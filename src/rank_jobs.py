import hashlib
import html
import json
import re
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pandas as pd

from src.ai_job_reviewer import (
    LEGACY_VALID_AI_REVIEW_SOURCES,
    apply_ai_batch_reviews,
    apply_ai_reviews,
    has_valid_ai_result,
    is_blank_value,
    is_truthy_value,
    manual_job_needs_ai_review,
)

try:
    import config
except ImportError:  # pragma: no cover - config exists in normal runs.
    config = None


# These paths are relative to the main project folder.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
RAW_JOBS_DIR = PROJECT_ROOT / "jobs" / "raw_jobs"
API_JOBS_FILE = PROJECT_ROOT / "jobs" / "api_jobs.csv"
PROFILE_FILE = PROJECT_ROOT / "profile" / "profile.json"
COMPANY_ALIASES_FILE = PROJECT_ROOT / "config" / "company_aliases.csv"
OUTPUT_FILE = PROJECT_ROOT / "output" / "ranked_jobs.xlsx"
ALL_OUTPUT_FILE = PROJECT_ROOT / "output" / "all_ranked_jobs.xlsx"
AI_REVIEW_QUEUE_FILE = PROJECT_ROOT / "output" / "ai_review_queue.xlsx"
MANUAL_AI_COVERAGE_AUDIT_FILE = PROJECT_ROOT / "output" / "manual_ai_coverage_audit.xlsx"
MANUAL_PRE_AI_EXCLUSION_AUDIT_FILE = PROJECT_ROOT / "output" / "manual_pre_ai_exclusion_audit.xlsx"
RANKED_JOBS_EXCLUSION_AUDIT_FILE = PROJECT_ROOT / "output" / "ranked_jobs_exclusion_audit.xlsx"
PIPELINE_DEBUG_AUDIT_FILE = PROJECT_ROOT / "output" / "pipeline_debug_audit.xlsx"
REPOST_TRACKER_AUDIT_FILE = PROJECT_ROOT / "output" / "repost_tracker_audit.xlsx"
API_LEADS_FILE = PROJECT_ROOT / "output" / "api_leads.xlsx"
TRACKER_FILE = PROJECT_ROOT / "tracker" / "applications.xlsx"
TRACKER_HISTORY_FILE = PROJECT_ROOT / "tracker" / "applications_history.xlsx"
STRICT_MODE = True
TRACKER_EXCLUDED_STATUSES = {
    "applied",
    "assessment",
    "interview",
    "final_interview",
    "offer",
    "rejected",
    "withdrawn",
    "expired",
    "closed",
}
TRACKER_SIMILAR_REVIEW_STATUSES = TRACKER_EXCLUDED_STATUSES
OPTIONAL_TRACKER_MATCH_COLUMNS = [
    "job_posted_date",
    "posted_date",
    "canonical_job_id",
    "location",
    "job_text",
    "job_description",
    "description_hash",
]


def cfg(name, default):
    return getattr(config, name, default) if config else default

# Simple UK locations to look for inside pasted job descriptions.
UK_LOCATIONS = [
    "Milton Keynes",
    "London",
    "Birmingham",
    "Luton",
    "Bedford",
    "Northampton",
    "Remote",
    "Hybrid",
]

# Role keywords that should fit this job search well.
GOOD_ROLE_KEYWORDS = [
    "junior data analyst",
    "graduate data analyst",
    "data analyst",
    "reporting analyst",
    "mi analyst",
    "bi analyst",
    "power bi analyst",
    "commercial analyst",
    "operations analyst",
    "insight analyst",
    "product analyst",
    "customer insight analyst",
    "people data analyst",
    "workforce analyst",
]

# Assistant roles can still be useful, but they should not score like analyst roles
# unless the description clearly includes data analysis work.
ASSISTANT_ROLE_KEYWORDS = [
    "assistant",
    "ecommerce assistant",
    "e-commerce assistant",
]

# Technical keywords for analyst and BI work.
TECHNICAL_KEYWORDS = [
    "sql",
    "kql",
    "power bi",
    "powerbi",
    "excel",
    "python",
    "bigquery",
    "dax",
    "power query",
    "dashboard",
    "dashboards",
    "large datasets",
    "metrics",
    "telemetry",
    "data visualisation",
    "data visualization",
    "reporting tools",
    "tableau",
    "looker",
    "analytics",
    "analysis",
    "stakeholder insights",
    "reporting",
    "kpi",
    "kpis",
    "data analysis",
]

# Business keywords that match commercial, retail and operations experience.
BUSINESS_KEYWORDS = [
    "commercial",
    "sales",
    "revenue",
    "customer",
    "retail",
    "ecommerce",
    "e-commerce",
    "logistics",
    "supply chain",
    "operations",
    "warehouse",
    "transport",
    "crm",
]

# These words show that a Commercial Analyst job is actually analytical.
COMMERCIAL_ANALYST_SIGNALS = [
    "data analysis",
    "analysis",
    "analytics",
    "reporting",
    "sql",
    "power bi",
    "powerbi",
    "dashboard",
    "dashboards",
    "kpi",
    "kpis",
    "revenue analysis",
    "sales analysis",
    "customer analysis",
]

# Words that suggest the role may be friendly to junior candidates.
JUNIOR_KEYWORDS = [
    "junior",
    "entry level",
    "graduate",
    "trainee",
    "1 year",
    "willingness to learn",
    "training",
    "grow",
]

SENIOR_TITLE_KEYWORDS = ["senior", "lead", "manager", "head of"]
SENIOR_LEADERSHIP_TITLE_PATTERNS = [
    r"\bhead\s+of\b",
    r"\bdirector\b",
    r"\blead\s+manager\b",
    r"\bsenior\s+manager\b",
    r"\bprincipal\b",
    r"\bstaff\b",
    r"\bvp\b",
    r"\bvice\s+president\b",
    r"\bchief\b",
    r"\bc[-\s]?suite\b",
    r"\bcto\b",
    r"\bcfo\b",
    r"\bcoo\b",
    r"\bcio\b",
]

BAD_FIT_KEYWORDS = [
    "marketing assistant",
    "admin assistant",
    "customer service",
    "sales assistant",
    "product upload",
    "product admin",
]

TENDER_ADMIN_KEYWORDS = [
    "tender",
    "tenders",
    "rfq",
    "rfqs",
    "proposal",
    "proposals",
    "crm updates",
    "sharepoint folders",
    "powerpoint presentations",
    "document templates",
    "admin support",
]

ASSISTANT_WEAK_FIT_KEYWORDS = [
    "merchandiser",
    "product upload",
    "uploading products",
    "product listings",
    "product admin",
]

# These title words keep low-scoring but possibly relevant analyst roles visible.
TARGET_TITLE_KEYWORDS = [
    "analyst",
    "data",
    "reporting",
    "bi",
    "mi",
    "insight",
    "commercial",
    "operations",
]

# These are common API results that are usually not relevant to this search.
IRRELEVANT_TITLE_KEYWORDS = [
    "sales assistant",
    "store assistant",
    "warehouse operative",
    "driver",
    "chef",
    "care assistant",
    "support worker",
    "teacher",
    "nurse",
    "receptionist",
    "cleaner",
]

# These are usually paid training/career-change programmes rather than normal
# job vacancies, so they should not compete with real analyst jobs.
TRAINING_PROGRAMME_KEYWORDS = [
    "placement programme",
    "no experience needed",
    "training provider",
    "it online learning",
    "newto training",
    "trainee programme",
    "career programme",
    "training course",
    "career package",
    "recruitment support package",
    "study fees",
    "finance terms",
    "money back guarantee",
    "comptia data+",
    "qualification package",
    "become a data analyst",
    "start your new career in data analysis",
]

# Strong data keywords can rescue an otherwise questionable title.
STRONG_DATA_KEYWORDS = [
    "sql",
    "power bi",
    "powerbi",
    "python",
    "dashboard",
    "dashboards",
    "data analysis",
    "analytics",
    "reporting",
    "kpi",
    "kpis",
]

# Recommendation rules. The script checks job titles first, then job text.
RECOMMENDATION_RULES = [
    {
        "role_type": "data analyst",
        "cv": "cvs/cv_data_analyst.pdf",
        "cover_letter": "cover_letters/cover_data_analyst.docx",
        "title_keywords": [
            "data analyst",
            "junior data analyst",
            "product data analyst",
            "analytics analyst",
        ],
        "description_keywords": [
            "sql",
            "python",
            "power bi",
            "powerbi",
            "bigquery",
            "data analysis",
        ],
    },
    {
        "role_type": "bi reporting analyst",
        "cv": "cvs/cv_bi_reporting_analyst.pdf",
        "cover_letter": "cover_letters/cover_bi_reporting_analyst.docx",
        "title_keywords": [
            "bi analyst",
            "reporting analyst",
            "mi analyst",
            "dashboard analyst",
        ],
        "description_keywords": [
            "power bi",
            "powerbi",
            "reporting",
            "dashboards",
            "dashboard",
            "kpis",
            "kpi",
            "dax",
            "power query",
        ],
    },
    {
        "role_type": "commercial analyst",
        "cv": "cvs/cv_commercial_analyst.pdf",
        "cover_letter": "cover_letters/cover_commercial_analyst.docx",
        "title_keywords": [
            "commercial analyst",
            "sales analyst",
            "revenue analyst",
            "customer analyst",
            "retail analyst",
        ],
        "description_keywords": [
            "commercial performance",
            "sales",
            "revenue",
            "customer behaviour",
            "customer behavior",
            "customer",
            "retail",
            "ecommerce",
            "e-commerce",
            "operations",
        ],
    },
    {
        "role_type": "business operations analyst",
        "cv": "cvs/cv_business_operations_analyst.pdf",
        "cover_letter": "cover_letters/cover_business_operations_analyst.docx",
        "title_keywords": [
            "business analyst",
            "operations analyst",
            "process analyst",
        ],
        "description_keywords": [
            "process improvement",
            "stakeholder communication",
            "operational reporting",
            "documentation",
        ],
    },
    {
        "role_type": "finance economics analyst",
        "cv": "cvs/cv_finance_economics_analyst.pdf",
        "cover_letter": "cover_letters/cover_finance_economics_analyst.docx",
        "title_keywords": [
            "finance analyst",
            "financial analyst",
            "economics analyst",
            "treasury analyst",
        ],
        "description_keywords": [
            "finance",
            "economics",
            "pricing",
            "forecasting",
            "budgeting",
            "financial reporting",
        ],
    },
]

# These are the status values the tracker understands. "new" is kept as an
# internal display value for jobs that are not yet in tracker/applications.xlsx.
VALID_STATUSES = [
    "new",
    "shortlisted",
    "applied",
    "assessment",
    "interview",
    "final_interview",
    "rejected",
    "offer",
    "withdrawn",
    "expired",
    "closed",
]

# These columns are used in tracker/applications.xlsx.
TRACKER_COLUMNS = [
    "job_id",
    "apply_link",
    "company",
    "job_title",
    "source",
    "record_source",
    "current_status",
    "highest_stage",
    "status_date",
    "apply_date",
    "cv_used",
    "cover_letter_used",
    "notes",
    "last_email_subject",
    "last_email_from",
    "last_email_date",
    "created_date",
    "updated_date",
]

API_COLUMNS = [
    "date_collected",
    "source",
    "job_title",
    "company",
    "location",
    "salary",
    "job_posted_date",
    "application_deadline",
    "contract_type",
    "job_description",
    "apply_link",
]


def load_profile():
    """Read the profile JSON file and return it as a Python dictionary."""
    with open(PROFILE_FILE, "r", encoding="utf-8") as file:
        return json.load(file)


def standardise_status(status):
    """Convert a status to one of the standard project status values."""
    status = str(status).strip().lower()

    # A few helpful conversions for old example statuses or typing variations.
    status_aliases = {
        "to apply": "new",
        "interested": "shortlisted",
        "maybe": "new",
        "apply": "applied",
        "a": "applied",
        "oa": "assessment",
        "as": "assessment",
        "online_assessment": "assessment",
        "i": "interview",
        "f": "final_interview",
        "fi": "final_interview",
        "r": "rejected",
        "o": "offer",
        "w": "withdrawn",
        "e": "expired",
        "expired": "expired",
        "closed": "closed",
        "c": "closed",
        "not interested": "not_interested",
        "not_interested": "not_interested",
        "skip": "skip",
        "s": "shortlisted",
    }

    status = status_aliases.get(status, status)

    if status in VALID_STATUSES:
        return status

    # If the status is blank or not recognised, keep it simple and use "new".
    return "new"


def detect_source_from_apply_link(apply_link, fallback_source="Unknown"):
    """Work out the job website from the apply link."""
    apply_link = str(apply_link).strip().lower()
    fallback_source = str(fallback_source).strip()

    if "linkedin.com" in apply_link:
        return "LinkedIn"
    if "indeed.com" in apply_link:
        return "Indeed"
    if "cord.co" in apply_link:
        return "CORD"
    if "reed.co.uk" in apply_link:
        return "Reed"
    if "civilservicejobs.service.gov.uk" in apply_link:
        return "Civil Service Jobs"
    if "totaljobs.com" in apply_link:
        return "Totaljobs"
    if "glassdoor.co.uk" in apply_link:
        return "Glassdoor"
    if "adzuna.co.uk" in apply_link:
        return "Adzuna"

    # Backward compatibility: older files may still have a source line.
    if fallback_source.lower() == "cord":
        return "CORD"
    if fallback_source and fallback_source != "Unknown":
        return fallback_source

    if apply_link:
        return "Manual"

    return "Unknown"


def clean_apply_link(apply_link):
    """Clean job URLs so duplicate links are easier to match."""
    apply_link = str(apply_link).strip()

    if "indeed." in apply_link.lower() and "vjk=" in apply_link.lower():
        parsed_url = urlparse(apply_link)
        query = parse_qs(parsed_url.query)
        vjk_values = query.get("vjk")

        if vjk_values and vjk_values[0]:
            return f"https://uk.indeed.com/viewjob?jk={vjk_values[0]}"

    return apply_link


def normalize_url_for_match(url):
    """Normalize job URLs for conservative exact matching."""
    url = clean_apply_link(url)
    text = str(url).strip()
    if not text:
        return ""
    parsed = urlparse(text)
    scheme = parsed.scheme.lower() or "https"
    netloc = parsed.netloc.lower()
    path = re.sub(r"/+$", "", parsed.path)
    query = parse_qs(parsed.query)
    useful_query = {}
    useful_query_keys = ["jk", "vjk", "currentJobId", "redirect_url"]
    if "cord.co" in netloc:
        useful_query_keys.extend(["jobId", "job_id", "job", "id"])
    for key in useful_query_keys:
        if key in query and query[key]:
            useful_query[key.lower()] = query[key][0]
    if useful_query:
        query_text = "&".join(f"{key}={useful_query[key]}" for key in sorted(useful_query))
        return f"{scheme}://{netloc}{path}?{query_text}".lower()
    return f"{scheme}://{netloc}{path}".lower()


def extract_url_job_identifiers(url):
    """Extract reliable external job ids from common UK job URLs."""
    text = str(url or "").strip()
    if not text:
        return set()
    parsed = urlparse(text)
    query = parse_qs(parsed.query)
    identifiers = set()
    host = parsed.netloc.lower()
    path = parsed.path.lower()

    for key in ["currentJobId", "jk", "vjk"]:
        for value in query.get(key, []):
            if str(value).strip():
                identifiers.add(f"{key.lower()}:{str(value).strip().lower()}")

    if "linkedin." in host:
        match = re.search(r"(?:currentjobid=|/jobs/view/)(\d+)", text, flags=re.IGNORECASE)
        if match:
            identifiers.add(f"linkedin:{match.group(1).lower()}")

    if "indeed." in host:
        match = re.search(r"(?:jk=|vjk=)([A-Za-z0-9_-]+)", text, flags=re.IGNORECASE)
        if match:
            identifiers.add(f"indeed:{match.group(1).lower()}")

    if "reed.co.uk" in host:
        match = re.search(r"/jobs/[^/?#]*/(\d+)", path)
        if not match:
            match = re.search(r"/jobs/(\d+)", path)
        if match:
            identifiers.add(f"reed:{match.group(1).lower()}")

    if "adzuna." in host:
        match = re.search(r"/details/(\d+)", path)
        if match:
            identifiers.add(f"adzuna:{match.group(1).lower()}")

    if "cord.co" in host:
        for key in ["jobId", "job_id", "job", "id"]:
            for value in query.get(key, []):
                if str(value).strip():
                    identifiers.add(f"cord:{str(value).strip().lower()}")
        for pattern in [
            r"/jobs/([a-z0-9_-]+)",
            r"/job/([a-z0-9_-]+)",
            r"/opportunities/([a-z0-9_-]+)",
            r"/opportunity/([a-z0-9_-]+)",
            r"/companies/[^/]+/jobs/([a-z0-9_-]+)",
            r"/company/[^/]+/jobs/([a-z0-9_-]+)",
        ]:
            match = re.search(pattern, path)
            if match:
                identifiers.add(f"cord:{match.group(1).lower()}")

    return identifiers


URL_PATTERN = re.compile(r"https?://\S+", flags=re.IGNORECASE)
METADATA_LABELS = {
    "job_title": "job_title",
    "job title": "job_title",
    "title": "job_title",
    "company": "company",
    "location": "location",
    "salary": "salary",
    "source": "source",
    "apply_link": "apply_link",
    "apply link": "apply_link",
    "status": "status",
}
FULL_JD_LABELS = {"full jd", "full job description", "job description", "description"}
PREFERRED_TITLE_PATTERNS = [
    "junior data analyst",
    "graduate data analyst",
    "data analyst",
    "reporting analyst",
    "mi analyst",
    "bi analyst",
    "power bi analyst",
    "commercial analyst",
    "operations analyst",
    "insight analyst",
    "product analyst",
    "customer insight analyst",
    "people data analyst",
    "workforce analyst",
]
LOCATION_PATTERN = re.compile(
    r"\b(?:Milton Keynes|Northampton|Bedford|Luton|Birmingham|London|Remote|Hybrid|United Kingdom|UK)"
    r"(?:\s*[,/|-]\s*(?:Remote|Hybrid|United Kingdom|UK|Milton Keynes|Northampton|Bedford|Luton|Birmingham|London))*\b",
    flags=re.IGNORECASE,
)
SALARY_PATTERN = re.compile(
    r"(?:£\s?\d{2,3}(?:,\d{3})?(?:\.\d+)?\s?k?(?:\s*(?:-|to)\s*£?\s?\d{2,3}(?:,\d{3})?(?:\.\d+)?\s?k?)?"
    r"(?:\s*(?:per hour|per annum|a year|pa|p\.a\.))?)|"
    r"(?:\b\d{5,6}\b(?:\s*(?:-|to)\s*\b\d{5,6}\b)?)",
    flags=re.IGNORECASE,
)


def parse_metadata_line(line):
    match = re.match(r"^\s*([A-Za-z_ ]{2,30})\s*:\s*(.*)$", str(line))
    if not match:
        return "", ""
    label = match.group(1).strip().lower().replace("-", " ")
    label = re.sub(r"\s+", " ", label)
    return METADATA_LABELS.get(label, ""), match.group(2).strip()


def strip_urls(line):
    return URL_PATTERN.sub("", str(line)).strip()


def clean_manual_lines_for_extraction(full_text):
    lines = []
    for line in str(full_text).splitlines():
        key, value = parse_metadata_line(line)
        if key in ["job_title", "company", "location", "salary"]:
            if value:
                lines.append(value)
            continue
        if key or line.strip().lower().rstrip(":") in FULL_JD_LABELS:
            continue
        cleaned = strip_urls(line)
        if cleaned:
            lines.append(cleaned)
    return lines


def extract_first_url(full_text):
    match = URL_PATTERN.search(str(full_text))
    return match.group(0).rstrip(").,]") if match else ""


def extract_salary_from_text(text):
    original_text = str(text)
    normalized_text = original_text.lower().replace(",", "")
    reject_terms = [
        "rsv",
        "brand",
        "revenue",
        "turnover",
        "sales",
        "bonus",
        "referral bonus",
        "voucher",
        "vouchers",
        "benefit",
        "benefits",
        "pension",
        "contribution",
    ]
    salary_context_terms = ["salary", "pay", "per year", "a year", "per annum", "pa", "p.a.", "annual"]
    patterns = [
        r"£\s?\d{2,3}(?:,\d{3})?(?:\.\d+)?\s?k\b(?:\s*(?:-|to)\s*£?\s?\d{2,3}(?:,\d{3})?(?:\.\d+)?\s?k\b)?",
        r"£\s?\d{2,3}(?:,\d{3})(?:\s*(?:-|to)\s*£?\s?\d{2,3}(?:,\d{3}))?",
        r"£\s?\d{2,3}(?:,\d{3})?\s*(?:per hour|per annum|a year|pa|p\.a\.)",
        r"\b\d{5,6}\b(?:\s*(?:-|to)\s*\b\d{5,6}\b)?",
    ]
    for pattern in patterns:
        for match in re.finditer(pattern, original_text, flags=re.IGNORECASE):
            lower_match = match.group(0).lower()
            window = normalized_text[max(0, match.start() - 80) : min(len(normalized_text), match.end() + 80)]
            if re.search(r"£\s?\d+(?:\.\d+)?\s*(?:m|million|bn|billion)\b", lower_match):
                continue
            if any(term in window for term in reject_terms) and not any(term in window for term in salary_context_terms):
                continue
            if not lower_match.startswith("£") and not any(term in window for term in salary_context_terms):
                continue
            return match.group(0).strip()
    return ""


def extract_manual_title(full_text, metadata, file_name):
    if metadata.get("job_title"):
        return metadata["job_title"], "High", ""
    lines = clean_manual_lines_for_extraction(full_text)[:30]
    for line in lines:
        line_lower = line.lower()
        if any(title in line_lower for title in PREFERRED_TITLE_PATTERNS):
            return line, "Medium", ""
    for line in lines:
        if re.search(r"\banalyst\b", line, flags=re.IGNORECASE):
            return re.sub(r"\s*-\s*job post\s*$", "", line, flags=re.IGNORECASE).strip(), "Medium", ""
    fallback = Path(file_name).stem.replace("_", " ").replace("-", " ").strip()
    return fallback or "Unknown", "Low", "title inferred from filename"


def extract_manual_company(full_text, metadata, title):
    if metadata.get("company"):
        return metadata["company"], "High", ""
    lines = clean_manual_lines_for_extraction(full_text)[:10]
    ignored = {str(title).strip().lower(), ""}
    ui_terms = {
        "share",
        "show more options",
        "easy apply",
        "save",
        "promoted by hirer",
        "your profile matches several required qualifications",
    }
    for line in lines:
        line = re.sub(r"\s+logo\s*$", "", line, flags=re.IGNORECASE).strip()
        line_lower = line.lower()
        if line_lower in ignored:
            continue
        if line_lower in ui_terms or line_lower.startswith("save "):
            continue
        if LOCATION_PATTERN.search(line) or SALARY_PATTERN.search(line):
            continue
        if any(token in line_lower for token in ["full job description", "job description", "apply now"]):
            continue
        return line, "Medium", ""
    return "Unknown", "Low", "company not found near top of file"


def extract_manual_location(full_text, metadata):
    if metadata.get("location"):
        return metadata["location"], "High", ""
    lines = clean_manual_lines_for_extraction(full_text)
    for line in lines[:40]:
        match = LOCATION_PATTERN.search(line)
        if match:
            return match.group(0).strip(), "Medium", ""
    match = LOCATION_PATTERN.search(str(full_text))
    if match:
        return match.group(0).strip(), "Medium", ""
    return "Unknown", "Low", "location not found"


def build_manual_job_text(full_text):
    description_lines = []
    for line in str(full_text).splitlines():
        clean_line = line.strip()
        key, value = parse_metadata_line(clean_line)
        label = clean_line.lower().rstrip(":")
        if key:
            continue
        if label in FULL_JD_LABELS:
            continue
        without_url = strip_urls(line)
        if without_url:
            description_lines.append(without_url)
    return "\n".join(description_lines).strip()


def parse_job_file(job_path):
    """Read one .txt file and extract metadata from either structured or URL-first text."""
    with open(job_path, "r", encoding="utf-8") as file:
        full_text = file.read()

    metadata = {}
    for line in full_text.splitlines():
        key, value = parse_metadata_line(line)
        if key:
            metadata[key] = value

    apply_link = clean_apply_link(metadata.get("apply_link", "") or extract_first_url(full_text))
    old_source = metadata.get("source", "Unknown")
    job_text = build_manual_job_text(full_text)
    job_title, title_confidence, title_note = extract_manual_title(full_text, metadata, job_path.name)
    company, company_confidence, company_note = extract_manual_company(full_text, metadata, job_title)
    location, location_confidence, location_note = extract_manual_location(full_text, metadata)
    salary = metadata.get("salary", "") or extract_salary_from_text(full_text)
    confidence_values = [title_confidence, company_confidence, location_confidence]
    low_count = confidence_values.count("Low")
    manual_parse_confidence = "Low" if low_count >= 2 else "Medium" if low_count == 1 else "High"
    notes = [note for note in [title_note, company_note, location_note] if note]
    if not apply_link:
        notes.append("apply link not found")
    if len(job_text) < 300:
        notes.append("job description text is short")
    if not job_text:
        job_text = full_text.strip()
        notes.append("job text fallback used full raw file content")

    return {
        "input_type": "manual",
        "file_name": job_path.name,
        "date_added": "",
        "source": detect_source_from_apply_link(apply_link, old_source),
        "apply_link": apply_link,
        "salary": salary,
        "contract_type": "",
        "job_title": job_title,
        "company": company,
        "location": location,
        "status": standardise_status(metadata.get("status", "new")),
        "job_posted_date": "Unknown",
        "job_age_days": "",
        "freshness": "Unknown",
        "posted_date_confidence": "none",
        "application_deadline": "",
        "deadline_confidence": "none",
        "job_text": job_text,
        "job_description": job_text,
        "manual_parse_confidence": manual_parse_confidence,
        "manual_parse_notes": "; ".join(notes),
        "needs_human_review": low_count >= 2,
    }


def fallback_manual_job(job_path, error):
    try:
        full_text = job_path.read_text(encoding="utf-8", errors="ignore")
    except OSError:
        full_text = ""
    apply_link = clean_apply_link(extract_first_url(full_text))
    file_title = job_path.stem.replace("_", " ").replace("-", " ").strip() or "Unknown"
    notes = [
        f"parser fallback used after error: {type(error).__name__}: {error}",
        "company not found",
        "location not found",
    ]
    if not apply_link:
        notes.append("apply link not found")
    return {
        "input_type": "manual",
        "file_name": job_path.name,
        "date_added": "",
        "source": detect_source_from_apply_link(apply_link, "Manual"),
        "apply_link": apply_link,
        "salary": "",
        "contract_type": "",
        "job_title": file_title,
        "company": "Unknown",
        "location": "Unknown",
        "status": "new",
        "job_posted_date": "Unknown",
        "job_age_days": "",
        "freshness": "Unknown",
        "posted_date_confidence": "none",
        "application_deadline": "",
        "deadline_confidence": "none",
        "job_text": full_text,
        "job_description": full_text,
        "manual_parse_confidence": "Low",
        "manual_parse_notes": "; ".join(notes),
        "needs_human_review": True,
    }


def load_raw_jobs():
    """Scan jobs/raw_jobs/ and load every .txt file as one job."""
    job_rows = []
    parse_failures = []
    raw_files = sorted(RAW_JOBS_DIR.glob("*.txt"))

    for job_path in raw_files:
        try:
            job_rows.append(parse_job_file(job_path))
        except Exception as error:
            parse_failures.append({"file_name": job_path.name, "reason": str(error)})
            job_rows.append(fallback_manual_job(job_path, error))

    raw_jobs = pd.DataFrame(job_rows)
    raw_jobs.attrs["raw_files_found"] = len(raw_files)
    raw_jobs.attrs["raw_files_parsed"] = len(job_rows)
    raw_jobs.attrs["raw_parse_failures"] = parse_failures
    return raw_jobs


def get_raw_job_files_modified_today():
    today = date.today()
    return {
        raw_file.name
        for raw_file in RAW_JOBS_DIR.glob("*.txt")
        if datetime.fromtimestamp(raw_file.stat().st_mtime).date() == today
    }


def load_api_jobs():
    """Read jobs collected from APIs and prepare them for ranking."""
    if not API_JOBS_FILE.exists():
        return pd.DataFrame()

    api_jobs = pd.read_csv(API_JOBS_FILE)

    for column in API_COLUMNS:
        if column not in api_jobs.columns:
            api_jobs[column] = ""

    api_jobs = api_jobs[API_COLUMNS].copy().fillna("")
    api_jobs["input_type"] = "api"
    api_jobs["file_name"] = ""
    api_jobs["manual_parse_confidence"] = ""
    api_jobs["manual_parse_notes"] = ""
    api_jobs["apply_link"] = api_jobs["apply_link"].apply(clean_apply_link)
    api_jobs["source"] = api_jobs.apply(
        lambda job: job["source"] or detect_source_from_apply_link(job["apply_link"], "API"),
        axis=1,
    )
    api_jobs["job_text"] = api_jobs.apply(
        lambda job: "\n".join(
            [
                str(job.get("job_title", "")),
                str(job.get("company", "")),
                str(job.get("location", "")),
                str(job.get("contract_type", "")),
                str(job.get("job_description", "")),
            ]
        ).strip(),
        axis=1,
    )

    api_jobs["date_added"] = api_jobs["date_collected"]
    return api_jobs[
        [
            "input_type",
            "file_name",
            "date_added",
            "source",
            "apply_link",
            "salary",
            "job_posted_date",
            "application_deadline",
            "contract_type",
            "job_text",
            "job_description",
            "job_title",
            "company",
            "location",
            "manual_parse_confidence",
            "manual_parse_notes",
        ]
    ]


def load_all_jobs():
    """Load manual pasted jobs and API jobs into one table."""
    manual_jobs = load_raw_jobs()
    api_jobs = load_api_jobs()

    if manual_jobs.empty and api_jobs.empty:
        return pd.DataFrame()

    jobs = pd.concat([manual_jobs, api_jobs], ignore_index=True)
    jobs.attrs["manual_load_summary"] = {
        "raw_files_found": manual_jobs.attrs.get("raw_files_found", 0),
        "raw_files_parsed": manual_jobs.attrs.get("raw_files_parsed", 0),
        "raw_parse_failures": manual_jobs.attrs.get("raw_parse_failures", []),
    }
    return jobs


def count_matches(text, keywords):
    """Count how many profile keywords appear in the job text."""
    text = str(text).lower()
    matches = 0

    for keyword in keywords:
        if str(keyword).lower() in text:
            matches += 1

    return matches


def find_matches(text, keywords):
    """Return the keywords found in the text."""
    text = str(text).lower()
    matches = []

    for keyword in keywords:
        if keyword.lower() in text:
            matches.append(keyword)

    return matches


CONTRACT_RED_FLAG_KEYWORDS = [
    "contract",
    "fixed term",
    "fixed-term",
    "ftc",
    "temporary",
    "temp",
    "interim",
    "freelance",
    "day rate",
    "daily rate",
    "per day",
    "inside ir35",
    "outside ir35",
    "3 month",
    "6 month",
    "12 month",
    "maternity cover",
    "cover role",
]


def combined_contract_text(job):
    fields = [
        "job_title",
        "job_text",
        "job_description",
        "contract_type",
        "salary",
        "source",
        "file_name",
    ]
    return "\n".join(str(job.get(field, "")) for field in fields)


def contract_keyword_matches(job):
    text = combined_contract_text(job).lower()
    permanent_patterns = [
        r"\bcontract\s*:\s*permanent\b",
        r"\bjob type\s*:?\s*(?:\n|\s|&nbsp;)*permanent\b",
        r"\bpermanent\s*,\s*full[-\s]time\b",
        r"\bfull[-\s]time\s*,\s*permanent\b",
        r"\bpermanent\b",
    ]
    is_permanent = any(re.search(pattern, text) for pattern in permanent_patterns)
    patterns = [
        r"\bfixed[-\s]term\b",
        r"\bftc\b",
        r"\btemporary\b",
        r"\btemp(?:orary)?\b",
        r"\binterim\b",
        r"\bfreelance\b",
        r"\bday rate\b",
        r"\bdaily rate\b",
        r"\bper day\b",
        r"\binside ir35\b",
        r"\boutside ir35\b",
        r"\b3[-\s]month\s+contract\b",
        r"\b6[-\s]month\s+contract\b",
        r"\b12[-\s]month\s+contract\b",
        r"\bmaternity cover\b",
        r"\bcover role\b",
    ]
    matches = [pattern for pattern in patterns if re.search(pattern, text)]
    if not is_permanent and re.search(r"\bcontract\b", text):
        matches.append(r"\bcontract\b")
    return matches


def has_contract_keyword(job):
    return bool(contract_keyword_matches(job))


def has_day_rate_salary(job):
    text = combined_contract_text(job).lower()
    salary_context = "\n".join(
        [
            str(job.get("salary", "")),
            str(job.get("job_title", "")),
            str(job.get("contract_type", "")),
            str(job.get("source", "")),
        ]
    ).lower()
    if re.search(r"(?:\D|^)\d{3}(?:\.0)?\s*[-]\s*(?:\D)?\d{3}(?:\.0)?", salary_context):
        return True
    if re.search(r"(?:\D|^)\d{3}(?:\.0)?\s*(?:per\s+day|/day|pd|p\.d\.)", text):
        return True
    return "day rate" in text or "daily rate" in text


def is_contract_or_temporary_role(job):
    return has_contract_keyword(job) or has_day_rate_salary(job)


JUNIOR_FRIENDLY_TITLE_WORDS = [
    "junior",
    "graduate",
    "entry",
    "trainee",
    "assistant",
    "associate",
]


EXPERIENCE_SENIORITY_PHRASES = [
    "extensive experience",
    "proven experience",
    "lead",
    "senior stakeholder",
    "own the reporting function",
    "manage junior analysts",
]


def has_junior_friendly_title(job):
    title = str(job.get("job_title", "")).lower()
    return any(re.search(rf"\b{re.escape(word)}\b", title) for word in JUNIOR_FRIENDLY_TITLE_WORDS)


def salary_values_from_text(text):
    values = []
    original_text = str(text)
    normalized_text = original_text.lower().replace(",", "")
    reject_terms = [
        "rsv",
        "brand",
        "brand value",
        "revenue",
        "turnover",
        "sales",
        "bonus",
        "referral bonus",
        "voucher",
        "vouchers",
        "benefit",
        "benefits",
        "pension",
        "contribution",
    ]
    salary_context_terms = ["salary", "pay", "per year", "a year", "per annum", "pa", "p.a.", "annual"]

    def is_rejected_context(match):
        window = normalized_text[max(0, match.start() - 80) : min(len(normalized_text), match.end() + 80)]
        if re.search(r"£\s?\d+(?:\.\d+)?\s*(?:m|million|bn|billion)\b", window):
            return True
        return any(term in window for term in reject_terms) and not any(term in window for term in salary_context_terms)

    def has_salary_context(match):
        window = normalized_text[max(0, match.start() - 80) : min(len(normalized_text), match.end() + 80)]
        return any(term in window for term in salary_context_terms)

    salary_patterns = [
        (r"£\s?(\d{2,3}(?:\.\d+)?)\s*k\b", 1000),
        (r"\b(\d{2,3}(?:\.\d+)?)\s*k\b", 1000),
        (r"£\s?(\d{5,6})(?:\.\d+)?\b", 1),
        (r"\b(\d{5,6})(?:\.\d+)?\b", 1),
    ]
    for pattern, multiplier in salary_patterns:
        for match in re.finditer(pattern, normalized_text):
            if is_rejected_context(match):
                continue
            if pattern.startswith(r"\b") and not has_salary_context(match):
                continue
            values.append(float(match.group(1)) * multiplier)
    return values


def get_salary_max(job):
    if has_day_rate_salary(job):
        return 0

    text = "\n".join(
        [
            str(job.get("salary", "")),
            str(job.get("job_title", "")),
            str(job.get("job_text", "")),
            str(job.get("job_description", "")),
        ]
    )
    values = salary_values_from_text(text)
    return int(max(values)) if values else 0


def has_high_salary_seniority_risk(job):
    salary_max = get_salary_max(job)
    return salary_max >= 50000 and not has_junior_friendly_title(job)


def has_senior_experience_wording(job):
    if has_junior_friendly_title(job):
        return False

    text = f"{job.get('job_title', '')}\n{job.get('job_text', '')}\n{job.get('job_description', '')}".lower()
    if re.search(r"\b[3-9]\+?\s+years?\b", text):
        return True
    return bool(find_matches(text, EXPERIENCE_SENIORITY_PHRASES))


def get_salary_seniority_risk(job):
    salary_max = get_salary_max(job)
    if not salary_max or has_junior_friendly_title(job):
        return ""
    if salary_max >= 60000:
        return "salary level suggests senior/mid-level role"
    if salary_max >= 50000:
        return "high salary suggests mid/senior level"
    if salary_max >= 45000:
        return "salary suggests possible mid-level role"
    return ""


def get_contract_risk(job):
    if is_contract_or_temporary_role(job):
        return "contract/FTC/temporary/day-rate role"
    return ""


def get_junior_fit_reason(job):
    title = str(job.get("job_title", ""))
    text = f"{title}\n{job.get('job_text', '')}"
    if has_junior_friendly_title(job):
        return "title contains junior-friendly wording"
    if has_target_title_keyword(title) and get_salary_max(job) <= 45000 and not has_senior_experience_wording(job):
        return "realistic analyst title and salary range"
    if has_senior_experience_wording(job):
        return "experience wording suggests mid/senior level"
    if find_matches(text, ["sql", "power bi", "excel", "reporting", "dashboard", "kpi"]):
        return "skills are relevant but junior fit needs review"
    return "junior fit unclear"


def has_developer_heavy_role(job):
    title = str(job.get("job_title", "")).lower()
    text = f"{title}\n{job.get('job_text', '')}".lower()
    developer_title = any(
        phrase in title
        for phrase in ["data engineer", "analytics engineer", "bi developer", "business intelligence developer"]
    )
    developer_heavy = len(find_matches(text, ["etl", "data pipeline", "data pipelines", "azure data factory", "dbt", "spark"])) >= 2
    return developer_title or developer_heavy


def has_domain_specific_hard_skip(job):
    text = f"{job.get('job_title', '')}\n{job.get('job_text', '')}\n{job.get('job_description', '')}".lower()
    domain_patterns = [
        r"\binsurance broking\b",
        r"\bacturis\b",
        r"\bsm&cr\b",
        r"\bt&c\b",
        r"\bcpd\b",
        r"\bcii\b",
        r"\bcisi\b",
        r"\bproject controls\b",
        r"\bevm\b",
        r"\blinear infrastructure\b",
    ]
    if any(re.search(pattern, text) for pattern in domain_patterns):
        return True
    if "hr" in text and "financial services" in text and any(
        re.search(pattern, text) for pattern in [r"\bsm&cr\b", r"\bt&c\b", r"\bcpd\b", r"\bcii\b", r"\bcisi\b"]
    ):
        return True
    if find_matches(text, ["crm analytics", "marketing analytics"]) and re.search(r"\b[3-9]\+?\s+years?\b", text):
        return True
    if re.search(r"\b[3-9]\+?\s+years?\b", text) and find_matches(
        text,
        ["insurance", "construction", "project controls", "acturis", "marketing analytics", "crm analytics"],
    ):
        return True
    return False


def has_core_skill_match(job):
    text = f"{job.get('job_title', '')}\n{job.get('job_text', '')}\n{job.get('job_description', '')}"
    return bool(find_matches(text, ["excel", "sql", "power bi", "powerbi", "dashboard", "dashboards", "reporting", "kpi", "kpis", "data quality"]))


def is_realistic_location(job):
    location_text = f"{job.get('location', '')}\n{job.get('job_text', '')}".lower()
    realistic_locations = [
        "milton keynes",
        "london",
        "remote",
        "hybrid",
        "luton",
        "bedford",
        "northampton",
        "birmingham",
        "united kingdom",
        " uk",
    ]
    return any(location in location_text for location in realistic_locations)


def get_location_fit_score(job):
    location_text = f"{job.get('location', '')}\n{job.get('job_text', '')}".lower()
    if "remote" in location_text:
        return 100
    if "hybrid" in location_text and any(city in location_text for city in ["milton keynes", "london", "birmingham", "northampton"]):
        return 95
    if "milton keynes" in location_text:
        return 90
    if any(city in location_text for city in ["northampton", "birmingham", "london", "luton", "bedford"]):
        return 75
    if "united kingdom" in location_text or " uk" in location_text:
        return 65
    return 20


def calculate_apply_priority_score(job):
    score = 0
    salary_max = get_salary_max(job)
    red_flags = str(job.get("red_flags", ""))
    text = f"{job.get('job_title', '')}\n{job.get('job_text', '')}"

    if has_target_title_keyword(job.get("job_title", "")):
        score += 25
    if has_junior_friendly_title(job):
        score += 18
    elif salary_max and salary_max <= 45000 and not has_senior_experience_wording(job):
        score += 10
    if not is_contract_or_temporary_role(job):
        score += 15
    if not salary_max:
        score += 6
    elif 28000 <= salary_max <= 45000:
        score += 15
    elif salary_max <= 35000:
        score += 12
    elif salary_max <= 50000:
        score += 4
    skill_matches = find_matches(text, ["sql", "power bi", "powerbi", "excel", "reporting", "dashboard", "dashboards", "kpi", "kpis"])
    score += min(18, len(skill_matches) * 4)
    if is_realistic_location(job):
        score += 10
    business_matches = find_matches(text, BUSINESS_KEYWORDS)
    score += min(10, len(business_matches) * 2)

    if "senior / lead / manager / head of" in red_flags:
        score -= 35
    if get_salary_seniority_risk(job):
        score -= 20
    if is_contract_or_temporary_role(job):
        score -= 60
    if str(job.get("status", "")) in ["expired", "closed"]:
        score -= 60
    if not is_realistic_location(job):
        score -= 10
    if has_senior_experience_wording(job):
        score -= 20
    if has_developer_heavy_role(job):
        score -= 45

    return max(0, min(100, int(score)))


def get_action_bucket(job):
    status = standardise_status(job.get("status", "new"))
    salary_max = get_salary_max(job)
    red_flags = str(job.get("red_flags", ""))

    if (
        is_contract_or_temporary_role(job)
        or status in ["expired", "closed"]
        or "senior / lead / manager / head of" in red_flags
        or (salary_max >= 60000 and not has_junior_friendly_title(job))
        or has_developer_heavy_role(job)
    ):
        return "Skip"

    if (
        (salary_max >= 50000 and not has_junior_friendly_title(job))
        or has_senior_experience_wording(job)
        or "high salary suggests mid/senior level" in red_flags
    ):
        return "Low Priority"

    priority_score = calculate_apply_priority_score(job)
    realistic_salary = not salary_max or salary_max <= 45000
    if (
        priority_score >= 72
        and realistic_salary
        and is_realistic_location(job)
        and not red_flags
    ):
        return "Apply Today"
    if priority_score >= 60:
        return "Strong Consider"
    if priority_score >= 40:
        return "Consider"
    return "Low Priority"


def get_action_reason(job):
    bucket = str(job.get("action_bucket", ""))
    if bucket == "Skip":
        if get_contract_risk(job):
            return "contract/FTC/temporary/day-rate role"
        if standardise_status(job.get("status", "new")) in ["expired", "closed"]:
            return "job is expired or closed"
        if has_developer_heavy_role(job):
            return "developer/engineering-heavy role"
        if "senior / lead / manager / head of" in str(job.get("red_flags", "")):
            return "senior title"
        return "not suitable for daily applications"
    if bucket == "Low Priority":
        return get_salary_seniority_risk(job) or get_junior_fit_reason(job)
    if bucket == "Apply Today":
        return "permanent junior-friendly analyst role with strong skills/location fit"
    return get_junior_fit_reason(job)


def get_action_bucket_rank(action_bucket):
    order = {
        "Apply Today": 1,
        "Strong Consider": 2,
        "Apply If Time": 3,
        "Pending AI Review": 4,
        "Manual Review": 4,
        "Consider": 4,
        "Low Priority": 5,
        "Skip": 6,
    }
    return order.get(str(action_bucket), 99)


def has_sponsorship_issue(job):
    text = f"{job.get('job_title', '')}\n{job.get('job_text', '')}\n{job.get('job_description', '')}".lower()
    sponsorship_phrases = [
        "unable to sponsor",
        "cannot sponsor",
        "do not sponsor",
        "does not sponsor",
        "will not sponsor",
        "must have the right to work",
        "must already have the right to work",
        "no sponsorship",
        "sponsorship is not available",
    ]
    return any(phrase in text for phrase in sponsorship_phrases)


def has_senior_leadership_title(job):
    title = str(job.get("job_title", "")).lower()
    return any(re.search(pattern, title) for pattern in SENIOR_LEADERSHIP_TITLE_PATTERNS)


def is_clearly_non_target_manual_role(job):
    title = str(job.get("job_title", ""))
    text = f"{title}\n{job.get('job_text', '')}\n{job.get('job_description', '')}"
    if has_target_title_keyword(title) or has_strong_data_keywords(text):
        return False
    return bool(find_matches(title, IRRELEVANT_TITLE_KEYWORDS + BAD_FIT_KEYWORDS))


def get_manual_pre_ai_exclusion(job):
    if not str(job.get("job_text", "")).strip():
        return "Empty or unreadable", "manual txt file is empty or unreadable"
    if bool(job.get("manual_duplicate", False)):
        return "Exact duplicate", "exact duplicate of another manual raw job"
    if bool(job.get("tracker_exact_match", False)) and bool(job.get("previously_seen", False)) and has_strong_exact_tracker_evidence(job):
        status = standardise_status(job.get("status", ""))
        return "Exact tracker exclusion", f"exact tracker match already exists as {status or 'existing tracker row'}"
    if is_contract_or_temporary_role(job):
        return "Contract / FTC skipped", "contract/FTC/temporary/day-rate role"
    if has_senior_leadership_title(job):
        return "Senior leadership skipped", "senior leadership title"
    if is_clearly_non_target_manual_role(job):
        return "Non-target role skipped", "clearly not an analyst/data/business/reporting/commercial/finance-related role"
    return "", ""


def get_manual_safety_skip_reason(job):
    _category, reason = get_manual_pre_ai_exclusion(job)
    return reason


def get_hard_skip_reason(job):
    status = standardise_status(job.get("status", "new"))
    salary_max = get_salary_max(job)
    red_flags = str(job.get("red_flags", ""))
    location_text = f"{job.get('location', '')}\n{job.get('job_text', '')}".lower()
    far_without_remote = "far location" in red_flags and not any(
        term in location_text for term in ["remote", "hybrid", "home based", "work from home"]
    )

    if status in ["expired", "closed"]:
        return "job is expired or closed"
    if (
        is_manual_job(job)
        and status in ["applied", "assessment", "interview", "final_interview", "offer", "rejected", "withdrawn", "not_interested", "skip"]
        and bool(job.get("tracker_exact_match", False))
        and not has_strong_exact_tracker_evidence(job)
    ):
        return ""
    if status in ["applied", "assessment", "interview", "final_interview", "offer", "rejected", "withdrawn", "not_interested", "skip"]:
        return f"already exists in tracker as {status}"
    if bool(job.get("previously_seen", False)):
        if is_manual_job(job) and bool(job.get("tracker_exact_match", False)) and not has_strong_exact_tracker_evidence(job):
            return ""
        return str(job.get("tracker_exclusion_reason", "")) or "already exists in tracker"
    if is_manual_job(job):
        if not str(job.get("job_text", "")).strip():
            return "manual txt file is empty or unreadable"
        if bool(job.get("manual_duplicate", False)):
            return "exact duplicate of another manual job"
        manual_safety_skip = get_manual_safety_skip_reason(job)
        if manual_safety_skip:
            return manual_safety_skip
        return ""

    if "training programme / not a normal job" in red_flags or has_training_programme_flag(job):
        return "training course / career package"
    if is_contract_or_temporary_role(job):
        return "contract/FTC/temporary/day-rate role"
    if "senior / lead / manager / head of" in red_flags:
        return "senior / lead / manager / head of title"
    if has_irrelevant_title_without_data(job):
        return "clearly non-data role"
    if salary_max >= 60000 and not has_junior_friendly_title(job):
        return "salary level suggests senior/mid-level role"
    if has_domain_specific_hard_skip(job):
        return "domain-specific experience requirement does not match target profile"
    if far_without_remote:
        return "far location with no realistic hybrid/remote setup"
    return ""


def is_api_job(job):
    return str(job.get("input_type", "")).lower() == "api"


def is_manual_job(job):
    return str(job.get("input_type", "")).lower() == "manual"


def has_valid_ai_decision(job):
    return has_valid_ai_result(job) and str(job.get("ai_final_action", "")) in [
        "Apply Today",
        "Strong Consider",
        "Apply If Time",
        "Consider",
        "Manual Review",
        "Low Priority",
        "Skip",
    ]


def is_extremely_strong_api_job(job):
    if not is_api_job(job):
        return False
    text = f"{job.get('job_title', '')}\n{job.get('location', '')}\n{job.get('job_text', '')}".lower()
    title_ok = has_junior_friendly_title(job) or any(
        phrase in text
        for phrase in ["junior data analyst", "graduate data analyst", "entry level data analyst", "assistant data analyst"]
    )
    location_ok = is_realistic_location(job)
    salary_max = get_salary_max(job)
    salary_ok = salary_max == 0 or 24000 <= salary_max <= 42000
    analyst_ok = has_pre_ai_target_title(job) and has_core_skill_match(job)
    return title_ok and location_ok and salary_ok and analyst_ok


def has_hard_skip_condition(job):
    return bool(get_hard_skip_reason(job))


def get_final_action(job):
    if get_hard_skip_reason(job):
        return "Skip"
    if is_manual_job(job) and not has_valid_ai_decision(job):
        return "Pending AI Review"
    if is_api_job(job) and not has_valid_ai_decision(job) and not is_extremely_strong_api_job(job):
        return "Consider" if is_pre_ai_borderline_candidate(job) else "Low Priority"

    legacy_skip = (
        str(job.get("apply_recommendation", "")).lower() == "skip"
        or "skip" in str(job.get("application_tier", "")).lower()
    )
    red_flags = str(job.get("red_flags", ""))
    risk_flags = [
        "weak technical analysis",
        "high salary suggests mid/senior level",
        "experience wording suggests mid/senior level",
    ]
    has_risk_flag = any(flag in red_flags for flag in risk_flags)
    strong_junior_signal = has_junior_friendly_title(job) and calculate_apply_priority_score(job) >= 55

    if legacy_skip:
        return "Low Priority"
    if has_risk_flag and not strong_junior_signal:
        return "Low Priority"

    salary_max = get_salary_max(job)
    text = f"{job.get('job_title', '')}\n{job.get('job_text', '')}"
    has_core_skill = bool(find_matches(text, ["sql", "power bi", "powerbi", "excel", "reporting", "dashboard", "dashboards", "kpi", "kpis"]))
    realistic_salary = not salary_max or 28000 <= salary_max <= 42000
    realistic_title = has_junior_friendly_title(job) or has_target_title_keyword(job.get("job_title", ""))
    no_serious_flags = not red_flags or red_flags == "weak technical analysis"

    if calculate_apply_priority_score(job) >= 78 and realistic_salary and no_serious_flags and realistic_title and has_core_skill and is_realistic_location(job):
        return "Apply Today"

    priority_score = calculate_apply_priority_score(job)
    if priority_score >= 68 and has_core_skill and is_realistic_location(job):
        return "Strong Consider"
    if priority_score >= 50:
        return "Consider"
    return "Low Priority"


def get_final_action_reason(job):
    final_action = str(job.get("final_action", ""))
    if final_action == "Skip":
        return get_hard_skip_reason(job) or "hard skip condition"
    if final_action == "Pending AI Review":
        return "manual full JD awaiting AI review"
    if final_action == "Manual Review":
        return "manual AI score requires human review"
    if final_action == "Low Priority":
        return get_salary_seniority_risk(job) or get_junior_fit_reason(job)
    if final_action == "Apply Today":
        return "permanent junior-friendly analyst role with realistic salary, skills and location fit"
    if final_action == "Apply If Time":
        return "manual full JD has moderate AI fit; apply if daily capacity allows"
    return get_junior_fit_reason(job)


def needs_human_review(job):
    if (
        is_manual_job(job)
        and str(job.get("manual_parse_confidence", "")) == "Low"
        and any(str(job.get(field, "")).strip().lower() in ["", "unknown"] for field in ["job_title", "company", "location"])
    ):
        return True
    if has_ai_review_inconsistency(job):
        return True
    if str(job.get("ai_final_action", "")) == "Manual Review":
        return True
    if str(job.get("final_action", "")) not in ["Apply Today", "Strong Consider"]:
        return False
    return bool(
        get_salary_seniority_risk(job)
        or has_senior_experience_wording(job)
        or str(job.get("junior_fit_reason", "")) in ["junior fit unclear", "skills are relevant but junior fit needs review"]
    )


def get_human_review_reason(job):
    if not needs_human_review(job):
        return ""
    if "AI cache invalid" in str(job.get("ai_red_flags", "")):
        return "AI cache invalid"
    if has_ai_review_inconsistency(job):
        return "AI score/action inconsistency"
    if str(job.get("ai_final_action", "")) == "Manual Review":
        return "AI requested manual review"
    return get_salary_seniority_risk(job) or get_junior_fit_reason(job) or "borderline fit"


def get_ai_fit_score(job):
    score = pd.to_numeric(job.get("ai_fit_score", ""), errors="coerce")
    return None if pd.isna(score) else int(max(0, min(100, round(float(score)))))


def get_manual_final_action_from_ai_score(job):
    score = get_ai_fit_score(job)
    if score is None:
        return ""
    if score >= 80:
        return "Apply Today"
    if score >= 70:
        return "Strong Consider"
    if score >= 60:
        return "Apply If Time"
    if score >= 50:
        return "Manual Review"
    return "Skip"


def get_manual_ai_threshold_reason(job):
    score = get_ai_fit_score(job)
    if score is None:
        return ""
    action = get_manual_final_action_from_ai_score(job)
    return f"manual AI score {score}; threshold action {action}"


def has_ai_review_inconsistency(job):
    action = str(job.get("ai_final_action", ""))
    if action not in ["Apply Today", "Strong Consider", "Apply If Time"]:
        return False
    score = get_ai_fit_score(job)
    if score is None:
        return True
    return (
        (action == "Apply Today" and score < 80)
        or (action == "Strong Consider" and score < 70)
        or (action == "Apply If Time" and score < 60)
        or (score < 50 and action in ["Apply Today", "Strong Consider", "Apply If Time"])
    )


def append_ai_red_flag(job, message):
    current = str(job.get("ai_red_flags", "")).strip()
    if message in current:
        return current
    return "; ".join(filter(None, [current, message]))


def has_pre_ai_target_title(job):
    title = str(job.get("job_title", "")).lower()
    target_titles = [
        "data analyst",
        "reporting analyst",
        "mi analyst",
        "insight analyst",
        "commercial analyst",
        "operations analyst",
        "product analyst",
        "crm analyst",
        "people data analyst",
        "hr data analyst",
        "workforce analyst",
    ]
    return any(target in title for target in target_titles)


def has_pre_ai_skill_signal(job):
    text = f"{job.get('job_title', '')}\n{job.get('job_text', '')}\n{job.get('job_description', '')}".lower()
    return bool(
        find_matches(
            text,
            [
                "excel",
                "sql",
                "power bi",
                "powerbi",
                "dashboard",
                "dashboards",
                "reporting",
                "kpi",
                "kpis",
                "analytics",
                "data quality",
                "insight",
                "large datasets",
            ],
        )
    )


def has_pre_ai_salary_fit(job):
    salary_max = get_salary_max(job)
    return salary_max == 0 or 25000 <= salary_max <= 50000


def is_pre_ai_borderline_candidate(job):
    return (
        has_pre_ai_target_title(job)
        and has_pre_ai_salary_fit(job)
        and is_realistic_location(job)
        and has_pre_ai_skill_signal(job)
        and not bool(job.get("previously_seen", False))
    )


def get_pre_ai_bucket(job):
    if get_hard_skip_reason(job):
        return "hard_skip"
    if is_manual_job(job) and not has_valid_ai_decision(job):
        return "ai_review_candidate"
    if str(job.get("final_action", "")) == "Apply Today":
        return "rule_apply_today"
    if str(job.get("final_action", "")) == "Strong Consider":
        return "rule_strong_consider"
    if is_pre_ai_borderline_candidate(job):
        return "ai_review_candidate"
    return "audit_only"


def get_pre_ai_reason(job):
    bucket = str(job.get("pre_ai_bucket", ""))
    if bucket == "hard_skip":
        return get_hard_skip_reason(job)
    if bucket == "rule_apply_today":
        return "rule-based high-confidence apply today candidate"
    if bucket == "rule_strong_consider":
        return "rule-based strong candidate"
    if bucket == "ai_review_candidate":
        if is_manual_job(job):
            return "manual full JD needs AI review"
        return "borderline analyst role with realistic salary, location and skill signals"
    return "audit only; insufficient pre-AI fit signals"


def would_send_to_ai(job):
    if str(job.get("ai_review_source", "rule_based")) in LEGACY_VALID_AI_REVIEW_SOURCES:
        return False
    return str(job.get("pre_ai_bucket", "")) in ["ai_review_candidate", "rule_apply_today", "rule_strong_consider"]


def calculate_ai_review_priority_score(job):
    if not would_send_to_ai(job):
        return 0
    score = 0
    source_priority_value = pd.to_numeric(job.get("source_priority", 3), errors="coerce")
    source_priority = 3 if pd.isna(source_priority_value) else int(source_priority_value)
    score += {1: 45, 2: 35, 3: 25, 4: 5}.get(source_priority, 5)
    location_score_value = pd.to_numeric(job.get("location_fit_score", 0), errors="coerce")
    priority_score_value = pd.to_numeric(job.get("apply_priority_score", 0), errors="coerce")
    location_score = 0 if pd.isna(location_score_value) else int(location_score_value)
    priority_score = 0 if pd.isna(priority_score_value) else int(priority_score_value)
    score += min(location_score, 100) * 0.2
    score += min(priority_score, 100) * 0.25
    if has_junior_friendly_title(job):
        score += 15
    if has_pre_ai_target_title(job):
        score += 15
    if has_pre_ai_skill_signal(job):
        score += 10
    if has_pre_ai_salary_fit(job):
        score += 10
    text = "\n".join(
        [
            str(job.get("red_flags", "")),
            str(job.get("salary_seniority_risk", "")),
            str(job.get("junior_fit_reason", "")),
            str(job.get("job_title", "")),
            str(job.get("job_text", "")),
            str(job.get("job_description", "")),
        ]
    ).lower()
    penalty_terms = [
        "weak technical analysis",
        "high salary suggests mid/senior level",
        "experience wording suggests mid/senior level",
        "acturis",
        "insurance broking",
        "evm",
        "project controls",
        "linear infrastructure",
        "schedule of rates",
        "sor",
        "subcontractor claims",
        "sm&cr",
        "t&c",
        "cpd",
        "cii",
        "cisi",
        "crm analytics",
        "marketing analytics",
    ]
    penalty = sum(12 for term in penalty_terms if term in text)
    if re.search(r"\b[3-9]\+?\s+years?\b", text) and ("crm" in text or "marketing analytics" in text):
        penalty += 25
    return round(max(0, score - penalty), 1)


def add_category_score(score, reasons, category, points, matches):
    """Add points and a clear reason for one scoring category."""
    if points <= 0:
        return score

    score += points
    match_text = ", ".join(matches[:5])

    if len(matches) > 5:
        match_text += f", +{len(matches) - 5} more"

    reasons.append(f"{category}: +{points} ({match_text})")
    return score


def find_best_recommendation_rule(job_title, job_description):
    """Find the best recommendation rule using title first, then description."""
    title = str(job_title).lower()
    description = str(job_description).lower()

    # Use the job title first because it is usually the strongest signal.
    for rule in RECOMMENDATION_RULES:
        matches = find_matches(title, rule["title_keywords"])
        if matches:
            reason = f"title matched {rule['role_type']}: {', '.join(matches)}"
            return rule, reason

    # If the title is unclear, use description keywords as a fallback.
    best_rule = None
    best_matches = []

    for rule in RECOMMENDATION_RULES:
        matches = find_matches(description, rule["description_keywords"])
        if len(matches) > len(best_matches):
            best_rule = rule
            best_matches = matches

    if best_rule and best_matches:
        reason = f"description matched {best_rule['role_type']}: {', '.join(best_matches[:5])}"
        return best_rule, reason

    # Default to a generic analyst cover letter and the general data analyst CV.
    default_rule = {
        "role_type": "generic analyst",
        "cv": "cvs/cv_data_analyst.pdf",
        "cover_letter": "cover_letters/cover_generic_analyst.docx",
    }
    return default_rule, "default: generic analyst recommendation"


def recommend_cv(job_title, job_description):
    """Recommend the best CV version for a job."""
    rule, reason = find_best_recommendation_rule(job_title, job_description)
    return rule["cv"], reason


def recommend_cover_letter(job_title, job_description):
    """Recommend the best cover letter template for a job."""
    rule, reason = find_best_recommendation_rule(job_title, job_description)
    return rule["cover_letter"], reason


def recommend_apply_action(score):
    """Turn the job score into a simple apply recommendation."""
    score = int(score)

    if score >= 75:
        return "Strong Consideration", "score is 75 or higher"
    if score >= 60:
        return "Consider", "score is between 60 and 74"
    if score >= 45:
        return "Low Priority", "score is between 45 and 59"

    return "Skip", "score is below 45"


def get_clean_lines(text):
    """Split pasted job text into useful non-empty lines."""
    lines = str(text).splitlines()
    return [line.strip() for line in lines if line.strip()]


def extract_job_title(job_text, profile):
    """Try to find the job title from the first few lines of pasted text."""
    lines = get_clean_lines(job_text)
    first_lines = lines[:5]

    # First, look for one of the target roles near the top of the pasted text.
    for line in first_lines:
        for role in profile.get("target_roles", []):
            if role.lower() in line.lower():
                return line

    # If no target role is found, use the first line as a simple guess.
    if first_lines:
        return first_lines[0]

    return "Unknown"


def extract_company(job_text):
    """Try to find the company name from a Company line or the first few lines."""
    lines = get_clean_lines(job_text)
    first_lines = lines[:8]

    # LinkedIn/Indeed text sometimes includes a line like "Company: Example Ltd".
    for line in first_lines:
        if line.lower().startswith("company:"):
            company = re.sub(r"(?i)^company\s*:\s*", "", line).strip()
            if company and company.lower() != "company":
                return company

    # If there is no Company line, the company is often the second line near the top.
    if len(lines) >= 2:
        return lines[1]

    return "Unknown"


def extract_location(job_text):
    """Try to find a known UK location inside the pasted job text."""
    text = str(job_text)

    for location in UK_LOCATIONS:
        if location.lower() in text.lower():
            return location

    return "Unknown"


def parse_clear_date(date_text):
    """Parse a clearly stated date into ISO format, otherwise return blank."""
    date_text = str(date_text).strip()
    date_text = re.sub(r"(\d+)(st|nd|rd|th)", r"\1", date_text, flags=re.IGNORECASE)
    date_text = date_text.strip(" .,:;")

    formats = [
        "%Y-%m-%d",
        "%Y-%m-%dT%H:%M:%S",
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d %B %Y",
        "%d %b %Y",
        "%B %d %Y",
        "%b %d %Y",
    ]

    for date_format in formats:
        try:
            parsed = datetime.strptime(date_text[:19], date_format).date()
            return parsed.isoformat()
        except ValueError:
            continue

    return ""


def extract_manual_posted_date(job_text):
    """Extract posted dates only when the text explicitly says posted."""
    text = str(job_text)
    today = date.today()

    days_match = re.search(r"\bposted\s+(\d+)\s+days?\s+ago\b", text, flags=re.IGNORECASE)
    if days_match:
        posted_date = today - timedelta(days=int(days_match.group(1)))
        return posted_date.isoformat(), "explicit"

    if re.search(r"\bposted\s+today\b", text, flags=re.IGNORECASE):
        return today.isoformat(), "explicit"

    if re.search(r"\bposted\s+yesterday\b", text, flags=re.IGNORECASE):
        return (today - timedelta(days=1)).isoformat(), "explicit"

    posted_date_match = re.search(
        r"\bposted(?:\s+on)?\s+([0-9]{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+[0-9]{4}|"
        r"[A-Za-z]+\s+[0-9]{1,2}(?:st|nd|rd|th)?\s+[0-9]{4}|"
        r"[0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{4}|[0-9]{4}-[0-9]{2}-[0-9]{2})",
        text,
        flags=re.IGNORECASE,
    )
    if posted_date_match:
        parsed_date = parse_clear_date(posted_date_match.group(1))
        if parsed_date:
            return parsed_date, "explicit"

    return "Unknown", "none"


def extract_application_deadline(job_text):
    """Extract deadline dates only when the text clearly labels them."""
    text = str(job_text)
    deadline_match = re.search(
        r"\b(?:closing date|closes on|application deadline)\s*:?\s*"
        r"([0-9]{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+[0-9]{4}|"
        r"[A-Za-z]+\s+[0-9]{1,2}(?:st|nd|rd|th)?\s+[0-9]{4}|"
        r"[0-9]{1,2}[/-][0-9]{1,2}[/-][0-9]{4}|[0-9]{4}-[0-9]{2}-[0-9]{2})",
        text,
        flags=re.IGNORECASE,
    )

    if deadline_match:
        parsed_date = parse_clear_date(deadline_match.group(1))
        if parsed_date:
            return parsed_date, "explicit"

    return "", "none"


def classify_freshness(job_posted_date):
    """Turn a known posted date into age, freshness label and score impact."""
    parsed_date = parse_clear_date(job_posted_date)

    if not parsed_date:
        return "", "Unknown", 0

    age_days = (date.today() - datetime.strptime(parsed_date, "%Y-%m-%d").date()).days

    if age_days < 0:
        return "", "Unknown", 0
    if age_days <= 3:
        return age_days, "Fresh", 10
    if age_days <= 7:
        return age_days, "Recent", 5
    if age_days <= 14:
        return age_days, "Aging", 0

    return age_days, "Old", -5


def get_freshness_fields(job):
    """Get freshness fields without guessing missing dates."""
    input_type = str(job.get("input_type", ""))
    job_text = str(job.get("job_text", ""))
    api_posted_date = str(job.get("job_posted_date", "")).strip()
    api_deadline = str(job.get("application_deadline", "")).strip()

    if input_type == "api" and api_posted_date:
        parsed_posted_date = parse_clear_date(api_posted_date)
        job_posted_date = parsed_posted_date or "Unknown"
        posted_confidence = "api" if parsed_posted_date else "none"
    elif input_type == "manual":
        job_posted_date, posted_confidence = extract_manual_posted_date(job_text)
    else:
        job_posted_date = "Unknown"
        posted_confidence = "none"

    if api_deadline:
        parsed_deadline = parse_clear_date(api_deadline)
        application_deadline = parsed_deadline
        deadline_confidence = "api" if parsed_deadline else "none"
    else:
        application_deadline, deadline_confidence = extract_application_deadline(job_text)

    job_age_days, freshness, freshness_score = classify_freshness(job_posted_date)

    return {
        "job_posted_date": job_posted_date,
        "job_age_days": job_age_days,
        "freshness": freshness,
        "posted_date_confidence": posted_confidence,
        "application_deadline": application_deadline,
        "deadline_confidence": deadline_confidence,
        "freshness_score": freshness_score,
    }


def score_job(job, profile):
    """Give one job a score from 0 to 100 and explain the score."""
    score = 0
    reasons = []

    # Keep the original pasted text for scoring.
    job_text = str(job.get("job_text", ""))
    job_title = str(job.get("job_title", ""))
    location = str(job.get("location", ""))
    all_text = f"{job_title}\n{location}\n{job_text}"
    title_lower = job_title.lower()
    text_lower = all_text.lower()

    # 1. Role/title match: up to 30 points.
    role_matches = find_matches(job_title, GOOD_ROLE_KEYWORDS)
    assistant_matches = find_matches(job_title, ASSISTANT_ROLE_KEYWORDS)
    assistant_weak_matches = find_matches(job_title, ASSISTANT_WEAK_FIT_KEYWORDS)
    assistant_strong_matches = find_matches(
        all_text,
        ["analytics", "dashboards", "dashboard", "sql", "power bi", "powerbi", "bi tools"],
    )
    assistant_data_matches = find_matches(
        all_text,
        ["data", "analysis", "analytics", "excel", "reporting"],
    )
    commercial_signal_matches = find_matches(all_text, COMMERCIAL_ANALYST_SIGNALS)

    if "commercial analyst" in title_lower and commercial_signal_matches:
        role_points = 24
        role_matches = ["commercial analyst"] + commercial_signal_matches[:2]
    elif "commercial analyst" in title_lower:
        role_points = 12
        role_matches = ["commercial analyst, weak analysis signals"]
    elif role_matches:
        role_points = 30
    elif assistant_matches and assistant_strong_matches:
        role_points = 12
        role_matches = assistant_matches + assistant_strong_matches[:2]
    elif assistant_matches and assistant_data_matches:
        role_points = 8
        role_matches = assistant_matches + assistant_data_matches[:2]
    elif assistant_matches:
        role_points = 3
        role_matches = assistant_matches
    else:
        role_points = 0

    score = add_category_score(score, reasons, "role/title", role_points, role_matches)

    # 2. Technical skills: up to 30 points.
    technical_matches = find_matches(all_text, TECHNICAL_KEYWORDS)
    strong_technical_matches = find_matches(
        all_text,
        [
            "sql",
            "kql",
            "power bi",
            "powerbi",
            "dashboard",
            "dashboards",
            "large datasets",
            "metrics",
            "telemetry",
            "data visualisation",
            "data visualization",
            "reporting tools",
            "tableau",
            "looker",
            "analytics",
            "stakeholder insights",
        ],
    )
    technical_points = min(30, len(technical_matches) * 3 + len(strong_technical_matches) * 2)
    score = add_category_score(score, reasons, "technical skills", technical_points, technical_matches)

    # Extra boost for true Data Analyst roles with SQL plus BI/reporting signals.
    data_analyst_boost_matches = []
    if "data analyst" in title_lower and "sql" in text_lower:
        bi_reporting_matches = find_matches(
            all_text,
            ["power bi", "powerbi", "dashboard", "dashboards", "reporting", "reporting tools"],
        )
        if bi_reporting_matches:
            data_analyst_boost_matches = ["data analyst", "sql"] + bi_reporting_matches[:2]
            score += 6
            reasons.append(
                f"boost: +6 data analyst with SQL and BI/reporting ({', '.join(data_analyst_boost_matches)})"
            )

    # 3. Business/commercial relevance: up to 20 points.
    business_matches = find_matches(all_text, BUSINESS_KEYWORDS)
    business_points = min(15, len(business_matches) * 2)
    if "commercial analyst" in title_lower and commercial_signal_matches:
        business_points = min(20, business_points + 5)
    score = add_category_score(score, reasons, "business relevance", business_points, business_matches)

    # 4. Junior friendliness: up to 10 points.
    junior_matches = find_matches(all_text, JUNIOR_KEYWORDS)
    junior_points = min(10, len(junior_matches) * 3)
    score = add_category_score(score, reasons, "junior friendly", junior_points, junior_matches)

    # 5. Location/work pattern: up to 10 points.
    location_matches = find_matches(all_text, UK_LOCATIONS)
    location_points = min(10, len(location_matches) * 5)
    score = add_category_score(score, reasons, "location/work pattern", location_points, location_matches)

    # 6. Penalties for roles that look too senior or clearly off-target.
    senior_matches = find_matches(job_title, SENIOR_TITLE_KEYWORDS)
    if senior_matches:
        score -= 20
        reasons.append(f"penalty: -20 senior title ({', '.join(senior_matches)})")

    bad_fit_matches = find_matches(title_lower, BAD_FIT_KEYWORDS)
    if bad_fit_matches:
        score -= 15
        reasons.append(f"penalty: -15 off-target role ({', '.join(bad_fit_matches)})")

    tender_admin_matches = find_matches(all_text, TENDER_ADMIN_KEYWORDS)
    if tender_admin_matches:
        tender_penalty = min(12, len(tender_admin_matches) * 3)
        score -= tender_penalty
        reasons.append(
            f"penalty: -{tender_penalty} tender/admin focus ({', '.join(tender_admin_matches[:5])})"
        )

    if assistant_matches and assistant_weak_matches and not assistant_strong_matches:
        score -= 10
        reasons.append(
            f"penalty: -10 assistant/merchandising role without strong BI tools ({', '.join(assistant_weak_matches)})"
        )

    # Make sure the score stays between 0 and 100.
    final_score = max(0, min(score, 100))

    if not reasons:
        reasons.append("few profile matches found")

    return final_score, "; ".join(reasons)


def detect_red_flags(job):
    """Find practical reasons to slow down before applying."""
    job_title = str(job.get("job_title", ""))
    location = str(job.get("location", ""))
    job_text = str(job.get("job_text", ""))
    all_text = f"{job_title}\n{location}\n{job_text}"
    text_lower = all_text.lower()
    flags = []

    flag_rules = [
        ("fixed-term contract", ["fixed-term contract", "fixed term contract", "fixed-term"]),
        ("12-month contract", ["12-month contract", "12 month contract", "12 months contract"]),
        ("experienced analyst requirement", ["experienced analyst", "proven experience", "3+ years", "three years"]),
        ("3 days onsite", ["3 days onsite", "3 days on-site", "three days onsite", "three days on-site"]),
        ("training programme / not a normal job", TRAINING_PROGRAMME_KEYWORDS),
        ("tender / RFQ / proposal heavy", TENDER_ADMIN_KEYWORDS),
        ("admin / product upload heavy", BAD_FIT_KEYWORDS + ASSISTANT_WEAK_FIT_KEYWORDS),
    ]

    for label, keywords in flag_rules:
        if find_matches(text_lower, keywords):
            flags.append(label)

    if is_contract_or_temporary_role(job):
        flags.append("contract/temporary role - user prefers permanent roles")

    if has_high_salary_seniority_risk(job):
        flags.append("high salary suggests mid/senior level")

    if has_senior_experience_wording(job):
        flags.append("experience wording suggests mid/senior level")

    # Senior words are only a red flag when they are in the title. Job
    # descriptions often mention "manager" as a stakeholder, which is normal.
    if find_matches(job_title, SENIOR_TITLE_KEYWORDS):
        flags.append("senior / lead / manager / head of")

    near_locations = ["milton keynes", "london", "birmingham", "luton", "bedford", "northampton", "remote", "hybrid"]
    location_lower = location.lower()
    is_near_location = any(near_location in location_lower for near_location in near_locations)
    if location and not is_near_location:
        flags.append("far location")

    technical_matches = find_matches(all_text, TECHNICAL_KEYWORDS)
    if len(technical_matches) < 2:
        flags.append("weak technical analysis")

    return "; ".join(dict.fromkeys(flags))


def has_target_title_keyword(job_title):
    """Keep likely analyst jobs visible even when the score is low."""
    title_lower = str(job_title).lower()

    for keyword in TARGET_TITLE_KEYWORDS:
        if keyword in ["bi", "mi"]:
            if re.search(rf"\b{keyword}\b", title_lower):
                return True
        elif keyword in title_lower:
            return True

    return False


def has_strong_data_keywords(text):
    """Check whether a job has enough real data signals to review."""
    return len(find_matches(text, STRONG_DATA_KEYWORDS)) >= 2


def has_irrelevant_title_without_data(job):
    """Find clearly irrelevant API titles unless the text has strong data keywords."""
    job_title = str(job.get("job_title", ""))
    job_text = str(job.get("job_text", ""))
    title_matches = find_matches(job_title, IRRELEVANT_TITLE_KEYWORDS)

    if not title_matches:
        return False

    return not has_strong_data_keywords(f"{job_title}\n{job_text}")


def has_training_programme_flag(job):
    """Identify training providers and career-change programmes."""
    job_title = str(job.get("job_title", ""))
    job_text = str(job.get("job_text", ""))
    return bool(find_matches(f"{job_title}\n{job_text}", TRAINING_PROGRAMME_KEYWORDS))


def has_senior_title_flag(job):
    """Only treat senior words as senior flags when they appear in the title."""
    return bool(find_matches(str(job.get("job_title", "")), SENIOR_TITLE_KEYWORDS))


def has_irrelevant_assistant_role(job):
    """Skip assistant roles unless they have strong data/BI signals."""
    job_title = str(job.get("job_title", ""))
    job_text = str(job.get("job_text", ""))

    if not find_matches(job_title, ASSISTANT_ROLE_KEYWORDS):
        return False

    return not has_strong_data_keywords(f"{job_title}\n{job_text}")


def calculate_experience_fit(job):
    """Estimate how realistic the experience requirement is for this profile."""
    job_title = str(job.get("job_title", ""))
    job_text = str(job.get("job_text", ""))
    all_text = f"{job_title}\n{job_text}".lower()
    score = 55

    if find_matches(all_text, JUNIOR_KEYWORDS):
        score += 20
    if re.search(r"\b(0-?1|1)\+?\s+years?\b", all_text):
        score += 15
    if re.search(r"\b2\+?\s+years?\b", all_text):
        score += 5
    if re.search(r"\b(3|4)\+?\s+years?\b", all_text):
        score -= 20
    if re.search(r"\b(5|6|7|8|9|10)\+?\s+years?\b", all_text):
        score -= 35
    if find_matches(job_title, SENIOR_TITLE_KEYWORDS):
        score -= 35
    if find_matches(all_text, ["no experience needed", "full training provided"]):
        score += 10
    if find_matches(all_text, ["proven experience", "experienced analyst"]):
        score -= 15
    if has_strong_data_keywords(all_text):
        score += 10

    return max(0, min(score, 100))


def calculate_industry_fit(job):
    """Estimate how well the industry/domain matches the profile background."""
    job_title = str(job.get("job_title", ""))
    company = str(job.get("company", ""))
    job_text = str(job.get("job_text", ""))
    all_text = f"{job_title}\n{company}\n{job_text}"
    matches = find_matches(all_text, BUSINESS_KEYWORDS)

    score = 40 + min(50, len(matches) * 8)

    if find_matches(all_text, ["retail", "ecommerce", "e-commerce", "commercial", "customer", "sales", "revenue"]):
        score += 10
    if find_matches(all_text, ["logistics", "supply chain", "warehouse", "transport", "operations"]):
        score += 5

    return max(0, min(score, 100))


def estimate_competition_score(job):
    """Estimate likely competition pressure from broad job-market signals."""
    job_title = str(job.get("job_title", ""))
    location = str(job.get("location", ""))
    source = str(job.get("source", ""))
    job_text = str(job.get("job_text", ""))
    all_text = f"{job_title}\n{location}\n{source}\n{job_text}".lower()
    score = 50

    if "london" in all_text:
        score += 15
    if "remote" in all_text:
        score += 20
    if find_matches(all_text, ["junior", "graduate", "trainee", "entry level"]):
        score += 15
    if str(source).lower() in ["reed", "indeed"]:
        score += 10
    if find_matches(job_title, ["bi analyst", "mi analyst", "reporting analyst"]):
        score -= 10
    if find_matches(all_text, ["contract", "fixed-term", "12 month", "12-month"]):
        score -= 10

    return max(0, min(score, 100))


def get_competition_level(job):
    """Convert competition score into a simple level for the spreadsheet."""
    competition_score = estimate_competition_score(job)

    if competition_score >= 70:
        return "High"
    if competition_score >= 45:
        return "Medium"
    return "Low"


def get_application_priority(job):
    """Create the V2 application priority tier."""
    score = int(job.get("score", 0))
    experience_fit = int(job.get("experience_fit", 0))
    industry_fit = int(job.get("industry_fit", 0))
    next_action = str(job.get("next_action", ""))
    red_flags = str(job.get("red_flags", ""))
    competition_level = str(job.get("competition_level", ""))

    if (
        next_action == "Skip"
        or "training programme / not a normal job" in red_flags
        or "senior / lead / manager / head of" in red_flags
    ):
        return "Tier 4 = Skip"
    if score >= 75 and experience_fit >= 60 and industry_fit >= 50 and competition_level != "High":
        return "Tier 1 = Apply Immediately"
    if score >= 60 and experience_fit >= 50:
        return "Tier 2 = Apply Regularly"
    if score >= 45:
        return "Tier 3 = Apply Only If High Score"

    return "Tier 4 = Skip"


def estimate_interview_probability(job):
    """Estimate interview probability as a rough percentage."""
    score = int(job.get("score", 0))
    experience_fit = int(job.get("experience_fit", 0))
    industry_fit = int(job.get("industry_fit", 0))
    competition_level = str(job.get("competition_level", ""))
    red_flags = str(job.get("red_flags", ""))

    probability = score * 0.35 + experience_fit * 0.30 + industry_fit * 0.20

    if competition_level == "Low":
        probability += 10
    elif competition_level == "High":
        probability -= 10

    if red_flags:
        probability -= 10
    if "training programme / not a normal job" in red_flags:
        probability = min(probability, 5)
    if "senior / lead / manager / head of" in red_flags:
        probability = min(probability, 10)

    return int(max(0, min(round(probability), 90)))


def should_show_in_main_output(job):
    """Decide whether a job belongs in the shorter ranked_jobs.xlsx file."""
    score = int(job.get("score", 0))
    next_action = str(job.get("next_action", ""))
    red_flags = str(job.get("red_flags", ""))

    if "training programme / not a normal job" in red_flags:
        return False
    if "contract/temporary role - user prefers permanent roles" in red_flags:
        return False
    if "salary level suggests senior/mid-level role" in str(job.get("reason", "")):
        return False
    if is_api_job(job) and not is_extremely_strong_api_job(job):
        return False

    return score >= 35 or next_action != "Skip"


def get_api_lead_next_action(job):
    if bool(job.get("previously_seen", False)):
        return "Already applied"
    if bool(job.get("previously_rejected_similar", False)) and not bool(job.get("tracker_exact_match", False)):
        return "Open link and check if this is a repost/new vacancy"
    if str(job.get("hard_skip_reason", "")).strip() or str(job.get("final_action", "")) == "Skip":
        return "Skip"
    if is_pre_ai_borderline_candidate(job) or is_extremely_strong_api_job(job):
        return "Open link and copy full JD"
    if pd.to_numeric(job.get("score", 0), errors="coerce") >= 35:
        return "Low priority"
    return "Skip"


def get_api_leads(jobs):
    if jobs.empty:
        return pd.DataFrame()
    api_leads = jobs[jobs.apply(is_api_job, axis=1)].copy()
    if api_leads.empty:
        return api_leads
    api_leads["suggested_next_action"] = api_leads.apply(get_api_lead_next_action, axis=1)
    api_leads = api_leads[
        api_leads["suggested_next_action"].isin(
            [
                "Open link and copy full JD",
                "Open link and check if this is a repost/new vacancy",
                "Low priority",
                "Already applied",
            ]
        )
    ]
    api_leads = api_leads.sort_values(
        by=["suggested_next_action", "apply_priority_score", "score", "location_fit_score"],
        ascending=[True, False, False, False],
    )
    columns = [
        "job_title",
        "company",
        "location",
        "salary",
        "source",
        "apply_link",
        "score",
        "apply_priority_score",
        "reason",
        "red_flags",
        "next_action",
        "suggested_next_action",
        "previously_rejected_similar",
        "previous_rejection_date",
        "previous_rejected_company",
        "previous_rejected_title",
        "previous_rejected_job_id",
        "repost_candidate",
        "tracker_exact_match",
        "tracker_similarity_match",
        "previous_similar_status",
        "previous_application_status",
        "tracker_overlay_reason",
    ]
    for column in columns:
        if column not in api_leads.columns:
            api_leads[column] = ""
    return api_leads[columns]


def is_actionable_status(job):
    """Only show jobs that still need review or application."""
    status = standardise_status(job.get("status", "new"))
    excluded_statuses = {
        "applied",
        "assessment",
        "interview",
        "final_interview",
        "offer",
        "rejected",
        "withdrawn",
        "not_interested",
        "skip",
        "expired",
        "closed",
    }
    return status not in excluded_statuses


def get_next_action_rank(next_action):
    """Sort the daily review file by the most useful action first."""
    order = {
        "Apply now": 1,
        "Strong consider": 2,
        "Review manually": 3,
        "Practice application only": 4,
        "Skip": 5,
    }
    return order.get(str(next_action), 99)


def normalize_dedupe_text(text):
    """Normalize company/title text for duplicate detection."""
    text = str(text).lower()
    text = re.sub(r"\b(ltd|limited|plc|recruitment|group uk)\b", " ", text)
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_for_match(value):
    """
    Normalise general text for tracker matching.
    Used for company, title, location and other non-URL text.
    """
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value)
    text = html.unescape(text)
    text = text.lower()
    text = re.sub(r"[\u200b-\u200f\ufeff]", "", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def load_company_aliases():
    aliases = {
        "7im": "seven investment management",
        "seven investment management": "seven investment management",
    }
    if COMPANY_ALIASES_FILE.exists():
        alias_rows = pd.read_csv(COMPANY_ALIASES_FILE).fillna("")
        if {"alias", "canonical_company"}.issubset(alias_rows.columns):
            for _, row in alias_rows.iterrows():
                alias = normalize_company_for_match(row["alias"], aliases={})
                canonical = normalize_company_for_match(row["canonical_company"], aliases={})
                if alias and canonical:
                    aliases[alias] = canonical
                    aliases[canonical] = canonical
    return aliases


def normalize_company_for_match(company, aliases=None):
    normalized = str(company).lower()
    normalized = re.sub(r"[^a-z0-9\s]", " ", normalized)
    normalized = re.sub(r"\b(ltd|limited|plc|group|uk|recruitment)\b", " ", normalized)
    normalized = re.sub(r"\s+", " ", normalized).strip()
    compact = normalized.replace(" ", "")
    aliases = aliases or {}
    if normalized in aliases:
        return aliases[normalized]
    if compact in aliases:
        return aliases[compact]
    return normalized


def normalize_title_for_match(title):
    normalized = str(title).lower()
    normalized = re.sub(r"[^a-z0-9\s]", " ", normalized)
    suffix_terms = [
        "contract",
        "ftc",
        "fixed term",
        "12 month",
        "12 months",
        "6 month",
        "3 month",
        "london",
        "hybrid",
        "remote",
    ]
    for term in suffix_terms:
        normalized = re.sub(rf"\b{re.escape(term)}\b", " ", normalized)
    return re.sub(r"\s+", " ", normalized).strip()


def token_similarity(left, right):
    left_tokens = set(str(left).split())
    right_tokens = set(str(right).split())
    if not left_tokens or not right_tokens:
        return 0
    return len(left_tokens & right_tokens) / max(len(left_tokens), len(right_tokens))


def strong_title_match(left, right):
    left_normalized = normalize_title_for_match(left)
    right_normalized = normalize_title_for_match(right)
    return bool(left_normalized and right_normalized and token_similarity(left_normalized, right_normalized) >= 0.8)


def make_dedupe_key(job):
    """Create a stable duplicate key from normalized company and title."""
    company = normalize_company_for_match(job.get("company", ""), load_company_aliases())
    job_title = normalize_title_for_match(job.get("job_title", ""))
    location = normalize_dedupe_text(job.get("location", ""))
    description_tokens = normalize_dedupe_text(job.get("job_text", "")).split()
    description_signature = " ".join(description_tokens[:20])
    if not company and not job_title:
        fallback = normalize_dedupe_text(job.get("apply_link", "") or job.get("file_name", ""))
        return f"unknown|{fallback}"
    return f"{company}|{job_title}|{location}|{description_signature}"


def count_populated_fields(job):
    """Count useful populated fields so richer API rows can win ties."""
    fields = [
        "source",
        "job_title",
        "company",
        "location",
        "salary",
        "job_text",
        "apply_link",
    ]
    return sum(bool(str(job.get(field, "")).strip()) for field in fields)


def add_duplicate_metadata(jobs):
    """Add duplicate metadata while keeping all rows."""
    jobs = jobs.copy()
    jobs["dedupe_key"] = jobs.apply(make_dedupe_key, axis=1)
    jobs["populated_field_count"] = jobs.apply(count_populated_fields, axis=1)

    duplicate_counts = jobs.groupby("dedupe_key")["dedupe_key"].transform("count")
    jobs["duplicate_count"] = duplicate_counts

    merged_sources = jobs.groupby("dedupe_key")["source"].transform(
        lambda sources: "|".join(sorted(set(str(source) for source in sources if str(source).strip())))
    )
    jobs["merged_sources"] = merged_sources

    return jobs


def deduplicate_ranked_jobs(jobs):
    """Keep the best row for each dedupe_key for the daily review file."""
    dedupe_sorted = jobs.sort_values(
        by=["dedupe_key", "final_action_sort", "apply_priority_score", "score", "populated_field_count", "source"],
        ascending=[True, True, False, False, False, True],
    )
    deduped_jobs = dedupe_sorted.drop_duplicates(subset=["dedupe_key"], keep="first")
    return deduped_jobs.sort_values(
        by=["final_action_sort", "apply_priority_score", "date_added_sort", "location_fit_score", "score", "source"],
        ascending=[True, False, False, False, False, True],
    )


def apply_ai_decisions_to_final_action(jobs):
    """Use cached/API AI decisions only after rule-based hard skips are known."""
    jobs = jobs.copy()
    if "rule_final_action" not in jobs.columns:
        jobs["rule_final_action"] = jobs["final_action"]
    if "rule_final_action_reason" not in jobs.columns:
        jobs["rule_final_action_reason"] = jobs["final_action_reason"]
    valid_actions = ["Apply Today", "Strong Consider", "Apply If Time", "Consider", "Low Priority", "Skip", "Manual Review"]
    ai_source_mask = jobs.apply(has_valid_ai_result, axis=1)
    ai_action_mask = jobs["ai_final_action"].isin(valid_actions)
    ai_decision_mask = ai_source_mask & ai_action_mask

    jobs.loc[ai_decision_mask, "final_action"] = jobs.loc[ai_decision_mask, "ai_final_action"]

    ai_reason_mask = ai_decision_mask & jobs["ai_main_reason"].astype(str).str.strip().ne("")
    jobs.loc[ai_reason_mask, "final_action_reason"] = jobs.loc[ai_reason_mask, "ai_main_reason"]

    manual_ai_decision_mask = ai_decision_mask & jobs.apply(is_manual_job, axis=1)
    if manual_ai_decision_mask.any():
        manual_threshold_actions = jobs.loc[manual_ai_decision_mask].apply(get_manual_final_action_from_ai_score, axis=1)
        manual_threshold_reasons = jobs.loc[manual_ai_decision_mask].apply(get_manual_ai_threshold_reason, axis=1)
        jobs.loc[manual_ai_decision_mask, "final_action"] = manual_threshold_actions
        reason_prefix = jobs.loc[manual_ai_decision_mask, "ai_main_reason"].astype(str).str.strip()
        jobs.loc[manual_ai_decision_mask, "final_action_reason"] = [
            "; ".join(filter(None, [prefix, threshold_reason]))
            for prefix, threshold_reason in zip(reason_prefix, manual_threshold_reasons)
        ]

    manual_review_mask = ai_decision_mask & jobs["ai_final_action"].eq("Manual Review") & ~jobs.apply(is_manual_job, axis=1)
    if manual_review_mask.any():
        priority_scores = pd.to_numeric(jobs.loc[manual_review_mask, "apply_priority_score"], errors="coerce").fillna(0)
        fallback_mask = manual_review_mask & jobs["ai_red_flags"].astype(str).str.contains(
            "AI cache invalid; using rule-based fallback",
            regex=False,
            na=False,
        )
        non_fallback_mask = manual_review_mask & ~fallback_mask
        fallback_rule_actions = jobs.loc[fallback_mask, "rule_final_action"].where(
            jobs.loc[fallback_mask, "rule_final_action"].isin(["Apply Today", "Strong Consider"]),
            "Strong Consider",
        )
        jobs.loc[fallback_mask, "final_action"] = fallback_rule_actions
        jobs.loc[fallback_mask, "final_action_reason"] = "AI cache invalid; using rule-based fallback"
        jobs.loc[fallback_mask, "needs_human_review"] = True
        jobs.loc[fallback_mask, "human_review_reason"] = "AI cache invalid; using rule-based fallback"
        non_fallback_indexes = jobs.index[non_fallback_mask]
        jobs.loc[non_fallback_mask, "final_action"] = [
            "Consider" if score >= 50 else "Low Priority" for score in priority_scores.reindex(non_fallback_indexes).fillna(0)
        ]
        jobs.loc[non_fallback_mask, "final_action_reason"] = "AI requested manual review"

    inconsistency_mask = jobs.apply(has_ai_review_inconsistency, axis=1) & ~jobs.apply(is_manual_job, axis=1)
    if inconsistency_mask.any():
        priority_scores = pd.to_numeric(jobs.loc[inconsistency_mask, "apply_priority_score"], errors="coerce").fillna(0)
        jobs.loc[inconsistency_mask, "ai_final_action"] = "Manual Review"
        jobs.loc[inconsistency_mask, "final_action"] = ["Consider" if score >= 50 else "Low Priority" for score in priority_scores]
        jobs.loc[inconsistency_mask, "final_action_reason"] = "AI score/action inconsistency"
        jobs.loc[inconsistency_mask, "needs_human_review"] = True
        jobs.loc[inconsistency_mask, "human_review_reason"] = "AI score/action inconsistency"
        jobs.loc[inconsistency_mask, "ai_red_flags"] = jobs.loc[inconsistency_mask].apply(
            lambda job: append_ai_red_flag(job, "AI score/action inconsistency"),
            axis=1,
        )

    hard_skip_mask = jobs["hard_skip_reason"].astype(str).str.strip().ne("")
    jobs.loc[hard_skip_mask, "final_action"] = "Skip"
    jobs.loc[hard_skip_mask, "final_action_reason"] = jobs.loc[hard_skip_mask, "hard_skip_reason"]
    return jobs


def add_manual_duplicate_metadata(jobs):
    jobs = jobs.copy()
    manual_mask = jobs["input_type"].astype(str).str.lower().eq("manual")
    jobs["manual_duplicate"] = False
    if not manual_mask.any():
        return jobs
    manual_keys = jobs.loc[manual_mask, "dedupe_key"].astype(str)
    jobs.loc[manual_mask, "manual_duplicate"] = manual_keys.duplicated(keep="first")
    link_keys = jobs.loc[manual_mask, "apply_link"].astype(str).str.strip().str.lower()
    duplicate_link_indexes = link_keys[link_keys.ne("") & link_keys.duplicated(keep="first")].index
    jobs.loc[duplicate_link_indexes, "manual_duplicate"] = True
    return jobs


def should_show_cache_invalid_rule_fallback(job):
    if str(job.get("ai_review_source", "")) != "cache_invalid":
        return True
    return (
        str(job.get("pre_ai_bucket", "")) == "rule_apply_today"
        and str(job.get("rule_final_action", "")) in ["Apply Today", "Strong Consider"]
        and not str(job.get("hard_skip_reason", "")).strip()
        and not str(job.get("salary_seniority_risk", "")).strip()
        and "weak technical analysis" not in str(job.get("red_flags", "")).lower()
        and "tender" not in str(job.get("red_flags", "")).lower()
        and "rfq" not in str(job.get("red_flags", "")).lower()
        and "proposal" not in str(job.get("red_flags", "")).lower()
    )


def get_ranked_jobs_exclusion_reason(job, included_indexes=None, deduped_indexes=None, allowed_actions=None):
    included_indexes = included_indexes or set()
    deduped_indexes = deduped_indexes or set()
    allowed_actions = allowed_actions or ["Apply Today", "Strong Consider", "Apply If Time"]
    reasons = []
    if job.name in included_indexes:
        return ""
    if not is_manual_job(job):
        reasons.append("not a manual raw job")
    if str(job.get("final_action", "")) not in allowed_actions:
        reasons.append("final_action is not a ranked_jobs action")
    if not has_valid_ai_decision(job):
        reasons.append("missing valid AI decision")
    if not is_actionable_status(job):
        reasons.append(f"tracker status is {standardise_status(job.get('status', ''))}")
    if bool(job.get("previously_seen", False)):
        reasons.append("exact tracker match already seen")
    if str(job.get("hard_skip_reason", "")).strip():
        reasons.append(str(job.get("hard_skip_reason", "")).strip())
    if str(job.get("human_review_reason", "")) == "AI score/action inconsistency":
        reasons.append("AI score/action inconsistency")
    if str(job.get("human_review_reason", "")) == "AI cache invalid" and not should_show_cache_invalid_rule_fallback(job):
        reasons.append("AI cache invalid")
    if not should_show_cache_invalid_rule_fallback(job):
        reasons.append("cache-invalid rule fallback not eligible")
    if job.name not in deduped_indexes:
        reasons.append("deduplicated in favour of another row")
    if not reasons:
        reasons.append("ranked_jobs filter mismatch")
    return "; ".join(dict.fromkeys(reasons))


def build_ranked_jobs_exclusion_audit(ranked_jobs, shown_ranked_jobs, deduped_ranked_jobs, allowed_daily_actions):
    columns = [
        "file_name",
        "job_title",
        "company",
        "final_action",
        "ai_fit_score",
        "should_be_in_ranked_jobs",
        "included_in_ranked_jobs",
        "exclusion_reason",
        "tracker_overlay_reason",
        "duplicate_reason",
        "queue_reason",
    ]
    if ranked_jobs.empty:
        return pd.DataFrame(columns=columns)
    included_indexes = set(shown_ranked_jobs.index)
    deduped_indexes = set(deduped_ranked_jobs.index)
    candidates = ranked_jobs[
        ranked_jobs.apply(is_manual_job, axis=1)
        & ranked_jobs["final_action"].isin(allowed_daily_actions)
    ].copy()
    audit_rows = []
    for _, job in candidates.iterrows():
        included = job.name in included_indexes
        if included:
            continue
        duplicate_reason = ""
        if pd.to_numeric(job.get("duplicate_count", 1), errors="coerce") > 1:
            duplicate_reason = f"duplicate_count={job.get('duplicate_count', '')}; merged_sources={job.get('merged_sources', '')}"
        if job.name not in deduped_indexes:
            duplicate_reason = append_text_flag(duplicate_reason, "deduplicated in favour of another row")
        queue_reason = ""
        if str(job.get("final_action", "")) in ["Manual Review", "Pending AI Review"]:
            queue_reason = str(job.get("final_action_reason", ""))
        elif str(job.get("human_review_reason", "")).strip():
            queue_reason = str(job.get("human_review_reason", "")).strip()
        audit_rows.append(
            {
                "file_name": job.get("file_name", ""),
                "job_title": job.get("job_title", ""),
                "company": job.get("company", ""),
                "final_action": job.get("final_action", ""),
                "ai_fit_score": job.get("ai_fit_score", ""),
                "should_be_in_ranked_jobs": True,
                "included_in_ranked_jobs": included,
                "exclusion_reason": get_ranked_jobs_exclusion_reason(
                    job,
                    included_indexes=included_indexes,
                    deduped_indexes=deduped_indexes,
                    allowed_actions=allowed_daily_actions,
                ),
                "tracker_overlay_reason": job.get("tracker_overlay_reason", ""),
                "duplicate_reason": duplicate_reason,
                "queue_reason": queue_reason,
            }
        )
    return pd.DataFrame(audit_rows, columns=columns)


def classify_tracker_match_type(reason, status_source=""):
    reason_text = str(reason or "").lower()
    if "apply_link" in reason_text:
        return "exact_apply_link"
    if "external job id" in reason_text:
        return "exact_external_job_id"
    if "canonical job_id" in reason_text:
        return "exact_canonical_job_id"
    if "description hash" in reason_text:
        return "exact_description_hash"
    if str(status_source or "") == "tracker_similarity":
        return "similar_tracker_match"
    if reason_text:
        return "exact_other"
    return ""


def has_strong_exact_tracker_evidence(job):
    return classify_tracker_match_type(job.get("tracker_overlay_reason", ""), job.get("status_source", "")) in [
        "exact_apply_link",
        "exact_external_job_id",
        "exact_canonical_job_id",
        "exact_description_hash",
    ]


def find_matched_tracker_row(job, tracker, aliases):
    if tracker.empty:
        return {}
    matched_job_id = str(job.get("matched_tracker_job_id", "")).strip()
    if matched_job_id and "job_id" in tracker.columns:
        matches = tracker[tracker["job_id"].astype(str).str.strip().eq(matched_job_id)]
        if not matches.empty:
            return matches.iloc[0].to_dict()
    for _, row in tracker.iterrows():
        if row_exact_match_reason(job, row, aliases):
            return row.to_dict()
    similar = tracker[tracker_similarity_mask(job, tracker, aliases)]
    if not similar.empty:
        return similar.iloc[0].to_dict()
    return {}


def get_external_job_id_text(job):
    return "; ".join(sorted(extract_url_job_identifiers(job.get("apply_link", ""))))


def get_ai_review_queue_reason(job, manual_needs_ai_review=False):
    if manual_needs_ai_review:
        return "pending AI review"
    if str(job.get("final_action", "")) == "Pending AI Review":
        return "final_action is Pending AI Review"
    if str(job.get("final_action", "")) == "Manual Review":
        return "final_action is Manual Review"
    if str(job.get("ai_final_action", "")) == "Manual Review":
        return "AI returned Manual Review"
    if str(job.get("ai_review_source", "")) == "cache_invalid":
        return "AI cache invalid"
    if str(job.get("human_review_reason", "")) == "AI score/action inconsistency":
        return "AI score/action inconsistency"
    if str(job.get("manual_parse_confidence", "")) == "Low":
        return "suspicious parser case"
    if bool(job.get("previously_rejected_similar", False)):
        return "similar rejected tracker history"
    if bool(job.get("tracker_similarity_match", False)):
        return "similar applied/interview tracker history"
    if bool(job.get("repost_candidate", False)):
        return "possible repost candidate"
    if is_api_job(job):
        return "API lead candidate"
    return ""


def should_be_in_ai_review_queue(job, manual_needs_ai_review=False):
    if str(job.get("hard_skip_reason", "")).strip():
        return False
    if is_api_job(job):
        if not bool(cfg("AI_BATCH_INCLUDE_API_LEADS", False)):
            return False
        return get_api_lead_next_action(job) in ["Open link and copy full JD", "Low priority"]
    if not is_manual_job(job):
        return False
    return bool(get_ai_review_queue_reason(job, manual_needs_ai_review))


def build_repost_tracker_audit(manual_rows):
    columns = [
        "file_name",
        "parsed_company",
        "parsed_job_title",
        "parsed_location",
        "apply_link",
        "external_job_id",
        "posted_date",
        "description_hash",
        "matched_tracker_company",
        "matched_tracker_job_title",
        "matched_tracker_status",
        "matched_tracker_apply_link",
        "matched_tracker_date",
        "tracker_match_type",
        "is_exact_tracker_match",
        "is_similar_tracker_match",
        "is_repost_candidate",
        "excluded_from_ai",
        "exclusion_reason",
        "recommended_action",
    ]
    tracker = load_applications_tracker()
    aliases = load_company_aliases()
    rows = []
    for _, job in manual_rows.sort_values(by=["file_name"]).iterrows():
        if not (bool(job.get("tracker_exact_match", False)) or bool(job.get("tracker_similarity_match", False)) or bool(job.get("repost_candidate", False))):
            continue
        tracker_row = find_matched_tracker_row(job, tracker, aliases)
        match_type = classify_tracker_match_type(job.get("tracker_overlay_reason", ""), job.get("status_source", ""))
        is_exact = bool(job.get("tracker_exact_match", False)) and has_strong_exact_tracker_evidence(job)
        excluded_from_ai = bool(job.get("previously_seen", False)) or bool(job.get("hard_skip_reason", ""))
        possible_repost = bool(job.get("tracker_similarity_match", False)) or bool(job.get("repost_candidate", False)) or (
            bool(job.get("tracker_exact_match", False)) and not is_exact
        )
        if is_exact and excluded_from_ai:
            recommended_action = "keep excluded; exact tracker evidence found"
        elif possible_repost:
            recommended_action = "treat as possible repost; keep for AI/manual review"
        else:
            recommended_action = "review tracker match"
        rows.append(
            {
                "file_name": job.get("file_name", ""),
                "parsed_company": job.get("company", ""),
                "parsed_job_title": job.get("job_title", ""),
                "parsed_location": job.get("location", ""),
                "apply_link": job.get("apply_link", ""),
                "external_job_id": get_external_job_id_text(job),
                "posted_date": job.get("job_posted_date", ""),
                "description_hash": get_description_hash_for_match(job),
                "matched_tracker_company": tracker_row.get("company", ""),
                "matched_tracker_job_title": tracker_row.get("job_title", ""),
                "matched_tracker_status": standardise_status(tracker_row.get("current_status", "")),
                "matched_tracker_apply_link": tracker_row.get("apply_link", ""),
                "matched_tracker_date": get_tracker_reference_date(tracker_row),
                "tracker_match_type": match_type,
                "is_exact_tracker_match": is_exact,
                "is_similar_tracker_match": bool(job.get("tracker_similarity_match", False)),
                "is_repost_candidate": possible_repost,
                "excluded_from_ai": excluded_from_ai,
                "exclusion_reason": job.get("tracker_overlay_reason", "") or job.get("hard_skip_reason", ""),
                "recommended_action": recommended_action,
            }
        )
    return pd.DataFrame(rows, columns=columns)


def build_pipeline_debug_audit(
    manual_rows,
    shown_ranked_jobs,
    ai_review_queue,
    manual_ai_coverage_audit,
    ranked_jobs_exclusion_audit,
    allowed_daily_actions,
):
    manual_jobs_columns = [
        "file_name",
        "input_type",
        "job_title",
        "company",
        "location",
        "final_action",
        "ai_final_action",
        "ai_fit_score",
        "ai_review_source",
        "next_action",
        "included_in_ranked_jobs",
        "included_in_ai_review_queue",
        "strict_excluded_before_ai",
        "strict_exclusion_reason",
        "tracker_exact_match",
        "tracker_match_type",
        "tracker_overlay_reason",
        "duplicate_reason",
        "coverage_status",
        "recommended_debug_action",
    ]
    ranked_check_columns = [
        "file_name",
        "job_title",
        "company",
        "final_action",
        "ai_fit_score",
        "should_be_in_ranked_jobs",
        "actually_in_ranked_jobs",
        "exclusion_reason",
        "investigation_note",
    ]
    queue_check_columns = [
        "file_name",
        "job_title",
        "company",
        "final_action",
        "ai_fit_score",
        "ai_review_source",
        "next_action",
        "reason_in_queue",
        "should_be_in_queue",
        "investigation_note",
    ]
    tracker_check_columns = [
        "file_name",
        "job_title",
        "company",
        "apply_link",
        "external_job_id",
        "description_hash",
        "matched_tracker_company",
        "matched_tracker_job_title",
        "matched_tracker_status",
        "matched_tracker_apply_link",
        "matched_tracker_job_id",
        "matched_tracker_description_hash",
        "tracker_match_type",
        "is_exact_tracker_match",
        "is_possible_repost",
        "excluded_from_ai",
        "exclusion_reason",
        "recommended_action",
    ]
    included_ranked_files = set(shown_ranked_jobs["file_name"].astype(str)) if "file_name" in shown_ranked_jobs.columns else set()
    included_queue_files = set(ai_review_queue["file_name"].astype(str)) if "file_name" in ai_review_queue.columns else set()
    coverage_by_file = {}
    if not manual_ai_coverage_audit.empty:
        coverage_by_file = dict(
            zip(
                manual_ai_coverage_audit["file_name"].astype(str),
                manual_ai_coverage_audit["coverage_status"].astype(str),
            )
        )
    ranked_exclusion_by_file = {}
    if not ranked_jobs_exclusion_audit.empty:
        ranked_exclusion_by_file = dict(
            zip(
                ranked_jobs_exclusion_audit["file_name"].astype(str),
                ranked_jobs_exclusion_audit["exclusion_reason"].astype(str),
            )
        )

    manual_job_rows = []
    ranked_check_rows = []
    tracker_rows = []
    tracker = load_applications_tracker()
    aliases = load_company_aliases()

    for _, job in manual_rows.sort_values(by=["file_name"]).iterrows():
        file_name = str(job.get("file_name", ""))
        strict_category, strict_reason = get_manual_pre_ai_exclusion(job)
        in_ranked = file_name in included_ranked_files
        in_queue = file_name in included_queue_files
        queue_reason = get_ai_review_queue_reason(job, manual_job_needs_ai_review(job))
        coverage_status = coverage_by_file.get(file_name, "")
        tracker_match_type = classify_tracker_match_type(job.get("tracker_overlay_reason", ""), job.get("status_source", ""))
        if coverage_status == "Pending AI review":
            debug_action = "rerun AI batch review"
        elif bool(job.get("tracker_exact_match", False)) and not has_strong_exact_tracker_evidence(job):
            debug_action = "possible repost; check tracker audit"
        elif str(job.get("final_action", "")) in allowed_daily_actions and not in_ranked:
            debug_action = "check ranked_jobs_exclusion_audit"
        elif in_queue:
            debug_action = "check ai_review_queue reason"
        else:
            debug_action = "no action"
        duplicate_reason = ""
        if pd.to_numeric(job.get("duplicate_count", 1), errors="coerce") > 1:
            duplicate_reason = f"duplicate_count={job.get('duplicate_count', '')}; merged_sources={job.get('merged_sources', '')}"
        manual_job_rows.append(
            {
                "file_name": file_name,
                "input_type": job.get("input_type", ""),
                "job_title": job.get("job_title", ""),
                "company": job.get("company", ""),
                "location": job.get("location", ""),
                "final_action": job.get("final_action", ""),
                "ai_final_action": job.get("ai_final_action", ""),
                "ai_fit_score": job.get("ai_fit_score", ""),
                "ai_review_source": job.get("ai_review_source", ""),
                "next_action": job.get("next_action", ""),
                "included_in_ranked_jobs": in_ranked,
                "included_in_ai_review_queue": in_queue,
                "strict_excluded_before_ai": bool(strict_category),
                "strict_exclusion_reason": strict_reason,
                "tracker_exact_match": job.get("tracker_exact_match", ""),
                "tracker_match_type": tracker_match_type,
                "tracker_overlay_reason": job.get("tracker_overlay_reason", ""),
                "duplicate_reason": duplicate_reason,
                "coverage_status": coverage_status,
                "recommended_debug_action": debug_action,
            }
        )
        if str(job.get("final_action", "")) in allowed_daily_actions:
            ranked_check_rows.append(
                {
                    "file_name": file_name,
                    "job_title": job.get("job_title", ""),
                    "company": job.get("company", ""),
                    "final_action": job.get("final_action", ""),
                    "ai_fit_score": job.get("ai_fit_score", ""),
                    "should_be_in_ranked_jobs": True,
                    "actually_in_ranked_jobs": in_ranked,
                    "exclusion_reason": ranked_exclusion_by_file.get(file_name, ""),
                    "investigation_note": "included" if in_ranked else ranked_exclusion_by_file.get(file_name, "missing from ranked_jobs without audit reason"),
                }
            )
        if bool(job.get("tracker_exact_match", False)) or bool(job.get("tracker_similarity_match", False)) or bool(job.get("previously_seen", False)):
            tracker_row = find_matched_tracker_row(job, tracker, aliases)
            is_exact = bool(job.get("tracker_exact_match", False)) and has_strong_exact_tracker_evidence(job)
            possible_repost = bool(job.get("tracker_similarity_match", False)) or bool(job.get("repost_candidate", False)) or (
                bool(job.get("tracker_exact_match", False)) and not is_exact
            )
            tracker_rows.append(
                {
                    "file_name": file_name,
                    "job_title": job.get("job_title", ""),
                    "company": job.get("company", ""),
                    "apply_link": job.get("apply_link", ""),
                    "external_job_id": get_external_job_id_text(job),
                    "description_hash": get_description_hash_for_match(job),
                    "matched_tracker_company": tracker_row.get("company", ""),
                    "matched_tracker_job_title": tracker_row.get("job_title", ""),
                    "matched_tracker_status": standardise_status(tracker_row.get("current_status", "")),
                    "matched_tracker_apply_link": tracker_row.get("apply_link", ""),
                    "matched_tracker_job_id": tracker_row.get("job_id", ""),
                    "matched_tracker_description_hash": get_description_hash_for_match(tracker_row) if tracker_row else "",
                    "tracker_match_type": tracker_match_type,
                    "is_exact_tracker_match": is_exact,
                    "is_possible_repost": possible_repost,
                    "excluded_from_ai": bool(job.get("previously_seen", False)) or bool(job.get("hard_skip_reason", "")),
                    "exclusion_reason": job.get("tracker_overlay_reason", "") or job.get("hard_skip_reason", ""),
                    "recommended_action": "keep excluded; exact tracker evidence found" if is_exact else "possible repost; review manually",
                }
            )

    queue_rows = []
    for _, job in ai_review_queue.sort_values(by=["file_name", "job_title"]).iterrows():
        manual_needs = manual_job_needs_ai_review(job) if is_manual_job(job) else False
        reason = get_ai_review_queue_reason(job, manual_needs)
        queue_rows.append(
            {
                "file_name": job.get("file_name", ""),
                "job_title": job.get("job_title", ""),
                "company": job.get("company", ""),
                "final_action": job.get("final_action", ""),
                "ai_fit_score": job.get("ai_fit_score", ""),
                "ai_review_source": job.get("ai_review_source", ""),
                "next_action": job.get("next_action", ""),
                "reason_in_queue": reason or "unknown queue reason",
                "should_be_in_queue": should_be_in_ai_review_queue(job, manual_needs),
                "investigation_note": "expected queue row" if reason else "unexpected queue row; check filter logic",
            }
        )

    return {
        "manual_jobs": pd.DataFrame(manual_job_rows, columns=manual_jobs_columns),
        "ranked_jobs_check": pd.DataFrame(ranked_check_rows, columns=ranked_check_columns),
        "ai_review_queue_check": pd.DataFrame(queue_rows, columns=queue_check_columns),
        "tracker_exclusion_check": pd.DataFrame(tracker_rows, columns=tracker_check_columns),
    }


def save_pipeline_debug_audit(path, sheets):
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path) as writer:
        for sheet_name, sheet_df in sheets.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name[:31], index=False)


def get_manual_ai_coverage_warning(job, coverage_status, selected_for_ai_queue=False):
    warnings = []
    source = str(job.get("ai_review_source", "")).strip()
    final_action = str(job.get("final_action", "")).strip()
    title = str(job.get("job_title", "")).strip()
    company = str(job.get("company", "")).strip()
    parse_confidence = str(job.get("manual_parse_confidence", "")).strip()
    strict_excluded_statuses = [
        "Strict excluded before AI",
        "Exact duplicate",
        "Empty or unreadable",
        "Contract / FTC skipped",
        "Exact tracker exclusion",
        "Senior leadership skipped",
        "Non-target role skipped",
        "Tracker excluded",
        "Parse failed",
    ]

    if is_blank_value(job.get("ai_fit_score", "")) and coverage_status not in strict_excluded_statuses:
        warnings.append("Manual job has no AI score")
    if source.lower() == "rule_based" and coverage_status not in strict_excluded_statuses:
        warnings.append("Rule-based result still present after AI batch run")
    if final_action == "Pending AI Review" and coverage_status not in strict_excluded_statuses:
        warnings.append("Pending AI Review remains")
    if (
        is_truthy_value(job.get("would_send_to_ai", False))
        and not selected_for_ai_queue
        and coverage_status not in ["AI reviewed"] + strict_excluded_statuses
    ):
        warnings.append("would_send_to_ai=True but not selected")
    if parse_confidence == "Low" and company.lower() in ["", "unknown"]:
        warnings.append("Suspicious company extraction")
    if parse_confidence == "Low" and title.lower() in ["", "unknown"]:
        warnings.append("Suspicious job title extraction")
    return "; ".join(dict.fromkeys(warnings))


def get_manual_ai_coverage_status(job, parse_failed=False, selected_for_ai_queue=False):
    tracker_excluded = (
        bool(job.get("tracker_exact_match", False))
        and bool(job.get("previously_seen", False))
        and has_strong_exact_tracker_evidence(job)
    )
    exclusion_category, _reason = get_manual_pre_ai_exclusion(job)
    ai_reviewed_source_and_score = (
        str(job.get("ai_review_source", "")).strip().lower() in LEGACY_VALID_AI_REVIEW_SOURCES
        and not is_blank_value(job.get("ai_fit_score", ""))
    )
    pending_signal = (
        str(job.get("final_action", "")).strip() == "Pending AI Review"
        or str(job.get("ai_review_source", "")).strip().lower() == "rule_based"
        or is_blank_value(job.get("ai_fit_score", ""))
    )
    warning = get_manual_ai_coverage_warning(job, "", selected_for_ai_queue)

    if parse_failed:
        return "Parse failed"
    if exclusion_category:
        return exclusion_category
    if tracker_excluded:
        return "Tracker excluded"
    if ai_reviewed_source_and_score:
        if pending_signal or is_truthy_value(job.get("would_send_to_ai", False)):
            return "Needs investigation"
        return "AI reviewed"
    if "would_send_to_ai=True but not selected" in warning:
        return "Needs investigation"
    if pending_signal:
        return "Pending AI review"
    return "Needs investigation"


def build_manual_pre_ai_exclusion_audit(manual_rows):
    columns = [
        "file_name",
        "parsed_job_title",
        "parsed_company",
        "parsed_location",
        "exclusion_status",
        "exclusion_reason",
        "would_send_to_ai",
        "strict_exclusion_category",
        "notes",
    ]
    audit_rows = []
    for _, job in manual_rows.sort_values(by=["file_name"]).iterrows():
        category, reason = get_manual_pre_ai_exclusion(job)
        would_send = is_truthy_value(job.get("would_send_to_ai", False))
        if category:
            status = "Strict excluded before AI"
        elif would_send:
            status = "Eligible for AI review"
        elif has_valid_ai_result(job):
            status = "AI reviewed"
        else:
            status = "Eligible pending review"
        audit_rows.append(
            {
                "file_name": job.get("file_name", ""),
                "parsed_job_title": job.get("job_title", ""),
                "parsed_company": job.get("company", ""),
                "parsed_location": job.get("location", ""),
                "exclusion_status": status,
                "exclusion_reason": reason,
                "would_send_to_ai": job.get("would_send_to_ai", ""),
                "strict_exclusion_category": category,
                "notes": job.get("manual_parse_notes", ""),
            }
        )
    return pd.DataFrame(audit_rows, columns=columns)


def build_manual_ai_coverage_audit(manual_rows, manual_load_summary, manual_queue_files):
    columns = [
        "file_name",
        "job_title",
        "company",
        "final_action",
        "ai_final_action",
        "ai_fit_score",
        "ai_review_source",
        "would_send_to_ai",
        "next_action",
        "tracker_exact_match",
        "tracker_overlay_reason",
        "parse_status",
        "coverage_status",
        "coverage_warning",
    ]
    parse_failures = {
        str(failure.get("file_name", "")): str(failure.get("reason", ""))
        for failure in manual_load_summary.get("raw_parse_failures", [])
    }
    audit_rows = []
    seen_files = set()

    for _, job in manual_rows.sort_values(by=["file_name"]).iterrows():
        file_name = str(job.get("file_name", ""))
        seen_files.add(file_name)
        parse_failed = file_name in parse_failures
        selected_for_ai_queue = file_name in manual_queue_files
        coverage_status = get_manual_ai_coverage_status(job, parse_failed, selected_for_ai_queue)
        coverage_warning = get_manual_ai_coverage_warning(job, coverage_status, selected_for_ai_queue)
        if parse_failed:
            coverage_warning = append_text_flag(coverage_warning, parse_failures[file_name])
        audit_rows.append(
            {
                "file_name": file_name,
                "job_title": job.get("job_title", ""),
                "company": job.get("company", ""),
                "final_action": job.get("final_action", ""),
                "ai_final_action": job.get("ai_final_action", ""),
                "ai_fit_score": job.get("ai_fit_score", ""),
                "ai_review_source": job.get("ai_review_source", ""),
                "would_send_to_ai": job.get("would_send_to_ai", ""),
                "next_action": job.get("next_action", ""),
                "tracker_exact_match": job.get("tracker_exact_match", ""),
                "tracker_overlay_reason": job.get("tracker_overlay_reason", ""),
                "parse_status": "Parse failed" if parse_failed else "Parsed",
                "coverage_status": coverage_status,
                "coverage_warning": coverage_warning,
            }
        )

    for file_name, reason in parse_failures.items():
        if file_name in seen_files:
            continue
        audit_rows.append(
            {
                "file_name": file_name,
                "job_title": "",
                "company": "",
                "final_action": "",
                "ai_final_action": "",
                "ai_fit_score": "",
                "ai_review_source": "",
                "would_send_to_ai": "",
                "next_action": "",
                "tracker_exact_match": "",
                "tracker_overlay_reason": "",
                "parse_status": "Parse failed",
                "coverage_status": "Parse failed",
                "coverage_warning": reason,
            }
        )

    return pd.DataFrame(audit_rows, columns=columns)


def recommend_next_action(
    score,
    red_flags,
    relevant_analyst_title=False,
    irrelevant_title=False,
    training_programme=False,
    senior_title=False,
    irrelevant_assistant=False,
):
    """Turn score plus red flags into a practical next step."""
    score = int(score)
    red_flags = str(red_flags).strip()

    if training_programme:
        return "Skip", "training programme / not a normal job"
    if senior_title:
        return "Skip", "senior, lead, manager or head of appears in the job title"
    if irrelevant_assistant:
        return "Skip", "assistant role without strong data signals"
    if irrelevant_title:
        return "Skip", "title is clearly irrelevant and does not contain strong data signals"
    if score < 35:
        return "Skip", "score is below 35"

    serious_flags = [
        "senior / lead / manager / head of",
        "admin / product upload heavy",
        "tender / RFQ / proposal heavy",
        "weak technical analysis",
        "training programme / not a normal job",
    ]
    has_serious_flag = any(flag in red_flags for flag in serious_flags)
    has_any_red_flag = bool(red_flags)
    has_contract_flag = (
        "fixed-term contract" in red_flags
        or "12-month contract" in red_flags
    )
    has_practice_flag = (
        has_any_red_flag
        or has_contract_flag
        or "weak technical analysis" in red_flags
        or "tender / RFQ / proposal heavy" in red_flags
    )

    if score >= 35 and has_practice_flag:
        return "Practice application only", "score is 35+ with red flags or weaker fit signals"
    if score >= 75 and not has_serious_flag:
        return "Apply now", "score is 75+ and no serious red flags were found"
    if score >= 60 and not has_serious_flag:
        return "Strong consider", "score is 60+ and no serious red flags were found"
    if score >= 45 and relevant_analyst_title:
        return "Review manually", "score is 45+ with a relevant analyst title"
    if score >= 35:
        return "Review manually", "score is 35+ and needs a manual check"

    return "Skip", "score is below 35"


def load_applications_tracker():
    """Load active and historical tracker rows for status matching."""
    tracker_tables = []

    for tracker_file in [TRACKER_FILE, TRACKER_HISTORY_FILE]:
        if not tracker_file.exists():
            continue

        tracker = pd.read_excel(tracker_file)

        if "status" in tracker.columns and "current_status" not in tracker.columns:
            tracker["current_status"] = tracker["status"]

        for column in TRACKER_COLUMNS:
            if column not in tracker.columns:
                tracker[column] = ""
        for column in OPTIONAL_TRACKER_MATCH_COLUMNS:
            if column not in tracker.columns:
                tracker[column] = ""

        tracker = tracker[TRACKER_COLUMNS + OPTIONAL_TRACKER_MATCH_COLUMNS].fillna("")
        tracker["apply_link"] = tracker["apply_link"].apply(clean_apply_link)
        tracker["current_status"] = tracker["current_status"].apply(standardise_status)
        tracker["matched_tracker_file"] = tracker_file.name
        tracker_tables.append(tracker)

    if not tracker_tables:
        return pd.DataFrame(columns=TRACKER_COLUMNS)

    tracker = pd.concat(tracker_tables, ignore_index=True)
    return tracker.drop_duplicates(subset=["job_id"], keep="last")


def get_tracker_reference_date(row):
    for column in ["status_date", "apply_date", "updated_date", "created_date"]:
        parsed = parse_clear_date(row.get(column, ""))
        if parsed:
            return parsed
    return ""


def get_repost_age_bucket(previous_rejection_date):
    parsed_date = parse_clear_date(previous_rejection_date)
    if not parsed_date:
        return "unknown"
    age_days = (date.today() - datetime.strptime(parsed_date, "%Y-%m-%d").date()).days
    if age_days <= 30:
        return "recent"
    if age_days <= 90:
        return "medium"
    return "old"


def repost_age_message(previous_rejection_date):
    bucket = get_repost_age_bucket(previous_rejection_date)
    if bucket == "recent":
        return "previous rejection within 30 days; low priority and review manually"
    if bucket == "medium":
        return "previous rejection 30-90 days ago; mild repost red flag"
    if bucket == "old":
        return "previous rejection more than 90 days ago; treat as a new opportunity with mild red flag"
    return "previous rejection date unknown; mild repost red flag"


def append_text_flag(existing_text, new_flag):
    existing_text = str(existing_text or "").strip()
    new_flag = str(new_flag or "").strip()
    if not new_flag:
        return existing_text
    if not existing_text:
        return new_flag
    if new_flag.lower() in existing_text.lower():
        return existing_text
    return f"{existing_text}; {new_flag}"


def same_known_posted_date(job, row):
    job_date = parse_clear_date(job.get("job_posted_date", ""))
    tracker_date = ""
    for column in ["job_posted_date", "posted_date"]:
        tracker_date = parse_clear_date(row.get(column, ""))
        if tracker_date:
            break
    return bool(job_date and tracker_date and job_date == tracker_date)


def normalize_description_for_match(value):
    text = normalize_for_match(value)
    return " ".join(text.split()[:800])


def get_description_hash_for_match(row):
    existing_hash = str(row.get("description_hash", "")).strip().lower()
    if existing_hash:
        return existing_hash
    description = str(row.get("job_text", "") or row.get("job_description", "") or "")
    normalized = normalize_description_for_match(description)
    if not normalized:
        return ""
    return hashlib.sha256(normalized.encode("utf-8", errors="ignore")).hexdigest()


def row_exact_match_reason(job, row, aliases):
    apply_link = normalize_url_for_match(job.get("apply_link", ""))
    tracker_link = normalize_url_for_match(row.get("apply_link", ""))
    if apply_link and tracker_link and apply_link == tracker_link:
        return "exact apply_link match"

    job_ids = extract_url_job_identifiers(job.get("apply_link", ""))
    tracker_ids = extract_url_job_identifiers(row.get("apply_link", ""))
    shared_ids = job_ids & tracker_ids
    if shared_ids:
        return f"exact external job id match ({sorted(shared_ids)[0]})"

    job_id = str(job.get("job_id", "") or job.get("canonical_job_id", "")).strip().lower()
    tracker_ids_to_check = [
        str(row.get("job_id", "")).strip().lower(),
        str(row.get("canonical_job_id", "")).strip().lower(),
    ]
    if job_id and job_id in tracker_ids_to_check:
        return "exact canonical job_id match"

    job_company = normalize_company_for_match(job.get("company", ""), aliases)
    tracker_company = normalize_company_for_match(row.get("company", ""), aliases)
    job_title = normalize_title_for_match(job.get("job_title", ""))
    tracker_title = normalize_title_for_match(row.get("job_title", ""))
    job_location = normalize_for_match(job.get("location", ""))
    tracker_location = normalize_for_match(row.get("location", ""))
    job_description_hash = get_description_hash_for_match(job)
    tracker_description_hash = get_description_hash_for_match(row)
    if (
        job_company
        and job_title
        and job_location
        and job_description_hash
        and tracker_location
        and tracker_description_hash
        and job_company == tracker_company
        and job_title == tracker_title
        and job_location == tracker_location
        and job_description_hash == tracker_description_hash
    ):
        return "exact company + title + location + description hash match"

    return ""


def tracker_similarity_mask(job, tracker, aliases):
    job_company = normalize_company_for_match(job.get("company", ""), aliases)
    job_title = normalize_title_for_match(job.get("job_title", ""))
    if not job_company or not job_title or tracker.empty:
        return pd.Series(False, index=tracker.index)

    normalized_company = tracker["company"].apply(lambda company: normalize_company_for_match(company, aliases))
    normalized_title = tracker["job_title"].apply(normalize_title_for_match)
    return (
        normalized_company.apply(lambda company: company == job_company or token_similarity(job_company, company) >= 0.75)
        & normalized_title.apply(lambda title: title == job_title or token_similarity(job_title, title) >= 0.8)
    )


def get_tracker_overlay(job, tracker, aliases):
    """Find exact tracker status and non-blocking similar-history signals."""

    empty = {
        "status": standardise_status(job.get("status", "new")),
        "matched_tracker_file": "",
        "matched_tracker_job_id": "",
        "previously_seen": False,
        "status_source": "",
        "tracker_overlay_reason": "",
        "previously_rejected_similar": False,
        "previous_rejection_date": "",
        "previous_rejected_company": "",
        "previous_rejected_title": "",
        "previous_rejected_job_id": "",
        "repost_candidate": False,
        "tracker_exact_match": False,
        "tracker_similarity_match": False,
        "previous_similar_status": "",
        "previous_application_status": "",
    }

    if tracker.empty:
        return empty

    def exact_overlay_from_row(row, reason):
        status = standardise_status(row.get("current_status", "new"))
        return {
            **empty,
            "status": status,
            "matched_tracker_file": row.get("matched_tracker_file", ""),
            "matched_tracker_job_id": row.get("job_id", ""),
            "previously_seen": status in TRACKER_EXCLUDED_STATUSES,
            "status_source": "tracker_overlay",
            "tracker_overlay_reason": reason,
            "tracker_exact_match": True,
            "tracker_similarity_match": False,
        }

    for _, row in tracker.iterrows():
        reason = row_exact_match_reason(job, row, aliases)
        if reason:
            return exact_overlay_from_row(row, reason)

    similar_rows = tracker[
        tracker_similarity_mask(job, tracker, aliases)
        & tracker["current_status"].apply(standardise_status).isin(TRACKER_SIMILAR_REVIEW_STATUSES)
    ].copy()
    if similar_rows.empty:
        return empty

    similar_rows["status_rank"] = similar_rows["current_status"].apply(
        lambda status: {"rejected": 0, "applied": 1, "assessment": 2, "interview": 3, "final_interview": 4, "offer": 5}.get(
            standardise_status(status),
            9,
        )
    )
    similar_rows = similar_rows.sort_values(by=["status_rank"])
    row = similar_rows.iloc[0]
    status = standardise_status(row.get("current_status", "new"))
    overlay = {
        **empty,
        "status": standardise_status(job.get("status", "new")),
        "matched_tracker_file": row.get("matched_tracker_file", ""),
        "matched_tracker_job_id": row.get("job_id", ""),
        "previously_seen": False,
        "status_source": "tracker_similarity",
        "tracker_overlay_reason": "Similar role was previously tracked, but this may be a repost or new vacancy",
        "tracker_exact_match": False,
        "tracker_similarity_match": True,
        "previous_similar_status": status,
        "previous_application_status": status,
    }
    if status == "rejected":
        overlay.update(
            {
                "previously_rejected_similar": True,
                "previous_rejection_date": get_tracker_reference_date(row),
                "previous_rejected_company": row.get("company", ""),
                "previous_rejected_title": row.get("job_title", ""),
                "previous_rejected_job_id": row.get("job_id", ""),
                "repost_candidate": True,
                "tracker_overlay_reason": "Similar role was previously rejected, but this may be a repost or new vacancy",
            }
        )
    return overlay

    return empty


def add_tracker_statuses(jobs):
    """Add a status column to ranked jobs by reading the applications tracker."""
    tracker = load_applications_tracker()
    aliases = load_company_aliases()
    overlays = jobs.apply(lambda job: get_tracker_overlay(job, tracker, aliases), axis=1)
    overlay_df = pd.DataFrame(list(overlays))
    overlay_df.index = jobs.index
    for column in overlay_df.columns:
        jobs[column] = overlay_df[column]
    return jobs


def apply_repost_candidate_annotations(jobs):
    """Annotate similar rejected tracker history without excluding new/reposted roles."""
    jobs = jobs.copy()
    if "previously_rejected_similar" not in jobs.columns:
        return jobs
    for column, default in [
        ("red_flags", ""),
        ("priority_reason", ""),
        ("final_action_reason", ""),
        ("human_review_reason", ""),
        ("needs_human_review", False),
        ("apply_recommendation", ""),
        ("application_tier", ""),
        ("next_action", ""),
    ]:
        if column not in jobs.columns:
            jobs[column] = default

    rejected_mask = jobs["previously_rejected_similar"].eq(True) & jobs["tracker_exact_match"].ne(True)
    if not rejected_mask.any():
        return jobs

    jobs.loc[rejected_mask, "tracker_overlay_reason"] = (
        "Similar role was previously rejected, but this may be a repost or new vacancy"
    )

    for row_index in jobs.index[rejected_mask]:
        previous_date = jobs.at[row_index, "previous_rejection_date"]
        age_note = repost_age_message(previous_date)
        jobs.at[row_index, "red_flags"] = append_text_flag(jobs.at[row_index, "red_flags"], age_note)
        jobs.at[row_index, "priority_reason"] = append_text_flag(
            jobs.at[row_index, "priority_reason"],
            "similar rejected role history; review repost manually",
        )
        jobs.at[row_index, "final_action_reason"] = append_text_flag(
            jobs.at[row_index, "final_action_reason"],
            "similar rejected role history; may be repost/new vacancy",
        )
        jobs.at[row_index, "human_review_reason"] = append_text_flag(
            jobs.at[row_index, "human_review_reason"],
            "similar rejected role history",
        )
        jobs.at[row_index, "needs_human_review"] = True

    recent_mask = rejected_mask & jobs["previous_rejection_date"].apply(get_repost_age_bucket).eq("recent")
    recent_non_manual = recent_mask & ~jobs.apply(is_manual_job, axis=1)
    if recent_non_manual.any():
        jobs.loc[recent_non_manual, "final_action"] = "Low Priority"
        jobs.loc[recent_non_manual, "apply_recommendation"] = "Low priority"
        jobs.loc[recent_non_manual, "application_tier"] = "Tier 4 = Low Priority"

    manual_mask = rejected_mask & jobs.apply(is_manual_job, axis=1)
    api_mask = rejected_mask & jobs.apply(is_api_job, axis=1)
    jobs.loc[manual_mask, "next_action"] = "Review repost manually / AI review full JD"
    jobs.loc[api_mask, "next_action"] = "Open link and check if this is a repost/new vacancy"
    return jobs


def main(ai_review=None, force_manual_review=False, force_files="", force_today=False, ai_batch_calls=None):
    """Read raw job files, rank the jobs, then save an Excel file."""
    profile = load_profile()
    jobs = load_all_jobs()
    manual_load_summary = jobs.attrs.get("manual_load_summary", {})

    if jobs.empty:
        print(f"No jobs found in {RAW_JOBS_DIR} or {API_JOBS_FILE}")
        raise SystemExit(1)

    # Extract simple fields from the full job description.
    if "job_title" not in jobs.columns:
        jobs["job_title"] = ""
    if "company" not in jobs.columns:
        jobs["company"] = ""
    if "location" not in jobs.columns:
        jobs["location"] = ""
    if "salary" not in jobs.columns:
        jobs["salary"] = ""
    if "date_added" not in jobs.columns:
        jobs["date_added"] = ""
    if "manual_parse_confidence" not in jobs.columns:
        jobs["manual_parse_confidence"] = ""
    if "manual_parse_notes" not in jobs.columns:
        jobs["manual_parse_notes"] = ""
    if "job_description" not in jobs.columns:
        jobs["job_description"] = ""

    jobs = jobs.fillna("")

    jobs["job_title"] = jobs.apply(
        lambda job: str(job["job_title"]).strip() or extract_job_title(job["job_text"], profile),
        axis=1,
    )
    jobs["company"] = jobs.apply(
        lambda job: str(job["company"]).strip() or extract_company(job["job_text"]),
        axis=1,
    )
    jobs["location"] = jobs.apply(
        lambda job: str(job["location"]).strip() or extract_location(job["job_text"]),
        axis=1,
    )
    jobs["salary"] = jobs.apply(
        lambda job: str(job["salary"]).strip() or extract_salary_from_text(job["job_text"]),
        axis=1,
    )
    jobs["job_description"] = jobs.apply(
        lambda job: str(job["job_description"]).strip() or str(job["job_text"]).strip(),
        axis=1,
    )

    # Score every raw job file and keep a short reason.
    score_results = jobs.apply(lambda job: score_job(job, profile), axis=1)
    jobs["score"] = score_results.apply(lambda result: result[0])
    jobs["reason"] = score_results.apply(lambda result: result[1])

    # Add freshness carefully. Unknown dates do not change the score.
    freshness_results = jobs.apply(get_freshness_fields, axis=1)
    freshness_fields = pd.DataFrame(list(freshness_results))
    for column in freshness_fields.columns:
        jobs[column] = freshness_fields[column]

    jobs["score"] = (jobs["score"] + jobs["freshness_score"]).clip(lower=0, upper=100)
    jobs["reason"] = jobs.apply(
        lambda job: (
            f"{job['reason']}; freshness: {int(job['freshness_score']):+d} ({job['freshness']})"
            if int(job["freshness_score"]) != 0
            else job["reason"]
        ),
        axis=1,
    )

    # Convert each score into an easy apply decision.
    apply_results = jobs["score"].apply(recommend_apply_action)
    jobs["apply_recommendation"] = apply_results.apply(lambda result: result[0])
    jobs["apply_reason"] = apply_results.apply(lambda result: result[1])

    # Add practical review guidance without changing the score itself.
    jobs["red_flags"] = jobs.apply(detect_red_flags, axis=1)
    contract_mask = jobs["red_flags"].astype(str).str.contains(
        "contract/temporary role - user prefers permanent roles",
        regex=False,
        na=False,
    )
    high_salary_mask = jobs["red_flags"].astype(str).str.contains(
        "high salary suggests mid/senior level",
        regex=False,
        na=False,
    )
    salary_55_mask = jobs.apply(lambda job: get_salary_max(job) >= 55000 and not has_junior_friendly_title(job), axis=1)
    salary_60_mask = jobs.apply(lambda job: get_salary_max(job) >= 60000 and not has_junior_friendly_title(job), axis=1)
    senior_experience_mask = jobs["red_flags"].astype(str).str.contains(
        "experience wording suggests mid/senior level",
        regex=False,
        na=False,
    )

    jobs.loc[contract_mask, "score"] = jobs.loc[contract_mask, "score"].clip(upper=30)
    jobs.loc[contract_mask, "reason"] = jobs.loc[contract_mask, "reason"].apply(
        lambda reason: f"{reason}; contract/temporary role - user prefers permanent roles"
    )
    jobs.loc[contract_mask, "apply_recommendation"] = "Skip"
    jobs.loc[contract_mask, "apply_reason"] = "contract/temporary role - user prefers permanent roles"

    jobs.loc[high_salary_mask, "reason"] = jobs.loc[high_salary_mask, "reason"].apply(
        lambda reason: f"{reason}; high salary suggests role may be above junior level"
    )
    jobs.loc[salary_55_mask, "score"] = jobs.loc[salary_55_mask, "score"].clip(upper=45)
    jobs.loc[salary_55_mask, "apply_recommendation"] = "Low priority"
    jobs.loc[salary_55_mask, "apply_reason"] = "high salary suggests role may be above junior level"
    jobs.loc[salary_60_mask, "score"] = jobs.loc[salary_60_mask, "score"].clip(upper=30)
    jobs.loc[salary_60_mask, "apply_recommendation"] = "Skip"
    jobs.loc[salary_60_mask, "apply_reason"] = "salary level suggests senior/mid-level role"
    jobs.loc[salary_60_mask, "reason"] = jobs.loc[salary_60_mask, "reason"].apply(
        lambda reason: f"{reason}; salary level suggests senior/mid-level role"
    )
    jobs.loc[senior_experience_mask, "score"] = jobs.loc[senior_experience_mask, "score"].clip(upper=55)
    jobs.loc[senior_experience_mask, "reason"] = jobs.loc[senior_experience_mask, "reason"].apply(
        lambda reason: f"{reason}; experience wording suggests role may be above junior level"
    )
    jobs["irrelevant_title_without_data"] = jobs.apply(has_irrelevant_title_without_data, axis=1)
    jobs["training_programme"] = jobs.apply(has_training_programme_flag, axis=1)
    jobs["senior_title"] = jobs.apply(has_senior_title_flag, axis=1)
    jobs["irrelevant_assistant"] = jobs.apply(has_irrelevant_assistant_role, axis=1)
    jobs["relevant_analyst_title"] = jobs["job_title"].apply(has_target_title_keyword)
    next_action_results = jobs.apply(
        lambda job: recommend_next_action(
            job["score"],
            job["red_flags"],
            job["relevant_analyst_title"],
            job["irrelevant_title_without_data"],
            job["training_programme"],
            job["senior_title"],
            job["irrelevant_assistant"],
        ),
        axis=1,
    )
    jobs["next_action"] = next_action_results.apply(lambda result: result[0])
    jobs["priority_reason"] = next_action_results.apply(lambda result: result[1])
    jobs.loc[contract_mask, "next_action"] = "Skip"
    jobs.loc[contract_mask, "priority_reason"] = "contract/temporary role - user prefers permanent roles"
    jobs.loc[salary_60_mask, "next_action"] = "Skip"
    jobs.loc[salary_60_mask, "priority_reason"] = "salary level suggests senior/mid-level role"
    jobs.loc[salary_55_mask & ~salary_60_mask, "next_action"] = "Review manually"
    jobs.loc[salary_55_mask & ~salary_60_mask, "priority_reason"] = "high salary suggests role may be above junior level"

    # Version 2.0 matching signals. These do not replace the core score; they
    # explain application fit from practical job-search angles.
    jobs["experience_fit"] = jobs.apply(calculate_experience_fit, axis=1)
    jobs["industry_fit"] = jobs.apply(calculate_industry_fit, axis=1)
    jobs["competition_level"] = jobs.apply(get_competition_level, axis=1)
    jobs["application_priority"] = jobs.apply(get_application_priority, axis=1)
    jobs["application_tier"] = jobs["application_priority"]
    jobs.loc[contract_mask, "application_priority"] = "Tier 4 = Skip"
    jobs.loc[contract_mask, "application_tier"] = "Skip"
    jobs.loc[salary_55_mask & ~salary_60_mask, "application_priority"] = "Tier 4 = Low Priority"
    jobs.loc[salary_55_mask & ~salary_60_mask, "application_tier"] = "Tier 4 = Low Priority"
    jobs.loc[salary_60_mask, "application_priority"] = "Tier 4 = Skip"
    jobs.loc[salary_60_mask, "application_tier"] = "Skip"
    jobs["estimated_interview_probability"] = jobs.apply(estimate_interview_probability, axis=1)

    # Recommend which CV version to use for each job.
    cv_results = jobs.apply(
        lambda job: recommend_cv(job["job_title"], job["job_text"]),
        axis=1,
    )
    jobs["recommended_cv"] = cv_results.apply(lambda result: result[0])
    jobs["cv_reason"] = cv_results.apply(lambda result: result[1])

    # Recommend which cover letter template to use for each job.
    cover_letter_results = jobs.apply(
        lambda job: recommend_cover_letter(job["job_title"], job["job_text"]),
        axis=1,
    )
    jobs["recommended_cover_letter"] = cover_letter_results.apply(lambda result: result[0])
    jobs["cover_letter_reason"] = cover_letter_results.apply(lambda result: result[1])

    # Add duplicate metadata before creating the daily review file. The full
    # audit file still keeps every row, including duplicates and score 0 jobs.
    jobs = add_duplicate_metadata(jobs)
    jobs = add_manual_duplicate_metadata(jobs)

    # Sort once by legacy action so tracker status matching receives stable rows.
    jobs["next_action_sort"] = jobs["next_action"].apply(get_next_action_rank)
    ranked_jobs = jobs.sort_values(
        by=["next_action_sort", "score", "source"],
        ascending=[True, False, True],
    )

    # Read current application statuses from tracker/applications.xlsx.
    ranked_jobs = add_tracker_statuses(ranked_jobs)
    ranked_jobs["salary_seniority_risk"] = ranked_jobs.apply(get_salary_seniority_risk, axis=1)
    ranked_jobs["contract_risk"] = ranked_jobs.apply(get_contract_risk, axis=1)
    ranked_jobs["junior_fit_reason"] = ranked_jobs.apply(get_junior_fit_reason, axis=1)
    ranked_jobs["tracker_exclusion_reason"] = ranked_jobs.apply(
        lambda job: (
            f"Excluded because already exists in tracker as {job['status']}"
            if bool(job.get("previously_seen", False))
            else ""
        ),
        axis=1,
    )
    ranked_jobs["apply_priority_score"] = ranked_jobs.apply(calculate_apply_priority_score, axis=1)
    ranked_jobs["action_bucket"] = ranked_jobs.apply(get_action_bucket, axis=1)
    ranked_jobs["action_reason"] = ranked_jobs.apply(get_action_reason, axis=1)
    ranked_jobs["action_bucket_sort"] = ranked_jobs["action_bucket"].apply(get_action_bucket_rank)
    ranked_jobs["final_action"] = ranked_jobs.apply(get_final_action, axis=1)
    ranked_jobs["final_action_reason"] = ranked_jobs.apply(get_final_action_reason, axis=1)
    ranked_jobs["rule_final_action"] = ranked_jobs["final_action"]
    ranked_jobs["rule_final_action_reason"] = ranked_jobs["final_action_reason"]
    ranked_jobs["hard_skip_reason"] = ranked_jobs.apply(get_hard_skip_reason, axis=1)
    ranked_jobs["needs_human_review"] = ranked_jobs.apply(needs_human_review, axis=1)
    ranked_jobs["human_review_reason"] = ranked_jobs.apply(get_human_review_reason, axis=1)
    ranked_jobs["location_fit_score"] = ranked_jobs.apply(get_location_fit_score, axis=1)
    ranked_jobs = apply_repost_candidate_annotations(ranked_jobs)
    ranked_jobs["pre_ai_bucket"] = ranked_jobs.apply(get_pre_ai_bucket, axis=1)
    ranked_jobs["pre_ai_reason"] = ranked_jobs.apply(get_pre_ai_reason, axis=1)

    ai_batch_status = {}
    if ai_review == "batch":
        force_today_files = sorted(get_raw_job_files_modified_today()) if force_today else None
        ranked_jobs = apply_ai_batch_reviews(
            ranked_jobs,
            enabled=True,
            force_manual_review=force_manual_review,
            force_files=force_files,
            force_today_files=force_today_files,
            manual_load_summary=manual_load_summary,
            max_api_calls_per_run_override=ai_batch_calls,
        )
        ai_batch_status = dict(ranked_jobs.attrs.get("ai_batch_status", {}))
    else:
        ranked_jobs = apply_ai_reviews(ranked_jobs, enabled=ai_review)
    ranked_jobs = apply_ai_decisions_to_final_action(ranked_jobs)
    ranked_jobs["needs_human_review"] = ranked_jobs.apply(needs_human_review, axis=1)
    ranked_jobs["human_review_reason"] = ranked_jobs.apply(get_human_review_reason, axis=1)
    review_apply_today_mask = ranked_jobs["needs_human_review"].eq(True) & ranked_jobs["final_action"].eq("Apply Today")
    ranked_jobs.loc[review_apply_today_mask, "final_action"] = "Strong Consider"
    ranked_jobs.loc[review_apply_today_mask, "final_action_reason"] = (
        ranked_jobs.loc[review_apply_today_mask, "final_action_reason"].astype(str)
        + "; human review warning"
    )
    ranked_jobs["would_send_to_ai"] = ranked_jobs.apply(would_send_to_ai, axis=1)
    ranked_jobs["ai_review_priority_score"] = ranked_jobs.apply(calculate_ai_review_priority_score, axis=1)

    ranked_jobs["action_bucket"] = ranked_jobs["final_action"]
    ranked_jobs["action_reason"] = ranked_jobs["final_action_reason"]
    ranked_jobs.loc[ranked_jobs["final_action"] == "Skip", "next_action"] = "Skip"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Skip", "apply_recommendation"] = "Skip"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Skip", "application_tier"] = "Skip"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Low Priority", "apply_recommendation"] = "Low priority"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Low Priority", "application_tier"] = "Tier 4 = Low Priority"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Manual Review", "apply_recommendation"] = "Manual review"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Manual Review", "application_tier"] = "Manual Review"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Manual Review", "next_action"] = "Review manually"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Consider", "apply_recommendation"] = "Consider"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Consider", "application_tier"] = "Tier 3 = Apply Only If High Score"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Consider", "next_action"] = "Review manually"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Apply If Time", "apply_recommendation"] = "Apply If Time"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Apply If Time", "application_tier"] = "Tier 3 = Apply If Time"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Apply If Time", "next_action"] = "Apply if time"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Pending AI Review", "apply_recommendation"] = "AI review"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Pending AI Review", "application_tier"] = "Pending AI Review"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Pending AI Review", "next_action"] = "AI review full JD"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Strong Consider", "apply_recommendation"] = "Strong Consideration"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Strong Consider", "application_tier"] = "Tier 2 = Apply Regularly"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Strong Consider", "next_action"] = "Strong consider"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Apply Today", "apply_recommendation"] = "Strong Consideration"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Apply Today", "application_tier"] = "Tier 1 = Apply Immediately"
    ranked_jobs.loc[ranked_jobs["final_action"] == "Apply Today", "next_action"] = "Apply now"
    ranked_jobs = apply_repost_candidate_annotations(ranked_jobs)
    ranked_jobs["final_action_sort"] = ranked_jobs["final_action"].apply(get_action_bucket_rank)
    ranked_jobs["date_added_sort"] = pd.to_datetime(ranked_jobs["date_added"], errors="coerce").fillna(pd.Timestamp.min)
    ranked_jobs = ranked_jobs.sort_values(
        by=["final_action_sort", "apply_priority_score", "date_added_sort", "location_fit_score", "score"],
        ascending=[True, False, False, False, False],
    )

    # Keep the output spreadsheet simple and easy to read.
    output_columns = [
        "input_type",
        "file_name",
        "date_added",
        "source",
        "dedupe_key",
        "merged_sources",
        "duplicate_count",
        "job_title",
        "company",
        "location",
        "salary",
        "manual_parse_confidence",
        "manual_parse_notes",
        "job_posted_date",
        "job_age_days",
        "freshness",
        "posted_date_confidence",
        "application_deadline",
        "deadline_confidence",
        "score",
        "apply_priority_score",
        "final_action",
        "final_action_reason",
        "hard_skip_reason",
        "pre_ai_bucket",
        "pre_ai_reason",
        "would_send_to_ai",
        "ai_review_priority_score",
        "needs_human_review",
        "human_review_reason",
        "location_fit_score",
        "ai_final_action",
        "ai_fit_score",
        "ai_fit_score_raw",
        "ai_fit_score_normalized",
        "ai_main_reason",
        "ai_red_flags",
        "ai_review_source",
        "source_priority",
        "ai_review_priority_reason",
        "action_bucket",
        "action_reason",
        "reason",
        "red_flags",
        "salary_seniority_risk",
        "contract_risk",
        "junior_fit_reason",
        "next_action",
        "priority_reason",
        "experience_fit",
        "industry_fit",
        "competition_level",
        "application_priority",
        "application_tier",
        "estimated_interview_probability",
        "apply_recommendation",
        "apply_reason",
        "recommended_cv",
        "cv_reason",
        "recommended_cover_letter",
        "cover_letter_reason",
        "status",
        "matched_tracker_file",
        "matched_tracker_job_id",
        "previously_seen",
        "status_source",
        "previously_rejected_similar",
        "previous_rejection_date",
        "previous_rejected_company",
        "previous_rejected_title",
        "previous_rejected_job_id",
        "repost_candidate",
        "tracker_exact_match",
        "tracker_similarity_match",
        "previous_similar_status",
        "previous_application_status",
        "tracker_overlay_reason",
        "tracker_exclusion_reason",
        "apply_link",
        "job_text",
        "job_description",
    ]
    all_ranked_jobs = ranked_jobs[output_columns]
    deduped_ranked_jobs = deduplicate_ranked_jobs(ranked_jobs)
    allowed_daily_actions = ["Apply Today", "Strong Consider", "Apply If Time"] if STRICT_MODE else ["Apply Today", "Strong Consider", "Apply If Time", "Consider", "Low Priority"]
    ranked_jobs_include_mask = (
        deduped_ranked_jobs.apply(is_manual_job, axis=1)
        & deduped_ranked_jobs.apply(has_valid_ai_decision, axis=1)
        & deduped_ranked_jobs.apply(is_actionable_status, axis=1)
        & deduped_ranked_jobs["final_action"].isin(allowed_daily_actions)
        & (deduped_ranked_jobs["previously_seen"] != True)
        & deduped_ranked_jobs["hard_skip_reason"].astype(str).str.strip().eq("")
        & (deduped_ranked_jobs["human_review_reason"] != "AI score/action inconsistency")
        & (
            (deduped_ranked_jobs["human_review_reason"] != "AI cache invalid")
            | deduped_ranked_jobs.apply(should_show_cache_invalid_rule_fallback, axis=1)
        )
        & deduped_ranked_jobs.apply(should_show_cache_invalid_rule_fallback, axis=1)
    )
    shown_ranked_jobs = deduped_ranked_jobs[ranked_jobs_include_mask][output_columns]
    ranked_jobs_exclusion_audit = build_ranked_jobs_exclusion_audit(
        ranked_jobs,
        shown_ranked_jobs,
        deduped_ranked_jobs,
        allowed_daily_actions,
    )
    include_api_leads_in_ai_queue = bool(cfg("AI_BATCH_INCLUDE_API_LEADS", False))
    api_leads = get_api_leads(ranked_jobs)
    manual_queue_candidate = ranked_jobs.apply(is_manual_job, axis=1)
    api_queue_candidate = ranked_jobs.apply(is_api_job, axis=1) & include_api_leads_in_ai_queue
    manual_needs_ai_review_mask = manual_queue_candidate & ranked_jobs.apply(manual_job_needs_ai_review, axis=1)
    queue_reason_series = ranked_jobs.apply(
        lambda job: get_ai_review_queue_reason(job, manual_job_needs_ai_review(job) if is_manual_job(job) else False),
        axis=1,
    )
    ai_review_queue = ranked_jobs[
        (
            (manual_queue_candidate & queue_reason_series.astype(str).str.strip().ne(""))
            | (
                api_queue_candidate
                & ranked_jobs.apply(get_api_lead_next_action, axis=1).isin(["Open link and copy full JD", "Low priority"])
            )
        )
        & (ranked_jobs["hard_skip_reason"].astype(str).str.strip() == "")
    ].sort_values(
        by=["source_priority", "ai_review_priority_score", "date_added_sort", "apply_priority_score"],
        ascending=[True, False, False, False],
    )[output_columns]
    manual_rows = ranked_jobs[ranked_jobs.apply(is_manual_job, axis=1)].copy()
    manual_queue_files = set(ai_review_queue[ai_review_queue["input_type"].astype(str).str.lower().eq("manual")]["file_name"].astype(str))
    tracker_exact_exclusions = int((ranked_jobs["tracker_exact_match"].eq(True) & ranked_jobs["previously_seen"].eq(True)).sum())
    similar_rejected_kept = int(
        (
            ranked_jobs["previously_rejected_similar"].eq(True)
            & ranked_jobs["tracker_exact_match"].ne(True)
            & ranked_jobs["previously_seen"].ne(True)
        ).sum()
    )
    similar_applied_interview_kept = int(
        (
            ranked_jobs["tracker_similarity_match"].eq(True)
            & ranked_jobs["tracker_exact_match"].ne(True)
            & ranked_jobs["previously_seen"].ne(True)
            & ranked_jobs["previous_similar_status"].isin(["applied", "assessment", "interview", "final_interview", "offer"])
        ).sum()
    )
    manual_excluded_tracker = int(
        (
            manual_rows["tracker_exact_match"].eq(True)
            & manual_rows["previously_seen"].eq(True)
        ).sum()
    ) if not manual_rows.empty else 0
    manual_kept_similar_rejected = int(
        (
            manual_rows["previously_rejected_similar"].eq(True)
            & manual_rows["tracker_exact_match"].ne(True)
            & manual_rows["previously_seen"].ne(True)
        ).sum()
    ) if not manual_rows.empty else 0
    manual_hard_skipped = int(manual_rows["hard_skip_reason"].astype(str).str.strip().ne("").sum()) if not manual_rows.empty else 0
    manual_sent_to_queue = len(manual_queue_files)
    manual_not_sent_to_queue = []
    for _, manual_job in manual_rows.iterrows():
        file_name = str(manual_job.get("file_name", ""))
        if file_name in manual_queue_files:
            continue
        reason = ""
        skipped_before_ai = False
        if bool(manual_job.get("previously_seen", False)):
            reason = str(manual_job.get("tracker_exclusion_reason", "")) or f"tracker status {manual_job.get('status', '')}"
        elif str(manual_job.get("hard_skip_reason", "")).strip():
            reason = str(manual_job.get("hard_skip_reason", "")).strip()
            skipped_before_ai = True
        elif has_valid_ai_result(manual_job):
            reason = "already has valid AI review result"
        else:
            reason = "not currently pending AI review"
        manual_not_sent_to_queue.append((file_name, reason, skipped_before_ai))
    manual_ai_coverage_audit = (
        build_manual_ai_coverage_audit(manual_rows, manual_load_summary, manual_queue_files)
        if ai_review == "batch"
        else pd.DataFrame()
    )
    manual_pre_ai_exclusion_audit = (
        build_manual_pre_ai_exclusion_audit(manual_rows)
        if ai_review == "batch"
        else pd.DataFrame()
    )
    repost_tracker_audit = build_repost_tracker_audit(manual_rows)
    pipeline_debug_audit = build_pipeline_debug_audit(
        manual_rows,
        shown_ranked_jobs,
        ai_review_queue,
        manual_ai_coverage_audit,
        ranked_jobs_exclusion_audit,
        allowed_daily_actions,
    )
    possible_repost_excluded = (
        repost_tracker_audit[
            repost_tracker_audit["excluded_from_ai"].eq(True)
            & repost_tracker_audit["is_exact_tracker_match"].ne(True)
        ]
        if not repost_tracker_audit.empty
        else pd.DataFrame()
    )

    # Create the output folder if it does not already exist.
    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)

    # Save both files:
    # - all_ranked_jobs.xlsx keeps every scored job for audit/history.
    # - ranked_jobs.xlsx is shorter and faster to review day to day.
    # If the Excel file is open, Windows may block Python from replacing it.
    try:
        all_ranked_jobs.to_excel(ALL_OUTPUT_FILE, index=False)
        shown_ranked_jobs.to_excel(OUTPUT_FILE, index=False)
        ranked_jobs_exclusion_audit.to_excel(RANKED_JOBS_EXCLUSION_AUDIT_FILE, index=False)
        ai_review_queue.to_excel(AI_REVIEW_QUEUE_FILE, index=False)
        api_leads.to_excel(API_LEADS_FILE, index=False)
        repost_tracker_audit.to_excel(REPOST_TRACKER_AUDIT_FILE, index=False)
        save_pipeline_debug_audit(PIPELINE_DEBUG_AUDIT_FILE, pipeline_debug_audit)
        if ai_review == "batch":
            manual_ai_coverage_audit.to_excel(MANUAL_AI_COVERAGE_AUDIT_FILE, index=False)
            manual_pre_ai_exclusion_audit.to_excel(MANUAL_PRE_AI_EXCLUSION_AUDIT_FILE, index=False)
    except PermissionError:
        print(
            f"Could not save {OUTPUT_FILE}, {ALL_OUTPUT_FILE}, {AI_REVIEW_QUEUE_FILE}, "
            f"{API_LEADS_FILE}, {RANKED_JOBS_EXCLUSION_AUDIT_FILE}, {MANUAL_AI_COVERAGE_AUDIT_FILE}, "
            f"{MANUAL_PRE_AI_EXCLUSION_AUDIT_FILE}, {PIPELINE_DEBUG_AUDIT_FILE}, or {REPOST_TRACKER_AUDIT_FILE}"
        )
        print("Please close the Excel files if they are open, then run this script again.")
        raise SystemExit(1)

    action_counts = all_ranked_jobs["next_action"].value_counts()
    final_action_counts = all_ranked_jobs["final_action"].value_counts()
    manual_final_action_counts = manual_rows["final_action"].value_counts() if not manual_rows.empty else pd.Series(dtype=int)
    duplicate_counts = pd.to_numeric(ranked_jobs["duplicate_count"], errors="coerce").fillna(1)
    invalid_cache_mask = ranked_jobs["ai_review_source"].eq("cache_invalid")
    invalid_cache_excluded = int(
        (
            invalid_cache_mask
            & ~ranked_jobs.index.isin(shown_ranked_jobs.index)
        ).sum()
    )
    cache_invalid_rule_kept = int(
        (
            invalid_cache_mask
            & ranked_jobs.apply(should_show_cache_invalid_rule_fallback, axis=1)
            & ranked_jobs["final_action"].isin(["Apply Today", "Strong Consider"])
        ).sum()
    )
    audit_status_counts = (
        manual_ai_coverage_audit["coverage_status"].value_counts()
        if ai_review == "batch" and not manual_ai_coverage_audit.empty
        else pd.Series(dtype=int)
    )
    parsed_manual_jobs_count = max(
        0,
        int(manual_load_summary.get("raw_files_found", 0)) - len(manual_load_summary.get("raw_parse_failures", [])),
    )
    manual_ai_unreviewed = (
        manual_ai_coverage_audit[
            manual_ai_coverage_audit["coverage_status"].isin(["Pending AI review", "Needs investigation"])
        ]
        if ai_review == "batch" and not manual_ai_coverage_audit.empty
        else pd.DataFrame()
    )
    manual_strict_excluded_before_ai = (
        int(manual_pre_ai_exclusion_audit["strict_exclusion_category"].astype(str).str.strip().ne("").sum())
        if ai_review == "batch" and not manual_pre_ai_exclusion_audit.empty
        else manual_hard_skipped
    )
    manual_eligible_for_ai_review = max(0, len(manual_rows) - manual_strict_excluded_before_ai)
    ranked_jobs_eligible = int(
        (
            ranked_jobs.apply(is_manual_job, axis=1)
            & ranked_jobs["final_action"].isin(allowed_daily_actions)
        ).sum()
    )
    ranked_jobs_written = len(shown_ranked_jobs)
    ranked_jobs_excluded = len(ranked_jobs_exclusion_audit)

    print()
    print("Ranking summary:")
    print(f"Manual raw job files found: {manual_load_summary.get('raw_files_found', 0)}")
    print(f"Manual raw job files parsed: {manual_load_summary.get('raw_files_parsed', 0)}")
    print(f"Manual raw job parse failures: {len(manual_load_summary.get('raw_parse_failures', []))}")
    print(f"Manual jobs added to candidate pool: {len(manual_rows)}")
    print(f"Tracker exact exclusions: {tracker_exact_exclusions}")
    print(f"Similar rejected repost candidates kept: {similar_rejected_kept}")
    print(f"Similar applied/interview candidates kept for review: {similar_applied_interview_kept}")
    print(f"Manual jobs excluded by exact tracker match: {manual_excluded_tracker}")
    print(f"Manual jobs kept despite similar rejected history: {manual_kept_similar_rejected}")
    print(f"Manual jobs strict excluded before AI: {manual_strict_excluded_before_ai}")
    print(f"Manual jobs eligible for AI review: {manual_eligible_for_ai_review}")
    print(f"Manual jobs sent to ai_review_queue: {manual_sent_to_queue}")
    for failure in manual_load_summary.get("raw_parse_failures", []):
        print(f"Manual raw parse fallback: {failure.get('file_name', '')} - {failure.get('reason', '')}")
    for file_name, reason, skipped_before_ai in manual_not_sent_to_queue:
        if skipped_before_ai:
            print(f"Manual raw job skipped before AI: {file_name} - {reason}")
        else:
            print(f"Manual raw job not sent to ai_review_queue: {file_name} - {reason}")
    print(f"Total jobs collected: {len(load_api_jobs())}")
    print(f"Hard skipped: {int((all_ranked_jobs['pre_ai_bucket'] == 'hard_skip').sum())}")
    print(f"Duplicates removed: {int((duplicate_counts > 1).sum())}")
    print(f"Previously applied excluded: {tracker_exact_exclusions}")
    print(f"ranked_jobs count: {len(shown_ranked_jobs)}")
    print(f"Ranked jobs eligible: {ranked_jobs_eligible}")
    print(f"Ranked jobs written: {ranked_jobs_written}")
    print(f"Ranked jobs excluded: {ranked_jobs_excluded}")
    print(f"Ranked jobs exclusion audit path: {RANKED_JOBS_EXCLUSION_AUDIT_FILE}")
    print(f"ai_review_queue count: {len(ai_review_queue)}")
    print(f"Pipeline debug audit path: {PIPELINE_DEBUG_AUDIT_FILE}")
    print(f"Repost tracker audit path: {REPOST_TRACKER_AUDIT_FILE}")
    if not possible_repost_excluded.empty:
        print("WARNING: Possible repost excluded. Check output/repost_tracker_audit.xlsx.")
    print(f"api_leads count: {len(api_leads)}")
    print(f"all_ranked_jobs count: {len(all_ranked_jobs)}")
    print(f"Overall final Apply Today count: {int(final_action_counts.get('Apply Today', 0))}")
    print(f"Overall final Strong Consider count: {int(final_action_counts.get('Strong Consider', 0))}")
    print(f"Overall final Apply If Time count: {int(final_action_counts.get('Apply If Time', 0))}")
    print(f"Manual Apply Today count: {int(manual_final_action_counts.get('Apply Today', 0))}")
    print(f"Manual Strong Consider count: {int(manual_final_action_counts.get('Strong Consider', 0))}")
    print(f"Manual Apply If Time count: {int(manual_final_action_counts.get('Apply If Time', 0))}")
    print(
        "Manual ranked-action total: "
        f"{int(manual_final_action_counts.get('Apply Today', 0)) + int(manual_final_action_counts.get('Strong Consider', 0)) + int(manual_final_action_counts.get('Apply If Time', 0))}"
    )
    print(
        "Overall ranked-action total including API leads: "
        f"{int(final_action_counts.get('Apply Today', 0)) + int(final_action_counts.get('Strong Consider', 0)) + int(final_action_counts.get('Apply If Time', 0))}"
    )
    print(f"Final Manual Review count: {int(final_action_counts.get('Manual Review', 0))}")
    print(f"Final Skip count: {int(final_action_counts.get('Skip', 0))}")
    print(f"Final Pending AI Review count: {int(final_action_counts.get('Pending AI Review', 0))}")
    print(f"Invalid cache rows excluded from ranked_jobs: {invalid_cache_excluded}")
    print(f"Rule-based fallback rows kept in ranked_jobs: {cache_invalid_rule_kept}")
    if ai_review == "batch":
        print()
        print("Manual AI coverage audit:")
        print(f"- Manual raw files found: {manual_load_summary.get('raw_files_found', 0)}")
        print(f"- Parsed manual jobs: {parsed_manual_jobs_count}")
        print(f"- AI reviewed: {int(audit_status_counts.get('AI reviewed', 0))}")
        print(f"- Pending AI review: {int(audit_status_counts.get('Pending AI review', 0))}")
        print(f"- Strict excluded before AI: {manual_strict_excluded_before_ai}")
        print(f"- Exact duplicate: {int(audit_status_counts.get('Exact duplicate', 0))}")
        print(f"- Empty or unreadable: {int(audit_status_counts.get('Empty or unreadable', 0))}")
        print(f"- Contract / FTC skipped: {int(audit_status_counts.get('Contract / FTC skipped', 0))}")
        print(f"- Exact tracker exclusion: {int(audit_status_counts.get('Exact tracker exclusion', 0))}")
        print(f"- Senior leadership skipped: {int(audit_status_counts.get('Senior leadership skipped', 0))}")
        print(f"- Non-target role skipped: {int(audit_status_counts.get('Non-target role skipped', 0))}")
        print(f"- Tracker excluded: {int(audit_status_counts.get('Tracker excluded', 0))}")
        print(f"- Parse failed: {int(audit_status_counts.get('Parse failed', 0))}")
        print(f"- Needs investigation: {int(audit_status_counts.get('Needs investigation', 0))}")
        if ai_batch_status.get("failed") or ai_batch_status.get("partial"):
            print()
            print("WARNING: AI batch review did not complete. Outputs may be partial.")
            print(f"jobs requested for AI review: {ai_batch_status.get('jobs_requested', 0)}")
            print(f"jobs successfully reviewed if known: {ai_batch_status.get('jobs_successfully_reviewed', 0)}")
            print(f"jobs still pending AI review: {ai_batch_status.get('jobs_still_pending', int(audit_status_counts.get('Pending AI review', 0)))}")
            if ai_batch_status.get("error_message"):
                print(f"AI batch error: {ai_batch_status.get('error_type', '')}: {ai_batch_status.get('error_message', '')}")
            print(f"next command to rerun: {ai_batch_status.get('next_command', 'python run_pipeline.py --ai-review-batch')}")
        if not manual_ai_unreviewed.empty:
            print("Some eligible manual jobs still need AI review. Re-run: python run_pipeline.py --ai-review-batch")
            print(
                "WARNING: Some manual jobs still do not have AI review results. "
                "Check output/manual_ai_coverage_audit.xlsx."
            )
    print(f"Apply now: {int(action_counts.get('Apply now', 0))}")
    print(f"Strong consider: {int(action_counts.get('Strong consider', 0))}")
    print(f"Apply if time: {int(action_counts.get('Apply if time', 0))}")
    print(f"Review manually: {int(action_counts.get('Review manually', 0))}")
    print(f"Skip: {int(action_counts.get('Skip', 0))}")
    print(f"Saved all ranked jobs to: {ALL_OUTPUT_FILE}")
    print(f"Saved ranked jobs to: {OUTPUT_FILE}")
    print(f"Saved ranked jobs exclusion audit to: {RANKED_JOBS_EXCLUSION_AUDIT_FILE}")
    print(f"Saved AI review queue to: {AI_REVIEW_QUEUE_FILE}")
    print(f"Saved pipeline debug audit to: {PIPELINE_DEBUG_AUDIT_FILE}")
    print(f"Saved repost tracker audit to: {REPOST_TRACKER_AUDIT_FILE}")
    if ai_review == "batch":
        print(f"Saved manual AI coverage audit to: {MANUAL_AI_COVERAGE_AUDIT_FILE}")
        print(f"Saved manual pre-AI exclusion audit to: {MANUAL_PRE_AI_EXCLUSION_AUDIT_FILE}")
    print(f"Saved API leads to: {API_LEADS_FILE}")

    return {
        "total_jobs_collected": len(load_api_jobs()),
        "total_jobs_scored": len(all_ranked_jobs),
        "jobs_shown": len(shown_ranked_jobs),
        "ai_review_queue": len(ai_review_queue),
        "api_leads": len(api_leads),
        "apply_now": int(action_counts.get("Apply now", 0)),
        "strong_consider": int(action_counts.get("Strong consider", 0)),
        "apply_if_time": int(action_counts.get("Apply if time", 0)),
        "review_manually": int(action_counts.get("Review manually", 0)),
        "skip": int(action_counts.get("Skip", 0)),
        "pending_ai_review": int(final_action_counts.get("Pending AI Review", 0)),
    }


if __name__ == "__main__":
    main()
