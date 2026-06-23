from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent
RAW_JOBS_DIR = PROJECT_ROOT / "jobs" / "raw_jobs"
ALL_RANKED_JOBS_FILE = PROJECT_ROOT / "output" / "all_ranked_jobs.xlsx"
AI_REVIEW_CACHE_FILE = PROJECT_ROOT / "data" / "ai_review_cache.csv"
MANUAL_AI_COVERAGE_AUDIT_FILE = PROJECT_ROOT / "output" / "manual_ai_coverage_audit.xlsx"

VALID_AI_REVIEW_SOURCES = {"cache_exact", "cache_similar", "openai_new", "api_new"}
AUDIT_COLUMNS = [
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


def is_blank(value):
    if pd.isna(value):
        return True
    return str(value).strip().lower() in ["", "nan", "none"]


def is_truthy(value):
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ["true", "1", "yes", "y"]


def normalize_text(value):
    text = str(value or "").lower()
    text = "".join(character if character.isalnum() else " " for character in text)
    return " ".join(text.split())


def canonical_job_key(row):
    return "|".join(
        [
            normalize_text(row.get("company", "")),
            normalize_text(row.get("job_title", "")),
            normalize_text(row.get("location", "")),
        ]
    )


def read_all_ranked_jobs():
    if not ALL_RANKED_JOBS_FILE.exists():
        return pd.DataFrame()
    jobs = pd.read_excel(ALL_RANKED_JOBS_FILE, dtype=object)
    jobs = jobs.where(pd.notna(jobs), "")
    if "input_type" not in jobs.columns:
        return pd.DataFrame()
    if "file_name" not in jobs.columns:
        return pd.DataFrame()
    return jobs[jobs["input_type"].astype(str).str.lower().eq("manual")].copy()


def read_ai_cache():
    if not AI_REVIEW_CACHE_FILE.exists():
        return pd.DataFrame()
    cache = pd.read_csv(AI_REVIEW_CACHE_FILE, dtype=object)
    cache = cache.where(pd.notna(cache), "")
    if "canonical_job_key" not in cache.columns:
        return pd.DataFrame()
    return cache


def latest_cache_result(cache, row):
    if cache.empty:
        return {}
    job_key = canonical_job_key(row)
    if not job_key or job_key == "||":
        return {}
    matches = cache[cache["canonical_job_key"].astype(str).eq(job_key)]
    if matches.empty:
        return {}
    if "last_seen_date" in matches.columns:
        matches = matches.sort_values(by=["last_seen_date"])
    return matches.iloc[-1].to_dict()


def append_warning(warnings, condition, message):
    if condition and message not in warnings:
        warnings.append(message)


def coverage_status(row, cache_row):
    tracker_excluded = is_truthy(row.get("tracker_exact_match", False)) and is_truthy(row.get("previously_seen", False))
    source = str(row.get("ai_review_source", "")).strip().lower()
    cache_score = cache_row.get("ai_fit_score", "") if cache_row else ""
    has_ranked_ai = source in VALID_AI_REVIEW_SOURCES and not is_blank(row.get("ai_fit_score", ""))
    has_cached_ai = not is_blank(cache_score)

    if tracker_excluded:
        return "Tracker excluded"
    if has_ranked_ai or has_cached_ai:
        if (
            str(row.get("final_action", "")).strip() == "Pending AI Review"
            or source == "rule_based"
            or is_truthy(row.get("would_send_to_ai", False))
        ):
            return "Needs investigation"
        return "AI reviewed"
    if (
        str(row.get("final_action", "")).strip() == "Pending AI Review"
        or source == "rule_based"
        or is_blank(row.get("ai_fit_score", ""))
    ):
        return "Pending AI review"
    return "Needs investigation"


def coverage_warning(row, status, cache_row, selected_in_ranked=True):
    warnings = []
    source = str(row.get("ai_review_source", "")).strip().lower()
    parse_confidence = str(row.get("manual_parse_confidence", "")).strip()
    company = str(row.get("company", "")).strip()
    title = str(row.get("job_title", "")).strip()

    append_warning(warnings, is_blank(row.get("ai_fit_score", "")) and status != "Tracker excluded", "Manual job has no AI score")
    append_warning(warnings, source == "rule_based", "Rule-based result still present after AI batch run")
    append_warning(warnings, str(row.get("final_action", "")).strip() == "Pending AI Review", "Pending AI Review remains")
    append_warning(
        warnings,
        is_truthy(row.get("would_send_to_ai", False)) and status != "AI reviewed",
        "would_send_to_ai=True but not selected" if not selected_in_ranked else "would_send_to_ai=True but still pending",
    )
    append_warning(warnings, parse_confidence == "Low" and company.lower() in ["", "unknown"], "Suspicious company extraction")
    append_warning(warnings, parse_confidence == "Low" and title.lower() in ["", "unknown"], "Suspicious job title extraction")
    append_warning(
        warnings,
        bool(cache_row) and not is_blank(cache_row.get("ai_fit_score", "")) and is_blank(row.get("ai_fit_score", "")),
        "AI cache result exists but all_ranked_jobs is not updated",
    )
    return "; ".join(warnings)


def audit_row_for_ranked_job(file_name, row, cache):
    cache_row = latest_cache_result(cache, row)
    status = coverage_status(row, cache_row)
    warning = coverage_warning(row, status, cache_row)
    return {
        "file_name": file_name,
        "job_title": row.get("job_title", ""),
        "company": row.get("company", ""),
        "final_action": row.get("final_action", ""),
        "ai_final_action": row.get("ai_final_action", cache_row.get("ai_final_action", "") if cache_row else ""),
        "ai_fit_score": row.get("ai_fit_score", cache_row.get("ai_fit_score", "") if cache_row else ""),
        "ai_review_source": row.get("ai_review_source", "cache_exact" if cache_row else ""),
        "would_send_to_ai": row.get("would_send_to_ai", ""),
        "next_action": row.get("next_action", ""),
        "tracker_exact_match": row.get("tracker_exact_match", ""),
        "tracker_overlay_reason": row.get("tracker_overlay_reason", ""),
        "parse_status": "Parsed",
        "coverage_status": status,
        "coverage_warning": warning,
    }


def audit_row_for_missing_file(file_name):
    return {
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
        "parse_status": "Missing from all_ranked_jobs",
        "coverage_status": "Parse missing",
        "coverage_warning": "Manual raw file is missing from all_ranked_jobs",
    }


def recommended_next_action(status_counts, suspicious_count, all_ranked_available):
    if not all_ranked_available:
        return "Run python run_pipeline.py --ai-review-batch to create output/all_ranked_jobs.xlsx."
    if int(status_counts.get("Needs investigation", 0)) > 0 or suspicious_count > 0:
        return "Open output/manual_ai_coverage_audit.xlsx and inspect Needs investigation or suspicious extraction rows."
    if int(status_counts.get("Pending AI review", 0)) > 0:
        return "Run python run_pipeline.py --ai-review-batch when OpenAI/network access is available."
    if int(status_counts.get("Parse missing", 0)) > 0:
        return "Run python run_pipeline.py --ai-review-batch and check raw job parsing errors."
    return "No action needed; manual raw jobs have AI coverage or valid exclusions."


def main():
    raw_files = sorted(path.name for path in RAW_JOBS_DIR.glob("*.txt")) if RAW_JOBS_DIR.exists() else []
    manual_jobs = read_all_ranked_jobs()
    cache = read_ai_cache()
    manual_by_file = {}

    if not manual_jobs.empty:
        manual_jobs = manual_jobs.copy()
        manual_jobs["file_name"] = manual_jobs["file_name"].astype(str)
        for _, row in manual_jobs.iterrows():
            file_name = str(row.get("file_name", ""))
            if file_name and file_name not in manual_by_file:
                manual_by_file[file_name] = row

    audit_rows = []
    for file_name in raw_files:
        row = manual_by_file.get(file_name)
        if row is None:
            audit_rows.append(audit_row_for_missing_file(file_name))
        else:
            audit_rows.append(audit_row_for_ranked_job(file_name, row, cache))

    audit = pd.DataFrame(audit_rows, columns=AUDIT_COLUMNS)
    MANUAL_AI_COVERAGE_AUDIT_FILE.parent.mkdir(parents=True, exist_ok=True)
    audit.to_excel(MANUAL_AI_COVERAGE_AUDIT_FILE, index=False)

    status_counts = audit["coverage_status"].value_counts() if not audit.empty else pd.Series(dtype=int)
    suspicious_count = int(audit["coverage_warning"].astype(str).str.contains("Suspicious", regex=False).sum()) if not audit.empty else 0
    next_action = recommended_next_action(status_counts, suspicious_count, ALL_RANKED_JOBS_FILE.exists())

    print("Manual raw files found:", len(raw_files))
    print("Manual jobs in all_ranked_jobs:", len(manual_by_file))
    print("AI reviewed:", int(status_counts.get("AI reviewed", 0)))
    print("Pending AI review:", int(status_counts.get("Pending AI review", 0)))
    print("Parse missing:", int(status_counts.get("Parse missing", 0)))
    print("Suspicious company/title:", suspicious_count)
    print("Recommended next action:", next_action)
    print(f"Saved manual AI coverage audit to: {MANUAL_AI_COVERAGE_AUDIT_FILE}")


if __name__ == "__main__":
    main()
