import hashlib
import json
import os
import socket
import time
from datetime import date, datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

import pandas as pd

try:
    import config
except ImportError:  # pragma: no cover - config exists in normal project runs.
    config = None


PROJECT_ROOT = Path(__file__).resolve().parents[1]
LOCAL_TIMEZONE = ZoneInfo("Europe/London")
VALID_AI_ACTIONS = ["Apply Today", "Strong Consider", "Apply If Time", "Consider", "Low Priority", "Skip", "Manual Review"]
VALID_AI_REVIEW_SOURCES = ["openai_new", "cache_exact", "cache_similar"]
LEGACY_VALID_AI_REVIEW_SOURCES = ["api_new", *VALID_AI_REVIEW_SOURCES]
PENDING_AI_REVIEW_SOURCES = ["", "nan", "none", "rule_based", "manual_pending"]
AI_REVIEW_CACHE_VERSION = 2
CACHE_COLUMNS = [
    "canonical_job_key",
    "description_hash",
    "normalized_company",
    "normalized_job_title",
    "normalized_location",
    "first_reviewed_date",
    "last_seen_date",
    "ai_final_action",
    "ai_fit_score",
    "ai_fit_score_raw",
    "ai_fit_score_normalized",
    "junior_realism",
    "skill_fit",
    "location_fit",
    "domain_fit",
    "ai_main_reason",
    "ai_red_flags",
    "ai_missing_requirements",
    "missing_requirements",
    "recommended_cv",
    "recommended_cover_letter",
    "cover_letter_angle",
    "model_used",
    "cache_version",
]
USAGE_LOG_COLUMNS = [
    "run_datetime",
    "mode",
    "batch_id",
    "batch_size",
    "canonical_job_key",
    "company",
    "job_title",
    "model_used",
    "ai_review_source",
    "input_chars",
    "description_chars",
    "api_call_made",
    "api_call_success",
    "error_type",
    "error_message",
    "tokens_used_if_available",
]
AI_OUTPUT_COLUMNS = [
    "ai_final_action",
    "ai_fit_score",
    "ai_fit_score_raw",
    "ai_fit_score_normalized",
    "ai_main_reason",
    "ai_red_flags",
    "ai_review_source",
    "source_priority",
    "ai_review_priority_reason",
]


def cfg(name, default):
    return getattr(config, name, default) if config else default


def project_path(path_value):
    path = Path(path_value)
    return path if path.is_absolute() else PROJECT_ROOT / path


def normalize_text(value):
    text = str(value or "").lower()
    text = "".join(character if character.isalnum() else " " for character in text)
    return " ".join(text.split())


def canonical_job_key(job):
    company = normalize_text(job.get("company", ""))
    title = normalize_text(job.get("job_title", ""))
    location = normalize_text(job.get("location", ""))
    return "|".join([company, title, location])


def get_description(job, max_chars):
    description = str(job.get("job_text", "") or job.get("description", "") or "")
    if len(description) > max_chars:
        return description[:max_chars]
    return description


def get_description_hash(description):
    return hashlib.sha256(description.encode("utf-8", errors="ignore")).hexdigest()


def ensure_ai_columns(jobs):
    for column in AI_OUTPUT_COLUMNS:
        if column not in jobs.columns:
            jobs[column] = ""
    source_replacements = {
        "": "rule_based",
        "api": "openai_new",
        "api_new": "openai_new",
        "cache": "cache_exact",
        "api_error": "skipped_quota",
    }
    jobs["ai_review_source"] = jobs["ai_review_source"].fillna("").astype(str).str.strip().replace(source_replacements)
    return jobs


def is_blank_value(value):
    if pd.isna(value):
        return True
    return str(value).strip().lower() in ["", "nan", "none"]


def is_truthy_value(value):
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in ["true", "1", "yes", "y"]


def has_valid_ai_result(job):
    return (
        str(job.get("ai_review_source", "")).strip().lower() in LEGACY_VALID_AI_REVIEW_SOURCES
        and str(job.get("ai_final_action", "")).strip() in VALID_AI_ACTIONS
        and not is_blank_value(job.get("ai_fit_score", ""))
    )


def manual_job_needs_ai_review(job):
    if str(job.get("input_type", "")).strip().lower() != "manual":
        return False
    source = str(job.get("ai_review_source", "")).strip().lower()
    return (
        source in PENDING_AI_REVIEW_SOURCES
        or str(job.get("final_action", "")).strip() == "Pending AI Review"
        or is_blank_value(job.get("ai_final_action", ""))
        or is_blank_value(job.get("ai_fit_score", ""))
        or is_truthy_value(job.get("would_send_to_ai", False))
    )


def read_table(path, columns):
    if not path.exists():
        return pd.DataFrame(columns=columns)
    table = pd.read_csv(path, dtype=object).fillna("")
    for column in columns:
        if column not in table.columns:
            table[column] = ""
    return table[columns]


def save_table(table, path, columns):
    path.parent.mkdir(parents=True, exist_ok=True)
    table = table.copy()
    for column in columns:
        if column not in table.columns:
            table[column] = ""
    table[columns].to_csv(path, index=False)


def get_local_now():
    return datetime.now(LOCAL_TIMEZONE).replace(microsecond=0)


def parse_usage_datetimes(usage_log):
    parsed = pd.to_datetime(usage_log["run_datetime"], errors="coerce")
    if getattr(parsed.dt, "tz", None) is None:
        return parsed.dt.tz_localize(LOCAL_TIMEZONE, nonexistent="shift_forward", ambiguous="NaT")
    return parsed.dt.tz_convert(LOCAL_TIMEZONE)


def successful_usage_mask(usage_log):
    return usage_log["api_call_success"].astype(str).str.lower().isin(["true", "1", "yes"])


def usage_period_mask(usage_log, period):
    if usage_log.empty:
        return pd.Series(False, index=usage_log.index)
    run_dates = parse_usage_datetimes(usage_log)
    today = get_local_now().date()
    if period == "day":
        return run_dates.dt.date == today
    return (run_dates.dt.year == today.year) & (run_dates.dt.month == today.month)


def usage_mode_masks(usage_log):
    if usage_log.empty:
        empty = pd.Series(False, index=usage_log.index)
        return empty, empty
    mode = (
        usage_log["mode"].astype(str).str.strip().str.lower()
        if "mode" in usage_log.columns
        else pd.Series("", index=usage_log.index)
    )
    batch_id = (
        usage_log["batch_id"].astype(str).str.strip()
        if "batch_id" in usage_log.columns
        else pd.Series("", index=usage_log.index)
    )
    batch_mask = mode.eq("batch") | batch_id.ne("")
    single_mask = mode.isin(["", "single_job"]) & ~batch_mask
    return single_mask, batch_mask


def count_batch_api_calls(usage_log, period):
    if usage_log.empty:
        return 0
    _, batch_mask = usage_mode_masks(usage_log)
    relevant = usage_log[successful_usage_mask(usage_log) & usage_period_mask(usage_log, period) & batch_mask]
    if relevant.empty:
        return 0
    batch_ids = relevant["batch_id"].astype(str).str.strip()
    unique_batch_ids = batch_ids[batch_ids.ne("")]
    fallback_rows = int(batch_ids.eq("").sum())
    return int(unique_batch_ids.nunique() + fallback_rows)


def count_single_api_calls(usage_log, period):
    if usage_log.empty:
        return 0
    single_mask, _ = usage_mode_masks(usage_log)
    relevant = successful_usage_mask(usage_log) & usage_period_mask(usage_log, period) & single_mask
    return int(relevant.sum())


def count_batch_jobs(usage_log, period):
    if usage_log.empty:
        return 0
    _, batch_mask = usage_mode_masks(usage_log)
    relevant = successful_usage_mask(usage_log) & usage_period_mask(usage_log, period) & batch_mask
    return int(relevant.sum())


def count_global_api_calls(usage_log, period):
    return count_single_api_calls(usage_log, period) + count_batch_api_calls(usage_log, period)


def usage_count(usage_log, period):
    return count_global_api_calls(usage_log, period)


def usage_quota_summary(usage_log):
    today = get_local_now().date()
    if usage_log.empty:
        return {
            "usage_date": "",
            "today": today.isoformat(),
            "daily_used_before_reset": 0,
            "daily_reset_applied": False,
            "daily_used_after_reset": 0,
            "monthly_used": 0,
        }

    run_dates = parse_usage_datetimes(usage_log)
    successful_calls = successful_usage_mask(usage_log)
    successful_dates = run_dates[successful_calls & run_dates.notna()].dt.date
    usage_date = successful_dates.max() if not successful_dates.empty else None
    daily_used_before_reset = (
        count_global_api_calls(usage_log, "day")
        if usage_date == today
        else int((successful_calls & (run_dates.dt.date == usage_date)).sum()) if usage_date else 0
    )
    daily_reset_applied = bool(usage_date and usage_date != today)
    daily_used_after_reset = 0 if daily_reset_applied else count_global_api_calls(usage_log, "day")
    monthly_used = count_global_api_calls(usage_log, "month")
    return {
        "usage_date": usage_date.isoformat() if usage_date else "",
        "today": today.isoformat(),
        "daily_used_before_reset": daily_used_before_reset,
        "daily_reset_applied": daily_reset_applied,
        "daily_used_after_reset": daily_used_after_reset,
        "monthly_used": monthly_used,
    }


def print_usage_reset_summary(summary):
    print("AI usage quota date check:")
    print("Timezone used: Europe/London")
    print(f"Usage date: {summary['usage_date'] or 'none'}")
    print(f"Today's date: {summary['today']}")
    print(f"Daily used before reset: {summary['daily_used_before_reset']}")
    print(f"Daily reset applied: {'YES' if summary['daily_reset_applied'] else 'NO'}")
    print(f"Daily used after reset: {summary['daily_used_after_reset']}")


def append_usage_rows(usage_path, rows):
    if not rows:
        return
    existing = read_table(usage_path, USAGE_LOG_COLUMNS)
    updated = pd.concat([existing, pd.DataFrame(rows)], ignore_index=True)
    save_table(updated, usage_path, USAGE_LOG_COLUMNS)


def is_retryable_openai_error(error):
    if isinstance(error, (TimeoutError, socket.timeout, URLError)):
        return True
    if isinstance(error, HTTPError):
        return error.code in [408, 429] or 500 <= error.code <= 599
    return False


def upsert_cache_result(cache, cache_row):
    exact_mask = (
        cache["canonical_job_key"].astype(str).eq(str(cache_row.get("canonical_job_key", "")))
        & cache["description_hash"].astype(str).eq(str(cache_row.get("description_hash", "")))
        & cache["model_used"].astype(str).eq(str(cache_row.get("model_used", "")))
    )
    if exact_mask.any():
        cache_index = cache[exact_mask].index[-1]
        for column, value in cache_row.items():
            if column in cache.columns:
                cache.at[cache_index, column] = value
        return cache
    return pd.concat([cache, pd.DataFrame([cache_row])], ignore_index=True)


def normalize_file_name(value):
    return str(value or "").strip().lower()


def parse_force_files(force_files):
    if not force_files:
        return set()
    if isinstance(force_files, str):
        raw_files = force_files.split(",")
    else:
        raw_files = force_files
    return {normalize_file_name(file_name) for file_name in raw_files if normalize_file_name(file_name)}


def should_force_manual_review(job, force_manual_review=False, force_files=None, force_today_files=None):
    if str(job.get("input_type", "")).lower() != "manual":
        return False
    file_name = normalize_file_name(job.get("file_name", ""))
    force_today_file_set = parse_force_files(force_today_files)
    if force_today_file_set:
        return file_name in force_today_file_set
    force_file_set = parse_force_files(force_files)
    if force_file_set:
        return file_name in force_file_set
    return bool(force_manual_review)


def is_ai_candidate(job):
    input_type = str(job.get("input_type", "")).lower()
    if input_type == "api" and not bool(cfg("AI_BATCH_INCLUDE_API_LEADS", False)):
        return False
    if input_type == "manual" and not manual_job_needs_ai_review(job):
        return False
    pre_ai_bucket = str(job.get("pre_ai_bucket", ""))
    if pre_ai_bucket:
        if pre_ai_bucket not in ["ai_review_candidate", "rule_apply_today", "rule_strong_consider"]:
            return False
    elif str(job.get("final_action", "")) not in ["Apply Today", "Strong Consider"]:
        return False
    if str(job.get("hard_skip_reason", "")).strip():
        return False
    if bool(job.get("previously_seen", False)):
        return False
    if str(job.get("status", "new")).lower() not in ["", "new", "shortlisted"]:
        return False
    if input_type != "manual" and str(job.get("contract_risk", "")).strip():
        return False
    return True


def get_source_priority(job):
    input_type = str(job.get("input_type", "")).lower()
    source = str(job.get("source", "")).lower()
    apply_link = str(job.get("apply_link", "")).lower()
    source_text = " ".join(
        [
            source,
            input_type,
            str(job.get("file_name", "")),
            apply_link,
            str(job.get("merged_sources", "")),
        ]
    ).lower()

    if input_type == "manual" and source not in ["linkedin", "indeed"]:
        return 1
    if any(term in source_text for term in ["company site", "company careers", "career site", "direct employer", "civil service jobs"]):
        return 2
    if input_type == "manual" and any(term in source_text for term in ["linkedin", "indeed"]):
        return 3
    if input_type == "manual":
        return 1
    if any(
        term in source_text
        for term in ["reed api", "adzuna api", "recruiter aggregator", "unknown api", "reed", "adzuna", "api"]
    ):
        return 4
    return 2


def get_ai_review_priority_reason(job):
    priority = get_source_priority(job)
    source = str(job.get("source", "") or job.get("input_type", "")).strip() or "unknown source"
    final_action = str(job.get("final_action", ""))
    junior_reason = str(job.get("junior_fit_reason", "")).strip()
    location_score = job.get("location_fit_score", "")

    if priority == 1:
        return (
            f"manual full JD prioritised before API jobs; {final_action}; "
            f"location_fit_score={location_score}; {junior_reason}"
        ).strip()
    if priority == 4:
        return (
            f"lower-priority API/aggregator source ({source}); review only after manual/direct jobs; "
            f"{final_action}; {junior_reason}"
        ).strip()
    if priority == 3:
        return f"manual LinkedIn/Indeed job; {final_action}; {junior_reason}".strip()
    return f"medium-priority source ({source}); {final_action}; {junior_reason}".strip()


def add_source_priority_columns(jobs):
    jobs = jobs.copy()
    jobs["source_priority"] = jobs.apply(get_source_priority, axis=1)
    jobs["ai_review_priority_reason"] = jobs.apply(get_ai_review_priority_reason, axis=1)
    return jobs


def get_ai_candidate_indexes(jobs, max_jobs, force_manual_review=False, force_files=None, force_today_files=None):
    candidates = jobs[jobs.apply(is_ai_candidate, axis=1)].copy()
    if candidates.empty:
        return []
    for column, default in [
        ("source_priority", 3),
        ("apply_priority_score", 0),
        ("location_fit_score", 0),
        ("score", 0),
        ("date_added", ""),
        ("ai_review_priority_score", 0),
    ]:
        if column not in candidates.columns:
            candidates[column] = default
    candidates["computed_ai_review_priority_score"] = candidates.apply(calculate_candidate_priority_score, axis=1)
    candidates["date_added_sort"] = pd.to_datetime(candidates["date_added"], errors="coerce").fillna(pd.Timestamp.min)
    candidates["forced_manual_review_sort"] = candidates.apply(
        lambda job: 1 if should_force_manual_review(job, force_manual_review, force_files, force_today_files) else 0,
        axis=1,
    )
    manual_candidates = candidates[candidates["input_type"].astype(str).str.lower().eq("manual")]
    if not manual_candidates.empty:
        candidates = manual_candidates
    candidates = candidates.sort_values(
        by=[
            "forced_manual_review_sort",
            "source_priority",
            "computed_ai_review_priority_score",
            "apply_priority_score",
            "location_fit_score",
            "score",
            "date_added_sort",
        ],
        ascending=[False, True, False, False, False, False, False],
    )
    if "dedupe_key" in candidates.columns:
        candidates = candidates.drop_duplicates(subset=["dedupe_key"], keep="first")
    if max_jobs is None:
        return list(candidates.index)
    return list(candidates.index[:max_jobs])


def build_force_today_summary(jobs, force_today_files):
    rows_by_file = {
        normalize_file_name(row.get("file_name", "")): row
        for _, row in jobs[jobs["input_type"].astype(str).str.lower().eq("manual")].iterrows()
    }
    eligible = []
    skipped = []
    for raw_file_name in sorted(force_today_files or []):
        normalized_name = normalize_file_name(raw_file_name)
        job = rows_by_file.get(normalized_name)
        if job is None:
            skipped.append((raw_file_name, "file was not parsed into a manual job row"))
            continue
        if is_ai_candidate(job):
            eligible.append(raw_file_name)
            continue
        reason = ""
        if str(job.get("hard_skip_reason", "")).strip():
            reason = str(job.get("hard_skip_reason", "")).strip()
        elif bool(job.get("previously_seen", False)):
            reason = str(job.get("tracker_overlay_reason", "")).strip() or "exact tracker exclusion"
        elif str(job.get("status", "new")).lower() not in ["", "new", "shortlisted"]:
            reason = f"status is {job.get('status', '')}"
        elif str(job.get("pre_ai_bucket", "")) not in ["ai_review_candidate", "rule_apply_today", "rule_strong_consider"]:
            reason = f"pre_ai_bucket is {job.get('pre_ai_bucket', '') or 'blank'}"
        else:
            reason = "not eligible for AI review"
        skipped.append((raw_file_name, reason))
    return {
        "files": sorted(force_today_files),
        "eligible": eligible,
        "skipped": skipped,
    }


def calculate_candidate_priority_score(job):
    configured_score = pd.to_numeric(job.get("ai_review_priority_score", 0), errors="coerce")
    if pd.notna(configured_score) and configured_score > 0:
        return float(configured_score)

    source_priority = pd.to_numeric(job.get("source_priority", 3), errors="coerce")
    apply_priority = pd.to_numeric(job.get("apply_priority_score", 0), errors="coerce")
    location_fit = pd.to_numeric(job.get("location_fit_score", 0), errors="coerce")
    match_score = pd.to_numeric(job.get("score", 0), errors="coerce")
    source_bonus = {1: 45, 2: 35, 3: 25, 4: 5}.get(4 if pd.isna(source_priority) else int(source_priority), 5)
    score = (
        source_bonus
        + (0 if pd.isna(apply_priority) else min(float(apply_priority), 100) * 0.3)
        + (0 if pd.isna(location_fit) else min(float(location_fit), 100) * 0.2)
        + (0 if pd.isna(match_score) else min(float(match_score), 100) * 0.1)
    )
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
    return max(0, score - penalty)


def build_prompt(job, description):
    job_payload = {
        "company": job.get("company", ""),
        "job_title": job.get("job_title", ""),
        "location": job.get("location", ""),
        "salary": job.get("salary", ""),
        "rule_based_final_action": job.get("final_action", ""),
        "rule_based_reason": job.get("final_action_reason", ""),
        "red_flags": job.get("red_flags", ""),
        "junior_fit_reason": job.get("junior_fit_reason", ""),
        "description": description,
    }
    return (
        "You are reviewing UK job adverts for a junior/early-career data analyst applicant. "
        "The applicant wants permanent junior-friendly analyst roles and does not want contract, "
        "temporary, training-course, senior, data-engineer, or domain-specific specialist roles. "
        "Return only valid JSON with these keys: ai_final_action, ai_fit_score, junior_realism, "
        "skill_fit, location_fit, domain_fit, ai_main_reason, ai_red_flags, missing_requirements, "
        "recommended_cv, recommended_cover_letter, cover_letter_angle. "
        "ai_final_action must be one of Apply Today, Strong Consider, Apply If Time, Consider, Low Priority, Skip, Manual Review. "
        "ai_fit_score must be an integer from 0 to 100. Apply Today normally requires 80+, "
        "Strong Consider normally requires 70+, Apply If Time normally requires 60+, Manual Review is for 50-59 or ambiguity. "
        "If the score and action do not align, use Manual Review. "
        "Do not give generic reasons. ai_main_reason must be one short, specific sentence based on the job description. "
        "It should mention concrete evidence such as graduate/junior/trainee wording, the exact location or hybrid/remote pattern, "
        "salary realism, Excel/SQL/Power BI/dashboard/reporting/KPI/data quality fit, domain fit or domain risk, "
        "and any experience requirement risk. Avoid repeating rule-based phrases unless the JD evidence supports them. "
        "Be strict: use Apply Today only for realistic permanent junior-friendly roles.\n\n"
        f"Job:\n{json.dumps(job_payload, ensure_ascii=False)}"
    )


def call_openai(api_key, model, prompt):
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": "Return concise, valid JSON only."},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }
    request = Request(
        "https://api.openai.com/v1/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urlopen(request, timeout=45) as response:
        response_payload = json.loads(response.read().decode("utf-8"))
    content = response_payload["choices"][0]["message"]["content"]
    return json.loads(content)


def call_openai_with_usage(api_key, model, prompt):
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": "Return concise, valid JSON only."},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0,
        "response_format": {"type": "json_object"},
    }
    request = Request(
        "https://api.openai.com/v1/chat/completions",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
        },
        method="POST",
    )
    with urlopen(request, timeout=60) as response:
        response_payload = json.loads(response.read().decode("utf-8"))
    content = response_payload["choices"][0]["message"]["content"]
    tokens_used = response_payload.get("usage", {}).get("total_tokens", "")
    return json.loads(content), tokens_used


def build_batch_prompt(batch_jobs):
    profile = {
        "name": "Wai Kit Chan",
        "location": "Milton Keynes, UK",
        "visa_sponsorship_required": False,
        "target_roles": [
            "Junior Data Analyst",
            "Graduate Data Analyst",
            "Data Analyst",
            "Reporting Analyst",
            "MI Analyst",
            "Junior BI Analyst",
            "Commercial Analyst",
            "Operations Analyst",
            "Customer Insight Analyst",
            "Product Insight Analyst",
            "People Data Analyst only if clearly analytics-focused",
        ],
        "skills": ["SQL", "BigQuery", "Power BI", "DAX basics", "Excel", "Python/pandas", "data cleaning", "dashboards", "KPI reporting", "Power Query"],
        "portfolio": ["UK Retail Analytics Dashboard", "Wine Analytics Dashboard", "Bellabeat case study"],
        "background": "Luxury retail sales, commercial/customer-facing experience, e-commerce/business experience",
        "preferred_locations": ["Milton Keynes", "Northampton", "Bedford", "Luton", "Birmingham", "London hybrid/remote", "remote"],
        "does_not_want": ["contract", "FTC", "temporary", "day-rate", "training courses", "career packages", "senior roles", "pure admin", "pure sales", "highly domain-specific roles requiring unavailable experience"],
    }
    return (
        "Review these UK job adverts for Wai Kit Chan. Use the profile, target roles, skip criteria, and scoring guide. "
        "Return exactly one JSON review per input job_id. Do not omit or duplicate any input job. Use this schema: "
        '{"reviews":[{"job_id":"string","canonical_job_key":"string","company":"string","job_title":"string",'
        '"ai_final_action":"Apply Today | Strong Consider | Apply If Time | Consider | Low Priority | Skip | Manual Review","ai_fit_score":0,'
        '"junior_realism":"High | Medium | Low","skill_fit":"High | Medium | Low","location_fit":"High | Medium | Low",'
        '"domain_fit":"High | Medium | Low","ai_main_reason":"string","ai_red_flags":["string"],'
        '"missing_requirements":["string"],"recommended_cv":"string","recommended_cover_letter":"string",'
        '"cover_letter_angle":"string"}]}. '
        "Use 0-100 scoring only: 90-100 Excellent/Apply Today, 80-89 Very strong/Apply Today or Strong Consider, "
        "70-79 Good/Strong Consider, 60-69 Borderline/Apply If Time or Consider, 50-59 Manual Review, 0-49 Skip or Low Priority. "
        "Do not use 1-10 scale. ai_main_reason must cite concrete JD evidence such as location/hybrid pattern, salary, "
        "Excel/SQL/Power BI/dashboard/reporting/KPI/data quality, graduate/junior wording, domain risk, or experience risk. "
        "Skip clearly unsuitable contract, training-course, senior, non-data, pure admin/sales, or impossible domain-specific roles.\n\n"
        f"User profile:\n{json.dumps(profile, ensure_ascii=False)}\n\n"
        f"Jobs:\n{json.dumps(batch_jobs, ensure_ascii=False)}"
    )


def get_cache_version():
    return int(cfg("AI_REVIEW_CACHE_VERSION", AI_REVIEW_CACHE_VERSION))


def normalize_ai_score(result):
    raw_value = result.get("ai_fit_score_raw", result.get("ai_fit_score", ""))
    normalized_value = result.get("ai_fit_score_normalized", "")
    raw_score = pd.to_numeric(raw_value, errors="coerce")
    normalized_score = pd.to_numeric(normalized_value, errors="coerce")
    action = str(result.get("ai_final_action", "")).strip()
    has_normalized_score = str(normalized_value).strip() != "" and pd.notna(normalized_score)

    if has_normalized_score:
        score = normalized_score
    elif pd.notna(raw_score) and raw_score <= 10 and (
        action in ["Apply Today", "Strong Consider", "Apply If Time", "Consider", "Low Priority"]
    ):
        score = raw_score * 10
    else:
        score = raw_score

    if pd.isna(score):
        return str(raw_value).strip(), None
    return str(raw_value).strip(), int(max(0, min(100, round(float(score)))))


def clean_ai_result(result, from_cache=False):
    action = str(result.get("ai_final_action", "")).strip()
    invalid_action = action not in VALID_AI_ACTIONS
    if invalid_action:
        action = "Manual Review"
    raw_score, score = normalize_ai_score(result)
    raw_red_flags = result.get("ai_red_flags", "")
    if isinstance(raw_red_flags, list):
        red_flags = "; ".join(str(flag).strip() for flag in raw_red_flags if str(flag).strip())
    else:
        red_flags = str(raw_red_flags).strip()
    raw_missing = result.get("missing_requirements", result.get("ai_missing_requirements", ""))
    if isinstance(raw_missing, list):
        missing_requirements = "; ".join(str(item).strip() for item in raw_missing if str(item).strip())
    else:
        missing_requirements = str(raw_missing).strip()
    normalized_score_changed = raw_score not in ["", str(score)] and score is not None
    inconsistent = (
        score is None
        or invalid_action
        or (from_cache and action == "Manual Review")
        or (action == "Apply Today" and score < 80)
        or (action == "Strong Consider" and score < 70)
        or (action == "Apply If Time" and score < 60)
        or (score < 50 and action in ["Apply Today", "Strong Consider", "Apply If Time"])
    )
    if inconsistent:
        action = "Manual Review"
        warning = "AI cache invalid" if from_cache else "AI score/action inconsistency"
        red_flags = "; ".join(filter(None, [red_flags, warning]))
    return {
        "ai_final_action": action,
        "ai_fit_score": None if score is None else str(score),
        "ai_fit_score_raw": raw_score,
        "ai_fit_score_normalized": None if score is None else str(score),
        "junior_realism": str(result.get("junior_realism", "")).strip(),
        "skill_fit": str(result.get("skill_fit", "")).strip(),
        "location_fit": str(result.get("location_fit", "")).strip(),
        "domain_fit": str(result.get("domain_fit", "")).strip(),
        "ai_main_reason": str(result.get("ai_main_reason", "")).strip(),
        "ai_red_flags": red_flags,
        "ai_missing_requirements": missing_requirements,
        "missing_requirements": missing_requirements,
        "recommended_cv": str(result.get("recommended_cv", "")).strip(),
        "recommended_cover_letter": str(result.get("recommended_cover_letter", "")).strip(),
        "cover_letter_angle": str(result.get("cover_letter_angle", "")).strip(),
        "cache_version": str(get_cache_version()),
        "ai_cache_valid": not inconsistent,
        "ai_score_normalized": normalized_score_changed,
    }


def apply_cache_result(jobs, row_index, cache_row, source="cache_exact"):
    jobs.at[row_index, "ai_final_action"] = cache_row.get("ai_final_action", "")
    jobs.at[row_index, "ai_fit_score"] = cache_row.get("ai_fit_score", "")
    jobs.at[row_index, "ai_fit_score_raw"] = cache_row.get("ai_fit_score_raw", "")
    jobs.at[row_index, "ai_fit_score_normalized"] = cache_row.get("ai_fit_score_normalized", "")
    jobs.at[row_index, "ai_main_reason"] = cache_row.get("ai_main_reason", "")
    jobs.at[row_index, "ai_red_flags"] = cache_row.get("ai_red_flags", "")
    jobs.at[row_index, "ai_review_source"] = source


def apply_review_result(jobs, row_index, result):
    jobs.at[row_index, "ai_final_action"] = result.get("ai_final_action", "")
    jobs.at[row_index, "ai_fit_score"] = result.get("ai_fit_score", "")
    jobs.at[row_index, "ai_fit_score_raw"] = result.get("ai_fit_score_raw", "")
    jobs.at[row_index, "ai_fit_score_normalized"] = result.get("ai_fit_score_normalized", "")
    jobs.at[row_index, "ai_main_reason"] = result.get("ai_main_reason", "")
    jobs.at[row_index, "ai_red_flags"] = result.get("ai_red_flags", "")
    jobs.at[row_index, "ai_review_source"] = "openai_new"


def description_similarity(left_job, right_job):
    left_tokens = set(normalize_text(get_description(left_job, 4000)).split())
    right_tokens = set(normalize_text(get_description(right_job, 4000)).split())
    if not left_tokens or not right_tokens:
        return 0
    return len(left_tokens & right_tokens) / len(left_tokens | right_tokens)


def propagate_review_to_duplicates(jobs, row_index, result, source):
    dedupe_key = str(jobs.at[row_index, "dedupe_key"]) if "dedupe_key" in jobs.columns else ""
    reviewed_job = jobs.loc[row_index]
    reviewed_key = canonical_job_key(reviewed_job)
    for duplicate_index, duplicate in jobs.iterrows():
        if duplicate_index == row_index:
            continue
        duplicate_source = str(duplicate.get("ai_review_source", "rule_based"))
        if duplicate_source not in ["", "rule_based", "skipped_quota", "skipped_no_key"]:
            continue
        duplicate_dedupe_key = str(duplicate.get("dedupe_key", ""))
        same_dedupe = bool(dedupe_key and duplicate_dedupe_key and duplicate_dedupe_key == dedupe_key)
        same_normalized_job = canonical_job_key(duplicate) == reviewed_key
        high_description_similarity = description_similarity(reviewed_job, duplicate) >= 0.85
        if same_dedupe or same_normalized_job or high_description_similarity:
            jobs.at[duplicate_index, "ai_final_action"] = result.get("ai_final_action", "")
            jobs.at[duplicate_index, "ai_fit_score"] = result.get("ai_fit_score", "")
            jobs.at[duplicate_index, "ai_fit_score_raw"] = result.get("ai_fit_score_raw", "")
            jobs.at[duplicate_index, "ai_fit_score_normalized"] = result.get("ai_fit_score_normalized", "")
            jobs.at[duplicate_index, "ai_main_reason"] = result.get("ai_main_reason", "")
            jobs.at[duplicate_index, "ai_red_flags"] = result.get("ai_red_flags", "")
            jobs.at[duplicate_index, "ai_review_source"] = source


def make_batch_job_payload(job, description, job_key):
    payload = {
        "job_id": str(job.get("job_id", "") or job.get("dedupe_key", "") or job_key),
        "canonical_job_key": job_key,
        "company": job.get("company", ""),
        "job_title": job.get("job_title", ""),
        "location": job.get("location", ""),
        "salary": job.get("salary", ""),
        "source": job.get("source", ""),
        "apply_priority_score": job.get("apply_priority_score", ""),
        "rule_final_action": job.get("final_action", ""),
        "rule_reason": job.get("final_action_reason", ""),
        "red_flags": job.get("red_flags", ""),
        "junior_fit_reason": job.get("junior_fit_reason", ""),
        "description": description,
    }
    return {key: "" if pd.isna(value) else str(value) for key, value in payload.items()}


def apply_ai_batch_reviews(
    jobs,
    enabled=None,
    require_confirmation=None,
    verbose=True,
    force_manual_review=False,
    force_files=None,
    force_today_files=None,
    manual_load_summary=None,
    max_api_calls_per_run_override=None,
):
    jobs = ensure_ai_columns(jobs.copy())
    jobs = add_source_priority_columns(jobs)
    default_enabled = bool(cfg("AI_REVIEW_ENABLED", False))
    enabled = default_enabled if enabled is None else bool(enabled)
    require_confirmation = bool(
        cfg("AI_BATCH_REQUIRE_CONFIRMATION", True) if require_confirmation is None else require_confirmation
    )
    if not enabled:
        if verbose:
            print("AI batch review disabled; using rule-based output only.")
        return jobs

    api_key = os.environ.get("OPENAI_API_KEY")
    model = cfg("AI_MODEL", "gpt-4o-mini")
    batch_size = int(cfg("AI_BATCH_SIZE", 10))
    max_jobs = int(cfg("AI_BATCH_MAX_JOBS_PER_RUN", 100))
    max_calls_per_run = (
        int(max_api_calls_per_run_override)
        if max_api_calls_per_run_override is not None
        else int(cfg("AI_BATCH_MAX_API_CALLS_PER_RUN", 10))
    )
    max_calls_per_run = max(0, max_calls_per_run)
    max_description_chars = int(cfg("AI_BATCH_MAX_DESCRIPTION_CHARS_PER_JOB", 2500))
    max_daily_api_calls = int(cfg("AI_BATCH_MAX_DAILY_API_CALLS", cfg("AI_REVIEW_MAX_DAILY_API_CALLS", 5)))
    max_daily_jobs = int(cfg("AI_BATCH_MAX_DAILY_JOBS", max_daily_api_calls * batch_size))
    max_monthly_api_calls = int(cfg("AI_BATCH_MAX_MONTHLY_API_CALLS", cfg("AI_REVIEW_MAX_MONTHLY_API_CALLS", 100)))
    max_monthly_jobs = int(cfg("AI_BATCH_MAX_MONTHLY_JOBS", max_monthly_api_calls * batch_size))
    global_max_daily_api_calls = int(cfg("AI_GLOBAL_MAX_DAILY_API_CALLS", cfg("AI_REVIEW_MAX_DAILY_API_CALLS", 5)))
    global_max_monthly_api_calls = int(cfg("AI_GLOBAL_MAX_MONTHLY_API_CALLS", cfg("AI_REVIEW_MAX_MONTHLY_API_CALLS", 100)))
    cache_path = project_path(cfg("AI_REVIEW_CACHE_PATH", "data/ai_review_cache.csv"))
    usage_path = project_path(cfg("AI_REVIEW_USAGE_LOG_PATH", "logs/ai_api_usage_log.csv"))
    retry_backoffs = [5, 15]
    batch_status = {
        "mode": "batch",
        "attempted": False,
        "completed": True,
        "partial": False,
        "failed": False,
        "error_type": "",
        "error_message": "",
        "jobs_requested": 0,
        "jobs_successfully_reviewed": 0,
        "jobs_still_pending": 0,
        "api_batches_requested": 0,
        "api_batches_completed": 0,
        "next_command": "python run_pipeline.py --ai-review-batch",
    }
    jobs.attrs["ai_batch_status"] = batch_status

    include_api_leads = bool(cfg("AI_BATCH_INCLUDE_API_LEADS", False))
    manual_load_summary = manual_load_summary or {}
    manual_raw_files_found = int(manual_load_summary.get("raw_files_found", 0) or 0)
    manual_raw_files_parsed = int(manual_load_summary.get("raw_files_parsed", 0) or 0)
    manual_candidate_mask = (
        jobs["input_type"].astype(str).str.lower().eq("manual")
        & jobs["hard_skip_reason"].astype(str).str.strip().eq("")
        & ~jobs["previously_seen"].eq(True)
    )
    manual_strict_excluded_mask = (
        jobs["input_type"].astype(str).str.lower().eq("manual")
        & jobs["hard_skip_reason"].astype(str).str.strip().ne("")
    )
    api_candidate_mask = (
        jobs["input_type"].astype(str).str.lower().eq("api")
        & jobs["pre_ai_bucket"].isin(["ai_review_candidate", "rule_apply_today", "rule_strong_consider"])
        & jobs["hard_skip_reason"].astype(str).str.strip().eq("")
        & ~jobs["previously_seen"].eq(True)
    )
    manual_full_jd_candidates = int(manual_candidate_mask.sum())
    manual_strict_excluded = int(manual_strict_excluded_mask.sum())
    api_lead_candidates = int(api_candidate_mask.sum())
    manual_valid_ai_result_before_cache = int(
        (manual_candidate_mask & jobs.apply(has_valid_ai_result, axis=1)).sum()
    )
    force_today_active = force_today_files is not None
    force_today_files = sorted(force_today_files or [])
    force_today_summary = build_force_today_summary(jobs, force_today_files) if force_today_active else None
    candidate_indexes = get_ai_candidate_indexes(jobs, None, force_manual_review, force_files, force_today_files)
    cache = read_table(cache_path, CACHE_COLUMNS)
    usage_log = read_table(usage_path, USAGE_LOG_COLUMNS)
    usage_rows = []
    pending = []
    cache_hits = 0
    manual_cache_hits = 0
    forced_manual_requeued = 0
    cache_invalid_rows = 0
    cache_invalid_requeued = 0
    cache_invalid_sent_to_api = 0
    score_normalized_rows = 0
    now = get_local_now().isoformat()
    today_text = get_local_now().date().isoformat()

    for row_index in candidate_indexes:
        job = jobs.loc[row_index]
        description = get_description(job, max_description_chars)
        job_key = canonical_job_key(job)
        desc_hash = get_description_hash(description)
        cache_match = pd.DataFrame()
        cache_source = "cache_exact"
        force_this_job = should_force_manual_review(job, force_manual_review, force_files, force_today_files)
        if not cache.empty and not force_this_job:
            cache_match = cache[
                (cache["canonical_job_key"] == job_key)
                & (cache["description_hash"] == desc_hash)
                & (cache["model_used"] == model)
            ]
            if cache_match.empty:
                cache_match = cache[(cache["canonical_job_key"] == job_key) & (cache["model_used"] == model)]
                cache_source = "cache_similar"
        if force_this_job:
            forced_manual_requeued += 1
            force_reason = "force_today" if parse_force_files(force_today_files) else "force_manual_review"
            pending.append((row_index, job_key, desc_hash, description, force_reason))
        elif not cache_match.empty:
            cache_row_index = cache_match.index[-1]
            cached_result = clean_ai_result(cache.loc[cache_row_index].to_dict(), from_cache=True)
            if cached_result.get("ai_score_normalized", False):
                score_normalized_rows += 1
            for result_key, result_value in cached_result.items():
                if result_key in cache.columns:
                    cache.at[cache_row_index, result_key] = result_value
            cache.at[cache_row_index, "last_seen_date"] = today_text
            if cached_result.get("ai_cache_valid", True):
                apply_cache_result(jobs, row_index, cached_result, source=cache_source)
                propagate_review_to_duplicates(jobs, row_index, cached_result, cache_source)
                cache_hits += 1
                if str(job.get("input_type", "")).lower() == "manual":
                    manual_cache_hits += 1
            else:
                jobs.at[row_index, "ai_review_source"] = "cache_invalid"
                jobs.at[row_index, "ai_final_action"] = "Manual Review"
                jobs.at[row_index, "ai_red_flags"] = cached_result.get("ai_red_flags", "AI cache invalid")
                cache_invalid_rows += 1
                pending.append((row_index, job_key, desc_hash, description, "cache_invalid"))
        else:
            pending.append((row_index, job_key, desc_hash, description, "new"))

    quota_summary = usage_quota_summary(usage_log)
    batch_api_calls_used_today = count_batch_api_calls(usage_log, "day")
    batch_api_calls_used_month = count_batch_api_calls(usage_log, "month")
    batch_jobs_used_today = count_batch_jobs(usage_log, "day")
    batch_jobs_used_month = count_batch_jobs(usage_log, "month")
    global_api_calls_used_today = count_global_api_calls(usage_log, "day")
    global_api_calls_used_month = count_global_api_calls(usage_log, "month")
    batch_daily_api_remaining = max(0, max_daily_api_calls - batch_api_calls_used_today)
    batch_monthly_api_remaining = max(0, max_monthly_api_calls - batch_api_calls_used_month)
    batch_daily_jobs_remaining = max(0, max_daily_jobs - batch_jobs_used_today)
    batch_monthly_jobs_remaining = max(0, max_monthly_jobs - batch_jobs_used_month)
    global_daily_remaining = max(0, global_max_daily_api_calls - global_api_calls_used_today)
    global_monthly_remaining = max(0, global_max_monthly_api_calls - global_api_calls_used_month)
    max_calls_allowed = min(
        max_calls_per_run,
        batch_daily_api_remaining,
        batch_monthly_api_remaining,
        global_daily_remaining,
        global_monthly_remaining,
    )
    max_jobs_reviewed = min(
        len(pending),
        max_jobs,
        max_calls_allowed * batch_size,
        batch_daily_jobs_remaining,
        batch_monthly_jobs_remaining,
    )
    pending_to_review = pending[:max_jobs_reviewed]
    batches = [pending_to_review[i : i + batch_size] for i in range(0, len(pending_to_review), batch_size)]
    manual_pending = [item for item in pending if str(jobs.loc[item[0]].get("input_type", "")).lower() == "manual"]
    manual_pending_to_review = [
        item for item in pending_to_review if str(jobs.loc[item[0]].get("input_type", "")).lower() == "manual"
    ]
    manual_blocked_by_limit = max(0, len(manual_pending) - len(manual_pending_to_review))
    cache_invalid_sent_to_api = sum(1 for item in pending_to_review if len(item) > 4 and item[4] == "cache_invalid")
    cache_invalid_requeued = max(0, cache_invalid_rows - cache_invalid_sent_to_api)
    batch_status["jobs_requested"] = len(pending_to_review)
    batch_status["jobs_still_pending"] = len(pending)
    batch_status["api_batches_requested"] = len(batches)
    jobs.attrs["ai_batch_status"] = batch_status

    print("AI batch review preflight:")
    print(f"OPENAI_API_KEY detected: {'YES' if api_key else 'NO'}")
    print(f"Manual raw job files found: {manual_raw_files_found}")
    print(f"Manual raw job files parsed: {manual_raw_files_parsed}")
    print(f"Manual jobs strict excluded before AI: {manual_strict_excluded}")
    print(f"Manual jobs eligible for AI review: {manual_full_jd_candidates}")
    print(f"Manual jobs with valid AI result: {manual_valid_ai_result_before_cache + manual_cache_hits}")
    print(f"Manual jobs pending AI review: {len(manual_pending)}")
    print(f"Manual jobs needing AI review this run: {len(manual_pending_to_review)}")
    print(f"Manual jobs blocked by per-run limit: {manual_blocked_by_limit}")
    if manual_blocked_by_limit:
        print(f"Manual jobs remaining pending after this run: {manual_blocked_by_limit}")
    print(f"Manual jobs force re-review: {forced_manual_requeued}")
    if force_today_summary:
        print(f"Manual raw files modified today: {len(force_today_summary['files'])}")
        print(f"Today files eligible for force review: {len(force_today_summary['eligible'])}")
        if force_today_summary["eligible"]:
            print("Today force review files:")
            for file_name in force_today_summary["eligible"]:
                print(f"- {file_name}")
        print(f"Today files skipped with reason: {len(force_today_summary['skipped'])}")
        for file_name, reason in force_today_summary["skipped"]:
            print(f"- {file_name}: {reason}")
    print(f"API lead candidates: {api_lead_candidates}")
    print(f"AI_BATCH_INCLUDE_API_LEADS: {include_api_leads}")
    print(f"Valid cache hits: {cache_hits}")
    print(f"New jobs to review: {len(pending)}")
    print(f"Batch API calls needed: {len(batches)}")
    print(f"Batch API calls allowed this run: {max_calls_allowed}")
    print(f"Model used: {model}")

    if not api_key:
        print("AI batch review skipped because OPENAI_API_KEY is not set; cached rows were still applied where available.")
        skipped_indexes = [item[0] for item in pending]
        jobs.loc[jobs.index.isin(skipped_indexes) & jobs["ai_review_source"].eq("rule_based"), "ai_review_source"] = "skipped_no_key"
        save_table(cache, cache_path, CACHE_COLUMNS)
        return jobs

    if not batches:
        save_table(cache, cache_path, CACHE_COLUMNS)
        return jobs

    if require_confirmation:
        answer = input(f"Proceed with {len(batches)} batch API calls reviewing {len(pending_to_review)} jobs? (Y/N) ").strip().lower()
        if answer != "y":
            print("AI batch review skipped by user; using rule-based output only for uncached rows.")
            skipped_indexes = [item[0] for item in pending]
            jobs.loc[jobs.index.isin(skipped_indexes) & jobs["ai_review_source"].eq("rule_based"), "ai_review_source"] = "skipped_quota"
            save_table(cache, cache_path, CACHE_COLUMNS)
            return jobs

    reviewed_row_indexes = set()
    failed_error = None
    for batch_number, batch in enumerate(batches, start=1):
        batch_id = f"{get_local_now().strftime('%Y%m%d%H%M%S')}_{batch_number}"
        batch_payload = []
        batch_lookup = {}
        expected_row_indexes = set()
        for row_index, job_key, desc_hash, description, _reason in batch:
            job = jobs.loc[row_index]
            payload = make_batch_job_payload(job, description, job_key)
            batch_payload.append(payload)
            expected_row_indexes.add(row_index)
            batch_lookup[str(payload["job_id"])] = (row_index, job_key, desc_hash, description)
            batch_lookup[job_key] = (row_index, job_key, desc_hash, description)
        prompt = build_batch_prompt(batch_payload)
        response = None
        tokens_used = ""
        for attempt_number in range(1, len(retry_backoffs) + 2):
            try:
                batch_status["attempted"] = True
                response, tokens_used = call_openai_with_usage(api_key, model, prompt)
                break
            except HTTPError as error:
                if error.code == 401:
                    print("OpenAI API authentication failed. Check OPENAI_API_KEY, project permissions, billing, and model access.")
                    failed_error = error
                    break
                if is_retryable_openai_error(error) and attempt_number <= len(retry_backoffs):
                    backoff_seconds = retry_backoffs[attempt_number - 1]
                    print(f"AI batch review attempt {attempt_number} failed: {error}. Retrying in {backoff_seconds} seconds...")
                    time.sleep(backoff_seconds)
                    continue
                failed_error = error
                break
            except (URLError, TimeoutError, socket.timeout) as error:
                if is_retryable_openai_error(error) and attempt_number <= len(retry_backoffs):
                    backoff_seconds = retry_backoffs[attempt_number - 1]
                    print(f"AI batch review attempt {attempt_number} failed: {error}. Retrying in {backoff_seconds} seconds...")
                    time.sleep(backoff_seconds)
                    continue
                failed_error = error
                break
            except (json.JSONDecodeError, KeyError) as error:
                failed_error = error
                break

        if response is None:
            error_type = type(failed_error).__name__ if failed_error else "UnknownError"
            error_message = str(failed_error or "AI batch review failed")
            print(f"AI batch review failed: {error_message}")
            batch_status["completed"] = False
            batch_status["partial"] = True
            batch_status["failed"] = True
            batch_status["error_type"] = error_type
            batch_status["error_message"] = error_message
            usage_rows.append(
                {
                    "run_datetime": now,
                    "mode": "batch",
                    "batch_id": batch_id,
                    "batch_size": len(batch),
                    "canonical_job_key": "",
                    "company": "",
                    "job_title": "",
                    "model_used": model,
                    "ai_review_source": "",
                    "input_chars": len(prompt),
                    "description_chars": "",
                    "api_call_made": True,
                    "api_call_success": False,
                    "error_type": error_type,
                    "error_message": error_message,
                    "tokens_used_if_available": "",
                }
            )
            break

        reviews = response.get("reviews", [])
        seen_row_indexes = set()
        for review in reviews:
            lookup_key = str(review.get("job_id", "")) or str(review.get("canonical_job_key", ""))
            if lookup_key not in batch_lookup and str(review.get("canonical_job_key", "")) in batch_lookup:
                lookup_key = str(review.get("canonical_job_key", ""))
            if lookup_key not in batch_lookup:
                continue
            row_index, job_key, desc_hash, description = batch_lookup[lookup_key]
            if row_index in seen_row_indexes:
                print(f"AI batch review returned duplicate review for {lookup_key}; ignoring duplicate.")
                continue
            seen_row_indexes.add(row_index)
            job = jobs.loc[row_index]
            result = clean_ai_result(review, from_cache=False)
            apply_review_result(jobs, row_index, result)
            propagate_review_to_duplicates(jobs, row_index, result, "cache_similar")
            reviewed_row_indexes.add(row_index)
            cache_row = {
                "canonical_job_key": job_key,
                "description_hash": desc_hash,
                "normalized_company": normalize_text(job.get("company", "")),
                "normalized_job_title": normalize_text(job.get("job_title", "")),
                "normalized_location": normalize_text(job.get("location", "")),
                "first_reviewed_date": today_text,
                "last_seen_date": today_text,
                "model_used": model,
                **result,
            }
            cache = upsert_cache_result(cache, cache_row)
            usage_rows.append(
                {
                    "run_datetime": now,
                    "mode": "batch",
                    "batch_id": batch_id,
                    "batch_size": len(batch),
                    "canonical_job_key": job_key,
                    "company": job.get("company", ""),
                    "job_title": job.get("job_title", ""),
                    "model_used": model,
                    "ai_review_source": "openai_new",
                    "input_chars": len(prompt),
                    "description_chars": len(description),
                    "api_call_made": True,
                    "api_call_success": True,
                    "error_type": "",
                    "error_message": "",
                    "tokens_used_if_available": tokens_used,
                }
            )
        missing_reviews = sorted(expected_row_indexes - seen_row_indexes)
        if missing_reviews:
            print(f"AI batch review missing {len(missing_reviews)} job review(s); leaving them in the AI review queue.")
            batch_status["partial"] = True
        batch_status["api_batches_completed"] += 1

    save_table(cache, cache_path, CACHE_COLUMNS)
    append_usage_rows(usage_path, usage_rows)
    still_pending_indexes = [
        item[0]
        for item in pending
        if not has_valid_ai_result(jobs.loc[item[0]])
    ]
    batch_status["jobs_successfully_reviewed"] = len(reviewed_row_indexes)
    batch_status["jobs_still_pending"] = len(still_pending_indexes)
    if batch_status["failed"] or batch_status["partial"]:
        print("WARNING: AI batch review did not complete. Outputs may be partial.")
        print(f"jobs requested for AI review: {batch_status['jobs_requested']}")
        print(f"jobs successfully reviewed if known: {batch_status['jobs_successfully_reviewed']}")
        print(f"jobs still pending AI review: {batch_status['jobs_still_pending']}")
        print(f"next command to rerun: {batch_status['next_command']}")
    jobs.attrs["ai_batch_status"] = batch_status
    return jobs


def apply_ai_reviews(jobs, enabled=None, require_confirmation=None, verbose=True):
    jobs = ensure_ai_columns(jobs.copy())
    jobs = add_source_priority_columns(jobs)
    default_enabled = bool(cfg("AI_REVIEW_ENABLED", False))
    enabled = default_enabled if enabled is None else bool(enabled)
    require_confirmation = bool(
        cfg("AI_REVIEW_REQUIRE_CONFIRMATION", True) if require_confirmation is None else require_confirmation
    )
    if not enabled:
        if verbose:
            print("AI review disabled; using rule-based output only.")
        return jobs

    api_key = os.environ.get("OPENAI_API_KEY")
    print(f"OPENAI_API_KEY detected: {'YES' if api_key else 'NO'}")
    if not api_key:
        candidate_mask = jobs.apply(is_ai_candidate, axis=1)
        jobs.loc[candidate_mask, "ai_review_source"] = "skipped_no_key"
        print("OPENAI_API_KEY is not set; skipping AI review and using rule-based output only.")
        return jobs

    model = cfg("AI_MODEL", "gpt-4o-mini")
    max_jobs = int(cfg("AI_REVIEW_MAX_JOBS_PER_RUN", 5))
    max_daily = int(cfg("AI_REVIEW_MAX_DAILY_API_CALLS", 5))
    max_monthly = int(cfg("AI_REVIEW_MAX_MONTHLY_API_CALLS", 100))
    global_max_daily_api_calls = int(cfg("AI_GLOBAL_MAX_DAILY_API_CALLS", max_daily))
    global_max_monthly_api_calls = int(cfg("AI_GLOBAL_MAX_MONTHLY_API_CALLS", max_monthly))
    max_description_chars = int(cfg("AI_REVIEW_MAX_DESCRIPTION_CHARS", 4000))
    cache_enabled = bool(cfg("AI_REVIEW_CACHE_ENABLED", True))
    cache_path = project_path(cfg("AI_REVIEW_CACHE_PATH", "data/ai_review_cache.csv"))
    usage_path = project_path(cfg("AI_REVIEW_USAGE_LOG_PATH", "logs/ai_api_usage_log.csv"))

    candidate_indexes = get_ai_candidate_indexes(jobs, max_jobs)
    cache = read_table(cache_path, CACHE_COLUMNS)
    usage_log = read_table(usage_path, USAGE_LOG_COLUMNS)
    usage_rows = []
    pending = []
    cache_hits = 0
    cache_valid_rows = 0
    cache_invalid_rows = 0
    cache_invalid_requeued = 0
    cache_invalid_sent_to_api = 0
    score_normalized_rows = 0
    now = get_local_now().isoformat()
    today_text = get_local_now().date().isoformat()

    for row_index in candidate_indexes:
        job = jobs.loc[row_index]
        description = get_description(job, max_description_chars)
        job_key = canonical_job_key(job)
        desc_hash = get_description_hash(description)
        cache_match = pd.DataFrame()
        cache_source = "cache_exact"
        if cache_enabled and not cache.empty:
            cache_match = cache[
                (cache["canonical_job_key"] == job_key)
                & (cache["description_hash"] == desc_hash)
                & (cache["model_used"] == model)
            ]
            if cache_match.empty:
                cache_match = cache[
                    (cache["canonical_job_key"] == job_key)
                    & (cache["model_used"] == model)
                ]
                cache_source = "cache_similar"
        if not cache_match.empty:
            cache_row_index = cache_match.index[-1]
            cache.at[cache_row_index, "last_seen_date"] = today_text
            cached_result = clean_ai_result(cache.loc[cache_row_index].to_dict(), from_cache=True)
            if cached_result.get("ai_cache_valid", True):
                cache_valid_rows += 1
            else:
                cache_invalid_rows += 1
            if cached_result.get("ai_score_normalized", False):
                score_normalized_rows += 1
            for result_key, result_value in cached_result.items():
                if result_key in cache.columns:
                    cache.at[cache_row_index, result_key] = result_value
            if not cached_result.get("ai_cache_valid", True):
                jobs.at[row_index, "ai_review_source"] = "cache_invalid"
                jobs.at[row_index, "ai_final_action"] = "Manual Review"
                jobs.at[row_index, "ai_fit_score"] = cached_result.get("ai_fit_score", "")
                jobs.at[row_index, "ai_fit_score_raw"] = cached_result.get("ai_fit_score_raw", "")
                jobs.at[row_index, "ai_fit_score_normalized"] = cached_result.get("ai_fit_score_normalized", "")
                jobs.at[row_index, "ai_red_flags"] = cached_result.get("ai_red_flags", "AI cache invalid")
                pending.append((row_index, job_key, desc_hash, description, "cache_invalid"))
                continue
            apply_cache_result(jobs, row_index, cached_result, source=cache_source)
            propagate_review_to_duplicates(jobs, row_index, cached_result, cache_source)
            cache_hits += 1
            usage_rows.append(
                {
                    "run_datetime": now,
                    "mode": "single_job",
                    "canonical_job_key": job_key,
                    "company": job.get("company", ""),
                    "job_title": job.get("job_title", ""),
                    "model_used": model,
                    "ai_review_source": cache_source,
                    "input_chars": 0,
                    "description_chars": len(description),
                    "api_call_made": False,
                    "api_call_success": False,
                    "error_type": "",
                    "error_message": "",
                }
            )
        else:
            pending.append((row_index, job_key, desc_hash, description, "new"))

    quota_summary = usage_quota_summary(usage_log)
    daily_used = count_single_api_calls(usage_log, "day")
    monthly_used = count_single_api_calls(usage_log, "month")
    global_daily_used = count_global_api_calls(usage_log, "day")
    global_monthly_used = count_global_api_calls(usage_log, "month")
    daily_remaining = max(0, max_daily - daily_used)
    monthly_remaining = max(0, max_monthly - monthly_used)
    global_daily_remaining = max(0, global_max_daily_api_calls - global_daily_used)
    global_monthly_remaining = max(0, global_max_monthly_api_calls - global_monthly_used)
    allowed_calls = min(len(pending), daily_remaining, monthly_remaining, global_daily_remaining, global_monthly_remaining)
    pending_to_call = pending[:allowed_calls]
    cache_invalid_sent_to_api = sum(1 for item in pending_to_call if len(item) > 4 and item[4] == "cache_invalid")
    cache_invalid_requeued = max(0, cache_invalid_rows - cache_invalid_sent_to_api)

    print("AI review preflight:")
    print_usage_reset_summary(quota_summary)
    print(f"Candidate jobs: {len(candidate_indexes)}")
    print(f"Cache hits: {cache_hits}")
    print(f"Valid cache hits: {cache_valid_rows}")
    print(f"AI cache valid rows: {cache_valid_rows}")
    print(f"AI cache invalid rows: {cache_invalid_rows}")
    print(f"Invalid cache rows requeued: {cache_invalid_requeued}")
    print(f"Invalid cache rows sent to API: {cache_invalid_sent_to_api}")
    print(f"AI score normalized rows: {score_normalized_rows}")
    print("AI fallback-to-rule rows: 0")
    print(f"New API calls needed: {allowed_calls}")
    print(f"Daily limit: {daily_used}/{max_daily} used")
    print(f"Monthly limit: {monthly_used}/{max_monthly} used")
    print(f"Global daily API limit: {global_daily_used}/{global_max_daily_api_calls} used")
    print(f"Global monthly API limit: {global_monthly_used}/{global_max_monthly_api_calls} used")
    print(f"Model used: {model}")

    if allowed_calls <= 0:
        jobs.loc[jobs["ai_review_source"].eq("rule_based") & jobs.index.isin([item[0] for item in pending]), "ai_review_source"] = "skipped_quota"
        save_table(cache, cache_path, CACHE_COLUMNS)
        append_usage_rows(usage_path, usage_rows)
        return jobs

    if require_confirmation:
        answer = input(f"Proceed with {allowed_calls} new OpenAI API calls? (Y/N) ").strip().lower()
        if answer != "y":
            print("AI review skipped by user; using rule-based output only for uncached rows.")
            jobs.loc[jobs["ai_review_source"].eq("rule_based") & jobs.index.isin([item[0] for item in pending]), "ai_review_source"] = "skipped_quota"
            save_table(cache, cache_path, CACHE_COLUMNS)
            append_usage_rows(usage_path, usage_rows)
            return jobs

    for pending_item in pending_to_call:
        row_index, job_key, desc_hash, description = pending_item[:4]
        job = jobs.loc[row_index]
        prompt = build_prompt(job, description)
        api_call_made = False
        api_call_success = False
        error_type = ""
        error_message = ""
        try:
            api_call_made = True
            result = clean_ai_result(call_openai(api_key, model, prompt), from_cache=False)
            api_call_success = True
            apply_review_result(jobs, row_index, result)
            propagate_review_to_duplicates(jobs, row_index, result, "cache_similar")
            cache_row = {
                "canonical_job_key": job_key,
                "description_hash": desc_hash,
                "normalized_company": normalize_text(job.get("company", "")),
                "normalized_job_title": normalize_text(job.get("job_title", "")),
                "normalized_location": normalize_text(job.get("location", "")),
                "first_reviewed_date": today_text,
                "last_seen_date": today_text,
                "model_used": model,
                **result,
            }
            cache = pd.concat([cache, pd.DataFrame([cache_row])], ignore_index=True)
        except HTTPError as error:
            jobs.at[row_index, "ai_review_source"] = "skipped_no_key" if error.code == 401 else "skipped_quota"
            jobs.at[row_index, "ai_red_flags"] = f"AI review failed: {error}"
            error_type = f"{error.code} {error.reason}".strip()
            error_message = str(error)
            if error.code == 401:
                error_type = "401 Unauthorized"
                print(
                    "OpenAI API authentication failed. Check OPENAI_API_KEY, project permissions, billing, and model access."
                )
                usage_rows.append(
                    {
                        "run_datetime": now,
                        "mode": "single_job",
                        "canonical_job_key": job_key,
                        "company": job.get("company", ""),
                        "job_title": job.get("job_title", ""),
                        "model_used": model,
                        "ai_review_source": jobs.at[row_index, "ai_review_source"],
                        "input_chars": len(prompt),
                        "description_chars": len(description),
                        "api_call_made": True,
                        "api_call_success": False,
                        "error_type": error_type,
                        "error_message": error_message,
                    }
                )
                break
            print(f"AI review failed for {job.get('company', '')} / {job.get('job_title', '')}: {error}")
        except (HTTPError, URLError, TimeoutError, json.JSONDecodeError, KeyError) as error:
            jobs.at[row_index, "ai_review_source"] = "skipped_quota"
            jobs.at[row_index, "ai_red_flags"] = f"AI review failed: {error}"
            error_type = type(error).__name__
            error_message = str(error)
            print(f"AI review failed for {job.get('company', '')} / {job.get('job_title', '')}: {error}")
        else:
            error_type = ""
            error_message = ""
        usage_rows.append(
            {
                "run_datetime": now,
                "mode": "single_job",
                "canonical_job_key": job_key,
                "company": job.get("company", ""),
                "job_title": job.get("job_title", ""),
                "model_used": model,
                "ai_review_source": jobs.at[row_index, "ai_review_source"],
                "input_chars": len(prompt),
                "description_chars": len(description),
                "api_call_made": api_call_made,
                "api_call_success": api_call_success,
                "error_type": error_type,
                "error_message": error_message,
            }
        )

    save_table(cache, cache_path, CACHE_COLUMNS)
    append_usage_rows(usage_path, usage_rows)
    return jobs
