import json
from base64 import b64encode
from datetime import date
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import quote_plus, urlencode
from urllib.request import Request, urlopen

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[1]
API_JOBS_FILE = PROJECT_ROOT / "jobs" / "api_jobs.csv"

SEARCH_TERMS = [
    "Junior Data Analyst",
    "Data Analyst",
    "Commercial Analyst",
    "MI Analyst",
    "Reporting Analyst",
    "BI Analyst",
    "Insight Analyst",
    "Operations Analyst",
]

LOCATIONS = [
    "Milton Keynes",
    "London",
    "Birmingham",
    "Luton",
    "Bedford",
    "Northampton",
    "Remote",
    "Hybrid",
]

REED_LOCATIONS = [
    "Milton Keynes",
    "London",
    "Birmingham",
    "Northampton",
    "Bedford",
    "Luton",
    "Remote",
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
    "job_description",
    "apply_link",
]


def load_config():
    """Read local API keys from config.py if it exists."""
    try:
        import config
    except ImportError:
        return "", "", ""

    return (
        getattr(config, "ADZUNA_APP_ID", ""),
        getattr(config, "ADZUNA_API_KEY", ""),
        getattr(config, "REED_API_KEY", ""),
    )


def fetch_json(url, headers=None):
    """Fetch one JSON URL and return a Python dictionary."""
    request = Request(url, headers=headers or {"User-Agent": "job-search-assistant"})

    with urlopen(request, timeout=20) as response:
        return json.loads(response.read().decode("utf-8"))


def make_salary(minimum, maximum):
    """Format salary fields from API responses."""
    if pd.isna(minimum) and pd.isna(maximum):
        return ""
    if minimum and maximum:
        return f"{minimum} - {maximum}"
    if minimum:
        return str(minimum)
    if maximum:
        return str(maximum)
    return ""


def collect_adzuna_jobs(app_id, api_key):
    """Collect jobs from Adzuna for all search terms and locations."""
    if not app_id or not api_key:
        print("Skipping Adzuna: ADZUNA_APP_ID or ADZUNA_API_KEY is missing.")
        return []

    rows = []

    for search_term in SEARCH_TERMS:
        for location in REED_LOCATIONS:
            params = urlencode(
                {
                    "app_id": app_id,
                    "app_key": api_key,
                    "what": search_term,
                    "where": location,
                    "results_per_page": 20,
                    "content-type": "application/json",
                }
            )
            url = f"https://api.adzuna.com/v1/api/jobs/gb/search/1?{params}"

            try:
                data = fetch_json(url)
            except (HTTPError, URLError, TimeoutError) as error:
                print(f"Adzuna request failed for {search_term} / {location}: {error}")
                continue

            for job in data.get("results", []):
                company = job.get("company", {}) or {}
                rows.append(
                    {
                        "date_collected": date.today().isoformat(),
                        "source": "Adzuna",
                        "job_title": job.get("title", ""),
                        "company": company.get("display_name", ""),
                        "location": (job.get("location", {}) or {}).get("display_name", location),
                        "salary": make_salary(job.get("salary_min"), job.get("salary_max")),
                        "job_posted_date": job.get("created", ""),
                        "application_deadline": job.get("valid_until", ""),
                        "job_description": job.get("description", ""),
                        "apply_link": job.get("redirect_url", ""),
                    }
                )

    return rows


def collect_reed_jobs(api_key):
    """Collect jobs from Reed for all search terms and locations."""
    if not api_key:
        print("Skipping Reed: REED_API_KEY is missing.")
        return []

    rows = []
    auth_token = b64encode(f"{api_key}:".encode("utf-8")).decode("utf-8")
    headers = {"Authorization": "Basic " + auth_token, "User-Agent": "job-search-assistant"}

    for search_term in SEARCH_TERMS:
        for location in LOCATIONS:
            params = urlencode(
                {
                    "keywords": search_term,
                    "locationName": location,
                    "resultsToTake": 20,
                }
            )
            url = f"https://www.reed.co.uk/api/1.0/search?{params}"

            try:
                data = fetch_json(url, headers=headers)
            except (HTTPError, URLError, TimeoutError) as error:
                print(f"Reed request failed for {search_term} / {location}: {error}")
                continue

            for job in data.get("results", []):
                rows.append(
                    {
                        "date_collected": date.today().isoformat(),
                        "source": "Reed",
                        "job_title": job.get("jobTitle", ""),
                        "company": job.get("employerName", ""),
                        "location": job.get("locationName", location),
                        "salary": make_salary(job.get("minimumSalary"), job.get("maximumSalary")),
                        "job_posted_date": job.get("datePosted", "") or job.get("date", ""),
                        "application_deadline": job.get("expirationDate", ""),
                        "job_description": job.get("jobDescription", ""),
                        "apply_link": job.get("jobUrl", ""),
                    }
                )

    return rows


def load_existing_api_jobs():
    """Read previous API jobs so new runs do not erase old results."""
    if API_JOBS_FILE.exists():
        jobs = pd.read_csv(API_JOBS_FILE)
    else:
        jobs = pd.DataFrame(columns=API_COLUMNS)

    for column in API_COLUMNS:
        if column not in jobs.columns:
            jobs[column] = ""

    return jobs[API_COLUMNS]


def deduplicate_jobs(jobs):
    """Remove duplicate jobs by URL first, then by company/title/location."""
    jobs = jobs.copy()
    jobs["apply_link"] = jobs["apply_link"].fillna("").astype(str).str.strip()

    # First remove exact URL duplicates. Blank links need separate handling so
    # missing URLs do not all collapse into one row.
    with_link = jobs[jobs["apply_link"] != ""].drop_duplicates(
        subset=["apply_link"],
        keep="first",
    )
    without_link = jobs[jobs["apply_link"] == ""]
    jobs = pd.concat([with_link, without_link], ignore_index=True)

    # Then remove repeated listings that have different redirect URLs but the
    # same company, title and location.
    jobs = jobs.drop_duplicates(
        subset=["company", "job_title", "location"],
        keep="first",
    )

    return jobs.reset_index(drop=True)[API_COLUMNS]


def collect_api_jobs():
    """Collect API jobs and save the merged, deduplicated CSV."""
    adzuna_app_id, adzuna_api_key, reed_api_key = load_config()
    adzuna_rows = collect_adzuna_jobs(adzuna_app_id, adzuna_api_key)
    reed_rows = collect_reed_jobs(reed_api_key)
    new_rows = adzuna_rows + reed_rows

    existing_jobs = load_existing_api_jobs()
    existing_count = len(deduplicate_jobs(existing_jobs))
    new_jobs = pd.DataFrame(new_rows, columns=API_COLUMNS)
    merged_jobs = deduplicate_jobs(pd.concat([existing_jobs, new_jobs], ignore_index=True))
    new_jobs_added = len(merged_jobs) - existing_count

    API_JOBS_FILE.parent.mkdir(parents=True, exist_ok=True)
    merged_jobs.to_csv(API_JOBS_FILE, index=False)

    print(f"Total Adzuna jobs collected: {len(adzuna_rows)}")
    print(f"Total Reed jobs collected: {len(reed_rows)}")
    print(f"Total API jobs collected this run: {len(new_jobs)}")
    print(f"New jobs added after deduplication: {new_jobs_added}")
    print(f"Saved {len(merged_jobs)} deduplicated API jobs to: {API_JOBS_FILE}")
    return merged_jobs


def main():
    collect_api_jobs()


if __name__ == "__main__":
    main()
