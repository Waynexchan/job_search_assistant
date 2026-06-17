import json
from base64 import b64encode
from datetime import date
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent
API_JOBS_FILE = PROJECT_ROOT / "jobs" / "api_jobs.csv"

SEARCH_TERMS = [
    "Junior Data Analyst",
    "Data Analyst",
    "Commercial Analyst",
    "Reporting Analyst",
    "MI Analyst",
    "BI Analyst",
    "Insight Analyst",
    "Operations Analyst",
]

LOCATIONS = [
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


def load_reed_api_key():
    """Read the Reed API key from local config.py."""
    try:
        import config
    except ImportError:
        return ""

    return getattr(config, "REED_API_KEY", "")


def make_salary(minimum, maximum):
    """Turn Reed salary fields into one easy-to-read value."""
    if minimum and maximum:
        return f"{minimum} - {maximum}"
    if minimum:
        return str(minimum)
    if maximum:
        return str(maximum)
    return ""


def fetch_reed_json(url, api_key):
    """Fetch one Reed API URL."""
    auth_token = b64encode(f"{api_key}:".encode("utf-8")).decode("utf-8")
    request = Request(
        url,
        headers={
            "Authorization": "Basic " + auth_token,
            "User-Agent": "job-search-assistant",
        },
    )

    with urlopen(request, timeout=20) as response:
        return json.loads(response.read().decode("utf-8"))


def collect_reed_jobs(api_key):
    """Collect Reed jobs for every search term and location."""
    rows = []

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
                data = fetch_reed_json(url, api_key)
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

    return pd.DataFrame(rows, columns=API_COLUMNS)


def load_existing_jobs():
    """Load the shared API job history file."""
    if API_JOBS_FILE.exists():
        jobs = pd.read_csv(API_JOBS_FILE)
    else:
        jobs = pd.DataFrame(columns=API_COLUMNS)

    for column in API_COLUMNS:
        if column not in jobs.columns:
            jobs[column] = ""

    return jobs[API_COLUMNS].fillna("")


def deduplicate_by_apply_link(jobs):
    """Keep one row per apply_link while preserving rows with blank links."""
    jobs = jobs.copy()
    jobs["apply_link"] = jobs["apply_link"].fillna("").astype(str).str.strip()

    with_link = jobs[jobs["apply_link"] != ""].drop_duplicates(
        subset=["apply_link"],
        keep="first",
    )
    without_link = jobs[jobs["apply_link"] == ""]

    return pd.concat([with_link, without_link], ignore_index=True)[API_COLUMNS]


def main():
    api_key = load_reed_api_key()

    if not api_key:
        print("REED_API_KEY is missing in config.py.")
        raise SystemExit(1)

    existing_jobs = load_existing_jobs()
    collected_jobs = collect_reed_jobs(api_key)
    merged_jobs = deduplicate_by_apply_link(
        pd.concat([existing_jobs, collected_jobs], ignore_index=True)
    )

    API_JOBS_FILE.parent.mkdir(parents=True, exist_ok=True)
    merged_jobs.to_csv(API_JOBS_FILE, index=False)

    new_jobs_added = len(merged_jobs) - len(deduplicate_by_apply_link(existing_jobs))

    print(f"Total jobs collected: {len(collected_jobs)}")
    print(f"New jobs added: {new_jobs_added}")
    print(f"Saved API jobs to: {API_JOBS_FILE}")


if __name__ == "__main__":
    main()
