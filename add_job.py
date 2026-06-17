from datetime import date
from pathlib import Path

import pandas as pd


# These paths are relative to this project folder.
PROJECT_ROOT = Path(__file__).resolve().parent
JOBS_FILE = PROJECT_ROOT / "jobs" / "manual_jobs.csv"

# The CSV should always use these columns.
COLUMNS = ["date_added", "source", "apply_link", "status", "job_file"]


def clean_job_file_name(job_file):
    """Make the job file path simple for the CSV."""
    job_file = job_file.strip()

    # If the user only types "job_005.txt", store it inside jobs/raw_jobs/.
    if "/" not in job_file and "\\" not in job_file:
        return f"jobs/raw_jobs/{job_file}"

    # Keep typed folder paths, but use forward slashes so the CSV stays tidy.
    return job_file.replace("\\", "/")


def load_jobs():
    """Read the existing manual_jobs.csv file."""
    if JOBS_FILE.exists():
        return pd.read_csv(JOBS_FILE)

    # If the CSV does not exist yet, start with an empty table.
    return pd.DataFrame(columns=COLUMNS)


def main():
    """Ask for job details and add one new row to manual_jobs.csv."""
    print("Add a new job to jobs/manual_jobs.csv")
    print()

    source = input("Source: ").strip()
    apply_link = input("Apply Link: ").strip()
    job_file = clean_job_file_name(input("Job File Name: "))

    jobs = load_jobs()

    new_job = {
        "date_added": date.today().isoformat(),
        "source": source,
        "apply_link": apply_link,
        "status": "new",
        "job_file": job_file,
    }

    # Append the new row. This does not overwrite existing records.
    jobs = pd.concat([jobs, pd.DataFrame([new_job])], ignore_index=True)

    # Keep the columns in the expected order.
    jobs = jobs[COLUMNS]

    # Save the updated CSV.
    JOBS_FILE.parent.mkdir(parents=True, exist_ok=True)
    jobs.to_csv(JOBS_FILE, index=False)

    print()
    print("Added job:")
    print(f"Date Added: {new_job['date_added']}")
    print(f"Source: {new_job['source']}")
    print(f"Apply Link: {new_job['apply_link']}")
    print(f"Status: {new_job['status']}")
    print(f"Job File: {new_job['job_file']}")


if __name__ == "__main__":
    main()
