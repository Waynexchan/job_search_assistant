from datetime import date
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parent
INPUT_FILE = PROJECT_ROOT / "tracker" / "legacy_applications.xlsx"
OUTPUT_FILE = PROJECT_ROOT / "tracker" / "applications_history.xlsx"

OUTPUT_COLUMNS = [
    "Company",
    "Job Title",
    "Apply Link",
    "Platform",
    "Date Applied",
    "Status",
    "record_source",
    "created_date",
]

STATUS_VALUES = ["Applied", "Rejected", "Interview", "Assessment", "Offer", "Unknown"]


def clean_text(value):
    """Make spreadsheet values safe to compare and save."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def standardize_status(status):
    """Convert messy historical statuses into the approved status values."""
    status_text = clean_text(status).lower()

    if not status_text:
        return "Unknown"

    if "reject" in status_text:
        return "Rejected"
    if "interview" in status_text:
        return "Interview"
    if "assessment" in status_text or "test" in status_text:
        return "Assessment"
    if "offer" in status_text:
        return "Offer"
    if "applied" in status_text or status_text == "apply":
        return "Applied"

    return "Unknown"


def first_existing_column(row, column_names):
    """Read the first available column from a row."""
    for column in column_names:
        if column in row:
            value = clean_text(row[column])
            if value:
                return value
    return ""


def normalize_history_rows(jobs, record_source):
    """Map source spreadsheet rows into the applications history schema."""
    rows = []
    import_date = date.today().isoformat()

    for _, row in jobs.fillna("").iterrows():
        rows.append(
            {
                "Company": first_existing_column(row, ["Company", "company"]),
                "Job Title": first_existing_column(row, ["Job Title", "job_title"]),
                "Apply Link": first_existing_column(row, ["Apply Link", "Link", "apply_link"]),
                "Platform": first_existing_column(row, ["Platform", "source"]),
                "Date Applied": first_existing_column(row, ["Date Applied", "apply_date"]),
                "Status": standardize_status(
                    first_existing_column(
                        row,
                        ["Status", "status", "狀態（Save / Apply / Skip）"],
                    )
                ),
                "record_source": first_existing_column(row, ["record_source"]) or record_source,
                "created_date": first_existing_column(row, ["created_date"]) or import_date,
            }
        )

    return pd.DataFrame(rows, columns=OUTPUT_COLUMNS)


def make_dedupe_key(row):
    """Prefer Apply Link for duplicates, then fallback to company/title/date."""
    apply_link = clean_text(row.get("Apply Link", "")).lower()

    if apply_link:
        return f"link|{apply_link}"

    company = clean_text(row.get("Company", "")).lower()
    job_title = clean_text(row.get("Job Title", "")).lower()
    date_applied = clean_text(row.get("Date Applied", "")).lower()
    return f"fallback|{company}|{job_title}|{date_applied}"


def load_existing_history():
    """Load existing history without dropping old rows."""
    if not OUTPUT_FILE.exists():
        return pd.DataFrame(columns=OUTPUT_COLUMNS)

    existing = pd.read_excel(OUTPUT_FILE)
    return normalize_history_rows(existing, "existing_history")


def load_legacy_applications():
    """Load tracker/legacy_applications.xlsx."""
    if not INPUT_FILE.exists():
        raise FileNotFoundError(f"Missing input file: {INPUT_FILE}")

    legacy = pd.read_excel(INPUT_FILE)
    return normalize_history_rows(legacy, "historical_import")


def import_history():
    """Import legacy applications into applications_history.xlsx."""
    existing_history = load_existing_history()
    legacy_history = load_legacy_applications()

    existing_keys = set(existing_history.apply(make_dedupe_key, axis=1))
    legacy_history["dedupe_key"] = legacy_history.apply(make_dedupe_key, axis=1)
    new_rows = legacy_history[~legacy_history["dedupe_key"].isin(existing_keys)].copy()
    new_rows = new_rows.drop_duplicates(subset=["dedupe_key"], keep="first")
    new_rows = new_rows[OUTPUT_COLUMNS]

    combined_history = pd.concat([existing_history, new_rows], ignore_index=True)
    combined_history = combined_history[OUTPUT_COLUMNS]

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)
    combined_history.to_excel(OUTPUT_FILE, index=False)

    return combined_history, len(new_rows)


def print_summary(history, new_rows_added):
    """Print validation counts requested by the user."""
    status_counts = history["Status"].value_counts()

    print("Import summary:")
    print(f"total rows: {len(history)}")
    print(f"new rows added: {new_rows_added}")
    print(f"rejected: {int(status_counts.get('Rejected', 0))}")
    print(f"interview: {int(status_counts.get('Interview', 0))}")
    print(f"offer: {int(status_counts.get('Offer', 0))}")
    print(f"applied: {int(status_counts.get('Applied', 0))}")


def main():
    try:
        history, new_rows_added = import_history()
    except FileNotFoundError as error:
        print(error)
        raise SystemExit(1)

    print_summary(history, new_rows_added)


if __name__ == "__main__":
    main()
