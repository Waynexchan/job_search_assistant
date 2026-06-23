from pathlib import Path
import sys

import pandas as pd

PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT))

from src.email_status_assistant import classify_email, find_best_tracker_match, make_review_row


EMAIL = {
    "message_id": "sample-innocent-rejection",
    "subject": "Update on your application status for the Junior Commercial Finance Analyst role",
    "sender": "careers@example.com",
    "email_date": "2026-06-23",
    "snippet": "Unfortunately, we will not be moving forward.",
    "body": "Thanks for applying for the Junior Commercial Finance Analyst role at innocent. Unfortunately, we will not be moving forward.",
}

TRACKER = pd.DataFrame(
    [
        {
            "job_id": "sample-001",
            "apply_link": "",
            "company": "innocent",
            "job_title": "Junior Commercial Finance Analyst",
            "current_status": "applied",
            "matched_tracker_file": "applications.xlsx",
        }
    ]
)


def main():
    classification = classify_email(EMAIL)
    row_index, confidence, reasons = find_best_tracker_match(EMAIL, TRACKER, classification)
    review_row = make_review_row(EMAIL, classification, confidence, reasons, TRACKER.loc[row_index])

    assert classification["detected_status"] == "rejected"
    assert classification["extracted_job_title_body"] == "Junior Commercial Finance Analyst"
    assert classification["extracted_company_body"] == "innocent"
    assert confidence >= 85
    assert review_row["safe_to_update"] is True

    print("Sample email status extraction passed.")


if __name__ == "__main__":
    main()
