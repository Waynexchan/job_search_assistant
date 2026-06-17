from pathlib import Path
import argparse
import sys
import traceback

from src.email_status_assistant import CREDENTIALS_FILE, process_emails


def safe_text(value):
    """Make terminal output safe on Windows consoles with limited encodings."""
    encoding = sys.stdout.encoding or "utf-8"
    return str(value).encode(encoding, errors="replace").decode(encoding, errors="replace")


def main():
    parser = argparse.ArgumentParser(description="Read Gmail application updates and match them to tracker rows.")
    parser.add_argument("--dry-run", action="store_true", help="Match emails without updating tracker files.")
    parser.add_argument("--days", type=int, default=30, help="Only search Gmail newer_than this many days.")
    parser.add_argument("--max-emails", type=int, default=20, help="Maximum number of emails to read.")
    args = parser.parse_args()

    print("Checking credentials.json")
    if not CREDENTIALS_FILE.exists():
        print(f"Missing {CREDENTIALS_FILE}")
        print("Add Gmail OAuth credentials.json before running the Email Assistant.")
        return

    try:
        summary = process_emails(
            dry_run=args.dry_run,
            days=args.days,
            max_emails=args.max_emails,
            verbose=True,
        )
    except Exception as error:
        print("Email Status Assistant stopped because an error occurred.")
        print(f"Error type: {type(error).__name__}")
        print(f"Error message: {error}")
        print("Full traceback:")
        print(traceback.format_exc())
        raise SystemExit(1)

    print("Email Status Assistant summary:")
    print(f"Dry run: {summary['dry_run']}")
    print(f"Emails checked: {summary['emails_checked']}")
    print(f"Ignored: {summary['counts']['ignored']}")
    print(f"Rejected detected: {summary['counts']['rejected']}")
    print(f"Assessment detected: {summary['counts']['assessment']}")
    print(f"Interview detected: {summary['counts']['interview'] + summary['counts']['final_interview']}")
    print(f"Offer detected: {summary['counts']['offer']}")
    print(f"Applied acknowledgements detected: {summary['counts'].get('applied_acknowledgement', 0)}")
    print(f"Tracker updates possible: {summary['counts']['tracker_updates_possible']}")
    print(f"Manual review required: {summary['counts']['manual_review_required']}")
    print(f"Tracker updates: {summary['tracker_updates']}")
    print(f"Manual review rows: {summary['manual_review']}")
    print(f"Review file: {Path(summary['review_file'])}")

    if summary["review_rows"]:
        print()
        print("Review rows:")
        for index, row in enumerate(summary["review_rows"], start=1):
            print(
                safe_text(
                    f"{index}. {row['detected_status']} | "
                    f"{row['match_confidence']} | "
                    f"{row['matched_company'] or 'No company'} | "
                    f"{row.get('matched_tracker_file', 'none')} | "
                    f"{row['email_subject']}"
                )
            )
            if row.get("ignore_reason"):
                print(safe_text(f"   ignore_reason: {row['ignore_reason']}"))
            if row.get("classification_evidence"):
                print(safe_text(f"   evidence: {row['classification_evidence']}"))
            print(safe_text(f"   reason: {row['match_reason']}"))
            print(safe_text(f"   safe_to_update: {row.get('safe_to_update', False)}"))
            print(safe_text(f"   suggested_action: {row['suggested_action']}"))

    if summary["dry_run"]:
        print()
        print("Dry run mode: tracker/applications.xlsx and tracker/status_history.xlsx were not changed.")


if __name__ == "__main__":
    main()
