from pathlib import Path
import argparse
import sys
import traceback

from src.email_status_assistant import CREDENTIALS_FILE, process_emails

try:
    from google.auth.exceptions import RefreshError
except ImportError:  # pragma: no cover - only absent when Gmail packages are not installed.
    RefreshError = None


PROJECT_ROOT = Path(__file__).resolve().parent


def safe_text(value):
    """Make terminal output safe on Windows consoles with limited encodings."""
    encoding = sys.stdout.encoding or "utf-8"
    return str(value).encode(encoding, errors="replace").decode(encoding, errors="replace")


def is_invalid_grant_error(error):
    if RefreshError is not None and not isinstance(error, RefreshError):
        return False
    message = str(error).lower()
    return "invalid_grant" in message or "expired or revoked" in message


def print_invalid_grant_help():
    print("Your Gmail token has expired or been revoked. Please delete token.json and rerun the script to re-authorise Gmail access.")
    print("Suggested PowerShell commands:")
    print(f'cd "{PROJECT_ROOT}"')
    print("rename-item token.json token_old_revoked.json")
    print("python run_email_status.py --dry-run")


def main():
    parser = argparse.ArgumentParser(description="Read Gmail application updates and match them to tracker rows.")
    parser.add_argument("--dry-run", action="store_true", help="Match emails without updating tracker files.")
    parser.add_argument("--apply", action="store_true", help="Apply safe tracker updates. Without this flag, the command runs as dry-run.")
    parser.add_argument("--days", type=int, default=30, help="Only search Gmail newer_than this many days.")
    parser.add_argument("--max-emails", type=int, default=20, help="Maximum number of emails to read.")
    args = parser.parse_args()
    if args.dry_run and args.apply:
        parser.error("Use either --dry-run or --apply, not both.")
    dry_run = not args.apply

    print("Checking credentials.json")
    if not CREDENTIALS_FILE.exists():
        print(f"Missing {CREDENTIALS_FILE}")
        print("Add Gmail OAuth credentials.json before running the Email Assistant.")
        return

    try:
        summary = process_emails(
            dry_run=dry_run,
            days=args.days,
            max_emails=args.max_emails,
            verbose=True,
        )
    except Exception as error:
        if is_invalid_grant_error(error):
            print_invalid_grant_help()
            raise SystemExit(1)
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
    print(f"High-confidence tracker matches: {summary['counts'].get('high_confidence_tracker_matches', 0)}")
    print(f"Tracker updates possible: {summary['counts']['tracker_updates_possible']}")
    if summary["dry_run"]:
        print(f"Tracker updates that would be made: {summary['counts']['tracker_updates_possible']}")
    print(f"Manual review required: {summary['counts']['manual_review_required']}")
    print(f"Tracker updates: {summary['tracker_updates']}")
    print(f"Manual review rows: {summary['manual_review']}")
    print(f"Review file: {Path(summary['review_file'])}")
    if not summary["dry_run"] and summary.get("applied_update_rows"):
        print(f"Applied update audit file: {Path(summary['applied_updates_file'])}")

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
        print("Dry-run only. No tracker files were modified. Review output/email_status_dry_run_review.xlsx. To apply safe updates, run: python run_email_status.py --apply")
    elif summary.get("applied_update_rows"):
        print()
        print("Tracker updates applied:")
        for index, row in enumerate(summary["applied_update_rows"], start=1):
            print(
                safe_text(
                    f"{index}. {row['email_date']} | {row['detected_status']} | "
                    f"{row['matched_tracker_company']} | {row['matched_tracker_job_title']} | "
                    f"{row['old_status']} -> {row['new_status']} | "
                    f"{row['matched_tracker_file']} | {row['match_confidence']} | "
                    f"{row['email_subject']}"
                )
            )
            print(safe_text(f"   match_explanation: {row['match_explanation']}"))
    else:
        print()
        print("No safe tracker updates were applied.")


if __name__ == "__main__":
    main()
