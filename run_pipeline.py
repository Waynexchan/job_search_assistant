import argparse

from pathlib import Path

from src.collect_api_jobs import collect_api_jobs
from src.rank_jobs import main as rank_jobs


PROJECT_ROOT = Path(__file__).resolve().parent
API_JOBS_FILE = PROJECT_ROOT / "jobs" / "api_jobs.csv"


def parse_args():
    parser = argparse.ArgumentParser(description="Run the job search assistant pipeline.")
    ai_group = parser.add_mutually_exclusive_group()
    ai_group.add_argument("--ai-review", action="store_true", dest="ai_review", help="Enable safe OpenAI job review.")
    ai_group.add_argument("--ai-review-batch", action="store_const", const="batch", dest="ai_review", help="Enable batch OpenAI job review.")
    ai_group.add_argument("--no-ai-review", action="store_false", dest="ai_review", help="Disable OpenAI job review.")
    parser.add_argument(
        "--force-manual-review",
        action="store_true",
        help="Force batch AI re-review for manual jobs from jobs/raw_jobs/ by ignoring their AI cache.",
    )
    parser.add_argument(
        "--force-files",
        default="",
        help='Comma-separated manual raw job filenames to force re-review, e.g. "Wasabi Sushi & Bento.txt,barclays.txt".',
    )
    parser.add_argument(
        "--force-today",
        action="store_true",
        help="Force batch AI re-review only for manual raw job files modified today.",
    )
    parser.add_argument(
        "--ai-batch-calls",
        type=int,
        default=None,
        help="Override AI batch API calls allowed this run, e.g. --ai-batch-calls 1 for smaller safer runs.",
    )
    parser.add_argument(
        "--skip-api",
        action="store_true",
        help="Skip Reed/Adzuna API collection and use existing jobs/api_jobs.csv if available.",
    )
    parser.set_defaults(ai_review=None)
    args = parser.parse_args()
    if args.force_today and args.ai_review != "batch":
        parser.error("--force-today can only be used with --ai-review-batch.")
    if args.force_today and args.force_manual_review:
        parser.error("Use either --force-today or --force-manual-review, not both.")
    if args.force_today and str(args.force_files).strip():
        parser.error("Use either --force-today or --force-files, not both.")
    return args


def main():
    """Collect API jobs when possible, then create the ranked Excel output."""
    args = parse_args()

    print("Starting Job Search Assistant V2 pipeline...")
    print()

    if args.skip_api:
        print("API collection skipped by --skip-api. Using existing jobs/api_jobs.csv.")
        if not API_JOBS_FILE.exists():
            print("No existing api_jobs.csv found. Continuing with manual jobs only.")
    else:
        try:
            collect_api_jobs()
        except Exception as error:
            # Ranking should still work from manual pasted jobs if APIs fail.
            print(f"WARNING: API collection failed, continuing with available jobs: {error}")
            if API_JOBS_FILE.exists():
                print("Preserving existing jobs/api_jobs.csv.")

    print()
    rank_jobs(
        ai_review=args.ai_review,
        force_manual_review=args.force_manual_review,
        force_files=args.force_files,
        force_today=args.force_today,
        ai_batch_calls=args.ai_batch_calls,
    )
    print()
    print("Pipeline complete.")


if __name__ == "__main__":
    main()
