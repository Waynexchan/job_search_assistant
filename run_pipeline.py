import argparse

from src.collect_api_jobs import collect_api_jobs
from src.rank_jobs import main as rank_jobs


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
    parser.set_defaults(ai_review=None)
    return parser.parse_args()


def main():
    """Collect API jobs when possible, then create the ranked Excel output."""
    args = parse_args()

    print("Starting Job Search Assistant V2 pipeline...")
    print()

    try:
        collect_api_jobs()
    except Exception as error:
        # Ranking should still work from manual pasted jobs if APIs fail.
        print(f"API collection failed, continuing with available jobs: {error}")

    print()
    rank_jobs(
        ai_review=args.ai_review,
        force_manual_review=args.force_manual_review,
        force_files=args.force_files,
    )
    print()
    print("Pipeline complete.")


if __name__ == "__main__":
    main()
