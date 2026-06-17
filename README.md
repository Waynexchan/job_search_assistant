# Job Search Assistant V2

## Overview
Job Search Assistant V2 is a local Python workflow for finding, ranking, reviewing, and tracking UK data analyst jobs. It combines API job discovery, manual full job-description review, OpenAI-assisted fit scoring, an applications tracker, and a Gmail status checker.

The project is designed for a junior or early-career UK data analyst job search where full job descriptions copied manually by the user are the most reliable source of truth.

## Key Features
- Loads manual job descriptions from `jobs/raw_jobs/*.txt`.
- Collects API job leads from Adzuna and Reed into `jobs/api_jobs.csv`.
- Scores jobs for analyst fit, location, salary seniority risk, contract risk, and junior suitability.
- Sends manual full JDs to OpenAI batch review when enabled.
- Keeps API jobs as discovery leads, not final application decisions.
- Tracks applications and highest stage in local tracker spreadsheets.
- Checks Gmail for likely application status updates through a dry-run review flow.
- Flags similar previously rejected jobs as possible reposts instead of excluding them automatically.

## Daily Workflow
1. Run Gmail status dry-run:
   `python run_email_status.py --dry-run`
2. Review `output/email_status_dry_run_review.xlsx`.
3. If correct, update tracker status:
   `python run_email_status.py`
4. Find new jobs from LinkedIn, Indeed, company career pages, and `output/api_leads.xlsx`.
5. Copy each link and full job description into `jobs/raw_jobs/`.
6. Run ranking with batch AI review:
   `python run_pipeline.py --ai-review-batch`
7. Open `output/ranked_jobs.xlsx`.
8. Apply to `Apply Today` and `Strong Consider` jobs.
9. Record applications:
   `python mark_status.py`

## Main Commands
```powershell
python run_pipeline.py --no-ai-review
python run_pipeline.py --ai-review-batch
python run_pipeline.py --ai-review-batch --force-files "file1.txt,file2.txt"
python run_pipeline.py --ai-review-batch --force-manual-review
python run_email_status.py --dry-run
python run_email_status.py
python mark_status.py
```

## Output Files
- `output/all_ranked_jobs.xlsx`: full audit output for every scored job.
- `output/ranked_jobs.xlsx`: short daily review list.
- `output/ai_review_queue.xlsx`: jobs needing manual or AI review.
- `output/api_leads.xlsx`: API discovery leads to open and inspect.
- `output/email_status_dry_run_review.xlsx`: Gmail status suggestions before tracker update.
- `output/email_status_review.xlsx`: Gmail status review output after update mode.

## Tracker Files
Tracker files live under `tracker/` and are local-only:
- `tracker/applications.xlsx`
- `tracker/status_history.xlsx`
- `tracker/applications_history.xlsx`

The tracker preserves highest-stage logic:
`offer > final_interview > interview > assessment > applied > new`.

Rejected is terminal only for the exact same job. Similar company/title matches are kept for review as possible reposts or new vacancies.

## Gmail Status Checker
`run_email_status.py` checks Gmail for likely application status updates. Use dry-run first:

```powershell
python run_email_status.py --dry-run
```

The checker must not send, archive, delete, or label emails. Review the dry-run spreadsheet before running update mode.

## OpenAI Batch Review
Manual full job descriptions are the primary OpenAI review source. Batch review is enabled with:

```powershell
python run_pipeline.py --ai-review-batch
```

Do not use force review daily. Use `--force-files` or `--force-manual-review` only after parser, prompt, or cache bugs.

## API Job Discovery
API discovery uses Adzuna and Reed where configured. API jobs are written to `jobs/api_jobs.csv` and summarized in `output/api_leads.xlsx`. API leads are not final apply decisions; open the link, inspect the live role, and copy the full JD into `jobs/raw_jobs/` before relying on AI review.

## Safety and Privacy
This project contains private job-search data. Do not commit credentials, OAuth tokens, tracker spreadsheets, generated outputs, raw job descriptions, AI caches, logs, CVs, cover letters, or personal application history.

## GitHub Setup Notes
Safe files to commit include source code, docs, `requirements.txt`, config templates without secrets, `.env.example`, and fake example files only.

Before publishing, run:

```powershell
git status
git check-ignore -v credentials.json
git check-ignore -v tracker/applications.xlsx
git check-ignore -v output/ranked_jobs.xlsx
```
