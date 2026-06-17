# App Overview

## API Discovery
API discovery pulls job leads from:
- Adzuna
- Reed

Collected API rows are cached in `jobs/api_jobs.csv` and surfaced for manual inspection in `output/api_leads.xlsx`.

API leads are discovery inputs only. They should be opened, checked, and copied into `jobs/raw_jobs/` with the full job description before final AI review or application decisions.

## Manual Full JD Workflow
Manual jobs are stored as text files in `jobs/raw_jobs/*.txt`.

The parser extracts:
- Link
- Source
- Title
- Company
- Location
- Salary
- Contract type
- Posted date when explicit
- Full job description

Manual full JDs are the primary AI review source because they contain richer evidence than API snippets.

## Ranking And Filtering
The main pipeline is `run_pipeline.py`.

Core ranking logic lives in `src/rank_jobs.py` and writes:
- `output/all_ranked_jobs.xlsx`
- `output/ranked_jobs.xlsx`
- `output/ai_review_queue.xlsx`
- `output/api_leads.xlsx`

The pipeline scores role fit, junior suitability, location, salary risk, contract risk, technical signals, and business-domain fit.

Tracker filtering now separates exact matches from similar history. Exact same jobs with terminal tracker statuses are excluded. Similar previously rejected roles are kept and flagged as possible reposts or new vacancies.

## OpenAI Review
OpenAI review logic lives in `src/ai_job_reviewer.py`.

Batch mode reviews manual full JDs, uses a local cache, and records API usage in logs. Force review options exist for parser, prompt, or cache fixes:

```powershell
python run_pipeline.py --ai-review-batch --force-files "file1.txt,file2.txt"
python run_pipeline.py --ai-review-batch --force-manual-review
```

Do not use force review as a daily default.

## Tracker
Tracker files are local-only:
- `tracker/applications.xlsx`
- `tracker/status_history.xlsx`
- `tracker/applications_history.xlsx`

Application status is managed through `mark_status.py`.

Highest-stage logic must not be downgraded:
`offer > final_interview > interview > assessment > applied > new`.

Rejected, withdrawn, expired, and closed block only exact same jobs, not similar reposted vacancies.

## Gmail Status Checker
Gmail status checking is run with `run_email_status.py` and implemented in `src/email_status_assistant.py`.

Outputs:
- `output/email_status_dry_run_review.xlsx`
- `output/email_status_review.xlsx`

Always run dry-run first and review the spreadsheet before update mode:

```powershell
python run_email_status.py --dry-run
python run_email_status.py
```
