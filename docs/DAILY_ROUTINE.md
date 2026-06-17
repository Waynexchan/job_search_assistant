# Daily Routine

## Step 1: Gmail Status Dry-Run
```powershell
python run_email_status.py --dry-run
```

## Step 2: Review Dry-Run Output
Open:

```text
output/email_status_dry_run_review.xlsx
```

Check that suggested status updates are correct.

## Step 3: Update Status If Correct
```powershell
python run_email_status.py
```

## Step 4: Find New Jobs
Use:
- LinkedIn
- Indeed
- Company career pages
- `output/api_leads.xlsx`

## Step 5: Copy Full JDs
Copy the job link and full job description into:

```text
jobs/raw_jobs/
```

## Step 6: Run Pipeline
```powershell
python run_pipeline.py --ai-review-batch
```

## Step 7: Open Ranked Jobs
Open:

```text
output/ranked_jobs.xlsx
```

## Step 8: Apply
Apply to jobs marked:
- `Apply Today`
- `Strong Consider`

## Step 9: Record Applications
```powershell
python mark_status.py
```

## Notes
- Do not use force review daily.
- Use force review only after parser, prompt, or cache bugs.
- API leads are not final apply decisions.
- Manual full JD is the source of truth for AI review.
