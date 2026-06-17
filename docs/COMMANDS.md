# Commands

## Pipeline
```powershell
python run_pipeline.py --no-ai-review
python run_pipeline.py --ai-review-batch
python run_pipeline.py --ai-review-batch --force-files "file1.txt,file2.txt"
python run_pipeline.py --ai-review-batch --force-manual-review
```

## Gmail Status Checker
```powershell
python run_email_status.py --dry-run
python run_email_status.py
```

## Tracker
```powershell
python mark_status.py
```

## Useful Check Commands
Check ranked jobs output:

```powershell
python -c "import pandas as pd; df=pd.read_excel('output/ranked_jobs.xlsx'); print(df[['company','job_title','final_action','next_action']].head(20).to_string(index=False))"
```

Check manual jobs in full audit output:

```powershell
python -c "import pandas as pd; df=pd.read_excel('output/all_ranked_jobs.xlsx'); print(df[df['input_type'].eq('manual')][['file_name','company','job_title','final_action','would_send_to_ai']].to_string(index=False))"
```

Check email review output:

```powershell
python -c "import pandas as pd; df=pd.read_excel('output/email_status_dry_run_review.xlsx'); print(df.head(20).to_string(index=False))"
```
