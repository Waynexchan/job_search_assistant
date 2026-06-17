# GitHub Safety Checklist

## Must Not Commit
- `credentials.json`
- `token.json`
- `.env`
- Any OpenAI API key
- Gmail OAuth credentials
- Gmail OAuth token
- `tracker/*.xlsx`
- `tracker/backups/`
- `output/*.xlsx`
- `jobs/raw_jobs/*.txt`
- `jobs/api_jobs.csv`
- `data/ai_review_cache.csv`
- `logs/`
- `backups/`
- `__pycache__/`
- `.venv/`
- `cvs/`
- `cover_letters/`
- Any personal CV or cover letter files containing address, phone, email, visa details, or other personal data
- Any personal job application history

## Safe To Commit
- Source code files
- Config template files without secrets
- `README.md`
- `docs/*.md`
- `requirements.txt`
- `.env.example`
- Example files with fake data only
- Sample raw job txt with fake company/job only

## Before Upload
Run:

```powershell
git status
git check-ignore -v credentials.json
git check-ignore -v token.json
git check-ignore -v tracker/applications.xlsx
git check-ignore -v output/ranked_jobs.xlsx
git check-ignore -v jobs/raw_jobs/sample.txt
```

If private files are already tracked, remove them from Git tracking without deleting local files:

```powershell
git rm --cached <file>
```

Do not run `git add .` until `git status` is clean of private/generated files.
