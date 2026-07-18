# Project Logic

## Project Purpose

This is a local-first Job Search Assistant V2 for UK Junior Data Analyst, BI, Reporting Analyst, and Commercial Analyst applications. It collects API job leads, parses manual full job descriptions, ranks jobs, reviews manual jobs with OpenAI, tracks applications, and checks Gmail for application status updates.

The project is intentionally local-first: tracker files, raw job descriptions, OAuth token files, AI cache files, usage logs, and generated review outputs live in this workspace and should not be deleted or rewritten unless the user explicitly asks.

## Main Folders and Files

- `run_pipeline.py`: main job collection, ranking, and AI review entry point.
- `run_email_status.py`: Gmail status check entry point. Use `--dry-run` before live updates.
- `mark_status.py`: updates application tracker status after the user applies or records a manual status change.
- `config.py`: local configuration and API keys/settings. This file may be untracked/local and should not be exposed.
- `src/rank_jobs.py`: core job parsing, scoring, tracker overlay, ranking output, and AI queue generation logic.
- `src/collect_api_jobs.py`: Reed/Adzuna API lead collection.
- `src/ai_job_reviewer.py`: OpenAI review, batch review, AI cache, usage limits, and AI review source handling.
- `src/email_status_assistant.py`: Gmail read-only fetch, status classification, tracker matching, dry-run output, and safe tracker update logic.
- `jobs/raw_jobs/`: manual full job description `.txt` files. These are the main AI review inputs.
- `jobs/applied_jobs/`: archived manual raw job files after `mark_status.py` marks a job as applied.
- `jobs/archived_jobs/`: optional archive target for withdrawn, not interested, or skipped jobs when enabled.
- `jobs/rejected_jobs/`: archive target for applied jobs later marked rejected, and optional target for rejected raw jobs when enabled.
- `jobs/raw_jobs/_invalid/`: holding area for invalid raw job files if manual cleanup is needed.
- `jobs/api_jobs.csv`: deduplicated API discovery leads from Reed/Adzuna.
- `output/ranked_jobs.xlsx`: main daily action list after ranking and AI review filtering.
- `output/ranked_jobs_exclusion_audit.xlsx`: audit of manual `Apply Today`, `Strong Consider`, or `Apply If Time` rows that were not written to `ranked_jobs.xlsx`.
- `output/quick_apply_jobs.xlsx`: optional volume list for low-friction LinkedIn Easy Apply, Indeed Apply, Easy Apply, or quick apply manual jobs.
- `output/application_pack_recommendations.xlsx`: ranked-action jobs with the recommended ATS CV, ATS cover letter, category, and local file-existence readiness check.
- `output/raw_job_trace_audit.xlsx`: one-row-per-manual-job trace showing parsing, scoring, AI/fallback status, export destination, and why a row was not in `ranked_jobs.xlsx`.
- `output/filtered_jobs.xlsx`: valid manual jobs that are not ranked, queued, duplicates, or parsing errors.
- `output/duplicates.xlsx`: duplicate manual/API rows removed from the daily ranked view, with duplicate evidence.
- `output/parsing_errors.xlsx`: genuinely unreadable manual raw records. Non-fatal missing salary/date/location/URL issues should stay in normal outputs with parser warnings.
- `output/ai_review_queue.xlsx`: manual review, borderline, and pending AI review list.
- `output/api_leads.xlsx`: API discovery leads only. These are not full JD AI review targets by default.
- `output/all_ranked_jobs.xlsx`: full audit output with every ranked row.
- `output/email_status_review.xlsx`: live Gmail status review/update output.
- `output/email_status_dry_run_review.xlsx`: dry-run Gmail status review output.
- `output/email_status_updates_applied.xlsx`: live-mode audit of tracker updates applied by the Gmail assistant.
- `tracker/applications.xlsx`: current active application tracker.
- `tracker/applications_history.xlsx`: historical application records used for matching and duplicate prevention.
- `tracker/status_history.xlsx`: status change history.
- `tracker/backups/`: backups created before tracker write operations.
- `data/ai_review_cache.csv`: cached AI review results to avoid repeated OpenAI costs.
- `logs/ai_api_usage_log.csv`: OpenAI usage log and quota accounting.
- `output/raw_job_file_movement_log.xlsx`: audit log for raw job lifecycle moves.
- `credentials.json`: Gmail OAuth client configuration.
- `token.json`: local Gmail OAuth token. If revoked/expired, rename it and re-authorise.

## Daily Workflow

### A. Email Status Check

Run dry-run first:

```powershell
python run_email_status.py
```

or explicitly:

```powershell
python run_email_status.py --dry-run
```

Review:

```text
output/email_status_dry_run_review.xlsx
```

If the matches and `safe_to_update` values are safe:

```powershell
python run_email_status.py --apply
```

### B. Job Ranking

Put full job descriptions into:

```text
jobs/raw_jobs/
```

Then run:

```powershell
python run_pipeline.py --ai-review-batch
```

For manual-only AI review, skip Reed/Adzuna collection:

```powershell
python run_pipeline.py --ai-review-batch --force-today --skip-api
```

### C. After Applying

```powershell
python mark_status.py
```

### D. Debug Only

```powershell
python run_pipeline.py --no-ai-review
```

Warning: `--no-ai-review` is diagnostic only. It should not be treated as the daily production ranking command because manual full JD jobs need AI review coverage.

## Job Pipeline Logic

- API jobs from Reed and Adzuna are discovery leads only.
- API jobs are saved into `jobs/api_jobs.csv` and `output/api_leads.xlsx`.
- API jobs should not consume AI review capacity when `AI_BATCH_INCLUDE_API_LEADS=False`.
- `--skip-api` skips Reed/Adzuna collection entirely and uses the existing `jobs/api_jobs.csv` if present. If it is missing, the pipeline continues with manual raw jobs only.
- If Reed/Adzuna returns HTTP 503 or another transient API error, the pipeline warns and continues with available data. Existing `jobs/api_jobs.csv` should be preserved when no new API jobs are collected.
- Manual full JD files in `jobs/raw_jobs/` are the main source for AI review and are assumed to have already been human-screened by the user.
- Active manual AI review scans `jobs/raw_jobs/*.txt` only. It does not scan `jobs/applied_jobs/`, `jobs/rejected_jobs/`, or `jobs/archived_jobs/` by default.
- Manual jobs should normally reach AI review unless there is an exact tracker exclusion, a parse failure/empty unreadable file, an exact duplicate raw file, a contract/FTC/temp/day-rate role when permanent-only is preferred, a senior leadership title, a clearly non-target role, or an already valid cached AI result.
- Manual jobs should not be hard skipped before AI only because of 3+ years wording, 2-4 years wording, mid-level wording, high salary, far location, hybrid uncertainty, domain mismatch, financial services/consulting/gaming/transport/healthcare domain, missing tools, or a stretch SQL/Python/Power BI requirement. Those concerns should be handled by AI scoring.
- Manual raw job source detection recognises common pasted job links including LinkedIn, Indeed, Reed, Civil Service Jobs, Totaljobs, Glassdoor, Adzuna, CORD (`cord.co`), and Welcome to the Jungle (`welcometothejungle.com`). A raw file can also specify `Source: CORD`, `Source: Welcome to the Jungle`, or `Source: WTTJ`.
- `data/ai_review_cache.csv` is used to avoid repeated OpenAI API costs.

## Raw Job Lifecycle Archive Logic

Raw job files should stay focused on active, unapplied manual jobs.

When `python mark_status.py` successfully updates a job to `applied`, the matching `.txt` file is moved from:

```text
jobs/raw_jobs/
```

to:

```text
jobs/applied_jobs/
```

The moved file is renamed as:

```text
YYYY-MM-DD__company__job_title__original_filename.txt
```

Unsafe filename characters are sanitized. If the target filename already exists, `__2`, `__3`, and so on are appended. The movement is logged in:

```text
output/raw_job_file_movement_log.xlsx
```

Tracker rows include raw job path fields when written by `mark_status.py`:

- `raw_job_file_path`
- `archived_raw_job_path`

Safety rules:

- Raw files are moved only after the tracker status save succeeds.
- Raw files are not moved during Gmail dry-run.
- Raw files are not moved during AI ranking.
- Raw files are never deleted by lifecycle logic.
- AI cache is not modified by lifecycle logic.
- OpenAI is not called by lifecycle logic.

Rejected and archived movement:

- If a job already archived in `jobs/applied_jobs/` is later marked `rejected`, it can move to `jobs/rejected_jobs/`.
- Raw jobs still in `jobs/raw_jobs/` are not moved to `jobs/rejected_jobs/` unless `MOVE_REJECTED_RAW_JOBS=True`.
- `withdrawn`, `not_interested`, and `skip` jobs are not moved to `jobs/archived_jobs/` unless `MOVE_ARCHIVED_RAW_JOBS=True`.
- Both settings default to `False`.

## Manual AI Review Coverage Logic

Every manual full JD should eventually have either:

- a valid AI result,
- a tracker exact exclusion,
- a strict pre-AI exclusion,
- or a parse failure/empty unreadable file.

A manual job is not fully reviewed if:

- `ai_review_source = rule_based`
- `final_action = Pending AI Review`
- `ai_fit_score` is blank
- `would_send_to_ai = True`

Valid AI review sources are:

- `cache_exact`
- `cache_similar`
- `openai_new`

Strict pre-AI exclusion statuses include:

- `Exact duplicate`
- `Empty or unreadable`
- `Contract / FTC skipped`
- `Exact tracker exclusion`
- `Senior leadership skipped`
- `Non-target role skipped`

Strictly excluded jobs are not counted as `Pending AI review`. If any eligible manual job remains `rule_based` or `Pending AI Review` after `--ai-review-batch`, this is a warning and should be checked in:

```text
output/manual_ai_coverage_audit.xlsx
```

The pre-AI exclusion audit explains which manual jobs did not reach AI and why:

```text
output/manual_pre_ai_exclusion_audit.xlsx
```

## AI Ranking Outputs

- `ranked_jobs.xlsx`: main daily action list.
- `ranked_jobs_exclusion_audit.xlsx`: explains why any manual `Apply Today`, `Strong Consider`, or `Apply If Time` job was not written to `ranked_jobs.xlsx`.
- `quick_apply_jobs.xlsx`: optional extra list for low-friction applications. It is for application volume, not the main priority list.
- `raw_job_trace_audit.xlsx`: source-of-truth trace for every manually pasted active raw job. It records raw file path, encoding, detected job blocks, parsed fields, active status, duplicate/hard-filter status, deterministic scores, AI/fallback status, recommendation, export destination, and why the row was not included in `ranked_jobs.xlsx`.
- `filtered_jobs.xlsx`: valid manual jobs that are visible but not priority-ranked or queued.
- `duplicates.xlsx`: rows treated as duplicates; deduplication must never silently delete a row.
- `parsing_errors.xlsx`: records that could not be read/parsed. Missing salary, dates, URL, or uncertain location are non-fatal parser warnings, not parsing errors.
- `ai_review_queue.xlsx`: manual review, borderline, and pending review list.
- `api_leads.xlsx`: API discovery leads only.
- `all_ranked_jobs.xlsx`: full audit file.
- `manual_ai_coverage_audit.xlsx`: one-row-per-manual-file AI coverage audit.
- `manual_pre_ai_exclusion_audit.xlsx`: strict pre-AI exclusion audit for manual files.
- `pipeline_debug_audit.xlsx`: multi-sheet debug workbook reconciling manual jobs, ranked output inclusion, AI review queue inclusion, and tracker exclusions.
- `repost_tracker_audit.xlsx`: audit of exact/similar tracker matches and possible repost handling.

`final_action` meanings:

- `Apply Today`: highest priority, ready to apply.
- `Strong Consider`: strong fit, usually worth applying.
- `Apply If Time`: manual full JD has a realistic but lower-priority AI fit; include in the daily list if there is application capacity.
- `Manual Review`: AI score or context needs human judgement before applying.
- `Consider`: possible API/rule-based fit, lower confidence or needs manual judgement.
- `Low Priority`: weak or risky fit.
- `Skip`: hard skip or clear mismatch.
- `Pending AI Review`: manual full JD still needs AI review.

Manual full JD jobs use AI fit score thresholds after cache/OpenAI review:

- `Apply Today`: `ai_fit_score >= 80`
- `Strong Consider`: `70 <= ai_fit_score < 80`
- `Apply If Time`: `60 <= ai_fit_score < 70`
- `Manual Review`: `50 <= ai_fit_score < 60`
- `Skip`: `ai_fit_score < 50`

Company website applications use these stricter thresholds because they usually take more time. Manual jobs with low-friction application methods such as LinkedIn Easy Apply, Indeed Apply, Easy Apply, or quick apply keep the same `Apply Today` and `Strong Consider` thresholds, but can become `Apply If Time` from `ai_fit_score >= 55`.

Low-friction manual jobs get `quick_apply_candidate=True`. `quick_apply_jobs.xlsx` includes quick-apply candidates with `Apply Today`, `Strong Consider`, `Apply If Time`, or `Manual Review` with `ai_fit_score >= 55`, while excluding senior/lead/manager/head/director roles, contract/FTC/day-rate/IR35 roles, clearly non-target roles, far office-based roles, and exact tracker matches.

`ranked_jobs.xlsx` includes manual jobs with `Apply Today`, `Strong Consider`, and `Apply If Time` when they have a valid AI decision and are not blocked by a clear exclusion such as exact tracker match, hard skip, invalid AI/cache inconsistency, or deduplication. The terminal summary separates overall action counts, which may include API leads, from manual ranked-action counts, which should reconcile with `ranked_jobs.xlsx` plus `ranked_jobs_exclusion_audit.xlsx`.

Manual raw jobs are human-screened before being added, so rule-based filters should be minimal before AI review. AI should decide whether stretch-but-relevant manual jobs become `Apply Today`, `Strong Consider`, `Apply If Time`, `Manual Review`, `Low Priority`, or `Skip`.

Manual raw jobs also receive deterministic component scores so AI failures, missing AI results, malformed AI responses, or low-confidence AI/cache output cannot make a valid job disappear. Component fields include role family, technical match, direct experience, transferable experience, location/commute, salary, career value, role level, gap penalties, final score, recommendation, recommendation reason, and `deterministic_fallback_used`.

Location scoring is intentionally local-first. Milton Keynes and close MK areas receive maximum priority. Bedford, Northampton, Luton, and Brixworth receive very high priority. Nearby drivable towns receive positive scores. London is assessed with commute and office-day caution rather than automatically outranking local roles by salary.

Salary is normally a soft factor. Missing salary is not a rejection. Local analyst roles around `£27,000-£29,999` can remain visible and may still be Apply/Manual Review when technical fit, location, and career value are strong.

Manual fixed-term or contract wording is no longer an automatic pre-AI exclusion by itself. Day-rate/IR35-style roles can still be strict excluded. Other contract/FTC signals should be visible as Manual Review/contract stability risk.

## Application Materials Mapping

The assistant recommends local two-page ATS CV files and matching ATS cover letter files by job category. These files are private application materials in `cvs/` and `cover_letters/`, and both folders are ignored by Git. Do not commit CVs, cover letters, `.docx`, or `.pdf` files.

Application material categories:

- `business_intelligence_officer`: `cv_two_page_business_intelligence_officer.docx` and `cover_letter_business_intelligence_officer_ats.docx`
- `data_analyst`: `cv_two_page_data_analyst.docx` and `cover_letter_data_analyst_ats.docx`
- `bi_reporting_mi_analyst`: `cv_two_page_bi_reporting_mi_analyst.docx` and `cover_letter_bi_reporting_mi_analyst_ats.docx`
- `commercial_analyst`: `cv_two_page_commercial_analyst.docx` and `cover_letter_commercial_analyst_ats.docx`
- `business_operations_analyst`: `cv_two_page_business_operations_analyst.docx` and `cover_letter_business_operations_analyst_ats.docx`
- `finance_economics_analyst`: `cv_two_page_finance_economics_analyst.docx` and `cover_letter_finance_economics_analyst_ats.docx`

Mapping logic:

- Business Intelligence Officer, BI Officer, performance/public-sector/council/local-government reporting roles use `business_intelligence_officer`.
- Data Analyst, Junior Data Analyst, Graduate Data Analyst, and Associate Data Analyst roles use `data_analyst`.
- BI Analyst, Reporting Analyst, MI Analyst, Power BI Analyst, Dashboard Analyst, and Management Information roles use `bi_reporting_mi_analyst`.
- Commercial Analyst, Sales Analyst, Pricing Analyst, Revenue Analyst, Customer Insight, and Retail Analyst roles use `commercial_analyst`.
- Business Analyst, Operations Analyst, Process Analyst, Performance Analyst, and Business Operations roles use `business_operations_analyst`.
- Finance Analyst, Economics Analyst, Treasury Analyst, Commercial Finance Analyst, and Junior Commercial Finance Analyst roles use `finance_economics_analyst`.

`ranked_jobs.xlsx`, `quick_apply_jobs.xlsx`, and `application_pack_recommendations.xlsx` show the recommended CV and cover letter filenames. `application_pack_ready=True` means both local files exist. If either file is missing, `application_pack_ready=False` and `missing_application_file_warning` lists the missing file.

Future OpenAI review prompts ask for `cv_category` and `cover_letter_category` from the six category IDs above, but deterministic code mapping remains the fallback when AI output is missing, invalid, or cached from an older run.

Batch AI review is configured as `AI_BATCH_SIZE=10`, `AI_BATCH_MAX_API_CALLS_PER_RUN=10`, and `AI_BATCH_MAX_JOBS_PER_RUN=100`, so one `--ai-review-batch` run can review up to 100 eligible jobs if daily/global quota remains available and the user confirms the API calls. If the network is unstable, run a smaller batch:

```powershell
python run_pipeline.py --ai-review-batch --ai-batch-calls 1
```

Timeout or network failures are retried twice with short backoff. If batch AI still fails or only partly completes, the pipeline may still write output files for auditability, but it prints `WARNING: AI batch review did not complete. Outputs may be partial.` The source of truth is `output/manual_ai_coverage_audit.xlsx`: eligible manual jobs with `Pending AI review` still need OpenAI review. Re-run:

```powershell
python run_pipeline.py --ai-review-batch
```

Force re-review options:

- `--force-today`: re-run OpenAI only for eligible manual raw job `.txt` files in `jobs/raw_jobs/` whose modified date is today. This ignores the AI cache for those files only and keeps normal cache behavior for all other manual jobs.
- `--force-files "file1.txt,file2.txt"`: re-run OpenAI only for specific named raw files.
- `--force-manual-review`: re-run OpenAI for all eligible manual raw jobs.

Do not combine these force options. `--force-today` is useful after adding or editing several raw job files on the same day:

```powershell
python run_pipeline.py --ai-review-batch --force-today
```

Recommended manual-only command when re-reviewing today's raw jobs and avoiding Reed/Adzuna delays:

```powershell
python run_pipeline.py --ai-review-batch --force-today --skip-api
```

API leads remain discovery-only in `api_leads.xlsx` unless the full JD is manually added to `jobs/raw_jobs/`.

`ai_review_queue.xlsx` keeps manual jobs needing attention, including `Manual Review`, `Pending AI Review`, `Needs investigation`, suspicious parse cases, contradictory AI results, and similar rejected/applied/interview history that needs human review. It should not include `Apply Today`, `Strong Consider`, or `Apply If Time` only because of generic salary/experience review notes. A manual row logged as `not sent to ai_review_queue` may already have a valid AI result; it does not mean the job was skipped or missed by AI.

## Tracker Logic

- `applications.xlsx` stores current applications.
- `applications_history.xlsx` stores historical records.
- `status_history.xlsx` records changes.
- `tracker/backups/` stores backups before write operations.
- Exact tracker matches should prevent duplicate applications.
- Similar rejected, applied, assessment, interview, final interview, or offer history should not automatically exclude new reposts. It should flag the row for review.
- Same company by itself must never be treated as a duplicate. The same company can have multiple different jobs.
- Exact duplicate matching should use strong evidence: same `apply_link`, same external job id, same canonical/requisition id, or same company, job title, location, and description hash.
- Same company and same title with a different apply link, job id, posted date, or description should be treated as a possible repost/new vacancy and allowed through review.
- Previously rejected similar jobs should be flagged with tracker/repost fields, not automatically blocked.
- If a manual job is excluded by tracker without strong exact evidence, print `WARNING: Possible repost excluded. Check output/repost_tracker_audit.xlsx.` and review `repost_tracker_audit.xlsx`.

Repost/audit fields include:

- `tracker_exact_match`
- `tracker_similarity_match`
- `repost_candidate`
- `previous_rejection_date`
- `previous_application_status`
- `tracker_overlay_reason`
- `next_action`

## GitHub Privacy Safety

Never commit private local job-search data:

- `jobs/raw_jobs/`
- `jobs/applied_jobs/`
- `jobs/archived_jobs/`
- `jobs/rejected_jobs/`
- `jobs/api_jobs.csv`
- `tracker/`
- `output/`
- `data/`
- `cache/`
- `logs/`
- CVs and cover letters
- `.xlsx`, `.pdf`, and `.docx` files
- `credentials.json`, `token.json`, `.env`, and any credential/token files

Commit source code, docs, README files, requirements/config templates, and fake examples only. The `examples/` folder must contain synthetic sample data, never real company/application history.

## Gmail Status Assistant Logic

- `run_email_status.py` uses Gmail API read-only access.
- `credentials.json` stores OAuth client info.
- `token.json` stores the local OAuth token.
- Dry-run is the default. `python run_email_status.py` does not modify tracker files.
- Live tracker updates require `python run_email_status.py --apply`.
- Tracker updates should only happen when `safe_to_update=True`.
- Acknowledgements should be conservative.
- Rejections, interviews, assessments, and offers can update tracker only when company and job title match with high confidence.

Current Gmail extraction uses sender display/domain, subject patterns, and body patterns. Company aliases include Legal & General, AXIS, Vitality, Houseful, Ministry of Justice, AIT Home Delivery, and innocent. The output review file includes sender, subject, body, final extracted company/title, alias used, matched tracker company/title, match explanation, unsafe reason, and suggested action columns.

In live `--apply` mode, each applied update is audited in `output/email_status_updates_applied.xlsx` with old status, new status, matched tracker company/title, matched tracker file, match confidence, and match explanation. If no safe updates are found, tracker backups and writes are skipped.

If Gmail fails with `invalid_grant` or an expired/revoked token, do not delete `credentials.json`. Rename `token.json` and re-run dry-run:

```powershell
cd "E:\Back up\IT learning\Data Analyst\job_search_assistant"
rename-item token.json token_old_revoked.json
python run_email_status.py --dry-run
```

## Gmail Safety Rules

- Never auto-update based on job title only.
- Company or trusted company alias must match.
- `match_confidence` should be `>= 85` for auto-update.
- If multiple tracker rows match, manual review only.
- If company is missing, manual review only.
- If status is `ignore`, never update.
- If status is `applied_acknowledgement`, keep conservative unless the match is very strong.
- Do not weaken safety rules just to reduce manual review.

## Known Historical Bugs Fixed

- Manual raw jobs not all being parsed.
- Manual full JD jobs stuck as `rule_based` / `Pending AI Review` but not selected for AI review.
- OpenAI JSON serialization issue with `int64`.
- Salary parser wrongly reading revenue numbers such as `£62m` as salary.
- Contract parser wrongly treating `Contract: Permanent` as contract role.
- API leads should not consume AI review capacity.
- Company/title parser may still need improvement for LinkedIn/Indeed UI text.
- Raw job lifecycle archiving added so applied manual `.txt` files leave `jobs/raw_jobs/` after `mark_status.py` successfully updates tracker status.

## Known Current Improvement Areas

- Company/title extraction sometimes reads UI text such as Apply, Saved, Show match details, or Your profile matches some required qualifications.
- Gmail email matching still needs better sender/subject/body extraction for new ATS templates.
- `invalid_grant` Gmail token errors should be handled gracefully and require token re-authorisation.
- `ranked_jobs` count and summary count may sometimes differ slightly and should be audited.
- pandas `FutureWarning` from concat should eventually be cleaned.
- Raw job archive matching depends on ranked output metadata; if a tracker row has no matching raw file, `mark_status.py` logs a warning and leaves the status update intact.

## Modification Rule for Future Codex Sessions

Before modifying code:

- Read `docs/PROJECT_LOGIC.md` first.
- Identify which workflow will be affected.
- Avoid changing tracker/cache/output files unless explicitly asked.
- Do not delete `credentials.json` or `token.json`.
- Do not call OpenAI API unless explicitly approved.
- After modifying code, update `docs/PROJECT_LOGIC.md` if logic changed.
- Run `py_compile` on modified Python files.
- Provide a short changelog summary.

## Change Log

### 2026-06-24

- Created `docs/PROJECT_LOGIC.md`.
- Improved Gmail email extraction and matching with sender/domain aliases, subject patterns, body patterns, high-confidence company/title matching, and safer title-only handling.
- Improved Gmail `invalid_grant` handling with friendly token re-authorisation instructions.
- Added manual raw job lifecycle archive logic for applied jobs, configurable rejected/archived moves, and `output/raw_job_file_movement_log.xlsx`.

### 2026-06-26

- Made Gmail status checking dry-run by default; live updates now require `--apply`.
- Added live Gmail update transparency and `output/email_status_updates_applied.xlsx`.
- Improved generic email extraction cleanup for invalid company values and Indeed job post subjects.

### 2026-06-28

- Added `Apply If Time` as a manual full JD action tier.
- Adjusted manual AI score thresholds so manually screened jobs produce a wider realistic application list while keeping hard skip safety filters.
- Kept API leads discovery-only; `ranked_jobs.xlsx` is manual-job focused.
- Relaxed manual pre-AI hard skips so human-screened raw jobs normally reach AI review; added `output/manual_pre_ai_exclusion_audit.xlsx`.
- Added `output/ranked_jobs_exclusion_audit.xlsx` and clarified that AI-reviewed rows not sent to `ai_review_queue` are not skipped before AI.

### 2026-07-03

- Added CORD (`cord.co`) source detection for manually pasted raw jobs and CORD URL job id extraction for safer duplicate/tracker matching.
- Increased batch OpenAI review capacity from 20 to 100 jobs per run by setting 10 batch calls x 10 jobs, with matching daily batch limits.
- Added AI batch retry/partial-output warnings, manual coverage rerun guidance, and `--ai-batch-calls` for smaller safer batch runs.

### 2026-07-05

- Added `output/pipeline_debug_audit.xlsx` and `output/repost_tracker_audit.xlsx`.
- Clarified ranked jobs reconciliation: overall action counts can include API leads; manual ranked-action counts reconcile with `ranked_jobs.xlsx`.
- Tightened AI review queue rules so generic salary/experience `needs_human_review` does not place otherwise actionable jobs into `ai_review_queue.xlsx`.
- Added a stronger guard around manual tracker exact exclusions so same company/title alone is treated as possible repost rather than automatic exclusion.
- Added `--force-today` for batch OpenAI re-review of eligible manual raw job files modified today.
- Added `--skip-api` so manual AI review can run without Reed/Adzuna API collection, and made API HTTP 503 handling preserve existing API jobs.

### 2026-07-10

- Added low-friction apply detection for LinkedIn Easy Apply, Indeed Apply, Easy Apply, and quick apply manual jobs.
- Lowered only the low-friction manual `Apply If Time` threshold to `ai_fit_score >= 55`; `Apply Today` and `Strong Consider` thresholds remain unchanged.
- Added `quick_apply_candidate` and `output/quick_apply_jobs.xlsx` as an optional quick-application volume list separate from the main `ranked_jobs.xlsx`.

### 2026-07-12

- Added Welcome to the Jungle (`welcometothejungle.com` / `WTTJ`) source detection for manually pasted raw job links.
- Added Welcome to the Jungle URL job slug/id extraction for duplicate and tracker matching evidence.

### 2026-07-15

- Added GitHub privacy safety documentation and a fake `examples/sample_raw_job.txt`.
- Updated `.gitignore` coverage for private raw/applied/archive/rejected job folders and local generated data.

### 2026-07-16

- Added robust manual raw job traceability with `raw_job_trace_audit.xlsx`, `filtered_jobs.xlsx`, `duplicates.xlsx`, and `parsing_errors.xlsx`.
- Improved LinkedIn/Indeed-style manual parser handling for `Save {title} at {company}` lines, UI noise, publication-vs-employer confusion, missing fields, and local location extraction.
- Expanded role-family recognition for BI, reporting, officer, executive, insight, business operations, HR business analyst, finance, and commercial analyst variants.
- Added deterministic component scoring for eligibility, title/role family, technical fit, direct and transferable experience, location, salary, career value, role level, and gap penalties.
- Relaxed manual contract/FTC handling so fixed-term wording is reviewable rather than silently excluded; day-rate/IR35-style roles remain strict exclusions.
- Added regression tests for the five missing-job scenarios and parser/audit edge cases.

### 2026-07-18

- Updated application-material mapping to the six two-page ATS CV categories and matching ATS cover letter categories.
- Added `application_pack_ready`, missing-file warnings, and `output/application_pack_recommendations.xlsx`.
- Updated future AI review prompts to request `cv_category` and `cover_letter_category` while keeping deterministic fallback mapping.
