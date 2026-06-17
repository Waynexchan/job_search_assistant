# Changelog

## 2026-06-17
- Manual `raw_jobs` txt loading fixed so all `.txt` files are loaded as manual jobs.
- Manual full JD jobs bypass hard filters before AI review where they need human or AI assessment.
- API jobs moved to discovery leads only instead of final ranked application decisions.
- AI batch review added and improved for manual full JDs.
- AI review cache and force review behavior improved.
- Salary parser fixed to ignore revenue, RSV, benefit, bonus, and similar non-salary amounts such as `£62m`.
- Contract parser fixed to treat `Contract: Permanent` correctly.
- Gmail email status checker identified as an existing module with dry-run and review outputs.
- Reposted job tracker filtering improved:
  - Exact same tracked jobs remain excluded for terminal statuses.
  - Similar previously rejected roles are kept for manual/AI review.
  - Repost flags and previous rejection context are added to outputs.
  - API leads for similar rejected roles are kept with a link-check action.

## Earlier Development
- Added local application tracker spreadsheets.
- Added status history and application history tracking.
- Added role scoring for analyst keywords, technical fit, business-domain fit, location, and salary risk.
- Added CV and cover letter recommendation fields.
- Added duplicate detection across job sources.
- Added manual raw job parsing from copied job descriptions.
- Added API collection from Reed and Adzuna.
- Added ranked daily output and full audit output spreadsheets.
