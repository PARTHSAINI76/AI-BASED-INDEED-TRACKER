# Indeed Recruitment Workflow - Tasks 1 to 5

A complete, runnable prototype of the five-task hiring workflow. It runs in **safe, simulated mode**:
nothing is posted to Indeed, no call is placed, no message is sent. Real accounts and vendors plug in
through small provider interfaces (see "Going live").

## Run it

```
pip install openpyxl
python run_all.py
```

Optional environment variables: `AUTHORISED_USERS` (comma list), `REQUESTING_USER`,
`CANDIDATE_DB_PASSWORD` (sheet/structure protection for the candidate database).

## What each task produces

| Task | Code | Output | Headline result (sample data) |
|---|---|---|---|
| 1 Job postings | `build_tracker.py`, `publish_to_indeed.py` | `Indeed_Job_Posting_Tracker.xlsx` | 3 postings (full content, keywords, approved-salary placeholder), tracker with date / reference / status |
| 2 Categorisation | `categorise_applicants.py` | `Indeed_Applicant_Summary.xlsx` | 15 applicants: 4 Relevant, 6 Partial, 4 Not Relevant, 1 Incomplete; Experienced 11 / Fresher 3; reasons per applicant |
| 3 Shortlist | `shortlist_candidates.py` | `Indeed_Candidate_Database.xlsx`, `Indeed_Shortlist_Masked.xlsx` | 19 applications -> 15 unique candidates; 4 Shortlisted, 7 On Hold, 4 Not Shortlisted; priorities P1-P3 |
| 4 Calling | `auto_calling.py` | `Calling_Prototype.xlsx` | Platform/integration/workflow/approvals plan, 29 passing safeguard tests, call log with all 7 outcomes |
| 5 Scheduling | `schedule_interviews.py` | `Interview_Schedule.xlsx`, `Interview_Schedule_HiringTeam.xlsx` | Scheduled / Rescheduled / Confirmation Pending / Declined shown; 19 passing tests incl. No Response |

Every workbook opens on a **Summary** sheet whose counts are formulas, so reviewer edits flow through.

## Design choices

- **One source of truth.** Job data lives in `build_tracker.py`; later tasks import it.
- **Pluggable providers.** Publisher (Task 1), ApplicantSource (2/3), CallProvider and DNDRegistry (4),
  Sender (5). Mock versions are tested; live versions are skeletons marked TODO.
- **Safeguards are code, not policy text.** All call rules sit in one `gate()` function and are covered by
  executable tests. The tests were mutation-checked (breaking a rule makes them fail).
- **Job-related criteria only.** Name, age, gender, etc. are never scored.

## Going live (nothing below is done yet)

1. **Indeed posting:** Indeed has no open self-serve posting API. Use its approved partner route (or post
   manually from the employer dashboard) and fill `IndeedLivePublisher._authenticate/_send` from Indeed's
   current partner documentation. No scraping or UI automation.
2. **Applicants:** export the applicant CSV from the employer dashboard; map its headers to
   `job_id,name,email,phone,location,experience_years,skills,education,expected_salary_lpa,notice_days,applied_date`;
   run with `APPLICANT_SOURCE=csv APPLICANT_CSV=file.csv AUTHORISED_USERS=...`.
3. **Calling:** choose a telephony vendor (approval needed), implement `LiveProvider`, load real consent
   records, then run with written approval (`CALL_APPROVAL_ID`, `CALL_APPROVED_BY`). Batch size is capped.
   Confirm calling hours, DND handling and consent wording with Legal first.
4. **Messages:** implement an approved email/SMS `Sender`. Replace the TBC interviewer, contact person,
   office address and placeholder meeting links.

## Assumptions to confirm

Salary bands, roles, locations, interviewers and all candidate data are samples. Score weights and
thresholds (Task 2), P1 cut-off 85 (Task 3), calling window 09:00-20:00 IST, 3 attempts, 4 h gap (Task 4),
and 48 h acknowledgement window, 45-minute slots (Task 5) are adjustable constants at the top of each file.
Excel sheet protection is a guard against accidental edits, not strong security: keep files on an
access-controlled drive. Retention periods and applicability of TRAI / DPDP Act rules to recruitment
calls must be confirmed with HR and Legal.
