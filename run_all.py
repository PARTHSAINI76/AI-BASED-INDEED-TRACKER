"""
Run the complete recruitment workflow (Tasks 1-5) end to end in SAFE / SIMULATED mode.

  python run_all.py

Task 1  build_tracker.py + publish_to_indeed.py   -> Indeed_Job_Posting_Tracker.xlsx
Task 2  categorise_applicants.py                  -> Indeed_Applicant_Summary.xlsx
Task 3  shortlist_candidates.py                   -> Indeed_Candidate_Database.xlsx, Indeed_Shortlist_Masked.xlsx
Task 4  auto_calling.py                           -> Calling_Prototype.xlsx
Task 5  schedule_interviews.py                    -> Interview_Schedule.xlsx, Interview_Schedule_HiringTeam.xlsx

Nothing is posted, called or emailed. Real accounts/vendors plug in through the provider interfaces
(see README.md). Optional env: AUTHORISED_USERS, REQUESTING_USER, CANDIDATE_DB_PASSWORD.
"""
import logging

import auto_calling
import build_tracker
import categorise_applicants
import publish_to_indeed
import schedule_interviews
import shortlist_candidates


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    print("== Task 1: job posting tracker ==")
    build_tracker.main()
    publish_to_indeed.main()
    print("== Task 2: applicant categorisation ==")
    categorise_applicants.main()
    print("== Task 3: consolidated shortlist ==")
    user = shortlist_candidates.authorise()
    cands = shortlist_candidates.main()
    shortlisted = [c for c in cands if c["status"] == "Shortlisted"]
    print("== Task 4: automated calling prototype ==")
    calling = auto_calling.run(shortlisted, user)
    print("== Task 5: interview scheduling ==")
    schedule_interviews.run(cands, calling, user)
    print("\nDone. Open the .xlsx files; start with the Summary sheet in each.")


if __name__ == "__main__":
    main()
