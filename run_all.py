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
