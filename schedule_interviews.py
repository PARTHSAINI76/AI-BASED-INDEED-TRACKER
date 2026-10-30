from __future__ import annotations

import logging
import os
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import datetime, time, timedelta

from auto_calling import IST, CallCampaign
from build_tracker import JOBS
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from shortlist_candidates import FONT, _body, _hdr, authorise, log_access

log = logging.getLogger("schedule")

SCHEDULED, PENDING, RESCHEDULED = "Scheduled", "Confirmation Pending", "Rescheduled"
NO_RESPONSE, DECLINED = "No Response", "Declined"
STATUSES = [SCHEDULED, PENDING, RESCHEDULED, NO_RESPONSE, DECLINED]

SLOT_HOURS = (10, 11, 12, 14, 15, 16)   # 45-min interviews; lunch 13:00-14:00 blocked; ends by 16:45
SLOT_MINUTES = 45
SEARCH_DAYS = 14
ACK_WINDOW = timedelta(hours=48)
FILE_HR, FILE_TEAM = "Interview_Schedule.xlsx", "Interview_Schedule_HiringTeam.xlsx"

ROLE = {j["id"]: j["title"] for j in JOBS}
INTERVIEWER = {"JP-001": "Engineering Manager (name TBC)", "JP-002": "Frontend Lead (name TBC)",
               "JP-003": "QA Lead (name TBC)"}
HR_CONTACT = "HR Coordinator (contact details TBC)"
OFFICE = "Company office, Gorakhpur (address TBC)"
MODE = {"JP-001": "Video call", "JP-002": "In-person", "JP-003": "Video call"}   # JP-002 is an on-site role


@dataclass
class Availability:
    weekdays: set[int] | None = None     # 0=Mon .. 4=Fri; None = any
    part: str | None = None              # "morning" | "afternoon" | None

    def accepts(self, dt: datetime) -> bool:
        if self.weekdays is not None and dt.weekday() not in self.weekdays:
            return False
        if self.part == "morning" and dt.hour >= 13:
            return False
        return not (self.part == "afternoon" and dt.hour < 13)


@dataclass
class Interview:
    iid: str
    cid: str
    name: str
    email: str
    phone: str
    job_id: str
    priority: str
    start: datetime | None = None
    mode: str = ""
    venue: str = ""
    interviewer: str = ""
    status: str = PENDING
    sent_at: datetime | None = None
    note: str = ""
    history: list = field(default_factory=list)

    @property
    def deadline(self) -> datetime | None:
        return self.sent_at + ACK_WINDOW if self.sent_at else None


# ----------------------------------------------------------------------------- messaging
class Sender(ABC):
    @abstractmethod
    def send(self, to_cid: str, channel: str, subject: str, body: str, now: datetime) -> None: ...


class MockSender(Sender):
    def __init__(self):
        self.outbox: list[dict] = []

    def send(self, to_cid, channel, subject, body, now):
        self.outbox.append({"time": now.replace(tzinfo=None), "cid": to_cid, "channel": channel,
                            "subject": subject, "body": body, "state": "SIMULATED - not sent"})


def confirmation_message(iv: Interview) -> tuple[str, str]:
    when = f"{iv.start:%A, %d %B %Y} at {iv.start:%I:%M %p} IST ({SLOT_MINUTES} minutes)"
    place = f"Venue: {iv.venue}" if iv.mode == "In-person" else f"Meeting link: {iv.venue}"
    subject = f"Interview Confirmation - {ROLE[iv.job_id]}"
    body = (f"Dear {iv.name},\n\nThank you for your interest in the {ROLE[iv.job_id]} position. "
            f"We would like to invite you for an interview.\n\nDate & time: {when}\nMode: {iv.mode}\n{place}\n"
            f"Interviewer: {iv.interviewer}\nContact person: {HR_CONTACT}\n\n"
            f"Please reply CONFIRM to acknowledge, or RESCHEDULE with your preferred times, by "
            f"{iv.deadline:%d %b %Y, %I:%M %p} IST. If we do not hear from you we will release the slot.\n\n"
            f"Regards,\nRecruitment Team")
    return subject, body


# ----------------------------------------------------------------------------- scheduler
class Scheduler:
    def __init__(self, sender: Sender):
        self.sender = sender
        self.busy: dict[str, set[datetime]] = {}
        self.interviews: list[Interview] = []
        self.unscheduled: list[dict] = []

    # --- slot grid
    @staticmethod
    def slot_grid(after: datetime, days: int = SEARCH_DAYS) -> list[datetime]:
        d = after.date() + timedelta(days=1)          # earliest: next calendar day
        out, seen = [], 0
        while seen < days:
            if d.weekday() < 5:
                out += [datetime.combine(d, time(h), tzinfo=IST) for h in SLOT_HOURS]
                seen += 1
            d += timedelta(days=1)
        return out

    def _free(self, interviewer: str, cand_busy: set[datetime], avail: Availability, after: datetime):
        taken = self.busy.setdefault(interviewer, set())
        for s in self.slot_grid(after):
            if s not in taken and s not in cand_busy and avail.accepts(s):
                return s
        return None

    def _candidate_slots(self, cid: str) -> set[datetime]:
        return {i.start for i in self.interviews if i.cid == cid and i.start}

    # --- lifecycle
    def _set(self, iv: Interview, status: str, now: datetime, note: str = ""):
        iv.status = status
        iv.history.append((now.replace(tzinfo=None), status, note))

    def _release(self, iv: Interview):
        if iv.start:
            self.busy.get(iv.interviewer, set()).discard(iv.start)

    def _book(self, iv: Interview, slot: datetime):
        iv.start, iv.mode = slot, MODE[iv.job_id]
        iv.interviewer = INTERVIEWER[iv.job_id]
        iv.venue = OFFICE if iv.mode == "In-person" else f"https://meet.example.com/{iv.iid.lower()} (placeholder link)"
        self.busy.setdefault(iv.interviewer, set()).add(slot)

    def propose(self, iid: str, cand: dict, avail: Availability, now: datetime, note: str = "") -> Interview | None:
        job = cand["best_job"]
        slot = self._free(INTERVIEWER[job], set(), avail, now)
        if slot is None:
            self.unscheduled.append({"cid": cand["cid"], "reason": "No free slot in next "
                                     f"{SEARCH_DAYS} working days - schedule manually"})
            return None
        iv = Interview(iid, cand["cid"], cand["name"], cand.get("email", ""), cand.get("phone", ""), job,
                       cand["priority"])
        self._book(iv, slot)
        iv.sent_at = now
        self.interviews.append(iv)
        self._set(iv, PENDING, now, note or "Slot proposed; confirmation requested")
        self.sender.send(iv.cid, "Email", *confirmation_message(iv), now)
        return iv

    def acknowledge(self, iv: Interview, now: datetime):
        if iv.status in (PENDING, RESCHEDULED):
            self._set(iv, SCHEDULED, now, "Candidate acknowledged")

    def reschedule(self, iv: Interview, avail: Availability, now: datetime) -> bool:
        old = iv.start
        self._release(iv)
        slot = self._free(iv.interviewer, set(), avail, now)
        if slot is None:       # keep the original slot rather than lose it
            self._book(iv, old)
            return False
        self._book(iv, slot)
        iv.sent_at = now
        self._set(iv, RESCHEDULED, now, f"Moved from {old:%d %b %H:%M} at candidate's request; awaiting re-acknowledgement")
        self.sender.send(iv.cid, "Email", *confirmation_message(iv), now)
        return True

    def decline(self, iv: Interview, now: datetime, note: str = "Candidate declined"):
        self._release(iv)
        self._set(iv, DECLINED, now, note)
        iv.note = "Slot released (original time kept for audit)"

    def expire_unanswered(self, now: datetime):
        for iv in self.interviews:
            if iv.status in (PENDING, RESCHEDULED) and iv.deadline and now >= iv.deadline:
                self._release(iv)
                self._set(iv, NO_RESPONSE, now, "No acknowledgement within 48 h; slot released")
                iv.note = "Slot released"


# ----------------------------------------------------------------------------- scenario
RESPONSES = {   # SAMPLE candidate behaviour, keyed by sample name
    "Sample Applicant 01": {"avail": Availability({0, 1, 2}, "morning"), "reply": "confirm"},
    "Sample Applicant 11": {"avail": Availability(None, None), "reply": None},
    "Sample Applicant 06": {"avail": Availability(None, "afternoon"), "reply": "reschedule",
                            "new_avail": Availability({2, 3}, "morning")},
    "Sample Applicant 02": {"reply": "decline", "avail": Availability(None, None)},
}


def build_schedule(cands: list[dict], calling: CallCampaign, now: datetime, sender: Sender) -> Scheduler:
    sch = Scheduler(sender)
    by_cid = {c["cid"]: c for c in cands}
    n = 0
    for contact in calling.contacts:
        c = by_cid[contact.cid]
        final = calling.final_status(c["cid"])
        resp = RESPONSES.get(c["name"], {"avail": Availability(), "reply": None})
        if final not in ("Interested", "Callback Requested", "Blocked"):
            sch.unscheduled.append({"cid": c["cid"], "reason": f"Call outcome '{final}' - not eligible"})
            continue
        # Callback Requested / Blocked(no consent): a human recruiter makes contact first (simulated here)
        if final == "Blocked" and resp["reply"] == "decline":
            n += 1
            iv = Interview(f"INT-{n:03d}", c["cid"], c["name"], c["email"], c["phone"], c["best_job"], c["priority"],
                           interviewer=INTERVIEWER[c["best_job"]])
            sch.interviews.append(iv)
            sch._set(iv, DECLINED, now, "No call consent; recruiter emailed - candidate declined")
            iv.note = "No slot held"
            continue
        n += 1
        iv = sch.propose(f"INT-{n:03d}", c, resp["avail"], now,
                         note=("Interested on IVR call" if final == "Interested" else "After recruiter callback"))
        if iv is None:
            continue
        reply = resp["reply"]
        if reply == "confirm":
            sch.acknowledge(iv, now + timedelta(hours=2))
        elif reply == "reschedule":
            sch.reschedule(iv, resp["new_avail"], now + timedelta(hours=3))
        elif reply == "decline":
            sch.decline(iv, now + timedelta(hours=3))
    sch.expire_unanswered(now)
    return sch


# ----------------------------------------------------------------------------- tests
def run_tests() -> list[tuple]:
    rows: list[tuple] = []

    def add(area, scenario, expected, actual):
        rows.append((f"S{len(rows) + 1:02d}", area, scenario, str(expected), str(actual),
                     "PASS" if str(expected) == str(actual) else "FAIL"))

    base = datetime(2026, 11, 6, 15, 0, tzinfo=IST)    # a Friday
    cand = lambda i, job="JP-001": {"cid": f"C{i}", "name": f"N{i}", "email": "", "phone": "", "best_job": job,
                                    "priority": "P1 - High"}
    grid = Scheduler.slot_grid(base)
    add("Slot grid", "No weekend slots", "True", all(s.weekday() < 5 for s in grid))
    add("Slot grid", "No slot in lunch hour / after 16:45", "True",
        all(s.hour != 13 and s.hour <= 16 for s in grid))
    add("Slot grid", "First slot is next working day (Mon)", "0", grid[0].weekday())

    sch = Scheduler(MockSender())
    ivs = [sch.propose(f"I{i}", cand(i), Availability(), base) for i in range(8)]
    starts = [i.start for i in ivs]
    add("No double booking", "8 candidates, one interviewer -> 8 distinct slots", "8", len(set(starts)))
    sch2 = Scheduler(MockSender())
    a = sch2.propose("A", cand(1, "JP-001"), Availability(), base)
    b = sch2.propose("B", cand(2, "JP-003"), Availability(), base)
    add("No double booking", "Different interviewers may share a time", "True", a.start == b.start)

    sch = Scheduler(MockSender())
    iv = sch.propose("A", cand(1), Availability({0, 1, 2}, "morning"), base)
    add("Availability", "Mon-Wed morning respected", "True", iv.start.weekday() <= 2 and iv.start.hour < 13)

    sch = Scheduler(MockSender())
    iv = sch.propose("A", cand(1), Availability(), base)
    add("Status", "New proposal is Confirmation Pending", PENDING, iv.status)
    sch.acknowledge(iv, base)
    add("Status", "Acknowledgement -> Scheduled", SCHEDULED, iv.status)

    sch = Scheduler(MockSender())
    iv = sch.propose("A", cand(1), Availability(), base)
    old = iv.start
    sch.reschedule(iv, Availability({3}, "morning"), base)
    add("Reschedule", "Status -> Rescheduled", RESCHEDULED, iv.status)
    add("Reschedule", "Old slot released and re-usable", "True",
        old not in sch.busy[iv.interviewer] and iv.start != old)
    other = sch.propose("B", cand(2), Availability(), base)
    add("Reschedule", "Released slot given to next candidate", "True", other.start == old)

    sch = Scheduler(MockSender())
    iv = sch.propose("A", cand(1), Availability(), base)
    sch.expire_unanswered(base + timedelta(hours=47))
    add("No response", "Still pending at 47 h", PENDING, iv.status)
    sch.expire_unanswered(base + timedelta(hours=49))
    add("No response", "No Response after 48 h", NO_RESPONSE, iv.status)
    add("No response", "Slot released after lapse", "True", iv.start not in sch.busy[iv.interviewer])

    sch = Scheduler(MockSender())
    iv = sch.propose("A", cand(1), Availability(), base)
    s = iv.start
    sch.decline(iv, base)
    add("Declined", "Status Declined, slot freed", "True", iv.status == DECLINED and s not in sch.busy[iv.interviewer])

    sch = Scheduler(MockSender())
    full = Availability({0}, "morning")   # Mondays 10-12 only: 3 slots/week x 2 weeks in window
    got = [sch.propose(f"I{i}", cand(i), full, base) for i in range(10)]
    add("Capacity", "Full calendar -> candidate left unscheduled (manual)", "True", got[-1] is None and len(sch.unscheduled) > 0)

    sch = Scheduler(MockSender())
    iv = sch.propose("A", cand(1, "JP-002"), Availability(), base)
    subj, body = confirmation_message(iv)
    needs = [f"{iv.start:%d %B %Y}", f"{iv.start:%I:%M %p}", iv.mode, iv.venue, iv.interviewer, HR_CONTACT, "CONFIRM"]
    add("Message", "Contains date, time, mode, venue, interviewer, contact, ack request", "True",
        all(x in body for x in needs))
    add("Message", "Drafts are simulated, not sent", "SIMULATED - not sent", sch.sender.outbox[0]["state"])
    add("Status", "All 5 required statuses defined", "5", len(set(STATUSES)))
    return rows


# ----------------------------------------------------------------------------- workbook
def _sheet(wb, title, heads, rows, widths):
    ws = wb.create_sheet(title)
    for c, h in enumerate(heads, 1):
        ws.cell(row=1, column=c, value=h)
        ws.column_dimensions[get_column_letter(c)].width = widths.get(c, 16)
    _hdr(ws, 1, len(heads))
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            ws.cell(row=r, column=c, value=v)
        _body(ws, r, len(heads))
    ws.freeze_panes = "A2"
    return ws


def write_workbooks(sch: Scheduler, tests: list[tuple], now: datetime, user: str) -> None:
    ivs = sorted(sch.interviews, key=lambda i: (i.start is None, i.start or now))
    wb = Workbook()
    wb.remove(wb.active)
    heads = ["Interview ID", "Candidate ID", "Candidate", "Email", "Phone", "Role", "Date", "Time (IST)", "Mode",
             "Venue / Meeting link", "Interviewer", "Contact person", "Confirmation Status", "Ack. due by", "Notes"]
    rows = [(i.iid, i.cid, i.name, i.email, i.phone, ROLE[i.job_id], i.start.date() if i.start else None,
             i.start.time() if i.start else None, i.mode, i.venue, i.interviewer, HR_CONTACT, i.status,
             i.deadline.replace(tzinfo=None) if i.deadline and i.status in (PENDING, RESCHEDULED) else None, i.note)
            for i in ivs]
    n = len(rows) + 1
    # Summary
    s = wb.create_sheet("Summary")
    s.column_dimensions["A"].width = 38
    s.column_dimensions["B"].width = 14
    s["A1"] = "Interview Scheduling - Summary"
    s["A1"].font = Font(name=FONT, bold=True, size=14)
    s["A2"] = f"Generated {now:%d-%b-%Y %H:%M} IST by {user}. Messages are SIMULATED drafts - nothing was sent."
    s["A2"].font = Font(name=FONT, italic=True, size=9)
    s["A4"], s["B4"] = "Candidates processed", f"=COUNTA(Schedule!A2:A{n})"
    s["A4"].font = s["B4"].font = Font(name=FONT, bold=True)
    for c, h in enumerate(["Confirmation status", "Count"], 1):
        s.cell(row=6, column=c, value=h)
    _hdr(s, 6, 2)
    for i, st in enumerate(STATUSES, 7):
        s.cell(row=i, column=1, value=st)
        s.cell(row=i, column=2, value=f"=COUNTIF(Schedule!$M$2:$M${n},A{i})")
        for c in (1, 2):
            s.cell(row=i, column=c).font = Font(name=FONT)
    s.cell(row=12, column=1, value="Check: statuses sum to total").font = Font(name=FONT, italic=True, size=9)
    s.cell(row=12, column=2, value='=IF(SUM(B7:B11)=B4,"OK","MISMATCH")').font = Font(name=FONT, italic=True, size=9)
    s["A14"] = "Not scheduled (not eligible / no slot)"
    s["A14"].font = Font(name=FONT, bold=True)
    for k, u in enumerate(sch.unscheduled, 15):
        s.cell(row=k, column=1, value=f"{u['cid']}: {u['reason']}").font = Font(name=FONT, size=10)

    ws = _sheet(wb, "Schedule", heads, rows, {1: 12, 2: 12, 3: 22, 4: 24, 5: 17, 6: 34, 7: 13, 8: 11, 9: 12, 10: 46,
                                              11: 30, 12: 30, 13: 22, 14: 18, 15: 44})
    for r in range(2, n + 1):
        ws.cell(row=r, column=7).number_format = "ddd dd-mmm-yyyy"
        ws.cell(row=r, column=8).number_format = "hh:mm"
        ws.cell(row=r, column=14).number_format = "dd-mmm-yyyy hh:mm"
        ws.row_dimensions[r].height = 34
    dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"M2:M{n}")
    _sheet(wb, "Messages (simulated)", ["Time", "Candidate ID", "Channel", "Subject", "Message", "State"],
           [(m["time"], m["cid"], m["channel"], m["subject"], m["body"], m["state"]) for m in sch.sender.outbox],
           {1: 18, 2: 12, 3: 9, 4: 36, 5: 90, 6: 22})
    hist = [(i.iid, i.cid, ts, st, nt) for i in sch.interviews for ts, st, nt in i.history]
    _sheet(wb, "Status History", ["Interview ID", "Candidate ID", "Time", "Status", "Note"], hist,
           {3: 18, 4: 22, 5: 70})
    _sheet(wb, "Test Results", ["Test", "Area", "Scenario", "Expected", "Actual", "Result"], tests,
           {1: 7, 2: 18, 3: 62, 4: 22, 5: 22, 6: 9})
    wb.move_sheet("Summary", offset=-(len(wb.sheetnames) - 1))
    wb.save(FILE_HR)
    os.chmod(FILE_HR, 0o600)

    # Hiring-team version: no candidate email/phone, only active interviews
    active = [i for i in ivs if i.status in (SCHEDULED, PENDING, RESCHEDULED)]
    team = Workbook()
    t = team.active
    t.title = "Interview Schedule"
    th = ["Date", "Time (IST)", "Candidate", "Role", "Mode", "Venue / Meeting link", "Interviewer",
          "Contact person", "Confirmation Status"]
    widths = {1: 16, 2: 11, 3: 22, 4: 34, 5: 12, 6: 46, 7: 30, 8: 30, 9: 22}
    for c, h in enumerate(th, 1):
        t.cell(row=1, column=c, value=h)
        t.column_dimensions[get_column_letter(c)].width = widths[c]
    _hdr(t, 1, len(th))
    for r, i in enumerate(active, 2):
        for c, v in enumerate([i.start.date(), i.start.time(), i.name, ROLE[i.job_id], i.mode, i.venue,
                               i.interviewer, HR_CONTACT, i.status], 1):
            t.cell(row=r, column=c, value=v)
        _body(t, r, len(th))
        t.cell(row=r, column=1).number_format = "ddd dd-mmm-yyyy"
        t.cell(row=r, column=2).number_format = "hh:mm"
    t.freeze_panes = "A2"
    t.cell(row=len(active) + 3, column=1,
           value="Contact details are held by HR only. Declined / lapsed candidates are not listed.").font = Font(name=FONT, italic=True, size=9)
    team.save(FILE_TEAM)
    os.chmod(FILE_TEAM, 0o600)


def run(cands: list[dict], calling: CallCampaign, user: str, now: datetime | None = None) -> Scheduler:
    now = now or datetime.now(IST)
    tests = run_tests()
    failed = [t for t in tests if t[5] != "PASS"]
    if failed:
        raise RuntimeError(f"{len(failed)} scheduling test(s) failed: {failed[0]}")
    log.info("Scheduling tests: %d/%d passed", len(tests), len(tests))
    sch = build_schedule(cands, calling, now, MockSender())
    write_workbooks(sch, tests, now, user)
    log_access(user, "interview_schedule", "sample", len(sch.interviews))
    for i in sch.interviews:
        log.info("%s %s %s %s -> %s", i.cid, ROLE[i.job_id][:22], f"{i.start:%a %d %b %H:%M}" if i.start else "-",
                 i.mode or "-", i.status)
    return sch


def main():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    from auto_calling import run as run_calls
    from shortlist_candidates import SampleSourceWithContacts, consolidate
    user = authorise()
    cands, _ = consolidate(SampleSourceWithContacts().fetch())
    short = [c for c in cands if c["status"] == "Shortlisted"]
    calling = run_calls(short, user)
    run(cands, calling, user)


if __name__ == "__main__":
    main()
