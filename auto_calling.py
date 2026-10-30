from __future__ import annotations

import csv
import logging
import os
import re
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from shortlist_candidates import FONT, _body, _hdr, authorise, log_access

log = logging.getLogger("calling")

IST = timezone(timedelta(hours=5, minutes=30))
CALL_WINDOW = (time(9, 0), time(20, 0))     # [start, end) in IST
MAX_ATTEMPTS = 3
MIN_GAP = timedelta(hours=4)
MAX_BATCH = int(os.getenv("CALL_MAX_BATCH", "10"))
OUTCOMES = ["Answered", "No Answer", "Busy", "Invalid Number",
            "Callback Requested", "Interested", "Not Interested"]
RETRY_OUTCOMES = {"No Answer", "Busy"}
DTMF_OUTCOME = {"1": "Interested", "2": "Not Interested", "3": "Callback Requested"}
OPT_OUT_KEY = "9"
OUTPUT_FILE = "Calling_Prototype.xlsx"
OPT_OUT_FILE = "opt_out_register.csv"

IVR_SCRIPT = ("Hello, this is [Company Name] recruitment about your application for [Role] on Indeed. "
              "Press 1 if you are interested, 2 if not interested, 3 to get a call back from a recruiter, "
              "or 9 to stop receiving automated calls. (Hindi/English versions to be approved.)")


# ----------------------------------------------------------------------------- number validation
def validate_indian_mobile(raw) -> tuple[bool, str, str]:
    """Return (ok, e164, reason). Accepts +91.., 0091.., 91.., 0.. and plain 10-digit forms."""
    s = re.sub(r"[\s\-().]", "", str(raw or ""))
    if not s:
        return False, "", "empty number"
    if s.startswith("+"):
        if not s.startswith("+91"):
            return False, "", "not an Indian (+91) number"
        s = s[3:]
    elif s.startswith("0091"):
        s = s[4:]
    elif len(s) == 12 and s.startswith("91"):
        s = s[2:]
    elif len(s) == 11 and s.startswith("0"):
        s = s[1:]
    if not s.isdigit():
        return False, "", "contains non-digit characters"
    if len(s) != 10:
        return False, "", f"expected 10 digits, found {len(s)}"
    if s[0] not in "6789":
        return False, "", "Indian mobile numbers start with 6, 7, 8 or 9"
    if len(set(s)) == 1:
        return False, "", "repeated-digit pattern"
    return True, "+91" + s, ""


def mask(e164: str) -> str:
    return f"+91 ******{e164[-4:]}" if e164 else ""


# ----------------------------------------------------------------------------- registers
@dataclass
class Consent:
    cid: str
    given: bool
    timestamp: str
    channel: str
    scope: str = "Recruitment calls/messages for the applied role"
    withdrawn: bool = False


class ConsentRegister:
    def __init__(self, records: list[Consent] | None = None):
        self.records = {c.cid: c for c in (records or [])}

    def add(self, c: Consent):
        self.records[c.cid] = c

    def valid(self, cid: str) -> tuple[bool, str]:
        c = self.records.get(cid)
        if c is None:
            return False, "no consent on record"
        if c.withdrawn:
            return False, "consent withdrawn"
        if not c.given:
            return False, "consent not given"
        return True, ""


class OptOutRegister:
    """Permanent suppression list keyed by E.164 number. Persists to CSV when a path is given."""

    def __init__(self, path: str | None = None):
        self.path = path
        self.entries: dict[str, dict] = {}
        if path and Path(path).exists():
            with open(path, newline="") as f:
                for r in csv.DictReader(f):
                    self.entries[r["number"]] = r

    def add(self, number: str, cid: str, ts: datetime, source: str):
        if number in self.entries:
            return
        self.entries[number] = {"number": number, "candidate_id": cid,
                                "timestamp": ts.isoformat(timespec="seconds"), "source": source}
        if self.path:
            new = not Path(self.path).exists()
            with open(self.path, "a", newline="") as f:
                w = csv.DictWriter(f, fieldnames=["number", "candidate_id", "timestamp", "source"])
                if new:
                    w.writeheader()
                w.writerow(self.entries[number])
            os.chmod(self.path, 0o600)

    def __contains__(self, number: str) -> bool:
        return number in self.entries


class DNDRegistry(ABC):
    @abstractmethod
    def is_registered(self, e164: str) -> bool: ...


class MockDND(DNDRegistry):
    def __init__(self, numbers: set[str] | None = None):
        self.numbers = numbers or set()

    def is_registered(self, e164: str) -> bool:
        return e164 in self.numbers


# ----------------------------------------------------------------------------- providers
@dataclass
class CallResult:
    status: str              # Answered | No Answer | Busy | Invalid Number
    dtmf: str | None = None  # key pressed in the IVR, if answered


class CallProvider(ABC):
    @abstractmethod
    def place_call(self, e164: str, script: str) -> CallResult: ...


class MockProvider(CallProvider):
    """Simulator. `scripts` maps number -> list of tokens consumed one per attempt:
    '1','2','3','9' (key presses), 'Answered', 'No Answer', 'Busy', 'Invalid Number'."""

    def __init__(self, scripts: dict[str, list[str]]):
        self.scripts = {k: list(v) for k, v in scripts.items()}
        self.calls = 0

    def place_call(self, e164: str, script: str) -> CallResult:
        self.calls += 1
        queue = self.scripts.get(e164) or ["No Answer"]
        tok = queue.pop(0) if len(queue) > 1 else queue[0]
        if tok in DTMF_OUTCOME or tok == OPT_OUT_KEY:
            return CallResult("Answered", tok)
        return CallResult(tok)


class LiveProvider(CallProvider):
    """Skeleton for an approved telephony vendor (REST call + signed status/DTMF webhooks)."""

    def __init__(self):
        self.key, self.secret = os.getenv("TELEPHONY_API_KEY", ""), os.getenv("TELEPHONY_API_SECRET", "")
        if not (self.key and self.secret):
            raise PermissionError("Live provider needs TELEPHONY_API_KEY and TELEPHONY_API_SECRET.")

    def place_call(self, e164: str, script: str) -> CallResult:
        # TODO: call the vendor's outbound-call API with the approved IVR flow, then receive the
        #       status + DTMF through the vendor's webhook and return it here.
        raise NotImplementedError("Implement against the vendor chosen after approval.")


# ----------------------------------------------------------------------------- policy + gate
@dataclass
class Policy:
    mode: str = "test"                      # test | pilot | live
    window: tuple[time, time] = CALL_WINDOW
    max_attempts: int = MAX_ATTEMPTS
    min_gap: timedelta = MIN_GAP
    approved_test_numbers: set[str] = field(default_factory=set)
    approval_id: str = ""
    approved_by: str = ""
    max_batch: int = MAX_BATCH


@dataclass
class Contact:
    cid: str
    name: str
    phone: str


def gate(contact: Contact, pol: Policy, consent: ConsentRegister, optout: OptOutRegister,
         dnd: DNDRegistry, history: list[tuple[datetime, str]], now: datetime) -> tuple[str, str]:
    """Return (decision, reason). decision: ALLOW | BLOCK (permanent) | DEFER (try later) | INVALID."""
    ok, e164, why = validate_indian_mobile(contact.phone)
    if pol.mode == "test" and (not ok or e164 not in pol.approved_test_numbers):
        return "BLOCK", "not an approved test number"
    if ok and e164 in optout:
        return "BLOCK", "candidate opted out"
    good, reason = consent.valid(contact.cid)
    if not good:
        return "BLOCK", reason
    if not ok:
        return "INVALID", why
    if dnd.is_registered(e164):
        return "BLOCK", "number on DND registry - use human/email outreach"
    if len(history) >= pol.max_attempts:
        return "BLOCK", f"attempt limit reached ({pol.max_attempts})"
    t = now.astimezone(IST).time()
    if not (pol.window[0] <= t < pol.window[1]):
        return "DEFER", f"outside calling hours {pol.window[0]:%H:%M}-{pol.window[1]:%H:%M} IST"
    if history and now - history[-1][0] < pol.min_gap:
        return "DEFER", f"less than {int(pol.min_gap.total_seconds() // 3600)} h since last attempt"
    return "ALLOW", ""


# ----------------------------------------------------------------------------- campaign
class CallCampaign:
    def __init__(self, run_type: str, provider: CallProvider, policy: Policy, consent: ConsentRegister,
                 optout: OptOutRegister, dnd: DNDRegistry, contacts: list[Contact]):
        if policy.mode == "live":
            if not (policy.approval_id and policy.approved_by):
                raise PermissionError("Live calling needs CALL_APPROVAL_ID and CALL_APPROVED_BY (written approval).")
            if len(contacts) > policy.max_batch:
                raise PermissionError(f"Batch of {len(contacts)} exceeds max {policy.max_batch}; split and re-approve.")
        self.run_type, self.provider, self.policy = run_type, provider, policy
        self.consent, self.optout, self.dnd, self.contacts = consent, optout, dnd, contacts
        self.log: list[dict] = []
        self.callbacks: list[dict] = []
        self.state = {c.cid: {"history": [], "final": None, "reason": ""} for c in contacts}

    def _log(self, now, c: Contact, attempt: int, result: str, detail: str):
        self.log.append({"id": f"{self.run_type[0]}-{len(self.log) + 1:03d}", "run": self.run_type,
                         "time": now.astimezone(IST).replace(tzinfo=None), "cid": c.cid,
                         "phone": mask(validate_indian_mobile(c.phone)[1]), "attempt": attempt,
                         "result": result, "detail": detail})

    def run_wave(self, now: datetime) -> None:
        for c in self.contacts:
            st = self.state[c.cid]
            if st["final"]:
                continue
            decision, reason = gate(c, self.policy, self.consent, self.optout, self.dnd, st["history"], now)
            n = len(st["history"])
            if decision == "INVALID":
                self._log(now, c, n, "Invalid Number", reason)
                st.update(final="Invalid Number", reason=reason)
            elif decision == "BLOCK":
                self._log(now, c, n, "Blocked", reason)
                st.update(final="Blocked", reason=reason)
            elif decision == "DEFER":
                self._log(now, c, n, "Deferred", reason)
            else:
                self._dial(c, st, now)

    def _dial(self, c: Contact, st: dict, now: datetime) -> None:
        e164 = validate_indian_mobile(c.phone)[1]
        res = self.provider.place_call(e164, IVR_SCRIPT)
        attempt = len(st["history"]) + 1
        detail = ""
        if res.status == "Answered":
            if res.dtmf == OPT_OUT_KEY:
                outcome, detail = "Not Interested", "Pressed 9 - OPTED OUT, number suppressed permanently"
                self.optout.add(e164, c.cid, now, "IVR key 9")
            elif res.dtmf in DTMF_OUTCOME:
                outcome, detail = DTMF_OUTCOME[res.dtmf], f"Pressed {res.dtmf}"
                if outcome == "Callback Requested":
                    self.callbacks.append({"cid": c.cid, "name": c.name, "requested": now,
                                           "action": "Human recruiter to call back within 1 working day"})
            else:
                outcome, detail = "Answered", "Picked up, no key pressed - human follow-up needed"
                self.callbacks.append({"cid": c.cid, "name": c.name, "requested": now,
                                       "action": "Answered without response - recruiter follow-up"})
        else:
            outcome = res.status
        st["history"].append((now, outcome))
        self._log(now, c, attempt, outcome, detail)
        if outcome in RETRY_OUTCOMES:
            if attempt >= self.policy.max_attempts:
                st.update(final=outcome, reason="max attempts reached - route to email/human outreach")
        else:
            st.update(final=outcome, reason=detail)

    def final_status(self, cid: str) -> str:
        st = self.state[cid]
        return st["final"] or (st["history"][-1][1] if st["history"] else "Not attempted")


def run_waves(campaign: CallCampaign, times: list[datetime]) -> None:
    for t in times:
        campaign.run_wave(t)


# ----------------------------------------------------------------------------- scenarios
def _dt(h: int, m: int = 0, day: date | None = None) -> datetime:
    return datetime.combine(day or date.today(), time(h, m), tzinfo=IST)


def waves(day: date | None = None) -> list[datetime]:
    return [_dt(10, 0, day), _dt(14, 30, day), _dt(19, 0, day)]


TEST_NUMBERS = {  # approved TEST numbers (fictitious) -> scripted behaviour per attempt
    "+919000000101": ["Answered"],
    "+919000000102": ["No Answer"],
    "+919000000103": ["Busy", "1"],
    "+919000000104": ["Invalid Number"],
    "+919000000105": ["3"],
    "+919000000106": ["1"],
    "+919000000107": ["2"],
    "+919000000108": ["9"],
}
TEST_DND = "+919000000109"


def make_test_run(optout: OptOutRegister) -> CallCampaign:
    contacts = [Contact(f"TEST-{i}", f"Test Contact {i}", n) for i, n in enumerate(TEST_NUMBERS, 1)]
    contacts.append(Contact("TEST-9", "Test DND Contact", TEST_DND))
    contacts.append(Contact("TEST-X", "Not-approved Number", "+919123456789"))  # must be blocked
    consent = ConsentRegister([Consent(c.cid, True, "TEST harness", "test") for c in contacts])
    pol = Policy(mode="test", approved_test_numbers=set(TEST_NUMBERS) | {TEST_DND})
    camp = CallCampaign("TEST", MockProvider(TEST_NUMBERS), pol, consent, optout, MockDND({TEST_DND}), contacts)
    run_waves(camp, waves())
    return camp


PILOT_SCRIPTS = {   # simulated behaviour for the sample shortlisted candidates (by phone)
    "+919000000001": ["1"],
    "+919000000006": ["No Answer", "3"],
    "+919000000011": ["Busy", "1"],
    "+919000000002": ["1"],
}


def make_pilot_run(shortlisted: list[dict], optout: OptOutRegister) -> CallCampaign:
    contacts = [Contact(c["cid"], c["name"], c["phone"]) for c in shortlisted]
    consent = ConsentRegister()
    # SAMPLE consent: first three gave consent on the application form; the last has none on record.
    for c in contacts[:-1]:
        consent.add(Consent(c.cid, True, f"{date.today():%Y-%m-%d} 09:00 IST", "Indeed application (sample)"))
    pol = Policy(mode="pilot")
    camp = CallCampaign("PILOT", MockProvider(PILOT_SCRIPTS), pol, consent, optout, MockDND(), contacts)
    run_waves(camp, waves())
    return camp


# ----------------------------------------------------------------------------- tests
def run_safeguard_tests(test_run: CallCampaign) -> list[tuple]:
    rows: list[tuple] = []

    def add(area, scenario, expected, actual):
        rows.append((f"T{len(rows) + 1:02d}", area, scenario, str(expected), str(actual),
                     "PASS" if str(expected) == str(actual) else "FAIL"))

    for raw, exp in [("+919876543210", "+919876543210"), ("09876543210", "+919876543210"),
                     ("919876543210", "+919876543210"), ("98765 43210", "+919876543210"),
                     ("0091-98765-43210", "+919876543210")]:
        add("Number format", f"Valid form '{raw}'", exp, validate_indian_mobile(raw)[1])
    for raw in ["5876543210", "987654321", "+14155550100", "0000000000", "9999999999", "98765abcde", ""]:
        add("Number format", f"Invalid '{raw}' rejected", "False", validate_indian_mobile(raw)[0])

    def one(pol_kw=None, consent=True, withdrawn=False, tokens=None, number="+919000000201"):
        c = Contact("T-1", "T", number)
        reg = ConsentRegister([Consent("T-1", consent, "t", "t", withdrawn=withdrawn)] if consent is not None else [])
        pol = Policy(mode="pilot", **(pol_kw or {}))
        camp = CallCampaign("TEST", MockProvider({number: tokens or ["No Answer"]}), pol, reg,
                            OptOutRegister(), MockDND(), [c])
        return camp, c

    d = date(2026, 11, 2)
    for h, m, exp in [(8, 59, "DEFER"), (9, 0, "ALLOW"), (19, 59, "ALLOW"), (20, 0, "DEFER")]:
        camp, c = one()
        add("Calling hours", f"Call at {h:02d}:{m:02d} IST", exp,
            gate(c, camp.policy, camp.consent, camp.optout, camp.dnd, [], _dt(h, m, d))[0])

    camp, c = one(consent=None)
    add("Consent", "No consent on record -> blocked", "BLOCK",
        gate(c, camp.policy, camp.consent, camp.optout, camp.dnd, [], _dt(10, 0, d))[0])
    camp, c = one(withdrawn=True)
    add("Consent", "Withdrawn consent -> blocked", "BLOCK",
        gate(c, camp.policy, camp.consent, camp.optout, camp.dnd, [], _dt(10, 0, d))[0])

    camp, c = one(tokens=["9", "1"])
    camp.run_wave(_dt(10, 0, d))
    camp.run_wave(_dt(15, 0, d))
    add("Opt-out", "Press 9 -> number suppressed", "True", validate_indian_mobile(c.phone)[1] in camp.optout)
    add("Opt-out", "No further call after opt-out", "1", camp.provider.calls)

    camp, c = one(tokens=["No Answer"])
    run_waves(camp, [_dt(10, 0, d), _dt(14, 30, d), _dt(19, 0, d), _dt(10, 0, date(2026, 11, 3))])
    add("Frequency", "4th attempt refused (limit 3)", "3", camp.provider.calls)
    camp, c = one(tokens=["Busy", "1"])
    run_waves(camp, [_dt(10, 0, d), _dt(11, 0, d)])
    add("Frequency", "Retry after 1 h refused (min gap 4 h)", "1", camp.provider.calls)

    camp, c = one(tokens=["3"])
    camp.run_wave(_dt(10, 0, d))
    add("Human option", "Press 3 -> callback task created", "1", len(camp.callbacks))

    pol = Policy(mode="test", approved_test_numbers=set(TEST_NUMBERS))
    camp = CallCampaign("TEST", MockProvider({}), pol, ConsentRegister([Consent("T-1", True, "t", "t")]),
                        OptOutRegister(), MockDND(), [Contact("T-1", "T", "+919123456789")])
    add("Test-number guard", "Non-approved number blocked in TEST mode", "BLOCK",
        gate(camp.contacts[0], pol, camp.consent, camp.optout, camp.dnd, [], _dt(10, 0, d))[0])

    for scen, kw, n_contacts in [("Live mode without approval", {}, 1),
                                 ("Live batch above cap", {"approval_id": "A1", "approved_by": "HR"}, 99)]:
        try:
            CallCampaign("LIVE", MockProvider({}), Policy(mode="live", **kw), ConsentRegister(), OptOutRegister(),
                         MockDND(), [Contact(f"C{i}", "x", "+919000000301") for i in range(n_contacts)])
            actual = "allowed"
        except PermissionError:
            actual = "refused"
        add("Live guard", f"{scen} refused", "refused", actual)

    try:
        os.environ.pop("TELEPHONY_API_KEY", None)
        LiveProvider()
        actual = "allowed"
    except PermissionError:
        actual = "refused"
    add("Live guard", "Live provider without credentials refused", "refused", actual)

    seen = {r["result"] for r in test_run.log}
    add("Outcomes", "TEST run produced all 7 required outcomes", "True", all(o in seen for o in OUTCOMES))
    add("DND", "DND-registered number not called", "Blocked",
        next(r["result"] for r in test_run.log if r["cid"] == "TEST-9"))
    return rows


# ----------------------------------------------------------------------------- workbook
PLAN_ROWS = [
    ("Platform", "Proposed calling platform",
     "Cloud telephony/IVR provider with Indian numbering, DLT-compliant routing and webhooks. Candidates to evaluate: ",
     "Exotel, Ozonetel, Knowlarity, Twilio (verify current India compliance support and pricing). Final choice needs IT-security and Legal sign-off."),
    ("Platform", "Integration method",
     "Our service calls the provider's outbound-call REST API with the approved IVR flow; the provider returns call status and key presses ",
     "through signed HTTPS webhooks, which write the call log. Credentials only via environment/secret manager. Prototype: provider-agnostic ",
     "CallProvider interface, MockProvider (tested), LiveProvider (skeleton, blocked without approval)."),
    ("Workflow", "1. Input", "Shortlisted candidates from Task 3 (one consolidated record each)."),
    ("Workflow", "2. Consent check", "Call only candidates with consent on record for recruitment calls; no consent -> human/email outreach asking for consent."),
    ("Workflow", "3. Number validation", "Normalise to +91XXXXXXXXXX; reject non-Indian, wrong length, or numbers not starting 6-9 (Invalid Number)."),
    ("Workflow", "4. Suppression", "Skip opted-out numbers and DND-registered numbers (scrub hook)."),
    ("Workflow", "5. Limits", "Calling hours 09:00-20:00 IST; max 3 attempts; min 4 h between attempts."),
    ("Workflow", "6. Call", "Play approved IVR script: 1 Interested, 2 Not interested, 3 Human callback, 9 Stop calls."),
    ("Workflow", "7. Record", "Outcome logged: Answered, No Answer, Busy, Invalid Number, Callback Requested, Interested, Not Interested."),
    ("Workflow", "8. Route", "Interested -> Task 5 scheduling | Callback/Answered -> recruiter callback queue | Not interested/opt-out -> closed + suppressed | ",
     "No Answer/Busy -> retry within limits, then email | Invalid -> request corrected number by email."),
    ("Approvals", "HR Head", "Approves purpose, IVR script (language, wording) and candidate list."),
    ("Approvals", "Legal / Data-protection owner", "Approves consent wording, DND/DLT handling and retention; confirms applicable TRAI and DPDP Act requirements."),
    ("Approvals", "IT / Security", "Approves vendor, data flow, credential handling."),
    ("Approvals", "Written approval reference", "Live mode needs CALL_APPROVAL_ID + CALL_APPROVED_BY, batch cap enforced. No bulk automated calls without approval."),
    ("IVR script (draft)", "For approval", IVR_SCRIPT),
]
SAFEGUARDS = [
    ("Consent", "Calls only with recorded consent; withdrawn or missing consent blocks the call."),
    ("Opt-out", "Key 9 suppresses the number permanently (opt_out_register.csv) across runs."),
    ("Human option", "Key 3 creates a human-callback task; no-input answers are also routed to a recruiter."),
    ("Number validation", "Only valid Indian mobile numbers are dialled."),
    ("Hours", "09:00-20:00 IST window (conservative; confirm with Legal)."),
    ("Frequency", "Max 3 attempts, min 4 h gap; retries only for No Answer / Busy."),
    ("DND", "DND registry hook; registered numbers get human/email outreach, not automated calls."),
    ("Test first", "TEST mode dials only approved test numbers; PILOT/LIVE only after tests pass."),
    ("Approval gate", "Live mode refuses to start without written approval ref; batch size capped."),
    ("Privacy", "Call log stores masked numbers and candidate IDs only; files saved owner-only; access limited to AUTHORISED_USERS and logged."),
]


def _sheet(wb, title, heads, rows, widths=None):
    ws = wb.create_sheet(title)
    for c, h in enumerate(heads, 1):
        ws.cell(row=1, column=c, value=h)
        ws.column_dimensions[get_column_letter(c)].width = (widths or {}).get(c, 18)
    _hdr(ws, 1, len(heads))
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            ws.cell(row=r, column=c, value=v)
        _body(ws, r, len(heads))
    ws.freeze_panes = "A2"
    return ws


def write_workbook(test_c, pilot_c, tests, optout, pilot_contacts_meta, user) -> None:
    wb = Workbook()
    wb.remove(wb.active)
    # Summary
    s = wb.create_sheet("Summary")
    s.column_dimensions["A"].width = 36
    for col in "BCD":
        s.column_dimensions[col].width = 14
    s["A1"] = "Automated Calling Prototype - Summary"
    s["A1"].font = Font(name=FONT, bold=True, size=14)
    s["A2"] = f"SIMULATION - no real calls were placed. Generated {date.today():%d-%b-%Y} by {user}."
    s["A2"].font = Font(name=FONT, italic=True, size=9)
    n_log = len(test_c.log) + len(pilot_c.log) + 1
    for c, h in enumerate(["Result", "TEST run", "PILOT run", "Total"], 1):
        s.cell(row=4, column=c, value=h)
    _hdr(s, 4, 4)
    results = OUTCOMES + ["Blocked", "Deferred"]
    for i, o in enumerate(results, 5):
        s.cell(row=i, column=1, value=o)
        for c, run in ((2, "TEST"), (3, "PILOT")):
            s.cell(row=i, column=c, value=f'=COUNTIFS(\'Call Log\'!$B$2:$B${n_log},"{run}",\'Call Log\'!$G$2:$G${n_log},$A{i})')
        s.cell(row=i, column=4, value=f"=B{i}+C{i}")
        for c in range(1, 5):
            s.cell(row=i, column=c).font = Font(name=FONT)
    end = 4 + len(results)
    s.cell(row=end + 1, column=1, value="Total log entries").font = Font(name=FONT, bold=True)
    for c, col in ((2, "B"), (3, "C"), (4, "D")):
        s.cell(row=end + 1, column=c, value=f"=SUM({col}5:{col}{end})").font = Font(name=FONT, bold=True)
    s.cell(row=end + 2, column=1, value="Check vs call-log rows").font = Font(name=FONT, italic=True, size=9)
    s.cell(row=end + 2, column=4, value=f'=IF(D{end + 1}=COUNTA(\'Call Log\'!A2:A{n_log}),"OK","MISMATCH")').font = Font(name=FONT, italic=True, size=9)
    s.cell(row=end + 4, column=1, value="Safeguard tests passed").font = Font(name=FONT, bold=True)
    s.cell(row=end + 4, column=2, value=f"=COUNTIF('Test Results'!F2:F{len(tests) + 1},\"PASS\")")
    s.cell(row=end + 4, column=3, value=f"=COUNTA('Test Results'!A2:A{len(tests) + 1})")
    s.cell(row=end + 4, column=4, value="of tests")

    # Plan / Safeguards
    ws = _sheet(wb, "Plan", ["Section", "Item", "Detail"], PLAN_ROWS, {1: 18, 2: 28, 3: 120})
    for r in range(2, len(PLAN_ROWS) + 2):
        ws.row_dimensions[r].height = 48
    _sheet(wb, "Safeguards", ["Safeguard", "How it is enforced"], SAFEGUARDS, {1: 22, 2: 120})
    _sheet(wb, "Test Results", ["Test", "Area", "Scenario", "Expected", "Actual", "Result"], tests,
           {1: 7, 2: 18, 3: 52, 4: 18, 5: 18, 6: 9})

    # Call log (TEST then PILOT)
    rows = [(r["id"], r["run"], r["time"], r["cid"], r["phone"], r["attempt"], r["result"], r["detail"])
            for r in test_c.log + pilot_c.log]
    ws = _sheet(wb, "Call Log", ["Call ID", "Run Type", "Time (IST)", "Candidate ID", "Phone (masked)",
                                 "Attempt #", "Result", "Detail"], rows,
                {1: 9, 2: 10, 3: 18, 4: 13, 5: 17, 6: 9, 7: 20, 8: 70})
    for r in range(2, len(rows) + 2):
        ws.cell(row=r, column=3).number_format = "dd-mmm-yyyy hh:mm"
    # Pilot candidate status
    prow = []
    for c in pilot_c.contacts:
        st = pilot_c.state[c.cid]
        meta = pilot_contacts_meta[c.cid]
        nxt = {"Interested": "Schedule interview (Task 5)",
               "Callback Requested": "Recruiter callback, then schedule",
               "Blocked": "Manual outreach: request consent by email/Indeed message",
               "Not Interested": "Close", "Invalid Number": "Request corrected number"}.get(pilot_c.final_status(c.cid), "Recruiter follow-up")
        prow.append((c.cid, meta["priority"], meta["job"], len(st["history"]), pilot_c.final_status(c.cid),
                     st["reason"], nxt))
    _sheet(wb, "Candidate Status", ["Candidate ID", "Priority", "Job", "Attempts", "Final Outcome", "Note", "Next Action"],
           prow, {1: 13, 2: 13, 3: 9, 4: 9, 5: 20, 6: 55, 7: 50})
    cons = [(cid, "Yes" if c.given else "No", c.timestamp, c.channel, c.scope, "Yes" if c.withdrawn else "No")
            for cid, c in {**test_c.consent.records, **pilot_c.consent.records}.items()]
    _sheet(wb, "Consent Register", ["Candidate ID", "Consent given", "Recorded at", "Channel", "Scope", "Withdrawn"],
           cons, {3: 24, 4: 28, 5: 48})
    oo = [(e["candidate_id"], mask(e["number"]), e["timestamp"], e["source"]) for e in optout.entries.values()]
    _sheet(wb, "Opt-out Register", ["Candidate ID", "Number (masked)", "Opted out at", "Source"], oo, {2: 20, 3: 24})
    cb = [(x["cid"], x["name"], x["requested"].astimezone(IST).replace(tzinfo=None), x["action"], "Open")
          for x in test_c.callbacks + pilot_c.callbacks]
    ws = _sheet(wb, "Callback Queue", ["Candidate ID", "Name", "Requested at", "Action", "Status"], cb,
                {2: 22, 3: 18, 4: 52})
    for r in range(2, len(cb) + 2):
        ws.cell(row=r, column=3).number_format = "dd-mmm-yyyy hh:mm"
    wb.move_sheet("Summary", offset=-(len(wb.sheetnames) - 1))
    wb.save(OUTPUT_FILE)
    os.chmod(OUTPUT_FILE, 0o600)


def run(shortlisted: list[dict], user: str):
    """Run TEST then PILOT. Returns (pilot_campaign, meta). Pilot runs only if every test passes."""
    test_opt = OptOutRegister()                 # in-memory: keeps the TEST run repeatable
    optout = OptOutRegister(OPT_OUT_FILE)       # persistent: real suppression list for PILOT/LIVE
    test_c = make_test_run(test_opt)
    tests = run_safeguard_tests(test_c)
    failed = [t for t in tests if t[5] != "PASS"]
    if failed:
        raise RuntimeError(f"{len(failed)} safeguard test(s) failed - pilot not run: {failed[0]}")
    log.info("Safeguard tests: %d/%d passed", len(tests), len(tests))
    pilot_c = make_pilot_run(shortlisted, optout)
    meta = {c["cid"]: {"priority": c["priority"], "job": c["best_job"]} for c in shortlisted}
    merged = OptOutRegister()
    merged.entries = {**test_opt.entries, **optout.entries}
    write_workbook(test_c, pilot_c, tests, merged, meta, user)
    log_access(user, "calling_prototype", "sample", len(shortlisted))
    for c in pilot_c.contacts:
        log.info("%s -> %s", c.cid, pilot_c.final_status(c.cid))
    return pilot_c


def main():
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    from categorise_applicants import classify  # noqa: F401  (ensures module import works standalone)
    from shortlist_candidates import SampleSourceWithContacts, consolidate
    user = authorise()
    cands, _ = consolidate(SampleSourceWithContacts().fetch())
    short = [c for c in cands if c["status"] == "Shortlisted"]
    run(short, user)


if __name__ == "__main__":
    main()
