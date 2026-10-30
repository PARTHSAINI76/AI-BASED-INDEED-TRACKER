"""
Task 3 - Applicant extraction & filtering: consolidated, access-controlled shortlist.

Design
------
Data in     : ONLY Indeed's permitted export (CSV from the employer dashboard) or the SAMPLE set.
              No scraping, no bots, no unofficial endpoints.
Data kept   : name, email, phone, location, relevant experience, skills, application date, status.
              Nothing else (no age, gender, photo, religion, marital status, etc.).
Consolidate : applications are linked by normalised email OR phone (union-find), so one candidate
              who applied twice, or to several jobs, gets ONE record.
Decision    : Shortlisted / On Hold / Not Shortlisted + priority + brief job-related reason.
Access      : run-time authorisation check, restricted file permissions, optional workbook password
              (from env, never in code), access log, and a masked export for wider sharing.

Usage
-----
  python shortlist_candidates.py                                    # sample data
  APPLICANT_SOURCE=csv APPLICANT_CSV=export.csv \
  AUTHORISED_USERS="hr.lead,recruiter1" REQUESTING_USER=recruiter1 \
  CANDIDATE_DB_PASSWORD='...' python shortlist_candidates.py        # real data
"""
from __future__ import annotations

import csv
import getpass
import logging
import os
import re
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.protection import WorkbookProtection
from openpyxl.worksheet.datavalidation import DataValidation

from build_tracker import JOBS
from categorise_applicants import JOB_CRITERIA, SampleSource, _skill_list, classify, get_source

log = logging.getLogger("shortlist")
DB_FILE = "Indeed_Candidate_Database.xlsx"
MASKED_FILE = "Indeed_Shortlist_Masked.xlsx"
ACCESS_LOG = "access_log.csv"

STATUSES = ["Shortlisted", "On Hold", "Not Shortlisted"]
PRIORITIES = ["P1 - High", "P2 - Medium", "P3 - Low", "N/A"]
P1_MIN_SCORE = 85  # assumption: strong matches get top priority


# ----------------------------------------------------------------------------- access control
def authorise() -> str:
    """Return the requesting user, or raise PermissionError."""
    user = os.getenv("REQUESTING_USER") or getpass.getuser()
    allowed = {u.strip() for u in os.getenv("AUTHORISED_USERS", "").split(",") if u.strip()}
    real_data = os.getenv("APPLICANT_SOURCE", "sample").lower() == "csv"
    if not allowed:
        if real_data:
            raise PermissionError("Real applicant data requires AUTHORISED_USERS to be set.")
        log.warning("AUTHORISED_USERS not set - allowed only because data is SAMPLE.")
    elif user not in allowed:
        raise PermissionError(f"User '{user}' is not authorised to process applicant data.")
    return user


def log_access(user: str, action: str, source: str, rows: int) -> None:
    new = not Path(ACCESS_LOG).exists()
    with open(ACCESS_LOG, "a", newline="") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["timestamp", "user", "action", "source", "records"])
        w.writerow([datetime.now().isoformat(timespec="seconds"), user, action, source, rows])
    os.chmod(ACCESS_LOG, 0o600)


# ----------------------------------------------------------------------------- sample data
class SampleSourceWithContacts(SampleSource):
    """SAMPLE data + fake phones + cross-job and repeat applications to exercise consolidation."""

    def fetch(self) -> list[dict]:
        apps = super().fetch()
        for a in apps:
            n = int(re.sub(r"\D", "", a["email"]) or 0)
            a["phone"] = f"+91-90000-{n:05d}"
        base = {a["email"]: a for a in apps}
        extra = []
        # s01 also applies to JP-003 (different job, same person)
        x = dict(base["s01@example.com"], job_id="JP-003", skills="Python;Selenium;SQL", experience_years=6)
        extra.append(x)
        # s06 re-applies with upper-case email and different phone format (same person, same job)
        x = dict(base["s06@example.com"], email="S06@Example.com", phone="90000 00006")
        extra.append(x)
        # s09 applies again to the same job using a new email but the same phone number
        x = dict(base["s09@example.com"], email="s09.alt@example.com")
        extra.append(x)
        for i, e in enumerate(extra, len(apps) + 1):
            e["applicant_id"] = f"AP-{i:03d}"
        return apps + extra


# ----------------------------------------------------------------------------- consolidation
def norm_email(v) -> str:
    return str(v or "").strip().lower()


def norm_phone(v) -> str:
    d = re.sub(r"\D", "", str(v or ""))
    return d[-10:] if len(d) >= 10 else ""


def group_candidates(apps: list[dict]) -> list[list[dict]]:
    """Union-find over shared email / phone. Apps with neither stay on their own."""
    parent = list(range(len(apps)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    seen: dict[tuple, int] = {}
    for i, a in enumerate(apps):
        for key in (("e", norm_email(a.get("email"))), ("p", norm_phone(a.get("phone")))):
            if not key[1]:
                continue
            if key in seen:
                parent[find(i)] = find(seen[key])
            else:
                seen[key] = i
    groups: dict[int, list[dict]] = {}
    for i, a in enumerate(apps):
        groups.setdefault(find(i), []).append(a)
    return list(groups.values())


def decide(best: dict) -> tuple[str, str, str]:
    """(status, priority, brief reason) from the best-scoring application."""
    crit = JOB_CRITERIA.get(best["job_id"])
    cat = best["category"]
    if crit is None or cat == "Incomplete Profile":
        return "On Hold", "P3 - Low", f"Profile incomplete ({best['reasons']}) - request missing details"
    have = _skill_list(best["skills"])
    matched = sum(1 for s in crit["skills"] if any(s in h or h in s for h in have))
    core = f"{matched}/{len(crit['skills'])} required skills, {best['experience_years']:g} yrs vs {crit['min_exp']} needed"
    flags = []
    sal, notice = best.get("expected_salary_lpa"), best.get("notice_days")
    if sal is not None and sal > crit["salary_max_lpa"]:
        flags.append(f"expected {sal:g} LPA above approved max {crit['salary_max_lpa']}")
    if notice is not None and notice > crit["max_notice_days"]:
        flags.append(f"notice {notice:g}d exceeds {crit['max_notice_days']}d")
    if not any(l in best["location"].lower() for l in crit["locations"]):
        flags.append(f"location {best['location']} outside job locations")
    extra = ("; " + "; ".join(flags)) if flags else ""
    if cat == "Relevant/Suitable":
        hard = [f for f in flags if "salary" in f or "notice" in f]
        if hard:
            return "On Hold", "P3 - Low", f"Strong fit ({core}) but {'; '.join(hard)} - confirm before proceeding"
        pr = "P1 - High" if best["score"] >= P1_MIN_SCORE else "P2 - Medium"
        return "Shortlisted", pr, f"Meets criteria: {core}{extra}"
    if cat == "Partially Relevant":
        return "On Hold", "P3 - Low", f"Partial fit: {core}{extra}"
    return "Not Shortlisted", "N/A", f"Does not meet core criteria: {core}{extra}"


def consolidate(apps: list[dict]) -> tuple[list[dict], list[dict]]:
    for a in apps:
        a.update(classify(a))
    candidates, links = [], []
    for g in group_candidates(apps):
        g.sort(key=lambda a: (a["score"], a["applied_date"] or date.min), reverse=True)
        best = g[0]

        def pick(field):  # first non-empty value, best application first
            return next((a[field] for a in g if a.get(field) not in (None, "")), "")

        jobs = sorted({a["job_id"] for a in g})
        status, prio, reason = decide(best)
        dates = [a["applied_date"] for a in g if a["applied_date"]]
        candidates.append({
            "name": pick("name"), "email": norm_email(pick("email")), "phone": pick("phone"),
            "location": pick("location"), "best_job": best["job_id"], "jobs": ", ".join(jobs),
            "n_apps": len(g), "n_dups": len(g) - len(jobs),
            "experience": max((a["experience_years"] for a in g if a.get("experience_years") is not None), default=""),
            "skills": "; ".join(sorted({s.title() for a in g for s in _skill_list(a["skills"])})),
            "first": min(dates) if dates else "", "last": max(dates) if dates else "",
            "score": best["score"], "status": status, "priority": prio, "reason": reason,
            "remarks": best["reasons"], "apps": g,
        })
    order = {p: i for i, p in enumerate(PRIORITIES)}
    candidates.sort(key=lambda c: (order[c["priority"]], -c["score"], c["name"]))
    for i, c in enumerate(candidates, 1):
        c["cid"] = f"CAND-{i:03d}"
        seen_jobs = set()
        for a in sorted(c["apps"], key=lambda a: a["applied_date"] or date.min):
            links.append({"applicant_id": a["applicant_id"], "cid": c["cid"], "job_id": a["job_id"],
                          "applied": a["applied_date"], "category": a["category"], "score": a["score"],
                          "repeat": "Yes" if a["job_id"] in seen_jobs else "No"})
            seen_jobs.add(a["job_id"])
    return candidates, links


# ----------------------------------------------------------------------------- workbook
FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
EDIT_FILL = PatternFill("solid", fgColor="FFF2CC")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def _hdr(ws, row, n):
    for c in range(1, n + 1):
        x = ws.cell(row=row, column=c)
        x.font = Font(name=FONT, bold=True, color="FFFFFF")
        x.fill = HEADER_FILL
        x.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        x.border = BORDER


def _body(ws, r, ncols):
    for c in range(1, ncols + 1):
        x = ws.cell(row=r, column=c)
        x.font = Font(name=FONT, size=10)
        x.alignment = Alignment(wrap_text=True, vertical="top")
        x.border = BORDER


def mask_email(e: str) -> str:
    if "@" not in e:
        return ""
    u, d = e.split("@", 1)
    return f"{u[:1]}***@{d}"


def mask_phone(p: str) -> str:
    d = norm_phone(p)
    return f"******{d[-4:]}" if d else ""


CAND_COLS = [("Candidate ID", 12), ("Name", 22), ("Email", 28), ("Phone", 17), ("Location", 14),
             ("Best-fit Job", 11), ("Jobs Applied", 16), ("Applications Received", 13),
             ("Duplicate (repeat) Applications", 14), ("Relevant Experience (yrs)", 12),
             ("Key Skills", 38), ("First Applied", 13), ("Latest Applied", 13), ("Match Score", 9),
             ("Status", 17), ("Shortlist Priority", 14), ("Reason (job-related)", 55),
             ("Screening Remarks", 60), ("Reviewer Notes", 28), ("Last Updated", 13)]
STATUS_COL, PRIO_COL, NOTES_COL = "O", "P", "S"


def build_db(cands, links, raw_count, user, protect_pw):
    wb = Workbook()
    ws = wb.active
    ws.title = "Candidates"
    for c, (h, w) in enumerate(CAND_COLS, 1):
        ws.cell(row=1, column=c, value=h)
        ws.column_dimensions[get_column_letter(c)].width = w
    _hdr(ws, 1, len(CAND_COLS))
    today = date.today()
    for r, c in enumerate(cands, 2):
        vals = [c["cid"], c["name"], c["email"], c["phone"], c["location"], c["best_job"], c["jobs"],
                c["n_apps"], c["n_dups"], c["experience"], c["skills"], c["first"], c["last"], c["score"],
                c["status"], c["priority"], c["reason"], c["remarks"], None, today]
        for i, v in enumerate(vals, 1):
            ws.cell(row=r, column=i, value=v)
        _body(ws, r, len(CAND_COLS))
        for col in (12, 13, 20):
            ws.cell(row=r, column=col).number_format = "dd-mmm-yyyy"
        for col in (15, 16, 19):  # reviewer-editable cells
            ws.cell(row=r, column=col).fill = EDIT_FILL
            ws.cell(row=r, column=col).protection = Protection(locked=False)
        ws.row_dimensions[r].height = 48
    last = 1 + len(cands)
    dv_s = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=False)
    dv_p = DataValidation(type="list", formula1='"' + ",".join(PRIORITIES) + '"', allow_blank=False)
    ws.add_data_validation(dv_s)
    ws.add_data_validation(dv_p)
    dv_s.add(f"{STATUS_COL}2:{STATUS_COL}{last}")
    dv_p.add(f"{PRIO_COL}2:{PRIO_COL}{last}")
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(CAND_COLS))}{last}"

    # Applications (traceability of merged duplicates)
    wa = wb.create_sheet("Applications")
    heads = [("Source Applicant ID", 18), ("Candidate ID", 13), ("Job", 9), ("Applied Date", 13),
             ("Task 2 Category", 20), ("Score", 8), ("Repeat for same job?", 16)]
    for c, (h, w) in enumerate(heads, 1):
        wa.cell(row=1, column=c, value=h)
        wa.column_dimensions[get_column_letter(c)].width = w
    _hdr(wa, 1, len(heads))
    for r, l in enumerate(sorted(links, key=lambda x: x["cid"]), 2):
        for i, v in enumerate([l["applicant_id"], l["cid"], l["job_id"], l["applied"], l["category"], l["score"], l["repeat"]], 1):
            wa.cell(row=r, column=i, value=v)
        _body(wa, r, len(heads))
        wa.cell(row=r, column=4).number_format = "dd-mmm-yyyy"
    wa.freeze_panes = "A2"
    la = 1 + len(links)

    # Summary (formulas)
    s = wb.create_sheet("Summary", 0)
    s.column_dimensions["A"].width = 46
    for col in "BCDE":
        s.column_dimensions[col].width = 16
    s["A1"] = "Indeed Candidate Shortlist - Summary"
    s["A1"].font = Font(name=FONT, bold=True, size=14)
    s["A2"] = f"Generated {today:%d-%b-%Y} by {user}. CONFIDENTIAL - authorised recruitment team only."
    s["A2"].font = Font(name=FONT, italic=True, size=9)
    rows = [("Applications received (before consolidation)", f"=COUNTA(Applications!A2:A{la})"),
            ("Unique candidates (after consolidation)", f"=COUNTA(Candidates!A2:A{last})"),
            ("Duplicate / repeat records merged", "=B4-B5")]
    for i, (k, f) in enumerate(rows, 4):
        s.cell(row=i, column=1, value=k).font = Font(name=FONT, bold=True)
        s.cell(row=i, column=2, value=f).font = Font(name=FONT, bold=True)
    s["A8"], s["B8"], s["C8"] = "Status", "Candidates", "% of total"
    _hdr(s, 8, 3)
    r = 9
    for st in STATUSES:
        s.cell(row=r, column=1, value=st)
        s.cell(row=r, column=2, value=f"=COUNTIF(Candidates!{STATUS_COL}2:{STATUS_COL}{last},A{r})")
        s.cell(row=r, column=3, value=f"=IF($B$5=0,0,B{r}/$B$5)").number_format = "0.0%"
        for c in range(1, 4):
            s.cell(row=r, column=c).font = Font(name=FONT)
            s.cell(row=r, column=c).border = BORDER
        r += 1
    s.cell(row=r, column=1, value="Check: statuses sum to unique candidates").font = Font(name=FONT, italic=True, size=9)
    s.cell(row=r, column=2, value='=IF(SUM(B9:B11)=B5,"OK","MISMATCH")').font = Font(name=FONT, italic=True, size=9)
    r += 2
    s.cell(row=r, column=1, value="Shortlist priority"); s.cell(row=r, column=2, value="Candidates")
    _hdr(s, r, 2)
    for p in PRIORITIES:
        r += 1
        s.cell(row=r, column=1, value=p)
        s.cell(row=r, column=2, value=f"=COUNTIF(Candidates!{PRIO_COL}2:{PRIO_COL}{last},A{r})")
        for c in (1, 2):
            s.cell(row=r, column=c).font = Font(name=FONT)
            s.cell(row=r, column=c).border = BORDER
    r += 2
    heads = ["Best-fit job"] + STATUSES + ["Total"]
    for c, h in enumerate(heads, 1):
        s.cell(row=r, column=c, value=h)
    _hdr(s, r, len(heads))
    for j in JOBS:
        r += 1
        s.cell(row=r, column=1, value=f"{j['id']} - {j['title']}")
        for c, st in enumerate(STATUSES, 2):
            s.cell(row=r, column=c, value=f'=COUNTIFS(Candidates!$F$2:$F${last},"{j["id"]}",Candidates!${STATUS_COL}$2:${STATUS_COL}${last},"{st}")')
        s.cell(row=r, column=5, value=f"=SUM(B{r}:D{r})")
        for c in range(1, 6):
            s.cell(row=r, column=c).font = Font(name=FONT)
            s.cell(row=r, column=c).border = BORDER
            s.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")

    # Access & handling
    h = wb.create_sheet("Access & Handling")
    h.column_dimensions["A"].width = 120
    allowed = os.getenv("AUTHORISED_USERS", "(not set - sample data only)")
    lines = [
        ("Access & data handling", True),
        (f"Authorised users: {allowed}", False),
        ("Source: Indeed's permitted employer-dashboard export only. No scraping, bots or unofficial access.", False),
        ("Captured: name, contact details, location, relevant experience, skills, application date, status, screening reason.", False),
        ("Not captured/used: age, gender, religion, marital status, photo, health, or any non-job-related attribute.", False),
        ("Sharing: use Indeed_Shortlist_Masked.xlsx (contacts masked) for hiring managers; never email the full database.", False),
        ("Protection: this file is saved owner-only; sheets are locked except yellow reviewer cells. Excel sheet protection is a guard against accidental edits, not strong security - keep the file on an access-controlled drive.", False),
        ("Retention: delete or anonymise candidate records after the period set by your company policy and applicable data-protection law (e.g. India's DPDP Act) - confirm with HR/legal.", False),
        ("Audit: each run is recorded in access_log.csv (time, user, records).", False),
    ]
    for i, (t, b) in enumerate(lines, 1):
        x = h.cell(row=i, column=1, value=t)
        x.font = Font(name=FONT, bold=b, size=11 if b else 10)
        x.alignment = Alignment(wrap_text=True, vertical="top")

    if protect_pw:
        for sh in wb.worksheets:
            sh.protection.sheet = True
            sh.protection.set_password(protect_pw)
            sh.protection.autoFilter = False   # allow filtering/sorting UI
            sh.protection.sort = False
        wb.security = WorkbookProtection(workbookPassword=protect_pw, lockStructure=True)
    wb.save(DB_FILE)
    os.chmod(DB_FILE, 0o600)


def build_masked(cands):
    wb = Workbook()
    ws = wb.active
    ws.title = "Shortlist (masked)"
    heads = [("Candidate ID", 13), ("Name", 22), ("Email (masked)", 26), ("Phone (masked)", 16),
             ("Location", 14), ("Best-fit Job", 11), ("Experience (yrs)", 11), ("Key Skills", 38),
             ("Status", 15), ("Priority", 13), ("Reason (job-related)", 60)]
    for c, (h, w) in enumerate(heads, 1):
        ws.cell(row=1, column=c, value=h)
        ws.column_dimensions[get_column_letter(c)].width = w
    _hdr(ws, 1, len(heads))
    r = 1
    for c in cands:
        if c["status"] == "Not Shortlisted":
            continue
        r += 1
        for i, v in enumerate([c["cid"], c["name"], mask_email(c["email"]), mask_phone(c["phone"]),
                               c["location"], c["best_job"], c["experience"], c["skills"],
                               c["status"], c["priority"], c["reason"]], 1):
            ws.cell(row=r, column=i, value=v)
        _body(ws, r, len(heads))
        ws.row_dimensions[r].height = 40
    ws.freeze_panes = "C2"
    wb.save(MASKED_FILE)
    os.chmod(MASKED_FILE, 0o600)


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    user = authorise()
    if os.getenv("APPLICANT_SOURCE", "sample").lower() == "csv":
        src, label = get_source(), "csv-export"
    else:
        src, label = SampleSourceWithContacts(), "sample"
    apps = src.fetch()
    cands, links = consolidate(apps)
    pw = os.getenv("CANDIDATE_DB_PASSWORD", "")
    if not pw:
        log.warning("CANDIDATE_DB_PASSWORD not set - workbook saved without sheet protection.")
    build_db(cands, links, len(apps), user, pw)
    build_masked(cands)
    log_access(user, "generate_shortlist", label, len(apps))
    counts = {s: sum(1 for c in cands if c["status"] == s) for s in STATUSES}
    log.info("%d applications -> %d unique candidates | %s", len(apps), len(cands),
             " | ".join(f"{k}: {v}" for k, v in counts.items()))
    return cands


if __name__ == "__main__":
    main()
