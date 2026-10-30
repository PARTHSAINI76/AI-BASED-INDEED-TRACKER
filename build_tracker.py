from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
INPUT_FILL = PatternFill("solid", fgColor="FFFF00")  # cells the user must fill in
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")

JOBS = [
    {
        "id": "JP-001",
        "title": "Senior Software Engineer (Backend - Python)",
        "type": "Full-time",
        "location": "Gorakhpur, Uttar Pradesh (Hybrid)",
        "experience": "4-7 years",
        "qualification": "B.Tech/B.E./MCA in Computer Science or related field",
        "salary": "₹12,00,000 - ₹20,00,000 per year (APPROVED RANGE - confirm with HR)",
        "skills": "Python, Django/FastAPI, REST APIs, PostgreSQL, Docker, Git, AWS, unit testing",
        "summary": (
            "We are looking for a Senior Backend Engineer to design, build and maintain "
            "scalable services that power our products. You will work with product and "
            "frontend teams and mentor junior developers."
        ),
        "responsibilities": (
            "- Design and develop secure, scalable REST APIs and backend services\n"
            "- Write clean, tested, well-documented code and perform code reviews\n"
            "- Optimise database queries and application performance\n"
            "- Collaborate with product, QA and DevOps on releases\n"
            "- Mentor junior engineers and contribute to technical design"
        ),
        "keywords": "software engineer, backend developer, python developer, django, fastapi, aws, rest api",
    },
    {
        "id": "JP-002",
        "title": "Frontend Developer (React.js)",
        "type": "Full-time",
        "location": "Gorakhpur, Uttar Pradesh (On-site)",
        "experience": "2-4 years",
        "qualification": "B.Tech/B.E./BCA/MCA or equivalent practical experience",
        "salary": "₹6,00,000 - ₹10,00,000 per year (APPROVED RANGE - confirm with HR)",
        "skills": "React.js, JavaScript (ES6+), TypeScript, HTML5, CSS3, Redux, REST integration, Git",
        "summary": (
            "Join our product team to build fast, accessible and responsive user interfaces. "
            "You will turn designs into production-quality components and work closely with "
            "designers and backend engineers."
        ),
        "responsibilities": (
            "- Build reusable, responsive UI components using React.js\n"
            "- Integrate frontend with REST APIs\n"
            "- Ensure cross-browser compatibility and accessibility\n"
            "- Write unit tests and fix UI bugs\n"
            "- Participate in sprint planning and code reviews"
        ),
        "keywords": "react developer, frontend developer, javascript, typescript, web developer, ui developer",
    },
    {
        "id": "JP-003",
        "title": "QA Automation Engineer",
        "type": "Full-time",
        "location": "Gorakhpur, Uttar Pradesh (Hybrid)",
        "experience": "2-5 years",
        "qualification": "B.Tech/B.E./MCA in Computer Science or related field",
        "salary": "₹5,00,000 - ₹9,00,000 per year (APPROVED RANGE - confirm with HR)",
        "skills": "Selenium/Playwright, Python or Java, API testing (Postman), CI/CD (Jenkins/GitHub Actions), SQL, test planning",
        "summary": (
            "We need a QA Automation Engineer to build and maintain automated test suites "
            "that keep our releases reliable and fast."
        ),
        "responsibilities": (
            "- Design, develop and maintain automated UI and API test suites\n"
            "- Create test plans and test cases from requirements\n"
            "- Integrate automated tests into CI/CD pipelines\n"
            "- Log, track and verify defects with the development team\n"
            "- Report test coverage and quality metrics"
        ),
        "keywords": "qa automation, test engineer, selenium, playwright, sdet, api testing, quality assurance",
    },
]

POSTING_COLS = [
    ("Ref ID", "id", 10),
    ("Job Title", "title", 34),
    ("Employment Type", "type", 16),
    ("Location", "location", 28),
    ("Experience", "experience", 14),
    ("Qualification", "qualification", 34),
    ("Salary (Approved)", "salary", 34),
    ("Key Skills", "skills", 40),
    ("Job Summary", "summary", 55),
    ("Responsibilities", "responsibilities", 60),
    ("SEO Keywords", "keywords", 40),
]

TRACKER_HEADERS = [
    ("Ref ID", 10), ("Job Title", 38), ("Posting Date", 14),
    ("Indeed Link / Reference", 42), ("Status", 20),
    ("Salary Approved By", 22), ("Proofread (Y/N)", 15), ("Remarks", 36),
]
STATUSES = ["Draft", "Ready to Publish", "Published", "Under Review", "Closed"]


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = BORDER


def build_postings(wb):
    ws = wb.active
    ws.title = "Postings"
    for c, (h, _, w) in enumerate(POSTING_COLS, 1):
        ws.cell(row=1, column=c, value=h)
        ws.column_dimensions[get_column_letter(c)].width = w
    style_header(ws, 1, len(POSTING_COLS))
    for r, job in enumerate(JOBS, 2):
        for c, (_, key, _) in enumerate(POSTING_COLS, 1):
            cell = ws.cell(row=r, column=c, value=job[key])
            cell.font = Font(name=FONT, size=10)
            cell.alignment = WRAP
            cell.border = BORDER
        ws.row_dimensions[r].height = 150
    ws.freeze_panes = "C2"


def build_tracker(wb):
    ws = wb.create_sheet("Tracker")
    ws["A1"] = "Indeed Job Posting Tracker"
    ws["A1"].font = Font(name=FONT, bold=True, size=14)
    # Summary block driven by formulas
    last = 5 + len(JOBS)
    ws["A2"] = "Total postings"
    ws["B2"] = f'=COUNTA(A6:A{last})'
    ws["C2"] = "Published"
    ws["D2"] = f'=COUNTIF(E6:E{last},"Published")'
    ws["E2"] = "Target met (>=3 published)?"
    ws["F2"] = '=IF(D2>=3,"YES","NO - pending")'
    for ref in ("A2", "C2", "E2"):
        ws[ref].font = Font(name=FONT, bold=True)
    for ref in ("B2", "D2", "F2"):
        ws[ref].font = Font(name=FONT)

    for c, (h, w) in enumerate(TRACKER_HEADERS, 1):
        ws.cell(row=5, column=c, value=h)
        ws.column_dimensions[get_column_letter(c)].width = w
    style_header(ws, 5, len(TRACKER_HEADERS))

    dv = DataValidation(type="list", formula1='"' + ",".join(STATUSES) + '"', allow_blank=True)
    yn = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    ws.add_data_validation(dv)
    ws.add_data_validation(yn)

    for i, job in enumerate(JOBS):
        r = 6 + i
        pr = 2 + i  # row in Postings sheet
        # Ref ID and title are linked to the Postings sheet (single source of truth)
        ws.cell(row=r, column=1, value=f"=Postings!A{pr}")
        ws.cell(row=r, column=2, value=f"=Postings!B{pr}")
        ws.cell(row=r, column=3, value=None)           # fill on publish (dd-mmm-yyyy)
        ws.cell(row=r, column=4, value=None)           # paste Indeed job URL / job key
        ws.cell(row=r, column=5, value="Ready to Publish")
        ws.cell(row=r, column=6, value=None)
        ws.cell(row=r, column=7, value="N")
        ws.cell(row=r, column=8, value="Pending Indeed employer account publish; salary range needs HR confirmation.")
        for c in range(1, len(TRACKER_HEADERS) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name=FONT, size=10)
            cell.alignment = WRAP
            cell.border = BORDER
        for c in (3, 4, 6):
            ws.cell(row=r, column=c).fill = INPUT_FILL
        ws.cell(row=r, column=3).number_format = "dd-mmm-yyyy"
        dv.add(ws.cell(row=r, column=5))
        yn.add(ws.cell(row=r, column=7))
        ws.row_dimensions[r].height = 36
    ws.freeze_panes = "A6"


def build_instructions(wb):
    ws = wb.create_sheet("Instructions")
    ws.column_dimensions["A"].width = 110
    lines = [
        ("How to use this tracker", True),
        ("1. Yellow cells on the 'Tracker' sheet are for you to fill in (posting date, Indeed link/reference, approver).", False),
        ("2. Copy each posting's fields from the 'Postings' sheet into Indeed (Employers > Post a job).", False),
        ("3. Confirm the salary range with HR before publishing; ranges shown are placeholders marked APPROVED RANGE - confirm.", False),
        ("4. Proofread (spelling, keywords, location, salary), then set 'Proofread' to Y.", False),
        ("5. After publishing, enter the posting date and the job URL (or Indeed job reference), then set Status = Published.", False),
        ("6. The summary at the top of 'Tracker' counts published postings automatically.", False),
        ("", False),
        ("Assumptions: job roles, locations and salary bands are illustrative drafts prepared by the assistant; replace with your company's real, approved details.", False),
        ("No links are pre-filled because the postings have not been published yet.", False),
    ]
    for i, (text, bold) in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=text)
        c.font = Font(name=FONT, bold=bold, size=12 if bold else 10)
        c.alignment = Alignment(wrap_text=True, vertical="top")


def main():
    wb = Workbook()
    build_postings(wb)
    build_tracker(wb)
    build_instructions(wb)
    wb.move_sheet("Tracker", offset=-1)  # Tracker first
    wb.save("Indeed_Job_Posting_Tracker.xlsx")
    print("Saved Indeed_Job_Posting_Tracker.xlsx")


if __name__ == "__main__":
    main()
