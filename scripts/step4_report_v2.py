"""
HFP AR Migration - Step 4 v2: Write hfp_Data_Analysis_v2.docx

Plain English document for the AR (Jim Hoban / Hoban Financial Planning Ltd).
No API / importer / UUID / cap language.
Rule 9: aggregates only, no PII.
"""
import os, json
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATA_ROOT    = r"C:\dev\AR Migration"
SUMMARY_PATH = os.path.join(DATA_ROOT, "step3_summary_v2.json")
OUT_PATH     = os.path.join(DATA_ROOT, "hfp_Data_Analysis_v2.docx")
TODAY        = "2026-10-14"

NAVY  = RGBColor(0x02, 0x08, 0x49)
BLACK = RGBColor(0x00, 0x00, 0x00)
GREY  = RGBColor(0x80, 0x80, 0x80)
SORA  = "Sora"

def p(*a): print(*a, flush=True)

with open(SUMMARY_PATH, encoding="utf-8") as f:
    S = json.load(f)

charges   = S["charges"]
plans_s   = S["plans"]
notes_s   = S["notes"]
deps_s    = S["dependants"]
files_s   = S["files"]
iae_s     = S["iae"]

# ── Doc helpers ───────────────────────────────────────────────────────────────

doc = Document()

def _set_font(run, size_pt=11, bold=False, italic=False, color=BLACK):
    run.font.name = SORA
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement("w:rFonts")
    rFonts.set(qn("w:asciiTheme"), "none")
    rFonts.set(qn("w:eastAsiaTheme"), "none")
    rFonts.set(qn("w:ascii"), SORA)
    rFonts.set(qn("w:hAnsi"), SORA)
    rFonts.set(qn("w:cs"), SORA)
    existing = rPr.find(qn("w:rFonts"))
    if existing is not None:
        rPr.remove(existing)
    rPr.insert(0, rFonts)

def heading(text, level=1):
    size = {1: 16, 2: 13, 3: 11}[level]
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(14 if level == 1 else 8)
    para.paragraph_format.space_after  = Pt(4)
    run = para.add_run(text)
    _set_font(run, size_pt=size, bold=True, color=NAVY)
    return para

def body(text, bold=False, italic=False, indent=False):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(4)
    if indent:
        para.paragraph_format.left_indent = Inches(0.3)
    run = para.add_run(text)
    _set_font(run, size_pt=11, bold=bold, italic=italic)
    return para

def bullet(text, bold_prefix=None):
    para = doc.add_paragraph(style="List Bullet")
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(2)
    for r in para.runs:
        r.clear()
    if bold_prefix:
        r1 = para.add_run(bold_prefix + ": ")
        _set_font(r1, size_pt=11, bold=True)
        r2 = para.add_run(text)
        _set_font(r2, size_pt=11)
    else:
        run = para.add_run(text)
        _set_font(run, size_pt=11)
    return para

def spacer():
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(0)
    run = para.add_run("")
    _set_font(run, size_pt=4)
    return para

def add_table(headers, rows_data, col_widths=None, zebra=False):
    table = doc.add_table(rows=1 + len(rows_data), cols=len(headers))
    table.style = "Table Grid"
    hdr_row = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.paragraphs[0].clear()
        run = cell.paragraphs[0].add_run(h)
        _set_font(run, size_pt=10, bold=True, color=NAVY)
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "EEF0F6")
        tcPr.append(shd)
    for r_idx, row in enumerate(rows_data, start=1):
        tr = table.rows[r_idx]
        fill = "F9F9F9" if (zebra and r_idx % 2 == 0) else "FFFFFF"
        for c_idx, val in enumerate(row):
            cell = tr.cells[c_idx]
            cell.paragraphs[0].clear()
            run = cell.paragraphs[0].add_run(str(val) if val is not None else "")
            _set_font(run, size_pt=10)
            if zebra:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), fill)
                tcPr.append(shd)
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Inches(w)
    spacer()
    return table

style = doc.styles["Normal"]
style.font.name = SORA
style.font.size = Pt(11)

# ── Cover ─────────────────────────────────────────────────────────────────────

heading("Data Migration Review", level=1)
body(f"Hoban Financial Planning Ltd    Prepared: {TODAY}", bold=False)
spacer()
body(
    "This document sets out what data has been prepared for transfer to Beyond Advice, "
    "what needs your review before the transfer runs, and what will stay on Plannr. "
    "No client names, dates of birth, addresses or policy numbers appear in this document.",
    italic=True
)

# ── 1. Summary ────────────────────────────────────────────────────────────────

heading("1. What has been prepared", level=2)
body(
    "The tables below cover every category of data that can move from Plannr to "
    "Beyond Advice. Each category has been checked against the Plannr records, "
    "cleaned where needed, and built into a workbook ready to load."
)

total_plans  = files_s.get("Plannr_Plans.xlsx (Plans)", 0)
total_charge = files_s.get("Plannr_Charges.xlsx", 0)
total_iae    = files_s.get("Plannr_Income_Employment.xlsx", 0)

summary_rows = [
    ("Clients",              f"{files_s.get('Plannr_Clients.xlsx', 0):,}",     "All active clients"),
    ("Plans",                f"{total_plans:,}",
     f"Includes {plans_s.get('new_post_snapshot', 0)} added in the last two days"),
    ("Charges",              f"{total_charge:,}",
     f"See Section 3 - {charges['rate_needed']['rows']} need a rate from you"),
    ("Circles",              f"{files_s.get('Plannr_Circles.xlsx', 0):,}",     "Groups with members"),
    ("Relationships",        f"{files_s.get('Plannr_Relationships.xlsx', 0):,}","Linked-client pairs"),
    ("Dependants",           f"{files_s.get('Plannr_Dependants.xlsx', 0):,}",  "Unique dependant persons"),
    ("Valuations",           f"{files_s.get('Plannr_Valuations.xlsx', 0):,}",  "Most recent valuation per holding"),
    ("Income and employment",f"{total_iae:,}",                                  "From your Plannr records"),
    ("Expenditure",          f"{files_s.get('Plannr_Expenditure.xlsx', 0):,}", "Mapped to Beyond Advice categories"),
    ("Premiums",             f"{files_s.get('Plannr_Premiums.xlsx', 0):,}",    "Protection premiums"),
]

add_table(
    ["Category", "Records", "Notes"],
    summary_rows,
    col_widths=[2.0, 1.0, 3.8],
    zebra=True,
)

# ── 2. Plans ─────────────────────────────────────────────────────────────────

heading("2. Plans", level=2)

body(
    f"There are {total_plans:,} plans on Plannr ready to transfer. "
    f"{plans_s.get('placeholder_pn_blanked', 0):,} of these have a policy number recorded as "
    "not applicable, TBC, or a similar placeholder - these have been left blank on the transfer "
    "file. The original entry is kept alongside so you can check or correct it before the load runs."
)
spacer()
body(
    f"A separate tab lists {files_s.get('Plannr_Plans.xlsx (Provider)', 0)} plans from your provider "
    "statements (Transact/Nucleus and Rathbones) that could not be matched to any plan on Plannr "
    "by policy number. These may be closed plans, plans held under a different number, or plans "
    "not yet on Plannr. Please review this tab and mark each row: Add to Plannr, Already in Plannr, "
    "Not ours, or Closed."
)

# ── 3. Charges ────────────────────────────────────────────────────────────────

heading("3. Charges", level=2)

sec_a_n   = charges["provider_statement"]["rows"]
sec_a_iai = charges["provider_statement"]["iai"]
sec_r_n   = charges["rathbones"]["rows"]
sec_r_iai = charges["rathbones"]["iai"]
sec_b_n   = charges["your_plannr"]["rows"]
sec_b_iai = charges["your_plannr"]["iai"]
sec_c_n   = charges["rate_needed"]["rows"]
total_iai = charges["total_iai"]

body(
    "Charges are split into three groups so you can see exactly where each rate came from "
    "and what still needs your attention."
)
spacer()

charge_rows = [
    ("Provider statement",
     f"{sec_a_n:,}",
     f"GBP{sec_a_iai:,.2f}",
     "Rates taken from your Transact/Nucleus statement. No action needed."),
    ("Rathbones",
     f"{sec_r_n:,}",
     f"GBP{sec_r_iai:,.2f}",
     "Rates taken from your Rathbones schedule. No action needed."),
    ("Your Plannr",
     f"{sec_b_n:,}",
     f"GBP{sec_b_iai:,.2f}",
     "Charge already on Plannr - kept as-is. No action needed."),
    ("Rate needed",
     f"{sec_c_n:,}",
     "-",
     "No rate was found for these plans. Please fill in the rate or delete the row."),
]

add_table(
    ["Source", "Plans", "Indicative annual income", "What to do"],
    charge_rows,
    col_widths=[1.8, 0.7, 1.8, 2.5],
    zebra=True,
)

body(
    f"Total indicative annual income from confirmed charges: GBP{total_iai:,.2f}. "
    "This figure covers the Provider statement, Rathbones and Plannr rows. "
    "It excludes the Rate needed rows until you fill them in."
)

# ── 4. Notes ─────────────────────────────────────────────────────────────────

heading("4. Notes and tasks", level=2)

exp_total  = notes_s.get("export_total", 0)
snap_total = notes_s.get("snapshot_total", 0)
by_model   = notes_s.get("by_model", {})

task_n      = by_model.get("task", 0)
case_n      = by_model.get("case", 0)
account_n   = by_model.get("account", 0)
plan_n      = by_model.get("plan", 0)
circle_n    = by_model.get("circle", 0)
risk_n      = by_model.get("risk", 0)
charge_n    = by_model.get("charge", 0)
board_n     = by_model.get("board_task", 0)

body(
    f"Plannr holds {exp_total:,} notes in total. They fall into these categories:"
)

notes_rows = [
    ("Tasks",               f"{task_n:,}",    "Can move to Beyond Advice"),
    ("Case notes",          f"{case_n:,}",    "Can move to Beyond Advice"),
    ("Client notes",        f"{account_n:,}", "Can move to Beyond Advice"),
    ("Plan notes",          f"{plan_n:,}",    "Can move to Beyond Advice"),
    ("Circle notes",        f"{circle_n:,}",  "Can move to Beyond Advice"),
    ("Risk notes",          f"{risk_n:,}",    "Can move to Beyond Advice"),
    ("Charge notes",        f"{charge_n:,}",  "Can move to Beyond Advice"),
    ("Board tasks",         f"{board_n:,}",   "Stay on Plannr - cannot be transferred"),
]

add_table(
    ["Type", "Count", "Transfer status"],
    notes_rows,
    col_widths=[2.0, 1.0, 3.8],
    zebra=True,
)

body(
    f"The {board_n:,} board tasks cannot be transferred to Beyond Advice as that platform "
    "does not have an equivalent. They will remain accessible on your old Plannr site. "
    "All other note types can be transferred, but the transfer runs by category - you "
    "either take all of a type or none of it."
)
spacer()
body("Please confirm which categories you would like to move:", bold=True)
moveable_total = task_n + case_n + account_n + plan_n + circle_n + risk_n + charge_n
bullet(
    f"Tasks and case notes only ({task_n + case_n:,} notes) - "
    "recommended if you want to keep the record clean",
    bold_prefix=None,
)
bullet(
    f"All transferable types ({moveable_total:,} notes) - "
    "includes client, plan, circle, risk and charge notes",
    bold_prefix=None,
)

# ── 5. Dependants ─────────────────────────────────────────────────────────────

heading("5. Dependants", level=2)

snap_dep = deps_s.get("snapshot", 249)
total_dep = deps_s.get("total", 250)
api_only  = deps_s.get("api_only", 1)

body(
    f"There are {total_dep} dependants on Plannr. {snap_dep} appeared in the full snapshot; "
    f"the remaining {api_only} was found only in the live download and has been included. "
    "Where a dependant is linked to more than one client household, all linked households "
    "are carried across together."
)

# ── 6. Income and employment ───────────────────────────────────────────────────

heading("6. Income and employment", level=2)

inc_n   = iae_s.get("income_rows", 0)
emp_n   = iae_s.get("employment_rows", 0)
stalone = iae_s.get("standalone_emp_rows", 0)

body(
    f"The income and employment file contains {inc_n:,} income records and "
    f"{emp_n:,} employment records. Where a client has both, their employment "
    "details sit alongside the matching income row. "
)
if stalone:
    body(
        f"{stalone} employment records had no income record to sit alongside - "
        "these have been given their own row with the income fields left blank. "
        "Please review these rows and either fill in the income or remove them."
    )

# ── 7. What stays on Plannr ───────────────────────────────────────────────────

heading("7. What stays on Plannr", level=2)

body(
    "The following will not move to Beyond Advice and will remain accessible "
    "on your current Plannr site:"
)
bullet(f"Board tasks ({board_n:,} items)")
bullet("Documents, emails and illustrations - these will be handled separately")
bullet("Any data added to Plannr after the transfer date")

# ── 8. Actions needed from you ────────────────────────────────────────────────

heading("8. Actions needed from you", level=2)

body("Before the transfer can run, please confirm or complete the following:")

action_rows = [
    ("1",
     "Charges - Rate needed tab",
     f"{sec_c_n} plans have no charge rate. Open Plannr_Charges.xlsx, "
     "filter the Basis column to 'Rate needed', and either fill in the rate "
     "or delete the row if no charge applies.",
     "Required before load"),
    ("2",
     "Provider plans not in Plannr",
     f"{files_s.get('Plannr_Plans.xlsx (Provider)', 0)} rows in the Provider_plans_not_in_Plannr "
     "tab of Plannr_Plans.xlsx need a decision. Mark each as: "
     "Add to Plannr / Already in Plannr / Not ours / Closed.",
     "Required before load"),
    ("3",
     "Notes - which types to transfer",
     "Confirm whether to move tasks and case notes only, "
     "or all transferable note types.",
     "Required before load"),
    ("4",
     "Plans with unclear policy numbers",
     f"{plans_s.get('placeholder_pn_blanked', 0)} plans had a placeholder in the policy number field "
     "(e.g. TBC or N/A). These have been left blank. If you know the correct "
     "policy number for any of them, add it to the Plans file before the load.",
     "Optional - fill in or leave blank"),
]

add_table(
    ["No.", "Item", "What to do", "Status"],
    action_rows,
    col_widths=[0.3, 1.7, 3.5, 1.3],
    zebra=False,
)

# ── 9. Data as at ─────────────────────────────────────────────────────────────

heading("9. Data as at", level=2)

body(
    "The transfer files were built on 2026-10-14 using a Plannr snapshot "
    "taken on 2026-10-12. Valuations, income and employment data are from "
    "the 12 October snapshot. Plans and charges include updates made up to "
    "2026-10-14. Any changes made to Plannr after today will not be included "
    "unless the files are rebuilt."
)

# ── Footer ────────────────────────────────────────────────────────────────────

spacer()
para = doc.add_paragraph()
run = para.add_run(
    f"Prepared {TODAY} - Hoban Financial Planning Ltd AR migration. "
    "Client data has been removed from this document in accordance with Rule 9."
)
_set_font(run, size_pt=9, italic=True, color=GREY)

# ── Save ─────────────────────────────────────────────────────────────────────

doc.save(OUT_PATH)
p(f"Saved: {OUT_PATH}")
p("\nStep 4 v2 complete.")
