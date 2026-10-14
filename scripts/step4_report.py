"""
HFP AR Migration - Step 4: Write hfp_Data_Analysis.docx

Beyond Advice house style:
- Font: Sora throughout
- Headings: navy #020849
- No em dashes (use : or - instead)
- Aggregates only (Rule 9), no PII
"""
import os, json
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DATA_ROOT    = r"C:\dev\AR Migration"
SUMMARY_PATH = os.path.join(DATA_ROOT, "step3_summary.json")
OUT_PATH     = os.path.join(DATA_ROOT, "hfp_Data_Analysis.docx")
TODAY        = "2026-10-14"

NAVY = RGBColor(0x02, 0x08, 0x49)
BLACK = RGBColor(0x00, 0x00, 0x00)
SORA = "Sora"

def p(*a): print(*a, flush=True)

with open(SUMMARY_PATH, encoding="utf-8") as f:
    S = json.load(f)

# ── Doc helpers ────────────────────────────────────────────────────────────────

doc = Document()

def _set_font(run, size_pt=11, bold=False, color=BLACK):
    run.font.name = SORA
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.color.rgb = color
    # Force East-Asian font to Sora too
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
    para.paragraph_format.space_before = Pt(12 if level == 1 else 6)
    para.paragraph_format.space_after  = Pt(4)
    run = para.add_run(text)
    _set_font(run, size_pt=size, bold=True, color=NAVY)
    return para

def body(text, bold=False, indent=False):
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(3)
    if indent:
        para.paragraph_format.left_indent = Inches(0.25)
    run = para.add_run(text)
    _set_font(run, size_pt=11, bold=bold)
    return para

def bullet(text):
    para = doc.add_paragraph(style="List Bullet")
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after  = Pt(2)
    run = para.runs[0] if para.runs else para.add_run("")
    # clear and reset
    for r in para.runs:
        r.clear()
    run = para.add_run(text)
    _set_font(run, size_pt=11)
    return para

def add_table(headers, rows_data, col_widths=None):
    """Add a table with headers and data rows."""
    table = doc.add_table(rows=1 + len(rows_data), cols=len(headers))
    table.style = "Table Grid"
    # Header row
    hdr_row = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.paragraphs[0].clear()
        run = cell.paragraphs[0].add_run(h)
        _set_font(run, size_pt=10, bold=True, color=NAVY)
        # Navy shading
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "EEF0F6")
        tcPr.append(shd)
    # Data rows
    for r_idx, row in enumerate(rows_data, start=1):
        tr = table.rows[r_idx]
        for c_idx, val in enumerate(row):
            cell = tr.cells[c_idx]
            cell.paragraphs[0].clear()
            run = cell.paragraphs[0].add_run(str(val) if val is not None else "")
            _set_font(run, size_pt=10)
    # Column widths
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Inches(w)
    doc.add_paragraph()  # spacer
    return table

# Set default paragraph font
style = doc.styles["Normal"]
style.font.name = SORA
style.font.size = Pt(11)

# ── Cover ─────────────────────────────────────────────────────────────────────

heading("HFP AR Migration - Data Analysis", level=1)
body(f"Hoban Financial Planning Ltd    |    Migration date: {TODAY}", bold=False)
body("Prepared by: Beyond Advice migration tooling    |    Rule 9 applied: aggregate counts only", bold=False)
doc.add_paragraph()

# ── 1. Scope ──────────────────────────────────────────────────────────────────

heading("1. Scope and source", level=2)
body(
    "This analysis covers the live Plannr export pulled on 2026-10-14 from HFP's "
    "source Plannr site (account db0ab7f5), cross-referenced against the deduplicated "
    "snapshot taken on 2026-10-12 at 09:09:06 UTC. Seven plans added after the snapshot "
    "are included. Snapshot-only entities (Relationships, Dependants, Valuations, "
    "Income and Employment) are sourced from the 12 Oct snapshot."
)

# ── 2. Importer workbook file list ────────────────────────────────────────────

heading("2. Importer workbook - file list", level=2)
body("Output file: hfp_importer_v1.xlsx (saved to DATA_ROOT)", bold=False)

sc = S["sheet_counts"]
table_rows = [
    (sheet, f"{count:,}") for sheet, count in sc.items()
]
add_table(
    ["Sheet", "Rows"],
    table_rows,
    col_widths=[3.5, 1.0],
)

# ── 3. Row count reconciliation ───────────────────────────────────────────────

heading("3. Row count reconciliation", level=2)

reconcile_data = [
    ("Clients",            375,    375,    0,   "Exact match"),
    ("Plans",             1871,   1864,    7,   "7 plans added after snapshot (all post-09:09 UTC)"),
    ("Charges",            372,    372,    0,   "Exact match"),
    ("Circles",             92,     92,    0,   "92 circles with members; 6 empty excluded"),
    ("Expenditure",       2381,   2381,    0,   "Exact match"),
    ("Plannr_Premiums",    173,    None,  None,  "No snapshot target; export count used"),
    ("Relationships",      190,    368,  None,   "190 unique pairs from 368 directional references"),
    ("Dependants",         249,    367,  None,   "249 unique dependants; 367 = total account-dependant links"),
    ("Valuations",        1528,   1528,    0,   "Exact match (latest per plan/sub-account)"),
    ("Income",             316,    316,    0,   "Exact match"),
    ("Employment",         184,    184,    0,   "Exact match"),
]
add_table(
    ["Sheet", "Importer rows", "Snapshot target", "Delta", "Note"],
    [r for r in reconcile_data],
    col_widths=[1.8, 1.1, 1.2, 0.7, 3.0],
)

# ── 4. Findings by label ──────────────────────────────────────────────────────

heading("4. Findings by label", level=2)

findings = [
    ("Plans",
     "placeholder_policy_numbers_blanked",
     S["placeholder_plans"],
     "policy_number blanked (TBC / N/A / short non-numeric). "
     "Original value kept in ORIGINAL_policy_number column."),
    ("Plans",
     "joint_plans_ownership_split_blank",
     S["joint_blank"],
     "Plans with 2 or more owners where ownership_split is null in source. "
     "client_uuids and client_ownership_percentage are populated from the export."),
    ("Provider matching",
     "provider_rows_not_matched_to_plan",
     S["provider_not_matched"],
     "Provider workbook rows (Rathbones + Transact/Nucleus) with no matching "
     "plan policy_number in the 1,871-plan export. "
     "See Provider_plans_not_in_Plannr sheet."),
    ("Provider matching",
     "provider_rows_matched",
     S["provider_matched"],
     "Provider workbook rows successfully matched to a plan policy_number."),
    ("Dependants",
     "unique_dependants_from_snapshot",
     249,
     "Source total was 367 account-dependant links across 249 unique persons. "
     "account_uuids populated with all accounts each dependant was found under. "
     "Note: /api/v2/dependants returns at most 250 per page regardless of per_page; "
     "use include=accounts and follow next-page links for full export in future pulls."),
    ("Circles",
     "circles_excluded_empty",
     6,
     "6 circles excluded: 5 with no members, 1 with members but no plans. "
     "92 circles written to importer."),
    ("Relationships",
     "directional_refs_deduplicated",
     368 - 190,
     "368 directional linked-client references reduced to 190 unique pairs "
     "by keeping one row per sorted (A, B) UUID pair."),
    ("Notes",
     "board_task_notes_not_importable",
     S["note_board_task"],
     "board_task notes exist in Plannr but are not supported by the "
     "notes importer (no model_type=board_task in importer spec). "
     "Excluded from importer sheet."),
    ("Notes",
     "export_only_notes_post_snapshot",
     S["note_export_only"],
     "Notes in export not in snapshot (all post-snapshot): "
     f"board_task {S['note_eo_by_model'].get('board_task',0)}, "
     f"task +{S['note_eo_by_model'].get('task',0)}, "
     f"case +{S['note_eo_by_model'].get('case',0)}, "
     f"account +{S['note_eo_by_model'].get('account',0)}, "
     f"circle +{S['note_eo_by_model'].get('circle',0)}."),
    ("Valuations",
     "pence_to_pounds_conversion",
     1528,
     "All 1,528 valuation amounts converted from pence (snapshot integer) "
     "to pounds. Original pence values in ORIGINAL_value column."),
    ("Income_and_Employment",
     "pence_to_pounds_conversion",
     316,
     "All 316 income amounts converted from pence to pounds. "
     "Original pence values in ORIGINAL_amount column."),
]

add_table(
    ["Sheet / area", "Finding", "Count", "Detail"],
    [(f, lbl, f"{ct:,}" if ct is not None else "-", detail)
     for f, lbl, ct, detail in findings],
    col_widths=[1.4, 2.1, 0.6, 3.0],
)

# ── 5. Notes analysis ─────────────────────────────────────────────────────────

heading("5. Notes analysis", level=2)
body(
    "Note export de-duplicated by UUID. All 6,792 snapshot notes are present "
    "in the export (0 snapshot-only). The export contains 2,130 additional notes "
    "created after the snapshot."
)

notes_table_data = [
    ("task",       S["note_export_by_model"].get("task", 0),
                   S["note_snap_by_type"].get("task", 0),     "Importable"),
    ("case",       S["note_export_by_model"].get("case", 0),
                   S["note_snap_by_type"].get("case", 0),     "Importable"),
    ("account",    S["note_export_by_model"].get("account", 0),
                   S["note_snap_by_type"].get("client", 0),   "Importable (model_type=account)"),
    ("plan",       S["note_export_by_model"].get("plan", 0),
                   S["note_snap_by_type"].get("plan", 0),     "Importable"),
    ("circle",     S["note_export_by_model"].get("circle", 0),
                   S["note_snap_by_type"].get("circle", 0),   "Importable"),
    ("risk",       S["note_export_by_model"].get("risk", 0),
                   S["note_snap_by_type"].get("risk", 0),     "Importable"),
    ("charge",     S["note_export_by_model"].get("charge", 0),
                   S["note_snap_by_type"].get("charge", 0),   "Importable"),
    ("board_task", S["note_export_by_model"].get("board_task", 0),
                   0,                                          "NOT importable - no importer spec"),
]

add_table(
    ["model_type", "Export count", "Snapshot count", "Import status"],
    [(mt, f"{ec:,}", f"{sc2:,}" if sc2 else "-", status)
     for mt, ec, sc2, status in notes_table_data],
    col_widths=[1.3, 1.1, 1.2, 3.0],
)

body(
    f"Recommended import: task ({S['note_task']:,}) and case ({S['note_case']:,}) notes. "
    f"Confirm with Jim Hoban whether to also import account, plan, circle, risk and charge notes "
    f"({sum(S['note_export_by_model'].get(t, 0) for t in ['account','plan','circle','risk','charge']):,} additional). "
    "Notes import is all-or-nothing per model_type in the BA importer."
)

# ── 6. Open items ─────────────────────────────────────────────────────────────

heading("6. Open items before import", level=2)

open_items = [
    ("Ownership split",
     f"{S['joint_blank']:,} joint plans have ownership_split=null in source Plannr. "
     "The importer field is blank for all 1,871 plans. "
     "Confirm whether BA requires this field or can derive from client_ownership_percentage."),
    ("51 provider plans not in Plannr",
     "Provider workbook rows not matched to any plan by policy_number. "
     "See Provider_plans_not_in_Plannr sheet. Adviser to confirm whether these are "
     "missing from Plannr, already closed, or use a different policy number."),
    ("Dependants: 249 unique vs 367 references",
     "The importer will receive 249 unique dependant rows. The 367 figure counts "
     "total account-dependant links. Confirm that BA importer expects one row per "
     "unique dependant person (with account_uuids as a multi-value field)."),
    ("Notes: board_task excluded",
     f"{S['note_board_task']:,} board_task notes cannot be imported. "
     "Confirm with Jim Hoban whether these are in scope or can be discarded."),
    ("Plans: 7 post-snapshot plans",
     "7 plans created after the snapshot are included in the Plans sheet. "
     "All 5 same-day plans verified via API as post-09:09 UTC. "
     "These plans have no charges and no valuations in the snapshot. "
     "Adviser to review and add charges as needed after import."),
    ("income_and_employment_export: server failure",
     "The Plannr API returned a server-side failure for the I&E export. "
     "This sheet was built from the snapshot instead (316 income, 184 employment rows). "
     "Snapshot data is as at 2026-10-12."),
    ("valuation_report: API 422",
     "The Plannr valuation_report endpoint returned HTTP 422 for all tested parameter "
     "combinations. Valuations sheet built from snapshot _valuations_latest "
     "(1,528 rows). Snapshot data is as at 2026-10-12."),
]

for title, detail in open_items:
    para = doc.add_paragraph()
    para.paragraph_format.space_before = Pt(4)
    para.paragraph_format.space_after  = Pt(2)
    r1 = para.add_run(title + ": ")
    _set_font(r1, size_pt=11, bold=True)
    r2 = para.add_run(detail)
    _set_font(r2, size_pt=11)

# ── 7. Technical notes ────────────────────────────────────────────────────────

heading("7. Technical notes", level=2)

tech_items = [
    "All amounts from snapshot JSON are in pence (stored as integer strings). "
    "Converted to pounds by dividing by 100. ORIGINAL_ columns retain the raw pence value.",
    "external_reference_company = data_migration on every sheet that supports it. "
    "external_reference_value = the source Plannr UUID. This enables UUID lookup "
    "between migration waves.",
    "account_uuids on Circles and Dependants sheets uses pipe ( | ) as separator "
    "for multi-value fields.",
    "Policy numbers matching PLACEHOLDER_UPPER set (TBC, N/A, Unknown, etc.) or "
    "short non-numeric strings are blanked in Plans. 345 plans affected.",
    "Circles: the circle export provides 185 client-circle membership rows; "
    "after grouping by circle_uuid, 92 unique circles with members are produced. "
    "group_uuids set to BA default group UUID b5749882.",
    "Relationships: 368 total directional references reduced to 190 unique pairs. "
    "inverse_relationship derived from the linked account's perspective in the snapshot.",
    "/api/v2/dependants pagination note: the endpoint hard-caps at 250 rows regardless "
    "of per_page parameter. To retrieve all dependants, use include=accounts and follow "
    "the next-page link on each response page.",
]

for item in tech_items:
    bullet(item)

# ── Footer ────────────────────────────────────────────────────────────────────

doc.add_paragraph()
para = doc.add_paragraph()
run = para.add_run(
    f"Generated {TODAY} by HFP-AR-Migration tooling. Rule 9 applied throughout."
)
_set_font(run, size_pt=9, color=RGBColor(0x80, 0x80, 0x80))

# ── Save ──────────────────────────────────────────────────────────────────────

doc.save(OUT_PATH)
p(f"Saved: {OUT_PATH}")
p("\nSTOP - Steps 3 and 4 complete.")
