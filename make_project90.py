"""
Generate PID-0090-Network_Infrastructure_Upgrade.xlsx
Pure stdlib — no openpyxl needed.

Run:  python3 make_project90.py
Output: uploads/PID-0090-Network_Infrastructure_Upgrade.xlsx
"""
import io, zipfile
from pathlib import Path

# ── Project data ──────────────────────────────────────────────────────────────
PROJECT_NAME = "PID-0090-Network Infrastructure Upgrade"

# Columns (must match what load_from_xlsx expects):
#  Sheet Name | Primary | Include in Dashboard Timeline | Start Date | End Date
#  Status | Category | % Complete | Past Due Helper | GN&T Impacts | Notes | Hierarchy

ROWS = [
    # name, include, start, end, status, category, pct, past_due, impacts, notes, hier
    ("Project Kickoff",              "1", "2026-03-01", "2026-03-01", "Complete",    "Milestone", "1",    "", "PMO",              "Approved & signed off",           "1"),
    ("Discovery & Requirements",     "1", "2026-03-02", "2026-03-31", "Complete",    "Task",      "1",    "", "Architecture",     "All stakeholders interviewed",    "1"),
    ("Network Architecture Design",  "1", "2026-04-01", "2026-04-30", "Complete",    "Task",      "1",    "", "Architecture",     "Design doc v2.0 approved",        "1"),
    ("Design Review & Approval",     "1", "2026-04-30", "2026-04-30", "Complete",    "Milestone", "1",    "", "PMO, Architecture","Signed off by VP",                "1"),
    ("Core Infrastructure Build",    "1", "2026-05-01", "2026-06-15", "In Progress", "Task",      "0.45", "", "Network, Cloud",   "Rack installation 45% done",      "1"),
    ("Security Configuration",       "1", "2026-05-15", "2026-06-30", "In Progress", "Task",      "0.20", "", "Security",         "Firewall rules in review",        "1"),
    ("Integration Testing Begin",    "1", "2026-06-01", "2026-06-30", "Not Started", "Task",      "0",    "", "QA",               "",                                "1"),
    ("Integration Complete",         "1", "2026-06-30", "2026-06-30", "Not Started", "Milestone", "0",    "", "QA, Network",      "Gate: all systems connected",     "1"),
    ("User Acceptance Testing",      "1", "2026-07-01", "2026-07-31", "Not Started", "Task",      "0",    "", "Business Units",   "",                                "1"),
    ("Performance & Load Testing",   "1", "2026-07-15", "2026-08-15", "Not Started", "Task",      "0",    "", "QA, Network",      "",                                "1"),
    ("UAT Sign-off",                 "1", "2026-08-15", "2026-08-15", "Not Started", "Milestone", "0",    "", "PMO, Business",    "Go/No-go decision point",         "1"),
    ("Production Deployment",        "1", "2026-09-01", "2026-09-30", "Not Started", "Task",      "0",    "", "Network, Ops",     "Phased rollout plan required",    "1"),
    ("Go Live",                      "1", "2026-10-01", "2026-10-01", "Not Started", "Milestone", "0",    "", "All",              "Full cutover to new infrastructure","1"),
    ("Post-Implementation Review",   "1", "2026-10-01", "2026-11-30", "Not Started", "Task",      "0",    "", "PMO",              "",                                "1"),
    ("Project Closeout",             "1", "2026-12-15", "2026-12-15", "Not Started", "Milestone", "0",    "", "PMO",              "Lessons learned & final report",  "1"),
]

# ── Build shared strings list ─────────────────────────────────────────────────
def build_data():
    """Return (shared_strings, all_rows_as_index_lists)."""
    str_idx: dict[str, int] = {}
    counter = [0]

    def sid(s: str) -> int:
        s = str(s)
        if s not in str_idx:
            str_idx[s] = counter[0]
            counter[0] += 1
        return str_idx[s]

    header = [
        "Sheet Name", "Primary", "Include in Dashboard Timeline",
        "Start Date", "End Date", "Status", "Category", "% Complete",
        "Past Due Helper", "GN&T Impacts", "Notes", "Hierarchy",
    ]
    for h in header:
        sid(h)

    encoded_rows = []
    # Header row
    encoded_rows.append([("s", sid(h)) for h in header])

    for (name, include, start, end, status, cat, pct, past_due, impacts, notes, hier) in ROWS:
        row_vals = [
            ("s", sid(PROJECT_NAME)),
            ("s", sid(name)),
            ("s", sid(include)),
            ("s", sid(start)),
            ("s", sid(end)),
            ("s", sid(status)),
            ("s", sid(cat)),
            ("s", sid(pct)),
            ("s", sid(past_due)),
            ("s", sid(impacts)),
            ("s", sid(notes)),
            ("s", sid(hier)),
        ]
        encoded_rows.append(row_vals)

    # Rebuild ordered list
    strings = [""] * len(str_idx)
    for s, i in str_idx.items():
        strings[i] = s

    return strings, encoded_rows


# ── XML helpers ───────────────────────────────────────────────────────────────
COL_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

def col_letter(n: int) -> str:
    """0-based column index → letter (A, B, …, Z, AA, …)"""
    result = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        result = COL_LETTERS[r] + result
    return result


def build_shared_strings_xml(strings: list[str]) -> bytes:
    count = len(strings)
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        f' count="{count}" uniqueCount="{count}">',
    ]
    for s in strings:
        safe = (s.replace("&", "&amp;").replace("<", "&lt;")
                  .replace(">", "&gt;").replace('"', "&quot;"))
        parts.append(f"<si><t>{safe}</t></si>")
    parts.append("</sst>")
    return "\n".join(parts).encode()


def build_sheet_xml(encoded_rows: list) -> bytes:
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        "<sheetData>",
    ]
    for ri, row in enumerate(encoded_rows, start=1):
        parts.append(f'<row r="{ri}">')
        for ci, (t, v) in enumerate(row):
            ref = f"{col_letter(ci)}{ri}"
            if t == "s":
                parts.append(f'<c r="{ref}" t="s"><v>{v}</v></c>')
            else:
                parts.append(f'<c r="{ref}"><v>{v}</v></c>')
        parts.append("</row>")
    parts += ["</sheetData>", "</worksheet>"]
    return "\n".join(parts).encode()


CONTENT_TYPES = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""

ROOT_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="xl/workbook.xml"/>
</Relationships>"""

WORKBOOK_XML = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""

WORKBOOK_RELS = b"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
    Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
    Target="sharedStrings.xml"/>
</Relationships>"""


# ── Write xlsx ────────────────────────────────────────────────────────────────
def make_xlsx(dest: Path):
    strings, encoded_rows = build_data()

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml",          CONTENT_TYPES)
        zf.writestr("_rels/.rels",                  ROOT_RELS)
        zf.writestr("xl/workbook.xml",              WORKBOOK_XML)
        zf.writestr("xl/_rels/workbook.xml.rels",   WORKBOOK_RELS)
        zf.writestr("xl/sharedStrings.xml",         build_shared_strings_xml(strings))
        zf.writestr("xl/worksheets/sheet1.xml",     build_sheet_xml(encoded_rows))

    dest.parent.mkdir(exist_ok=True)
    dest.write_bytes(buf.getvalue())
    print(f"Created: {dest}  ({dest.stat().st_size:,} bytes)")
    print(f"Project: {PROJECT_NAME}")
    print(f"Rows:    {len(ROWS)} (tasks + milestones)")


if __name__ == "__main__":
    out = Path(__file__).parent / "uploads" / "PID-0090-Network_Infrastructure_Upgrade.xlsx"
    make_xlsx(out)
    print("\nDone! Upload this file in the web UI via Path B → xlsx upload.")
