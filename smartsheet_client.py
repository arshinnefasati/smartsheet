"""
Fetch project data from Smartsheet API.
Falls back to local xlsx when SMARTSHEET_API_KEY is not set (demo mode).
"""
from __future__ import annotations
import re
import zipfile
import xml.etree.ElementTree as ET
from datetime import date, timedelta, datetime
from pathlib import Path
from typing import Any

import config


# ── Excel serial date → Python date ──────────────────────────────────────────
def _xl_date(serial) -> date | None:
    """Convert Excel serial number OR ISO date string to Python date."""
    if serial is None or serial == "":
        return None
    # Try Excel serial number first (real Smartsheet exports)
    try:
        n = float(serial)
        return (datetime(1899, 12, 30) + timedelta(days=n)).date()
    except (ValueError, TypeError):
        pass
    # Fallback: common date string formats (test xlsx, API strings, etc.)
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y"):
        try:
            return datetime.strptime(str(serial).strip(), fmt).date()
        except ValueError:
            continue
    return None


# ── Parse xlsx (offline / demo mode) ─────────────────────────────────────────
def load_from_xlsx(path: Path, use_name_override: bool = True) -> dict[str, Any]:
    """Read the Timeline Extract xlsx and return the same dict shape as the API.

    use_name_override=False: always use the project name stored inside the xlsx
    (e.g. when loading an uploaded file that is NOT the default configured sheet).
    """

    ns_ss = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

    with zipfile.ZipFile(path) as z:
        # Shared strings
        with z.open("xl/sharedStrings.xml") as f:
            ss_root = ET.fromstring(f.read())
        strings: list[str] = []
        for si in ss_root.findall("x:si", ns_ss):
            t_nodes = si.findall(".//x:t", ns_ss)
            strings.append("".join(t.text or "" for t in t_nodes))

        # Sheet 1
        with z.open("xl/worksheets/sheet1.xml") as f:
            ws_root = ET.fromstring(f.read())

    rows: list[list[str]] = []
    for row in ws_root.findall(".//x:row", ns_ss):
        cells: list[str] = []
        for c in row.findall("x:c", ns_ss):
            t_attr = c.get("t", "")
            v = c.find("x:v", ns_ss)
            if v is None:
                cells.append("")
            elif t_attr == "s":
                cells.append(strings[int(v.text)])
            else:
                cells.append(v.text or "")
        if any(x.strip() for x in cells):
            rows.append(cells)

    if not rows:
        return _empty_project()

    headers = rows[0]
    col = {name: idx for idx, name in enumerate(headers)}

    def cell(row: list[str], name: str) -> str:
        idx = col.get(name)
        return row[idx].strip() if idx is not None and idx < len(row) else ""

    tasks: list[dict] = []
    milestones: list[dict] = []
    all_notes: list[str] = []

    for row in rows[1:]:
        name = cell(row, "Primary")
        if not name:
            continue

        include = cell(row, "Include in Dashboard Timeline")
        if include not in ("1", "1.0"):
            continue

        end_serial   = cell(row, "End Date")
        start_serial = cell(row, "Start Date")
        end_date     = _xl_date(end_serial)
        start_date   = _xl_date(start_serial)
        if end_date is None:
            continue

        status   = cell(row, "Status")
        category = cell(row, "Category")
        pct_raw  = cell(row, "% Complete")
        past_due = cell(row, "Past Due Helper")
        impacts  = cell(row, "GN&T Impacts")
        notes    = cell(row, "Notes")
        hier_raw = cell(row, "Hierarchy")

        try:
            pct = float(pct_raw) if pct_raw else 0.0
        except ValueError:
            pct = 0.0
        try:
            hierarchy = int(float(hier_raw)) if hier_raw else 2
        except ValueError:
            hierarchy = 2

        item = {
            "name":       name,
            "start":      start_date,
            "end":        end_date,
            "status":     status,
            "type":       category,           # "Task" | "Milestone"
            "pct":        pct,
            "past_due":   past_due == "1",
            "impacts":    impacts,
            "notes":      notes,
            "hierarchy":  hierarchy,
            "is_done":    status == "Complete",
            "is_wip":     status == "In Progress",
        }

        if notes:
            all_notes.append(notes)

        if category == "Milestone":
            milestones.append(item)
        else:
            tasks.append(item)

    # Derive project meta from data
    raw_sheet    = cell(rows[1], "Sheet Name") if len(rows) > 1 else "Project"
    sheet_name   = raw_sheet.replace(" - Project Schedule", "").strip() or "Project"
    project_name = (config.PROJECT_NAME_OVERRIDE if use_name_override else None) or sheet_name

    return {
        "project_name": project_name,
        "pm_name":       "Steven Bruchey",
        "reference":     "PID-0085",
        "tasks":         tasks,
        "milestones":    milestones,
        "all_items":     tasks + milestones,
        "notes":         all_notes,
        "source":        "xlsx",
    }


# ── Smartsheet API helpers ────────────────────────────────────────────────────
def _api_get(path: str, api_key: str) -> dict:
    """Make a GET request to the Smartsheet API and return parsed JSON."""
    import json
    import ssl
    import urllib.request
    import urllib.error

    # Bypass SSL verification — needed when VPN does SSL inspection
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    url = f"https://api.smartsheet.com/2.0{path}"
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {api_key}"})
    try:
        with urllib.request.urlopen(req, timeout=30, context=ctx) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        raise RuntimeError(f"Smartsheet API {e.code}: {body}") from e


def list_sheets(api_key: str) -> list[dict]:
    """Return list of {id, name} for all sheets the token can access."""
    data = _api_get("/sheets", api_key)
    return [{"id": s["id"], "name": s["name"]} for s in data.get("data", [])]


def list_sights(api_key: str) -> list[dict]:
    """Return list of {id, name, permalink} for all dashboards the token can access."""
    data = _api_get("/sights", api_key)
    return [
        {"id": s["id"], "name": s["name"], "permalink": s.get("permalink", "")}
        for s in data.get("data", [])
    ]


def load_from_published_url(publish_url: str) -> dict[str, Any]:
    """Read a publicly published Smartsheet (no API key needed).

    Smartsheet publish URLs look like:
      https://app.smartsheet.com/b/publish?EQBCT=xxxxxxxxxxxxxxxx

    The published page contains a JSON bundle with all sheet data.
    No authentication required — works even with VPN on.
    """
    import ssl, urllib.request, json as _json

    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    req = urllib.request.Request(
        publish_url,
        headers={"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"}
    )
    try:
        with urllib.request.urlopen(req, timeout=30, context=ctx) as resp:
            html = resp.read().decode("utf-8", errors="replace")
    except Exception as e:
        raise RuntimeError(f"Could not fetch published sheet: {e}")

    # Smartsheet embeds sheet data as JSON in a <script> tag
    # Look for: window.ss.settings = {...}  or  "primaryColumnIndex"
    patterns = [
        r'window\.ss\.settings\s*=\s*(\{.*?\});',
        r'"rows"\s*:\s*(\[.*?\])\s*,\s*"totalRowCount"',
        r'bootstrapData\s*=\s*(\{.*?\})\s*;',
    ]
    data_json = None
    for pat in patterns:
        m = re.search(pat, html, re.DOTALL)
        if m:
            try:
                data_json = _json.loads(m.group(1))
                break
            except Exception:
                continue

    if data_json is None:
        # Try to extract rows table from HTML directly
        raise RuntimeError(
            "Could not parse data from published Smartsheet page. "
            "Make sure the sheet is published as a grid/table view."
        )

    # Normalise whatever structure we found into our standard format
    return _parse_published_json(data_json, publish_url)


def _parse_published_json(raw: dict, source_url: str) -> dict[str, Any]:
    """Convert the published-page JSON into the same dict format as load_from_xlsx."""
    # The structure varies by Smartsheet version; try common paths
    sheet_name = (
        raw.get("name") or
        raw.get("sheet", {}).get("name") or
        raw.get("title") or
        "Project"
    )

    rows_raw = (
        raw.get("rows") or
        raw.get("sheet", {}).get("rows") or
        []
    )

    columns = (
        raw.get("columns") or
        raw.get("sheet", {}).get("columns") or
        []
    )

    col_map = {c.get("id"): c.get("title", "") for c in columns}
    col_idx  = {c.get("title", "").lower(): c.get("id") for c in columns}

    def col_val(row, name_fragment):
        """Find cell value by partial column name match."""
        cells = {c.get("columnId"): c.get("displayValue") or c.get("value") for c in row.get("cells", [])}
        for title_lower, cid in col_idx.items():
            if name_fragment.lower() in title_lower:
                v = cells.get(cid)
                if v is not None:
                    return v
        return None

    tasks, milestones, all_items = [], [], []
    pm_name = ""
    reference = ""

    for row in rows_raw:
        name = col_val(row, "name") or col_val(row, "task") or ""
        if not name:
            continue

        start_raw = col_val(row, "start")
        end_raw   = col_val(row, "finish") or col_val(row, "end") or col_val(row, "due")
        status    = col_val(row, "status") or "Not Started"
        pct_raw   = col_val(row, "complete") or col_val(row, "percent") or 0
        row_type  = col_val(row, "type") or ""

        start_d = _parse_date_str(str(start_raw)) if start_raw else None
        end_d   = _parse_date_str(str(end_raw))   if end_raw   else None
        if end_d is None and start_d is not None:
            end_d = start_d

        try:
            pct = float(str(pct_raw).replace("%","")) / 100 if pct_raw else 0.0
        except Exception:
            pct = 0.0

        is_milestone = "milestone" in str(row_type).lower() or pct_raw == "" and start_d == end_d
        is_done = pct >= 1.0 or "complete" in str(status).lower() or "done" in str(status).lower()
        past_due = end_d is not None and end_d < date.today() and not is_done

        item = {
            "name": str(name),
            "start": start_d or date.today(),
            "end":   end_d   or date.today(),
            "status": str(status),
            "type": "Milestone" if is_milestone else "Task",
            "pct": pct,
            "past_due": past_due,
            "impacts": "",
            "notes": "",
            "hierarchy": row.get("parentId") and 2 or 1,
            "is_done": is_done,
            "is_wip": 0 < pct < 1.0,
        }
        all_items.append(item)
        if is_milestone:
            milestones.append(item)
        else:
            tasks.append(item)

    return {
        "project_name": sheet_name,
        "pm_name":      pm_name or "Unknown",
        "reference":    reference or sheet_name[:20],
        "tasks":        tasks,
        "milestones":   milestones,
        "all_items":    all_items,
        "notes":        "",
        "source":       f"published:{source_url}",
    }


def _parse_date_str(s: str) -> date | None:
    """Parse common date string formats."""
    for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d/%m/%Y", "%B %d, %Y", "%b %d, %Y"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except Exception:
            continue
    return None


def parse_smartsheet_url(url_or_id: str) -> dict:
    """Parse a Smartsheet URL or raw ID into {type, id}.

    Supported formats:
      - Published URL: https://app.smartsheet.com/b/publish?EQBCT=...
      - Sheet URL:     https://app.smartsheet.com/sheets/<permalink>
      - Dashboard URL: https://app.smartsheet.com/dashboards/<permalink>
      - Numeric ID:    1234567890123456  (direct API sheet ID)
    """
    s = url_or_id.strip()
    # Published sheet (no API key needed)
    m = re.search(r"smartsheet\.com/b/publish", s)
    if m:
        return {"type": "published", "id": s}
    m = re.search(r"smartsheet\.com/sheets/([A-Za-z0-9]+)", s)
    if m:
        return {"type": "sheet", "id": m.group(1)}
    m = re.search(r"smartsheet\.com/dashboards/([A-Za-z0-9]+)", s)
    if m:
        return {"type": "dashboard", "id": m.group(1)}
    if re.match(r"^\d+$", s):
        return {"type": "sheet_numeric", "id": s}
    return {"type": "unknown", "id": s}


def resolve_to_sheet_id(url_or_id: str, api_key: str) -> str:
    """Resolve a Smartsheet URL or ID to a numeric sheet ID string.

    - Sheet URL permalink → matched against sheet list by trying direct API call
    - Dashboard URL → calls Sights API, finds first linked sheet widget
    - Numeric ID → returned as-is
    """
    parsed = parse_smartsheet_url(url_or_id)

    if parsed["type"] == "sheet_numeric":
        return parsed["id"]

    if parsed["type"] == "sheet":
        # Smartsheet permalink IDs are base-62; the API accepts them directly
        # via the /sheets endpoint (returns 404 if not found).
        try:
            sheet = _api_get(f"/sheets/{parsed['id']}", api_key)
            return str(sheet["id"])
        except Exception:
            # Fallback: search the sheet list for a matching permalink
            sheets = list_sheets(api_key)
            for s in sheets:
                if str(s["id"]) == parsed["id"]:
                    return str(s["id"])
            raise ValueError(
                f"Sheet not found for URL token '{parsed['id']}'. "
                "Use 'Load Projects' to pick from your accessible sheets."
            )

    if parsed["type"] == "dashboard":
        masked_id = parsed["id"]
        # The API requires a numeric sight ID — resolve by listing all sights
        # and matching the masked URL token against each sight's permalink URL.
        sights = list_sights(api_key)
        matched = next((s for s in sights if masked_id in s.get("permalink", "")), None)
        if not matched:
            names = ", ".join(s["name"] for s in sights) or "none found"
            raise ValueError(
                f"Dashboard not matched. Available dashboards: {names}"
            )
        # Fetch full sight using its numeric ID to find linked sheet widgets
        sight = _api_get(f"/sights/{matched['id']}", api_key)
        for widget in sight.get("widgets", []):
            contents = widget.get("contents", {})
            sheet_id = contents.get("sheetId")
            if sheet_id:
                return str(sheet_id)
        raise ValueError(
            f"No linked sheet found in dashboard '{matched['name']}'. "
            "The dashboard may only contain text/image widgets with no sheet data."
        )

    raise ValueError(
        f"Unrecognized URL format. Paste a Smartsheet sheet or dashboard URL, "
        f"or a numeric sheet ID."
    )


def _api_post(path: str, api_key: str, payload: dict) -> dict:
    """Make a POST request to the Smartsheet API and return parsed JSON."""
    import json as _json
    import ssl
    import urllib.request
    import urllib.error

    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    url  = f"https://api.smartsheet.com/2.0{path}"
    data = _json.dumps(payload).encode()
    req  = urllib.request.Request(
        url, data=data,
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=30, context=ctx) as resp:
            return _json.loads(resp.read().decode())
    except urllib.error.HTTPError as e:
        body = e.read().decode()
        raise RuntimeError(f"Smartsheet API {e.code}: {body}") from e


def push_timeline_to_dashboard(
    dashboard_url: str,
    timeline_public_url: str,
    widget_title: str,
    api_key: str,
) -> dict:
    """Add or update a Web Content widget on a Smartsheet dashboard.

    Finds an existing widget titled `widget_title` and updates its URL,
    or creates a new one if not found.
    Returns the API response dict.
    """
    # Resolve dashboard masked ID → numeric sight ID
    parsed = parse_smartsheet_url(dashboard_url)
    if parsed["type"] != "dashboard":
        raise ValueError("Please provide a dashboard URL (app.smartsheet.com/dashboards/...)")

    sights = list_sights(api_key)
    matched = next((s for s in sights if parsed["id"] in s.get("permalink", "")), None)
    if not matched:
        raise ValueError("Dashboard not found in your account.")

    sight_id = matched["id"]

    # Check if a widget with this title already exists
    sight = _api_get(f"/sights/{sight_id}", api_key)
    existing = next(
        (w for w in sight.get("widgets", []) if w.get("title") == widget_title),
        None,
    )

    widget_payload = {
        "type": "WEBCONTENT",
        "title": widget_title,
        "showTitle": True,
        "contents": {"url": timeline_public_url},
    }

    if existing:
        # Update existing widget
        result = _api_post(
            f"/sights/{sight_id}/widgets/{existing['id']}",
            api_key, widget_payload,
        )
    else:
        # Add new widget
        result = _api_post(
            f"/sights/{sight_id}/widgets",
            api_key, widget_payload,
        )

    return result


# ── Sheet Summary (project metadata sidebar) ──────────────────────────────────
def _fetch_sheet_summary(sheet_id: str, api_key: str) -> dict[str, str]:
    """Fetch Sheet Summary fields — returns {field_title: display_value}.

    Sheet Summary is the metadata sidebar in Smartsheet (Status, PM, Tech Lead,
    Clarity ID, % Complete, etc.). Returns empty dict if not available or error.
    """
    try:
        data = _api_get(f"/sheets/{sheet_id}/summary", api_key)
        result = {}
        for field in data.get("fields", []):
            title = field.get("title", "")
            # displayValue is formatted (e.g. "45%", "In Progress")
            # value is raw (contact object, number, etc.)
            val = field.get("displayValue") or field.get("value")
            # Contact fields return an object like {"name": "Steven Bruchey"}
            if isinstance(val, dict):
                val = val.get("name") or val.get("email") or str(val)
            if title and val is not None:
                result[title] = str(val)
        return result
    except Exception as e:
        print(f"[_fetch_sheet_summary] Could not fetch summary: {e}")
        return {}


# ── Smartsheet API ────────────────────────────────────────────────────────────
def load_from_api(sheet_id: str, api_key: str) -> dict[str, Any]:
    """Fetch sheet data from Smartsheet REST API v2 using urllib (no SDK needed)."""
    sheet = _api_get(f"/sheets/{sheet_id}", api_key)

    # Build column id → title map
    col_map = {col["id"]: col["title"] for col in sheet.get("columns", [])}
    print(f"[load_from_api] Sheet columns: {list(col_map.values())}")
    col_titles_lower = {v.lower(): k for k, v in col_map.items()}

    def cell_val(row: dict, col_name: str) -> str:
        """Exact column name match."""
        for c in row.get("cells", []):
            if col_map.get(c.get("columnId")) == col_name:
                return str(c.get("displayValue") or c.get("value") or "")
        return ""

    def cell_fuzzy(row: dict, *fragments: str) -> str:
        """Return first cell whose column title contains any of the fragments (case-insensitive)."""
        for frag in fragments:
            for cid, title in col_map.items():
                if frag.lower() in title.lower():
                    for c in row.get("cells", []):
                        if c.get("columnId") == cid:
                            v = str(c.get("displayValue") or c.get("value") or "")
                            if v:
                                return v
        return ""

    # Only filter by "Include" checkbox if the column actually exists in this sheet
    include_col_exists = any(
        "include" in t.lower() and "dashboard" in t.lower()
        for t in col_map.values()
    )

    tasks, milestones, all_notes = [], [], []

    for row in sheet.get("rows", []):
        name = cell_val(row, "Primary") or cell_fuzzy(row, "task name", "name", "primary")
        if not name:
            continue

        # Skip rows that don't have "Include" checked — but only if the column exists
        if include_col_exists:
            include = cell_val(row, "Include in Dashboard Timeline") or cell_fuzzy(row, "include")
            if include not in ("1", "1.0", "true", "True"):
                continue

        status    = cell_val(row, "Status")    or cell_fuzzy(row, "status")
        category  = cell_val(row, "Category")  or cell_fuzzy(row, "category", "type")
        end_raw   = cell_val(row, "End Date")  or cell_fuzzy(row, "finish", "end date", "due date", "end")
        start_raw = cell_val(row, "Start Date") or cell_fuzzy(row, "start date", "begin", "start")
        pct_raw   = cell_val(row, "% Complete")
        past_due  = cell_val(row, "Past Due Helper")
        impacts   = cell_val(row, "GN&T Impacts")
        notes     = cell_val(row, "Notes")
        hier_raw  = cell_val(row, "Hierarchy")

        # Parse dates (Smartsheet returns YYYY-MM-DD strings)
        try:
            end_date   = date.fromisoformat(end_raw[:10]) if end_raw else None
            start_date = date.fromisoformat(start_raw[:10]) if start_raw else None
        except ValueError:
            end_date = start_date = None

        if end_date is None:
            continue

        try:
            pct = float(pct_raw.replace("%", "")) if pct_raw else 0.0
        except ValueError:
            pct = 0.0
        try:
            hierarchy = int(float(hier_raw)) if hier_raw else 2
        except ValueError:
            hierarchy = 2

        item = {
            "name":      name,
            "start":     start_date,
            "end":       end_date,
            "status":    status,
            "type":      category,
            "pct":       pct,
            "past_due":  past_due in ("1", "true", "True"),
            "impacts":   impacts,
            "notes":     notes,
            "hierarchy": hierarchy,
            "is_done":   status == "Complete",
            "is_wip":    status == "In Progress",
        }

        if notes:
            all_notes.append(notes)

        if category == "Milestone":
            milestones.append(item)
        else:
            tasks.append(item)

    raw_name   = sheet.get("name", "Project").replace(" - Project Schedule", "").strip()
    sheet_name = config.PROJECT_NAME_OVERRIDE or raw_name

    # ── Sheet Summary fields (project metadata sidebar) ───────────────────────
    meta = _fetch_sheet_summary(sheet_id, api_key)
    print(f"[load_from_api] Sheet summary fields: {list(meta.keys())}")

    def _m(*keys) -> str:
        """Return first non-empty summary value matching any of the key fragments."""
        for k in keys:
            for title, val in meta.items():
                if k.lower() in title.lower() and val:
                    return val
        return ""

    pm_name          = _m("project manager", "pm name", "tpd pm")
    tech_lead        = _m("tech lead", "technical lead")
    project_status   = _m("status")
    exec_status      = _m("executive status", "exec status")
    pct_complete_raw = _m("% complete", "percent complete", "pct")
    clarity_id       = _m("clarity id", "clarity")
    unique_id        = _m("unique id", "uid")
    project_start    = _m("start date", "start")
    project_end      = _m("end date", "finish", "end")
    # Collect all impacted partner fields (any field with "impact" or "partner")
    impacted = [v for t, v in meta.items()
                if ("impact" in t.lower() or "partner" in t.lower()) and v]

    try:
        pct_complete = float(str(pct_complete_raw).replace("%", "")) / 100
    except (ValueError, TypeError):
        pct_complete = 0.0

    # Reference = first PID-XXXX token found in the sheet name or clarity id
    import re as _re
    ref_match = _re.search(r"PID-\d+", sheet_name) or _re.search(r"PID-\d+", clarity_id)
    reference = ref_match.group(0) if ref_match else clarity_id or sheet_name[:15]

    return {
        "project_name":    sheet_name,
        "pm_name":         pm_name or "Unknown",
        "tech_lead":       tech_lead,
        "reference":       reference,
        "project_status":  project_status,
        "exec_status":     exec_status,
        "pct_complete":    pct_complete,
        "clarity_id":      clarity_id,
        "unique_id":       unique_id,
        "project_start":   project_start,
        "project_end":     project_end,
        "impacted":        impacted,
        "tasks":           tasks,
        "milestones":      milestones,
        "all_items":       tasks + milestones,
        "notes":           all_notes,
        "source":          "api",
    }


# ── Public entry point ────────────────────────────────────────────────────────
def get_project_data(sheet_id: str | None = None) -> dict[str, Any]:
    """Return project data dict — from API (with optional sheet_id override) or xlsx fallback."""
    effective_id = sheet_id or config.SMARTSHEET_SHEET_ID

    if config.DEMO_MODE or not effective_id:
        print("[smartsheet_client] Demo mode — loading from xlsx")
        return load_from_xlsx(config.SOURCE_XLSX)

    print(f"[smartsheet_client] Live mode — fetching sheet {effective_id}")
    try:
        data = load_from_api(effective_id, config.SMARTSHEET_API_KEY)
        print(f"[smartsheet_client] API OK — {len(data['all_items'])} items from '{data['project_name']}'")
        return data
    except Exception as exc:
        print(f"[smartsheet_client] API failed: {exc}")
        print("[smartsheet_client] Trying to list accessible sheets for diagnostics...")
        try:
            sheets = list_sheets(config.SMARTSHEET_API_KEY)
            print(f"[smartsheet_client] Sheets accessible with this token:")
            for s in sheets[:10]:
                print(f"  id={s['id']}  name={s['name']}")
        except Exception as list_exc:
            print(f"[smartsheet_client] Could not list sheets either: {list_exc}")
            print("[smartsheet_client] Network may be blocked or API key is invalid.")
        print("[smartsheet_client] Falling back to xlsx...")
        return load_from_xlsx(config.SOURCE_XLSX)


def _empty_project() -> dict:
    return {
        "project_name": "Unnamed Project",
        "pm_name": "",
        "reference": "",
        "tasks": [],
        "milestones": [],
        "all_items": [],
        "notes": [],
        "source": "empty",
    }
