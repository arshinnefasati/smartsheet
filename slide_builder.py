"""
Populate the One Page Summary PPTX using raw zipfile + XML manipulation.
Zero external dependencies — stdlib only.

All timeline dimensions matched pixel-perfect to PID-0085-Dummy_FINAL.pptx
(source slide: 10" × 5.625" = 9144000 × 5143500 EMU)
"""
from __future__ import annotations

import calendar
import textwrap
import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import date
from pathlib import Path
from typing import Any

import config

# ── Namespace URIs ────────────────────────────────────────────────────────────
P_NS  = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS  = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

for _pfx, _uri in [("a", A_NS), ("p", P_NS), ("r", R_NS)]:
    ET.register_namespace(_pfx, _uri)

EMU = 914400  # 1 inch in EMU

# ── Timeline deletion zone ────────────────────────────────────────────────────
TL_TOP_EMU = int(0.62 * EMU)   # 0.62" — catches HIGH_ABOVE names at 0.682"
TL_BOT_EMU = int(2.50 * EMU)   # 2.50" — below all BELOW name boxes

# ── Text match signatures — matched against shapes in the source template ──────
# Order matters: more specific matches first.
TEXT_SIGS = [
    ("TPD PM:",                         "pm_name"),
    ("Updated:",                        "updated_date"),
    ("Reference:",                      "reference"),
    # project title — any shape whose full text looks like a project/PID name
    # and is NOT one of the protected labels
    ("PID-",                            "project_title"),
    # status / risks / summary — matched by partial content
    ("COMPLETE:",                       "status"),
    ("NOT STARTED",                     "status"),
    ("IN PROGRESS",                     "status"),
    ("ON TRACK",                        "status"),
    ("Schedule risk",                   "risks"),
    ("Dependency",                      "risks"),
    ("No risks",                        "risks"),
]
PROTECTED = {"Status", "Risks / Dependencies", "Project Summary & Reference"}


# ── XML tag helper ────────────────────────────────────────────────────────────
def Q(ns: str, local: str) -> str:
    return f"{{{ns}}}{local}"


def _shape_top(elem) -> int | None:
    xfrm = elem.find(f".//{Q(A_NS,'xfrm')}")
    if xfrm is None:
        return None
    off = xfrm.find(Q(A_NS, "off"))
    if off is None:
        return None
    try:
        return int(off.get("y", "0"))
    except (ValueError, TypeError):
        return None


def _get_all_text(elem) -> str:
    return " ".join(
        t.text for t in elem.findall(f".//{Q(A_NS,'t')}") if t.text
    ).strip()


def _replace_text(sp_elem, new_text: str):
    txBody = None
    for tag in [Q(P_NS, "txBody"), Q(A_NS, "txBody")]:
        txBody = sp_elem.find(f".//{tag}")
        if txBody is not None:
            break
    if txBody is None:
        return
    sz = None
    for rPr in txBody.findall(f".//{Q(A_NS,'rPr')}"):
        if rPr.get("sz"):
            sz = rPr.get("sz")
            break
    a_p = Q(A_NS, "p")
    for child in list(txBody):
        if child.tag == a_p:
            txBody.remove(child)
    for line in (new_text.split("\n") if new_text else [""]):
        p = ET.SubElement(txBody, a_p)
        r = ET.SubElement(p, Q(A_NS, "r"))
        if sz:
            rPr_el = ET.SubElement(r, Q(A_NS, "rPr"))
            rPr_el.set("lang", "en-US")
            rPr_el.set("sz", sz)
        t = ET.SubElement(r, Q(A_NS, "t"))
        t.text = line


# ── Native shape builders ─────────────────────────────────────────────────────
def _sp(sid, name, x, y, cx, cy, fill=None, line_w=None, line_color=None,
        prst="rect", txBox=False) -> ET.Element:
    """Create a p:sp element."""
    sp = ET.Element(Q(P_NS, "sp"))

    nv = ET.SubElement(sp, Q(P_NS, "nvSpPr"))
    cp = ET.SubElement(nv, Q(P_NS, "cNvPr"))
    cp.set("id", str(sid)); cp.set("name", name)
    csp = ET.SubElement(nv, Q(P_NS, "cNvSpPr"))
    if txBox:
        csp.set("txBox", "1")
    else:
        lk = ET.SubElement(csp, Q(A_NS, "spLocks"))
        lk.set("noGrp", "1")
    ET.SubElement(nv, Q(P_NS, "nvPr"))

    spr = ET.SubElement(sp, Q(P_NS, "spPr"))
    xf = ET.SubElement(spr, Q(A_NS, "xfrm"))
    of = ET.SubElement(xf, Q(A_NS, "off")); of.set("x", str(x)); of.set("y", str(y))
    ex = ET.SubElement(xf, Q(A_NS, "ext")); ex.set("cx", str(cx)); ex.set("cy", str(cy))
    gm = ET.SubElement(spr, Q(A_NS, "prstGeom")); gm.set("prst", prst)
    ET.SubElement(gm, Q(A_NS, "avLst"))

    if fill:
        sf = ET.SubElement(spr, Q(A_NS, "solidFill"))
        ET.SubElement(sf, Q(A_NS, "srgbClr")).set("val", fill)
    else:
        ET.SubElement(spr, Q(A_NS, "noFill"))

    if line_color:
        ln = ET.SubElement(spr, Q(A_NS, "ln")); ln.set("w", str(line_w or 9525))
        sf = ET.SubElement(ln, Q(A_NS, "solidFill"))
        ET.SubElement(sf, Q(A_NS, "srgbClr")).set("val", line_color)
    else:
        ln = ET.SubElement(spr, Q(A_NS, "ln"))
        ET.SubElement(ln, Q(A_NS, "noFill"))

    return sp


def _txbody(sp, lines, sz=900, bold=False, color="000000", align="l",
            anchor="t", insets=(45720, 45720, 22860, 22860)):
    """Append txBody (no custom font)."""
    return _txbody_f(sp, lines, sz=sz, bold=bold, color=color,
                     align=align, anchor=anchor, insets=insets, typeface=None)


def _txbody_f(sp, lines, sz=900, bold=False, color="000000", align="l",
              anchor="t", insets=(45720, 45720, 22860, 22860),
              typeface="Verizon NHG DS"):
    """Append txBody with Verizon NHG DS font and full run properties."""
    tb = ET.SubElement(sp, Q(P_NS, "txBody"))
    bp = ET.SubElement(tb, Q(A_NS, "bodyPr"))
    bp.set("anchor", anchor)
    bp.set("lIns", str(insets[0])); bp.set("rIns", str(insets[1]))
    bp.set("tIns", str(insets[2])); bp.set("bIns", str(insets[3]))
    bp.set("wrap", "square")
    ET.SubElement(bp, Q(A_NS, "normAutofit"))
    ET.SubElement(tb, Q(A_NS, "lstStyle"))

    for line in (lines if isinstance(lines, list) else [lines]):
        para = ET.SubElement(tb, Q(A_NS, "p"))
        pp = ET.SubElement(para, Q(A_NS, "pPr"))
        pp.set("algn", align)
        pp.set("indent", "0"); pp.set("marL", "0")
        ET.SubElement(pp, Q(A_NS, "spcBef")).append(
            _spc_pts(0))
        run = ET.SubElement(para, Q(A_NS, "r"))
        rp = ET.SubElement(run, Q(A_NS, "rPr"))
        rp.set("lang", "en-US"); rp.set("sz", str(sz))
        rp.set("dirty", "0"); rp.set("b", "1" if bold else "0")
        sf = ET.SubElement(rp, Q(A_NS, "solidFill"))
        ET.SubElement(sf, Q(A_NS, "srgbClr")).set("val", color)
        if typeface:
            for tag in ["latin", "ea", "cs"]:
                tf_el = ET.SubElement(rp, Q(A_NS, tag))
                tf_el.set("typeface", typeface)
        ET.SubElement(run, Q(A_NS, "t")).text = str(line)
    return tb


def _spc_pts(val: int) -> ET.Element:
    sp = ET.Element(Q(A_NS, "spcPts"))
    sp.set("val", str(val))
    return sp


def _empty_txbody(sp):
    tb = ET.SubElement(sp, Q(P_NS, "txBody"))
    ET.SubElement(tb, Q(A_NS, "bodyPr"))
    ET.SubElement(tb, Q(A_NS, "lstStyle"))
    ET.SubElement(tb, Q(A_NS, "p"))


# ── Native timeline builder ───────────────────────────────────────────────────
def _build_native_timeline(data: dict) -> list[ET.Element]:
    """
    Build the project timeline as native editable PPTX shapes,
    pixel-perfect matched to PID-0085-Dummy_FINAL.pptx reference.

    Slot system (3 vertical levels):
      Slot 0 = LOW_ABOVE  — name/date ABOVE bar, near bar
      Slot 1 = BELOW      — name/date BELOW bar
      Slot 2 = HIGH_ABOVE — name/date ABOVE bar, high up (overflow tier)
    """
    items = [i for i in data.get("all_items", []) if i.get("end")]
    if not items:
        return []

    dates = [i["end"] for i in items]
    t0    = min(dates).replace(day=1)
    t1    = max(dates).replace(
        day=calendar.monthrange(max(dates).year, max(dates).month)[1])
    span  = max(1, (t1 - t0).days)
    today = date.today()

    # ── Pixel-perfect layout constants (from reference FINAL.pptx) ──────────
    # Timeline bar
    BAR_LEFT  = 455700       # 0.498"  left edge of bar
    BAR_RIGHT = 8767200      # 9.588"  right edge of bar
    BAR_WIDTH = BAR_RIGHT - BAR_LEFT   # 8311500 = 9.090"
    BAR_TOP   = 1380317      # 1.510"  top of bar
    BAR_H     = 175800       # 0.192"  bar height
    BAR_BOT   = BAR_TOP + BAR_H        # 1556117 = 1.702"

    # Diamond marker
    DIAM_SZ   = 91641        # 0.100"  diamond bounding box side
    HALF_D    = DIAM_SZ // 2  # 45820

    # Date badge
    DATE_CX   = 457200       # 0.500"  wide
    DATE_CY   = 87900        # 0.096"  tall
    HALF_DX   = DATE_CX // 2  # 228600

    # Name label
    NAME_CX   = 783000       # 0.856"  wide
    NAME_CY   = 191700       # 0.210"  tall (2 lines)
    HALF_N    = NAME_CX // 2  # 391500

    # Connector (thin vertical line)
    CONN_W    = 9600         # 0.0105" wide

    # ── Y positions per slot ─────────────────────────────────────────────────
    #          Slot 0 (LOW_ABOVE)  Slot 1 (BELOW)  Slot 2 (HIGH_ABOVE)
    NAME_Y  = [952085,             1864983,         623925]   # name box top
    DATE_Y  = [1147557,            1748913,         800358]   # date badge top
    DIAM_Y  = [BAR_TOP - HALF_D,   BAR_BOT - HALF_D, BAR_TOP - HALF_D]
    #           1334497              1510297           1334497

    # Connector range calculator
    def conn_range(slot: int, dy: int):
        """Return (conn_y_top, conn_height) for connecting date badge to diamond."""
        if slot in (0, 2):  # above — connector from date-bottom to diamond-top
            ct = dy + DATE_CY       # bottom of date badge
            cb = DIAM_Y[slot]       # top of diamond
            h  = cb - ct
        else:               # below — connector from diamond-bottom to date-top
            ct = DIAM_Y[slot] + DIAM_SZ  # bottom of diamond
            cb = dy                  # top of date badge
            h  = cb - ct
        return ct, max(h, 0)

    # ── Colors (exact from reference) ────────────────────────────────────────
    BLACK  = "000000"
    WHITE  = "FFFFFF"
    GREEN  = "00B845"   # completed / on-track
    BLUE   = "0089EC"   # committed / in-progress
    RED    = "EE001E"   # at-risk / delayed / today marker
    AMBER  = "FFCD27"   # past-due
    YELLOW = "F8FF3C"   # date badge fill (exact from reference)
    GRAY33 = "333333"   # month separator ticks
    DARK   = "1A1A1A"   # dark text on badges
    FONT   = "Verizon NHG DS"

    def xp(d: date) -> int:
        return BAR_LEFT + int(BAR_WIDTH * (d - t0).days / span)

    shapes: list[ET.Element] = []
    sid = [500]
    def nid() -> int:
        sid[0] += 1
        return sid[0]

    # ── 1. Black timeline bar ─────────────────────────────────────────────────
    bar = _sp(nid(), "tl_bar", BAR_LEFT, BAR_TOP, BAR_WIDTH, BAR_H, fill=BLACK)
    _empty_txbody(bar)
    shapes.append(bar)

    # ── 2. Month labels (white bold text) + separator ticks ──────────────────
    y_m, m_m = t0.year, t0.month
    while True:
        m_date = date(y_m, m_m, 1)
        if m_date > t1:
            break
        last_day = calendar.monthrange(y_m, m_m)[1]
        m_end    = date(y_m, m_m, last_day)
        x_s      = max(BAR_LEFT,  xp(m_date))
        x_e      = min(BAR_RIGHT, xp(m_end))
        m_w      = x_e - x_s

        if m_w > int(0.08 * EMU):
            lbl = _sp(nid(), f"tl_mon_{y_m}_{m_m:02d}",
                      x_s, BAR_TOP, m_w, BAR_H, fill=None, txBox=True)
            _txbody_f(lbl, m_date.strftime("%b"), sz=630, bold=True,
                      color=WHITE, align="ctr", anchor="ctr",
                      insets=(0, 0, 0, 0), typeface=FONT)
            shapes.append(lbl)

        # Separator tick at LEFT edge of each month (except first)
        if m_date > t0:
            tx = xp(m_date)
            if BAR_LEFT < tx < BAR_RIGHT:
                TICK_W = 7800          # 0.0085"
                TICK_H = int(0.140 * EMU)  # 0.140" — full bar height + slight overhang
                tick = _sp(nid(), f"tl_tick_{y_m}_{m_m:02d}",
                           tx - TICK_W // 2, BAR_TOP,
                           TICK_W, TICK_H, fill=GRAY33)
                _empty_txbody(tick)
                shapes.append(tick)

        if m_m == 12:
            y_m += 1; m_m = 1
        else:
            m_m += 1

    # ── 3. Today progress marker (thin red bar below main bar) ────────────────
    if t0 <= today <= t1:
        xt = xp(today)
        prog_w = max(7800, xt - BAR_LEFT)
        prog = _sp(nid(), "tl_today",
                   BAR_LEFT, BAR_BOT + int(0.025 * EMU),
                   prog_w, int(0.040 * EMU), fill=RED)
        _empty_txbody(prog)
        shapes.append(prog)

    # ── 4. Cluster spreading (same date → spread labels by NAME_CX) ──────────
    sorted_items = sorted(items, key=lambda i: i["end"])
    date_groups: dict = defaultdict(list)
    for it in sorted_items:
        date_groups[it["end"]].append(it)

    STEP   = NAME_CX          # 783000 = 0.856" — one label width per step

    label_x: dict[int, int] = {}
    for d, grp in date_groups.items():
        n  = len(grp)
        cx = xp(d)
        if n == 1:
            raw = [cx]
        else:
            half = STEP * (n - 1) // 2
            raw  = [cx - half + k * STEP for k in range(n)]

        # Clamp entire cluster to stay within bar bounds
        lo = min(raw) - HALF_N
        hi = max(raw) + HALF_N
        shift = 0
        if lo < BAR_LEFT:
            shift = BAR_LEFT - lo
        elif hi > BAR_RIGHT:
            shift = BAR_RIGHT - hi
        for k, it in enumerate(grp):
            label_x[id(it)] = raw[k] + shift

    # ── 5. Slot assignment ────────────────────────────────────────────────────
    # Within each date cluster: pre-assign slots cyclically [0=LOW_A, 1=BELOW, 2=HIGH_A]
    # Then resolve cross-cluster conflicts with greedy re-assignment.
    item_slot: dict[int, int] = {}
    for d in sorted(date_groups.keys()):
        for k, it in enumerate(date_groups[d]):
            item_slot[id(it)] = k % 3   # 0, 1, 2, 0, 1, ...

    SLOT_SEP = int(0.78 * EMU)  # 0.78" min center separation in same slot
    ordered  = sorted(sorted_items, key=lambda i: label_x[id(i)])
    last_x_s = [BAR_LEFT - SLOT_SEP * 10] * 3   # last placed x per slot

    for it in ordered:
        lx   = label_x[id(it)]
        slot = item_slot[id(it)]
        if lx - last_x_s[slot] >= SLOT_SEP:
            last_x_s[slot] = lx
        else:
            # Try other slots; keep original if all occupied
            for s in [0, 1, 2]:
                if s != slot and lx - last_x_s[s] >= SLOT_SEP:
                    item_slot[id(it)] = s
                    last_x_s[s] = lx
                    break
            else:
                last_x_s[slot] = lx  # force original slot

    # ── 6. Draw each milestone ────────────────────────────────────────────────
    for it in sorted_items:
        d    = it["end"]
        lx   = label_x[id(it)]
        slot = item_slot[id(it)]
        name = it.get("name", "")
        ds   = d.strftime("%-m/%y")

        # Diamond fill color
        if it.get("is_done"):
            dc = GREEN
        elif it.get("past_due"):
            dc = AMBER
        elif it.get("status", "") in ("On Hold", "At Risk", "Delayed"):
            dc = RED
        else:
            dc = BLUE

        # Clamp diamond to bar bounds
        dbar_x = max(BAR_LEFT + HALF_D, min(BAR_RIGHT - HALF_D, lx))
        # Clamp label center to bar bounds
        clx    = max(BAR_LEFT + HALF_N, min(BAR_RIGHT - HALF_N, lx))

        # Diamond (on bar)
        dy = DIAM_Y[slot]
        diam = _sp(nid(), f"tl_d_{sid[0]}",
                   dbar_x - HALF_D, dy, DIAM_SZ, DIAM_SZ,
                   fill=dc, prst="diamond")
        _empty_txbody(diam)
        shapes.append(diam)

        # Connector (same color as diamond)
        date_y_val = DATE_Y[slot]
        ct, ch = conn_range(slot, date_y_val)
        if ch > 100:   # only draw if visible
            conn = _sp(nid(), f"tl_c_{sid[0]}",
                       clx - CONN_W // 2, ct, CONN_W, ch, fill=dc)
            _empty_txbody(conn)
            shapes.append(conn)

        # Date badge (yellow, bold)
        dbox = _sp(nid(), f"tl_dt_{sid[0]}",
                   clx - HALF_DX, date_y_val, DATE_CX, DATE_CY, fill=YELLOW)
        _txbody_f(dbox, ds, sz=570, bold=True, color=DARK,
                  align="ctr", anchor="ctr",
                  insets=(18288, 18288, 9144, 9144), typeface=FONT)
        shapes.append(dbox)

        # Name label (max 2 lines, small bold text)
        lines = textwrap.wrap(name, width=18)[:2] or [name[:24]]
        nlbl  = _sp(nid(), f"tl_n_{sid[0]}",
                    clx - HALF_N, NAME_Y[slot],
                    NAME_CX, NAME_CY,
                    fill=None, txBox=True)
        anchor_v = "b" if slot in (0, 2) else "t"
        _txbody_f(nlbl, lines, sz=500, bold=True, color=DARK,
                  align="ctr", anchor=anchor_v,
                  insets=(0, 0, 0, 0), typeface=FONT)
        shapes.append(nlbl)

    print(f"[slide_builder] Built {len(shapes)} native timeline shapes "
          f"({len(sorted_items)} milestones, 3-slot system)")
    return shapes


# ── Main entry point ──────────────────────────────────────────────────────────
def build_slide(
    data: dict[str, Any],
    texts: dict[str, str],
    timeline_png: bytes,        # kept for API compatibility, not used
    output_pptx: Path,
) -> Path:

    if not config.SOURCE_PPTX.exists():
        raise FileNotFoundError(f"PPTX not found: {config.SOURCE_PPTX}")

    today_str = date.today().strftime("%-m/%-d/%Y")
    payload = {
        "project_title": data.get("project_name", "Project Status Report"),
        "updated_date":  f"Updated: {today_str}",
        "pm_name":       f"TPD PM: {data.get('pm_name', '')}",
        "reference":     f"Reference: {data.get('reference', '')}",
        "status":        texts.get("status", ""),
        "risks":         texts.get("risks", ""),
        "summary":       texts.get("summary", ""),
    }

    # ── Read entire PPTX into memory ─────────────────────────────────────────
    with zipfile.ZipFile(config.SOURCE_PPTX, "r") as zin:
        files = {n: zin.read(n) for n in zin.namelist()}

    # ── Parse slide XML ───────────────────────────────────────────────────────
    slide_key  = "ppt/slides/slide1.xml"
    slide_root = ET.fromstring(files[slide_key])

    sp_tree = None
    for elem in slide_root.iter():
        if elem.tag.endswith("}spTree"):
            sp_tree = elem
            break
    if sp_tree is None:
        raise RuntimeError("spTree not found in slide1.xml")

    METADATA_TAGS = {Q(P_NS, "nvGrpSpPr"), Q(P_NS, "grpSpPr")}

    # ── Delete all shapes in the timeline zone (sp, grpSp, graphicFrame …) ───
    removed = 0
    for child in list(sp_tree):
        if child.tag in METADATA_TAGS:
            continue
        top = _shape_top(child)
        if top is None:
            # No position — check for Status/Trending table by text
            txt = _get_all_text(child)
            if "Current" in txt and "Trending" in txt:
                sp_tree.remove(child)
                removed += 1
            continue
        if TL_TOP_EMU <= top <= TL_BOT_EMU:
            txt = _get_all_text(child)
            if txt.strip() in PROTECTED:
                continue
            sp_tree.remove(child)
            removed += 1
    print(f"[slide_builder] Removed {removed} old timeline/header shapes")

    # ── Insert native timeline shapes ─────────────────────────────────────────
    for shape in _build_native_timeline(data):
        sp_tree.append(shape)

    # ── Replace dynamic text shapes ───────────────────────────────────────────
    sp_tag  = Q(P_NS, "sp")
    updated: set[str] = set()
    for sp in sp_tree.findall(sp_tag):
        full = _get_all_text(sp)
        if full in PROTECTED:
            continue
        for sig, field in TEXT_SIGS:
            if sig.lower() in full.lower() and field not in updated:
                val = payload.get(field, "")
                if val:
                    _replace_text(sp, val)
                    updated.add(field)
                    print(f"[slide_builder] Replaced '{field}'")
                break

    # ── Repack zip ────────────────────────────────────────────────────────────
    files[slide_key] = ET.tostring(
        slide_root, xml_declaration=True, encoding="UTF-8"
    )

    with zipfile.ZipFile(output_pptx, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, content in files.items():
            zout.writestr(name, content)

    print(f"[slide_builder] Saved → {output_pptx}")
    return output_pptx


# ── PDF export ────────────────────────────────────────────────────────────────
def export_pdf(pptx_path: Path) -> Path | None:
    import subprocess, shutil as sh
    pdf_path = pptx_path.with_suffix(".pdf")
    for cmd in ["libreoffice", "soffice"]:
        if sh.which(cmd):
            r = subprocess.run(
                [cmd, "--headless", "--convert-to", "pdf",
                 "--outdir", str(pdf_path.parent), str(pptx_path)],
                capture_output=True, timeout=60,
            )
            if r.returncode == 0 and pdf_path.exists():
                print(f"[slide_builder] PDF → {pdf_path}")
                return pdf_path
    return None
