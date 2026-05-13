"""
Generate the project timeline as an SVG.
Zero dependencies — pure Python stdlib.
Ports the exact rendering logic from PID-0085-Timeline-auto.html.

Returns SVG bytes that can be:
  - Embedded directly in a PPTX (modern PowerPoint/Google Slides support SVG)
  - Written to a .svg file for browser preview
"""
from __future__ import annotations

import textwrap
from calendar import monthrange
from datetime import date
from pathlib import Path
from typing import Any


# ── Colors ────────────────────────────────────────────────────────────────────
BG     = "#F6F0E2"
NAVY   = "#000000"
BLUE   = "#0089EC"
GREEN  = "#00B845"
AMBER  = "#FFCD27"
RED    = "#EE001E"
YELLOW = "#F8FF3C"
DARK   = "#1a1a1a"
WHITE  = "#ffffff"

# ── Layout (matches HTML exactly) ─────────────────────────────────────────────
W, H    = 1150, 380
ML, MR  = 44, 44
TLW     = W - ML - MR
BAR_Y   = 172
BAR_H   = 26
BMID    = BAR_Y + BAR_H / 2
LVL     = [55, 100, 145]
MIN_GAP = 56
LH      = 13
FS      = 10.5
MN      = ["Jan","Feb","Mar","Apr","May","Jun",
           "Jul","Aug","Sep","Oct","Nov","Dec"]


# ── Helpers ───────────────────────────────────────────────────────────────────
def _e(s: str) -> str:
    """Escape text for SVG."""
    return (s.replace("&", "&amp;")
             .replace("<", "&lt;")
             .replace(">", "&gt;")
             .replace('"', "&quot;"))


def generate_timeline_svg(
    items: list[dict[str, Any]],
    t0: date | None = None,
    t1: date | None = None,
    output_path: Path | None = None,
) -> bytes:
    """
    Build and return an SVG timeline as UTF-8 bytes.

    items: list of project task/milestone dicts with keys:
      name, end (date), type, is_done, is_wip
    """
    if t0 is None:
        t0 = date.today().replace(day=1)
    if t1 is None:
        t1 = date(t0.year, 12, 31)

    today  = date.today()
    tspan  = (t1 - t0).days or 1

    # Filter + sort
    tl = [it for it in items if isinstance(it.get("end"), date) and t0 <= it["end"] <= t1]
    tl = sorted(tl, key=lambda x: x["end"])

    # Assign above/below alternately
    for i, it in enumerate(tl):
        it = dict(it)   # don't mutate original
        tl[i] = it
        it["_pos"] = "above" if i % 2 == 0 else "below"
        it["_m"]   = (it.get("type", "Task") == "Milestone")

    # Assign stagger levels (avoid label overlap)
    def date_x(d: date) -> float:
        return ML + ((d - t0).days / tspan) * TLW

    slots: dict[str, list[float]] = {"above": [], "below": []}
    for it in tl:
        x   = date_x(it["end"])
        arr = slots[it["_pos"]]
        lvl = -1
        for i, sx in enumerate(arr):
            if x - sx >= MIN_GAP:
                lvl = i; arr[i] = x; break
        if lvl == -1:
            lvl = len(arr); arr.append(x)
        it["_lvl"] = min(lvl, len(LVL) - 1)

    # ── SVG elements accumulator ──────────────────────────────────────────────
    E: list[str] = []

    def diamond(cx, cy, r, color):
        pts = f"{cx},{cy-r} {cx+r},{cy} {cx},{cy+r} {cx-r},{cy}"
        E.append(f'<polygon points="{pts}" fill="{color}"/>')

    def txt(x, y, s, sz, color, weight="normal", anchor="middle"):
        E.append(
            f'<text x="{x:.1f}" y="{y:.1f}" text-anchor="{anchor}" '
            f'font-size="{sz}" font-weight="{weight}" fill="{color}" '
            f'font-family="Calibri,Arial,sans-serif">{_e(s)}</text>'
        )

    def date_box(cx, cy, s):
        bw = max(len(s) * 6.8 + 14, 36)
        bh = 16
        E.append(
            f'<rect x="{cx - bw/2:.1f}" y="{cy - 12:.1f}" '
            f'width="{bw:.1f}" height="{bh}" rx="3" fill="{YELLOW}"/>'
        )
        txt(cx, cy, s, 10, NAVY, "bold")

    # ── SVG open ──────────────────────────────────────────────────────────────
    E.append(
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'width="{W}" height="{H}" viewBox="0 0 {W} {H}">'
    )
    E.append(f'<rect width="{W}" height="{H}" fill="{BG}"/>')

    # Timeline bar
    E.append(
        f'<rect x="{ML}" y="{BAR_Y}" width="{TLW}" height="{BAR_H}" '
        f'fill="{NAVY}" rx="4"/>'
    )

    # Red progress line to today
    today_x = min(date_x(today), ML + TLW)
    E.append(
        f'<line x1="{ML}" y1="{BMID:.1f}" x2="{today_x:.1f}" y2="{BMID:.1f}" '
        f'stroke="{RED}" stroke-width="5" stroke-linecap="round"/>'
    )

    # Month labels + separators
    mo_offset = t0.month - 1
    yr = t0.year
    while True:
        year  = yr + mo_offset // 12
        month = mo_offset % 12 + 1
        d1 = date(year, month, 1)
        if d1 > t1:
            break
        _, last = monthrange(year, month)
        d2   = date(year, month, last)
        x1   = ML + ((d1 - t0).days / tspan) * TLW
        x2   = ML + ((d2 - t0).days / tspan) * TLW
        cx   = (x1 + x2) / 2
        if d1 > t0:
            E.append(
                f'<line x1="{x1:.1f}" y1="{BAR_Y+4}" '
                f'x2="{x1:.1f}" y2="{BAR_Y+BAR_H-4}" '
                f'stroke="rgba(255,255,255,0.25)" stroke-width="1"/>'
            )
        txt(cx, BMID + 4, MN[month - 1], 11, WHITE, "bold")
        mo_offset += 1

    # ── Draw tasks & milestones ───────────────────────────────────────────────
    for it in tl:
        tx    = date_x(it["end"])
        above = it["_pos"] == "above"
        DR    = 9 if it["_m"] else 7
        ds    = f"{it['end'].month}/{it['end'].day}"
        name  = it["name"]
        wt    = "bold" if it["_m"] else "normal"
        stem  = LVL[it["_lvl"]]

        # Color
        if it.get("is_done"):
            color = GREEN
        elif it.get("is_wip"):
            color = AMBER
        elif above:
            color = BLUE
        else:
            color = GREEN

        sy0 = BAR_Y if above else BAR_Y + BAR_H
        sy1 = BAR_Y - stem if above else BAR_Y + BAR_H + stem

        # Stem line
        E.append(
            f'<line x1="{tx:.1f}" y1="{sy0}" x2="{tx:.1f}" y2="{sy1}" '
            f'stroke="{color}" stroke-width="1.3"/>'
        )
        diamond(tx, sy0, DR, color)

        # Wrap long names
        lines = textwrap.wrap(name, width=20) or [name]

        if above:
            name_btm = sy1 - 5
            for i, line in enumerate(lines):
                y_pos = name_btm - (len(lines) - 1 - i) * LH
                txt(tx, y_pos, line, FS, DARK, wt)
            date_box(tx, name_btm - len(lines) * LH - 3, ds)
        else:
            name_top = sy1 + 5
            for i, line in enumerate(lines):
                txt(tx, name_top + (i + 1) * LH, line, FS, DARK, wt)
            date_box(tx, name_top + (len(lines) + 1) * LH + 2, ds)

    E.append("</svg>")

    svg_bytes = "\n".join(E).encode("utf-8")

    if output_path:
        Path(output_path).write_bytes(svg_bytes)
        print(f"[timeline_builder] Saved SVG → {output_path}")

    return svg_bytes


def svg_to_png(svg_bytes: bytes, width: int = 1800) -> bytes | None:
    """
    Convert SVG bytes to PNG bytes using macOS sips (built-in, zero deps).
    Returns None if conversion fails.
    """
    import subprocess
    import tempfile
    import os

    with tempfile.TemporaryDirectory() as tmp:
        svg_path = os.path.join(tmp, "timeline.svg")
        png_path = os.path.join(tmp, "timeline.png")

        with open(svg_path, "wb") as f:
            f.write(svg_bytes)

        try:
            r = subprocess.run(
                ["sips", "-s", "format", "png",
                 "--resampleWidth", str(width),
                 svg_path, "--out", png_path],
                capture_output=True, timeout=15,
            )
            if r.returncode == 0 and os.path.exists(png_path):
                with open(png_path, "rb") as f:
                    png_bytes = f.read()
                print(f"[timeline_builder] SVG→PNG via sips ({len(png_bytes)//1024}KB)")
                return png_bytes
        except Exception as e:
            print(f"[timeline_builder] sips conversion failed: {e}")

    return None


def generate_timeline_png(
    items: list[dict],
    t0=None,
    t1=None,
    output_path=None,
) -> bytes:
    """Generate timeline as PNG (falls back to SVG if sips unavailable)."""
    svg = generate_timeline_svg(items, t0, t1, output_path=None)
    png = svg_to_png(svg)
    if png:
        if output_path:
            Path(output_path).with_suffix(".png").write_bytes(png)
        return png
    # Fallback to SVG if conversion failed
    if output_path:
        Path(output_path).write_bytes(svg)
    return svg


def generate_timeline_html(
    items: list[dict],
    project_name: str = "Project Timeline",
    t0: date | None = None,
    t1: date | None = None,
    output_path: Path | None = None,
) -> bytes:
    """
    Generate a self-contained HTML timeline file (same visual as PID-0085-Timeline.html)
    dynamically populated from live Smartsheet data.
    Returns UTF-8 bytes.
    """
    import json

    today = date.today()
    if t0 is None:
        dates = [it["end"] for it in items if isinstance(it.get("end"), date)]
        t0 = (min(dates).replace(day=1)) if dates else today.replace(day=1)
    if t1 is None:
        dates = [it["end"] for it in items if isinstance(it.get("end"), date)]
        last  = max(dates) if dates else today
        t1    = date(last.year, 12, 31)

    # Build JS-compatible TASKS array
    tl = [it for it in items if isinstance(it.get("end"), date)]
    tl.sort(key=lambda x: x["end"])

    tasks_js = []
    for i, it in enumerate(tl):
        name = it.get("name", "")
        # Wrap long names with \n for the HTML label
        words = name.split()
        line1, line2 = [], []
        for w in words:
            if len(" ".join(line1 + [w])) <= 22:
                line1.append(w)
            else:
                line2.append(w)
        label = " ".join(line1) + ("\n" + " ".join(line2) if line2 else "")

        pos  = "above" if i % 2 == 0 else "below"
        done = it.get("is_done", False)
        ms   = it.get("type", "Task") == "Milestone"
        tasks_js.append({
            "n": label,
            "d": it["end"].strftime("%Y-%m-%d"),
            "p": pos,
            "m": ms,
            "done": done,
        })

    tasks_json = json.dumps(tasks_js, indent=2)
    today_str  = today.strftime("%Y-%m-%d")
    t0_str     = t0.strftime("%Y-%m-%d")
    t1_str     = t1.strftime("%Y-%m-%d")

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>{project_name} — Timeline</title>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: Calibri, 'Segoe UI', Arial, sans-serif;
    background: #F6F0E2;
    padding: 28px 36px 36px;
  }}
  h2 {{ color: #000000; font-size: 15px; margin-bottom: 2px; font-weight: 700; }}
  p.sub {{ font-size: 10.5px; color: #666; margin-bottom: 22px; }}
  #chart {{ overflow-x: auto; }}
  svg {{ display: block; }}
</style>
</head>
<body>

<h2>{project_name} — Timeline</h2>
<p class="sub">Generated {today.strftime("%-m/%-d/%Y")} from Smartsheet data.</p>
<div id="chart"></div>

<script>
const TASKS = {tasks_json};

const W      = 1150;
const H      = 370;
const ML     = 40;
const MR     = 40;
const TLW    = W - ML - MR;

const BAR_Y  = 168;
const BAR_H  = 26;
const BMID   = BAR_Y + BAR_H / 2;

const T0     = new Date("{t0_str}");
const T1     = new Date("{t1_str}");
const TSPAN  = T1 - T0;

const TODAY  = new Date("{today_str}");

const LVL    = [55, 100, 145];
const LH     = 13;
const FS     = 10.5;
const MIN_GAP= 56;

const NAVY   = "#000000";
const BLUE   = "#0089EC";
const GREEN  = "#00B845";
const RED    = "#EE001E";
const YELLOW = "#F8FF3C";

function dateX(dateStr) {{
  return ML + ((new Date(dateStr) - T0) / TSPAN) * TLW;
}}
function fmt(dateStr) {{
  const d = new Date(dateStr);
  return `${{d.getMonth()+1}}/${{String(d.getFullYear()).slice(2)}}`;
}}

TASKS.sort((a, b) => new Date(a.d) - new Date(b.d));
const slotRight = {{ above: [], below: [] }};

TASKS.forEach(t => {{
  const x   = dateX(t.d);
  const arr = slotRight[t.p];
  let lvl   = -1;
  for (let i = 0; i < arr.length; i++) {{
    if (x - arr[i] >= MIN_GAP) {{ lvl = i; arr[i] = x; break; }}
  }}
  if (lvl === -1) {{ lvl = arr.length; arr.push(x); }}
  t.lvl = Math.min(lvl, LVL.length - 1);
}});

const E = [];
const o = s => E.push(s);

function diamond(cx, cy, r, color) {{
  o(`<polygon points="${{cx}},${{cy-r}} ${{cx+r}},${{cy}} ${{cx}},${{cy+r}} ${{cx-r}},${{cy}}" fill="${{color}}"/>`);
}}
function text(x, y, txt, size, color, weight, anchor) {{
  const fw = weight || "normal";
  const ta = anchor || "middle";
  o(`<text x="${{x}}" y="${{y}}" text-anchor="${{ta}}" font-size="${{size}}" font-weight="${{fw}}" fill="${{color}}" font-family="Calibri,Arial,sans-serif">${{txt}}</text>`);
}}
function dateBox(cx, cy, txt) {{
  const bw = txt.length * 6.8 + 14;
  const bh = 16;
  o(`<rect x="${{(cx - bw/2).toFixed(1)}}" y="${{(cy - 12).toFixed(1)}}" width="${{bw.toFixed(1)}}" height="${{bh}}" rx="3" fill="${{YELLOW}}"/>`);
  o(`<text x="${{cx}}" y="${{cy}}" text-anchor="middle" font-size="10" font-weight="bold" fill="#000000" font-family="Calibri,Arial,sans-serif">${{txt}}</text>`);
}}

o(`<svg xmlns="http://www.w3.org/2000/svg" width="${{W}}" height="${{H}}" viewBox="0 0 ${{W}} ${{H}}">`);
o(`<rect width="${{W}}" height="${{H}}" fill="#F6F0E2"/>`);
o(`<rect x="${{ML}}" y="${{BAR_Y}}" width="${{TLW}}" height="${{BAR_H}}" fill="${{NAVY}}" rx="4"/>`);

const todayX = Math.min(dateX(TODAY.toISOString().slice(0,10)), ML + TLW);
o(`<line x1="${{ML}}" y1="${{BMID}}" x2="${{todayX}}" y2="${{BMID}}" stroke="${{RED}}" stroke-width="5" stroke-linecap="round"/>`);

const MONTHS_LBL = ["","","","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const startM = T0.getMonth() + 1;
const endM   = T1.getMonth() + 1;
const startY = T0.getFullYear();
const endY   = T1.getFullYear();
for (let yr = startY; yr <= endY; yr++) {{
  const mStart = (yr === startY) ? startM : 1;
  const mEnd   = (yr === endY)   ? endM   : 12;
  for (let m = mStart; m <= mEnd; m++) {{
    const d1  = new Date(yr, m - 1, 1);
    const d2  = new Date(yr, m,     0);
    const x1  = ML + ((d1 - T0) / TSPAN) * TLW;
    const x2  = ML + ((d2 - T0) / TSPAN) * TLW;
    const cx  = (x1 + x2) / 2;
    if (d1 > T0) {{
      o(`<line x1="${{x1}}" y1="${{BAR_Y+4}}" x2="${{x1}}" y2="${{BAR_Y+BAR_H-4}}" stroke="rgba(255,255,255,0.25)" stroke-width="1"/>`);
    }}
    const lbl = MONTHS_LBL[m] || String(m);
    text(cx, BMID + 4, lbl, 11, "white", "bold");
  }}
}}

TASKS.forEach(t => {{
  const tx    = dateX(t.d);
  const above = t.p === "above";
  const DR    = t.m ? 9 : 7;
  const color = t.done ? GREEN : (above ? BLUE : GREEN);
  const stem  = LVL[t.lvl];
  const lines = t.n.split("\\n");
  const dateStr = fmt(t.d);

  const sy0 = above ? BAR_Y : BAR_Y + BAR_H;
  const sy1 = above ? BAR_Y - stem : BAR_Y + BAR_H + stem;
  o(`<line x1="${{tx}}" y1="${{sy0}}" x2="${{tx}}" y2="${{sy1}}" stroke="${{color}}" stroke-width="1.3"/>`);
  diamond(tx, sy0, DR, color);

  if (above) {{
    const nameBtm = sy1 - 5;
    for (let li = 0; li < lines.length; li++) {{
      const y = nameBtm - (lines.length - 1 - li) * LH;
      text(tx, y, lines[li], FS, "#1a1a1a");
    }}
    dateBox(tx, nameBtm - lines.length * LH - 3, dateStr);
  }} else {{
    const nameTop = sy1 + 5;
    for (let li = 0; li < lines.length; li++) {{
      text(tx, nameTop + (li + 1) * LH, lines[li], FS, "#1a1a1a");
    }}
    dateBox(tx, nameTop + (lines.length + 1) * LH + 2, dateStr);
  }}
}});

o(`</svg>`);
document.getElementById("chart").innerHTML = E.join("");
</script>

<div style="margin-top:16px; display:flex; gap:28px; font-size:10.5px; color:#000000; flex-wrap:wrap;">
  <span style="display:flex;align-items:center;gap:6px;">
    <svg width="14" height="14"><polygon points="7,0 14,7 7,14 0,7" fill="#0089EC"/></svg>
    Not Started (above bar)
  </span>
  <span style="display:flex;align-items:center;gap:6px;">
    <svg width="14" height="14"><polygon points="7,0 14,7 7,14 0,7" fill="#00B845"/></svg>
    Not Started (below bar) / Complete
  </span>
  <span style="display:flex;align-items:center;gap:6px;">
    <svg width="40" height="14"><line x1="0" y1="7" x2="40" y2="7" stroke="#EE001E" stroke-width="4"/></svg>
    Progress to today
  </span>
  <span style="display:flex;align-items:center;gap:6px;">
    <svg width="32" height="16"><rect x="0" y="0" width="32" height="16" rx="3" fill="#F8FF3C"/></svg>
    Date badge
  </span>
</div>

</body>
</html>
"""
    html_bytes = html.encode("utf-8")
    if output_path:
        Path(output_path).write_bytes(html_bytes)
        print(f"[timeline_builder] Saved HTML → {output_path}")
    return html_bytes


# ── Quick test ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    sample = [
        {"name": "Project Kickoff",    "end": date(2026,  3, 19), "type": "Task",      "is_done": True,  "is_wip": False},
        {"name": "Requirements Review","end": date(2026,  3, 26), "type": "Task",      "is_done": False, "is_wip": False},
        {"name": "Business Case",      "end": date(2026,  5, 19), "type": "Milestone", "is_done": False, "is_wip": False},
        {"name": "GTS FD Approval",    "end": date(2026,  5, 26), "type": "Task",      "is_done": False, "is_wip": False},
        {"name": "Big Team Review",    "end": date(2026,  6,  8), "type": "Task",      "is_done": False, "is_wip": False},
        {"name": "Committed Date GTS", "end": date(2026,  9, 22), "type": "Milestone", "is_done": False, "is_wip": False},
        {"name": "Committed Date TPD", "end": date(2026,  9, 22), "type": "Milestone", "is_done": False, "is_wip": False},
        {"name": "Prod Testing",       "end": date(2026, 11,  3), "type": "Milestone", "is_done": False, "is_wip": False},
        {"name": "Tech Operational",   "end": date(2026, 12,  8), "type": "Milestone", "is_done": False, "is_wip": False},
        {"name": "Project End",        "end": date(2026, 12, 22), "type": "Milestone", "is_done": False, "is_wip": False},
    ]
    out = Path("output/test_timeline.svg")
    out.parent.mkdir(exist_ok=True)
    generate_timeline_svg(sample, date(2026, 3, 1), date(2026, 12, 31), output_path=out)
    print(f"Open in browser: file://{out.resolve()}")
