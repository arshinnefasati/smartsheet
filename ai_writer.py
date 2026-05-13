"""
Generate VP-ready text for the three slide sections.
Uses Claude API via stdlib urllib (no anthropic SDK needed).
Falls back to rule-based generation if no API key.
"""
from __future__ import annotations
import json
import urllib.request
import urllib.error
from datetime import date
from typing import Any

import config


# ── Rule-based (no AI) ────────────────────────────────────────────────────────
def _status_from_data(data: dict) -> str:
    all_items = data.get("all_items", [])
    complete  = [t["name"] for t in all_items if t.get("is_done")]
    overdue   = [t["name"] for t in all_items if t.get("past_due") and not t.get("is_done")]
    in_prog   = [t["name"] for t in all_items if t.get("is_wip")]
    next_ms   = next((m for m in data.get("milestones", []) if not m.get("is_done")), None)

    lines = []
    if complete:
        lines.append(f"• COMPLETE: {', '.join(complete[:2])}.")
    if overdue:
        lines.append(f"• PAST DUE / ACTION NEEDED: {', '.join(overdue[:2])} — requires immediate follow-up.")
    if in_prog:
        lines.append(f"• IN PROGRESS: {', '.join(in_prog[:2])}.")
    if next_ms:
        d = next_ms["end"]
        ds = d.strftime("%-m/%-d/%Y") if isinstance(d, date) else str(d)
        lines.append(f"• NEXT MILESTONE: {next_ms['name']} — target {ds}.")
    if not lines:
        lines.append("• Project is On Track. No blockers identified at this time.")
    return "\n".join(lines)


def _risks_from_data(data: dict) -> str:
    overdue = [t["name"] for t in data.get("all_items", [])
               if t.get("past_due") and not t.get("is_done")]
    notes   = [n for n in data.get("notes", []) if n.strip()]
    lines   = []
    if overdue:
        lines.append(
            f"• Schedule risk: {len(overdue)} task(s) past due "
            f"({', '.join(overdue[:2])}). Escalation may be required."
        )
    for n in notes[:2]:
        lines.append(f"• {n.strip()}")
    if not lines:
        lines.append("• No critical risks or dependencies identified at this time.")
    return "\n".join(lines)


def _summary_from_data(data: dict) -> str:
    name    = data.get("project_name", "This project")
    items   = data.get("all_items", [])
    total   = len(items)
    done    = sum(1 for t in items if t.get("is_done"))
    status  = data.get("project_status") or "In Progress"
    impacted = data.get("impacted") or []
    teams   = ", ".join(impacted[:4]) if impacted else "GTS, NI, and TPD"
    return (
        f"{name} — {done} of {total} deliverables complete. "
        f"Project teams include {teams}. "
        "Reference dashboard for full schedule and milestone tracking."
    )


# ── Claude API via stdlib urllib ──────────────────────────────────────────────
def _claude(prompt: str, max_tokens: int = 300) -> str:
    payload = json.dumps({
        "model":      "claude-opus-4-6",
        "max_tokens": max_tokens,
        "messages":   [{"role": "user", "content": prompt}],
    }).encode()

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=payload,
        headers={
            "x-api-key":         config.ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type":      "application/json",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = json.loads(resp.read())
    return body["content"][0]["text"].strip()


def _ai_status(data: dict) -> str:
    all_items = data.get("all_items", [])
    complete  = [t["name"] for t in all_items if t.get("is_done")]
    overdue   = [t["name"] for t in all_items if t.get("past_due") and not t.get("is_done")]
    next_ms   = next((m for m in data.get("milestones", []) if not m.get("is_done")), None)
    return _claude(
        f"You are a Verizon senior project manager writing a VP status update.\n"
        f"Project: {data.get('project_name','N/A')}\n"
        f"Completed: {complete or 'None'}\n"
        f"Past-due: {overdue or 'None'}\n"
        f"Next milestone: {next_ms['name'] + ' on ' + str(next_ms['end']) if next_ms else 'N/A'}\n\n"
        f"Write 2-4 bullet points (format: • LABEL: detail.). Verizon professional tone. No markdown headers.",
        300
    )


def _ai_risks(data: dict) -> str:
    notes   = data.get("notes", [])
    overdue = [t["name"] for t in data.get("all_items", [])
               if t.get("past_due") and not t.get("is_done")]
    return _claude(
        f"You are a Verizon project manager writing Risks & Dependencies for a VP.\n"
        f"Overdue items: {overdue or 'None'}\n"
        f"Notes: {notes or 'None'}\n\n"
        f"Write 2-3 bullet points (• Risk. Mitigation needed.). Concise and professional.",
        250
    )


def _ai_summary(data: dict) -> str:
    items    = data.get("all_items", [])
    impacted = ", ".join(data.get("impacted", [])[:4]) or "GTS, NI, TPD"
    return _claude(
        f"You are a Verizon PM writing a 2-sentence project summary for a VP.\n"
        f"Project: {data.get('project_name','N/A')}\n"
        f"Status: {data.get('project_status','In Progress')}\n"
        f"Progress: {sum(1 for t in items if t.get('is_done'))}/{len(items)} deliverables complete.\n"
        f"Impacted teams: {impacted}\n\n"
        f"Write exactly 2 sentences: what the project does and its current state.",
        150
    )


# ── Public API ────────────────────────────────────────────────────────────────
def generate_texts(data: dict) -> dict[str, str]:
    if config.USE_AI:
        print("[ai_writer] Using Claude API")
        try:
            return {
                "status":  _ai_status(data),
                "risks":   _ai_risks(data),
                "summary": _ai_summary(data),
            }
        except Exception as e:
            print(f"[ai_writer] Claude error: {e} — using rule-based fallback")

    print("[ai_writer] Using rule-based text")
    return {
        "status":  _status_from_data(data),
        "risks":   _risks_from_data(data),
        "summary": _summary_from_data(data),
    }
