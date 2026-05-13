"""Central configuration — reads from .env file (no external deps)."""
import os
from pathlib import Path

# ── Simple .env reader (no python-dotenv needed) ─────────────────────────────
def _load_env(path: Path):
    if not path.exists():
        return
    for line in path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, val = line.partition("=")
        key = key.strip()
        val = val.strip().strip('"').strip("'")
        if key and key not in os.environ:      # don't override real env vars
            os.environ[key] = val

_load_env(Path(__file__).parent / ".env")

# ── API Keys ──────────────────────────────────────────────────────────────────
SMARTSHEET_API_KEY  = os.getenv("SMARTSHEET_API_KEY", "")
ANTHROPIC_API_KEY   = os.getenv("ANTHROPIC_API_KEY", "")
SMARTSHEET_SHEET_ID = os.getenv("SMARTSHEET_SHEET_ID", "")

# ── Optional overrides ────────────────────────────────────────────────────────
# If Smartsheet truncates the sheet name, set this to the full project name
PROJECT_NAME_OVERRIDE = os.getenv("PROJECT_NAME", "")

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

SOURCE_PPTX = Path(os.getenv(
    "SOURCE_PPTX",
    "/Users/nefasar/Desktop/STEVE - Smartsheet/Existing Google Slide One Page Summary.pptx"
))
SOURCE_XLSX = Path(os.getenv(
    "SOURCE_XLSX",
    "/Users/nefasar/Desktop/STEVE - Smartsheet/PID-0085-Dummy Smartsheet Proj - Timeline Extract (2).xlsx"
))

# ── Modes ─────────────────────────────────────────────────────────────────────
USE_AI    = bool(ANTHROPIC_API_KEY)
DEMO_MODE = not bool(SMARTSHEET_API_KEY)
