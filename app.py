"""
Solution 1 — Python Agent Dashboard
Zero-dependency HTTP server (stdlib only + Pillow).

Run:   python3 app.py
Open:  http://localhost:5001
"""
from __future__ import annotations

import json
import mimetypes
import os
import traceback
import urllib.parse
from datetime import date
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

import config
import smartsheet_client
import timeline_builder
import ai_writer
import slide_builder

PORT = 5001


# ── Request handler ───────────────────────────────────────────────────────────
class Handler(BaseHTTPRequestHandler):

    def log_message(self, fmt, *args):
        pass  # suppress default access log spam

    # ── Route dispatch ────────────────────────────────────────────────────────
    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path

        if path == "/":
            self._serve_index()
        elif path == "/status":
            self._serve_status()
        elif path == "/reports":
            self._serve_reports()
        elif path == "/sheets":
            self._serve_sheets()
        elif path.startswith("/download/"):
            fname = urllib.parse.unquote(path[len("/download/"):])
            self._serve_file(config.OUTPUT_DIR / fname)
        else:
            self._send_404()

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        if path == "/generate":
            self._handle_generate()
        elif path == "/resolve":
            self._handle_resolve()
        elif path == "/upload":
            self._handle_upload()
        elif path == "/push-dashboard":
            self._handle_push_dashboard()
        else:
            self._send_404()

    # ── Index page ────────────────────────────────────────────────────────────
    def _serve_index(self):
        html_path = Path(__file__).parent / "templates" / "index.html"
        html = html_path.read_text(encoding="utf-8")
        # Inject template variables (simple string replace, no Jinja needed)
        mode  = "DEMO (xlsx)" if config.DEMO_MODE else "LIVE (Smartsheet API)"
        ai_on = "ON (Claude)"  if config.USE_AI    else "OFF (rule-based)"
        html  = html.replace("{{ mode }}", mode).replace("{{ ai_on }}", ai_on)
        self._send(200, html.encode(), "text/html")

    # ── /status ───────────────────────────────────────────────────────────────
    def _serve_status(self):
        self._json(200, {
            "demo_mode":   config.DEMO_MODE,
            "ai_enabled":  config.USE_AI,
            "source_pptx": str(config.SOURCE_PPTX),
            "pptx_exists": config.SOURCE_PPTX.exists(),
            "xlsx_exists": config.SOURCE_XLSX.exists(),
            "output_dir":  str(config.OUTPUT_DIR),
        })

    # ── /sheets ───────────────────────────────────────────────────────────────
    def _serve_sheets(self):
        """Return list of Smartsheet sheets accessible with the configured API key."""
        if not config.SMARTSHEET_API_KEY:
            self._json(200, {"sheets": [], "error": "No API key configured"})
            return
        try:
            sheets = smartsheet_client.list_sheets(config.SMARTSHEET_API_KEY)
            self._json(200, {"sheets": sheets})
        except Exception as e:
            self._json(200, {"sheets": [], "error": str(e)})

    # ── /upload ───────────────────────────────────────────────────────────────
    def _handle_upload(self):
        """Accept a multipart/form-data xlsx file upload, save to uploads/, preview data."""
        import cgi, io
        try:
            ctype = self.headers.get("Content-Type", "")
            length = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(length)

            # Parse multipart
            environ = {"REQUEST_METHOD": "POST", "CONTENT_TYPE": ctype, "CONTENT_LENGTH": str(length)}
            fs = cgi.FieldStorage(fp=io.BytesIO(raw), environ=environ)
            item = fs["file"]
            if not item.filename or not item.filename.endswith(".xlsx"):
                self._json(400, {"ok": False, "error": "Please upload an .xlsx file"})
                return

            # Save to uploads/
            uploads_dir = Path(__file__).parent / "uploads"
            uploads_dir.mkdir(exist_ok=True)
            safe = item.filename.replace(" ", "_").replace("/", "-")
            dest = uploads_dir / safe
            dest.write_bytes(item.file.read())

            # Preview data (use xlsx's own project name, not the .env override)
            data = smartsheet_client.load_from_xlsx(dest, use_name_override=False)
            self._json(200, {
                "ok":        True,
                "filename":  safe,
                "project":   data["project_name"],
                "tasks":     len(data["tasks"]),
                "milestones": len(data["milestones"]),
            })
        except Exception as e:
            self._json(500, {"ok": False, "error": str(e)})

    # ── /resolve ──────────────────────────────────────────────────────────────
    def _handle_resolve(self):
        """Resolve a Smartsheet sheet URL or dashboard URL to a numeric sheet ID."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length).decode())
            url = (body.get("url") or "").strip()
        except Exception:
            self._json(400, {"error": "Invalid request body"})
            return

        if not url:
            self._json(400, {"error": "No URL provided"})
            return
        if not config.SMARTSHEET_API_KEY:
            self._json(400, {"error": "No API key configured — set SMARTSHEET_API_KEY in .env"})
            return

        try:
            sheet_id = smartsheet_client.resolve_to_sheet_id(url, config.SMARTSHEET_API_KEY)
            # Look up the sheet name
            try:
                sheets = smartsheet_client.list_sheets(config.SMARTSHEET_API_KEY)
                name = next((s["name"] for s in sheets if str(s["id"]) == str(sheet_id)), sheet_id)
            except Exception:
                name = sheet_id
            self._json(200, {"sheet_id": sheet_id, "name": name})
        except Exception as e:
            self._json(200, {"error": str(e)})

    # ── /push-dashboard ───────────────────────────────────────────────────────
    def _handle_push_dashboard(self):
        """Push the timeline HTML to a Smartsheet dashboard as a Web Content widget."""
        try:
            length = int(self.headers.get("Content-Length", 0))
            body   = json.loads(self.rfile.read(length).decode())
            dashboard_url      = (body.get("dashboard_url") or "").strip()
            timeline_public_url = (body.get("timeline_url") or "").strip()
            widget_title       = body.get("widget_title") or "BPE Timeline"

            if not dashboard_url:
                self._json(400, {"ok": False, "error": "dashboard_url required"})
                return
            if not timeline_public_url:
                self._json(400, {"ok": False, "error": "timeline_url required — start ngrok first"})
                return
            if not config.SMARTSHEET_API_KEY:
                self._json(400, {"ok": False, "error": "No API key configured"})
                return

            result = smartsheet_client.push_timeline_to_dashboard(
                dashboard_url, timeline_public_url, widget_title,
                config.SMARTSHEET_API_KEY,
            )
            self._json(200, {"ok": True, "result": result})
        except Exception as e:
            self._json(200, {"ok": False, "error": str(e)})

    # ── /reports ──────────────────────────────────────────────────────────────
    def _serve_reports(self):
        files = sorted(config.OUTPUT_DIR.glob("*.pptx"), reverse=True)[:10]
        self._json(200, [
            {"name": f.name, "size_kb": round(f.stat().st_size / 1024, 1)}
            for f in files
        ])

    # ── /generate ────────────────────────────────────────────────────────────
    def _handle_generate(self):
        try:
            print("\n" + "="*56)
            print("GENERATING REPORT")
            print("="*56)

            # Read optional sheet_id / uploaded_file / published_url from POST body (JSON)
            sheet_id_override = None
            uploaded_file     = None
            published_url     = None
            try:
                length = int(self.headers.get("Content-Length", 0))
                if length > 0:
                    body = json.loads(self.rfile.read(length).decode())
                    sheet_id_override = body.get("sheet_id") or None
                    uploaded_file     = body.get("uploaded_file") or None
                    published_url     = body.get("published_url") or None
            except Exception:
                pass

            # 1. Fetch data — priority: published_url > uploaded xlsx > sheet_id > default
            if published_url:
                print(f"  Using published URL: {published_url[:70]}")
                data = smartsheet_client.load_from_published_url(published_url)
            elif uploaded_file:
                xlsx_path = Path(__file__).parent / "uploads" / uploaded_file
                print(f"  Using uploaded file: {uploaded_file}")
                data = smartsheet_client.load_from_xlsx(xlsx_path, use_name_override=False)
            else:
                if sheet_id_override:
                    print(f"  Sheet override: {sheet_id_override}")
                data = smartsheet_client.get_project_data(sheet_id=sheet_id_override)
            print(f"  Tasks: {len(data['tasks'])}  Milestones: {len(data['milestones'])}")

            # 2. Timeline bounds
            all_items = data["all_items"]
            dates = [it["end"] for it in all_items if isinstance(it.get("end"), date)]
            if dates:
                t0 = min(dates).replace(day=1)
                t1 = date(max(dates).year, 12, 31)
            else:
                t0 = date(2026, 3, 1)
                t1 = date(2026, 12, 31)

            # 3. Timeline PNG
            timeline_png = timeline_builder.generate_timeline_png(
                all_items, t0=t0, t1=t1
            )
            print(f"  Timeline PNG: {len(timeline_png):,} bytes")

            # 4. Text (AI or rule-based)
            texts = ai_writer.generate_texts(data)
            print(f"  Status text: {len(texts['status'])} chars")

            # 5. Build PPTX
            ts        = date.today().strftime("%Y-%m-%d")
            safe_name = (data["project_name"]
                         .replace(" ", "_").replace("/", "-"))[:40]
            out_name  = f"{safe_name}_{ts}.pptx"
            out_path  = config.OUTPUT_DIR / out_name

            slide_builder.build_slide(data, texts, timeline_png, out_path)

            # 6. Timeline HTML (standalone, always generated)
            tl_name = out_name.replace(".pptx", "_timeline.html")
            tl_path = config.OUTPUT_DIR / tl_name
            timeline_builder.generate_timeline_html(
                data["all_items"],
                project_name=data["project_name"],
                t0=t0, t1=t1,
                output_path=tl_path,
            )

            # 7. PDF (optional)
            pdf_path = slide_builder.export_pdf(out_path)
            has_pdf  = pdf_path is not None

            print("DONE\n")
            self._json(200, {
                "ok":             True,
                "pptx":           out_name,
                "pdf":            pdf_path.name if has_pdf else None,
                "has_pdf":        has_pdf,
                "timeline_html":  tl_name,
                "project":        data["project_name"],
                "tasks":          len(data["tasks"]),
                "milestones":     len(data["milestones"]),
                "source":         data["source"],
                "status_preview": texts["status"][:200],
            })

        except Exception as exc:
            tb = traceback.format_exc()
            print(tb)
            self._json(500, {"ok": False, "error": str(exc), "trace": tb})

    # ── File download ─────────────────────────────────────────────────────────
    def _serve_file(self, path: Path):
        if not path.exists():
            self._send_404()
            return
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Content-Disposition", f'attachment; filename="{path.name}"')
        self.end_headers()
        self.wfile.write(data)

    # ── Low-level helpers ─────────────────────────────────────────────────────
    def _send(self, code: int, body: bytes, content_type: str = "text/plain"):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code: int, obj):
        body = json.dumps(obj, default=str).encode()
        self._send(code, body, "application/json")

    def _send_404(self):
        self._send(404, b"Not found")


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("\n" + "="*56)
    print("  BPE Report Builder — Solution 1")
    print("="*56)
    print(f"  Mode   : {'DEMO (xlsx)' if config.DEMO_MODE else 'LIVE (Smartsheet API)'}")
    print(f"  AI     : {'Claude ON'   if config.USE_AI    else 'Rule-based (no API key)'}")
    print(f"  PPTX   : {config.SOURCE_PPTX}")
    print(f"  Engine : Pure Python SVG (no Pillow needed)")
    print(f"\n  Open   → http://localhost:{PORT}\n")

    server = HTTPServer(("0.0.0.0", PORT), Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Server stopped.")
