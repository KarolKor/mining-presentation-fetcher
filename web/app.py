from __future__ import annotations

import json
import os
import secrets
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

from flask import Flask, abort, render_template, request, send_file, url_for


BASE_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = BASE_DIR.parent
SCRIPT_PATH = PROJECT_ROOT / "scripts" / "fetch_latest_presentations.py"
RUNS_DIR = BASE_DIR / "runs"

RUNS_DIR.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024
EXCHANGE_OPTIONS = [
    "NYSE",
    "NASDAQ",
    "AMEX",
    "TSX",
    "TSXV",
    "CSE",
    "ASX",
    "LSE",
]


def _now_id() -> str:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    return f"{stamp}-{secrets.token_hex(4)}"


def _safe_rel(base: Path, target: Path) -> str | None:
    try:
        rel = target.resolve().relative_to(base.resolve())
    except ValueError:
        return None
    return rel.as_posix()


def _read_json(path: Path, fallback: object) -> object:
    if not path.exists():
        return fallback
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return fallback


def _build_download_url(job_id: str, output_dir: Path, path_value: str | None) -> str | None:
    if not path_value:
        return None
    candidate = Path(path_value)
    rel = _safe_rel(output_dir, candidate)
    if rel is None:
        return None
    return url_for("download", job_id=job_id, file_path=rel)


def _collect_result(job_id: str) -> dict:
    job_dir = RUNS_DIR / job_id
    output_dir = job_dir / "output"
    manifest_path = output_dir / "manifest.json"
    run_meta_path = job_dir / "run.json"

    records = _read_json(manifest_path, [])
    if not isinstance(records, list):
        records = []

    for item in records:
        if not isinstance(item, dict):
            continue
        item["local_file_url"] = _build_download_url(
            job_id=job_id,
            output_dir=output_dir,
            path_value=item.get("local_file"),
        )
        item["text_file_url"] = _build_download_url(
            job_id=job_id,
            output_dir=output_dir,
            path_value=item.get("text_file"),
        )

    run_meta = _read_json(run_meta_path, {})
    if not isinstance(run_meta, dict):
        run_meta = {}

    return {
        "job_id": job_id,
        "output_dir": output_dir,
        "records": records,
        "manifest_exists": manifest_path.exists(),
        "run_meta": run_meta,
    }


@app.get("/")
def index():
    job_id = request.args.get("job_id", "").strip()
    result = _collect_result(job_id) if job_id else None
    return render_template(
        "index.html",
        result=result,
        exchange_options=EXCHANGE_OPTIONS,
        form_values={"ticker": "", "exchange": "TSX"},
    )


@app.post("/run")
def run_fetch():
    ticker = (request.form.get("ticker") or "").strip().upper()
    exchange = (request.form.get("exchange") or "").strip().upper()
    if exchange not in EXCHANGE_OPTIONS:
        exchange = "TSX"

    specs: list[str] = []
    if ticker:
        specs.append(f"{ticker}|{exchange}")

    include_sec = request.form.get("include_sec") == "on"
    playwright_fallback = request.form.get("playwright_fallback") == "on"
    disable_search_fallback = request.form.get("disable_search_fallback") == "on"
    dry_run = request.form.get("dry_run") == "on"

    latest_count_raw = (request.form.get("latest_count") or "1").strip()
    timeout_raw = (request.form.get("timeout") or "20").strip()

    try:
        latest_count = max(1, min(10, int(latest_count_raw)))
    except ValueError:
        latest_count = 1

    try:
        timeout = max(5, min(120, int(timeout_raw)))
    except ValueError:
        timeout = 20

    upload = request.files.get("companies_csv")
    has_upload = bool(upload and upload.filename and upload.filename.strip())
    if not specs and not has_upload:
        return render_template(
            "index.html",
            form_error="Provide a ticker + exchange or upload a CSV file.",
            result=None,
            exchange_options=EXCHANGE_OPTIONS,
            form_values={"ticker": ticker, "exchange": exchange},
        )

    job_id = _now_id()
    job_dir = RUNS_DIR / job_id
    output_dir = job_dir / "output"
    job_dir.mkdir(parents=True, exist_ok=True)

    csv_path = None
    if has_upload and upload is not None:
        csv_path = job_dir / "companies.csv"
        upload.save(csv_path)

    cmd = [
        sys.executable,
        str(SCRIPT_PATH),
        "--output-dir",
        str(output_dir),
        "--latest-count",
        str(latest_count),
        "--timeout",
        str(timeout),
        "--verbose",
    ]

    if include_sec:
        cmd.append("--include-sec")
    if playwright_fallback:
        cmd.append("--playwright-fallback")
    if disable_search_fallback:
        cmd.append("--disable-search-fallback")
    if dry_run:
        cmd.append("--dry-run")
    if csv_path is not None:
        cmd.extend(["--companies-csv", str(csv_path)])
    for spec in specs:
        cmd.extend(["--company", spec])

    completed = subprocess.run(
        cmd,
        cwd=PROJECT_ROOT,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60 * 30,
    )

    run_meta = {
        "command": cmd,
        "returncode": completed.returncode,
        "stdout": completed.stdout,
        "stderr": completed.stderr,
        "submitted_at_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "options": {
            "ticker": ticker,
            "exchange": exchange,
            "latest_count": latest_count,
            "timeout": timeout,
            "include_sec": include_sec,
            "playwright_fallback": playwright_fallback,
            "disable_search_fallback": disable_search_fallback,
            "dry_run": dry_run,
        },
    }
    (job_dir / "run.json").write_text(json.dumps(run_meta, indent=2), encoding="utf-8")

    result = _collect_result(job_id)
    if completed.returncode != 0:
        return render_template(
            "index.html",
            form_error="Fetcher exited with an error. Check logs below.",
            result=result,
            exchange_options=EXCHANGE_OPTIONS,
            form_values={"ticker": ticker, "exchange": exchange},
        )
    return render_template(
        "index.html",
        result=result,
        exchange_options=EXCHANGE_OPTIONS,
        form_values={"ticker": ticker, "exchange": exchange},
    )


@app.get("/files/<job_id>/<path:file_path>")
def download(job_id: str, file_path: str):
    job_dir = RUNS_DIR / job_id
    output_dir = job_dir / "output"
    if not output_dir.exists():
        abort(404)

    target = (output_dir / file_path).resolve()
    try:
        target.relative_to(output_dir.resolve())
    except ValueError:
        abort(404)

    if not target.exists() or not target.is_file():
        abort(404)
    return send_file(target)


if __name__ == "__main__":
    debug_enabled = os.getenv("APP_DEBUG", "").strip() in {"1", "true", "True"}
    app.run(host="127.0.0.1", port=5050, debug=debug_enabled, use_reloader=debug_enabled)
