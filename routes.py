"""
routes.py — FastAPI route definitions
"""

import asyncio
import io
from datetime import date, datetime

from fastapi import APIRouter, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from jinja2 import Environment, FileSystemLoader

from processors import process_current, process_eod, process_range, process_single_date

router = APIRouter()
env    = Environment(loader=FileSystemLoader("templates"))


# ── Helpers ───────────────────────────────────────────────────────────────────
def _validate_upload(file: UploadFile) -> None:
    if not file.filename.endswith((".xlsx", ".xlsm")):
        raise HTTPException(400, "Only .xlsx files are supported.")


def _stream_response(result_bytes: bytes, stats: dict, mode_prefix: str) -> StreamingResponse:
    safe_date  = stats["date"].replace(" ", "_").replace("→", "to").replace("-", "_")
    # HTTP headers are latin-1 only — strip the arrow and any non-latin-1 chars
    header_date = stats["date"].replace("→", "to").encode("latin-1", errors="replace").decode("latin-1")
    filename   = f"stocks_{mode_prefix}_{safe_date}.xlsx"
    return StreamingResponse(
        io.BytesIO(result_bytes),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition":   f"attachment; filename={filename}",
            "X-Stats-Date":          header_date,
            "X-Stats-Fetched":       str(stats["fetched"]),
            "X-Stats-Failed":        str(stats["failed"]),
            "X-Stats-Total":         str(stats["total"]),
            "X-Stats-Columns":       str(stats.get("columns", 1)),
            # Range-only: comma-separated trading days & raw from/to for calendar UI
            "X-Stats-Trading-Days":  ",".join(stats.get("trading_days", [])),
            "X-Stats-From":          stats.get("from_date", ""),
            "X-Stats-To":            stats.get("to_date", ""),
        },
    )


# ── Routes ────────────────────────────────────────────────────────────────────
@router.get("/", response_class=HTMLResponse)
async def index(request: Request):
    template = env.get_template("index.html")
    return HTMLResponse(content=template.render())


@router.post("/upload")
async def upload_eod(file: UploadFile = File(...)):
    """Mode 1 — Append today's EOD prices."""
    _validate_upload(file)
    contents = await file.read()
    if len(contents) > 20 * 1024 * 1024:
        raise HTTPException(400, "File too large (max 20 MB).")
    try:
        result_bytes, stats = await asyncio.to_thread(process_eod, contents)
    except ValueError as e:
        raise HTTPException(422, str(e))
    except Exception as e:
        raise HTTPException(500, f"Processing error: {e}")
    return _stream_response(result_bytes, stats, "eod")


@router.post("/get-current")
async def upload_current(file: UploadFile = File(...)):
    """Mode 2 — Fetch live/current prices."""
    _validate_upload(file)
    contents = await file.read()
    try:
        result_bytes, stats = await asyncio.to_thread(process_current, contents)
    except ValueError as e:
        raise HTTPException(422, str(e))
    except Exception as e:
        raise HTTPException(500, f"Processing error: {e}")
    return _stream_response(result_bytes, stats, "live")


@router.post("/get-date")
async def upload_date(
    file: UploadFile = File(...),
    target_date: str = Form(...),
):
    """Mode 3 — Fetch closing price for a specific date."""
    _validate_upload(file)
    try:
        tdate = datetime.strptime(target_date, "%Y-%m-%d").date()
    except ValueError:
        raise HTTPException(400, "Invalid date format. Expected YYYY-MM-DD.")
    if tdate.weekday() >= 5:
        raise HTTPException(422, f"{tdate.strftime('%d-%b-%Y')} is a weekend. Markets are closed.")
    if tdate > date.today():
        raise HTTPException(422, "Future dates are not supported.")
    contents = await file.read()
    try:
        result_bytes, stats = await asyncio.to_thread(process_single_date, contents, tdate)
    except ValueError as e:
        raise HTTPException(422, str(e))
    except Exception as e:
        raise HTTPException(500, f"Processing error: {e}")
    return _stream_response(result_bytes, stats, "date")


@router.post("/get-range")
async def upload_range(
    file: UploadFile = File(...),
    from_date: str   = Form(...),
    to_date: str     = Form(...),
):
    """Mode 4 — Fetch closing prices for a date range."""
    _validate_upload(file)
    try:
        fdate = datetime.strptime(from_date, "%Y-%m-%d").date()
        tdate = datetime.strptime(to_date,   "%Y-%m-%d").date()
    except ValueError:
        raise HTTPException(400, "Invalid date format. Expected YYYY-MM-DD.")
    if fdate > tdate:
        raise HTTPException(422, "From date must be before To date.")
    if (tdate - fdate).days > 90:
        raise HTTPException(422, "Range too large. Keep it within 90 days.")
    if tdate > date.today():
        raise HTTPException(422, "To date cannot be in the future.")
    contents = await file.read()
    try:
        result_bytes, stats = await asyncio.to_thread(process_range, contents, fdate, tdate)
    except ValueError as e:
        raise HTTPException(422, str(e))
    except Exception as e:
        raise HTTPException(500, f"Processing error: {e}")
    return _stream_response(result_bytes, stats, "range")