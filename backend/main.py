from __future__ import annotations

import io
import math
import os
import re
import tempfile
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pydantic import BaseModel, Field
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Table, TableStyle

app = FastAPI(title="Dynamic Excel Builder API", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

ALLOWED_EXTENSIONS = {".xlsx", ".xls"}
MAX_UPLOAD_BYTES = 100 * 1024 * 1024


def safe_filename(name: str, default: str = "export") -> str:
    name = (name or default).strip()
    name = re.sub(r"[^\w\-.\u0080-\uffff ]+", "_", name, flags=re.UNICODE)
    return name or default


def unique_columns(columns: List[Any]) -> List[str]:
    seen: Dict[str, int] = {}
    result: List[str] = []
    for i, col in enumerate(columns):
        base = str(col).strip() if col is not None and not (isinstance(col, float) and math.isnan(col)) else ""
        if not base or base.lower().startswith("unnamed:"):
            base = f"Column_{i + 1}"
        count = seen.get(base, 0)
        result.append(base if count == 0 else f"{base}_{count}")
        seen[base] = count + 1
    return result


def json_safe(value: Any) -> Any:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return value


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df = df.dropna(how="all")
    df = df.dropna(axis=1, how="all")
    df.columns = unique_columns(list(df.columns))
    return df


def read_excel_upload(file: UploadFile, content: bytes, header_row: int, formulas: str = "values") -> pd.DataFrame:
    suffix = Path(file.filename or "").suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        raise HTTPException(status_code=400, detail="Unsupported file type. Please upload a valid Excel file.")
    if len(content) == 0:
        raise HTTPException(status_code=400, detail="No data found.")
    if len(content) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="File too large. Maximum supported upload is 100 MB.")
    try:
        # pandas/openpyxl imports cached formula values where available. Macro/script execution never occurs.
        df = pd.read_excel(io.BytesIO(content), header=header_row, dtype=object)
    except ValueError as exc:
        msg = str(exc).lower()
        if "password" in msg or "encrypted" in msg:
            raise HTTPException(status_code=400, detail="Password-protected files are not supported.")
        raise HTTPException(status_code=400, detail="Unable to read this Excel file. It may be corrupted or unsupported.")
    except Exception as exc:
        msg = str(exc).lower()
        if "password" in msg or "encrypted" in msg or "file is not a zip" in msg:
            raise HTTPException(status_code=400, detail="Password-protected or corrupted files are not supported.")
        raise HTTPException(status_code=400, detail=f"Unable to process Excel file: {exc}")
    df = normalize_dataframe(df)
    if df.empty or len(df.columns) == 0:
        raise HTTPException(status_code=400, detail="No data found.")
    return df


@app.get("/health")
def health() -> Dict[str, str]:
    return {"status": "ok"}


@app.post("/api/parse-excel")
async def parse_excel(
    file: UploadFile = File(...),
    header_row: int = Form(0),
    formulas: str = Form("values"),
) -> Dict[str, Any]:
    content = await file.read()
    df = read_excel_upload(file, content, max(header_row, 0), formulas)
    rows = [{col: json_safe(row[col]) for col in df.columns} for _, row in df.iterrows()]
    columns = [
        {
            "id": re.sub(r"\W+", "_", str(col).strip().lower(), flags=re.UNICODE).strip("_") or f"column_{idx+1}",
            "originalName": col,
            "displayName": col,
            "dataType": str(df[col].dropna().map(type).iloc[0].__name__) if not df[col].dropna().empty else "text",
        }
        for idx, col in enumerate(df.columns)
    ]
    return {
        "fileName": file.filename,
        "sheetName": "Sheet1",
        "headerRowIndex": header_row,
        "columns": columns,
        "rows": rows,
        "rowCount": len(rows),
        "columnCount": len(columns),
    }


@app.post("/api/import-working-sheet")
async def import_working_sheet(file: UploadFile = File(...), header_row: int = Form(0)) -> Dict[str, Any]:
    content = await file.read()
    df = read_excel_upload(file, content, max(header_row, 0), "values")
    rows = [[json_safe(row[col]) for col in df.columns] for _, row in df.iterrows()]
    return {"columns": [str(c) for c in df.columns], "rows": rows, "rowCount": len(rows)}


class ExportColumn(BaseModel):
    id: str
    displayName: str
    sourceColumn: Optional[str] = None
    isCustom: bool = False


class ExportPayload(BaseModel):
    filename: str = "export"
    columns: List[ExportColumn]
    rows: List[Dict[str, Any]] = Field(default_factory=list)


def row_value(row: Dict[str, Any], col_id: str) -> Any:
    val = row.get(col_id, "")
    if val is None:
        return ""
    if isinstance(val, (dict, list)):
        return str(val)
    return val


@app.post("/api/export/excel")
def export_excel(payload: ExportPayload) -> Response:
    wb = Workbook()
    ws = wb.active
    ws.title = "Working Sheet"
    headers = [c.displayName for c in payload.columns]
    ws.append(headers)

    for cell in ws[1]:
        cell.alignment = Alignment(wrap_text=True, vertical="center")

    for item in payload.rows:
        ws.append([row_value(item, c.id) for c in payload.columns])

    for col_idx, col in enumerate(payload.columns, start=1):
        letter = get_column_letter(col_idx)
        max_len = len(col.displayName)
        for row_idx in range(2, ws.max_row + 1):
            value = ws.cell(row=row_idx, column=col_idx).value
            max_len = max(max_len, len(str(value)) if value is not None else 0)
            ws.cell(row=row_idx, column=col_idx).alignment = Alignment(wrap_text=True, vertical="top")
        ws.column_dimensions[letter].width = min(max(max_len + 2, 12), 60)
    ws.freeze_panes = "A2"

    stream = io.BytesIO()
    wb.save(stream)
    name = safe_filename(payload.filename)
    if not name.lower().endswith(".xlsx"):
        name += ".xlsx"
    return Response(
        stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{name}"'},
    )


def register_unicode_font() -> str:
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/arial.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("AppUnicode", path))
                return "AppUnicode"
            except Exception:
                continue
    return "Helvetica"


@app.post("/api/export/pdf")
def export_pdf(payload: ExportPayload) -> Response:
    font_name = register_unicode_font()
    stream = io.BytesIO()
    doc = SimpleDocTemplate(
        stream,
        pagesize=landscape(A4),
        leftMargin=8 * mm,
        rightMargin=8 * mm,
        topMargin=8 * mm,
        bottomMargin=8 * mm,
    )
    styles = getSampleStyleSheet()
    cell_style = ParagraphStyle(
        "cell",
        parent=styles["BodyText"],
        fontName=font_name,
        fontSize=7,
        leading=9,
        alignment=TA_LEFT,
        wordWrap="CJK",
    )
    header_style = ParagraphStyle(
        "header",
        parent=cell_style,
        fontSize=7,
        leading=9,
    )

    def para(value: Any, style: ParagraphStyle) -> Paragraph:
        text = "" if value is None else str(value)
        text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        return Paragraph(text, style)

    data: List[List[Any]] = [[para(c.displayName, header_style) for c in payload.columns]]
    for row in payload.rows:
        data.append([para(row_value(row, c.id), cell_style) for c in payload.columns])

    page_width = landscape(A4)[0] - 16 * mm
    ncols = max(len(payload.columns), 1)
    col_width = page_width / ncols
    widths = [col_width] * ncols

    table = Table(data, colWidths=widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 3),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    doc.build([table])
    name = safe_filename(payload.filename)
    if not name.lower().endswith(".pdf"):
        name += ".pdf"
    return Response(
        stream.getvalue(),
        media_type="application/pdf",
        headers={"Content-Disposition": f'attachment; filename="{name}"'},
    )
