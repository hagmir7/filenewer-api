"""
PDF → Word conversion service.

Strategy:
  1. pdfplumber  – extract text with layout + tables
  2. pypdf       – fallback for metadata / basic text
  3. python-docx – build the .docx output
"""

import io
import os
import re
import logging
from pathlib import Path
import pandas as pd
import json

import pdfplumber
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────
# PDF TO WORD
# ─────────────────────────────────────────────


def _looks_like_heading(text: str) -> bool:
    """Heuristic: short, no period at end, possible ALL CAPS."""
    text = text.strip()
    if not text:
        return False
    if len(text) > 120:
        return False
    if text.endswith("."):
        return False
    if text.isupper() and len(text) > 3:
        return True
    if re.match(r"^(\d+[\.\)]\s+|\d+\.\d+\s+)", text):  # "1. Title" / "1.2 Section"
        return True
    return False


def _add_table_to_doc(doc: Document, table_data: list[list]) -> None:
    """Add a pdfplumber table into the Word document."""
    if not table_data:
        return

    rows = len(table_data)
    cols = max(len(row) for row in table_data)
    if rows == 0 or cols == 0:
        return

    word_table = doc.add_table(rows=rows, cols=cols)
    word_table.style = "Table Grid"

    for r_idx, row in enumerate(table_data):
        for c_idx, cell_text in enumerate(row):
            if c_idx >= cols:
                break
            cell = word_table.cell(r_idx, c_idx)
            cell.text = str(cell_text) if cell_text is not None else ""
            # Bold first row (header)
            if r_idx == 0:
                for run in cell.paragraphs[0].runs:
                    run.bold = True


def convert_pdf_to_docx(pdf_bytes: bytes) -> bytes:
    """
    Convert PDF bytes → DOCX bytes.

    Returns the raw bytes of the generated .docx file.
    """
    doc = Document()

    # ── Default style tweaks ──────────────────
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(11)

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        total_pages = len(pdf.pages)
        logger.info("Converting PDF: %d page(s)", total_pages)

        for page_num, page in enumerate(pdf.pages, start=1):
            # ── Extract tables first ─────────────
            tables = page.extract_tables() or []
            table_bboxes = []

            for table_data in tables:
                if table_data:
                    _add_table_to_doc(doc, table_data)
                    doc.add_paragraph()  # breathing room after table

            # ── Extract text (excluding table regions) ──
            # Crop away table areas to avoid duplicating content
            cropped_page = page
            for table in page.find_tables():
                try:
                    cropped_page = cropped_page.outside_bbox(table.bbox)
                except Exception:
                    pass  # some versions don't support outside_bbox; fall through

            raw_text = cropped_page.extract_text(x_tolerance=3, y_tolerance=3) or ""

            if not raw_text.strip() and not tables:
                # Blank page – add a soft separator
                doc.add_paragraph()
                continue

            for line in raw_text.splitlines():
                line = line.strip()
                if not line:
                    doc.add_paragraph()
                    continue

                if _looks_like_heading(line):
                    heading_level = 1 if line.isupper() else 2
                    doc.add_heading(line, level=heading_level)
                else:
                    para = doc.add_paragraph(line)
                    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

            # ── Page break between pages (except last) ──
            if page_num < total_pages:
                doc.add_page_break()

    # ── Serialize to bytes ────────────────────
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()


# ─────────────────────────────────────────────
# CSV TO SQL
# ─────────────────────────────────────────────

def sanitize_name(name: str) -> str:
    name = name.strip().lower()
    name = re.sub(r"[^a-z0-9_]", "_", name)
    if name[0].isdigit():
        name = "col_" + name
    return name


def infer_sql_type(dtype) -> str:
    dtype = str(dtype)
    if "int" in dtype:
        return "INTEGER"
    if "float" in dtype:
        return "REAL"
    if "bool" in dtype:
        return "BOOLEAN"
    if "datetime" in dtype:
        return "TIMESTAMP"
    return "TEXT"


def csv_to_sql(source, table_name: str, dialect: str = 'sqlite', separator: str = None) -> str:
    """
    Convert CSV (file or raw text string) → SQL statements.

    Args:
        source     : uploaded file object OR raw CSV string
        table_name : name for the SQL table
        dialect    : 'sqlite' | 'postgresql' | 'mysql'

    Returns:
        SQL string (CREATE TABLE + INSERT statements)
    """
    # ── Read CSV ─────────────────────────────────
    if isinstance(source, str):
        df = pd.read_csv(io.StringIO(source), sep=separator, engine='python')
    else:
        # ── File: read bytes first, then decode ──
        raw = source.read().decode('utf-8')          # ← read raw bytes
        df  = pd.read_csv(io.StringIO(raw), sep=separator, engine='python') 

    df.columns = [sanitize_name(c) for c in df.columns]
    table_name = sanitize_name(table_name)

    # ── Dialect config ────────────────────────────
    if dialect == "postgresql":
        pk = "id SERIAL PRIMARY KEY"
        str_quote = "E'{}'"
    elif dialect == "mysql":
        pk = "id INT AUTO_INCREMENT PRIMARY KEY"
        str_quote = "'{}'"
    else:  # sqlite (default)
        pk = "id INTEGER PRIMARY KEY AUTOINCREMENT"
        str_quote = "'{}'"

    # ── CREATE TABLE ──────────────────────────────
    col_defs = ",\n  ".join(
        f'"{col}" {infer_sql_type(dtype)}' for col, dtype in df.dtypes.items()
    )
    sql_lines = [
        f"-- Generated SQL ({dialect}) for table: {table_name}",
        f"-- Rows: {len(df)}  |  Columns: {len(df.columns)}\n",
        f'CREATE TABLE IF NOT EXISTS "{table_name}" (',
        f"  {pk},",
        f"  {col_defs}",
        ");\n",
    ]

    # ── INSERT statements ─────────────────────────
    col_str = ", ".join(f'"{c}"' for c in df.columns)

    for _, row in df.iterrows():
        values = []
        for val in row:
            if pd.isna(val):
                values.append("NULL")
            elif isinstance(val, bool):
                values.append("TRUE" if val else "FALSE")
            elif isinstance(val, (int, float)):
                values.append(str(val))
            else:
                escaped = str(val).replace("'", "''")
                values.append(f"'{escaped}'")

        val_str = ", ".join(values)
        sql_lines.append(f'INSERT INTO "{table_name}" ({col_str}) VALUES ({val_str});')

    return "\n".join(sql_lines)


# ─────────────────────────────────────────────
# CSV TO JSON
# ─────────────────────────────────────────────
def csv_to_json(source, separator: str = None, orient: str = "records") -> list | dict:
    """
    Convert CSV (file or raw text string) → JSON.

    Args:
        source    : uploaded file object OR raw CSV string
        separator : column separator (None = auto-detect)
        orient    : 'records'  → [ {col: val}, ... ]         (default)
                    'columns'  → { col: {index: val}, ... }
                    'values'   → [ [val, val], ... ]
                    'index'    → { index: {col: val}, ... }

    Returns:
        parsed Python object (list or dict) ready for JSON response
    """
    # ── Read CSV ──────────────────────────────────
    if isinstance(source, str):
        df = pd.read_csv(io.StringIO(source), sep=separator, engine="python")
    else:
        raw = source.read().decode("utf-8")
        df = pd.read_csv(io.StringIO(raw), sep=separator, engine="python")

    # ── Clean column names ─────────────────────────
    df.columns = [sanitize_name(c) for c in df.columns]

    # ── Replace NaN with None (JSON null) ─────────
    df = df.where(pd.notnull(df), None)

    # ── Convert to JSON-serializable object ────────
    return df.to_dict(orient=orient)


def json_to_csv(source, separator: str = ",") -> str:
    """
    Convert JSON (file, raw text, or dict/list) → CSV string.

    Args:
        source    : uploaded file object | raw JSON string | list | dict
        separator : column separator (default: ',')

    Returns:
        CSV string
    """
    # ── Parse input ───────────────────────────────
    if isinstance(source, (list, dict)):
        data = source  # already parsed
    elif isinstance(source, str):
        data = json.loads(source)  # raw JSON string
    else:
        raw = source.read().decode("utf-8")  # file object
        data = json.loads(raw)

    # ── Handle nested/wrapped JSON ─────────────────
    # e.g. { "data": [ {...}, {...} ] }
    if isinstance(data, dict):
        # find the first key whose value is a list
        for key, val in data.items():
            if isinstance(val, list):
                data = val
                break
        else:
            data = [data]  # single object → wrap in list

    if not isinstance(data, list):
        raise ValueError("JSON must be an array of objects or a wrapped array.")

    # ── Convert to DataFrame → CSV ─────────────────
    df = pd.json_normalize(data)  # flattens nested objects
    df.columns = [sanitize_name(c) for c in df.columns]

    return df.to_csv(index=False, sep=separator)


def excel_to_csv(source, sheet_name=None, separator: str = ",") -> dict:
    """
    Convert Excel file (.xlsx / .xls) → CSV string(s).

    Args:
        source     : uploaded file object or file path
        sheet_name : specific sheet name or index (None = all sheets)
        separator  : column separator (default: ',')

    Returns:
        {
            'sheets': {
                'Sheet1': 'csv string...',
                'Sheet2': 'csv string...',
            },
            'total_sheets': 2,
        }
    """
    # ── Read Excel ────────────────────────────────
    if hasattr(source, "read"):
        raw = source.read()
        xls = pd.ExcelFile(io.BytesIO(raw))  # ← BytesIO for binary
    else:
        xls = pd.ExcelFile(source)

    available_sheets = xls.sheet_names

    # ── Decide which sheets to convert ────────────
    if sheet_name is not None:
        # single sheet by name or index
        if isinstance(sheet_name, int):
            if sheet_name >= len(available_sheets):
                raise ValueError(
                    f"Sheet index {sheet_name} out of range. "
                    f"Available: 0-{len(available_sheets)-1}"
                )
            sheet_name = available_sheets[sheet_name]

        if sheet_name not in available_sheets:
            raise ValueError(
                f"Sheet '{sheet_name}' not found. "
                f"Available sheets: {available_sheets}"
            )
        sheets_to_convert = [sheet_name]
    else:
        sheets_to_convert = available_sheets

    # ── Convert each sheet → CSV ──────────────────
    result = {}
    for name in sheets_to_convert:
        df = pd.read_excel(xls, sheet_name=name, engine=None)

        # Clean column names
        df.columns = [sanitize_name(str(c)) for c in df.columns]

        # Replace NaN with empty string for clean CSV
        df = df.fillna("")

        result[name] = df.to_csv(index=False, sep=separator)

    return {
        "sheets": result,
        "total_sheets": len(result),
        "sheet_names": list(result.keys()),
    }
