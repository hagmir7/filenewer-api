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


def json_to_excel(source, sheet_name: str = "Sheet1") -> bytes:
    """
    Convert JSON (file, raw text, list, or dict) → Excel bytes.

    Args:
        source     : uploaded file object | raw JSON string | list | dict
        sheet_name : name of the Excel sheet (default: Sheet1)

    Returns:
        Raw bytes of the generated .xlsx file
    """
    # ── Parse input ───────────────────────────────
    if isinstance(source, (list, dict)):
        data = source  # already parsed
    elif isinstance(source, str):
        data = json.loads(source)  # raw JSON string
    else:
        raw = source.read().decode("utf-8")  # file object
        data = json.loads(raw)

    # ── Handle all JSON shapes ─────────────────────
    if isinstance(data, dict):
        # { "data": [{...}] }  → unwrap
        for key, val in data.items():
            if isinstance(val, list):
                data = val
                break
        else:
            data = [data]  # single object → wrap in list

    if not isinstance(data, list):
        raise ValueError("JSON must be an array of objects or a wrapped array.")

    # ── Flatten nested objects ─────────────────────
    df = pd.json_normalize(data)
    df.columns = [sanitize_name(c) for c in df.columns]

    # ── Write to Excel in memory ───────────────────
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        # ── Auto-fit column widths ─────────────────
        worksheet = writer.sheets[sheet_name]
        for col_cells in worksheet.columns:
            max_len = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in col_cells
            )
            col_letter = col_cells[0].column_letter
            worksheet.column_dimensions[col_letter].width = min(max_len + 4, 50)

        # ── Style header row ───────────────────────
        from openpyxl.styles import Font, PatternFill, Alignment

        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(fill_type="solid", fgColor="2E75B6")
        header_align = Alignment(horizontal="center", vertical="center")

        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align

    buffer.seek(0)
    return buffer.read()


def json_to_excel_multisheets(sources: dict) -> bytes:
    """
    Convert multiple JSON arrays → multi-sheet Excel file.

    Args:
        sources : { 'SheetName': list_or_dict, ... }

    Returns:
        Raw bytes of the generated .xlsx file
    """
    from openpyxl.styles import Font, PatternFill, Alignment

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet, data in sources.items():

            # ── Handle shapes ──────────────────────
            if isinstance(data, dict):
                for key, val in data.items():
                    if isinstance(val, list):
                        data = val
                        break
                else:
                    data = [data]

            df = pd.json_normalize(data)
            df.columns = [sanitize_name(c) for c in df.columns]
            df.to_excel(
                writer, index=False, sheet_name=sheet[:31]
            )  # Excel max 31 chars

            # ── Auto-fit + style header ────────────
            worksheet = writer.sheets[sheet[:31]]

            for col_cells in worksheet.columns:
                max_len = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in col_cells
                )
                col_letter = col_cells[0].column_letter
                worksheet.column_dimensions[col_letter].width = min(max_len + 4, 50)

            for cell in worksheet[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(fill_type="solid", fgColor="2E75B6")
                cell.alignment = Alignment(horizontal="center", vertical="center")

    buffer.seek(0)
    return buffer.read()


def csv_to_excel(source, sheet_name: str = "Sheet1", separator: str = None) -> bytes:
    """
    Convert CSV (file or raw text string) → Excel bytes.

    Args:
        source     : uploaded file object OR raw CSV string
        sheet_name : name of the Excel sheet (default: Sheet1)
        separator  : column separator (None = auto-detect)

    Returns:
        Raw bytes of the generated .xlsx file
    """
    # ── Read CSV ──────────────────────────────────
    if isinstance(source, str):
        df = pd.read_csv(io.StringIO(source), sep=separator, engine="python")
    else:
        raw = source.read().decode("utf-8")
        df = pd.read_csv(io.StringIO(raw), sep=separator, engine="python")

    # ── Clean column names ─────────────────────────
    df.columns = [sanitize_name(str(c)) for c in df.columns]

    # ── Replace NaN with empty string ─────────────
    df = df.fillna("")

    # ── Write to Excel in memory ───────────────────
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)

        worksheet = writer.sheets[sheet_name]

        # ── Auto-fit column widths ─────────────────
        for col_cells in worksheet.columns:
            max_len = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in col_cells
            )
            col_letter = col_cells[0].column_letter
            worksheet.column_dimensions[col_letter].width = min(max_len + 4, 50)

        # ── Style header row ───────────────────────
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

        header_font = Font(bold=True, color="FFFFFF", size=11)
        header_fill = PatternFill(fill_type="solid", fgColor="2E75B6")
        header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        thin_border = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )

        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = thin_border

        # ── Style data rows (alternating colors) ───
        from openpyxl.styles import PatternFill as PF

        even_fill = PF(fill_type="solid", fgColor="DCE6F1")  # light blue
        odd_fill = PF(fill_type="solid", fgColor="FFFFFF")  # white

        for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
            fill = even_fill if row_idx % 2 == 0 else odd_fill
            for cell in row:
                cell.fill = fill
                cell.border = thin_border
                cell.alignment = Alignment(vertical="center")

        # ── Freeze header row ──────────────────────
        worksheet.freeze_panes = "A2"

        # ── Set row height for header ──────────────
        worksheet.row_dimensions[1].height = 30

    buffer.seek(0)
    return buffer.read()


def csv_to_excel_multisheets(sources: dict, separator: str = None) -> bytes:
    """
    Convert multiple CSV strings → multi-sheet Excel file.

    Args:
        sources   : { 'SheetName': 'csv string...', ... }
        separator : column separator (None = auto-detect)

    Returns:
        Raw bytes of the generated .xlsx file
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet, csv_data in sources.items():

            # ── Read CSV ───────────────────────────
            if isinstance(csv_data, str):
                df = pd.read_csv(io.StringIO(csv_data), sep=separator, engine="python")
            else:
                raw = csv_data.read().decode("utf-8")
                df = pd.read_csv(io.StringIO(raw), sep=separator, engine="python")

            df.columns = [sanitize_name(str(c)) for c in df.columns]
            df = df.fillna("")

            sheet_label = sheet[:31]  # Excel max 31 chars
            df.to_excel(writer, index=False, sheet_name=sheet_label)

            worksheet = writer.sheets[sheet_label]

            # ── Auto-fit columns ───────────────────
            for col_cells in worksheet.columns:
                max_len = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in col_cells
                )
                col_letter = col_cells[0].column_letter
                worksheet.column_dimensions[col_letter].width = min(max_len + 4, 50)

            # ── Style header ───────────────────────
            thin = Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            )
            for cell in worksheet[1]:
                cell.font = Font(bold=True, color="FFFFFF", size=11)
                cell.fill = PatternFill(fill_type="solid", fgColor="2E75B6")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = thin

            # ── Alternating row colors ─────────────
            for row_idx, row in enumerate(worksheet.iter_rows(min_row=2), start=2):
                fill_color = "DCE6F1" if row_idx % 2 == 0 else "FFFFFF"
                for cell in row:
                    cell.fill = PatternFill(fill_type="solid", fgColor=fill_color)
                    cell.border = thin
                    cell.alignment = Alignment(vertical="center")

            worksheet.freeze_panes = "A2"
            worksheet.row_dimensions[1].height = 30

    buffer.seek(0)
    return buffer.read()


def pdf_to_excel(source, password: str = None) -> bytes:
    """
    Convert PDF (file or bytes) → Excel bytes.

    Strategy:
        1. Extract tables  → each table gets its own sheet
        2. Extract text    → one 'Text' sheet with all plain text
        3. Style headers   → bold white on blue, alternating rows, auto-width

    Args:
        source   : uploaded file object OR raw bytes
        password : PDF password if encrypted (default: None)

    Returns:
        Raw bytes of the generated .xlsx file
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl import Workbook

    # ── Read PDF bytes ────────────────────────────
    if hasattr(source, "read"):
        pdf_bytes = source.read()
    else:
        pdf_bytes = source

    # ── Styles ────────────────────────────────────
    def make_header_style():
        return {
            "font": Font(bold=True, color="FFFFFF", size=11),
            "fill": PatternFill(fill_type="solid", fgColor="2E75B6"),
            "align": Alignment(horizontal="center", vertical="center", wrap_text=True),
            "border": Border(
                left=Side(style="thin"),
                right=Side(style="thin"),
                top=Side(style="thin"),
                bottom=Side(style="thin"),
            ),
        }

    def apply_header(ws):
        style = make_header_style()
        for cell in ws[1]:
            cell.font = style["font"]
            cell.fill = style["fill"]
            cell.alignment = style["align"]
            cell.border = style["border"]
        ws.row_dimensions[1].height = 25

    def apply_rows(ws):
        thin = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
            color = "DCE6F1" if row_idx % 2 == 0 else "FFFFFF"
            for cell in row:
                cell.fill = PatternFill(fill_type="solid", fgColor=color)
                cell.border = thin
                cell.alignment = Alignment(vertical="center")

    def auto_fit(ws):
        for col_cells in ws.columns:
            max_len = max(
                len(str(cell.value)) if cell.value is not None else 0
                for cell in col_cells
            )
            ws.column_dimensions[col_cells[0].column_letter].width = min(
                max_len + 4, 60
            )

    # ── Parse PDF ─────────────────────────────────
    wb = Workbook()
    wb.remove(wb.active)  # remove default empty sheet

    table_count = 0
    all_text = []

    open_kwargs = {"stream": io.BytesIO(pdf_bytes)}
    pdf_stream = io.BytesIO(pdf_bytes)
    if password:
        open_kwargs["password"] = password

    with pdfplumber.open(pdf_stream, password=password) as pdf:
        total_pages = len(pdf.pages)

        for page_num, page in enumerate(pdf.pages, start=1):

            # ── Extract tables ─────────────────────
            tables = page.extract_tables() or []
            for table in tables:
                if not table:
                    continue

                # clean None cells
                cleaned = [
                    [str(cell).strip() if cell is not None else "" for cell in row]
                    for row in table
                ]

                table_count += 1
                sheet_label = f"Table_{table_count}_P{page_num}"[:31]
                ws = wb.create_sheet(title=sheet_label)

                for row in cleaned:
                    ws.append(row)

                apply_header(ws)
                apply_rows(ws)
                auto_fit(ws)
                ws.freeze_panes = "A2"

            # ── Extract plain text ─────────────────
            # Exclude table regions to avoid duplication
            cropped = page
            for tbl in page.find_tables():
                try:
                    cropped = cropped.outside_bbox(tbl.bbox)
                except Exception:
                    pass

            text = cropped.extract_text(x_tolerance=3, y_tolerance=3) or ""
            if text.strip():
                all_text.append(f"--- Page {page_num} ---")
                all_text.append(text.strip())
                all_text.append("")

    # ── Text sheet ────────────────────────────────
    if all_text:
        ws_text = wb.create_sheet(title="Text")
        ws_text.column_dimensions["A"].width = 100

        # Header
        ws_text.append(["Extracted Text"])
        apply_header(ws_text)

        for line in all_text:
            ws_text.append([line])

        # Light style for text rows
        thin = Border(
            left=Side(style="thin"),
            right=Side(style="thin"),
            top=Side(style="thin"),
            bottom=Side(style="thin"),
        )
        for row in ws_text.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.border = thin

    # ── Summary sheet (first sheet) ───────────────
    ws_summary = wb.create_sheet(title="Summary", index=0)
    ws_summary.column_dimensions["A"].width = 25
    ws_summary.column_dimensions["B"].width = 40

    summary_data = [
        ["Property", "Value"],
        ["Total Pages", total_pages],
        ["Tables Found", table_count],
        ["Has Text", "Yes" if all_text else "No"],
        ["Sheets", wb.sheetnames.__len__()],
    ]

    for row in summary_data:
        ws_summary.append(row)

    apply_header(ws_summary)
    apply_rows(ws_summary)
    auto_fit(ws_summary)

    # ── Serialize ─────────────────────────────────
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer.read()


def pdf_to_jpg(source, dpi: int = 200, quality: int = 85, password: str = None) -> list[dict]:
    """
    Convert PDF → JPG using pymupdf (no poppler / no system deps needed).

    Args:
        source   : uploaded file object OR raw bytes
        dpi      : image resolution (default: 200)
        quality  : JPG quality 1-95 (default: 85)
        password : PDF password if encrypted (default: None)

    Returns:
        list of {
            'page'    : int,
            'bytes'   : bytes,
            'width'   : int,
            'height'  : int,
            'filename': str,
        }
    """
    import fitz          # pymupdf
    from PIL import Image

    # ── Read PDF bytes ────────────────────────────
    if hasattr(source, 'read'):
        pdf_bytes = source.read()
    else:
        pdf_bytes = source

    # ── Open PDF ──────────────────────────────────
    doc = fitz.open(stream=pdf_bytes, filetype='pdf')

    if password:
        if not doc.authenticate(password):
            raise ValueError('Invalid PDF password.')

    # ── Convert each page → JPG ───────────────────
    zoom   = dpi / 72        # pymupdf base DPI is 72
    matrix = fitz.Matrix(zoom, zoom)

    results = []

    for page_num in range(len(doc)):
        page    = doc[page_num]
        pixmap  = page.get_pixmap(matrix=matrix, alpha=False)

        # Pixmap → PIL Image → JPG bytes
        image = Image.frombytes(
            'RGB',
            [pixmap.width, pixmap.height],
            pixmap.samples,
        )

        buffer = io.BytesIO()
        image.save(
            buffer,
            format  ='JPEG',
            quality =quality,
            optimize=True,
        )
        buffer.seek(0)
        jpg_bytes = buffer.read()

        results.append({
            'page'    : page_num + 1,
            'bytes'   : jpg_bytes,
            'width'   : pixmap.width,
            'height'  : pixmap.height,
            'filename': f'page_{page_num + 1}.jpg',
        })

    doc.close()
    return results


# ── Arabic support ─────────────────────────────────
try:
    from bidi.algorithm import get_display
    import arabic_reshaper

    ARABIC_SUPPORT = True
except ImportError:
    ARABIC_SUPPORT = False


def is_arabic(text: str) -> bool:
    """Check if text contains Arabic characters."""
    return bool(re.search(r"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF]", text))


def process_arabic_text(text: str) -> str:
    """
    Reshape and reorder Arabic text for correct PDF rendering.
    Arabic needs two steps:
        1. Reshape  → connect letters correctly
        2. Bidi     → reverse for RTL display
    """
    if not ARABIC_SUPPORT or not text.strip():
        return text

    try:
        if is_arabic(text):
            reshaped = arabic_reshaper.reshape(text)
            return get_display(reshaped)
        return text
    except Exception:
        return text


def register_arabic_fonts(fonts_dir: str = None) -> dict:
    """
    Register Arabic fonts with reportlab.
    Returns dict of available font names.
    """
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont

    if fonts_dir is None:
        # Default: fonts/ folder next to manage.py
        fonts_dir = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "fonts"
        )

    registered = {}
    font_files = {
        "Amiri": "Amiri-Regular.ttf",
        "Amiri-Bold": "Amiri-Bold.ttf",
        "Amiri-Italic": "Amiri-Italic.ttf",
        "Amiri-BoldItalic": "Amiri-BoldItalic.ttf",
    }

    for font_name, font_file in font_files.items():
        font_path = os.path.join(fonts_dir, font_file)
        if os.path.exists(font_path):
            try:
                pdfmetrics.registerFont(TTFont(font_name, font_path))
                registered[font_name] = font_path
            except Exception:
                pass

    return registered


def word_to_pdf(source, filename: str = "document.docx") -> bytes:
    """
    Convert Word (.docx) → PDF with full Arabic / RTL support.
    """
    from docx import Document as DocxDocument
    from docx.text.paragraph import Paragraph as DocxPara
    from docx.table import Table as DocxTable
    from docx.oxml.ns import qn
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_JUSTIFY
    from reportlab.platypus import (
        SimpleDocTemplate,
        Paragraph,
        Spacer,
        Table,
        TableStyle,
        PageBreak,
    )

    # ── Read DOCX ─────────────────────────────────
    if hasattr(source, "read"):
        file_bytes = source.read()
    else:
        file_bytes = source

    docx = DocxDocument(io.BytesIO(file_bytes))

    # ── Register Arabic fonts ──────────────────────
    registered_fonts = register_arabic_fonts()
    has_arabic_font = bool(registered_fonts)

    arabic_font = "Amiri" if "Amiri" in registered_fonts else "Helvetica"
    arabic_font_bold = (
        "Amiri-Bold" if "Amiri-Bold" in registered_fonts else "Helvetica-Bold"
    )
    latin_font = "Helvetica"
    latin_font_bold = "Helvetica-Bold"

    # ── Helper: pick font based on content ─────────
    def pick_font(text: str, bold: bool = False) -> str:
        if is_arabic(text):
            return arabic_font_bold if bold else arabic_font
        return latin_font_bold if bold else latin_font

    # ── Helper: process text (Arabic or Latin) ─────
    def prepare_text(text: str) -> str:
        if is_arabic(text):
            return process_arabic_text(text)
        return text

    # ── Helper: escape XML special chars ───────────
    def escape(text: str) -> str:
        return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    # ── Custom styles ──────────────────────────────
    def make_style(name, font, font_bold, size, color, before, after, align=TA_LEFT):
        return ParagraphStyle(
            name,
            fontName=font_bold,
            fontSize=size,
            textColor=color,
            spaceBefore=before,
            spaceAfter=after,
            leading=size * 1.4,
            alignment=align,
            wordWrap="RTL" if font == arabic_font else "LTR",
        )

    # ── Process paragraph runs → markup ────────────
    def runs_to_markup(para) -> tuple[str, bool]:
        """
        Returns (markup_string, has_arabic).
        Uses correct font per run.
        """
        parts = []
        has_arabic = False

        for run in para.runs:
            raw = run.text
            if not raw:
                continue

            arabic = is_arabic(raw)
            if arabic:
                has_arabic = True
                raw = process_arabic_text(raw)

            text = escape(raw)
            font = pick_font(run.text, bold=run.bold)

            # Wrap in font tag
            text = f'<font name="{font}">{text}</font>'

            if run.bold:
                text = f"<b>{text}</b>"
            if run.italic:
                text = f"<i>{text}</i>"
            if run.underline:
                text = f"<u>{text}</u>"

            parts.append(text)

        return "".join(parts), has_arabic

    # ── Pick paragraph style ───────────────────────
    def pick_style(para, has_arabic: bool) -> ParagraphStyle:
        name = (para.style.name or "Normal").lower()
        align = TA_RIGHT if has_arabic else TA_JUSTIFY
        font = arabic_font if has_arabic else latin_font
        bold = arabic_font_bold if has_arabic else latin_font_bold
        wrap = "RTL" if has_arabic else "LTR"

        base = {
            "fontName": font,
            "leading": 20,
            "wordWrap": wrap,
            "alignment": align,
        }

        if "heading 1" in name:
            return ParagraphStyle(
                "H1",
                **base,
                fontSize=22,
                textColor=colors.HexColor("#2E75B6"),
                fontName=bold,
                spaceBefore=18,
                spaceAfter=14,
            )

        if "heading 2" in name:
            return ParagraphStyle(
                "H2",
                **base,
                fontSize=16,
                textColor=colors.HexColor("#2E75B6"),
                fontName=bold,
                spaceBefore=14,
                spaceAfter=10,
            )

        if "heading 3" in name:
            return ParagraphStyle(
                "H3",
                **base,
                fontSize=13,
                textColor=colors.HexColor("#1F4E79"),
                fontName=bold,
                spaceBefore=10,
                spaceAfter=8,
            )

        if "list bullet" in name:
            return ParagraphStyle(
                "Bullet", **base, fontSize=11, leftIndent=20, spaceAfter=4
            )

        if "list number" in name:
            return ParagraphStyle(
                "Number", **base, fontSize=11, leftIndent=20, spaceAfter=4
            )

        return ParagraphStyle("Body", **base, fontSize=11, spaceAfter=6)

    # ── Process table ──────────────────────────────
    def process_table(tbl) -> Table:
        data = []
        col_cnt = max(len(row.cells) for row in tbl.rows)
        col_w = (A4[0] - 2 * inch) / col_cnt

        for r_idx, row in enumerate(tbl.rows):
            row_data = []
            for cell in row.cells:
                cell_text = "\n".join(p.text for p in cell.paragraphs)
                arabic = is_arabic(cell_text)

                if arabic:
                    cell_text = process_arabic_text(cell_text)

                font = (
                    arabic_font_bold
                    if (r_idx == 0 and arabic)
                    else (
                        arabic_font
                        if arabic
                        else (latin_font_bold if r_idx == 0 else latin_font)
                    )
                )

                align = TA_RIGHT if arabic else TA_LEFT
                wrap = "RTL" if arabic else "LTR"

                style = ParagraphStyle(
                    f"cell_{r_idx}",
                    fontName=font,
                    fontSize=10,
                    textColor=colors.white if r_idx == 0 else colors.black,
                    alignment=align,
                    wordWrap=wrap,
                    leading=14,
                )
                row_data.append(Paragraph(escape(cell_text), style))
            data.append(row_data)

        pdf_tbl = Table(data, colWidths=[col_w] * col_cnt, repeatRows=1)
        pdf_tbl.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2E75B6")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                    (
                        "ROWBACKGROUNDS",
                        (0, 1),
                        (-1, -1),
                        [colors.HexColor("#DCE6F1"), colors.white],
                    ),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#AAAAAA")),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("PADDING", (0, 0), (-1, -1), 6),
                ]
            )
        )
        return pdf_tbl

    # ── Build story ────────────────────────────────
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=inch,
        leftMargin=inch,
        topMargin=inch,
        bottomMargin=inch,
        title=filename.replace(".docx", "").replace(".doc", ""),
    )

    story = []

    for element in docx.element.body:
        tag = element.tag.split("}")[-1]

        # ── Paragraph ─────────────────────────────
        if tag == "p":
            para = DocxPara(element, docx)

            # Page break
            if any(
                br.get(qn("w:type")) == "page"
                for br in element.findall(".//" + qn("w:br"))
            ):
                story.append(PageBreak())
                continue

            # Empty
            if not para.text.strip():
                story.append(Spacer(1, 8))
                continue

            markup, has_arabic = runs_to_markup(para)
            style = pick_style(para, has_arabic)

            # Bullet
            name = (para.style.name or "").lower()
            if "list bullet" in name:
                prefix = "• " if not has_arabic else " •"
                try:
                    story.append(Paragraph(prefix + markup, style))
                except Exception:
                    story.append(
                        Paragraph(prefix + escape(prepare_text(para.text)), style)
                    )
                continue

            # Numbered
            if "list number" in name:
                try:
                    story.append(Paragraph("1. " + markup, style))
                except Exception:
                    story.append(
                        Paragraph("1. " + escape(prepare_text(para.text)), style)
                    )
                continue

            try:
                story.append(Paragraph(markup, style))
            except Exception:
                story.append(Paragraph(escape(prepare_text(para.text)), style))

        # ── Table ──────────────────────────────────
        elif tag == "tbl":
            try:
                tbl = DocxTable(element, docx)
                story.append(Spacer(1, 8))
                story.append(process_table(tbl))
                story.append(Spacer(1, 12))
            except Exception as e:
                story.append(
                    Paragraph(
                        f"[Table error: {e}]",
                        ParagraphStyle("err", fontName="Helvetica", fontSize=10),
                    )
                )

        elif tag == "sectPr":
            pass

    doc.build(story)
    buffer.seek(0)
    return buffer.read()
