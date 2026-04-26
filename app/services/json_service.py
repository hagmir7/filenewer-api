"""
JSON service functions.
"""

import io
import json
import logging

import pandas as pd

from .csv_service import sanitize_name

logger = logging.getLogger(__name__)


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


def format_json(
    source,
    indent: int = 4,
    sort_keys: bool = False,
    minify: bool = False,
    ensure_ascii: bool = False,
) -> dict:
    """
    Format / validate / minify JSON from file or raw text.

    Args:
        source       : uploaded file object OR raw JSON string
        indent       : indentation spaces (default: 4)
        sort_keys    : sort keys alphabetically (default: False)
        minify       : minify JSON (overrides indent) (default: False)
        ensure_ascii : escape non-ASCII chars (default: False)

    Returns:
        {
            'formatted'  : str,
            'is_valid'   : bool,
            'key_count'  : int,
            'size_original' : int,
            'size_formatted': int,
            'type'       : 'object' | 'array' | 'other',
            'depth'      : int,
        }
    """
    # ── Read source ───────────────────────────────
    if hasattr(source, "read"):
        raw = source.read()
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8")
    else:
        raw = source

    raw = raw.strip()

    if not raw:
        raise ValueError("Empty input. Provide a valid JSON string or file.")

    # ── Parse JSON ────────────────────────────────
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as e:
        return {
            "formatted": None,
            "is_valid": False,
            "error": str(e),
            "error_line": e.lineno,
            "error_column": e.colno,
            "error_position": e.pos,
            "size_original": len(raw),
        }

    # ── Format ────────────────────────────────────
    if minify:
        formatted = json.dumps(
            parsed,
            separators=(",", ":"),
            ensure_ascii=ensure_ascii,
        )
    else:
        formatted = json.dumps(
            parsed,
            indent=indent,
            sort_keys=sort_keys,
            ensure_ascii=ensure_ascii,
        )

    # ── Analyze structure ─────────────────────────
    def get_depth(obj, level=0):
        if isinstance(obj, dict):
            if not obj:
                return level
            return max(get_depth(v, level + 1) for v in obj.values())
        if isinstance(obj, list):
            if not obj:
                return level
            return max(get_depth(i, level + 1) for i in obj)
        return level

    def count_keys(obj):
        if isinstance(obj, dict):
            return len(obj) + sum(count_keys(v) for v in obj.values())
        if isinstance(obj, list):
            return sum(count_keys(i) for i in obj)
        return 0

    json_type = (
        "object"
        if isinstance(parsed, dict)
        else "array" if isinstance(parsed, list) else "other"
    )

    return {
        "formatted": formatted,
        "is_valid": True,
        "type": json_type,
        "depth": get_depth(parsed),
        "key_count": count_keys(parsed),
        "item_count": len(parsed) if isinstance(parsed, (dict, list)) else 1,
        "size_original": len(raw),
        "size_formatted": len(formatted),
        "size_original_kb": round(len(raw) / 1024, 2),
        "size_formatted_kb": round(len(formatted) / 1024, 2),
        "minified": minify,
        "sorted_keys": sort_keys,
        "indent": indent,
    }


def json_to_yaml(
    source,
    indent: int = 2,
    sort_keys: bool = False,
    allow_unicode: bool = True,
    default_flow: bool = False,
    encoding: str = "utf-8",
) -> dict:
    """
    Convert JSON (file, text, dict, list) → YAML string.

    Args:
        source        : file object | raw JSON string | dict | list
        indent        : YAML indentation spaces        (default: 2)
        sort_keys     : sort keys alphabetically       (default: False)
        allow_unicode : allow unicode characters       (default: True)
        default_flow  : use flow style                 (default: False)
        encoding      : output encoding               (default: utf-8)

    Returns:
        {
            'yaml'         : str,
            'input_type'   : str,
            'key_count'    : int,
            'depth'        : int,
            'size_original': int,
            'size_yaml'    : int,
        }
    """
    import yaml

    # ── Read source ───────────────────────────────
    input_type = "text"

    if hasattr(source, "read"):
        raw = source.read()
        input_type = "file"
        if isinstance(raw, bytes):
            raw = raw.decode(encoding, errors="replace")
        data = json.loads(raw)

    elif isinstance(source, (dict, list)):
        data = source
        input_type = "object"
        raw = json.dumps(source)

    elif isinstance(source, str):
        raw = source.strip()
        data = json.loads(raw)

    elif isinstance(source, bytes):
        raw = source.decode(encoding, errors="replace")
        input_type = "bytes"
        data = json.loads(raw)

    else:
        raise ValueError("source must be a string, file, bytes, dict, or list.")

    if data is None:
        raise ValueError("Input is null/None — nothing to convert.")

    # ── Convert to YAML ───────────────────────────
    yaml_str = yaml.dump(
        data,
        indent=indent,
        sort_keys=sort_keys,
        allow_unicode=allow_unicode,
        default_flow_style=default_flow,
        explicit_start=False,
        explicit_end=False,
    )

    # ── Analyze structure ─────────────────────────
    def get_depth(obj, level=0):
        if isinstance(obj, dict):
            if not obj:
                return level
            return max(get_depth(v, level + 1) for v in obj.values())
        if isinstance(obj, list):
            if not obj:
                return level
            return max(get_depth(i, level + 1) for i in obj)
        return level

    def count_keys(obj):
        if isinstance(obj, dict):
            return len(obj) + sum(count_keys(v) for v in obj.values())
        if isinstance(obj, list):
            return sum(count_keys(i) for i in obj)
        return 0

    return {
        "yaml": yaml_str,
        "input_type": input_type,
        "type": (
            "object"
            if isinstance(data, dict)
            else "array" if isinstance(data, list) else "other"
        ),
        "key_count": count_keys(data),
        "item_count": len(data) if isinstance(data, (dict, list)) else 1,
        "depth": get_depth(data),
        "size_original": len(raw),
        "size_yaml": len(yaml_str),
        "size_original_kb": round(len(raw) / 1024, 2),
        "size_yaml_kb": round(len(yaml_str) / 1024, 2),
        "sort_keys": sort_keys,
        "indent": indent,
        "encoding": encoding,
    }
