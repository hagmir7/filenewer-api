"""
Services package — re-exports all public functions for backward compatibility.

Import everything from individual service modules so that existing code using
    from app.services import some_function
continues to work unchanged.
"""

# ── PDF services ──────────────────────────────────────────────────────────────
from .pdf_service import (
    _looks_like_heading,
    _add_table_to_doc,
    convert_pdf_to_docx,
    pdf_to_excel,
    pdf_to_jpg,
    ARABIC_SUPPORT,
    is_arabic,
    process_arabic_text,
    register_arabic_fonts,
    encrypt_pdf,
    build_permissions,
    decrypt_pdf,
    get_pdf_info,
    compress_pdf,
    pdf_to_png,
    rotate_pdf,
    get_pdf_page_info,
    watermark_pdf,
    merge_pdfs,
    _add_page_numbers_to_writer,
    split_pdf,
)

# ── CSV services ──────────────────────────────────────────────────────────────
from .csv_service import (
    sanitize_name,
    infer_sql_type,
    csv_to_sql,
    csv_to_json,
    csv_to_excel,
    csv_to_excel_multisheets,
    view_csv,
)

# ── JSON services ─────────────────────────────────────────────────────────────
from .json_service import (
    json_to_csv,
    json_to_excel,
    json_to_excel_multisheets,
    format_json,
)

# ── Excel services ────────────────────────────────────────────────────────────
from .excel_service import (
    excel_to_csv,
    excel_to_markdown,
    _safe_auto_fit,
    markdown_to_excel,
)

# ── Word / DOCX services ──────────────────────────────────────────────────────
from .word_service import (
    word_to_pdf,
    word_to_jpg,
    word_to_txt,
    txt_to_word,
    merge_docx,
    _add_toc,
    split_docx,
    word_to_markdown,
    markdown_to_word,
)

# ── Base64 services ───────────────────────────────────────────────────────────
from .base64_service import (
    base64_encode,
    base64_decode,
    base64_validate,
)

# ── UUID services ─────────────────────────────────────────────────────────────
from .uuid_service import (
    generate_uuid,
    _generate_uuid6,
    _generate_uuid7,
    validate_uuid,
    bulk_generate_uuids,
)

# ── Crypto / file encryption services ────────────────────────────────────────
from .crypto_service import (
    encrypt_file,
    decrypt_file,
    get_file_hash,
    generate_checksum,
)

# ── Password services ─────────────────────────────────────────────────────────
from .password_service import (
    generate_password,
    _calculate_strength,
    _estimate_crack_time,
    check_password_strength,
    generate_passphrase,
    generate_hash,
    compare_hashes,
)

# ── Timestamp services ────────────────────────────────────────────────────────
from .timestamp_service import (
    convert_timestamp,
    _parse_datetime_auto,
    _tz_offset,
    _get_relative_time,
    _is_dst,
    batch_convert_timestamps,
)

# ── Compare / diff services ───────────────────────────────────────────────────
from .compare_service import (
    compare_texts,
    _generate_unified_diff,
    _generate_html_diff,
    _generate_side_by_side,
    compare_files,
)

# ── OCR services ──────────────────────────────────────────────────────────────
from .ocr_service import (
    ocr_pdf,
    _preprocess_image_for_ocr,
)

__all__ = [
    # pdf_service
    "_looks_like_heading",
    "_add_table_to_doc",
    "convert_pdf_to_docx",
    "pdf_to_excel",
    "pdf_to_jpg",
    "ARABIC_SUPPORT",
    "is_arabic",
    "process_arabic_text",
    "register_arabic_fonts",
    "encrypt_pdf",
    "build_permissions",
    "decrypt_pdf",
    "get_pdf_info",
    "compress_pdf",
    "pdf_to_png",
    "rotate_pdf",
    "get_pdf_page_info",
    "watermark_pdf",
    "merge_pdfs",
    "_add_page_numbers_to_writer",
    "split_pdf",
    # csv_service
    "sanitize_name",
    "infer_sql_type",
    "csv_to_sql",
    "csv_to_json",
    "csv_to_excel",
    "csv_to_excel_multisheets",
    "view_csv",
    # json_service
    "json_to_csv",
    "json_to_excel",
    "json_to_excel_multisheets",
    "format_json",
    # excel_service
    "excel_to_csv",
    "excel_to_markdown",
    "_safe_auto_fit",
    "markdown_to_excel",
    # word_service
    "word_to_pdf",
    "word_to_jpg",
    "word_to_txt",
    "txt_to_word",
    "merge_docx",
    "_add_toc",
    "split_docx",
    "word_to_markdown",
    "markdown_to_word",
    # base64_service
    "base64_encode",
    "base64_decode",
    "base64_validate",
    # uuid_service
    "generate_uuid",
    "_generate_uuid6",
    "_generate_uuid7",
    "validate_uuid",
    "bulk_generate_uuids",
    # crypto_service
    "encrypt_file",
    "decrypt_file",
    "get_file_hash",
    "generate_checksum",
    # password_service
    "generate_password",
    "_calculate_strength",
    "_estimate_crack_time",
    "check_password_strength",
    "generate_passphrase",
    "generate_hash",
    "compare_hashes",
    # timestamp_service
    "convert_timestamp",
    "_parse_datetime_auto",
    "_tz_offset",
    "_get_relative_time",
    "_is_dst",
    "batch_convert_timestamps",
    # compare_service
    "compare_texts",
    "_generate_unified_diff",
    "_generate_html_diff",
    "_generate_side_by_side",
    "compare_files",
    # ocr_service
    "ocr_pdf",
    "_preprocess_image_for_ocr",
]
