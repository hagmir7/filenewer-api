"""
Base64 service functions.
"""

import logging

logger = logging.getLogger(__name__)


def base64_encode(
    source,
    encoding: str = "standard",
    chunk_size: int = 0,
    filename: str = None,
) -> dict:
    """
    Encode text, file, or bytes to Base64.

    Args:
        source     : text string | file object | bytes
        encoding   : 'standard' | 'url_safe' | 'mime'  (default: standard)
        chunk_size : split output into chunks of N chars (default: 0 = no split)
        filename   : original filename for file inputs

    Returns:
        {
            'encoded'        : str,
            'encoding'       : str,
            'original_size'  : int,
            'encoded_size'   : int,
            'is_file'        : bool,
            'filename'       : str,
            'mime_type'      : str,
            'data_uri'       : str,
        }
    """
    import base64
    import mimetypes

    # ── Read source ───────────────────────────────
    is_file = False
    mime_type = "text/plain"

    if hasattr(source, "read"):
        # File object
        raw_bytes = source.read()
        is_file = True
        if filename:
            mime_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"
    elif isinstance(source, bytes):
        raw_bytes = source
    elif isinstance(source, str):
        raw_bytes = source.encode("utf-8")
    else:
        raise ValueError("source must be a string, bytes, or file object.")

    if not raw_bytes:
        raise ValueError("Empty input. Nothing to encode.")

    original_size = len(raw_bytes)

    # ── Encode ────────────────────────────────────
    if encoding == "url_safe":
        encoded_bytes = base64.urlsafe_b64encode(raw_bytes)
    elif encoding == "mime":
        # MIME encoding adds line breaks every 76 chars
        import base64 as b64

        encoded_bytes = b64.encodebytes(raw_bytes)
        encoded_str = encoded_bytes.decode("utf-8").strip()
    else:
        encoded_bytes = base64.b64encode(raw_bytes)

    if encoding != "mime":
        encoded_str = encoded_bytes.decode("utf-8")

    # ── Chunk output ──────────────────────────────
    if chunk_size > 0 and encoding != "mime":
        encoded_str = "\n".join(
            encoded_str[i : i + chunk_size]
            for i in range(0, len(encoded_str), chunk_size)
        )

    # ── Data URI ──────────────────────────────────
    clean_encoded = encoded_str.replace("\n", "")
    data_uri = f"data:{mime_type};base64,{clean_encoded}"

    return {
        "encoded": encoded_str,
        "encoding": encoding,
        "original_size": original_size,
        "encoded_size": len(encoded_str),
        "original_size_kb": round(original_size / 1024, 2),
        "encoded_size_kb": round(len(encoded_str) / 1024, 2),
        "is_file": is_file,
        "filename": filename or "",
        "mime_type": mime_type,
        "data_uri": data_uri if is_file else "",
        "overhead_percent": round(
            (len(encoded_str) - original_size) / original_size * 100, 2
        ),
    }


def base64_decode(
    source,
    encoding: str = "standard",
    as_text: bool = True,
    text_encoding: str = "utf-8",
) -> dict:
    """
    Decode Base64 string to text or bytes.

    Args:
        source        : Base64 encoded string | file object
        encoding      : 'standard' | 'url_safe'    (default: standard)
        as_text       : try to decode as text       (default: True)
        text_encoding : text encoding to use        (default: utf-8)

    Returns:
        {
            'decoded_text'  : str | None,
            'decoded_bytes' : bytes,
            'is_text'       : bool,
            'original_size' : int,
            'decoded_size'  : int,
            'encoding'      : str,
        }
    """
    import base64

    # ── Read source ───────────────────────────────
    if hasattr(source, "read"):
        raw = source.read()
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8")
    elif isinstance(source, str):
        raw = source
    else:
        raise ValueError("source must be a string or file object.")

    # ── Clean input ───────────────────────────────
    # Remove whitespace, newlines, data URI prefix
    raw = raw.strip()
    if raw.startswith("data:"):
        # Strip data URI prefix: data:image/png;base64,xxxxx
        raw = raw.split(",", 1)[-1]

    # Remove all whitespace (handles MIME line breaks)
    raw = "".join(raw.split())

    if not raw:
        raise ValueError("Empty input. Nothing to decode.")

    # ── Add padding if needed ─────────────────────
    missing_padding = len(raw) % 4
    if missing_padding:
        raw += "=" * (4 - missing_padding)

    # ── Decode ────────────────────────────────────
    try:
        if encoding == "url_safe":
            decoded_bytes = base64.urlsafe_b64decode(raw)
        else:
            decoded_bytes = base64.b64decode(raw)
    except Exception as e:
        raise ValueError(f"Invalid Base64 input: {e}")

    decoded_size = len(decoded_bytes)
    original_size = len(raw)

    # ── Try decode as text ────────────────────────
    decoded_text = None
    is_text = False

    if as_text:
        try:
            decoded_text = decoded_bytes.decode(text_encoding)
            is_text = True
        except (UnicodeDecodeError, LookupError):
            # Binary content — not text
            is_text = False
            decoded_text = None

    return {
        "decoded_text": decoded_text,
        "decoded_bytes": decoded_bytes,
        "is_text": is_text,
        "encoding": encoding,
        "original_size": original_size,
        "decoded_size": decoded_size,
        "original_size_kb": round(original_size / 1024, 2),
        "decoded_size_kb": round(decoded_size / 1024, 2),
        "text_encoding": text_encoding if is_text else None,
    }


def base64_validate(source) -> dict:
    """
    Validate whether a string is valid Base64.

    Args:
        source : string to validate

    Returns:
        { 'is_valid', 'encoding', 'decoded_size', 'message' }
    """
    import base64 as b64

    if hasattr(source, "read"):
        raw = source.read().decode("utf-8")
    else:
        raw = str(source).strip()

    # Strip data URI prefix
    if raw.startswith("data:"):
        raw = raw.split(",", 1)[-1]

    raw = "".join(raw.split())

    # Check character set
    standard_chars = set(
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/="
    )
    url_safe_chars = set(
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_="
    )

    is_url_safe = all(c in url_safe_chars for c in raw) and ("-" in raw or "_" in raw)

    # Add padding
    padded = raw + "=" * (-len(raw) % 4)

    for enc_name, decode_fn in [
        ("standard", b64.b64decode),
        ("url_safe", b64.urlsafe_b64decode),
    ]:
        try:
            decoded = decode_fn(padded)
            decoded_size = len(decoded)
            return {
                "is_valid": True,
                "encoding": enc_name,
                "decoded_size": decoded_size,
                "decoded_size_kb": round(decoded_size / 1024, 2),
                "input_size": len(raw),
                "message": f"Valid {enc_name} Base64 string.",
            }
        except Exception:
            continue

    return {
        "is_valid": False,
        "encoding": None,
        "message": "Invalid Base64 string.",
        "input_size": len(raw),
    }
