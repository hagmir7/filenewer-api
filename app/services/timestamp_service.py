"""
Timestamp conversion service functions.
"""

import logging

logger = logging.getLogger(__name__)


def convert_timestamp(
    value: str,
    from_tz: str = "UTC",
    to_tz: str = "UTC",
    from_format: str = None,
    to_format: str = None,
) -> dict:
    """
    Convert timestamps between formats and timezones.

    Args:
        value       : timestamp value to convert
                      - Unix timestamp  : '1704067200'
                      - ISO 8601        : '2024-01-01T00:00:00Z'
                      - Human readable  : '2024-01-01 12:00:00'
                      - Date only       : '2024-01-01'
                      - Relative        : 'now', 'today', 'yesterday'
        from_tz     : source timezone               (default: UTC)
        to_tz       : target timezone               (default: UTC)
        from_format : input format (strptime)       (default: auto-detect)
        to_format   : output format (strftime)      (default: all formats)

    Returns:
        dict with all converted formats
    """
    from datetime import datetime, timezone, timedelta
    import time
    import calendar

    try:
        import zoneinfo

        def get_tz(name):
            if name.upper() == "UTC":
                return timezone.utc
            return zoneinfo.ZoneInfo(name)

    except ImportError:
        try:
            import pytz

            def get_tz(name):
                if name.upper() == "UTC":
                    return pytz.utc
                return pytz.timezone(name)

        except ImportError:

            def get_tz(name):
                return timezone.utc

    # ── Parse relative values ─────────────────────
    now = datetime.now(timezone.utc)

    relative_map = {
        "now": now,
        "today": now.replace(hour=0, minute=0, second=0, microsecond=0),
        "yesterday": (now - timedelta(days=1)).replace(
            hour=0, minute=0, second=0, microsecond=0
        ),
        "tomorrow": (now + timedelta(days=1)).replace(
            hour=0, minute=0, second=0, microsecond=0
        ),
    }

    value = str(value).strip()

    if value.lower() in relative_map:
        dt = relative_map[value.lower()]

    # ── Parse Unix timestamp ───────────────────────
    elif value.lstrip("-").replace(".", "").isdigit():
        ts = float(value)

        # Detect milliseconds vs seconds
        if abs(ts) > 1e10:
            ts = ts / 1000  # convert ms → seconds

        dt = datetime.fromtimestamp(ts, tz=timezone.utc)

    # ── Parse with custom format ───────────────────
    elif from_format:
        try:
            dt = datetime.strptime(value, from_format)
            if dt.tzinfo is None:
                src_tz = get_tz(from_tz)
                dt = dt.replace(tzinfo=src_tz)
        except ValueError as e:
            raise ValueError(f'Cannot parse "{value}" with format "{from_format}": {e}')

    # ── Auto-detect format ─────────────────────────
    else:
        dt = _parse_datetime_auto(value, from_tz, get_tz)

    # ── Convert to target timezone ─────────────────
    try:
        target_tz = get_tz(to_tz)
        dt_target = dt.astimezone(target_tz)
    except Exception:
        dt_target = dt

    # ── Unix timestamps ────────────────────────────
    unix_seconds = int(dt_target.timestamp())
    unix_ms = int(dt_target.timestamp() * 1000)
    unix_ns = int(dt_target.timestamp() * 1_000_000_000)

    # ── Format output ─────────────────────────────
    formats = {
        "unix_seconds": unix_seconds,
        "unix_ms": unix_ms,
        "unix_ns": unix_ns,
        "iso_8601": dt_target.strftime("%Y-%m-%dT%H:%M:%S") + _tz_offset(dt_target),
        "iso_8601_utc": dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "rfc_2822": dt_target.strftime("%a, %d %b %Y %H:%M:%S ")
        + _tz_offset(dt_target),
        "rfc_3339": dt_target.strftime("%Y-%m-%dT%H:%M:%S") + _tz_offset(dt_target),
        "human_readable": dt_target.strftime("%B %d, %Y %I:%M:%S %p"),
        "date_only": dt_target.strftime("%Y-%m-%d"),
        "time_only": dt_target.strftime("%H:%M:%S"),
        "datetime_local": dt_target.strftime("%Y-%m-%d %H:%M:%S"),
        "day_of_week": dt_target.strftime("%A"),
        "day_of_year": dt_target.timetuple().tm_yday,
        "week_number": dt_target.strftime("%W"),
        "quarter": f"Q{(dt_target.month - 1) // 3 + 1}",
        "relative": _get_relative_time(dt_target),
        "utc_offset": _tz_offset(dt_target),
        "timezone": to_tz,
    }

    if to_format:
        try:
            formats["custom"] = dt_target.strftime(to_format)
        except Exception as e:
            formats["custom_error"] = str(e)

    return {
        "input": value,
        "from_timezone": from_tz,
        "to_timezone": to_tz,
        "is_dst": _is_dst(dt_target),
        "formats": formats,
    }


def _parse_datetime_auto(value: str, from_tz: str, get_tz) -> "datetime":
    """Auto-detect and parse datetime string."""
    from datetime import datetime, timezone

    # Common formats to try in order
    format_attempts = [
        "%Y-%m-%dT%H:%M:%SZ",  # ISO 8601 UTC
        "%Y-%m-%dT%H:%M:%S%z",  # ISO 8601 with offset
        "%Y-%m-%dT%H:%M:%S",  # ISO 8601 no tz
        "%Y-%m-%dT%H:%M:%S.%fZ",  # ISO 8601 with microseconds UTC
        "%Y-%m-%dT%H:%M:%S.%f%z",  # ISO 8601 with microseconds
        "%Y-%m-%d %H:%M:%S",  # Common datetime
        "%Y-%m-%d %H:%M",  # No seconds
        "%Y-%m-%d",  # Date only
        "%d/%m/%Y %H:%M:%S",  # DD/MM/YYYY
        "%d/%m/%Y %H:%M",  # DD/MM/YYYY no seconds
        "%d/%m/%Y",  # DD/MM/YYYY date only
        "%m/%d/%Y %H:%M:%S",  # US format
        "%m/%d/%Y %H:%M",  # US format no seconds
        "%m/%d/%Y",  # US date only
        "%d-%m-%Y %H:%M:%S",  # DD-MM-YYYY
        "%d-%m-%Y",  # DD-MM-YYYY date only
        "%B %d, %Y %I:%M:%S %p",  # Human readable
        "%B %d, %Y",  # Month Day Year
        "%b %d, %Y %H:%M:%S",  # Short month
        "%b %d, %Y",  # Short month date only
        "%a, %d %b %Y %H:%M:%S %z",  # RFC 2822
        "%Y%m%d",  # Compact date
        "%Y%m%dT%H%M%S",  # Compact datetime
    ]

    for fmt in format_attempts:
        try:
            dt = datetime.strptime(value, fmt)
            if dt.tzinfo is None:
                src_tz = get_tz(from_tz)
                dt = dt.replace(tzinfo=src_tz)
            return dt
        except ValueError:
            continue

    raise ValueError(
        f'Cannot parse timestamp: "{value}". ' f"Try providing from_format explicitly."
    )


def _tz_offset(dt) -> str:
    """Get timezone offset string like +05:00 or Z."""
    from datetime import timezone

    if dt.tzinfo is None:
        return ""
    offset = dt.utcoffset()
    if offset is None:
        return ""
    total_seconds = int(offset.total_seconds())
    if total_seconds == 0:
        return "+00:00"
    sign = "+" if total_seconds >= 0 else "-"
    abs_sec = abs(total_seconds)
    hours = abs_sec // 3600
    minutes = (abs_sec % 3600) // 60
    return f"{sign}{hours:02d}:{minutes:02d}"


def _get_relative_time(dt) -> str:
    """Get human-readable relative time like '2 hours ago'."""
    from datetime import datetime, timezone

    now = datetime.now(timezone.utc)
    delta = now - dt.astimezone(timezone.utc)
    secs = int(delta.total_seconds())
    abs_s = abs(secs)
    future = secs < 0

    def fmt(n, unit):
        label = f'{n} {unit}{"s" if n != 1 else ""}'
        return f"in {label}" if future else f"{label} ago"

    if abs_s < 5:
        return "just now"
    if abs_s < 60:
        return fmt(abs_s, "second")
    if abs_s < 3600:
        return fmt(abs_s // 60, "minute")
    if abs_s < 86400:
        return fmt(abs_s // 3600, "hour")
    if abs_s < 604800:
        return fmt(abs_s // 86400, "day")
    if abs_s < 2592000:
        return fmt(abs_s // 604800, "week")
    if abs_s < 31536000:
        return fmt(abs_s // 2592000, "month")
    return fmt(abs_s // 31536000, "year")


def _is_dst(dt) -> bool:
    """Check if datetime is in daylight saving time."""
    try:
        import time

        ts = dt.timestamp()
        local_tm = time.localtime(ts)
        return bool(local_tm.tm_isdst)
    except Exception:
        return False


def batch_convert_timestamps(
    values: list,
    from_tz: str = "UTC",
    to_tz: str = "UTC",
    to_format: str = None,
) -> list:
    """
    Convert multiple timestamps at once.

    Args:
        values    : list of timestamp strings
        from_tz   : source timezone
        to_tz     : target timezone
        to_format : output format (strftime)

    Returns:
        list of conversion results
    """
    results = []
    for value in values:
        try:
            result = convert_timestamp(
                value,
                from_tz=from_tz,
                to_tz=to_tz,
                to_format=to_format,
            )
            results.append({"input": value, "success": True, **result})
        except Exception as e:
            results.append(
                {
                    "input": value,
                    "success": False,
                    "error": str(e),
                }
            )
    return results
