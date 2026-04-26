"""
UUID service functions.
"""

import logging

logger = logging.getLogger(__name__)


def generate_uuid(
    version: int = 4,
    count: int = 1,
    uppercase: bool = False,
    hyphens: bool = True,
    braces: bool = False,
    prefix: str = "",
    suffix: str = "",
    namespace: str = None,
    name: str = None,
    seed: int = None,
) -> dict:
    """
    Generate UUID(s) with multiple options.

    Args:
        version   : UUID version 1|3|4|5|6|7    (default: 4)
        count     : number of UUIDs to generate  (default: 1)
        uppercase : output in uppercase           (default: False)
        hyphens   : include hyphens              (default: True)
        braces    : wrap in curly braces {}       (default: False)
        prefix    : add prefix to each UUID       (default: '')
        suffix    : add suffix to each UUID       (default: '')
        namespace : namespace for v3/v5           (default: DNS)
                    dns | url | oid | x500 | custom UUID
        name      : name for v3/v5               (required for v3/v5)
        seed      : random seed for reproducible  (default: None)

    UUID Versions:
        v1  → time-based + MAC address
        v3  → MD5 hash of namespace + name
        v4  → random (most common)
        v5  → SHA-1 hash of namespace + name
        v6  → reordered time-based (sortable)
        v7  → Unix timestamp-based (sortable, modern)

    Returns:
        {
            'uuids'    : list,
            'version'  : int,
            'count'    : int,
            'options'  : dict,
        }
    """
    import uuid
    import random
    import time
    import os

    # ── Validate ──────────────────────────────────
    if version not in (1, 3, 4, 5, 6, 7):
        raise ValueError(f"Invalid version: {version}. " f"Supported: 1, 3, 4, 5, 6, 7")

    if not (1 <= count <= 1000):
        raise ValueError("count must be between 1 and 1000.")

    if version in (3, 5) and not name:
        raise ValueError(f'UUID v{version} requires a "name" parameter.')

    # ── Namespace for v3/v5 ───────────────────────
    namespace_map = {
        "dns": uuid.NAMESPACE_DNS,
        "url": uuid.NAMESPACE_URL,
        "oid": uuid.NAMESPACE_OID,
        "x500": uuid.NAMESPACE_X500,
    }

    if version in (3, 5):
        if namespace is None or namespace.lower() in namespace_map:
            ns = namespace_map.get(
                (namespace or "dns").lower(),
                uuid.NAMESPACE_DNS,
            )
        else:
            # Try custom UUID namespace
            try:
                ns = uuid.UUID(namespace)
            except ValueError:
                raise ValueError(
                    f'Invalid namespace: "{namespace}". '
                    f"Use: dns, url, oid, x500, or a valid UUID string."
                )

    # ── Seed for reproducible results ─────────────
    if seed is not None:
        random.seed(seed)

    # ── Generate UUIDs ────────────────────────────
    generated = []

    for _ in range(count):
        if version == 1:
            uid = uuid.uuid1()

        elif version == 3:
            uid = uuid.uuid3(ns, name)

        elif version == 4:
            if seed is not None:
                # Seeded random UUID
                rand_bytes = bytes(random.randint(0, 255) for _ in range(16))
                uid = uuid.UUID(bytes=rand_bytes, version=4)
            else:
                uid = uuid.uuid4()

        elif version == 5:
            uid = uuid.uuid5(ns, name)

        elif version == 6:
            uid = _generate_uuid6()

        elif version == 7:
            uid = _generate_uuid7()

        # ── Format output ──────────────────────────
        uid_str = str(uid)

        if not hyphens:
            uid_str = uid_str.replace("-", "")

        if uppercase:
            uid_str = uid_str.upper()

        if braces:
            uid_str = "{" + uid_str + "}"

        if prefix:
            uid_str = prefix + uid_str

        if suffix:
            uid_str = uid_str + suffix

        generated.append(uid_str)

    # ── Build info ────────────────────────────────
    version_info = {
        1: "Time-based + MAC address",
        3: "MD5 hash (namespace + name)",
        4: "Random (cryptographically secure)",
        5: "SHA-1 hash (namespace + name)",
        6: "Reordered time-based (sortable)",
        7: "Unix timestamp-based (sortable, modern)",
    }

    return {
        "uuids": generated,
        "version": version,
        "count": len(generated),
        "description": version_info.get(version, ""),
        "options": {
            "uppercase": uppercase,
            "hyphens": hyphens,
            "braces": braces,
            "prefix": prefix,
            "suffix": suffix,
            "namespace": namespace,
            "name": name,
            "seed": seed,
        },
    }


def _generate_uuid6() -> "uuid.UUID":
    """
    Generate UUID version 6 (reordered time-based, sortable).
    Reorders UUID v1 timestamp for lexicographic sorting.
    """
    import uuid
    import time

    # Get UUID v1 and reorder timestamp
    uid1 = uuid.uuid1()
    uid1_int = uid1.int

    # Extract time fields from v1
    time_low = (uid1_int >> 96) & 0xFFFFFFFF
    time_mid = (uid1_int >> 80) & 0xFFFF
    time_hi = (uid1_int >> 64) & 0x0FFF

    # Reorder: time_hi + time_mid + time_low (sortable)
    time_v6 = (time_hi << 48) | (time_mid << 32) | time_low

    # Rebuild UUID with v6
    clock_seq = (uid1_int >> 48) & 0x3FFF
    node = uid1_int & 0xFFFFFFFFFFFF

    uid6_int = (
        (time_v6 & 0x0FFFFFFFFFFFFFFF) << 64
        | 0x6000_0000_0000_0000
        | clock_seq << 48
        | node
    )

    return uuid.UUID(int=uid6_int)


def _generate_uuid7() -> "uuid.UUID":
    """
    Generate UUID version 7 (Unix timestamp milliseconds, sortable).
    Modern replacement for v1/v6 — monotonically increasing.
    """
    import uuid
    import time
    import os

    # 48-bit Unix timestamp in milliseconds
    ms = int(time.time() * 1000)
    rand_a = int.from_bytes(os.urandom(2), "big") & 0x0FFF
    rand_b = int.from_bytes(os.urandom(8), "big") & 0x3FFFFFFFFFFFFFFF

    uid7_int = (
        (ms & 0xFFFFFFFFFFFF) << 80
        | 0x7000_0000_0000_0000_0000
        | (rand_a & 0x0FFF) << 64
        | 0x8000_0000_0000_0000
        | (rand_b & 0x3FFFFFFFFFFFFFFF)
    )

    return uuid.UUID(int=uid7_int)


def validate_uuid(value: str) -> dict:
    """
    Validate a UUID string and return its info.

    Args:
        value : UUID string to validate

    Returns:
        { 'is_valid', 'version', 'variant', 'uuid', 'formatted' }
    """
    import uuid

    value = str(value).strip()

    # Strip braces and formatting
    clean = value.strip("{}").replace("-", "").strip()

    try:
        # Try parsing with hyphens first
        try:
            uid = uuid.UUID(value.strip("{}"))
        except ValueError:
            # Try without hyphens
            uid = uuid.UUID(clean)

        # ── Format all variants ────────────────────
        uid_str = str(uid)
        formatted = {
            "standard": uid_str,
            "uppercase": uid_str.upper(),
            "no_hyphens": uid_str.replace("-", ""),
            "braces": "{" + uid_str + "}",
            "urn": f"urn:uuid:{uid_str}",
            "int": uid.int,
            "hex": uid.hex,
            "bytes": list(uid.bytes),
        }

        version_info = {
            1: "Time-based + MAC address",
            3: "MD5 hash (namespace + name)",
            4: "Random",
            5: "SHA-1 hash (namespace + name)",
        }

        return {
            "is_valid": True,
            "uuid": uid_str,
            "version": uid.version,
            "variant": str(uid.variant),
            "description": version_info.get(uid.version, "Unknown"),
            "formatted": formatted,
        }

    except ValueError:
        return {
            "is_valid": False,
            "uuid": value,
            "error": f'"{value}" is not a valid UUID.',
        }


def bulk_generate_uuids(
    version: int = 4,
    count: int = 10,
    format: str = "standard",
) -> dict:
    """
    Generate bulk UUIDs in multiple export formats.

    Args:
        version : UUID version                (default: 4)
        count   : number of UUIDs 1-1000     (default: 10)
        format  : standard | csv | json | sql (default: standard)

    Returns:
        { 'uuids', 'export', 'format', 'count' }
    """
    import uuid as _uuid
    import json as _json

    if not (1 <= count <= 1000):
        raise ValueError("count must be between 1 and 1000.")

    result = generate_uuid(version=version, count=count)
    uuids = result["uuids"]

    if format == "csv":
        export = "uuid\n" + "\n".join(uuids)

    elif format == "json":
        export = _json.dumps(uuids, indent=2)

    elif format == "sql":
        rows = ",\n".join(f"  ('{u}')" for u in uuids)
        export = f"INSERT INTO uuids (id) VALUES\n{rows};"

    elif format == "array":
        items = ", ".join(f'"{u}"' for u in uuids)
        export = f"[{items}]"

    else:
        export = "\n".join(uuids)

    return {
        "uuids": uuids,
        "export": export,
        "format": format,
        "count": len(uuids),
        "version": version,
    }
