"""
Password generation and strength analysis service functions.
"""

import logging

logger = logging.getLogger(__name__)


def generate_password(
    length: int = 16,
    count: int = 1,
    uppercase: bool = True,
    lowercase: bool = True,
    digits: bool = True,
    symbols: bool = True,
    exclude_chars: str = "",
    exclude_similar: bool = False,
    exclude_ambiguous: bool = False,
    custom_chars: str = "",
    prefix: str = "",
    suffix: str = "",
    no_repeat: bool = False,
) -> dict:
    """
    Generate secure password(s) with full customization.

    Args:
        length           : password length 4-256         (default: 16)
        count            : number of passwords 1-100     (default: 1)
        uppercase        : include A-Z                   (default: True)
        lowercase        : include a-z                   (default: True)
        digits           : include 0-9                   (default: True)
        symbols          : include !@#$%^&*              (default: True)
        exclude_chars    : specific chars to exclude     (default: '')
        exclude_similar  : exclude 0O1lI                 (default: False)
        exclude_ambiguous: exclude {}[]()/'"`~,;:.<>     (default: False)
        custom_chars     : use only these characters     (default: '')
        prefix           : add prefix to password        (default: '')
        suffix           : add suffix to password        (default: '')
        no_repeat        : no repeating characters       (default: False)

    Returns:
        {
            'passwords'  : list,
            'count'      : int,
            'length'     : int,
            'strength'   : str,
            'entropy'    : float,
            'options'    : dict,
        }
    """
    import secrets
    import string
    import math

    # ── Validate ──────────────────────────────────
    if not (4 <= length <= 256):
        raise ValueError("length must be between 4 and 256.")

    if not (1 <= count <= 100):
        raise ValueError("count must be between 1 and 100.")

    # ── Build character pool ───────────────────────
    if custom_chars:
        pool = custom_chars
    else:
        pool = ""
        if uppercase:
            pool += string.ascii_uppercase
        if lowercase:
            pool += string.ascii_lowercase
        if digits:
            pool += string.digits
        if symbols:
            pool += "!@#$%^&*()-_=+[]{}|;:,.<>?"

        if not pool:
            raise ValueError("At least one character type must be selected.")

    # ── Exclude similar characters ─────────────────
    if exclude_similar:
        for ch in "0O1lI":
            pool = pool.replace(ch, "")

    # ── Exclude ambiguous characters ───────────────
    if exclude_ambiguous:
        for ch in r"{}[]()\/'\"`~,;:.<>":
            pool = pool.replace(ch, "")

    # ── Exclude specific characters ────────────────
    for ch in exclude_chars:
        pool = pool.replace(ch, "")

    if not pool:
        raise ValueError(
            "Character pool is empty after exclusions. " "Please adjust your settings."
        )

    # ── Check no_repeat feasibility ────────────────
    actual_length = length - len(prefix) - len(suffix)

    if no_repeat and len(pool) < actual_length:
        raise ValueError(
            f"Cannot generate a {actual_length}-char password "
            f"without repeating from a pool of {len(pool)} characters. "
            f"Reduce length or disable no_repeat."
        )

    # ── Generate passwords ─────────────────────────
    passwords = []
    pool_list = list(pool)

    for _ in range(count):
        if no_repeat:
            # Sample without replacement
            chars = secrets.SystemRandom().sample(pool_list, actual_length)
        else:
            chars = [secrets.choice(pool_list) for _ in range(actual_length)]

        # Ensure at least one char from each required set
        if not custom_chars and not no_repeat:
            required = []
            if uppercase and not exclude_similar:
                uc = [
                    c
                    for c in string.ascii_uppercase
                    if c not in exclude_chars and c not in "0O1lI" * exclude_similar
                ]
                if uc:
                    required.append(secrets.choice(uc))
            if lowercase and not exclude_similar:
                lc = [
                    c
                    for c in string.ascii_lowercase
                    if c not in exclude_chars and c not in "0O1lI" * exclude_similar
                ]
                if lc:
                    required.append(secrets.choice(lc))
            if digits:
                dg = [c for c in string.digits if c not in exclude_chars]
                if dg:
                    required.append(secrets.choice(dg))
            if symbols:
                sy = [c for c in "!@#$%^&*()-_=+[]{}|;:,.<>?" if c not in exclude_chars]
                if sy:
                    required.append(secrets.choice(sy))

            # Replace random positions with required chars
            for i, req_char in enumerate(required):
                if i < len(chars):
                    pos = secrets.randbelow(len(chars))
                    chars[pos] = req_char

        password = prefix + "".join(chars) + suffix
        passwords.append(password)

    # ── Calculate entropy ──────────────────────────
    pool_size = len(pool)
    entropy = actual_length * math.log2(pool_size) if pool_size > 1 else 0

    # ── Calculate strength ─────────────────────────
    strength = _calculate_strength(entropy)

    # ── Crack time estimate ───────────────────────
    crack_time = _estimate_crack_time(entropy)

    return {
        "passwords": passwords,
        "count": len(passwords),
        "length": length,
        "pool_size": pool_size,
        "entropy": round(entropy, 2),
        "strength": strength,
        "crack_time": crack_time,
        "options": {
            "uppercase": uppercase,
            "lowercase": lowercase,
            "digits": digits,
            "symbols": symbols,
            "exclude_similar": exclude_similar,
            "exclude_ambiguous": exclude_ambiguous,
            "exclude_chars": exclude_chars,
            "no_repeat": no_repeat,
            "prefix": prefix,
            "suffix": suffix,
            "custom_chars": custom_chars,
        },
    }


def _calculate_strength(entropy: float) -> str:
    """Calculate password strength from entropy bits."""
    if entropy < 28:
        return "very_weak"
    if entropy < 36:
        return "weak"
    if entropy < 60:
        return "moderate"
    if entropy < 128:
        return "strong"
    return "very_strong"


def _estimate_crack_time(entropy: float) -> dict:
    """
    Estimate time to crack password at different attack speeds.

    Attack speeds (guesses/second):
        online_throttled  : 100/s      (online with rate limiting)
        online_unthrottled: 10,000/s   (online without limiting)
        offline_slow      : 1M/s       (bcrypt/scrypt)
        offline_fast      : 100B/s     (MD5/SHA1 GPU)
        massive_attack    : 100T/s     (nation state)
    """
    import math

    combinations = 2**entropy

    speeds = {
        "online_throttled": 100,
        "online_unthrottled": 10_000,
        "offline_slow": 1_000_000,
        "offline_fast": 100_000_000_000,
        "massive_attack": 100_000_000_000_000,
    }

    def format_time(seconds: float) -> str:
        if seconds < 1:
            return "less than a second"
        if seconds < 60:
            return f"{int(seconds)} seconds"
        if seconds < 3600:
            return f"{int(seconds/60)} minutes"
        if seconds < 86400:
            return f"{int(seconds/3600)} hours"
        if seconds < 2592000:
            return f"{int(seconds/86400)} days"
        if seconds < 31536000:
            return f"{int(seconds/2592000)} months"
        if seconds < 3153600000:
            return f"{int(seconds/31536000)} years"
        centuries = seconds / 3153600000
        if centuries < 1000:
            return f"{int(centuries)} centuries"
        return "longer than the age of the universe"

    return {
        name: format_time(combinations / speed / 2) for name, speed in speeds.items()
    }


def check_password_strength(password: str) -> dict:
    """
    Analyze the strength of an existing password.

    Args:
        password : password string to analyze

    Returns:
        {
            'strength'   : str,
            'entropy'    : float,
            'score'      : int,
            'issues'     : list,
            'suggestions': list,
            'crack_time' : dict,
        }
    """
    import math
    import string
    import re

    if not password:
        raise ValueError("Password cannot be empty.")

    length = len(password)

    # ── Build charset size ────────────────────────
    charset = 0
    has_upper = bool(re.search(r"[A-Z]", password))
    has_lower = bool(re.search(r"[a-z]", password))
    has_digit = bool(re.search(r"\d", password))
    has_symbol = bool(re.search(r"[^A-Za-z0-9]", password))
    has_space = " " in password

    if has_upper:
        charset += 26
    if has_lower:
        charset += 26
    if has_digit:
        charset += 10
    if has_symbol:
        charset += 32
    if has_space:
        charset += 1

    if charset == 0:
        charset = 1

    entropy = length * math.log2(charset)

    # ── Score (0-100) ─────────────────────────────
    score = 0
    score += min(length * 4, 40)  # length up to 40pts
    score += 10 if has_upper else 0
    score += 10 if has_lower else 0
    score += 10 if has_digit else 0
    score += 15 if has_symbol else 0
    score += 5 if length >= 20 else 0
    score = min(score, 100)

    # ── Issues ────────────────────────────────────
    issues = []
    suggestions = []

    if length < 8:
        issues.append("Too short (less than 8 characters).")
        suggestions.append("Use at least 8 characters.")

    if length < 12:
        suggestions.append("Consider using 12+ characters for better security.")

    if not has_upper:
        issues.append("No uppercase letters.")
        suggestions.append("Add uppercase letters (A-Z).")

    if not has_lower:
        issues.append("No lowercase letters.")
        suggestions.append("Add lowercase letters (a-z).")

    if not has_digit:
        issues.append("No digits.")
        suggestions.append("Add numbers (0-9).")

    if not has_symbol:
        issues.append("No special characters.")
        suggestions.append("Add symbols (!@#$%^&*).")

    # Check for repeating patterns
    if re.search(r"(.)\1{2,}", password):
        issues.append('Contains repeating characters (e.g. "aaa").')
        suggestions.append("Avoid repeating the same character multiple times.")

    # Check for sequential chars
    if re.search(
        r"(012|123|234|345|456|567|678|789|890|abc|bcd|cde)", password.lower()
    ):
        issues.append('Contains sequential characters (e.g. "123", "abc").')
        suggestions.append("Avoid predictable sequences.")

    # Common patterns
    common_patterns = [
        "password",
        "qwerty",
        "admin",
        "123456",
        "letmein",
        "welcome",
        "monkey",
        "dragon",
    ]
    if any(p in password.lower() for p in common_patterns):
        issues.append("Contains common password patterns.")
        suggestions.append("Avoid common words and patterns.")

    strength = _calculate_strength(entropy)
    crack_time = _estimate_crack_time(entropy)

    return {
        "password": "*" * min(len(password), 4) + "...",  # masked
        "length": length,
        "strength": strength,
        "score": score,
        "entropy": round(entropy, 2),
        "charset_size": charset,
        "has_uppercase": has_upper,
        "has_lowercase": has_lower,
        "has_digits": has_digit,
        "has_symbols": has_symbol,
        "issues": issues,
        "suggestions": suggestions,
        "crack_time": crack_time,
    }


def generate_passphrase(
    words: int = 4,
    count: int = 1,
    separator: str = "-",
    capitalize: bool = True,
    include_digit: bool = True,
) -> dict:
    """
    Generate memorable passphrase using random words.

    Args:
        words        : number of words 3-10     (default: 4)
        count        : number of passphrases    (default: 1)
        separator    : word separator           (default: -)
        capitalize   : capitalize each word     (default: True)
        include_digit: append random digit      (default: True)

    Returns:
        { 'passphrases', 'count', 'words', 'entropy', 'strength' }
    """
    import secrets
    import math

    if not (3 <= words <= 10):
        raise ValueError("words must be between 3 and 10.")
    if not (1 <= count <= 100):
        raise ValueError("count must be between 1 and 100.")

    # ── Word list (EFF large wordlist subset) ──────
    word_list = [
        "apple",
        "brave",
        "cloud",
        "dance",
        "earth",
        "flame",
        "grace",
        "happy",
        "ivory",
        "joker",
        "kneel",
        "lemon",
        "magic",
        "noble",
        "ocean",
        "peace",
        "queen",
        "river",
        "stone",
        "tiger",
        "ultra",
        "vapor",
        "water",
        "xenon",
        "yacht",
        "zebra",
        "amber",
        "blaze",
        "crisp",
        "daisy",
        "eagle",
        "frost",
        "gloom",
        "haste",
        "input",
        "jolly",
        "karma",
        "lunar",
        "marsh",
        "nerve",
        "orbit",
        "pixel",
        "quill",
        "rainy",
        "solar",
        "thorn",
        "umbra",
        "vivid",
        "windy",
        "xeric",
        "young",
        "zesty",
        "agile",
        "blend",
        "coral",
        "drift",
        "elite",
        "fauna",
        "globe",
        "haven",
        "ideal",
        "jazzy",
        "knack",
        "light",
        "mocha",
        "night",
        "olive",
        "prism",
        "quirk",
        "radio",
        "sigma",
        "token",
        "ultra",
        "vague",
        "woken",
        "axiom",
        "birch",
        "cedar",
        "depot",
        "epoch",
        "flair",
        "giant",
        "hyper",
        "index",
        "joint",
        "kraft",
        "lilac",
        "maple",
        "north",
        "oasis",
        "piano",
        "quota",
        "radar",
        "shelf",
        "trove",
        "unity",
        "viola",
        "waltz",
        "exist",
        "years",
        "azure",
        "boxer",
        "cider",
        "delta",
        "ember",
        "field",
        "grand",
        "herbs",
        "inlet",
        "jewel",
        "kapok",
        "largo",
        "mango",
        "niche",
        "onset",
        "perch",
        "quaff",
        "renew",
        "swamp",
        "tryst",
        "upper",
        "verve",
        "wagon",
        "extra",
        "yield",
    ]

    passphrases = []

    for _ in range(count):
        selected = [secrets.choice(word_list) for _ in range(words)]

        if capitalize:
            selected = [w.capitalize() for w in selected]

        phrase = separator.join(selected)

        if include_digit:
            phrase += separator + str(secrets.randbelow(9999)).zfill(4)

        passphrases.append(phrase)

    # ── Entropy calculation ────────────────────────
    # Each word adds log2(wordlist_size) bits
    word_entropy = words * math.log2(len(word_list))
    digit_entropy = math.log2(9999) if include_digit else 0
    total_entropy = word_entropy + digit_entropy

    return {
        "passphrases": passphrases,
        "count": len(passphrases),
        "words": words,
        "separator": separator,
        "entropy": round(total_entropy, 2),
        "strength": _calculate_strength(total_entropy),
        "crack_time": _estimate_crack_time(total_entropy),
    }


def generate_hash(
    source,
    algorithms_list: list = None,
    encoding: str = "utf-8",
    hmac_key: str = None,
    output_format: str = "hex",
) -> dict:
    """
    Generate hash(es) from text or file.

    Args:
        source          : text string | file object | bytes
        algorithms_list : list of algorithms to use
                          md5 | sha1 | sha224 | sha256 | sha384 | sha512
                          sha3_224 | sha3_256 | sha3_384 | sha3_512
                          blake2b | blake2s | shake_128 | shake_256
                          default: all
        encoding        : text encoding                 (default: utf-8)
        hmac_key        : HMAC secret key              (default: None)
        output_format   : hex | base64 | base64url | int (default: hex)

    Returns:
        {
            'hashes'      : dict,
            'input_type'  : str,
            'input_size'  : int,
            'encoding'    : str,
            'output_format': str,
            'is_hmac'     : bool,
        }
    """
    import hashlib
    import hmac as hmac_lib
    import base64

    # ── Read source ───────────────────────────────
    input_type = "text"

    if hasattr(source, "read"):
        data = source.read()
        input_type = "file"
        if isinstance(data, str):
            data = data.encode(encoding)
    elif isinstance(source, bytes):
        data = source
        input_type = "bytes"
    elif isinstance(source, str):
        data = source.encode(encoding)
        input_type = "text"
    else:
        raise ValueError("source must be a string, bytes, or file object.")

    if not data:
        raise ValueError("Empty input. Nothing to hash.")

    # ── Validate output format ─────────────────────
    valid_formats = ("hex", "base64", "base64url", "int")
    if output_format not in valid_formats:
        raise ValueError(
            f'Invalid output_format: "{output_format}". '
            f"Must be one of: {valid_formats}"
        )

    # ── All supported algorithms ───────────────────
    all_algorithms = {
        "md5": {"fn": hashlib.md5, "bits": 128, "secure": False},
        "sha1": {"fn": hashlib.sha1, "bits": 160, "secure": False},
        "sha224": {"fn": hashlib.sha224, "bits": 224, "secure": True},
        "sha256": {"fn": hashlib.sha256, "bits": 256, "secure": True},
        "sha384": {"fn": hashlib.sha384, "bits": 384, "secure": True},
        "sha512": {"fn": hashlib.sha512, "bits": 512, "secure": True},
        "sha3_224": {"fn": hashlib.sha3_224, "bits": 224, "secure": True},
        "sha3_256": {"fn": hashlib.sha3_256, "bits": 256, "secure": True},
        "sha3_384": {"fn": hashlib.sha3_384, "bits": 384, "secure": True},
        "sha3_512": {"fn": hashlib.sha3_512, "bits": 512, "secure": True},
        "blake2b": {"fn": hashlib.blake2b, "bits": 512, "secure": True},
        "blake2s": {"fn": hashlib.blake2s, "bits": 256, "secure": True},
        "shake_128": {"fn": hashlib.shake_128, "bits": 128, "secure": True},
        "shake_256": {"fn": hashlib.shake_256, "bits": 256, "secure": True},
    }

    # ── Select algorithms ─────────────────────────
    if algorithms_list is None:
        selected = list(all_algorithms.keys())
    else:
        selected = [a.lower().strip() for a in algorithms_list]
        invalid = [a for a in selected if a not in all_algorithms]
        if invalid:
            raise ValueError(
                f"Unsupported algorithm(s): {invalid}. "
                f"Supported: {list(all_algorithms.keys())}"
            )

    # ── Format digest ─────────────────────────────
    def format_digest(digest_bytes: bytes) -> str:
        if output_format == "hex":
            return digest_bytes.hex()
        elif output_format == "base64":
            return base64.b64encode(digest_bytes).decode("utf-8")
        elif output_format == "base64url":
            return base64.urlsafe_b64encode(digest_bytes).decode("utf-8")
        elif output_format == "int":
            return str(int.from_bytes(digest_bytes, "big"))
        return digest_bytes.hex()

    # ── Compute hashes ────────────────────────────
    hashes = {}
    is_hmac = bool(hmac_key)
    hmac_key_b = hmac_key.encode(encoding) if hmac_key else None

    for algo in selected:
        info = all_algorithms[algo]
        try:
            if is_hmac:
                # HMAC mode
                if algo in ("shake_128", "shake_256"):
                    hashes[algo] = {
                        "hash": "[HMAC not supported for SHAKE]",
                        "bits": info["bits"],
                        "secure": info["secure"],
                        "error": True,
                    }
                    continue

                h = hmac_lib.new(
                    hmac_key_b,
                    data,
                    info["fn"],
                )
                digest = h.digest()

            else:
                # Regular hash
                if algo == "shake_128":
                    h = hashlib.shake_128(data)
                    digest = h.digest(16)  # 128 bits = 16 bytes
                elif algo == "shake_256":
                    h = hashlib.shake_256(data)
                    digest = h.digest(32)  # 256 bits = 32 bytes
                else:
                    h = info["fn"](data)
                    digest = h.digest()

            hashes[algo] = {
                "hash": format_digest(digest),
                "bits": info["bits"],
                "length": (
                    len(digest) * 2
                    if output_format == "hex"
                    else len(format_digest(digest))
                ),
                "secure": info["secure"],
                "error": False,
            }

        except Exception as e:
            hashes[algo] = {
                "hash": None,
                "error": True,
                "message": str(e),
            }

    return {
        "hashes": hashes,
        "input_type": input_type,
        "input_size": len(data),
        "input_size_kb": round(len(data) / 1024, 2),
        "encoding": encoding,
        "output_format": output_format,
        "is_hmac": is_hmac,
        "algorithm_count": len(hashes),
    }


def compare_hashes(
    hash1: str,
    hash2: str,
) -> dict:
    """
    Securely compare two hash strings.
    Uses constant-time comparison to prevent timing attacks.

    Args:
        hash1 : first hash string
        hash2 : second hash string

    Returns:
        { 'match', 'hash1_length', 'hash2_length' }
    """
    import hmac

    h1 = hash1.strip().lower()
    h2 = hash2.strip().lower()

    # Constant-time comparison
    match = hmac.compare_digest(h1, h2)

    return {
        "match": match,
        "hash1_length": len(h1),
        "hash2_length": len(h2),
        "message": "Hashes match." if match else "Hashes do not match.",
    }
