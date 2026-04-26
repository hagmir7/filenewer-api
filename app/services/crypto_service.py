"""
Cryptography / file encryption service functions.
"""

import logging

logger = logging.getLogger(__name__)


def encrypt_file(
    source,
    password: str,
    algorithm: str = "AES-256-GCM",
    filename: str = None,
) -> dict:
    """
    Encrypt any file using symmetric encryption.

    Args:
        source    : uploaded file object OR raw bytes
        password  : encryption password
        algorithm : AES-256-GCM | AES-256-CBC | ChaCha20   (default: AES-256-GCM)
        filename  : original filename

    Encryption details:
        AES-256-GCM   → authenticated encryption, best security (recommended)
        AES-256-CBC   → classic encryption, widely compatible
        ChaCha20      → fast, secure, good for mobile/embedded

    File format (encrypted output):
        [4 bytes]  magic header  "ENC1"
        [1 byte]   algorithm ID  (1=GCM, 2=CBC, 3=ChaCha20)
        [1 byte]   filename_len
        [N bytes]  filename (utf-8)
        [16 bytes] salt
        [12/16 bytes] IV / nonce
        [16 bytes] tag (GCM only, else empty)
        [N bytes]  encrypted data

    Returns:
        {
            'encrypted_bytes' : bytes,
            'algorithm'       : str,
            'original_size'   : int,
            'encrypted_size'  : int,
            'filename'        : str,
        }
    """
    import os
    import struct
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes, padding
    from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
    from cryptography.hazmat.backends import default_backend

    # ── Read source ───────────────────────────────
    if hasattr(source, "read"):
        raw_bytes = source.read()
        if filename is None and hasattr(source, "name"):
            filename = source.name
    elif isinstance(source, bytes):
        raw_bytes = source
    else:
        raw_bytes = source.encode("utf-8")

    if not raw_bytes:
        raise ValueError("Empty file. Nothing to encrypt.")

    if not password:
        raise ValueError("Password is required.")

    filename = filename or "encrypted_file"
    original_size = len(raw_bytes)

    # ── Validate algorithm ─────────────────────────
    valid_algorithms = ("AES-256-GCM", "AES-256-CBC", "ChaCha20")
    if algorithm not in valid_algorithms:
        raise ValueError(
            f'Invalid algorithm: "{algorithm}". ' f"Must be one of: {valid_algorithms}"
        )

    # ── Derive key from password ───────────────────
    salt = os.urandom(16)
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=600_000,
        backend=default_backend(),
    )
    key = kdf.derive(password.encode("utf-8"))

    # ── Encrypt ───────────────────────────────────
    tag = b""
    algo_id = 0
    encrypted_data = b""

    if algorithm == "AES-256-GCM":
        algo_id = 1
        iv = os.urandom(12)
        cipher = Cipher(
            algorithms.AES(key),
            modes.GCM(iv),
            backend=default_backend(),
        )
        encryptor = cipher.encryptor()
        encrypted_data = encryptor.update(raw_bytes) + encryptor.finalize()
        tag = encryptor.tag  # 16 bytes

    elif algorithm == "AES-256-CBC":
        algo_id = 2
        iv = os.urandom(16)

        # PKCS7 padding
        padder = padding.PKCS7(128).padder()
        padded_data = padder.update(raw_bytes) + padder.finalize()

        cipher = Cipher(
            algorithms.AES(key),
            modes.CBC(iv),
            backend=default_backend(),
        )
        encryptor = cipher.encryptor()
        encrypted_data = encryptor.update(padded_data) + encryptor.finalize()

    elif algorithm == "ChaCha20":
        algo_id = 3
        iv = os.urandom(16)  # ChaCha20 nonce is 16 bytes
        cipher = Cipher(
            algorithms.ChaCha20(key, iv),
            mode=None,
            backend=default_backend(),
        )
        encryptor = cipher.encryptor()
        encrypted_data = encryptor.update(raw_bytes) + encryptor.finalize()

    # ── Build output file ──────────────────────────
    # Header structure:
    # magic(4) + algo_id(1) + filename_len(1) + filename(N)
    # + salt(16) + iv(len) + tag(16 or 0) + encrypted_data
    fname_bytes = filename.encode("utf-8")[:255]
    fname_len = len(fname_bytes)

    header = (
        b"ENC1"  # magic
        + struct.pack("B", algo_id)  # algorithm ID
        + struct.pack("B", fname_len)  # filename length
        + fname_bytes  # filename
        + salt  # 16 bytes salt
        + iv  # IV/nonce
        + tag  # auth tag (GCM only, else b'')
    )

    output_bytes = header + encrypted_data

    return {
        "encrypted_bytes": output_bytes,
        "algorithm": algorithm,
        "original_size": original_size,
        "encrypted_size": len(output_bytes),
        "original_size_kb": round(original_size / 1024, 2),
        "encrypted_size_kb": round(len(output_bytes) / 1024, 2),
        "filename": filename,
        "output_filename": filename + ".enc",
    }


def decrypt_file(
    source,
    password: str,
) -> dict:
    """
    Decrypt a file encrypted by encrypt_file().

    Args:
        source   : uploaded .enc file object OR raw bytes
        password : decryption password

    Returns:
        {
            'decrypted_bytes' : bytes,
            'algorithm'       : str,
            'original_filename': str,
            'original_size'   : int,
            'encrypted_size'  : int,
        }
    """
    import struct
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes, padding
    from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
    from cryptography.hazmat.backends import default_backend
    from cryptography.exceptions import InvalidTag

    # ── Read source ───────────────────────────────
    if hasattr(source, "read"):
        enc_bytes = source.read()
    elif isinstance(source, bytes):
        enc_bytes = source
    else:
        raise ValueError("Invalid source.")

    if not password:
        raise ValueError("Password is required.")

    encrypted_size = len(enc_bytes)

    # ── Validate magic header ──────────────────────
    if len(enc_bytes) < 4 or enc_bytes[:4] != b"ENC1":
        raise ValueError(
            "Invalid encrypted file. " "File was not encrypted with this tool."
        )

    # ── Parse header ──────────────────────────────
    offset = 4

    algo_id = struct.unpack("B", enc_bytes[offset : offset + 1])[0]
    offset += 1

    fname_len = struct.unpack("B", enc_bytes[offset : offset + 1])[0]
    offset += 1

    filename = enc_bytes[offset : offset + fname_len].decode("utf-8", errors="replace")
    offset += fname_len

    salt = enc_bytes[offset : offset + 16]
    offset += 16

    # ── Algorithm map ─────────────────────────────
    algo_map = {
        1: "AES-256-GCM",
        2: "AES-256-CBC",
        3: "ChaCha20",
    }

    if algo_id not in algo_map:
        raise ValueError(f"Unknown algorithm ID: {algo_id}.")

    algorithm = algo_map[algo_id]

    # ── IV size by algorithm ───────────────────────
    iv_sizes = {
        "AES-256-GCM": 12,
        "AES-256-CBC": 16,
        "ChaCha20": 16,
    }
    iv_size = iv_sizes[algorithm]

    iv = enc_bytes[offset : offset + iv_size]
    offset += iv_size

    # ── Auth tag (GCM only) ────────────────────────
    tag = b""
    if algorithm == "AES-256-GCM":
        tag = enc_bytes[offset : offset + 16]
        offset += 16

    encrypted_data = enc_bytes[offset:]

    # ── Derive key ────────────────────────────────
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=600_000,
        backend=default_backend(),
    )
    key = kdf.derive(password.encode("utf-8"))

    # ── Decrypt ───────────────────────────────────
    try:
        if algorithm == "AES-256-GCM":
            cipher = Cipher(
                algorithms.AES(key),
                modes.GCM(iv, tag),
                backend=default_backend(),
            )
            decryptor = cipher.decryptor()
            decrypted_data = decryptor.update(encrypted_data) + decryptor.finalize()

        elif algorithm == "AES-256-CBC":
            cipher = Cipher(
                algorithms.AES(key),
                modes.CBC(iv),
                backend=default_backend(),
            )
            decryptor = cipher.decryptor()
            padded_data = decryptor.update(encrypted_data) + decryptor.finalize()

            # Remove PKCS7 padding
            unpadder = padding.PKCS7(128).unpadder()
            decrypted_data = unpadder.update(padded_data) + unpadder.finalize()

        elif algorithm == "ChaCha20":
            cipher = Cipher(
                algorithms.ChaCha20(key, iv),
                mode=None,
                backend=default_backend(),
            )
            decryptor = cipher.decryptor()
            decrypted_data = decryptor.update(encrypted_data) + decryptor.finalize()

    except InvalidTag:
        raise ValueError("Wrong password or corrupted file. " "Authentication failed.")
    except Exception as e:
        raise ValueError(f"Decryption failed: {e}")

    return {
        "decrypted_bytes": decrypted_data,
        "algorithm": algorithm,
        "original_filename": filename,
        "original_size": len(decrypted_data),
        "encrypted_size": encrypted_size,
        "original_size_kb": round(len(decrypted_data) / 1024, 2),
        "encrypted_size_kb": round(encrypted_size / 1024, 2),
    }


def get_file_hash(
    source,
    algorithms_list: list = None,
) -> dict:
    """
    Calculate hash(es) of a file.

    Args:
        source          : file object OR bytes OR string
        algorithms_list : list of hash algorithms to compute
                          md5 | sha1 | sha256 | sha512 | sha3_256
                          default: all

    Returns:
        { 'hashes': { 'md5': '...', 'sha256': '...' }, 'size': int }
    """
    import hashlib

    if hasattr(source, "read"):
        data = source.read()
    elif isinstance(source, bytes):
        data = source
    elif isinstance(source, str):
        data = source.encode("utf-8")
    else:
        raise ValueError("Invalid source.")

    if not data:
        raise ValueError("Empty input.")

    supported = {
        "md5": hashlib.md5,
        "sha1": hashlib.sha1,
        "sha256": hashlib.sha256,
        "sha512": hashlib.sha512,
        "sha3_256": hashlib.sha3_256,
        "sha3_512": hashlib.sha3_512,
        "blake2b": hashlib.blake2b,
    }

    if algorithms_list is None:
        algorithms_list = list(supported.keys())

    hashes = {}
    for algo in algorithms_list:
        if algo.lower() in supported:
            h = supported[algo.lower()]()
            h.update(data)
            hashes[algo.lower()] = h.hexdigest()
        else:
            hashes[algo.lower()] = f"unsupported algorithm: {algo}"

    return {
        "hashes": hashes,
        "size": len(data),
        "size_kb": round(len(data) / 1024, 2),
        "size_mb": round(len(data) / (1024 * 1024), 2),
    }


def generate_checksum(
    source,
    algorithm: str = "sha256",
) -> dict:
    """
    Generate a file checksum for integrity verification.

    Args:
        source    : file object OR bytes
        algorithm : hash algorithm to use    (default: sha256)

    Returns:
        { 'checksum', 'algorithm', 'filename', 'size' }
    """
    from .password_service import generate_hash

    result = generate_hash(
        source,
        algorithms_list=[algorithm],
        output_format="hex",
    )

    return {
        "checksum": result["hashes"][algorithm]["hash"],
        "algorithm": algorithm,
        "size_kb": result["input_size_kb"],
        "size": result["input_size"],
    }
