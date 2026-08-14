"""Fernet encryption utility for the Halo Outlook Watcher.

Provides token encryption/decryption compatible with the Express server's
Node.js fernet.js implementation. Both use AES-128-CBC with HMAC-SHA256
per the Fernet specification.

Usage:
    from watcher.crypto import get_fernet, encrypt_token, decrypt_token

    key = generate_key()  # or read from FERNET_KEY env var
    encrypted = encrypt_token("my-refresh-token", key)
    decrypted = decrypt_token(encrypted, key)
"""

from __future__ import annotations

import os

from cryptography.fernet import Fernet


def generate_key() -> str:
    """Generate a new Fernet key (44-char base64, 32 bytes)."""
    return Fernet.generate_key().decode()


def get_fernet() -> Fernet:
    """Get a Fernet instance from the FERNET_KEY environment variable.

    Raises RuntimeError if FERNET_KEY is not set.
    """
    key = os.environ.get("FERNET_KEY")
    if not key:
        raise RuntimeError("FERNET_KEY environment variable is required")
    return Fernet(key.encode() if isinstance(key, str) else key)


def encrypt_token(token: str, key: str) -> str:
    """Encrypt a refresh token for storage at rest.

    Returns base64 Fernet token string.
    """
    f = Fernet(key.encode() if isinstance(key, str) else key)
    return f.encrypt(token.encode()).decode()


def decrypt_token(encrypted: str, key: str) -> str:
    """Decrypt a Fernet-encrypted token.

    Returns the original plaintext token string.
    Raises cryptography.fernet.InvalidToken on wrong key or tampered data.
    """
    f = Fernet(key.encode() if isinstance(key, str) else key)
    return f.decrypt(encrypted.encode()).decode()