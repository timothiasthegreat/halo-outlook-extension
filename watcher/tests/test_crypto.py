"""Tests for Fernet encryption utility."""

import os
import tempfile

import pytest
from cryptography.fernet import Fernet

from watcher.crypto import decrypt_token, encrypt_token, generate_key, get_fernet


class TestCrypto:
    def test_generate_key_returns_valid_fernet_key(self):
        key = generate_key()
        # Fernet.generate_key produces 44-char base64 of 32 bytes
        assert len(key) == 44
        # Verify the key is usable
        f = Fernet(key.encode())
        ct = f.encrypt(b"test")
        assert len(ct) > 0

    def test_encrypt_decrypt_roundtrip(self):
        key = generate_key()
        token = "my-refresh-token-abc-123"

        encrypted = encrypt_token(token, key)
        assert encrypted != token
        assert len(encrypted) > 50  # Fernet tokens are at least ~120 chars base64

        decrypted = decrypt_token(encrypted, key)
        assert decrypted == token

    def test_encrypt_produces_different_ciphertext_each_time(self):
        key = generate_key()
        token = "same-plaintext"

        ct1 = encrypt_token(token, key)
        ct2 = encrypt_token(token, key)
        assert ct1 != ct2  # different IV → different ciphertext

    def test_decrypt_with_wrong_key_raises(self):
        key1 = generate_key()
        key2 = generate_key()

        encrypted = encrypt_token("secret", key1)
        with pytest.raises(Exception):
            decrypt_token(encrypted, key2)

    def test_decrypt_tampered_token_raises(self):
        key = generate_key()
        encrypted = encrypt_token("secret", key)
        tampered = encrypted[:-4] + "AAAA"
        with pytest.raises(Exception):
            decrypt_token(tampered, key)

    def test_get_fernet_from_env(self, monkeypatch):
        key = generate_key()
        monkeypatch.setenv("FERNET_KEY", key)

        f = get_fernet()
        assert isinstance(f, Fernet)
        # Verify it works
        ct = f.encrypt(b"test")
        pt = f.decrypt(ct)
        assert pt == b"test"

    def test_get_fernet_missing_key_raises(self, monkeypatch):
        monkeypatch.delenv("FERNET_KEY", raising=False)
        with pytest.raises(RuntimeError, match="FERNET_KEY"):
            get_fernet()

    def test_encrypt_with_fernet_roundtrip(self):
        """Direct roundtrip using get_fernet/encrypt/decrypt."""
        key = generate_key()
        monkeypatch = pytest.MonkeyPatch()
        monkeypatch.setenv("FERNET_KEY", key)

        try:
            token = "refresh-token-value"
            encrypted = encrypt_token(token, key)
            decrypted = decrypt_token(encrypted, key)
            assert decrypted == token

            # Also test via get_fernet
            f = get_fernet()
            ct = encrypt_token("hello", key)
            pt = f.decrypt(ct.encode()).decode()
            assert pt == "hello"
        finally:
            monkeypatch.undo()