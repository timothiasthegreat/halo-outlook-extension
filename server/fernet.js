"use strict";

/**
 * Minimal Fernet-compatible encryption for Node.js.
 *
 * Implements the Fernet specification (AES-128-CBC with HMAC-SHA256)
 * using Node's built-in crypto module. Compatible with Python's
 * cryptography.fernet so tokens encrypted here can be decrypted
 * by the watcher and vice versa.
 *
 * Usage:
 *   const fernet = createFernet(process.env.FERNET_KEY);
 *   const ciphertext = fernet.encrypt("plaintext");
 *   const plaintext = fernet.decrypt(ciphertext);
 */

const crypto = require("crypto");

const FERNET_VERSION = 0x80;

function createFernet(keyB64) {
  if (!keyB64) throw new Error("FERNET_KEY is required");

  const key = Buffer.from(keyB64, "base64");
  if (key.length !== 32) throw new Error("Fernet key must be 32 bytes (base64-encoded)");

  const signingKey = key.subarray(0, 16);    // first 128 bits
  const encryptionKey = key.subarray(16, 32); // last 128 bits

  function encrypt(plaintext) {
    const iv = crypto.randomBytes(16);
    const now = Math.floor(Date.now() / 1000);

    const timestamp = Buffer.alloc(8);
    timestamp.writeBigInt64BE(BigInt(now));

    // Pad the plaintext to AES block size (PKCS7)
    const plainBuf = Buffer.from(plaintext, "utf8");
    const padLen = 16 - (plainBuf.length % 16);
    const padByte = padLen;
    const padding = Buffer.alloc(padLen, padByte);
    const padded = Buffer.concat([plainBuf, padding]);

    // Encrypt
    const cipher = crypto.createCipheriv("aes-128-cbc", encryptionKey, iv);
    cipher.setAutoPadding(false);
    const ciphertext = Buffer.concat([cipher.update(padded), cipher.final()]);

    // Assemble: version (1) || timestamp (8) || iv (16) || ciphertext
    const payload = Buffer.concat([Buffer.from([FERNET_VERSION]), timestamp, iv, ciphertext]);

    // HMAC
    const hmac = crypto.createHmac("sha256", signingKey);
    hmac.update(payload);
    const signature = hmac.digest();

    const token = Buffer.concat([payload, signature]);
    return token.toString("base64url");
  }

  function decrypt(tokenB64) {
    const raw = Buffer.from(tokenB64, "base64url");
    if (raw.length < 73) throw new Error("Token too short");

    const payload = raw.subarray(0, raw.length - 32);
    const signature = raw.subarray(raw.length - 32);

    // Verify HMAC
    const hmac = crypto.createHmac("sha256", signingKey);
    hmac.update(payload);
    const expected = hmac.digest();
    if (!crypto.timingSafeEqual(signature, expected)) {
      throw new Error("Invalid signature — token may be tampered");
    }

    // Parse payload
    if (payload[0] !== FERNET_VERSION) throw new Error("Unsupported Fernet version");
    const iv = payload.subarray(9, 25);
    const ciphertext = payload.subarray(25);

    const decipher = crypto.createDecipheriv("aes-128-cbc", encryptionKey, iv);
    decipher.setAutoPadding(false);
    const padded = Buffer.concat([decipher.update(ciphertext), decipher.final()]);

    // Remove PKCS7 padding
    const padLen = padded[padded.length - 1];
    if (padLen < 1 || padLen > 16) throw new Error("Invalid padding");
    const plaintext = padded.subarray(0, padded.length - padLen);
    return plaintext.toString("utf8");
  }

  return { encrypt, decrypt };
}

module.exports = { createFernet };