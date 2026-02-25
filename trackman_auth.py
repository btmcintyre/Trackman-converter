import base64
import ctypes
import ctypes.wintypes
import logging
import os
import sqlite3
import json
import shutil
from pathlib import Path

import sys

log = logging.getLogger(__name__)


BASE_DIR = Path(getattr(sys, '_MEIPASS', Path(__file__).parent))

# Store the token in a stable user-data directory that never moves regardless
# of where the executable is placed.  In development (not frozen) fall back to
# the repo root so the existing dev workflow is unchanged.
if getattr(sys, 'frozen', False):
    _TOKEN_DIR = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")) / "TrackmanConverter"
    _TOKEN_DIR.mkdir(parents=True, exist_ok=True)
    TOKEN_FILE = _TOKEN_DIR / "trackman_token.txt"
else:
    TOKEN_FILE = BASE_DIR / "trackman_token.txt"


def _is_valid_token(token: str) -> bool:
    """Return True if the token string looks safe to use as an HTTP header value."""
    if not token:
        return False
    # HTTP header values must not contain control characters (< 0x20) or DEL (0x7f)
    # Allow printable ASCII + common extended chars used in JWT/bearer tokens
    return all(0x20 <= ord(c) < 0x7f or ord(c) > 0x7f for c in token)


def get_saved_token():
    """Return saved token if it exists.

    Also migrates a token from the old location (next to the exe) to the new
    stable %APPDATA% location so existing users don't need to re-authenticate.
    """
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r", encoding="utf-8") as f:
            token = f.read().strip()
        if _is_valid_token(token):
            return token
        # Token is corrupted — delete it so we re-extract on next call
        try:
            os.remove(TOKEN_FILE)
        except Exception:
            pass
        return None

    # One-time migration: look for a token next to the running exe (old location)
    if getattr(sys, 'frozen', False):
        old_location = Path(sys.executable).parent / "trackman_token.txt"
        if old_location.exists() and old_location != TOKEN_FILE:
            try:
                token = old_location.read_text(encoding="utf-8").strip()
                if token:
                    save_token(token)  # write to new stable location
                    old_location.unlink(missing_ok=True)  # remove old file
                    return token
            except Exception:
                pass

    return None

def save_token(token):
    """Save token for reuse"""
    with open(TOKEN_FILE, "w", encoding="utf-8") as f:
        f.write(token)

def _dpapi_decrypt(ciphertext: bytes) -> bytes | None:
    """Decrypt bytes using Windows DPAPI (CryptUnprotectData)."""
    class DATA_BLOB(ctypes.Structure):
        _fields_ = [("cbData", ctypes.wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_char))]

    buf = ctypes.create_string_buffer(ciphertext, len(ciphertext))
    blobin = DATA_BLOB(ctypes.sizeof(buf), buf)
    blobout = DATA_BLOB()
    ok = ctypes.windll.crypt32.CryptUnprotectData(
        ctypes.byref(blobin), None, None, None, None, 0, ctypes.byref(blobout)
    )
    if not ok:
        return None
    result = ctypes.string_at(blobout.pbData, blobout.cbData)
    ctypes.windll.kernel32.LocalFree(blobout.pbData)
    return result


def _get_chrome_master_key() -> bytes | None:
    """Read and decrypt Chrome's AES master key from Local State."""
    local_state_path = (
        Path(os.getenv("LOCALAPPDATA", ""))
        / "Google" / "Chrome" / "User Data" / "Local State"
    )
    try:
        with open(local_state_path, "r", encoding="utf-8") as f:
            local_state = json.load(f)
        enc_key_b64 = local_state["os_crypt"]["encrypted_key"]
    except Exception:
        log.exception("_get_chrome_master_key: failed to read Local State from %s", local_state_path)
        return None

    try:
        # The key is base64-encoded with a "DPAPI" prefix (5 bytes) prepended.
        enc_key = base64.b64decode(enc_key_b64)[5:]
        key = _dpapi_decrypt(enc_key)
        if key is None:
            log.error("_get_chrome_master_key: DPAPI decrypt returned None")
        else:
            log.debug("_get_chrome_master_key: master key obtained, len=%d", len(key))
        return key
    except Exception:
        log.exception("_get_chrome_master_key: failed to decrypt master key")
        return None


def _decrypt_chrome_cookie(encrypted_value: bytes) -> str | None:
    """Decrypt a Chrome cookie encrypted_value field.

    Chrome on Windows uses AES-256-GCM for cookies with v10/v20 prefix,
    with the key stored in Local State protected by DPAPI.
    Older cookies may be protected directly by DPAPI.
    """
    if not encrypted_value:
        log.warning("_decrypt_chrome_cookie: empty encrypted_value")
        return None

    raw = bytes(encrypted_value)
    prefix = raw[:3]
    log.debug("_decrypt_chrome_cookie: prefix=%r, total_len=%d", prefix, len(raw))

    # Modern Chrome: v10/v20 prefix + 12-byte nonce + ciphertext+tag
    if prefix in (b"v10", b"v20"):
        master_key = _get_chrome_master_key()
        if master_key is None:
            log.error("_decrypt_chrome_cookie: could not obtain Chrome master key")
            return None
        try:
            from cryptography.hazmat.primitives.ciphers.aead import AESGCM
            nonce = raw[3:15]
            ciphertext = raw[15:]
            decrypted = AESGCM(master_key).decrypt(nonce, ciphertext, None)
            result = decrypted.decode("utf-8", errors="ignore")
            log.debug("_decrypt_chrome_cookie: AES-GCM decryption succeeded, token_len=%d", len(result))
            return result
        except Exception:
            log.exception("_decrypt_chrome_cookie: AES-GCM decryption failed")
            return None

    # Legacy Chrome: raw DPAPI blob
    try:
        result = _dpapi_decrypt(raw)
        if result:
            decoded = result.decode("utf-8", errors="ignore")
            log.debug("_decrypt_chrome_cookie: DPAPI decryption succeeded, token_len=%d", len(decoded))
            return decoded
    except Exception:
        log.exception("_decrypt_chrome_cookie: DPAPI decryption failed")

    log.error("_decrypt_chrome_cookie: all decryption methods failed for prefix=%r", prefix)
    return None


def get_chrome_cookie_path():
    """Locate Chrome's cookie file"""
    local_app_data = os.getenv("LOCALAPPDATA")
    cookie_path = Path(local_app_data) / "Google/Chrome/User Data/Default/Network/Cookies"
    if cookie_path.exists():
        return cookie_path
    return None

def extract_token_from_chrome():
    """
    Extract TrackMan token from Chrome cookies manually (even if Chrome is running).
    """
    cookie_db_path = get_chrome_cookie_path()
    if not cookie_db_path:
        log.error("extract_token_from_chrome: Chrome cookie DB not found")
        return None

    log.info("extract_token_from_chrome: cookie DB path=%s", cookie_db_path)

    tmp_copy = Path(os.environ.get("TEMP", os.environ.get("TMP", "."))) / "chrome_cookies_copy.db"

    try:
        with open(cookie_db_path, "rb") as src, open(tmp_copy, "wb") as dst:
            dst.write(src.read())
        log.info("extract_token_from_chrome: copied cookie DB to %s", tmp_copy)
    except PermissionError:
        log.error("extract_token_from_chrome: PermissionError copying cookie DB — Chrome may be running")
        return None
    except Exception:
        log.exception("extract_token_from_chrome: failed to copy cookie DB")
        return None

    token = None
    conn = None
    try:
        conn = sqlite3.connect(tmp_copy)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT name, encrypted_value FROM cookies WHERE host_key LIKE '%trackmangolf.com%'"
        )
        rows = cursor.fetchall()
        log.info("extract_token_from_chrome: found %d trackmangolf cookie row(s)", len(rows))
        for name, value in rows:
            log.debug("extract_token_from_chrome: cookie name=%r, encrypted_len=%d", name, len(bytes(value)) if value else 0)
            if name.lower() == "appsession":
                try:
                    token = _decrypt_chrome_cookie(value)
                    if token:
                        log.info("extract_token_from_chrome: token decrypted successfully, len=%d", len(token))
                        break
                    else:
                        log.warning("extract_token_from_chrome: decryption returned None for appsession cookie")
                except Exception:
                    log.exception("extract_token_from_chrome: exception decrypting appsession")
        if token is None:
            log.warning("extract_token_from_chrome: no appsession cookie found or all decryptions failed")
    except Exception:
        log.exception("extract_token_from_chrome: error reading cookie DB")
    finally:
        if conn:
            conn.close()
        try:
            if tmp_copy.exists():
                tmp_copy.unlink()
        except Exception:
            log.warning("extract_token_from_chrome: could not delete temp copy")

    return token


def login_via_browser():
    """Try to reuse login from browser, fallback to manual entry"""
    print("Checking Chrome cookies for TrackMan login...")
    token = extract_token_from_chrome()
    if token:
        print("Found token from Chrome session!")
        save_token(token)
        return token

    # In a frozen GUI app there is no console — return None so the GUI layer
    # can show a proper token-paste dialog instead of crashing.
    if getattr(sys, 'frozen', False):
        log.warning("login_via_browser: could not extract token from Chrome in frozen build, returning None")
        return None

    print("Could not auto-detect login. Please paste manually.")
    manual = input("Paste your TrackMan Bearer token: ").strip()
    if manual:
        save_token(manual)
        return manual
    return None
