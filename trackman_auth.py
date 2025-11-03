import os
import sqlite3
import json
import shutil
from pathlib import Path

TOKEN_FILE = "trackman_token.txt"

def get_saved_token():
    """Return saved token if it exists"""
    if os.path.exists(TOKEN_FILE):
        with open(TOKEN_FILE, "r") as f:
            return f.read().strip()
    return None

def save_token(token):
    """Save token for reuse"""
    with open(TOKEN_FILE, "w") as f:
        f.write(token)

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
        print("Could not find Chrome cookie database.")
        return None

    tmp_copy = Path("chrome_cookies_copy.db")

    try:
        # Copy cookie DB so we don't lock the original file
        with open(cookie_db_path, "rb") as src, open(tmp_copy, "wb") as dst:
            dst.write(src.read())
    except PermissionError:
        print("Chrome is still using the cookie file. Please close Chrome completely and retry.")
        return None

    token = None
    conn = None
    try:
        conn = sqlite3.connect(tmp_copy)
        cursor = conn.cursor()
        cursor.execute(
            "SELECT name, encrypted_value FROM cookies WHERE host_key LIKE '%trackmangolf.com%'"
        )
        for name, value in cursor.fetchall():
            if name.lower() == "appsession":
                try:
                    token = value.decode("utf-8", errors="ignore")
                    if token:
                        break
                except Exception:
                    pass
    except Exception as e:
        print("Error reading Chrome cookies:", e)
    finally:
        if conn:
            conn.close()
        try:
            if tmp_copy.exists():
                tmp_copy.unlink()
        except Exception as cleanup_error:
            print("Couldn't delete temp copy:", cleanup_error)

    return token


def login_via_browser():
    """Try to reuse login from browser, fallback to manual entry"""
    print("Checking Chrome cookies for TrackMan login...")
    token = extract_token_from_chrome()
    if token:
        print("Found token from Chrome session!")
        save_token(token)
        return token

    print("Could not auto-detect login. Please paste manually.")
    manual = input("Paste your TrackMan Bearer token: ").strip()
    if manual:
        save_token(manual)
        return manual
    return None
