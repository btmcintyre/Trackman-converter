import requests
import json
from pathlib import Path
import sqlite3
import re
import shutil
import browser_cookie3


# ✅ Main TrackMan API endpoint
TRACKMAN_API_URL = "https://golf-player-activities.trackmangolf.com/api/reports/getreport"


def download_report(token: str, report_id: str) -> str:
    """Downloads a TrackMan report by ID and saves it as a JSON file."""
    payload = {
        "ReportId": report_id,  # Capital R
        "dm": True,
        "nd": True,
        "nd_ballType": "Premium",
        "nd_altitude": 0,
        "nd_temperature": 25,
        "nd_temperatureUnit": "Celsius",
        "lop": True,
        "sro": False,
        "do": True,
        "nd_pressure": 1013,
        "nd_wind": 0,
        "nd_humidity": 50
    }

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    print(f"📡 Sending request to: {TRACKMAN_API_URL}")
    response = requests.post(TRACKMAN_API_URL, headers=headers, json=payload)

    print("Status:", response.status_code)
    if response.status_code == 200:
        data = response.json()
        keys = list(data.keys())
        print(f"Success — keys: {keys[:10]}")

        out_path = Path("trackman_full_report.json")
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        print(f"Saved as {out_path}")
        return str(out_path)
    else:
        raise Exception(f"Error {response.status_code}: {response.text}")


def get_latest_report_id_from_chrome() -> str:
    """
    Reads Chrome history for the latest TrackMan Multi Group Report URL
    and extracts the report ID (?r=...).
    """
    print("Searching Chrome history for latest TrackMan report...")

    history_path = Path.home() / "AppData/Local/Google/Chrome/User Data/Default/History"
    temp_copy = Path("chrome_history_copy.db")

    try:
        # Copy history so we can read it while Chrome is open
        shutil.copyfile(history_path, temp_copy)

        conn = sqlite3.connect(temp_copy)
        cursor = conn.cursor()
        cursor.execute(
            """
            SELECT url, last_visit_time FROM urls
            WHERE url LIKE '%trackmangolf.com%'
            ORDER BY last_visit_time DESC
            LIMIT 50
            """
        )
        rows = cursor.fetchall()
        conn.close()

        for url, _ in rows:
            match = re.search(r"[?&]r=([0-9a-fA-F-]{36})", url)
            if match:
                report_id = match.group(1)
                print(f"Found recent report ID: {report_id}")
                return report_id

        print("⚠️ No valid TrackMan report URL found in recent Chrome history.")
        return None

    except Exception as e:
        print(f"Error reading Chrome history: {e}")
        return None

    finally:
        try:
            if temp_copy.exists():
                temp_copy.unlink()
        except:
            pass
