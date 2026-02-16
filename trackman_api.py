import requests
import json
from pathlib import Path
import sqlite3
import re
import shutil
import tempfile
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed


TRACKMAN_API_URL = "https://golf-player-activities.trackmangolf.com/api/reports/getreport"



def download_report(token: str, report_id: str) -> str:
    """Downloads a TrackMan report by ID and saves it as a JSON file."""
    payload = {
        "ReportId": report_id,
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
        "nd_humidity": 50,
    }

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    print(f" Sending request to: {TRACKMAN_API_URL}")
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
    Reads Chrome history for the latest TrackMan report URL
    and extracts the report ID (?r=... or /reports/<uuid>).
    """
    print("Searching Chrome history for latest TrackMan report...")

    history_path = Path.home() / "AppData/Local/Google/Chrome/User Data/Default/History"
    temp_copy = Path(tempfile.gettempdir()) / "chrome_history_copy.db"

    try:
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
            match = re.search(r"(?:reports/|[?&]r=)([0-9a-fA-F-]{36})", url)
            if match:
                report_id = match.group(1)
                print(f"Found recent report ID: {report_id}")
                return report_id

        print("No valid TrackMan report URL found in recent Chrome history.")
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



def get_all_report_ids_from_chrome(limit=15):
    """
    Scans Chrome's history for all TrackMan report URLs
    and returns a list of dicts: [{'id': 'uuid', 'url': '...', 'time': datetime}, ...]
    """
    chrome_path = Path.home() / "AppData/Local/Google/Chrome/User Data/Default/History"
    if not chrome_path.exists():
        raise Exception("Chrome history not found. Make sure Chrome is installed and used.")

    tmp_copy = Path(tempfile.gettempdir()) / "chrome_history_copy.db"
    try:
        shutil.copyfile(chrome_path, tmp_copy)
    except Exception as e:
        raise Exception(f"Failed to copy Chrome history — close Chrome and retry.\n{e}")

    try:
        conn = sqlite3.connect(tmp_copy)
        cursor = conn.cursor()
        cursor.execute("""
            SELECT url, last_visit_time
            FROM urls
            WHERE url LIKE '%trackmangolf.com/reports/%' OR url LIKE '%trackmangolf.com%r=%'
            ORDER BY last_visit_time DESC
            LIMIT ?
        """, (limit,))
        rows = cursor.fetchall()
        conn.close()
    except Exception as e:
        raise Exception(f"Error reading Chrome history: {e}")
    finally:
        try:
            tmp_copy.unlink(missing_ok=True)
        except:
            pass

    def chrome_time_to_datetime(chrome_time):

        return datetime(1601, 1, 1) + timedelta(microseconds=chrome_time)

    results = []
    for url, visit_time in rows:
        match = re.search(r"(?:reports/|[?&]r=)([0-9a-fA-F-]{36})", url)
        if match:
            report_id = match.group(1)
            results.append({
                "id": report_id,
                "url": url,
                "time": chrome_time_to_datetime(visit_time)
            })

    if not results:
        print(" No TrackMan reports found in Chrome history.")
    else:
        print(f" Found {len(results)} recent TrackMan reports:")
        for r in results:
            print(f" - {r['time']} — {r['id']}")

    return results

def fetch_report_metadata(token: str, report_id: str) -> dict | None:
    """
    Fetch minimal info for a given report — just enough to get its true creation time.
    """
    url = "https://golf-player-activities.trackmangolf.com/api/reports/getreport"
    payload = {"ReportId": report_id, "dm": False}
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    try:
        resp = requests.post(url, headers=headers, json=payload, timeout=10)
        if resp.status_code != 200:
            return None
        data = resp.json()
        return {
            "id": report_id,
            "created": data.get("Time") or data.get("Updated"),
            "kind": data.get("Kind"),
        }
    except Exception:
        return None


def fetch_report_metadata_batch(token: str, report_ids: list, max_workers: int = 20) -> list:
    """
    Fetch metadata for multiple reports in parallel.
    
    Args:
        token: Authorization bearer token
        report_ids: List of report IDs to fetch
        max_workers: Number of concurrent requests (default 5, TrackMan API friendly)
    
    Returns:
        List of metadata dicts with same length as input, None entries for failed requests
    """
    def fetch_single(report_id):
        return fetch_report_metadata(token, report_id)
    
    results = [None] * len(report_ids)
    
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_index = {executor.submit(fetch_single, rid): idx for idx, rid in enumerate(report_ids)}
        
        for future in as_completed(future_to_index):
            idx = future_to_index[future]
            try:
                results[idx] = future.result()
            except Exception:
                results[idx] = None
    
    return results
