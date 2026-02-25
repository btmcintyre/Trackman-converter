import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import json
import logging
import time
from pathlib import Path
import sqlite3
import re
import shutil
import tempfile
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import asyncio

log = logging.getLogger(__name__)

try:
    import aiohttp
except Exception:
    aiohttp = None


TRACKMAN_API_URL = "https://golf-player-activities.trackmangolf.com/api/reports/getreport"


# Shared requests session for connection reuse and pooling across threads
_SESSION = requests.Session()
_RETRY = Retry(total=2, backoff_factor=0.2, status_forcelist=(429, 500, 502, 503, 504), allowed_methods=frozenset(["GET", "POST"]))
_ADAPTER = HTTPAdapter(pool_connections=20, pool_maxsize=50, max_retries=_RETRY)
_SESSION.mount("https://", _ADAPTER)
_SESSION.mount("http://", _ADAPTER)

# Simple in-memory cache for fetched club lists to avoid repeated full downloads.
# Cache maps report_id -> (timestamp_seconds, clubs_list)
CLUBS_CACHE_TTL = 60 * 60  # 1 hour

# Persistent cache file stored next to this module
CACHE_FILE = Path(__file__).parent / "clubs_cache.json"

def _load_clubs_cache() -> dict:
    try:
        if CACHE_FILE.exists():
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
            cache = {}
            for k, v in raw.items():
                if isinstance(v, list) and len(v) == 2:
                    try:
                        ts = float(v[0])
                        clubs = v[1]
                        cache[k] = (ts, clubs)
                    except Exception:
                        continue
            return cache
    except Exception:
        pass
    return {}

def _save_clubs_cache() -> None:
    try:
        serial = {k: [v[0], v[1]] for k, v in _CLUBS_CACHE.items()}
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(serial, f)
    except Exception:
        pass

# load cache at import
_CLUBS_CACHE: dict = _load_clubs_cache()


def clear_clubs_cache() -> None:
    """Clear the in-memory and on-disk clubs cache."""
    try:
        _CLUBS_CACHE.clear()
    except Exception:
        pass
    try:
        if CACHE_FILE.exists():
            CACHE_FILE.unlink()
    except Exception:
        pass



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
            LIMIT 200
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



def get_all_report_ids_from_chrome(limit=200):
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
        resp = _SESSION.post(url, headers=headers, json=payload, timeout=10)
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
    # Prefer async aiohttp implementation when available for better throughput.
    if aiohttp is not None:
        async def _fetch_one(session: aiohttp.ClientSession, rid: str, retries: int = 2):
            url = "https://golf-player-activities.trackmangolf.com/api/reports/getreport"
            payload = {"ReportId": rid, "dm": False}
            headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
            backoff_base = 0.2
            for attempt in range(retries + 1):
                try:
                    async with session.post(url, json=payload, headers=headers) as resp:
                        if resp.status != 200:
                            if attempt == retries:
                                return None
                            await asyncio.sleep(backoff_base * (2 ** attempt))
                            continue
                        data = await resp.json()
                        return {
                            "id": rid,
                            "created": data.get("Time") or data.get("Updated"),
                            "kind": data.get("Kind"),
                        }
                except asyncio.TimeoutError:
                    if attempt == retries:
                        return None
                    await asyncio.sleep(backoff_base * (2 ** attempt))
                except Exception:
                    if attempt == retries:
                        return None
                    await asyncio.sleep(backoff_base * (2 ** attempt))

        async def _fetch_batch_async(rids, workers):
            timeout = aiohttp.ClientTimeout(sock_connect=5, sock_read=10)
            connector = aiohttp.TCPConnector(limit=workers)
            async with aiohttp.ClientSession(connector=connector, timeout=timeout) as session:
                tasks = [_fetch_one(session, rid) for rid in rids]
                results = await asyncio.gather(*tasks)
                return results

        # Run the async batch and return results synchronously to callers.
        try:
            return asyncio.run(_fetch_batch_async(report_ids, max_workers))
        except Exception:
            # Fall back to threaded requests-based approach on any failure
            pass

    # Fallback: use threaded requests (existing behavior)
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


def extract_clubs_from_report_json(data: dict) -> list:
    """Return a sorted list of unique club names found in a full report JSON."""
    clubs = set()
    # StrokeGroups may contain Strokes with Club or Measurement entries with Club
    for sg in data.get("StrokeGroups", []):
        # group-level Club field
        gclub = sg.get("Club")
        if gclub:
            clubs.add(gclub)

        # strokes
        for s in sg.get("Strokes", []) or []:
            # some reports use a nested Measurement with Club
            m = s.get("Measurement") or {}
            c = s.get("Club") or m.get("Club")
            if c:
                clubs.add(c)

    return sorted([c for c in clubs if c])


def fetch_report_clubs(token: str, report_id: str) -> list:
    """Fetch the full report for `report_id` and return the club list.

    This requests the full report (dm=True) which is larger than the
    lightweight metadata request. Use sparingly or cache results.
    """
    url = "https://golf-player-activities.trackmangolf.com/api/reports/getreport"
    payload = {"ReportId": report_id, "dm": True}
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # Check cache first
    try:
        import time
        cached = _CLUBS_CACHE.get(report_id)
        if cached:
            ts, clubs = cached
            if time.time() - ts < CLUBS_CACHE_TTL:
                return clubs
    except Exception:
        pass

    try:
        resp = _SESSION.post(url, headers=headers, json=payload, timeout=15)
        if resp.status_code != 200:
            log.warning("fetch_report_clubs: HTTP %s for report %s", resp.status_code, report_id)
            return []
        data = resp.json()
        clubs = extract_clubs_from_report_json(data)
        try:
            _CLUBS_CACHE[report_id] = (time.time(), clubs)
            _save_clubs_cache()
        except Exception:
            pass
        return clubs
    except Exception:
        log.exception("fetch_report_clubs failed for report %s", report_id)
        return []
