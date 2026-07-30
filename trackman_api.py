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

# Reports reached by activity id use a different endpoint and a different request
# body. Both are taken from the dynamic-reports web client, which does:
#   activityId ? {ActivityId, Altitude, Temperature, BallType} -> /getactivityreport
#              : {ReportId,   Altitude, Temperature, BallType} -> /getreport
TRACKMAN_ACTIVITY_API_URL = "https://golf-player-activities.trackmangolf.com/api/reports/getactivityreport"


# TrackMan report URLs appear in three forms:
#   legacy portal      https://.../reports/<uuid>
#   multi group report https://web-dynamic-reports.trackmangolf.com/?r=<uuid>
#   shot analysis      https://web-dynamic-reports.trackmangolf.com/?a=<uuid>
# `r` is a report id and `a` an activity id. They are separate namespaces and are
# NOT interchangeable, so callers are told which one a URL yielded.
SOURCE_REPORT = "r"
SOURCE_ACTIVITY = "a"

# Strict 8-4-4-4-12 UUID. Deliberately not `[0-9a-fA-F-]{36}`, which also matches
# arbitrary runs of hex and dashes and produced false positives on other params.
_UUID = r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}"

# The `sgos[]=<uuid>` params in dynamic-report URLs are stroke group ids, not report
# ids. Anchoring the parameter match to `?`/`&` followed by exactly `r` or `a`
# keeps them (and every other param) out.
_REPORT_URL_RE = re.compile(
    rf"(?:reports/(?P<legacy>{_UUID})|[?&](?P<param>[ra])=(?P<ident>{_UUID}))"
)


def parse_report_url(url: str):
    """Extract a TrackMan report identifier from a URL.

    Returns a ``(identifier, source)`` tuple where ``source`` is
    :data:`SOURCE_REPORT` ("r") or :data:`SOURCE_ACTIVITY` ("a"), or ``None`` when
    the URL carries no report identifier.
    """
    match = _REPORT_URL_RE.search(url or "")
    if not match:
        return None
    legacy = match.group("legacy")
    if legacy:
        return legacy, SOURCE_REPORT
    return match.group("ident"), match.group("param")


# Shared requests session for connection reuse and pooling across threads
_SESSION = requests.Session()
_RETRY = Retry(total=2, backoff_factor=0.2, status_forcelist=(429, 500, 502, 503, 504), allowed_methods=frozenset(["GET", "POST"]))
_ADAPTER = HTTPAdapter(pool_connections=20, pool_maxsize=50, max_retries=_RETRY)
_SESSION.mount("https://", _ADAPTER)
_SESSION.mount("http://", _ADAPTER)

# Simple in-memory cache of per-report summaries (players + clubs) to avoid
# repeated full downloads. Maps report_id -> (timestamp_seconds, summary_dict)
CLUBS_CACHE_TTL = 60 * 60  # 1 hour

# Persistent cache file stored next to this module
CACHE_FILE = Path(__file__).parent / "clubs_cache.json"

def _load_summary_cache() -> dict:
    try:
        if CACHE_FILE.exists():
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                raw = json.load(f)
            cache = {}
            for k, v in raw.items():
                # Entries written before the player column existed hold a bare
                # club list; skip them so the summary is refetched once.
                if isinstance(v, list) and len(v) == 2 and isinstance(v[1], dict):
                    try:
                        cache[k] = (float(v[0]), v[1])
                    except Exception:
                        continue
            return cache
    except Exception:
        pass
    return {}

def _save_summary_cache() -> None:
    try:
        serial = {k: [v[0], v[1]] for k, v in _SUMMARY_CACHE.items()}
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(serial, f)
    except Exception:
        pass

# load cache at import
_SUMMARY_CACHE: dict = _load_summary_cache()


def clear_clubs_cache() -> None:
    """Clear the in-memory and on-disk report summary cache."""
    try:
        _SUMMARY_CACHE.clear()
    except Exception:
        pass
    try:
        if CACHE_FILE.exists():
            CACHE_FILE.unlink()
    except Exception:
        pass



def build_report_request(report_id: str, source: str = SOURCE_REPORT, detailed: bool = True):
    """Return the ``(url, payload)`` needed to fetch `report_id`.

    `source` selects the identifier namespace: :data:`SOURCE_REPORT` ("r") for a
    report id, :data:`SOURCE_ACTIVITY` ("a") for an activity id. Passing an
    activity id to the report endpoint returns HTTP 404, so the two must not be
    mixed up.
    """
    if source == SOURCE_ACTIVITY:
        return TRACKMAN_ACTIVITY_API_URL, {
            "ActivityId": report_id,
            "Altitude": 0,
            "Temperature": 25,
            "BallType": "Premium",
        }

    payload = {
        "ReportId": report_id,
        "dm": detailed,
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
    return TRACKMAN_API_URL, payload


def download_report(token: str, report_id: str, source: str = SOURCE_REPORT) -> str:
    """Downloads a TrackMan report by ID and saves it as a JSON file."""
    url, payload = build_report_request(report_id, source)

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }

    print(f" Sending request to: {url}")
    response = requests.post(url, headers=headers, json=payload)

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
            parsed = parse_report_url(url)
            if parsed:
                report_id, source = parsed
                print(f"Found recent report ID: {report_id} (source={source})")
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



def get_all_report_ids_from_chrome(limit=200, scan_limit=5000):
    """
    Scans Chrome's history for all TrackMan report URLs and returns a list of
    dicts: [{'id': 'uuid', 'url': '...', 'time': datetime, 'source': 'r'|'a'}, ...]

    `scan_limit` bounds how many trackmangolf.com history rows are examined;
    `limit` bounds how many report URLs are returned. Matching is done in Python
    rather than SQL so that every URL form is handled by one regex.
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
            WHERE url LIKE '%trackmangolf.com%'
            ORDER BY last_visit_time DESC
            LIMIT ?
        """, (scan_limit,))
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
        parsed = parse_report_url(url)
        if not parsed:
            continue
        report_id, source = parsed
        results.append({
            "id": report_id,
            "url": url,
            "time": chrome_time_to_datetime(visit_time),
            "source": source,
        })
        if len(results) >= limit:
            break

    # Log both counts: a large raw count with zero matches means the URL patterns
    # have drifted, which is otherwise invisible from the post-filter count alone.
    log.info(
        "get_all_report_ids_from_chrome: %d trackmangolf.com history row(s) scanned, %d report URL(s) matched",
        len(rows), len(results),
    )

    if not results:
        print(" No TrackMan reports found in Chrome history.")
    else:
        print(f" Found {len(results)} recent TrackMan reports:")
        for r in results:
            print(f" - {r['time']} — {r['id']} ({r['source']})")

    return results

def fetch_report_metadata(token: str, report_id: str, source: str = SOURCE_REPORT) -> dict | None:
    """
    Fetch minimal info for a given report — just enough to get its true creation time.
    """
    url, payload = build_report_request(report_id, source, detailed=False)
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    try:
        resp = _SESSION.post(url, headers=headers, json=payload, timeout=10)
        if resp.status_code != 200:
            return None
        data = resp.json()
        return {
            "id": report_id,
            "source": source,
            "created": data.get("Time") or data.get("Updated"),
            "kind": data.get("Kind"),
        }
    except Exception:
        return None


def _as_id_source_pairs(reports: list) -> list:
    """Normalise a list of report dicts or bare id strings to ``(id, source)`` pairs."""
    pairs = []
    for item in reports:
        if isinstance(item, dict):
            pairs.append((item.get("id"), item.get("source") or SOURCE_REPORT))
        else:
            pairs.append((item, SOURCE_REPORT))
    return pairs


def fetch_report_metadata_batch(token: str, reports: list, max_workers: int = 20) -> list:
    """
    Fetch metadata for multiple reports in parallel.

    Args:
        token: Authorization bearer token
        reports: List of report dicts with 'id' and 'source' keys, or bare id
            strings (treated as report ids)
        max_workers: Number of concurrent requests (default 5, TrackMan API friendly)

    Returns:
        List of metadata dicts with same length as input, None entries for failed requests
    """
    pairs = _as_id_source_pairs(reports)

    # Prefer async aiohttp implementation when available for better throughput.
    if aiohttp is not None:
        async def _fetch_one(session: aiohttp.ClientSession, rid: str, source: str, retries: int = 2):
            url, payload = build_report_request(rid, source, detailed=False)
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
                            "source": source,
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

        async def _fetch_batch_async(id_sources, workers):
            timeout = aiohttp.ClientTimeout(sock_connect=5, sock_read=10)
            connector = aiohttp.TCPConnector(limit=workers)
            async with aiohttp.ClientSession(connector=connector, timeout=timeout) as session:
                tasks = [_fetch_one(session, rid, source) for rid, source in id_sources]
                results = await asyncio.gather(*tasks)
                return results

        # Run the async batch and return results synchronously to callers.
        try:
            return asyncio.run(_fetch_batch_async(pairs, max_workers))
        except Exception:
            # Fall back to threaded requests-based approach on any failure
            pass

    # Fallback: use threaded requests (existing behavior)
    def fetch_single(pair):
        rid, source = pair
        return fetch_report_metadata(token, rid, source)

    results = [None] * len(pairs)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        future_to_index = {executor.submit(fetch_single, pair): idx for idx, pair in enumerate(pairs)}

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


def extract_players_from_report_json(data: dict) -> list:
    """Return a sorted list of unique player names found in a full report JSON."""
    players = set()
    for sg in data.get("StrokeGroups", []) or []:
        name = ((sg.get("Player") or {}).get("Name") or "").strip()
        if name:
            players.add(name)
    return sorted(players)


def fetch_report_summary(token: str, report_id: str, source: str = SOURCE_REPORT) -> dict:
    """Fetch the full report for `report_id` and return its players and clubs.

    This requests the full report, which is much larger than the lightweight
    metadata request, so results are cached. Players and clubs are read from the
    same download because two sessions on the same date are otherwise
    indistinguishable in the report list.

    Returns ``{"players": [...], "clubs": [...]}``; both lists are empty on failure.
    """
    url, payload = build_report_request(report_id, source)
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    # Check cache first
    try:
        cached = _SUMMARY_CACHE.get(report_id)
        if cached:
            ts, summary = cached
            if time.time() - ts < CLUBS_CACHE_TTL:
                return summary
    except Exception:
        pass

    try:
        resp = _SESSION.post(url, headers=headers, json=payload, timeout=15)
        if resp.status_code != 200:
            log.warning("fetch_report_summary: HTTP %s for report %s", resp.status_code, report_id)
            return {"players": [], "clubs": []}
        data = resp.json()
        summary = {
            "players": extract_players_from_report_json(data),
            "clubs": extract_clubs_from_report_json(data),
        }
        try:
            _SUMMARY_CACHE[report_id] = (time.time(), summary)
            _save_summary_cache()
        except Exception:
            pass
        return summary
    except Exception:
        log.exception("fetch_report_summary failed for report %s", report_id)
        return {"players": [], "clubs": []}


def prefer_shot_analysis_per_date(reports: list) -> list:
    """Keep one report per calendar date, preferring shot analysis (activity) ones.

    Every practice session is recorded as a shot analysis, whereas a multi group
    report only exists when someone explicitly created one. Where a date has both,
    they contain the same strokes, so the shot analysis is used and the duplicate
    multi group report is dropped. Dates that only have a multi group report keep
    it, and reports without a usable date are always kept rather than silently
    discarded.

    Input order is preserved. Reports are expected to have 'id', 'time' and
    'source' keys.
    """
    by_date: dict = {}
    keep = set()

    for index, report in enumerate(reports):
        when = report.get("time")
        if when is None:
            keep.add(index)  # no date to group on — never drop it
            continue
        by_date.setdefault(when.date(), []).append(index)

    for date, indexes in by_date.items():
        activities = [i for i in indexes if reports[i].get("source") == SOURCE_ACTIVITY]
        chosen = activities or indexes
        keep.update(chosen)
        dropped = len(indexes) - len(chosen)
        if dropped:
            log.info(
                "prefer_shot_analysis_per_date: %s — kept %d shot analysis report(s), dropped %d duplicate(s)",
                date, len(chosen), dropped,
            )

    return [r for i, r in enumerate(reports) if i in keep]
