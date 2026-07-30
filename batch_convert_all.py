"""batch_convert_all.py
Headless script that discovers every TrackMan session visible in Chrome
history, downloads each one, converts it to a dated Excel session file,
and appends the data into the per-club master workbook.

Already-converted sessions whose session file already exists in the output
directory are skipped by default (use --force to re-convert them).

Usage
-----
    python batch_convert_all.py [--force] [--out-dir PATH] [--limit N]

Options
    --force       Re-convert even if the session xlsx already exists.
    --out-dir     Where to write session files and the master workbook.
                  Defaults to C:\\Trackman\\Data
    --limit       Maximum number of reports to process (newest first).
                  Defaults to 200.
"""

import argparse
import json
import logging
import sys
import time
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# Logging — write to console so progress is visible when run from a terminal
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)


def parse_args():
    p = argparse.ArgumentParser(description="Batch-convert all TrackMan sessions to Excel.")
    p.add_argument("--force", action="store_true", help="Re-convert sessions that already have an xlsx file.")
    p.add_argument("--out-dir", default=r"C:\Trackman\Data", help="Output directory for xlsx files.")
    p.add_argument("--limit", type=int, default=200, help="Max number of reports to process (newest first).")
    return p.parse_args()


def main():
    args = parse_args()
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # -----------------------------------------------------------------------
    # Imports — done here so errors surface cleanly before any work starts
    # -----------------------------------------------------------------------
    import trackman_auth
    import trackman_api
    from trackman_api import (
        download_report,
        get_all_report_ids_from_chrome,
        fetch_report_metadata_batch,
        extract_clubs_from_report_json,
        prefer_shot_analysis_per_date,
        SOURCE_ACTIVITY,
        SOURCE_REPORT,
    )
    from converter import build_workbook_per_club, append_to_master_workbook, split_by_player, _safe_filename_part

    # -----------------------------------------------------------------------
    # Authenticate
    # -----------------------------------------------------------------------
    log.info("Obtaining authentication token…")
    token = trackman_auth.get_saved_token()
    if not token:
        log.error(
            "No saved token found.  Run the GUI app once and save your "
            "token, then retry this script."
        )
        sys.exit(1)
    log.info("Token obtained.")

    # -----------------------------------------------------------------------
    # Discover report IDs from Chrome history
    # -----------------------------------------------------------------------
    log.info("Scanning Chrome history for TrackMan report URLs…")
    raw_reports = get_all_report_ids_from_chrome(limit=args.limit)
    if not raw_reports:
        log.error("No reports found in Chrome history.  Open some reports in Chrome and retry.")
        sys.exit(1)

    # Deduplicate
    seen: set[str] = set()
    unique_reports: list[dict] = []
    for r in raw_reports:
        rid = r.get("id")
        if rid and rid not in seen:
            seen.add(rid)
            unique_reports.append(r)
    log.info(f"Found {len(unique_reports)} unique report(s) in Chrome history.")

    # -----------------------------------------------------------------------
    # Fetch metadata (upload dates) in parallel
    # -----------------------------------------------------------------------
    log.info("Fetching report metadata (upload dates)…")
    metadata_list = fetch_report_metadata_batch(token, unique_reports, max_workers=5)

    enriched: list[dict] = []
    for r, meta in zip(unique_reports, metadata_list):
        if meta and meta.get("created"):
            try:
                created = datetime.fromisoformat(meta["created"].replace("Z", "+00:00"))
                enriched.append({"id": r["id"], "source": r.get("source", SOURCE_REPORT), "time": created})
            except Exception:
                pass

    if not enriched:
        log.error("Could not retrieve metadata for any report.  Check your token.")
        sys.exit(1)

    enriched = prefer_shot_analysis_per_date(enriched)

    # Sort newest-first so the master workbook gets rows in chronological order
    # when processed oldest-to-newest below.
    enriched.sort(key=lambda x: x["time"])
    log.info(f"{len(enriched)} report(s) with valid metadata, processing oldest → newest.")

    # -----------------------------------------------------------------------
    # Process each report
    # -----------------------------------------------------------------------
    success = 0
    skipped = 0
    failed = 0

    for i, report in enumerate(enriched, start=1):
        date_str = report["time"].strftime("%Y_%m_%d")
        report_id = report["id"]
        report_source = report.get("source", SOURCE_REPORT)
        kind = "shot analysis" if report_source == SOURCE_ACTIVITY else "multi group report"
        log.info(f"[{i}/{len(enriched)}] {kind} {report_id}  ({date_str})")

        # Skip if any session file for this date already exists (covers all players).
        # Master workbooks are named Trackman_Master_*.xlsx so they won't match the
        # date_YYYY_MM_DD prefix and don't need to be excluded explicitly.
        existing = list(out_dir.glob(f"{date_str}_*.xlsx"))

        if existing and not args.force:
            log.info(f"  Skipping — session file already exists: {existing[0].name}")
            skipped += 1
            continue

        # Download
        try:
            log.info(f"  Downloading…")
            json_path = download_report(token, report_id, report_source)
        except Exception as e:
            log.warning(f"  Download failed: {e}")
            failed += 1
            # Brief pause before the next request so we don't hammer the API
            time.sleep(1)
            continue

        # Read the downloaded JSON
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception as e:
            log.warning(f"  Could not read JSON: {e}")
            failed += 1
            continue

        # Split by player and produce one session file + master per player
        players = split_by_player(data)
        report_ok = True

        for player_name, player_data in players.items():
            safe_player = _safe_filename_part(player_name)
            try:
                clubs = extract_clubs_from_report_json(player_data)
                clubs_str = "_".join(clubs) if clubs else ""
            except Exception:
                clubs_str = ""

            session_name = (
                f"{date_str}_{safe_player}_{clubs_str}.xlsx" if clubs_str
                else f"{date_str}_{safe_player}.xlsx"
            )
            out_path = out_dir / session_name

            # Convert to session xlsx
            try:
                log.info(f"  [{player_name}] Converting → {session_name}")
                wb = build_workbook_per_club(player_data)
                wb.save(str(out_path))
            except Exception as e:
                log.warning(f"  [{player_name}] Conversion failed: {e}")
                report_ok = False
                continue

            # Append to per-player master
            player_master = out_dir / f"Trackman_Master_{safe_player}.xlsx"
            try:
                log.info(f"  [{player_name}] Updating master → {player_master.name}")
                append_to_master_workbook(player_data, player_master)
            except Exception as e:
                log.warning(f"  [{player_name}] Master update failed: {e}")

        if report_ok:
            success += 1
        else:
            failed += 1
            continue
        # Small pause to be polite to the API between downloads
        time.sleep(0.5)

    # -----------------------------------------------------------------------
    # Summary
    # -----------------------------------------------------------------------
    log.info("─" * 60)
    log.info(f"Done.  Converted: {success}  Skipped: {skipped}  Failed: {failed}")
    masters = sorted(out_dir.glob("Trackman_Master_*.xlsx"))
    for m in masters:
        log.info(f"Master workbook: {m}")
    log.info(f"Session files:   {out_dir}")


if __name__ == "__main__":
    main()
