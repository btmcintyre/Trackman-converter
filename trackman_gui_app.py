# TrackMan Report Converter GUI Application
# This module provides the main GUI interface for downloading and converting TrackMan reports to Excel format.

# UI Framework and dialogs
import customtkinter as ctk
from tkinter import messagebox, filedialog

# Data handling
import json
import logging
import os
import sys
from pathlib import Path
from datetime import datetime
import pandas as pd
import threading

# Set up file logging so errors are visible even in console=False frozen builds
_LOG_DIR = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")) / "TrackmanConverter"
_LOG_DIR.mkdir(parents=True, exist_ok=True)
logging.basicConfig(
    filename=str(_LOG_DIR / "trackman.log"),
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger(__name__)
log.info("TrackmanConverter starting up")

# In a frozen PyInstaller build the certifi CA bundle is extracted alongside
# the exe. Point requests and aiohttp at it via env vars so HTTPS works.
if getattr(sys, 'frozen', False):
    # In a one-dir build, bundled datas land next to the exe (sys.executable parent)
    _ca = os.path.join(os.path.dirname(sys.executable), 'certifi', 'cacert.pem')
    if not os.path.exists(_ca):
        # Fallback: one-file build extracts to sys._MEIPASS
        _ca = os.path.join(getattr(sys, '_MEIPASS', ''), 'certifi', 'cacert.pem')
    os.environ.setdefault('SSL_CERT_FILE', _ca)
    os.environ.setdefault('REQUESTS_CA_BUNDLE', _ca)
    log.info(f"SSL_CERT_FILE set to: {_ca} (exists={os.path.exists(_ca)})")

# Excel workbook creation and formatting
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

# Project-specific modules
import trackman_auth
import trackman_api
from trackman_api import download_report, get_latest_report_id_from_chrome, fetch_report_clubs
from converter import build_workbook_per_club, append_to_master_workbook, split_by_player, _safe_filename_part

# Application-wide constants and theme configuration
APP_FOOTER_TEXT = "© 2026 TrackMan Converter by Tom McIntyre and Brian McIntyre. All rights reserved."
TRACKMAN_COLOUR = "#2563EB"   # Modern blue – primary action colour
HEADER_BG    = "#1A1F3A"   # Deep navy – header background
DARK_BG      = "#F4F6F9"   # Off-white – main window background
SUBTLE_BTN   = "#2D3561"   # Muted navy – secondary/utility buttons in header
SUBTLE_HOVER = "#3D4571"   # Lighter navy – hover for subtle buttons
ACCENT_HOVER = "#1D4ED8"   # Darker blue – hover for primary action buttons
TEXT_PRIMARY = "#1A1F3A"   # Dark navy – headings and bold labels
TEXT_BODY    = "#374151"   # Dark grey – regular body text
TEXT_MUTED   = "#6B7280"   # Medium grey – secondary / loading text

# NOTE: The modal overlay window was removed. A small inline spinner
# is displayed inside the main window during long-running operations.


def convert_json_to_excel(json_path: str, out_path: str = None):
    """Convert a TrackMan JSON report to a formatted Excel workbook.
    
    Args:
        json_path: Path to the downloaded TrackMan report JSON file
        out_path: Optional output path for the Excel file. If provided, saves directly without a dialog.
                  If None, shows a file save dialog to the user.
    
    Returns:
        Path: The path where the Excel file was saved, or None if the user cancelled.
    """
    # Load the JSON data from disk
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # Build the workbook with per-club sheets using the converter module
    wb = build_workbook_per_club(data)

    # If a path was provided, save directly
    if out_path:
        wb.save(out_path)
        return Path(out_path)

    # Otherwise, show file save dialog
    default_name = datetime.now().strftime("Trackman_Report_%Y%m%d_%H%M%S.xlsx")
    default_dir = str(Path.home() / "Documents")

    save_path = filedialog.asksaveasfilename(
        title="Save Converted Excel File",
        defaultextension=".xlsx",
        initialfile=default_name,
        initialdir=default_dir,
        filetypes=[("Excel Files", "*.xlsx")],
    )
    if not save_path:
        messagebox.showinfo("Cancelled", "Save cancelled. File not created.")
        return None

    wb.save(save_path)
    return Path(save_path)


# Configure the appearance theme for the application
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


class TrackmanApp(ctk.CTk):
    """Main application window for the TrackMan Report Converter.
    
    Single window application that:
    1. Automatically fetches TrackMan reports from Chrome history on startup
    2. Displays a grid of available reports for selection
    3. Downloads and converts the selected report to Excel
    4. Refreshes the report list after each conversion
    
    All UI remains within a single window without modal dialogs.
    """
    
    def __init__(self):
        super().__init__()
        self.title("TrackMan Converter")
        #self.geometry("800x600")
        # Allow the main window to be resized by the user so the report
        # selection area can expand or shrink as needed.
        self.resizable(True, True)
        self.configure(fg_color=DARK_BG)
        self.overlay = None  # Loading overlay reference
        self.token = None  # Store token for report selection
        self.content_frame = None  # Current content display frame
        self._cached_reports = None  # Enriched report list from last full metadata fetch

        # Create header with TrackMan branding (persistent across all views)
        header = ctk.CTkFrame(self, fg_color=HEADER_BG, corner_radius=0, height=30)
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="TrackMan Report Converter",
            font=("Segoe UI", 28, "bold"),
            text_color="white",
        ).pack(pady=12)

        # Inline spinner label for long-running tasks (hidden by default)
        self.spinner_label = ctk.CTkLabel(header, text="", font=("Segoe UI", 12), text_color="white")
        # show spinner text on the right side of the header so it's always visible
        self.spinner_label.pack(side="right", padx=12, pady=4)

        # Clear clubs cache button (moved to left side of header)
        clear_btn = ctk.CTkButton(header, text="Clear Clubs Cache", width=140, height=26, fg_color=SUBTLE_BTN, text_color="white", hover_color=SUBTLE_HOVER, command=self.clear_clubs_cache)
        clear_btn.pack(side="left", padx=(12, 6), pady=6)

        # Create main content container that will hold different views
        self.main_content = ctk.CTkFrame(self, fg_color=DARK_BG)
        self.main_content.pack(expand=True, fill="both", padx=0, pady=0)

        # Footer with copyright info (persistent across all views)
        ctk.CTkLabel(
            self,
            text=APP_FOOTER_TEXT,
            font=("Segoe UI", 11, "italic"),
            text_color=TEXT_MUTED,
        ).pack(side="bottom", pady=8)

        # Center the window on screen
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 1000
        window_height = 900
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        self.geometry(f"{window_width}x{window_height}+{x}+{y}")

        # Automatically start the report discovery process
        self.after(100, self.handle_cloud)


    def show_report_selector(self, reports, token):
        """Display the report selector grid directly within the main window.
        
        Args:
            reports: List of report dictionaries with 'id' and 'time' keys
            token: Authentication token for downloading reports from TrackMan API
        """
        log.info(f"show_report_selector: called with {len(reports)} report(s)")
        self._clear_content()
        log.info("show_report_selector: content cleared")
        self.token = token
        
        # Create container for the selector
        selector_container = ctk.CTkFrame(self.main_content, fg_color=DARK_BG)
        selector_container.pack(fill="both", expand=True, padx=0, pady=0)
        self.content_frame = selector_container

        # Create scrollable area to accommodate many reports
        scroll_area = ctk.CTkScrollableFrame(selector_container, fg_color=DARK_BG)
        scroll_area.pack(fill="both", expand=True, padx=20, pady=10)

        container = ctk.CTkFrame(scroll_area, fg_color=DARK_BG)
        container.pack(anchor="n", fill="x")

        # Sort reports by date (newest first)
        log.info("show_report_selector: sorting reports")
        reports.sort(key=lambda r: r["time"], reverse=True)
        log.info("show_report_selector: building header row")

        # Configure container as a simple rows table: # | Date | Report | Clubs | Action
        container.grid_columnconfigure(0, weight=0, minsize=40)
        container.grid_columnconfigure(1, weight=0)
        container.grid_columnconfigure(2, weight=1)
        container.grid_columnconfigure(3, weight=1)
        container.grid_columnconfigure(4, weight=0)

        # Header row
        ctk.CTkLabel(container, text="#", font=("Segoe UI", 16, "bold"), text_color=TEXT_MUTED).grid(row=0, column=0, padx=(8, 4), pady=4, sticky="e")
        ctk.CTkLabel(container, text="Date", font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY).grid(row=0, column=1, padx=28, pady=4, sticky="w")
        ctk.CTkLabel(container, text="Report", font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY).grid(row=0, column=2, padx=28, pady=4, sticky="w")
        ctk.CTkLabel(container, text="Clubs", font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY).grid(row=0, column=3, padx=48, pady=4, sticky="w")
        ctk.CTkLabel(container, text="Action", font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY).grid(row=0, column=4, padx=28, pady=4, sticky="w")
        log.info("show_report_selector: header row built, starting report rows")

        # Rows
        for i, r in enumerate(reports, start=1):
            date = r["time"]
            month = date.strftime("%b").upper()
            day = date.strftime("%d")
            year = date.strftime("%Y")

            # Sequence number cell
            ctk.CTkLabel(container, text=str(i), font=("Segoe UI", 13), text_color=TEXT_MUTED).grid(row=i, column=0, padx=(8, 4), pady=6, sticky="e")

            # Date cell (bigger font)
            ctk.CTkLabel(container, text=f"{month} {day} {year}", font=("Segoe UI", 16, "bold"), text_color=TEXT_PRIMARY).grid(row=i, column=1, padx=8, pady=6, sticky="w")

            # Report description cell
            ctk.CTkLabel(container, text="Multi Group Report", font=("Segoe UI", 12), text_color=TEXT_BODY).grid(row=i, column=2, padx=8, pady=6, sticky="w")

            # Clubs placeholder (will be filled asynchronously)
            clubs_lbl = ctk.CTkLabel(container, text="Loading...", font=("Segoe UI", 11), text_color=TEXT_MUTED)
            clubs_lbl.grid(row=i, column=3, padx=8, pady=6, sticky="w")

            # Action button
            btn = ctk.CTkButton(container, text="Select", fg_color=TRACKMAN_COLOUR, hover_color=ACCENT_HOVER, text_color="white", width=90, height=28, font=("Segoe UI", 11, "bold"), command=lambda rep=r: self.on_report_selected(rep))
            btn.grid(row=i, column=4, padx=8, pady=4)

            # Start background thread to fetch clubs for this report (uses cache)
            def fetch_and_update(rep_id, label):
                try:
                    clubs = fetch_report_clubs(self.token, rep_id)
                    text = ", ".join(clubs) if clubs else "—"
                except Exception:
                    text = "—"
                # schedule UI update on main thread
                self.after(0, lambda: label.configure(text=text))

            threading.Thread(target=fetch_and_update, args=(r["id"], clubs_lbl), daemon=True).start()

        log.info(f"show_report_selector: all {len(reports)} row(s) rendered successfully")
        # Force tkinter to recalculate layout and update the scrollable canvas scroll region
        self.update_idletasks()


    def on_report_selected(self, report):
        """Handle report selection: run download+convert in a background thread.

        UI updates are scheduled back onto the main thread via `self.after`.
        """
        def worker():
            try:
                json_path = download_report(self.token, report["id"])
                # update spinner text while converting
                self.after(0, lambda: self.show_overlay(" Converting to formatted Excel..."))

                # Load downloaded JSON once
                with open(json_path, "r", encoding="utf-8") as f:
                    report_data = json.load(f)

                out_dir = Path(r"C:\Trackman\Data")
                out_dir.mkdir(parents=True, exist_ok=True)
                date_str = report['time'].strftime('%Y_%m_%d')

                # Split by player and produce one session file + one master per player
                players = split_by_player(report_data)
                session_paths: list[Path] = []
                master_paths: list[Path] = []

                for player_name, player_data in players.items():
                    safe_player = _safe_filename_part(player_name)
                    try:
                        clubs = trackman_api.extract_clubs_from_report_json(player_data)
                        clubs_str = "_".join(clubs) if clubs else ""
                    except Exception:
                        clubs_str = ""

                    session_name = (
                        f"{date_str}_{safe_player}_{clubs_str}.xlsx" if clubs_str
                        else f"{date_str}_{safe_player}.xlsx"
                    )
                    out_path = out_dir / session_name

                    self.after(0, lambda pn=player_name: self.show_overlay(f" Converting {pn}..."))
                    wb = build_workbook_per_club(player_data)
                    wb.save(str(out_path))
                    session_paths.append(out_path)
                    log.info(f"on_report_selected: session saved -> {out_path}")

                    master_path = out_dir / f"Trackman_Master_{safe_player}.xlsx"
                    self.after(0, lambda pn=player_name: self.show_overlay(f" Updating master for {pn}..."))
                    try:
                        append_to_master_workbook(player_data, master_path)
                        master_paths.append(master_path)
                        log.info(f"on_report_selected: master updated -> {master_path}")
                    except Exception as _me:
                        log.warning(f"on_report_selected: master update failed for {player_name}: {_me}")

                _session_paths = session_paths
                _master_paths = master_paths

                def on_success():
                    self.hide_overlay()
                    msg = "Downloaded and converted!"
                    msg += "\n\nSession files:\n" + "\n".join(str(p) for p in _session_paths)
                    if _master_paths:
                        msg += "\n\nMaster workbooks updated:\n" + "\n".join(str(p) for p in _master_paths)
                    messagebox.showinfo("Success", msg)
                    # refresh report list but skip re-fetching metadata (it's unchanged)
                    self.handle_cloud(fetch_metadata=False)

                self.after(0, on_success)
            except Exception as e:
                # Capture message NOW — Python 3 deletes 'e' after the except block.
                err_msg = str(e)
                log.exception("on_report_selected worker failed")
                def on_error():
                    self.hide_overlay()
                    messagebox.showerror("Error", err_msg)
                self.after(0, on_error)

        # start background worker and show spinner
        self.show_overlay("Downloading selected report...")
        threading.Thread(target=worker, daemon=True).start()


    def _clear_content(self):
        """Remove the current content frame."""
        if self.content_frame:
            self.content_frame.destroy()
            self.content_frame = None
        # Clear all widgets from main_content
        for widget in self.main_content.winfo_children():
            widget.destroy()


    def clear_clubs_cache(self):
        """Clear the persisted and in-memory clubs cache via trackman_api."""
        try:
            trackman_api.clear_clubs_cache()
            messagebox.showinfo("Cache Cleared", "Clubs cache cleared.")
            # Show a brief overlay while we refresh club data
            self.show_overlay(" Refreshing club data...")
            # Refresh the report list so clubs are re-fetched (fresh cache)
            self.after(100, lambda: self.handle_cloud(fetch_metadata=True))
        except Exception as e:
            messagebox.showerror("Error", f"Failed to clear clubs cache:\n{e}")


    def show_overlay(self, text="Loading..."):
        """Start an animated inline spinner with base `text`."""
        self._spinner_base = text
        if getattr(self, "_spinner_running", False):
            return
        self._spinner_running = True
        self._spinner_dots = 0
        self._spinner_after_id = None
        self._spinner_loop()

    def hide_overlay(self):
        """Stop spinner animation and clear the label."""
        if not getattr(self, "_spinner_running", False):
            return
        self._spinner_running = False
        try:
            if getattr(self, "_spinner_after_id", None):
                self.after_cancel(self._spinner_after_id)
        except Exception:
            pass
        self._spinner_after_id = None
        self.spinner_label.configure(text="")

    def _spinner_loop(self):
        if not getattr(self, "_spinner_running", False):
            return
        dots = "." * (self._spinner_dots % 4)
        text = f"{getattr(self, '_spinner_base', '')} {dots}"
        self.spinner_label.configure(text=text)
        self._spinner_dots += 1
        self._spinner_after_id = self.after(300, self._spinner_loop)

    def _ask_for_token(self, on_success):
        """Show a dialog asking the user to paste their TrackMan bearer token.

        `on_success(token)` is called on the main thread once a valid token is
        saved, allowing the caller to resume whatever flow needs authentication.
        """
        dialog = ctk.CTkToplevel(self)
        dialog.title("TrackMan Login Required")
        dialog.geometry("560x320")
        dialog.resizable(False, False)
        dialog.grab_set()  # modal

        ctk.CTkLabel(
            dialog,
            text="Could not read your TrackMan login from Chrome cookies.\n\n"
                 "To get your Bearer token:\n"
                 "  1. Open Chrome and go to a TrackMan report page.\n"
                 "  2. Press F12 → Application → Cookies → trackmangolf.com\n"
                 "  3. Find 'appSession' and copy its Value.",
            font=("Segoe UI", 12),
            justify="left",
            wraplength=520,
        ).pack(padx=20, pady=(20, 10), anchor="w")

        entry = ctk.CTkEntry(dialog, width=520, placeholder_text="Paste token here…")
        entry.pack(padx=20, pady=(0, 12))

        status_lbl = ctk.CTkLabel(dialog, text="", font=("Segoe UI", 11), text_color="#e05555")
        status_lbl.pack()

        def _on_submit():
            token = entry.get().strip()
            if not token:
                status_lbl.configure(text="Token cannot be empty.")
                return
            if not trackman_auth._is_valid_token(token):
                status_lbl.configure(text="Token contains invalid characters. Please re-copy from Chrome.")
                return
            trackman_auth.save_token(token)
            log.info("_ask_for_token: token manually saved, len=%d", len(token))
            dialog.destroy()
            on_success(token)

        ctk.CTkButton(dialog, text="Save & Continue", command=_on_submit,
                      fg_color=TRACKMAN_COLOUR, hover_color=ACCENT_HOVER,
                      text_color="white").pack(pady=(4, 0))

        entry.bind("<Return>", lambda _e: _on_submit())

    def handle_cloud(self, fetch_metadata: bool = True):
        """Fetch TrackMan reports from Chrome history and display report selector.
        
        This method orchestrates the entire discovery and selection flow:
        1. Check for saved authentication token (or prompt user to login)
        2. Search Chrome history for TrackMan report URLs
        3. Fetch metadata (upload dates) from TrackMan API
        4. Display the report selector grid in the main window
        
        This is called on app startup and again after each successful conversion.
        """
        # Run the discovery flow in a background thread so UI remains responsive
        def worker():
            try:
                from trackman_api import get_all_report_ids_from_chrome, fetch_report_metadata_batch

                log.info("handle_cloud: getting token")
                token = trackman_auth.get_saved_token() or trackman_auth.login_via_browser()
                if not token:
                    log.warning("handle_cloud: no token available, showing manual input dialog")
                    # Schedule the token dialog on the main thread; resume handle_cloud after success.
                    def _resume_after_token(saved_token):
                        self.token = saved_token
                        self.hide_overlay()
                        self.handle_cloud(fetch_metadata=fetch_metadata)
                    self.after(0, lambda: [self.hide_overlay(), self._ask_for_token(_resume_after_token)])
                    return
                log.info("handle_cloud: token obtained")

                # If skipping metadata and we have a cached list, go straight to the selector
                if not fetch_metadata and self._cached_reports is not None:
                    log.info("handle_cloud: using cached report list (%d reports)", len(self._cached_reports))
                    cached = self._cached_reports
                    def _show_cached():
                        self.hide_overlay()
                        self.show_report_selector(cached, token)
                    self.after(0, _show_cached)
                    return

                # search chrome history
                self.after(0, lambda: self.show_overlay(" Searching Chrome history for TrackMan reports..."))
                raw_reports = get_all_report_ids_from_chrome(limit=200)
                log.info(f"handle_cloud: found {len(raw_reports) if raw_reports else 0} raw report(s) in Chrome history")

                if not raw_reports:
                    def no_reports():
                        self.hide_overlay()
                        messagebox.showerror(
                            "No Reports Found",
                            "No recent TrackMan reports were found in Chrome history.\n"
                            "Please open a TrackMan report in Chrome and try again."
                        )
                    self.after(0, no_reports)
                    return

                # dedupe
                seen = set()
                unique_reports = []
                for r in raw_reports:
                    rid = r.get("id")
                    if rid and rid not in seen:
                        seen.add(rid)
                        unique_reports.append(r)
                log.info(f"handle_cloud: {len(unique_reports)} unique report(s) after dedup")

                # fetch metadata (optional)
                if fetch_metadata:
                    self.after(0, lambda: self.show_overlay(" Getting upload dates from TrackMan..."))
                    report_ids = [r["id"] for r in unique_reports]
                    metadata_list = fetch_report_metadata_batch(token, report_ids, max_workers=5)
                    log.info(f"handle_cloud: metadata batch returned {sum(1 for m in metadata_list if m)} non-None result(s)")

                    enriched = []
                    for r, meta in zip(unique_reports, metadata_list):
                        if meta and meta.get("created"):
                            try:
                                meta["time"] = datetime.fromisoformat(meta["created"].replace("Z", "+00:00"))
                            except Exception:
                                meta["time"] = r.get("time", datetime.utcnow())
                            enriched.append(meta)
                        else:
                            # Metadata fetch failed — fall back to Chrome history visit time
                            enriched.append({"id": r["id"], "time": r.get("time", datetime.utcnow())})
                    self._cached_reports = enriched  # cache for subsequent refreshes
                else:
                    # Skip fetching metadata; use Chrome history visit time for each report
                    enriched = [{"id": r["id"], "time": r.get("time", datetime.utcnow())} for r in unique_reports]

                log.info(f"handle_cloud: showing selector with {len(enriched)} report(s)")
                # show selector on main thread
                def _show_selector():
                    try:
                        self.hide_overlay()
                        self.show_report_selector(enriched, token)
                    except Exception as _e:
                        log.exception("show_report_selector raised an exception")
                        messagebox.showerror("Error", str(_e))
                self.after(0, _show_selector)
            except Exception as e:
                # Capture message NOW — Python 3 deletes 'e' after the except block,
                # so referencing it inside a lambda scheduled via after() would raise NameError.
                err_msg = str(e)
                log.exception("handle_cloud worker failed")
                self.after(0, lambda: [self.hide_overlay(), messagebox.showerror("Error", err_msg)])

        # start worker thread and show initial spinner
        self.show_overlay(" Checking TrackMan login...")
        threading.Thread(target=worker, daemon=True).start()

# Entry point for the application
if __name__ == "__main__":
    # Create and run the main GUI window
    app = TrackmanApp()
    app.mainloop()
