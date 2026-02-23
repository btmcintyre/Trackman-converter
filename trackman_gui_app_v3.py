# TrackMan Report Converter GUI Application
# This module provides the main GUI interface for downloading and converting TrackMan reports to Excel format.

# UI Framework and dialogs
import customtkinter as ctk
from tkinter import messagebox, filedialog

# Data handling
import json
from pathlib import Path
from datetime import datetime
import pandas as pd
import threading

# Excel workbook creation and formatting
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

# Project-specific modules
import trackman_auth
from trackman_api import download_report, get_latest_report_id_from_chrome
from converter import build_workbook_per_club

# Application-wide constants and theme configuration
APP_FOOTER_TEXT = "© 2026 TrackMan Converter by Tom McIntyre and Brian McIntyre. All rights reserved."
TRACKMAN_COLOUR = "#001AFF"  # Primary blue color used throughout the UI
DARK_BG = "#FFFFFF"  # Main background color for the application (white)

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

        # Create header with TrackMan branding (persistent across all views)
        header = ctk.CTkFrame(self, fg_color=TRACKMAN_COLOUR, corner_radius=0, height=30)
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

        # Create main content container that will hold different views
        self.main_content = ctk.CTkFrame(self, fg_color=DARK_BG)
        self.main_content.pack(expand=True, fill="both", padx=0, pady=0)

        # Footer with copyright info (persistent across all views)
        ctk.CTkLabel(
            self,
            text=APP_FOOTER_TEXT,
            font=("Segoe UI", 11, "italic"),
            text_color="#333333",
        ).pack(side="bottom", pady=8)

        # Center the window on screen
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        window_width = 600
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
        self._clear_content()
        self.token = token
        
        # Create container for the selector
        selector_container = ctk.CTkFrame(self.main_content, fg_color=DARK_BG)
        selector_container.pack(fill="both", expand=True, padx=0, pady=0)
        self.content_frame = selector_container

        # Create scrollable area to accommodate many reports
        scroll_area = ctk.CTkScrollableFrame(selector_container, fg_color=DARK_BG)
        scroll_area.pack(fill="both", expand=True, padx=20, pady=10)

        container = ctk.CTkFrame(scroll_area, fg_color=DARK_BG)
        container.pack(anchor="center", expand=True)

        # Sort reports by date (newest first)
        reports.sort(key=lambda r: r["time"], reverse=True)

        # Configure container as a simple rows table: Date | Report | Year | Action
        container.grid_columnconfigure(0, weight=0)
        container.grid_columnconfigure(1, weight=1)
        container.grid_columnconfigure(2, weight=0)
        container.grid_columnconfigure(3, weight=0)

        # Header row
        ctk.CTkLabel(container, text="Date", font=("Segoe UI", 14, "bold"), text_color="black").grid(row=0, column=0, padx=8, pady=4, sticky="w")
        ctk.CTkLabel(container, text="Report", font=("Segoe UI", 14, "bold"), text_color="black").grid(row=0, column=1, padx=8, pady=4, sticky="w")
        ctk.CTkLabel(container, text="Action", font=("Segoe UI", 14, "bold"), text_color="black").grid(row=0, column=3, padx=8, pady=4, sticky="w")

        # Rows
        for i, r in enumerate(reports, start=1):
            date = r["time"]
            month = date.strftime("%b").upper()
            day = date.strftime("%d")
            year = date.strftime("%Y")

            # Date cell (bigger font)
            ctk.CTkLabel(container, text=f"{month} {day} {year}", font=("Segoe UI", 16, "bold"), text_color="black").grid(row=i, column=0, padx=8, pady=6, sticky="w")

            # Report description cell
            ctk.CTkLabel(container, text="Multi Group Report", font=("Segoe UI", 12), text_color="#333333").grid(row=i, column=1, padx=8, pady=6, sticky="w")

            # Action button
            btn = ctk.CTkButton(container, text="Select", fg_color=TRACKMAN_COLOUR, hover_color="#FF8533", text_color="white", width=90, height=28, font=("Segoe UI", 11, "bold"), command=lambda rep=r: self.on_report_selected(rep))
            btn.grid(row=i, column=3, padx=8, pady=4)


    def on_report_selected(self, report):
        """Handle report selection: run download+convert in a background thread.

        UI updates are scheduled back onto the main thread via `self.after`.
        """
        def worker():
            try:
                json_path = download_report(self.token, report["id"])
                # update spinner text while converting
                self.after(0, lambda: self.show_overlay(" Converting to formatted Excel..."))

                out_dir = Path(r"C:\Trackman\Data")
                out_dir.mkdir(parents=True, exist_ok=True)
                default_name = f"{report['time'].strftime('%Y_%m_%d')}.xlsx"
                out_path = out_dir / default_name

                result = convert_json_to_excel(json_path, str(out_path))

                def on_success():
                    self.hide_overlay()
                    messagebox.showinfo("Success", f" Downloaded and converted!\nSaved as:\n{result}")
                    # refresh report list but skip re-fetching metadata (it's unchanged)
                    self.handle_cloud(fetch_metadata=False)

                self.after(0, on_success)
            except Exception as e:
                def on_error():
                    self.hide_overlay()
                    messagebox.showerror("Error", str(e))
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

                token = trackman_auth.get_saved_token() or trackman_auth.login_via_browser()
                if not token:
                    raise Exception("Could not retrieve TrackMan token.")

                # search chrome history
                self.after(0, lambda: self.show_overlay(" Searching Chrome history for TrackMan reports..."))
                raw_reports = get_all_report_ids_from_chrome(limit=50)

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

                # fetch metadata (optional)
                if fetch_metadata:
                    self.after(0, lambda: self.show_overlay(" Getting upload dates from TrackMan..."))
                    report_ids = [r["id"] for r in unique_reports]
                    metadata_list = fetch_report_metadata_batch(token, report_ids, max_workers=5)

                    enriched = []
                    for r, meta in zip(unique_reports, metadata_list):
                        if meta and meta.get("created"):
                            try:
                                meta["time"] = datetime.fromisoformat(meta["created"].replace("Z", "+00:00"))
                            except Exception:
                                meta["time"] = datetime.utcnow()
                            enriched.append(meta)
                        else:
                            enriched.append({"id": r["id"], "time": datetime.utcnow()})
                else:
                    # Skip fetching metadata; provide fallback timestamps so UI can display
                    enriched = [{"id": r["id"], "time": datetime.utcnow()} for r in unique_reports]

                # show selector on main thread
                self.after(0, lambda: [self.hide_overlay(), self.show_report_selector(enriched, token)])
            except Exception as e:
                self.after(0, lambda: [self.hide_overlay(), messagebox.showerror("Error", str(e))])

        # start worker thread and show initial spinner
        self.show_overlay(" Checking TrackMan login...")
        threading.Thread(target=worker, daemon=True).start()

# Entry point for the application
if __name__ == "__main__":
    # Create and run the main GUI window
    app = TrackmanApp()
    app.mainloop()
