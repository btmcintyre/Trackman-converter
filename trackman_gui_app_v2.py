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
DARK_BG = "#1E1E1E"  # Dark background color for the main application

# Overlay window displayed during long-running operations (downloading, converting)
class LoadingOverlay(ctk.CTkToplevel):

    def __init__(self, parent, text="Loading..."):
        super().__init__(parent)
        # Position the overlay to match the parent window
        self.geometry(
            f"{parent.winfo_width()}x{parent.winfo_height()}+"
            f"{parent.winfo_rootx()}+{parent.winfo_rooty()}"
        )
        self.overrideredirect(True)  # Remove window decorations
        self.configure(bg="#000000")
        self.attributes("-topmost", True)  # Keep overlay on top
        self.attributes("-alpha", 0.65)  # Make it semi-transparent

        # Create frame and labels for the loading message
        frame = ctk.CTkFrame(self, fg_color="#1a1a1a", corner_radius=16)
        frame.place(relx=0.5, rely=0.5, anchor="center")

        spinner = ctk.CTkLabel(frame, text=".", font=("Segoe UI Emoji", 36))
        spinner.pack(pady=(20, 5))
        self.label = ctk.CTkLabel(frame, text=text, font=("Segoe UI", 14))
        self.label.pack(pady=(0, 20))
        self.update_idletasks()

    def update_text(self, text: str):
        """Update the loading message displayed in the overlay."""
        self.label.configure(text=text)
        self.update_idletasks()


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
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


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
        self.geometry("700x600")
        self.resizable(False, False)
        self.configure(fg_color=DARK_BG)
        self.overlay = None  # Loading overlay reference
        self.token = None  # Store token for report selection
        self.content_frame = None  # Current content display frame

        # Create header with TrackMan branding (persistent across all views)
        header = ctk.CTkFrame(self, fg_color=TRACKMAN_COLOUR, corner_radius=0, height=90)
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="TrackMan Report Converter",
            font=("Segoe UI", 28, "bold"),
            text_color="white",
        ).pack(pady=25)

        # Create main content container that will hold different views
        self.main_content = ctk.CTkFrame(self, fg_color=DARK_BG)
        self.main_content.pack(expand=True, fill="both", padx=0, pady=0)

        # Footer with copyright info (persistent across all views)
        ctk.CTkLabel(
            self,
            text=APP_FOOTER_TEXT,
            font=("Segoe UI", 11, "italic"),
            text_color="#666666",
        ).pack(side="bottom", pady=8)

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
        selector_container = ctk.CTkFrame(self.main_content, fg_color="#1E1E1E")
        selector_container.pack(fill="both", expand=True, padx=0, pady=0)
        self.content_frame = selector_container

        # Add title
        title_label = ctk.CTkLabel(
            selector_container,
            text="Select a TrackMan Report",
            font=("Segoe UI", 20, "bold"),
            text_color="white",
        )
        title_label.pack(pady=(15, 10))

        # Create scrollable area to accommodate many reports
        scroll_area = ctk.CTkScrollableFrame(selector_container, fg_color="#1E1E1E")
        scroll_area.pack(fill="both", expand=True, padx=20, pady=10)

        container = ctk.CTkFrame(scroll_area, fg_color="#1E1E1E")
        container.pack(anchor="center")

        # Sort reports by date (newest first)
        reports.sort(key=lambda r: r["time"], reverse=True)
        
        # Display reports in a grid layout (3 columns)
        cols = 3
        for i, r in enumerate(reports):
            date = r["time"]
            month = date.strftime("%b").upper()
            day = date.strftime("%d")
            year = date.strftime("%Y")

            # Create a card for each report
            frame = ctk.CTkFrame(
                container,
                fg_color="#2A2A2A",
                corner_radius=12,
                border_color="#444",
                border_width=1,
                width=180,
                height=160,
            )
            frame.grid(row=i // cols, column=i % cols, padx=18, pady=18)
            
            # Display date in a prominent way
            ctk.CTkLabel(
                frame,
                text=f"{month}\n{day}",
                font=("Segoe UI", 20, "bold"),
                text_color=TRACKMAN_COLOUR,
                justify="center",
            ).pack(pady=(10, 4))

            # Label the report type
            ctk.CTkLabel(
                frame,
                text="Multi Group Report",
                font=("Segoe UI", 13),
                text_color="white",
            ).pack()

            # Display year in italics
            ctk.CTkLabel(
                frame,
                text=year,
                font=("Segoe UI", 11, "italic"),
                text_color="#AAAAAA",
            ).pack(pady=(0, 8))

            # Select button for this report
            ctk.CTkButton(
                frame,
                text="Select",
                fg_color=TRACKMAN_COLOUR,
                hover_color="#FF8533",
                text_color="white",
                corner_radius=8,
                height=28,
                width=100,
                font=("Segoe UI", 13, "bold"),
                command=lambda rep=r: self.on_report_selected(rep),
            ).pack(pady=(5, 10))


    def on_report_selected(self, report):
        """Handle report selection: download JSON from TrackMan API and convert to Excel."""
        self.show_overlay("Downloading selected report...")
        try:
            # Download the JSON report from TrackMan
            json_path = download_report(self.token, report["id"])
            self.overlay.update_text(" Converting to formatted Excel...")
            
            # Create output directory if it doesn't exist
            out_dir = Path(r"C:\Trackman\Data")
            out_dir.mkdir(parents=True, exist_ok=True)
            
            # Generate output filename based on report date
            default_name = f"{report['time'].strftime('%Y_%m_%d')}.xlsx"
            out_path = out_dir / default_name
            
            # Convert JSON to Excel and save
            result = convert_json_to_excel(json_path, str(out_path))
            self.hide_overlay()
            messagebox.showinfo("Success", f" Downloaded and converted!\nSaved as:\n{result}")
            # Refresh the report list after successful conversion
            self.handle_cloud()
        except Exception as e:
            self.hide_overlay()
            messagebox.showerror("Error", str(e))


    def _clear_content(self):
        """Remove the current content frame."""
        if self.content_frame:
            self.content_frame.destroy()
            self.content_frame = None
        # Clear all widgets from main_content
        for widget in self.main_content.winfo_children():
            widget.destroy()


    def show_overlay(self, text="Loading..."):
        """Display a loading overlay with the given text during long operations."""
        if self.overlay:
            self.overlay.destroy()
        self.overlay = LoadingOverlay(self, text)
        self.overlay.update()

    def hide_overlay(self):
        """Hide and destroy the loading overlay."""
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None

    def handle_cloud(self):
        """Fetch TrackMan reports from Chrome history and display report selector.
        
        This method orchestrates the entire discovery and selection flow:
        1. Check for saved authentication token (or prompt user to login)
        2. Search Chrome history for TrackMan report URLs
        3. Fetch metadata (upload dates) from TrackMan API
        4. Display the report selector grid in the main window
        
        This is called on app startup and again after each successful conversion.
        """
        try:
            from trackman_api import get_all_report_ids_from_chrome, fetch_report_metadata_batch

            # Step 1: Ensure user is authenticated
            self.show_overlay(" Checking TrackMan login...")
            token = trackman_auth.get_saved_token() or trackman_auth.login_via_browser()
            if not token:
                raise Exception("Could not retrieve TrackMan token.")

            # Step 2: Scan Chrome history for TrackMan reports
            self.overlay.update_text(" Searching Chrome history for TrackMan reports...")
            raw_reports = get_all_report_ids_from_chrome(limit=50)

            if not raw_reports:
                self.hide_overlay()
                messagebox.showerror(
                    "No Reports Found",
                    "No recent TrackMan reports were found in Chrome history.\n"
                    "Please open a TrackMan report in Chrome and try again."
                )
                return

            # Remove duplicate report IDs while preserving order
            seen = set()
            unique_reports = []
            for r in raw_reports:
                rid = r.get("id")
                if rid and rid not in seen:
                    seen.add(rid)
                    unique_reports.append(r)

            # Step 3: Fetch creation dates and other metadata from TrackMan API
            self.overlay.update_text("Getting upload dates from TrackMan...")
            report_ids = [r["id"] for r in unique_reports]
            metadata_list = fetch_report_metadata_batch(token, report_ids, max_workers=5)
            
            # Enrich reports with metadata (creation timestamps)
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

            # Step 4: Hide overlay and show report selector in the main window
            self.hide_overlay()
            self.show_report_selector(enriched, token)

        except Exception as e:
            self.hide_overlay()
            messagebox.showerror("Error", str(e))

# Entry point for the application
if __name__ == "__main__":
    # Create and run the main GUI window
    app = TrackmanApp()
    app.mainloop()
