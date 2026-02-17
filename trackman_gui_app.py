import customtkinter as ctk
from tkinter import messagebox, filedialog
import json
from pathlib import Path
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter

import trackman_auth
from trackman_api import download_report, get_latest_report_id_from_chrome
from converter import build_workbook_per_club



class LoadingOverlay(ctk.CTkToplevel):
    def __init__(self, parent, text="Loading..."):
        super().__init__(parent)
        self.geometry(
            f"{parent.winfo_width()}x{parent.winfo_height()}+"
            f"{parent.winfo_rootx()}+{parent.winfo_rooty()}"
        )
        self.overrideredirect(True)
        self.configure(bg="#000000")
        self.attributes("-topmost", True)
        self.attributes("-alpha", 0.65)

        frame = ctk.CTkFrame(self, fg_color="#1a1a1a", corner_radius=16)
        frame.place(relx=0.5, rely=0.5, anchor="center")

        spinner = ctk.CTkLabel(frame, text=".", font=("Segoe UI Emoji", 36))
        spinner.pack(pady=(20, 5))
        self.label = ctk.CTkLabel(frame, text=text, font=("Segoe UI", 14))
        self.label.pack(pady=(0, 20))
        self.update_idletasks()

    def update_text(self, text: str):
        self.label.configure(text=text)
        self.update_idletasks()



APP_FOOTER_TEXT = "© 2026 TrackMan Converter by Tom McIntyre and Brian McIntyre. All rights reserved."
TRACKMAN_COLOUR = "#001AFF"
DARK_BG = "#1E1E1E"



def convert_json_to_excel(json_path: str, out_path: str = None):
    """Convert a TrackMan JSON to Excel.

    If `out_path` is provided, the workbook is saved there without showing a file dialog.
    If `out_path` is None, the GUI save dialog is shown (original behavior).
    Returns a `Path` to the saved file, or `None` if the user cancelled the dialog.
    """
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    wb = build_workbook_per_club(data)

    if out_path:
        wb.save(out_path)
        return Path(out_path)

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



ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")


class TrackmanApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("TrackMan Converter")
        self.geometry("700x480")
        self.resizable(False, False)
        self.configure(fg_color=DARK_BG)
        self.overlay = None

     
        header = ctk.CTkFrame(self, fg_color=TRACKMAN_COLOUR, corner_radius=0, height=90)
        header.pack(fill="x")
        ctk.CTkLabel(
            header,
            text="TrackMan Report Converter",
            font=("Segoe UI", 28, "bold"),
            text_color="white",
        ).pack(pady=25)

  
        content = ctk.CTkFrame(self, fg_color=DARK_BG)
        content.pack(expand=True, fill="both")

        ctk.CTkLabel(
            content,
            text="Automatically fetches your latest TrackMan report\n"
                 "and converts it to a formatted Excel file.",
            font=("Segoe UI", 16),
            text_color="#BBBBBB",
            justify="center",
        ).pack(pady=(60, 30))

        ctk.CTkButton(
            content,
            text=" Download & Convert Latest Report",
            width=320,
            height=60,
            corner_radius=12,
            font=("Segoe UI", 18, "bold"),
            fg_color=TRACKMAN_COLOUR,
            hover_color="#FF8533",
            text_color="white",
            command=self.handle_cloud,
        ).pack(pady=10)

        self.status_label = ctk.CTkLabel(
            content, text="", font=("Segoe UI", 14), text_color="#AAAAAA"
        )
        self.status_label.pack(pady=(30, 10))

        ctk.CTkLabel(
            self,
            text=APP_FOOTER_TEXT,
            font=("Segoe UI", 11, "italic"),
            text_color="#666666",
        ).pack(side="bottom", pady=8)


    def show_overlay(self, text="Loading..."):
        if self.overlay:
            self.overlay.destroy()
        self.overlay = LoadingOverlay(self, text)
        self.overlay.update()

    def hide_overlay(self):
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None

    def handle_cloud(self):
        try:
            from trackman_api import get_all_report_ids_from_chrome, fetch_report_metadata_batch

            self.show_overlay(" Checking TrackMan login...")
            token = trackman_auth.get_saved_token() or trackman_auth.login_via_browser()
            if not token:
                raise Exception("Could not retrieve TrackMan token.")

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

            seen = set()
            unique_reports = []
            for r in raw_reports:
                rid = r.get("id")
                if rid and rid not in seen:
                    seen.add(rid)
                    unique_reports.append(r)

            self.overlay.update_text("Getting upload dates from TrackMan...")
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
                    enriched.append({"id": r["id"], "time": datetime.utcnow()})  # fallback

            reports = enriched
            self.hide_overlay()

            #show_selector(reports, token, self)
            selector = ctk.CTkToplevel(self)
            selector.title("Select TrackMan Report")
            selector.geometry("720x560")
            selector.grab_set()
            selector.configure(fg_color="#1E1E1E")

            title_label = ctk.CTkLabel(
                selector,
                text="Select a TrackMan Report",
                font=("Segoe UI", 22, "bold"),
                text_color="white",
            )
            title_label.pack(pady=(20, 5))

            scroll_area = ctk.CTkScrollableFrame(selector, fg_color="#1E1E1E")
            scroll_area.pack(fill="both", expand=True, padx=40, pady=20)

            container = ctk.CTkFrame(scroll_area, fg_color="#1E1E1E")
            container.pack(anchor="center")

            

          
            reports.sort(key=lambda r: r["time"], reverse=True)

           
            cols = 3
            for i, r in enumerate(reports):
                date = r["time"]
                month = date.strftime("%b").upper()
                day = date.strftime("%d")
                year = date.strftime("%Y")

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

                
                ctk.CTkLabel(
                    frame,
                    text=f"{month}\n{day}",
                    font=("Segoe UI", 20, "bold"),
                    text_color=TRACKMAN_COLOUR,
                    justify="center",
                ).pack(pady=(10, 4))

                ctk.CTkLabel(
                    frame,
                    text="Multi Group Report",
                    font=("Segoe UI", 13),
                    text_color="white",
                ).pack()

                ctk.CTkLabel(
                    frame,
                    text=year,
                    font=("Segoe UI", 11, "italic"),
                    text_color="#AAAAAA",
                ).pack(pady=(0, 8))

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
                    command=lambda rep=r: on_select(rep),
                ).pack(pady=(5, 10))


                def on_select(report):
                    selector.destroy()
                    self.show_overlay("Downloading selected report...")
                    try:
                        json_path = download_report(token, report["id"])
                        self.overlay.update_text(" Converting to formatted Excel...")
                        out_dir = Path(r"C:\Trackman\Data")  # or Path.home() / "Documents"
                        out_dir.mkdir(parents=True, exist_ok=True)
                        default_name = f"{report['time'].strftime('%Y_%m_%d')}.xlsx"
                        out_path = out_dir / default_name
                        result = convert_json_to_excel(json_path, str(out_path))
                        self.hide_overlay()
                        messagebox.showinfo("Success", f" Downloaded and converted!\nSaved as:\n{result}")
                    except Exception as e:
                        self.hide_overlay()
                        messagebox.showerror("Error", str(e))

        except Exception as e:
            self.hide_overlay()
            messagebox.showerror("Error", str(e))

        def on_select(report):
            selector.destroy()
            self.show_overlay("Downloading selected report...")
            try:
                json_path = download_report(token, report["id"])
                self.overlay.update_text(" Converting to formatted Excel...")
                out_dir = Path(r"C:\Trackman\Data")  # or Path.home() / "Documents"
                out_dir.mkdir(parents=True, exist_ok=True)
                default_name = f"{report['time'].strftime('%Y_%m_%d')}.xlsx"
                out_path = out_dir / default_name
                result = convert_json_to_excel(json_path, str(out_path))
                self.hide_overlay()
                messagebox.showinfo("Success", f" Downloaded and converted!\nSaved as:\n{result}")
            except Exception as e:
                self.hide_overlay()
                messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = TrackmanApp()
    app.mainloop()
