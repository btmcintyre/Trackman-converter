import os, sys, builtins, logging
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

# -----------------------------
# SAFETY + LOGGING (for .exe)
# -----------------------------
def init_safe_logging():
    if not sys.stdin:
        sys.stdin = open(os.devnull, "r")
    if not sys.stdout:
        sys.stdout = open(os.devnull, "w", encoding="utf-8", errors="ignore")
    if not sys.stderr:
        sys.stderr = open(os.devnull, "w", encoding="utf-8", errors="ignore")
    builtins.input = lambda *a, **kw: ""

    log_dir = Path(os.getenv("LOCALAPPDATA", Path.home())) / "TrackmanConverter" / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / "app.log"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8")]
    )
    logging.info("=== TrackmanConverter started ===")
    return log_file

LOG_PATH = init_safe_logging()

# -----------------------------
# CONSTANTS
# -----------------------------
APP_FOOTER_TEXT = "© 2025 TrackMan Converter by Tom McIntyre"
TRACKMAN_ORANGE = "#FF6600"
DARK_BG = "#1E1E1E"

# -----------------------------
# LOADING OVERLAY
# -----------------------------
class LoadingOverlay(ctk.CTkToplevel):
    def __init__(self, parent, text="Loading..."):
        super().__init__(parent)
        self.geometry(f"{parent.winfo_width()}x{parent.winfo_height()}+{parent.winfo_rootx()}+{parent.winfo_rooty()}")
        self.overrideredirect(True)
        self.configure(bg="#000000")
        self.attributes("-topmost", True)
        self.attributes("-alpha", 0.65)

        frame = ctk.CTkFrame(self, fg_color="#1a1a1a", corner_radius=16)
        frame.place(relx=0.5, rely=0.5, anchor="center")

        spinner = ctk.CTkLabel(frame, text="⏳", font=("Segoe UI Emoji", 36))
        spinner.pack(pady=(20, 5))

        self.label = ctk.CTkLabel(frame, text=text, font=("Segoe UI", 14))
        self.label.pack(pady=(0, 20))
        self.update_idletasks()

    def update_text(self, text):
        self.label.configure(text=text)
        self.update_idletasks()

# -----------------------------
# EXCEL CONVERSION LOGIC
# -----------------------------
COLUMNS = [
    "Time", "Club Speed (Mph)", "Ball Speed (Mph)", "Smash Factor",
    "Carry (Yds)", "Total (Yds)", "Impact Height (mm)", "Impact Offset (mm)",
    "Club Path (Deg)", "Face Angle (Deg)", "Face To Path (Deg)",
    "Launch Direction (Deg)", "Attack Angle (Deg)", "Dynamic Loft (Deg)",
    "Launch Angle (Deg)", "Spin Loft (Deg)", "Spin Rate (Rpm)",
    "Spin Axis (Deg)", "Curve (Ft)", "Carry Side (Ft)", "Total Side (Ft)",
    "Max Height (Ft)", "Landing Angle (Deg)", "Swing Direction (Deg)",
    "Swing Plane (Deg)", "Swing Radius", "DPlane Tilt", "Low Point (In)",
    "Landing Height", "Hang Time (Sec)", "Dynamic Lie (Deg)"
]

def _fmt_time(iso: str):
    if not iso: return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except:
        return iso

def _conv(v, factor=1):
    if v is None: return ""
    try: return f"{float(v) * factor:.2f}"
    except: return ""

def convert_measurement_to_row(m: dict):
    if not m:
        return {k: "" for k in COLUMNS}
    return {
        "Time": _fmt_time(m.get("Time")),
        "Club Speed (Mph)": _conv(m.get("ClubSpeed"), 2.23694),
        "Ball Speed (Mph)": _conv(m.get("BallSpeed"), 2.23694),
        "Smash Factor": _conv(m.get("SmashFactor")),
        "Carry (Yds)": _conv(m.get("Carry"), 1.09361),
        "Total (Yds)": _conv(m.get("Total"), 1.09361),
        "Impact Height (mm)": _conv(m.get("ImpactHeight"), 1000),
        "Impact Offset (mm)": _conv(m.get("ImpactOffset"), 1000),
        "Club Path (Deg)": _conv(m.get("ClubPath")),
        "Face Angle (Deg)": _conv(m.get("FaceAngle")),
        "Face To Path (Deg)": _conv(m.get("FaceToPath")),
        "Launch Direction (Deg)": _conv(m.get("LaunchDirection")),
        "Attack Angle (Deg)": _conv(m.get("AttackAngle")),
        "Dynamic Loft (Deg)": _conv(m.get("DynamicLoft")),
        "Launch Angle (Deg)": _conv(m.get("LaunchAngle")),
        "Spin Loft (Deg)": _conv(m.get("SpinLoft")),
        "Spin Rate (Rpm)": _conv(m.get("SpinRate")),
        "Spin Axis (Deg)": _conv(m.get("SpinAxis")),
        "Curve (Ft)": _conv(m.get("Curve"), 3.28084),
        "Carry Side (Ft)": _conv(m.get("CarrySide"), 3.28084),
        "Total Side (Ft)": _conv(m.get("TotalSide"), 3.28084),
        "Max Height (Ft)": _conv(m.get("MaxHeight"), 3.28084),
        "Landing Angle (Deg)": _conv(m.get("LandingAngle")),
        "Swing Direction (Deg)": _conv(m.get("SwingDirection")),
        "Swing Plane (Deg)": _conv(m.get("SwingPlane")),
        "Swing Radius": _conv(m.get("SwingRadius")),
        "DPlane Tilt": _conv(m.get("DPlaneTilt")),
        "Low Point (In)": _conv(m.get("LowPointDistance"), 39.3701),
        "Landing Height": _conv(m.get("LandingHeight")),
        "Hang Time (Sec)": _conv(m.get("HangTime")),
        "Dynamic Lie (Deg)": _conv(m.get("DynamicLie")),
    }

def build_workbook(data: dict):
    wb = Workbook()
    wb.remove(wb.active)
    groups = data.get("StrokeGroups", []) or []
    all_rows = []

    for g in groups:
        club = str(g.get("Club", "Unknown Club"))[:31]
        ws = wb.create_sheet(title=club)
        rows = [convert_measurement_to_row(s.get("Measurement", {})) for s in g.get("Strokes", [])]
        df = pd.DataFrame(rows, columns=COLUMNS)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)
        all_rows.extend(rows)
        ws.freeze_panes = "A2"

    if all_rows:
        ws_all = wb.create_sheet("All Data")
        df = pd.DataFrame(all_rows, columns=COLUMNS)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws_all.append(r)

    return wb

def convert_json_to_excel(json_path: str) -> Path:
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    wb = build_workbook(data)

    default_name = f"Trackman_Report_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    save_path = filedialog.asksaveasfilename(
        title="Save Converted Excel File",
        defaultextension=".xlsx",
        initialfile=default_name,
        initialdir=str(Path.home() / "Documents"),
        filetypes=[("Excel Files", "*.xlsx")],
    )
    if not save_path:
        messagebox.showinfo("Cancelled", "Save cancelled.")
        return None

    wb.save(save_path)
    return save_path

# -----------------------------
# MODERN GUI
# -----------------------------
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class TrackmanApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("TrackMan Converter")
        self.geometry("700x500")
        self.resizable(False, False)
        self.configure(fg_color=DARK_BG)
        self.overlay = None

        # Header
        header = ctk.CTkFrame(self, fg_color=TRACKMAN_ORANGE, corner_radius=0, height=90)
        header.pack(fill="x")

        title = ctk.CTkLabel(header, text="TrackMan Report Converter", font=("Segoe UI", 28, "bold"), text_color="white")
        title.pack(pady=25)

        # Content
        content = ctk.CTkFrame(self, fg_color=DARK_BG)
        content.pack(expand=True, fill="both")

        desc = ctk.CTkLabel(content,
            text="Automatically fetches your latest TrackMan report\nand converts it to a formatted Excel file.",
            font=("Segoe UI", 16), text_color="#BBBBBB", justify="center"
        )
        desc.pack(pady=(60, 30))

        # Buttons
        self.convert_btn = ctk.CTkButton(
            content, text="Download & Convert Latest Report",
            width=320, height=60, corner_radius=12,
            font=("Segoe UI", 18, "bold"),
            fg_color=TRACKMAN_ORANGE, hover_color="#FF8533",
            text_color="white", command=self.handle_cloud
        )
        self.convert_btn.pack(pady=10)

        self.token_btn = ctk.CTkButton(
            content, text="🔑 Paste TrackMan Token Manually",
            width=320, height=40, corner_radius=10,
            font=("Segoe UI", 15),
            fg_color="#444444", hover_color="#666666",
            text_color="white", command=self.manual_token_entry
        )
        self.token_btn.pack(pady=8)

        self.status_label = ctk.CTkLabel(content, text="", font=("Segoe UI", 14), text_color="#AAAAAA")
        self.status_label.pack(pady=(30, 10))

        footer = ctk.CTkLabel(self, text=APP_FOOTER_TEXT, font=("Segoe UI", 11, "italic"), text_color="#666666")
        footer.pack(side="bottom", pady=8)

    # Overlay methods
    def show_overlay(self, text="Loading..."):
        if self.overlay:
            self.overlay.destroy()
        self.overlay = LoadingOverlay(self, text)
        self.overlay.update()

    def hide_overlay(self):
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None

    # Manual token entry
    def manual_token_entry(self):
        import tkinter.simpledialog as sd
        token_input = sd.askstring("Enter TrackMan Token", "Paste your TrackMan Bearer token below:")
        if not token_input:
            messagebox.showinfo("Cancelled", "No token entered.")
            return

        token_clean = token_input.strip()
        if token_clean.lower().startswith("bearer "):
            token_clean = token_clean.split(" ", 1)[1].strip()

        try:
            with open("trackman_token.txt", "w", encoding="utf-8") as f:
                f.write(token_clean)
            messagebox.showinfo("Success", "✅ Token saved successfully!\nYou can now download reports.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save token:\n{e}")

    # Download + Convert handler
    def handle_cloud(self):
        try:
            self.show_overlay("Checking TrackMan login...")
            token = trackman_auth.get_saved_token() or trackman_auth.login_via_browser()
            if not token:
                raise Exception("Could not retrieve TrackMan token.")

            self.overlay.update_text("Detecting latest TrackMan report in Chrome...")
            report_id = get_latest_report_id_from_chrome()
            if not report_id:
                raise Exception("Couldn't find any recent TrackMan report in Chrome. Open one in your browser first.")

            self.overlay.update_text("Downloading report data...")
            json_path = download_report(token, report_id)

            self.overlay.update_text("Converting to formatted Excel...")
            result = convert_json_to_excel(json_path)

            self.hide_overlay()
            messagebox.showinfo("Success", f"Downloaded and converted!\nSaved as:\n{result}")
        except Exception as e:
            self.hide_overlay()
            messagebox.showerror("Error", str(e))

# -----------------------------
# RUN APP
# -----------------------------
if __name__ == "__main__":
    app = TrackmanApp()
    app.mainloop()
