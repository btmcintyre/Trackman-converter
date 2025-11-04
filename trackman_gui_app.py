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
# =========================================
# Constants
# =========================================
APP_FOOTER_TEXT = "© 2025 TrackMan Converter by Tom McIntyre"
TRACKMAN_ORANGE = "#FF6600"
DARK_BG = "#1E1E1E"

HEADER_FILL = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
ALT_FILL = PatternFill(start_color="E8F0FA", end_color="E8F0FA", fill_type="solid")
BORDER = Border(left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin"))


COLUMNS = [
    "Time",
    "Club Speed (Mph)", "Ball Speed (Mph)", "Smash Factor",
    "Carry (Yds)", "Total (Yds)",
    "Impact Height (mm)", "Impact Offset (mm)",
    "Club Path (Deg)", "Face Angle (Deg)", "Face To Path (Deg)",
    "Launch Direction (Deg)", "Attack Angle (Deg)",
    "Dynamic Loft (Deg)", "Launch Angle (Deg)", "Spin Loft (Deg)",
    "Spin Rate (Rpm)", "Spin Axis (Deg)",
    "Curve (Ft)", "Carry Side (Ft)", "Total Side (Ft)",
    "Max Height (Ft)", "Landing Angle (Deg)",
    "Swing Direction (Deg)", "Swing Plane (Deg)", "Swing Radius",
    "DPlane Tilt", "Low Point (In)", "Landing Height", "Hang Time (Sec)",
    "Dynamic Lie (Deg)"
]


# =========================================
# Helpers
# =========================================
def _fmt_time(iso: str) -> str:
    if not iso:
        return ""
    try:
        dt = datetime.fromisoformat(iso.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        return iso


def _conv(v, factor=1.0) -> str:
    if v is None:
        return ""
    try:
        return f"{float(v) * factor:.2f}"
    except Exception:
        return ""


def convert_measurement_to_row(m: dict) -> dict:
    if not m:
        return {k: "" for k in COLUMNS}
    return {
        "Time": _fmt_time(m.get("Time", "")),
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


# =========================================
# Sheet Styling + Stats
# =========================================
def style_and_finalize_sheet(ws, header_row_idx: int, n_cols: int, n_rows: int):
    # Header
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=header_row_idx, column=c)
        cell.font = Font(bold=True, color="000000")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="thin", color="888888"))

    # Alternating fill
    data_start = header_row_idx + 1
    data_end = header_row_idx + n_rows
    for r in range(data_start, data_end + 1):
        fill = ALT_FILL if (r - data_start) % 2 == 0 else None
        for c in range(1, n_cols + 1):
            if fill:
                ws.cell(row=r, column=c).fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")

    ws.freeze_panes = f"A{header_row_idx + 1}"
    last_col_letter = get_column_letter(n_cols)
    ws.auto_filter.ref = f"A{header_row_idx}:{last_col_letter}{data_end}"


def add_summary_rows(ws, df: pd.DataFrame):
    """Adds Pos/Neg/Av/Spread/% rows based on dynamic filters"""
    if df.empty:
        return

    def meets_filter(row):
        try:
            return (
                float(row["Smash Factor"]) >= 1.45 and
                abs(float(row["Impact Height (mm)"])) <= 10 and
                abs(float(row["Impact Offset (mm)"])) <= 10 and
                abs(float(row["Club Path (Deg)"])) <= 4 and
                abs(float(row["Face Angle (Deg)"])) <= 2
            )
        except Exception:
            return False

    df_numeric = df.apply(pd.to_numeric, errors="ignore")
    pos_df = df_numeric[df_numeric.apply(meets_filter, axis=1)]
    neg_df = df_numeric[~df_numeric.apply(meets_filter, axis=1)]

    rows_to_add = {
        "Pos Av": pos_df.mean(numeric_only=True),
        "Neg Av": neg_df.mean(numeric_only=True),
        "Av": df_numeric.mean(numeric_only=True),
        "Spread": df_numeric.max(numeric_only=True) - df_numeric.min(numeric_only=True),
        "% Pos": (len(pos_df) / len(df_numeric) * 100) if len(df_numeric) else 0,
        "% Neg": (len(neg_df) / len(df_numeric) * 100) if len(df_numeric) else 0,
    }

    start_row = ws.max_row + 2
    for label, vals in rows_to_add.items():
        ws.cell(row=start_row, column=1, value=label)
        ws.cell(row=start_row, column=1).font = Font(bold=True)
        if isinstance(vals, pd.Series):
            for i, v in enumerate(vals.tolist(), start=2):
                if isinstance(v, (int, float)):
                    ws.cell(row=start_row, column=i, value=round(v, 2))
        elif isinstance(vals, (int, float)):
            ws.cell(row=start_row, column=2, value=round(vals, 2))
        start_row += 1


# =========================================
# Workbook Build
# =========================================
def build_workbook_per_club(data: dict) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    stroke_groups = data.get("StrokeGroups", []) or []

    for g in stroke_groups:
        club = str(g.get("Club", "Unknown Club"))
        ws = wb.create_sheet(title=club[:31])
        rows = [convert_measurement_to_row(s.get("Measurement")) for s in g.get("Strokes", []) if s.get("Measurement")]
        if not rows:
            continue

        df = pd.DataFrame(rows, columns=COLUMNS)
        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

        style_and_finalize_sheet(ws, 1, len(COLUMNS), len(df.index))
        add_summary_rows(ws, df)

    # All Data (unchanged)
    all_rows = [convert_measurement_to_row(s.get("Measurement"))
                for g in stroke_groups for s in g.get("Strokes", []) if s.get("Measurement")]
    if all_rows:
        ws_all = wb.create_sheet("All Data")
        df_all = pd.DataFrame(all_rows, columns=COLUMNS)
        for r in dataframe_to_rows(df_all, index=False, header=True):
            ws_all.append(r)
        style_and_finalize_sheet(ws_all, 1, len(COLUMNS), len(df_all.index))

    return wb


# =========================================
# JSON → Excel Conversion
# =========================================
def convert_json_to_excel(json_path: str) -> Path:
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    wb = build_workbook_per_club(data)

    default_name = f"Trackman_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    save_path = filedialog.asksaveasfilename(
        title="Save Converted Excel File",
        defaultextension=".xlsx",
        initialfile=default_name,
        initialdir=str(Path.home() / "Documents"),
        filetypes=[("Excel Files", "*.xlsx")],
    )

    if not save_path:
        messagebox.showinfo("Cancelled", "Save cancelled. File not created.")
        return None

    wb.save(save_path)
    return save_path


# =========================================
# GUI
# =========================================
class TrackmanApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("TrackMan Converter")
        self.geometry("700x480")
        self.resizable(False, False)
        self.configure(fg_color=DARK_BG)
        self.overlay = None

        # Header
        header = ctk.CTkFrame(self, fg_color=TRACKMAN_ORANGE, corner_radius=0, height=90)
        header.pack(fill="x")

        title = ctk.CTkLabel(header, text="TrackMan Report Converter",
                             font=("Segoe UI", 28, "bold"), text_color="white")
        title.pack(pady=25)

        # Content
        content = ctk.CTkFrame(self, fg_color=DARK_BG)
        content.pack(expand=True, fill="both")

        desc = ctk.CTkLabel(content,
                            text="Automatically fetches your latest TrackMan report\nand converts it to a formatted Excel file.",
                            font=("Segoe UI", 16), text_color="#BBBBBB", justify="center")
        desc.pack(pady=(60, 30))

        self.convert_btn = ctk.CTkButton(content,
                                         text="☁️  Download & Convert Latest Report",
                                         width=320, height=60, corner_radius=12,
                                         font=("Segoe UI", 18, "bold"),
                                         fg_color=TRACKMAN_ORANGE, hover_color="#FF8533",
                                         text_color="white", command=self.handle_cloud)
        self.convert_btn.pack(pady=10)

        self.status_label = ctk.CTkLabel(content, text="", font=("Segoe UI", 14), text_color="#AAAAAA")
        self.status_label.pack(pady=(30, 10))

        footer = ctk.CTkLabel(self, text=APP_FOOTER_TEXT,
                              font=("Segoe UI", 11, "italic"), text_color="#666666")
        footer.pack(side="bottom", pady=8)

    # -----------------------------
    def show_overlay(self, text="Loading..."):
        if self.overlay:
            self.overlay.destroy()
        self.overlay = LoadingOverlay(self, text)
        self.overlay.update()

    def hide_overlay(self):
        if self.overlay:
            self.overlay.destroy()
            self.overlay = None

    # -----------------------------
    def handle_cloud(self):
        try:
            self.show_overlay("🔐 Checking TrackMan login...")
            token = trackman_auth.get_saved_token() or trackman_auth.login_via_browser()
            if not token:
                raise Exception("Could not retrieve TrackMan token.")

            self.overlay.update_text("🔍 Detecting latest TrackMan report in Chrome...")
            report_id = get_latest_report_id_from_chrome()
            if not report_id:
                raise Exception("Couldn't find any recent TrackMan report in Chrome. Open one in your browser first.")

            self.overlay.update_text("📡 Downloading report data...")
            json_path = download_report(token, report_id)

            self.overlay.update_text("📊 Converting to formatted Excel...")
            result = convert_json_to_excel(json_path)

            self.hide_overlay()
            messagebox.showinfo("Success", f"✅ Downloaded and converted!\nSaved as:\n{result}")

        except Exception as e:
            self.hide_overlay()
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    app = TrackmanApp()
    app.mainloop()
