import streamlit as st
import pandas as pd
import io
import os
import sys
import base64
import datetime
import subprocess
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_DATE_DDMMYY

# ── Patch Streamlit's index.html to embed manifest link in raw HTML ───────────
# This runs once per Streamlit session restart, before any page is served.
# It is idempotent and resilient to package reinstalls.
try:
    import patch_pwa
    patch_pwa.patch()
except Exception:
    pass

st.set_page_config(
    page_title="EngiTrack Timesheet Combiner",
    page_icon="📋",
    layout="wide"
)

# ── Service Worker registration (blob URL so it can claim root scope) ─────────
st.markdown("""
<script>
(function () {
  var SW_CODE = [
    'var CACHE="engitrack-v2";',
    'self.addEventListener("install",function(e){self.skipWaiting();});',
    'self.addEventListener("activate",function(e){',
    '  e.waitUntil(caches.keys().then(function(ks){',
    '    return Promise.all(ks.filter(function(k){return k!==CACHE;}).map(function(k){return caches.delete(k);}));',
    '  }));',
    '  return self.clients.claim();',
    '});',
    'self.addEventListener("fetch",function(e){',
    '  if(e.request.method!=="GET")return;',
    '  e.respondWith(fetch(e.request).catch(function(){',
    '    return new Response("Offline \u2013 please reconnect.",{status:503,headers:{"Content-Type":"text/plain"}});',
    '  }));',
    '});'
  ].join("\\n");

  if ("serviceWorker" in navigator) {
    try {
      var blob = new Blob([SW_CODE], { type:"application/javascript" });
      var swUrl = URL.createObjectURL(blob);
      navigator.serviceWorker.register(swUrl, { scope:"/" })
        .then(function(){ console.log("EngiTrack SW registered"); })
        .catch(function(err){ console.log("EngiTrack SW:", err.message); });
    } catch(e) { console.log("SW setup:", e); }
  }
})();
</script>
""", unsafe_allow_html=True)

st.title("EngiTrack Timesheet Combiner")
st.markdown("Upload up to **30 engineer timesheet CSVs** from your EngiTrack app to combine them into a single Excel workbook.")

st.divider()

# ── Helpers ──────────────────────────────────────────────────────────────────

def to_excel_serial(date_val):
    """Convert a date to Excel date serial number."""
    if isinstance(date_val, str):
        try:
            date_val = datetime.date.fromisoformat(date_val)
        except Exception:
            return None
    if isinstance(date_val, datetime.datetime):
        date_val = date_val.date()
    if isinstance(date_val, datetime.date):
        return (date_val - datetime.date(1899, 12, 30)).days
    return None


def sanitise_sheet_name(name):
    for ch in r"/\?*[]:'":
        name = name.replace(ch, "-")
    return name[:31].strip()


def parse_engitrack_csv(file_obj):
    """
    Parse an EngiTrack CSV and return a dict with all extracted fields.
    Layout:
      Row 0 : title
      Row 1 : Engineer, <name>
      Row 2 : Week Commencing, <date>
      Row 3 : blank
      Row 4 : column headers (Date, Day, ...)
      Row 5-11: daily rows
      Row 12: blank
      Row 13: Weekly Totals (label)
      Row 14+: key,value pairs for totals (some values are multi-line quoted)
    """
    import csv as _csv

    raw = file_obj.read()
    if isinstance(raw, bytes):
        raw = raw.decode("utf-8", errors="replace")

    # Use csv.reader so multi-line quoted fields are handled correctly
    reader = list(_csv.reader(io.StringIO(raw)))

    engineer_name = "Unknown"
    week_commencing_str = ""

    for row in reader[:5]:
        if len(row) >= 2:
            key = row[0].strip().lower()
            val = row[1].strip()
            if key == "engineer":
                engineer_name = val
            elif key == "week commencing":
                week_commencing_str = val

    # Parse daily rows (skip 4 header lines, stop at blank or non-date row)
    daily_df = pd.read_csv(io.StringIO(raw), skiprows=4, on_bad_lines="skip")
    if "Date" in daily_df.columns:
        daily_df = daily_df[daily_df["Date"].astype(str).str.match(r"\d{4}-\d{2}-\d{2}")]
    else:
        daily_df = daily_df.iloc[:7]

    # Parse weekly totals section using csv.reader rows
    # (multi-line quoted fields are already merged by csv.reader)
    totals = {}
    in_totals = False
    for row in reader:
        if not row:
            continue
        cell0 = row[0].strip()
        if cell0.lower() == "weekly totals":
            in_totals = True
            continue
        if not in_totals:
            continue
        if len(row) >= 2 and cell0:
            totals[cell0] = row[1].strip()
        elif len(row) == 1 and cell0:
            totals[cell0] = ""

    def safe_float(key, default=0.0):
        try:
            return float(totals.get(key, default))
        except (ValueError, TypeError):
            return default

    def safe_int(key, default=0):
        try:
            return int(float(totals.get(key, default)))
        except (ValueError, TypeError):
            return default

    # Count holiday and sick days from daily data
    holiday_days = 0
    sick_days = 0
    if "Holiday" in daily_df.columns:
        holiday_days = int((daily_df["Holiday"].astype(str).str.strip().str.lower() == "yes").sum())
    if "Sickness" in daily_df.columns:
        sick_days = int((daily_df["Sickness"].astype(str).str.strip().str.lower() == "yes").sum())

    return {
        "engineer_name": engineer_name,
        "week_commencing_str": week_commencing_str,
        "total_hours": safe_float("Total Hours"),
        "standard_hours": safe_float("Standard Hours (Mon-Fri)"),
        "weekday_ot": safe_float("Weekday Overtime"),
        "saturday_hours": safe_float("Saturday Hours"),
        "sunday_hours": safe_float("Sunday Hours"),
        "bh_hours": safe_float("Bank Holiday Hours"),
        "holiday_days": holiday_days,
        "sick_days": sick_days,
        "repairs_logged": safe_int("Repairs Logged"),
        "repairs_notes": totals.get("Repairs Notes", ""),
        "extra_jobs_logged": safe_int("Extra Jobs Logged"),
        "extra_jobs_notes": totals.get("Extra Jobs Notes", ""),
        "daily_df": daily_df,
    }


# ── Styles ───────────────────────────────────────────────────────────────────

NAV   = "1F4E79"   # dark navy
MID   = "2E75B6"   # mid blue
LIGHT = "D6E4F0"   # light blue row alt
WHITE = "FFFFFF"
GOLD  = "FFD700"
GREY  = "F2F2F2"

def hdr_cell(ws, row, col, value, bg=NAV, fg=WHITE, bold=True, center=True, size=11):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Calibri", bold=bold, color=fg, size=size)
    c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
    c.alignment = Alignment(horizontal="center" if center else "left",
                            vertical="center", wrap_text=True)
    c.border = thin_border()
    return c


def thin_border():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)


def data_cell(ws, row, col, value, alt=False, wrap=False, number_format=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = Font(name="Calibri", size=10)
    if alt:
        c.fill = PatternFill(start_color=LIGHT, end_color=LIGHT, fill_type="solid")
    c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=wrap)
    c.border = thin_border()
    if number_format:
        c.number_format = number_format
    return c


def set_col_width(ws, col_letter, width):
    ws.column_dimensions[col_letter].width = width


# ── Workbook builder ─────────────────────────────────────────────────────────

def build_workbook(timesheets):
    wb = Workbook()
    wb.remove(wb.active)

    _build_instructions(wb)
    _build_engineer_list(wb, timesheets)
    _build_week_list(wb, timesheets)
    _build_weekly_summary(wb, timesheets)
    _build_weekly_data(wb, timesheets)
    _build_daily_data(wb, timesheets)

    return wb


def _build_instructions(wb):
    ws = wb.create_sheet("Instructions")
    ws.sheet_view.showGridLines = False

    lines = [
        ("EngiTrack Weekly Timesheet Combiner — up to 30 engineers", True, 14, NAV, WHITE),
        ("", False, 11, WHITE, "000000"),
        ("How to use", True, 12, MID, WHITE),
        ("1)  Upload your EngiTrack weekly CSV exports in the app above.", False, 11, GREY, "000000"),
        ("2)  Click Combine — the workbook is generated automatically.", False, 11, WHITE, "000000"),
        ("3)  Open Weekly_Summary — each week's engineers and totals are pre-filled.", False, 11, GREY, "000000"),
        ("4)  WeeklyData and DailyData hold the raw row-level data for deeper analysis.", False, 11, WHITE, "000000"),
        ("", False, 11, WHITE, "000000"),
        ("Notes", True, 12, MID, WHITE),
        ("•  Weekly_Summary groups engineers by week with auto-calculated column totals.", False, 11, GREY, "000000"),
        ("•  WeeklyData holds one row per engineer per week.", False, 11, WHITE, "000000"),
        ("•  DailyData holds one row per day per engineer.", False, 11, GREY, "000000"),
    ]

    for i, (text, bold, size, bg, fg) in enumerate(lines, 1):
        c = ws.cell(row=i, column=1, value=text)
        c.font = Font(name="Calibri", bold=bold, size=size, color=fg)
        c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        c.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws.row_dimensions[i].height = 20 if text else 8

    set_col_width(ws, "A", 80)


def _build_engineer_list(wb, timesheets):
    ws = wb.create_sheet("Engineer_List")
    ws.sheet_view.showGridLines = False

    hdr_cell(ws, 1, 1, "Engineer")
    set_col_width(ws, "A", 30)
    ws.row_dimensions[1].height = 22

    seen = []
    for ts in timesheets:
        name = ts["engineer_name"]
        if name not in seen:
            seen.append(name)

    for i, name in enumerate(seen[:30], 2):
        c = ws.cell(row=i, column=1, value=name)
        c.font = Font(name="Calibri", size=10)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = thin_border()
        if i % 2 == 0:
            c.fill = PatternFill(start_color=LIGHT, end_color=LIGHT, fill_type="solid")


def _build_week_list(wb, timesheets):
    ws = wb.create_sheet("Week_List")
    ws.sheet_view.showGridLines = False

    hdr_cell(ws, 1, 1, "Week Commencing (YYYY-MM-DD)")
    set_col_width(ws, "A", 30)
    ws.row_dimensions[1].height = 22

    seen = []
    for ts in timesheets:
        wc = ts["week_commencing_str"]
        if wc and wc not in seen:
            seen.append(wc)
    seen.sort()

    for i, wc in enumerate(seen, 2):
        c = ws.cell(row=i, column=1, value=wc)
        c.font = Font(name="Calibri", size=10)
        c.alignment = Alignment(horizontal="left", vertical="center")
        c.border = thin_border()
        if i % 2 == 0:
            c.fill = PatternFill(start_color=LIGHT, end_color=LIGHT, fill_type="solid")


def _build_weekly_summary(wb, timesheets):
    ws = wb.create_sheet("Weekly_Summary")
    ws.sheet_view.showGridLines = False

    col_headers = ["Engineer", "Total Hours", "Standard Hours", "Weekday OT",
                   "Sat Hours", "Sun Hours", "BH Hours", "Holiday Days",
                   "Sick Days", "Repairs", "Extra Jobs"]
    col_widths   = [28, 13, 16, 13, 11, 11, 11, 14, 11, 10, 11]
    num_cols = len(col_headers)

    for col_idx, w in enumerate(col_widths, 1):
        set_col_width(ws, get_column_letter(col_idx), w)

    # Group timesheets by week commencing (preserve insertion order)
    weeks = {}
    for ts in timesheets:
        wc = ts["week_commencing_str"]
        weeks.setdefault(wc, []).append(ts)

    # Main title
    ws.merge_cells(f"A1:{get_column_letter(num_cols)}1")
    c = ws.cell(row=1, column=1, value="Weekly Summary")
    c.font = Font(name="Calibri", bold=True, size=16, color=WHITE)
    c.fill = PatternFill(start_color=NAV, end_color=NAV, fill_type="solid")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    current_row = 2

    for week_idx, (wc_str, week_ts) in enumerate(weeks.items()):
        # Gap between weeks
        if week_idx > 0:
            ws.row_dimensions[current_row].height = 10
            current_row += 1

        # Week heading row
        try:
            wc_date = datetime.date.fromisoformat(wc_str)
            wc_label = wc_date.strftime("Week commencing  %d %B %Y")
        except Exception:
            wc_label = f"Week commencing  {wc_str}"

        ws.merge_cells(f"A{current_row}:{get_column_letter(num_cols)}{current_row}")
        wh = ws.cell(row=current_row, column=1, value=wc_label)
        wh.font = Font(name="Calibri", bold=True, size=12, color=WHITE)
        wh.fill = PatternFill(start_color=MID, end_color=MID, fill_type="solid")
        wh.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        wh.border = thin_border()
        ws.row_dimensions[current_row].height = 26
        current_row += 1

        # Column headers
        for col_idx, h in enumerate(col_headers, 1):
            hdr_cell(ws, current_row, col_idx, h, bg="2E75B6", size=10)
        ws.row_dimensions[current_row].height = 24
        data_start_row = current_row + 1
        current_row += 1

        # Accumulate totals
        totals = [0.0] * (num_cols - 1)

        # Engineer rows
        for eng_idx, ts in enumerate(week_ts):
            alt = (eng_idx % 2 == 0)
            row_vals = [
                ts["engineer_name"],
                ts["total_hours"],
                ts["standard_hours"],
                ts["weekday_ot"],
                ts["saturday_hours"],
                ts["sunday_hours"],
                ts["bh_hours"],
                ts["holiday_days"],
                ts["sick_days"],
                ts["repairs_logged"],
                ts["extra_jobs_logged"],
            ]
            for col_idx, val in enumerate(row_vals, 1):
                c = ws.cell(row=current_row, column=col_idx, value=val)
                c.font = Font(name="Calibri", size=10,
                              bold=(col_idx == 1))
                c.fill = PatternFill(
                    start_color=LIGHT if alt else WHITE,
                    end_color=LIGHT if alt else WHITE,
                    fill_type="solid")
                c.alignment = Alignment(
                    horizontal="left" if col_idx == 1 else "center",
                    vertical="center")
                c.border = thin_border()
                if col_idx > 1:
                    try:
                        totals[col_idx - 2] += float(val or 0)
                    except (TypeError, ValueError):
                        pass
            ws.row_dimensions[current_row].height = 18
            current_row += 1

        # Totals row
        tc = ws.cell(row=current_row, column=1, value="TOTAL")
        tc.font = Font(name="Calibri", bold=True, size=11, color=WHITE)
        tc.fill = PatternFill(start_color=NAV, end_color=NAV, fill_type="solid")
        tc.alignment = Alignment(horizontal="center", vertical="center")
        tc.border = thin_border()

        for col_idx, total_val in enumerate(totals, 2):
            display = int(total_val) if total_val == int(total_val) else round(total_val, 2)
            c = ws.cell(row=current_row, column=col_idx, value=display)
            c.font = Font(name="Calibri", bold=True, size=11, color=WHITE)
            c.fill = PatternFill(start_color=NAV, end_color=NAV, fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = thin_border()

        ws.row_dimensions[current_row].height = 22
        current_row += 1

    ws.freeze_panes = "A2"


def _build_weekly_data(wb, timesheets):
    ws = wb.create_sheet("WeeklyData")

    headers = [
        "Engineer", "Week Commencing", "Total Hours", "Standard Hours (Mon-Fri)",
        "Weekday Overtime", "Saturday Hours", "Sunday Hours", "Bank Holiday Hours",
        "Holiday Days", "Sickness Days", "Repairs Logged", "Repairs Notes",
        "Extra Jobs Logged", "Extra Jobs Notes", "Source File", "Key (Engineer|Serial)"
    ]
    col_widths = [28, 18, 13, 22, 16, 15, 13, 18, 13, 13, 14, 30, 15, 30, 36, 30]

    for col_idx, (h, w) in enumerate(zip(headers, col_widths), 1):
        hdr_cell(ws, 1, col_idx, h, size=10)
        set_col_width(ws, get_column_letter(col_idx), w)
    ws.row_dimensions[1].height = 22

    for row_idx, ts in enumerate(timesheets, 2):
        wc = ts["week_commencing_str"]
        try:
            wc_date = datetime.date.fromisoformat(wc)
        except Exception:
            wc_date = None

        serial = to_excel_serial(wc_date) if wc_date else ""
        key = f"{ts['engineer_name']}|{serial}" if serial else ""
        alt = (row_idx % 2 == 0)

        row_values = [
            ts["engineer_name"],
            wc_date,
            ts["total_hours"],
            ts["standard_hours"],
            ts["weekday_ot"],
            ts["saturday_hours"],
            ts["sunday_hours"],
            ts["bh_hours"],
            ts["holiday_days"],
            ts["sick_days"],
            ts["repairs_logged"],
            ts["repairs_notes"],
            ts["extra_jobs_logged"],
            ts["extra_jobs_notes"],
            ts.get("source_file", ""),
            key,
        ]

        for col_idx, val in enumerate(row_values, 1):
            c = data_cell(ws, row_idx, col_idx, val, alt=alt,
                          wrap=(col_idx in (12, 14)))
            if col_idx == 2 and wc_date:
                c.number_format = "YYYY-MM-DD"

        ws.row_dimensions[row_idx].height = 18

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


def _build_daily_data(wb, timesheets):
    ws = wb.create_sheet("DailyData")

    headers = [
        "Engineer", "Week Commencing", "Date", "Day",
        "Start Time", "End Time", "Total Hours",
        "Bank Holiday", "Holiday", "Sickness", "Source File"
    ]
    col_widths = [28, 18, 13, 12, 11, 11, 13, 14, 10, 10, 36]

    for col_idx, (h, w) in enumerate(zip(headers, col_widths), 1):
        hdr_cell(ws, 1, col_idx, h, size=10)
        set_col_width(ws, get_column_letter(col_idx), w)
    ws.row_dimensions[1].height = 22

    row_idx = 2
    for ts in timesheets:
        daily_df = ts["daily_df"]
        wc = ts["week_commencing_str"]
        engineer = ts["engineer_name"]
        source = ts.get("source_file", "")

        for _, day_row in daily_df.iterrows():
            alt = (row_idx % 2 == 0)
            date_val = str(day_row.get("Date", "")).strip()
            try:
                date_obj = datetime.date.fromisoformat(date_val)
            except Exception:
                date_obj = date_val

            row_values = [
                engineer,
                wc,
                date_obj,
                str(day_row.get("Day", "")).strip(),
                str(day_row.get("Start Time", "")).strip(),
                str(day_row.get("End Time", "")).strip(),
                day_row.get("Total Hours", 0),
                str(day_row.get("Bank Holiday", "")).strip(),
                str(day_row.get("Holiday", "")).strip(),
                str(day_row.get("Sickness", "")).strip(),
                source,
            ]

            for col_idx, val in enumerate(row_values, 1):
                c = data_cell(ws, row_idx, col_idx, val, alt=alt)
                if col_idx == 3 and isinstance(date_obj, datetime.date):
                    c.number_format = "YYYY-MM-DD"

            ws.row_dimensions[row_idx].height = 16
            row_idx += 1

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"


# ── UI ────────────────────────────────────────────────────────────────────────

uploaded_files = st.file_uploader(
    "Select timesheet CSV files",
    type=["csv"],
    accept_multiple_files=True,
    help="Select all the EngiTrack CSV timesheets you want to combine (max 30 files)"
)

if uploaded_files:
    if len(uploaded_files) > 30:
        st.error(f"Too many files selected ({len(uploaded_files)}). Please select a maximum of 30 timesheets.")
        st.stop()

    timesheets = []
    sheet_errors = []
    seen_names = {}

    for f in uploaded_files:
        try:
            ts = parse_engitrack_csv(f)
            ts["source_file"] = f.name

            base = sanitise_sheet_name(ts["engineer_name"])
            if base in seen_names:
                seen_names[base] += 1
                ts["sheet_name"] = sanitise_sheet_name(f"{ts['engineer_name']} ({seen_names[base]})")
            else:
                seen_names[base] = 1
                ts["sheet_name"] = base

            timesheets.append(ts)
        except Exception as e:
            sheet_errors.append((f.name, str(e)))

    for fname, err in sheet_errors:
        st.warning(f"Could not read **{fname}**: {err}")

    if timesheets:
        st.success(f"{len(timesheets)} timesheet(s) loaded.")
        st.subheader("Preview")

        tab_labels = [ts["engineer_name"] for ts in timesheets]
        tabs = st.tabs(tab_labels)

        for tab, ts in zip(tabs, timesheets):
            with tab:
                wc = ts["week_commencing_str"]
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Total Hours", ts["total_hours"])
                c2.metric("Standard Hours", ts["standard_hours"])
                c3.metric("Weekday OT", ts["weekday_ot"])
                c4.metric("Week Commencing", wc)
                c5, c6, c7, c8 = st.columns(4)
                c5.metric("Sat Hours", ts["saturday_hours"])
                c6.metric("Sun Hours", ts["sunday_hours"])
                c7.metric("Holiday Days", ts["holiday_days"])
                c8.metric("Sick Days", ts["sick_days"])
                st.dataframe(ts["daily_df"], use_container_width=True, height=260)

        st.divider()

        col1, col2 = st.columns([2, 1])
        with col1:
            workbook_name = st.text_input(
                "Output workbook filename",
                value="EngiTrack_Combined",
                help="Name for the downloaded Excel workbook (no extension needed)"
            )

        current_file_key = tuple(sorted(f.name for f in uploaded_files))
        if st.session_state.get("_wb_file_key") != current_file_key:
            st.session_state.pop("_wb_bytes", None)
            st.session_state.pop("_wb_filename", None)
            st.session_state["_wb_file_key"] = current_file_key

        if st.button("Combine into Excel Workbook", type="primary", use_container_width=True):
            with st.spinner("Building workbook..."):
                wb = build_workbook(timesheets)
                buffer = io.BytesIO()
                wb.save(buffer)
                safe_name = workbook_name.strip().replace(" ", "_") or "EngiTrack_Combined"
                st.session_state["_wb_bytes"] = buffer.getvalue()
                st.session_state["_wb_filename"] = f"{safe_name}.xlsx"

        if "_wb_bytes" in st.session_state:
            wb_bytes = st.session_state["_wb_bytes"]
            fname = st.session_state["_wb_filename"]

            if getattr(sys, "frozen", False):
                downloads_dir = Path.home() / "Downloads"
                downloads_dir.mkdir(parents=True, exist_ok=True)
                output_path = downloads_dir / fname
                output_path.write_bytes(wb_bytes)
                st.success(
                    f"Workbook saved to your Downloads folder:\n\n"
                    f"**{output_path}**"
                )
                if st.button("Open file location", key="open_folder"):
                    subprocess.Popen(
                        ["explorer", "/select,", str(output_path)]
                    )
            else:
                b64 = base64.b64encode(wb_bytes).decode()
                mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                st.markdown(
                    f"""
                    <a href="data:{mime};base64,{b64}" download="{fname}"
                       style="display:block;width:100%;padding:0.6rem 1rem;
                              background:#FF4B4B;color:white;text-align:center;
                              font-weight:600;font-size:1rem;border-radius:0.5rem;
                              text-decoration:none;">
                        ⬇ Download {fname}
                    </a>
                    """,
                    unsafe_allow_html=True,
                )
                st.success(
                    f"Workbook ready — {len(timesheets)} engineer(s) across 6 sheets "
                    f"(Instructions, Engineer_List, Week_List, Weekly_Summary, WeeklyData, DailyData)."
                )

else:
    st.info("Upload your EngiTrack CSV timesheet files above to get started.")
    st.markdown("""
    **How it works:**
    1. Upload your EngiTrack CSV files (up to 30)
    2. Check the preview — each tab shows an engineer's hours summary and daily breakdown
    3. Click **Combine into Excel Workbook** to generate the workbook
    4. Open **Weekly_Summary** in the downloaded file and pick any week to see all engineers at a glance

    **Output sheets:**  Instructions · Engineer_List · Week_List · Weekly_Summary · WeeklyData · DailyData
    """)
