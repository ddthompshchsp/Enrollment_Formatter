# app.py
from datetime import datetime, date
from zoneinfo import ZoneInfo

import re
import pandas as pd
import streamlit as st
from PIL import Image

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel
from openpyxl.drawing.image import Image as XLImage  # for embedding the PNG chart

import matplotlib
matplotlib.use("Agg")  # headless backend for servers/Streamlit
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, PercentFormatter

st.set_page_config(page_title="Enrollment Formatter", layout="centered")

# ---------------- Header / UI ----------------
try:
    logo = Image.open("header_logo.png")
    st.image(logo, width=300)
except Exception:
    pass

st.title("HCHSP Enrollment Checklist (2025–2026)")
st.markdown("Upload your **Enrollment.xlsx** file to receive a formatted version.")

uploaded_file = st.file_uploader("Upload Enrollment.xlsx", type=["xlsx"])

# ---------------- Helpers ----------------
def coerce_to_dt(v):
    if pd.isna(v):
        return None
    if isinstance(v, datetime):
        return v
    if isinstance(v, date):
        return datetime(v.year, v.month, v.day)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            return from_excel(v)
        except Exception:
            return None
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return None
        for fmt in ("%m/%d/%Y", "%m-%d-%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(s, fmt)
            except Exception:
                continue
    return None

def most_recent(series):
    dates, texts = [], []
    for v in pd.unique(series.dropna()):
        dt = coerce_to_dt(v)
        if dt:
            dates.append(dt)
        else:
            s = str(v).strip()
            if s:
                texts.append(s)
    if dates:
        return max(dates)
    return texts[0] if texts else None

def normalize(s: str) -> str:
    s = s.lower()
    s = re.sub(r"[\s\-\–\—_:()]+", " ", s)
    return s.strip()

def find_cols(cols, keywords):
    out = []
    for c in cols:
        if not isinstance(c, str):
            continue
        n = normalize(c)
        if any(k in n for k in keywords):
            out.append(c)
    return out

def collapse_row_values(row, col_names):
    vals = []
    for c in col_names:
        if c in row and pd.notna(row[c]) and str(row[c]).strip() != "":
            vals.append(row[c])
    if not vals:
        return None
    dts = [coerce_to_dt(v) for v in vals]
    dts = [d for d in dts if d]
    if dts:
        return max(dts)
    return str(vals[0]).strip()

# ---------------- Main ----------------
if uploaded_file:
    # 1) Find header row via "ST: Participant PID"
    wb_src = load_workbook(uploaded_file, data_only=True)
    ws_src = wb_src.active

    header_row = None
    for row in ws_src.iter_rows(min_row=1, max_row=30):
        for cell in row:
            if isinstance(cell.value, str) and "ST: Participant PID" in cell.value:
                header_row = cell.row
                break
        if header_row:
            break

    if not header_row:
        st.error("Couldn't find 'ST: Participant PID' in the file.")
        st.stop()

    uploaded_file.seek(0)

    # 2) Load & normalize
    df = pd.read_excel(uploaded_file, header=header_row - 1)
    df.columns = [c.replace("ST: ", "") if isinstance(c, str) else c for c in df.columns]

    general_cutoff = datetime(2025, 5, 11)  # other date fields < this => "X"
    field_cutoff   = datetime(2025, 8, 1)   # Immunizations/TB/Lead < this => red date (keep value)

    if "Participant PID" not in df.columns:
        st.error("The file is missing 'Participant PID'.")
        st.stop()

    # One row per PID (most recent across dups)
    df = (
        df.dropna(subset=["Participant PID"])
          .groupby("Participant PID", as_index=False)
          .agg(most_recent)
    )

    # 3) Collapse Immunizations, TB, Lead
    all_cols = list(df.columns)
    immun_cols = find_cols(all_cols, ["immun"])
    tb_cols    = find_cols(all_cols, ["tb", "tuberc", "ppd"])
    lead_cols  = find_cols(all_cols, ["lead", "pb"])

    if immun_cols:
        df["Immunizations"] = df.apply(lambda r: collapse_row_values(r, immun_cols), axis=1)
        df.drop(columns=[c for c in immun_cols if c in df.columns], inplace=True)

    if tb_cols:
        df["TB Test"] = df.apply(lambda r: collapse_row_values(r, tb_cols), axis=1)
        df.drop(columns=[c for c in tb_cols if c in df.columns], inplace=True)

    if lead_cols:
        df["Lead Test"] = df.apply(lambda r: collapse_row_values(r, lead_cols), axis=1)
        df.drop(columns=[c for c in lead_cols if c in df.columns], inplace=True)

    # 4) Write workbook scaffold (basic with pandas + openpyxl styling after)
    title_text = "Enrollment Checklist 2025–2026"
    central_now = datetime.now(ZoneInfo("America/Chicago"))
    timestamp_text = central_now.strftime("Generated on %B %d, %Y at %I:%M %p %Z")

    temp_path = "Enrollment_Cleaned.xlsx"
    with pd.ExcelWriter(temp_path, engine="openpyxl") as writer:
        pd.DataFrame([[title_text]]).to_excel(writer, index=False, header=False, startrow=0)
        pd.DataFrame([[timestamp_text]]).to_excel(writer, index=False, header=False, startrow=1)
        df.to_excel(writer, index=False, startrow=3)

    # 5) Style + rules + dynamic totals
    wb = load_workbook(temp_path)
    ws = wb.active

    filter_row = 4
    data_start = filter_row + 1
    data_end = ws.max_row
    base_max_col = ws.max_column  # remember before helper column

    # Freeze rows 1–4 only
    ws.freeze_panes = "A5"

    # AutoFilter
    ws.auto_filter.ref = f"A{filter_row}:{get_column_letter(base_max_col)}{data_end}"

    # Title & timestamp style
    title_fill = PatternFill(start_color="EFEFEF", end_color="EFEFEF", fill_type="solid")
    ts_fill = PatternFill(start_color="F7F7F7", end_color="F7F7F7", fill_type="solid")

    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=base_max_col)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=base_max_col)

    tcell = ws.cell(row=1, column=1); tcell.value = title_text
    tcell.font = Font(size=14, bold=True)
    tcell.alignment = Alignment(horizontal="center", vertical="center")
    tcell.fill = title_fill

    scell = ws.cell(row=2, column=1); scell.value = timestamp_text
    scell.font = Font(size=10, italic=True, color="555555")
    scell.alignment = Alignment(horizontal="center", vertical="center")
    scell.fill = ts_fill

    # Header style
    header_fill = PatternFill(start_color="305496", end_color="305496", fill_type="solid")
    for cell in ws[filter_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = header_fill

    # Borders / fonts
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    red_font = Font(color="FF0000", bold=True)

    # Column locations
    headers = [ws.cell(row=filter_row, column=c).value for c in range(1, base_max_col + 1)]

    def find_idx_exact(name):
        for i, h in enumerate(headers, start=1):
            if h == name:
                return i
        return None

    def find_idx_sub(sub):
        for i, h in enumerate(headers, start=1):
            if isinstance(h, str) and sub in h.lower():
                return i
        return None

    name_col_idx = next((i for i, h in enumerate(headers, 1)
                         if isinstance(h, str) and "name" in h.lower()), 2)
    immun_idx = find_idx_exact("Immunizations") or find_idx_sub("immun")
    tb_idx    = find_idx_exact("TB Test")       or find_idx_sub("tb")
    lead_idx  = find_idx_exact("Lead Test")     or find_idx_sub("lead")

    # Find center/campus column for summary
    center_idx = next((i for i, h in enumerate(headers, 1)
                       if isinstance(h, str) and ("center" in h.lower() or "campus" in h.lower() or "school" in h.lower())), None)

    # Clean any stray "Filtered Total"
    for r in range(1, ws.max_row + 1):
        for c in range(1, base_max_col + 1):
            v = ws.cell(row=r, column=c).value
            if isinstance(v, str) and "filtered total" in v.lower():
                ws.cell(row=r, column=c).value = None

    # Apply cell rules
    for r in range(data_start, data_end + 1):
        for c in range(1, base_max_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            cell.border = thin_border

            if val in (None, "", "nan", "NaT"):
                cell.value = "X"
                cell.font = red_font
                continue

            dt = coerce_to_dt(val)

            # Immunizations: keep info; red date if < 8/1/2025
            if immun_idx and c == immun_idx:
                if dt:
                    cell.value = dt
                    cell.number_format = "m/d/yy"
                    if dt < field_cutoff:
                        cell.font = red_font
                continue

            # TB: keep info; red date if < 8/1/2025
            if tb_idx and c == tb_idx:
                if dt:
                    cell.value = dt
                    cell.number_format = "m/d/yy"
                    if dt < field_cutoff:
                        cell.font = red_font
                continue

            # Lead: keep info; red date if < 8/1/2025
            if lead_idx and c == lead_idx:
                if dt:
                    cell.value = dt
                    cell.number_format = "m/d/yy"
                    if dt < field_cutoff:
                        cell.font = red_font
                continue

            # General rule for other date fields
            if dt:
                if dt < general_cutoff:
                    cell.value = "X"
                    cell.font = red_font
                else:
                    cell.value = dt
                    cell.number_format = "m/d/yy"

    # Column widths
    width_map = {1: 16, 2: 22}
    for c in range(1, base_max_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = width_map.get(c, 14)

    # ---- Helper column for dynamic filtering (add AFTER styling so base_max_col stays table width)
    helper_col = base_max_col + 1
    helper_letter = get_column_letter(helper_col)
    ws.cell(row=filter_row, column=helper_col, value="VisibleFlag").font = Font(bold=True)
    anchor = f"$A${data_start}"
    for r in range(data_start, data_end + 1):
        ws.cell(row=r, column=helper_col).value = f'=SUBTOTAL(103,OFFSET({anchor},ROW()-ROW({anchor}),0))'
    ws.column_dimensions[helper_letter].hidden = True

    # ---- Dynamic Grand Total
    total_row = ws.max_row + 2
    ws.cell(row=total_row, column=1, value="Grand Total").font = Font(bold=True)
    ws.cell(row=total_row, column=1).alignment = Alignment(horizontal="left", vertical="center")

    center_align = Alignment(horizontal="center", vertical="center")
    top_border = Border(top=Side(style="thin"))
    vis_range = f"${helper_letter}${data_start}:${helper_letter}${data_end}"

    for c in range(1, base_max_col + 1):
        if c <= name_col_idx:
            continue
        col_letter = get_column_letter(c)
        data_range = f"${col_letter}${data_start}:${col_letter}${data_end}"
        formula = f'=SUMPRODUCT(--({vis_range}=1),--({data_range}<>""),--({data_range}<>"X"))'
        cell = ws.cell(row=total_row, column=c)
        cell.value = formula
        cell.alignment = center_align
        cell.font = Font(bold=True)
        cell.border = top_border

    # ---- Center Summary sheet (H–P completion -> completion rate)
    # Use H..P columns (8..16). Clamp to existing real columns (not including helper).
    H_idx = 8
    P_idx = min(16, base_max_col)
    req_cols = [c for c in range(H_idx, P_idx + 1) if c <= base_max_col]

    ws_summary = wb.create_sheet(title="Center Summary")

    # Only: Center/Campus, Completion Rate of Enrollment, Is 100%?
    ws_summary.append(["Center/Campus", "Completion Rate of Enrollment", "Is 100%?"])
    for c in range(1, 4):
        ws_summary.cell(row=1, column=c).font = Font(bold=True)
        ws_summary.cell(row=1, column=c).alignment = Alignment(horizontal="center", vertical="center")

    # Aggregate per center
    center_stats = {}
    for r in range(data_start, data_end + 1):
        center_name = ws.cell(row=r, column=center_idx).value if center_idx else "Unknown"
        if center_name is None or str(center_name).strip() == "":
            center_name = "Unknown"
        center_stats.setdefault(center_name, {"total": 0, "completed": 0})
        center_stats[center_name]["total"] += 1

        row_complete = True
        for c in req_cols:
            v = ws.cell(row=r, column=c).value
            if v in (None, "", "X"):
                row_complete = False
                break
        if row_complete:
            center_stats[center_name]["completed"] += 1

    # Sort centers by completion rate (desc) for a cleaner visual
    sorted_centers = sorted(
        center_stats.items(),
        key=lambda kv: (0 if kv[1]["total"] == 0 else kv[1]["completed"]/kv[1]["total"]),
        reverse=True
    )

    # Write rows and also capture arrays for plotting
    names, rates = [], []
    row_i = 2
    for center_name, stats in sorted_centers:
        total = stats["total"]
        completed = stats["completed"]
        rate = 0 if total == 0 else completed / total

        ws_summary.cell(row=row_i, column=1, value=center_name)
        cell_rate = ws_summary.cell(row=row_i, column=2, value=rate)
        cell_rate.number_format = "0%"
        ws_summary.cell(row=row_i, column=3, value=(completed == total))

        names.append(str(center_name))
        rates.append(rate)
        row_i += 1

    last_row = row_i - 1

    # Sheet widths & filter
    ws_summary.auto_filter.ref = f"A1:C{last_row}"
    ws_summary.column_dimensions["A"].width = 36
    ws_summary.column_dimensions["B"].width = 26
    ws_summary.column_dimensions["C"].width = 14

    # ---- Build chart as an IMAGE (robust against Excel "repair")
    if len(names) > 0:
        fig, ax = plt.subplots(figsize=(13, 7))  # larger but not huge

        # Unique colors (tab20 cycle)
        cmap = plt.get_cmap("tab20")
        colors = [cmap(i % 20) for i in range(len(names))]

        bars = ax.bar(range(len(names)), rates, color=colors)

        # Title & axes
        ax.set_title("Completion Rate of Enrollment", fontsize=16, pad=14)
        ax.set_ylabel("Enrollment Percentage")
        ax.set_ylim(0, max(1.0, min(1.05, max(rates) + 0.05)) )  # small headroom
        ax.yaxis.set_major_locator(MultipleLocator(0.25))
        ax.yaxis.set_major_formatter(PercentFormatter(xmax=1.0, decimals=0))

        # X tick labels (rotated for readability)
        ax.set_xticks(range(len(names)))
        ax.set_xticklabels(names, rotation=20, ha="right")

        # Gridlines for cleaner read
        ax.grid(axis="y", linestyle="-", linewidth=0.5, alpha=0.5)

        # Percentage labels on top of each bar
        for rect, val in zip(bars, rates):
            height = rect.get_height()
            ax.annotate(f"{val:.0%}",
                        xy=(rect.get_x() + rect.get_width() / 2, height),
                        xytext=(0, 5),
                        textcoords="offset points",
                        ha="center", va="bottom", fontsize=9)

        plt.tight_layout()

        chart_path = "center_completion_chart.png"
        fig.savefig(chart_path, dpi=180)
        plt.close(fig)

        # Embed the PNG into the worksheet
        img = XLImage(chart_path)
        # size tweak
        img.width = 1100
        img.height = 600
        ws_summary.add_image(img, "E2")

    # Save & download
    final_output = "Formatted_Enrollment_Checklist.xlsx"
    wb.save(final_output)

    with open(final_output, "rb") as f:
        st.download_button("📥 Download Formatted Excel", f, file_name=final_output)


