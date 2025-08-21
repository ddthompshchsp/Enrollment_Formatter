import io
import re
from pathlib import Path
from datetime import date, datetime
from zoneinfo import ZoneInfo
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="HCHSP Enrollment", layout="wide")

# ----------------------------
# Streamlit header (UI only)
# ----------------------------
logo_path = Path("header_logo.png")
hdr_l, hdr_c, hdr_r = st.columns([1, 2, 1])
with hdr_c:
    if logo_path.exists():
        st.image(str(logo_path), width=320)
    st.markdown(
        "<h1 style='text-align:center; margin: 8px 0 4px;'>Hidalgo County Head Start — Enrollment Formatter</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <p style='text-align:center; font-size:16px; margin-top:0;'>
        Upload the VF Average Funded Enrollment report and the 25–26 Applied/Accepted report.
        </p>
        """,
        unsafe_allow_html=True,
    )

st.divider()

# ----------------------------
# Inputs
# ----------------------------
inp_l, inp_c, inp_r = st.columns([1, 2, 1])
with inp_c:
    vf_file = st.file_uploader("Upload *VF_Average_Funded_Enrollment_Level.xlsx*", type=["xlsx"], key="vf")
    aa_file = st.file_uploader("Upload *25-26 Applied/Accepted.xlsx*", type=["xlsx"], key="aa")
    process = st.button("Process & Download")

# ----------------------------
# Static Lic. Cap values
# ----------------------------
LIC_CAPS = {
    "Alvarez-McAllen ISD": 138,
    "Camarena-La Joya ISD": 192,
    "Chapa-La Joya ISD": 154,
    "Edinburg": 232,
    "Edinburg North": 147,
    "Escandon-McAllen ISD": 131,
    "Farias-PSJA ISD": 153,
    "Guerra-PSJA ISD": 144,
    "Guzman-Donna ISD": 373,
    "Longoria-PSJA ISD": 125,
    "Mercedes-Mercedes ISD": 213,
    "Mission-Mission CISD": 165,
    "Monte Alto-Monte Alto ISD": 100,
    "Palacios-PSJA ISD": 135,
    "Salinas-Mission CISD": 90,
    "Sam Fordyce-La Joya ISD": 131,
    "Sam Houston-McAllen ISD": 90,
    "San Carlos-Edinburg CISD": 105,
    "San Juan-PSJA ISD": 180,
    "Seguin-La Joya ISD": 150,
    "Singleterry-Donna ISD": 130,
    "Thigpen-Zavala-McAllen ISD": 119,
    "Wilson-McAllen ISD": 119,
}

def _norm_key(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = re.sub(r"^HCHSP --\s*", "", s, flags=re.I)
    s = re.sub(r"\b(head\s*start|elem(?:entary)?|isd|cisd|isd\.?)\b", "", s, flags=re.I)
    s = s.replace("&", "and")
    s = re.sub(r"[-–—_/]", " ", s)
    s = re.sub(r"\s+", "", s).lower()
    return s

CAPS_LOOKUP = {_norm_key(k): v for k, v in LIC_CAPS.items()}

def cap_for_center(center: str):
    nk = _norm_key(center)
    if nk in CAPS_LOOKUP:
        return CAPS_LOOKUP[nk]
    for ck, cv in CAPS_LOOKUP.items():
        if ck in nk or nk in ck:
            return cv
    return None

# ----------------------------
# Parsers
# ----------------------------
def parse_vf(vf_df_raw: pd.DataFrame) -> pd.DataFrame:
    """Parse VF report (header=None) into per-class rows: Center | Class | Funded | Enrolled"""
    records = []
    current_center = None
    current_class = None

    for i in range(len(vf_df_raw)):
        c0 = vf_df_raw.iloc[i, 0]
        if isinstance(c0, str) and c0.startswith("HCHSP --"):
            current_center = c0.strip()
        elif isinstance(c0, str) and re.match(r"^Class \d+", c0):
            current_class = c0.split(" ", 1)[1].strip()

        if c0 == "Class Totals:" and current_center and current_class:
            row = vf_df_raw.iloc[i]
            funded = pd.to_numeric(row.iloc[4], errors="coerce")
            enrolled = pd.to_numeric(row.iloc[3], errors="coerce")
            center_clean = re.sub(r"^HCHSP --\s*", "", current_center)
            records.append({
                "Center": center_clean,
                "Class": f"Class {current_class}",
                "Funded": 0 if pd.isna(funded) else float(funded),
                "Enrolled": 0 if pd.isna(enrolled) else float(enrolled),
            })

    tidy = pd.DataFrame(records)
    if tidy.empty:
        raise ValueError("Could not find any 'Class Totals:' rows in the VF report. Check that you're uploading the correct file.")
    return tidy


def parse_applied_accepted(aa_df_raw: pd.DataFrame) -> pd.DataFrame:
    """Parse Applied/Accepted (header=None) to per-center counts; only blank 'ST: Status End Date' rows kept."""
    header_row_idx = aa_df_raw.index[aa_df_raw.iloc[:, 0].astype(str).str.startswith("ST: Participant PID", na=False)]
    if len(header_row_idx) == 0:
        raise ValueError("Could not find header row in Applied/Accepted report (expected a row starting with 'ST: Participant PID').")
    header_row_idx = int(header_row_idx[0])
    headers = aa_df_raw.iloc[header_row_idx].tolist()
    body = pd.DataFrame(aa_df_raw.iloc[header_row_idx + 1:].values, columns=headers)

    center_col = "ST: Center Name"
    status_col = "ST: Status"
    date_col = "ST: Status End Date"

    is_blank_date = body[date_col].isna() | body[date_col].astype(str).str.strip().eq("")
    body = body[is_blank_date].copy()
    body[center_col] = body[center_col].astype(str).str.replace(r"^HCHSP --\s*", "", regex=True)

    counts = body.groupby(center_col)[status_col].value_counts().unstack(fill_value=0)
    for c in ["Accepted", "Applied"]:
        if c not in counts.columns:
            counts[c] = 0

    return counts[["Accepted", "Applied"]].astype(int).reset_index().rename(columns={center_col: "Center"})

# ----------------------------
# Builder
# ----------------------------
def build_output_table(vf_tidy: pd.DataFrame, counts: pd.DataFrame) -> pd.DataFrame:
    """
    - Class rows: keep # Classrooms/Lic. Cap/Applied/Accepted/Lacking/Overage/Waitlist blank
    - Center totals:
        * # Classrooms = number of class rows for that center
        * Lic. Cap = from LIC_CAPS (blank if unknown)
        * Waitlist = Accepted if Enrolled > Funded else blank
        * Lacking/Overage = Funded - Enrolled (can be negative)
    - Agency total:
        * Lic. Cap blank
        * Waitlist = sum of center waitlists
        * Lacking/Overage = Funded - Enrolled
    """
    merged = vf_tidy.merge(counts, on="Center", how="left").fillna({"Accepted": 0, "Applied": 0})
    merged["% Enrolled of Funded"] = np.where(
        merged["Funded"] > 0,
        (merged["Enrolled"] / merged["Funded"] * 100).round(0).astype("Int64"),
        pd.NA
    )

    applied_by_center = merged.groupby("Center")["Applied"].max()
    accepted_by_center = merged.groupby("Center")["Accepted"].max()

    rows = []
    waitlist_totals = 0
    agency_classrooms_total = 0

    for center, group in merged.groupby("Center", sort=True):
        # Class rows
        for _, r in group.iterrows():
            rows.append({
                "Center": r["Center"],
                "Class": r["Class"],
                "# Classrooms": "",
                "Lic. Cap": "",
                "Funded": int(r["Funded"]),
                "Enrolled": int(r["Enrolled"]),
                "Applied": "",
                "Accepted": "",
                "Lacking/Overage": "",
                "Waitlist": "",
                "% Enrolled of Funded": int(r["% Enrolled of Funded"]) if pd.notna(r["% Enrolled of Funded"]) else pd.NA
            })

        # Center totals
        funded_sum   = int(group["Funded"].sum())
        enrolled_sum = int(group["Enrolled"].sum())
        pct_total    = int(round(enrolled_sum / funded_sum * 100, 0)) if funded_sum > 0 else pd.NA

        accepted_val = int(accepted_by_center.get(center, 0))
        applied_val  = int(applied_by_center.get(center, 0))
        waitlist_val = accepted_val if enrolled_sum > funded_sum else ""
        lacking_over = funded_sum - enrolled_sum

        class_count = int(len(group))  # number of class rows
        agency_classrooms_total += class_count

        lic_cap_val = cap_for_center(center)
        if waitlist_val != "":
            waitlist_totals += waitlist_val

        rows.append({
            "Center": f"{center} Total",
            "Class": "",
            "# Classrooms": class_count,
            "Lic. Cap": ("" if lic_cap_val is None else int(lic_cap_val)),
            "Funded": funded_sum,
            "Enrolled": enrolled_sum,
            "Applied": applied_val,
            "Accepted": accepted_val,
            "Lacking/Overage": lacking_over,
            "Waitlist": waitlist_val,
            "% Enrolled of Funded": pct_total
        })

    final = pd.DataFrame(rows)

    # Agency totals (Lic. Cap intentionally blank)
    agency_funded   = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Funded"].sum())
    agency_enrolled = int(final.loc[final["Center"].str.endswith(" Total", na=False), "Enrolled"].sum())
    agency_applied  = int(counts["Applied"].sum())
    agency_accepted = int(counts["Accepted"].sum())
    agency_pct      = int(round(agency_enrolled / agency_funded * 100, 0)) if agency_funded > 0 else pd.NA
    agency_lacking  = agency_funded - agency_enrolled

    final = pd.concat([final, pd.DataFrame([{
        "Center": "Agency Total",
        "Class": "",
        "# Classrooms": agency_classrooms_total,
        "Lic. Cap": "",
        "Funded": agency_funded,
        "Enrolled": agency_enrolled,
        "Applied": agency_applied,
        "Accepted": agency_accepted,
        "Lacking/Overage": agency_lacking,
        "Waitlist": waitlist_totals,
        "% Enrolled of Funded": agency_pct
    }])], ignore_index=True)

    # Final column order (Waitlist AFTER Lacking/Overage)
    final = final[[
        "Center","Class","# Classrooms","Lic. Cap",
        "Funded","Enrolled","Applied","Accepted","Lacking/Overage","Waitlist","% Enrolled of Funded"
    ]]
    return final

# ----------------------------
# Excel Writer (Power BI)
# ----------------------------
def to_styled_excel(df: pd.DataFrame) -> bytes:
    """
    Logo at A1; titles merged in B..last; thick outer box from row 1 (continuous across title);
    borders on table; gridlines outside kept; Named Range for PBI; freeze top 4 rows;
    subtitle shows date and Central time in 12-hour format.
    """
    def col_letter(n: int) -> str:
        s = ""
        while n >= 0:
            s = chr(n % 26 + 65) + s
            n = n // 26 - 1
        return s

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Head Start Enrollment", startrow=3)
        wb = writer.book
        ws = writer.sheets["Head Start Enrollment"]

        # Keep default gridlines visible outside the table
        ws.hide_gridlines(0)

        # Title area row heights
        ws.set_row(0, 24)  # title
        ws.set_row(1, 22)  # subtitle
        ws.set_row(2, 20)  # spacer above header

        # --- LOGO in column A (A1) ---
        logo = Path("header_logo.png")
        if logo.exists():
            ws.set_column(0, 0, 7)  # column A width for logo
            ws.insert_image(0, 0, str(logo), {
                "x_offset": 4, "y_offset": 3,
                "x_scale": 0.53, "y_scale": 0.53,
                "object_position": 1
            })

        # --- Titles + Central timestamp ---
        today = date.today()
        date_str = f"{today.month}.{today.day}.{str(today.year % 100).zfill(2)}"
        now_ct = datetime.now(ZoneInfo("America/Chicago"))
        time_str = now_ct.strftime("%I:%M %p").lstrip("0")  # e.g., 8:05 PM
        tz_abbr = now_ct.strftime("%Z")  # CST or CDT

        title_fmt = wb.add_format({"bold": True, "font_size": 14, "align": "center"})
        subtitle_fmt = wb.add_format({"bold": True, "font_size": 12, "align": "center"})
        red_fmt = wb.add_format({"bold": True, "font_size": 12, "font_color": "#C00000"})

        last_col_0 = len(df.columns) - 1
        last_col_letter = col_letter(last_col_0)

        ws.merge_range(0, 1, 0, last_col_0, "Hidalgo County Head Start Program", title_fmt)
        ws.merge_range(1, 1, 1, last_col_0, "", subtitle_fmt)
        ws.write_rich_string(
            1, 1,
            subtitle_fmt, "Head Start - 2025-2026 Campus Classroom Enrollment as of ",
            red_fmt, f"({date_str}, {time_str} {tz_abbr})",
            subtitle_fmt
        )

        # --- Header row (blue) ---
        header_fmt = wb.add_format({
            "bold": True, "font_color": "white", "bg_color": "#305496",
            "align": "center", "valign": "vcenter", "text_wrap": True,
            "border": 1
        })
        ws.set_row(3, 26)
        for c, col in enumerate(df.columns):
            ws.write(3, c, col, header_fmt)

        last_row_0 = len(df) + 3
        last_excel_row = last_row_0 + 1

        # Named range for Power BI
        wb.define_name("EnrollmentRange", f"='Head Start Enrollment'!$A$4:${last_col_letter}${last_excel_row}")

        # Filters + Freeze panes (top 4 rows)
        ws.autofilter(3, 0, last_row_0, last_col_0)
        ws.freeze_panes(4, 0)

        # Column widths
        widths = {
            "Center": 28, "Class": 14, "# Classrooms": 12, "Lic. Cap": 12,
            "Funded": 12, "Enrolled": 12, "Applied": 12, "Accepted": 12,
            "Lacking/Overage": 14, "Waitlist": 12
        }
        for name, width in widths.items():
            if name in df.columns:
                idx = df.columns.get_loc(name)
                ws.set_column(idx, idx, width)
        pct_idx = df.columns.get_loc("% Enrolled of Funded")
        ws.set_column(pct_idx, pct_idx, 16)

        # Borders on every header+data cell
        border_all = wb.add_format({"border": 1})
        ws.conditional_format(f"A4:{last_col_letter}{last_excel_row}",
                              {"type": "formula", "criteria": "TRUE", "format": border_all})

        # % display & colors
        pct_letter = col_letter(pct_idx)
        pct_range = f"{pct_letter}5:{pct_letter}{last_excel_row}"
        ws.conditional_format(pct_range, {"type": "cell", "criteria": "<", "value": 100,
                                          "format": wb.add_format({"font_color": "red"})})
        ws.conditional_format(pct_range, {"type": "cell", "criteria": ">", "value": 100,
                                          "format": wb.add_format({"font_color": "blue"})})
        ws.conditional_format(pct_range, {"type": "formula", "criteria": "TRUE",
                                          "format": wb.add_format({'num_format': '0"%"', 'align': 'center'})})

        # Bold center totals + agency total
        bold_row = wb.add_format({"bold": True})
        for ridx, val in enumerate(df["Center"].tolist()):
            if isinstance(val, str) and (val.endswith(" Total") or val == "Agency Total"):
                ws.set_row(ridx + 4, None, bold_row)

        # ===== Thick outer box from row 1 to the end =====
        top    = wb.add_format({"top": 2})
        bottom = wb.add_format({"bottom": 2})
        left   = wb.add_format({"left": 2})
        right  = wb.add_format({"right": 2})

        # Top edge across A1..last
        ws.conditional_format(f"A1:{last_col_letter}1",
                              {"type": "formula", "criteria": "TRUE", "format": top})
        # Left & right edges from row 1 to bottom
        ws.conditional_format(f"A1:A{last_excel_row}",
                              {"type": "formula", "criteria": "TRUE", "format": left})
        ws.conditional_format(f"{last_col_letter}1:{last_col_letter}{last_excel_row}",
                              {"type": "formula", "criteria": "TRUE", "format": right})
        # Bottom edge at the end of the table
        ws.conditional_format(f"A{last_excel_row}:{last_col_letter}{last_excel_row}",
                              {"type": "formula", "criteria": "TRUE", "format": bottom})

        # Bulletproof edges for merged title area (ensure corners/edges render)
        ws.write(0, last_col_0, "", wb.add_format({"right": 2, "top": 2}))  # top-right corner
        ws.write(1, last_col_0, "", wb.add_format({"right": 2}))
        ws.write(2, last_col_0, "", wb.add_format({"right": 2}))
        ws.write(0, 0, "", wb.add_format({"left": 2, "top": 2}))            # top-left corner
        ws.write(1, 0, "", wb.add_format({"left": 2}))
        ws.write(2, 0, "", wb.add_format({"left": 2}))

    return output.getvalue()

# ----------------------------
# Main
# ----------------------------
if process and vf_file and aa_file:
    try:
        vf_raw = pd.read_excel(vf_file, sheet_name=0, header=None)
        aa_raw = pd.read_excel(aa_file, sheet_name=0, header=None)

        vf_tidy = parse_vf(vf_raw)
        aa_counts = parse_applied_accepted(aa_raw)
        final_df = build_output_table(vf_tidy, aa_counts)

        st.success("Preview below. Use the download button to get the Excel file.")
        preview_df = final_df.copy()
        pct_col = "% Enrolled of Funded"
        preview_df[pct_col] = preview_df[pct_col].apply(lambda v: "" if pd.isna(v) else f"{int(v)}%")
        preview_df = preview_df[
            ["Center","Class","# Classrooms","Lic. Cap",
             "Funded","Enrolled","Applied","Accepted","Lacking/Overage","Waitlist",pct_col]
        ]
        st.dataframe(preview_df, use_container_width=True)

        xlsx_bytes = to_styled_excel(final_df)
        st.download_button(
            "Download Formatted Excel",
            data=xlsx_bytes,
            file_name="HCHSP_Enrollment_Formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Processing error: {e}")





