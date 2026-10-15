import os
import io
import pandas as pd
import streamlit as st
import plotly.express as px
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

st.set_page_config(
    page_title="Feedback Analyzer & Dual-Closure Report Tool",
    layout="wide",
    page_icon="📊"
)

ROSTER_FILE = "Master_Roster_DB.xlsx"

DEFAULT_HC = [
    {"LOB": "Chat Team", "Team HC": 5},
    {"LOB": "CHR - Egypt", "Team HC": 5},
    {"LOB": "CHR - MEA", "Team HC": 7},
    {"LOB": "CSE Team", "Team HC": 5},
    {"LOB": "E-Commerce", "Team HC": 9},
    {"LOB": "Egypt Hotline", "Team HC": 19},
    {"LOB": "Egypt Supervisor", "Team HC": 1},
    {"LOB": "E-Mail Team", "Team HC": 3},
    {"LOB": "MEA Hotline", "Team HC": 7},
    {"LOB": "Outbound", "Team HC": 10},
    {"LOB": "Quality Department", "Team HC": 6},
    {"LOB": "Social Media Arabic", "Team HC": 2},
    {"LOB": "Social Media English", "Team HC": 9},
    {"LOB": "South Africa Team", "Team HC": 9},
    {"LOB": "Training Department", "Team HC": 3},
    {"LOB": "UAE Supervisor", "Team HC": 1},
    {"LOB": "WhatsApp", "Team HC": 7},
]

def load_master_roster():
    if os.path.exists(ROSTER_FILE):
        xls = pd.ExcelFile(ROSTER_FILE)
        agents_df = pd.read_excel(xls, sheet_name="Agents_Roster") if "Agents_Roster" in xls.sheet_names else pd.read_excel(xls, sheet_name=0)
        hc_df = pd.read_excel(xls, sheet_name="Team_HC") if "Team_HC" in xls.sheet_names else pd.DataFrame(DEFAULT_HC)
    else:
        agents_df = pd.DataFrame(columns=["Consignee", "Submitter Name", "Submitter LOB"])
        hc_df = pd.DataFrame(DEFAULT_HC)

    for col in ["Consignee", "Submitter Name", "Submitter LOB"]:
        if col in agents_df.columns:
            agents_df[col] = agents_df[col].astype(str).str.strip()
    if "LOB" in hc_df.columns:
        hc_df["LOB"] = hc_df["LOB"].astype(str).str.strip()
    return agents_df, hc_df

def save_master_roster(agents_df, hc_df):
    with pd.ExcelWriter(ROSTER_FILE, engine="openpyxl") as writer:
        agents_df.to_excel(writer, sheet_name="Agents_Roster", index=False)
        hc_df.to_excel(writer, sheet_name="Team_HC", index=False)

agents_df, hc_df = load_master_roster()

# --- Sidebar: Master Roster & Team HC ---
st.sidebar.title("⚙️ Master Roster & Team HC")
st.sidebar.markdown(
    f"**Registered Agents:** `{len(agents_df)}`  \n"
    f"**Registered Teams (LOB):** `{len(hc_df)}`"
)

uploaded_roster = st.sidebar.file_uploader("🔄 Upload Updated Roster Sheet (Optional)", type=["xlsx", "xls", "csv"])
if uploaded_roster is not None:
    try:
        if uploaded_roster.name.endswith(".csv"):
            new_agents = pd.read_csv(uploaded_roster)
            new_hc = hc_df
        else:
            xls_up = pd.ExcelFile(uploaded_roster)
            if "Agents_Roster" in xls_up.sheet_names:
                new_agents = pd.read_excel(xls_up, "Agents_Roster")
                new_hc = pd.read_excel(xls_up, "Team_HC") if "Team_HC" in xls_up.sheet_names else hc_df
            else:
                new_agents = pd.read_excel(xls_up, 0)
                new_hc = hc_df

        if {"Consignee", "Submitter Name", "Submitter LOB"}.issubset(set(new_agents.columns)):
            new_agents["Consignee"] = new_agents["Consignee"].astype(str).str.strip()
            new_agents["Submitter Name"] = new_agents["Submitter Name"].astype(str).str.strip()
            new_agents["Submitter LOB"] = new_agents["Submitter LOB"].astype(str).str.strip()
            combined_agents = pd.concat([agents_df, new_agents], ignore_index=True).drop_duplicates(subset=["Consignee"], keep="last")
            save_master_roster(combined_agents, new_hc)
            agents_df, hc_df = load_master_roster()
            st.sidebar.success("✅ Master Roster updated and saved!")
        else:
            st.sidebar.error("Sheet must contain: Consignee, Submitter Name, Submitter LOB")
    except Exception as e:
        st.sidebar.error(f"Error reading Roster file: {e}")

with st.sidebar.expander("📝 View / Edit Roster & Team HC"):
    st.caption("Edit or add agents/teams directly below and click Save:")
    edited_agents = st.data_editor(agents_df, num_rows="dynamic", key="edit_agents")
    edited_hc = st.data_editor(hc_df, num_rows="dynamic", key="edit_hc")
    if st.button("💾 Save Roster Changes"):
        save_master_roster(edited_agents, edited_hc)
        agents_df, hc_df = load_master_roster()
        st.success("Saved successfully!")

# --- Smart Pre-Classification Engine ---
def smart_pre_classify_row(row):
    sol1 = str(row.get("Responsible person replies with a handling solution", "SNULL")).strip()
    sol2 = str(row.get("Responsible person replies with a handling solution.1", sol1)).strip()
    resolved = str(row.get("Is it resolved?", "SNULL")).strip()
    loop_st = str(row.get("Has the responsible person promptly closed the loop?", "")).strip()
    upg_res = str(row.get("Upgrade of requirements, human handling of results", "SNULL")).strip()
    sys_res = str(row.get("The responsible person replies with the handling result.", "SNULL")).strip()
    reply_txt = str(row.get("Responsible Person's Reply and Comments", "SNULL")).strip()
    issue_type = str(row.get("Issue Type", "")).strip()

    # Rule 1: If already explicitly labeled Reasonable or Unreasonable in the uploaded Raw Data
    if sys_res in ["Reasonable", "Unreasonable"]:
        sev = "High" if (sys_res == "Unreasonable" or issue_type == "The content is incorrect.") else ("Medium" if issue_type == "Situation encountered by uncovered users" else "Low")
        esc = "Yes" if sys_res == "Unreasonable" else "No"
        return sys_res, "100% (Explicit in Raw Data)", "", sev, esc

    # Rule 2: If Handling Solution is Corrected Knowledge or Optimized knowledge -> Reasonable
    if sol1 in ["Corrected Knowledge", "Optimized knowledge"] or sol2 in ["Corrected Knowledge", "Optimized knowledge"] or upg_res == "The complaint is reasonable":
        return "Reasonable", "High (Auto: Knowledge Updated)", "", "Low", "No"

    # Rule 3: If Upgrade result says unreasonable OR Is it resolved? == 'The issue remains unresolved.' -> Unreasonable
    if upg_res == "The complaint is unreasonable." or resolved == "The issue remains unresolved.":
        return "Unreasonable", "High (Auto: Unresolved / Rejected)", "", "High", "Yes"

    # Rule 4: Check handler reply keywords if solution is SNULL
    reply_lower = reply_txt.lower()
    if reply_txt != "SNULL" and reply_txt != "" and reply_txt != "nan":
        reject_phrases = ["sorry", "no changes", "not available", "will not", "already modified", "originally in english", "ineffective", "not a knowledge"]
        accept_phrases = ["done", "changed to", "updated", "added", "modified the knowledge", "received feedback and changed"]
        if any(p in reply_lower for p in reject_phrases):
            return "Unreasonable", "Medium-High (Auto: Rejection Reply)", "", "High", "Yes"
        if any(p in reply_lower for p in accept_phrases):
            return "Reasonable", "Medium-High (Auto: Positive Reply)", "", "Low", "No"

    # Rule 5: If Loop is still 'The electronic current has not been fully completed.' and no reply -> SNULL
    if loop_st == "The electronic current has not been fully completed." and (reply_txt == "SNULL" or not reply_txt):
        sev = "High" if issue_type == "The content is incorrect." else ("Medium" if issue_type == "Situation encountered by uncovered users" else "Low")
        return "SNULL", "High (Auto: Loop Incomplete & No Reply)", "", sev, "Yes"

    return "SNULL", "Low (Check E-Flow Link)", "", "Medium", "Yes"


def enrich_and_classify_dataframe(df, agents_df):
    df = df.copy()
    df["Consignee"] = df["Consignee"].astype(str).str.strip()

    roster_map_name = dict(zip(agents_df["Consignee"].str.upper(), agents_df["Submitter Name"]))
    roster_map_lob = dict(zip(agents_df["Consignee"].str.upper(), agents_df["Submitter LOB"]))

    mapped_names = df["Consignee"].str.upper().map(roster_map_name)
    mapped_lobs = df["Consignee"].str.upper().map(roster_map_lob)

    if "Submitter Name" in df.columns:
        df["Submitter Name"] = mapped_names.fillna(df["Submitter Name"].astype(str).str.strip())
    else:
        c_idx = df.columns.get_loc("Consignee")
        df.insert(c_idx + 1, "Submitter Name", mapped_names.fillna("Unknown Agent"))

    if "Submitter LOB" in df.columns:
        df["Submitter LOB"] = mapped_lobs.fillna(df["Submitter LOB"].astype(str).str.strip())
    else:
        n_idx = df.columns.get_loc("Submitter Name")
        df.insert(n_idx + 1, "Submitter LOB", mapped_lobs.fillna("Unassigned LOB"))

    df["Submitter Name"] = df["Submitter Name"].astype(str).str.strip()
    df["Submitter LOB"] = df["Submitter LOB"].astype(str).str.strip()

    results, confs, comms, sevs, escs = [], [], [], [], []
    for _, row in df.iterrows():
        r_val, c_val, cm_val, sv_val, es_val = smart_pre_classify_row(row)
        results.append(r_val)
        confs.append(c_val)
        comms.append(cm_val)
        sevs.append(sv_val)
        escs.append(es_val)

    res_col = "The responsible person replies with the handling result."
    df[res_col] = results
    df["Auto_Detection_Confidence"] = confs
    df["Comment"] = df["Comment"] if "Comment" in df.columns else comms
    df["Severity"] = df["Severity"] if "Severity" in df.columns else sevs
    df["Escalation status"] = df["Escalation status"] if "Escalation status" in df.columns else escs
    df["Relevant Team"] = df["Submitter LOB"]
    return df


def apply_optional_pending_review_updates(base_df, reviewed_df):
    """
    Applies user updates from the optional Pending_Cases_Review sheet onto the Raw Data.
    If the user changes a ticket from Reasonable -> Unreasonable, Unreasonable -> Reasonable,
    or SNULL -> Reasonable/Unreasonable, it updates the main dataset so the final report reflects it!
    """
    df = base_df.copy()
    rev = reviewed_df.copy()

    res_col = "The responsible person replies with the handling result."
    loop_col = "Has the responsible person promptly closed the loop?"

    df["Order No."] = df["Order No."].astype(str).str.strip()
    rev["Order No."] = rev["Order No."].astype(str).str.strip()
    rev = rev.drop_duplicates(subset=["Order No."], keep="last")

    # Map updated columns by Order No.
    cols_to_sync = [
        res_col,
        loop_col,
        "Comment",
        "Severity",
        "Escalation status",
        "Relevant Team",
        "Submitter LOB",
        "Submitter Name"
    ]
    # Also support 'Escalation status ' with trailing space
    if "Escalation status " in rev.columns and "Escalation status" not in rev.columns:
        rev["Escalation status"] = rev["Escalation status "]

    for col in cols_to_sync:
        if col in rev.columns:
            valid_rev = rev[rev[col].notna() & (rev[col].astype(str).str.strip() != "") & (rev[col].astype(str).str.strip() != "nan")]
            update_map = dict(zip(valid_rev["Order No."], valid_rev[col].astype(str).str.strip()))
            if col in df.columns:
                df[col] = df["Order No."].map(update_map).fillna(df[col])
            else:
                df[col] = df["Order No."].map(update_map).fillna("")

    return df


def build_pending_review_19_df(df):
    res_col = "The responsible person replies with the handling result."
    p_sub = df[df[res_col].isin(["Unreasonable", "SNULL"])].copy()
    sol_col_1 = "Responsible person replies with a handling solution"
    sol_col_2 = "Responsible person replies with a handling solution.1" if "Responsible person replies with a handling solution.1" in p_sub.columns else sol_col_1

    return pd.DataFrame({
        "Order No.": p_sub.get("Order No.", ""),
        "Bill of Lading Time": p_sub.get("Bill of Lading Time", ""),
        "Issue Type": p_sub.get("Issue Type", ""),
        "Country/Region Where the Issue Occurred": p_sub.get("Country/Region Where the Issue Occurred", ""),
        "Process Type": p_sub.get("Process Type", ""),
        "Consignee": p_sub.get("Consignee", ""),
        "Submitter Name": p_sub.get("Submitter Name", ""),
        "Submitter LOB": p_sub.get("Submitter LOB", ""),
        "Knowledge ID": p_sub.get("Knowledge ID", ""),
        "Responsible person replies with a handling solution (Initial)": p_sub.get(sol_col_1, "SNULL"),
        "Has the responsible person promptly closed the loop?": p_sub.get("Has the responsible person promptly closed the loop?", ""),
        "Person Responsible": p_sub.get("Person Responsible", "SNULL"),
        "Responsible person replies with a handling solution (Final)": p_sub.get(sol_col_2, "SNULL"),
        "The responsible person replies with the handling result.": p_sub.get(res_col, "SNULL"),
        "Electronic Flow Link": p_sub.get("Electronic Flow Link", ""),
        "Comment": p_sub.get("Comment", ""),
        "Severity": p_sub.get("Severity", ""),
        "Escalation status": p_sub.get("Escalation status", "Yes"),
        "Relevant Team": p_sub.get("Relevant Team", p_sub.get("Submitter LOB", ""))
    })


def create_styled_excel(sheets_dict, add_dropdowns_on_sheet=None):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    editable_header_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    regular_font = Font(name="Calibri", size=11, color="1E293B")
    bold_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    link_font = Font(name="Calibri", size=10, color="2563EB", underline="single")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    editable_cols = {
        "Comment", "Severity", "Escalation status", "Relevant Team",
        "The responsible person replies with the handling result."
    }

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(title=sheet_name)
        cols = list(df.columns)

        for col_idx, col_name in enumerate(cols, start=1):
            cell = ws.cell(1, col_idx, col_name)
            cell.fill = editable_header_fill if (sheet_name == add_dropdowns_on_sheet and col_name in editable_cols) else header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            for c_idx, val in enumerate(row, start=1):
                col_name = cols[c_idx - 1]
                clean_v = "" if pd.isna(val) else val
                cell = ws.cell(r_idx, c_idx, clean_v)

                if col_name == "Electronic Flow Link" and str(clean_v).startswith("http"):
                    cell.hyperlink = str(clean_v)
                    cell.font = link_font
                else:
                    cell.font = bold_font if c_idx == 1 else regular_font

                if r_idx % 2 == 1:
                    cell.fill = zebra_fill
                cell.border = thin_border
                if col_name in ["Comment", "Electronic Flow Link", "Issue Type", "Responsible Person's Reply and Comments", "E-Flow Problem Description"]:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=(col_name in ["Comment", "Responsible Person's Reply and Comments", "E-Flow Problem Description"]))
                else:
                    cell.alignment = Alignment(horizontal="center" if c_idx > 1 else "left", vertical="center")

        ws.freeze_panes = "A2"
        if len(df) > 0:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(df) + 1}"

        for c_idx, col_name in enumerate(cols, start=1):
            col_letter = get_column_letter(c_idx)
            if col_name in ["Comment", "Responsible Person's Reply and Comments", "E-Flow Problem Description"]:
                ws.column_dimensions[col_letter].width = 42
            elif col_name in ["Issue Type", "Has the responsible person promptly closed the loop?"]:
                ws.column_dimensions[col_letter].width = 32
            elif col_name in ["Submitter Name", "Electronic Flow Link"]:
                ws.column_dimensions[col_letter].width = 28
            else:
                ws.column_dimensions[col_letter].width = 22

        if sheet_name == add_dropdowns_on_sheet and len(df) > 0:
            max_r = len(df) + 1
            res_col_name = "The responsible person replies with the handling result."
            if res_col_name in cols:
                r_col = get_column_letter(cols.index(res_col_name) + 1)
                dv_res = DataValidation(type="list", formula1='"Reasonable,Unreasonable,SNULL"', allow_blank=True)
                ws.add_data_validation(dv_res)
                dv_res.add(f"{r_col}2:{r_col}{max_r}")
            if "Severity" in cols:
                sev_col = get_column_letter(cols.index("Severity") + 1)
                dv_sev = DataValidation(type="list", formula1='"High,Medium,Low"', allow_blank=True)
                ws.add_data_validation(dv_sev)
                dv_sev.add(f"{sev_col}2:{sev_col}{max_r}")
            if "Escalation status" in cols:
                esc_col = get_column_letter(cols.index("Escalation status") + 1)
                dv_esc = DataValidation(type="list", formula1='"Yes,No"', allow_blank=True)
                ws.add_data_validation(dv_esc)
                dv_esc.add(f"{esc_col}2:{esc_col}{max_r}")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# --- Main UI ---
st.title("📊 Feedback Analysis & Dual-Closure Reporting Tool")
st.markdown(
    "**How It Works:**  \n"
    "1. **Upload Raw Data Sheet:** Automatically matches `Consignee` to `Submitter Name` & `Submitter LOB` (and sets `Relevant Team = Submitter LOB`), extracts `Pending_Cases_Review.xlsx` (`Unreasonable` & `SNULL` cases), and computes **Both Closure Status Methods** (Old Loop-Based & New Handling-Result-Based) side by side.  \n"
    "2. **Optional Upload of Updated `Pending_Cases_Review.xlsx`:** After reviewing or updating `Pending_Cases_Review.xlsx` manually (e.g., changing a case from `Unreasonable` to `Reasonable` or vice versa, or adding `Comment` / `Severity` / `Escalation status`), upload it in Box 2 to immediately recalculate the final **Summary**, **Charts**, **PDF Report**, and **Excel Workbook**."
)

u_col1, u_col2 = st.columns(2)
with u_col1:
    raw_file = st.file_uploader(
        "📂 1. Upload Raw Data Sheet (Required - System Export or Raw Data):",
        type=["xlsx", "xls", "csv"],
        key="raw_uploader"
    )
with u_col2:
    optional_review_file = st.file_uploader(
        "📝 2. (Optional) Upload Updated 'Pending_Cases_Review' Sheet to Recalculate Final Report:",
        type=["xlsx", "xls", "csv"],
        key="review_uploader"
    )

if raw_file is not None:
    # 1. Read Raw Data ONLY (Auto multi-tab sync disabled as requested)
    if raw_file.name.endswith(".csv"):
        raw_input_df = pd.read_csv(raw_file)
    else:
        xls_raw = pd.ExcelFile(raw_file)
        if "Raw Data" in xls_raw.sheet_names:
            raw_input_df = pd.read_excel(xls_raw, sheet_name="Raw Data")
        elif "Schedule of Details" in xls_raw.sheet_names:
            raw_input_df = pd.read_excel(xls_raw, sheet_name="Schedule of Details")
        elif "Enriched_Raw_Data" in xls_raw.sheet_names:
            raw_input_df = pd.read_excel(xls_raw, sheet_name="Enriched_Raw_Data")
        else:
            raw_input_df = pd.read_excel(xls_raw, sheet_name=xls_raw.sheet_names[0])

    if "Consignee" not in raw_input_df.columns:
        st.error("❌ Column 'Consignee' was not found in the uploaded Raw Data sheet!")
        st.stop()

    working_df = enrich_and_classify_dataframe(raw_input_df, agents_df)

    # 2. If user uploaded the optional Updated Pending_Cases_Review file, apply its changes!
    if optional_review_file is not None:
        if optional_review_file.name.endswith(".csv"):
            rev_input_df = pd.read_csv(optional_review_file)
        else:
            xls_rev = pd.ExcelFile(optional_review_file)
            rev_sheet = "Pending_Cases_Review" if "Pending_Cases_Review" in xls_rev.sheet_names else xls_rev.sheet_names[0]
            rev_input_df = pd.read_excel(xls_rev, sheet_name=rev_sheet)

        if "Order No." in rev_input_df.columns:
            working_df = apply_optional_pending_review_updates(working_df, rev_input_df)
            st.success(f"✅ Applied manual updates from `{optional_review_file.name}`! Any status changes (Reasonable ↔ Unreasonable / SNULL) are now reflected in the final analysis below.")
        else:
            st.error("❌ The uploaded review sheet must contain an 'Order No.' column.")

    # Check for unmapped Consignees
    unmapped_mask = (
        working_df["Submitter Name"].isna()
        | (working_df["Submitter Name"] == "Unknown Agent")
        | (working_df["Submitter Name"] == "nan")
    )
    if unmapped_mask.any():
        unmapped_ids = working_df.loc[unmapped_mask, "Consignee"].unique().tolist()
        with st.expander(f"⚠️ {len(unmapped_ids)} Unmapped Consignee IDs Detected! Click here to add them to Master Roster:", expanded=True):
            new_roster_rows = pd.DataFrame({
                "Consignee": unmapped_ids,
                "Submitter Name": ["" for _ in unmapped_ids],
                "Submitter LOB": [hc_df["LOB"].iloc[0] if len(hc_df) > 0 else "Egypt Hotline" for _ in unmapped_ids]
            })
            edited_new_agents = st.data_editor(
                new_roster_rows,
                column_config={
                    "Submitter LOB": st.column_config.SelectboxColumn("Submitter LOB", options=sorted(hc_df["LOB"].unique().tolist()))
                },
                use_container_width=True,
                key="quick_map_editor"
            )
            if st.button("✅ Save New Agents to Master Roster & Re-Map"):
                valid_new = edited_new_agents[edited_new_agents["Submitter Name"].astype(str).str.strip() != ""]
                if len(valid_new) > 0:
                    combined = pd.concat([agents_df, valid_new], ignore_index=True).drop_duplicates(subset=["Consignee"], keep="last")
                    save_master_roster(combined, hc_df)
                    st.rerun()

    # --- Optional In-App Quick Editor (Allows changing Reasonable <-> Unreasonable live as well) ---
    res_col = "The responsible person replies with the handling result."
    loop_col = "Has the responsible person promptly closed the loop?"

    with st.expander("✏️ Optional: Quick In-App Case Editor (Change Reasonable ↔ Unreasonable / SNULL or Add Comments Live)", expanded=False):
        edit_filter = st.radio(
            "Show Cases:",
            ["Unreasonable & SNULL Only", "All Cases (Including Reasonable)"],
            horizontal=True
        )
        e_mask = working_df[res_col].isin(["Unreasonable", "SNULL"]) if edit_filter == "Unreasonable & SNULL Only" else pd.Series([True] * len(working_df), index=working_df.index)
        edit_cols = [
            "Order No.", "Bill of Lading Time", "Consignee", "Submitter Name", "Submitter LOB",
            "Issue Type", "Knowledge ID", loop_col, "Electronic Flow Link",
            res_col, "Comment", "Severity", "Escalation status", "Relevant Team"
        ]
        edit_cols = [c for c in edit_cols if c in working_df.columns]
        edited_live = st.data_editor(
            working_df.loc[e_mask, edit_cols],
            column_config={
                "Electronic Flow Link": st.column_config.LinkColumn("Electronic Flow Link", display_text="🔗 Open Link"),
                res_col: st.column_config.SelectboxColumn("Handling Result", options=["Reasonable", "Unreasonable", "SNULL"], required=True),
                "Severity": st.column_config.SelectboxColumn("Severity", options=["", "High", "Medium", "Low"]),
                "Escalation status": st.column_config.SelectboxColumn("Escalation status", options=["Yes", "No"]),
            },
            use_container_width=True,
            key="live_case_editor"
        )
        working_df.loc[edited_live.index, edit_cols] = edited_live

    # Build Pending_Cases_Review (19-column structure for Unreasonable & SNULL cases)
    # Plus if user uploaded an external review sheet that had some cases changed to Reasonable, we also keep a full view of reviewed rows if needed, while Pending_Cases_Review holds current Unreasonable + SNULL
    pending_review_19_df = build_pending_review_19_df(working_df)

    # Stage 1 Export Button: Download Pending_Cases_Review.xlsx for Manual Review
    stage1_excel_bytes = create_styled_excel(
        {
            "Pending_Cases_Review": pending_review_19_df,
            "Enriched_Raw_Data": working_df
        },
        add_dropdowns_on_sheet="Pending_Cases_Review"
    )
    st.download_button(
        label="📥 Stage 1: Download 'Pending_Cases_Review.xlsx' (Unreasonable & SNULL Cases) for Manual Review & Update",
        data=stage1_excel_bytes,
        file_name="Pending_Cases_Review.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    st.divider()

    # ==========================================================================
    # COMPUTE SUMMARY & PIVOT WITH BOTH CLOSURE STATUS METHODS (OLD & NEW)
    # ==========================================================================
    all_lobs = sorted(list(set(hc_df["LOB"].tolist() + working_df["Submitter LOB"].unique().tolist())))
    hc_map = dict(zip(hc_df["LOB"], hc_df["Team HC"]))

    summary_rows, pivot_rows = [], []

    for lob in all_lobs:
        sub = working_df[working_df["Submitter LOB"] == lob]
        hc = int(hc_map.get(lob, max(1, len(sub["Submitter Name"].unique()))))
        part = sub["Submitter Name"].nunique() if len(sub) > 0 else 0
        total_tix = len(sub)
        usage_rate = (part / hc * 100) if hc > 0 else 0.0

        reas_cnt = int((sub[res_col] == "Reasonable").sum())
        snull_cnt = int((sub[res_col] == "SNULL").sum())
        unreas_cnt = int((sub[res_col] == "Unreasonable").sum())

        # Method 1 (Old Rule - Loop Status Based): 'Timely closure of the loop' + 'Not through the owner phase'
        closed_old_cnt = int(sub[loop_col].isin(["Timely closure of the loop", "Not through the owner phase"]).sum())
        clos_old_rate = (closed_old_cnt / total_tix * 100) if total_tix > 0 else None

        # Method 2 (New Rule - Handling Result Based): (Reasonable + Unreasonable) / Total, and 100% if Total == 0
        closed_new_cnt = reas_cnt + unreas_cnt
        clos_new_rate = (closed_new_cnt / total_tix * 100) if total_tix > 0 else 100.0

        reas_rate = (reas_cnt / total_tix * 100) if total_tix > 0 else None

        summary_rows.append({
            "LOB": lob,
            "Team HC": hc,
            "Participation": part,
            "Usage Rate per Team (%)": round(usage_rate, 1),
            "Reasonability (%)": round(reas_rate, 1) if reas_rate is not None else None,
            "Closure Status - Loop Based (Old %)": round(clos_old_rate, 1) if clos_old_rate is not None else None,
            "Closure Status - Result Based (New %)": round(clos_new_rate, 1),
            "Total Feedbacks": total_tix
        })

        if total_tix > 0:
            pivot_rows.append({
                "Row Labels": lob,
                "Reasonable": reas_cnt,
                "SNULL": snull_cnt,
                "Unreasonable": unreas_cnt,
                "Grand Total": total_tix,
                "Reasonability (%)": round(reas_rate, 1),
                "Closure Status - Loop (Old %)": round(clos_old_rate, 1) if clos_old_rate is not None else 0.0,
                "Closure Status - Result (New %)": round(clos_new_rate, 1)
            })

    summary_df_out = pd.DataFrame(summary_rows)
    pivot_df_out = pd.DataFrame(pivot_rows)

    # Top Executive KPIs
    total_hc_all = int(summary_df_out["Team HC"].sum())
    total_part_all = int(working_df["Submitter Name"].nunique())
    total_tix_all = len(working_df)
    total_reas_all = int((working_df[res_col] == "Reasonable").sum())
    total_unreas_all = int((working_df[res_col] == "Unreasonable").sum())
    total_snull_all = int((working_df[res_col] == "SNULL").sum())

    total_closed_old_all = int(working_df[loop_col].isin(["Timely closure of the loop", "Not through the owner phase"]).sum())
    total_closed_new_all = total_reas_all + total_unreas_all
    escalated_cnt = int((pending_review_19_df["Escalation status"].astype(str).str.strip().str.upper() == "YES").sum())

    st.subheader("📈 Executive KPIs (Featuring Both Old & New Closure Status Metrics)")
    k1, k2, k3, k4, k5, k6, k7 = st.columns(7)
    k1.metric("Total Feedbacks", total_tix_all)
    k2.metric("Active Participants", f"{total_part_all} / {total_hc_all}")
    k3.metric("Usage Rate", f"{(total_part_all/total_hc_all*100):.1f}%" if total_hc_all else "0%")
    k4.metric("Reasonability Rate", f"{(total_reas_all/total_tix_all*100):.1f}%" if total_tix_all else "0%", f"{total_reas_all} Reasonable")
    k5.metric("Closure (Old - Loop)", f"{(total_closed_old_all/total_tix_all*100):.1f}%" if total_tix_all else "0%", f"{total_closed_old_all} Closed")
    k6.metric("Closure (New - Result)", f"{(total_closed_new_all/total_tix_all*100):.1f}%" if total_tix_all else "0%", f"{total_closed_new_all} Replied")
    k7.metric("Unreasonable / SNULL", f"{total_unreas_all} U / {total_snull_all} S")

    # Charts
    ch1, ch2 = st.columns(2)
    with ch1:
        active_sum = summary_df_out[summary_df_out["Total Feedbacks"] > 0]
        fig_bar = px.bar(
            active_sum,
            x="LOB",
            y=[
                "Usage Rate per Team (%)",
                "Reasonability (%)",
                "Closure Status - Loop Based (Old %)",
                "Closure Status - Result Based (New %)"
            ],
            barmode="group",
            title="Team Performance Comparison (Usage, Reasonability, Old Closure & New Closure %)",
            color_discrete_sequence=["#2563eb", "#10b981", "#0d9488", "#8b5cf6"]
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with ch2:
        fig_pie = px.pie(
            working_df,
            names=res_col,
            title="Handling Result Distribution (Reasonable vs. Unreasonable vs. SNULL)",
            color=res_col,
            color_discrete_map={"Reasonable": "#10b981", "Unreasonable": "#ef4444", "SNULL": "#f59e0b"}
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    # Tabs for Tables
    tab1, tab2, tab3, tab4 = st.tabs([
        "📊 1. Summary (Old & New Closure)",
        "⏳ 2. Pending_Cases_Review (Unreasonable & SNULL)",
        "📈 3. Handling Result Pivot",
        "🗂️ 4. Enriched Raw Data"
    ])
    with tab1:
        st.dataframe(summary_df_out, use_container_width=True)
    with tab2:
        st.dataframe(pending_review_19_df, use_container_width=True)
    with tab3:
        st.dataframe(pivot_df_out, use_container_width=True)
    with tab4:
        st.dataframe(working_df, use_container_width=True)

    # --- PDF Report Generator (100% English, Both Closure Columns, Charts & Pending Cases) ---
    def generate_final_pdf_bytes():
        buf = io.BytesIO()
        doc = SimpleDocTemplate(buf, pagesize=landscape(A4), rightMargin=22, leftMargin=22, topMargin=22, bottomMargin=22)
        styles = getSampleStyleSheet()
        story = []

        title_style = ParagraphStyle("ExecTitle", parent=styles["Heading1"], fontSize=17, textColor=colors.HexColor("#0F172A"), spaceAfter=6)
        sub_style = ParagraphStyle("ExecSub", parent=styles["Normal"], fontSize=9.5, textColor=colors.HexColor("#334155"), spaceAfter=10)
        cell_wrap_style = ParagraphStyle("CellWrap", parent=styles["Normal"], fontSize=7, leading=8.5, textColor=colors.HexColor("#1E293B"))

        story.append(Paragraph("Executive Feedback Analysis Report", title_style))
        story.append(Paragraph(
            f"<b>Total Feedbacks:</b> {total_tix_all} &nbsp;|&nbsp; "
            f"<b>Participants:</b> {total_part_all}/{total_hc_all} ({(total_part_all/total_hc_all*100 if total_hc_all else 0):.1f}% Usage) &nbsp;|&nbsp; "
            f"<b>Reasonability:</b> {(total_reas_all/total_tix_all*100 if total_tix_all else 0):.1f}% ({total_reas_all}) &nbsp;|&nbsp; "
            f"<b>Closure (Old Loop):</b> {(total_closed_old_all/total_tix_all*100 if total_tix_all else 0):.1f}% &nbsp;|&nbsp; "
            f"<b>Closure (New Result):</b> {(total_closed_new_all/total_tix_all*100 if total_tix_all else 0):.1f}% &nbsp;|&nbsp; "
            f"<b>Unreasonable:</b> {total_unreas_all} &nbsp;|&nbsp; <b>SNULL:</b> {total_snull_all}",
            sub_style
        ))

        # Matplotlib Charts for PDF
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11.3, 3.4), gridspec_kw={"width_ratios": [2.3, 1]})
        active_s = summary_df_out[summary_df_out["Total Feedbacks"] > 0].fillna(0)
        if len(active_s) == 0:
            active_s = summary_df_out.fillna(0)
        x = range(len(active_s))
        w = 0.20
        ax1.bar([i - 1.5*w for i in x], active_s["Usage Rate per Team (%)"], width=w, label="Usage Rate (%)", color="#2563eb")
        ax1.bar([i - 0.5*w for i in x], active_s["Reasonability (%)"], width=w, label="Reasonability (%)", color="#10b981")
        ax1.bar([i + 0.5*w for i in x], active_s["Closure Status - Loop Based (Old %)"], width=w, label="Closure Old (Loop %)", color="#0d9488")
        ax1.bar([i + 1.5*w for i in x], active_s["Closure Status - Result Based (New %)"], width=w, label="Closure New (Result %)", color="#8b5cf6")
        ax1.set_xticks(list(x))
        ax1.set_xticklabels(active_s["LOB"], rotation=28, ha="right", fontsize=8)
        ax1.set_ylim(0, 115)
        ax1.set_ylabel("Percentage (%)")
        ax1.set_title("Active Teams Performance Breakdown (%)", fontsize=10, fontweight="bold")
        ax1.legend(fontsize=7.5, loc="upper right", ncol=2)
        ax1.grid(axis="y", linestyle="--", alpha=0.4)

        res_counts = working_df[res_col].value_counts()
        pie_colors = [{"Reasonable": "#10b981", "Unreasonable": "#ef4444", "SNULL": "#f59e0b"}.get(k, "#64748b") for k in res_counts.index]
        ax2.pie(res_counts.values, labels=res_counts.index, autopct="%1.1f%%", startangle=140, colors=pie_colors, textprops={"fontsize": 8.5})
        ax2.set_title("Handling Result Share", fontsize=10, fontweight="bold")

        plt.tight_layout()
        img_buf = io.BytesIO()
        plt.savefig(img_buf, format="png", dpi=160)
        plt.close(fig)
        img_buf.seek(0)
        story.append(Image(img_buf, width=765, height=220))
        story.append(Spacer(1, 8))

        # Section 1: Summary Table with Both Closure Columns
        story.append(Paragraph("1. Team Summary Table (With Old & New Closure Status)", styles["Heading2"]))
        sum_table_data = [[
            "LOB (Team)", "Team HC", "Participation", "Usage Rate (%)",
            "Reasonability (%)", "Closure Old (Loop %)", "Closure New (Result %)", "Total Feedbacks"
        ]]
        for _, r in summary_df_out.iterrows():
            sum_table_data.append([
                str(r["LOB"]),
                str(r["Team HC"]),
                str(r["Participation"]),
                f"{r['Usage Rate per Team (%)']:.1f}%",
                f"{r['Reasonability (%)']:.1f}%" if pd.notna(r["Reasonability (%)"]) else "-",
                f"{r['Closure Status - Loop Based (Old %)']:.1f}%" if pd.notna(r["Closure Status - Loop Based (Old %)"]) else "-",
                f"{r['Closure Status - Result Based (New %)']:.1f}%",
                str(r["Total Feedbacks"])
            ])
        sum_table_data.append([
            "Grand Total",
            str(total_hc_all),
            str(total_part_all),
            f"{(total_part_all/total_hc_all*100 if total_hc_all else 0):.1f}%",
            f"{(total_reas_all/total_tix_all*100 if total_tix_all else 0):.1f}%",
            f"{(total_closed_old_all/total_tix_all*100 if total_tix_all else 0):.1f}%",
            f"{(total_closed_new_all/total_tix_all*100 if total_tix_all else 0):.1f}%",
            str(total_tix_all)
        ])

        t_sum = Table(sum_table_data, colWidths=[135, 58, 72, 85, 95, 108, 112, 85], repeatRows=1)
        t_sum.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2E8F0")),
            ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
            ("ALIGN", (1, 0), (-1, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F8FAFC")]),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ]))
        story.append(t_sum)
        story.append(PageBreak())

        # Section 2: Pending & Unreasonable Cases Table
        story.append(Paragraph("2. Pending & Unreasonable Cases (Pending_Cases_Review)", styles["Heading2"]))
        act_data = [[
            "Order No.", "Submitter Name", "Issue Type", "Knowledge ID",
            "Result", "Severity", "Escalated", "Relevant Team", "Review Comment / Action Taken"
        ]]
        for _, r in pending_review_19_df.iterrows():
            comment_txt = str(r.get("Comment", "") if pd.notna(r.get("Comment", "")) else "").strip()
            if comment_txt == "nan":
                comment_txt = ""
            act_data.append([
                str(r.get("Order No.", "")),
                Paragraph(str(r.get("Submitter Name", "")), cell_wrap_style),
                Paragraph(str(r.get("Issue Type", "")), cell_wrap_style),
                str(r.get("Knowledge ID", "")),
                str(r.get(res_col, "")),
                str(r.get("Severity", "")) if pd.notna(r.get("Severity", "")) else "-",
                str(r.get("Escalation status", "")) if pd.notna(r.get("Escalation status", "")) else "-",
                Paragraph(str(r.get("Relevant Team", "")), cell_wrap_style),
                Paragraph(comment_txt[:260], cell_wrap_style)
            ])

        t_act = Table(act_data, colWidths=[82, 95, 105, 68, 62, 48, 48, 85, 195], repeatRows=1)
        t_act.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (3, 0), (6, -1), "CENTER"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
            ("FONTSIZE", (0, 0), (-1, -1), 7.5),
        ]))
        story.append(t_act)

        doc.build(story)
        buf.seek(0)
        return buf.getvalue()

    dl1, dl2 = st.columns(2)
    with dl1:
        st.download_button(
            label="📄 Download Final Executive PDF Report (With Both Closure Metrics & Updated Results)",
            data=generate_final_pdf_bytes(),
            file_name="Final_Feedback_Analysis_Report.pdf",
            mime="application/pdf",
            type="primary",
            use_container_width=True
        )
    with dl2:
        final_wb_bytes = create_styled_excel(
            {
                "Summary": summary_df_out,
                "Handling_Result_Pivot": pivot_df_out,
                "Pending_Cases_Review": pending_review_19_df,
                "Enriched_Raw_Data": working_df
            },
            add_dropdowns_on_sheet="Pending_Cases_Review"
        )
        st.download_button(
            label="📥 Download Final Complete Excel Workbook (.xlsx)",
            data=final_wb_bytes,
            file_name="Final_Feedback_Analysis_Workbook.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
