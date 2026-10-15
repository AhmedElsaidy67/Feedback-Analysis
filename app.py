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
    page_title="Feedback Analysis & Two-Stage Workflow Tool",
    layout="wide",
    page_icon="📊"
)

ROSTER_FILE = "Master_Roster_DB.xlsx"

# --- Default Fallback Roster & Team HC ---
DEFAULT_HC = [
    {"LOB": "Chat Team", "Team HC": 5},
    {"LOB": "CHR - Egypt", "Team HC": 5},
    {"LOB": "CHR - MEA", "Team HC": 7},
    {"LOB": "CSE Team", "Team HC": 5},
    {"LOB": "E-Commerce", "Team HC": 9},
    {"LOB": "Egypt Hotline", "Team HC": 14},
    {"LOB": "Egypt Supervisor", "Team HC": 1},
    {"LOB": "E-Mail Team", "Team HC": 4},
    {"LOB": "MEA Hotline", "Team HC": 8},
    {"LOB": "Outbound", "Team HC": 9},
    {"LOB": "Quality Department", "Team HC": 6},
    {"LOB": "Social Media Arabic", "Team HC": 2},
    {"LOB": "Social Media English", "Team HC": 11},
    {"LOB": "South Africa Team", "Team HC": 9},
    {"LOB": "Training Department", "Team HC": 3},
    {"LOB": "UAE Supervisor", "Team HC": 1},
    {"LOB": "WhatsApp", "Team HC": 8},
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

# --- Sidebar: Master Roster & Team HC Management ---
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
            st.sidebar.success("✅ Master Roster database updated and saved!")
        else:
            st.sidebar.error("Sheet must contain columns: Consignee, Submitter Name, Submitter LOB")
    except Exception as e:
        st.sidebar.error(f"Error reading Roster file: {e}")

with st.sidebar.expander("📝 View / Edit Roster & Team HC"):
    st.caption("Edit or add rows directly below and click Save:")
    edited_agents = st.data_editor(agents_df, num_rows="dynamic", key="edit_agents")
    edited_hc = st.data_editor(hc_df, num_rows="dynamic", key="edit_hc")
    if st.button("💾 Save Roster Changes"):
        save_master_roster(edited_agents, edited_hc)
        agents_df, hc_df = load_master_roster()
        st.success("Saved successfully!")

# --- Helper Function: Map Consignee to Submitter Name & LOB ---
def enrich_raw_data_with_roster(df, agents_df):
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
    return df

# --- Helper Function: Build Pending Cases DataFrame (Exact 19-Column Structure) ---
def build_pending_cases_df(raw_df):
    res_col = "The responsible person replies with the handling result."
    raw_df[res_col] = raw_df[res_col].fillna("SNULL").astype(str).str.strip()

    pending_raw = raw_df[raw_df[res_col].isin(["Unreasonable", "SNULL"])].copy()

    # Exact columns from original 'Data till 12 Oct' + Severity, Escalation status, Relevant Team (= Submitter LOB)
    sol_col_1 = "Responsible person replies with a handling solution"
    sol_col_2 = "Responsible person replies with a handling solution.1" if "Responsible person replies with a handling solution.1" in pending_raw.columns else sol_col_1

    out_df = pd.DataFrame()
    out_df["Order No."] = pending_raw.get("Order No.", "")
    out_df["Bill of Lading Time"] = pending_raw.get("Bill of Lading Time", "")
    out_df["Issue Type"] = pending_raw.get("Issue Type", "")
    out_df["Country/Region Where the Issue Occurred"] = pending_raw.get("Country/Region Where the Issue Occurred", "")
    out_df["Process Type"] = pending_raw.get("Process Type", "")
    out_df["Consignee"] = pending_raw.get("Consignee", "")
    out_df["Submitter Name"] = pending_raw.get("Submitter Name", "")
    out_df["Submitter LOB"] = pending_raw.get("Submitter LOB", "")
    out_df["Knowledge ID"] = pending_raw.get("Knowledge ID", "")
    out_df["Responsible person replies with a handling solution (Initial)"] = pending_raw.get(sol_col_1, "SNULL")
    out_df["Has the responsible person promptly closed the loop?"] = pending_raw.get("Has the responsible person promptly closed the loop?", "")
    out_df["Person Responsible"] = pending_raw.get("Person Responsible", "SNULL")
    out_df["Responsible person replies with a handling solution (Final)"] = pending_raw.get(sol_col_2, "SNULL")
    out_df["The responsible person replies with the handling result."] = pending_raw.get(res_col, "SNULL")
    out_df["Electronic Flow Link"] = pending_raw.get("Electronic Flow Link", "")
    out_df["Comment"] = ""
    out_df["Severity"] = ""
    out_df["Escalation status"] = "Yes"
    # Relevant Team is automatically set to Submitter LOB as requested
    out_df["Relevant Team"] = pending_raw.get("Submitter LOB", "")

    return out_df

# --- Helper Function: Style Excel Export for Stage 1 & Stage 2 ---
def create_styled_excel(sheets_dict, add_dropdowns_on_sheet=None):
    wb = openpyxl.Workbook()
    # remove default sheet
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    editable_header_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    regular_font = Font(name="Calibri", size=11, color="1E293B")
    bold_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    editable_cols = {"Comment", "Severity", "Escalation status", "Relevant Team", "The responsible person replies with the handling result."}

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(title=sheet_name)
        cols = list(df.columns)

        # Write Header
        for col_idx, col_name in enumerate(cols, start=1):
            cell = ws.cell(1, col_idx, col_name)
            cell.fill = editable_header_fill if (sheet_name == add_dropdowns_on_sheet and col_name in editable_cols) else header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        # Write Rows
        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            for c_idx, val in enumerate(row, start=1):
                cell = ws.cell(r_idx, c_idx, "" if pd.isna(val) else val)
                cell.font = bold_font if c_idx == 1 else regular_font
                if r_idx % 2 == 1:
                    cell.fill = zebra_fill
                cell.border = thin_border
                col_name = cols[c_idx - 1]
                if col_name in ["Comment", "Electronic Flow Link", "Issue Type"]:
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=(col_name == "Comment"))
                else:
                    cell.alignment = Alignment(horizontal="center" if c_idx > 1 else "left", vertical="center")

        ws.freeze_panes = "A2"
        if len(df) > 0:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(df) + 1}"

        # Adjust column widths
        for c_idx, col_name in enumerate(cols, start=1):
            col_letter = get_column_letter(c_idx)
            if col_name == "Comment":
                ws.column_dimensions[col_letter].width = 45
            elif col_name in ["Issue Type", "Has the responsible person promptly closed the loop?"]:
                ws.column_dimensions[col_letter].width = 34
            elif col_name in ["Submitter Name", "Electronic Flow Link"]:
                ws.column_dimensions[col_letter].width = 28
            else:
                ws.column_dimensions[col_letter].width = 20

        # Add Excel Dropdowns (DataValidation) on Pending_Cases_Review sheet for fast manual review
        if sheet_name == add_dropdowns_on_sheet and len(df) > 0:
            max_r = len(df) + 1
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

# --- Main Page UI ---
st.title("📊 Feedback Analysis & Two-Stage Workflow Automation")
st.markdown(
    "This tool operates in **Two Seamless Stages**:  \n"
    "1. **Stage 1 (Enrichment & Pending Cases Extraction):** Upload the raw export (even without `Submitter Name` or `Submitter LOB`). The tool matches `Consignee` against the Master Roster, populates `Submitter Name` & `Submitter LOB` (and sets `Relevant Team = Submitter LOB`), and exports the **`Pending_Cases_Review.xlsx`** file (`Unreasonable` & `SNULL` cases) for your manual review.  \n"
    "2. **Stage 2 (Final Analysis & Executive Reporting):** Upload the manually updated `Pending_Cases_Review.xlsx` file to automatically compute the final **Summary**, **Handling Result Breakdown**, interactive **Charts**, **PDF Executive Report**, and final **Excel Workbook**."
)

stage1_tab, stage2_tab = st.tabs([
    "🔹 Stage 1: Upload Raw Data & Export Pending Cases for Manual Review",
    "🔸 Stage 2: Upload Reviewed Sheet & Generate Final Analysis + PDF Report"
])

# ==============================================================================
# STAGE 1: UPLOAD RAW DATA -> MATCH CONSIGNEE -> EXPORT PENDING CASES
# ==============================================================================
with stage1_tab:
    st.subheader("Step 1: Upload Raw Data Sheet")
    raw_file_s1 = st.file_uploader(
        "📂 Upload Raw Data (.xlsx, .xls, or .csv) — Submitter Name & Submitter LOB will be added automatically via Consignee:",
        type=["xlsx", "xls", "csv"],
        key="s1_uploader"
    )

    if raw_file_s1 is not None:
        if raw_file_s1.name.endswith(".csv"):
            raw_df_s1 = pd.read_csv(raw_file_s1)
        else:
            xls_s1 = pd.ExcelFile(raw_file_s1)
            sheet_to_read = "Raw Data" if "Raw Data" in xls_s1.sheet_names else xls_s1.sheet_names[0]
            raw_df_s1 = pd.read_excel(xls_s1, sheet_name=sheet_to_read)

        if "Consignee" not in raw_df_s1.columns:
            st.error("❌ Column 'Consignee' was not found in the uploaded file!")
            st.stop()

        enriched_raw_df = enrich_raw_data_with_roster(raw_df_s1, agents_df)
        pending_df_s1 = build_pending_cases_df(enriched_raw_df)

        # Save in session_state so Stage 2 can also access the enriched Raw Data automatically
        st.session_state["enriched_raw_df"] = enriched_raw_df

        # Alert for any unmapped Consignees
        unmapped_mask = (
            enriched_raw_df["Submitter Name"].isna()
            | (enriched_raw_df["Submitter Name"] == "Unknown Agent")
            | (enriched_raw_df["Submitter Name"] == "nan")
        )
        if unmapped_mask.any():
            unmapped_ids = enriched_raw_df.loc[unmapped_mask, "Consignee"].unique().tolist()
            st.warning(
                f"⚠️ **Unmapped Consignee IDs Detected ({len(unmapped_ids)}):** `{', '.join(unmapped_ids)}`  \n"
                "Please add them in the Sidebar under **'View / Edit Roster & Team HC'** so they are permanently mapped."
            )
        else:
            st.success(f"✅ All `{len(enriched_raw_df)}` rows successfully matched with `Submitter Name` and `Submitter LOB`!")

        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.metric("Total Raw Data Rows", len(enriched_raw_df))
        c_m2.metric("Extracted Pending Cases (Unreasonable / SNULL)", len(pending_df_s1))
        c_m3.metric("Unique Participating Agents", enriched_raw_df["Submitter Name"].nunique())

        st.markdown("#### 📋 Preview of Extracted `Pending_Cases_Review` (Unreasonable & SNULL Cases)")
        st.caption("Notice that `Submitter Name`, `Submitter LOB`, and `Relevant Team` (= `Submitter LOB`) are automatically populated. Dropdowns for `Severity` and `Escalation status` are included in the downloaded Excel file.")
        st.dataframe(pending_df_s1, use_container_width=True)

        # Create Stage 1 Excel Workbook containing BOTH 'Pending_Cases_Review' and 'Enriched_Raw_Data'
        stage1_excel_bytes = create_styled_excel(
            {
                "Pending_Cases_Review": pending_df_s1,
                "Enriched_Raw_Data": enriched_raw_df
            },
            add_dropdowns_on_sheet="Pending_Cases_Review"
        )

        st.download_button(
            label="📥 Download 'Pending_Cases_Review.xlsx' for Manual Review & Update",
            data=stage1_excel_bytes,
            file_name="Pending_Cases_Review.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )

# ==============================================================================
# STAGE 2: UPLOAD REVIEWED PENDING CASES -> FINAL ANALYSIS, CHARTS, PDF & EXCEL
# ==============================================================================
with stage2_tab:
    st.subheader("Step 2: Upload Manually Updated `Pending_Cases_Review.xlsx`")
    st.markdown(
        "Upload the **`Pending_Cases_Review.xlsx`** file after completing your manual updates (`Comment`, `Severity`, `Escalation status`, `Relevant Team`, or updated handling results).  \n"
        "*Note: Because the Stage 1 file includes the `Enriched_Raw_Data` tab inside it, you only need to upload that single file here!*"
    )

    reviewed_file = st.file_uploader(
        "📂 Upload Reviewed `Pending_Cases_Review.xlsx` File:",
        type=["xlsx", "xls"],
        key="s2_uploader"
    )

    optional_raw_s2 = st.file_uploader(
        "📂 (Optional) Upload Raw Data if your reviewed file does not contain the 'Enriched_Raw_Data' tab:",
        type=["xlsx", "xls", "csv"],
        key="s2_raw_optional"
    )

    if reviewed_file is not None:
        xls_rev = pd.ExcelFile(reviewed_file)

        # Read Reviewed Pending Cases sheet
        rev_sheet_name = "Pending_Cases_Review" if "Pending_Cases_Review" in xls_rev.sheet_names else xls_rev.sheet_names[0]
        reviewed_pending_df = pd.read_excel(xls_rev, sheet_name=rev_sheet_name)

        # Determine Full Raw Data source:
        # Priority 1: Optional Raw Data uploaded in Stage 2
        # Priority 2: 'Enriched_Raw_Data' or 'Raw Data' sheet inside the uploaded Stage 2 workbook
        # Priority 3: Session state from Stage 1
        if optional_raw_s2 is not None:
            if optional_raw_s2.name.endswith(".csv"):
                base_raw_df = pd.read_csv(optional_raw_s2)
            else:
                xls_opt = pd.ExcelFile(optional_raw_s2)
                s_opt = "Raw Data" if "Raw Data" in xls_opt.sheet_names else xls_opt.sheet_names[0]
                base_raw_df = pd.read_excel(xls_opt, sheet_name=s_opt)
            full_raw_df = enrich_raw_data_with_roster(base_raw_df, agents_df)
        elif "Enriched_Raw_Data" in xls_rev.sheet_names:
            full_raw_df = pd.read_excel(xls_rev, sheet_name="Enriched_Raw_Data")
        elif "Raw Data" in xls_rev.sheet_names:
            full_raw_df = enrich_raw_data_with_roster(pd.read_excel(xls_rev, sheet_name="Raw Data"), agents_df)
        elif "enriched_raw_df" in st.session_state:
            full_raw_df = st.session_state["enriched_raw_df"].copy()
        else:
            st.warning("⚠️ Full Raw Data tab was not found in the workbook. Computing analysis from the uploaded sheet only, or please upload the Raw Data file above.")
            full_raw_df = enrich_raw_data_with_roster(reviewed_pending_df.copy(), agents_df)

        # Synchronize any manual edits made in reviewed_pending_df back into full_raw_df by Order No.
        res_col = "The responsible person replies with the handling result."
        loop_col = "Has the responsible person promptly closed the loop?"

        reviewed_pending_df["Order No."] = reviewed_pending_df["Order No."].astype(str).str.strip()
        full_raw_df["Order No."] = full_raw_df["Order No."].astype(str).str.strip()

        #Ensure Relevant Team defaults to Submitter LOB if left blank
        if "Relevant Team" in reviewed_pending_df.columns and "Submitter LOB" in reviewed_pending_df.columns:
            reviewed_pending_df["Relevant Team"] = (
                reviewed_pending_df["Relevant Team"]
                .fillna(reviewed_pending_df["Submitter LOB"])
                .replace("", pd.NA)
                .fillna(reviewed_pending_df["Submitter LOB"])
            )

        # Sync updated Handling Result or Submitter LOB if changed during manual review
        if res_col in reviewed_pending_df.columns:
            res_update_map = dict(zip(reviewed_pending_df["Order No."], reviewed_pending_df[res_col].astype(str).str.strip()))
            full_raw_df[res_col] = full_raw_df["Order No."].map(res_update_map).fillna(full_raw_df[res_col])

        # Also merge manual columns (Comment, Severity, Escalation status, Relevant Team) onto full_raw_df
        manual_cols = ["Order No."] + [c for c in ["Comment", "Severity", "Escalation status", "Relevant Team"] if c in reviewed_pending_df.columns]
        manual_sub = reviewed_pending_df[manual_cols].drop_duplicates(subset=["Order No."])
        full_raw_df = pd.merge(
            full_raw_df.drop(columns=[c for c in ["Comment", "Severity", "Escalation status", "Relevant Team"] if c in full_raw_df.columns], errors="ignore"),
            manual_sub,
            on="Order No.",
            how="left"
        )

        full_raw_df[res_col] = full_raw_df[res_col].fillna("SNULL").astype(str).str.strip()
        full_raw_df[loop_col] = full_raw_df[loop_col].fillna("").astype(str).str.strip()
        full_raw_df["Submitter LOB"] = full_raw_df["Submitter LOB"].astype(str).str.strip()

        # --- Compute Summary & Pivot Tables ---
        all_lobs = sorted(list(set(hc_df["LOB"].tolist() + full_raw_df["Submitter LOB"].unique().tolist())))
        hc_map = dict(zip(hc_df["LOB"], hc_df["Team HC"]))

        summary_rows = []
        pivot_rows = []

        for lob in all_lobs:
            sub = full_raw_df[full_raw_df["Submitter LOB"] == lob]
            hc = int(hc_map.get(lob, len(sub["Submitter Name"].unique())))
            part = sub["Submitter Name"].nunique() if len(sub) > 0 else 0
            total_tix = len(sub)
            usage_rate = (part / hc * 100) if hc > 0 else 0.0

            reas_cnt = int((sub[res_col] == "Reasonable").sum())
            snull_cnt = int((sub[res_col] == "SNULL").sum())
            unreas_cnt = int((sub[res_col] == "Unreasonable").sum())
            closed_cnt = int(sub[loop_col].isin(["Timely closure of the loop", "Not through the owner phase"]).sum())

            reas_rate = (reas_cnt / total_tix * 100) if total_tix > 0 else None
            clos_rate = (closed_cnt / total_tix * 100) if total_tix > 0 else None

            summary_rows.append({
                "LOB": lob,
                "Team HC": hc,
                "Participation": part,
                "Usage Rate per Team (%)": round(usage_rate, 1),
                "Reasonability (%)": round(reas_rate, 1) if reas_rate is not None else None,
                "Closure Status (%)": round(clos_rate, 1) if clos_rate is not None else None,
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
                    "Closure Status (%)": round(clos_rate, 1)
                })

        summary_df_out = pd.DataFrame(summary_rows)
        pivot_df_out = pd.DataFrame(pivot_rows)

        # --- Top Executive KPIs ---
        total_hc_all = int(summary_df_out["Team HC"].sum())
        total_part_all = int(full_raw_df["Submitter Name"].nunique())
        total_tix_all = len(full_raw_df)
        total_reas_all = int((full_raw_df[res_col] == "Reasonable").sum())
        total_closed_all = int(full_raw_df[loop_col].isin(["Timely closure of the loop", "Not through the owner phase"]).sum())
        escalated_cnt = int((reviewed_pending_df.get("Escalation status", pd.Series(dtype=str)).astype(str).str.strip().str.upper() == "YES").sum())

        st.success("✅ Final Analysis Completed Successfully!")
        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("Total Feedbacks", total_tix_all)
        k2.metric("Active Participants", f"{total_part_all} / {total_hc_all}")
        k3.metric("Overall Usage Rate", f"{(total_part_all/total_hc_all*100):.1f}%" if total_hc_all else "0%")
        k4.metric("Reasonability Rate", f"{(total_reas_all/total_tix_all*100):.1f}%" if total_tix_all else "0%")
        k5.metric("Loop Closure Rate", f"{(total_closed_all/total_tix_all*100):.1f}%" if total_tix_all else "0%")
        k6.metric("Escalated Cases", f"{escalated_cnt} / {len(reviewed_pending_df)}")

        # --- Interactive Charts ---
        ch1, ch2 = st.columns(2)
        with ch1:
            active_sum = summary_df_out[summary_df_out["Total Feedbacks"] > 0]
            fig_bar = px.bar(
                active_sum,
                x="LOB",
                y=["Usage Rate per Team (%)", "Reasonability (%)", "Closure Status (%)"],
                barmode="group",
                title="Team Performance Comparison (Usage vs. Reasonability vs. Closure %)",
                color_discrete_sequence=["#2563eb", "#10b981", "#14b8a6"]
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        with ch2:
            fig_pie = px.pie(
                full_raw_df,
                names=res_col,
                title="Overall Handling Result Distribution",
                color=res_col,
                color_discrete_map={"Reasonable": "#10b981", "Unreasonable": "#ef4444", "SNULL": "#f59e0b"}
            )
            st.plotly_chart(fig_pie, use_container_width=True)

        ch3, ch4 = st.columns(2)
        with ch3:
            if "Severity" in reviewed_pending_df.columns and reviewed_pending_df["Severity"].dropna().astype(str).str.strip().ne("").any():
                sev_counts = reviewed_pending_df["Severity"].fillna("Unspecified").replace("", "Unspecified").value_counts().reset_index()
                sev_counts.columns = ["Severity", "Count"]
                fig_sev = px.bar(
                    sev_counts,
                    x="Severity",
                    y="Count",
                    color="Severity",
                    title="Reviewed Pending Cases by Severity",
                    color_discrete_map={"High": "#ef4444", "Medium": "#f59e0b", "Low": "#3b82f6", "Unspecified": "#94a3b8"}
                )
                st.plotly_chart(fig_sev, use_container_width=True)
        with ch4:
            if "Relevant Team" in reviewed_pending_df.columns:
                rel_counts = reviewed_pending_df["Relevant Team"].fillna("Unspecified").value_counts().reset_index()
                rel_counts.columns = ["Relevant Team", "Count"]
                fig_rel = px.bar(
                    rel_counts,
                    x="Relevant Team",
                    y="Count",
                    title="Pending & Unreasonable Cases by Relevant Team (Submitter LOB)",
                    color_discrete_sequence=["#6366f1"]
                )
                st.plotly_chart(fig_rel, use_container_width=True)

        # --- Data Preview Tabs ---
        rt1, rt2, rt3, rt4 = st.tabs([
            "📊 1. Team Summary",
            "📝 2. Reviewed Pending Cases (Unreasonable & SNULL)",
            "📈 3. Handling Result Pivot",
            "🗂️ 4. Enriched Raw Data"
        ])
        with rt1:
            st.dataframe(summary_df_out, use_container_width=True)
        with rt2:
            st.dataframe(reviewed_pending_df, use_container_width=True)
        with rt3:
            st.dataframe(pivot_df_out, use_container_width=True)
        with rt4:
            st.dataframe(full_raw_df, use_container_width=True)

        # --- PDF Report Generator (100% English, Charts + Percentages + Reviewed Table) ---
        def generate_final_pdf_bytes():
            buf = io.BytesIO()
            doc = SimpleDocTemplate(
                buf,
                pagesize=landscape(A4),
                rightMargin=24,
                leftMargin=24,
                topMargin=24,
                bottomMargin=24
            )
            styles = getSampleStyleSheet()
            story = []

            title_style = ParagraphStyle(
                "ExecTitle",
                parent=styles["Heading1"],
                fontSize=18,
                textColor=colors.HexColor("#0F172A"),
                spaceAfter=6
            )
            sub_style = ParagraphStyle(
                "ExecSub",
                parent=styles["Normal"],
                fontSize=10,
                textColor=colors.HexColor("#334155"),
                spaceAfter=12
            )
            cell_wrap_style = ParagraphStyle(
                "CellWrap",
                parent=styles["Normal"],
                fontSize=7,
                leading=8.5,
                textColor=colors.HexColor("#1E293B")
            )

            story.append(Paragraph("Executive Feedback Analysis Report", title_style))
            story.append(Paragraph(
                f"<b>Total Feedbacks:</b> {total_tix_all} &nbsp;|&nbsp; "
                f"<b>Active Participants:</b> {total_part_all} / {total_hc_all} ({(total_part_all/total_hc_all*100 if total_hc_all else 0):.1f}% Usage Rate) &nbsp;|&nbsp; "
                f"<b>Overall Reasonability:</b> {(total_reas_all/total_tix_all*100 if total_tix_all else 0):.1f}% &nbsp;|&nbsp; "
                f"<b>Overall Closure Rate:</b> {(total_closed_all/total_tix_all*100 if total_tix_all else 0):.1f}% &nbsp;|&nbsp; "
                f"<b>Reviewed Pending/Unreasonable Cases:</b> {len(reviewed_pending_df)} (Escalated: {escalated_cnt})",
                sub_style
            ))

            # Render High-Res Matplotlib Charts for PDF
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11.2, 3.5), gridspec_kw={"width_ratios": [2.2, 1]})
            active_s = summary_df_out[summary_df_out["Total Feedbacks"] > 0].fillna(0)
            if len(active_s) == 0:
                active_s = summary_df_out.fillna(0)
            x = range(len(active_s))
            w = 0.26
            ax1.bar([i - w for i in x], active_s["Usage Rate per Team (%)"], width=w, label="Usage Rate (%)", color="#2563eb")
            ax1.bar(x, active_s["Reasonability (%)"], width=w, label="Reasonability (%)", color="#10b981")
            ax1.bar([i + w for i in x], active_s["Closure Status (%)"], width=w, label="Closure Status (%)", color="#14b8a6")
            ax1.set_xticks(list(x))
            ax1.set_xticklabels(active_s["LOB"], rotation=30, ha="right", fontsize=8)
            ax1.set_ylim(0, 115)
            ax1.set_ylabel("Percentage (%)")
            ax1.set_title("Active Teams Performance Breakdown (%)", fontsize=10, fontweight="bold")
            ax1.legend(fontsize=8, loc="upper right")
            ax1.grid(axis="y", linestyle="--", alpha=0.4)

            res_counts = full_raw_df[res_col].value_counts()
            pie_colors = [{"Reasonable": "#10b981", "Unreasonable": "#ef4444", "SNULL": "#f59e0b"}.get(k, "#64748b") for k in res_counts.index]
            ax2.pie(
                res_counts.values,
                labels=res_counts.index,
                autopct="%1.1f%%",
                startangle=140,
                colors=pie_colors,
                textprops={"fontsize": 8.5}
            )
            ax2.set_title("Handling Result Distribution", fontsize=10, fontweight="bold")

            plt.tight_layout()
            img_buf = io.BytesIO()
            plt.savefig(img_buf, format="png", dpi=160)
            plt.close(fig)
            img_buf.seek(0)
            story.append(Image(img_buf, width=760, height=225))
            story.append(Spacer(1, 10))

            # Section 1: Team Summary Table
            story.append(Paragraph("1. Team Summary & KPIs Table", styles["Heading2"]))
            sum_table_data = [[
                "LOB (Team)", "Team HC", "Participation", "Usage Rate (%)",
                "Reasonability (%)", "Closure Status (%)", "Total Feedbacks"
            ]]
            for _, r in summary_df_out.iterrows():
                sum_table_data.append([
                    str(r["LOB"]),
                    str(r["Team HC"]),
                    str(r["Participation"]),
                    f"{r['Usage Rate per Team (%)']:.1f}%",
                    f"{r['Reasonability (%)']:.1f}%" if pd.notna(r["Reasonability (%)"]) else "-",
                    f"{r['Closure Status (%)']:.1f}%" if pd.notna(r["Closure Status (%)"]) else "-",
                    str(r["Total Feedbacks"])
                ])
            # Add Grand Total Row
            sum_table_data.append([
                "Grand Total",
                str(total_hc_all),
                str(total_part_all),
                f"{(total_part_all/total_hc_all*100 if total_hc_all else 0):.1f}%",
                f"{(total_reas_all/total_tix_all*100 if total_tix_all else 0):.1f}%",
                f"{(total_closed_all/total_tix_all*100 if total_tix_all else 0):.1f}%",
                str(total_tix_all)
            ])

            t_sum = Table(sum_table_data, colWidths=[150, 75, 85, 105, 110, 110, 100], repeatRows=1)
            t_sum.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E2E8F0")),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -2), [colors.white, colors.HexColor("#F8FAFC")]),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ]))
            story.append(t_sum)
            story.append(PageBreak())

            # Section 2: Reviewed Pending & Unreasonable Cases Table
            story.append(Paragraph("2. Reviewed Pending & Unreasonable Cases (Manual Review Outcomes)", styles["Heading2"]))
            act_data = [[
                "Order No.", "Submitter Name", "Issue Type", "Knowledge ID",
                "Result", "Severity", "Escalated", "Relevant Team", "Review Comment / Action Taken"
            ]]
            for _, r in reviewed_pending_df.iterrows():
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

        # --- Final Export Buttons ---
        exp_c1, exp_c2 = st.columns(2)
        with exp_c1:
            pdf_report_bytes = generate_final_pdf_bytes()
            st.download_button(
                label="📄 Download Final Executive PDF Report (Charts, KPIs & Reviewed Cases)",
                data=pdf_report_bytes,
                file_name="Final_Feedback_Analysis_Report.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
        with exp_c2:
            final_excel_bytes = create_styled_excel(
                {
                    "Summary": summary_df_out,
                    "Handling_Result_Pivot": pivot_df_out,
                    "Pending_Cases_Review": reviewed_pending_df,
                    "Enriched_Raw_Data": full_raw_df
                },
                add_dropdowns_on_sheet="Pending_Cases_Review"
            )
            st.download_button(
                label="📥 Download Final Complete Excel Workbook (.xlsx)",
                data=final_excel_bytes,
                file_name="Final_Feedback_Analysis_Workbook.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
