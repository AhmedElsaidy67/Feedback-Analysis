import io
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from shared_utils import load_master_roster, save_master_roster

TOP3_POSITIONS = {"0", "1", "2", "0.0", "1.0", "2.0"}
TOP10_POSITIONS = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "0.0", "1.0", "2.0", "3.0", "4.0", "5.0", "6.0", "7.0", "8.0", "9.0"}

def assign_week_bucket(date_val):
    """
    Converts integer/string date like '20260901' or '2026-09-01' into standard weekly period buckets
    matching 'Trend Per Team' & 'Trend Per Language Searching' (e.g. '1 September to 7 September').
    """
    s = str(date_val).strip().replace("-", "").replace("/", "")[:8]
    if len(s) == 8 and s.isdigit():
        dt = pd.to_datetime(s, format="%Y%m%d", errors="coerce")
    else:
        dt = pd.to_datetime(date_val, errors="coerce")
    if pd.isna(dt):
        return "Unknown Period"
    month_name = dt.strftime("%B")
    day = dt.day
    last_day = dt.days_in_month
    if day <= 7:
        return f"1 {month_name} to 7 {month_name}"
    elif day <= 14:
        return f"8 {month_name} to 14 {month_name}"
    elif day <= 21:
        return f"15 {month_name} to 21 {month_name}"
    else:
        return f"22 {month_name} to {last_day} {month_name}"


def compute_pos_metrics(sub_df):
    """
    Computes exact click position metrics matching the manual Excel formulas:
    - Grand Total = Total rows
    - Top 3 Clicks = Position in {0, 1, 2}
    - Top 10 Clicks = Position in {0, 1, 2, 3, 4, 5, 6, 7, 8, 9}
    - No Click = Position == 'No Click'
    - Top 3 Clicks % = Top 3 Clicks / Grand Total
    - Top 10 Clicks % = Top 10 Clicks / Grand Total
    - No Click % = No Click / Grand Total
    - Achievement Status = 'Yes' if Top 3 Clicks % >= 90% (0.90) else 'No'
    """
    tot = len(sub_df)
    if tot == 0:
        return {
            "Grand Total": 0, "Top 3 Clicks": 0, "Top 10 Clicks": 0, "No Click": 0,
            "Top 3 Clicks %": None, "Top 10 Clicks %": None, "No Click %": None, "Achievement": "-"
        }
    pos_series = sub_df["Position_Clean"]
    top3 = int(pos_series.isin(TOP3_POSITIONS).sum())
    top10 = int(pos_series.isin(TOP10_POSITIONS).sum())
    noclick = int((pos_series == "No Click").sum())

    top3_pct = round(top3 / tot * 100, 2)
    top10_pct = round(top10 / tot * 100, 2)
    noclick_pct = round(noclick / tot * 100, 2)
    ach = "Yes" if (top3 / tot) >= 0.90 else "No"
    return {
        "Grand Total": tot,
        "Top 3 Clicks": top3,
        "Top 10 Clicks": top10,
        "No Click": noclick,
        "Top 3 Clicks %": top3_pct,
        "Top 10 Clicks %": top10_pct,
        "No Click %": noclick_pct,
        "Achievement": ach
    }


def build_languages_data_pivot_sheet(df):
    """
    Reconstructs 'Languages Data' hierarchical sheet (ar-eg by LOB, en-us by LOB, plus Grand Total)
    with individual Position columns (0, 1, 10, 2, ..., No Click) + Grand Total + Top 3/10 Clicks & %.
    """
    if len(df) == 0:
        return pd.DataFrame()

    pos_Order = ["0", "1", "10", "2", "3", "4", "5", "6", "7", "8", "9", "No Click"]
    present_pos = [p for p in pos_Order if p in df["Position_Clean"].unique()]

    rows = []
    for lang_val in ["ar-eg", "en-us"]:
        sub_lang = df[df["Language"] == lang_val]
        if len(sub_lang) == 0:
            continue
        # Language header summary row
        m_lang = compute_pos_metrics(sub_lang)
        lang_row = {"Row Labels": lang_val}
        for p in present_pos:
            lang_row[p] = int((sub_lang["Position_Clean"] == p).sum())
        lang_row["Grand Total"] = m_lang["Grand Total"]
        lang_row["Top 3 Clicks"] = m_lang["Top 3 Clicks"]
        lang_row["Top 10 Clicks"] = m_lang["Top 10 Clicks"]
        lang_row["TOP 3 Clicks %"] = f"{m_lang['Top 3 Clicks %']:.2f}%" if m_lang["Top 3 Clicks %"] is not None else "0.00%"
        lang_row["TOP 10 Clicks %"] = f"{m_lang['Top 10 Clicks %']:.2f}%" if m_lang["Top 10 Clicks %"] is not None else "0.00%"
        lang_row["No Click %"] = f"{m_lang['No Click %']:.2f}%" if m_lang["No Click %"] is not None else "0.00%"
        rows.append(lang_row)

        # LOB breakdown under each language
        for lob_val, grp_df in sub_lang.groupby("LOB"):
            r_dict = {"Row Labels": lob_val}
            for p in present_pos:
                r_dict[p] = int((grp_df["Position_Clean"] == p).sum())
            m = compute_pos_metrics(grp_df)
            r_dict["Grand Total"] = m["Grand Total"]
            r_dict["Top 3 Clicks"] = m["Top 3 Clicks"]
            r_dict["Top 10 Clicks"] = m["Top 10 Clicks"]
            r_dict["TOP 3 Clicks %"] = f"{m['Top 3 Clicks %']:.2f}%" if m["Top 3 Clicks %"] is not None else "0.00%"
            r_dict["TOP 10 Clicks %"] = f"{m['Top 10 Clicks %']:.2f}%" if m["Top 10 Clicks %"] is not None else "0.00%"
            r_dict["No Click %"] = f"{m['No Click %']:.2f}%" if m["No Click %"] is not None else "0.00%"
            rows.append(r_dict)

    out_df = pd.DataFrame(rows)
    if len(out_df) > 0:
        tot_m = compute_pos_metrics(df)
        gt_row = {"Row Labels": "Grand Total"}
        for p in present_pos:
            gt_row[p] = int((df["Position_Clean"] == p).sum())
        gt_row["Grand Total"] = tot_m["Grand Total"]
        gt_row["Top 3 Clicks"] = tot_m["Top 3 Clicks"]
        gt_row["Top 10 Clicks"] = tot_m["Top 10 Clicks"]
        gt_row["TOP 3 Clicks %"] = f"{tot_m['Top 3 Clicks %']:.2f}%" if tot_m["Top 3 Clicks %"] is not None else "0.00%"
        gt_row["TOP 10 Clicks %"] = f"{tot_m['Top 10 Clicks %']:.2f}%" if tot_m["Top 10 Clicks %"] is not None else "0.00%"
        gt_row["No Click %"] = f"{tot_m['No Click %']:.2f}%" if tot_m["No Click %"] is not None else "0.00%"
        out_df = pd.concat([out_df, pd.DataFrame([gt_row])], ignore_index=True)

    return out_df


def export_multi_sheet_searching_excel(sheets_dict):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    total_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    regular_font = Font(name="Calibri", size=11, color="1E293B")
    bold_font = Font(name="Calibri", size=11, bold=True, color="0F172A")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    for sheet_name, df in sheets_dict.items():
        ws = wb.create_sheet(title=sheet_name[:31])
        cols = list(df.columns)

        for col_idx, col_name in enumerate(cols, start=1):
            cell = ws.cell(1, col_idx, col_name)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border

        for r_idx, row in enumerate(df.itertuples(index=False), start=2):
            is_gt = str(row[0]).strip() in ["Grand Total", "Total", "ar-eg", "en-us"]
            for c_idx, val in enumerate(row, start=1):
                clean_v = "" if pd.isna(val) else val
                cell = ws.cell(r_idx, c_idx, clean_v)
                cell.font = bold_font if (c_idx == 1 or is_gt) else regular_font
                if is_gt:
                    cell.fill = total_fill
                elif r_idx % 2 == 1:
                    cell.fill = zebra_fill
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center" if c_idx > 1 else "left", vertical="center")

        ws.freeze_panes = "A2"
        if len(df) > 0:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(df) + 1}"

        for c_idx, col_name in enumerate(cols, start=1):
            col_letter = get_column_letter(c_idx)
            ws.column_dimensions[col_letter].width = 28 if col_name in ["Name", "User Name", "Row Labels", "Title", "Query"] else 18

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def render_report(agents_df, hc_df):
    st.title("🔍 Searching Accuracy Report & Executive PDF Generator")
    st.markdown(
        "**Automated End-to-End Workflow for `Searching accuracy report.xlsx`:**  \n"
        "1. Upload the raw **`Total RawData`** export (without `User Name` & `LOB`) or the full **`Searching accuracy report.xlsx`** workbook.  \n"
        "2. The tool automatically matches **`User W3 Account`** against **`Master_Roster_DB.xlsx`**, populates **`User Name`** and **`LOB`**, and organizes every remaining report section **paired side-by-side with its dedicated Chart** both in the interactive view and in the **Executive PDF Report** (ready to send via email to all teams)."
    )

    col_up1, col_up2 = st.columns(2)
    with col_up1:
        search_file = st.file_uploader(
            "📂 1. Upload 'Total RawData' Export or 'Searching accuracy report.xlsx':",
            type=["xlsx", "xls", "csv"],
            key="search_raw_uploader"
        )
    with col_up2:
        hist_trend_file = st.file_uploader(
            "📈 2. (Optional) Upload Previous 'Searching accuracy report.xlsx' to Append Historical Weekly Trends:",
            type=["xlsx", "xls"],
            key="search_hist_trend_uploader"
        )

    if search_file is not None:
        prev_trend_team_df = None
        prev_trend_lang_df = None

        if search_file.name.endswith(".csv"):
            raw_df = pd.read_csv(search_file)
        else:
            xls = pd.ExcelFile(search_file)
            if "Trend Per Team" in xls.sheet_names:
                prev_trend_team_df = pd.read_excel(xls, sheet_name="Trend Per Team")
            if "Trend Per Language Searching" in xls.sheet_names:
                prev_trend_lang_df = pd.read_excel(xls, sheet_name="Trend Per Language Searching")

            target_raw_sheet = "Total RawData" if "Total RawData" in xls.sheet_names else xls.sheet_names[-1]
            for s in xls.sheet_names:
                if "raw" in s.lower() or "total" in s.lower():
                    target_raw_sheet = s
                    break
            raw_df = pd.read_excel(xls, sheet_name=target_raw_sheet)

        if hist_trend_file is not None:
            xls_ht = pd.ExcelFile(hist_trend_file)
            if "Trend Per Team" in xls_ht.sheet_names:
                prev_trend_team_df = pd.read_excel(xls_ht, sheet_name="Trend Per Team")
            if "Trend Per Language Searching" in xls_ht.sheet_names:
                prev_trend_lang_df = pd.read_excel(xls_ht, sheet_name="Trend Per Language Searching")

        raw_df.columns = [str(c).strip() for c in raw_df.columns]

        # Identify User W3 Account column
        w3_col = None
        for cand in ["User W3 Account", "Consignee", "User Account", "Account", "User ID"]:
            if cand in raw_df.columns:
                w3_col = cand
                break
        if w3_col is None:
            st.error(f"❌ Could not find 'User W3 Account' column in uploaded sheet! Found: {list(raw_df.columns)}")
            st.stop()

        raw_df[w3_col] = raw_df[w3_col].astype(str).str.strip()

        # Map User W3 Account -> User Name & LOB from Master Roster DB
        roster_map_name = dict(zip(agents_df["Consignee"].str.upper(), agents_df["Submitter Name"]))
        roster_map_lob = dict(zip(agents_df["Consignee"].str.upper(), agents_df["Submitter LOB"]))

        mapped_names = raw_df[w3_col].str.upper().map(roster_map_name)
        mapped_lobs = raw_df[w3_col].str.upper().map(roster_map_lob)

        if "User Name" in raw_df.columns:
            raw_df["User Name"] = mapped_names.fillna(raw_df["User Name"].astype(str).str.strip())
        else:
            idx_w3 = raw_df.columns.get_loc(w3_col)
            raw_df.insert(idx_w3 + 1, "User Name", mapped_names.fillna("Unknown User"))

        if "LOB" in raw_df.columns:
            raw_df["LOB"] = mapped_lobs.fillna(raw_df["LOB"].astype(str).str.strip())
        else:
            idx_un = raw_df.columns.get_loc("User Name")
            raw_df.insert(idx_un + 1, "LOB", mapped_lobs.fillna("Unassigned LOB"))

        raw_df["User Name"] = raw_df["User Name"].astype(str).str.strip()
        raw_df["LOB"] = raw_df["LOB"].astype(str).str.strip()

        # Alert if any User W3 Account is unmapped
        unmapped_mask = (
            raw_df["User Name"].isna()
            | (raw_df["User Name"] == "Unknown User")
            | (raw_df["User Name"] == "nan")
            | (raw_df["User Name"] == "")
        )
        if unmapped_mask.any():
            unmapped_ids = raw_df.loc[unmapped_mask, w3_col].unique().tolist()
            with st.expander(f"⚠️ {len(unmapped_ids)} New User W3 Accounts Found! Assign their User Name & LOB once:", expanded=True):
                new_r_df = pd.DataFrame({
                    "Consignee": unmapped_ids,
                    "Submitter Name": ["" for _ in unmapped_ids],
                    "Submitter LOB": [hc_df["LOB"].iloc[0] if len(hc_df) > 0 else "Egypt Hotline" for _ in unmapped_ids]
                })
                edited_new_r = st.data_editor(
                    new_r_df,
                    column_config={
                        "Submitter LOB": st.column_config.SelectboxColumn("LOB", options=sorted(hc_df["LOB"].unique().tolist()))
                    },
                    use_container_width=True,
                    key="w3_quick_map"
                )
                if st.button("✅ Save New W3 Accounts to Master Roster & Re-Map"):
                    valid_rows = edited_new_r[edited_new_r["Submitter Name"].astype(str).str.strip() != ""]
                    if len(valid_rows) > 0:
                        valid_rows["Consignee"] = valid_rows["Consignee"].astype(str).str.strip().str.upper()
                        comb = pd.concat([agents_df, valid_rows], ignore_index=True).drop_duplicates(subset=["Consignee"], keep="last")
                        save_master_roster(comb, hc_df)
                        st.rerun()

        # Clean Language & Position columns
        raw_df["Language"] = raw_df["Language"].fillna("en-us").astype(str).str.strip().str.lower()
        raw_df["Position_Clean"] = (
            raw_df["Position"]
            .fillna("No Click")
            .astype(str)
            .str.strip()
            .apply(lambda x: x[:-2] if x.endswith(".0") else x)
        )

        # ======================================================================
        # BUILD ALL REMAINING REPORT TABLES (Removed Arabic & English Agents Data)
        # ======================================================================
        teams_list = sorted([t for t in raw_df["LOB"].unique() if t and t != "Unassigned LOB"])

        # 1. Per Team Table
        per_team_rows = []
        for lob in teams_list:
            sub_t = raw_df[raw_df["LOB"] == lob]
            ar_m = compute_pos_metrics(sub_t[sub_t["Language"] == "ar-eg"])
            en_m = compute_pos_metrics(sub_t[sub_t["Language"] == "en-us"])

            ar_t3_pct = ar_m["Top 3 Clicks %"]
            en_t3_pct = en_m["Top 3 Clicks %"]
            valid_pcts = [p for p in [ar_t3_pct, en_t3_pct] if p is not None]
            avg_t3_pct = round(sum(valid_pcts) / len(valid_pcts), 2) if valid_pcts else 0.0

            per_team_rows.append({
                "Row Labels": lob,
                "Arabic Grand Total": ar_m["Grand Total"] if ar_m["Grand Total"] > 0 else None,
                "Arabic Top 3 Click": ar_m["Top 3 Clicks"] if ar_m["Grand Total"] > 0 else None,
                "Arabic No Click": ar_m["No Click"] if ar_m["Grand Total"] > 0 else None,
                "Arabic Top 3 Clicks %": ar_t3_pct,
                "Arabic No Click %": ar_m["No Click %"],
                "Arabic Achievement Status": ar_m["Achievement"] if ar_m["Grand Total"] > 0 else None,
                "English Grand Total": en_m["Grand Total"] if en_m["Grand Total"] > 0 else None,
                "English Top 3 Click": en_m["Top 3 Clicks"] if en_m["Grand Total"] > 0 else None,
                "English No Click": en_m["No Click"] if en_m["Grand Total"] > 0 else None,
                "English Top 3 Clicks %": en_t3_pct,
                "English No Click %": en_m["No Click %"],
                "English Achievement Status": en_m["Achievement"] if en_m["Grand Total"] > 0 else None,
                "Average Top 3 Clicks %": avg_t3_pct
            })
        per_team_df = pd.DataFrame(per_team_rows).sort_values(by="English Top 3 Clicks %", ascending=False)

        # 2. Per Agent Table
        per_agent_rows = []
        for (w3_acc, u_name, u_lob), grp in raw_df.groupby([w3_col, "User Name", "LOB"]):
            ar_m = compute_pos_metrics(grp[grp["Language"] == "ar-eg"])
            en_m = compute_pos_metrics(grp[grp["Language"] == "en-us"])
            per_agent_rows.append({
                "Row Labels": w3_acc.lower(),
                "Name": u_name,
                "LOB": u_lob,
                "Arabic Grand Total": ar_m["Grand Total"] if ar_m["Grand Total"] > 0 else None,
                "Arabic Top 3 Clicks": ar_m["Top 3 Clicks"] if ar_m["Grand Total"] > 0 else None,
                "Arabic Top 10 Clicks": ar_m["Top 10 Clicks"] if ar_m["Grand Total"] > 0 else None,
                "Arabic No Click": ar_m["No Click"] if ar_m["Grand Total"] > 0 else None,
                "Arabic Top 3 Clicks %": ar_m["Top 3 Clicks %"],
                "Arabic TOP 10 Clicks %": ar_m["Top 10 Clicks %"],
                "Arabic No Click %": ar_m["No Click %"],
                "Arabic TOP 3 Clicks Achievement Status": ar_m["Achievement"] if ar_m["Grand Total"] > 0 else None,
                "English Grand Total": en_m["Grand Total"] if en_m["Grand Total"] > 0 else None,
                "English Top 3 Clicks": en_m["Top 3 Clicks"] if en_m["Grand Total"] > 0 else None,
                "English Top 10 Clicks": en_m["Top 10 Clicks"] if en_m["Grand Total"] > 0 else None,
                "English No Click": en_m["No Click"] if en_m["Grand Total"] > 0 else None,
                "English Top 3 Clicks %": en_m["Top 3 Clicks %"],
                "English TOP 10 Clicks %": en_m["Top 10 Clicks %"],
                "English No Click %": en_m["No Click %"],
                "English TOP 3 Clicks Achievement Status": en_m["Achievement"] if en_m["Grand Total"] > 0 else None,
            })
        per_agent_df = pd.DataFrame(per_agent_rows).sort_values(by=["English Top 3 Clicks %", "English Grand Total"], ascending=[False, False])

        # 3. Per Language Table
        per_lang_rows = []
        for lang_code in ["ar-eg", "en-us"]:
            sub_l = raw_df[raw_df["Language"] == lang_code]
            for lob, grp_l in sub_l.groupby("LOB"):
                m = compute_pos_metrics(grp_l)
                per_lang_rows.append({
                    "Language Group": lang_code,
                    "Row Labels": lob,
                    "Grand Total": m["Grand Total"],
                    "Top 3 Clicks": m["Top 3 Clicks"],
                    "No Click": m["No Click"],
                    "TOP 3 Clicks %": m["Top 3 Clicks %"],
                    "No Click %": m["No Click %"],
                    "Achievement": m["Achievement"]
                })
            tot_m = compute_pos_metrics(sub_l)
            per_lang_rows.append({
                "Language Group": lang_code,
                "Row Labels": lang_code,
                "Grand Total": tot_m["Grand Total"],
                "Top 3 Clicks": tot_m["Top 3 Clicks"],
                "No Click": tot_m["No Click"],
                "TOP 3 Clicks %": tot_m["Top 3 Clicks %"],
                "No Click %": tot_m["No Click %"],
                "Achievement": tot_m["Achievement"]
            })
        per_language_df = pd.DataFrame(per_lang_rows)

        # 4. Language Searching Rate Table
        en_total_cnt = int((raw_df["Language"] == "en-us").sum())
        ar_total_cnt = int((raw_df["Language"] == "ar-eg").sum())
        all_total_cnt = len(raw_df)
        lang_rate_df = pd.DataFrame([
            {"Language Searching Rate": "English", "Count of Language": en_total_cnt, "%": round(en_total_cnt / all_total_cnt * 100, 2) if all_total_cnt else 0.0},
            {"Language Searching Rate": "Arabic", "Count of Language": ar_total_cnt, "%": round(ar_total_cnt / all_total_cnt * 100, 2) if all_total_cnt else 0.0},
            {"Language Searching Rate": "Total", "Count of Language": all_total_cnt, "%": 100.0}
        ])

        # 5. Trend Per Team & 6. Trend Per Language Searching
        raw_df["Week_Bucket"] = raw_df["Date"].apply(assign_week_bucket)
        week_order = raw_df.sort_values("Date")["Week_Bucket"].unique().tolist()

        if prev_trend_team_df is not None and len(prev_trend_team_df) > 0:
            trend_team_df = prev_trend_team_df.copy()
        else:
            t_rows = []
            for lob in teams_list:
                r_d = {"Team": lob}
                for wk in week_order:
                    sub_wk = raw_df[(raw_df["LOB"] == lob) & (raw_df["Week_Bucket"] == wk)]
                    m_wk = compute_pos_metrics(sub_wk)
                    r_d[wk] = round(m_wk["Top 3 Clicks %"] / 100.0, 4) if m_wk["Top 3 Clicks %"] is not None else None
                t_rows.append(r_d)
            trend_team_df = pd.DataFrame(t_rows)

        if prev_trend_lang_df is not None and len(prev_trend_lang_df) > 0:
            trend_lang_df = prev_trend_lang_df.copy()
        else:
            l_rows = [{"Row Labels": "English"}, {"Row Labels": "Arabic"}]
            for wk in week_order:
                sub_wk = raw_df[raw_df["Week_Bucket"] == wk]
                tot_w = len(sub_wk)
                en_w = int((sub_wk["Language"] == "en-us").sum())
                ar_w = int((sub_wk["Language"] == "ar-eg").sum())
                l_rows[0][wk] = round(en_w / tot_w, 4) if tot_w else 0.0
                l_rows[1][wk] = round(ar_w / tot_w, 4) if tot_w else 0.0
            trend_lang_df = pd.DataFrame(l_rows)

        # 7. Languages Data (Position 0..10 & No Click Matrix by Language & LOB)
        lang_lob_pivot = build_languages_data_pivot_sheet(raw_df)
        clean_raw_out = raw_df.drop(columns=["Position_Clean", "Week_Bucket"], errors="ignore")

        # ======================================================================
        # TOP EXECUTIVE KPIs
        # ======================================================================
        overall_m = compute_pos_metrics(raw_df)
        en_overall_m = compute_pos_metrics(raw_df[raw_df["Language"] == "en-us"])
        ar_overall_m = compute_pos_metrics(raw_df[raw_df["Language"] == "ar-eg"])

        k1, k2, k3, k4, k5, k6 = st.columns(6)
        k1.metric("Total Searches", f"{all_total_cnt:,}")
        k2.metric("Overall Top 3 Clicks %", f"{overall_m['Top 3 Clicks %']:.2f}%", f"Target >= 90% ({overall_m['Achievement']})")
        k3.metric("Overall Top 10 Clicks %", f"{overall_m['Top 10 Clicks %']:.2f}%")
        k4.metric("Overall No Click %", f"{overall_m['No Click %']:.2f}%", f"{overall_m['No Click']:,} No Clicks")
        k5.metric("English Top 3 % (en-us)", f"{en_overall_m['Top 3 Clicks %']:.2f}%" if en_overall_m["Top 3 Clicks %"] is not None else "-", f"{en_total_cnt:,} searches")
        k6.metric("Arabic Top 3 % (ar-eg)", f"{ar_overall_m['Top 3 Clicks %']:.2f}%" if ar_overall_m["Top 3 Clicks %"] is not None else "-", f"{ar_total_cnt:,} searches")

        st.divider()

        # ======================================================================
        # EACH REPORT SECTION PAIRED WITH ITS DEDICATED CHART (IN-APP & IN PDF)
        # ======================================================================
        tabs = st.tabs([
            "🏆 1. Per Team (+Chart)",
            "🌐 2. Per Language (+Chart)",
            "📊 3. Language Searching Rate (+Chart)",
            "📈 4. Trend Per Team (+Chart)",
            "📉 5. Trend Per Language (+Chart)",
            "🔢 6. Languages Data Matrix (+Chart)",
            "👤 7. Per Agent (+Chart)",
            "🗂️ 8. Total RawData"
        ])

        # --- Section 1: Per Team ---
        with tabs[0]:
            st.subheader("1. Searching Accuracy Per Team (Chart + Table)")
            chart_team_df = per_team_df.sort_values(by="Average Top 3 Clicks %", ascending=False).fillna(0)
            fig_team = px.bar(
                chart_team_df,
                x="Row Labels",
                y=["Arabic Top 3 Clicks %", "Arabic No Click %", "English Top 3 Clicks %", "English No Click %"],
                barmode="group",
                title="Searching Accuracy Per Team (Arabic & English Top 3 Clicks % vs. No Click %)",
                color_discrete_sequence=["#10b981", "#f59e0b", "#2563eb", "#ef4444"],
                text_auto=".1f"
            )
            fig_team.add_hline(y=90, line_dash="dash", line_color="#0f172a", annotation_text="90% Target")
            st.plotly_chart(fig_team, use_container_width=True)
            st.dataframe(per_team_df, use_container_width=True)

        # --- Section 2: Per Language ---
        with tabs[1]:
            st.subheader("2. Per Language Breakdown by Team (Chart + Table)")
            pl_chart_df = per_language_df[~per_language_df["Row Labels"].isin(["ar-eg", "en-us"])].copy()
            fig_pl = px.bar(
                pl_chart_df,
                x="Row Labels",
                y=["TOP 3 Clicks %", "No Click %"],
                facet_col="Language Group",
                barmode="group",
                title="Per Language Searching Accuracy (ar-eg vs. en-us) Across Teams",
                color_discrete_sequence=["#10b981", "#ef4444"],
                text_auto=".1f"
            )
            fig_pl.add_hline(y=90, line_dash="dash", line_color="#0f172a")
            st.plotly_chart(fig_pl, use_container_width=True)
            st.dataframe(per_language_df, use_container_width=True)

        # --- Section 3: Language Searching Rate ---
        with tabs[2]:
            st.subheader("3. Language Searching Rate (Chart + Table)")
            c_lr1, c_lr2 = st.columns([1.3, 1])
            with c_lr1:
                fig_lang_rate = px.pie(
                    lang_rate_df[lang_rate_df["Language Searching Rate"] != "Total"],
                    names="Language Searching Rate",
                    values="Count of Language",
                    hole=0.38,
                    title="Language Searching Share (English vs. Arabic)",
                    color="Language Searching Rate",
                    color_discrete_map={"English": "#2563eb", "Arabic": "#10b981"}
                )
                fig_lang_rate.update_traces(textinfo="label+percent+value")
                st.plotly_chart(fig_lang_rate, use_container_width=True)
            with c_lr2:
                st.markdown("##### Language Searching Rate Table")
                st.dataframe(lang_rate_df, use_container_width=True)

        # --- Section 4: Trend Per Team ---
        with tabs[3]:
            st.subheader("4. Weekly Searching Accuracy Trend Per Team (Chart + Table)")
            wk_cols_t = [c for c in trend_team_df.columns if c not in ["Team", "Trend"]]
            if len(wk_cols_t) > 0:
                melted_t = trend_team_df.melt(id_vars=["Team"], value_vars=wk_cols_t, var_name="Period", value_name="Top 3 Accuracy")
                melted_t["Top 3 Accuracy (%)"] = pd.to_numeric(melted_t["Top 3 Accuracy"], errors="coerce") * 100.0
                fig_trend_t = px.line(
                    melted_t,
                    x="Period",
                    y="Top 3 Accuracy (%)",
                    color="Team",
                    markers=True,
                    title="Weekly Top 3 Clicks Accuracy Trend Per Team (%)"
                )
                fig_trend_t.add_hline(y=90, line_dash="dash", line_color="#0f172a", annotation_text="90% Target")
                st.plotly_chart(fig_trend_t, use_container_width=True)
            st.dataframe(trend_team_df, use_container_width=True)

        # --- Section 5: Trend Per Language Searching ---
        with tabs[4]:
            st.subheader("5. Weekly Trend Per Language Searching (Chart + Table)")
            wk_cols_l = [c for c in trend_lang_df.columns if c != "Row Labels"]
            if len(wk_cols_l) > 0:
                melted_l = trend_lang_df.melt(id_vars=["Row Labels"], value_vars=wk_cols_l, var_name="Period", value_name="Share")
                melted_l["Searching Share (%)"] = pd.to_numeric(melted_l["Share"], errors="coerce") * 100.0
                fig_trend_l = px.bar(
                    melted_l,
                    x="Period",
                    y="Searching Share (%)",
                    color="Row Labels",
                    barmode="group",
                    text_auto=".2f",
                    title="Weekly Language Searching Share Trend (English vs. Arabic %)",
                    color_discrete_map={"English": "#2563eb", "Arabic": "#10b981"}
                )
                st.plotly_chart(fig_trend_l, use_container_width=True)
            st.dataframe(trend_lang_df, use_container_width=True)

        # --- Section 6: Languages Data Matrix ---
        with tabs[5]:
            st.subheader("6. Languages Data Click Position Matrix (Chart + Table)")
            pos_order_chart = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "No Click"]
            pos_dist = raw_df.groupby(["LOB", "Position_Clean"]).size().reset_index(name="Count")
            fig_pos = px.bar(
                pos_dist,
                x="LOB",
                y="Count",
                color="Position_Clean",
                category_orders={"Position_Clean": pos_order_chart},
                title="Click Position Distribution (0, 1, 2 ... No Click) Across Teams",
                color_discrete_sequence=px.colors.qualitative.Prism
            )
            st.plotly_chart(fig_pos, use_container_width=True)
            st.dataframe(lang_lob_pivot, use_container_width=True)

        # --- Section 7: Per Agent ---
        with tabs[6]:
            st.subheader("7. Per Agent Searching Accuracy (Chart + Table)")
            ag_chart_df = per_agent_df.dropna(subset=["English Top 3 Clicks %"]).copy()
            fig_ag = px.bar(
                ag_chart_df.head(35),
                x="Name",
                y=["English Top 3 Clicks %", "English TOP 10 Clicks %", "English No Click %"],
                barmode="group",
                title="Agent Searching Accuracy Comparison (Top 35 Agents by Top 3 Clicks %)",
                color_discrete_sequence=["#10b981", "#2563eb", "#ef4444"]
            )
            fig_ag.add_hline(y=90, line_dash="dash", line_color="#0f172a", annotation_text="90% Target")
            fig_ag.update_layout(xaxis_tickangle=-35)
            st.plotly_chart(fig_ag, use_container_width=True)
            st.dataframe(per_agent_df, use_container_width=True)

        # --- Section 8: Total RawData ---
        with tabs[7]:
            st.subheader("8. Enriched Total RawData (With User Name & LOB Added)")
            st.dataframe(clean_raw_out.head(500), use_container_width=True)
            st.caption(f"Showing first 500 rows of {len(clean_raw_out):,} total rows. All rows are included in the Excel export below.")

        st.divider()

        # ======================================================================
        # PDF EXECUTIVE REPORT GENERATOR (EVERY SECTION PAIRED WITH ITS CHART!)
        # ======================================================================
        def fig_to_rl_image(fig, width_pt=750, height_pt=210):
            img_buf = io.BytesIO()
            fig.savefig(img_buf, format="png", dpi=180, bbox_inches="tight")
            plt.close(fig)
            img_buf.seek(0)
            return Image(img_buf, width=width_pt, height=height_pt)

        def style_pdf_table(tbl, header_bg="#1E293B", font_size=7.5):
            tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_bg)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("FONTSIZE", (0, 0), (-1, -1), font_size),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]))

        def generate_searching_pdf_bytes():
            buf = io.BytesIO()
            doc = SimpleDocTemplate(
                buf,
                pagesize=landscape(A4),
                rightMargin=20,
                leftMargin=20,
                topMargin=20,
                bottomMargin=20
            )
            styles = getSampleStyleSheet()
            story = []

            title_style = ParagraphStyle("STitle", parent=styles["Heading1"], fontSize=18, textColor=colors.HexColor("#0F172A"), spaceAfter=4)
            sec_style = ParagraphStyle("SSec", parent=styles["Heading2"], fontSize=13, textColor=colors.HexColor("#1E3A8A"), spaceBefore=4, spaceAfter=6)
            sub_style = ParagraphStyle("SSub", parent=styles["Normal"], fontSize=9.5, textColor=colors.HexColor("#334155"), spaceAfter=10)

            # --- PAGE 1: SECTION 1 - PER TEAM (CHART + TABLE) ---
            story.append(Paragraph("Executive Searching Accuracy Report (All Teams Overview)", title_style))
            story.append(Paragraph(
                f"<b>Total Searches:</b> {all_total_cnt:,} &nbsp;|&nbsp; "
                f"<b>Overall Top 3 Accuracy:</b> {overall_m['Top 3 Clicks %']:.2f}% ({overall_m['Top 3 Clicks']:,}) &nbsp;|&nbsp; "
                f"<b>Overall Top 10 Accuracy:</b> {overall_m['Top 10 Clicks %']:.2f}% &nbsp;|&nbsp; "
                f"<b>No Click Rate:</b> {overall_m['No Click %']:.2f}% ({overall_m['No Click']:,}) &nbsp;|&nbsp; "
                f"<b>English Share:</b> {lang_rate_df.iloc[0]['%']:.2f}% ({en_total_cnt:,}) &nbsp;|&nbsp; "
                f"<b>Arabic Share:</b> {lang_rate_df.iloc[1]['%']:.2f}% ({ar_total_cnt:,})",
                sub_style
            ))

            story.append(Paragraph("Section 1: Searching Accuracy Per Team (Chart & Table)", sec_style))

            # Chart 1: Per Team
            fig1, ax1 = plt.subplots(figsize=(11.0, 3.1))
            ct = per_team_df.sort_values(by="Average Top 3 Clicks %", ascending=False).fillna(0)
            x = range(len(ct))
            w = 0.20
            b_ar3 = ax1.bar([i - 1.5*w for i in x], ct["Arabic Top 3 Clicks %"], width=w, label="Arabic Top 3 Clicks %", color="#10b981")
            b_arn = ax1.bar([i - 0.5*w for i in x], ct["Arabic No Click %"], width=w, label="Arabic No Click %", color="#f59e0b")
            b_en3 = ax1.bar([i + 0.5*w for i in x], ct["English Top 3 Clicks %"], width=w, label="English Top 3 Clicks %", color="#2563eb")
            b_enn = ax1.bar([i + 1.5*w for i in x], ct["English No Click %"], width=w, label="English No Click %", color="#ef4444")
            ax1.axhline(90, color="#0f172a", linestyle="--", linewidth=1.0, label="90% Target")
            ax1.set_xticks(list(x))
            ax1.set_xticklabels(ct["Row Labels"], rotation=15, ha="right", fontsize=8.5, fontweight="bold")
            ax1.set_ylim(0, 115)
            ax1.set_ylabel("Percentage (%)")
            ax1.set_title("Searching Accuracy Per Team (Arabic & English Top 3 Clicks % vs. No Click %)", fontsize=10.5, fontweight="bold")
            ax1.legend(fontsize=8, loc="upper right", ncol=5)
            ax1.grid(axis="y", linestyle="--", alpha=0.4)
            for bars in [b_ar3, b_en3]:
                for bar in bars:
                    h = bar.get_height()
                    if h > 0:
                        ax1.annotate(f"{h:.1f}%", xy=(bar.get_x() + bar.get_width()/2, h), xytext=(0, 2), textcoords="offset points", ha="center", va="bottom", fontsize=6.5, fontweight="bold")
            story.append(fig_to_rl_image(fig1, width_pt=770, height_pt=215))
            story.append(Spacer(1, 8))

            # Table 1: Per Team
            t_team_data = [[
                "Team (LOB)", "Ar Total", "Ar Top 3", "Ar NoClick", "Ar Top 3 %", "Ar NoClick %", "Ar Ach.",
                "En Total", "En Top 3", "En NoClick", "En Top 3 %", "En NoClick %", "En Ach.", "Average Top 3 %"
            ]]
            for _, r in per_team_df.iterrows():
                t_team_data.append([
                    str(r["Row Labels"]),
                    str(int(r["Arabic Grand Total"])) if pd.notna(r["Arabic Grand Total"]) else "-",
                    str(int(r["Arabic Top 3 Click"])) if pd.notna(r["Arabic Top 3 Click"]) else "-",
                    str(int(r["Arabic No Click"])) if pd.notna(r["Arabic No Click"]) else "-",
                    f"{r['Arabic Top 3 Clicks %']:.2f}%" if pd.notna(r["Arabic Top 3 Clicks %"]) else "-",
                    f"{r['Arabic No Click %']:.2f}%" if pd.notna(r["Arabic No Click %"]) else "-",
                    str(r["Arabic Achievement Status"]) if pd.notna(r["Arabic Achievement Status"]) else "-",
                    str(int(r["English Grand Total"])) if pd.notna(r["English Grand Total"]) else "-",
                    str(int(r["English Top 3 Click"])) if pd.notna(r["English Top 3 Click"]) else "-",
                    str(int(r["English No Click"])) if pd.notna(r["English No Click"]) else "-",
                    f"{r['English Top 3 Clicks %']:.2f}%" if pd.notna(r["English Top 3 Clicks %"]) else "-",
                    f"{r['English No Click %']:.2f}%" if pd.notna(r["English No Click %"]) else "-",
                    str(r["English Achievement Status"]) if pd.notna(r["English Achievement Status"]) else "-",
                    f"{r['Average Top 3 Clicks %']:.2f}%"
                ])
            tbl_team = Table(t_team_data, colWidths=[110, 42, 42, 46, 52, 54, 38, 48, 48, 48, 52, 54, 38, 65], repeatRows=1)
            style_pdf_table(tbl_team, font_size=7.5)
            story.append(tbl_team)
            story.append(PageBreak())

            # --- PAGE 2: SECTION 2 (PER LANGUAGE) & SECTION 3 (LANGUAGE SEARCHING RATE) ---
            story.append(Paragraph("Section 2 & 3: Per Language Breakdown & Language Searching Rate (Charts & Tables)", sec_style))

            fig2, (ax2_l, ax2_r) = plt.subplots(1, 2, figsize=(11.0, 2.9), gridspec_kw={"width_ratios": [2.2, 1]})
            en_sub_pl = per_language_df[(per_language_df["Language Group"] == "en-us") & (per_language_df["Row Labels"] != "en-us")].sort_values("TOP 3 Clicks %", ascending=False)
            x_pl = range(len(en_sub_pl))
            ax2_l.bar([i - 0.18 for i in x_pl], en_sub_pl["TOP 3 Clicks %"].fillna(0), width=0.35, label="English Top 3 %", color="#2563eb")
            ax2_l.bar([i + 0.18 for i in x_pl], en_sub_pl["No Click %"].fillna(0), width=0.35, label="English No Click %", color="#ef4444")
            ax2_l.axhline(90, color="#0f172a", linestyle="--", linewidth=0.9, label="90% Target")
            ax2_l.set_xticks(list(x_pl))
            ax2_l.set_xticklabels(en_sub_pl["Row Labels"], rotation=20, ha="right", fontsize=8)
            ax2_l.set_ylim(0, 115)
            ax2_l.set_title("Per Language (en-us) Top 3 % vs. No Click % by Team", fontsize=9.5, fontweight="bold")
            ax2_l.legend(fontsize=7.5, loc="upper right")
            ax2_l.grid(axis="y", linestyle="--", alpha=0.4)

            ax2_r.pie(
                [en_total_cnt, max(0, ar_total_cnt)],
                labels=[f"English ({lang_rate_df.iloc[0]['%']:.2f}%)", f"Arabic ({lang_rate_df.iloc[1]['%']:.2f}%)"],
                autopct="%1.2f%%",
                startangle=140,
                colors=["#2563eb", "#10b981"],
                textprops={"fontsize": 8, "fontweight": "bold"}
            )
            ax2_r.set_title("Language Searching Rate Share", fontsize=9.5, fontweight="bold")
            story.append(fig_to_rl_image(fig2, width_pt=770, height_pt=195))
            story.append(Spacer(1, 6))

            t_pl_data = [["Language", "Team (Row Labels)", "Grand Total", "Top 3 Clicks", "No Click", "TOP 3 Clicks %", "No Click %", "Achievement"]]
            for _, r in per_language_df.iterrows():
                t_pl_data.append([
                    str(r["Language Group"]),
                    str(r["Row Labels"]),
                    f"{int(r['Grand Total']):,}",
                    f"{int(r['Top 3 Clicks']):,}",
                    f"{int(r['No Click']):,}",
                    f"{r['TOP 3 Clicks %']:.2f}%" if pd.notna(r["TOP 3 Clicks %"]) else "-",
                    f"{r['No Click %']:.2f}%" if pd.notna(r["No Click %"]) else "-",
                    str(r["Achievement"])
                ])
            tbl_pl = Table(t_pl_data, colWidths=[65, 130, 75, 75, 65, 85, 75, 75], repeatRows=1)
            style_pdf_table(tbl_pl, font_size=7.2)
            story.append(tbl_pl)
            story.append(Spacer(1, 6))

            t_lr_data = [["Language Searching Rate", "Count of Language", "Share (%)"]]
            for _, r in lang_rate_df.iterrows():
                t_lr_data.append([str(r["Language Searching Rate"]), f"{int(r['Count of Language']):,}", f"{r['%']:.2f}%"])
            tbl_lr = Table(t_lr_data, colWidths=[180, 140, 120], repeatRows=1)
            style_pdf_table(tbl_lr, header_bg="#0F766E", font_size=7.5)
            story.append(tbl_lr)
            story.append(PageBreak())

            # --- PAGE 3: SECTION 4 - TREND PER TEAM (CHART + TABLE) ---
            story.append(Paragraph("Section 4: Weekly Searching Accuracy Trend Per Team (Chart & Table)", sec_style))
            wk_cols_t = [c for c in trend_team_df.columns if c not in ["Team", "Trend"]]
            if len(wk_cols_t) > 0:
                fig3, ax3 = plt.subplots(figsize=(11.0, 3.2))
                for _, r in trend_team_df.iterrows():
                    team_name = str(r["Team"])
                    y_vals = [float(r[c])*100.0 if pd.notna(r[c]) and str(r[c]).strip() != "" else None for c in wk_cols_t]
                    ax3.plot(range(len(wk_cols_t)), y_vals, marker="o", linewidth=1.8, markersize=4.5, label=team_name)
                ax3.axhline(90, color="#0f172a", linestyle="--", linewidth=1.0, label="90% Target")
                ax3.set_xticks(range(len(wk_cols_t)))
                ax3.set_xticklabels(wk_cols_t, rotation=20, ha="right", fontsize=8)
                ax3.set_ylim(60, 105)
                ax3.set_ylabel("Top 3 Clicks Accuracy (%)")
                ax3.set_title("Weekly Searching Accuracy Trend Per Team (%)", fontsize=10.5, fontweight="bold")
                ax3.legend(fontsize=7.2, loc="lower left", ncol=4)
                ax3.grid(True, linestyle="--", alpha=0.4)
                story.append(fig_to_rl_image(fig3, width_pt=770, height_pt=225))
                story.append(Spacer(1, 8))

            t_tt_headers = ["Team"] + wk_cols_t
            t_tt_data = [t_tt_headers]
            for _, r in trend_team_df.iterrows():
                row_vals = [str(r["Team"])]
                for c in wk_cols_t:
                    val = r[c]
                    if pd.notna(val) and str(val).strip() != "":
                        num_v = float(val)
                        row_vals.append(f"{num_v*100:.1f}%" if num_v <= 1.0 else f"{num_v:.1f}%")
                    else:
                        row_vals.append("-")
                t_tt_data.append(row_vals)
            col_w_tt = [105] + [int(660 / max(1, len(wk_cols_t)))] * len(wk_cols_t)
            tbl_tt = Table(t_tt_data, colWidths=col_w_tt, repeatRows=1)
            style_pdf_table(tbl_tt, font_size=6.8)
            story.append(tbl_tt)
            story.append(PageBreak())

            # --- PAGE 4: SECTION 5 (TREND PER LANGUAGE) & SECTION 6 (LANGUAGES DATA MATRIX) ---
            story.append(Paragraph("Section 5 & 6: Trend Per Language Searching & Languages Click Position Matrix", sec_style))
            wk_cols_l = [c for c in trend_lang_df.columns if c != "Row Labels"]
            if len(wk_cols_l) > 0:
                fig4, (ax4_l, ax4_r) = plt.subplots(1, 2, figsize=(11.0, 2.7), gridspec_kw={"width_ratios": [1.5, 1.5]})
                for idx_l, r in trend_lang_df.iterrows():
                    lbl = str(r["Row Labels"]).strip()
                    y_vals = [float(r[c])*100.0 if pd.notna(r[c]) and str(r[c]).strip() != "" else 0.0 for c in wk_cols_l]
                    color_line = "#2563eb" if "eng" in lbl.lower() else "#10b981"
                    ax4_l.plot(range(len(wk_cols_l)), y_vals, marker="s", linewidth=2.0, color=color_line, label=lbl)
                ax4_l.set_xticks(range(len(wk_cols_l)))
                ax4_l.set_xticklabels(wk_cols_l, rotation=22, ha="right", fontsize=7.5)
                ax4_l.set_ylabel("Share (%)")
                ax4_l.set_title("Weekly Trend Per Language Searching (%)", fontsize=9.5, fontweight="bold")
                ax4_l.legend(fontsize=8)
                ax4_l.grid(True, linestyle="--", alpha=0.4)

                pos_order_plot = ["0", "1", "2", "3", "4", "5", "No Click"]
                pos_counts = [int((raw_df["Position_Clean"] == p).sum()) for p in pos_order_plot]
                bars_p = ax4_r.bar(pos_order_plot, pos_counts, color=["#10b981", "#34d399", "#6ee7b7", "#60a5fa", "#93c5fd", "#bfdbfe", "#ef4444"])
                ax4_r.set_title("Overall Click Position Distribution (Positions 0-5 vs. No Click)", fontsize=9.5, fontweight="bold")
                ax4_r.set_ylabel("Search Count")
                ax4_r.grid(axis="y", linestyle="--", alpha=0.4)
                for b in bars_p:
                    h = b.get_height()
                    ax4_r.annotate(f"{h:,}", xy=(b.get_x() + b.get_width()/2, h), xytext=(0, 2), textcoords="offset points", ha="center", va="bottom", fontsize=7, fontweight="bold")

                story.append(fig_to_rl_image(fig4, width_pt=770, height_pt=185))
                story.append(Spacer(1, 6))

                t_tl_headers = ["Language"] + wk_cols_l
                t_tl_data = [t_tl_headers]
                for _, r in trend_lang_df.iterrows():
                    rv = [str(r["Row Labels"]).strip()]
                    for c in wk_cols_l:
                        v = r[c]
                        if pd.notna(v) and str(v).strip() != "":
                            nv = float(v)
                            rv.append(f"{nv*100:.2f}%" if nv <= 1.0 else f"{nv:.2f}%")
                        else:
                            rv.append("-")
                    t_tl_data.append(rv)
                col_w_tl = [95] + [int(670 / max(1, len(wk_cols_l)))] * len(wk_cols_l)
                tbl_tl = Table(t_tl_data, colWidths=col_w_tl, repeatRows=1)
                style_pdf_table(tbl_tl, header_bg="#0F766E", font_size=6.8)
                story.append(tbl_tl)
                story.append(Spacer(1, 8))

            if len(lang_lob_pivot) > 0:
                ld_cols = list(lang_lob_pivot.columns)
                t_ld_data = [ld_cols]
                for _, r in lang_lob_pivot.iterrows():
                    t_ld_data.append([str(r[c]) if pd.notna(r[c]) else "0" for c in ld_cols])
                w_first = 115
                w_rest = int(655 / max(1, len(ld_cols) - 1))
                tbl_ld = Table(t_ld_data, colWidths=[w_first] + [w_rest]*(len(ld_cols)-1), repeatRows=1)
                style_pdf_table(tbl_ld, font_size=6.5)
                story.append(tbl_ld)
            story.append(PageBreak())

            # --- PAGE 5: SECTION 7 - PER AGENT (CHART + TABLE) ---
            story.append(Paragraph("Section 7: Searching Accuracy Per Agent (Chart & Complete Table)", sec_style))
            ag_plot_df = per_agent_df.dropna(subset=["English Top 3 Clicks %"]).head(25)
            if len(ag_plot_df) > 0:
                fig5, ax5 = plt.subplots(figsize=(11.0, 2.7))
                x_ag = range(len(ag_plot_df))
                ax5.bar([i - 0.18 for i in x_ag], ag_plot_df["English Top 3 Clicks %"], width=0.35, label="English Top 3 Clicks %", color="#10b981")
                ax5.bar([i + 0.18 for i in x_ag], ag_plot_df["English TOP 10 Clicks %"], width=0.35, label="English Top 10 Clicks %", color="#2563eb")
                ax5.axhline(90, color="#0f172a", linestyle="--", linewidth=0.9, label="90% Target")
                ax5.set_xticks(list(x_ag))
                ax5.set_xticklabels([str(n)[:16] for n in ag_plot_df["Name"]], rotation=28, ha="right", fontsize=7.2)
                ax5.set_ylim(0, 112)
                ax5.set_ylabel("Accuracy (%)")
                ax5.set_title("Top 25 Agents Searching Accuracy (English Top 3 % & Top 10 %)", fontsize=9.5, fontweight="bold")
                ax5.legend(fontsize=7.5, loc="lower right")
                ax5.grid(axis="y", linestyle="--", alpha=0.4)
                story.append(fig_to_rl_image(fig5, width_pt=770, height_pt=180))
                story.append(Spacer(1, 6))

            t_ag_data = [[
                "W3 Account", "Agent Name", "LOB", "Ar Total", "Ar Top 3 %", "Ar Ach.",
                "En Total", "En Top 3", "En Top 10", "En NoClick", "En Top 3 %", "En Top 10 %", "En NoClick %", "En Ach."
            ]]
            for _, r in per_agent_df.iterrows():
                t_ag_data.append([
                    str(r["Row Labels"]),
                    str(r["Name"])[:24],
                    str(r["LOB"]),
                    str(int(r["Arabic Grand Total"])) if pd.notna(r["Arabic Grand Total"]) else "-",
                    f"{r['Arabic Top 3 Clicks %']:.1f}%" if pd.notna(r["Arabic Top 3 Clicks %"]) else "-",
                    str(r["Arabic TOP 3 Clicks Achievement Status"]) if pd.notna(r["Arabic TOP 3 Clicks Achievement Status"]) else "-",
                    str(int(r["English Grand Total"])) if pd.notna(r["English Grand Total"]) else "-",
                    str(int(r["English Top 3 Clicks"])) if pd.notna(r["English Top 3 Clicks"]) else "-",
                    str(int(r["English Top 10 Clicks"])) if pd.notna(r["English Top 10 Clicks"]) else "-",
                    str(int(r["English No Click"])) if pd.notna(r["English No Click"]) else "-",
                    f"{r['English Top 3 Clicks %']:.1f}%" if pd.notna(r["English Top 3 Clicks %"]) else "-",
                    f"{r['English TOP 10 Clicks %']:.1f}%" if pd.notna(r["English TOP 10 Clicks %"]) else "-",
                    f"{r['English No Click %']:.1f}%" if pd.notna(r["English No Click %"]) else "-",
                    str(r["English TOP 3 Clicks Achievement Status"]) if pd.notna(r["English TOP 3 Clicks Achievement Status"]) else "-"
                ])

            tbl_ag = Table(t_ag_data, colWidths=[64, 115, 90, 38, 46, 35, 44, 44, 46, 46, 48, 50, 50, 38], repeatRows=1)
            style_pdf_table(tbl_ag, font_size=6.8)
            story.append(tbl_ag)

            doc.build(story)
            buf.seek(0)
            return buf.getvalue()

        dl_c1, dl_c2 = st.columns(2)
        with dl_c1:
            st.download_button(
                label="📄 Download Complete Searching Accuracy PDF Report (Every Section Paired with Its Chart - Ready for Email)",
                data=generate_searching_pdf_bytes(),
                file_name="Searching_Accuracy_Executive_Report.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
        with dl_c2:
            excel_Bytes = export_multi_sheet_searching_excel({
                "Languages Data": lang_lob_pivot,
                "Per Agent": per_agent_df,
                "Per Team": per_team_df,
                "Per Language": per_language_df,
                "Trend Per Team": trend_team_df,
                "Language Searching Rate": lang_rate_df,
                "Trend Per Language Searching": trend_lang_df,
                "Total RawData": clean_raw_out
            })
            st.download_button(
                label="📥 Download Complete 8-Sheet Searching Accuracy Workbook (.xlsx)",
                data=excel_Bytes,
                file_name="Searching_Accuracy_Report_Automated.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
