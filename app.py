import pandas as pd
import streamlit as st
from shared_utils import load_master_roster, save_master_roster
import report_feedback
import report_searching_accuracy

st.set_page_config(
    page_title="Smart Feedback & Searching Accuracy Hub",
    layout="wide",
    page_icon="⚡"
)

agents_df, hc_df = load_master_roster()

# ==============================================================================
# REPORTS REGISTRY (Add any future reports here in 1 line!)
# ==============================================================================
REPORTS_REGISTRY = {
    "⚡ 1. Feedback Analysis Report": report_feedback.render_report,
    "🔍 2. Searching Accuracy Report": report_searching_accuracy.render_report,
    # Future reports can be registered here easily:
    # "📊 3. New Report Name": report_new_module.render_report,
}

# --- Sidebar: Report Selector Dropdown ---
st.sidebar.title("📌 Select Report Module")
selected_report = st.sidebar.selectbox(
    "Choose Report to Run:",
    options=list(REPORTS_REGISTRY.keys())
)
st.sidebar.divider()

# --- Sidebar: Master Roster & Historical Memory ---
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

        # Support both Consignee and User W3 Account column names
        col_map = {}
        for c in new_agents.columns:
            cl = str(c).strip().lower()
            if cl in ["consignee", "user w3 account", "user account", "user id", "account", "staff id"]:
                col_map[c] = "Consignee"
            elif cl in ["submitter name", "user name", "name", "agent name"]:
                col_map[c] = "Submitter Name"
            elif cl in ["submitter lob", "lob", "team", "department"]:
                col_map[c] = "Submitter LOB"
        new_agents = new_agents.rename(columns=col_map)

        if {"Consignee", "Submitter Name", "Submitter LOB"}.issubset(set(new_agents.columns)):
            new_agents["Consignee"] = new_agents["Consignee"].astype(str).str.strip().str.upper()
            new_agents["Submitter Name"] = new_agents["Submitter Name"].astype(str).str.strip()
            new_agents["Submitter LOB"] = new_agents["Submitter LOB"].astype(str).str.strip()
            combined_agents = pd.concat([agents_df, new_agents], ignore_index=True).drop_duplicates(subset=["Consignee"], keep="last")
            save_master_roster(combined_agents, new_hc)
            agents_df, hc_df = load_master_roster()
            st.sidebar.success("✅ Master Roster updated and saved!")
        else:
            st.sidebar.error("Sheet must contain: Consignee (or User W3 Account), Submitter Name, Submitter LOB")
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

# --- Render the Selected Report Module ---
REPORTS_REGISTRY[selected_report](agents_df, hc_df)
