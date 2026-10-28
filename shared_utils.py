import os
import io
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROSTER_FILE = "Master_Roster_DB.xlsx"

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
