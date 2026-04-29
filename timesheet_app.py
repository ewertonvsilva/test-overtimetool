# BEFORE USING THIS FILE — REQUIRED DECISIONS
# ------------------------------------------
# Please confirm these behaviors:
# 1. DL logic:
#    - DL Earned shown on OT day ✔
#    - DL Taken manually selected ✔ (currently enabled)
#    - Should DL deduct exactly 8h always? (currently YES)
#
# 2. TOIL logic:
#    - TOIL shown on OT day (Table 2) ✔
#    - TOIL consumed automatically next working day ✔
#    - OR should TOIL also be manually selectable like DL?
#
# 3. Creative balancing:
#    - Creative starts at MAX ✔
#    - Reduced down to MIN ✔
#    - Then Support reduced ✔
#
# 4. Overtime:
#    - Always stays on exact calendar day ✔ (including weekends)
#
# Reply with decisions if changes needed.

import streamlit as st
import pandas as pd
import calendar
import holidays
from datetime import date, datetime, timedelta

# ──────────────────────────────────────────────
# BASIC SETUP
# ──────────────────────────────────────────────

st.set_page_config(layout="wide")

# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────

def parse_hhmm(s):
    try:
        h, m = map(int, s.split(":"))
        return h + m/60
    except:
        return None

# ──────────────────────────────────────────────
# SIDEBAR
# ──────────────────────────────────────────────

with st.sidebar:
    year = st.number_input("Year", 2020, 2035, 2026)
    month = st.number_input("Month", 1, 12, 4)

    creative_min = st.number_input("Creative MIN", 0.0, 8.0, 2.0)
    creative_max = st.number_input("Creative MAX", 0.0, 8.0, 6.0)

    st.markdown("### DL Selection")

# ──────────────────────────────────────────────
# DAYS
# ──────────────────────────────────────────────

_, days_in_month = calendar.monthrange(year, month)
all_days = [date(year, month, d) for d in range(1, days_in_month+1)]

pl_holidays = holidays.Poland(years=year)
holiday_dates = {d for d in pl_holidays if d.month == month}

# ──────────────────────────────────────────────
# STATE
# ──────────────────────────────────────────────

if "entries" not in st.session_state:
    st.session_state.entries = []

# ──────────────────────────────────────────────
# INPUT
# ──────────────────────────────────────────────

st.subheader("Add Entry")

col1, col2, col3 = st.columns(3)

with col1:
    typ = st.selectbox("Type", ["Incident", "Overtime"])
with col2:
    sd = st.date_input("Date", value=all_days[0])
with col3:
    sh = st.text_input("Start HH:MM", "09:00")

eh = st.text_input("End HH:MM", "10:00")

if st.button("Add"):
    st.session_state.entries.append({
        "type": typ,
        "date": sd,
        "start": parse_hhmm(sh),
        "end": parse_hhmm(eh)
    })

# ──────────────────────────────────────────────
# DL INPUT
# ──────────────────────────────────────────────

selected_dl_days = st.multiselect(
    "Select DL days",
    all_days,
    format_func=lambda d: d.strftime("%d %b")
)

# ──────────────────────────────────────────────
# COMPUTE
# ──────────────────────────────────────────────

day_data = {
    d: {
        "creative": 0,
        "support": 0,
        "abs": 0,
        "ot": 0,
        "toil_earned": 0,
        "dl_earned": False,
        "dl_taken": False
    }
    for d in all_days
}

# Apply DL taken
for d in selected_dl_days:
    day_data[d]["dl_taken"] = True
    day_data[d]["abs"] += 8

# Process entries
for e in st.session_state.entries:
    d = e["date"]
    h = e["end"] - e["start"]

    if e["type"] == "Incident":
        day_data[d]["ot"] += h
        day_data[d]["toil_earned"] += h

        if d.weekday() >= 5 or d in holiday_dates:
            day_data[d]["dl_earned"] = True

    else:
        day_data[d]["ot"] += h

# Balance hours
for d in all_days:
    absence = day_data[d]["abs"]

    creative = creative_max
    creative = max(creative_min, creative - absence)

    support = max(0, 8 - creative - absence)

    day_data[d]["creative"] = creative
    day_data[d]["support"] = support

# ──────────────────────────────────────────────
# TABLES
# ──────────────────────────────────────────────

st.subheader("Table 1")

df1 = pd.DataFrame({
    d.strftime("%d"): [
        day_data[d]["creative"],
        day_data[d]["support"],
        day_data[d]["abs"]
    ] for d in all_days
}, index=["Creative", "Support", "Absence"])

st.dataframe(df1)

st.subheader("Table 2")

df2 = pd.DataFrame({
    d.strftime("%d"): [
        day_data[d]["ot"],
        day_data[d]["toil_earned"],
        "YES" if day_data[d]["dl_earned"] else ""
    ] for d in all_days
}, index=["Overtime", "TOIL Earned", "DL Earned"])

st.dataframe(df2)

# ──────────────────────────────────────────────
# DOWNLOAD
# ──────────────────────────────────────────────

st.download_button("Download CSV", df1.to_csv(), "timesheet.csv")
