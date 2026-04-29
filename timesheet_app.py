import streamlit as st
import pandas as pd
import calendar
import holidays
from datetime import date, datetime, timedelta
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Timesheet Calculator", page_icon="🕐", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
:root{--bg:#0f1117;--surface:#1a1d27;--surface2:#242838;--accent:#4f9eff;--accent2:#ff6b6b;--yellow:#ffd166;--green:#06d6a0;--text:#e8eaf0;--text2:#8b8fa8;--border:#2e3248;}
html,body,[class*="css"]{font-family:'IBM Plex Sans',sans-serif;color:var(--text);}
.stApp{background:var(--bg);}
h1,h2,h3{font-family:'IBM Plex Mono',monospace;}
.stSidebar{background:var(--surface) !important;border-right:1px solid var(--border);}
.stSidebar [data-testid="stSidebarContent"]{padding:1.5rem 1rem;}
.ts-header{background:linear-gradient(135deg,var(--surface) 0%,var(--surface2) 100%);border:1px solid var(--border);border-left:4px solid var(--accent);padding:1.5rem 2rem;border-radius:8px;margin-bottom:1.5rem;}
.ts-header h1{color:var(--accent);margin:0;font-size:1.6rem;letter-spacing:-0.5px;}
.ts-header p{color:var(--text2);margin:0.25rem 0 0;font-size:0.85rem;}
.ts-section-title{font-family:'IBM Plex Mono',monospace;font-size:0.75rem;font-weight:600;letter-spacing:2px;text-transform:uppercase;color:var(--accent);margin-bottom:1rem;padding-bottom:0.5rem;border-bottom:1px solid var(--border);}
.metric-row{display:flex;gap:1rem;margin:1rem 0;flex-wrap:wrap;}
.metric-card{flex:1;min-width:130px;background:var(--surface2);border:1px solid var(--border);border-radius:6px;padding:0.75rem 1rem;text-align:center;}
.metric-val{font-family:'IBM Plex Mono',monospace;font-size:1.6rem;font-weight:600;}
.metric-lbl{font-size:0.7rem;color:var(--text2);letter-spacing:1px;text-transform:uppercase;margin-top:2px;}
.stButton button{background:var(--accent) !important;color:#000 !important;font-family:'IBM Plex Mono',monospace !important;font-weight:600 !important;border:none !important;border-radius:4px !important;padding:0.4rem 1.2rem !important;}
.stTabs [data-baseweb="tab-list"]{background:var(--surface) !important;border-radius:6px;padding:4px;}
.stTabs [data-baseweb="tab"]{color:var(--text2) !important;}
.stTabs [aria-selected="true"]{color:var(--accent) !important;background:var(--surface2) !important;border-radius:4px !important;}
.desc-box{background:var(--surface2);border:1px solid var(--border);border-left:3px solid var(--yellow);border-radius:6px;padding:1rem 1.25rem;font-family:'IBM Plex Mono',monospace;font-size:0.8rem;line-height:1.7;white-space:pre-wrap;color:var(--text);margin-top:0.5rem;}
</style>
""", unsafe_allow_html=True)

# ── HELPERS ──────────────────────────────────────────────────────────────────

def get_polish_holidays(year, month):
    pl = holidays.Poland(years=year)
    return {d: n for d, n in pl.items() if d.month == month}

def is_non_working(d, holiday_dates):
    return d.weekday() >= 5 or d in holiday_dates

def next_working_day(d, holiday_dates):
    nd = d + timedelta(days=1)
    while is_non_working(nd, holiday_dates):
        nd += timedelta(days=1)
    return nd

def parse_hhmm(s):
    try:
        p = s.strip().split(":")
        return int(p[0]) + int(p[1]) / 60
    except Exception:
        return None

def hhmm(h):
    hh = int(h)
    mm = int(round((h - hh) * 60))
    return f"{hh:02d}:{mm:02d}"

def fmt(v):
    if not v:
        return ""
    return str(round(float(v), 2))

def split_hours(s_d, s_h, e_d, e_h):
    """Return (night_h, day_h) — mutually exclusive. night=22:00-04:00, day=everything else."""
    night_h = 0.0
    day_h   = 0.0
    cur = datetime(s_d.year, s_d.month, s_d.day) + timedelta(hours=s_h)
    end = datetime(e_d.year, e_d.month, e_d.day) + timedelta(hours=e_h)
    while cur < end:
        nxt = min(cur + timedelta(minutes=30), end)
        mid = cur + timedelta(minutes=15)
        hm = mid.hour + mid.minute / 60
        chunk = (nxt - cur).total_seconds() / 3600
        if hm >= 22 or hm < 4:
            night_h += chunk
        else:
            day_h += chunk
        cur = nxt
    return night_h, day_h

# ── SESSION STATE ────────────────────────────────────────────────────────────

for k, v in {"incidents": [], "overtime_entries": [], "dl_taken": []}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ── SIDEBAR ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### 📅 Period")
    c1, c2 = st.columns(2)
    year  = int(c1.number_input("Year",  min_value=2020, max_value=2035, value=date.today().year,  step=1))
    month = int(c2.number_input("Month", min_value=1,    max_value=12,   value=date.today().month, step=1))

    _, days_in_month = calendar.monthrange(year, month)
    all_days = [date(year, month, d) for d in range(1, days_in_month + 1)]

    pl_raw = get_polish_holidays(year, month)
    holiday_dates_set = set(pl_raw.keys())

    st.markdown("### 🇵🇱 Polish Holidays")
    if pl_raw:
        hol_override = {hd: st.checkbox(f"{hd.strftime('%d')} – {hn}", value=True, key=f"hol_{hd}") for hd, hn in pl_raw.items()}
        holiday_dates_set = {d for d, on in hol_override.items() if on}
    else:
        st.caption("No national holidays this month.")

    st.markdown("### ⏱ Daily Hours")
    creative_min = st.number_input("Creative MIN (h/day)", min_value=0.0, max_value=8.0, value=2.0, step=0.5)
    creative_max = st.number_input("Creative MAX (h/day)", min_value=0.0, max_value=8.0, value=4.0, step=0.5)
    support_default = st.number_input("Support (h/day)",   min_value=0.0, max_value=8.0, value=4.0, step=0.5)

    st.markdown("### ☕ Break Time")
    add_break = st.checkbox("Add 0.5h break (Others) each working day", value=False)

# ── HEADER ───────────────────────────────────────────────────────────────────

working_days = sum(1 for d in all_days if not is_non_working(d, holiday_dates_set))
st.markdown(f"""
<div class="ts-header">
  <h1>🕐 TIMESHEET CALCULATOR</h1>
  <p>{calendar.month_name[month]} {year} &nbsp;·&nbsp; Poland &nbsp;·&nbsp; {working_days} working days</p>
</div>
""", unsafe_allow_html=True)

tab1, tab2, tab3, tab4 = st.tabs(["📋 Timesheets", "🚨 Incidents & Overtime", "📝 Description", "📥 Export"])

# ── TAB 2: INPUT ─────────────────────────────────────────────────────────────

with tab2:
    st.markdown('<div class="ts-section-title">ADD INCIDENT / OVERTIME</div>', unsafe_allow_html=True)

    with st.form("entry_form", clear_on_submit=True):
        f1, f2, f3, f4 = st.columns(4)
        entry_type     = f1.selectbox("Type", ["Incident", "Overtime"])
        entry_id       = f2.text_input("ID / Reference", placeholder="INC0001234")
        start_date_inp = f3.date_input("Start date", value=date(year, month, 1),
                                       min_value=date(year, month, 1), max_value=date(year, month, days_in_month))
        start_time_inp = f3.text_input("Start time (HH:MM)", value="17:00")
        end_date_inp   = f4.date_input("End date", value=date(year, month, 1),
                                       min_value=date(year, month, 1), max_value=date(year, month, days_in_month))
        end_time_inp   = f4.text_input("End time (HH:MM)", value="19:00")
        if st.form_submit_button("➕ Add Entry"):
            sh = parse_hhmm(start_time_inp)
            eh = parse_hhmm(end_time_inp)
            if sh is None or eh is None:
                st.error("Invalid time format. Use HH:MM")
            else:
                entry = {"type": entry_type, "id": entry_id or entry_type,
                         "start_date": start_date_inp, "start_h": sh,
                         "end_date": end_date_inp, "end_h": eh}
                if entry_type == "Incident":
                    st.session_state.incidents.append(entry)
                else:
                    st.session_state.overtime_entries.append(entry)
                st.success(f"{entry_type} added!")

    st.markdown('<div class="ts-section-title" style="margin-top:1.5rem">RECORD DL TAKEN DAY</div>', unsafe_allow_html=True)
    st.caption("Record when the employee actually takes a DL day off. Independent from when it was earned.")

    with st.form("dl_form", clear_on_submit=True):
        d1, d2 = st.columns(2)
        dl_date = d1.date_input("Day taken as DL", value=date(year, month, 1),
                                min_value=date(year, month, 1), max_value=date(year, month, days_in_month))
        dl_ref  = d2.text_input("Reference (INC ID that earned this DL)", placeholder="INC0001234")
        if st.form_submit_button("➕ Add DL Taken Day"):
            if is_non_working(dl_date, holiday_dates_set):
                st.error("Cannot take DL on a weekend or holiday.")
            else:
                st.session_state.dl_taken.append({"date": dl_date, "ref": dl_ref or "DL"})
                st.success(f"DL taken on {dl_date.strftime('%d/%m/%Y')} recorded.")

    all_rows = []
    for e in st.session_state.incidents + st.session_state.overtime_entries:
        total_h = (e["end_date"] - e["start_date"]).days * 24 + e["end_h"] - e["start_h"]
        all_rows.append({"Type": e["type"], "ID": e["id"],
                         "Start": f"{e['start_date']} {hhmm(e['start_h'])}",
                         "End":   f"{e['end_date']} {hhmm(e['end_h'])}",
                         "Hours": round(total_h, 2)})
    for dl in st.session_state.dl_taken:
        all_rows.append({"Type": "DL Taken", "ID": dl["ref"],
                         "Start": str(dl["date"]), "End": str(dl["date"]), "Hours": 8.0})

    if all_rows:
        st.dataframe(pd.DataFrame(all_rows), use_container_width=True)
        bc1, bc2 = st.columns(2)
        if bc1.button("🗑 Clear incidents & overtime"):
            st.session_state.incidents = []
            st.session_state.overtime_entries = []
            st.rerun()
        if bc2.button("🗑 Clear DL taken days"):
            st.session_state.dl_taken = []
            st.rerun()

# ── BUSINESS LOGIC ───────────────────────────────────────────────────────────

day_data = {}
for d in all_days:
    day_data[d] = {
        "is_non_working": is_non_working(d, holiday_dates_set),
        # Regular table
        "creative": 0.0, "support": 0.0, "others": 0.0,
        "abs_rt": 0.0, "abs_toil": 0.0, "abs_etoil": 0.0, "abs_dl": 0.0,
        # Overtime table (shown on the day OT/INC actually happened, incl. weekends)
        "ot_support": 0.0, "ot_others": 0.0, "ot_night": 0.0,
        "ot_toil": 0.0,        # TOIL earned — on the OT day column
        "dl_earned": False,    # DL earned flag — on the OT day column
        "dl_earned_ref": "",
        "description_lines": [],
    }

# Seed regular hours (will be adjusted after absences are computed)
for d in all_days:
    if not day_data[d]["is_non_working"]:
        day_data[d]["creative"] = creative_max
        day_data[d]["support"]  = support_default
        if add_break:
            day_data[d]["others"] = 0.5


def process_entry(entry):
    sd, ed = entry["start_date"], entry["end_date"]
    sh, eh = entry["start_h"],   entry["end_h"]
    etype, eid = entry["type"],  entry["id"]

    total_h = (ed - sd).days * 24 + eh - sh
    if total_h <= 0:
        return

    if etype == "Incident":
        # Attribution: only re-attribute to previous day when the incident starts
        # before 09:00 on a plain weekday (not weekend, not holiday).
        # Weekend/holiday incidents always report on the day they actually occurred.
        if sh < 9.0 and sd.weekday() < 5 and sd not in holiday_dates_set:
            ot_day = sd - timedelta(days=1)
            while ot_day not in day_data:
                ot_day -= timedelta(days=1)
        else:
            ot_day = sd

        if ot_day not in day_data:
            ot_day = sd

        is_dl_eligible = (
            sd.weekday() >= 5 or sd in holiday_dates_set
            or (ed != sd and (ed.weekday() >= 5 or ed in holiday_dates_set))
        )

        night_h, day_h = split_hours(sd, sh, ed, eh)

        # OT on the actual day — night and non-night hours are independent buckets
        day_data[ot_day]["ot_others"] += day_h
        day_data[ot_day]["ot_night"]  += night_h
        if not is_dl_eligible:
            day_data[ot_day]["ot_toil"] += total_h  # TOIL earned = full incident hours

        if is_dl_eligible:
            day_data[ot_day]["dl_earned"]     = True
            day_data[ot_day]["dl_earned_ref"] = eid

        # Description for the incident itself (always logged)
        reported_on = f" (reported on {ot_day.strftime('%d/%m')})" if ot_day != sd else ""
        day_data[ot_day]["description_lines"].append((
            "INC",
            f"{sd.strftime('%d/%m/%Y')} - INC - {hhmm(sh)} to {hhmm(eh)}{reported_on} - {hhmm(total_h)}"
        ))

        if is_dl_eligible:
            # DL only — no TOIL, E-TOIL or RT
            day_data[ot_day]["description_lines"].append((
                "DL_EARNED",
                f"{ot_day.strftime('%d/%m/%Y')} - DL EARNED FOR INC {eid}"
            ))
        else:
            # TOIL / E-TOIL / RT → next working day (regular table)
            toil_day = next_working_day(ot_day, holiday_dates_set)
            if toil_day in day_data:
                day_data[toil_day]["abs_toil"]  += total_h
                day_data[toil_day]["abs_etoil"] += round(total_h * 0.5, 2)
                if night_h > 0:
                    day_data[toil_day]["abs_rt"] += night_h
                toil_line = (f"{toil_day.strftime('%d/%m/%Y')} - "
                             f"TOIL {hhmm(total_h)} + E-TOIL {hhmm(round(total_h*0.5,2))}")
                if night_h > 0:
                    toil_line += f" + RT {hhmm(night_h)}"
                day_data[toil_day]["description_lines"].append(("TOIL", toil_line))

    else:  # Overtime — no TOIL
        if sd in day_data:
            day_data[sd]["ot_support"] += total_h
            day_data[sd]["description_lines"].append((
                "OT",
                f"{sd.strftime('%d/%m/%Y')} - OT - {hhmm(sh)} to {hhmm(eh)} - {hhmm(total_h)}"
            ))


for e in st.session_state.incidents:
    process_entry(e)
for e in st.session_state.overtime_entries:
    process_entry(e)

# Apply DL taken days
for dl in st.session_state.dl_taken:
    d, ref = dl["date"], dl["ref"]
    if d in day_data and not day_data[d]["is_non_working"]:
        day_data[d]["abs_dl"] = 8.0
        day_data[d]["description_lines"].append((
            "DL_TAKEN",
            f"{d.strftime('%d/%m/%Y')} - DL TAKEN (ref: {ref})"
        ))

# Adjust creative/support so daily total = 8h after absences
TARGET = 8.0
for d in all_days:
    dd = day_data[d]
    if dd["is_non_working"]:
        continue
    absences = dd["abs_rt"] + dd["abs_toil"] + dd["abs_etoil"] + dd["abs_dl"]
    available = max(0.0, TARGET - absences - dd["others"])
    creative = max(creative_min, min(creative_max, available))
    support  = max(0.0, min(support_default, available - creative))
    dd["creative"] = round(creative, 2)
    dd["support"]  = round(support,  2)

# ── COLUMN LABELS ─────────────────────────────────────────────────────────────

col_labels = [f"{d.day}\n{calendar.day_abbr[d.weekday()]}" for d in all_days]

# ── SHARED STYLER ─────────────────────────────────────────────────────────────

def make_styler(df, col_labels, all_days, holiday_dates_set, total_row=None):
    def _style(df_):
        styles = pd.DataFrame("", index=df_.index, columns=df_.columns)
        for col in df_.columns:
            try:
                idx = col_labels.index(col)
                d = all_days[idx]
                if is_non_working(d, holiday_dates_set):
                    styles[col] = "background-color:#3d3800; color:#ffd166;"
            except ValueError:
                pass
        if total_row and total_row in styles.index:
            for col in styles.columns:
                styles.loc[total_row, col] += " font-weight:bold;"
        return styles
    return df.style.apply(_style, axis=None)

# ── TAB 1: TIMESHEETS ─────────────────────────────────────────────────────────

with tab1:

    # TABLE 1
    st.markdown('<div class="ts-section-title">TABLE 1 — REGULAR WORKING HOURS</div>', unsafe_allow_html=True)

    rows_t1 = {
        "ECoE / Creative Work": [], "ECoE / Support": [], "ECoE / Others": [],
        "Absence / RT": [], "Absence / TOIL": [], "Absence / E-TOIL": [], "Absence / DL": [],
        "── TOTAL ──": [],
    }

    for d in all_days:
        dd  = day_data[d]
        nw  = dd["is_non_working"]
        rows_t1["ECoE / Creative Work"].append("" if nw else fmt(dd["creative"]))
        rows_t1["ECoE / Support"].append(       "" if nw else fmt(dd["support"]))
        rows_t1["ECoE / Others"].append(        "" if nw else fmt(dd["others"]))
        rows_t1["Absence / RT"].append(         "" if nw else fmt(dd["abs_rt"]))
        rows_t1["Absence / TOIL"].append(       "" if nw else fmt(dd["abs_toil"]))
        rows_t1["Absence / E-TOIL"].append(     "" if nw else fmt(dd["abs_etoil"]))
        rows_t1["Absence / DL"].append(         "" if nw else fmt(dd["abs_dl"]))
        if nw:
            rows_t1["── TOTAL ──"].append("")
        else:
            total = (dd["creative"] + dd["support"] + dd["others"]
                     + dd["abs_rt"] + dd["abs_toil"] + dd["abs_etoil"] + dd["abs_dl"])
            rows_t1["── TOTAL ──"].append(fmt(total))

    df1 = pd.DataFrame(rows_t1, index=col_labels).T
    st.dataframe(
        make_styler(df1, col_labels, all_days, holiday_dates_set, total_row="── TOTAL ──"),
        use_container_width=True, height=320,
    )

    total_project = sum(day_data[d]["creative"] + day_data[d]["support"] + day_data[d]["others"]
                        for d in all_days if not day_data[d]["is_non_working"])
    total_absence = sum(day_data[d]["abs_rt"] + day_data[d]["abs_toil"]
                        + day_data[d]["abs_etoil"] + day_data[d]["abs_dl"]
                        for d in all_days)

    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card"><div class="metric-val" style="color:#4f9eff">{total_project:.1f}h</div><div class="metric-lbl">Project Hours</div></div>
      <div class="metric-card"><div class="metric-val" style="color:#ff6b6b">{total_absence:.1f}h</div><div class="metric-lbl">Absence Hours</div></div>
      <div class="metric-card"><div class="metric-val" style="color:#06d6a0">{total_project+total_absence:.1f}h</div><div class="metric-lbl">Total</div></div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    # TABLE 2
    st.markdown('<div class="ts-section-title">TABLE 2 — OVERTIME HOURS</div>', unsafe_allow_html=True)

    rows_t2 = {
        "ECoE / Support": [],
        "ECoE / Others": [],
        "ECoE / Others / Night Time": [],
        "DL": [],
    }

    for d in all_days:
        dd = day_data[d]
        rows_t2["ECoE / Support"].append(fmt(dd["ot_support"]))
        rows_t2["ECoE / Others"].append(fmt(dd["ot_others"]))
        rows_t2["ECoE / Others / Night Time"].append(fmt(dd["ot_night"]))
        rows_t2["DL"].append("X" if dd["dl_earned"] else "")

    df2 = pd.DataFrame(rows_t2, index=col_labels).T

    def style_t2(df_):
        styles = pd.DataFrame("", index=df_.index, columns=df_.columns)
        for col in df_.columns:
            try:
                idx = col_labels.index(col)
                if is_non_working(all_days[idx], holiday_dates_set):
                    styles[col] = "background-color:#3d3800; color:#ffd166;"
            except ValueError:
                pass
        return styles

    st.dataframe(df2.style.apply(style_t2, axis=None), use_container_width=True, height=200)

    total_ot       = sum(day_data[d]["ot_support"] + day_data[d]["ot_others"] for d in all_days)
    total_dl_earn  = sum(1 for d in all_days if day_data[d]["dl_earned"])
    total_dl_taken = len(st.session_state.dl_taken)

    st.markdown(f"""
    <div class="metric-row">
      <div class="metric-card"><div class="metric-val" style="color:#ff6b6b">{total_ot:.1f}h</div><div class="metric-lbl">Total Overtime</div></div>
      <div class="metric-card"><div class="metric-val" style="color:#ffd166">{total_dl_earn}</div><div class="metric-lbl">DL Earned</div></div>
      <div class="metric-card"><div class="metric-val" style="color:#ffd166">{total_dl_taken}</div><div class="metric-lbl">DL Taken</div></div>
    </div>
    """, unsafe_allow_html=True)

# ── TAB 3: DESCRIPTION ────────────────────────────────────────────────────────

with tab3:
    st.markdown('<div class="ts-section-title">AUTO-GENERATED DESCRIPTION</div>', unsafe_allow_html=True)
    desc_lines = []
    for d in sorted(all_days):
        for _, line in day_data[d]["description_lines"]:
            desc_lines.append(line)
    desc_lines.append("SPREADSHEET IS ATTACHED")
    desc_text = "\n".join(desc_lines)

    st.markdown(f'<div class="desc-box">{desc_text}</div>', unsafe_allow_html=True)
    st.text_area("Edit Description (optional)", value=desc_text, height=300,
                 key="desc_edit", label_visibility="collapsed")

# ── TAB 4: EXPORT ─────────────────────────────────────────────────────────────

def build_excel():
    wb = openpyxl.Workbook()
    ws1 = wb.active;  ws1.title = "Regular Hours"
    ws2 = wb.create_sheet("Overtime Hours")
    ws3 = wb.create_sheet("Description")

    YEL_F  = PatternFill("solid", fgColor="FFD166")
    HDR_F  = PatternFill("solid", fgColor="1A1D27")
    TOT_F  = PatternFill("solid", fgColor="242838")
    BOD_F  = PatternFill("solid", fgColor="0F1117")
    HDR_FN = Font(name="Calibri", bold=True, color="4F9EFF", size=9)
    TOT_FN = Font(name="Calibri", bold=True, color="FFFFFF", size=9)
    BOD_FN = Font(name="Calibri", size=9, color="E8EAF0")
    YEL_FN = Font(name="Calibri", bold=True, color="3D3800", size=9)
    CTR    = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin   = Side(style="thin", color="2E3248")
    brd    = Border(left=thin, right=thin, top=thin, bottom=thin)

    def write_table(ws, rows_dict, grey_weekends=True):
        ws.cell(1, 1, "Category").font = HDR_FN
        ws.cell(1, 1).fill = HDR_F
        ws.cell(1, 1).alignment = CTR
        ws.cell(1, 1).border = brd
        ws.column_dimensions["A"].width = 24

        for ci, (d, lbl) in enumerate(zip(all_days, col_labels), start=2):
            nw = is_non_working(d, holiday_dates_set)
            c = ws.cell(1, ci, lbl)
            c.fill = YEL_F if nw else HDR_F
            c.font = YEL_FN if nw else HDR_FN
            c.alignment = CTR; c.border = brd
            ws.column_dimensions[get_column_letter(ci)].width = 5.5

        for ri, (rname, vals) in enumerate(rows_dict.items(), start=2):
            is_tot = "TOTAL" in rname
            ws.cell(ri, 1, rname).font = HDR_FN
            ws.cell(ri, 1).fill = TOT_F if is_tot else HDR_F
            ws.cell(ri, 1).alignment = CTR; ws.cell(ri, 1).border = brd
            for ci, (val, d) in enumerate(zip(vals, all_days), start=2):
                nw = is_non_working(d, holiday_dates_set)
                c = ws.cell(ri, ci, val if val != "" else None)
                if nw and grey_weekends:
                    c.fill = YEL_F; c.font = YEL_FN
                elif is_tot:
                    c.fill = TOT_F; c.font = TOT_FN
                else:
                    c.fill = BOD_F; c.font = BOD_FN
                c.alignment = CTR; c.border = brd

    write_table(ws1, rows_t1, grey_weekends=True)
    write_table(ws2, rows_t2, grey_weekends=False)

    ws3.column_dimensions["A"].width = 90
    for i, line in enumerate(desc_lines, start=1):
        c = ws3.cell(i, 1, line)
        c.font = Font(name="Courier New", size=10, color="E8EAF0")
        c.fill = PatternFill("solid", fgColor="0F1117")
        ws3.row_dimensions[i].height = 16

    buf = BytesIO()
    wb.save(buf); buf.seek(0)
    return buf


with tab4:
    st.markdown('<div class="ts-section-title">EXPORT TIMESHEET</div>', unsafe_allow_html=True)
    st.markdown("""
    <div style="background:#1a1d27;border:1px solid #2e3248;border-radius:6px;padding:1rem 1.25rem;margin-bottom:1rem;">
    <p style="margin:0;color:#8b8fa8;font-size:0.85rem;">
    Exports <strong style="color:#4f9eff">.xlsx</strong> with: Regular Hours · Overtime Hours · Description
    </p></div>
    """, unsafe_allow_html=True)

    fname = f"timesheet_{year}_{month:02d}.xlsx"
    st.download_button(
        label="📥 Download Excel Timesheet",
        data=build_excel(),
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption(f"File: `{fname}`")
