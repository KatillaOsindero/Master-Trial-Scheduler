# visit_scheduler.py — Single-patient only
# Features: protocol dropdown, single/range blackouts, anchor suggestions,
# red-dot flag for out-of-window, Start Time & Duration dropdowns,
# participant Excel handout (Visit Name + Chosen Date), Outlook .ics export,
# mm/dd/yyyy formatting.

import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta, time as dtime
from pathlib import Path
import io

st.set_page_config(page_title="Visit Scheduler", layout="wide")
st.markdown("# 🧬 Visit Scheduler")
st.caption("Choose a protocol (IGC / ATAI / Reunion ADCO), set a patient and constraints, pick dates, then export. No file uploads needed.")

REQUIRED_COLS = ["Day From Baseline", "Window Minus", "Window Plus"]

# ---------- Helpers ----------
def _to_date(x):
    if pd.isna(x) or x == "":
        return None
    if isinstance(x, (date, datetime)):
        return pd.to_datetime(x).date()
    try:
        return pd.to_datetime(x).date()
    except Exception:
        return None

def _to_time(x):
    if x is None or (isinstance(x, float) and pd.isna(x)) or str(x).strip() == "":
        return None
    s = str(x).strip()
    for fmt in ["%I:%M %p", "%H:%M", "%I:%M%p", "%H%M"]:
        try:
            return datetime.strptime(s, fmt).time()
        except Exception:
            continue
    try:
        return pd.to_datetime(s).time()
    except Exception:
        return None

def nth_weekday_of_month(year, month, weekday, n):
    d = date(year, month, 1)
    while d.weekday() != weekday:
        d += timedelta(days=1)
    d += timedelta(weeks=n-1)
    return d

def last_weekday_of_month(year, month, weekday):
    d = date(year, month + 1, 1) - timedelta(days=1) if month < 12 else date(year, 12, 31)
    while d.weekday() != weekday:
        d -= timedelta(days=1)
    return d

def observed(dt: date):
    if dt.weekday() == 5:  # Sat
        return dt - timedelta(days=1)
    if dt.weekday() == 6:  # Sun
        return dt + timedelta(days=1)
    return dt

def us_federal_holidays(year: int):
    hol = set()
    hol.add(observed(date(year, 1, 1)))
    hol.add(nth_weekday_of_month(year, 1, 0, 3))      # MLK
    hol.add(nth_weekday_of_month(year, 2, 0, 3))      # Presidents
    hol.add(last_weekday_of_month(year, 5, 0))        # Memorial
    hol.add(observed(date(year, 6, 19)))              # Juneteenth
    hol.add(observed(date(year, 7, 4)))               # Independence
    hol.add(nth_weekday_of_month(year, 9, 0, 1))      # Labor
    hol.add(nth_weekday_of_month(year, 10, 0, 2))     # Columbus
    hol.add(observed(date(year, 11, 11)))             # Veterans
    hol.add(nth_weekday_of_month(year, 11, 3, 4))     # Thanksgiving
    hol.add(observed(date(year, 12, 25)))             # Christmas
    return hol

def build_holiday_set(date_min: date, date_max: date, include_us_federal: bool):
    if not include_us_federal or date_min is None or date_max is None:
        return set()
    years = set(range(date_min.year, date_max.year + 1))
    out = set()
    for y in years:
        out |= us_federal_holidays(y)
    return out

def nearest_allowed_date(target, earliest, latest, disallow_weekends, holiday_set, custom_blackouts):
    if target is None or earliest is None or latest is None:
        return None
    target = pd.to_datetime(target).date()
    earliest = pd.to_datetime(earliest).date()
    latest = pd.to_datetime(latest).date()

    def allowed(d):
        if d < earliest or d > latest:
            return False
        if disallow_weekends and d.weekday() >= 5:
            return False
        if d in holiday_set:
            return False
        if d in custom_blackouts:
            return False
        return True

    if allowed(target):
        return target

    span = (latest - earliest).days
    for k in range(1, span + 1):
        plus = target + timedelta(days=k)
        minus = target - timedelta(days=k)
        if allowed(plus):
            return plus
        if allowed(minus):
            return minus
    for cand in [earliest, latest]:
        if (not disallow_weekends or cand.weekday() < 5) and cand not in holiday_set and cand not in custom_blackouts:
            return cand
    return None

def window_has_allowed_date(earliest, latest, disallow_weekends, holiday_set, custom_blackouts):
    if earliest is None or latest is None:
        return False
    e = pd.to_datetime(earliest).date()
    l = pd.to_datetime(latest).date()
    d = e
    while d <= l:
        if (not disallow_weekends or d.weekday() < 5) and (d not in holiday_set) and (d not in custom_blackouts):
            return True
        d += timedelta(days=1)
    return False

def suggest_anchor_dates(anchor, schedule_df, disallow_weekends, include_us_holidays, custom_blackouts, search_days=60):
    anchor = _to_date(anchor)
    if anchor is None:
        return None, None

    min_day = int(schedule_df["Day From Baseline"].min() - schedule_df["Window Minus"].max() - 7)
    max_day = int(schedule_df["Day From Baseline"].max() + schedule_df["Window Plus"].max() + 7)

    def all_visits_feasible(proposed_anchor):
        date_min = proposed_anchor + timedelta(days=min_day)
        date_max = proposed_anchor + timedelta(days=max_day)
        holiday_set = build_holiday_set(date_min, date_max, include_us_holidays)

        tmp = schedule_df.copy()
        tmp["Target Date"] = pd.to_datetime(proposed_anchor) + pd.to_timedelta(tmp["Day From Baseline"], unit="D")
        tmp["Earliest"] = tmp["Target Date"] - pd.to_timedelta(tmp["Window Minus"], unit="D")
        tmp["Latest"]   = tmp["Target Date"] + pd.to_timedelta(tmp["Window Plus"], unit="D")

        for _, r in tmp.iterrows():
            if not window_has_allowed_date(r["Earliest"], r["Latest"], disallow_weekends, holiday_set, custom_blackouts):
                return False
        return True

    earlier, later = None, None
    for delta in range(1, search_days + 1):
        cand_e = anchor - timedelta(days=delta)
        if earlier is None and all_visits_feasible(cand_e):
            earlier = cand_e
        cand_l = anchor + timedelta(days=delta)
        if later is None and all_visits_feasible(cand_l):
            later = cand_l
        if earlier is not None and later is not None:
            break
    return earlier, later

def make_ics(events, cal_name="Visit Schedule"):
    def dtstamp():
        return datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    lines = [
        "BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//Visit Scheduler//EN",f"X-WR-CALNAME:{cal_name}"
    ]
    for ev in events:
        start_t = ev.get("start_time") or dtime(hour=9, minute=0)
        dur_min = int(ev.get("duration_min") or 60)
        start_dt = datetime.combine(ev["date"], start_t)
        end_dt = start_dt + timedelta(minutes=dur_min)
        lines += [
            "BEGIN:VEVENT",
            f"UID:{hash((ev['summary'], ev['date'], start_t, dur_min))}@visitscheduler",
            f"DTSTAMP:{dtstamp()}",
            f"DTSTART:{start_dt.strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{end_dt.strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{ev['summary']}",
        ]
        if ev.get("description"):
            desc = str(ev["description"]).replace(",", r"\,").replace(";", r"\;")
            lines.append(f"DESCRIPTION:{desc}")
        lines.append("END:VEVENT")
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines).encode("utf-8")

# ---------- Sidebar ----------
with st.sidebar:
    st.subheader("⚙️ Settings")
    disallow_weekends = st.toggle("Disallow weekends", value=True)
    include_us_holidays = st.toggle("Exclude US Federal Holidays", value=True)
    st.caption("Holiday dates are observed when they fall on a weekend.")

# ---------- Protocol loader ----------
def list_protocol_csvs():
    protodir = Path(__file__).parent / "protocols"
    mapping = {"IGC": protodir / "IGC.csv","ATAI": protodir / "ATAI.csv","Reunion ADCO": protodir / "Reunion_ADCO.csv"}
    return {label: p for label, p in mapping.items() if p.exists()}

def load_protocol(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    df.columns = df.columns.str.strip()
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"CSV missing required column(s): {', '.join(missing)}")
    if "Visit Name" not in df.columns:
        df["Visit Name"] = [f"Visit {i+1}" for i in range(len(df))]
    for c in REQUIRED_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0).astype(int)
    if "Start Time" not in df.columns: df["Start Time"] = ""
    if "Visit Duration (min)" not in df.columns: df["Visit Duration (min)"] = ""
    return df

protocols = list_protocol_csvs()
if not protocols:
    st.error("No protocols found. Ensure at least one of: `protocols/IGC.csv`, `protocols/ATAI.csv`, `protocols/Reunion_ADCO.csv`.")
    st.stop()

proto_name = st.selectbox("📄 Choose protocol", list(protocols.keys()))
try:
    schedule = load_protocol(protocols[proto_name])
except Exception as e:
    st.error(f"Error loading protocol: {e}")
    st.stop()

# ---------- Patient ----------
st.markdown("## 👤 Patient")
c1, c2 = st.columns([1, 1])
with c1:
    patient_id = st.text_input("Patient ID")
with c2:
    anchor_date = st.date_input("Anchor Date", value=date.today(), key="anchor_date_input")
notes = st.text_area("Optional Notes (internal)")
ready = bool(patient_id and anchor_date)

# ---------- Blackouts ----------
st.markdown("## 🚫 Blackouts & Constraints")
c_top1, c_top2 = st.columns([1, 1])
with c_top1:
    st.markdown("**Custom blackout dates (single-day)**")
    single_seed = pd.DataFrame({"Blackout Date": pd.Series([], dtype="datetime64[ns]")})
    single_blackouts_df = st.data_editor(
        single_seed, num_rows="dynamic", use_container_width=True,
        column_config={"Blackout Date": st.column_config.DateColumn()},
        key="blackouts_single_editor"
    ).dropna(subset=["Blackout Date"])

with c_top2:
    st.markdown("**Custom blackout ranges (start–end, inclusive)**")
    range_seed = pd.DataFrame({
        "Start": pd.Series([], dtype="datetime64[ns]"),
        "End":   pd.Series([], dtype="datetime64[ns]"),
    })
    range_blackouts_df = st.data_editor(
        range_seed, num_rows="dynamic", use_container_width=True,
        column_config={"Start": st.column_config.DateColumn(), "End": st.column_config.DateColumn()},
        key="blackouts_range_editor"
    ).dropna(subset=["Start", "End"])

def expand_ranges(df):
    dates = set()
    for _, r in df.iterrows():
        s = _to_date(r["Start"]); e = _to_date(r["End"])
        if s and e and e >= s:
            d = s
            while d <= e:
                dates.add(d)
                d += timedelta(days=1)
    return dates

custom_blackouts = set(d.date() for d in pd.to_datetime(single_blackouts_df["Blackout Date"]).dropna().tolist())
custom_blackouts |= expand_ranges(range_blackouts_df)

# ---------- Core compute ----------
def compute_visits_for_patient(anchor: date):
    min_day = int(schedule["Day From Baseline"].min() - schedule["Window Minus"].max() - 7)
    max_day = int(schedule["Day From Baseline"].max() + schedule["Window Plus"].max() + 7)
    date_min = anchor + timedelta(days=min_day)
    date_max = anchor + timedelta(days=max_day)
    holiday_set = build_holiday_set(date_min, date_max, include_us_holidays)

    out = schedule.copy()
    out["Target Date"] = pd.to_datetime(anchor) + pd.to_timedelta(out["Day From Baseline"], unit="D")
    out["Earliest"]    = out["Target Date"] - pd.to_timedelta(out["Window Minus"], unit="D")
    out["Latest"]      = out["Target Date"] + pd.to_timedelta(out["Window Plus"], unit="D")

    chosen = []
    for _, r in out.iterrows():
        ch = nearest_allowed_date(
            r["Target Date"].date(),
            r["Earliest"].date(),
            r["Latest"].date(),
            disallow_weekends,
            holiday_set,
            custom_blackouts
        )
        chosen.append(ch)
    out["Chosen Date"] = chosen
    if "Start Time" not in out.columns: out["Start Time"] = ""
    if "Visit Duration (min)" not in out.columns: out["Visit Duration (min)"] = ""
    return out

def _coerce_to_date_cols(df, cols):
    for c in cols:
        df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
    return df

def _annotate_status_and_flag(df):
    def status_row(r):
        cd = r.get("Chosen Date")
        e, l = r.get("Earliest"), r.get("Latest")
        cd, e, l = _to_date(cd), _to_date(e), _to_date(l)
        if cd is None or e is None or l is None:
            return "⏳ Not set", ""
        return ("✅ In window", "") if (e <= cd <= l) else ("🔴 Out of window", "🔴")

    df = df.copy()
    statuses, flags = [], []
    for _, r in df.iterrows():
        s, f = status_row(r)
        statuses.append(s); flags.append(f)
    df["Status"] = statuses
    if "⚠" in df.columns:
        df.drop(columns=["⚠"], inplace=True)
    df.insert(df.columns.get_loc("Chosen Date"), "⚠", flags)
    return df, any(s == "🔴 Out of window" for s in statuses)

def _format_mmddyyyy(df, cols):
    df = df.copy()
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.strftime("%m/%d/%Y")
    return df

# Dropdown options
TIME_OPTS = [f"{(h%12) or 12}:{m:02d} {'AM' if h < 12 else 'PM'}" for h in range(24) for m in (0, 30)]
DUR_OPTS  = list(range(30, 721, 30))  # minutes up to 12 hours

# ---------- Schedule & adjust ----------
if ready:
    st.markdown("## 📅 Schedule & Adjust")
    st.info(
        "How to choose a date:\n"
        "1) Click the **Chosen Date** cell for a visit.\n"
        "2) Pick a date from the calendar popup (mm/dd/yyyy).\n"
        "3) A red dot **🔴** next to the date means it is **out of window**.\n"
        "4) (Optional) Pick **Start Time** and **Visit Duration** to include in Outlook export."
    )

    visits = compute_visits_for_patient(anchor_date)
    table = _coerce_to_date_cols(visits.copy(), ["Target Date", "Earliest", "Latest", "Chosen Date"])

    table = st.data_editor(
        table,
        use_container_width=True,
        num_rows="fixed",
        column_config={
            "Visit Name": st.column_config.TextColumn(disabled=True),
            "Day From Baseline": st.column_config.NumberColumn(disabled=True),
            "Window Minus": st.column_config.NumberColumn(disabled=True),
            "Window Plus": st.column_config.NumberColumn(disabled=True),
            "Target Date": st.column_config.DateColumn(disabled=True),
            "Earliest": st.column_config.DateColumn(disabled=True),
            "Latest": st.column_config.DateColumn(disabled=True),
            "Chosen Date": st.column_config.DateColumn(),
            "Start Time": st.column_config.SelectboxColumn(options=TIME_OPTS, required=False),
            "Visit Duration (min)": st.column_config.SelectboxColumn(options=DUR_OPTS, required=False),
        },
        key="single_visits_editor"
    )
    table_with_status, any_out = _annotate_status_and_flag(table)

    # Conflict detection (no allowed day in a window)
    min_day = int(schedule["Day From Baseline"].min() - schedule["Window Minus"].max() - 7)
    max_day = int(schedule["Day From Baseline"].max() + schedule["Window Plus"].max() + 7)
    date_min = anchor_date + timedelta(days=min_day)
    date_max = anchor_date + timedelta(days=max_day)
    holiday_set_now = build_holiday_set(date_min, date_max, include_us_holidays)
    windows_ok = True
    tmp = schedule.copy()
    tmp["Target Date"] = pd.to_datetime(anchor_date) + pd.to_timedelta(tmp["Day From Baseline"], unit="D")
    tmp["Earliest"] = tmp["Target Date"] - pd.to_timedelta(tmp["Window Minus"], unit="D")
    tmp["Latest"]   = tmp["Target Date"] + pd.to_timedelta(tmp["Window Plus"], unit="D")
    for _, r in tmp.iterrows():
        if not window_has_allowed_date(r["Earliest"], r["Latest"], disallow_weekends, holiday_set_now, custom_blackouts):
            windows_ok = False
            break

    if any_out:
        st.warning("Some visits are **out of window** (🔴). Please adjust Chosen Date between Earliest and Latest.")
    if not windows_ok:
        earlier_sug, later_sug = suggest_anchor_dates(anchor_date, schedule, disallow_weekends, include_us_holidays, custom_blackouts, search_days=60)
        with st.container(border=True):
            st.error("Blackouts/constraints remove all allowed days for at least one visit window.")
            msg = "Suggested anchor dates:"
            if earlier_sug: msg += f" **Earlier:** {earlier_sug.strftime('%m/%d/%Y')}"
            if later_sug:   msg += f" | **Later:** {later_sug.strftime('%m/%d/%Y')}"
            st.write(msg)
            c1, c2 = st.columns(2)
            if earlier_sug:
                if c1.button(f"Apply earlier anchor ({earlier_sug.strftime('%m/%d/%Y')})"):
                    st.session_state["anchor_date_input"] = earlier_sug
                    st.rerun()
            if later_sug:
                if c2.button(f"Apply later anchor ({later_sug.strftime('%m/%d/%Y')})"):
                    st.session_state["anchor_date_input"] = later_sug
                    st.rerun()

    st.session_state["single_result"] = {
        "patient_id": patient_id,
        "anchor_date": anchor_date,
        "notes": (notes or ""),
        "df": table_with_status.copy()
    }

# ---------- Export & print ----------
st.markdown("## 🖨️ Export & Print")
role = st.radio("Role", ["Coordinator view", "Participant handout"], horizontal=True)

def coordinator_view(df):
    cols = ["Visit Name","Day From Baseline","Target Date","Earliest","Latest","⚠","Chosen Date",
            "Status","Window Minus","Window Plus","Start Time","Visit Duration (min)"]
    out = df.copy()
    out = out[[c for c in cols if c in out.columns]].copy()
    out = _format_mmddyyyy(out, ["Target Date","Earliest","Latest","Chosen Date"])
    return out

def participant_view(df):
    out = df.copy()[["Visit Name","Chosen Date"]].copy()
    out = _format_mmddyyyy(out, ["Chosen Date"])
    return out

left, right = st.columns([2, 1])

with left:
    if "single_result" in st.session_state and ready:
        res = st.session_state["single_result"]
        df = res["df"].copy()
        if role == "Coordinator view":
            table = coordinator_view(df)
        else:
            table = participant_view(df)

        st.dataframe(table, use_container_width=True)

        # CSV
        st.download_button(
            "⬇️ Download CSV",
            data=table.to_csv(index=False).encode("utf-8"),
            file_name=f"{res['patient_id']}_{'coordinator' if role=='Coordinator view' else 'participant'}.csv",
            mime="text/csv"
        )
        # Excel
        xbuf = io.BytesIO()
        with pd.ExcelWriter(xbuf, engine="xlsxwriter") as writer:
            sheet = "Schedule" if role == "Coordinator view" else "Participant Handout"
            table.to_excel(writer, index=False, sheet_name=sheet)
        st.download_button(
            "⬇️ Download Excel",
            data=xbuf.getvalue(),
            file_name=f"{res['patient_id']}_{'coordinator' if role=='Coordinator view' else 'participant'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Outlook ICS (only uses Chosen Date rows; optional time/duration if present)
        if role == "Coordinator view":
            events = []
            for _, r in df.iterrows():
                cd = _to_date(r.get("Chosen Date"))
                if not cd:
                    continue
                start_t = _to_time(r.get("Start Time"))
                dur = r.get("Visit Duration (min)")
                try:
                    dur = int(dur) if str(dur).strip() != "" else None
                except Exception:
                    dur = None
                summary = f"{res['patient_id']} · {r.get('Visit Name','Visit')}"
                desc = f"{proto_name} — Window {r.get('Earliest')} to {r.get('Latest')}"
                events.append({"summary": summary, "date": cd, "start_time": start_t, "duration_min": dur, "description": desc})
            if events:
                ics_data = make_ics(events, cal_name=f"{proto_name} - {res['patient_id']}")
                st.download_button("📅 Export to Outlook calendar", data=ics_data,
                                   file_name=f"{res['patient_id']}_schedule.ics", mime="text/calendar")
            else:
                st.info("Set at least one **Chosen Date** to enable calendar export.")

with right:
    st.markdown("### 🧾 Tips")
    st.write("- **Single-day** and **range** blackouts are supported; ranges are inclusive.")
    st.write("- If blackouts make a visit impossible, you’ll see **suggested anchor dates** (apply with one click).")
    st.write("- **Red dot (🔴)** next to a date means it’s **out of window**.")
    st.write("- Use **Start Time** and **Visit Duration** dropdowns if you want them included in Outlook exports.")
    st.write("- All dates display/export as **mm/dd/yyyy**.")
