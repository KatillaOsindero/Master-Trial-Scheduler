# visit_scheduler.py — single patient; no procedures/staff; blackout dates & ranges;
# weekend exclusion; per-holiday toggles (observed); status in editor;
# reordered columns; taller tables; anchor suggestions; Outlook export; Word export (participant);
# protocol version shown ONLY in the protocol dropdown label. Participant handout labels "Visit Date".

import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta, time as dtime
from pathlib import Path
import io

# For Word export (participant handout)
from docx import Document

st.set_page_config(page_title="Visit Scheduler", layout="wide")
st.markdown("# 🧬 Visit Scheduler")
st.caption("Choose protocol → set patient/constraints → pick dates → export (Outlook/Excel/Word). No uploads needed.")

with st.container(border=True):
    st.markdown(
        "### ⚠️ Privacy Notice\n"
        "- **Do not enter Protected Health Information (PHI).** Use de-identified IDs only.\n"
        "- Hosting platforms may keep logs/telemetry. For HIPAA workflows, use de-identified inputs and export locally."
    )

# -------------------- constants & helpers --------------------
REQUIRED_COLS = ["Day From Baseline", "Window Minus", "Window Plus"]

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
    for fmt in ["%I:%M %p", "%I:%M%p", "%H:%M", "%H%M"]:
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
    d = date(year + (1 if month == 12 else 0), (month % 12) + 1, 1) - timedelta(days=1)
    while d.weekday() != weekday:
        d -= timedelta(days=1)
    return d

def observed(dt: date):
    # Observed rules: if holiday falls on Sat -> observed Fri; on Sun -> observed Mon
    if dt.weekday() == 5:  # Sat
        return dt - timedelta(days=1)
    if dt.weekday() == 6:  # Sun
        return dt + timedelta(days=1)
    return dt

# ---- Per-holiday date generators (observed) ----
def holiday_dates_for_year(year: int, flags: dict[str, bool]) -> set[date]:
    s = set()
    if flags.get("new_years", True):
        s.add(observed(date(year, 1, 1)))
    if flags.get("mlk", True):
        s.add(nth_weekday_of_month(year, 1, 0, 3))  # 3rd Monday Jan
    if flags.get("presidents", True):
        s.add(nth_weekday_of_month(year, 2, 0, 3))  # 3rd Monday Feb
    if flags.get("memorial", True):
        s.add(last_weekday_of_month(year, 5, 0))    # last Monday May
    if flags.get("juneteenth", True):
        s.add(observed(date(year, 6, 19)))
    if flags.get("independence", True):
        s.add(observed(date(year, 7, 4)))
    if flags.get("labor", True):
        s.add(nth_weekday_of_month(year, 9, 0, 1))  # 1st Monday Sep
    if flags.get("indigenous_columbus", True):
        s.add(nth_weekday_of_month(year, 10, 0, 2)) # 2nd Monday Oct
    if flags.get("veterans", True):
        s.add(observed(date(year, 11, 11)))
    if flags.get("thanksgiving", True):
        s.add(nth_weekday_of_month(year, 11, 3, 4)) # 4th Thursday Nov
    if flags.get("christmas", True):
        s.add(observed(date(year, 12, 25)))
    return s

def build_holiday_set(date_min: date, date_max: date, holiday_flags: dict[str, bool]):
    if date_min is None or date_max is None:
        return set()
    years = set(range(date_min.year, date_max.year + 1))
    out = set()
    for y in years:
        out |= holiday_dates_for_year(y, holiday_flags)
    return out

def nearest_allowed_date(target, earliest, latest, disallow_weekends, holiday_set, custom_blackouts):
    if target is None or earliest is None or latest is None:
        return None
    target = pd.to_datetime(target).date()
    earliest = pd.to_datetime(earliest).date()
    latest = pd.to_datetime(latest).date()

    def allowed(d):
        if d < earliest or d > latest: return False
        if disallow_weekends and d.weekday() >= 5: return False
        if d in holiday_set or d in custom_blackouts: return False
        return True

    if allowed(target):
        return target
    span = (latest - earliest).days
    for k in range(1, span + 1):
        for cand in (target + timedelta(days=k), target - timedelta(days=k)):
            if allowed(cand): return cand
    for cand in (earliest, latest):
        if allowed(cand): return cand
    return None

def window_has_allowed_date(earliest, latest, disallow_weekends, holiday_set, custom_blackouts):
    if earliest is None or latest is None: return False
    d = pd.to_datetime(earliest).date()
    latest = pd.to_datetime(latest).date()
    while d <= latest:
        if (not disallow_weekends or d.weekday() < 5) and (d not in holiday_set) and (d not in custom_blackouts):
            return True
        d += timedelta(days=1)
    return False

def suggest_anchor_dates(anchor, schedule_df, disallow_weekends, holiday_flags, custom_blackouts, search_days=60):
    anchor = _to_date(anchor)
    if anchor is None: return None, None

    min_day = int(schedule_df["Day From Baseline"].min() - schedule_df["Window Minus"].max() - 7)
    max_day = int(schedule_df["Day From Baseline"].max() + schedule_df["Window Plus"].max() + 7)

    def all_visits_feasible(proposed_anchor):
        date_min = proposed_anchor + timedelta(days=min_day)
        date_max = proposed_anchor + timedelta(days=max_day)
        holiday_set = build_holiday_set(date_min, date_max, holiday_flags)
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
        ce, cl = anchor - timedelta(days=delta), anchor + timedelta(days=delta)
        if earlier is None and all_visits_feasible(ce): earlier = ce
        if later   is None and all_visits_feasible(cl): later   = cl
        if earlier and later: break
    return earlier, later

def make_ics(events, cal_name="Visit Schedule"):
    def dtstamp(): return datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")
    lines = ["BEGIN:VCALENDAR","VERSION:2.0","PRODID:-//Visit Scheduler//EN",f"X-WR-CALNAME:{cal_name}"]
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

def duration_minutes_to_label(m):
    m = int(m)
    if m < 60: return f"{m} minutes"
    h, rem = divmod(m, 60)
    if rem == 0: return f"{h} hour" if h == 1 else f"{h} hours"
    return f"{h} hour {rem} minutes" if h == 1 else f"{h} hours {rem} minutes"

def duration_label_to_minutes(label):
    s = str(label).lower().strip()
    if "hour" in s or "minute" in s:
        h = m = 0
        parts = s.replace("minutes","minute").replace("hours","hour").split()
        try:
            if "hour" in parts: h = int(parts[parts.index("hour")-1])
            if "minute" in parts: m = int(parts[parts.index("minute")-1])
            return h*60 + m
        except Exception: pass
    try: return int(float(label))
    except Exception: return None

# durations: 30 min to 12 hours
DUR_MINUTES = list(range(30, 721, 30))
DUR_LABELS  = [duration_minutes_to_label(m) for m in DUR_MINUTES]
LABEL_TO_MIN = {lab: mins for lab, mins in zip(DUR_LABELS, DUR_MINUTES)}

# start times: every 30 minutes, 12h clock with AM/PM
TIME_OPTS = [f"{(h%12) or 12}:{m:02d} {'AM' if h < 12 else 'PM'}" for h in range(24) for m in (0, 30)]

def _format_mmddyyyy(df, cols):
    df = df.copy()
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce").dt.strftime("%m/%d/%Y")
    return df

# -------------------- sidebar settings --------------------
with st.sidebar:
    st.subheader("⚙️ Settings")
    disallow_weekends = st.toggle("Disallow weekends", value=True)

    st.subheader("🏳️‍🌈 Holidays to exclude (observed)")
    # Defaults: ON for all federal holidays; user can toggle each off
    c1, c2 = st.columns(2)
    with c1:
        h_new_years   = st.checkbox("New Year's Day", True)
        h_mlk         = st.checkbox("MLK Day", True)
        h_presidents  = st.checkbox("Presidents' Day", True)
        h_memorial    = st.checkbox("Memorial Day", True)
        h_juneteenth  = st.checkbox("Juneteenth", True)
        h_independence= st.checkbox("Independence Day", True)
    with c2:
        h_labor       = st.checkbox("Labor Day", True)
        h_indig_col   = st.checkbox("Indigenous Peoples’ Day (Columbus)", True)
        h_veterans    = st.checkbox("Veterans Day", True)
        h_thanks      = st.checkbox("Thanksgiving", True)
        h_christmas   = st.checkbox("Christmas Day", True)

    holiday_flags = {
        "new_years": h_new_years,
        "mlk": h_mlk,
        "presidents": h_presidents,
        "memorial": h_memorial,
        "juneteenth": h_juneteenth,
        "independence": h_independence,
        "labor": h_labor,
        "indigenous_columbus": h_indig_col,
        "veterans": h_veterans,
        "thanksgiving": h_thanks,
        "christmas": h_christmas,
    }

    st.caption("Observed rules: if a holiday falls on Sat → observed Fri; on Sun → observed Mon.")
    st.caption("**PHI:** Use de-identified patient IDs only.")

# -------------------- protocol dropdown (version shown only here) --------------------
def read_protocol_version(csv_path: Path):
    try:
        df = pd.read_csv(csv_path, nrows=10)
        df.columns = df.columns.str.strip()

        def norm(s): return "".join(str(s).strip().lower().replace("_", "").split())
        target = norm("Protocol Version")
        col = None
        for c in df.columns:
            if norm(c) == target:
                col = c; break
        if col is None:
            for cand in ["Version", "ProtocolVersion", "Protocol_Version"]:
                for c in df.columns:
                    if norm(c) == norm(cand):
                        col = c; break
                if col: break
        if not col:
            return None

        vals = [str(v).strip() for v in df[col].dropna().tolist() if str(v).strip() and str(v).strip().lower() != "nan"]
        return vals[0] if vals else None
    except Exception:
        return None

def list_protocol_csvs():
    protodir = Path(__file__).parent / "protocols"
    return {
        "IGC": protodir / "IGC.csv",
        "ATAI": protodir / "ATAI.csv",
        "Reunion ADCO": protodir / "Reunion_ADCO.csv",
    }

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
    # optional columns used in app
    if "Start Time" not in df.columns:     df["Start Time"] = pd.Series([None]*len(df), dtype="object")
    if "Visit Duration" not in df.columns: df["Visit Duration"] = pd.Series([None]*len(df), dtype="object")
    if "Protocol Version" not in df.columns: df["Protocol Version"] = ""
    return df

prot_map = list_protocol_csvs()
prot_map = {k:v for k,v in prot_map.items() if v.exists()}
if not prot_map:
    st.error("No protocols found. Ensure protocols/IGC.csv, protocols/ATAI.csv, protocols/Reunion_ADCO.csv exist.")
    st.stop()

label_map = {}
for base, p in prot_map.items():
    ver = read_protocol_version(p)
    label_map[f"{base} — {ver}" if ver else base] = {"base": base, "path": p, "version": ver}

proto_label = st.selectbox("📄 Choose protocol", list(label_map.keys()))
selected = label_map[proto_label]
proto_base = selected["base"]      # keep base for filenames

try:
    schedule = load_protocol(selected["path"])
except Exception as e:
    st.error(f"Error loading protocol: {e}")
    st.stop()

# -------------------- patient inputs --------------------
st.markdown("## 👤 Patient")
c1, c2 = st.columns([1, 1])
with c1:
    st.markdown("**Patient ID — :red[DO NOT ENTER PHI]**")
    patient_id = st.text_input("Patient ID", label_visibility="collapsed")
with c2:
    anchor_date = st.date_input("Anchor Date", value=date.today(), key="anchor_date_input")
notes = st.text_area("Optional Notes (internal)")
ready = bool(patient_id and anchor_date)

# -------------------- blackouts --------------------
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
    range_seed = pd.DataFrame({"Start": pd.Series([], dtype="datetime64[ns]"), "End": pd.Series([], dtype="datetime64[ns]")})
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
                dates.add(d); d += timedelta(days=1)
    return dates

custom_blackouts = set(d.date() for d in pd.to_datetime(single_blackouts_df["Blackout Date"]).dropna().tolist())
custom_blackouts |= expand_ranges(range_blackouts_df)

# -------------------- core compute --------------------
def compute_visits(anchor: date):
    min_day = int(schedule["Day From Baseline"].min() - schedule["Window Minus"].max() - 7)
    max_day = int(schedule["Day From Baseline"].max() + schedule["Window Plus"].max() + 7)
    date_min = anchor + timedelta(days=min_day)
    date_max = anchor + timedelta(days=max_day)
    holiday_set = build_holiday_set(date_min, date_max, holiday_flags)

    out = schedule.copy()
    out["Target Date"] = pd.to_datetime(anchor) + pd.to_timedelta(out["Day From Baseline"], unit="D")
    out["Earliest"]    = out["Target Date"] - pd.to_timedelta(out["Window Minus"], unit="D")
    out["Latest"]      = out["Target Date"] + pd.to_timedelta(out["Window Plus"], unit="D")

    chosen = []
    for _, r in out.iterrows():
        ch = nearest_allowed_date(r["Target Date"].date(), r["Earliest"].date(), r["Latest"].date(),
                                  disallow_weekends, holiday_set, custom_blackouts)
        chosen.append(ch)
    out["Chosen Date"] = chosen

    if "Start Time" not in out.columns:     out["Start Time"] = pd.Series([None]*len(out), dtype="object")
    if "Visit Duration" not in out.columns: out["Visit Duration"] = pd.Series([None]*len(out), dtype="object")
    return out

def _coerce_dates(df, cols):
    for c in cols: df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
    return df

def _annotate(df):
    def row_status(r):
        cd = _to_date(r.get("Chosen Date"))
        e  = _to_date(r.get("Earliest"))
        l  = _to_date(r.get("Latest"))
        if cd is None or e is None or l is None:
            return "⏳ Not set"
        return "✅ In window" if (e <= cd <= l) else "🔴 Out of window"

    df = df.copy()
    df["Status"] = [row_status(r) for _, r in df.iterrows()]
    any_out = (df["Status"] == "🔴 Out of window").any()
    return df, any_out

# -------------------- schedule & adjust --------------------
if ready:
    st.markdown("## 📅 Schedule & Adjust")
    st.info(
        "**How to choose a date:** click **Chosen Date** and pick from the calendar (mm/dd/yyyy). "
        "Set **Start Time** and **Visit Duration** if you want calendar export and handouts. "
        "The **Status** column shows whether a visit is in window."
    )

    # 1) Compute the schedule (auto-picked dates) and coerce date columns
    visits = compute_visits(anchor_date)
    base_table = _coerce_dates(
        visits.copy(),
        ["Target Date", "Earliest", "Latest", "Chosen Date"]
    )

    # 2) Pre-compute Status so it shows *inside* the editor
    seed_for_editor, _ = _annotate(base_table)

    # 3) Show the editor with Status next to Chosen Date
    # Order: Visit, Chosen Date, Status, Start Time, Duration, Target, Earliest, Latest, Day/Windows at end
    column_order = [
        "Visit Name",
        "Chosen Date",
        "Status",
        "Start Time",
        "Visit Duration",
        "Target Date",
        "Earliest",
        "Latest",
        "Day From Baseline",
        "Window Minus",
        "Window Plus",
    ]

    edited = st.data_editor(
        seed_for_editor,
        use_container_width=True,
        height=800,  # tall to minimize scrolling
        hide_index=True,
        num_rows="fixed",
        column_order=[c for c in column_order if c in seed_for_editor.columns],
        column_config={
            "Visit Name":        st.column_config.TextColumn(disabled=True),
            "Day From Baseline": st.column_config.NumberColumn(disabled=True),
            "Window Minus":      st.column_config.NumberColumn(disabled=True),
            "Window Plus":       st.column_config.NumberColumn(disabled=True),
            "Target Date":       st.column_config.DateColumn(disabled=True),
            "Earliest":          st.column_config.DateColumn(disabled=True),
            "Latest":            st.column_config.DateColumn(disabled=True),
            "Chosen Date":       st.column_config.DateColumn(),
            "Status":            st.column_config.TextColumn(disabled=True),
            "Start Time":        st.column_config.SelectboxColumn(options=[None]+TIME_OPTS, required=False),
            "Visit Duration":    st.column_config.SelectboxColumn(options=[None]+DUR_LABELS, required=False),
        },
        key="single_visits_editor"
    )

    # 4) Recompute Status based on the user’s latest edits (for exports & warnings)
    table_with_status, any_out = _annotate(edited)

    # 5) Feasibility check under current blackouts/constraints
    min_day = int(schedule["Day From Baseline"].min() - schedule["Window Minus"].max() - 7)
    max_day = int(schedule["Day From Baseline"].max() + schedule["Window Plus"].max() + 7)
    date_min = anchor_date + timedelta(days=min_day)
    date_max = anchor_date + timedelta(days=max_day)
    holiday_set_now = build_holiday_set(date_min, date_max, holiday_flags)
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
        st.warning("Some visits are out of window. Please adjust **Chosen Date** to fall between **Earliest** and **Latest**.")
    if not windows_ok:
        earlier_sug, later_sug = suggest_anchor_dates(
            anchor_date, schedule, disallow_weekends, holiday_flags, custom_blackouts
        )
        with st.container(border=True):
            st.error("Blackouts/constraints remove all allowed days for at least one visit window.")
            msg = "Suggested anchor dates:"
            if earlier_sug: msg += f" **Earlier:** {earlier_sug.strftime('%m/%d/%Y')}"
            if later_sug:   msg += f" | **Later:** {later_sug.strftime('%m/%d/%Y')}"
            st.write(msg)
            c1, c2 = st.columns(2)
            if earlier_sug and c1.button(f"Apply earlier anchor ({earlier_sug.strftime('%m/%d/%Y')})"):
                st.session_state["anchor_date_input"] = earlier_sug
                st.rerun()
            if later_sug and c2.button(f"Apply later anchor ({later_sug.strftime('%m/%d/%Y')})"):
                st.session_state["anchor_date_input"] = later_sug
                st.rerun()

    # 6) Store for Export & Print
    st.session_state["single_result"] = {
        "patient_id": patient_id,
        "anchor_date": anchor_date,
        "df": table_with_status.copy()
    }

# -------------------- export & print --------------------
st.markdown("## 🖨️ Export & Print")
role = st.radio("Role", ["Coordinator view", "Participant handout"], horizontal=True)

def coordinator_view(df):
    # Order: Visit Name, Chosen Date, Start Time, Visit Duration, Status, Target, Earliest, Latest, Window-, Window+
    cols = [
        "Visit Name",
        "Chosen Date",
        "Start Time",
        "Visit Duration",
        "Status",
        "Target Date",
        "Earliest",
        "Latest",
        "Window Minus",
        "Window Plus",
    ]
    out = df[[c for c in cols if c in df.columns]].copy()
    out = _format_mmddyyyy(out, ["Target Date","Earliest","Latest","Chosen Date"])
    return out

def participant_view(df):
    # Start with Chosen Date, then rename to Visit Date for user-facing outputs
    out = df.copy()[["Visit Name","Chosen Date","Start Time","Visit Duration"]].copy()
    out = _format_mmddyyyy(out, ["Chosen Date"])
    # compute Expected End Time if start time & duration are set
    end_times = []
    for _, r in out.iterrows():
        cd = _to_date(r.get("Chosen Date")); stime = r.get("Start Time"); dlabel = r.get("Visit Duration")
        if cd and stime and dlabel:
            start_dt = datetime.combine(cd, _to_time(stime) or dtime(9,0))
            dur_min = LABEL_TO_MIN.get(dlabel, duration_label_to_minutes(dlabel))
            if dur_min:
                end_times.append((start_dt + timedelta(minutes=int(dur_min))).strftime("%I:%M %p"))
            else:
                end_times.append("")
        else:
            end_times.append("")
    out["Expected End Time"] = end_times
    # Rename for display/export
    out.rename(columns={"Chosen Date": "Visit Date"}, inplace=True)
    out = out[["Visit Name","Visit Date","Start Time","Expected End Time","Visit Duration"]]
    return out

left, right = st.columns([2, 1])
with left:
    if "single_result" in st.session_state and ready:
        res = st.session_state["single_result"]; df = res["df"].copy()
        table = coordinator_view(df) if role == "Coordinator view" else participant_view(df)
        st.dataframe(table, use_container_width=True, height=700)

        # CSV & Excel
        fname_root = f"{res['patient_id']}_{'coordinator' if role=='Coordinator view' else 'participant'}"
        st.download_button("⬇️ Download CSV", data=table.to_csv(index=False).encode("utf-8"),
                           file_name=f"{fname_root}.csv", mime="text/csv")
        xbuf = io.BytesIO()
        with pd.ExcelWriter(xbuf, engine="xlsxwriter") as writer:
            sheet = "Schedule" if role == "Coordinator view" else "Participant Handout"
            if role != "Coordinator view":
                # add disclaimer sheet
                pd.DataFrame({"Note":["Times are estimates. Actual end time may vary."]}).to_excel(writer, index=False, sheet_name="Disclaimer")
            table.to_excel(writer, index=False, sheet_name=sheet)
        st.download_button("⬇️ Download Excel", data=xbuf.getvalue(),
                           file_name=f"{fname_root}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        # Word export (participant only) — uses "Visit Date" label
        if role == "Participant handout":
            doc = Document()
            doc.add_heading('Participant Visit Handout', level=1)
            doc.add_paragraph("Times are estimates and may vary.").runs[0].italic = True

            headers = ["Visit Name", "Visit Date", "Start Time", "Expected End Time", "Visit Duration"]
            t = doc.add_table(rows=1, cols=len(headers))
            t.style = "Table Grid"
            hdr_cells = t.rows[0].cells
            for i, h in enumerate(headers):
                hdr_cells[i].text = h
            for _, r in table.iterrows():
                row = t.add_row().cells
                row[0].text = str(r.get("Visit Name",""))
                row[1].text = str(r.get("Visit Date",""))
                row[2].text = str(r.get("Start Time",""))
                row[3].text = str(r.get("Expected End Time",""))
                row[4].text = str(r.get("Visit Duration",""))

            bio = io.BytesIO()
            doc.save(bio)
            st.download_button(
                "⬇️ Download Word",
                data=bio.getvalue(),
                file_name=f"{fname_root}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )

        # Outlook ICS (Coordinator view only). (No version in description per your preference.)
        if role == "Coordinator view":
            events = []
            for _, r in df.iterrows():
                cd = _to_date(r.get("Chosen Date"))
                if not cd: continue
                start_t = _to_time(r.get("Start Time"))
                dlabel = r.get("Visit Duration")
                dur = LABEL_TO_MIN.get(dlabel, duration_label_to_minutes(dlabel)) if dlabel else None

                window_txt = f"Window {r.get('Earliest')} to {r.get('Latest')}"
                desc = f"{proto_base}\\n{window_txt}"  # version intentionally not included
                summary = f"{res['patient_id']} · {r.get('Visit Name','Visit')}"
                events.append({"summary": summary, "date": cd, "start_time": start_t, "duration_min": dur, "description": desc})
            if events:
                ics = make_ics(events, cal_name=f"{proto_base} - {res['patient_id']}")
                st.download_button("📅 Export to Outlook calendar", data=ics,
                                   file_name=f"{res['patient_id']}_schedule.ics", mime="text/calendar")
            else:
                st.info("Set at least one **Chosen Date** to enable calendar export.")

with right:
    st.markdown("### 🧾 Tips")
    st.write("- Toggle specific holidays off if your site is open that day.")
    st.write("- Add **Start Time** and **Visit Duration** to include times in Outlook and participant handouts.")
    st.write("- **Blackouts**: add single days or ranges; anchor suggestions appear if a window becomes impossible.")
    st.write("- All dates display as **mm/dd/yyyy**.")
