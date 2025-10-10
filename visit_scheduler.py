# visit_scheduler.py — fixed labels (IGC / ATAI / Reunion ADCO), no duration column
import streamlit as st
import pandas as pd
from datetime import date, datetime, timedelta, time as dtime
from pathlib import Path
import io
import zipfile

# ----------------------
# Page config & header
# ----------------------
st.set_page_config(page_title="Visit Scheduler", layout="wide")
st.markdown("# 🧬 Visit Scheduler")
st.caption("Choose a protocol (IGC / ATAI / Reunion ADCO), add patient(s), apply constraints, pick dates, then export. No file uploads needed.")

# ----------------------
# Required columns
# ----------------------
REQUIRED_COLS = ["Day From Baseline", "Window Minus", "Window Plus"]

# ----------------------
# Helpers
# ----------------------
def _to_date(x):
    if pd.isna(x) or x == "":
        return None
    if isinstance(x, (date, datetime)):
        return pd.to_datetime(x).date()
    try:
        return pd.to_datetime(x).date()
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
    hol.add(observed(date(year, 1, 1)))                          # New Year’s Day
    hol.add(nth_weekday_of_month(year, 1, 0, 3))                 # MLK Day (3rd Mon Jan)
    hol.add(nth_weekday_of_month(year, 2, 0, 3))                 # Washington’s Birthday
    hol.add(last_weekday_of_month(year, 5, 0))                   # Memorial Day (last Mon May)
    hol.add(observed(date(year, 6, 19)))                         # Juneteenth
    hol.add(observed(date(year, 7, 4)))                          # Independence Day
    hol.add(nth_weekday_of_month(year, 9, 0, 1))                 # Labor Day
    hol.add(nth_weekday_of_month(year, 10, 0, 2))                # Columbus Day
    hol.add(observed(date(year, 11, 11)))                        # Veterans Day
    hol.add(nth_weekday_of_month(year, 11, 3, 4))                # Thanksgiving (4th Thu)
    hol.add(observed(date(year, 12, 25)))                        # Christmas
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
    """
    Find the nearest allowed date to 'target' within [earliest, latest]
    avoiding weekends/holidays/blackouts. Search order: 0, +1, -1, +2, -2, ...
    """
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
    # fall back to boundaries if possible
    for cand in [earliest, latest]:
        if (not disallow_weekends or cand.weekday() < 5) and cand not in holiday_set and cand not in custom_blackouts:
            return cand
    return None

def make_ics(events, cal_name="Visit Schedule"):
    """
    events: list of dicts with keys:
        - summary (str)
        - date (date)
        - description (str, optional)
        - duration_min (int, optional; default 60 if omitted)
    """
    def dtstamp():
        return datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//Visit Scheduler//EN",
        f"X-WR-CALNAME:{cal_name}",
    ]
    for ev in events:
        start_dt = datetime.combine(ev["date"], dtime(hour=9, minute=0))  # default 9:00 AM
        dur_min = int(ev.get("duration_min") or 60)
        end_dt = start_dt + timedelta(minutes=dur_min)
        lines += [
            "BEGIN:VEVENT",
            f"UID:{hash((ev['summary'], ev['date']))}@visitscheduler",
            f"DTSTAMP:{dtstamp()}",
            f"DTSTART:{start_dt.strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{end_dt.strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{ev['summary']}",
        ]
        if ev.get("description"):
            desc = ev["description"].replace(",", r"\,").replace(";", r"\;")
            lines.append(f"DESCRIPTION:{desc}")
        lines.append("END:VEVENT")
    lines.append("END:VCALENDAR")
    return "\r\n".join(lines).encode("utf-8")

# ----------------------
# Sidebar: global toggles
# ----------------------
with st.sidebar:
    st.subheader("⚙️ Settings")
    disallow_weekends = st.toggle("Disallow weekends", value=True)
    include_us_holidays = st.toggle("Exclude US Federal Holidays", value=True)
    st.caption("Holiday dates are observed when they fall on a weekend.")

# ----------------------
# Protocol loader (fixed labels)
# ----------------------
def list_protocol_csvs():
    protodir = Path(__file__).parent / "protocols"
    mapping = {
        "IGC": protodir / "IGC.csv",
        "ATAI": protodir / "ATAI.csv",
        "Reunion ADCO": protodir / "Reunion_ADCO.csv",
    }
    # Only include files that actually exist
    existing = {label: path for label, path in mapping.items() if path.exists()}
    return existing

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
    return df

protocols = list_protocol_csvs()
if not protocols:
    st.error("No protocols found. Ensure these files exist: "
             "`protocols/IGC.csv`, `protocols/ATAI.csv`, `protocols/Reunion_ADCO.csv` (at least one).")
    st.stop()

proto_name = st.selectbox("📄 Choose protocol", list(protocols.keys()))
try:
    schedule = load_protocol(protocols[proto_name])
except Exception as e:
    st.error(f"Error loading protocol: {e}")
    st.stop()

# ----------------------
# Patients (single or batch)
# ----------------------
st.markdown("## 👤 Patients")
mode = st.radio("Mode", options=["Single", "Batch"], horizontal=True)

if mode == "Single":
    c1, c2 = st.columns([1, 1])
    with c1:
        patient_id = st.text_input("Patient ID")
    with c2:
        anchor_date = st.date_input("Anchor Date", value=date.today())
    notes = st.text_area("Optional Notes (internal)")
    ready = bool(patient_id and anchor_date)
else:
    st.caption("Enter multiple patients and anchor dates. Add or delete rows as needed.")
    default_rows = [{"Patient ID": "", "Anchor Date": date.today()}]
    batch_df = st.data_editor(
        pd.DataFrame(default_rows),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Patient ID": st.column_config.TextColumn(required=True, help="De-identified or site ID"),
            "Anchor Date": st.column_config.DateColumn(required=True),
        },
        key="batch_editor"
    )
    batch_df["Patient ID"] = batch_df["Patient ID"].astype(str).str.strip()
    batch_df["Anchor Date"] = batch_df["Anchor Date"].apply(_to_date)
    batch_df = batch_df.dropna(subset=["Anchor Date"])
    batch_df = batch_df[batch_df["Patient ID"] != ""]
    ready = len(batch_df) > 0

# ----------------------
# Blackouts & constraints (FIXED: pre-typed datetime column for editor)
# ----------------------
st.markdown("## 🚫 Blackouts & Constraints")
cA, cB = st.columns([1, 1])

with cA:
    st.markdown("**Custom blackout dates**")
    st.caption("Add dates the site or participant cannot attend.")

    # Seed with a real datetime64[ns] dtype so DateColumn works even when empty
    blackout_seed = pd.DataFrame({
        "Blackout Date": pd.Series([], dtype="datetime64[ns]")
    })

    blackout_df = st.data_editor(
        blackout_seed,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Blackout Date": st.column_config.DateColumn()
        },
        key="blackouts_editor"
    )

    # Clean NA rows
    blackout_df = blackout_df.dropna(subset=["Blackout Date"])

with cB:
    st.markdown("**Notes**")
    st.caption("Optional scheduling context.")
    if mode == "Single":
        st.session_state["notes_text"] = notes
    else:
        st.info("Use per-patient notes outside the app if needed (batch mode).")

# Build blackout set safely (dates only)
custom_blackouts = set(
    d.date() for d in pd.to_datetime(blackout_df["Blackout Date"]).dropna().tolist()
)

# ----------------------
# Core compute
# ----------------------
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
    return out

# ----------------------
# Schedule & adjust
# ----------------------
if ready:
    st.markdown("## 📅 Schedule & Adjust")

    if mode == "Single":
        visits = compute_visits_for_patient(anchor_date)
        table = visits.copy()
        for c in ["Target Date", "Earliest", "Latest", "Chosen Date"]:
            table[c] = table[c].dt.date
        st.caption("Edit **Chosen Date**. Suggested dates avoid weekends/US holidays/blackouts when possible; overrides are allowed.")
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
                "Chosen Date": st.column_config.DateColumn(help="Pick the actual appointment date (ideally within window)"),
            },
            key="single_visits_editor"
        )
        st.session_state["single_result"] = {
            "patient_id": patient_id,
            "anchor_date": anchor_date,
            "notes": (st.session_state.get("notes_text") or ""),
            "df": table.copy()
        }

    else:
        rows = []
        for _, r in batch_df.iterrows():
            pid = str(r["Patient ID"]).strip()
            ad = r["Anchor Date"]
            v = compute_visits_for_patient(ad).copy()
            for c in ["Target Date", "Earliest", "Latest", "Chosen Date"]:
                v[c] = v[c].dt.date
            v["Patient ID"] = pid
            v["Anchor Date"] = ad
            rows.append(v)
        batch_table = pd.concat(rows, ignore_index=True) if rows else pd.DataFrame()
        st.caption("You can sort/filter and edit **Chosen Date** per visit.")
        batch_table = st.data_editor(
            batch_table,
            use_container_width=True,
            num_rows="fixed",
            column_config={
                "Patient ID": st.column_config.TextColumn(),
                "Anchor Date": st.column_config.DateColumn(),
                "Visit Name": st.column_config.TextColumn(),
                "Day From Baseline": st.column_config.NumberColumn(),
                "Window Minus": st.column_config.NumberColumn(),
                "Window Plus": st.column_config.NumberColumn(),
                "Target Date": st.column_config.DateColumn(disabled=True),
                "Earliest": st.column_config.DateColumn(disabled=True),
                "Latest": st.column_config.DateColumn(disabled=True),
                "Chosen Date": st.column_config.DateColumn(),
            },
            key="batch_visits_editor"
        )
        st.session_state["batch_result"] = batch_table.copy()

# ----------------------
# Export & print
# ----------------------
st.markdown("## 🖨️ Export & Print")
role = st.radio("Role", ["Coordinator view", "Participant handout"], horizontal=True)

def coordinator_view(df, include_patient=True):
    cols = ["Visit Name", "Day From Baseline", "Target Date", "Earliest", "Latest", "Chosen Date", "Window Minus", "Window Plus"]
    out = df.copy()
    if include_patient and "Patient ID" in out.columns:
        cols = ["Patient ID", "Anchor Date"] + cols
    return out[[c for c in cols if c in out.columns]]

def participant_view(df, include_patient=False):
    cols = ["Visit Name", "Chosen Date", "Earliest", "Latest"]
    out = df.copy()
    if include_patient and "Patient ID" in out.columns:
        cols = ["Patient ID"] + cols
    return out[[c for c in cols if c in out.columns]]

left, right = st.columns([2, 1])

with left:
    if mode == "Single" and "single_result" in st.session_state:
        res = st.session_state["single_result"]
        df = res["df"].copy()
        table = coordinator_view(df, include_patient=False) if role == "Coordinator view" else participant_view(df, include_patient=False)
        st.dataframe(table, use_container_width=True)

        # CSV
        csv_bytes = table.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Download CSV",
            data=csv_bytes,
            file_name=f"{res['patient_id']}_{role.replace(' ','_')}.csv",
            mime="text/csv"
        )
        # Excel
        xbuf = io.BytesIO()
        with pd.ExcelWriter(xbuf, engine="xlsxwriter") as writer:
            table.to_excel(writer, index=False, sheet_name="Schedule")
        st.download_button(
            "⬇️ Download Excel",
            data=xbuf.getvalue(),
            file_name=f"{res['patient_id']}_{role.replace(' ','_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # ICS from Chosen Dates
        events = []
        for _, r in df.iterrows():
            cd = _to_date(r.get("Chosen Date"))
            if not cd:
                continue
            summary = f"{r.get('Visit Name','Visit')}"
            if role == "Coordinator view":
                summary = f"{res['patient_id']} · {summary}"
            desc = f"{proto_name} — Window {r.get('Earliest')} to {r.get('Latest')}"
            # duration hidden from UI; default to 60 minutes internally
            events.append({"summary": summary, "date": cd, "description": desc})
        if events:
            ics_data = make_ics(events, cal_name=f"{proto_name} - {res['patient_id']}")
            st.download_button("📅 Download .ics (Chosen Dates)", data=ics_data, file_name=f"{res['patient_id']}_schedule.ics", mime="text/calendar")
        else:
            st.info("Set at least one **Chosen Date** to enable calendar export.")

    elif mode == "Batch" and "batch_result" in st.session_state:
        df = st.session_state["batch_result"].copy()
        if df.empty:
            st.info("Add patients in the table above.")
        else:
            table = coordinator_view(df, include_patient=True) if role == "Coordinator view" else participant_view(df, include_patient=True)
            st.dataframe(table, use_container_width=True)

            # CSV / Excel
            csv_bytes = table.to_csv(index=False).encode("utf-8")
            st.download_button("⬇️ Download CSV (All Patients)", data=csv_bytes, file_name=f"batch_{role.replace(' ','_')}.csv", mime="text/csv")

            xbuf = io.BytesIO()
            with pd.ExcelWriter(xbuf, engine="xlsxwriter") as writer:
                table.to_excel(writer, index=False, sheet_name="Schedules")
            st.download_button("⬇️ Download Excel (All Patients)", data=xbuf.getvalue(), file_name=f"batch_{role.replace(' ','_')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # ZIP of per-patient ICS
            patients = sorted(df["Patient ID"].unique())
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.Zip_DEFLATED) as zf:
                for pid in patients:
                    sub = df[df["Patient ID"] == pid].copy()
                    events = []
                    for _, r in sub.iterrows():
                        cd = _to_date(r.get("Chosen Date"))
                        if not cd:
                            continue
                        summary = f"{r.get('Visit Name','Visit')}"
                        if role == "Coordinator view":
                            summary = f"{pid} · {summary}"
                        desc = f"{proto_name} — Window {r.get('Earliest')} to {r.get('Latest')}"
                        events.append({"summary": summary, "date": cd, "description": desc})
                    if events:
                        ics_bytes = make_ics(events, cal_name=f"{proto_name} - {pid}")
                        zf.writestr(f"{pid}_schedule.ics", ics_bytes)
            if zip_buf.getbuffer().nbytes > 0:
                st.download_button("📦 Download .ics (ZIP per patient)", data=zip_buf.getvalue(), file_name="batch_schedules_ics.zip", mime="application/zip")
            else:
                st.info("Set **Chosen Date** for at least one row to enable calendar ZIP export.")

with right:
    st.markdown("### 🧾 Print Tips")
    st.write("- Use **Participant handout** to hide internal fields.")
    st.write("- Then use your browser’s **Print** (Ctrl/Cmd + P).")
    st.write("- Calendar export uses your **Chosen Dates**.")

    st.write("- Calendar export uses your **Chosen Dates**.")

