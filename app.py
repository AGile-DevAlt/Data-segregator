import re
from io import BytesIO
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Event Participant Segregator", page_icon="📋", layout="wide")

EVENT_GROUPS = {
    "Yuva Sansad": "DAY 1 & 2",
    "The mentalist": "DAY 1 & 2",
    "Drishti": "DAY 1 & 2",
    "Kalakriti": "DAY 1 ONLY",
    "Natyalaya": "DAY 1 ONLY",
    "JAM": "DAY 1 ONLY",
    "Pitch Perfect": "DAY 2 ONLY",
    "Reel making": "DAY 2 ONLY",
    "Nukkad Natak": "DAY 2 ONLY",
    "Behaviour Scientist": "DAY 2 ONLY",
}

OUTPUT_COLUMNS = [
    "Full name", "Email", "Grade", "Section", "Contact Number",
    "Admission Number", "Events (DAY 1 & 2)", "Events (Day 1 ONLY)",
    "Events (Day 2 ONLY)"
]

ALIASES = {
    "Full name": ["full name", "name", "participant", "student name", "student"],
    "Email": ["email", "email address", "e-mail", "mail id"],
    "Grade": ["grade", "class", "year"],
    "Section": ["section", "sec"],
    "Contact Number": ["contact number", "phone", "mobile", "contact", "phone number", "mobile number"],
    "Admission Number": ["admission number", "admission no", "admission no.", "adm no", "admission id", "roll number", "roll no"],
}

def norm(value):
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()

def normalize_header(value):
    return re.sub(r"[^a-z0-9]+", " ", norm(value).lower()).strip()

def find_columns(df):
    found = {}
    normalized = {col: normalize_header(col) for col in df.columns}
    for target, aliases in ALIASES.items():
        for col, h in normalized.items():
            if any(h == normalize_header(alias) or normalize_header(alias) in h for alias in aliases):
                found[target] = col
                break
    return found

def events_in_row(row):
    cells = [norm(v) for v in row.tolist() if norm(v)]
    haystack = " | ".join(cells).lower()
    found = []
    for event in EVENT_GROUPS:
        # Match event names case-insensitively; permit harmless spaces around words.
        pattern = re.escape(event).replace(r"\ ", r"\s+")
        if re.search(r"(?<![a-z0-9])" + pattern + r"(?![a-z0-9])", haystack, re.I):
            found.append(event)
    return found

def build_output(df, mapping, selected_event_columns=None):
    selected_event_columns = selected_event_columns or []
    # Search every cell by default, including columns not recognized as participant details.
    records = []
    for _, row in df.iterrows():
        event_hits = events_in_row(row)
        # If the user picked specific event columns, only use those columns to detect events.
        if selected_event_columns:
            event_text = " | ".join(norm(row[c]) for c in selected_event_columns if c in df.columns)
            event_hits = []
            for event in EVENT_GROUPS:
                pattern = re.escape(event).replace(r"\ ", r"\s+")
                if re.search(r"(?<![a-z0-9])" + pattern + r"(?![a-z0-9])", event_text, re.I):
                    event_hits.append(event)
        if not event_hits:
            continue
        record = {}
        for out_col in OUTPUT_COLUMNS[:6]:
            source_col = mapping.get(out_col)
            record[out_col] = norm(row[source_col]) if source_col in df.columns else ""
        # Avoid blank/non-participant rows such as title rows that mention an event.
        if not record.get("Full name"):
            continue
        for group in ("DAY 1 & 2", "DAY 1 ONLY", "DAY 2 ONLY"):
            record_col = {
                "DAY 1 & 2": "Events (DAY 1 & 2)",
                "DAY 1 ONLY": "Events (Day 1 ONLY)",
                "DAY 2 ONLY": "Events (Day 2 ONLY)",
            }[group]
            record[record_col] = ", ".join(e for e in event_hits if EVENT_GROUPS[e] == group)
        records.append((record, event_hits))

    # Remove exact duplicate participant/event rows while preserving the first occurrence.
    event_frames = {event: [] for event in EVENT_GROUPS}
    all_records = []
    seen = set()
    for record, hits in records:
        for event in hits:
            key = (event, tuple(record.get(c, "") for c in OUTPUT_COLUMNS[:6]))
            if key not in seen:
                event_frames[event].append({**record, "Event": event, "Event Day": EVENT_GROUPS[event]})
                seen.add(key)
        all_records.append(record)

    event_frames = {k: pd.DataFrame(v) for k, v in event_frames.items()}
    day_frames = {}
    for group, out_col in [
        ("DAY 1 & 2", "Events (DAY 1 & 2)"),
        ("DAY 1 ONLY", "Events (Day 1 ONLY)"),
        ("DAY 2 ONLY", "Events (Day 2 ONLY)"),
    ]:
        rows = []
        for event in EVENT_GROUPS:
            if EVENT_GROUPS[event] == group:
                rows.extend(event_frames[event].to_dict("records"))
        day_frames[group] = pd.DataFrame(rows)
    return event_frames, day_frames

def to_excel(event_frames, day_frames):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for event, frame in event_frames.items():
            sheet = re.sub(r"[\[\]\*:/\\?]", "-", event)[:31]
            if frame.empty:
                frame = pd.DataFrame(columns=OUTPUT_COLUMNS + ["Event", "Event Day"])
            frame.to_excel(writer, sheet_name=sheet, index=False)
        for group, frame in day_frames.items():
            sheet = {"DAY 1 & 2": "Day 1 and 2", "DAY 1 ONLY": "Day 1 Only", "DAY 2 ONLY": "Day 2 Only"}[group]
            if frame.empty:
                frame = pd.DataFrame(columns=OUTPUT_COLUMNS + ["Event", "Event Day"])
            frame.to_excel(writer, sheet_name=sheet, index=False)
    output.seek(0)
    return output.getvalue()

st.title("📋 Event Participant Segregator")
st.write("Upload a mixed Excel list. The app detects event names, places participants into each relevant event sheet, and includes common participants in every event they are listed for.")

st.info("Event groups: **Day 1 & 2** — Yuva Sansad, The mentalist, Drishti · **Day 1 only** — Kalakriti, Natyalaya, JAM · **Day 2 only** — Pitch Perfect, Reel making, Nukkad Natak, Behaviour Scientist")

uploaded = st.file_uploader("Upload Excel file (.xlsx or .xls)", type=["xlsx", "xls"])
if uploaded:
    try:
        workbook = pd.ExcelFile(uploaded)
        sheet_name = st.selectbox("Choose the sheet containing the mixed data", workbook.sheet_names)
        df = pd.read_excel(uploaded, sheet_name=sheet_name, dtype=str).fillna("")
        if df.empty:
            st.warning("This sheet has no rows.")
            st.stop()

        st.subheader("Preview")
        st.dataframe(df.head(10), use_container_width=True)

        detected = find_columns(df)
        st.subheader("Match participant columns")
        st.caption("We try to identify these automatically. Check the selections and correct any that are wrong.")
        mapping = {}
        cols = ["(not present)"] + list(df.columns)
        grid = st.columns(3)
        for i, target in enumerate(OUTPUT_COLUMNS[:6]):
            default = detected.get(target, "(not present)")
            index = cols.index(default) if default in cols else 0
            with grid[i % 3]:
                selected = st.selectbox(target, cols, index=index, key=f"map_{target}")
            mapping[target] = None if selected == "(not present)" else selected

        st.subheader("Event detection")
        st.write("By default, the app searches **all cells in each row** for event names. If your sheet has a dedicated event column, you can select it to avoid accidental matches in other fields.")
        event_cols = st.multiselect("Optional: select column(s) that contain event names", list(df.columns), default=[])

        if st.button("Segregate participants", type="primary"):
            event_frames, day_frames = build_output(df, mapping, event_cols)
            total = sum(len(frame) for frame in event_frames.values())
            if total == 0:
                st.error("No participants were matched. Check the event names in the file and confirm the Full name column is mapped correctly.")
            else:
                st.success(f"Done! Created {len(EVENT_GROUPS)} event sheets and 3 day-group sheets.")
                summary = pd.DataFrame([
                    {"Event": event, "Day group": EVENT_GROUPS[event], "Participant rows": len(event_frames[event])}
                    for event in EVENT_GROUPS
                ])
                st.dataframe(summary, use_container_width=True, hide_index=True)
                st.download_button(
                    "Download segregated Excel workbook",
                    data=to_excel(event_frames, day_frames),
                    file_name="segregated_event_participants.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.caption("Each participant is included in every event detected on their source row. Exact duplicate rows within the same event are removed. Rows without a detected event or without a full name are skipped.")
    except Exception as exc:
        st.error(f"Could not read this workbook: {exc}")
        st.write("Try saving the source file as .xlsx and uploading it again.")
else:
    st.markdown("### How to use")
    st.markdown("1. Install Python 3.10 or newer.\n2. Unzip this package.\n3. Run `pip install -r requirements.txt`.\n4. Run `streamlit run app.py`.\n5. Upload your Excel file, verify column matches, and click **Segregate participants**.")
