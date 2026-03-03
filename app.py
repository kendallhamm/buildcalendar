import streamlit as st
import csv
import uuid
import re
from datetime import datetime
from zoneinfo import available_timezones, ZoneInfo
from io import StringIO

st.set_page_config(page_title="CSV to ICS Generator", layout="wide")

TIMEZONE_DEFAULT = "America/New_York"

# ---------------------------------
# Header Normalization
# ---------------------------------

def normalize_header(name):
    if name is None:
        return ""
    name = str(name).strip().lower()
    name = re.sub(r"\s+", " ", name)
    return name

HEADER_ALIASES = {
    "date": ["date"],
    "start": ["start", "start time", "begin", "begin time"],
    "end": ["end", "end time", "stop"],
    "location": ["location", "room", "venue"],
    "activity": ["activity", "event", "event title", "title"],
    "poc": ["poc", "point of contact", "contact"],
    "comments": ["comments", "notes", "description", "details"]
}

def map_columns(headers):
    normalized = {normalize_header(h): h for h in headers}
    mapped = {}

    for standard_name, aliases in HEADER_ALIASES.items():
        for alias in aliases:
            if alias in normalized:
                mapped[standard_name] = normalized[alias]
                break

    missing = [k for k in HEADER_ALIASES if k not in mapped]
    return mapped, missing

# ---------------------------------
# Timezone Block
# ---------------------------------

def build_vtimezone(tzid):
    return f"""BEGIN:VTIMEZONE
TZID:{tzid}
END:VTIMEZONE
"""

# ---------------------------------
# Parsing Helpers
# ---------------------------------

def parse_date(value):
    if not value:
        return None

    value = value.strip()

    for fmt in ["%d-%b-%y", "%m/%d/%Y", "%Y-%m-%d"]:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue

    return None

def parse_time(value):
    if not value:
        return None

    value = value.strip()

    if ":" in value:
        return datetime.strptime(value, "%H:%M").strftime("%H%M")

    return value.zfill(4)

def format_dt(date_obj, time_str):
    dt = datetime.strptime(
        f"{date_obj.strftime('%Y-%m-%d')} {time_str}",
        "%Y-%m-%d %H%M"
    )
    return dt.strftime("%Y%m%dT%H%M%S")

# ---------------------------------
# UI
# ---------------------------------

st.title("CSV to ICS Calendar Generator")

st.write("Upload a CSV file formatted with columns: Date, Start, End, Location, Activity, POC, Comments.")

uploaded_file = st.file_uploader("Upload CSV", type=["csv"])

tz_list = sorted(available_timezones())
timezone = st.selectbox(
    "Select Timezone",
    tz_list,
    index=tz_list.index(TIMEZONE_DEFAULT) if TIMEZONE_DEFAULT in tz_list else 0
)

generate_button = st.button("Generate ICS")

# ---------------------------------
# Processing
# ---------------------------------

if uploaded_file and generate_button:

    text_data = uploaded_file.read().decode("utf-8-sig")
    reader = csv.DictReader(StringIO(text_data))

    column_map, missing = map_columns(reader.fieldnames)

    if missing:
        st.error(f"Missing required columns: {missing}")
        st.stop()

    events = []
    current_date = None

    for row in reader:

        raw_date = row[column_map["date"]]

        if raw_date.strip():
            current_date = parse_date(raw_date)

        if current_date is None:
            continue

        start = parse_time(row[column_map["start"]])
        end = parse_time(row[column_map["end"]])

        if not start or not end:
            continue

        events.append({
            "date": current_date,
            "start": start,
            "end": end,
            "location": row[column_map["location"]].strip(),
            "activity": row[column_map["activity"]].strip(),
            "poc": row[column_map["poc"]].strip(),
            "comments": row[column_map["comments"]].strip(),
        })

    if not events:
        st.warning("No valid events found.")
        st.stop()

    # Build ICS
    ics = "BEGIN:VCALENDAR\r\n"
    ics += "VERSION:2.0\r\n"
    ics += "CALSCALE:GREGORIAN\r\n"
    ics += "PRODID:-//Streamlit CSV to ICS//EN\r\n"
    ics += build_vtimezone(timezone)

    now_stamp = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    for e in events:
        uid = str(uuid.uuid4())
        description = f"POC: {e['poc']}\\nComments: {e['comments']}"

        ics += "BEGIN:VEVENT\r\n"
        ics += f"UID:{uid}\r\n"
        ics += f"DTSTAMP:{now_stamp}\r\n"
        ics += f"DTSTART;TZID={timezone}:{format_dt(e['date'], e['start'])}\r\n"
        ics += f"DTEND;TZID={timezone}:{format_dt(e['date'], e['end'])}\r\n"
        ics += "TRANSP:OPAQUE\r\n"
        ics += f"SUMMARY:{e['activity']}\r\n"

        if e["location"]:
            ics += f"LOCATION:{e['location']}\r\n"

        ics += f"DESCRIPTION:{description}\r\n"
        ics += "END:VEVENT\r\n"

    ics += "END:VCALENDAR\r\n"

    st.success(f"Generated {len(events)} events.")

    st.download_button(
        label="Download ICS File",
        data=ics,
        file_name="calendar_output.ics",
        mime="text/calendar"
    )
