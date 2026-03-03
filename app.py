import streamlit as st
import csv
import uuid
import re
from datetime import datetime
from zoneinfo import available_timezones, ZoneInfo
from io import StringIO

st.set_page_config(page_title="CSV to ICS Generator", layout="wide")

TIMEZONE_DEFAULT = "America/New_York"

# Analytics
GOOGLE_APP_URL = "https://script.google.com/macros/s/AKfycbyAaJudgxvFC1qh1GyypKgeGjbO-6iHVgABZa0_svbadwgB7YJOFH5Wc5hH_Y98ZKxqyw/exec"
def log_schedule_generated():
    try:
        requests.post(GOOGLE_APP_URL, timeout=5)
    except Exception:
        # Fail silently: logging should never break scheduling
        pass

# Refresh of 'total schedules generated' value function definition
@st.cache_data(ttl=300, show_spinner=False)  # refresh every 5 minutes
def get_total_schedules_generated():
    try:
        r = requests.get(GOOGLE_APP_URL, timeout=15) # timeout at 15 seconds. 5 was too short for stale script.
        r.raise_for_status()
        return int(r.text)
    except Exception:
        # Analytics should never break the app
        return None


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

    US_TZ_BLOCKS = {

        "America/New_York": """BEGIN:VTIMEZONE
TZID:America/New_York
BEGIN:DAYLIGHT
TZOFFSETFROM:-0500
TZOFFSETTO:-0400
TZNAME:EDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0400
TZOFFSETTO:-0500
TZNAME:EST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE
""",

        "America/Chicago": """BEGIN:VTIMEZONE
TZID:America/Chicago
BEGIN:DAYLIGHT
TZOFFSETFROM:-0600
TZOFFSETTO:-0500
TZNAME:CDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0500
TZOFFSETTO:-0600
TZNAME:CST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE
""",

        "America/Denver": """BEGIN:VTIMEZONE
TZID:America/Denver
BEGIN:DAYLIGHT
TZOFFSETFROM:-0700
TZOFFSETTO:-0600
TZNAME:MDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0600
TZOFFSETTO:-0700
TZNAME:MST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE
""",

        "America/Los_Angeles": """BEGIN:VTIMEZONE
TZID:America/Los_Angeles
BEGIN:DAYLIGHT
TZOFFSETFROM:-0800
TZOFFSETTO:-0700
TZNAME:PDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0700
TZOFFSETTO:-0800
TZNAME:PST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE
""",

        "America/Anchorage": """BEGIN:VTIMEZONE
TZID:America/Anchorage
BEGIN:DAYLIGHT
TZOFFSETFROM:-0900
TZOFFSETTO:-0800
TZNAME:AKDT
DTSTART:19700308T020000
RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=2SU
END:DAYLIGHT
BEGIN:STANDARD
TZOFFSETFROM:-0800
TZOFFSETTO:-0900
TZNAME:AKST
DTSTART:19701101T020000
RRULE:FREQ=YEARLY;BYMONTH=11;BYDAY=1SU
END:STANDARD
END:VTIMEZONE
""",

        "Pacific/Honolulu": """BEGIN:VTIMEZONE
TZID:Pacific/Honolulu
BEGIN:STANDARD
TZOFFSETFROM:-1000
TZOFFSETTO:-1000
TZNAME:HST
DTSTART:19700101T000000
END:STANDARD
END:VTIMEZONE
"""
    }

    # Return full US block if available
    if tzid in US_TZ_BLOCKS:
        return US_TZ_BLOCKS[tzid]

    # Fallback for non-US zones
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
st. write("The .ics file will export with the date/start/end times correct. The Activity will be the event name, the POC and comments will be in the description block.")

st.write("For U.S. time zones, the app embeds official DST transition rules directly into the .ics file to prevent 1-hour time shifts when importing into Outlook. Users select the region only; DST is handled automatically. Applies to Eastern, Central, Mountain, Pacific, Alaska, and Hawaii. ")

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
