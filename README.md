# CSV to ICS Calendar Generator (Streamlit)

A Streamlit web application that converts a properly formatted CSV schedule into a standards-compliant `.ics` calendar file.

Designed for departmental use where users:

- Upload a CSV export from Excel
- Select a timezone
- Download a ready-to-import calendar file
- Ensure events display as **Busy**
- Preserve correct daylight savings time behavior

---

## Features

- Upload CSV directly in browser
- Automatic header normalization (case and spacing insensitive)
- Carries forward blank Date cells
- Skips incomplete rows
- Timezone selection
- Proper DST handling
- Outlook-compatible `.ics` generation
- All events set to `Busy` (`TRANSP:OPAQUE`)
- Clean formatting:
  - **SUMMARY = Activity**
  - **LOCATION = Location**
  - **DESCRIPTION = POC + Comments**
  - No quotation marks around values

---

## Expected CSV Format

Your CSV must contain columns equivalent to:

| Date | Start | End | Location | Activity | POC | Comments |

Header matching is flexible:

- Case insensitive (`date`, `DATE`, `Date`)
- Extra spaces ignored
- Accepts aliases:
  - `Activity`, `Event`, `Title`
  - `POC`, `Point of Contact`
  - `Comments`, `Notes`, `Details`

Blank Date cells are allowed and will inherit the most recent valid date above them.


- Works with Outlook Desktop
- Works with Outlook Web
- Works with Google Calendar
- All events import as Busy
- No meeting-request behavior (import-only calendar file, add invitees manually)

## Time Zone & Daylight Saving Time (DST) Handling

Calendar applications interpret time zone data differently, and some clients—especially Outlook Desktop—require explicit daylight saving time (DST) rules to prevent 1-hour shifts when importing `.ics` files.

To ensure consistent behavior, this application:

- Writes full DST transition rules for U.S. time zones:
  - America/New_York (Eastern)
  - America/Chicago (Central)
  - America/Denver (Mountain)
  - America/Los_Angeles (Pacific)
  - America/Anchorage (Alaska)
  - Pacific/Honolulu (Hawaii, no DST)
- Encodes the correct U.S. DST rules:
  - Second Sunday in March
  - First Sunday in November
- Automatically applies the correct offset for the event date

Users only select the regional time zone (e.g., `America/New_York`).  
They do **not** need to choose between standard time (EST) or daylight time (EDT).  
The calendar client applies DST automatically based on the event date.

This approach ensures reliable imports across:
- Outlook Desktop
- Outlook Web
- Google Calendar
- Apple Calendar
---

## Installation

### Requirements

Python 3.9 or newer (required for `zoneinfo`)

Install dependencies:

```bash
pip install -r requirements.txt