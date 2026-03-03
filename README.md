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
---

## Installation

### Requirements

Python 3.9 or newer (required for `zoneinfo`)

Install dependencies:

```bash
pip install -r requirements.txt