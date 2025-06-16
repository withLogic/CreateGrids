# Event Grid Generator

This script generates a printable Excel grid schedule for the Mensa events, including the Annual Gathering. It reads session data from a `csv` or `xls` file and outputs a well-formatted `xlsx` file suitable for printing and distribution.

## Features

- Supports CSV and XLS input formats
- Automatically generates a daily program grid
- Dynamically lays out sessions by time and room
- Hides rooms with no scheduled sessions (optional)
- Adds styled headers, footers, and gridlines
- Supports merging time blocks for longer sessions
- Highlights and skips unpublished sessions

## Input File Requirements

Your input file (CSV or XLS) **must** contain the following columns:

- `Session Title`
- `Session Start Date` (Format: MM/DD/YYYY)
- `Session Start Time` (Format: HH:MM AM/PM)
- `Session End Time` (Format: HH:MM AM/PM)
- `Room`

**Note:** XLS files must have the session data as the first table in the spreadsheet.

## Configuration

These constants can be customized in the script:

```python
EVENT_NAME = ""
START_DATE = datetime.datetime(2025, 7, 2)
ROOMS_TO_SUPPRESS = ['', '']
HIDE_ROOMS_WITH_NO_SESSIONS = 1
INTERVAL = 15  # Time grid interval in minutes
FOOTER_TEXT 
```

## Output
The script creates a file called:
`event_grid__out.xlsx`

Each worksheet represents one day, with rooms across the top and time slots on the side. Sessions are placed in the corresponding room and time cells, with multi-row merging for session duration.

## Usage
### Prerequisites
* Python 3.7+
* Install dependencies: `pip install openpyxl pandas`

### Update the variables
Each sheet includes a custom footer:

`Shop the Mensa Store and wear your brain on your sleeve with our exclusive licensed apparel.`

This can be customized by updating the `FOOTER_TEXT` variable in the script.

The event name is printed in the header of the sheet. This can be set by updating the `AG_NAME` variable.

If there are rooms that have sessions that should not be printed on the grids, you can supress them by putting the room names in the `ROOMS_TO_SUPPRESS`variable.

### Run the Script
`python createEventProgramGrids.py your_schedule_file.csv`
or
`python createEventProgramGrids.py your_schedule_file.xls`


## Known Limitations
* Sessions with overlapping time blocks in the same room are skipped and flagged.
* Date formats must be consistent across input rows.
* XLS files must have structured table data (first table only is parsed).