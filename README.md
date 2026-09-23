# Sheet Ops Toolkit

A small collection of Google Apps Script utilities that automate repetitive operations tasks inside Google Sheets, built while managing end-to-end lifecycles for 7+ parallel class programs and 700+ students at once.

Adds a **Custom Function** menu to any Google Sheet with five tools:

| Tool                                   | What it does                                                                                                                                                                                              |
| --------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Get Links from Drive**                | Lists every file in a Drive folder as a hyperlink + file name, in two columns you choose.                                                                                                                 |
| **Parse Json**                          | Reads a column of JSON strings and expands each into structured label/url rows.                                                                                                                           |
| **Get Name File from URL**              | Resolves a human-readable name from a raw URL in Drive file, Google Form title, Colab notebook, Drive folder, YouTube playlist, or a Bitly destination.                                                  |
| **List Filtered Events from Gcal**      | Pulls every calendar event in a date range into a sheet (title, description, location, start/end time, invited emails), used as a weekly audit to catch invite or scheduling errors before they reach students and instructors. |
| **Bulk Update Event Descriptions**      | Reads a sheet of (title, link, start time, description) rows and writes each description + link into the matching calendar event, used to bulk-attach recording/material links to a batch of class sessions at once. |

## Sheet Structure Reference

Most tools prompt for cells/columns interactively, so there's no fixed layout to follow but a couple of tools expect specific columns. Reference below is based on what the code actually reads/writes.

### Parse Json (output layout)

Input: one JSON array per row, e.g. `[{"label":"...","url":"..."}]`, in whatever column you point it at.

Output: starting from the cell you choose, each array item takes **3 columns** (label, url, blank spacer). Rows with more items than others just use more columns to the right — width auto-expands to the widest row.

| Label 1 | URL 1 | *(blank)* | Label 2 | URL 2 | *(blank)* | ... |
| ------- | ----- | --------- | ------- | ----- | --------- | --- |

### List Filtered Events from Gcal (output layout)

Writes to the sheet name you give it, overwriting its contents, with this fixed header row:

| Event ID | Title | Description | Location/Link | Start Time | End Time | Invited Emails |
| -------- | ----- | ------------ | -------------- | ---------- | -------- | --------------- |

### Bulk Update Event Descriptions (required input layout)

Reads from the sheet name you give it (row 1 = header, data starts row 2):

| Column | Content |
| ------ | ------- |
| A | Event title, must match the calendar event's title exactly (case-insensitive) |
| B | Link to append to the description |
| C | Event start date/time, matched to the calendar event's start time to the minute |
| D | New description text (this **replaces** the event's existing description) |

⚠️ Multiple events with the same title + start time will all get updated, and a warning is logged for that case — check Logs after running if you have recurring events with duplicate titles.

## Why I built this

As Class Operations for a multi-program academy, I was manually cross-checking calendar invites, matching Drive links to file names, and reformatting JSON exports every week. Each of these tools replaced a specific chunk of that manual work, the calendar auditor alone cut a task that used to take about an hour down to under a minute.

## Setup

1. Open (or create) a Google Sheet.
2. **Extensions → Apps Script**, paste `code.gs` in as the script, and save.
3. Reload the Sheet, a **Custom Function** menu appears in the menu bar.
4. *(Optional, only for the calendar tools)* In the Apps Script editor, under **Services**, add the **Google Calendar API** advanced service, and make sure the calendar ID you enter is one your account has access to.
5. *(Optional, only for resolving YouTube playlist names)* Under **Services**, add **YouTube Data API v3**.

## License

MIT - use it, fork it, adapt it.
