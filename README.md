# Sheet Ops Toolkit

A small collection of Google Apps Script utilities that automate repetitive operations tasks inside Google Sheets, built while managing end-to-end lifecycles for 7+ parallel class programs and 700+ students at once.

Adds a **Custom Function** menu to any Google Sheet with four tools:

| Tool | What it does |
|---|---|
| **Get Links from Drive** | Lists every file in a Drive folder as a hyperlink + file name, in two columns you choose. |
| **Parse Json** | Reads a column of JSON strings and expands each into structured label/url rows. |
| **Get Name File from URL** | Resolves a human-readable name from a raw URL in Drive file, Google Form title, Colab notebook, Drive folder, YouTube playlist, or a Bitly destination. |
| **List Filtered Events from Gcal** | Pulls every calendar event in a date range into a sheet (title, description, location, start/end time, invited emails), used as a weekly audit to catch invite or scheduling errors before they reach students and instructors. |

## Why I built this

As Class Operations for a multi-program academy, I was manually cross-checking calendar invites, matching Drive links to file names, and reformatting JSON exports every week. Each of these tools replaced a specific chunk of that manual work, the calendar auditor alone cut a task that used to take about an hour down to under a minute.

## Setup

1. Open (or create) a Google Sheet.
2. **Extensions → Apps Script**, paste `code.gs` in as the script, and save.
3. Reload the Sheet, a **Custom Function** menu appears in the menu bar.
4. *(Optional, only for the calendar auditor)* In the Apps Script editor, under **Services**, add the **Google Calendar API** advanced service.
5. *(Optional, only for resolving YouTube playlist names)* Under **Services**, add **YouTube Data API v3**.

## License

MIT - use it, fork it, adapt it.
