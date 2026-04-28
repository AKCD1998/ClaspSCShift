# SC Shift Apps Script Reader

This folder contains a minimal Google Apps Script project for reading one worksheet
from the shift schedule spreadsheet.

## Target

- Spreadsheet ID: `1-t9xFhQzIqUCHZSlzLDnKa7l_YUb7vXB5d2RU4UEflw`
- Worksheet name: `แก้ไข เมษายน 69`
- Current implementation: opens the target spreadsheet by ID using
  `SpreadsheetApp.openById()`

## Project overview for Codex

This is a Google Apps Script + clasp project that reads a company shift
schedule from Google Sheets and renders it as a web app. The worksheet is a
visual schedule for pharmacy-chain employees. Cell text and cell colors both
matter because shift type, branch, leave, meeting, and time interpretation can
depend on the displayed code and formatting.

Main files:

- `Code.js`: Apps Script server code. It opens the target spreadsheet, reads the
  worksheet values and formatting, exposes inspection helpers, and serves the
  web app.
- `Index.html`: Web app UI. It calls `google.script.run.getWorksheetWebAppData()`
  and renders a spreadsheet-like table with a loading overlay while data is
  being fetched.
- `appsscript.json`: Apps Script manifest. It currently requests spreadsheet
  access because the script opens the target spreadsheet by ID.
- `.claspignore`: Limits pushed files to Apps Script-compatible files.

Important server functions:

- `doGet()`: Serves `Index.html` for the web app.
- `getWorksheetWebAppData()`: Returns the full worksheet matrix for the UI,
  including displayed values, background colors, font colors, font weights,
  alignments, row/column positions, A1 notation, merged-cell metadata, and
  available 2026 month tabs. It accepts an optional worksheet name; no argument
  still defaults to `แก้ไข เมษายน 69`.
- `inspectShiftSheet()`: Logs a compact worksheet summary.
- `inspectShiftSheetWithColors()`: Returns employee/day structured data with
  values and colors. This can be too large for Apps Script logs.
- `createBlankMonthlySheetsMayToDec2026()`: Copies the April 2569 worksheet as a
  template and creates blank monthly worksheet tabs from May through December
  2026. It updates the month label, day numbers, and Thai weekdays, then clears
  schedule values/colors in the roster grid while preserving the overall sheet
  structure. If a previous run failed after creating a future-month tab, the
  function repairs that partial tab when its month label is still wrong.
- `repairScheduleMonthLayouts2026()`: Repairs existing 2026 month tabs without
  clearing real roster data. It keeps H:AL reserved for days 1-31, moves the
  note/holiday area to the right of that grid when needed, clears only
  impossible future-day cells for 28/30-day months, clears stray day/weekday
  headers from note columns, and breaks merged ranges anchored inside the
  schedule-day body.

Data details:

- The schedule day columns start at column H and always reserve columns H:AL
  for days 1-31. Months with fewer than 31 days keep blank day columns so the
  note/holiday area stays aligned across February, 30-day months, and 31-day
  months.
- Employee metadata is mostly in columns B through G.
- Notes/holiday leave/summary fields are to the right of the day columns.
- The web app intentionally renders the raw worksheet matrix instead of only
  normalized employee rows so it stays visually close to the original sheet.

Current implementation status:

- The web app has an Edit Mode toggle. Editable schedule cells can be selected,
  changed locally, cleared locally, reviewed, and then saved back to the
  monthly roster worksheet.
- Pending edits are stored on the client until the user confirms them. The
  payload sent to Apps Script includes sheet name, row, column, previous raw
  code, next raw code, action type, employee/date labels, and optional shift
  metadata.
- Confirmed edits write raw codes back to the source Google Sheet through
  `saveScheduleChanges()`. Successful changes are appended to
  `Schedule Audit Log`; audit failure is reported per item but does not block a
  successful sheet write.
- Shift metadata can come from `Shift Dictionary` / `รหัสเวร`, with fallback
  items in Apps Script. The UI catalog uses this metadata for labels,
  descriptions, and colors.
- Monthly regular Off days are computed per employee and per selected month as
  `Off1` through `Off6`. Users choose "Set Off"; the system assigns the number.
- Annual Leave is computed per employee and per calendar year as `AL1` through
  `AL6`. Users choose the catalog option; the system assigns the number.
- Public Holiday leave is computed per employee and per calendar year as
  `Off1` through `Off13`, using the holiday color/metadata so it stays distinct
  from regular monthly Off. Users choose the catalog option
  `ตั้งนักขัตฤกษ์`; the system assigns the number.
- The lightweight pseudo-relational layer currently includes normalized
  employees, schedule entries, shift dictionary rows, optional draft changes,
  and audit log repository helpers. The current monthly roster worksheets remain
  the operational source of truth.

## Setup

1. Enable the Apps Script API for your Google account:
   https://script.google.com/home/usersettings
2. Log in with the Google account that has access to the spreadsheet:

   ```powershell
   clasp login
   ```

3. Create a bound Apps Script project on the spreadsheet:

   ```powershell
   clasp create --type sheets --title "SC Shift Reader" --parentId 1-t9xFhQzIqUCHZSlzLDnKa7l_YUb7vXB5d2RU4UEflw
   ```

   If `clasp` asks whether to overwrite `appsscript.json`, keep this local
   file. It contains the spreadsheet OAuth scope needed by `openById()`.

4. Push the local files:

   ```powershell
   clasp push
   ```

5. Open the Apps Script editor, select `inspectShiftSheet`, and run it:

   ```powershell
   clasp open-script
   ```

Run `inspectShiftSheetWithColors` when you need the structured employee/day
grid. Each day cell includes the displayed value, background color, font color,
spreadsheet row/column, and A1 notation.

Run `createBlankMonthlySheetsMayToDec2026` from the Apps Script editor when you
need blank future-month tabs. Existing tabs with the target names are skipped,
not overwritten when their month label is already correct.

Run `repairBlankMonthlySheetsMayToDec2026` as a convenient alias when a previous
generation attempt failed partway through. It runs the same repair-safe
generator.

Run `repairScheduleMonthLayouts2026` when existing month tabs have shifted
note/holiday columns or merged cells inside the day grid. This is non-destructive
for real roster values; it only inserts missing reserved day columns, clears
future-day cells that cannot exist for that month, clears duplicate day/weekday
headers in note columns, and breaks merges anchored in the schedule-day body.

## Web app

The project includes `doGet()` and `Index.html`. The web app reads the worksheet
through `getWorksheetWebAppData()` and renders the sheet matrix with cell values,
background colors, text colors, alignments, and merged cells.

The month label cell in the roster header opens a month picker modal for 2026.
Selecting an available month reloads the table from that worksheet tab. Missing
month tabs are shown as disabled options.

To deploy it from Apps Script:

1. Open the Apps Script editor:

   ```powershell
   clasp open-script
   ```

2. Select **Deploy > New deployment**.
3. Choose **Web app** as the deployment type.
4. Set **Execute as** to your account.
5. Set access to the users who should be able to view the schedule.
6. Deploy, authorize, then open the web app URL.

Do not open URLs that look like this:

```text
https://script.google.com/macros/library/d/<scriptId>/<version>
```

That is the Apps Script library/version page and it will only list functions
such as `doGet()` and `getWorksheetWebAppData()`. The web app URL must look
like this:

```text
https://script.google.com/macros/s/<deploymentId>/exec
```

From this folder you can open a deployed web app with:

```powershell
clasp open-web-app AKfycbxmcfQH9yYRYPnJq3eN7kHttk6mKvZYfIbvYRLbLRMIEupnxzeUus0jJ66JChQvkina
```

Current web app URL:

```text
https://script.google.com/macros/s/AKfycbxmcfQH9yYRYPnJq3eN7kHttk6mKvZYfIbvYRLbLRMIEupnxzeUus0jJ66JChQvkina/exec
```

The first run asks for authorization. The manifest requests spreadsheet access
because the script opens the target spreadsheet by ID. The Google account that
runs the function must have access to the target spreadsheet.
