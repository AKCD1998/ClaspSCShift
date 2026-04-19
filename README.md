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

Data details:

- The schedule day columns start at column H. The web UI now detects the last
  day column from the date header row, so 30-day months end at AK and 31-day
  months end at AL. For generated 31-day month tabs, the generator inserts a
  real day-31 column before the note/holiday area instead of reusing the old
  blank separator column from the 30-day April template.
- Employee metadata is mostly in columns B through G.
- Notes/holiday leave/summary fields are to the right of the day columns.
- The web app intentionally renders the raw worksheet matrix instead of only
  normalized employee rows so it stays visually close to the original sheet.

Current next task:

- Improve web text readability. Some cells inherit the original sheet font
  colors and can become hard to read in the web app. A practical next step is to
  add a contrast helper in `Index.html` that chooses black or white text based
  on each cell background, while preserving explicit important colors only when
  they remain readable.

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
