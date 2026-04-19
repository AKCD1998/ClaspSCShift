const SPREADSHEET_ID = '1-t9xFhQzIqUCHZSlzLDnKa7l_YUb7vXB5d2RU4UEflw';
const SHEET_NAME = 'แก้ไข เมษายน 69';
const MONTH_TEMPLATE_YEAR = 2026;
const MONTH_TEMPLATE_YEAR_BE = 2569;
const MONTH_TEMPLATE_YEAR_SHORT_BE = '69';
const MONTH_TEMPLATE_START_MONTH = 5;
const MONTH_TEMPLATE_END_MONTH = 12;
const MONTH_LABEL_ROW = 3;
const DAY_NUMBER_ROW = 4;
const WEEKDAY_ROW = 5;
const DAY_START_COLUMN = 8;
const MAX_DAY_COLUMNS = 31;
const DAY_END_COLUMN = DAY_START_COLUMN + MAX_DAY_COLUMNS - 1;
const NOTE_START_COLUMN = DAY_END_COLUMN + 1;
const HOLIDAY_LEAVE_COLUMN = NOTE_START_COLUMN + 1;
const FIRST_SCHEDULE_ROW = 6;
const THIRTIETH_DAY_COLUMN = DAY_START_COLUMN + 30 - 1;
const THIRTY_FIRST_DAY_COLUMN = DAY_END_COLUMN;
const HOLIDAY_LEAVE_LABEL = 'นักขัตฤกษ์ / ลาพักร้อน';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SC Shift Schedule')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWorksheetWebAppData(sheetName) {
  const spreadsheet = getTargetSpreadsheet_();
  const targetSheetName = getReadableScheduleSheetName_(spreadsheet, sheetName);
  const sheet = spreadsheet.getSheetByName(targetSheetName);
  assertSheet_(spreadsheet, sheet, targetSheetName);

  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  const backgrounds = range.getBackgrounds();
  const fontColors = getFontColorHexGrid_(range);
  const fontWeights = range.getFontWeights();
  const horizontalAlignments = range.getHorizontalAlignments();
  const verticalAlignments = range.getVerticalAlignments();

  return {
    spreadsheetName: spreadsheet.getName(),
    spreadsheetId: spreadsheet.getId(),
    sheetName: sheet.getName(),
    sheetId: sheet.getSheetId(),
    fetchedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    rowCount: values.length,
    columnCount: values[0] ? values[0].length : 0,
    monthOptions: getScheduleMonthOptions_(spreadsheet, targetSheetName),
    rows: buildWorksheetRows_(values, backgrounds, fontColors, fontWeights, horizontalAlignments, verticalAlignments),
    merges: getMergeMeta_(range),
  };
}

function inspectShiftSheet() {
  const spreadsheet = getTargetSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  assertSheet_(spreadsheet, sheet);

  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  const summary = summarizeShiftSheet_(spreadsheet.getName(), sheet.getName(), values);

  console.log(JSON.stringify(summary, null, 2));
  return summary;
}

function inspectShiftSheetWithColors() {
  const spreadsheet = getTargetSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(SHEET_NAME);
  assertSheet_(spreadsheet, sheet);

  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  const backgrounds = range.getBackgrounds();
  const fontColors = getFontColorHexGrid_(range);
  const shiftGrid = buildShiftGrid_(spreadsheet, sheet, values, backgrounds, fontColors);

  console.log(JSON.stringify(shiftGrid, null, 2));
  return shiftGrid;
}

function createBlankMonthlySheetsMayToDec2026() {
  const spreadsheet = getTargetSpreadsheet_();
  const templateSheet = spreadsheet.getSheetByName(SHEET_NAME);
  assertSheet_(spreadsheet, templateSheet);

  const created = [];
  const repaired = [];
  const skipped = [];

  for (let month = MONTH_TEMPLATE_START_MONTH; month <= MONTH_TEMPLATE_END_MONTH; month += 1) {
    const monthMeta = getMonthTemplateMeta_(month);
    const existingSheet = spreadsheet.getSheetByName(monthMeta.sheetName);

    if (existingSheet) {
      if (shouldRepairExistingMonthlySheet_(existingSheet, monthMeta)) {
        prepareBlankMonthlySheet_(existingSheet, monthMeta);
        repaired.push(monthMeta.sheetName);
      } else {
        skipped.push(monthMeta.sheetName);
      }
      continue;
    }

    const sheet = templateSheet.copyTo(spreadsheet);
    sheet.setName(monthMeta.sheetName);
    prepareBlankMonthlySheet_(sheet, monthMeta);
    created.push(monthMeta.sheetName);
  }

  const result = {
    created,
    repaired,
    skipped,
    createdCount: created.length,
    repairedCount: repaired.length,
    skippedCount: skipped.length,
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function getTargetSpreadsheet_() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet && activeSpreadsheet.getId() === SPREADSHEET_ID) {
    return activeSpreadsheet;
  }

  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getMonthTemplateMeta_(month) {
  const thaiMonthNames = [
    '',
    'มกราคม',
    'กุมภาพันธ์',
    'มีนาคม',
    'เมษายน',
    'พฤษภาคม',
    'มิถุนายน',
    'กรกฎาคม',
    'สิงหาคม',
    'กันยายน',
    'ตุลาคม',
    'พฤศจิกายน',
    'ธันวาคม',
  ];
  const daysInMonth = new Date(MONTH_TEMPLATE_YEAR, month, 0).getDate();

  return {
    month,
    thaiMonthName: thaiMonthNames[month],
    daysInMonth,
    sheetName: `แก้ไข ${thaiMonthNames[month]} ${MONTH_TEMPLATE_YEAR_SHORT_BE}`,
    monthLabel: `เดือน  ${thaiMonthNames[month]} ${MONTH_TEMPLATE_YEAR_BE}`,
  };
}

function repairBlankMonthlySheetsMayToDec2026() {
  return createBlankMonthlySheetsMayToDec2026();
}

function repairScheduleMonthLayouts2026() {
  const spreadsheet = getTargetSpreadsheet_();
  const repaired = [];
  const skipped = [];

  for (let month = 1; month <= 12; month += 1) {
    const monthMeta = getMonthTemplateMeta_(month);
    const sheet = spreadsheet.getSheetByName(monthMeta.sheetName);

    if (!sheet) {
      skipped.push({
        sheetName: monthMeta.sheetName,
        reason: 'missing sheet',
      });
      continue;
    }

    repaired.push(repairScheduleMonthLayout_(sheet, monthMeta));
  }

  const result = {
    repaired,
    skipped,
    repairedCount: repaired.length,
    skippedCount: skipped.length,
  };
  console.log(JSON.stringify(result, null, 2));
  return result;
}

function prepareBlankMonthlySheet_(sheet, monthMeta) {
  const lastScheduleRow = findLastScheduleTemplateRow_(sheet);

  normalizeMonthDayGridLayout_(sheet, monthMeta, lastScheduleRow);
  writeMonthHeaders_(sheet, monthMeta);
  clearScheduleTemplateArea_(sheet, lastScheduleRow);
}

function shouldRepairExistingMonthlySheet_(sheet, monthMeta) {
  return (
    sheet.getRange(MONTH_LABEL_ROW, DAY_START_COLUMN).getDisplayValue() !== monthMeta.monthLabel ||
    needsMonthDayGridLayoutRepair_(sheet, monthMeta)
  );
}

function repairScheduleMonthLayout_(sheet, monthMeta) {
  const lastScheduleRow = findLastScheduleTemplateRow_(sheet);
  const insertedDayColumns = ensureReservedDayColumns_(sheet, monthMeta);
  writeMonthHeaders_(sheet, monthMeta);
  const clearedExtraNoteHeaders = clearExtraNoteHeaderValues_(sheet);
  formatFutureDayColumns_(sheet, monthMeta);
  const futureDayNoteRescueCandidates = getFutureDayNoteRescueCandidates_(sheet, monthMeta, lastScheduleRow);
  const brokenScheduleMerges = breakMergedRangesInArea_(
    sheet,
    FIRST_SCHEDULE_ROW,
    Math.max(lastScheduleRow - FIRST_SCHEDULE_ROW + 1, 1),
    DAY_START_COLUMN,
    MAX_DAY_COLUMNS,
  );
  const rescuedFutureDayNotes = rescueFutureDayNoteValues_(sheet, futureDayNoteRescueCandidates);
  clearFutureDayColumns_(sheet, monthMeta, lastScheduleRow);

  return {
    sheetName: sheet.getName(),
    daysInMonth: monthMeta.daysInMonth,
    insertedDayColumns,
    clearedExtraNoteHeaders,
    brokenScheduleMerges,
    rescuedFutureDayNotes,
    holidayLeaveColumn: findHolidayLeaveColumn_(sheet),
  };
}

function normalizeMonthDayGridLayout_(sheet, monthMeta, lastScheduleRow) {
  const insertedDayColumns = ensureReservedDayColumns_(sheet, monthMeta);
  formatThirtyFirstDayColumn_(sheet, lastScheduleRow);
  formatFutureDayColumns_(sheet, monthMeta);
  clearExtraNoteHeaderValues_(sheet);
  return insertedDayColumns;
}

function needsMonthDayGridLayoutRepair_(sheet, monthMeta) {
  const holidayColumn = findHolidayLeaveColumn_(sheet);
  return (
    Boolean(holidayColumn && holidayColumn < HOLIDAY_LEAVE_COLUMN) ||
    hasMergedRangesAnchoredInScheduleBody_(sheet) ||
    hasIncorrectDayHeaders_(sheet, monthMeta) ||
    hasExtraNoteHeaderValues_(sheet)
  );
}

function isHolidayLeaveColumn_(sheet, column) {
  return sheet.getRange(DAY_NUMBER_ROW, column).getDisplayValue().indexOf(HOLIDAY_LEAVE_LABEL) !== -1;
}

function findHolidayLeaveColumn_(sheet) {
  const values = sheet.getRange(DAY_NUMBER_ROW, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];

  for (let index = 0; index < values.length; index += 1) {
    if (values[index].indexOf(HOLIDAY_LEAVE_LABEL) !== -1) {
      return index + 1;
    }
  }

  return 0;
}

function ensureReservedDayColumns_(sheet, monthMeta) {
  let holidayColumn = findHolidayLeaveColumn_(sheet);
  let insertedColumns = 0;

  while (holidayColumn && holidayColumn < HOLIDAY_LEAVE_COLUMN) {
    const firstNonDayColumn = Math.max(DAY_START_COLUMN + monthMeta.daysInMonth, holidayColumn - 1);
    sheet.insertColumnBefore(firstNonDayColumn);
    insertedColumns += 1;
    holidayColumn += 1;
  }

  return insertedColumns;
}

function clearExtraNoteHeaderValues_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < NOTE_START_COLUMN) {
    return 0;
  }

  const headerRange = sheet.getRange(DAY_NUMBER_ROW, NOTE_START_COLUMN, 2, lastColumn - NOTE_START_COLUMN + 1);
  const values = headerRange.getDisplayValues();
  let clearedCount = 0;

  values.forEach((rowValues, rowOffset) => {
    rowValues.forEach((value, columnOffset) => {
      const row = DAY_NUMBER_ROW + rowOffset;
      const column = NOTE_START_COLUMN + columnOffset;

      if (!isExtraNoteHeaderValue_(value, row)) {
        return;
      }

      sheet.getRange(row, column).clearContent();
      clearedCount += 1;
    });
  });

  return clearedCount;
}

function formatFutureDayColumns_(sheet, monthMeta) {
  const blankColumnCount = MAX_DAY_COLUMNS - monthMeta.daysInMonth;
  if (blankColumnCount <= 0) {
    return;
  }

  const firstBlankColumn = DAY_START_COLUMN + monthMeta.daysInMonth;
  const sourceColumn = Math.max(DAY_START_COLUMN, firstBlankColumn - 1);
  const firstFormatRow = DAY_NUMBER_ROW;
  const rowCount = WEEKDAY_ROW - DAY_NUMBER_ROW + 1;

  for (let column = firstBlankColumn; column <= DAY_END_COLUMN; column += 1) {
    copyColumnFormat_(sheet, sourceColumn, column, firstFormatRow, rowCount);
    sheet.setColumnWidth(column, sheet.getColumnWidth(sourceColumn));
  }
}

function copyColumnFormat_(sheet, sourceColumn, targetColumn, firstRow, rowCount) {
  const sourceRange = sheet.getRange(firstRow, sourceColumn, rowCount, 1);
  const targetRange = sheet.getRange(firstRow, targetColumn, rowCount, 1);

  targetRange
    .setBackgrounds(sourceRange.getBackgrounds())
    .setFontColors(sourceRange.getFontColors())
    .setFontWeights(sourceRange.getFontWeights())
    .setHorizontalAlignments(sourceRange.getHorizontalAlignments())
    .setVerticalAlignments(sourceRange.getVerticalAlignments())
    .setNumberFormats(sourceRange.getNumberFormats());
}

function clearFutureDayColumns_(sheet, monthMeta, lastScheduleRow) {
  const blankColumnCount = MAX_DAY_COLUMNS - monthMeta.daysInMonth;
  if (blankColumnCount <= 0 || lastScheduleRow < FIRST_SCHEDULE_ROW) {
    return;
  }

  const firstBlankColumn = DAY_START_COLUMN + monthMeta.daysInMonth;
  const rowCount = lastScheduleRow - FIRST_SCHEDULE_ROW + 1;
  const actualDayRange = sheet.getRange(FIRST_SCHEDULE_ROW, DAY_START_COLUMN, rowCount, monthMeta.daysInMonth);
  const actualDayValues = actualDayRange.getDisplayValues();
  const actualDayBackgrounds = actualDayRange.getBackgrounds();
  const blankValues = [];
  const blankBackgrounds = [];
  const fontColors = [];
  const fontWeights = [];
  const alignments = [];

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const baseBackground = getBaseScheduleBackground_(actualDayValues[rowIndex], actualDayBackgrounds[rowIndex]);

    blankValues.push(Array(blankColumnCount).fill(''));
    blankBackgrounds.push(Array(blankColumnCount).fill(baseBackground));
    fontColors.push(Array(blankColumnCount).fill('#000000'));
    fontWeights.push(Array(blankColumnCount).fill('normal'));
    alignments.push(Array(blankColumnCount).fill('center'));
  }

  sheet
    .getRange(FIRST_SCHEDULE_ROW, firstBlankColumn, rowCount, blankColumnCount)
    .setValues(blankValues)
    .setBackgrounds(blankBackgrounds)
    .setFontColors(fontColors)
    .setFontWeights(fontWeights)
    .setHorizontalAlignments(alignments);
}

function getFutureDayNoteRescueCandidates_(sheet, monthMeta, lastScheduleRow) {
  const blankColumnCount = MAX_DAY_COLUMNS - monthMeta.daysInMonth;
  if (blankColumnCount <= 0 || lastScheduleRow < FIRST_SCHEDULE_ROW) {
    return [];
  }

  const firstFutureDayColumn = DAY_START_COLUMN + monthMeta.daysInMonth;
  const lastFutureDayColumn = DAY_END_COLUMN;
  const lastScheduleBodyRow = lastScheduleRow;

  return sheet
    .getDataRange()
    .getMergedRanges()
    .reduce((candidates, mergedRange) => {
      const mergedRow = mergedRange.getRow();
      const mergedEndRow = mergedRow + mergedRange.getNumRows() - 1;
      const mergedColumn = mergedRange.getColumn();
      const mergedEndColumn = mergedColumn + mergedRange.getNumColumns() - 1;
      const startsInFutureDayColumn =
        mergedColumn >= firstFutureDayColumn && mergedColumn <= lastFutureDayColumn;
      const intersectsScheduleBody =
        mergedRow <= lastScheduleBodyRow && mergedEndRow >= FIRST_SCHEDULE_ROW;
      const crossesIntoNoteArea = mergedEndColumn >= NOTE_START_COLUMN;

      if (startsInFutureDayColumn && intersectsScheduleBody && crossesIntoNoteArea) {
        candidates.push({
          sourceRow: Math.max(mergedRow, FIRST_SCHEDULE_ROW),
          sourceColumn: mergedColumn,
          rowSpan: Math.min(mergedEndRow, lastScheduleBodyRow) - Math.max(mergedRow, FIRST_SCHEDULE_ROW) + 1,
          targetColumn: NOTE_START_COLUMN,
        });
      }

      return candidates;
    }, []);
}

function rescueFutureDayNoteValues_(sheet, candidates) {
  let rescuedCount = 0;

  candidates.forEach((candidate) => {
    const sourceCell = sheet.getRange(candidate.sourceRow, candidate.sourceColumn);
    const targetCell = sheet.getRange(candidate.sourceRow, candidate.targetColumn);
    const sourceValue = sourceCell.getDisplayValue();

    if (!sourceValue || targetCell.getDisplayValue()) {
      return;
    }

    const targetRange = sheet.getRange(candidate.sourceRow, candidate.targetColumn, candidate.rowSpan, 1);
    targetRange.breakApart();

    if (candidate.rowSpan > 1) {
      targetRange.merge();
    }

    sourceCell.copyTo(targetCell, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    sourceCell.copyTo(targetCell, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    rescuedCount += 1;
  });

  return rescuedCount;
}

function getBaseScheduleBackground_(rowValues, rowBackgrounds) {
  let explicitBlankIndex = -1;
  for (let columnIndex = rowValues.length - 1; columnIndex >= 0; columnIndex -= 1) {
    if (rowValues[columnIndex] === '') {
      explicitBlankIndex = columnIndex;
      break;
    }
  }

  if (explicitBlankIndex !== -1) {
    return rowBackgrounds[explicitBlankIndex];
  }

  return getMostCommonValue_(rowBackgrounds) || '#ffffff';
}

function hasMergedRangesAnchoredInScheduleBody_(sheet) {
  return sheet
    .getDataRange()
    .getMergedRanges()
    .some((mergedRange) => {
      const mergedRow = mergedRange.getRow();
      const mergedEndRow = mergedRow + mergedRange.getNumRows() - 1;
      const mergedColumn = mergedRange.getColumn();

      return (
        mergedEndRow >= FIRST_SCHEDULE_ROW &&
        mergedColumn >= DAY_START_COLUMN &&
        mergedColumn <= DAY_END_COLUMN
      );
    });
}

function hasIncorrectDayHeaders_(sheet, monthMeta) {
  const dayNumbers = sheet.getRange(DAY_NUMBER_ROW, DAY_START_COLUMN, 1, MAX_DAY_COLUMNS).getDisplayValues()[0];
  const weekdays = sheet.getRange(WEEKDAY_ROW, DAY_START_COLUMN, 1, MAX_DAY_COLUMNS).getDisplayValues()[0];

  for (let day = 1; day <= MAX_DAY_COLUMNS; day += 1) {
    const index = day - 1;
    const expectedDayNumber = day <= monthMeta.daysInMonth ? String(day) : '';
    const expectedWeekday =
      day <= monthMeta.daysInMonth
        ? getThaiWeekday_(new Date(MONTH_TEMPLATE_YEAR, monthMeta.month - 1, day))
        : '';

    if (dayNumbers[index] !== expectedDayNumber || weekdays[index] !== expectedWeekday) {
      return true;
    }
  }

  return false;
}

function hasExtraNoteHeaderValues_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < NOTE_START_COLUMN) {
    return false;
  }

  const values = sheet
    .getRange(DAY_NUMBER_ROW, NOTE_START_COLUMN, 2, lastColumn - NOTE_START_COLUMN + 1)
    .getDisplayValues();

  return values.some((rowValues, rowOffset) =>
    rowValues.some((value) => isExtraNoteHeaderValue_(value, DAY_NUMBER_ROW + rowOffset)),
  );
}

function isExtraNoteHeaderValue_(value, row) {
  const text = String(value || '').trim();
  if (!text) {
    return false;
  }

  if (row === DAY_NUMBER_ROW) {
    const dayNumber = Number(text);
    return Number.isInteger(dayNumber) && dayNumber >= 1 && dayNumber <= MAX_DAY_COLUMNS;
  }

  if (row === WEEKDAY_ROW) {
    return isThaiWeekdayHeaderValue_(text);
  }

  return false;
}

function isThaiWeekdayHeaderValue_(value) {
  return ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'].indexOf(value) !== -1;
}

function formatThirtyFirstDayColumn_(sheet, lastScheduleRow) {
  const firstFormatRow = DAY_NUMBER_ROW;
  const rowCount = Math.max(lastScheduleRow - firstFormatRow + 1, 1);
  const sourceRange = sheet.getRange(firstFormatRow, THIRTIETH_DAY_COLUMN, rowCount, 1);
  const targetRange = sheet.getRange(firstFormatRow, THIRTY_FIRST_DAY_COLUMN, rowCount, 1);

  targetRange
    .setBackgrounds(sourceRange.getBackgrounds())
    .setFontColors(sourceRange.getFontColors())
    .setFontWeights(sourceRange.getFontWeights())
    .setHorizontalAlignments(sourceRange.getHorizontalAlignments())
    .setVerticalAlignments(sourceRange.getVerticalAlignments())
    .setNumberFormats(sourceRange.getNumberFormats());
  sheet.setColumnWidth(THIRTY_FIRST_DAY_COLUMN, sheet.getColumnWidth(THIRTIETH_DAY_COLUMN));
}

function writeMonthHeaders_(sheet, monthMeta) {
  const dayNumbers = [];
  const weekdays = [];

  for (let day = 1; day <= MAX_DAY_COLUMNS; day += 1) {
    if (day <= monthMeta.daysInMonth) {
      const date = new Date(MONTH_TEMPLATE_YEAR, monthMeta.month - 1, day);
      dayNumbers.push(String(day));
      weekdays.push(getThaiWeekday_(date));
    } else {
      dayNumbers.push('');
      weekdays.push('');
    }
  }

  sheet.getRange(MONTH_LABEL_ROW, DAY_START_COLUMN).setValue(monthMeta.monthLabel);
  sheet.getRange(DAY_NUMBER_ROW, DAY_START_COLUMN, 1, MAX_DAY_COLUMNS).setValues([dayNumbers]);
  sheet.getRange(WEEKDAY_ROW, DAY_START_COLUMN, 1, MAX_DAY_COLUMNS).setValues([weekdays]);
}

function clearScheduleTemplateArea_(sheet, lastScheduleRow) {
  if (lastScheduleRow < FIRST_SCHEDULE_ROW) {
    return;
  }

  const rowCount = lastScheduleRow - FIRST_SCHEDULE_ROW + 1;
  breakMergedRangesInArea_(sheet, FIRST_SCHEDULE_ROW, rowCount, DAY_START_COLUMN, MAX_DAY_COLUMNS);
  const range = sheet.getRange(FIRST_SCHEDULE_ROW, DAY_START_COLUMN, rowCount, MAX_DAY_COLUMNS);
  const values = range.getDisplayValues();
  const backgrounds = range.getBackgrounds();
  const baseBackgrounds = backgrounds.map((rowBackgrounds, index) =>
    getBaseScheduleBackground_(values[index], rowBackgrounds),
  );

  const blankValues = [];
  const blankBackgrounds = [];
  const fontColors = [];
  const fontWeights = [];
  const alignments = [];

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    blankValues.push(Array(MAX_DAY_COLUMNS).fill(''));
    blankBackgrounds.push(Array(MAX_DAY_COLUMNS).fill(baseBackgrounds[rowIndex]));
    fontColors.push(Array(MAX_DAY_COLUMNS).fill('#000000'));
    fontWeights.push(Array(MAX_DAY_COLUMNS).fill('normal'));
    alignments.push(Array(MAX_DAY_COLUMNS).fill('center'));
  }

  range
    .setValues(blankValues)
    .setBackgrounds(blankBackgrounds)
    .setFontColors(fontColors)
    .setFontWeights(fontWeights)
    .setHorizontalAlignments(alignments);
}

function breakMergedRangesInArea_(sheet, row, numRows, column, numColumns) {
  if (numRows <= 0 || numColumns <= 0) {
    return 0;
  }

  const targetEndRow = row + numRows - 1;
  const targetEndColumn = column + numColumns - 1;
  const mergedRanges = sheet.getDataRange().getMergedRanges();
  let brokenCount = 0;

  mergedRanges.forEach((mergedRange) => {
    const mergedRow = mergedRange.getRow();
    const mergedEndRow = mergedRow + mergedRange.getNumRows() - 1;
    const mergedColumn = mergedRange.getColumn();
    const mergedEndColumn = mergedColumn + mergedRange.getNumColumns() - 1;
    const intersectsRows = mergedRow <= targetEndRow && mergedEndRow >= row;
    const intersectsColumns = mergedColumn <= targetEndColumn && mergedEndColumn >= column;

    if (intersectsRows && intersectsColumns) {
      mergedRange.breakApart();
      brokenCount += 1;
    }
  });

  return brokenCount;
}

function findLastScheduleTemplateRow_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();

  for (let rowIndex = 0; rowIndex < values.length; rowIndex += 1) {
    if ((values[rowIndex][1] || '').trim() === 'หมายเหตุ') {
      return rowIndex;
    }
  }

  return values.length;
}

function getThaiWeekday_(date) {
  const weekdays = ['อา.', 'จ.', 'อ.', 'พ.', 'พฤ.', 'ศ.', 'ส.'];
  return weekdays[date.getDay()];
}

function getMostCommonValue_(values) {
  const counts = values.reduce((map, value) => {
    map[value] = (map[value] || 0) + 1;
    return map;
  }, {});

  return Object.keys(counts).sort((left, right) => counts[right] - counts[left])[0] || '';
}

function getReadableScheduleSheetName_(spreadsheet, sheetName) {
  const requestedSheetName = String(sheetName || '').trim();
  if (!requestedSheetName) {
    return SHEET_NAME;
  }

  const readableSheetNames = getScheduleMonthOptions_(spreadsheet).map((option) => option.sheetName);
  if (readableSheetNames.indexOf(requestedSheetName) === -1) {
    throw new Error(`Sheet "${requestedSheetName}" is not a supported 2026 schedule month.`);
  }

  return requestedSheetName;
}

function getScheduleMonthOptions_(spreadsheet, selectedSheetName) {
  const options = [];

  for (let month = 1; month <= 12; month += 1) {
    const meta = getMonthTemplateMeta_(month);
    const exists = Boolean(spreadsheet.getSheetByName(meta.sheetName));

    options.push({
      month,
      year: MONTH_TEMPLATE_YEAR,
      yearBE: MONTH_TEMPLATE_YEAR_BE,
      thaiMonthName: meta.thaiMonthName,
      label: `${meta.thaiMonthName} ${MONTH_TEMPLATE_YEAR_BE}`,
      sheetName: meta.sheetName,
      available: exists,
      selected: meta.sheetName === selectedSheetName,
    });
  }

  return options;
}

function assertSheet_(spreadsheet, sheet, expectedSheetName) {
  if (sheet) {
    return;
  }

  const availableSheets = spreadsheet.getSheets().map((item) => item.getName());
  const sheetName = expectedSheetName || SHEET_NAME;
  throw new Error(
    `Sheet "${sheetName}" was not found in spreadsheet "${spreadsheet.getName()}" ` +
      `(${spreadsheet.getId()}). Available sheets: ${availableSheets.join(', ')}`,
  );
}

function summarizeShiftSheet_(spreadsheetName, sheetName, values) {
  const title = values[1] && values[1][1] ? values[1][1] : '';
  const monthLabel = values[2] && values[2][7] ? values[2][7] : '';
  const dayNumbers = values[3] ? values[3].slice(7, 37).filter(String) : [];
  const sections = [];
  const sectionSet = new Set();

  values.forEach((row) => {
    const label = (row[1] || '').trim();
    const hasName = (row[2] || '').trim();
    const looksLikeSection =
      label &&
      !hasName &&
      (label === 'เภสัชกร' ||
        label === 'ฝ่ายขายและการตลาด' ||
        label === 'หมายเหตุ' ||
        /^0\d{2}\s/.test(label));

    if (looksLikeSection && !sectionSet.has(label)) {
      sectionSet.add(label);
      sections.push(label);
    }
  });

  return {
    spreadsheetName,
    sheetName,
    title,
    monthLabel,
    rows: values.length,
    columns: values[0] ? values[0].length : 0,
    dayCount: dayNumbers.length,
    sections,
    preview: values.slice(0, 12),
  };
}

function buildShiftGrid_(spreadsheet, sheet, values, backgrounds, fontColors) {
  const dayStartIndex = 7;
  const dayEndIndex = 37;
  const dayHeaderRowIndex = 3;
  const weekdayHeaderRowIndex = 4;
  const rows = [];
  let section = '';

  values.forEach((row, rowIndex) => {
    const sectionLabel = getSectionLabel_(row);
    if (sectionLabel) {
      section = sectionLabel;
      return;
    }

    const sequence = (row[1] || '').trim();
    const fullName = (row[2] || '').trim();
    if (!/^\d+$/.test(sequence) || !fullName || section === 'หมายเหตุ') {
      return;
    }

    const dayCells = [];
    for (let columnIndex = dayStartIndex; columnIndex <= dayEndIndex; columnIndex += 1) {
      const day = values[dayHeaderRowIndex][columnIndex];
      const weekday = values[weekdayHeaderRowIndex][columnIndex];
      const value = row[columnIndex] || '';

      dayCells.push({
        day,
        weekday,
        value,
        background: backgrounds[rowIndex][columnIndex],
        fontColor: fontColors[rowIndex][columnIndex],
        row: rowIndex + 1,
        column: columnIndex + 1,
        a1: toA1_(rowIndex + 1, columnIndex + 1),
      });
    }

    rows.push({
      section,
      sequence,
      fullName,
      nickname: row[3] || '',
      birthDate: row[4] || '',
      position: row[5] || '',
      team: row[6] || '',
      note: row[37] || '',
      holidayLeave: row[39] || row[38] || '',
      total: row[41] || row[40] || '',
      sourceRow: rowIndex + 1,
      days: dayCells,
    });
  });

  return {
    spreadsheetName: spreadsheet.getName(),
    spreadsheetId: spreadsheet.getId(),
    sheetName: sheet.getName(),
    sheetId: sheet.getSheetId(),
    title: values[1] && values[1][1] ? values[1][1] : '',
    monthLabel: values[2] && values[2][7] ? values[2][7] : '',
    rows: rows.length,
    dayCount: dayEndIndex - dayStartIndex + 1,
    data: rows,
  };
}

function buildWorksheetRows_(values, backgrounds, fontColors, fontWeights, horizontalAlignments, verticalAlignments) {
  return values.map((row, rowIndex) => ({
    rowNumber: rowIndex + 1,
    cells: row.map((value, columnIndex) => ({
      value: sanitizeWorksheetDisplayValue_(value, rowIndex + 1, columnIndex + 1),
      background: backgrounds[rowIndex][columnIndex],
      fontColor: fontColors[rowIndex][columnIndex],
      fontWeight: fontWeights[rowIndex][columnIndex],
      horizontalAlignment: horizontalAlignments[rowIndex][columnIndex],
      verticalAlignment: verticalAlignments[rowIndex][columnIndex],
      row: rowIndex + 1,
      column: columnIndex + 1,
      a1: toA1_(rowIndex + 1, columnIndex + 1),
    })),
  }));
}

function sanitizeWorksheetDisplayValue_(value, row, column) {
  if (column >= NOTE_START_COLUMN && isExtraNoteHeaderValue_(value, row)) {
    return '';
  }

  return value;
}

function getMergeMeta_(range) {
  return range.getMergedRanges().map((mergedRange) => ({
    row: mergedRange.getRow(),
    column: mergedRange.getColumn(),
    rowSpan: mergedRange.getNumRows(),
    colSpan: mergedRange.getNumColumns(),
    a1: mergedRange.getA1Notation(),
  }));
}

function getSectionLabel_(row) {
  const label = (row[1] || '').trim();
  const hasName = (row[2] || '').trim();
  if (!label || hasName) {
    return '';
  }

  if (
    label === 'เภสัชกร' ||
    label === 'ฝ่ายขายและการตลาด' ||
    label === 'หมายเหตุ' ||
    /^0\d{2}\s/.test(label)
  ) {
    return label;
  }

  return '';
}

function getFontColorHexGrid_(range) {
  const colorObjects = range.getFontColorObjects();
  return colorObjects.map((row) => row.map(colorObjectToString_));
}

function colorObjectToString_(color) {
  if (!color) {
    return '';
  }

  try {
    if (color.getColorType() === SpreadsheetApp.ColorType.RGB) {
      return color.asRgbColor().asHexString();
    }

    if (color.getColorType() === SpreadsheetApp.ColorType.THEME) {
      return `theme:${color.asThemeColor().getThemeColorType()}`;
    }
  } catch (error) {
    return '';
  }

  return '';
}

function toA1_(row, column) {
  let columnName = '';
  let cursor = column;

  while (cursor > 0) {
    const remainder = (cursor - 1) % 26;
    columnName = String.fromCharCode(65 + remainder) + columnName;
    cursor = Math.floor((cursor - 1) / 26);
  }

  return `${columnName}${row}`;
}
