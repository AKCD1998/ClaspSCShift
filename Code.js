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
const EMPLOYEES_SHEET_NAME = 'Employees';
const SHIFT_DICTIONARY_SHEET_NAMES = ['Shift Dictionary', 'รหัสเวร'];
const SCHEDULE_DRAFT_CHANGES_SHEET_NAME = 'Schedule Draft Changes';
const SCHEDULE_AUDIT_LOG_SHEET_NAME = 'Schedule Audit Log';
const PUBLIC_HOLIDAY_BACKGROUND = '#bf9000';
const PUBLIC_HOLIDAY_INK = '#ffffff';
const REGULAR_OFF_BACKGROUNDS = ['#00a9e0', '#00b0f0', '#00ace6'];
const PUBLIC_HOLIDAY_BACKGROUNDS = ['#bf9000', '#c49a00', '#c79a00', '#cc9900'];
const MAX_PUBLIC_HOLIDAY_LEAVE_COUNT = 13;
const MAX_ANNUAL_LEAVE_COUNT = 6;
const PILOT_ATTENDANCE_EMPLOYEE_FULL_NAME = 'ณัฐธนนท์ มานะกุล';
const PILOT_ATTENDANCE_EMPLOYEE_FIRST_NAME = 'ณัฐธนนท์';
const PILOT_ATTENDANCE_EMPLOYEE_NUMBER = '26001';
const PILOT_ATTENDANCE_DEPARTMENT = 'OMO';
const PILOT_ATTENDANCE_MONTH = 4;
const PILOT_ATTENDANCE_COMPARE_END_DATE = '2026-04-28';
const PILOT_ATTENDANCE_EXPORT_CUTOFF = '2026-04-28 11:13';

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
  const normalizedModel = buildNormalizedScheduleModel_(spreadsheet, sheet, values);
  const shiftMetaByCell = buildShiftMetaByCell_(normalizedModel.scheduleEntries);

  return {
    spreadsheetName: spreadsheet.getName(),
    spreadsheetId: spreadsheet.getId(),
    sheetName: sheet.getName(),
    sheetId: sheet.getSheetId(),
    fetchedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    rowCount: values.length,
    columnCount: values[0] ? values[0].length : 0,
    monthOptions: getScheduleMonthOptions_(spreadsheet, targetSheetName),
    shiftDictionarySource: normalizedModel.shiftDictionarySource,
    shiftDictionaryItems: normalizedModel.shiftDictionaryItems,
    rows: buildWorksheetRows_(
      values,
      backgrounds,
      fontColors,
      fontWeights,
      horizontalAlignments,
      verticalAlignments,
      shiftMetaByCell,
    ),
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

function getNormalizedScheduleData(sheetName) {
  const spreadsheet = getTargetSpreadsheet_();
  const targetSheetName = getReadableScheduleSheetName_(spreadsheet, sheetName);
  const sheet = spreadsheet.getSheetByName(targetSheetName);
  assertSheet_(spreadsheet, sheet, targetSheetName);

  const range = sheet.getDataRange();
  const values = range.getDisplayValues();
  const model = buildNormalizedScheduleModel_(spreadsheet, sheet, values);

  console.log(JSON.stringify({
    sheetName: model.sheetName,
    employeeCount: model.employees.length,
    scheduleEntryCount: model.scheduleEntries.length,
    shiftDictionaryItemCount: model.shiftDictionaryItems.length,
  }, null, 2));
  return model;
}

function inspectPilotAttendanceEmployee(sheetName) {
  const spreadsheet = getTargetSpreadsheet_();
  const targetSheetName = getReadableScheduleSheetName_(spreadsheet, sheetName || SHEET_NAME);
  const sheet = spreadsheet.getSheetByName(targetSheetName);
  assertSheet_(spreadsheet, sheet, targetSheetName);

  const values = sheet.getDataRange().getDisplayValues();
  const model = buildNormalizedScheduleModel_(spreadsheet, sheet, values);
  const inspection = resolvePilotAttendanceEmployee_(model.employees);

  console.log(JSON.stringify(inspection, null, 2));
  return inspection;
}

function getPilotAttendancePocData(sheetName) {
  const spreadsheet = getTargetSpreadsheet_();
  const targetSheetName = getReadableScheduleSheetName_(spreadsheet, sheetName || SHEET_NAME);
  const sheet = spreadsheet.getSheetByName(targetSheetName);
  assertSheet_(spreadsheet, sheet, targetSheetName);

  const values = sheet.getDataRange().getDisplayValues();
  const model = buildNormalizedScheduleModel_(spreadsheet, sheet, values);
  const employeeInspection = resolvePilotAttendanceEmployee_(model.employees);
  const expectedSchedule = employeeInspection.reliable
    ? buildPilotExpectedSchedule_(model, employeeInspection.selectedEmployee)
    : [];

  return {
    spreadsheetName: model.spreadsheetName,
    spreadsheetId: model.spreadsheetId,
    sheetName: model.sheetName,
    sheetId: model.sheetId,
    fetchedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    pilot: {
      employeeFullName: PILOT_ATTENDANCE_EMPLOYEE_FULL_NAME,
      employeeFirstName: PILOT_ATTENDANCE_EMPLOYEE_FIRST_NAME,
      employeeNumber: PILOT_ATTENDANCE_EMPLOYEE_NUMBER,
      department: PILOT_ATTENDANCE_DEPARTMENT,
      month: PILOT_ATTENDANCE_MONTH,
      year: MONTH_TEMPLATE_YEAR,
      compareEndDate: PILOT_ATTENDANCE_COMPARE_END_DATE,
      exportCutoff: PILOT_ATTENDANCE_EXPORT_CUTOFF,
    },
    employeeInspection,
    expectedSchedule,
  };
}

function getScheduleRepositorySnapshot() {
  const spreadsheet = getTargetSpreadsheet_();
  const employees = getEmployeesRepository_(spreadsheet);
  const shiftDictionary = getShiftDictionaryRepository_(spreadsheet);
  const draftChanges = getScheduleDraftChangesRepository_(spreadsheet);
  const auditLog = getScheduleAuditLogRepository_(spreadsheet);

  return {
    spreadsheetName: spreadsheet.getName(),
    spreadsheetId: spreadsheet.getId(),
    fetchedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    sources: {
      employees: employees.source,
      shiftDictionary: shiftDictionary.source,
      draftChanges: draftChanges.source,
      auditLog: auditLog.source,
    },
    counts: {
      employees: employees.items.length,
      shiftDictionaryItems: shiftDictionary.items.length,
      draftChanges: draftChanges.items.length,
      auditLog: auditLog.items.length,
    },
    employees: employees.items,
    shiftDictionaryItems: shiftDictionary.items,
    scheduleDraftChanges: draftChanges.items,
    auditLog: auditLog.items,
  };
}

function getEmployeeRepositoryData() {
  return getEmployeesRepository_(getTargetSpreadsheet_());
}

function getShiftDictionaryRepositoryData() {
  return getShiftDictionaryRepository_(getTargetSpreadsheet_());
}

function getScheduleDraftChangeRepositoryData() {
  return getScheduleDraftChangesRepository_(getTargetSpreadsheet_());
}

function getScheduleAuditLogRepositoryData() {
  return getScheduleAuditLogRepository_(getTargetSpreadsheet_());
}

function getEmployeeAnnualLeaveYearCells(request) {
  const spreadsheet = getTargetSpreadsheet_();
  const params = request || {};
  const rowNumber = Number(params.row);
  const targetSheetName = getReadableScheduleSheetName_(spreadsheet, params.sheetName);

  if (!Number.isInteger(rowNumber) || rowNumber < FIRST_SCHEDULE_ROW) {
    throw new Error('A valid employee row is required for annual leave lookup.');
  }

  const cells = [];
  const sourceSheets = [];

  for (let month = 1; month <= 12; month += 1) {
    const monthMeta = getMonthTemplateMeta_(month);
    const sheet = spreadsheet.getSheetByName(monthMeta.sheetName);
    if (!sheet || rowNumber > sheet.getLastRow()) {
      continue;
    }

    const dayNumbers = sheet.getRange(DAY_NUMBER_ROW, DAY_START_COLUMN, 1, MAX_DAY_COLUMNS).getDisplayValues()[0];
    const rowValues = sheet.getRange(rowNumber, DAY_START_COLUMN, 1, MAX_DAY_COLUMNS).getDisplayValues()[0];
    const employeeLabel = getEmployeeLabelFromSheetRow_(sheet, rowNumber);
    sourceSheets.push({
      month,
      sheetName: sheet.getName(),
      available: true,
    });

    for (let offset = 0; offset < MAX_DAY_COLUMNS; offset += 1) {
      const rawCode = String(rowValues[offset] || '').trim();
      if (!isAnnualLeaveRawCode_(rawCode)) {
        continue;
      }

      const dayNumber = Number(dayNumbers[offset]);
      if (!Number.isInteger(dayNumber) || dayNumber < 1 || dayNumber > monthMeta.daysInMonth) {
        continue;
      }

      const column = DAY_START_COLUMN + offset;
      cells.push({
        sheetName: sheet.getName(),
        month,
        year: MONTH_TEMPLATE_YEAR,
        yearBE: MONTH_TEMPLATE_YEAR_BE,
        row: rowNumber,
        column,
        a1: toA1_(rowNumber, column),
        dayNumber,
        workDate: toIsoDate_(MONTH_TEMPLATE_YEAR, month, dayNumber),
        rawCode,
        previousRawCode: rawCode,
        employeeLabel,
        dateLabel: `${dayNumber}`,
        source: {
          type: 'sheet',
          sheetName: sheet.getName(),
        },
      });
    }
  }

  return {
    year: MONTH_TEMPLATE_YEAR,
    yearBE: MONTH_TEMPLATE_YEAR_BE,
    sourceSheetName: targetSheetName,
    row: rowNumber,
    employeeLabel: String(params.employeeLabel || '').trim(),
    maxAnnualLeaveCount: MAX_ANNUAL_LEAVE_COUNT,
    cells,
    sourceSheets,
  };
}

function isAnnualLeaveRawCode_(rawCode) {
  return /^AL(?:[1-6])?$/i.test(String(rawCode || '').trim());
}

function getEmployeePublicHolidayYearCells(request) {
  const spreadsheet = getTargetSpreadsheet_();
  const params = request || {};
  const rowNumber = Number(params.row);
  const targetSheetName = getReadableScheduleSheetName_(spreadsheet, params.sheetName);

  if (!Number.isInteger(rowNumber) || rowNumber < FIRST_SCHEDULE_ROW) {
    throw new Error('A valid employee row is required for public holiday lookup.');
  }

  const cells = [];
  const sourceSheets = [];

  for (let month = 1; month <= 12; month += 1) {
    const monthMeta = getMonthTemplateMeta_(month);
    const sheet = spreadsheet.getSheetByName(monthMeta.sheetName);
    if (!sheet || rowNumber > sheet.getLastRow()) {
      continue;
    }

    const dayNumbers = sheet.getRange(DAY_NUMBER_ROW, DAY_START_COLUMN, 1, MAX_DAY_COLUMNS).getDisplayValues()[0];
    const rowRange = sheet.getRange(rowNumber, DAY_START_COLUMN, 1, MAX_DAY_COLUMNS);
    const rowValues = rowRange.getDisplayValues()[0];
    const rowBackgrounds = rowRange.getBackgrounds()[0];
    const employeeLabel = getEmployeeLabelFromSheetRow_(sheet, rowNumber);
    sourceSheets.push({
      month,
      sheetName: sheet.getName(),
      available: true,
    });

    for (let offset = 0; offset < MAX_DAY_COLUMNS; offset += 1) {
      const rawCode = String(rowValues[offset] || '').trim();
      const background = String(rowBackgrounds[offset] || '').trim();
      if (!isPublicHolidayLeaveRawCode_(rawCode) || !isPublicHolidayLeaveBackground_(background)) {
        continue;
      }

      const dayNumber = Number(dayNumbers[offset]);
      if (!Number.isInteger(dayNumber) || dayNumber < 1 || dayNumber > monthMeta.daysInMonth) {
        continue;
      }

      const column = DAY_START_COLUMN + offset;
      cells.push({
        sheetName: sheet.getName(),
        month,
        year: MONTH_TEMPLATE_YEAR,
        yearBE: MONTH_TEMPLATE_YEAR_BE,
        row: rowNumber,
        column,
        a1: toA1_(rowNumber, column),
        dayNumber,
        workDate: toIsoDate_(MONTH_TEMPLATE_YEAR, month, dayNumber),
        rawCode,
        previousRawCode: rawCode,
        employeeLabel,
        dateLabel: `${dayNumber}`,
        background,
        shiftMeta: createPublicHolidayShiftMeta_(rawCode),
        source: {
          type: 'sheet',
          sheetName: sheet.getName(),
        },
      });
    }
  }

  return {
    year: MONTH_TEMPLATE_YEAR,
    yearBE: MONTH_TEMPLATE_YEAR_BE,
    sourceSheetName: targetSheetName,
    row: rowNumber,
    employeeLabel: String(params.employeeLabel || '').trim(),
    maxPublicHolidayCount: MAX_PUBLIC_HOLIDAY_LEAVE_COUNT,
    cells,
    sourceSheets,
  };
}

function isPublicHolidayLeaveRawCode_(rawCode) {
  return /^Off(?:[1-9]|1[0-3])$/i.test(String(rawCode || '').trim());
}

function isPublicHolidayLeaveBackground_(background) {
  const normalizedBackground = String(background || '').trim().toLowerCase();
  return PUBLIC_HOLIDAY_BACKGROUNDS
    .concat(['#ffff00', '#fff200'])
    .indexOf(normalizedBackground) !== -1;
}

function isPublicHolidayShiftMeta_(shiftMeta) {
  if (!shiftMeta) {
    return false;
  }

  return (
    shiftMeta.segmentType === 'public_holiday' ||
    shiftMeta.displayMode === 'public_holiday' ||
    shiftMeta.leaveType === 'public_holiday'
  );
}

function createPublicHolidayShiftMeta_(rawCode, baseMeta) {
  const code = String(rawCode || '').trim();
  const base = baseMeta || {};
  return {
    rawCode: code,
    displayLabel: code || 'PL',
    displayMode: 'leave',
    roleGroup: base.roleGroup || 'any',
    segmentType: 'public_holiday',
    locationCode: '',
    start: '',
    end: '',
    impliedBaseCode: '',
    color: PUBLIC_HOLIDAY_BACKGROUND,
    description: 'วันหยุดนักขัตฤกษ์',
    leaveType: 'public_holiday',
    dictionarySource: base.dictionarySource || base.source || {
      type: 'computed',
      sheetName: '',
      sourceRow: 0,
    },
  };
}

function getEmployeeLabelFromSheetRow_(sheet, rowNumber) {
  const values = sheet.getRange(rowNumber, 2, 1, 2).getDisplayValues()[0];
  return values.filter(Boolean).join(' ');
}

function saveScheduleChanges(changes) {
  const spreadsheet = getTargetSpreadsheet_();
  const items = Array.isArray(changes) ? changes : [];
  const sheetCache = {};
  const successfulChanges = [];
  const result = {
    successCount: 0,
    failedCount: 0,
    items: [],
    summaryUpdates: [],
  };

  if (!Array.isArray(changes)) {
    return {
      successCount: 0,
      failedCount: 1,
      items: [
        {
          index: -1,
          ok: false,
          status: 'failed',
          message: 'Payload must be an array of schedule changes.',
        },
      ],
    };
  }

  items.forEach((change, index) => {
    try {
      const target = resolveScheduleChangeTarget_(spreadsheet, change, sheetCache);
      const previousRawCode = String(change.previousRawCode || '').trim();
      const nextRawCode = String(change.nextRawCode || '').trim();
      const currentRawCode = target.values[target.row - 1][target.column - 1] || '';

      if (currentRawCode !== previousRawCode) {
        throw new Error(`Cell ${target.a1} changed since review. Current value is "${currentRawCode || '(blank)'}".`);
      }

      const targetRange = target.sheet.getRange(target.row, target.column);
      if (target.actionType === 'delete') {
        targetRange.clearContent();
      } else {
        targetRange.setValue(nextRawCode);
        applyScheduleCellFormatting_(spreadsheet, targetRange, nextRawCode, change.roleGroup, change.nextShiftMeta);
      }

      let auditStatus = 'logged';
      let auditMessage = '';
      try {
        appendScheduleAuditLogRow_(spreadsheet, target, change, previousRawCode, nextRawCode);
      } catch (auditError) {
        auditStatus = 'failed';
        auditMessage = auditError && auditError.message ? auditError.message : String(auditError);
      }

      result.successCount += 1;
      successfulChanges.push({
        change,
        target,
      });
      result.items.push({
        index,
        ok: true,
        status: 'success',
        actionType: target.actionType,
        sheetName: target.sheet.getName(),
        row: target.row,
        column: target.column,
        a1: target.a1,
        previousRawCode,
        nextRawCode,
        auditStatus,
        auditMessage,
        message: auditStatus === 'logged' ? 'Saved.' : `Saved. Audit log failed: ${auditMessage}`,
      });
    } catch (error) {
      result.failedCount += 1;
      result.items.push({
        index,
        ok: false,
        status: 'failed',
        actionType: change && change.actionType ? String(change.actionType) : '',
        sheetName: change && change.sheetName ? String(change.sheetName) : '',
        row: change && change.row ? Number(change.row) : 0,
        column: change && change.column ? Number(change.column) : 0,
        a1: change && change.a1 ? String(change.a1) : '',
        previousRawCode: change && change.previousRawCode !== undefined ? String(change.previousRawCode) : '',
        nextRawCode: change && change.nextRawCode !== undefined ? String(change.nextRawCode) : '',
        message: error && error.message ? error.message : String(error),
      });
    }
  });

  result.summaryUpdates = updateHolidayLeaveSnapshotsForChanges_(spreadsheet, successfulChanges);

  return result;
}

function applyScheduleCellFormatting_(spreadsheet, range, rawCode, roleGroup, clientShiftMeta) {
  const code = String(rawCode || '').trim();
  if (!code) {
    return;
  }

  const shiftDictionary = getShiftDictionaryAdapter_(spreadsheet);
  const dictionaryItem = resolveShiftDictionaryItem_(shiftDictionary.items, code, String(roleGroup || '').trim());
  const colorSource = isPublicHolidayShiftMeta_(clientShiftMeta)
    ? createPublicHolidayShiftMeta_(code, clientShiftMeta)
    : dictionaryItem;
  const colorScheme = getScheduleCellColorScheme_(colorSource);
  if (!colorScheme.background) {
    return;
  }

  range
    .setBackground(colorScheme.background)
    .setFontColor(colorScheme.ink)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
}

function updateHolidayLeaveSnapshotsForChanges_(spreadsheet, successfulChanges) {
  const updatesBySheetAndRow = new Map();
  const updates = [];

  successfulChanges.forEach((item) => {
    const change = item.change || {};
    const target = item.target || {};
    const summarySheetName = getHolidayLeaveSummarySheetName_(spreadsheet, change, target);
    const row = Number(target.row || change.row);

    if (!summarySheetName || !Number.isInteger(row) || row < FIRST_SCHEDULE_ROW) {
      return;
    }

    const key = `${summarySheetName}:${row}`;
    if (!updatesBySheetAndRow.has(key)) {
      updatesBySheetAndRow.set(key, {
        sheetName: summarySheetName,
        row,
      });
    }
  });

  updatesBySheetAndRow.forEach((item) => {
    try {
      const sheet = spreadsheet.getSheetByName(item.sheetName);
      if (!sheet) {
        throw new Error(`Sheet "${item.sheetName}" was not found.`);
      }

      const snapshot = calculateHolidayLeaveSnapshot_(spreadsheet, sheet, item.row);
      const holidayLeaveColumn = snapshot.holidayLeaveColumn;
      if (!holidayLeaveColumn) {
        throw new Error(`Could not find "${HOLIDAY_LEAVE_LABEL}" column.`);
      }

      const range = sheet.getRange(item.row, holidayLeaveColumn);
      range
        .setValue(snapshot.displayValue)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');

      updates.push({
        ok: true,
        sheetName: sheet.getName(),
        row: item.row,
        column: holidayLeaveColumn,
        a1: toA1_(item.row, holidayLeaveColumn),
        displayValue: snapshot.displayValue,
        publicHolidayUsed: snapshot.publicHolidayUsed,
        annualLeaveUsed: snapshot.annualLeaveUsed,
        publicHolidayRemaining: snapshot.publicHolidayRemaining,
        annualLeaveRemaining: snapshot.annualLeaveRemaining,
      });
    } catch (error) {
      updates.push({
        ok: false,
        sheetName: item.sheetName,
        row: item.row,
        message: error && error.message ? error.message : String(error),
      });
    }
  });

  return updates;
}

function getHolidayLeaveSummarySheetName_(spreadsheet, change, target) {
  const sourceSheetName = String(change.sourceSheetName || change.summarySheetName || '').trim();
  if (sourceSheetName) {
    return getReadableScheduleSheetName_(spreadsheet, sourceSheetName);
  }

  if (target && target.sheet) {
    return target.sheet.getName();
  }

  return '';
}

function calculateHolidayLeaveSnapshot_(spreadsheet, summarySheet, rowNumber) {
  const summaryMonth = getMonthNumberFromSheetName_(summarySheet.getName());
  if (!summaryMonth) {
    throw new Error(`Cannot resolve month from sheet "${summarySheet.getName()}".`);
  }

  const counts = countEmployeeYearLeaveUsageThroughMonth_(spreadsheet, rowNumber, summaryMonth);
  const publicHolidayRemaining = Math.max(MAX_PUBLIC_HOLIDAY_LEAVE_COUNT - counts.publicHolidayUsed, 0);
  const annualLeaveRemaining = Math.max(MAX_ANNUAL_LEAVE_COUNT - counts.annualLeaveUsed, 0);

  return {
    holidayLeaveColumn: findHolidayLeaveColumn_(summarySheet) || HOLIDAY_LEAVE_COLUMN,
    publicHolidayUsed: counts.publicHolidayUsed,
    annualLeaveUsed: counts.annualLeaveUsed,
    publicHolidayRemaining,
    annualLeaveRemaining,
    displayValue: `${publicHolidayRemaining}/${annualLeaveRemaining}`,
  };
}

function countEmployeeYearLeaveUsageThroughMonth_(spreadsheet, rowNumber, throughMonth) {
  let publicHolidayUsed = 0;
  let annualLeaveUsed = 0;

  for (let month = 1; month <= throughMonth; month += 1) {
    const monthMeta = getMonthTemplateMeta_(month);
    const sheet = spreadsheet.getSheetByName(monthMeta.sheetName);
    if (!sheet || rowNumber > sheet.getLastRow()) {
      continue;
    }

    const dayNumbers = sheet.getRange(DAY_NUMBER_ROW, DAY_START_COLUMN, 1, MAX_DAY_COLUMNS).getDisplayValues()[0];
    const rowRange = sheet.getRange(rowNumber, DAY_START_COLUMN, 1, MAX_DAY_COLUMNS);
    const rowValues = rowRange.getDisplayValues()[0];
    const rowBackgrounds = rowRange.getBackgrounds()[0];

    for (let offset = 0; offset < MAX_DAY_COLUMNS; offset += 1) {
      const dayNumber = Number(dayNumbers[offset]);
      if (!Number.isInteger(dayNumber) || dayNumber < 1 || dayNumber > monthMeta.daysInMonth) {
        continue;
      }

      const rawCode = String(rowValues[offset] || '').trim();
      const background = String(rowBackgrounds[offset] || '').trim();
      const annualLeaveOrdinal = getAnnualLeaveOrdinal_(rawCode);
      if (annualLeaveOrdinal) {
        annualLeaveUsed = Math.max(annualLeaveUsed, annualLeaveOrdinal);
        continue;
      }

      const publicHolidayOrdinal = getPublicHolidayLeaveOrdinal_(rawCode, background);
      if (publicHolidayOrdinal) {
        publicHolidayUsed = Math.max(publicHolidayUsed, publicHolidayOrdinal);
      }
    }
  }

  return {
    publicHolidayUsed,
    annualLeaveUsed,
  };
}

function getAnnualLeaveOrdinal_(rawCode) {
  const match = String(rawCode || '').trim().match(/^AL([1-6])?$/i);
  if (!match) {
    return 0;
  }

  return match[1] ? Number(match[1]) : 1;
}

function getPublicHolidayLeaveOrdinal_(rawCode, background) {
  if (!isPublicHolidayLeaveBackground_(background)) {
    return 0;
  }

  const match = String(rawCode || '').trim().match(/^Off([1-9]|1[0-3])$/i);
  return match ? Number(match[1]) : 0;
}

function getScheduleCellColorScheme_(shiftMeta) {
  if (isPublicHolidayShiftMeta_(shiftMeta)) {
    return {
      background: isHexColor_(shiftMeta.color) ? shiftMeta.color : PUBLIC_HOLIDAY_BACKGROUND,
      ink: PUBLIC_HOLIDAY_INK,
    };
  }

  if (shiftMeta && isHexColor_(shiftMeta.color)) {
    return {
      background: shiftMeta.color,
      ink: readableTextColorForHex_(shiftMeta.color),
    };
  }

  const rawCode = String((shiftMeta && shiftMeta.rawCode) || '').trim();
  const upperCode = rawCode.toUpperCase();
  const displayMode = shiftMeta && shiftMeta.displayMode ? shiftMeta.displayMode : '';
  const segmentType = shiftMeta && shiftMeta.segmentType ? shiftMeta.segmentType : '';

  if (/^OFF\d*$/i.test(rawCode)) {
    return { background: '#00a9e0', ink: '#ffffff' };
  }

  if (/^(AL|BL|ML|PL|SL)\d*$/i.test(rawCode)) {
    return { background: '#ff1010', ink: '#ffffff' };
  }

  if (upperCode === 'RXM') {
    return { background: '#075be8', ink: '#ffffff' };
  }

  if (upperCode === 'HM') {
    return { background: '#6b3fa0', ink: '#ffffff' };
  }

  if (upperCode === 'PRODUCT TRAINING') {
    return { background: '#e77bea', ink: '#111827' };
  }

  if (upperCode === 'OFFICE' || upperCode === 'O-1' || segmentType === 'office') {
    return { background: '#fff200', ink: '#101820' };
  }

  if (/^1\/3H$/i.test(rawCode)) {
    return { background: '#2f5f25', ink: '#ffffff' };
  }

  if (/^1[123]H/i.test(rawCode) || upperCode === '12H') {
    return { background: '#00ff00', ink: '#002b18' };
  }

  if (/^\d+H-\d+$/i.test(rawCode) || /^3H$/i.test(rawCode) || /^4H$/i.test(rawCode)) {
    return { background: '#002060', ink: '#ffffff' };
  }

  if (upperCode === '10T') {
    return { background: '#fff200', ink: '#101820' };
  }

  if (upperCode === '0T') {
    return { background: '#dff0e8', ink: '#063f2c' };
  }

  if (rawCode === '1' || rawCode === '2' || displayMode === 'single' || segmentType === 'regular') {
    return { background: '#dbeaf1', ink: '#063f2c' };
  }

  return { background: '#e7f2ed', ink: '#164f3b' };
}

function isHexColor_(color) {
  return /^#[0-9a-f]{6}$/i.test(String(color || ''));
}

function readableTextColorForHex_(background) {
  const rgb = parseHexColor_(background);
  if (!rgb) {
    return '#1f2933';
  }

  return relativeLuminance_(rgb) > 0.54 ? '#17202a' : '#ffffff';
}

function parseHexColor_(hex) {
  if (!isHexColor_(hex)) {
    return null;
  }

  return {
    r: parseInt(hex.slice(1, 3), 16),
    g: parseInt(hex.slice(3, 5), 16),
    b: parseInt(hex.slice(5, 7), 16),
  };
}

function relativeLuminance_(rgb) {
  const convert = (value) => {
    const channel = value / 255;
    return channel <= 0.03928
      ? channel / 12.92
      : Math.pow((channel + 0.055) / 1.055, 2.4);
  };

  return 0.2126 * convert(rgb.r) + 0.7152 * convert(rgb.g) + 0.0722 * convert(rgb.b);
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

function repairExtraNoteHeaderValues2026() {
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

    repaired.push({
      sheetName: sheet.getName(),
      daysInMonth: monthMeta.daysInMonth,
      clearedExtraNoteHeaders: clearExtraNoteHeaderValues_(sheet),
    });
  }

  const result = {
    repaired,
    skipped,
    repairedCount: repaired.length,
    skippedCount: skipped.length,
    clearedCount: repaired.reduce((total, item) => total + item.clearedExtraNoteHeaders, 0),
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

function resolveScheduleChangeTarget_(spreadsheet, change, sheetCache) {
  validateScheduleChangeShape_(change);

  const sheetName = getReadableScheduleSheetName_(spreadsheet, change.sheetName);
  const cache = getScheduleSheetCache_(spreadsheet, sheetName, sheetCache);
  const row = Number(change.row);
  const column = Number(change.column);
  const a1 = toA1_(row, column);

  if (change.a1 && String(change.a1) !== a1) {
    throw new Error(`A1 mismatch. Payload says ${change.a1}, resolved target is ${a1}.`);
  }

  validateScheduleEditTarget_(cache.values, row, column);

  return {
    sheet: cache.sheet,
    values: cache.values,
    row,
    column,
    a1,
    actionType: normalizeScheduleChangeAction_(change),
  };
}

function validateScheduleChangeShape_(change) {
  if (!change || typeof change !== 'object') {
    throw new Error('Change item must be an object.');
  }

  if (!change.sheetName) {
    throw new Error('Missing sheetName.');
  }

  if (!Number.isInteger(Number(change.row)) || Number(change.row) < FIRST_SCHEDULE_ROW) {
    throw new Error('Invalid row.');
  }

  if (!Number.isInteger(Number(change.column))) {
    throw new Error('Invalid column.');
  }

  normalizeScheduleChangeAction_(change);

  if (change.previousRawCode === undefined || change.nextRawCode === undefined) {
    throw new Error('Missing previousRawCode or nextRawCode.');
  }
}

function normalizeScheduleChangeAction_(change) {
  const requestedActionType = String(change.actionType || '').trim();
  if (['update', 'clear', 'delete', 'create'].indexOf(requestedActionType) === -1) {
    const actionType = requestedActionType;
    throw new Error(`Unsupported actionType "${actionType}".`);
  }

  const actionType = requestedActionType === 'clear' ? 'delete' : requestedActionType;
  const previousRawCode = String(change.previousRawCode || '').trim();
  const nextRawCode = String(change.nextRawCode || '').trim();
  if (actionType === 'delete' && nextRawCode) {
    throw new Error('Delete action must have an empty nextRawCode.');
  }

  if (actionType === 'delete' && !previousRawCode) {
    throw new Error('Delete action must have a previousRawCode.');
  }

  if (actionType === 'create' && (previousRawCode || !nextRawCode)) {
    throw new Error('Create action must have an empty previousRawCode and a nextRawCode.');
  }

  if (actionType === 'update' && (!previousRawCode || !nextRawCode)) {
    throw new Error('Update action must have both previousRawCode and nextRawCode.');
  }

  return actionType;
}

function getScheduleSheetCache_(spreadsheet, sheetName, sheetCache) {
  if (sheetCache[sheetName]) {
    return sheetCache[sheetName];
  }

  const sheet = spreadsheet.getSheetByName(sheetName);
  assertSheet_(spreadsheet, sheet, sheetName);
  sheetCache[sheetName] = {
    sheet,
    values: sheet.getDataRange().getDisplayValues(),
  };
  return sheetCache[sheetName];
}

function validateScheduleEditTarget_(values, row, column) {
  if (column < DAY_START_COLUMN || column > DAY_END_COLUMN) {
    throw new Error('Target column is outside the schedule day grid.');
  }

  const dayNumber = Number((values[DAY_NUMBER_ROW - 1] || [])[column - 1]);
  if (!Number.isInteger(dayNumber) || dayNumber < 1 || dayNumber > MAX_DAY_COLUMNS) {
    throw new Error('Target column does not have a valid day header.');
  }

  const targetRow = values[row - 1];
  if (!targetRow) {
    throw new Error('Target row is outside the sheet data range.');
  }

  const sequence = (targetRow[1] || '').trim();
  const employeeName = (targetRow[2] || '').trim();
  if (!/^\d+$/.test(sequence) || !employeeName) {
    throw new Error('Target row is not a recognized employee schedule row.');
  }

  if (getSectionForSheetRow_(values, row) === 'หมายเหตุ') {
    throw new Error('Target row belongs to the note section, not the schedule table.');
  }
}

function getSectionForSheetRow_(values, rowNumber) {
  let section = '';

  for (let rowIndex = 0; rowIndex < rowNumber && rowIndex < values.length; rowIndex += 1) {
    const sectionLabel = getSectionLabel_(values[rowIndex]);
    if (sectionLabel) {
      section = sectionLabel;
    }
  }

  return section;
}

function appendScheduleAuditLogRow_(spreadsheet, target, change, previousRawCode, nextRawCode) {
  const auditSheet = getOrCreateScheduleAuditLogSheet_(spreadsheet);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  const source = getScheduleAuditSource_(change);

  auditSheet.appendRow([
    timestamp,
    getCurrentUserLabel_(),
    target.sheet.getName(),
    getScheduleAuditEmployeeLabel_(target, change),
    getScheduleAuditDateLabel_(target, change),
    target.row,
    target.column,
    target.a1,
    previousRawCode,
    nextRawCode,
    target.actionType,
    source.page,
    source.mode,
  ]);
}

function getOrCreateScheduleAuditLogSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(SCHEDULE_AUDIT_LOG_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(SCHEDULE_AUDIT_LOG_SHEET_NAME);
  }

  const headers = [
    'timestamp',
    'userLabel',
    'sheetName',
    'employeeLabel',
    'dateLabel',
    'row',
    'column',
    'a1',
    'oldRawCode',
    'newRawCode',
    'actionType',
    'sourcePage',
    'sourceMode',
  ];
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  const existingHeaders = headerRange.getDisplayValues()[0];
  const hasHeaders = existingHeaders.some((value) => String(value || '').trim());
  if (!hasHeaders) {
    headerRange.setValues([headers]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function getCurrentUserLabel_() {
  try {
    const activeUser = Session.getActiveUser();
    return activeUser ? activeUser.getEmail() : '';
  } catch (error) {
    return '';
  }
}

function getScheduleAuditEmployeeLabel_(target, change) {
  if (change && change.employeeLabel) {
    return String(change.employeeLabel).trim();
  }

  return ((target.values[target.row - 1] || [])[2] || '').trim();
}

function getScheduleAuditDateLabel_(target, change) {
  if (change && change.dateLabel) {
    return String(change.dateLabel).trim();
  }

  const day = ((target.values[DAY_NUMBER_ROW - 1] || [])[target.column - 1] || '').trim();
  const weekday = ((target.values[WEEKDAY_ROW - 1] || [])[target.column - 1] || '').trim();
  return [day, weekday].filter(String).join(' ');
}

function getScheduleAuditSource_(change) {
  const sourcePage = change && change.sourcePage ? String(change.sourcePage).trim() : 'schedule-web-app';
  const sourceMode = change && change.sourceMode ? String(change.sourceMode).trim() : 'edit-mode';

  return {
    page: sourcePage,
    mode: sourceMode,
  };
}

function getEmployeesRepository_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(EMPLOYEES_SHEET_NAME);
  if (!sheet) {
    return createMissingRepositoryResult_('employees', [EMPLOYEES_SHEET_NAME]);
  }

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return createSheetRepositoryResult_('employees', sheet, []);
  }

  const headerMap = buildRepositoryHeaderMap_(values[0], {
    employeeId: ['employeeid', 'employee_id', 'id', 'รหัสพนักงาน'],
    employeeName: ['employeename', 'employee_name', 'name', 'fullname', 'full_name', 'ชื่อ', 'ชื่อพนักงาน'],
    roleGroup: ['rolegroup', 'role_group', 'role', 'group', 'กลุ่มงาน', 'สายงาน'],
    active: ['active', 'enabled', 'status', 'สถานะ', 'ใช้งาน'],
  });

  const items = values.slice(1).reduce((employees, row, rowIndex) => {
    const employeeName = getRepositoryCellValue_(row, headerMap, 'employeeName');
    const providedEmployeeId = getRepositoryCellValue_(row, headerMap, 'employeeId');
    if (!employeeName && !providedEmployeeId) {
      return employees;
    }

    const sourceRow = rowIndex + 2;
    employees.push({
      employeeId: providedEmployeeId || `employee:${sheet.getSheetId()}:${sourceRow}`,
      employeeName,
      roleGroup: getRepositoryCellValue_(row, headerMap, 'roleGroup') || 'unknown',
      active: parseRepositoryBoolean_(getRepositoryCellValue_(row, headerMap, 'active'), true),
      source: {
        type: 'sheet',
        sheetName: sheet.getName(),
        sourceRow,
      },
    });
    return employees;
  }, []);

  return createSheetRepositoryResult_('employees', sheet, items);
}

function getShiftDictionaryRepository_(spreadsheet) {
  const sheet = getOptionalSheetByNames_(spreadsheet, SHIFT_DICTIONARY_SHEET_NAMES);
  if (!sheet) {
    return createMissingRepositoryResult_('shiftDictionary', SHIFT_DICTIONARY_SHEET_NAMES);
  }

  return createSheetRepositoryResult_('shiftDictionary', sheet, readShiftDictionaryItemsFromSheet_(sheet));
}

function getScheduleDraftChangesRepository_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SCHEDULE_DRAFT_CHANGES_SHEET_NAME);
  if (!sheet) {
    return createMissingRepositoryResult_('scheduleDraftChanges', [SCHEDULE_DRAFT_CHANGES_SHEET_NAME]);
  }

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return createSheetRepositoryResult_('scheduleDraftChanges', sheet, []);
  }

  const headerMap = buildRepositoryHeaderMap_(values[0], {
    draftId: ['draftid', 'draft_id', 'id', 'รหัสดราฟต์'],
    employeeId: ['employeeid', 'employee_id', 'รหัสพนักงาน'],
    employeeName: ['employeename', 'employee_name', 'name', 'ชื่อพนักงาน'],
    workDate: ['workdate', 'work_date', 'date', 'วันที่'],
    oldRawCode: ['oldrawcode', 'old_raw_code', 'previousrawcode', 'previous_raw_code', 'oldcode', 'old_code', 'รหัสเดิม'],
    newRawCode: ['newrawcode', 'new_raw_code', 'nextrawcode', 'next_raw_code', 'newcode', 'new_code', 'รหัสใหม่'],
    status: ['status', 'สถานะ'],
    actionType: ['actiontype', 'action_type', 'action', 'ประเภท'],
    sheetName: ['sheetname', 'sheet_name', 'worksheet', 'ชีต'],
    row: ['row', 'แถว'],
    column: ['column', 'col', 'คอลัมน์'],
    a1: ['a1', 'cell', 'cellref', 'cell_ref', 'เซลล์'],
  });

  const items = values.slice(1).reduce((drafts, row, rowIndex) => {
    const employeeId = getRepositoryCellValue_(row, headerMap, 'employeeId');
    const workDate = getRepositoryCellValue_(row, headerMap, 'workDate');
    const oldRawCode = getRepositoryCellValue_(row, headerMap, 'oldRawCode');
    const newRawCode = getRepositoryCellValue_(row, headerMap, 'newRawCode');
    if (!employeeId && !workDate && !oldRawCode && !newRawCode) {
      return drafts;
    }

    const sourceRow = rowIndex + 2;
    drafts.push({
      draftId: getRepositoryCellValue_(row, headerMap, 'draftId') || `draft:${sheet.getSheetId()}:${sourceRow}`,
      employeeId,
      employeeName: getRepositoryCellValue_(row, headerMap, 'employeeName'),
      workDate,
      oldRawCode,
      newRawCode,
      status: getRepositoryCellValue_(row, headerMap, 'status') || 'draft',
      actionType: getRepositoryCellValue_(row, headerMap, 'actionType'),
      sheetName: getRepositoryCellValue_(row, headerMap, 'sheetName'),
      row: getRepositoryCellValue_(row, headerMap, 'row'),
      column: getRepositoryCellValue_(row, headerMap, 'column'),
      a1: getRepositoryCellValue_(row, headerMap, 'a1'),
      source: {
        type: 'sheet',
        sheetName: sheet.getName(),
        sourceRow,
      },
    });
    return drafts;
  }, []);

  return createSheetRepositoryResult_('scheduleDraftChanges', sheet, items);
}

function getScheduleAuditLogRepository_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(SCHEDULE_AUDIT_LOG_SHEET_NAME);
  if (!sheet) {
    return createMissingRepositoryResult_('scheduleAuditLog', [SCHEDULE_AUDIT_LOG_SHEET_NAME]);
  }

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return createSheetRepositoryResult_('scheduleAuditLog', sheet, []);
  }

  const headerMap = buildRepositoryHeaderMap_(values[0], {
    auditId: ['auditid', 'audit_id', 'id'],
    timestamp: ['timestamp', 'createdat', 'created_at', 'เวลา', 'วันที่เวลา'],
    userLabel: ['userlabel', 'user_label', 'user', 'email', 'ผู้ใช้'],
    sheetName: ['sheetname', 'sheet_name', 'worksheet', 'ชีต'],
    employeeLabel: ['employeelabel', 'employee_label', 'employee', 'พนักงาน'],
    dateLabel: ['datelabel', 'date_label', 'date', 'วันที่'],
    row: ['row', 'แถว'],
    column: ['column', 'col', 'คอลัมน์'],
    a1: ['a1', 'cell', 'cellref', 'cell_ref', 'เซลล์'],
    oldRawCode: ['oldrawcode', 'old_raw_code', 'previousrawcode', 'previous_raw_code', 'oldcode', 'old_code', 'รหัสเดิม'],
    newRawCode: ['newrawcode', 'new_raw_code', 'nextrawcode', 'next_raw_code', 'newcode', 'new_code', 'รหัสใหม่'],
    actionType: ['actiontype', 'action_type', 'action', 'ประเภท'],
    sourcePage: ['sourcepage', 'source_page', 'page', 'หน้า'],
    sourceMode: ['sourcemode', 'source_mode', 'mode', 'โหมด'],
  });

  const items = values.slice(1).reduce((auditItems, row, rowIndex) => {
    const timestamp = getRepositoryCellValue_(row, headerMap, 'timestamp');
    const actionType = getRepositoryCellValue_(row, headerMap, 'actionType');
    const oldRawCode = getRepositoryCellValue_(row, headerMap, 'oldRawCode');
    const newRawCode = getRepositoryCellValue_(row, headerMap, 'newRawCode');
    if (!timestamp && !actionType && !oldRawCode && !newRawCode) {
      return auditItems;
    }

    const sourceRow = rowIndex + 2;
    auditItems.push({
      auditId: getRepositoryCellValue_(row, headerMap, 'auditId') || `audit:${sheet.getSheetId()}:${sourceRow}`,
      timestamp,
      userLabel: getRepositoryCellValue_(row, headerMap, 'userLabel'),
      sheetName: getRepositoryCellValue_(row, headerMap, 'sheetName'),
      employeeLabel: getRepositoryCellValue_(row, headerMap, 'employeeLabel'),
      dateLabel: getRepositoryCellValue_(row, headerMap, 'dateLabel'),
      row: getRepositoryCellValue_(row, headerMap, 'row'),
      column: getRepositoryCellValue_(row, headerMap, 'column'),
      a1: getRepositoryCellValue_(row, headerMap, 'a1'),
      oldRawCode,
      newRawCode,
      actionType,
      sourcePage: getRepositoryCellValue_(row, headerMap, 'sourcePage'),
      sourceMode: getRepositoryCellValue_(row, headerMap, 'sourceMode'),
      source: {
        type: 'sheet',
        sheetName: sheet.getName(),
        sourceRow,
      },
    });
    return auditItems;
  }, []);

  return createSheetRepositoryResult_('scheduleAuditLog', sheet, items);
}

function getOptionalSheetByNames_(spreadsheet, sheetNames) {
  for (let index = 0; index < sheetNames.length; index += 1) {
    const sheet = spreadsheet.getSheetByName(sheetNames[index]);
    if (sheet) {
      return sheet;
    }
  }

  return null;
}

function createMissingRepositoryResult_(entityName, expectedSheetNames) {
  return {
    entityName,
    source: {
      type: 'missing',
      sheetName: '',
      expectedSheetNames,
      message: `Optional repository sheet was not found: ${expectedSheetNames.join(' / ')}`,
    },
    items: [],
  };
}

function createSheetRepositoryResult_(entityName, sheet, items) {
  return {
    entityName,
    source: {
      type: 'sheet',
      sheetName: sheet.getName(),
      sheetId: sheet.getSheetId(),
      message: `Loaded ${items.length} ${entityName} item(s).`,
    },
    items,
  };
}

function buildRepositoryHeaderMap_(headerRow, aliases) {
  const headerMap = {};

  headerRow.forEach((header, index) => {
    const normalizedHeader = normalizeDictionaryHeader_(header);
    Object.keys(aliases).forEach((field) => {
      if (headerMap[field] === undefined && aliases[field].indexOf(normalizedHeader) !== -1) {
        headerMap[field] = index;
      }
    });
  });

  return headerMap;
}

function getRepositoryCellValue_(row, headerMap, field) {
  const index = headerMap[field];
  if (index === undefined) {
    return '';
  }

  return String(row[index] || '').trim();
}

function parseRepositoryBoolean_(value, defaultValue) {
  const normalizedValue = String(value || '').trim().toLowerCase();
  if (!normalizedValue) {
    return defaultValue;
  }

  if (['true', 'yes', 'y', '1', 'active', 'enabled', 'ใช้งาน'].indexOf(normalizedValue) !== -1) {
    return true;
  }

  if (['false', 'no', 'n', '0', 'inactive', 'disabled', 'ไม่ใช้งาน', 'ยกเลิก'].indexOf(normalizedValue) !== -1) {
    return false;
  }

  return defaultValue;
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

function buildNormalizedScheduleModel_(spreadsheet, sheet, values) {
  const shiftDictionary = getShiftDictionaryAdapter_(spreadsheet);
  const shiftDictionaryItems = shiftDictionary.items;
  const employees = buildEmployeeModels_(sheet, values);
  const scheduleEntries = buildScheduleEntryModels_(sheet, values, employees, shiftDictionaryItems);

  return {
    spreadsheetName: spreadsheet.getName(),
    spreadsheetId: spreadsheet.getId(),
    sheetName: sheet.getName(),
    sheetId: sheet.getSheetId(),
    roleGroups: getRoleGroupItems_(),
    shiftDictionarySource: shiftDictionary.source,
    shiftDictionaryItems,
    employees,
    scheduleEntries,
  };
}

function buildEmployeeModels_(sheet, values) {
  const employees = [];
  let section = '';

  values.forEach((row, rowIndex) => {
    const sectionLabel = getSectionLabel_(row);
    if (sectionLabel) {
      section = sectionLabel;
      return;
    }

    const sequence = (row[1] || '').trim();
    const employeeName = (row[2] || '').trim();
    if (!/^\d+$/.test(sequence) || !employeeName || section === 'หมายเหตุ') {
      return;
    }

    const sourceRow = rowIndex + 1;
    const position = row[5] || '';
    const team = row[6] || '';
    const roleGroup = inferRoleGroup_(section, position, team);

    employees.push({
      employeeId: `employee:${sheet.getSheetId()}:${sourceRow}`,
      employeeName,
      sequence,
      nickname: row[3] || '',
      birthDate: row[4] || '',
      position,
      team,
      section,
      roleGroup,
      sourceRow,
      sourceColumns: {
        sequence: 2,
        employeeName: 3,
        nickname: 4,
        birthDate: 5,
        position: 6,
        team: 7,
      },
    });
  });

  return employees;
}

function buildScheduleEntryModels_(sheet, values, employees, shiftDictionaryItems) {
  const entries = [];
  const monthNumber = getMonthNumberFromSheetName_(sheet.getName());
  const dayHeaderRow = values[DAY_NUMBER_ROW - 1] || [];
  const weekdayHeaderRow = values[WEEKDAY_ROW - 1] || [];

  employees.forEach((employee) => {
    const row = values[employee.sourceRow - 1] || [];

    for (let column = DAY_START_COLUMN; column <= DAY_END_COLUMN; column += 1) {
      const dayNumber = Number(dayHeaderRow[column - 1]);
      if (!Number.isInteger(dayNumber) || dayNumber < 1 || dayNumber > MAX_DAY_COLUMNS) {
        continue;
      }

      const rawCode = row[column - 1] || '';
      const dictionaryItem = resolveShiftDictionaryItem_(shiftDictionaryItems, rawCode, employee.roleGroup);
      const a1 = toA1_(employee.sourceRow, column);

      entries.push({
        scheduleEntryId: `scheduleEntry:${sheet.getSheetId()}:${employee.sourceRow}:${column}`,
        employeeId: employee.employeeId,
        employeeName: employee.employeeName,
        roleGroup: employee.roleGroup,
        workDate: monthNumber ? toIsoDate_(MONTH_TEMPLATE_YEAR, monthNumber, dayNumber) : '',
        dayNumber,
        weekday: weekdayHeaderRow[column - 1] || '',
        rawCode,
        shiftDictionaryItemId: dictionaryItem.shiftDictionaryItemId,
        displayLabel: dictionaryItem.displayLabel,
        displayMode: dictionaryItem.displayMode,
        segmentType: dictionaryItem.segmentType,
        locationCode: dictionaryItem.locationCode,
        start: dictionaryItem.start,
        breakStart: dictionaryItem.breakStart,
        breakEnd: dictionaryItem.breakEnd,
        end: dictionaryItem.end,
        expectedWorkMinutes: dictionaryItem.expectedWorkMinutes,
        impliedBaseCode: dictionaryItem.impliedBaseCode,
        color: dictionaryItem.color,
        description: dictionaryItem.description,
        dictionarySource: dictionaryItem.source,
        sheetName: sheet.getName(),
        sheetId: sheet.getSheetId(),
        sourceRow: employee.sourceRow,
        sourceColumn: column,
        a1,
      });
    }
  });

  return entries;
}

function resolvePilotAttendanceEmployee_(employees) {
  const exactMatches = employees.filter((employee) =>
    normalizeHumanName_(employee.employeeName) === normalizeHumanName_(PILOT_ATTENDANCE_EMPLOYEE_FULL_NAME)
  );
  const firstNameMatches = employees.filter((employee) =>
    normalizeHumanName_(employee.employeeName).indexOf(normalizeHumanName_(PILOT_ATTENDANCE_EMPLOYEE_FIRST_NAME)) !== -1
  );
  const selectedEmployee = exactMatches.length === 1
    ? exactMatches[0]
    : exactMatches.length === 0 && firstNameMatches.length === 1
      ? firstNameMatches[0]
      : null;
  const matchMethod = exactMatches.length === 1
    ? 'exact_full_name'
    : selectedEmployee
      ? 'first_name'
      : 'none';

  return {
    targetFullName: PILOT_ATTENDANCE_EMPLOYEE_FULL_NAME,
    targetFirstName: PILOT_ATTENDANCE_EMPLOYEE_FIRST_NAME,
    reliable: Boolean(selectedEmployee),
    matchMethod,
    selectedEmployee: selectedEmployee ? createPilotEmployeeInspectionItem_(selectedEmployee) : null,
    exactMatchCount: exactMatches.length,
    firstNameMatchCount: firstNameMatches.length,
    candidates: firstNameMatches.map(createPilotEmployeeInspectionItem_),
    message: selectedEmployee
      ? `Pilot employee resolved by ${matchMethod}.`
      : 'Pilot employee row could not be resolved reliably. Inspect candidates before comparing.',
  };
}

function createPilotEmployeeInspectionItem_(employee) {
  return {
    employeeId: employee.employeeId,
    employeeName: employee.employeeName,
    sequence: employee.sequence,
    nickname: employee.nickname,
    position: employee.position,
    team: employee.team,
    section: employee.section,
    roleGroup: employee.roleGroup,
    sourceRow: employee.sourceRow,
  };
}

function normalizeHumanName_(name) {
  return String(name || '').replace(/\s+/g, '').trim();
}

function buildPilotExpectedSchedule_(model, employeeInspectionItem) {
  const employeeId = employeeInspectionItem && employeeInspectionItem.employeeId;
  const expectedMonthPrefix = `${MONTH_TEMPLATE_YEAR}-${pad2_(PILOT_ATTENDANCE_MONTH)}-`;

  return model.scheduleEntries
    .filter((entry) => entry.employeeId === employeeId && String(entry.workDate || '').indexOf(expectedMonthPrefix) === 0)
    .sort((left, right) => String(left.workDate).localeCompare(String(right.workDate)))
    .map((entry) => {
      const timing = getShiftTimingWithBreak_(entry);

      return {
        expectedWorkDate: entry.workDate,
        employeeName: entry.employeeName,
        rawCode: entry.rawCode,
        expectedStart: timing.expectedStart,
        breakStart: timing.breakStart,
        breakEnd: timing.breakEnd,
        expectedEnd: timing.expectedEnd,
        expectedWorkMinutes: timing.expectedWorkMinutes,
        expectedStatus: timing.expectedStatus,
        sourceRow: entry.sourceRow,
        sourceColumn: entry.sourceColumn,
        a1: entry.a1,
      };
    });
}

function getShiftTimingWithBreak_(entry) {
  const rawCode = String((entry && entry.rawCode) || '').trim();
  if (!rawCode) {
    return createExpectedShiftTiming_('BLANK', '', '', '', '', 0);
  }

  if (/^off(?:\d+)?$/i.test(rawCode)) {
    return createExpectedShiftTiming_('OFF', '', '', '', '', 0);
  }

  if (rawCode === 'O-1') {
    return createExpectedShiftTiming_(
      'WORK',
      entry.start || '08:00',
      entry.breakStart || '12:00',
      entry.breakEnd || '13:00',
      entry.end || '17:00',
      entry.expectedWorkMinutes || 480,
    );
  }

  return createExpectedShiftTiming_('UNSUPPORTED', '', '', '', '', 0);
}

function createExpectedShiftTiming_(expectedStatus, expectedStart, breakStart, breakEnd, expectedEnd, expectedWorkMinutes) {
  return {
    expectedStatus,
    expectedStart,
    breakStart,
    breakEnd,
    expectedEnd,
    expectedWorkMinutes: Number(expectedWorkMinutes) || 0,
  };
}

function getRoleGroupItems_() {
  return [
    {
      roleGroup: 'senior_pharmacist',
      displayLabel: 'เภสัชกรอาวุโส',
      description: 'เภสัชกรที่มีงานสำนักงานหรือสาขาต่อท้ายในบางวัน',
    },
    {
      roleGroup: 'operation_pharmacist',
      displayLabel: 'เภสัชปฏิบัติการ',
      description: 'เภสัชกรปฏิบัติการตามรหัสเวรมาตรฐาน',
    },
    {
      roleGroup: 'assistant',
      displayLabel: 'ผู้ช่วยเภสัชกร',
      description: 'ผู้ช่วยเภสัชกรหรือพนักงานสนับสนุนหน้าร้าน',
    },
    {
      roleGroup: 'sales_marketing',
      displayLabel: 'ฝ่ายขายและการตลาด',
      description: 'กลุ่มฝ่ายขายและการตลาด',
    },
    {
      roleGroup: 'branch',
      displayLabel: 'พนักงานสาขา',
      description: 'พนักงานประจำสาขา',
    },
    {
      roleGroup: 'unknown',
      displayLabel: 'ไม่ระบุกลุ่ม',
      description: 'ยังจับกลุ่มจากแถวปัจจุบันไม่ได้',
    },
  ];
}

function getShiftDictionaryAdapter_(spreadsheet) {
  const dictionarySheet = getShiftDictionarySheet_(spreadsheet);
  if (!dictionarySheet) {
    return {
      source: {
        type: 'fallback',
        sheetName: '',
        message: 'Shift Dictionary sheet was not found. Using built-in fallback mapping.',
      },
      items: getFallbackShiftDictionaryItems_(),
    };
  }

  const sheetItems = readShiftDictionaryItemsFromSheet_(dictionarySheet);
  if (sheetItems.length === 0) {
    return {
      source: {
        type: 'fallback',
        sheetName: dictionarySheet.getName(),
        message: 'Shift Dictionary sheet exists but has no valid rows. Using built-in fallback mapping.',
      },
      items: getFallbackShiftDictionaryItems_(),
    };
  }

  return {
    source: {
      type: 'sheet',
      sheetName: dictionarySheet.getName(),
      message: 'Shift Dictionary loaded from worksheet.',
    },
    items: sheetItems,
  };
}

function getShiftDictionarySheet_(spreadsheet) {
  for (let index = 0; index < SHIFT_DICTIONARY_SHEET_NAMES.length; index += 1) {
    const sheet = spreadsheet.getSheetByName(SHIFT_DICTIONARY_SHEET_NAMES[index]);
    if (sheet) {
      return sheet;
    }
  }

  return null;
}

function readShiftDictionaryItemsFromSheet_(sheet) {
  const values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) {
    return [];
  }

  const headerMap = buildShiftDictionaryHeaderMap_(values[0]);
  if (headerMap.rawCode === undefined) {
    return [];
  }

  return values.slice(1).reduce((items, row, rowIndex) => {
    const rawCode = getDictionaryCellValue_(row, headerMap, 'rawCode');
    if (!rawCode) {
      return items;
    }

    const roleGroup = getDictionaryCellValue_(row, headerMap, 'roleGroup') || 'any';
    items.push(createShiftDictionaryItem_({
      rawCode,
      roleGroup,
      displayLabel: getDictionaryCellValue_(row, headerMap, 'displayLabel') || rawCode,
      displayMode: getDictionaryCellValue_(row, headerMap, 'displayMode') || 'unknown',
      segmentType: getDictionaryCellValue_(row, headerMap, 'segmentType') || 'unknown',
      locationCode: getDictionaryCellValue_(row, headerMap, 'locationCode'),
      start: getDictionaryCellValue_(row, headerMap, 'start'),
      breakStart: getDictionaryCellValue_(row, headerMap, 'breakStart'),
      breakEnd: getDictionaryCellValue_(row, headerMap, 'breakEnd'),
      end: getDictionaryCellValue_(row, headerMap, 'end'),
      expectedWorkMinutes: getDictionaryCellValue_(row, headerMap, 'expectedWorkMinutes'),
      impliedBaseCode: getDictionaryCellValue_(row, headerMap, 'impliedBaseCode'),
      color: getDictionaryCellValue_(row, headerMap, 'color'),
      description: getDictionaryCellValue_(row, headerMap, 'description'),
      source: {
        type: 'sheet',
        sheetName: sheet.getName(),
        sourceRow: rowIndex + 2,
      },
    }));
    return items;
  }, []);
}

function buildShiftDictionaryHeaderMap_(headerRow) {
  const aliases = {
    rawCode: ['rawcode', 'raw_code', 'code', 'shiftcode', 'shift_code', 'รหัส', 'รหัสเวร'],
    roleGroup: ['rolegroup', 'role_group', 'role', 'group', 'กลุ่ม', 'สายงาน'],
    displayLabel: ['displaylabel', 'display_label', 'label', 'ชื่อแสดง', 'แสดงผล'],
    displayMode: ['displaymode', 'display_mode', 'mode', 'รูปแบบแสดงผล'],
    segmentType: ['segmenttype', 'segment_type', 'type', 'ประเภท', 'ประเภทงาน'],
    locationCode: ['locationcode', 'location_code', 'location', 'branch', 'branchcode', 'branch_code', 'สาขา', 'รหัสสาขา'],
    start: ['start', 'starttime', 'start_time', 'เริ่ม', 'เวลาเริ่ม'],
    breakStart: ['breakstart', 'break_start', 'lunchstart', 'lunch_start', 'พักเริ่ม', 'เริ่มพัก'],
    breakEnd: ['breakend', 'break_end', 'lunchend', 'lunch_end', 'พักจบ', 'จบพัก'],
    end: ['end', 'endtime', 'end_time', 'จบ', 'เวลาจบ'],
    expectedWorkMinutes: ['expectedworkminutes', 'expected_work_minutes', 'workminutes', 'work_minutes', 'นาทีทำงาน'],
    impliedBaseCode: ['impliedbasecode', 'implied_base_code', 'basecode', 'base_code', 'รหัสพื้นฐาน'],
    color: ['color', 'colour', 'สี'],
    description: ['description', 'desc', 'รายละเอียด', 'คำอธิบาย'],
  };
  const headerMap = {};

  headerRow.forEach((header, index) => {
    const normalizedHeader = normalizeDictionaryHeader_(header);
    Object.keys(aliases).forEach((field) => {
      if (headerMap[field] === undefined && aliases[field].indexOf(normalizedHeader) !== -1) {
        headerMap[field] = index;
      }
    });
  });

  return headerMap;
}

function normalizeDictionaryHeader_(header) {
  return String(header || '').toLowerCase().replace(/[\s\-_./]/g, '');
}

function getDictionaryCellValue_(row, headerMap, field) {
  const index = headerMap[field];
  if (index === undefined) {
    return '';
  }

  return String(row[index] || '').trim();
}

// Fallback adapter keeps the normalized layer useful before a real dictionary sheet exists.
function getFallbackShiftDictionaryItems_() {
  return [
    createShiftDictionaryItem_({ rawCode: '', roleGroup: 'any', displayLabel: '(ว่าง)', displayMode: 'empty', segmentType: 'empty', description: 'ไม่มีรหัสเวร' }),
    createShiftDictionaryItem_({ rawCode: '1', roleGroup: 'any', displayLabel: '1', displayMode: 'single', segmentType: 'regular', start: '08:00', end: '17:00', description: 'ทำงาน 08.00-17.00' }),
    createShiftDictionaryItem_({ rawCode: '2', roleGroup: 'any', displayLabel: '2', displayMode: 'single', segmentType: 'regular', start: '11:00', end: '20:00', description: 'ทำงาน 11.00-20.00' }),
    createShiftDictionaryItem_({ rawCode: '1/3H', roleGroup: 'any', displayLabel: '1/3H', displayMode: 'addon', segmentType: 'ot', start: '17:00', end: '20:00', impliedBaseCode: '1', description: 'OT 17.00-20.00' }),
    createShiftDictionaryItem_({ rawCode: 'O-1', roleGroup: 'senior_pharmacist', displayLabel: 'O-1', displayMode: 'base', segmentType: 'office', locationCode: 'office', start: '08:00', breakStart: '12:00', breakEnd: '13:00', end: '17:00', expectedWorkMinutes: 480, description: 'ทำงานสำนักงาน 08.00-17.00 พัก 12.00-13.00' }),
    createShiftDictionaryItem_({ rawCode: '4H-1', roleGroup: 'senior_pharmacist', displayLabel: '4H-1', displayMode: 'composite', segmentType: 'branch_ot', locationCode: '001', start: '17:00', end: '21:00', impliedBaseCode: 'O-1', description: 'สำนักงาน + สาขา 001 17.00-21.00' }),
    createShiftDictionaryItem_({ rawCode: '3H-4', roleGroup: 'senior_pharmacist', displayLabel: '3H-4', displayMode: 'composite', segmentType: 'branch_ot', locationCode: '004', start: '17:00', end: '20:00', impliedBaseCode: 'O-1', description: 'สำนักงาน + สาขา 004 17.00-20.00' }),
    createShiftDictionaryItem_({ rawCode: '11H-3', roleGroup: 'senior_pharmacist', displayLabel: '11H-3', displayMode: 'single', segmentType: 'branch', locationCode: '003', description: 'สาขา 003 รวม 11 ชั่วโมง' }),
    createShiftDictionaryItem_({ rawCode: 'AL1', roleGroup: 'any', displayLabel: 'AL1', displayMode: 'leave', segmentType: 'leave', description: 'ลาพักร้อน' }),
    createShiftDictionaryItem_({ rawCode: 'AL2', roleGroup: 'any', displayLabel: 'AL2', displayMode: 'leave', segmentType: 'leave', description: 'ลาพักร้อน' }),
    createShiftDictionaryItem_({ rawCode: 'AL3', roleGroup: 'any', displayLabel: 'AL3', displayMode: 'leave', segmentType: 'leave', description: 'ลาพักร้อน' }),
    createShiftDictionaryItem_({ rawCode: 'AL4', roleGroup: 'any', displayLabel: 'AL4', displayMode: 'leave', segmentType: 'leave', description: 'ลาพักร้อน' }),
    createShiftDictionaryItem_({ rawCode: 'AL5', roleGroup: 'any', displayLabel: 'AL5', displayMode: 'leave', segmentType: 'leave', description: 'ลาพักร้อน' }),
    createShiftDictionaryItem_({ rawCode: 'AL6', roleGroup: 'any', displayLabel: 'AL6', displayMode: 'leave', segmentType: 'leave', description: 'ลาพักร้อน' }),
    createShiftDictionaryItem_({ rawCode: 'BL', roleGroup: 'any', displayLabel: 'BL', displayMode: 'leave', segmentType: 'leave', description: 'ลา' }),
    createShiftDictionaryItem_({ rawCode: 'RxM', roleGroup: 'any', displayLabel: 'RxM', displayMode: 'assignment', segmentType: 'meeting', description: 'งาน RxM' }),
    createShiftDictionaryItem_({ rawCode: 'Off1', roleGroup: 'any', displayLabel: 'Off1', displayMode: 'off', segmentType: 'off', description: 'วันหยุด' }),
    createShiftDictionaryItem_({ rawCode: 'Off2', roleGroup: 'any', displayLabel: 'Off2', displayMode: 'off', segmentType: 'off', description: 'วันหยุด' }),
    createShiftDictionaryItem_({ rawCode: 'Off3', roleGroup: 'any', displayLabel: 'Off3', displayMode: 'off', segmentType: 'off', description: 'วันหยุด' }),
    createShiftDictionaryItem_({ rawCode: 'Off4', roleGroup: 'any', displayLabel: 'Off4', displayMode: 'off', segmentType: 'off', description: 'วันหยุด' }),
    createShiftDictionaryItem_({ rawCode: 'Off5', roleGroup: 'any', displayLabel: 'Off5', displayMode: 'off', segmentType: 'off', description: 'วันหยุด' }),
    createShiftDictionaryItem_({ rawCode: 'Off6', roleGroup: 'any', displayLabel: 'Off6', displayMode: 'off', segmentType: 'off', description: 'วันหยุด' }),
  ];
}

function createShiftDictionaryItem_(item) {
  const rawCode = item.rawCode || '';
  const roleGroup = item.roleGroup || 'any';

  return {
    shiftDictionaryItemId: `shift:${roleGroup}:${rawCode || 'blank'}`,
    rawCode,
    roleGroup,
    displayLabel: item.displayLabel || rawCode || '(ว่าง)',
    displayMode: item.displayMode || 'unknown',
    segmentType: item.segmentType || 'unknown',
    locationCode: item.locationCode || '',
    start: item.start || '',
    breakStart: item.breakStart || '',
    breakEnd: item.breakEnd || '',
    end: item.end || '',
    expectedWorkMinutes: item.expectedWorkMinutes || '',
    impliedBaseCode: item.impliedBaseCode || '',
    color: item.color || '',
    description: item.description || '',
    source: item.source || {
      type: 'fallback',
      sheetName: '',
      sourceRow: 0,
    },
  };
}

function resolveShiftDictionaryItem_(shiftDictionaryItems, rawCode, roleGroup) {
  const code = rawCode || '';
  const roleMatch = shiftDictionaryItems.find((item) => item.rawCode === code && item.roleGroup === roleGroup);
  if (roleMatch) {
    return roleMatch;
  }

  const genericMatch = shiftDictionaryItems.find((item) => item.rawCode === code && item.roleGroup === 'any');
  if (genericMatch) {
    return genericMatch;
  }

  return createShiftDictionaryItem_({
    rawCode: code,
    roleGroup: roleGroup || 'unknown',
    displayLabel: code || '(ว่าง)',
    displayMode: code ? 'unknown' : 'empty',
    segmentType: code ? 'unknown' : 'empty',
    locationCode: inferLocationCodeFromRawCode_(code),
    description: code ? 'ยังไม่มีข้อมูลใน Shift Dictionary' : 'ไม่มีรหัสเวร',
    source: {
      type: 'unmapped',
      sheetName: '',
      sourceRow: 0,
    },
  });
}

function inferRoleGroup_(section, position, team) {
  const text = `${section || ''} ${position || ''} ${team || ''}`;

  if (text.indexOf('อาวุโส') !== -1) {
    return 'senior_pharmacist';
  }

  if (text.indexOf('ปฏิบัติการ') !== -1) {
    return 'operation_pharmacist';
  }

  if (text.indexOf('ผู้ช่วย') !== -1) {
    return 'assistant';
  }

  if (section === 'ฝ่ายขายและการตลาด') {
    return 'sales_marketing';
  }

  if (/^0\d{2}\s/.test(section || '')) {
    return 'branch';
  }

  if (section === 'เภสัชกร') {
    return 'operation_pharmacist';
  }

  return 'unknown';
}

function inferLocationCodeFromRawCode_(rawCode) {
  const match = String(rawCode || '').match(/-(\d+)$/);
  if (!match) {
    return '';
  }

  return match[1].padStart(3, '0');
}

function getMonthNumberFromSheetName_(sheetName) {
  for (let month = 1; month <= 12; month += 1) {
    if (getMonthTemplateMeta_(month).sheetName === sheetName) {
      return month;
    }
  }

  return 0;
}

function toIsoDate_(year, month, day) {
  return `${year}-${pad2_(month)}-${pad2_(day)}`;
}

function pad2_(value) {
  return String(value).padStart(2, '0');
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

function buildShiftMetaByCell_(scheduleEntries) {
  return scheduleEntries.reduce((map, entry) => {
    map[`${entry.sourceRow}:${entry.sourceColumn}`] = {
      rawCode: entry.rawCode,
      displayLabel: entry.displayLabel,
      displayMode: entry.displayMode,
      roleGroup: entry.roleGroup,
      segmentType: entry.segmentType,
      locationCode: entry.locationCode,
      start: entry.start,
      breakStart: entry.breakStart,
      breakEnd: entry.breakEnd,
      end: entry.end,
      expectedWorkMinutes: entry.expectedWorkMinutes,
      impliedBaseCode: entry.impliedBaseCode,
      color: entry.color,
      description: entry.description,
      dictionarySource: entry.dictionarySource,
    };
    return map;
  }, {});
}

function buildWorksheetRows_(
  values,
  backgrounds,
  fontColors,
  fontWeights,
  horizontalAlignments,
  verticalAlignments,
  shiftMetaByCell,
) {
  const metaByCell = shiftMetaByCell || {};
  return values.map((row, rowIndex) => ({
    rowNumber: rowIndex + 1,
    cells: row.map((value, columnIndex) => {
      const sourceRow = rowIndex + 1;
      const sourceColumn = columnIndex + 1;
      return {
        value: sanitizeWorksheetDisplayValue_(value, sourceRow, sourceColumn),
        background: backgrounds[rowIndex][columnIndex],
        fontColor: fontColors[rowIndex][columnIndex],
        fontWeight: fontWeights[rowIndex][columnIndex],
        horizontalAlignment: horizontalAlignments[rowIndex][columnIndex],
        verticalAlignment: verticalAlignments[rowIndex][columnIndex],
        row: sourceRow,
        column: sourceColumn,
        a1: toA1_(sourceRow, sourceColumn),
        shiftMeta: getWorksheetCellShiftMeta_(
          metaByCell[`${sourceRow}:${sourceColumn}`] || null,
          value,
          backgrounds[rowIndex][columnIndex],
        ),
      };
    }),
  }));
}

function getWorksheetCellShiftMeta_(shiftMeta, rawCode, background) {
  const code = String(rawCode || '').trim();
  if (isPublicHolidayLeaveRawCode_(code) && isPublicHolidayLeaveBackground_(background)) {
    return createPublicHolidayShiftMeta_(code, shiftMeta);
  }

  return shiftMeta || null;
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
