/** Sheet-only identity validation. Existing-data migration belongs to the worker EXE. */
function isExternalCalendarOwner_() {
  // Calendar operations never belong to GAS.
  return true;
}

function externalCalendarResult_(context) {
  return { ok: true, status: 'external_calendar_owner',
    eventId: context && context.eventId || '',
    message: '刀表已保存；由指定 Windows 主機監控並同步日曆。' };
}

function externalCalendarMenu_() {
  SpreadsheetApp.getUi().alert('GAS 不再操作日曆。請在指定 Windows 主機查看同步狀態。');
  return externalCalendarResult_();
}

function newScheduleSyncId_() {
  const value = String(Utilities.getUuid()).replace(/-/g, '').toLowerCase();
  if (!/^[a-f0-9]{32}$/.test(value)) throw new Error('無法產生有效排程識別碼。');
  return value;
}

function getScheduleIdentitySheetInfo_(sheet) {
  const name = sheet.getName();
  const archive = isAnnualArchiveSheet_(sheet);
  const kind = archive ? 'ARCHIVE' : isMainTrackingSheetName_(name) ? 'FU'
    : isMonthlySheetName_(name) ? 'MONTHLY' : '';
  if (!kind) return null;
  const headerRow = archive ? ANNUAL_ARCHIVE_HEADER_ROW : 1;
  const lastColumn = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0].map(toCellText_);
  const byHeader = {};
  headers.forEach((value, index) => {
    if (!value) return;
    if (byHeader[value]) throw new Error(`${name} 欄名重複：${value}。`);
    byHeader[value] = index + 1;
  });
  const headerByKey = kind === 'FU' ? CONFIG.FIELD_HEADERS : CONFIG.MONTHLY_FIELD_HEADERS;
  const columns = {};
  if (kind !== 'ARCHIVE') {
    const managed = kind === 'FU' ? getFuColumnInfo_(sheet) : getMonthlyColumnInfo_(sheet);
    assertUniqueHeaders_(managed, name);
    Object.assign(columns, managed.columns);
  } else {
    Object.keys(headerByKey).forEach(key => { columns[key] = byHeader[headerByKey[key]] || 0; });
  }
  if (!columns.CHART_NO || !columns.NAME || (kind === 'MONTHLY' && !columns.TIME)) {
    throw new Error(`${name} 缺少辨識病人／刀日的必要欄位，請先修復標題。`);
  }
  const firstRow = headerRow + 1;
  const count = Math.max(0, sheet.getLastRow() - headerRow);
  const values = count ? sheet.getRange(firstRow, 1, count, lastColumn).getValues() : [];
  return { sheet, name, kind, headerRow, headers, byHeader, columns, lastColumn, firstRow, values };
}

function classifyScheduleIdentityRow_(info, values) {
  if (info.kind === 'ARCHIVE') {
    const rowType = toCellText_(values[(info.byHeader[ANNUAL_ARCHIVE_ROW_TYPE_HEADER] || 0) - 1]);
    return rowType === ARCHIVE_ROW_TYPES.PATIENT ? 'PATIENT'
      : rowType === ARCHIVE_ROW_TYPES.DATE ? 'DATE_HEADER' : '';
  }
  const chart = toCellText_(getRowFieldValue_(values, info.columns, 'CHART_NO'));
  const name = toCellText_(getRowFieldValue_(values, info.columns, 'NAME'));
  if (info.kind === 'MONTHLY') {
    const classification = classifyMonthlyRowValues_(values, info.columns);
    if (classification.type === MONTHLY_ROW_TYPES.HYBRID_CONFLICT) {
      throw new Error(`${info.name} 有不完整或混合刀日標題，請先修復。`);
    }
    if (classification.type === MONTHLY_ROW_TYPES.DATE_HEADER) {
      return 'DATE_HEADER';
    }
  }
  return chart || name ? 'PATIENT' : '';
}

function scanScheduleIdentities_(spreadsheet) {
  const sheets = spreadsheet.getSheets().map(getScheduleIdentitySheetInfo_).filter(Boolean);
  const seen = {};
  const conflicts = [];
  let assigned = 0;
  let missing = 0;
  sheets.forEach(info => {
    info.rows = [];
    info.values.forEach((values, index) => {
      const row = info.firstRow + index;
      const type = classifyScheduleIdentityRow_(info, values);
      const id = toCellText_(getRowFieldValue_(values, info.columns, 'SYNC_ID'));
      if (id) {
        if (!/^[a-f0-9]{32}$/.test(id)) conflicts.push(`${info.name} 第 ${row} 列識別碼格式無效`);
        if (seen[id]) conflicts.push(`識別碼重複：${seen[id]} 與 ${info.name} 第 ${row} 列`);
        else seen[id] = `${info.name} 第 ${row} 列`;
        if (!type) conflicts.push(`${info.name} 第 ${row} 列只剩識別碼，需人工確認`);
        assigned++;
      } else if (type) missing++;
      info.rows.push({ row, type, id });
    });
  });
  return { sheets, conflicts, assigned, missing };
}

// Stale menus/old callers remain harmless after code deployment.
function migrateScheduleSyncIds() { return externalCalendarMenu_(); }
function checkScheduleSyncIds() { return externalCalendarMenu_(); }
function handoverCalendarToApplication() { return externalCalendarMenu_(); }
function showCalendarHandover() { return externalCalendarMenu_(); }
function restoreGasCalendarOwner() { return externalCalendarMenu_(); }
function processMonthlyStructureChange() { return externalCalendarResult_(); }
function processCalendarChange() { return externalCalendarResult_(); }
function exportUpcomingWeekData() { return externalCalendarResult_(); }
function runPendingUpcomingWeekExport() { return externalCalendarResult_(); }
function autoExportByEdit() { return externalCalendarResult_(); }
