/**
 * Formatting-only GAS, protocol 2. Calendar API access is deliberately absent.
 * ScheduleSyncId identifies both a patient row and a ◆ 刀日 row independently.
 */
const EXTERNAL_CALENDAR_OWNER_PROPERTY = 'CALENDAR_OWNER_APPLICATION_V1';
const EXTERNAL_CALENDAR_CONTROL_KEY = 'vghks.calendar.control';
const SCHEDULE_SYNC_HEADER = 'ScheduleSyncId';

function isExternalCalendarOwner_() {
  // Always true, including before handover and after a partial migration.
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
  Object.keys(headerByKey).forEach(key => { columns[key] = byHeader[headerByKey[key]] || 0; });
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
    const date = parseFullDate_(getRowFieldValue_(values, info.columns, 'TIME'));
    if (date) {
      if (chart !== MONTHLY_DATE_HEADER_MARKER ||
          toCellText_(getRowFieldValue_(values, info.columns, 'EVENT_ID'))) {
        throw new Error(`${info.name} 有不完整或混合刀日標題，請先修復。`);
      }
      return 'DATE_HEADER';
    }
    if (chart === MONTHLY_DATE_HEADER_MARKER) {
      throw new Error(`${info.name} 有缺少有效日期的刀日標題，請先修復。`);
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

function migrateScheduleSyncIds_(spreadsheet) {
  const scan = scanScheduleIdentities_(spreadsheet);
  if (scan.conflicts.length) throw new Error('尚未修改資料：\n' + scan.conflicts.join('\n'));
  // All-sheet duplicate/structure validation precedes any write. Existing IDs,
  // clinical values, notes, formulas and CalendarEventId are never rewritten.
  let created = 0;
  scan.sheets.forEach(info => {
    let column = info.columns.SYNC_ID;
    if (!column) {
      column = info.lastColumn + 1;
      ensureColumnCapacity_(info.sheet, column);
      info.sheet.getRange(info.headerRow, column).setValue(SCHEDULE_SYNC_HEADER);
    }
    info.rows.forEach(item => {
      if (!item.type || item.id) return;
      const id = newScheduleSyncId_();
      info.sheet.getRange(item.row, column).setValue(id);
      if (toCellText_(info.sheet.getRange(item.row, column).getValue()) !== id) {
        throw new Error('識別碼回讀不符；已完成項目保留，請重新執行遷移。');
      }
      created++;
    });
    hideAndProtectSystemColumns_(info.sheet, { SYNC_ID: column }, ['SYNC_ID']);
  });
  const verified = scanScheduleIdentities_(spreadsheet);
  if (verified.conflicts.length || verified.missing) throw new Error('遷移未完整，請修正後重試。');
  return { ok: true, created, retained: scan.assigned, sheetCount: scan.sheets.length };
}

function migrateScheduleSyncIds() {
  return runMenuAction_('遷移固定排程識別欄', () => withSheetOperationLock_(() => {
    assertNoActiveAnnualArchiveTransaction_();
    const result = migrateScheduleSyncIds_(SpreadsheetApp.getActiveSpreadsheet());
    SpreadsheetApp.getUi().alert(`遷移完成：新增 ${result.created}、保留 ${result.retained} 個識別碼。\n` +
      '既有病人資料、公式與 CalendarEventId 均保留；請授權監控服務帳戶編輯系統欄。');
    return result;
  }));
}

function checkScheduleSyncIds() {
  return runMenuAction_('檢查排程識別欄', () => {
    const scan = scanScheduleIdentities_(SpreadsheetApp.getActiveSpreadsheet());
    SpreadsheetApp.getUi().alert(scan.conflicts.length ? scan.conflicts.join('\n')
      : `識別碼 ${scan.assigned} 個；待配置 ${scan.missing} 個。`);
    return { ok: !scan.conflicts.length, assigned: scan.assigned, missing: scan.missing,
      conflicts: scan.conflicts };
  });
}

function writeExternalCalendarControl_(spreadsheet, value) {
  const old = spreadsheet.getDeveloperMetadata().filter(item => item.getKey() === EXTERNAL_CALENDAR_CONTROL_KEY);
  if (old.length > 1) throw new Error('交接設定重複，請先由擁有者檢查。');
  const serialized = JSON.stringify(value);
  if (old.length) old[0].setValue(serialized);
  else spreadsheet.addDeveloperMetadata(EXTERNAL_CALENDAR_CONTROL_KEY,
    serialized, SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT);
  const verified = spreadsheet.getDeveloperMetadata().filter(item => item.getKey() === EXTERNAL_CALENDAR_CONTROL_KEY);
  if (verified.length !== 1 || verified[0].getValue() !== serialized) {
    throw new Error('交接設定回讀失敗；GAS 仍不操作日曆，請重新執行交接。');
  }
}



function handoverCalendarToApplication() {
  return runMenuAction_('交由本系統同步日曆', () => {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    if (Session.getScriptTimeZone() !== 'Asia/Taipei' || spreadsheet.getSpreadsheetTimeZone() !== 'Asia/Taipei') {
      throw new Error('GAS 與刀表時區須為 Asia/Taipei。');
    }
    let calendarId = getConfiguredCalendarId_();
    const calendarReply = ui.prompt('指定日曆', '請輸入監控 EXE 使用的 Calendar ID：\n' + calendarId, ui.ButtonSet.OK_CANCEL);
    if (calendarReply.getSelectedButton() !== ui.Button.OK) return { ok: false, cancelled: true };
    calendarId = toCellText_(calendarReply.getResponseText());
    if (!calendarId || /[\r\n]/.test(calendarId)) throw new Error('Calendar ID 不可空白或包含換行。');
    const reply = ui.prompt('指定唯一日曆同步主機',
      '請貼上 Windows 主機 init 顯示的 worker_id（32 位小寫十六進位）。\n更換前先停止舊主機。', ui.ButtonSet.OK_CANCEL);
    if (reply.getSelectedButton() !== ui.Button.OK) return { ok: false, cancelled: true };
    const workerId = toCellText_(reply.getResponseText());
    if (!/^[a-f0-9]{32}$/.test(workerId)) throw new Error('worker_id 格式不正確。');
    return withSheetOperationLock_(() => {
      assertNoActiveAnnualArchiveTransaction_();
      const identities = scanScheduleIdentities_(spreadsheet);
      if (identities.conflicts.length) throw new Error(identities.conflicts.join('\n'));
      if (identities.sheets.some(info => !info.columns.SYNC_ID)) {
        throw new Error('請先執行「遷移固定排程識別欄」。');
      }
      assertScheduleIntegrity_(spreadsheet);
      const control = { protocol: 2, owner: 'application', gas_calendar_disabled: true,
        calendar_id: calendarId, time_zone: 'Asia/Taipei', gas_version: CONFIG.VERSION, worker_id: workerId };
      PropertiesService.getScriptProperties().setProperty(EXTERNAL_CALENDAR_OWNER_PROPERTY, JSON.stringify(control));
      PropertiesService.getScriptProperties().setProperty(CALENDAR_ID_PROPERTY, calendarId);
      writeExternalCalendarControl_(spreadsheet, control);
      installOrRepairTriggers_(false);
      onOpen();
      ui.alert('交接完成。GAS 只處理表格，未存取或修改日曆事件。\n請由指定主機 check 後再啟用。');
      return { ok: true, workerId, calendarId };
    });
  });
}

function showCalendarHandover() {
  const value = PropertiesService.getScriptProperties().getProperty(EXTERNAL_CALENDAR_OWNER_PROPERTY);
  SpreadsheetApp.getUi().alert(value ? 'GAS 不操作日曆；目前交接設定：\n' + value
    : 'GAS 不操作日曆；尚未設定指定同步主機。');
}

function restoreGasCalendarOwner() {
  // Old deployments may retain this menu entry; never restore an API writer.
  return externalCalendarMenu_();
}

// Compatibility entrypoints for triggers installed by older versions/users.
function processMonthlyStructureChange() { return externalCalendarResult_(); }
function processCalendarChange() { return externalCalendarResult_(); }
function exportUpcomingWeekData() { return externalCalendarResult_(); }
function runPendingUpcomingWeekExport() { return externalCalendarResult_(); }
function autoExportByEdit() { return externalCalendarResult_(); }
