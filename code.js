/**
 * 手術排程系統 Google Apps Script
 *
 * 首次安裝流程：
 * 1. 將本檔貼到試算表綁定的 Apps Script 專案。
 * 2. 手動執行 setup() 一次，完成授權、表格初始化與 trigger 安裝。
 * 3. 若 setup 尚未提示設定，回到試算表後可用「維護工具」的「設定 Calendar ID」儲存手術日曆 ID。
 *
 * 重要假設：
 * - 主要資料表名稱為 CONFIG.SHEET_OP，預設是「OP」。
 * - 系統欄位由第 1 列欄位名稱辨識，欄位順序可以調整或插入自訂欄位。
 * - 病歷號是同步 Calendar 的必要欄位；有日期但沒有病歷號時不會新增或更新事件。
 */
const CONFIG = {
  VERSION: '2026.05.09',
  CALENDAR_ID: 'YOUR_CALENDAR_ID_HERE', // fallback：優先使用 ScriptProperties 內的 Calendar ID
  SHEET_OP: 'OP',
  SHEET_OUT: '輸出表單',
  FIELD_KEYS: [
    'CHART_NO',
    'NAME',
    'TEL',
    'TAG',
    'COND',
    'DATE',
    'TIME',
    'PLAN',
    'MEMO',
    'EVENT_ID',
    'SHEET_WRITE_UPDATED'
  ],
  SYSTEM_FIELD_KEYS: ['EVENT_ID', 'SHEET_WRITE_UPDATED'],
  FIELD_HEADERS: {
    CHART_NO: '病歷號',
    NAME: '姓名',
    TEL: 'TEL',
    TAG: 'Tag',
    COND: 'Condition',
    DATE: '日期',
    TIME: '時間',
    PLAN: 'Plan',
    MEMO: '心得',
    EVENT_ID: 'CalendarEventId',
    SHEET_WRITE_UPDATED: 'CalendarSheetWriteUpdated'
  },

  HEADERS: [
    '病歷號',
    '姓名',
    'TEL',
    'Tag',
    'Condition',
    '日期',
    '時間',
    'Plan',
    '心得',
    'CalendarEventId',
    'CalendarSheetWriteUpdated'
  ],
  TAG_OPTIONS: ['CATA', 'Eyelid', 'Retina', 'OP', 'FU', 'Suture IOL', 'Complication'],

  // 條件格式化關鍵字：
  // Plan 欄包含黃色關鍵字時套淡黃；包含綠色關鍵字時套淡綠。
  PLAN_YELLOW_KEYWORDS: ['問', '補', '通知'],
  PLAN_GREEN_KEYWORDS: ['APPLY'],
  TIME_GA_KEYWORDS: ['GA'],

  COLORS: {
    NO_TIME: '5',
    OP: '11',
    DEFAULT: '9'
  }
};

const AUTO_EXPORT_UPCOMING_DAYS = 7;
const AUTO_EXPORT_DAILY_HOUR = 6;
const AUTO_EXPORT_DAILY_MINUTE = 0;
const AUTO_EXPORT_EDIT_DELAY_MS = 5 * 60 * 1000;
const AUTO_EXPORT_DAILY_HANDLER = 'exportUpcomingWeekData';
const AUTO_EXPORT_PENDING_HANDLER = 'runPendingUpcomingWeekExport';

const TIME_ERROR_NOTE_PREFIX = '系統時間檢查：';
const CALENDAR_SYNC_NOTE_PREFIX = '系統日曆同步：';
const CALENDAR_SYNC_TOKEN_PROPERTY_PREFIX = 'CALENDAR_SYNC_TOKEN_';
const CALENDAR_REVERSE_SYNC_HANDLER = 'processCalendarChange';
const CALENDAR_TITLE_SEPARATOR = '|';
const CALENDAR_ID_SCRIPT_PROPERTY_KEY = 'CALENDAR_ID';
const LEGACY_EVENT_ID_HEADERS = ['日曆eventID', '日曆apiEventID'];
const SYSTEM_COLUMN_PROTECTION_PREFIX = 'OP_SPREADSHEET_SYSTEM_COLUMN:';
const SYSTEM_COLUMN_PROTECTION_EDITORS_PROPERTY_KEY = 'SYSTEM_COLUMN_PROTECTION_EDITORS';

// 條件格式沒有 metadata 可標記來源，因此在 custom formula 中放入 N("...") marker。
// initializeSheet() 重新執行時會用這些 marker 辨識並替換系統規則，同時保留使用者自訂規則。
const CONDITIONAL_FORMAT_MARKERS = {
  OP_RED: 'SYSTEM_CF_OP_RED',
  PLAN_YELLOW: 'SYSTEM_CF_PLAN_YELLOW',
  PLAN_YELLOW_TIME: 'SYSTEM_CF_PLAN_YELLOW_TIME',
  PLAN_GREEN: 'SYSTEM_CF_PLAN_GREEN',
  PLAN_GREEN_TIME: 'SYSTEM_CF_PLAN_GREEN_TIME',
  TIME_GA: 'SYSTEM_CF_TIME_GA',
  PLAN_YELLOW_MAIN: 'SYSTEM_CF_PLAN_YELLOW_MAIN',
  PLAN_YELLOW_TAG: 'SYSTEM_CF_PLAN_YELLOW_TAG'
};

function deleteTriggersByHandler_(handlerFunctionName) {
  let deletedCount = 0;

  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === handlerFunctionName)
    .forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    });

  return deletedCount;
}

/**
 * 安裝 processRowChange 的 installable onEdit trigger。
 *
 * GAS 的 simple trigger 無法使用需要授權的 Calendar API，所以這裡必須建立
 * installable onEdit trigger。每次安裝前會先刪除同 handler 的舊 trigger，
 * 避免同一次編輯觸發多次同步。
 */
function installProcessRowChangeTrigger_(showAlert) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 避免重複安裝同一個 handler，造成一次編輯觸發多次同步
  ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'processRowChange')
    .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('processRowChange')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  if (showAlert) {
    SpreadsheetApp.getUi().alert('已安裝 processRowChange 的 onEdit 觸發器。');
  }
}

function installProcessRowChangeTrigger() {
  installProcessRowChangeTrigger_(true);
}

function isCalendarAdvancedServiceAvailable_() {
  return typeof Calendar !== 'undefined' &&
    Calendar.Events &&
    typeof Calendar.Events.list === 'function' &&
    typeof Calendar.Events.get === 'function' &&
    typeof Calendar.Events.update === 'function' &&
    typeof Calendar.Events.insert === 'function' &&
    typeof Calendar.Events.remove === 'function';
}

function getCalendarSyncTokenPropertyKey_(calendarId) {
  return CALENDAR_SYNC_TOKEN_PROPERTY_PREFIX + String(calendarId || '')
    .replace(/[^A-Za-z0-9_]/g, '_');
}

function getCalendarSyncToken_(calendarId) {
  return PropertiesService.getScriptProperties()
    .getProperty(getCalendarSyncTokenPropertyKey_(calendarId));
}

function setCalendarSyncToken_(calendarId, syncToken) {
  PropertiesService.getScriptProperties()
    .setProperty(getCalendarSyncTokenPropertyKey_(calendarId), syncToken);
}

function parseCalendarUpdatedMillis_(updated) {
  const text = toCellText_(updated);
  if (!text) return NaN;

  const millis = new Date(text).getTime();
  return isNaN(millis) ? NaN : millis;
}

function compareCalendarUpdated_(eventUpdated, rememberedUpdated) {
  const eventUpdatedText = toCellText_(eventUpdated);
  const rememberedUpdatedText = toCellText_(rememberedUpdated);
  if (!eventUpdatedText || !rememberedUpdatedText) return null;

  const eventMillis = parseCalendarUpdatedMillis_(eventUpdatedText);
  const rememberedMillis = parseCalendarUpdatedMillis_(rememberedUpdatedText);

  if (!isNaN(eventMillis) && !isNaN(rememberedMillis)) {
    if (eventMillis < rememberedMillis) return -1;
    if (eventMillis > rememberedMillis) return 1;
    return 0;
  }

  return eventUpdatedText === rememberedUpdatedText ? 0 : null;
}

function buildCalendarApiEventResourceFromSheetRow_(sheet, row, columns) {
  const rowData = sheet.getRange(row, 1, 1, getTableLastColumn_(sheet, columns)).getValues()[0];
  const chartNoText = toCellText_(getRowFieldValue_(rowData, columns, 'CHART_NO'));
  const patientNameText = toSingleLineText_(getRowFieldValue_(rowData, columns, 'NAME'));
  const telText = toCellText_(getRowFieldValue_(rowData, columns, 'TEL'));
  const tagText = toCellText_(getRowFieldValue_(rowData, columns, 'TAG'));
  const conditionText = toSingleLineText_(getRowFieldValue_(rowData, columns, 'COND'));
  const dateVal = getRowFieldValue_(rowData, columns, 'DATE');
  const timeVal = getRowFieldValue_(rowData, columns, 'TIME');
  const planText = toCellText_(getRowFieldValue_(rowData, columns, 'PLAN'));

  if (!chartNoText || !dateVal) return null;

  const dateObj = new Date(dateVal);
  if (isNaN(dateObj.getTime())) return null;

  const timeInfo = resolveCalendarTime_(timeVal);
  if (timeInfo.errorMessage) return null;

  return buildCalendarApiEventResource_(
    buildCalendarTitle_(chartNoText, patientNameText, conditionText),
    buildCalendarDescription_(telText, planText, timeInfo.timeNote),
    dateObj,
    timeInfo,
    tagText
  );
}

function sameCalendarDateTimeValue_(left, right) {
  const leftText = toCellText_(left);
  const rightText = toCellText_(right);
  if (!leftText || !rightText) return leftText === rightText;

  const leftMillis = new Date(leftText).getTime();
  const rightMillis = new Date(rightText).getTime();
  if (!isNaN(leftMillis) && !isNaN(rightMillis)) {
    return leftMillis === rightMillis;
  }

  return leftText === rightText;
}

function calendarEventEndpointMatchesResource_(eventEndpoint, resourceEndpoint) {
  const eventValue = eventEndpoint || {};
  const resourceValue = resourceEndpoint || {};

  if (resourceValue.date) {
    return toCellText_(eventValue.date) === toCellText_(resourceValue.date);
  }

  if (resourceValue.dateTime) {
    return sameCalendarDateTimeValue_(eventValue.dateTime, resourceValue.dateTime);
  }

  return !eventValue.date && !eventValue.dateTime;
}

function calendarEventMatchesResource_(event, resource) {
  if (!event || !resource) return false;

  return toCellText_(event.summary) === toCellText_(resource.summary) &&
    toCellText_(event.description) === toCellText_(resource.description) &&
    toCellText_(event.colorId || CONFIG.COLORS.DEFAULT) === toCellText_(resource.colorId || CONFIG.COLORS.DEFAULT) &&
    calendarEventEndpointMatchesResource_(event.start, resource.start) &&
    calendarEventEndpointMatchesResource_(event.end, resource.end);
}

function calendarEventMatchesCurrentSheetProjection_(sheet, row, columns, event) {
  const resource = buildCalendarApiEventResourceFromSheetRow_(sheet, row, columns);
  return calendarEventMatchesResource_(event, resource);
}

function isSheetOriginatedCalendarEcho_(sheet, row, columns, event) {
  if (!row || !columns.SHEET_WRITE_UPDATED) return false;

  const comparison = compareCalendarUpdated_(
    event && event.updated,
    sheet.getRange(row, columns.SHEET_WRITE_UPDATED).getValue()
  );

  if (comparison === null) return false;
  if (comparison < 0) return true;
  if (comparison > 0) return false;

  return calendarEventMatchesCurrentSheetProjection_(sheet, row, columns, event);
}

function writeSheetOriginatedCalendarState_(sheet, row, columns, event, fallbackEventId) {
  const eventId = toCellText_(event && event.id) || toCellText_(fallbackEventId);
  const updated = toCellText_(event && event.updated);

  if (eventId) {
    sheet.getRange(row, columns.EVENT_ID).setValue(eventId);
  }

  sheet.getRange(row, columns.SHEET_WRITE_UPDATED).setValue(updated);
}

function clearSheetOriginatedCalendarState_(sheet, row, columns) {
  sheet.getRange(row, columns.EVENT_ID).clearContent();
  sheet.getRange(row, columns.SHEET_WRITE_UPDATED).clearContent();
}

function normalizeCalendarId_(calendarId) {
  const text = toCellText_(calendarId);
  return text && text !== 'YOUR_CALENDAR_ID_HERE' ? text : '';
}

function getConfiguredCalendarId_() {
  const propertyCalendarId = PropertiesService.getScriptProperties()
    .getProperty(CALENDAR_ID_SCRIPT_PROPERTY_KEY);

  return normalizeCalendarId_(propertyCalendarId) || normalizeCalendarId_(CONFIG.CALENDAR_ID);
}

function setConfiguredCalendarId_(calendarId) {
  const normalizedCalendarId = normalizeCalendarId_(calendarId);
  if (!normalizedCalendarId) return '';

  PropertiesService.getScriptProperties()
    .setProperty(CALENDAR_ID_SCRIPT_PROPERTY_KEY, normalizedCalendarId);

  return normalizedCalendarId;
}

function isCalendarIdConfigured_(calendarId) {
  return Boolean(arguments.length > 0 ? normalizeCalendarId_(calendarId) : getConfiguredCalendarId_());
}

function promptAndSaveCalendarId_(options) {
  const settings = options || {};
  const ui = SpreadsheetApp.getUi();
  const currentCalendarId = getConfiguredCalendarId_();
  const message = settings.message || (currentCalendarId
    ? `目前 Calendar ID：${currentCalendarId}\n請輸入要使用的手術日曆 Calendar ID：`
    : '請輸入要使用的手術日曆 Calendar ID：');
  const response = ui.prompt('設定 Calendar ID', message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) {
    return {
      ok: false,
      status: 'cancelled',
      calendarId: currentCalendarId,
      message: 'Calendar ID 設定已取消。'
    };
  }

  const savedCalendarId = setConfiguredCalendarId_(response.getResponseText());
  if (!savedCalendarId) {
    return {
      ok: false,
      status: 'blank',
      calendarId: currentCalendarId,
      message: 'Calendar ID 不可空白，未更新設定。'
    };
  }

  return {
    ok: true,
    status: 'saved',
    calendarId: savedCalendarId,
    message: `Calendar ID 已儲存：${savedCalendarId}`
  };
}

function promptAndSaveCalendarId() {
  const ui = SpreadsheetApp.getUi();
  const saveResult = promptAndSaveCalendarId_();

  if (!saveResult.ok) {
    ui.alert(saveResult.message);
    return saveResult;
  }

  installProcessRowChangeTrigger_(false);
  const reverseSyncResult = setupCalendarReverseSync_(false);
  const reverseSyncMessage = reverseSyncResult.ok
    ? reverseSyncResult.message
    : `Calendar 反向同步未啟用：${reverseSyncResult.message}`;
  const result = {
    ok: true,
    bindingOk: reverseSyncResult.ok,
    calendarId: saveResult.calendarId,
    saveResult,
    reverseSyncResult
  };

  ui.alert(
    'Calendar ID 設定完成',
    `${saveResult.message}\n` +
      '已安裝 Sheet 自動同步觸發器。\n' +
      `${reverseSyncMessage}\n` +
      '未自動批次同步既有列；若需同步既有資料，請執行「一鍵安裝」。',
    ui.ButtonSet.OK
  );

  return result;
}

function initializeCalendarSyncToken_(calendarId) {
  if (!isCalendarIdConfigured_(calendarId)) {
    return {
      ok: false,
      message: 'Calendar ID 尚未設定，無法建立 Calendar 反向同步 baseline。'
    };
  }

  if (!isCalendarAdvancedServiceAvailable_()) {
    return {
      ok: false,
      message: '尚未啟用 Calendar Advanced Service，無法建立 Calendar syncToken。'
    };
  }

  let pageToken = '';
  let response = null;
  let scannedCount = 0;

  try {
    do {
      const options = {
        maxResults: 2500,
        showDeleted: false
      };

      if (pageToken) {
        options.pageToken = pageToken;
      }

      response = Calendar.Events.list(calendarId, options);
      scannedCount += (response.items || []).length;
      pageToken = response.nextPageToken || '';
    } while (pageToken);
  } catch (err) {
    return {
      ok: false,
      message: `Calendar syncToken baseline 建立失敗：${err.message || err}`
    };
  }

  if (!response || !response.nextSyncToken) {
    return {
      ok: false,
      message: 'Calendar API 未回傳 nextSyncToken，反向同步 baseline 未建立。'
    };
  }

  setCalendarSyncToken_(calendarId, response.nextSyncToken);

  return {
    ok: true,
    scannedCount,
    message: `已建立 Calendar 反向同步 baseline，掃描 ${scannedCount} 個既有事件。`
  };
}

function initializeCalendarSyncToken() {
  const result = initializeCalendarSyncToken_(getConfiguredCalendarId_());
  SpreadsheetApp.getUi().alert(result.message);
}

function installCalendarChangeTrigger_(showAlert) {
  const calendarId = getConfiguredCalendarId_();

  if (!isCalendarIdConfigured_(calendarId)) {
    const result = {
      ok: false,
      message: 'Calendar ID 尚未設定，未安裝 Calendar 反向同步觸發器。'
    };
    if (showAlert) SpreadsheetApp.getUi().alert(result.message);
    return result;
  }

  if (!isCalendarAdvancedServiceAvailable_()) {
    const result = {
      ok: false,
      message: '尚未啟用 Calendar Advanced Service，未安裝 Calendar 反向同步觸發器。'
    };
    if (showAlert) SpreadsheetApp.getUi().alert(result.message);
    return result;
  }

  try {
    ScriptApp.getProjectTriggers()
      .filter(trigger => trigger.getHandlerFunction() === CALENDAR_REVERSE_SYNC_HANDLER)
      .forEach(trigger => ScriptApp.deleteTrigger(trigger));

    ScriptApp.newTrigger(CALENDAR_REVERSE_SYNC_HANDLER)
      .forUserCalendar(calendarId)
      .onEventUpdated()
      .create();

    const result = {
      ok: true,
      message: '已安裝 Calendar 反向同步觸發器。'
    };
    if (showAlert) SpreadsheetApp.getUi().alert(result.message);
    return result;
  } catch (err) {
    const result = {
      ok: false,
      message: `Calendar 反向同步觸發器安裝失敗：${err.message || err}`
    };
    if (showAlert) SpreadsheetApp.getUi().alert(result.message);
    return result;
  }
}

function installCalendarChangeTrigger() {
  installCalendarChangeTrigger_(true);
}

function setupCalendarReverseSync_(showAlert) {
  const calendarId = getConfiguredCalendarId_();
  const tokenResult = initializeCalendarSyncToken_(calendarId);
  if (!tokenResult.ok) {
    if (showAlert) SpreadsheetApp.getUi().alert(tokenResult.message);
    return tokenResult;
  }

  const triggerResult = installCalendarChangeTrigger_(false);
  const result = triggerResult.ok
    ? {
        ok: true,
        message: `${tokenResult.message}\n${triggerResult.message}`
      }
    : triggerResult;

  if (showAlert) {
    SpreadsheetApp.getUi().alert(result.message);
  }

  return result;
}

function setupCalendarReverseSync() {
  setupCalendarReverseSync_(true);
}

/**
 * 一鍵安裝入口。
 *
 * 第一次安裝建議只手動執行這個函式：
 * - 初始化 OP sheet 標題、格式、資料驗證與條件格式
 * - 安裝 Calendar 同步用 installable onEdit trigger
 * - 安裝未來一周清單自動輸出 time-driven trigger
 * - 立即刷新自訂選單
 * - 批次同步既有列，讓 event title/description 等新邏輯套到既有日曆事件
 *
 * 若之後有修改程式、換日曆或懷疑 trigger 重複，也可以重新執行。
 */
function setup() {
  const ui = SpreadsheetApp.getUi();
  const initResult = initializeSheet_(false);

  if (!initResult.ok) {
    ui.alert(initResult.message);
    return;
  }

  let calendarIdPromptMessage = '';
  if (!isCalendarIdConfigured_()) {
    const calendarIdResult = promptAndSaveCalendarId_({
      message: 'Calendar ID 尚未設定。\n請輸入要使用的手術日曆 Calendar ID；若取消，Calendar 同步會略過。'
    });
    calendarIdPromptMessage = `\n${calendarIdResult.message}`;
  }

  const hasCalendarId = isCalendarIdConfigured_();
  const syncResult = hasCalendarId
    ? batchSyncAllEvents_(false)
    : {
        ok: false,
        message: 'Calendar ID 尚未設定。',
        processedCount: 0,
        statusCounts: {}
      };
  const syncMessage = syncResult.ok
    ? '\n' + buildBatchSyncSummaryMessage_(syncResult)
    : `\n批次同步未執行：${syncResult.message}`;
  installProcessRowChangeTrigger_(false);
  const reverseSyncResult = hasCalendarId
    ? setupCalendarReverseSync_(false)
    : {
        ok: false,
        message: 'Calendar ID 尚未設定。'
      };
  const reverseSyncMessage = reverseSyncResult.ok
    ? `\n${reverseSyncResult.message}`
    : `\nCalendar 反向同步未啟用：${reverseSyncResult.message}`;
  const autoExportResult = installAutoExportTriggers_(false);
  const autoExportMessage = autoExportResult.ok
    ? `\n${autoExportResult.message}`
    : `\n自動輸出觸發器未啟用：${autoExportResult.message}`;
  onOpen();

  ui.alert(
    '一鍵安裝完成',
    `已完成表格初始化、Sheet 自動同步觸發器安裝，並刷新「手術排程系統」選單。\n` +
      `時間欄位已轉換 ${initResult.timeResult.normalizedCount} 格，格式錯誤 ${initResult.timeResult.errorCount} 格。` +
      calendarIdPromptMessage +
      syncMessage +
      reverseSyncMessage +
      autoExportMessage +
      initResult.mismatchMessage,
    ui.ButtonSet.OK
  );
}

/**
 * 避免多人或批次同步時，同時建立/更新同一批 Calendar events。
 */
function withCalendarSyncLock_(callback) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    console.warn('無法取得日曆同步鎖，本次執行略過。');
    return false;
  }

  try {
    callback();
    return true;
  } finally {
    lock.releaseLock();
  }
}

function withAutoExportLock_(callback) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();

  if (!lock.tryLock(10000)) {
    console.warn('無法取得自動輸出鎖，本次輸出略過。');
    return false;
  }

  try {
    callback();
    return true;
  } finally {
    lock.releaseLock();
  }
}

/**
 * 解析時間欄位。
 *
 * 接受：
 * - 空白：代表全天事件
 * - 08:30
 * - 8:30
 * - 0830
 * - 830 會視為 08:30
 * - Date 物件形式的時間
 */
function parseTime_(value) {
  if (value === null || value === undefined || value === '') return null;

  if (value instanceof Date && !isNaN(value.getTime())) {
    const hour = value.getHours();
    const minute = value.getMinutes();

    return {
      hour,
      minute,
      text: Utilities.formatString('%02d:%02d', hour, minute)
    };
  }

  // 若 Google Sheets 以小數時間序列值回傳，例如 0.354166... = 08:30
  if (typeof value === 'number' && isFinite(value) && value >= 0 && value < 1) {
    const totalMinutes = Math.round(value * 24 * 60);
    const hour = Math.floor(totalMinutes / 60) % 24;
    const minute = totalMinutes % 60;

    return {
      hour,
      minute,
      text: Utilities.formatString('%02d:%02d', hour, minute)
    };
  }

  const raw = String(value).trim().replace(/：/g, ':');
  if (!raw) return null;

  let hour;
  let minute;

  const colonMatch = raw.match(/^(\d{1,2}):(\d{2})$/);

  if (colonMatch) {
    hour = Number(colonMatch[1]);
    minute = Number(colonMatch[2]);
  } else {
    const compactMatch = raw.match(/^(\d{3,4})$/);

    if (!compactMatch) {
      throw new Error(`時間格式錯誤：${raw}。請使用 HH:mm 或 HHmm，例如 08:30 或 0830。`);
    }

    const compact = compactMatch[1].padStart(4, '0');
    hour = Number(compact.slice(0, 2));
    minute = Number(compact.slice(2, 4));
  }

  if (
    !Number.isInteger(hour) ||
    !Number.isInteger(minute) ||
    hour < 0 ||
    hour > 23 ||
    minute < 0 ||
    minute > 59
  ) {
    throw new Error(`時間範圍錯誤：${raw}。小時需為 0-23，分鐘需為 0-59。`);
  }

  return {
    hour,
    minute,
    text: Utilities.formatString('%02d:%02d', hour, minute)
  };
}

function getRawTimeText_(value) {
  return value === null || value === undefined ? '' : String(value).trim().replace(/：/g, ':');
}

function looksLikeTimeInput_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return true;
  if (typeof value === 'number' && isFinite(value)) return true;

  const raw = getRawTimeText_(value);
  return /^(\d{1,2}):(\d{2})$/.test(raw) || /^\d{3,4}$/.test(raw);
}

/**
 * 解析 Calendar 同步用的時間資訊。
 *
 * 與 parseTime_ 不同：這裡允許 GA 這類非時間文字作為全天事件備註；
 * 但 2460、25:00 等明顯像時間卻無效的輸入仍視為錯誤。
 */
function resolveCalendarTime_(value) {
  const raw = getRawTimeText_(value);

  if (!raw) {
    return {
      parsedTime: null,
      timeNote: '',
      errorMessage: ''
    };
  }

  try {
    return {
      parsedTime: parseTime_(value),
      timeNote: '',
      errorMessage: ''
    };
  } catch (err) {
    if (looksLikeTimeInput_(value)) {
      return {
        parsedTime: null,
        timeNote: '',
        errorMessage: err.message
      };
    }

    return {
      parsedTime: null,
      timeNote: raw,
      errorMessage: ''
    };
  }
}

/**
 * 在時間欄位加上系統錯誤 note。
 */
function setTimeErrorNote_(range, message) {
  range.setNote(TIME_ERROR_NOTE_PREFIX + message);
}

/**
 * 只清除系統產生的時間錯誤 note。
 */
function clearTimeErrorNote_(range) {
  const note = range.getNote();

  if (note && String(note).startsWith(TIME_ERROR_NOTE_PREFIX)) {
    range.clearNote();
  }
}

/**
 * 在病歷號欄位加上系統日曆同步提醒 note。
 */
function setCalendarSyncNote_(range, message) {
  range.setNote(CALENDAR_SYNC_NOTE_PREFIX + message);
}

/**
 * 只清除系統產生的日曆同步提醒 note。
 */
function clearCalendarSyncNote_(range) {
  const note = range.getNote();

  if (note && String(note).startsWith(CALENDAR_SYNC_NOTE_PREFIX)) {
    range.clearNote();
  }
}

/**
 * 將欄位數字轉為 A1 欄名。
 * 例如 8 -> H。
 */
function columnToLetter_(column) {
  let letter = '';
  let temp;

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}

function columnLetterToNumber_(letter) {
  const text = String(letter || '').trim().toUpperCase();
  if (!/^[A-Z]+$/.test(text)) return 0;

  let column = 0;
  for (let i = 0; i < text.length; i++) {
    column = column * 26 + text.charCodeAt(i) - 64;
  }

  return column;
}

/**
 * 將字串安全轉成 Google Sheets 公式字串。
 */
function toSheetFormulaString_(value) {
  return `"${String(value).replace(/"/g, '""')}"`;
}

function toCellText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function toSingleLineText_(value) {
  return toCellText_(value).replace(/\s+/g, ' ');
}

function mergeText_(a, b) {
  a = String(a || '').trim();
  b = String(b || '').trim();
  if (!a) return b;
  if (!b) return a;
  if (a === b) return a;
  return a + "\n" + b;
}

function getFieldHeaderToKeyMap_() {
  const map = {};

  CONFIG.FIELD_KEYS.forEach(fieldKey => {
    map[CONFIG.FIELD_HEADERS[fieldKey]] = fieldKey;
  });

  return map;
}

function getDefaultColumnForField_(fieldKey) {
  const index = CONFIG.HEADERS.indexOf(CONFIG.FIELD_HEADERS[fieldKey]);
  return index === -1 ? 0 : index + 1;
}

function getSheetHeaderValues_(sheet) {
  const scanColumnCount = Math.max(
    1,
    Math.min(sheet.getMaxColumns(), Math.max(sheet.getLastColumn(), CONFIG.HEADERS.length))
  );

  return sheet.getRange(1, 1, 1, scanColumnCount).getValues()[0];
}

function getLastHeaderColumnFromValues_(headerValues) {
  for (let i = headerValues.length - 1; i >= 0; i--) {
    if (toCellText_(headerValues[i])) return i + 1;
  }

  return 0;
}

function getLastHeaderColumn_(sheet) {
  const headerValues = getSheetHeaderValues_(sheet);
  return getLastHeaderColumnFromValues_(headerValues) || Math.min(sheet.getMaxColumns(), CONFIG.HEADERS.length);
}

function formatColumnList_(columns) {
  return columns.map(columnToLetter_).join(', ');
}

function formatFieldHeaderList_(fieldKeys) {
  return fieldKeys
    .map(fieldKey => `「${CONFIG.FIELD_HEADERS[fieldKey] || fieldKey}」`)
    .join('、');
}

function buildFieldColumnInfo_(sheet) {
  const headerValues = getSheetHeaderValues_(sheet);
  const headerToKey = getFieldHeaderToKeyMap_();
  const columns = {};
  const canonicalColumns = {};
  const duplicateColumnsByHeader = {};
  const legacyEventIdColumns = [];

  headerValues.forEach((value, index) => {
    const header = toCellText_(value);
    if (!header) return;

    if (LEGACY_EVENT_ID_HEADERS.indexOf(header) !== -1) {
      legacyEventIdColumns.push(index + 1);
    }

    const fieldKey = headerToKey[header];
    if (!fieldKey) return;

    if (columns[fieldKey]) {
      if (!duplicateColumnsByHeader[header]) {
        duplicateColumnsByHeader[header] = [columns[fieldKey]];
      }
      duplicateColumnsByHeader[header].push(index + 1);
      return;
    }

    columns[fieldKey] = index + 1;
    canonicalColumns[fieldKey] = index + 1;
  });

  if (!columns.EVENT_ID && legacyEventIdColumns.length > 0) {
    columns.EVENT_ID = legacyEventIdColumns[0];
  }

  const missingKeys = CONFIG.FIELD_KEYS.filter(fieldKey => !columns[fieldKey]);
  const missingCanonicalKeys = CONFIG.FIELD_KEYS.filter(fieldKey => !canonicalColumns[fieldKey]);
  const duplicateMessages = Object.keys(duplicateColumnsByHeader).map(header => {
    return `「${header}」出現在 ${formatColumnList_(duplicateColumnsByHeader[header])} 欄，系統會使用最左側欄位。`;
  });

  return {
    columns,
    canonicalColumns,
    missingKeys,
    missingCanonicalKeys,
    duplicateMessages,
    legacyEventIdColumns,
    lastHeaderColumn: getLastHeaderColumnFromValues_(headerValues) || CONFIG.HEADERS.length
  };
}

function getRequiredSheetColumnInfo_(sheet) {
  const info = buildFieldColumnInfo_(sheet);

  if (info.missingKeys.length > 0) {
    throw new Error(`缺少必要欄位 ${formatFieldHeaderList_(info.missingKeys)}，請先執行「一鍵安裝」或「初始化表格格式」。`);
  }

  return info;
}

function getRequiredSheetColumns_(sheet) {
  return getRequiredSheetColumnInfo_(sheet).columns;
}

function getEventIdColumnForSheet_(sheet) {
  const info = buildFieldColumnInfo_(sheet);
  return info.canonicalColumns.EVENT_ID || info.columns.EVENT_ID || 0;
}

function getTableLastColumn_(sheet, columns) {
  const columnValues = columns ? Object.values(columns) : [];
  const requiredLastColumn = columnValues.length > 0
    ? Math.max.apply(null, columnValues)
    : CONFIG.HEADERS.length;

  return Math.max(getLastHeaderColumn_(sheet), requiredLastColumn);
}

function getRowFieldValue_(rowValues, columns, fieldKey) {
  return rowValues[columns[fieldKey] - 1];
}

function setRowFieldValue_(rowValues, columns, fieldKey, value) {
  const index = columns[fieldKey] - 1;

  while (rowValues.length <= index) {
    rowValues.push('');
  }

  rowValues[index] = value;
}

function getOpSheetOrThrow_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_OP);

  if (!sheet) {
    throw new Error(`找不到「${CONFIG.SHEET_OP}」工作表。`);
  }

  return sheet;
}

function getOpSheetContext_() {
  const sheet = getOpSheetOrThrow_();

  return {
    sheet,
    columns: getRequiredSheetColumns_(sheet)
  };
}

function runMenuAction_(title, callback) {
  const ui = SpreadsheetApp.getUi();

  try {
    return callback(ui);
  } catch (err) {
    ui.alert(title, err.message || String(err), ui.ButtonSet.OK);
    return null;
  }
}

function assertActiveOpSheet_(sheet) {
  if (!sheet || sheet.getName() !== CONFIG.SHEET_OP) {
    throw new Error(`請先切換到「${CONFIG.SHEET_OP}」工作表後再執行。`);
  }
}

function buildCalendarTitle_(chartNoText, patientNameText, conditionText) {
  return [chartNoText, patientNameText, conditionText]
    .map(value => toSingleLineText_(value))
    .join(` ${CALENDAR_TITLE_SEPARATOR} `);
}

function buildCalendarDescription_(telText, planText, timeNote) {
  const lines = [
    `TEL: ${telText}`,
    `Plan: ${planText}`
  ];

  if (timeNote) {
    lines.push(`Time note: ${timeNote}`);
  }

  return lines.join('\n');
}

function parseCalendarTitle_(title) {
  const raw = toSingleLineText_(title);
  if (!raw) {
    return {
      ok: false,
      message: '標題空白。'
    };
  }

  const pipeParts = raw.replace(/｜/g, CALENDAR_TITLE_SEPARATOR)
    .split(CALENDAR_TITLE_SEPARATOR)
    .map(part => part.trim());

  if (pipeParts.length >= 2 && pipeParts[0]) {
    return {
      ok: true,
      chartNo: pipeParts[0],
      patientName: pipeParts[1] || '',
      condition: pipeParts.slice(2).join(` ${CALENDAR_TITLE_SEPARATOR} `).trim(),
      format: 'pipe'
    };
  }

  const legacyMatch = raw.match(/^(\S+)\s+(.+?)(?:\s+-\s+(.+))?$/);
  if (legacyMatch && legacyMatch[1] && legacyMatch[2]) {
    return {
      ok: true,
      chartNo: legacyMatch[1].trim(),
      patientName: legacyMatch[2].trim(),
      condition: (legacyMatch[3] || '').trim(),
      format: 'legacy'
    };
  }

  return {
    ok: false,
    message: '標題格式錯誤。請使用「病歷號 | 姓名 | Condition」。'
  };
}

function parseCalendarDescription_(description) {
  const lines = String(description || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const result = {
    hasTel: false,
    hasPlan: false,
    hasTimeNote: false,
    tel: '',
    plan: '',
    timeNote: ''
  };
  let activeField = '';

  lines.forEach(line => {
    const telMatch = line.match(/^TEL\s*[:：]\s*(.*)$/i);
    if (telMatch) {
      activeField = 'tel';
      result.hasTel = true;
      result.tel = telMatch[1].trim();
      return;
    }

    const planMatch = line.match(/^Plan\s*[:：]\s*(.*)$/i);
    if (planMatch) {
      activeField = 'plan';
      result.hasPlan = true;
      result.plan = planMatch[1].trim();
      return;
    }

    const timeNoteMatch = line.match(/^Time note\s*[:：]\s*(.*)$/i);
    if (timeNoteMatch) {
      activeField = '';
      result.hasTimeNote = true;
      result.timeNote = timeNoteMatch[1].trim();
      return;
    }

    if (activeField === 'plan') {
      result.plan = result.plan ? `${result.plan}\n${line}` : line;
    } else if (activeField === 'tel' && line.trim()) {
      result.tel = result.tel ? `${result.tel} ${line.trim()}` : line.trim();
    }
  });

  return result;
}

/**
 * 建立試算表上方的自訂選單。
 *
 * onOpen 只負責 UI 選單；它不會安裝 onEdit trigger。
 * 第一次使用或需要重新授權時，仍需執行 setup() 或選單中的一鍵安裝。
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const maintenanceMenu = ui.createMenu('維護工具')
    .addItem('完整安裝／批次同步', 'setup')
    .addSeparator()
    .addItem('設定日曆', 'promptAndSaveCalendarId')
    .addItem('初始化表格', 'initializeSheet')
    .addItem('資料遷移（不批次同步）', 'runMigrations')
    .addItem('安裝同步', 'installProcessRowChangeTrigger')
    .addItem('安裝反向同步', 'setupCalendarReverseSync')
    .addItem('安裝自動輸出', 'installAutoExportTriggers')
    .addItem('清除舊事件', 'clearCalendarEventsFromSpecifiedColumn')
    .addSeparator()
    .addItem(`版本：${CONFIG.VERSION}`, 'showVersionInfo');

  ui.createMenu('手術排程系統')
    .addItem('輸出日期', 'exportDateData')
    .addItem('輸出一週', 'exportUpcomingWeekDataFromMenu')
    .addItem('複製列', 'duplicateRow')
    .addItem('歸人整合', 'mergePatientRecords')
    .addItem('日期排序', 'sortSheetByDateTime')
    .addSeparator()
    .addSubMenu(maintenanceMenu)
    .addToUi();
}

function showVersionInfo() {
  SpreadsheetApp.getUi().alert(`手術排程系統版本：${CONFIG.VERSION}`);
}

function splitFirstLineAndRest_(value) {
  const lines = String(value || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const firstLine = (lines.shift() || '').trim();
  const rest = lines.join('\n').trim();

  return {
    firstLine,
    rest
  };
}

function joinNonEmptyLines_(values) {
  return values
    .map(value => String(value || '').trim())
    .filter(Boolean)
    .join('\n');
}

/**
 * 1: 輸出指令日期資料
 */
function formatExportDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function formatExportDateWithWeekday_(date) {
  const dateText = formatExportDate_(date);
  const weekdayText = Utilities.formatDate(date, Session.getScriptTimeZone(), 'E');
  return `${dateText}(${weekdayText})`;
}

function parseExportDate_(value) {
  const date = new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

function buildSurgeryExportItemsFromRow_(row, cols) {
  const conditionStr = String(getRowFieldValue_(row, cols, 'COND'));
  let disease = conditionStr;
  let surgery = '';

  if (conditionStr.includes('s/p')) {
    const parts = conditionStr.split('s/p');
    disease = parts[0].trim();
    surgery = parts.slice(1).join('s/p').trim();
  }

  const surgeryParts = splitFirstLineAndRest_(surgery);
  const surgeryMain = surgeryParts.firstLine;
  const supplementalNote = joinNonEmptyLines_([
    surgeryParts.rest,
    getRowFieldValue_(row, cols, 'PLAN')
  ]);
  const patientRow = [
    getRowFieldValue_(row, cols, 'CHART_NO'),
    getRowFieldValue_(row, cols, 'NAME'),
    getRowFieldValue_(row, cols, 'TEL'),
    getRowFieldValue_(row, cols, 'TIME'),
    disease,
    surgeryMain,
    supplementalNote
  ];
  const iolMatch = surgery.match(/IOL\(([^)]+)\)/i) || surgery.match(/MSICS\(([^)]+)\)/i);
  let iolRow = null;

  if (iolMatch) {
    const iolDetails = iolMatch[1].trim().split(/\s+/);
    iolRow = [
      getRowFieldValue_(row, cols, 'NAME'),
      iolDetails[0] || '',
      iolDetails[2] || '',
      iolDetails[1] || ''
    ];
  }

  return {
    iolRow,
    patientRow
  };
}

function buildSurgeryExportDataForDate_(data, cols, targetDateObj) {
  const targetDateText = formatExportDate_(targetDateObj);
  const iolList = [];
  const patientList = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tagVal = String(getRowFieldValue_(row, cols, 'TAG')).trim();
    const dateVal = getRowFieldValue_(row, cols, 'DATE');

    if (!dateVal || tagVal !== 'OP') continue;

    const rowDateObj = new Date(dateVal);
    if (isNaN(rowDateObj.getTime())) continue;

    if (formatExportDate_(rowDateObj) !== targetDateText) continue;

    const items = buildSurgeryExportItemsFromRow_(row, cols);
    if (items.iolRow) iolList.push(items.iolRow);
    patientList.push(items.patientRow);
  }

  iolList.sort((a, b) => {
    const brandCompare = String(a[1] || '').localeCompare(String(b[1] || ''), 'zh-Hant');
    if (brandCompare !== 0) return brandCompare;

    return String(a[0] || '').localeCompare(String(b[0] || ''), 'zh-Hant');
  });

  return {
    date: new Date(targetDateObj.getTime()),
    dateText: targetDateText,
    dateLabel: formatExportDateWithWeekday_(targetDateObj),
    iolList,
    patientList
  };
}

function buildUpcomingExportDates_(startDateObj, dayCount) {
  const baseDate = new Date(startDateObj.getTime());
  const dates = [];

  baseDate.setHours(0, 0, 0, 0);

  for (let i = 0; i < dayCount; i++) {
    dates.push(addDays_(baseDate, i));
  }

  return dates;
}

function getSurgeryExportSourceData_() {
  const context = getOpSheetContext_();
  const sheetOp = context.sheet;
  const cols = context.columns;
  const lastRow = sheetOp.getLastRow();
  const data = lastRow > 0
    ? sheetOp.getRange(1, 1, lastRow, getTableLastColumn_(sheetOp, cols)).getValues()
    : [];

  return {
    columns: cols,
    data
  };
}

function getOrCreateOutputSheet_(spreadsheet) {
  return spreadsheet.getSheetByName(CONFIG.SHEET_OUT) || spreadsheet.insertSheet(CONFIG.SHEET_OUT);
}

function writeSurgeryExportSection_(sheetOut, startRow, exportData) {
  const iolStartCol = 1;
  const patientStartCol = 6;

  sheetOut.getRange(startRow, iolStartCol)
    .setValue(`[ ${exportData.dateLabel} 水晶體清單 ]`)
    .setFontWeight('bold');

  if (exportData.iolList.length > 0) {
    sheetOut.getRange(startRow + 1, iolStartCol, 1, 4)
      .setValues([['姓名', '品牌', '水晶體度數', '目標度數']])
      .setBackground('#efefef');
    sheetOut.getRange(startRow + 2, iolStartCol, exportData.iolList.length, 4)
      .setValues(exportData.iolList);
  } else {
    sheetOut.getRange(startRow + 1, iolStartCol).setValue('本日無水晶體資料');
  }

  sheetOut.getRange(startRow, patientStartCol)
    .setValue(`[ ${exportData.dateLabel} 病人清單 ]`)
    .setFontWeight('bold');

  if (exportData.patientList.length > 0) {
    sheetOut.getRange(startRow + 1, patientStartCol, 1, 7)
      .setValues([['病歷號', '姓名', 'TEL', '時間', '疾病', '術式', '補充說明']])
      .setBackground('#efefef');
    sheetOut.getRange(startRow + 2, patientStartCol, exportData.patientList.length, 7)
      .setValues(exportData.patientList);
  } else {
    sheetOut.getRange(startRow + 1, patientStartCol).setValue('本日無病人資料');
  }

  const iolBlockHeight = exportData.iolList.length > 0 ? exportData.iolList.length + 2 : 2;
  const patientBlockHeight = exportData.patientList.length > 0 ? exportData.patientList.length + 2 : 2;

  return startRow + Math.max(iolBlockHeight, patientBlockHeight) + 2;
}

function writeSingleDateExport_(targetDateObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = getSurgeryExportSourceData_();
  const sheetOut = getOrCreateOutputSheet_(ss);
  const exportData = buildSurgeryExportDataForDate_(source.data, source.columns, targetDateObj);

  sheetOut.clear();
  writeSurgeryExportSection_(sheetOut, 1, exportData);
  sheetOut.autoResizeColumns(1, 12);

  return {
    ok: true,
    startDateText: exportData.dateText,
    endDateText: exportData.dateText,
    dayCount: 1,
    exports: [exportData]
  };
}

function writeUpcomingWeekExport_(startDateObj) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = getSurgeryExportSourceData_();
  const sheetOut = getOrCreateOutputSheet_(ss);
  const targetDates = buildUpcomingExportDates_(startDateObj, AUTO_EXPORT_UPCOMING_DAYS);
  const exports = [];
  let nextRow = 1;

  sheetOut.clear();

  targetDates.forEach(date => {
    const exportData = buildSurgeryExportDataForDate_(source.data, source.columns, date);
    if (exportData.patientList.length === 0) return;

    exports.push(exportData);
    nextRow = writeSurgeryExportSection_(sheetOut, nextRow, exportData);
  });

  sheetOut.autoResizeColumns(1, 12);

  return {
    ok: true,
    startDateText: exports.length ? exports[0].dateText : '',
    endDateText: exports.length ? exports[exports.length - 1].dateText : '',
    dayCount: exports.length,
    exports
  };
}

function exportSingleDateData_(targetDateObj) {
  let result = null;
  const executed = withAutoExportLock_(() => {
    result = writeSingleDateExport_(targetDateObj);
  });

  if (!executed) {
    throw new Error('目前有另一個輸出作業正在執行，本次輸出未執行。');
  }

  return result;
}

function exportUpcomingWeekData_() {
  let result = null;
  const executed = withAutoExportLock_(() => {
    result = writeUpcomingWeekExport_(new Date());
  });

  if (!executed) {
    return {
      ok: false,
      status: 'locked',
      message: '目前有另一個輸出作業正在執行，本次自動輸出未執行。'
    };
  }

  return result;
}

function exportDateData() {
  return runMenuAction_('輸出指令日期資料', ui => {
    const response = ui.prompt('輸出指定日期資料', '請輸入日期 (格式如: 2026/5/11):', ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() !== ui.Button.OK) return null;

    const targetDateObj = parseExportDate_(response.getResponseText().trim());
    if (!targetDateObj) {
      ui.alert('錯誤', '日期格式不正確，請重新執行。', ui.ButtonSet.OK);
      return null;
    }

    const result = exportSingleDateData_(targetDateObj);
    ui.alert('完成', `已成功將 ${result.startDateText} 的資料分析匯出至「輸出表單」。`, ui.ButtonSet.OK);
    return result;
  });
}

function exportUpcomingWeekData() {
  return exportUpcomingWeekData_();
}

function exportUpcomingWeekDataFromMenu() {
  return runMenuAction_('立即輸出未來一周', ui => {
    const result = exportUpcomingWeekData_();

    if (!result.ok) {
      throw new Error(result.message);
    }

    ui.alert(
      '完成',
      `已成功將 ${result.startDateText} 至 ${result.endDateText} 的資料匯出至「輸出表單」。`,
      ui.ButtonSet.OK
    );

    return result;
  });
}

function deletePendingAutoExportTriggers_() {
  return deleteTriggersByHandler_(AUTO_EXPORT_PENDING_HANDLER);
}

function scheduleUpcomingWeekExport_() {
  try {
    deletePendingAutoExportTriggers_();
    ScriptApp.newTrigger(AUTO_EXPORT_PENDING_HANDLER)
      .timeBased()
      .after(AUTO_EXPORT_EDIT_DELAY_MS)
      .create();

    return {
      ok: true,
      message: '已排程 5 分鐘後更新未來一周輸出表單。'
    };
  } catch (err) {
    const message = `排程未來一周自動輸出失敗：${err.message || err}`;
    console.warn(message);

    return {
      ok: false,
      message
    };
  }
}

function runPendingUpcomingWeekExport() {
  deletePendingAutoExportTriggers_();
  const result = exportUpcomingWeekData_();

  if (result && result.status === 'locked') {
    scheduleUpcomingWeekExport_();
  }

  return result;
}

function installAutoExportTriggers_(showAlert) {
  let deletedCount = 0;

  try {
    deletedCount += deleteTriggersByHandler_(AUTO_EXPORT_DAILY_HANDLER);
    deletedCount += deletePendingAutoExportTriggers_();

    ScriptApp.newTrigger(AUTO_EXPORT_DAILY_HANDLER)
      .timeBased()
      .atHour(AUTO_EXPORT_DAILY_HOUR)
      .nearMinute(AUTO_EXPORT_DAILY_MINUTE)
      .everyDays(1)
      .create();

    const result = {
      ok: true,
      deletedCount,
      message: '已安裝未來一周自動輸出觸發器：每日約 06:00 執行，修改後會延遲 5 分鐘更新。'
    };

    if (showAlert) {
      SpreadsheetApp.getUi().alert(result.message);
    }

    return result;
  } catch (err) {
    const result = {
      ok: false,
      deletedCount,
      message: err.message || String(err)
    };

    if (showAlert) {
      SpreadsheetApp.getUi().alert(`自動輸出觸發器安裝失敗：${result.message}`);
    }

    return result;
  }
}

function installAutoExportTriggers() {
  return runMenuAction_('安裝自動輸出觸發器', () => installAutoExportTriggers_(true));
}

/**
 * 2: 複製一列
 */
function duplicateRow() {
  return runMenuAction_('複製一列', () => {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  assertActiveOpSheet_(sheet);
  const activeCell = sheet.getActiveCell();
  const activeRow = activeCell.getRow();

  if (activeRow < 2) return;

  const lastCol = sheet.getLastColumn();
  const cols = getRequiredSheetColumns_(sheet);
  sheet.insertRowAfter(activeRow);

  const sourceRange = sheet.getRange(activeRow, 1, 1, lastCol);
  const targetRange = sheet.getRange(activeRow + 1, 1, 1, lastCol);

  sourceRange.copyTo(targetRange);

  sheet.getRange(activeRow + 1, cols.DATE).clearContent();
  sheet.getRange(activeRow + 1, cols.TIME).clearContent();
  sheet.getRange(activeRow + 1, cols.EVENT_ID).clearContent();
  sheet.getRange(activeRow + 1, cols.SHEET_WRITE_UPDATED).clearContent();
  });
}

/**
 * 3: 歸人整合
 */
function mergePatientRecords() {
  return runMenuAction_('歸人整合', ui => {
  const context = getOpSheetContext_();
  const sheet = context.sheet;
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  if (data.length < 2) return;

  const cols = context.columns;
  const result = buildMergedPatientRows_(data, cols);

  if (!result.ok) {
    ui.alert(
      '歸人整合中止',
      '以下將被合併移除的列含有日期、時間或日曆 eventID，為避免日曆事件失去追蹤，本次未更動資料：\n' +
        result.blockedRows.join(', '),
      ui.ButtonSet.OK
    );
    return;
  }

  const finalRows = result.rows;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // 從第 2 列開始清除，所以列數應該是 lastRow - 1
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }

  if (finalRows.length > 0) {
    sheet.getRange(2, 1, finalRows.length, data[0].length).setValues(finalRows);
  }

  ui.alert('歸人整合完成', '所有非 OP 的同病人紀錄已合併。', ui.ButtonSet.OK);
  });
}

function buildMergedPatientRows_(data, cols) {
  const grouped = {};
  const finalRows = [];
  const blockedRows = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i].slice();
    const chartNo = getRowFieldValue_(row, cols, 'CHART_NO');
    const name = getRowFieldValue_(row, cols, 'NAME');
    const tag = getRowFieldValue_(row, cols, 'TAG');

    if (tag === 'OP' || !chartNo || !name) {
      finalRows.push(row);
    } else {
      const key = `${chartNo}_${name}`;
      if (!grouped[key]) {
        grouped[key] = row;
        finalRows.push(grouped[key]);
      } else {
        if (rowHasCalendarTrackingData_(row, cols)) {
          blockedRows.push(i + 1);
          continue;
        }

        setRowFieldValue_(grouped[key], cols, 'TEL', mergeText_(getRowFieldValue_(grouped[key], cols, 'TEL'), getRowFieldValue_(row, cols, 'TEL')));
        setRowFieldValue_(grouped[key], cols, 'COND', mergeText_(getRowFieldValue_(grouped[key], cols, 'COND'), getRowFieldValue_(row, cols, 'COND')));
        setRowFieldValue_(grouped[key], cols, 'PLAN', mergeText_(getRowFieldValue_(grouped[key], cols, 'PLAN'), getRowFieldValue_(row, cols, 'PLAN')));
        setRowFieldValue_(grouped[key], cols, 'MEMO', mergeText_(getRowFieldValue_(grouped[key], cols, 'MEMO'), getRowFieldValue_(row, cols, 'MEMO')));

        if (getRowFieldValue_(grouped[key], cols, 'TAG') !== getRowFieldValue_(row, cols, 'TAG')) {
          setRowFieldValue_(grouped[key], cols, 'TAG', mergeText_(getRowFieldValue_(grouped[key], cols, 'TAG'), getRowFieldValue_(row, cols, 'TAG')));
        }
      }
    }
  }

  return {
    ok: blockedRows.length === 0,
    rows: finalRows,
    blockedRows
  };
}

function rowHasCalendarTrackingData_(row, cols) {
  return Boolean(
    getRowFieldValue_(row, cols, 'DATE') ||
      getRowFieldValue_(row, cols, 'TIME') ||
      toCellText_(getRowFieldValue_(row, cols, 'EVENT_ID'))
  );
}

/**
 * 4: 依照日期與時間排序
 */
function sortSheetByDateTime() {
  return runMenuAction_('依照日期與時間排序', () => {
  const context = getOpSheetContext_();
  const sheet = context.sheet;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  const cols = context.columns;
  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  range.sort([
    { column: cols.DATE, ascending: true },
    { column: cols.TIME, ascending: true }
  ]);
  });
}

function ensureMinimumSheetSize_(sheet) {
  const requiredRows = 2;
  const requiredCols = CONFIG.HEADERS.length;

  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  if (sheet.getMaxColumns() < requiredCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), requiredCols - sheet.getMaxColumns());
  }
}

function ensureSheetHasColumn_(sheet, column) {
  if (sheet.getMaxColumns() < column) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), column - sheet.getMaxColumns());
  }
}

function migrateLegacyEventIdColumns_(sheet) {
  const info = buildFieldColumnInfo_(sheet);
  let currentEventIdColumn = info.canonicalColumns.EVENT_ID || 0;
  const legacyColumns = info.legacyEventIdColumns.slice();

  if (!currentEventIdColumn || legacyColumns.length === 0) {
    return {
      changed: false,
      message: ''
    };
  }

  let deletedLegacyCount = 0;
  let copiedValueCount = 0;
  const conflictMessages = [];

  legacyColumns.sort((a, b) => b - a).forEach(legacyColumn => {
    const lastRow = sheet.getLastRow();
    const legacyHeader = toCellText_(sheet.getRange(1, legacyColumn).getValue());
    const conflictedRows = [];

    if (lastRow >= 2 && currentEventIdColumn && legacyColumn !== currentEventIdColumn) {
      const currentRange = sheet.getRange(2, currentEventIdColumn, lastRow - 1, 1);
      const legacyValues = sheet.getRange(2, legacyColumn, lastRow - 1, 1).getValues();
      const currentValues = currentRange.getValues();
      let hasCopiedValues = false;

      legacyValues.forEach((rowValues, index) => {
        const legacyValue = toCellText_(rowValues[0]);
        const currentValue = toCellText_(currentValues[index][0]);

        if (!legacyValue) return;

        if (!currentValue) {
          currentValues[index][0] = legacyValue;
          copiedValueCount++;
          hasCopiedValues = true;
        } else if (currentValue !== legacyValue) {
          conflictedRows.push(index + 2);
        }
      });

      if (hasCopiedValues) {
        currentRange.setValues(currentValues);
      }
    }

    if (conflictedRows.length > 0) {
      conflictMessages.push(
        `「${legacyHeader}」${columnToLetter_(legacyColumn)} 欄有 ${conflictedRows.length} 列與「${CONFIG.FIELD_HEADERS.EVENT_ID}」不同，已保留舊欄。`
      );
      return;
    }

    sheet.deleteColumn(legacyColumn);
    deletedLegacyCount++;

    if (legacyColumn < currentEventIdColumn) {
      currentEventIdColumn--;
    }
  });

  const messageParts = [];
  if (deletedLegacyCount > 0) {
    messageParts.push(`已移除 ${deletedLegacyCount} 個舊系統欄位（${LEGACY_EVENT_ID_HEADERS.map(header => `「${header}」`).join('、')}）。`);
  }
  if (copiedValueCount > 0) {
    messageParts.push(`已從舊欄位補回 ${copiedValueCount} 格 Calendar event ID。`);
  }
  conflictMessages.forEach(message => messageParts.push(message));

  return {
    changed: messageParts.length > 0,
    message: messageParts.join('\n')
  };
}

function ensureHeaders_(sheet) {
  const headerValues = getSheetHeaderValues_(sheet);
  const hasAnyHeader = headerValues.some(value => Boolean(toCellText_(value)));
  const addedHeaders = [];

  if (!hasAnyHeader) {
    ensureSheetHasColumn_(sheet, CONFIG.HEADERS.length);
    sheet.getRange(1, 1, 1, CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);

    const emptySheetInfo = buildFieldColumnInfo_(sheet);
    return {
      columns: emptySheetInfo.columns,
      duplicateMessages: emptySheetInfo.duplicateMessages,
      addedHeaders: CONFIG.HEADERS.slice()
    };
  }

  let info = buildFieldColumnInfo_(sheet);
  let nextColumn = getLastHeaderColumn_(sheet) + 1;

  info.missingCanonicalKeys.forEach(fieldKey => {
    ensureSheetHasColumn_(sheet, nextColumn);
    const header = CONFIG.FIELD_HEADERS[fieldKey];
    sheet.getRange(1, nextColumn).setValue(header);
    addedHeaders.push(header);
    nextColumn++;
  });

  info = buildFieldColumnInfo_(sheet);

  return {
    columns: info.columns,
    duplicateMessages: info.duplicateMessages,
    addedHeaders
  };
}

function applyDataFormats_(sheet, columns) {
  const dataRowCount = sheet.getMaxRows() - 1;
  if (dataRowCount <= 0) return;

  sheet.getRange(2, columns.TEL, dataRowCount, 1).setNumberFormat('@');
  sheet.getRange(2, columns.TIME, dataRowCount, 1).setNumberFormat('@');
  sheet.getRange(2, columns.DATE, dataRowCount, 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(2, columns.EVENT_ID, dataRowCount, 1).setNumberFormat('@');
  sheet.getRange(2, columns.SHEET_WRITE_UPDATED, dataRowCount, 1).setNumberFormat('@');

  const tagRange = sheet.getRange(2, columns.TAG, dataRowCount, 1);
  const ruleValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.TAG_OPTIONS, true)
    .setAllowInvalid(true)
    .build();
  tagRange.setDataValidation(ruleValidation);
}

function applyTableAlignment_(sheet, columns) {
  const lastColumn = getTableLastColumn_(sheet, columns);
  if (lastColumn <= 0) return;

  const rowCount = Math.max(sheet.getLastRow(), 2);
  sheet.getRange(1, 1, rowCount, lastColumn)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
}

function applyRowDataFormats_(sheet, row, columns) {
  const cols = columns || getRequiredSheetColumns_(sheet);
  sheet.getRange(row, cols.TEL).setNumberFormat('@');
  sheet.getRange(row, cols.TIME).setNumberFormat('@');
  sheet.getRange(row, cols.DATE).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(row, cols.EVENT_ID).setNumberFormat('@');
  sheet.getRange(row, cols.SHEET_WRITE_UPDATED).setNumberFormat('@');

  const ruleValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.TAG_OPTIONS, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(row, cols.TAG).setDataValidation(ruleValidation);

  sheet.getRange(row, 1, 1, getTableLastColumn_(sheet, cols))
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
}

function hideSystemColumns_(sheet, columns) {
  CONFIG.SYSTEM_FIELD_KEYS.forEach(fieldKey => {
    if (columns[fieldKey]) {
      sheet.hideColumns(columns[fieldKey]);
    }
  });
}

function getSystemColumnProtectionDescription_(fieldKey) {
  return SYSTEM_COLUMN_PROTECTION_PREFIX + fieldKey;
}

function removeManagedSystemColumnProtections_(sheet, description) {
  if (!SpreadsheetApp.ProtectionType) return;

  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(protection => protection.getDescription() === description)
    .forEach(protection => protection.remove());
}

function getUserEmail_(user) {
  if (!user) return '';
  if (typeof user === 'string') return user.trim();

  try {
    return toCellText_(user.getEmail && user.getEmail()).trim();
  } catch (err) {
    return '';
  }
}

function getConfiguredSystemColumnProtectionEditors_() {
  try {
    const value = PropertiesService.getScriptProperties()
      .getProperty(SYSTEM_COLUMN_PROTECTION_EDITORS_PROPERTY_KEY);
    return String(value || '')
      .split(/[\n,;]/)
      .map(email => email.trim())
      .filter(Boolean);
  } catch (err) {
    return [];
  }
}

function getSystemColumnProtectionEditors_() {
  const editors = [];

  try {
    const effectiveUser = Session.getEffectiveUser();
    if (effectiveUser) editors.push(effectiveUser);
  } catch (err) {
    // Apps Script may hide the effective user in some trigger contexts.
  }

  getConfiguredSystemColumnProtectionEditors_().forEach(email => editors.push(email));

  const seenEmails = {};
  return editors.filter(editor => {
    const email = getUserEmail_(editor).toLowerCase();
    if (!email) return true;
    if (seenEmails[email]) return false;
    seenEmails[email] = true;
    return true;
  });
}

function protectSystemColumn_(sheet, column, description) {
  removeManagedSystemColumnProtections_(sheet, description);

  const protection = sheet.getRange(1, column, sheet.getMaxRows(), 1)
    .protect()
    .setDescription(description)
    .setWarningOnly(false);
  const allowedEditors = getSystemColumnProtectionEditors_();
  const allowedEmails = {};

  allowedEditors.forEach(editor => {
    const email = getUserEmail_(editor).toLowerCase();
    if (email) allowedEmails[email] = true;

    try {
      protection.addEditor(editor);
    } catch (err) {
      try {
        protection.setWarningOnly(true);
      } catch (warningErr) {
        // Ignore cleanup errors; the caller will surface the original failure.
      }
      try {
        protection.remove();
      } catch (removeErr) {
        // Ignore cleanup errors; the caller will surface the original failure.
      }
      throw new Error(`無法保留系統欄保護編輯者「${email || editor}」：${err.message || err}`);
    }
  });

  const allowedEmailList = Object.keys(allowedEmails);
  if (allowedEmailList.length > 0) {
    const removableEditors = protection.getEditors().filter(editor => {
      const email = getUserEmail_(editor).toLowerCase();
      return email && !allowedEmails[email];
    });

    if (removableEditors.length > 0) {
      protection.removeEditors(removableEditors);
    }
  }

  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

function protectSystemColumns_(sheet, columns) {
  const failedMessages = [];

  CONFIG.SYSTEM_FIELD_KEYS.forEach(fieldKey => {
    const column = columns[fieldKey];
    if (!column) return;

    try {
      protectSystemColumn_(sheet, column, getSystemColumnProtectionDescription_(fieldKey));
    } catch (err) {
      failedMessages.push(`${CONFIG.FIELD_HEADERS[fieldKey]}：${err.message || err}`);
    }
  });

  return {
    ok: failedMessages.length === 0,
    message: failedMessages.length > 0
      ? `系統欄位保護設定失敗：\n${failedMessages.join('\n')}`
      : ''
  };
}

function getConditionalRuleFormula_(rule) {
  const condition = rule.getBooleanCondition();

  if (!condition || condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
    return '';
  }

  const values = condition.getCriteriaValues();
  return values && values.length > 0 ? String(values[0]) : '';
}

function conditionalRuleHasSystemMarker_(rule) {
  const formula = getConditionalRuleFormula_(rule);
  return Object.values(CONDITIONAL_FORMAT_MARKERS).some(marker => formula.indexOf(marker) !== -1);
}

function rangeMatchesSingleColumn_(range, column) {
  return range.getRow() === 2 && range.getColumn() === column && range.getNumColumns() === 1;
}

function rangeMatchesMainTable_(range, sheet) {
  return range.getRow() === 2 &&
    range.getColumn() === 1 &&
    range.getNumColumns() === getLastHeaderColumn_(sheet);
}

function rangeMatchesLegacyMainTable_(range) {
  return range.getRow() === 2 &&
    range.getColumn() === 1 &&
    (range.getNumColumns() === CONFIG.HEADERS.length || range.getNumColumns() === CONFIG.HEADERS.length + 1);
}

function isLegacyOpConditionalRule_(rule, columns) {
  const condition = rule.getBooleanCondition();

  if (!condition || condition.getCriteriaType() !== SpreadsheetApp.BooleanCriteria.TEXT_EQUAL_TO) {
    return false;
  }

  const values = condition.getCriteriaValues();
  const ranges = rule.getRanges();

  return values &&
    values.length > 0 &&
    String(values[0]) === 'OP' &&
    ranges.length > 0 &&
    ranges.every(range => {
      return rangeMatchesSingleColumn_(range, columns.TAG) ||
        rangeMatchesSingleColumn_(range, getDefaultColumnForField_('TAG'));
    });
}

function isLegacyPlanConditionalRule_(rule, sheet, columns) {
  const formula = getConditionalRuleFormula_(rule);
  const planReferences = [
    '$' + columnToLetter_(columns.PLAN) + '2',
    '$' + columnToLetter_(getDefaultColumnForField_('PLAN')) + '2'
  ];

  if (!formula || formula.indexOf('SEARCH(') === -1 || !planReferences.some(reference => formula.indexOf(reference) !== -1)) {
    return false;
  }

  const ranges = rule.getRanges();
  return ranges.length === 1 && (rangeMatchesMainTable_(ranges[0], sheet) || rangeMatchesLegacyMainTable_(ranges[0]));
}

function isManagedConditionalRule_(rule, sheet, columns) {
  return conditionalRuleHasSystemMarker_(rule) ||
    isLegacyOpConditionalRule_(rule, columns) ||
    isLegacyPlanConditionalRule_(rule, sheet, columns);
}

function getRetainedConditionalRules_(sheet, columns) {
  return sheet.getConditionalFormatRules()
    .filter(rule => !isManagedConditionalRule_(rule, sheet, columns));
}

/**
 * 將關鍵字轉成可套在第 2 列起始的條件格式公式片段。
 *
 * 條件格式公式會從範圍左上角的列開始相對套用，所以欄位固定寫成 $欄名2。
 */
function buildKeywordExpression_(column, keywords) {
  const activeKeywords = (keywords || [])
    .map(keyword => String(keyword).trim())
    .filter(Boolean);

  if (activeKeywords.length === 0) return '';

  const columnLetter = columnToLetter_(column);
  const formulaParts = activeKeywords.map(keyword => {
    return `ISNUMBER(SEARCH(${toSheetFormulaString_(keyword)}, $${columnLetter}2))`;
  });

  return formulaParts.length === 1
    ? formulaParts[0]
    : `OR(${formulaParts.join(',')})`;
}

function buildTableRangesExcludingColumns_(sheet, startRow, numRows, lastColumn, excludedColumns) {
  const excluded = {};
  const ranges = [];
  let column = 1;

  excludedColumns.forEach(excludedColumn => {
    excluded[excludedColumn] = true;
  });

  while (column <= lastColumn) {
    while (column <= lastColumn && excluded[column]) {
      column++;
    }

    const startColumn = column;

    while (column <= lastColumn && !excluded[column]) {
      column++;
    }

    if (startColumn <= lastColumn) {
      ranges.push(sheet.getRange(startRow, startColumn, numRows, column - startColumn));
    }
  }

  return ranges;
}

/**
 * 套用系統條件格式，並保留使用者自己新增的其他條件格式。
 *
 * Plan 命中 APPLY 時淡綠優先；命中問/補/通知時淡黃。
 * Plan 底色不套到 Tag 欄；時間欄命中 GA 時只讓時間格變紅，避免和整列底色互相覆蓋。
 */
function applyConditionalFormatting_(sheet, columns) {
  const dataRowCount = sheet.getMaxRows() - 1;
  if (dataRowCount <= 0) return;

  const retainedRules = getRetainedConditionalRules_(sheet, columns);
  const systemRules = [];
  const lastColumn = getTableLastColumn_(sheet, columns);
  const tagColLetter = columnToLetter_(columns.TAG);
  const tagRange = sheet.getRange(2, columns.TAG, dataRowCount, 1);
  const timeRange = sheet.getRange(2, columns.TIME, dataRowCount, 1);
  const applyExpression = buildKeywordExpression_(columns.PLAN, CONFIG.PLAN_GREEN_KEYWORDS);
  const yellowExpression = buildKeywordExpression_(columns.PLAN, CONFIG.PLAN_YELLOW_KEYWORDS);
  const timeGaExpression = buildKeywordExpression_(columns.TIME, CONFIG.TIME_GA_KEYWORDS);

  systemRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($${tagColLetter}2="OP", N(${toSheetFormulaString_(CONDITIONAL_FORMAT_MARKERS.OP_RED)})=0)`)
    .setBackground('#F4CCCC')
    .setFontColor('#CC0000')
    .setRanges([tagRange])
    .build()
  );

  if (timeGaExpression) {
    systemRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND(${timeGaExpression}, N(${toSheetFormulaString_(CONDITIONAL_FORMAT_MARKERS.TIME_GA)})=0)`)
      .setBackground('#F4CCCC')
      .setRanges([timeRange])
      .build()
    );
  }

  const planColorRanges = buildTableRangesExcludingColumns_(sheet, 2, dataRowCount, lastColumn, [
    columns.TAG,
    columns.TIME
  ]);

  if (applyExpression) {
    const greenFormula = `=AND(${applyExpression}, N(${toSheetFormulaString_(CONDITIONAL_FORMAT_MARKERS.PLAN_GREEN)})=0)`;
    const greenTimeFormula = `=AND(${applyExpression}, NOT(${timeGaExpression || 'FALSE'}), N(${toSheetFormulaString_(CONDITIONAL_FORMAT_MARKERS.PLAN_GREEN_TIME)})=0)`;

    systemRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(greenFormula)
      .setBackground('#D9EAD3')
      .setRanges(planColorRanges)
      .build()
    );

    systemRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(greenTimeFormula)
      .setBackground('#D9EAD3')
      .setRanges([timeRange])
      .build()
    );
  }

  if (yellowExpression) {
    const notApplyExpression = applyExpression ? `NOT(${applyExpression})` : 'TRUE';
    const yellowFormula = `=AND(${notApplyExpression}, ${yellowExpression}, N(${toSheetFormulaString_(CONDITIONAL_FORMAT_MARKERS.PLAN_YELLOW)})=0)`;
    const yellowTimeFormula = `=AND(${notApplyExpression}, ${yellowExpression}, NOT(${timeGaExpression || 'FALSE'}), N(${toSheetFormulaString_(CONDITIONAL_FORMAT_MARKERS.PLAN_YELLOW_TIME)})=0)`;

    systemRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(yellowFormula)
      .setBackground('#FFF2CC')
      .setRanges(planColorRanges)
      .build()
    );

    systemRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(yellowTimeFormula)
      .setBackground('#FFF2CC')
      .setRanges([timeRange])
      .build()
    );
  }

  sheet.setConditionalFormatRules(retainedRules.concat(systemRules));
}

/**
 * 將既有時間欄資料正規化為 HH:mm 文字。
 *
 * 支援使用者常見輸入：1330、830、13:30、Sheets 時間序列值與 Date 物件。
 * 無法解析的值不覆蓋，改在時間欄留下系統 note。
 */
function normalizeExistingTimes_(sheet, columns) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { normalizedCount: 0, errorCount: 0 };
  }

  const timeRange = sheet.getRange(2, columns.TIME, lastRow - 1, 1);
  const values = timeRange.getValues();
  let normalizedCount = 0;
  let errorCount = 0;

  values.forEach((rowValues, index) => {
    const row = index + 2;
    const value = rowValues[0];
    const timeCell = sheet.getRange(row, columns.TIME);

    if (value === null || value === undefined || value === '') {
      clearTimeErrorNote_(timeCell);
      return;
    }

    const timeInfo = resolveCalendarTime_(value);

    if (timeInfo.errorMessage) {
      setTimeErrorNote_(timeCell, timeInfo.errorMessage);
      errorCount++;
      return;
    }

    clearTimeErrorNote_(timeCell);

    if (timeInfo.parsedTime && !(typeof value === 'string' && value.trim() === timeInfo.parsedTime.text)) {
      timeCell.setValue(timeInfo.parsedTime.text);
      normalizedCount++;
    }
  });

  return { normalizedCount, errorCount };
}

/**
 * 初始化 OP sheet 的表格結構與格式。
 *
 * showAlert=false 是給 setup() 使用，避免一鍵安裝過程跳出多個 alert。
 * 此函式不會安裝 trigger，也不會批次同步既有 Calendar event。
 */
function initializeSheet_(showAlert) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_OP);

  if (!sheet) {
    return {
      ok: false,
      message: `找不到「${CONFIG.SHEET_OP}」工作表，初始化未執行。`
    };
  }

  ensureMinimumSheetSize_(sheet);
  const headerResult = ensureHeaders_(sheet);
  const legacyColumnResult = migrateLegacyEventIdColumns_(sheet);
  const columns = getRequiredSheetColumns_(sheet);
  applyDataFormats_(sheet, columns);
  applyTableAlignment_(sheet, columns);
  const timeResult = normalizeExistingTimes_(sheet, columns);
  applyConditionalFormatting_(sheet, columns);
  hideSystemColumns_(sheet, columns);
  const protectionResult = protectSystemColumns_(sheet, columns);

  const addedHeaderMessage = headerResult.addedHeaders.length > 0
    ? `\n已補齊欄位：${headerResult.addedHeaders.map(header => `「${header}」`).join('、')}。`
    : '';
  const duplicateHeaderMessage = headerResult.duplicateMessages.length > 0
    ? `\n\n以下必要欄位名稱重複，系統會使用最左側欄位：\n${headerResult.duplicateMessages.join('\n')}`
    : '';
  const legacyColumnMessage = legacyColumnResult.message
    ? '\n' + legacyColumnResult.message
    : '';
  const protectionMessage = protectionResult.message
    ? '\n' + protectionResult.message
    : '';

  if (showAlert) {
    ui.alert(
      '初始化完成',
      `已完成欄位確認、欄位格式、靠左靠上對齊、Tag 下拉選單、條件格式與既有時間正規化。\n` +
        `時間欄位已轉換 ${timeResult.normalizedCount} 格，格式錯誤 ${timeResult.errorCount} 格。` +
        addedHeaderMessage +
        legacyColumnMessage +
        protectionMessage +
        duplicateHeaderMessage,
      ui.ButtonSet.OK
    );
  }

  return {
    ok: true,
    timeResult,
    mismatchMessage: addedHeaderMessage + legacyColumnMessage + protectionMessage + duplicateHeaderMessage
  };
}

function initializeSheet() {
  initializeSheet_(true);
}

function runMigrations_(showAlert) {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_OP);

  if (!sheet) {
    return {
      ok: false,
      message: `找不到「${CONFIG.SHEET_OP}」工作表，資料遷移未執行。`
    };
  }

  ensureMinimumSheetSize_(sheet);
  const headerResult = ensureHeaders_(sheet);
  const legacyColumnResult = migrateLegacyEventIdColumns_(sheet);
  const columns = getRequiredSheetColumns_(sheet);
  applyDataFormats_(sheet, columns);
  hideSystemColumns_(sheet, columns);
  const protectionResult = protectSystemColumns_(sheet, columns);
  const messageParts = [];

  if (headerResult.addedHeaders.length > 0) {
    messageParts.push(`已補齊欄位：${headerResult.addedHeaders.map(header => `「${header}」`).join('、')}。`);
  }

  if (legacyColumnResult.message) {
    messageParts.push(legacyColumnResult.message);
  }

  if (protectionResult.message) {
    messageParts.push(protectionResult.message);
  }

  if (headerResult.duplicateMessages.length > 0) {
    messageParts.push(`以下必要欄位名稱重複，系統會使用最左側欄位：\n${headerResult.duplicateMessages.join('\n')}`);
  }

  const message = messageParts.length > 0
    ? messageParts.join('\n')
    : '沒有需要遷移的資料；系統欄位已確認、隱藏並嘗試保護。';

  if (showAlert) {
    ui.alert('資料遷移完成', message, ui.ButtonSet.OK);
  }

  return {
    ok: true,
    message,
    headerResult,
    legacyColumnResult,
    protectionResult
  };
}

function runMigrations() {
  const result = runMigrations_(true);
  if (!result.ok) {
    SpreadsheetApp.getUi().alert(result.message);
  }
  return result;
}

/**
 * 舊函式相容：請改用 initializeSheet。
 */
function initializeConditionalFormatting() {
  initializeSheet();
}

/**
 * 舊函式相容：請改用 initializeSheet。
 */
function initializeDataFormats() {
  initializeSheet();
}

/**
 * 批次同步日曆事件 (輔助工具，處理大量貼上時使用)
 *
 * setup() 會呼叫這個流程，讓更新程式碼後既有 Calendar event 也能套用新邏輯。
 * 也可以從選單或 Apps Script 手動執行 batchSyncAllEvents()。
 * 統計只計算真正新增、更新或刪除的列，缺病歷號等略過情境會另外計數。
 */
function buildBatchSyncSummaryMessage_(syncResult) {
  const statusCounts = syncResult.statusCounts || {};

  return `批次同步新增、更新，或刪除 ${syncResult.processedCount} 列；` +
    `缺病歷號略過 ${statusCounts.skipped_missing_chart_no || 0} 列；` +
    `時間錯誤 ${statusCounts.skipped_invalid_time || 0} 列。`;
}

function batchSyncAllEvents_(showAlert) {
  const sheet = getOpSheetOrThrow_();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    const emptyResult = {
      ok: true,
      processedCount: 0,
      statusCounts: {}
    };

    if (showAlert) {
      SpreadsheetApp.getUi().alert(buildBatchSyncSummaryMessage_(emptyResult));
    }

    return emptyResult;
  }

  if (!isCalendarIdConfigured_()) {
    const noCalendarIdResult = {
      ok: false,
      message: 'Calendar ID 尚未設定，批次同步未執行。',
      processedCount: 0,
      statusCounts: {}
    };

    if (showAlert) {
      SpreadsheetApp.getUi().alert(noCalendarIdResult.message);
    }

    return noCalendarIdResult;
  }

  const cols = getRequiredSheetColumns_(sheet);
  let processedCount = 0;
  const statusCounts = {};

  const executed = withCalendarSyncLock_(() => {
    const data = sheet.getRange(2, 1, lastRow - 1, getTableLastColumn_(sheet, cols)).getValues();

    for (let i = 0; i < data.length; i++) {
      const row = i + 2;
      const dateVal = getRowFieldValue_(data[i], cols, 'DATE');
      const eventId = getRowFieldValue_(data[i], cols, 'EVENT_ID');

      // 有日期：新增或更新
      // 無日期但有 eventId：刪除既有事件
      if (dateVal || eventId) {
        const status = syncToCalendar(sheet, row, cols);
        statusCounts[status] = (statusCounts[status] || 0) + 1;

        if (['created', 'updated', 'deleted'].indexOf(status) !== -1) {
          processedCount++;
        }
      }
    }
  });

  if (!executed) {
    const lockedResult = {
      ok: false,
      message: '目前有另一個日曆同步作業正在執行，本次批次同步未執行。',
      processedCount: 0,
      statusCounts
    };

    if (showAlert) {
      SpreadsheetApp.getUi().alert(lockedResult.message);
    }

    return lockedResult;
  }

  const result = {
    ok: true,
    processedCount,
    statusCounts
  };

  if (showAlert) {
    SpreadsheetApp.getUi().alert('批次處理完成。' + buildBatchSyncSummaryMessage_(result));
  }

  return result;
}

function batchSyncAllEvents() {
  return runMenuAction_('批次同步日曆事件', () => batchSyncAllEvents_(true));
}

function isCalendarSyncTokenInvalidError_(err) {
  const message = String((err && err.message) || err || '');
  return message.indexOf('410') !== -1 ||
    message.toLowerCase().indexOf('sync token') !== -1 ||
    message.toLowerCase().indexOf('gone') !== -1;
}

function buildCalendarEventRowIndex_(sheet) {
  const index = {
    byEventId: {},
    duplicateEventIds: {},
    columns: getRequiredSheetColumns_(sheet)
  };
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return index;

  const values = sheet.getRange(2, index.columns.EVENT_ID, lastRow - 1, 1).getValues();

  values.forEach((rowValues, indexOffset) => {
    const row = indexOffset + 2;
    const eventId = toCellText_(rowValues[0]);

    if (!eventId) return;

    if (!index.byEventId[eventId]) {
      index.byEventId[eventId] = row;
      return;
    }

    if (!index.duplicateEventIds[eventId]) {
      index.duplicateEventIds[eventId] = [index.byEventId[eventId]];
    }
    index.duplicateEventIds[eventId].push(row);
  });

  return index;
}

function updateCalendarEventRowIndex_(rowIndex, row, event) {
  const eventId = toCellText_(event && event.id);

  if (eventId) rowIndex.byEventId[eventId] = row;
}

function findSheetRowForCalendarApiEvent_(rowIndex, event) {
  const eventId = toCellText_(event && event.id);

  if (eventId && rowIndex.byEventId[eventId]) {
    return rowIndex.byEventId[eventId];
  }

  return 0;
}

function getDuplicateRowsForCalendarApiEvent_(rowIndex, event) {
  const eventId = toCellText_(event && event.id);
  if (!eventId || !rowIndex.duplicateEventIds) return [];

  return rowIndex.duplicateEventIds[eventId] || [];
}

function markDuplicateCalendarEventIdRows_(sheet, rows, columns, eventId) {
  const rowList = rows.join(', ');
  const note = `${CALENDAR_SYNC_NOTE_PREFIX}CalendarEventId 重複：${eventId} 同時出現在第 ${rowList} 列。請保留正確列的 ID，清空其他列後再同步。`;

  rows.forEach(row => {
    const cell = sheet.getRange(row, columns.EVENT_ID);
    cell.setNote(note);
  });
}

function parseCalendarDateOnly_(dateText) {
  const match = String(dateText || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return null;

  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
}

function getCalendarEventDateTimeValues_(event) {
  if (!event || !event.start) {
    return {
      ok: false,
      message: 'Calendar event 缺少 start 欄位。'
    };
  }

  if (event.start.date) {
    const date = parseCalendarDateOnly_(event.start.date);
    return date
      ? {
          ok: true,
          date,
          timeText: '',
          isAllDay: true
        }
      : {
          ok: false,
          message: `全天事件日期格式錯誤：${event.start.date}`
        };
  }

  if (event.start.dateTime) {
    const start = new Date(event.start.dateTime);
    if (isNaN(start.getTime())) {
      return {
        ok: false,
        message: `事件開始時間格式錯誤：${event.start.dateTime}`
      };
    }

    const timeZone = event.start.timeZone || Session.getScriptTimeZone();
    const dateText = Utilities.formatDate(start, timeZone, 'yyyy-MM-dd');
    const date = parseCalendarDateOnly_(dateText);

    return {
      ok: true,
      date,
      timeText: Utilities.formatDate(start, timeZone, 'HH:mm'),
      isAllDay: false
    };
  }

  return {
    ok: false,
    message: 'Calendar event start 欄位沒有 date 或 dateTime。'
  };
}

function isRecurringCalendarApiEvent_(event) {
  return Boolean(event && (event.recurringEventId || event.recurrence));
}

function tagFromCalendarApiEvent_(event) {
  const colorId = toCellText_(event && event.colorId);
  const opColor = toCellText_(CONFIG.COLORS.OP);

  return colorId && colorId === opColor ? 'OP' : '';
}

function addCalendarDeletedMemoNote_(range) {
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
  const message = `${CALENDAR_SYNC_NOTE_PREFIX}日曆已刪除：${timestamp}`;
  const existingNote = range.getNote();

  range.setNote(existingNote ? `${existingNote}\n${message}` : message);
}

function markSheetRowCalendarDeleted_(sheet, row, columns) {
  const cols = columns || getRequiredSheetColumns_(sheet);
  sheet.getRange(row, cols.DATE).clearContent();
  sheet.getRange(row, cols.TIME).clearContent();
  clearSheetOriginatedCalendarState_(sheet, row, cols);
  addCalendarDeletedMemoNote_(sheet.getRange(row, cols.MEMO));
}

function getSheetTimeValueFromCalendarEvent_(dateTimeInfo, descriptionInfo) {
  if (dateTimeInfo && dateTimeInfo.isAllDay && descriptionInfo && descriptionInfo.hasTimeNote) {
    return descriptionInfo.timeNote;
  }

  return dateTimeInfo ? dateTimeInfo.timeText : '';
}

function updateSheetRowFromCalendarApiEvent_(sheet, row, event, dateTimeInfo, columns) {
  const cols = columns || getRequiredSheetColumns_(sheet);
  const rowRange = sheet.getRange(row, 1, 1, getTableLastColumn_(sheet, cols));
  const values = rowRange.getValues()[0];
  const titleInfo = parseCalendarTitle_(event.summary || '');
  const descriptionInfo = parseCalendarDescription_(event.description || '');

  if (titleInfo.ok) {
    setRowFieldValue_(values, cols, 'CHART_NO', titleInfo.chartNo);
    setRowFieldValue_(values, cols, 'NAME', titleInfo.patientName);
    setRowFieldValue_(values, cols, 'COND', titleInfo.condition);
  } else {
    console.warn(`第 ${row} 列 Calendar 標題解析失敗，保留原病人欄位。`);
  }

  if (descriptionInfo.hasTel) {
    setRowFieldValue_(values, cols, 'TEL', descriptionInfo.tel);
  }

  if (descriptionInfo.hasPlan) {
    setRowFieldValue_(values, cols, 'PLAN', descriptionInfo.plan);
  }

  setRowFieldValue_(values, cols, 'DATE', dateTimeInfo.date);
  setRowFieldValue_(values, cols, 'TIME', getSheetTimeValueFromCalendarEvent_(dateTimeInfo, descriptionInfo));
  setRowFieldValue_(values, cols, 'EVENT_ID', toCellText_(event.id) || getRowFieldValue_(values, cols, 'EVENT_ID'));

  rowRange.setValues([values]);
  applyRowDataFormats_(sheet, row, cols);
}

function appendSheetRowFromCalendarApiEvent_(sheet, event, dateTimeInfo, columns) {
  const cols = columns || getRequiredSheetColumns_(sheet);
  const titleInfo = parseCalendarTitle_(event.summary || '');

  if (!titleInfo.ok) {
    console.warn('新增 Calendar event 未寫入 Sheet：標題格式不符合系統規則。');
    return 0;
  }

  const descriptionInfo = parseCalendarDescription_(event.description || '');
  const rowValues = new Array(getTableLastColumn_(sheet, cols)).fill('');
  setRowFieldValue_(rowValues, cols, 'CHART_NO', titleInfo.chartNo);
  setRowFieldValue_(rowValues, cols, 'NAME', titleInfo.patientName);
  setRowFieldValue_(rowValues, cols, 'TEL', descriptionInfo.hasTel ? descriptionInfo.tel : '');
  setRowFieldValue_(rowValues, cols, 'TAG', tagFromCalendarApiEvent_(event));
  setRowFieldValue_(rowValues, cols, 'COND', titleInfo.condition);
  setRowFieldValue_(rowValues, cols, 'DATE', dateTimeInfo.date);
  setRowFieldValue_(rowValues, cols, 'TIME', getSheetTimeValueFromCalendarEvent_(dateTimeInfo, descriptionInfo));
  setRowFieldValue_(rowValues, cols, 'PLAN', descriptionInfo.hasPlan ? descriptionInfo.plan : '');
  setRowFieldValue_(rowValues, cols, 'EVENT_ID', toCellText_(event.id));

  const row = sheet.getLastRow() + 1;
  if (row > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  }

  sheet.getRange(row, 1, 1, rowValues.length).setValues([rowValues]);
  applyRowDataFormats_(sheet, row, cols);

  return row;
}

function processCalendarApiEventChange_(sheet, rowIndex, event, calendarId) {
  if (!event || !event.id) return 'skipped_missing_event_id';

  const row = findSheetRowForCalendarApiEvent_(rowIndex, event);
  const duplicateRows = getDuplicateRowsForCalendarApiEvent_(rowIndex, event);
  if (duplicateRows.length > 0) {
    markDuplicateCalendarEventIdRows_(sheet, duplicateRows, rowIndex.columns, event.id);
    console.warn(`Calendar event ${event.id} 對應到重複 CalendarEventId 列：${duplicateRows.join(', ')}`);
    return 'skipped_duplicate_event_id';
  }

  if (row && isSheetOriginatedCalendarEcho_(sheet, row, rowIndex.columns, event)) {
    return 'skipped_sheet_echo';
  }

  if (event.status === 'cancelled') {
    if (!row) return 'skipped_deleted_unmatched';
    markSheetRowCalendarDeleted_(sheet, row, rowIndex.columns);
    return 'deleted_marked';
  }

  if (isRecurringCalendarApiEvent_(event)) {
    return 'skipped_recurring';
  }

  const dateTimeInfo = getCalendarEventDateTimeValues_(event);
  if (!dateTimeInfo.ok) {
    console.warn(`Calendar event ${event.id} 時間解析失敗：${dateTimeInfo.message}`);
    return 'skipped_invalid_time';
  }

  if (row) {
    updateSheetRowFromCalendarApiEvent_(sheet, row, event, dateTimeInfo, rowIndex.columns);
    updateCalendarEventRowIndex_(rowIndex, row, event);
    return 'updated';
  }

  const newRow = appendSheetRowFromCalendarApiEvent_(sheet, event, dateTimeInfo, rowIndex.columns);
  if (!newRow) return 'skipped_invalid_title';

  updateCalendarEventRowIndex_(rowIndex, newRow, event);
  return 'created';
}

function syncCalendarChangesToSheet_(calendarId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_OP);
  if (!sheet) {
    return {
      ok: false,
      message: `找不到「${CONFIG.SHEET_OP}」工作表，Calendar 反向同步未執行。`
    };
  }

  if (!isCalendarIdConfigured_(calendarId)) {
    return {
      ok: false,
      message: 'Calendar ID 尚未設定，Calendar 反向同步未執行。'
    };
  }

  if (!isCalendarAdvancedServiceAvailable_()) {
    return {
      ok: false,
      message: '尚未啟用 Calendar Advanced Service，Calendar 反向同步未執行。'
    };
  }

  const syncToken = getCalendarSyncToken_(calendarId);
  if (!syncToken) {
    const baselineResult = initializeCalendarSyncToken_(calendarId);
    return {
      ok: baselineResult.ok,
      message: baselineResult.ok
        ? 'Calendar syncToken 不存在，已建立 baseline；本次不回寫事件。'
        : baselineResult.message,
      processedCount: 0,
      statusCounts: {}
    };
  }

  const rowIndex = buildCalendarEventRowIndex_(sheet);
  const statusCounts = {};
  let processedCount = 0;
  let pageToken = '';
  let response = null;

  try {
    do {
      const options = {
        maxResults: 2500,
        syncToken
      };

      if (pageToken) {
        options.pageToken = pageToken;
      }

      response = Calendar.Events.list(calendarId, options);

      (response.items || []).forEach(event => {
        const status = processCalendarApiEventChange_(sheet, rowIndex, event, calendarId);
        statusCounts[status] = (statusCounts[status] || 0) + 1;

        if (['updated', 'created', 'deleted_marked'].indexOf(status) !== -1) {
          processedCount++;
        }
      });

      pageToken = response.nextPageToken || '';
    } while (pageToken);
  } catch (err) {
    if (isCalendarSyncTokenInvalidError_(err)) {
      const baselineResult = initializeCalendarSyncToken_(calendarId);
      return {
        ok: false,
        message: baselineResult.ok
          ? 'Calendar syncToken 已失效，已重新建立 baseline；本次不回寫事件。'
          : `Calendar syncToken 已失效，但 baseline 重建失敗：${baselineResult.message}`,
        processedCount,
        statusCounts
      };
    }

    return {
      ok: false,
      message: `Calendar 反向同步失敗：${err.message || err}`,
      processedCount,
      statusCounts
    };
  }

  if (response && response.nextSyncToken) {
    setCalendarSyncToken_(calendarId, response.nextSyncToken);
  }

  return {
    ok: true,
    message: `Calendar 反向同步完成，回寫 ${processedCount} 筆變更。`,
    processedCount,
    statusCounts
  };
}

function syncCalendarChangesToSheet() {
  const result = syncCalendarChangesToSheet_(getConfiguredCalendarId_());
  SpreadsheetApp.getUi().alert(result.message);
}

function processCalendarChange(e) {
  const calendarId = e && e.calendarId ? e.calendarId : getConfiguredCalendarId_();
  let syncResult = null;

  const executed = withCalendarSyncLock_(() => {
    syncResult = syncCalendarChangesToSheet_(calendarId);
    if (!syncResult.ok) {
      console.warn(syncResult.message);
    }
  });

  if (executed && syncResult && syncResult.ok && syncResult.processedCount > 0) {
    scheduleUpcomingWeekExport_();
  }
}

/**
 * 核心觸發器：編輯時更新日曆
 *
 * 注意：
 * - 這個函式需要由 installable onEdit trigger 呼叫，建議透過 setup() 安裝。
 * - 不要任意改名；已安裝 trigger 會記住 handler function name。
 */
function processRowChange(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.SHEET_OP) return;
  let shouldScheduleExport = false;

  withCalendarSyncLock_(() => {
    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();
    const startCol = e.range.getColumn();
    const numCols = e.range.getNumColumns();
    const endCol = startCol + numCols - 1;
    let cols;

    try {
      cols = getRequiredSheetColumns_(sheet);
    } catch (err) {
      console.warn(err.message || err);
      return;
    }

    // 檢查編輯範圍是否涵蓋我們關注的欄位
    const watchCols = Object.values(cols);
    const isWatchColEdited = watchCols.some(col => col >= startCol && col <= endCol);

    if (!isWatchColEdited) return;

    // 逐列處理 (支援同時貼上多列的情況)
    for (let i = 0; i < numRows; i++) {
      const currentRow = startRow + i;
      if (currentRow < 2) continue;

      shouldScheduleExport = true;

      // 處理手打時間轉換
      // 僅在「單一時間儲存格」編輯時自動改寫顯示值，避免整批貼上時改壞範圍。
      if (numRows === 1 && numCols === 1 && startCol === cols.TIME) {
        const timeInfo = resolveCalendarTime_(e.range.getValue());

        if (timeInfo.errorMessage) {
          setTimeErrorNote_(e.range, timeInfo.errorMessage);
          console.warn(`第 ${currentRow} 列：${timeInfo.errorMessage}`);
          return;
        }

        clearTimeErrorNote_(e.range);

        if (timeInfo.parsedTime) {
          e.range.setValue(timeInfo.parsedTime.text);
        }
      }

      syncToCalendar(sheet, currentRow, cols);
    }
  });

  if (shouldScheduleExport) {
    scheduleUpcomingWeekExport_();
  }
}

function isCalendarEventNotFoundError_(err) {
  const message = String((err && err.message) || err || '').toLowerCase();
  return message.indexOf('404') !== -1 || message.indexOf('not found') !== -1;
}

function isLegacyCalendarEventId_(eventId) {
  return String(eventId || '').indexOf('@') !== -1;
}

function formatCalendarApiDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatCalendarApiDateTime_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function addDays_(date, days) {
  const nextDate = new Date(date.getTime());
  nextDate.setDate(nextDate.getDate() + days);
  return nextDate;
}

function buildCalendarApiEventResource_(title, description, dateObj, timeInfo, tagText) {
  const resource = {
    summary: title,
    description,
    colorId: CONFIG.COLORS.DEFAULT
  };

  if (timeInfo.parsedTime === null) {
    resource.colorId = CONFIG.COLORS.NO_TIME;
    resource.start = {
      date: formatCalendarApiDate_(dateObj)
    };
    resource.end = {
      date: formatCalendarApiDate_(addDays_(dateObj, 1))
    };
    return resource;
  }

  const startTime = new Date(dateObj.getTime());
  startTime.setHours(timeInfo.parsedTime.hour, timeInfo.parsedTime.minute, 0, 0);
  const endTime = new Date(startTime.getTime() + 60 * 60 * 1000);

  if (tagText === 'OP') {
    resource.colorId = CONFIG.COLORS.OP;
  }

  resource.start = {
    dateTime: formatCalendarApiDateTime_(startTime),
    timeZone: Session.getScriptTimeZone()
  };
  resource.end = {
    dateTime: formatCalendarApiDateTime_(endTime),
    timeZone: Session.getScriptTimeZone()
  };

  return resource;
}

function buildCalendarApiUpdateResource_(existingEvent, eventResource) {
  const updateResource = Object.assign({}, existingEvent || {});

  updateResource.summary = eventResource.summary;
  updateResource.description = eventResource.description;
  updateResource.colorId = eventResource.colorId;
  updateResource.start = Object.assign({}, eventResource.start);
  updateResource.end = Object.assign({}, eventResource.end);

  return updateResource;
}

function updateCalendarApiEvent_(calendarId, eventId, eventResource) {
  const existingEvent = Calendar.Events.get(calendarId, eventId);
  const updateResource = buildCalendarApiUpdateResource_(existingEvent, eventResource);

  return Calendar.Events.update(updateResource, calendarId, eventId);
}

function removeCalendarApiEvent_(calendarId, eventId) {
  if (!eventId || !isCalendarAdvancedServiceAvailable_()) {
    return {
      ok: false,
      message: 'Calendar Advanced Service 尚未啟用或 event ID 空白。'
    };
  }

  try {
    Calendar.Events.remove(calendarId, eventId);
    return {
      ok: true,
      status: 'deleted'
    };
  } catch (err) {
    if (isCalendarEventNotFoundError_(err)) {
      return {
        ok: true,
        status: 'already_missing'
      };
    }

    return {
      ok: false,
      message: err.message || String(err)
    };
  }
}

function findCalendarApiEventsByICalUID_(calendarId, iCalUID) {
  const events = [];
  let pageToken = '';

  do {
    const options = {
      iCalUID,
      maxResults: 2500,
      showDeleted: false
    };

    if (pageToken) {
      options.pageToken = pageToken;
    }

    const response = Calendar.Events.list(calendarId, options);
    (response.items || []).forEach(event => events.push(event));
    pageToken = response.nextPageToken || '';
  } while (pageToken);

  return events;
}

function deleteCalendarEventByPossibleId_(calendarId, rawEventId) {
  const eventId = toCellText_(rawEventId);
  if (!eventId) {
    return {
      ok: true,
      status: 'blank'
    };
  }

  const directDeleteResult = removeCalendarApiEvent_(calendarId, eventId);
  if (directDeleteResult.ok && directDeleteResult.status === 'deleted') return directDeleteResult;

  try {
    const matchingEvents = findCalendarApiEventsByICalUID_(calendarId, eventId);
    if (matchingEvents.length === 0) {
      return {
        ok: true,
        status: 'already_missing'
      };
    }

    let deletedCount = 0;
    for (let i = 0; i < matchingEvents.length; i++) {
      const matchedEventId = toCellText_(matchingEvents[i].id);
      if (!matchedEventId) continue;

      const deleteResult = removeCalendarApiEvent_(calendarId, matchedEventId);
      if (!deleteResult.ok) {
        return deleteResult;
      }
      deletedCount++;
    }

    return {
      ok: true,
      status: deletedCount > 0 ? 'deleted' : 'already_missing',
      deletedCount
    };
  } catch (err) {
    if (directDeleteResult.ok) return directDeleteResult;

    return {
      ok: false,
      message: err.message || String(err)
    };
  }
}

function clearSystemCalendarSyncNote_(range) {
  const note = range.getNote();
  if (note && String(note).startsWith(CALENDAR_SYNC_NOTE_PREFIX)) {
    range.clearNote();
  }
}

function countNonEmptyEventIds_(values) {
  return values.reduce((count, rowValues) => count + (toCellText_(rowValues[0]) ? 1 : 0), 0);
}

function clearCalendarEventsInEventIdRange_(idRange, calendarId, deleteEventFn, sheetWriteUpdatedRange) {
  const values = idRange.getValues();
  const result = {
    scannedCount: 0,
    clearedCount: 0,
    deletedCount: 0,
    alreadyMissingCount: 0,
    failedCount: 0
  };
  const deleteFn = deleteEventFn || deleteCalendarEventByPossibleId_;

  values.forEach((rowValues, index) => {
    const eventId = toCellText_(rowValues[0]);
    if (!eventId) return;

    result.scannedCount++;
    const cell = idRange.getCell(index + 1, 1);
    const deleteResult = deleteFn(calendarId, eventId);

    if (deleteResult.ok) {
      cell.clearContent();
      clearSystemCalendarSyncNote_(cell);

      if (sheetWriteUpdatedRange) {
        const sheetWriteUpdatedCell = sheetWriteUpdatedRange.getCell(index + 1, 1);
        sheetWriteUpdatedCell.clearContent();
        clearSystemCalendarSyncNote_(sheetWriteUpdatedCell);
      }

      result.clearedCount++;

      if (deleteResult.status === 'deleted') {
        result.deletedCount++;
      } else if (deleteResult.status === 'already_missing') {
        result.alreadyMissingCount++;
      }
      return;
    }

    result.failedCount++;
    cell.setNote(`${CALENDAR_SYNC_NOTE_PREFIX}刪除日曆事件失敗：${deleteResult.message}`);
  });

  return result;
}

function clearCalendarEventsFromSpecifiedColumn() {
  const ui = SpreadsheetApp.getUi();
  const calendarId = getConfiguredCalendarId_();

  if (!isCalendarIdConfigured_(calendarId)) {
    ui.alert('Calendar ID 尚未設定，未清除日曆事件。');
    return;
  }

  if (!isCalendarAdvancedServiceAvailable_()) {
    ui.alert('尚未啟用 Calendar Advanced Service，未清除日曆事件。');
    return;
  }

  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_OP);
  if (!targetSheet) {
    ui.alert(`找不到「${CONFIG.SHEET_OP}」工作表，未清除日曆事件。`);
    return;
  }

  const column = getEventIdColumnForSheet_(targetSheet);
  if (!column) {
    ui.alert(`「${CONFIG.SHEET_OP}」工作表第 1 列找不到「${CONFIG.FIELD_HEADERS.EVENT_ID}」欄位，未清除日曆事件。`);
    return;
  }

  const lastRow = targetSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert(`「${CONFIG.SHEET_OP}」工作表沒有可處理的資料列。`);
    return;
  }

  const idRange = targetSheet.getRange(2, column, lastRow - 1, 1);
  const columnInfo = buildFieldColumnInfo_(targetSheet);
  const sheetWriteUpdatedColumn = columnInfo.canonicalColumns.SHEET_WRITE_UPDATED || 0;
  const sheetWriteUpdatedRange = sheetWriteUpdatedColumn
    ? targetSheet.getRange(2, sheetWriteUpdatedColumn, lastRow - 1, 1)
    : null;
  const nonEmptyCount = countNonEmptyEventIds_(idRange.getValues());
  if (nonEmptyCount === 0) {
    ui.alert(`「${CONFIG.SHEET_OP}」工作表的「${CONFIG.FIELD_HEADERS.EVENT_ID}」欄沒有可清除的事件 ID。`);
    return;
  }

  const clearStateMessage = sheetWriteUpdatedRange
    ? `此動作會刪除 Calendar 事件，並清空對應 ${CONFIG.FIELD_HEADERS.EVENT_ID} 與 ${CONFIG.FIELD_HEADERS.SHEET_WRITE_UPDATED} 儲存格。`
    : `此動作會刪除 Calendar 事件並清空對應 ${CONFIG.FIELD_HEADERS.EVENT_ID} 儲存格。`;
  const confirmResponse = ui.alert(
    '清除舊事件',
    `即將針對「${CONFIG.SHEET_OP}」工作表 ${columnToLetter_(column)} 欄（${CONFIG.FIELD_HEADERS.EVENT_ID}）清除 ${nonEmptyCount} 筆日曆事件。\n` +
      clearStateMessage,
    ui.ButtonSet.OK_CANCEL
  );
  if (confirmResponse !== ui.Button.OK) return;

  let result;
  const executed = withCalendarSyncLock_(() => {
    result = clearCalendarEventsInEventIdRange_(idRange, calendarId, null, sheetWriteUpdatedRange);
  });

  if (!executed) {
    ui.alert('目前有另一個日曆同步作業正在執行，本次清除未執行。');
    return;
  }

  ui.alert(
    '清除完成',
    `工作表「${CONFIG.SHEET_OP}」${columnToLetter_(column)} 欄；` +
      `掃描 ${result.scannedCount} 格；` +
      `刪除 ${result.deletedCount} 筆；` +
      `已不存在 ${result.alreadyMissingCount} 筆；` +
      `清空 ID ${result.clearedCount} 格；` +
      `失敗 ${result.failedCount} 格。`,
    ui.ButtonSet.OK
  );
}

/**
 * 依單列資料建立、更新或刪除 Calendar event。
 *
 * 同步規則：
 * - 沒有日期：若有 eventID，刪除既有事件並清空 eventID。
 * - 有日期但沒有病歷號：暫停同步，保留 eventID，並在病歷號欄加 note。
 * - 有日期且有病歷號：依 Calendar API event.id 更新既有事件，或建立新事件。
 *
 * 回傳狀態字串給 batchSyncAllEvents() 統計用。
 */
function syncToCalendar(sheet, row, columns) {
  const cols = columns || getRequiredSheetColumns_(sheet);
  const rowData = sheet.getRange(row, 1, 1, getTableLastColumn_(sheet, cols)).getValues()[0];
  const chartNo = getRowFieldValue_(rowData, cols, 'CHART_NO');
  const patientName = getRowFieldValue_(rowData, cols, 'NAME');
  const tel = getRowFieldValue_(rowData, cols, 'TEL');
  const tag = getRowFieldValue_(rowData, cols, 'TAG');
  const condition = getRowFieldValue_(rowData, cols, 'COND');
  const dateVal = getRowFieldValue_(rowData, cols, 'DATE');
  const timeVal = getRowFieldValue_(rowData, cols, 'TIME');
  const plan = getRowFieldValue_(rowData, cols, 'PLAN');
  const eventId = toCellText_(getRowFieldValue_(rowData, cols, 'EVENT_ID'));
  const chartNoText = toCellText_(chartNo);
  const patientNameText = toSingleLineText_(patientName);
  const telText = toCellText_(tel);
  const tagText = toCellText_(tag);
  const conditionText = toSingleLineText_(condition);
  const planText = toCellText_(plan);
  const chartNoCell = sheet.getRange(row, cols.CHART_NO);
  const eventIdCell = sheet.getRange(row, cols.EVENT_ID);
  const calendarId = getConfiguredCalendarId_();

  if (!isCalendarIdConfigured_(calendarId)) {
    setCalendarSyncNote_(eventIdCell, 'Calendar ID 尚未設定，請先用「設定 Calendar ID」儲存手術日曆 ID。');
    return 'skipped_no_calendar_id';
  }

  if (!isCalendarAdvancedServiceAvailable_()) {
    setCalendarSyncNote_(eventIdCell, '需啟用 Calendar Advanced Service 才能同步日曆。');
    return 'skipped_no_calendar_service';
  }

  if (!dateVal) {
    clearCalendarSyncNote_(chartNoCell);

    if (eventId) {
      const deleteResult = deleteCalendarEventByPossibleId_(calendarId, eventId);
      if (!deleteResult.ok) {
        setCalendarSyncNote_(eventIdCell, `刪除日曆事件失敗：${deleteResult.message}`);
        return 'skipped_delete_failed';
      }

      clearSheetOriginatedCalendarState_(sheet, row, cols);
      clearSystemCalendarSyncNote_(eventIdCell);
      return 'deleted';
    }

    return 'skipped_no_date';
  }

  if (!chartNoText) {
    setCalendarSyncNote_(chartNoCell, '需填病歷號才會同步日曆。');
    return 'skipped_missing_chart_no';
  }

  clearCalendarSyncNote_(chartNoCell);

  if (eventId && isLegacyCalendarEventId_(eventId)) {
    setCalendarSyncNote_(eventIdCell, '這是舊 iCalUID，請先用「清除舊事件」刪除並清空後再重建。');
    return 'skipped_legacy_event_id';
  }

  const dateObj = new Date(dateVal);
  if (isNaN(dateObj.getTime())) return 'skipped_invalid_date';

  const timeCell = sheet.getRange(row, cols.TIME);
  const timeInfo = resolveCalendarTime_(timeVal);

  if (timeInfo.errorMessage) {
    setTimeErrorNote_(timeCell, timeInfo.errorMessage);
    console.warn(`第 ${row} 列：${timeInfo.errorMessage}`);
    return 'skipped_invalid_time';
  }

  clearTimeErrorNote_(timeCell);

  const title = buildCalendarTitle_(chartNoText, patientNameText, conditionText);
  const description = buildCalendarDescription_(telText, planText, timeInfo.timeNote);
  const eventResource = buildCalendarApiEventResource_(title, description, dateObj, timeInfo, tagText);

  if (eventId) {
    try {
      const updatedEvent = updateCalendarApiEvent_(calendarId, eventId, eventResource);
      writeSheetOriginatedCalendarState_(sheet, row, cols, updatedEvent, eventId);
      clearSystemCalendarSyncNote_(eventIdCell);
      return 'updated';
    } catch (err) {
      if (!isCalendarEventNotFoundError_(err)) {
        setCalendarSyncNote_(eventIdCell, `更新日曆事件失敗：${err.message || err}`);
        return 'skipped_update_failed';
      }
    }
  }

  try {
    const createdEvent = Calendar.Events.insert(eventResource, calendarId);
    writeSheetOriginatedCalendarState_(sheet, row, cols, createdEvent, '');
    clearSystemCalendarSyncNote_(eventIdCell);
    return 'created';
  } catch (err) {
    setCalendarSyncNote_(eventIdCell, `建立日曆事件失敗：${err.message || err}`);
    return 'skipped_create_failed';
  }
}
