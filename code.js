/**
 * 手術排程系統 Google Apps Script
 *
 * 首次安裝流程：
 * 1. 將本檔貼到試算表綁定的 Apps Script 專案。
 * 2. 將 CONFIG.CALENDAR_ID 改成要同步的 Google Calendar ID。
 * 3. 手動執行 setup() 一次，完成授權、表格初始化與 onEdit trigger 安裝。
 *
 * 重要假設：
 * - 主要資料表名稱為 CONFIG.SHEET_OP，預設是「OP」。
 * - A:J 是系統管理欄位，欄位順序由 CONFIG.COLS 固定。
 * - 病歷號是同步 Calendar 的必要欄位；有日期但沒有病歷號時不會新增或更新事件。
 */
const CONFIG = {
  CALENDAR_ID: 'YOUR_CALENDAR_ID_HERE', // 請替換成專屬手術日曆的ID
  SHEET_OP: 'OP',
  SHEET_OUT: '輸出表單',
  COLS: {
    CHART_NO: 1, // A欄: 病歷號
    NAME: 2,     // B欄: 姓名
    TEL: 3,      // C欄: TEL
    TAG: 4,      // D欄: Tag
    COND: 5,     // E欄: Condition
    DATE: 6,     // F欄: 日期
    TIME: 7,     // G欄: 時間
    PLAN: 8,     // H欄: Plan
    MEMO: 9,     // I欄: 心得
    EVENT_ID: 10 // J欄: 隱藏的 Calendar API event ID
  },

  HEADERS: ['病歷號', '姓名', 'TEL', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', '日曆eventID'],
  TAG_OPTIONS: ['CATA', 'Eyelid', 'Retina', 'OP', 'FU', 'Suture IOL', 'Complication'],

  // 條件格式化關鍵字：
  // 只要 Plan 欄包含以下任一字彙，就會套用條件格式。
  CONDITIONAL_FORMAT_KEYWORDS: ['通知', 'APPLY'],

  COLORS: {
    NO_TIME: '5',
    OP: '11',
    DEFAULT: '9'
  }
};

const TIME_ERROR_NOTE_PREFIX = '系統時間檢查：';
const CALENDAR_SYNC_NOTE_PREFIX = '系統日曆同步：';
const CALENDAR_SYNC_TOKEN_PROPERTY_PREFIX = 'CALENDAR_SYNC_TOKEN_';
const CALENDAR_REVERSE_SYNC_HANDLER = 'processCalendarChange';
const CALENDAR_TITLE_SEPARATOR = '|';
const LEGACY_API_EVENT_ID_HEADER = '日曆apiEventID';

// 條件格式沒有 metadata 可標記來源，因此在 custom formula 中放入 N("...") marker。
// initializeSheet() 重新執行時會用這些 marker 辨識並替換系統規則，同時保留使用者自訂規則。
const CONDITIONAL_FORMAT_MARKERS = {
  OP_RED: 'SYSTEM_CF_OP_RED',
  PLAN_YELLOW_MAIN: 'SYSTEM_CF_PLAN_YELLOW_MAIN',
  PLAN_YELLOW_TAG: 'SYSTEM_CF_PLAN_YELLOW_TAG'
};

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
    typeof Calendar.Events.list === 'function';
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

function isCalendarIdConfigured_() {
  return CONFIG.CALENDAR_ID && CONFIG.CALENDAR_ID !== 'YOUR_CALENDAR_ID_HERE';
}

function initializeCalendarSyncToken_(calendarId) {
  if (!isCalendarIdConfigured_()) {
    return {
      ok: false,
      message: 'CONFIG.CALENDAR_ID 尚未設定，無法建立 Calendar 反向同步 baseline。'
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
  const result = initializeCalendarSyncToken_(CONFIG.CALENDAR_ID);
  SpreadsheetApp.getUi().alert(result.message);
}

function installCalendarChangeTrigger_(showAlert) {
  if (!isCalendarIdConfigured_()) {
    const result = {
      ok: false,
      message: 'CONFIG.CALENDAR_ID 尚未設定，未安裝 Calendar 反向同步觸發器。'
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
      .forUserCalendar(CONFIG.CALENDAR_ID)
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
  const tokenResult = initializeCalendarSyncToken_(CONFIG.CALENDAR_ID);
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
 * 一鍵安裝/初始化入口。
 *
 * 第一次安裝建議只手動執行這個函式：
 * - 初始化 OP sheet 標題、格式、資料驗證與條件格式
 * - 安裝 Calendar 同步用 installable onEdit trigger
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

  const syncResult = batchSyncAllEvents_(false);
  const syncMessage = syncResult.ok
    ? '\n' + buildBatchSyncSummaryMessage_(syncResult)
    : `\n批次同步未執行：${syncResult.message}`;
  installProcessRowChangeTrigger_(false);
  const reverseSyncResult = setupCalendarReverseSync_(false);
  const reverseSyncMessage = reverseSyncResult.ok
    ? `\n${reverseSyncResult.message}`
    : `\nCalendar 反向同步未啟用：${reverseSyncResult.message}`;
  onOpen();

  ui.alert(
    '一鍵設定完成',
    `已完成表格初始化、Sheet 自動同步觸發器安裝，並刷新「手術排程系統」選單。\n` +
      `時間欄位已轉換 ${initResult.timeResult.normalizedCount} 格，格式錯誤 ${initResult.timeResult.errorCount} 格。` +
      syncMessage +
      reverseSyncMessage +
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

function buildCalendarTitle_(chartNoText, patientNameText, conditionText) {
  return [chartNoText, patientNameText, conditionText]
    .map(value => toSingleLineText_(value))
    .filter(Boolean)
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

  if (pipeParts.length >= 2 && pipeParts[0] && pipeParts[1]) {
    return {
      ok: true,
      chartNo: pipeParts[0],
      patientName: pipeParts[1],
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
    tel: '',
    plan: ''
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

    if (/^Time note\s*[:：]/i.test(line)) {
      activeField = '';
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
  const advancedMenu = ui.createMenu('進階/維護工具')
    .addItem('初始化表格格式', 'initializeSheet')
    .addItem('安裝自動同步觸發器', 'installProcessRowChangeTrigger')
    .addItem('安裝日曆反向同步觸發器', 'setupCalendarReverseSync')
    .addItem('清除指定欄位舊日曆事件', 'clearCalendarEventsFromSpecifiedColumn');

  ui.createMenu('手術排程系統')
    .addItem('一鍵安裝/初始化', 'setup')
    .addItem('輸出指令日期資料', 'exportDateData')
    .addItem('複製一列 (清空時間)', 'duplicateRow')
    .addItem('歸人整合 (同病歷號合併)', 'mergePatientRecords')
    .addItem('依照日期與時間排序', 'sortSheetByDateTime')
    .addSeparator()
    .addSubMenu(advancedMenu)
    .addToUi();
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
function exportDateData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('輸出指定日期資料', '請輸入日期 (格式如: 2026/5/11):', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return;
  const targetDateStr = response.getResponseText().trim();
  const targetDateObj = new Date(targetDateStr);

  if (isNaN(targetDateObj.getTime())) {
    ui.alert('錯誤', '日期格式不正確，請重新執行。', ui.ButtonSet.OK);
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetOp = ss.getSheetByName(CONFIG.SHEET_OP);
  let sheetOut = ss.getSheetByName(CONFIG.SHEET_OUT);

  if (!sheetOut) {
    sheetOut = ss.insertSheet(CONFIG.SHEET_OUT);
  } else {
    sheetOut.clear();
  }

  const data = sheetOp.getDataRange().getValues();
  const targetDateValue = Utilities.formatDate(targetDateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd');

  let iolList = [];
  let patientList = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tagVal = String(row[CONFIG.COLS.TAG - 1]).trim();
    const dateVal = row[CONFIG.COLS.DATE - 1];
    if (!dateVal) continue;

    // 強化日期物件解析
    const rowDateObj = new Date(dateVal);
    if (isNaN(rowDateObj.getTime())) continue;
    const rowDateStr = Utilities.formatDate(rowDateObj, Session.getScriptTimeZone(), 'yyyy/MM/dd');

    if (rowDateStr === targetDateValue && tagVal === 'OP') {
      const conditionStr = String(row[CONFIG.COLS.COND - 1]);

      let disease = conditionStr;
      let surgery = "";
      if (conditionStr.includes("s/p")) {
        const parts = conditionStr.split("s/p");
        disease = parts[0].trim();
        surgery = parts.slice(1).join("s/p").trim();
      }

      const surgeryParts = splitFirstLineAndRest_(surgery);
      const surgeryMain = surgeryParts.firstLine;
      const supplementalNote = joinNonEmptyLines_([
        surgeryParts.rest,
        row[CONFIG.COLS.PLAN - 1]
      ]);
      let brand = "", target = "", power = "";
      const iolMatch = surgery.match(/IOL\(([^)]+)\)/i) || surgery.match(/MSICS\(([^)]+)\)/i);
      if (iolMatch) {
        // 強化：先 trim 再 split 避免多餘空白導致陣列錯位
        const iolDetails = iolMatch[1].trim().split(/\s+/);
        brand = iolDetails[0] || "";
        target = iolDetails[1] || "";
        power = iolDetails[2] || "";
        iolList.push([row[CONFIG.COLS.NAME - 1], brand, target, power]);
      }

      patientList.push([
        row[CONFIG.COLS.CHART_NO - 1],
        row[CONFIG.COLS.NAME - 1],
        row[CONFIG.COLS.TEL - 1],
        row[CONFIG.COLS.TIME - 1],
        disease,
        surgeryMain,
        supplementalNote
      ]);
    }
  }

  iolList.sort((a, b) => {
    const brandCompare = String(a[1] || '').localeCompare(String(b[1] || ''), 'zh-Hant');
    if (brandCompare !== 0) return brandCompare;

    return String(a[0] || '').localeCompare(String(b[0] || ''), 'zh-Hant');
  });

  const iolStartRow = 1;
  const iolStartCol = 1;
  sheetOut.getRange(iolStartRow, iolStartCol).setValue(`[ ${targetDateStr} 水晶體清單 ]`).setFontWeight('bold');
  if (iolList.length > 0) {
    sheetOut.getRange(iolStartRow + 1, iolStartCol, 1, 4).setValues([["姓名", "品牌", "目標度數", "水晶體度數"]]).setBackground('#efefef');
    sheetOut.getRange(iolStartRow + 2, iolStartCol, iolList.length, 4).setValues(iolList);
  } else {
    sheetOut.getRange(iolStartRow + 1, iolStartCol).setValue("本日無水晶體資料");
  }

  const patientStartRow = 1;
  const patientStartCol = 6;
  sheetOut.getRange(patientStartRow, patientStartCol).setValue(`[ ${targetDateStr} 病人清單 ]`).setFontWeight('bold');
  if (patientList.length > 0) {
    sheetOut.getRange(patientStartRow + 1, patientStartCol, 1, 7)
      .setValues([["病歷號", "姓名", "TEL", "時間", "疾病", "術式", "補充說明"]])
      .setBackground('#efefef');
    sheetOut.getRange(patientStartRow + 2, patientStartCol, patientList.length, 7).setValues(patientList);
  } else {
    sheetOut.getRange(patientStartRow + 1, patientStartCol).setValue("本日無病人資料");
  }

  sheetOut.autoResizeColumns(1, 12);
  ui.alert('完成', `已成功將 ${targetDateStr} 的資料分析匯出至「輸出表單」。`, ui.ButtonSet.OK);
}

/**
 * 2: 複製一列
 */
function duplicateRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeCell = sheet.getActiveCell();
  const activeRow = activeCell.getRow();

  if (activeRow < 2) return;

  const lastCol = sheet.getLastColumn();
  sheet.insertRowAfter(activeRow);

  const sourceRange = sheet.getRange(activeRow, 1, 1, lastCol);
  const targetRange = sheet.getRange(activeRow + 1, 1, 1, lastCol);

  sourceRange.copyTo(targetRange);

  sheet.getRange(activeRow + 1, CONFIG.COLS.DATE).clearContent();
  sheet.getRange(activeRow + 1, CONFIG.COLS.TIME).clearContent();
  if (lastCol >= CONFIG.COLS.EVENT_ID) {
    sheet.getRange(activeRow + 1, CONFIG.COLS.EVENT_ID).clearContent();
  }
}

/**
 * 3: 歸人整合
 */
function mergePatientRecords() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_OP);
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  if (data.length < 2) return;

  const headers = data[0];
  let grouped = {};
  let finalRows = [];

  const mergeText = (a, b) => {
    a = String(a).trim();
    b = String(b).trim();
    if (!a) return b;
    if (!b) return a;
    if (a === b) return a;
    return a + "\n" + b;
  };

  for (let i = 1; i < data.length; i++) {
    let row = [...data[i]];
    const chartNo = row[CONFIG.COLS.CHART_NO - 1];
    const name = row[CONFIG.COLS.NAME - 1];
    const tag = row[CONFIG.COLS.TAG - 1];

    if (tag === 'OP' || !chartNo || !name) {
      finalRows.push(row);
    } else {
      const key = `${chartNo}_${name}`;
      if (!grouped[key]) {
        grouped[key] = row;
        finalRows.push(grouped[key]);
      } else {
        grouped[key][CONFIG.COLS.COND - 1] = mergeText(grouped[key][CONFIG.COLS.COND - 1], row[CONFIG.COLS.COND - 1]);
        grouped[key][CONFIG.COLS.PLAN - 1] = mergeText(grouped[key][CONFIG.COLS.PLAN - 1], row[CONFIG.COLS.PLAN - 1]);
        grouped[key][CONFIG.COLS.MEMO - 1] = mergeText(grouped[key][CONFIG.COLS.MEMO - 1], row[CONFIG.COLS.MEMO - 1]);

        if (grouped[key][CONFIG.COLS.TAG - 1] !== row[CONFIG.COLS.TAG - 1]) {
          grouped[key][CONFIG.COLS.TAG - 1] = mergeText(grouped[key][CONFIG.COLS.TAG - 1], row[CONFIG.COLS.TAG - 1]);
        }
      }
    }
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  // 從第 2 列開始清除，所以列數應該是 lastRow - 1
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  }

  if (finalRows.length > 0) {
    sheet.getRange(2, 1, finalRows.length, headers.length).setValues(finalRows);
  }

  SpreadsheetApp.getUi().alert('歸人整合完成', '所有非 OP 的同病人紀錄已合併。', SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 4: 依照日期與時間排序
 */
function sortSheetByDateTime() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_OP);
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
  range.sort([
    { column: CONFIG.COLS.DATE, ascending: true },
    { column: CONFIG.COLS.TIME, ascending: true }
  ]);
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

function removeLegacyApiEventIdColumn_(sheet) {
  const legacyColumn = CONFIG.COLS.EVENT_ID + 1;

  if (sheet.getMaxColumns() < legacyColumn) return false;

  const headerText = String(sheet.getRange(1, legacyColumn).getValue() || '').trim();
  if (headerText !== LEGACY_API_EVENT_ID_HEADER) return false;

  sheet.deleteColumn(legacyColumn);
  return true;
}

function ensureHeaders_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length);
  const values = headerRange.getValues()[0];
  const nextValues = values.slice();
  const mismatches = [];
  let hasChanges = false;

  CONFIG.HEADERS.forEach((expectedHeader, index) => {
    const currentValue = values[index];
    const currentText = currentValue === null || currentValue === undefined
      ? ''
      : String(currentValue).trim();

    if (!currentText) {
      nextValues[index] = expectedHeader;
      hasChanges = true;
      return;
    }

    if (currentText !== expectedHeader) {
      mismatches.push(`${columnToLetter_(index + 1)}1：目前為「${currentText}」，建議為「${expectedHeader}」`);
    }
  });

  if (hasChanges) {
    headerRange.setValues([nextValues]);
  }

  return mismatches;
}

function applyDataFormats_(sheet) {
  const dataRowCount = sheet.getMaxRows() - 1;

  sheet.getRange(2, CONFIG.COLS.TEL, dataRowCount, 1).setNumberFormat('@');
  sheet.getRange(2, CONFIG.COLS.TIME, dataRowCount, 1).setNumberFormat('@');
  sheet.getRange(2, CONFIG.COLS.DATE, dataRowCount, 1).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(2, CONFIG.COLS.EVENT_ID, dataRowCount, 1).setNumberFormat('@');

  const tagRange = sheet.getRange(2, CONFIG.COLS.TAG, dataRowCount, 1);
  const ruleValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.TAG_OPTIONS, true)
    .setAllowInvalid(true)
    .build();
  tagRange.setDataValidation(ruleValidation);
}

function applyRowDataFormats_(sheet, row) {
  sheet.getRange(row, CONFIG.COLS.TEL).setNumberFormat('@');
  sheet.getRange(row, CONFIG.COLS.TIME).setNumberFormat('@');
  sheet.getRange(row, CONFIG.COLS.DATE).setNumberFormat('yyyy/mm/dd');
  sheet.getRange(row, CONFIG.COLS.EVENT_ID).setNumberFormat('@');

  const ruleValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.TAG_OPTIONS, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(row, CONFIG.COLS.TAG).setDataValidation(ruleValidation);
}

function hideSystemColumns_(sheet) {
  sheet.hideColumns(CONFIG.COLS.EVENT_ID);
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

function rangeMatchesMainTable_(range) {
  return range.getRow() === 2 && range.getColumn() === 1 && range.getNumColumns() === CONFIG.HEADERS.length;
}

function rangeMatchesLegacyMainTable_(range) {
  return range.getRow() === 2 &&
    range.getColumn() === 1 &&
    (range.getNumColumns() === CONFIG.COLS.EVENT_ID || range.getNumColumns() === CONFIG.COLS.EVENT_ID + 1);
}

function isLegacyOpConditionalRule_(rule) {
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
    ranges.every(range => rangeMatchesSingleColumn_(range, CONFIG.COLS.TAG));
}

function isLegacyPlanConditionalRule_(rule) {
  const formula = getConditionalRuleFormula_(rule);

  if (!formula || formula.indexOf('SEARCH(') === -1 || formula.indexOf('$' + columnToLetter_(CONFIG.COLS.PLAN) + '2') === -1) {
    return false;
  }

  const ranges = rule.getRanges();
  return ranges.length === 1 && (rangeMatchesMainTable_(ranges[0]) || rangeMatchesLegacyMainTable_(ranges[0]));
}

function isManagedConditionalRule_(rule) {
  return conditionalRuleHasSystemMarker_(rule) ||
    isLegacyOpConditionalRule_(rule) ||
    isLegacyPlanConditionalRule_(rule);
}

function getRetainedConditionalRules_(sheet) {
  return sheet.getConditionalFormatRules()
    .filter(rule => !isManagedConditionalRule_(rule));
}

/**
 * 將 CONFIG.CONDITIONAL_FORMAT_KEYWORDS 轉成可套在第 2 列起始的條件格式公式片段。
 *
 * 條件格式公式會從範圍左上角的列開始相對套用，所以 Plan 欄固定寫成 $H2。
 */
function buildPlanKeywordExpression_() {
  const keywords = (CONFIG.CONDITIONAL_FORMAT_KEYWORDS || [])
    .map(keyword => String(keyword).trim())
    .filter(Boolean);

  if (keywords.length === 0) return '';

  const planColLetter = columnToLetter_(CONFIG.COLS.PLAN);
  const formulaParts = keywords.map(keyword => {
    return `ISNUMBER(SEARCH(${toSheetFormulaString_(keyword)}, $${planColLetter}2))`;
  });

  return formulaParts.length === 1
    ? formulaParts[0]
    : `OR(${formulaParts.join(',')})`;
}

/**
 * 套用系統條件格式，並保留使用者自己新增的其他條件格式。
 *
 * OP 紅底要優先於 Plan 黃底：
 * - A:C、E:J 命中 Plan 關鍵字時套黃底。
 * - D 欄只有在 Tag 不是 OP 時才套 Plan 黃底。
 * - D 欄 Tag=OP 時永遠套紅底。
 */
function applyConditionalFormatting_(sheet) {
  const dataRowCount = sheet.getMaxRows() - 1;
  const retainedRules = getRetainedConditionalRules_(sheet);
  const systemRules = [];
  const tagColLetter = columnToLetter_(CONFIG.COLS.TAG);
  const tagRange = sheet.getRange(2, CONFIG.COLS.TAG, dataRowCount, 1);

  systemRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied(`=AND($${tagColLetter}2="OP", N(${toSheetFormulaString_(CONDITIONAL_FORMAT_MARKERS.OP_RED)})=0)`)
    .setBackground('#F4CCCC')
    .setFontColor('#CC0000')
    .setRanges([tagRange])
    .build()
  );

  const planKeywordExpression = buildPlanKeywordExpression_();

  if (planKeywordExpression) {
    const planMainRanges = [
      sheet.getRange(2, 1, dataRowCount, CONFIG.COLS.TAG - 1),
      sheet.getRange(2, CONFIG.COLS.COND, dataRowCount, CONFIG.HEADERS.length - CONFIG.COLS.COND + 1)
    ];
    const planMainFormula = `=AND(${planKeywordExpression}, N(${toSheetFormulaString_(CONDITIONAL_FORMAT_MARKERS.PLAN_YELLOW_MAIN)})=0)`;
    const planTagFormula = `=AND($${tagColLetter}2<>"OP", ${planKeywordExpression}, N(${toSheetFormulaString_(CONDITIONAL_FORMAT_MARKERS.PLAN_YELLOW_TAG)})=0)`;

    systemRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(planMainFormula)
      .setBackground('#FFF2CC')
      .setRanges(planMainRanges)
      .build()
    );

    systemRules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(planTagFormula)
      .setBackground('#FFF2CC')
      .setRanges([tagRange])
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
function normalizeExistingTimes_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { normalizedCount: 0, errorCount: 0 };
  }

  const timeRange = sheet.getRange(2, CONFIG.COLS.TIME, lastRow - 1, 1);
  const values = timeRange.getValues();
  let normalizedCount = 0;
  let errorCount = 0;

  values.forEach((rowValues, index) => {
    const row = index + 2;
    const value = rowValues[0];
    const timeCell = sheet.getRange(row, CONFIG.COLS.TIME);

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
  const removedLegacyColumn = removeLegacyApiEventIdColumn_(sheet);
  const headerMismatches = ensureHeaders_(sheet);
  applyDataFormats_(sheet);
  const timeResult = normalizeExistingTimes_(sheet);
  applyConditionalFormatting_(sheet);
  hideSystemColumns_(sheet);

  const mismatchMessage = headerMismatches.length > 0
    ? `\n\n以下標題已存在且未覆蓋：\n${headerMismatches.join('\n')}`
    : '';
  const legacyColumnMessage = removedLegacyColumn
    ? '\n已移除舊系統欄位「日曆apiEventID」。'
    : '';

  if (showAlert) {
    ui.alert(
      '初始化完成',
      `已完成標題補齊、欄位格式、Tag 下拉選單、條件格式與既有時間正規化。\n` +
        `時間欄位已轉換 ${timeResult.normalizedCount} 格，格式錯誤 ${timeResult.errorCount} 格。` +
        legacyColumnMessage +
        mismatchMessage,
      ui.ButtonSet.OK
    );
  }

  return {
    ok: true,
    timeResult,
    mismatchMessage: legacyColumnMessage + mismatchMessage
  };
}

function initializeSheet() {
  initializeSheet_(true);
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_OP);
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

  let processedCount = 0;
  const statusCounts = {};

  const executed = withCalendarSyncLock_(() => {
    const data = sheet.getRange(2, 1, lastRow - 1, CONFIG.COLS.EVENT_ID).getValues();

    for (let i = 0; i < data.length; i++) {
      const row = i + 2;
      const dateVal = data[i][CONFIG.COLS.DATE - 1];
      const eventId = data[i][CONFIG.COLS.EVENT_ID - 1];

      // 有日期：新增或更新
      // 無日期但有 eventId：刪除既有事件
      if (dateVal || eventId) {
        const status = syncToCalendar(sheet, row);
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
  batchSyncAllEvents_(true);
}

function isCalendarSyncTokenInvalidError_(err) {
  const message = String((err && err.message) || err || '');
  return message.indexOf('410') !== -1 ||
    message.toLowerCase().indexOf('sync token') !== -1 ||
    message.toLowerCase().indexOf('gone') !== -1;
}

function buildCalendarEventRowIndex_(sheet) {
  const index = {
    byEventId: {}
  };
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return index;

  const values = sheet.getRange(2, CONFIG.COLS.EVENT_ID, lastRow - 1, 1).getValues();

  values.forEach((rowValues, indexOffset) => {
    const row = indexOffset + 2;
    const eventId = toCellText_(rowValues[0]);

    if (eventId && !index.byEventId[eventId]) {
      index.byEventId[eventId] = row;
    }
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

function markSheetRowCalendarDeleted_(sheet, row) {
  sheet.getRange(row, CONFIG.COLS.DATE).clearContent();
  sheet.getRange(row, CONFIG.COLS.TIME).clearContent();
  sheet.getRange(row, CONFIG.COLS.EVENT_ID).clearContent();
  addCalendarDeletedMemoNote_(sheet.getRange(row, CONFIG.COLS.MEMO));
}

function updateSheetRowFromCalendarApiEvent_(sheet, row, event, dateTimeInfo) {
  const rowRange = sheet.getRange(row, 1, 1, CONFIG.COLS.EVENT_ID);
  const values = rowRange.getValues()[0];
  const titleInfo = parseCalendarTitle_(event.summary || '');
  const descriptionInfo = parseCalendarDescription_(event.description || '');

  if (titleInfo.ok) {
    values[CONFIG.COLS.CHART_NO - 1] = titleInfo.chartNo;
    values[CONFIG.COLS.NAME - 1] = titleInfo.patientName;
    values[CONFIG.COLS.COND - 1] = titleInfo.condition;
  } else {
    console.warn(`第 ${row} 列 Calendar 標題解析失敗，保留原病人欄位。`);
  }

  if (descriptionInfo.hasTel) {
    values[CONFIG.COLS.TEL - 1] = descriptionInfo.tel;
  }

  if (descriptionInfo.hasPlan) {
    values[CONFIG.COLS.PLAN - 1] = descriptionInfo.plan;
  }

  values[CONFIG.COLS.DATE - 1] = dateTimeInfo.date;
  values[CONFIG.COLS.TIME - 1] = dateTimeInfo.timeText;
  values[CONFIG.COLS.EVENT_ID - 1] = toCellText_(event.id) || values[CONFIG.COLS.EVENT_ID - 1];

  rowRange.setValues([values]);
  applyRowDataFormats_(sheet, row);
}

function appendSheetRowFromCalendarApiEvent_(sheet, event, dateTimeInfo) {
  const titleInfo = parseCalendarTitle_(event.summary || '');

  if (!titleInfo.ok) {
    console.warn('新增 Calendar event 未寫入 Sheet：標題格式不符合系統規則。');
    return 0;
  }

  const descriptionInfo = parseCalendarDescription_(event.description || '');
  const rowValues = new Array(CONFIG.HEADERS.length).fill('');
  rowValues[CONFIG.COLS.CHART_NO - 1] = titleInfo.chartNo;
  rowValues[CONFIG.COLS.NAME - 1] = titleInfo.patientName;
  rowValues[CONFIG.COLS.TEL - 1] = descriptionInfo.hasTel ? descriptionInfo.tel : '';
  rowValues[CONFIG.COLS.TAG - 1] = tagFromCalendarApiEvent_(event);
  rowValues[CONFIG.COLS.COND - 1] = titleInfo.condition;
  rowValues[CONFIG.COLS.DATE - 1] = dateTimeInfo.date;
  rowValues[CONFIG.COLS.TIME - 1] = dateTimeInfo.timeText;
  rowValues[CONFIG.COLS.PLAN - 1] = descriptionInfo.hasPlan ? descriptionInfo.plan : '';
  rowValues[CONFIG.COLS.EVENT_ID - 1] = toCellText_(event.id);

  const row = sheet.getLastRow() + 1;
  if (row > sheet.getMaxRows()) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  }

  sheet.getRange(row, 1, 1, CONFIG.HEADERS.length).setValues([rowValues]);
  applyRowDataFormats_(sheet, row);

  return row;
}

function processCalendarApiEventChange_(sheet, rowIndex, event) {
  if (!event || !event.id) return 'skipped_missing_event_id';

  const row = findSheetRowForCalendarApiEvent_(rowIndex, event);

  if (event.status === 'cancelled') {
    if (!row) return 'skipped_deleted_unmatched';
    markSheetRowCalendarDeleted_(sheet, row);
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
    updateSheetRowFromCalendarApiEvent_(sheet, row, event, dateTimeInfo);
    updateCalendarEventRowIndex_(rowIndex, row, event);
    return 'updated';
  }

  const newRow = appendSheetRowFromCalendarApiEvent_(sheet, event, dateTimeInfo);
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
        const status = processCalendarApiEventChange_(sheet, rowIndex, event);
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
  const result = syncCalendarChangesToSheet_(CONFIG.CALENDAR_ID);
  SpreadsheetApp.getUi().alert(result.message);
}

function processCalendarChange(e) {
  const calendarId = e && e.calendarId ? e.calendarId : CONFIG.CALENDAR_ID;

  withCalendarSyncLock_(() => {
    const result = syncCalendarChangesToSheet_(calendarId);
    if (!result.ok) {
      console.warn(result.message);
    }
  });
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

  withCalendarSyncLock_(() => {
    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();
    const startCol = e.range.getColumn();
    const numCols = e.range.getNumColumns();
    const endCol = startCol + numCols - 1;

    // 檢查編輯範圍是否涵蓋我們關注的欄位
    const watchCols = Object.values(CONFIG.COLS);
    const isWatchColEdited = watchCols.some(col => col >= startCol && col <= endCol);

    if (!isWatchColEdited) return;

    // 逐列處理 (支援同時貼上多列的情況)
    for (let i = 0; i < numRows; i++) {
      const currentRow = startRow + i;
      if (currentRow < 2) continue;

      // 處理手打時間轉換
      // 僅在「單一時間儲存格」編輯時自動改寫顯示值，避免整批貼上時改壞範圍。
      if (numRows === 1 && numCols === 1 && startCol === CONFIG.COLS.TIME) {
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

      syncToCalendar(sheet, currentRow);
    }
  });
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

function extractSpreadsheetId_(input) {
  const text = String(input || '').trim();
  const match = text.match(/\/spreadsheets\/d\/([A-Za-z0-9_-]+)/);
  return match ? match[1] : text;
}

function promptText_(ui, title, message) {
  const response = ui.prompt(title, message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return null;

  return response.getResponseText().trim();
}

function clearSystemCalendarSyncNote_(range) {
  const note = range.getNote();
  if (note && String(note).startsWith(CALENDAR_SYNC_NOTE_PREFIX)) {
    range.clearNote();
  }
}

function clearCalendarEventsFromSpecifiedColumn() {
  const ui = SpreadsheetApp.getUi();

  if (!isCalendarIdConfigured_()) {
    ui.alert('CONFIG.CALENDAR_ID 尚未設定，未清除日曆事件。');
    return;
  }

  if (!isCalendarAdvancedServiceAvailable_()) {
    ui.alert('尚未啟用 Calendar Advanced Service，未清除日曆事件。');
    return;
  }

  const spreadsheetInput = promptText_(ui, '清除舊日曆事件', '請輸入 Spreadsheet URL 或 ID：');
  if (spreadsheetInput === null) return;

  const sheetName = promptText_(ui, '清除舊日曆事件', '請輸入工作表名稱：');
  if (sheetName === null) return;

  const columnInput = promptText_(ui, '清除舊日曆事件', '請輸入日曆 ID 欄位字母，例如 J：');
  if (columnInput === null) return;

  const spreadsheetId = extractSpreadsheetId_(spreadsheetInput);
  const column = columnLetterToNumber_(columnInput);

  if (!spreadsheetId || !column) {
    ui.alert('Spreadsheet ID 或欄位字母格式不正確，未清除日曆事件。');
    return;
  }

  let targetSpreadsheet;
  try {
    targetSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (err) {
    ui.alert(`無法開啟指定 Spreadsheet：${err.message || err}`);
    return;
  }

  const targetSheet = targetSpreadsheet.getSheetByName(sheetName);
  if (!targetSheet) {
    ui.alert(`找不到工作表「${sheetName}」，未清除日曆事件。`);
    return;
  }

  const lastRow = targetSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('指定工作表沒有可處理的資料列。');
    return;
  }

  const result = {
    scannedCount: 0,
    clearedCount: 0,
    deletedCount: 0,
    alreadyMissingCount: 0,
    failedCount: 0
  };

  const executed = withCalendarSyncLock_(() => {
    const idRange = targetSheet.getRange(2, column, lastRow - 1, 1);
    const values = idRange.getValues();

    values.forEach((rowValues, index) => {
      const eventId = toCellText_(rowValues[0]);
      if (!eventId) return;

      result.scannedCount++;
      const cell = idRange.getCell(index + 1, 1);
      const deleteResult = deleteCalendarEventByPossibleId_(CONFIG.CALENDAR_ID, eventId);

      if (deleteResult.ok) {
        cell.clearContent();
        clearSystemCalendarSyncNote_(cell);
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
  });

  if (!executed) {
    ui.alert('目前有另一個日曆同步作業正在執行，本次清除未執行。');
    return;
  }

  ui.alert(
    '清除完成',
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
function syncToCalendar(sheet, row) {
  const rowData = sheet.getRange(row, 1, 1, Math.max(CONFIG.COLS.EVENT_ID, sheet.getLastColumn())).getValues()[0];
  const chartNo = rowData[CONFIG.COLS.CHART_NO - 1];
  const patientName = rowData[CONFIG.COLS.NAME - 1];
  const tel = rowData[CONFIG.COLS.TEL - 1];
  const tag = rowData[CONFIG.COLS.TAG - 1];
  const condition = rowData[CONFIG.COLS.COND - 1];
  const dateVal = rowData[CONFIG.COLS.DATE - 1];
  const timeVal = rowData[CONFIG.COLS.TIME - 1];
  const plan = rowData[CONFIG.COLS.PLAN - 1];
  const eventId = toCellText_(rowData[CONFIG.COLS.EVENT_ID - 1]);
  const chartNoText = toCellText_(chartNo);
  const patientNameText = toSingleLineText_(patientName);
  const telText = toCellText_(tel);
  const tagText = toCellText_(tag);
  const conditionText = toSingleLineText_(condition);
  const planText = toCellText_(plan);
  const chartNoCell = sheet.getRange(row, CONFIG.COLS.CHART_NO);
  const eventIdCell = sheet.getRange(row, CONFIG.COLS.EVENT_ID);

  if (!isCalendarAdvancedServiceAvailable_()) {
    setCalendarSyncNote_(eventIdCell, '需啟用 Calendar Advanced Service 才能同步日曆。');
    return 'skipped_no_calendar_service';
  }

  if (!dateVal) {
    clearCalendarSyncNote_(chartNoCell);

    if (eventId) {
      const deleteResult = deleteCalendarEventByPossibleId_(CONFIG.CALENDAR_ID, eventId);
      if (!deleteResult.ok) {
        setCalendarSyncNote_(eventIdCell, `刪除日曆事件失敗：${deleteResult.message}`);
        return 'skipped_delete_failed';
      }

      eventIdCell.clearContent();
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
    setCalendarSyncNote_(eventIdCell, '這是舊 iCalUID，請先用「清除指定欄位舊日曆事件」刪除並清空後再重建。');
    return 'skipped_legacy_event_id';
  }

  const dateObj = new Date(dateVal);
  if (isNaN(dateObj.getTime())) return 'skipped_invalid_date';

  const timeCell = sheet.getRange(row, CONFIG.COLS.TIME);
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
      const updatedEvent = Calendar.Events.patch(eventResource, CONFIG.CALENDAR_ID, eventId);
      eventIdCell.setValue(updatedEvent.id || eventId);
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
    const createdEvent = Calendar.Events.insert(eventResource, CONFIG.CALENDAR_ID);
    eventIdCell.setValue(createdEvent.id);
    clearSystemCalendarSyncNote_(eventIdCell);
    return 'created';
  } catch (err) {
    setCalendarSyncNote_(eventIdCell, `建立日曆事件失敗：${err.message || err}`);
    return 'skipped_create_failed';
  }
}
