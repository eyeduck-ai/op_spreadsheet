/**
 * 手術排程與 FU 追蹤系統
 *
 * 模組分工：
 * - code.js：入口、選單、設定與共用工具
 * - sheet_model.js：欄位模型、格式、日期區塊與時間正規化
 * - calendar_sync.js：Calendar 生命週期與 CALENDAR_ROW_REGISTRY_V2
 * - workflows.js：FU／月表快照、月份模板與排序
 * - annual_archive.js：私人備份與永久年度封存
 */
const CONFIG = {
  VERSION: '2026.09.19',
  CALENDAR_ID: 'YOUR_CALENDAR_ID_HERE',
  SHEET_ALL: 'All',
  SHEET_FU: 'FU',
  LEGACY_SHEET_OP: 'OP',
  SHEET_IVI: 'IVI',
  FIELD_KEYS: [
    'CHART_NO',
    'NAME',
    'TEL',
    'HOSPITAL',
    'TAG',
    'COND',
    'DATE',
    'PLAN',
    'MEMO',
    'EVENT_ID'
  ],
  SYSTEM_FIELD_KEYS: ['EVENT_ID'],
  FIELD_HEADERS: {
    CHART_NO: '病歷號',
    NAME: '姓名',
    TEL: 'TEL',
    HOSPITAL: '醫院',
    TAG: 'Tag',
    COND: 'Condition',
    DATE: '日期',
    PLAN: 'Plan',
    MEMO: '心得',
    EVENT_ID: 'CalendarEventId'
  },
  HEADERS: [
    '病歷號',
    '姓名',
    'TEL',
    '醫院',
    'Tag',
    'Condition',
    '日期',
    'Plan',
    '心得',
    'CalendarEventId'
  ],
  MONTHLY_FIELD_KEYS: [
    'TIME',
    'HOSPITAL',
    'CHART_NO',
    'NAME',
    'TEL',
    'GA',
    'SIDE',
    'DIAGNOSIS',
    'GRADE',
    'PROCEDURE',
    'IOL',
    'IOL_TARGET',
    'IOL_FINAL',
    'AXIS',
    'PLAN',
    'MEMO',
    'SN',
    'CDE',
    'ENERGY_TIME',
    'ENERGY_PERCENT',
    'REFRACTION',
    'EVENT_ID'
  ],
  MONTHLY_SYSTEM_FIELD_KEYS: ['EVENT_ID'],
  MONTHLY_FIELD_HEADERS: {
    TIME: '日期／報到時間',
    HOSPITAL: '醫院',
    CHART_NO: '病歷號',
    NAME: '姓名',
    TEL: 'TEL',
    GA: 'GA',
    SIDE: '側別',
    DIAGNOSIS: '診斷',
    GRADE: 'Grade',
    PROCEDURE: '術式',
    PLAN: 'Plan',
    IOL: 'IOL',
    IOL_TARGET: 'IOL Target',
    IOL_FINAL: 'IOL Final',
    AXIS: 'Axis',
    MEMO: '心得',
    SN: 'SN',
    CDE: 'CDE',
    ENERGY_TIME: 'Energy Time',
    ENERGY_PERCENT: 'Energy %',
    REFRACTION: '屈光數據',
    EVENT_ID: 'CalendarEventId'
  },
  MONTHLY_HEADERS: [
    '日期／報到時間',
    '醫院',
    '病歷號',
    '姓名',
    'TEL',
    'GA',
    '側別',
    '診斷',
    'Grade',
    '術式',
    'IOL',
    'IOL Target',
    'IOL Final',
    'Axis',
    'Plan',
    '心得',
    '',
    'SN',
    'CDE',
    'Energy Time',
    'Energy %',
    '屈光數據',
    'CalendarEventId'
  ],
  HOSPITAL_OPTIONS: ['高榮', '聯醫'],
  FU_TAG_OPTIONS: [
    'FU',
    'Complication',
    'CATA',
    'Cornea',
    'Retina',
    'Plasty',
    'Neuro',
    'Refraction',
    'Suture IOL',
    'IVI'
  ],
  MONTHLY_GA_OPTIONS: ['GA'],
  MONTHLY_SIDE_OPTIONS: ['OD', 'OS', 'OU'],
  MONTHLY_DIAGNOSIS_OPTIONS: [
    'CATA',
    'Subluxated IOL',
    'VH',
    'ERM',
    'RD',
    'Ptosis',
    'Dermatochalasis',
    'Entropion'
  ],
  MONTHLY_DIAGNOSIS_SUMMARY_GROUPS: [
    {
      label: 'CATA',
      keywords: ['CATA', 'Cataract']
    },
    {
      label: 'Retina',
      keywords: [
        'VH',
        'ERM',
        'Subluxation',
        'Dislocation',
        'RRD',
        'TRD',
        'Subluxated IOL',
        'RD'
      ]
    },
    {
      label: 'Plasty',
      keywords: ['Dermatochalasis', 'Ptosis', 'Dacryocystitis', 'Entropion']
    }
  ],
  MONTHLY_PROCEDURE_OPTIONS: [
    'Phaco-IOL',
    'LenSx-Phaco-IOL',
    'MSICS-IOL',
    'VT',
    'VT+MP',
    'LMR',
    'correction',
    'Suture IOL'
  ],
  PLAN_RED_KEYWORDS: ['!'],
  PLAN_YELLOW_KEYWORDS: ['#'],
  PLAN_GREEN_KEYWORDS: ['APPLY'],
  COLORS: {
    MONTHLY_TIMED: '10',
    MONTHLY_UNDECIDED: '2',
    MONTHLY_GA: '4',
    MONTHLY_CANCELLED: '8',
    FU: '8'
  }
};

const CALENDAR_ID_PROPERTY = 'CALENDAR_ID';
const LEGACY_EVENT_ID_HEADERS = ['日曆eventID', '日曆apiEventID'];
const LEGACY_MONTHLY_TIME_HEADERS = ['時間'];
const MONTHLY_DATE_HEADER_MARKER = '◆ 刀日';
const MONTHLY_DIAGNOSIS_SUMMARY_FORMULA_MARKER =
  'SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V2';
const MONTHLY_DIAGNOSIS_SUMMARY_NOTE_PREFIX = '系統月表診斷統計：';
const HOSPITAL_KAOH = '高榮';
const HOSPITAL_UNION = '聯醫';
const MONTHLY_TEMPLATE_BLANK_ROWS = 5;
const MANAGED_FONT_FAMILY = 'Arial';
const MANAGED_FONT_SIZE = 10;
const TABLE_HEADER_BACKGROUND = '#B7B7B7';
const MONTHLY_DATE_HEADER_BACKGROUND = '#D9D9D9';
const MONTHLY_GA_BACKGROUND = '#F4CCCC';
const PLAN_RED_BACKGROUND = '#F4CCCC';
const PLAN_GREEN_BACKGROUND = '#D9EAD3';
const PLAN_YELLOW_BACKGROUND = '#FFF2CC';
const SIDE_OD_BACKGROUND = '#D9EAD3';
const SIDE_OU_BACKGROUND = '#D9D2E9';
const TIME_ERROR_NOTE_PREFIX = '系統時間檢查：';
const CALENDAR_SYNC_NOTE_PREFIX = '系統日曆同步：';
const CALENDAR_CONFLICT_NOTE_PREFIX = '系統同步衝突：';
const MONTHLY_STRUCTURE_NOTE_PREFIX = '系統刀日結構：';
const CALENDAR_TITLE_SEPARATOR = '|';
const CALENDAR_PRIVATE_KIND_KEY = 'surgerySyncKind';
const CALENDAR_PRIVATE_STATE_KEY = 'surgerySyncState';
const CALENDAR_PRIVATE_MONTHLY_KIND = 'MONTHLY';
const CALENDAR_PRIVATE_CANCELLED_STATE = 'CANCELLED';
const CALENDAR_ON_EDIT_HANDLER = 'processRowChange';
const CALENDAR_ON_CHANGE_HANDLER = 'processCalendarStructureChange';
const CALENDAR_PENDING_RETRY_HANDLER = 'retryPendingCalendarSync';
const CALENDAR_TRIGGER_LOCK_WAIT_MS = 750;
const CALENDAR_PENDING_RETRY_DELAY_MS = 60 * 1000;
const CALENDAR_PENDING_RETRY_LOCK_WAIT_MS = 20 * 1000;
const CALENDAR_PENDING_QUEUE_PROPERTY = 'CALENDAR_SYNC_PENDING_V3';
const CALENDAR_PENDING_SPREADSHEET_ID_PROPERTY =
  'CALENDAR_SYNC_PENDING_SPREADSHEET_ID_V1';
const CALENDAR_PENDING_QUEUE_VERSION = 1;
const CALENDAR_PENDING_QUEUE_MAX_ROWS = 500;
const LEGACY_TRIGGER_HANDLERS = [
  'processMonthlyStructureChange',
  'processCalendarChange',
  'exportUpcomingWeekData',
  'runPendingUpcomingWeekExport',
  'autoExportByEdit'
];
const CROSS_SHEET_CACHE_SECONDS = 10 * 60;
const MONTHLY_TO_FU_CACHE_PREFIX = 'MONTHLY_TO_FU_';
const FU_TO_MONTHLY_CACHE_PREFIX = 'FU_TO_MONTHLY_';
const ANNUAL_ARCHIVE_PREFIX = '刀表封存_';
const ANNUAL_ARCHIVE_NOTE_MARKER = 'SYSTEM_ANNUAL_SURGERY_ARCHIVE_V1';
const ANNUAL_ARCHIVE_HEADER_ROW = 6;
const ANNUAL_ARCHIVE_SOURCE_MONTH_HEADER = 'ArchiveSourceMonth';
const ANNUAL_ARCHIVE_ROW_TYPE_HEADER = 'ArchiveRowType';
const ANNUAL_ARCHIVE_TRANSACTION_PROPERTY =
  'ANNUAL_SURGERY_ARCHIVE_TRANSACTION_V1';
const LEGACY_ARCHIVED_MONTHS_PROPERTY = 'ARCHIVED_MONTHLY_SHEETS_V1';

function toCellText_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function toSingleLineText_(value) {
  return toCellText_(value).replace(/\s+/g, ' ');
}

function getRowFieldValue_(rowValues, columns, key) {
  const column = columns[key];
  return column ? rowValues[column - 1] : '';
}

function setRowFieldValue_(rowValues, columns, key, value) {
  const column = columns[key];
  if (!column) return;
  while (rowValues.length < column) rowValues.push('');
  rowValues[column - 1] = value;
}

function columnToLetter_(column) {
  let value = Number(column) || 0;
  let result = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function formatDateKey_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatMonthKey_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMM');
}

function formatCompactDate_(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyyMMdd');
}

function safeJsonParse_(rawValue, fallback) {
  try {
    return JSON.parse(rawValue);
  } catch (err) {
    return fallback;
  }
}

function sha256Hex_(value) {
  const text = String(value === undefined ? '' : value);
  if (
    typeof Utilities !== 'undefined' &&
    Utilities.computeDigest &&
    Utilities.DigestAlgorithm
  ) {
    const bytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      text,
      Utilities.Charset.UTF_8
    );
    return bytes
      .map(byte => {
        const unsigned = byte < 0 ? byte + 256 : byte;
        return unsigned.toString(16).padStart(2, '0');
      })
      .join('');
  }

  // Node smoke tests use this deterministic fallback.
  let hash = 2166136261;
  for (let index = 0; index < text.length; index++) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return ('00000000' + (hash >>> 0).toString(16)).slice(-8);
}

function runMenuAction_(title, callback) {
  const ui = SpreadsheetApp.getUi();
  const startedAt = Date.now();
  let succeeded = false;
  try {
    const result = callback();
    succeeded = !result || result.ok !== false;
    return result;
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    ui.alert(
      title,
      (err && err.message) || String(err),
      ui.ButtonSet.OK
    );
    return {
      ok: false,
      message: (err && err.message) || String(err)
    };
  } finally {
    console.info(JSON.stringify({
      metric: 'SURGERY_SYSTEM_ACTION',
      action: toSingleLineText_(title),
      ok: succeeded,
      durationMs: Date.now() - startedAt
    }));
  }
}

function createPerformancePhaseTimer_(action) {
  const startedAt = Date.now();
  let checkpointAt = startedAt;
  const phases = {};
  return {
    mark(name) {
      const now = Date.now();
      phases[toSingleLineText_(name)] = now - checkpointAt;
      checkpointAt = now;
      return this;
    },
    log() {
      console.info(JSON.stringify({
        metric: 'SURGERY_SYSTEM_PHASES',
        action: toSingleLineText_(action),
        durationMs: Date.now() - startedAt,
        phases
      }));
      return { ...phases };
    }
  };
}

function withCalendarSyncLock_(callback) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('目前有另一個同步作業正在執行，請稍後再試。');
  }
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function tryWithCalendarSyncLock_(callback, waitMs) {
  const lock = LockService.getDocumentLock() || LockService.getScriptLock();
  if (!lock.tryLock(
    waitMs === undefined ? CALENDAR_TRIGGER_LOCK_WAIT_MS : waitMs
  )) {
    return { acquired: false, value: null };
  }
  try {
    return { acquired: true, value: callback() };
  } finally {
    lock.releaseLock();
  }
}

function getConfiguredCalendarId_() {
  const propertyValue = PropertiesService.getScriptProperties()
    .getProperty(CALENDAR_ID_PROPERTY);
  const configured = toCellText_(propertyValue) || toCellText_(CONFIG.CALENDAR_ID);
  return configured === 'YOUR_CALENDAR_ID_HERE' ? '' : configured;
}

function isCalendarAdvancedServiceAvailable_() {
  return typeof Calendar !== 'undefined' &&
    Calendar.Events &&
    typeof Calendar.Events.get === 'function' &&
    typeof Calendar.Events.insert === 'function' &&
    typeof Calendar.Events.update === 'function' &&
    typeof Calendar.Events.remove === 'function';
}

function validateCalendarAccess_(calendarId) {
  if (!isCalendarAdvancedServiceAvailable_()) {
    throw new Error('Calendar Advanced Service 尚未啟用。');
  }
  try {
    if (
      Calendar.Calendars &&
      typeof Calendar.Calendars.get === 'function'
    ) {
      Calendar.Calendars.get(calendarId);
    } else if (typeof Calendar.Events.list === 'function') {
      Calendar.Events.list(calendarId, {
        maxResults: 1,
        singleEvents: true
      });
    } else {
      throw new Error('目前的 Calendar Advanced Service 無法執行唯讀驗證。');
    }
  } catch (err) {
    throw new Error(
      `無法讀取指定 Calendar；原設定未變更：${err.message || err}`
    );
  }
  return true;
}

function promptAndSaveCalendarId() {
  return runMenuAction_('設定日曆', () => {
    const ui = SpreadsheetApp.getUi();
    const current = getConfiguredCalendarId_();
    const response = ui.prompt(
      '設定日曆',
      (current ? `目前 Calendar ID：${current}\n` : '') +
        '請輸入要同步的 Google Calendar ID：',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      return { ok: false, cancelled: true };
    }
    const value = toCellText_(response.getResponseText());
    if (!value) throw new Error('Calendar ID 不可空白。');
    validateCalendarAccess_(value);
    PropertiesService.getScriptProperties()
      .setProperty(CALENDAR_ID_PROPERTY, value);
    installOrRepairTriggers_(false);
    ui.alert('Calendar ID 已儲存；既有資料不會被自動批次重寫。');
    return { ok: true, calendarId: value };
  });
}

function deleteTriggersByHandler_(handlerName) {
  let count = 0;
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
      count++;
    }
  });
  return count;
}

function installOrRepairTriggers_(showAlert) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  [
    CALENDAR_ON_EDIT_HANDLER,
    CALENDAR_ON_CHANGE_HANDLER,
    CALENDAR_PENDING_RETRY_HANDLER
  ].concat(LEGACY_TRIGGER_HANDLERS).forEach(deleteTriggersByHandler_);

  ScriptApp.newTrigger(CALENDAR_ON_EDIT_HANDLER)
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();
  ScriptApp.newTrigger(CALENDAR_ON_CHANGE_HANDLER)
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();
  if (pendingCalendarQueueHasWork_(readPendingCalendarQueue_())) {
    ensurePendingCalendarRetryScheduled_(spreadsheet);
  }

  const result = {
    ok: true,
    onEdit: CALENDAR_ON_EDIT_HANDLER,
    onChange: CALENDAR_ON_CHANGE_HANDLER,
    pendingRetry: CALENDAR_PENDING_RETRY_HANDLER
  };
  if (showAlert) {
    SpreadsheetApp.getUi().alert(
      '已安裝／修復各一個 onEdit 與 onChange 自動同步觸發器；' +
        '有待處理資料時會建立一次性的延遲重試。'
    );
  }
  return result;
}

function installOrRepairSystem() {
  return setup();
}

function setup() {
  return runMenuAction_('安裝／修復系統', () => {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    assertNoLegacyOpOnly_(spreadsheet);
    assertNoLegacyArchiveState_();
    assertNoActiveAnnualArchiveTransaction_();
    assertCalendarRegistrySafeToRebuild_(
      spreadsheet,
      '安裝／修復',
      { allowRepairableStructureActions: true }
    );
    const fuResult = initializeMainTrackingSheet_(spreadsheet, false);
    const monthlyResult = initializeAllMonthlySheets_(
      spreadsheet,
      false,
      { normalizeTimes: false }
    );
    installOrRepairTriggers_(false);
    const registryResult = rebuildCalendarRegistry_(spreadsheet, {
      allowConflicts: false
    });
    reorderManagedSheets_(spreadsheet);
    onOpen();

    const message = [
      '已完成格式、驗證、觸發器、V2 同步索引及分頁順序修復。',
      '本動作沒有新增、更新或刪除 Calendar 事件。',
      `FU：${fuResult.message}`,
      `月份表：${monthlyResult.message}`,
      `索引：${registryResult.eventCount} 個 CalendarEventId。`
    ].join('\n');
    SpreadsheetApp.getUi().alert('安裝／修復完成', message, SpreadsheetApp.getUi().ButtonSet.OK);
    return {
      ok: true,
      fuResult,
      monthlyResult,
      registryResult,
      message
    };
  });
}

function assertNoLegacyOpOnly_(spreadsheet) {
  const hasCurrent =
    spreadsheet.getSheetByName(CONFIG.SHEET_FU) ||
    spreadsheet.getSheetByName(CONFIG.SHEET_ALL);
  if (!hasCurrent && spreadsheet.getSheetByName(CONFIG.LEGACY_SHEET_OP)) {
    throw new Error(
      '偵測到舊分頁「OP」。請先手動將它改名為「All」，系統不會自動改名。'
    );
  }
}

function getMainTrackingSheet_(spreadsheet) {
  const ss = spreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(CONFIG.SHEET_FU) ||
    ss.getSheetByName(CONFIG.SHEET_ALL) ||
    null;
}

function isMainTrackingSheetName_(sheetName) {
  return sheetName === CONFIG.SHEET_FU || sheetName === CONFIG.SHEET_ALL;
}

function isMonthlySheetName_(sheetName) {
  const match = String(sheetName || '').match(/^(\d{4})(\d{2})$/);
  return Boolean(match && Number(match[2]) >= 1 && Number(match[2]) <= 12);
}

function showVersionInfo() {
  SpreadsheetApp.getUi().alert(
    `手術排程系統版本：${CONFIG.VERSION}\n` +
      '系統欄：CalendarEventId\n' +
      '同步索引：CALENDAR_ROW_REGISTRY_V2'
  );
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const maintenance = ui.createMenu('維護工具')
    .addItem('設定日曆', 'promptAndSaveCalendarId')
    .addItem('安裝／修復系統', 'installOrRepairSystem')
    .addItem(
      '套用所有 FU／月表建議欄寬',
      'applyAllRecommendedColumnWidths'
    )
    .addSeparator()
    .addItem('檢查同步健康', 'checkSyncHealth')
    .addItem('同步待處理變更', 'syncPendingChanges')
    .addSeparator()
    .addItem(`版本：${CONFIG.VERSION}`, 'showVersionInfo');

  ui.createMenu('手術排程系統')
    .addItem('建立新月刀表', 'createMonthlySurgerySheet')
    .addItem('選取列新增刀日', 'insertSurgeryDateAtSelection')
    .addItem('整理目前分頁', 'sortCurrentScheduleSheet')
    .addItem('水晶體清單（選刀日／複製）', 'showIolListDialog')
    .addItem('修復選取列同步', 'repairSelectedCalendarRows')
    .addSeparator()
    .addItem('月刀表 => FU', 'addSelectedMonthlyRowToFu')
    .addItem('FU => 月刀表', 'showFuToMonthlyDialog')
    .addSeparator()
    .addItem('彙整舊月刀表', 'showMonthlyRollupDialog')
    .addSubMenu(maintenance)
    .addToUi();
}
