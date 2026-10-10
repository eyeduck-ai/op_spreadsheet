/**
 * 手術排程與 FU 追蹤系統
 *
 * 模組分工：
 * - code.js：入口、選單、設定與共用工具
 * - sheet_model.js：欄位模型、格式、日期區塊與時間正規化
 * - schedule_integrity.js：排程識別、重複綁定檢查與人工輸入格式
 * - workflows.js：FU／月表快照、月份模板與排序
 * - annual_archive.js：私人備份與永久年度封存
 */
const CONFIG = {
  VERSION: '2026.10.11.2',
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
    'EVENT_ID',
    'SYNC_ID'
  ],
  SYSTEM_FIELD_KEYS: ['EVENT_ID', 'SYNC_ID'],
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
    EVENT_ID: 'CalendarEventId',
    SYNC_ID: 'ScheduleSyncId'
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
    'CalendarEventId',
    'ScheduleSyncId'
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
    'EVENT_ID',
    'SYNC_ID'
  ],
  MONTHLY_SYSTEM_FIELD_KEYS: ['EVENT_ID', 'SYNC_ID'],
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
    EVENT_ID: 'CalendarEventId',
    SYNC_ID: 'ScheduleSyncId'
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
    'CalendarEventId',
    'ScheduleSyncId'
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
  PLAN_GREEN_KEYWORDS: ['APPLY']
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
const MONTHLY_STRUCTURE_NOTE_PREFIX = '系統刀日結構：';
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

function withSheetOperationLock_(callback) {
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



function getConfiguredCalendarId_() {
  return toCellText_(PropertiesService.getScriptProperties().getProperty(CALENDAR_ID_PROPERTY));
}





function promptAndSaveCalendarId() {
  return externalCalendarMenu_(); // Legacy menu entry: no Calendar access.
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
  ['processScheduleEdit', 'processRowChange', 'processCalendarStructureChange',
    'retryPendingCalendarSync'].concat(LEGACY_TRIGGER_HANDLERS).forEach(deleteTriggersByHandler_);
  ScriptApp.newTrigger('processScheduleEdit').forSpreadsheet(spreadsheet).onEdit().create();
  const properties = PropertiesService.getScriptProperties();
  Object.keys(properties.getProperties()).filter(key =>
    /^CALENDAR_ROW_REGISTRY_V2(?:_|$)/.test(key) || /^CALENDAR_SYNC_PENDING/.test(key)
  ).forEach(key => properties.deleteProperty(key));
  if (showAlert) SpreadsheetApp.getUi().alert('已修復 onEdit 格式處理並清除舊日曆索引／重試設定。');
  return { ok: true, onEdit: 'processScheduleEdit', onChange: '', pendingRetry: '' };
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
    assertScheduleIntegrity_(
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
    const identityResult = assertScheduleIntegrity_(spreadsheet, {
      allowConflicts: false
    });
    reorderManagedSheets_(spreadsheet);
    onOpen();

    const message = [
      '已完成格式、驗證、觸發器、排程識別及分頁順序修復。',
      '本動作沒有新增、更新或刪除 Calendar 事件。',
      `FU：${fuResult.message}`,
      `月份表：${monthlyResult.message}`,
      `綁定：${identityResult.eventCount} 個 CalendarEventId。`
    ].join('\n');
    SpreadsheetApp.getUi().alert('安裝／修復完成', message, SpreadsheetApp.getUi().ButtonSet.OK);
    return {
      ok: true,
      fuResult,
      monthlyResult,
      identityResult,
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
      '系統欄：CalendarEventId、ScheduleSyncId\n' +
      '日曆由指定 Windows 主機同步；GAS 只處理刀表。'
  );
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const maintenance = ui.createMenu('維護工具')
    .addItem('遷移固定排程識別欄', 'migrateScheduleSyncIds')
    .addItem('交由本系統同步日曆', 'handoverCalendarToApplication')
    .addItem('顯示日曆交接狀態', 'showCalendarHandover')
    .addItem('安裝／修復系統', 'installOrRepairSystem')
    .addItem(
      '套用所有 FU／月表建議欄寬',
      'applyAllRecommendedColumnWidths'
    )
    .addSeparator()
    .addItem('檢查排程識別欄', 'checkScheduleSyncIds')
    .addSeparator()
    .addItem(`版本：${CONFIG.VERSION}`, 'showVersionInfo');

  ui.createMenu('手術排程系統')
    .addItem('建立新月刀表', 'createMonthlySurgerySheet')
    .addItem('選取列新增刀日', 'insertSurgeryDateAtSelection')
    .addItem('整理目前分頁', 'sortCurrentScheduleSheet')
    .addItem('水晶體清單（選刀日／複製）', 'showIolListDialog')
    .addSeparator()
    .addItem('月刀表 => FU', 'addSelectedMonthlyRowToFu')
    .addItem('FU => 月刀表', 'showFuToMonthlyDialog')
    .addSeparator()
    .addItem('彙整舊月刀表', 'showMonthlyRollupDialog')
    .addSubMenu(maintenance)
    .addToUi();
}
