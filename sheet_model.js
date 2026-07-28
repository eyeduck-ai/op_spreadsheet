/**
 * 表格模型、格式、日期區塊與時間正規化。
 */
const CONDITIONAL_FORMAT_MARKER_PREFIX = 'SURGERY_SYSTEM_CF_';
const LEGACY_CONDITIONAL_FORMAT_MARKER_PREFIX = 'SYSTEM_CF_';
const FU_WRAP_KEYS = ['TEL', 'COND', 'PLAN', 'MEMO'];
const MONTHLY_WRAP_KEYS = [
  'TEL',
  'DIAGNOSIS',
  'GRADE',
  'PROCEDURE',
  'PLAN',
  'IOL',
  'MEMO',
  'REFRACTION'
];
const MONTHLY_TEXT_KEYS = [
  'IOL',
  'IOL_TARGET',
  'IOL_FINAL',
  'AXIS',
  'SN',
  'CDE',
  'ENERGY_TIME',
  'ENERGY_PERCENT'
];
const FU_CALENDAR_KEYS = [
  'CHART_NO',
  'NAME',
  'TEL',
  'HOSPITAL',
  'COND',
  'DATE',
  'PLAN'
];
const MONTHLY_CALENDAR_KEYS = [
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
  'PLAN',
  'IOL',
  'IOL_TARGET',
  'IOL_FINAL',
  'AXIS'
];

const FU_COLUMN_WIDTHS = {
  CHART_NO: 76,
  NAME: 72,
  TEL: 96,
  HOSPITAL: 58,
  TAG: 90,
  COND: 320,
  DATE: 96,
  PLAN: 250,
  MEMO: 250,
  EVENT_ID: 120
};

const MONTHLY_ROW_TYPES = Object.freeze({
  DATE_HEADER: 'DATE_HEADER',
  PATIENT: 'PATIENT',
  HYBRID_CONFLICT: 'HYBRID_CONFLICT',
  BLANK: 'BLANK'
});

const MONTHLY_COLUMN_WIDTHS = {
  TIME: 115,
  HOSPITAL: 58,
  CHART_NO: 76,
  NAME: 72,
  TEL: 96,
  GA: 42,
  SIDE: 46,
  DIAGNOSIS: 160,
  GRADE: 82,
  PROCEDURE: 180,
  PLAN: 260,
  IOL: 96,
  IOL_TARGET: 80,
  IOL_FINAL: 76,
  AXIS: 52,
  MEMO: 260,
  SN: 65,
  CDE: 65,
  ENERGY_TIME: 90,
  ENERGY_PERCENT: 72,
  REFRACTION: 180,
  EVENT_ID: 120
};

function getSheetHeaderValues_(sheet) {
  const count = Math.max(1, sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, count).getValues()[0];
}

function getLastHeaderColumnFromValues_(headers) {
  for (let index = headers.length - 1; index >= 0; index--) {
    if (toCellText_(headers[index])) return index + 1;
  }
  return 1;
}

function getLastHeaderColumn_(sheet) {
  return getLastHeaderColumnFromValues_(getSheetHeaderValues_(sheet));
}

function buildHeaderColumnMap_(sheet, headerByKey, fieldKeys, options) {
  const settings = options || {};
  const headers = getSheetHeaderValues_(sheet);
  const headerToKey = {};
  (fieldKeys || []).forEach(key => {
    headerToKey[headerByKey[key]] = key;
  });
  (settings.aliases || []).forEach(alias => {
    headerToKey[alias.header] = alias.key;
  });

  const columns = {};
  const occurrences = {};
  headers.forEach((value, index) => {
    const header = toCellText_(value);
    if (!header) return;
    if (!occurrences[header]) occurrences[header] = [];
    occurrences[header].push(index + 1);
    const key = headerToKey[header];
    if (key && !columns[key]) columns[key] = index + 1;
  });

  const duplicateMessages = Object.keys(occurrences)
    .filter(header => occurrences[header].length > 1)
    .map(header => {
      return `「${header}」重複出現在 ${occurrences[header]
        .map(columnToLetter_)
        .join('、')} 欄`;
    });
  const missingKeys = (fieldKeys || []).filter(key => !columns[key]);

  return {
    columns,
    headers,
    missingKeys,
    duplicateMessages,
    lastColumn: Math.max(getLastHeaderColumnFromValues_(headers), 1)
  };
}

function getFuColumnInfo_(sheet) {
  return buildHeaderColumnMap_(
    sheet,
    CONFIG.FIELD_HEADERS,
    CONFIG.FIELD_KEYS,
    {
      aliases: LEGACY_EVENT_ID_HEADERS.map(header => ({
        header,
        key: 'EVENT_ID'
      }))
    }
  );
}

function getMonthlyColumnInfo_(sheet) {
  return buildHeaderColumnMap_(
    sheet,
    CONFIG.MONTHLY_FIELD_HEADERS,
    CONFIG.MONTHLY_FIELD_KEYS,
    {
      aliases: LEGACY_MONTHLY_TIME_HEADERS.map(header => ({
        header,
        key: 'TIME'
      })).concat(
        LEGACY_EVENT_ID_HEADERS.map(header => ({
          header,
          key: 'EVENT_ID'
        }))
      )
    }
  );
}

function assertUniqueHeaders_(info, label) {
  if (info.duplicateMessages.length) {
    throw new Error(`${label}欄名必須唯一：\n${info.duplicateMessages.join('\n')}`);
  }
}

function getRequiredFuColumns_(sheet) {
  const info = getFuColumnInfo_(sheet);
  assertUniqueHeaders_(info, 'FU ');
  if (info.missingKeys.length) {
    throw new Error(
      `FU 缺少必要欄位：${info.missingKeys
        .map(key => CONFIG.FIELD_HEADERS[key])
        .join('、')}。請先執行「安裝／修復系統」。`
    );
  }
  return info.columns;
}

function getRequiredMonthlyColumns_(sheet) {
  const info = getMonthlyColumnInfo_(sheet);
  assertUniqueHeaders_(info, '月份刀表 ');
  if (info.missingKeys.length) {
    throw new Error(
      `月份刀表缺少必要欄位：${info.missingKeys
        .map(key => CONFIG.MONTHLY_FIELD_HEADERS[key])
        .join('、')}。請先執行「安裝／修復系統」。`
    );
  }
  return info.columns;
}

function ensureColumnCapacity_(sheet, requiredColumns) {
  const shortage = requiredColumns - sheet.getMaxColumns();
  if (shortage > 0) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), shortage);
  }
}

function ensureManagedHeaders_(sheet, kind) {
  const isFu = kind === 'FU';
  const fieldKeys = isFu ? CONFIG.FIELD_KEYS : CONFIG.MONTHLY_FIELD_KEYS;
  const headerByKey = isFu
    ? CONFIG.FIELD_HEADERS
    : CONFIG.MONTHLY_FIELD_HEADERS;
  const preferredHeaders = isFu ? CONFIG.HEADERS : CONFIG.MONTHLY_HEADERS;
  let info = isFu ? getFuColumnInfo_(sheet) : getMonthlyColumnInfo_(sheet);

  assertUniqueHeaders_(info, isFu ? 'FU ' : '月份刀表 ');
  const hasAnyHeader = info.headers.some(value => toCellText_(value));
  if (!hasAnyHeader) {
    ensureColumnCapacity_(sheet, preferredHeaders.length);
    sheet.getRange(1, 1, 1, preferredHeaders.length)
      .setValues([preferredHeaders]);
  } else {
    // 舊欄名只作一次性相容升級，不建立第二個 event ID 欄。
    if (!info.columns.EVENT_ID) {
      const legacyColumn = info.headers.findIndex(value => {
        return LEGACY_EVENT_ID_HEADERS.indexOf(toCellText_(value)) !== -1;
      });
      if (legacyColumn >= 0) {
        sheet.getRange(1, legacyColumn + 1)
          .setValue(headerByKey.EVENT_ID);
      }
    } else {
      const currentHeader = toCellText_(info.headers[info.columns.EVENT_ID - 1]);
      if (LEGACY_EVENT_ID_HEADERS.indexOf(currentHeader) !== -1) {
        sheet.getRange(1, info.columns.EVENT_ID)
          .setValue(headerByKey.EVENT_ID);
      }
    }

    if (!isFu && info.columns.TIME) {
      const currentTimeHeader = toCellText_(info.headers[info.columns.TIME - 1]);
      if (LEGACY_MONTHLY_TIME_HEADERS.indexOf(currentTimeHeader) !== -1) {
        sheet.getRange(1, info.columns.TIME)
          .setValue(CONFIG.MONTHLY_FIELD_HEADERS.TIME);
      }
    }

    info = isFu ? getFuColumnInfo_(sheet) : getMonthlyColumnInfo_(sheet);
    let appendColumn = Math.max(info.lastColumn, 1);
    info.missingKeys.forEach(key => {
      appendColumn++;
      ensureColumnCapacity_(sheet, appendColumn);
      sheet.getRange(1, appendColumn).setValue(headerByKey[key]);
    });
  }

  info = isFu ? getFuColumnInfo_(sheet) : getMonthlyColumnInfo_(sheet);
  assertUniqueHeaders_(info, isFu ? 'FU ' : '月份刀表 ');
  if (info.missingKeys.length) {
    throw new Error(`無法補齊欄位：${info.missingKeys.join('、')}`);
  }
  hideAndProtectSystemColumns_(sheet, info.columns, ['EVENT_ID']);
  return info.columns;
}

function hideAndProtectSystemColumns_(sheet, columns, fieldKeys) {
  (fieldKeys || []).forEach(key => {
    const column = columns[key];
    if (!column) return;
    sheet.hideColumns(column);
    const description = `SURGERY_SYSTEM_COLUMN:${key}`;
    const existing = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .find(item => item.getDescription() === description);
    if (!existing) {
      const protection = sheet.getRange(1, column, sheet.getMaxRows(), 1)
        .protect()
        .setDescription(description);
      if (typeof protection.setWarningOnly === 'function') {
        protection.setWarningOnly(true);
      }
    }
  });
}

function formatTimeText_(hour, minute) {
  return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
}

function getRawTimeText_(value) {
  return value === null || value === undefined
    ? ''
    : String(value).trim().replace(/：/g, ':');
}

function parseTimeInput_(value) {
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date && !isNaN(value.getTime())) {
    return {
      hour: value.getHours(),
      minute: value.getMinutes(),
      text: formatTimeText_(value.getHours(), value.getMinutes())
    };
  }

  if (typeof value === 'number' && isFinite(value)) {
    if (value >= 0 && value < 1) {
      const totalMinutes = Math.round(value * 24 * 60);
      const hour = Math.floor(totalMinutes / 60) % 24;
      const minute = totalMinutes % 60;
      return { hour, minute, text: formatTimeText_(hour, minute) };
    }
    if (Number.isInteger(value) && value >= 0 && value <= 9999) {
      value = String(value);
    }
  }

  const raw = getRawTimeText_(value);
  if (!raw) return null;
  let hour;
  let minute;
  const colon = raw.match(/^(\d{1,2}):(\d{2})$/);
  if (colon) {
    hour = Number(colon[1]);
    minute = Number(colon[2]);
  } else if (/^\d{3,4}$/.test(raw)) {
    const compact = raw.padStart(4, '0');
    hour = Number(compact.slice(0, 2));
    minute = Number(compact.slice(2));
  } else {
    throw new Error(
      `時間格式錯誤：${raw}。請輸入 HH:mm 或 HHmm，例如 08:00 或 0800。`
    );
  }

  if (hour < 0 || hour > 23 || minute < 0 || minute > 59) {
    throw new Error(`時間範圍錯誤：${raw}。小時需為 0–23，分鐘需為 0–59。`);
  }
  return { hour, minute, text: formatTimeText_(hour, minute) };
}

function looksLikeTimeInput_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return true;
  if (typeof value === 'number' && isFinite(value)) return true;
  const raw = getRawTimeText_(value);
  return /^\d{3,4}$/.test(raw) || /^\d{1,2}:\d{2}$/.test(raw);
}

function resolveCalendarTime_(value) {
  const raw = getRawTimeText_(value);
  if (!raw) {
    return { parsedTime: null, timeNote: '', errorMessage: '' };
  }
  try {
    return {
      parsedTime: parseTimeInput_(value),
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

function getSystemNoteBlockMarkers_(prefix) {
  const label = toCellText_(prefix).replace(/[：:\s]+$/g, '') || '系統訊息';
  return {
    start: `【${label}】`,
    end: `【/${label}】`
  };
}

function removeSystemNoteBlockText_(note, prefix) {
  const markers = getSystemNoteBlockMarkers_(prefix);
  const lines = toCellText_(note).replace(/\r\n/g, '\n').split('\n');
  const kept = [];
  let inBlock = false;
  lines.forEach(line => {
    if (line === markers.start) {
      inBlock = true;
      return;
    }
    if (inBlock && line === markers.end) {
      inBlock = false;
      return;
    }
    if (!inBlock) kept.push(line);
  });

  const firstTextLine = kept.findIndex(line => toCellText_(line));
  if (
    firstTextLine >= 0 &&
    toCellText_(kept[firstTextLine]).indexOf(toCellText_(prefix)) === 0
  ) {
    kept.splice(firstTextLine, 1);
  }
  return kept.join('\n')
    .replace(/^\s*\n+|\n+\s*$/g, '')
    .replace(/\n{3,}/g, '\n\n');
}

function setSystemNote_(range, prefix, message) {
  range.setNote(buildSystemNoteText_(range.getNote(), prefix, message));
}

function buildSystemNoteText_(existingNote, prefix, message) {
  const existing = removeSystemNoteBlockText_(existingNote, prefix);
  const text = toCellText_(message);
  if (!text) return existing;
  const markers = getSystemNoteBlockMarkers_(prefix);
  const block = [
    markers.start,
    text,
    markers.end
  ].join('\n');
  return [existing, block].filter(Boolean).join('\n\n');
}

function clearSystemNote_(range, prefixes) {
  const original = toCellText_(range.getNote());
  const updated = (prefixes || []).reduce((note, prefix) => {
    return removeSystemNoteBlockText_(note, prefix);
  }, original);
  if (updated !== original) {
    if (updated) range.setNote(updated);
    else range.clearNote();
  }
}

function parseFullDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    const year = value.getFullYear();
    if (
      year >= 2000 &&
      year <= 2099 &&
      value.getHours() === 0 &&
      value.getMinutes() === 0 &&
      value.getSeconds() === 0
    ) {
      return new Date(year, value.getMonth(), value.getDate());
    }
    return null;
  }

  const text = toCellText_(value);
  if (!text) return null;
  let weekdayText = '';
  let match = text.match(
    /^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})(?:\s+(Sun|Mon|Tue|Wed|Thu|Fri|Sat))?$/i
  );
  if (match) weekdayText = String(match[4] || '').toLowerCase();
  if (!match) match = text.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day);
  if (
    year < 2000 ||
    year > 2099 ||
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }
  if (weekdayText) {
    const weekdayByName = {
      sun: 0,
      mon: 1,
      tue: 2,
      wed: 3,
      thu: 4,
      fri: 5,
      sat: 6
    };
    if (weekdayByName[weekdayText] !== date.getDay()) return null;
  }
  return date;
}

function isCompletelyBlankRow_(rowValues) {
  return (rowValues || []).every(value => !toCellText_(value));
}

function classifyMonthlyRowValues_(rowValues, columns) {
  const date = parseFullDate_(getRowFieldValue_(rowValues, columns, 'TIME'));
  const eventId = toCellText_(
    getRowFieldValue_(rowValues, columns, 'EVENT_ID')
  );
  if (date) {
    return {
      type: eventId
        ? MONTHLY_ROW_TYPES.HYBRID_CONFLICT
        : MONTHLY_ROW_TYPES.DATE_HEADER,
      date,
      eventId
    };
  }
  return {
    type: isCompletelyBlankRow_(rowValues)
      ? MONTHLY_ROW_TYPES.BLANK
      : MONTHLY_ROW_TYPES.PATIENT,
    date: null,
    eventId
  };
}

function getMonthlyHeaderMarkerState_(rowValues, columns) {
  const value = toCellText_(
    getRowFieldValue_(rowValues, columns, 'CHART_NO')
  );
  return {
    value,
    missing: !value,
    valid: value === MONTHLY_DATE_HEADER_MARKER,
    conflict: Boolean(value && value !== MONTHLY_DATE_HEADER_MARKER)
  };
}

function hasMonthlyPatientContent_(rowValues, columns) {
  if (toCellText_(getRowFieldValue_(rowValues, columns, 'EVENT_ID'))) return true;
  return CONFIG.MONTHLY_FIELD_KEYS.some(key => {
    if (key === 'HOSPITAL' || key === 'EVENT_ID') return false;
    return Boolean(toCellText_(getRowFieldValue_(rowValues, columns, key)));
  });
}

function scanMonthlyBlocksFromValues_(
  sheet,
  columns,
  values,
  startRow,
  lastColumn,
  lastRow
) {
  const cols = columns || getRequiredMonthlyColumns_(sheet);
  const firstRow = Number(startRow) || 2;
  const rows = values || [];
  const resolvedLastColumn = Number(lastColumn) || Math.max(
    getLastHeaderColumn_(sheet),
    ...Object.values(cols)
  );
  const resolvedLastRow = Number(lastRow) || Math.max(
    firstRow - 1,
    firstRow + rows.length - 1
  );
  const blocks = [];
  const orphans = [];
  const hybrids = [];
  const markerIssues = [];
  const duplicateBlocks = [];
  const blockRowsByKey = {};
  const rowValuesByRow = {};
  let currentBlock = null;

  rows.forEach((rowValues, index) => {
    const row = index + firstRow;
    rowValuesByRow[row] = rowValues;
    const classification = classifyMonthlyRowValues_(rowValues, cols);
    if (classification.type === MONTHLY_ROW_TYPES.DATE_HEADER) {
      const date = classification.date;
      const hospital =
        toCellText_(getRowFieldValue_(rowValues, cols, 'HOSPITAL')) ||
        HOSPITAL_KAOH;
      const blockKey = `${formatDateKey_(date)}|${hospital}`;
      const occurrence = (blockRowsByKey[blockKey] || []).length + 1;
      currentBlock = {
        row,
        date,
        dateKey: formatDateKey_(date),
        hospital,
        blockKey,
        occurrence,
        blockRef: {
          headerRow: row,
          dateKey: formatDateKey_(date),
          hospital,
          occurrence
        },
        values: rowValues,
        patientRows: [],
        blankRows: []
      };
      blocks.push(currentBlock);
      if (!blockRowsByKey[blockKey]) blockRowsByKey[blockKey] = [];
      blockRowsByKey[blockKey].push(currentBlock);
      const markerState = getMonthlyHeaderMarkerState_(rowValues, cols);
      if (!markerState.valid) {
        markerIssues.push({
          row,
          type: markerState.conflict
            ? 'monthly_header_marker_conflict'
            : 'monthly_header_marker_missing',
          value: markerState.value,
          blockKey
        });
      }
      return;
    }

    if (classification.type === MONTHLY_ROW_TYPES.HYBRID_CONFLICT) {
      hybrids.push({
        row,
        values: rowValues,
        date: classification.date,
        dateKey: formatDateKey_(classification.date),
        eventId: classification.eventId
      });
      currentBlock = null;
      return;
    }

    if (classification.type === MONTHLY_ROW_TYPES.BLANK) {
      if (currentBlock) currentBlock.blankRows.push(row);
      return;
    }
    // 自訂欄可能是該列唯一有值的欄位；仍須把整列視為資料列，
    // 才不會在月份排序時遺失使用者補充資訊。
    const patient = { row, values: rowValues };
    if (currentBlock) {
      currentBlock.patientRows.push(patient);
    } else {
      orphans.push(patient);
    }
  });

  Object.keys(blockRowsByKey).forEach(blockKey => {
    if (blockRowsByKey[blockKey].length > 1) {
      duplicateBlocks.push({
        blockKey,
        rows: blockRowsByKey[blockKey].map(block => block.row),
        blocks: blockRowsByKey[blockKey]
      });
    }
  });

  return {
    columns: cols,
    lastColumn: resolvedLastColumn,
    lastRow: resolvedLastRow,
    firstRow,
    values: rows,
    blocks,
    orphans,
    hybrids,
    markerIssues,
    duplicateBlocks,
    rowValuesByRow
  };
}

function scanMonthlyBlocks_(sheet, columns) {
  const cols = columns || getRequiredMonthlyColumns_(sheet);
  const lastColumn = Math.max(getLastHeaderColumn_(sheet), ...Object.values(cols));
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues()
    : [];
  return scanMonthlyBlocksFromValues_(
    sheet,
    cols,
    values,
    2,
    lastColumn,
    lastRow
  );
}

function getMonthlyStructureIssues_(sheet, scan) {
  const result = scan || scanMonthlyBlocks_(
    sheet,
    getRequiredMonthlyColumns_(sheet)
  );
  const issues = [];
  result.hybrids.forEach(item => {
    issues.push({
      severity: 'conflict',
      type: 'monthly_hybrid_date_event_id',
      row: item.row,
      eventId: item.eventId,
      message:
        '第一欄是完整日期但該列已有 CalendarEventId；未將它當作日期標題或病人列。'
    });
  });
  result.orphans.forEach(item => {
    if (!hasMonthlyPatientContent_(item.values, result.columns)) return;
    issues.push({
      severity: 'conflict',
      type: 'monthly_orphan_patient',
      row: item.row,
      eventId: toCellText_(
        getRowFieldValue_(item.values, result.columns, 'EVENT_ID')
      ),
      message: '病人資料列缺少上方有效日期標題。'
    });
  });
  result.markerIssues.forEach(item => {
    issues.push({
      severity: item.type === 'monthly_header_marker_missing'
        ? 'action'
        : 'conflict',
      type: item.type,
      row: item.row,
      eventId: '',
      message: item.type === 'monthly_header_marker_missing'
        ? '日期標題尚未顯示 ◆ 刀日，等待安裝／修復或編輯時自動補齊。'
        : '日期標題的病歷號格含既有內容；系統未覆寫。'
    });
  });
  Object.keys(result.rowValuesByRow).forEach(rowKey => {
    const row = Number(rowKey);
    const values = result.rowValuesByRow[row];
    const classification = classifyMonthlyRowValues_(
      values,
      result.columns
    );
    if (
      !classification.date &&
      toCellText_(
        getRowFieldValue_(values, result.columns, 'CHART_NO')
      ) === MONTHLY_DATE_HEADER_MARKER
    ) {
      issues.push({
        severity: 'action',
        type: 'monthly_stale_header_marker',
        row,
        eventId: '',
        message: '該列已不是完整日期標題，等待編輯或安裝／修復時清除舊的 ◆ 刀日標記。'
      });
    }
  });
  result.blocks.forEach(block => {
    if (formatMonthKey_(block.date) === sheet.getName()) return;
    issues.push({
      severity: 'conflict',
      type: 'monthly_cross_month_header',
      row: block.row,
      eventId: '',
      message: `日期 ${block.dateKey} 不屬於 ${sheet.getName()}。`
    });
  });
  return issues;
}

function getDefaultMonthlyDates_(monthName) {
  if (!/^\d{6}$/.test(String(monthName || ''))) return [];
  const year = Number(String(monthName).slice(0, 4));
  const month = Number(String(monthName).slice(4, 6));
  if (month < 1 || month > 12) return [];
  const weekdayOccurrences = { 1: [], 4: [] };
  const cursor = new Date(year, month - 1, 1);
  while (cursor.getMonth() === month - 1) {
    const weekday = cursor.getDay();
    if (weekdayOccurrences[weekday]) {
      weekdayOccurrences[weekday].push(new Date(cursor.getTime()));
    }
    cursor.setDate(cursor.getDate() + 1);
  }
  return [
    weekdayOccurrences[1][1],
    weekdayOccurrences[4][1],
    weekdayOccurrences[1][3],
    weekdayOccurrences[4][3]
  ].filter(Boolean).sort((left, right) => left.getTime() - right.getTime());
}

function applyColumnWidths_(sheet, columns, widthMap) {
  Object.keys(widthMap).forEach(fieldKey => {
    if (columns[fieldKey]) {
      sheet.setColumnWidth(columns[fieldKey], widthMap[fieldKey]);
    }
  });
  const headers = getSheetHeaderValues_(sheet);
  const lastHeaderColumn = getLastHeaderColumn_(sheet);
  const internalBlankColumns = [];
  for (let column = 1; column <= lastHeaderColumn; column++) {
    if (toCellText_(headers[column - 1])) continue;
    const rowCount = Math.max(0, sheet.getLastRow() - 1);
    const hasData = rowCount
      ? sheet.getRange(2, column, rowCount, 1)
        .getDisplayValues()
        .some(row => toCellText_(row[0]))
      : false;
    if (!hasData) internalBlankColumns.push(column);
  }
  if (internalBlankColumns.length === 1) {
    sheet.setColumnWidth(internalBlankColumns[0], 18);
  }
  return { applied: true };
}

function buildListValidation_(values, allowInvalid) {
  return SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(Boolean(allowInvalid))
    .build();
}

function clearMonthlyHeaderPatientValidations_(sheet, row, columns) {
  ['GA', 'SIDE', 'DIAGNOSIS', 'PROCEDURE'].forEach(key => {
    if (columns[key]) {
      sheet.getRange(row, columns[key]).clearDataValidations();
    }
  });
}

function needsMonthlyHeaderPresentation_(rowValues, columns) {
  const classification = classifyMonthlyRowValues_(rowValues, columns);
  if (
    classification.type === MONTHLY_ROW_TYPES.DATE_HEADER ||
    classification.type === MONTHLY_ROW_TYPES.HYBRID_CONFLICT
  ) {
    return true;
  }
  return toCellText_(
    getRowFieldValue_(rowValues, columns, 'CHART_NO')
  ) === MONTHLY_DATE_HEADER_MARKER;
}

function ensureMonthlyHeaderPresentation_(sheet, row, columns, rowValues) {
  const lastColumn = Math.max(
    getLastHeaderColumn_(sheet),
    ...Object.values(columns)
  );
  const values = rowValues || sheet.getRange(row, 1, 1, lastColumn)
    .getValues()[0];
  const classification = classifyMonthlyRowValues_(values, columns);
  const markerCell = sheet.getRange(row, columns.CHART_NO);
  if (classification.type !== MONTHLY_ROW_TYPES.DATE_HEADER) {
    const markerValue = toCellText_(markerCell.getValue());
    if (!classification.date && markerValue === MONTHLY_DATE_HEADER_MARKER) {
      markerCell.clearContent();
    }
    if (!classification.date) {
      markerCell.clearDataValidations();
      clearSystemNote_(markerCell, [MONTHLY_STRUCTURE_NOTE_PREFIX]);
    }
    return {
      ok: classification.type !== MONTHLY_ROW_TYPES.HYBRID_CONFLICT,
      type: classification.type,
      conflict: classification.type === MONTHLY_ROW_TYPES.HYBRID_CONFLICT
    };
  }

  const dateCell = sheet.getRange(row, columns.TIME);
  const dateFormula = typeof dateCell.getFormula === 'function'
    ? toCellText_(dateCell.getFormula())
    : '';
  if (
    !dateFormula &&
    !(
      dateCell.getValue() instanceof Date &&
      !isNaN(dateCell.getValue().getTime())
    )
  ) {
    dateCell.setValue(classification.date);
  }
  dateCell.setNumberFormat('yyyy/m/d ddd');
  const hospitalCell = sheet.getRange(row, columns.HOSPITAL);
  const hospitalFormula = typeof hospitalCell.getFormula === 'function'
    ? toCellText_(hospitalCell.getFormula())
    : '';
  if (!toCellText_(hospitalCell.getValue()) && !hospitalFormula) {
    hospitalCell.setValue(HOSPITAL_KAOH);
  }
  hospitalCell.setDataValidation(
    buildListValidation_(CONFIG.HOSPITAL_OPTIONS, false)
  );
  clearMonthlyHeaderPatientValidations_(sheet, row, columns);

  const markerState = getMonthlyHeaderMarkerState_(values, columns);
  const markerFormula = typeof markerCell.getFormula === 'function'
    ? toCellText_(markerCell.getFormula())
    : '';
  if (markerState.conflict || markerFormula) {
    setSystemNote_(
      markerCell,
      MONTHLY_STRUCTURE_NOTE_PREFIX,
      '完整日期列的「病歷號」格保留給 ◆ 刀日；既有內容或公式未被覆寫，請移至第四欄以後。'
    );
    return {
      ok: false,
      type: classification.type,
      conflict: true,
      markerValue: markerState.value
    };
  }
  markerCell
    .setValue(MONTHLY_DATE_HEADER_MARKER)
    .setDataValidation(buildListValidation_([MONTHLY_DATE_HEADER_MARKER], false));
  clearSystemNote_(markerCell, [MONTHLY_STRUCTURE_NOTE_PREFIX]);
  return {
    ok: true,
    type: classification.type,
    conflict: false,
    markerValue: MONTHLY_DATE_HEADER_MARKER
  };
}

function applyColumnRangeFormat_(sheet, column, options) {
  if (!column) return;
  const range = sheet.getRange(1, column, sheet.getMaxRows(), 1);
  range
    .setFontFamily(MANAGED_FONT_FAMILY)
    .setFontSize(MANAGED_FONT_SIZE)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top');
  if (options && options.wrap) range.setWrap(true);
}

function applyMonthlyTextNumberFormats_(
  sheet,
  columns,
  startRow,
  rowCount
) {
  if (!rowCount || rowCount < 1) return { applied: false, ranges: [] };
  const endRow = startRow + rowCount - 1;
  const addresses = MONTHLY_TEXT_KEYS
    .map(key => columns[key])
    .filter(Boolean)
    .map(column => {
      const letter = columnToLetter_(column);
      return rowCount === 1
        ? `${letter}${startRow}`
        : `${letter}${startRow}:${letter}${endRow}`;
    });
  if (!addresses.length) return { applied: false, ranges: [] };
  sheet.getRangeList(addresses).setNumberFormat('@');
  return { applied: true, ranges: addresses };
}

function applyHeaderFormat_(sheet, lastColumn) {
  sheet.getRange(1, 1, 1, Math.max(1, lastColumn))
    .setBackground(TABLE_HEADER_BACKGROUND)
    .setFontColor('#000000')
    .setFontFamily(MANAGED_FONT_FAMILY)
    .setFontSize(MANAGED_FONT_SIZE)
    .setFontWeight('bold')
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top')
    .setWrap(true);
}

function removeSystemConditionalFormats_(rules) {
  return (rules || []).filter(rule => {
    try {
      const formula = rule.getBooleanCondition() &&
        rule.getBooleanCondition().getCriteriaValues()[0];
      const text = String(formula || '');
      return (
        text.indexOf(CONDITIONAL_FORMAT_MARKER_PREFIX) === -1 &&
        text.indexOf(LEGACY_CONDITIONAL_FORMAT_MARKER_PREFIX) === -1
      );
    } catch (err) {
      return true;
    }
  });
}

function buildVisibleRangesExcludingColumns_(sheet, startRow, endRow, lastColumn, excluded) {
  const excludedMap = {};
  (excluded || []).forEach(column => {
    if (column) excludedMap[column] = true;
  });
  const ranges = [];
  let segmentStart = 1;
  for (let column = 1; column <= lastColumn + 1; column++) {
    if (column <= lastColumn && !excludedMap[column]) continue;
    if (segmentStart < column) {
      ranges.push(
        sheet.getRange(
          startRow,
          segmentStart,
          endRow - startRow + 1,
          column - segmentStart
        )
      );
    }
    segmentStart = column + 1;
  }
  return ranges;
}

function buildPlanContainsAnyFormula_(cellReference, keywords) {
  const clauses = (keywords || [])
    .map(keyword => String(keyword || '').replace(/"/g, '""'))
    .filter(Boolean)
    .map(keyword => `ISNUMBER(SEARCH("${keyword}",${cellReference}))`);
  if (!clauses.length) return 'FALSE';
  return clauses.length === 1 ? clauses[0] : `OR(${clauses.join(',')})`;
}

function applyFuConditionalFormats_(sheet, columns, lastColumn) {
  const maxRows = Math.max(sheet.getMaxRows(), 2);
  const dataRows = buildVisibleRangesExcludingColumns_(
    sheet,
    2,
    maxRows,
    lastColumn,
    [columns.EVENT_ID]
  );
  const planLetter = columnToLetter_(columns.PLAN);
  const planReference = `$${planLetter}2`;
  const redMatch = buildPlanContainsAnyFormula_(
    planReference,
    CONFIG.PLAN_RED_KEYWORDS
  );
  const greenMatch = buildPlanContainsAnyFormula_(
    planReference,
    CONFIG.PLAN_GREEN_KEYWORDS
  );
  const yellowMatch = buildPlanContainsAnyFormula_(
    planReference,
    CONFIG.PLAN_YELLOW_KEYWORDS
  );
  const redFormula =
    `=AND(${redMatch},` +
    `N("${CONDITIONAL_FORMAT_MARKER_PREFIX}FU_PLAN_RED")=0)`;
  const greenFormula =
    `=AND(NOT(${redMatch}),${greenMatch},` +
    `N("${CONDITIONAL_FORMAT_MARKER_PREFIX}FU_PLAN_GREEN")=0)`;
  const yellowFormula =
    `=AND(NOT(${redMatch}),NOT(${greenMatch}),${yellowMatch},` +
    `N("${CONDITIONAL_FORMAT_MARKER_PREFIX}FU_PLAN_YELLOW")=0)`;
  const systemRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(redFormula)
      .setBackground(PLAN_RED_BACKGROUND)
      .setRanges(dataRows)
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(greenFormula)
      .setBackground(PLAN_GREEN_BACKGROUND)
      .setRanges(dataRows)
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(yellowFormula)
      .setBackground(PLAN_YELLOW_BACKGROUND)
      .setRanges(dataRows)
      .build()
  ];
  sheet.setConditionalFormatRules(
    systemRules.concat(removeSystemConditionalFormats_(sheet.getConditionalFormatRules()))
  );
}

function applyMonthlyConditionalFormats_(sheet, columns, lastColumn) {
  const maxRows = Math.max(sheet.getMaxRows(), 2);
  const fullRanges = buildVisibleRangesExcludingColumns_(
    sheet,
    2,
    maxRows,
    lastColumn,
    [columns.EVENT_ID]
  );
  const planRanges = buildVisibleRangesExcludingColumns_(
    sheet,
    2,
    maxRows,
    lastColumn,
    [columns.GA, columns.SIDE, columns.EVENT_ID]
  );
  const timeLetter = columnToLetter_(columns.TIME);
  const eventIdLetter = columnToLetter_(columns.EVENT_ID);
  const planLetter = columnToLetter_(columns.PLAN);
  const gaLetter = columnToLetter_(columns.GA);
  const sideLetter = columnToLetter_(columns.SIDE);
  const marker = name => `${CONDITIONAL_FORMAT_MARKER_PREFIX}${name}`;
  const planReference = `$${planLetter}2`;
  const redMatch = buildPlanContainsAnyFormula_(
    planReference,
    CONFIG.PLAN_RED_KEYWORDS
  );
  const greenMatch = buildPlanContainsAnyFormula_(
    planReference,
    CONFIG.PLAN_GREEN_KEYWORDS
  );
  const yellowMatch = buildPlanContainsAnyFormula_(
    planReference,
    CONFIG.PLAN_YELLOW_KEYWORDS
  );
  const headerPredicate =
    `AND(ISNUMBER($${timeLetter}2),YEAR($${timeLetter}2)>=2000,` +
    `YEAR($${timeLetter}2)<=2099,MOD($${timeLetter}2,1)=0,` +
    `$${eventIdLetter}2="")`;
  const headerFormula =
    `=AND(${headerPredicate},` +
    `N("${marker('MONTH_HEADER')}")=0)`;
  const redFormula =
    `=AND(NOT(${headerPredicate}),${redMatch},` +
    `N("${marker('MONTH_PLAN_RED')}")=0)`;
  const greenFormula =
    `=AND(NOT(${headerPredicate}),` +
    `NOT(${redMatch}),${greenMatch},` +
    `N("${marker('MONTH_PLAN_GREEN')}")=0)`;
  const yellowFormula =
    `=AND(NOT(${headerPredicate}),` +
    `NOT(${redMatch}),NOT(${greenMatch}),${yellowMatch},` +
    `N("${marker('MONTH_PLAN_YELLOW')}")=0)`;
  const systemRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(headerFormula)
      .setBackground(MONTHLY_DATE_HEADER_BACKGROUND)
      .setBold(true)
      .setRanges(fullRanges)
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        `=AND(NOT(${headerPredicate}),$${gaLetter}2="GA",` +
        `N("${marker('GA')}")=0)`
      )
      .setBackground(MONTHLY_GA_BACKGROUND)
      .setRanges([sheet.getRange(2, columns.GA, maxRows - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        `=AND(NOT(${headerPredicate}),$${sideLetter}2="OD",` +
        `N("${marker('OD')}")=0)`
      )
      .setBackground(SIDE_OD_BACKGROUND)
      .setRanges([sheet.getRange(2, columns.SIDE, maxRows - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(
        `=AND(NOT(${headerPredicate}),$${sideLetter}2="OU",` +
        `N("${marker('OU')}")=0)`
      )
      .setBackground(SIDE_OU_BACKGROUND)
      .setRanges([sheet.getRange(2, columns.SIDE, maxRows - 1, 1)])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(redFormula)
      .setBackground(PLAN_RED_BACKGROUND)
      .setRanges(planRanges)
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(greenFormula)
      .setBackground(PLAN_GREEN_BACKGROUND)
      .setRanges(planRanges)
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(yellowFormula)
      .setBackground(PLAN_YELLOW_BACKGROUND)
      .setRanges(planRanges)
      .build()
  ];
  sheet.setConditionalFormatRules(
    systemRules.concat(removeSystemConditionalFormats_(sheet.getConditionalFormatRules()))
  );
}

function applyFuFormatting_(sheet, options) {
  const settings = options || {};
  const columns = ensureManagedHeaders_(sheet, 'FU');
  const lastColumn = Math.max(getLastHeaderColumn_(sheet), ...Object.values(columns));
  applyHeaderFormat_(sheet, lastColumn);
  CONFIG.FIELD_KEYS.forEach(key => {
    applyColumnRangeFormat_(sheet, columns[key], {
      wrap: FU_WRAP_KEYS.indexOf(key) !== -1
    });
    if (sheet.getMaxRows() > 1) {
      sheet.getRange(2, columns[key], sheet.getMaxRows() - 1, 1)
        .setBackground('#FFFFFF');
    }
  });
  sheet.getRange(2, columns.DATE, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setNumberFormat('yyyy/m/d');
  sheet.getRange(2, columns.TAG, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setDataValidation(buildListValidation_(CONFIG.FU_TAG_OPTIONS, true));
  sheet.getRange(2, columns.HOSPITAL, Math.max(1, sheet.getMaxRows() - 1), 1)
    .setDataValidation(buildListValidation_(CONFIG.HOSPITAL_OPTIONS, true));
  applyFuConditionalFormats_(sheet, columns, lastColumn);
  const widths = settings.forceWidths
    ? applyColumnWidths_(sheet, columns, FU_COLUMN_WIDTHS)
    : { applied: false };
  return { columns, lastColumn, widths };
}

function applyMonthlyFormatting_(sheet, options) {
  const settings = options || {};
  const columns = ensureManagedHeaders_(sheet, 'MONTHLY');
  const lastColumn = Math.max(getLastHeaderColumn_(sheet), ...Object.values(columns));
  applyHeaderFormat_(sheet, lastColumn);
  CONFIG.MONTHLY_FIELD_KEYS.forEach(key => {
    applyColumnRangeFormat_(sheet, columns[key], {
      wrap: MONTHLY_WRAP_KEYS.indexOf(key) !== -1
    });
    if (sheet.getMaxRows() > 1) {
      sheet.getRange(2, columns[key], sheet.getMaxRows() - 1, 1)
        .setBackground('#FFFFFF');
    }
  });
  const dataRowCount = Math.max(1, sheet.getMaxRows() - 1);
  sheet.getRange(2, columns.TIME, dataRowCount, 1).setNumberFormat('@');
  applyMonthlyTextNumberFormats_(sheet, columns, 2, dataRowCount);
  sheet.getRange(2, columns.HOSPITAL, dataRowCount, 1)
    .setDataValidation(buildListValidation_(CONFIG.HOSPITAL_OPTIONS, false));
  sheet.getRange(2, columns.GA, dataRowCount, 1)
    .setDataValidation(buildListValidation_(CONFIG.MONTHLY_GA_OPTIONS, false));
  sheet.getRange(2, columns.SIDE, dataRowCount, 1)
    .setDataValidation(buildListValidation_(CONFIG.MONTHLY_SIDE_OPTIONS, false));
  sheet.getRange(2, columns.DIAGNOSIS, dataRowCount, 1)
    .setDataValidation(buildListValidation_(CONFIG.MONTHLY_DIAGNOSIS_OPTIONS, true));
  sheet.getRange(2, columns.PROCEDURE, dataRowCount, 1)
    .setDataValidation(buildListValidation_(CONFIG.MONTHLY_PROCEDURE_OPTIONS, true));
  applyMonthlyConditionalFormats_(sheet, columns, lastColumn);
  const widths = settings.forceWidths
    ? applyColumnWidths_(sheet, columns, MONTHLY_COLUMN_WIDTHS)
    : { applied: false };

  const scan = scanMonthlyBlocks_(sheet, columns);
  scan.blocks.forEach(block => {
    sheet.getRange(block.row, 1, 1, lastColumn)
      .setFontFamily(MANAGED_FONT_FAMILY)
      .setFontSize(MANAGED_FONT_SIZE)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
    ensureMonthlyHeaderPresentation_(
      sheet,
      block.row,
      columns,
      block.values
    );
  });
  return { columns, lastColumn, widths };
}

function applyManagedRowFormat_(sheet, row, kind, columns) {
  if (row < 2) return;
  const isFu = kind === 'FU';
  const cols = columns || (
    isFu ? getRequiredFuColumns_(sheet) : getRequiredMonthlyColumns_(sheet)
  );
  const keys = isFu ? CONFIG.FIELD_KEYS : CONFIG.MONTHLY_FIELD_KEYS;
  const wrapKeys = isFu ? FU_WRAP_KEYS : MONTHLY_WRAP_KEYS;
  const managedAddresses = keys
    .map(key => cols[key])
    .filter(Boolean)
    .map(column => `${columnToLetter_(column)}${row}`);
  if (managedAddresses.length) {
    sheet.getRangeList(managedAddresses)
      .setFontFamily(MANAGED_FONT_FAMILY)
      .setFontSize(MANAGED_FONT_SIZE)
      .setFontWeight('normal')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
  }
  const wrapAddresses = wrapKeys
    .map(key => cols[key])
    .filter(Boolean)
    .map(column => `${columnToLetter_(column)}${row}`);
  if (wrapAddresses.length) {
    sheet.getRangeList(wrapAddresses).setWrap(true);
  }

  if (isFu) {
    sheet.getRange(row, cols.TAG)
      .setDataValidation(buildListValidation_(CONFIG.FU_TAG_OPTIONS, true));
  } else {
    sheet.getRange(row, cols.HOSPITAL)
      .setDataValidation(buildListValidation_(CONFIG.HOSPITAL_OPTIONS, false));
    sheet.getRange(row, cols.GA)
      .setDataValidation(buildListValidation_(CONFIG.MONTHLY_GA_OPTIONS, false));
    sheet.getRange(row, cols.SIDE)
      .setDataValidation(buildListValidation_(CONFIG.MONTHLY_SIDE_OPTIONS, false));
    sheet.getRange(row, cols.DIAGNOSIS)
      .setDataValidation(buildListValidation_(CONFIG.MONTHLY_DIAGNOSIS_OPTIONS, true));
    sheet.getRange(row, cols.PROCEDURE)
      .setDataValidation(buildListValidation_(CONFIG.MONTHLY_PROCEDURE_OPTIONS, true));
    sheet.getRange(row, cols.TIME).setNumberFormat('@');
    applyMonthlyTextNumberFormats_(sheet, cols, row, 1);
  }
}

function compactRowSegments_(rows) {
  const sorted = Array.from(new Set(
    (rows || []).map(Number).filter(row => row >= 2)
  )).sort((left, right) => left - right);
  const segments = [];
  sorted.forEach(row => {
    const current = segments[segments.length - 1];
    if (current && current.start + current.count === row) {
      current.count++;
    } else {
      segments.push({ start: row, count: 1 });
    }
  });
  return segments;
}

function buildColumnSegmentAddresses_(column, segments) {
  const letter = columnToLetter_(column);
  return (segments || []).map(segment => {
    const end = segment.start + segment.count - 1;
    return segment.count === 1
      ? `${letter}${segment.start}`
      : `${letter}${segment.start}:${letter}${end}`;
  });
}

function applyManagedRowsFormatBatch_(sheet, rows, kind, columns) {
  const segments = compactRowSegments_(rows);
  if (!segments.length) return { applied: false, rowCount: 0 };
  const isFu = kind === 'FU';
  const cols = columns || (
    isFu ? getRequiredFuColumns_(sheet) : getRequiredMonthlyColumns_(sheet)
  );
  const keys = isFu ? CONFIG.FIELD_KEYS : CONFIG.MONTHLY_FIELD_KEYS;
  const wrapKeys = isFu ? FU_WRAP_KEYS : MONTHLY_WRAP_KEYS;
  const managedAddresses = [];
  keys.forEach(key => {
    if (!cols[key]) return;
    buildColumnSegmentAddresses_(cols[key], segments)
      .forEach(address => managedAddresses.push(address));
  });
  if (managedAddresses.length) {
    sheet.getRangeList(managedAddresses)
      .setFontFamily(MANAGED_FONT_FAMILY)
      .setFontSize(MANAGED_FONT_SIZE)
      .setFontWeight('normal')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
  }
  const wrapAddresses = [];
  wrapKeys.forEach(key => {
    if (!cols[key]) return;
    buildColumnSegmentAddresses_(cols[key], segments)
      .forEach(address => wrapAddresses.push(address));
  });
  if (wrapAddresses.length) {
    sheet.getRangeList(wrapAddresses).setWrap(true);
  }

  if (isFu) {
    segments.forEach(segment => {
      sheet.getRange(segment.start, cols.TAG, segment.count, 1)
        .setDataValidation(buildListValidation_(CONFIG.FU_TAG_OPTIONS, true));
    });
  } else {
    const validations = [
      ['HOSPITAL', CONFIG.HOSPITAL_OPTIONS, false],
      ['GA', CONFIG.MONTHLY_GA_OPTIONS, false],
      ['SIDE', CONFIG.MONTHLY_SIDE_OPTIONS, false],
      ['DIAGNOSIS', CONFIG.MONTHLY_DIAGNOSIS_OPTIONS, true],
      ['PROCEDURE', CONFIG.MONTHLY_PROCEDURE_OPTIONS, true]
    ];
    segments.forEach(segment => {
      validations.forEach(([key, options, allowInvalid]) => {
        sheet.getRange(segment.start, cols[key], segment.count, 1)
          .setDataValidation(buildListValidation_(options, allowInvalid));
      });
    });
    const textAddresses = buildColumnSegmentAddresses_(
      cols.TIME,
      segments
    );
    MONTHLY_TEXT_KEYS.forEach(key => {
      buildColumnSegmentAddresses_(cols[key], segments)
        .forEach(address => textAddresses.push(address));
    });
    if (textAddresses.length) {
      sheet.getRangeList(textAddresses).setNumberFormat('@');
    }
  }
  return {
    applied: true,
    rowCount: segments.reduce((sum, segment) => sum + segment.count, 0),
    segmentCount: segments.length
  };
}

function getTouchedManagedKeys_(range, columns, keys) {
  const start = range.getColumn();
  const end = start + range.getNumColumns() - 1;
  return (keys || []).filter(key => {
    const column = columns[key];
    return column && column >= start && column <= end;
  });
}

function applyTouchedManagedRowsFormatBatch_(
  sheet,
  rows,
  kind,
  columns,
  touchedKeys
) {
  const segments = compactRowSegments_(rows);
  const keys = Array.from(new Set(touchedKeys || []));
  if (!segments.length || !keys.length) {
    return { applied: false, rowCount: 0, keys: [] };
  }
  const isFu = kind === 'FU';
  const wrapKeys = isFu ? FU_WRAP_KEYS : MONTHLY_WRAP_KEYS;
  const managedAddresses = [];
  keys.forEach(key => {
    if (!columns[key]) return;
    buildColumnSegmentAddresses_(columns[key], segments)
      .forEach(address => managedAddresses.push(address));
  });
  if (managedAddresses.length) {
    sheet.getRangeList(managedAddresses)
      .setFontFamily(MANAGED_FONT_FAMILY)
      .setFontSize(MANAGED_FONT_SIZE)
      .setFontWeight('normal')
      .setHorizontalAlignment('left')
      .setVerticalAlignment('top');
  }
  const wrapAddresses = [];
  keys.filter(key => wrapKeys.indexOf(key) !== -1).forEach(key => {
    buildColumnSegmentAddresses_(columns[key], segments)
      .forEach(address => wrapAddresses.push(address));
  });
  if (wrapAddresses.length) {
    sheet.getRangeList(wrapAddresses).setWrap(true);
  }

  const applyValidation = (key, options, allowInvalid) => {
    if (keys.indexOf(key) === -1 || !columns[key]) return;
    segments.forEach(segment => {
      sheet.getRange(segment.start, columns[key], segment.count, 1)
        .setDataValidation(buildListValidation_(options, allowInvalid));
    });
  };
  if (isFu) {
    applyValidation('TAG', CONFIG.FU_TAG_OPTIONS, true);
    applyValidation('HOSPITAL', CONFIG.HOSPITAL_OPTIONS, true);
    if (keys.indexOf('DATE') !== -1) {
      sheet.getRangeList(
        buildColumnSegmentAddresses_(columns.DATE, segments)
      ).setNumberFormat('yyyy/m/d');
    }
  } else {
    applyValidation('HOSPITAL', CONFIG.HOSPITAL_OPTIONS, false);
    applyValidation('GA', CONFIG.MONTHLY_GA_OPTIONS, false);
    applyValidation('SIDE', CONFIG.MONTHLY_SIDE_OPTIONS, false);
    applyValidation(
      'DIAGNOSIS',
      CONFIG.MONTHLY_DIAGNOSIS_OPTIONS,
      true
    );
    applyValidation(
      'PROCEDURE',
      CONFIG.MONTHLY_PROCEDURE_OPTIONS,
      true
    );
    const textKeys = ['TIME'].concat(MONTHLY_TEXT_KEYS)
      .filter(key => keys.indexOf(key) !== -1 && columns[key]);
    const textAddresses = [];
    textKeys.forEach(key => {
      buildColumnSegmentAddresses_(columns[key], segments)
        .forEach(address => textAddresses.push(address));
    });
    if (textAddresses.length) {
      sheet.getRangeList(textAddresses).setNumberFormat('@');
    }
  }
  return {
    applied: true,
    rowCount: segments.reduce((sum, segment) => sum + segment.count, 0),
    keys
  };
}

function initializeMainTrackingSheet_(spreadsheet, showAlert) {
  assertNoLegacyOpOnly_(spreadsheet);
  let sheet = getMainTrackingSheet_(spreadsheet);
  const created = !sheet;
  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONFIG.SHEET_FU);
  }
  const result = applyFuFormatting_(sheet, { forceWidths: created });
  const normalization = { normalized: 0, errors: 0 };
  const message = `${sheet.getName()} 已初始化；FU 事件使用全天模式。`;
  if (showAlert) SpreadsheetApp.getUi().alert(message);
  return {
    ok: true,
    sheet,
    created,
    columns: result.columns,
    normalization,
    message
  };
}

function applyAllRecommendedColumnWidths() {
  return runMenuAction_('套用所有 FU／月表建議欄寬', () => {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      '套用建議欄寬',
      '這會重設 FU 與所有未封存月份表的標準欄寬；' +
        '自訂欄不會變更。是否繼續？',
      ui.ButtonSet.YES_NO
    );
    if (response !== ui.Button.YES) {
      return { ok: false, cancelled: true };
    }
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const applied = [];
    const fu = getMainTrackingSheet_(spreadsheet);
    if (fu) {
      applyColumnWidths_(
        fu,
        getRequiredFuColumns_(fu),
        FU_COLUMN_WIDTHS
      );
      applied.push(fu.getName());
    }
    getActiveMonthlySheets_(spreadsheet).forEach(sheet => {
      applyColumnWidths_(
        sheet,
        getRequiredMonthlyColumns_(sheet),
        MONTHLY_COLUMN_WIDTHS
      );
      applied.push(sheet.getName());
    });
    ui.alert(
      applied.length
        ? `已套用 ${applied.length} 張表的建議欄寬。`
        : '沒有可套用的 FU 或月份刀表。'
    );
    return { ok: true, appliedSheets: applied };
  });
}

function initializeAllMonthlySheets_(spreadsheet, showAlert, options) {
  const settings = options || {};
  const normalizeTimes = settings.normalizeTimes !== false;
  const sheets = getActiveMonthlySheets_(spreadsheet);
  let normalized = 0;
  let errors = 0;
  sheets.forEach(sheet => {
    const result = applyMonthlyFormatting_(sheet, { forceWidths: false });
    if (normalizeTimes) {
      const timeResult = normalizeExistingTimeColumn_(
        sheet,
        'MONTHLY',
        result.columns
      );
      normalized += timeResult.normalized;
      errors += timeResult.errors;
    }
  });
  const message = normalizeTimes
    ? (
      `${sheets.length} 張月份刀表已初始化；時間正規化 ${normalized} 格，` +
      `錯誤 ${errors} 格。`
    )
    : `${sheets.length} 張月份刀表已初始化；既有報到時間未批次改寫。`;
  if (showAlert) SpreadsheetApp.getUi().alert(message);
  return { ok: true, count: sheets.length, normalized, errors, message };
}

function normalizeTimeCell_(sheet, row, column, kind, columns) {
  const cell = sheet.getRange(row, column);
  const value = cell.getValue();
  if (
    kind === 'MONTHLY' &&
    column === columns.TIME
  ) {
    const rowValues = sheet.getRange(
      row,
      1,
      1,
      Math.max(getLastHeaderColumn_(sheet), ...Object.values(columns))
    ).getValues()[0];
    const classification = classifyMonthlyRowValues_(rowValues, columns);
    if (classification.type === MONTHLY_ROW_TYPES.DATE_HEADER) {
      const presentation = ensureMonthlyHeaderPresentation_(
        sheet,
        row,
        columns,
        rowValues
      );
      clearSystemNote_(cell, [TIME_ERROR_NOTE_PREFIX]);
      return {
        header: true,
        normalized: false,
        error: false,
        conflict: !presentation.ok
      };
    }
    if (classification.type === MONTHLY_ROW_TYPES.HYBRID_CONFLICT) {
      clearSystemNote_(cell, [TIME_ERROR_NOTE_PREFIX]);
      setSystemNote_(
        sheet.getRange(row, columns.CHART_NO),
        MONTHLY_STRUCTURE_NOTE_PREFIX,
        '第一欄是完整日期但該列已有 CalendarEventId；系統未將它當作日期標題或報到時間。'
      );
      return {
        header: false,
        normalized: false,
        error: false,
        conflict: true
      };
    }
    ensureMonthlyHeaderPresentation_(sheet, row, columns, rowValues);
  }

  const info = resolveCalendarTime_(value);
  if (info.errorMessage) {
    setSystemNote_(cell, TIME_ERROR_NOTE_PREFIX, info.errorMessage);
    return { header: false, normalized: false, error: true };
  }
  clearSystemNote_(cell, [TIME_ERROR_NOTE_PREFIX]);
  if (info.parsedTime && toCellText_(value) !== info.parsedTime.text) {
    cell.setNumberFormat('@').setValue(info.parsedTime.text);
    return { header: false, normalized: true, error: false };
  }
  return { header: false, normalized: false, error: false };
}

function normalizeExistingTimeColumn_(sheet, kind, columns) {
  const lastRow = sheet.getLastRow();
  const timeColumn = columns.TIME;
  if (!timeColumn) return { normalized: 0, errors: 0 };
  let normalized = 0;
  let errors = 0;
  for (let row = 2; row <= lastRow; row++) {
    const result = normalizeTimeCell_(sheet, row, timeColumn, kind, columns);
    if (result.normalized) normalized++;
    if (result.error) errors++;
  }
  return { normalized, errors };
}

function normalizeEditedTimes_(event, sheet, kind, columns) {
  const range = event.range;
  if (!columns.TIME) {
    return { normalized: 0, errors: 0, headers: [] };
  }
  const startColumn = range.getColumn();
  const endColumn = startColumn + range.getNumColumns() - 1;
  if (columns.TIME < startColumn || columns.TIME > endColumn) {
    return { normalized: 0, errors: 0, headers: [] };
  }
  let normalized = 0;
  let errors = 0;
  const headers = [];
  const firstRow = Math.max(2, range.getRow());
  const lastRow = range.getRow() + range.getNumRows() - 1;
  const rowCount = Math.max(0, lastRow - firstRow + 1);
  if (!rowCount) return { normalized, errors, headers };
  const lastColumn = Math.max(
    getLastHeaderColumn_(sheet),
    ...Object.values(columns)
  );
  const rowValues = sheet.getRange(
    firstRow,
    1,
    rowCount,
    lastColumn
  ).getValues();
  const timeRange = sheet.getRange(
    firstRow,
    columns.TIME,
    rowCount,
    1
  );
  const originalNotes = timeRange.getNotes();
  const updatedNotes = originalNotes.map(row => row.slice());
  let notesChanged = false;
  const normalizedByRow = {};

  rowValues.forEach((values, index) => {
    const row = firstRow + index;
    const value = getRowFieldValue_(values, columns, 'TIME');
    if (kind === 'MONTHLY') {
      const classification = classifyMonthlyRowValues_(values, columns);
      if (classification.type === MONTHLY_ROW_TYPES.DATE_HEADER) {
        ensureMonthlyHeaderPresentation_(sheet, row, columns, values);
        headers.push(row);
        const cleared = buildSystemNoteText_(
          originalNotes[index][0],
          TIME_ERROR_NOTE_PREFIX,
          ''
        );
        if (cleared !== originalNotes[index][0]) {
          updatedNotes[index][0] = cleared;
          notesChanged = true;
        }
        return;
      }
      if (classification.type === MONTHLY_ROW_TYPES.HYBRID_CONFLICT) {
        setSystemNote_(
          sheet.getRange(row, columns.CHART_NO),
          MONTHLY_STRUCTURE_NOTE_PREFIX,
          '第一欄是完整日期但該列已有 CalendarEventId；系統未將它當作日期標題或報到時間。'
        );
        return;
      }
      if (needsMonthlyHeaderPresentation_(values, columns)) {
        ensureMonthlyHeaderPresentation_(sheet, row, columns, values);
      }
    }

    const info = resolveCalendarTime_(value);
    const note = buildSystemNoteText_(
      originalNotes[index][0],
      TIME_ERROR_NOTE_PREFIX,
      info.errorMessage || ''
    );
    if (note !== originalNotes[index][0]) {
      updatedNotes[index][0] = note;
      notesChanged = true;
    }
    if (info.errorMessage) {
      errors++;
      return;
    }
    if (info.parsedTime && toCellText_(value) !== info.parsedTime.text) {
      normalizedByRow[row] = info.parsedTime.text;
      normalized++;
    }
  });

  compactRowSegments_(Object.keys(normalizedByRow).map(Number))
    .forEach(segment => {
      const values = [];
      for (
        let row = segment.start;
        row < segment.start + segment.count;
        row++
      ) {
        values.push([normalizedByRow[row]]);
      }
      sheet.getRange(
        segment.start,
        columns.TIME,
        segment.count,
        1
      ).setNumberFormat('@').setValues(values);
    });
  if (notesChanged) {
    timeRange.setNotes(updatedNotes);
  }
  return { normalized, errors, headers };
}

function isEditTouchingColumns_(range, columns, keys) {
  const start = range.getColumn();
  const end = start + range.getNumColumns() - 1;
  return (keys || []).some(key => {
    const column = columns[key];
    return column && column >= start && column <= end;
  });
}

function appendMonthlyBlock_(sheet, date, hospital, columns, options) {
  const cols = columns || getRequiredMonthlyColumns_(sheet);
  const settings = options || {};
  const dateKey = formatDateKey_(date);
  const normalizedHospital = toCellText_(hospital) || HOSPITAL_KAOH;
  const currentScan = settings.scan || scanMonthlyBlocks_(sheet, cols);
  const existing = currentScan.blocks.find(block => {
    return block.dateKey === dateKey && block.hospital === normalizedHospital;
  });
  if (existing && !settings.skipExistingCheck) return existing;

  const lastBlock = currentScan.blocks.length
    ? currentScan.blocks[currentScan.blocks.length - 1]
    : null;
  const lastPatientRow = lastBlock && lastBlock.patientRows.length
    ? Math.max(...lastBlock.patientRows.map(item => item.row))
    : lastBlock
      ? lastBlock.row
      : Math.max(1, sheet.getLastRow());
  const row = lastBlock
    ? lastPatientRow + MONTHLY_TEMPLATE_BLANK_ROWS + 1
    : Math.max(2, sheet.getLastRow() + 1);
  const requiredRows = row + MONTHLY_TEMPLATE_BLANK_ROWS;
  if (requiredRows > sheet.getMaxRows()) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      requiredRows - sheet.getMaxRows()
    );
  }
  sheet.getRange(row, cols.TIME)
    .setValue(date)
    .setNumberFormat('yyyy/m/d ddd');
  sheet.getRange(row, cols.HOSPITAL).setValue(normalizedHospital);
  sheet.getRange(row, cols.CHART_NO).setValue(MONTHLY_DATE_HEADER_MARKER);
  applyManagedRowFormat_(sheet, row, 'MONTHLY', cols);
  ensureMonthlyHeaderPresentation_(sheet, row, cols);
  return {
    row,
    date,
    dateKey,
    hospital: normalizedHospital,
    blockKey: `${dateKey}|${normalizedHospital}`,
    patientRows: [],
    blankRows: Array.from(
      { length: MONTHLY_TEMPLATE_BLANK_ROWS },
      (_, index) => row + index + 1
    )
  };
}

function ensureDefaultMonthlyBlocks_(sheet, columns) {
  const monthName = sheet.getName();
  const dates = getDefaultMonthlyDates_(monthName);
  const scan = scanMonthlyBlocks_(sheet, columns);
  const existingKeys = {};
  scan.blocks.forEach(block => {
    existingKeys[block.blockKey] = true;
  });
  let created = 0;
  dates.forEach(date => {
    const blockKey = `${formatDateKey_(date)}|${HOSPITAL_KAOH}`;
    if (!existingKeys[blockKey]) {
      const block = appendMonthlyBlock_(
        sheet,
        date,
        HOSPITAL_KAOH,
        columns,
        { skipExistingCheck: true, scan }
      );
      scan.blocks.push(block);
      existingKeys[blockKey] = true;
      created++;
    }
  });
  return created;
}

function getMonthlyTimeSortKey_(value, originalIndex) {
  const info = resolveCalendarTime_(value);
  if (info.errorMessage) {
    return { group: 1, minutes: 0, text: getRawTimeText_(value), originalIndex };
  }
  if (info.parsedTime) {
    return {
      group: 0,
      minutes: info.parsedTime.hour * 60 + info.parsedTime.minute,
      text: '',
      originalIndex
    };
  }
  const text = getRawTimeText_(value);
  return {
    group: text ? 1 : 2,
    minutes: 0,
    text,
    originalIndex
  };
}

function compareMonthlyPatientRows_(left, right, columns) {
  const leftKey = getMonthlyTimeSortKey_(
    getRowFieldValue_(left.values, columns, 'TIME'),
    left.originalIndex
  );
  const rightKey = getMonthlyTimeSortKey_(
    getRowFieldValue_(right.values, columns, 'TIME'),
    right.originalIndex
  );
  if (leftKey.group !== rightKey.group) return leftKey.group - rightKey.group;
  if (leftKey.minutes !== rightKey.minutes) {
    return leftKey.minutes - rightKey.minutes;
  }
  return leftKey.originalIndex - rightKey.originalIndex;
}

function captureMonthlyRows_(sheet, lastRow, lastColumn) {
  if (lastRow < 2) return [];
  const range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
  const values = range.getValues();
  const formulas = range.getFormulasR1C1();
  const notes = range.getNotes();
  const backgrounds = range.getBackgrounds();
  const fontColors = range.getFontColors();
  const fontFamilies = range.getFontFamilies();
  const fontSizes = range.getFontSizes();
  const fontWeights = range.getFontWeights();
  const numberFormats = range.getNumberFormats();
  const horizontal = range.getHorizontalAlignments();
  const vertical = range.getVerticalAlignments();
  const wraps = range.getWrapStrategies();
  const validations = range.getDataValidations();
  return values.map((rowValues, index) => ({
    sourceRow: index + 2,
    rowHeight: typeof sheet.getRowHeight === 'function'
      ? sheet.getRowHeight(index + 2)
      : null,
    values: rowValues.slice(),
    formulas: formulas[index].slice(),
    notes: notes[index].slice(),
    backgrounds: backgrounds[index].slice(),
    fontColors: fontColors[index].slice(),
    fontFamilies: fontFamilies[index].slice(),
    fontSizes: fontSizes[index].slice(),
    fontWeights: fontWeights[index].slice(),
    numberFormats: numberFormats[index].slice(),
    horizontal: horizontal[index].slice(),
    vertical: vertical[index].slice(),
    wraps: wraps[index].slice(),
    validations: validations[index].slice()
  }));
}

function normalizeSortFingerprintCell_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return ['date', value.getTime()];
  }
  if (typeof value === 'number') {
    return ['number', isFinite(value) ? value : String(value)];
  }
  if (typeof value === 'boolean') return ['boolean', value];
  return ['text', value === null || value === undefined ? '' : String(value)];
}

function fingerprintSortableRangeState_(values, formulas, notes) {
  const fingerprints = [];
  (values || []).forEach((rowValues, rowIndex) => {
    const rowFormulas = (formulas || [])[rowIndex] || [];
    const rowNotes = (notes || [])[rowIndex] || [];
    const cells = rowValues.map((value, columnIndex) => {
      const formula = toCellText_(rowFormulas[columnIndex]);
      return {
        content: formula
          ? ['formulaR1C1', formula]
          : normalizeSortFingerprintCell_(value),
        note:
          rowNotes[columnIndex] === null ||
          rowNotes[columnIndex] === undefined
            ? ''
            : String(rowNotes[columnIndex])
      };
    });
    const meaningful = cells.some(cell => {
      return cell.content[0] === 'formulaR1C1' ||
        toCellText_(cell.content[1]) ||
        cell.note !== '';
    });
    if (meaningful) {
      fingerprints.push(sha256Hex_(JSON.stringify(cells)));
    }
  });
  return fingerprints.sort();
}

function captureSortableRangeState_(sheet, startRow, rowCount, lastColumn) {
  if (!rowCount || rowCount < 1) {
    return {
      values: [],
      formulas: [],
      notes: [],
      fingerprints: []
    };
  }
  const range = sheet.getRange(startRow, 1, rowCount, lastColumn);
  const values = range.getValues();
  const formulas = typeof range.getFormulasR1C1 === 'function'
    ? range.getFormulasR1C1()
    : range.getFormulas();
  const notes = range.getNotes();
  return {
    values,
    formulas,
    notes,
    fingerprints: fingerprintSortableRangeState_(
      values,
      formulas,
      notes
    )
  };
}

function areSortKeysInAscendingOrder_(sortKeys) {
  for (let index = 1; index < (sortKeys || []).length; index++) {
    if (String(sortKeys[index - 1]) > String(sortKeys[index])) return false;
  }
  return true;
}

function sortRowsWithTemporaryKeys_(
  sheet,
  startRow,
  rowCount,
  lastColumn,
  sortKeys,
  verifyCallback
) {
  if (!rowCount || rowCount < 2) {
    return {
      ok: true,
      sorted: false,
      verification: typeof verifyCallback === 'function'
        ? verifyCallback()
        : null
    };
  }
  if (!Array.isArray(sortKeys) || sortKeys.length !== rowCount) {
    throw new Error('暫時排序鍵數量與資料列數不一致。');
  }

  const helperStartColumn = lastColumn + 1;
  const originalOrderColumn = helperStartColumn + 1;
  const widthWithHelpers = originalOrderColumn;
  let helpersInserted = false;
  let sorted = false;
  let result = null;
  let failure = null;

  try {
    sheet.insertColumnsAfter(lastColumn, 2);
    helpersInserted = true;
    const helperValues = sortKeys.map((key, index) => [
      toCellText_(key),
      String(index).padStart(10, '0')
    ]);
    sheet.getRange(
      startRow,
      helperStartColumn,
      rowCount,
      2
    ).setValues(helperValues);
    sheet.getRange(
      startRow,
      1,
      rowCount,
      widthWithHelpers
    ).sort({
      column: helperStartColumn,
      ascending: true
    });
    sorted = true;
    result = typeof verifyCallback === 'function'
      ? verifyCallback()
      : null;
  } catch (err) {
    failure = err;
    if (sorted) {
      try {
        sheet.getRange(
          startRow,
          1,
          rowCount,
          widthWithHelpers
        ).sort({
          column: originalOrderColumn,
          ascending: true
        });
      } catch (rollbackError) {
        failure = new Error(
          `排序失敗且原順序還原失敗：${err.message || err}；` +
          `${rollbackError.message || rollbackError}`
        );
      }
    }
  }

  if (helpersInserted) {
    try {
      sheet.deleteColumns(helperStartColumn, 2);
    } catch (cleanupError) {
      if (!failure) {
        failure = new Error(
          `排序完成，但暫時排序欄清理失敗：${cleanupError.message || cleanupError}`
        );
      }
    }
  }
  if (failure) throw failure;
  return {
    ok: true,
    sorted: true,
    verification: result
  };
}

function buildMonthlySnapshotGroups_(scan, snapshotsByRow) {
  return (scan.blocks || []).map((block, blockIndex) => {
    const group = {
      date: block.date,
      dateKey: block.dateKey,
      hospital: block.hospital,
      header: snapshotsByRow[block.row],
      originalHeaderRow: block.row,
      originalBlockIndex: blockIndex,
      patients: []
    };
    block.patientRows.forEach((patient, index) => {
      group.patients.push({
        ...snapshotsByRow[patient.row],
        originalIndex: patient.row * 1000 + index
      });
    });
    return group;
  });
}

function compareMonthlySnapshotGroups_(left, right) {
  const dateCompare = left.date.getTime() - right.date.getTime();
  if (dateCompare) return dateCompare;
  const hospitalOrder = hospital => (
    hospital === HOSPITAL_KAOH ? 0 : hospital === HOSPITAL_UNION ? 1 : 2
  );
  return hospitalOrder(left.hospital) - hospitalOrder(right.hospital) ||
    left.hospital.localeCompare(right.hospital) ||
    left.originalHeaderRow - right.originalHeaderRow ||
    left.originalBlockIndex - right.originalBlockIndex;
}

function buildMonthlyNativeSortPlan_(scan, columns) {
  const orderedGroups = (scan.blocks || []).map((block, index) => ({
    ...block,
    originalHeaderRow: block.row,
    originalBlockIndex: index
  })).sort(compareMonthlySnapshotGroups_);
  const ranksBySourceRow = {};
  const generatedSourceRows = [];
  const generatedFinalRows = [];
  const finalHeaderRows = [];
  let nextGeneratedRow = scan.lastRow + 1;
  let rank = 0;
  let patientCount = 0;

  const assign = (sourceRow, generated) => {
    ranksBySourceRow[sourceRow] = rank;
    if (generated) {
      generatedSourceRows.push(sourceRow);
      generatedFinalRows.push(rank + 2);
    }
    rank++;
  };

  orderedGroups.forEach(group => {
    finalHeaderRows.push(rank + 2);
    assign(group.row, false);
    group.patientRows
      .map(patient => ({
        ...patient,
        originalIndex: patient.row
      }))
      .sort((left, right) => {
        return compareMonthlyPatientRows_(left, right, columns);
      })
      .forEach(patient => {
        assign(patient.row, false);
        patientCount++;
      });
    const blankRows = group.blankRows.slice();
    while (blankRows.length < MONTHLY_TEMPLATE_BLANK_ROWS) {
      blankRows.push(nextGeneratedRow);
      nextGeneratedRow++;
    }
    blankRows.forEach(row => {
      assign(row, row >= scan.lastRow + 1);
    });
  });

  const targetLastRow = Math.max(
    scan.lastRow,
    nextGeneratedRow - 1
  );
  for (let row = 2; row <= targetLastRow; row++) {
    if (Object.prototype.hasOwnProperty.call(ranksBySourceRow, row)) {
      continue;
    }
    assign(row, false);
  }
  const rowCount = targetLastRow - 1;
  const sortKeys = [];
  for (let row = 2; row <= targetLastRow; row++) {
    sortKeys.push(String(ranksBySourceRow[row]).padStart(12, '0'));
  }
  return {
    orderedGroups,
    rowCount,
    targetLastRow,
    sortKeys,
    generatedSourceRows,
    generatedFinalRows,
    finalHeaderRows,
    patientCount
  };
}

function validateMonthlyNativeSortScan_(sheet, scan, expectedBlockCount) {
  if (scan.hybrids.length || scan.orphans.length) {
    throw new Error('原生排序後出現混合列或孤立病人列。');
  }
  const markerConflicts = scan.markerIssues.filter(item => {
    return item.type === 'monthly_header_marker_conflict';
  });
  if (markerConflicts.length) {
    throw new Error('原生排序後日期標題標記發生衝突。');
  }
  if (scan.blocks.length !== expectedBlockCount) {
    throw new Error(
      `原生排序後日期區塊數量由 ${expectedBlockCount} 變成 ${scan.blocks.length}。`
    );
  }
  scan.blocks.forEach((block, blockIndex) => {
    if (formatMonthKey_(block.date) !== sheet.getName()) {
      throw new Error(`排序後第 ${block.row} 列日期不屬於 ${sheet.getName()}。`);
    }
    if (block.blankRows.length < MONTHLY_TEMPLATE_BLANK_ROWS) {
      throw new Error(
        `排序後第 ${block.row} 列刀日只有 ${block.blankRows.length} 列空白輸入列。`
      );
    }
    if (blockIndex > 0) {
      const previous = scan.blocks[blockIndex - 1];
      const compare = compareMonthlySnapshotGroups_(
        {
          ...previous,
          originalHeaderRow: previous.row,
          originalBlockIndex: blockIndex - 1
        },
        {
          ...block,
          originalHeaderRow: block.row,
          originalBlockIndex: blockIndex
        }
      );
      if (compare > 0) {
        throw new Error('排序後日期或醫院區塊順序不正確。');
      }
    }
    let previousTimeKey = null;
    block.patientRows.forEach((patient, patientIndex) => {
      const currentTimeKey = getMonthlyTimeSortKey_(
        getRowFieldValue_(patient.values, scan.columns, 'TIME'),
        patientIndex
      );
      if (
        previousTimeKey &&
        (
          currentTimeKey.group < previousTimeKey.group ||
          (
            currentTimeKey.group === previousTimeKey.group &&
            currentTimeKey.minutes < previousTimeKey.minutes
          )
        )
      ) {
        throw new Error(`第 ${block.row} 列刀日內報到時間順序不正確。`);
      }
      previousTimeKey = currentTimeKey;
    });
  });
  return true;
}

function rebuildMonthlySheetSorted_(sheet, options) {
  const settings = options || {};
  const columns = getRequiredMonthlyColumns_(sheet);
  const scan = settings.precomputedScan ||
    scanMonthlyBlocks_(sheet, columns);
  if (scan.hybrids.length) {
    throw new Error(
      `${sheet.getName()} 有完整日期與 CalendarEventId 並存的混合列：` +
      scan.hybrids.map(item => item.row).join('、')
    );
  }
  const markerConflicts = scan.markerIssues.filter(item => {
    return item.type === 'monthly_header_marker_conflict';
  });
  if (markerConflicts.length) {
    throw new Error(
      `${sheet.getName()} 的刀日標記格含既有內容：` +
      markerConflicts.map(item => item.row).join('、')
    );
  }
  if (scan.orphans.length) {
    throw new Error(
      `${sheet.getName()} 有缺少上方日期標題的資料列：` +
      scan.orphans.map(item => item.row).join('、')
    );
  }
  scan.blocks.forEach(block => {
    if (formatMonthKey_(block.date) !== sheet.getName()) {
      throw new Error(
        `${sheet.getName()} 第 ${block.row} 列日期 ${block.dateKey} 不屬於該月份。`
      );
    }
  });
  if (!scan.blocks.length) {
    return {
      ok: true,
      blockCount: 0,
      patientCount: 0,
      strategy: 'native_range_sort'
    };
  }

  let normalizedHeaderHospitals = 0;
  scan.blocks.forEach(block => {
    const hospital = toCellText_(
      getRowFieldValue_(block.values, columns, 'HOSPITAL')
    );
    if (!hospital) {
      normalizedHeaderHospitals++;
      sheet.getRange(block.row, columns.HOSPITAL).setValue(HOSPITAL_KAOH);
      setRowFieldValue_(block.values, columns, 'HOSPITAL', HOSPITAL_KAOH);
      setRowFieldValue_(
        scan.rowValuesByRow[block.row],
        columns,
        'HOSPITAL',
        HOSPITAL_KAOH
      );
    }
  });

  const lastColumn = Math.max(
    scan.lastColumn,
    sheet.getLastColumn(),
    ...Object.values(columns)
  );
  const plan = buildMonthlyNativeSortPlan_(scan, columns);
  let insertedRows = 0;
  if (plan.targetLastRow > sheet.getMaxRows()) {
    insertedRows = plan.targetLastRow - sheet.getMaxRows();
    sheet.insertRowsAfter(sheet.getMaxRows(), insertedRows);
  }
  // 最後一個區塊後方的空白輸入列通常不會被 getLastRow() 計入，
  // 但工作表本身已有可用空白列。若所有來源列已在正確順位，就不必
  // 只為這些尾端空白列執行一次昂貴的結構排序。
  const changed = !areSortKeysInAscendingOrder_(plan.sortKeys);
  if (!changed) {
    if (insertedRows) {
      applyManagedRowsFormatBatch_(
        sheet,
        plan.generatedFinalRows,
        'MONTHLY',
        columns
      );
    }
    const normalizedSheetScan = normalizedHeaderHospitals
      ? buildMonthlySheetCalendarScanFromBlockScan_(
        sheet,
        columns,
        scan
      )
      : null;
    return {
      ok: true,
      blockCount: scan.blocks.length,
      patientCount: plan.patientCount,
      generatedBlankRows: insertedRows,
      changed: false,
      registryRefreshRequired: Boolean(normalizedHeaderHospitals),
      sheetScan:
        normalizedSheetScan ||
        settings.precomputedCalendarScan ||
        null,
      strategy: normalizedHeaderHospitals
        ? 'normalized_header'
        : 'precomputed_noop'
    };
  }
  const before = captureSortableRangeState_(
    sheet,
    2,
    plan.rowCount,
    lastColumn
  );
  let verifiedScan = null;
  sortRowsWithTemporaryKeys_(
    sheet,
    2,
    plan.rowCount,
    lastColumn,
    plan.sortKeys,
    () => {
      const written = captureSortableRangeState_(
        sheet,
        2,
        plan.rowCount,
        lastColumn
      );
      if (
        JSON.stringify(written.fingerprints) !==
        JSON.stringify(before.fingerprints)
      ) {
        throw new Error('排序後資料、公式或 notes 核對不一致。');
      }
      verifiedScan = scanMonthlyBlocks_(sheet, columns);
      validateMonthlyNativeSortScan_(
        sheet,
        verifiedScan,
        scan.blocks.length
      );
      return {
        fingerprints: written.fingerprints,
        blockCount: verifiedScan.blocks.length
      };
    }
  );
  applyManagedRowsFormatBatch_(
    sheet,
    plan.generatedFinalRows,
    'MONTHLY',
    columns
  );
  if (typeof sheet.autoResizeRows === 'function') {
    sheet.autoResizeRows(2, plan.rowCount);
  }
  return {
    ok: true,
    blockCount: verifiedScan
      ? verifiedScan.blocks.length
      : scan.blocks.length,
    patientCount: plan.patientCount,
    generatedBlankRows: plan.generatedFinalRows.length,
    changed: true,
    registryRefreshRequired: true,
    sheetScan: verifiedScan
      ? buildMonthlySheetCalendarScanFromBlockScan_(
        sheet,
        columns,
        verifiedScan
      )
      : null,
    strategy: 'native_range_sort'
  };
}

function findOrCreateMonthlyBlock_(
  sheet,
  date,
  hospital,
  columns,
  precomputedScan
) {
  const dateKey = formatDateKey_(date);
  const normalizedHospital = toCellText_(hospital) || HOSPITAL_KAOH;
  const scan = precomputedScan || scanMonthlyBlocks_(sheet, columns);
  const matches = scan.blocks.filter(block => {
    return block.dateKey === dateKey && block.hospital === normalizedHospital;
  });
  if (matches.length > 1) {
    throw new Error(
      `${dateKey}／${normalizedHospital} 有多個刀日區塊；` +
      '請從視窗的現有日期區塊清單選擇明確列號。'
    );
  }
  if (matches[0]) return matches[0];
  const created = appendMonthlyBlock_(
    sheet,
    date,
    normalizedHospital,
    columns,
    { scan }
  );
  scan.blocks.push(created);
  return created;
}

function resolveMonthlyBlockRef_(sheet, blockRef, columns, precomputedScan) {
  const ref = blockRef || {};
  const row = Number(ref.headerRow || ref.row);
  if (!row || row < 2) return null;
  const scan = precomputedScan || scanMonthlyBlocks_(sheet, columns);
  const block = scan.blocks.find(item => item.row === row);
  if (!block) return null;
  if (
    ref.dateKey &&
    block.dateKey !== toCellText_(ref.dateKey)
  ) {
    return null;
  }
  if (
    ref.hospital &&
    block.hospital !== toCellText_(ref.hospital)
  ) {
    return null;
  }
  return block;
}

function insertPatientAtBlockEnd_(
  sheet,
  block,
  valuesByKey,
  columns,
  precomputedScan
) {
  const scan = precomputedScan || scanMonthlyBlocks_(sheet, columns);
  const currentBlock = scan.blocks.find(item => {
    return item.row === Number(block && block.row) &&
      item.dateKey === toCellText_(block && block.dateKey) &&
      item.hospital === toCellText_(block && block.hospital);
  });
  if (!currentBlock) throw new Error('找不到指定日期區塊。');
  const nextBlock = scan.blocks.find(item => item.row > currentBlock.row);
  const lastPatientRow = currentBlock.patientRows.length
    ? Math.max(...currentBlock.patientRows.map(item => item.row))
    : currentBlock.row;
  const targetRow = lastPatientRow + 1;
  let nextHeaderRow = nextBlock ? nextBlock.row : 0;
  if (nextHeaderRow && targetRow >= nextHeaderRow) {
    sheet.insertRowBefore(nextHeaderRow);
    nextHeaderRow++;
  } else if (!nextHeaderRow && targetRow > sheet.getMaxRows()) {
    const currentMaxRows = sheet.getMaxRows();
    sheet.insertRowsAfter(currentMaxRows, targetRow - currentMaxRows);
  }

  const lastColumn = Math.max(getLastHeaderColumn_(sheet), ...Object.values(columns));
  const rowValues = Array(lastColumn).fill('');
  Object.keys(valuesByKey || {}).forEach(key => {
    setRowFieldValue_(rowValues, columns, key, valuesByKey[key]);
  });
  sheet.getRange(targetRow, 1, 1, lastColumn).setValues([rowValues]);
  applyManagedRowFormat_(sheet, targetRow, 'MONTHLY', columns);

  if (nextHeaderRow) {
    const blankCount = Math.max(0, nextHeaderRow - targetRow - 1);
    const shortage = MONTHLY_TEMPLATE_BLANK_ROWS - blankCount;
    if (shortage > 0) {
      sheet.insertRowsBefore(nextHeaderRow, shortage);
    }
  } else {
    const blankCount = Math.max(0, sheet.getMaxRows() - targetRow);
    const shortage = MONTHLY_TEMPLATE_BLANK_ROWS - blankCount;
    if (shortage > 0) {
      sheet.insertRowsAfter(sheet.getMaxRows(), shortage);
    }
  }
  return targetRow;
}

function getMonthlyBlockOptions_(sheet) {
  const columns = getRequiredMonthlyColumns_(sheet);
  const blocks = scanMonthlyBlocks_(sheet, columns).blocks;
  const counts = {};
  blocks.forEach(block => {
    counts[block.blockKey] = (counts[block.blockKey] || 0) + 1;
  });
  return blocks.map(block => ({
    date: block.dateKey,
    hospital: block.hospital,
    headerRow: block.row,
    occurrence: block.occurrence,
    label:
      `${block.dateKey}｜${block.hospital}` +
      (counts[block.blockKey] > 1 ? `｜第 ${block.row} 列` : '')
  }));
}
