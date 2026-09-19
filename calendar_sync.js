/**
 * Calendar 事件生命週期與 CALENDAR_ROW_REGISTRY_V2。
 *
 * Registry 只保存 Event ID、表單位置、日期區塊、內容／綁定雜湊
 * 與安全的 pending 狀態；
 * 不保存病歷號、姓名、診斷或 Plan 原文。
 */
const CALENDAR_REGISTRY_VERSION = 2;
const CALENDAR_REGISTRY_INDEX_KEY = 'CALENDAR_ROW_REGISTRY_V2_INDEX';
const CALENDAR_REGISTRY_CHUNK_PREFIX = 'CALENDAR_ROW_REGISTRY_V2_';
const CALENDAR_REGISTRY_CHUNK_SIZE = 7000;
const CALENDAR_PENDING_QUEUE_LOCK_WAIT_MS = 1000;
const CALENDAR_REPAIR_MAX_PREVIEW_ROWS = 40;
const CALENDAR_FAST_FINGERPRINT_VERSION = 2;

function buildCalendarTitle_(chartNo, patientName, condition) {
  return [
    toSingleLineText_(chartNo),
    toSingleLineText_(patientName),
    toSingleLineText_(condition)
  ].join(` ${CALENDAR_TITLE_SEPARATOR} `);
}

function normalizeCalendarBindingCondition_(condition) {
  return toSingleLineText_(condition)
    .replace(/\s+\d{8}\s*$/, '')
    .trim();
}

function buildCalendarBindingSignature_(chartNo, patientName, condition) {
  return sha256Hex_(JSON.stringify([
    toSingleLineText_(chartNo),
    toSingleLineText_(patientName),
    normalizeCalendarBindingCondition_(condition)
  ]));
}

function getContextCalendarBindingSignature_(context) {
  return buildCalendarBindingSignature_(
    context && context.chartNo,
    context && context.patientName,
    context && context.condition
  );
}

function getEventCalendarBindingSignature_(event) {
  const slots = toCellText_(event && event.summary).split(/\s*\|\s*/);
  if (slots.length !== 3) return '';
  return buildCalendarBindingSignature_(slots[0], slots[1], slots[2]);
}

function getCalendarPrivateProperty_(event, key) {
  return toCellText_(
    event &&
    event.extendedProperties &&
    event.extendedProperties.private &&
    event.extendedProperties.private[key]
  );
}

function isCancelledMonthlyCalendarEvent_(event) {
  return Boolean(
    getCalendarPrivateProperty_(event, CALENDAR_PRIVATE_KIND_KEY) ===
      CALENDAR_PRIVATE_MONTHLY_KIND &&
    getCalendarPrivateProperty_(event, CALENDAR_PRIVATE_STATE_KEY) ===
      CALENDAR_PRIVATE_CANCELLED_STATE
  );
}

function buildCalendarDescription_(items) {
  return (items || [])
    .filter(item => item && toCellText_(item.value))
    .map(item => `${item.label}: ${toCellText_(item.value)}`)
    .join('\n');
}

function buildMonthlyCondition_(data, date) {
  const value = data || {};
  const side = toSingleLineText_(value.side);
  const diagnosis = toSingleLineText_(value.diagnosis);
  const grade = toSingleLineText_(value.grade);
  const procedure = toSingleLineText_(value.procedure);
  const iol = toSingleLineText_(value.iol);
  const target = toSingleLineText_(value.iolTarget);
  const finalValue = toSingleLineText_(value.iolFinal);
  const axisRaw = toSingleLineText_(value.axis).replace(/^@/, '');
  const sidePattern = /(?:^|\s)(OD|OS|OU)(?:\s|$)/i;
  const diagnosisSide = side && !sidePattern.test(diagnosis) ? ` ${side}` : '';
  const procedureSide = side && !sidePattern.test(procedure) ? ` ${side}` : '';
  const diagnosisText = diagnosis
    ? `${diagnosis}${grade ? `(${grade})` : ''}${diagnosisSide}`
    : '';
  const iolParts = [iol, target, finalValue, axisRaw ? `@${axisRaw}` : '']
    .filter(Boolean);
  const procedureText = procedure
    ? `${procedure}${iolParts.length ? `(${iolParts.join(' ')})` : ''}${procedureSide}`
    : iolParts.join(' ');
  const segments = [];
  if (diagnosisText) segments.push(diagnosisText);
  if (procedureText) {
    if (segments.length) segments.push('s/p');
    segments.push(procedureText);
  }
  const compactDate = formatCompactDate_(date);
  if (compactDate) segments.push(compactDate);
  return segments.join(' ').replace(/\s+/g, ' ').trim();
}

function parseEventDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    const year = value.getFullYear();
    if (year < 2000 || year > 2099) return null;
    return new Date(year, value.getMonth(), value.getDate());
  }
  return parseFullDate_(value);
}

function formatCalendarDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatCalendarDateTime_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd'T'HH:mm:ssXXX"
  );
}

function addDays_(date, days) {
  const result = new Date(date.getTime());
  result.setDate(result.getDate() + days);
  return result;
}

function getCalendarColorId_(context) {
  if (context.kind === 'MONTHLY' && context.cancelled) {
    return CONFIG.COLORS.MONTHLY_CANCELLED;
  }
  if (context.kind === 'FU') return CONFIG.COLORS.FU;
  if (toCellText_(context.ga)) return CONFIG.COLORS.MONTHLY_GA;
  return context.timeInfo && context.timeInfo.parsedTime
    ? CONFIG.COLORS.MONTHLY_TIMED
    : CONFIG.COLORS.MONTHLY_UNDECIDED;
}

function buildCalendarResource_(context) {
  const title = buildCalendarTitle_(
    context.chartNo,
    context.patientName,
    context.condition
  );
  const descriptionItems = [];
  if (context.kind === 'MONTHLY') {
    descriptionItems.push({ label: '醫院', value: context.hospital });
  } else if (context.hospital) {
    descriptionItems.push({ label: '醫院', value: context.hospital });
  }
  descriptionItems.push({ label: 'TEL', value: context.tel });
  descriptionItems.push({ label: 'Plan', value: context.plan });
  if (context.ga) descriptionItems.push({ label: 'GA', value: context.ga });

  const forceAllDay = context.kind === 'FU' || Boolean(context.ga);
  if (forceAllDay && context.timeInfo.parsedTime) {
    descriptionItems.push({
      label: '原報到時間',
      value: context.timeInfo.parsedTime.text
    });
  } else if (context.timeInfo.timeNote) {
    descriptionItems.push({
      label: 'Time note',
      value: context.timeInfo.timeNote
    });
  }

  const resource = {
    summary: title,
    description: buildCalendarDescription_(descriptionItems),
    colorId: getCalendarColorId_(context)
  };
  if (context.kind === 'MONTHLY' && context.cancelled) {
    resource.extendedProperties = {
      private: {
        [CALENDAR_PRIVATE_KIND_KEY]: CALENDAR_PRIVATE_MONTHLY_KIND,
        [CALENDAR_PRIVATE_STATE_KEY]: CALENDAR_PRIVATE_CANCELLED_STATE
      }
    };
  }
  if (!forceAllDay && context.timeInfo.parsedTime) {
    const start = new Date(context.date.getTime());
    start.setHours(
      context.timeInfo.parsedTime.hour,
      context.timeInfo.parsedTime.minute,
      0,
      0
    );
    const end = new Date(start.getTime() + 30 * 60 * 1000);
    resource.start = {
      dateTime: formatCalendarDateTime_(start),
      timeZone: Session.getScriptTimeZone()
    };
    resource.end = {
      dateTime: formatCalendarDateTime_(end),
      timeZone: Session.getScriptTimeZone()
    };
  } else {
    resource.start = { date: formatCalendarDate_(context.date) };
    resource.end = { date: formatCalendarDate_(addDays_(context.date, 1)) };
  }
  return resource;
}

function calendarEndpointMatches_(actual, expected) {
  const left = actual || {};
  const right = expected || {};
  if (right.date) return toCellText_(left.date) === toCellText_(right.date);
  if (right.dateTime) {
    const leftTime = new Date(left.dateTime).getTime();
    const rightTime = new Date(right.dateTime).getTime();
    return !isNaN(leftTime) && !isNaN(rightTime) && leftTime === rightTime;
  }
  return !left.date && !left.dateTime;
}

function calendarEventMatchesResource_(event, resource) {
  const expectedCancelled = isCancelledMonthlyCalendarEvent_(resource);
  const actualCancelled = isCancelledMonthlyCalendarEvent_(event);
  return Boolean(
    event &&
    toCellText_(event.summary) === toCellText_(resource.summary) &&
    toCellText_(event.description) === toCellText_(resource.description) &&
    toCellText_(event.colorId) === toCellText_(resource.colorId) &&
    expectedCancelled === actualCancelled &&
    calendarEndpointMatches_(event.start, resource.start) &&
    calendarEndpointMatches_(event.end, resource.end)
  );
}

function isCalendarNotFoundError_(err) {
  const message = String((err && err.message) || err || '').toLowerCase();
  return message.indexOf('404') !== -1 ||
    message.indexOf('not found') !== -1;
}

function getContextNoteColumn_(context) {
  return (
    context.chartNo && context.columns.CHART_NO
  ) || (
    context.patientName && context.columns.NAME
  ) ||
    context.columns.CHART_NO ||
    context.columns.NAME ||
    context.columns.TIME ||
    context.columns.EVENT_ID;
}

function getContextNoteCell_(context) {
  return context.sheet.getRange(
    context.row,
    getContextNoteColumn_(context)
  );
}

function clearContextSystemNotes_(context) {
  const cells = [getContextNoteCell_(context)];
  if (context.columns.EVENT_ID) {
    cells.push(context.sheet.getRange(context.row, context.columns.EVENT_ID));
  }
  cells.forEach(cell => {
    clearSystemNote_(
      cell,
      [CALENDAR_SYNC_NOTE_PREFIX, CALENDAR_CONFLICT_NOTE_PREFIX]
    );
  });
  if (context.columns.TIME) {
    clearSystemNote_(
      context.sheet.getRange(context.row, context.columns.TIME),
      [TIME_ERROR_NOTE_PREFIX]
    );
  }
}

function isCalendarNoopDraftContext_(context) {
  if (!context || context.eventId) return false;
  return Boolean(
    context.invalidReason === 'no_identity' ||
    (context.kind === 'FU' && context.invalidReason === 'no_date')
  );
}

function clearContextsSystemNotesBatch_(contexts) {
  const grouped = {};
  (contexts || []).forEach(context => {
    if (!context || !context.sheet || !context.row) return;
    const key = String(context.sheetId || context.sheet.getSheetId());
    if (!grouped[key]) {
      grouped[key] = {
        sheet: context.sheet,
        contexts: []
      };
    }
    grouped[key].contexts.push(context);
  });

  let cleared = 0;
  Object.keys(grouped).forEach(key => {
    const group = grouped[key];
    const targets = {};
    group.contexts.forEach(context => {
      const addTarget = (column, prefixes) => {
        if (!column) return;
        const targetKey = `${context.row}:${column}`;
        if (!targets[targetKey]) {
          targets[targetKey] = {
            row: context.row,
            column,
            prefixes: []
          };
        }
        prefixes.forEach(prefix => {
          if (targets[targetKey].prefixes.indexOf(prefix) === -1) {
            targets[targetKey].prefixes.push(prefix);
          }
        });
      };
      addTarget(
        getContextNoteColumn_(context),
        [CALENDAR_SYNC_NOTE_PREFIX, CALENDAR_CONFLICT_NOTE_PREFIX]
      );
      addTarget(
        context.columns.EVENT_ID,
        [CALENDAR_SYNC_NOTE_PREFIX, CALENDAR_CONFLICT_NOTE_PREFIX]
      );
      if (context.columns.TIME) {
        addTarget(context.columns.TIME, [TIME_ERROR_NOTE_PREFIX]);
      }
    });
    const targetList = Object.values(targets);
    if (!targetList.length) return;
    const minRow = Math.min(...targetList.map(target => target.row));
    const maxRow = Math.max(...targetList.map(target => target.row));
    const minColumn = Math.min(...targetList.map(target => target.column));
    const maxColumn = Math.max(...targetList.map(target => target.column));
    const range = group.sheet.getRange(
      minRow,
      minColumn,
      maxRow - minRow + 1,
      maxColumn - minColumn + 1
    );
    const notes = range.getNotes();
    let changed = false;
    targetList.forEach(target => {
      const rowIndex = target.row - minRow;
      const columnIndex = target.column - minColumn;
      const original = toCellText_(notes[rowIndex][columnIndex]);
      const updated = target.prefixes.reduce((note, prefix) => {
        return removeSystemNoteBlockText_(note, prefix);
      }, original);
      if (updated === original) return;
      notes[rowIndex][columnIndex] = updated;
      changed = true;
      cleared++;
    });
    if (changed) range.setNotes(notes);
  });
  return cleared;
}

function getContextLocationText_(context) {
  if (!context) return '';
  const sheetName = toCellText_(context.sheetName) ||
    (context.sheet && context.sheet.getName
      ? context.sheet.getName()
      : '');
  const row = Number(context.row);
  return sheetName && row
    ? `「${sheetName}」第 ${row} 列`
    : sheetName
      ? `「${sheetName}」`
      : row
        ? `第 ${row} 列`
        : '';
}

function summarizeVisibleSyncMessage_(message, conflict, context) {
  const raw = sanitizeHealthMessage_(
    toSingleLineText_(message),
    context && context.eventId
  );
  if (!raw) {
    return conflict
      ? '資料列發現同步衝突；請執行「修復選取列同步」。'
      : 'Calendar 同步失敗；請執行「修復選取列同步」。';
  }
  if (
    /修復選取列同步|同步待處理變更|檢查同步健康|已排入重試/.test(raw)
  ) {
    return raw;
  }
  const nextStep = conflict
    ? '請執行「修復選取列同步」查看並確認修復。'
    : '請執行「修復選取列同步」重試；若仍失敗，再執行「檢查同步健康」。';
  return `${raw} ${nextStep}`;
}

function setContextSyncNote_(context, message, conflict) {
  let eventIdCell = null;
  let eventId = toCellText_(context && context.eventId);
  if (context.columns.EVENT_ID) {
    eventIdCell = context.sheet.getRange(
      context.row,
      context.columns.EVENT_ID
    );
    if (!eventId) eventId = toCellText_(eventIdCell.getValue());
    clearSystemNote_(
      eventIdCell,
      [CALENDAR_SYNC_NOTE_PREFIX, CALENDAR_CONFLICT_NOTE_PREFIX]
    );
  }
  const noteCell = getContextNoteCell_(context);
  clearSystemNote_(
    noteCell,
    [CALENDAR_SYNC_NOTE_PREFIX, CALENDAR_CONFLICT_NOTE_PREFIX]
  );
  setSystemNote_(
    noteCell,
    conflict ? CALENDAR_CONFLICT_NOTE_PREFIX : CALENDAR_SYNC_NOTE_PREFIX,
    summarizeVisibleSyncMessage_(
      message,
      conflict,
      eventId ? { ...context, eventId } : context
    )
  );
}

function buildDuplicateEventIdMessage_(context, locations) {
  const related = (locations || [])
    .filter(location => {
      return !context ||
        Number(location.sheetId) !== Number(context.sheetId) ||
        Number(location.row) !== Number(context.row);
    })
    .map(getContextLocationText_)
    .filter(Boolean);
  const suffix = related.length
    ? `與 ${related.join('、')} 重複`
    : '出現在多列';
  return `CalendarEventId ${suffix}；相關列已停止同步。` +
    '請執行「修復選取列同步」查看並確認修復。';
}

function getContextExpectedScheduleText_(context) {
  if (!context || !context.date) return '';
  const dateText = formatDateKey_(context.date);
  const timeText = context.timeInfo && context.timeInfo.parsedTime
    ? context.timeInfo.parsedTime.text
    : '全天';
  const hospital = context.kind === 'MONTHLY' && context.hospital
    ? `／${context.hospital}`
    : '';
  return `${dateText} ${timeText}${hospital}`;
}

function buildFuRowContext_(sheet, row, columns, rowValues) {
  const cols = columns || getRequiredFuColumns_(sheet);
  // Batch callers already supply a fresh row snapshot; do not reread headers per row.
  const values = rowValues || sheet.getRange(row, 1, 1,
    Math.max(getLastHeaderColumn_(sheet), ...Object.values(cols))).getValues()[0];
  const chartNo = toCellText_(getRowFieldValue_(values, cols, 'CHART_NO'));
  const patientName = toSingleLineText_(getRowFieldValue_(values, cols, 'NAME'));
  const date = parseEventDate_(getRowFieldValue_(values, cols, 'DATE'));
  const timeInfo = resolveCalendarTime_('');
  const eventId = toCellText_(getRowFieldValue_(values, cols, 'EVENT_ID'));
  const hasIdentity = Boolean(chartNo || patientName);
  let invalidReason = '';
  if (!date) invalidReason = 'no_date';
  else if (!hasIdentity) invalidReason = 'no_identity';
  const hashPayload = {
    kind: 'FU',
    date: date ? formatDateKey_(date) : '',
    time: '',
    chartNo,
    patientName,
    tel: toCellText_(getRowFieldValue_(values, cols, 'TEL')),
    hospital: toCellText_(getRowFieldValue_(values, cols, 'HOSPITAL')),
    condition: toCellText_(getRowFieldValue_(values, cols, 'COND')),
    plan: toCellText_(getRowFieldValue_(values, cols, 'PLAN'))
  };
  return {
    sheet,
    sheetId: sheet.getSheetId(),
    sheetName: sheet.getName(),
    kind: 'FU',
    archived: false,
    row,
    columns: cols,
    values,
    eventId,
    chartNo,
    patientName,
    tel: hashPayload.tel,
    hospital: hashPayload.hospital,
    condition: hashPayload.condition,
    plan: hashPayload.plan,
    ga: '',
    date,
    dateKey: date ? formatDateKey_(date) : '',
    blockKey: date ? formatDateKey_(date) : '',
    timeInfo,
    valid: !invalidReason,
    invalidReason,
    bindingHash: buildCalendarBindingSignature_(
      chartNo,
      patientName,
      hashPayload.condition
    ),
    rowHash: sha256Hex_(JSON.stringify(hashPayload))
  };
}

function isMonthlyCancelledFontLine_(value) {
  return toCellText_(value).toLowerCase() === 'line-through';
}

function readMonthlyCancelledRows_(
  sheet,
  columns,
  startRow,
  rowCount
) {
  const firstRow = Number(startRow) || 2;
  const count = Math.max(0, Number(rowCount) || 0);
  const cancelledRows = {};
  if (!count || !columns || !columns.CHART_NO) return cancelledRows;
  const fontLines = sheet
    .getRange(firstRow, columns.CHART_NO, count, 1)
    .getFontLines();
  (fontLines || []).forEach((values, index) => {
    if (isMonthlyCancelledFontLine_((values || [])[0])) {
      cancelledRows[firstRow + index] = true;
    }
  });
  return cancelledRows;
}

function buildMonthlyRowContext_(
  sheet,
  patient,
  block,
  columns,
  archived,
  cancelledRows
) {
  const values = patient.values;
  const chartNo = toCellText_(getRowFieldValue_(values, columns, 'CHART_NO'));
  const patientName = toSingleLineText_(
    getRowFieldValue_(values, columns, 'NAME')
  );
  const eventId = toCellText_(getRowFieldValue_(values, columns, 'EVENT_ID'));
  const timeValue = getRowFieldValue_(values, columns, 'TIME');
  const timeInfo = resolveCalendarTime_(timeValue);
  const data = {
    side: getRowFieldValue_(values, columns, 'SIDE'),
    diagnosis: getRowFieldValue_(values, columns, 'DIAGNOSIS'),
    grade: getRowFieldValue_(values, columns, 'GRADE'),
    procedure: getRowFieldValue_(values, columns, 'PROCEDURE'),
    iol: getRowFieldValue_(values, columns, 'IOL'),
    iolTarget: getRowFieldValue_(values, columns, 'IOL_TARGET'),
    iolFinal: getRowFieldValue_(values, columns, 'IOL_FINAL'),
    axis: getRowFieldValue_(values, columns, 'AXIS')
  };
  const date = block ? block.date : null;
  const cancelled = Boolean(
    !archived &&
    cancelledRows &&
    cancelledRows[patient.row]
  );
  const hasIdentity = Boolean(chartNo || patientName);
  let invalidReason = '';
  if (!block) invalidReason = 'no_block';
  else if (!hasIdentity) invalidReason = 'no_identity';
  else if (timeInfo.errorMessage) invalidReason = 'invalid_time';
  const hashPayload = {
    kind: 'MONTHLY',
    time: timeInfo.parsedTime
      ? timeInfo.parsedTime.text
      : getRawTimeText_(timeValue),
    chartNo,
    patientName,
    tel: toCellText_(getRowFieldValue_(values, columns, 'TEL')),
    ga: toCellText_(getRowFieldValue_(values, columns, 'GA')),
    side: toCellText_(data.side),
    diagnosis: toCellText_(data.diagnosis),
    grade: toCellText_(data.grade),
    procedure: toCellText_(data.procedure),
    plan: toCellText_(getRowFieldValue_(values, columns, 'PLAN')),
    iol: toCellText_(data.iol),
    iolTarget: toCellText_(data.iolTarget),
    iolFinal: toCellText_(data.iolFinal),
    axis: toCellText_(data.axis)
  };
  if (cancelled) hashPayload.cancelled = true;
  return {
    sheet,
    sheetId: sheet.getSheetId(),
    sheetName: sheet.getName(),
    kind: 'MONTHLY',
    archived: Boolean(archived),
    cancelled,
    row: patient.row,
    columns,
    values,
    eventId,
    chartNo,
    patientName,
    tel: hashPayload.tel,
    hospital: block ? block.hospital : '',
    condition: buildMonthlyCondition_(data, date),
    plan: hashPayload.plan,
    ga: hashPayload.ga,
    date,
    dateKey: date ? formatDateKey_(date) : '',
    blockKey: block ? block.blockKey : '',
    timeInfo,
    valid: !invalidReason,
    invalidReason,
    bindingHash: buildCalendarBindingSignature_(
      chartNo,
      patientName,
      buildMonthlyCondition_(data, date)
    ),
    rowHash: sha256Hex_(JSON.stringify(hashPayload))
  };
}

function buildMonthlyStructureContext_(sheet, issue, columns, values) {
  const rowValues = values || [];
  const chartNo = toCellText_(
    getRowFieldValue_(rowValues, columns, 'CHART_NO')
  );
  const patientName = toSingleLineText_(
    getRowFieldValue_(rowValues, columns, 'NAME')
  );
  const date = parseFullDate_(
    getRowFieldValue_(rowValues, columns, 'TIME')
  );
  return {
    sheet,
    sheetId: sheet.getSheetId(),
    sheetName: sheet.getName(),
    kind: 'MONTHLY',
    archived: false,
    row: issue.row,
    columns,
    values: rowValues,
    eventId: toCellText_(issue.eventId),
    chartNo,
    patientName,
    tel: '',
    hospital: '',
    condition: '',
    plan: '',
    ga: '',
    date,
    dateKey: date ? formatDateKey_(date) : '',
    blockKey: '',
    timeInfo: resolveCalendarTime_(''),
    valid: false,
    invalidReason: issue.type,
    structural: true,
    bindingHash: '',
    rowHash: sha256Hex_(JSON.stringify({
      kind: 'MONTHLY_STRUCTURE',
      sheetId: sheet.getSheetId(),
      row: issue.row,
      type: issue.type,
      eventId: toCellText_(issue.eventId)
    }))
  };
}

function normalizeFastFingerprintCell_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return ['date', value.getTime()];
  }
  if (typeof value === 'number') {
    return ['number', isFinite(value) ? value : String(value)];
  }
  if (typeof value === 'boolean') return ['boolean', value];
  return [
    'text',
    value === null || value === undefined ? '' : String(value)
  ];
}

function buildManagedSheetFastFingerprint_(
  kind,
  columns,
  values,
  startRow,
  cancelledRows
) {
  const keys = (
    kind === 'FU' ? FU_CALENDAR_KEYS : MONTHLY_CALENDAR_KEYS
  ).concat(['EVENT_ID']);
  const seen = {};
  const columnSpec = keys
    .filter(key => {
      if (!columns[key] || seen[key]) return false;
      seen[key] = true;
      return true;
    })
    .map(key => [key, Number(columns[key])]);
  const firstRow = Number(startRow) || 2;
  const rows = (values || []).map((rowValues, index) => {
    return [
      firstRow + index,
      columnSpec.map(item => {
        return normalizeFastFingerprintCell_(
          rowValues[item[1] - 1]
        );
      })
    ];
  });
  return sha256Hex_(JSON.stringify([
    CALENDAR_FAST_FINGERPRINT_VERSION,
    kind,
    columnSpec,
    rows,
    Object.keys(cancelledRows || {}).map(Number).sort((left, right) => {
      return left - right;
    })
  ]));
}

function buildFuSheetCalendarScanFromValues_(
  sheet,
  columns,
  values,
  startRow
) {
  const contexts = [];
  const firstRow = Number(startRow) || 2;
  const fastFingerprint = buildManagedSheetFastFingerprint_(
    'FU',
    columns,
    values,
    firstRow
  );
  (values || []).forEach((rowValues, index) => {
    if (isCompletelyBlankRow_(rowValues)) return;
    contexts.push(
      buildFuRowContext_(sheet, firstRow + index, columns, rowValues)
    );
  });
  const eventLocations = {};
  contexts.forEach(context => {
    if (!context.eventId) return;
    if (!eventLocations[context.eventId]) eventLocations[context.eventId] = [];
    eventLocations[context.eventId].push(context);
  });
  return {
    contexts,
    sheetRecords: [{
      sheetId: sheet.getSheetId(),
      sheetName: sheet.getName(),
      kind: 'FU',
      archived: false,
      fastFingerprint
    }],
    eventLocations,
    duplicateIds: Object.keys(eventLocations)
      .filter(eventId => eventLocations[eventId].length > 1),
    structuralActions: [],
    structuralConflicts: [],
    blockScan: null,
    fastFingerprint
  };
}

function buildMonthlySheetCalendarScanFromBlockScan_(
  sheet,
  columns,
  blockScan
) {
  const contexts = [];
  const structuralActions = [];
  const structuralConflicts = [];
  const firstRow = blockScan.firstRow || 2;
  const cancelledRows = readMonthlyCancelledRows_(
    sheet,
    columns,
    firstRow,
    (blockScan.values || []).length
  );
  const fastFingerprint = buildManagedSheetFastFingerprint_(
    'MONTHLY',
    columns,
    blockScan.values || [],
    firstRow,
    cancelledRows
  );
  blockScan.blocks.forEach(block => {
    block.patientRows.forEach(patient => {
      contexts.push(
        buildMonthlyRowContext_(
          sheet,
          patient,
          block,
          columns,
          false,
          cancelledRows
        )
      );
    });
  });
  blockScan.orphans.forEach(patient => {
    contexts.push(
      buildMonthlyRowContext_(
        sheet,
        patient,
        null,
        columns,
        false,
        cancelledRows
      )
    );
  });
  getMonthlyStructureIssues_(sheet, blockScan).forEach(issue => {
    const context = buildMonthlyStructureContext_(
      sheet,
      issue,
      columns,
      blockScan.rowValuesByRow[issue.row] || []
    );
    const target = {
      type: issue.type,
      eventId: issue.eventId || '',
      context,
      previous: null,
      structural: true,
      message: issue.message
    };
    if (issue.severity === 'action') structuralActions.push(target);
    else structuralConflicts.push(target);
  });
  blockScan.hybrids.forEach(item => {
    const issue = structuralConflicts.find(candidate => {
      return candidate.type === 'monthly_hybrid_date_event_id' &&
        candidate.context.row === item.row;
    });
    if (issue && item.eventId) contexts.push(issue.context);
  });
  const eventLocations = {};
  contexts.forEach(context => {
    if (!context.eventId) return;
    if (!eventLocations[context.eventId]) eventLocations[context.eventId] = [];
    eventLocations[context.eventId].push(context);
  });
  return {
    contexts,
    sheetRecords: [{
      sheetId: sheet.getSheetId(),
      sheetName: sheet.getName(),
      kind: 'MONTHLY',
      archived: false,
      fastFingerprint
    }],
    eventLocations,
    duplicateIds: Object.keys(eventLocations)
      .filter(eventId => eventLocations[eventId].length > 1),
    structuralActions,
    structuralConflicts,
    blockScan,
    cancelledRows,
    fastFingerprint
  };
}

function buildManagedSheetCalendarScan_(sheet) {
  const contexts = [];
  const structuralActions = [];
  const structuralConflicts = [];
  let blockScan = null;
  const sheetName = sheet.getName();
  const isFu = isMainTrackingSheetName_(sheetName);
  const isMonthly = isMonthlySheetName_(sheetName);
  if (!isFu && !isMonthly) {
    throw new Error(`${sheetName} 不是 FU 或有效月份刀表。`);
  }

  if (isFu) {
    const columns = getRequiredFuColumns_(sheet);
    const lastColumn = Math.max(
      getLastHeaderColumn_(sheet),
      ...Object.values(columns)
    );
    const rowCount = Math.max(0, sheet.getLastRow() - 1);
    const values = rowCount
      ? sheet.getRange(2, 1, rowCount, lastColumn).getValues()
      : [];
    return buildFuSheetCalendarScanFromValues_(
      sheet,
      columns,
      values,
      2
    );
  } else {
    const columns = getRequiredMonthlyColumns_(sheet);
    blockScan = scanMonthlyBlocks_(sheet, columns);
    return buildMonthlySheetCalendarScanFromBlockScan_(
      sheet,
      columns,
      blockScan
    );
  }
}

function buildGlobalEventIdLocationsLightweight_(spreadsheet) {
  const eventLocations = {};
  const addLocation = location => {
    const eventId = toCellText_(location && location.eventId);
    if (!eventId) return;
    if (!eventLocations[eventId]) eventLocations[eventId] = [];
    eventLocations[eventId].push(location);
  };
  spreadsheet.getSheets().forEach(sheet => {
    const sheetName = sheet.getName();
    const isFu = isMainTrackingSheetName_(sheetName);
    const isMonthly = isMonthlySheetName_(sheetName);
    if (!isFu && !isMonthly) return;
    const columns = isFu
      ? getRequiredFuColumns_(sheet)
      : getRequiredMonthlyColumns_(sheet);
    const rowCount = Math.max(0, sheet.getLastRow() - 1);
    if (!rowCount) return;
    sheet.getRange(2, columns.EVENT_ID, rowCount, 1)
      .getDisplayValues()
      .forEach((values, index) => {
        const eventId = toCellText_(values[0]);
        if (!eventId) return;
        addLocation({
          eventId,
          sheetId: sheet.getSheetId(),
          sheetName,
          kind: isFu ? 'FU' : 'MONTHLY',
          archived: false,
          row: index + 2
        });
      });
  });
  getAnnualArchiveEventIdLocations_(spreadsheet).forEach(addLocation);
  return {
    eventLocations,
    duplicateIds: Object.keys(eventLocations)
      .filter(eventId => eventLocations[eventId].length > 1)
  };
}

function getSortedRegistryEventIds_(entries) {
  return (entries || [])
    .map(entry => toCellText_(entry && entry.eventId))
    .filter(Boolean)
    .sort();
}

function buildGlobalEventIdLocationsFromRegistry_(
  baseline,
  sheet,
  sheetScan,
  pendingQueue
) {
  if (!baseline || !baseline.ok || !baseline.index) return null;
  if (pendingCalendarQueueHasWork_(pendingQueue)) return null;
  if ((baseline.entries || []).some(entry => {
    return Boolean(entry && (
      entry.pendingSync ||
      entry.pendingDelete ||
      entry.pendingResolution
    ));
  })) {
    return null;
  }

  const sheetId = Number(sheet.getSheetId());
  const metadata = getRegistrySheetMetadata_(baseline, sheetId);
  if (!metadata) return null;
  const previousEntries = (baseline.entries || []).filter(entry => {
    return Number(entry && entry.sheetId) === sheetId;
  });
  const currentContexts = (sheetScan.contexts || []).filter(context => {
    return Boolean(toCellText_(context && context.eventId));
  });
  const previousIds = getSortedRegistryEventIds_(previousEntries);
  const currentIds = getSortedRegistryEventIds_(currentContexts);
  if (JSON.stringify(previousIds) !== JSON.stringify(currentIds)) {
    return null;
  }

  const eventLocations = {};
  const addLocation = location => {
    const eventId = toCellText_(location && location.eventId);
    if (!eventId) return;
    if (!eventLocations[eventId]) eventLocations[eventId] = [];
    eventLocations[eventId].push(location);
  };
  (baseline.entries || []).forEach(entry => {
    if (Number(entry && entry.sheetId) === sheetId) return;
    addLocation(entry);
  });
  currentContexts.forEach(addLocation);
  return {
    eventLocations,
    duplicateIds: Object.keys(eventLocations)
      .filter(eventId => eventLocations[eventId].length > 1),
    baseline,
    source: 'registry_v2'
  };
}

function resolveGlobalEventIdLocationsForSync_(
  spreadsheet,
  sheet,
  sheetScan
) {
  const baseline = readCalendarRegistryStore_();
  const pendingQueue = readPendingCalendarQueue_();
  const indexed = buildGlobalEventIdLocationsFromRegistry_(
    baseline,
    sheet,
    sheetScan,
    pendingQueue
  );
  if (indexed) return indexed;
  const scanned = buildGlobalEventIdLocationsLightweight_(spreadsheet);
  return {
    ...scanned,
    baseline,
    source: 'live_sheet_scan'
  };
}

function buildManagedSheetFastState_(sheet) {
  const sheetName = sheet.getName();
  const isFu = isMainTrackingSheetName_(sheetName);
  const isMonthly = isMonthlySheetName_(sheetName);
  if (!isFu && !isMonthly) {
    throw new Error(`${sheetName} 不是 FU 或有效月份刀表。`);
  }
  const kind = isFu ? 'FU' : 'MONTHLY';
  const columns = isFu
    ? getRequiredFuColumns_(sheet)
    : getRequiredMonthlyColumns_(sheet);
  const lastColumn = Math.max(
    getLastHeaderColumn_(sheet),
    ...Object.values(columns)
  );
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const values = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues()
    : [];
  const cancelledRows = isMonthly
    ? readMonthlyCancelledRows_(sheet, columns, 2, values.length)
    : {};
  const fastFingerprint = buildManagedSheetFastFingerprint_(
    kind,
    columns,
    values,
    2,
    cancelledRows
  );
  const eventLocations = {};
  const addEventId = (eventId, row) => {
    const normalized = toCellText_(eventId);
    if (!normalized) return;
    if (!eventLocations[normalized]) eventLocations[normalized] = [];
    eventLocations[normalized].push(row);
  };
  values.forEach((rowValues, index) => {
    addEventId(
      getRowFieldValue_(rowValues, columns, 'EVENT_ID'),
      index + 2
    );
  });

  let pendingCreates = 0;
  let invalidExistingEvents = 0;
  let blockScan = null;
  let structuralActions = [];
  let structuralConflicts = [];
  if (isFu) {
    values.forEach(rowValues => {
      if (isCompletelyBlankRow_(rowValues)) return;
      const eventId = toCellText_(
        getRowFieldValue_(rowValues, columns, 'EVENT_ID')
      );
      const date = parseEventDate_(
        getRowFieldValue_(rowValues, columns, 'DATE')
      );
      const hasIdentity = Boolean(
        toCellText_(getRowFieldValue_(rowValues, columns, 'CHART_NO')) ||
        toCellText_(getRowFieldValue_(rowValues, columns, 'NAME'))
      );
      const valid = Boolean(date && hasIdentity);
      if (!eventId && valid) pendingCreates++;
      if (eventId && !valid) invalidExistingEvents++;
    });
  } else {
    blockScan = scanMonthlyBlocksFromValues_(
      sheet,
      columns,
      values,
      2,
      lastColumn,
      lastRow
    );
    const issues = getMonthlyStructureIssues_(sheet, blockScan);
    structuralActions = issues.filter(item => item.severity === 'action');
    structuralConflicts = issues.filter(item => item.severity !== 'action');
    const inspectPatient = (patient, hasBlock) => {
      const rowValues = patient.values || [];
      const eventId = toCellText_(
        getRowFieldValue_(rowValues, columns, 'EVENT_ID')
      );
      const hasIdentity = Boolean(
        toCellText_(getRowFieldValue_(rowValues, columns, 'CHART_NO')) ||
        toCellText_(getRowFieldValue_(rowValues, columns, 'NAME'))
      );
      const timeInfo = resolveCalendarTime_(
        getRowFieldValue_(rowValues, columns, 'TIME')
      );
      const valid = Boolean(
        hasBlock && hasIdentity && !timeInfo.errorMessage
      );
      if (!eventId && valid) pendingCreates++;
      if (eventId && !valid) invalidExistingEvents++;
    };
    blockScan.blocks.forEach(block => {
      block.patientRows.forEach(patient => inspectPatient(patient, true));
    });
    blockScan.orphans.forEach(patient => inspectPatient(patient, false));
    blockScan.hybrids.forEach(patient => {
      const eventId = toCellText_(patient.eventId);
      if (eventId) invalidExistingEvents++;
    });
  }

  return {
    kind,
    columns,
    lastColumn,
    lastRow,
    values,
    blockScan,
    cancelledRows,
    fastFingerprint,
    eventLocations,
    eventIds: Object.keys(eventLocations),
    duplicateIds: Object.keys(eventLocations)
      .filter(eventId => eventLocations[eventId].length > 1),
    pendingCreates,
    invalidExistingEvents,
    structuralActions,
    structuralConflicts
  };
}

function getRegistrySheetMetadata_(baseline, sheetId) {
  if (!baseline || !baseline.index) return null;
  return (baseline.index.sheets || []).find(item => {
    return Number(item.sheetId) === Number(sheetId);
  }) || null;
}

function assertManagedSheetRegistrySafeForReorder_(
  spreadsheet,
  sheet,
  actionLabel
) {
  const label = toCellText_(actionLabel) || '排序';
  const baseline = readCalendarRegistryStore_();
  if (!baseline.ok) {
    throw new Error(
      `${label}前尚未建立同步索引；請先執行「安裝／修復系統」。`
    );
  }

  const registryEventLocations = {};
  baseline.entries.forEach(entry => {
    const eventId = toCellText_(entry && entry.eventId);
    if (!eventId) return;
    if (!registryEventLocations[eventId]) {
      registryEventLocations[eventId] = [];
    }
    registryEventLocations[eventId].push(entry);
  });
  const registryDuplicateIds = Object.keys(registryEventLocations)
    .filter(eventId => registryEventLocations[eventId].length > 1);
  if (registryDuplicateIds.length) {
    throw new Error(
      `${label}前同步索引已有 ${registryDuplicateIds.length} 組重複 ` +
      'CalendarEventId；請先執行「檢查同步健康」。'
    );
  }

  const fastState = buildManagedSheetFastState_(sheet);
  if (fastState.duplicateIds.length) {
    throw new Error(
      `${sheet.getName()} 有 ${fastState.duplicateIds.length} 組重複 ` +
      `CalendarEventId；${label}未執行。`
    );
  }
  if (
    fastState.structuralActions.length ||
    fastState.structuralConflicts.length
  ) {
    throw new Error(
      `${sheet.getName()} 有 ${fastState.structuralActions.length} 項待修復結構、` +
      `${fastState.structuralConflicts.length} 項結構衝突；` +
      `${label}未執行。`
    );
  }

  const sheetId = Number(sheet.getSheetId());
  const previousEntries = baseline.entries.filter(entry => {
    return Number(entry.sheetId) === sheetId;
  });
  const previousById = {};
  previousEntries.forEach(entry => {
    if (entry.eventId) previousById[entry.eventId] = entry;
  });
  const currentById = {};
  fastState.eventIds.forEach(eventId => {
    currentById[eventId] = true;
  });

  const pendingEntries = previousEntries.filter(entry => {
    return Boolean(
      entry.pendingSync ||
      entry.pendingDelete ||
      entry.pendingResolution
    );
  });
  const missingFromSheet = previousEntries.filter(entry => {
    return entry.eventId && !currentById[entry.eventId];
  });
  const missingFromRegistry = Object.keys(currentById).filter(eventId => {
    return !previousById[eventId];
  });
  const sheetMetadata = getRegistrySheetMetadata_(
    baseline,
    sheet.getSheetId()
  );
  const fingerprintMatches = Boolean(
    sheetMetadata &&
    sheetMetadata.fastFingerprint &&
    sheetMetadata.fastFingerprint === fastState.fastFingerprint
  );
  let sheetScan = null;
  let changedEntries = [];
  if (!fingerprintMatches) {
    sheetScan = fastState.kind === 'FU'
      ? buildFuSheetCalendarScanFromValues_(
        sheet,
        fastState.columns,
        fastState.values,
        2
      )
      : buildMonthlySheetCalendarScanFromBlockScan_(
        sheet,
        fastState.columns,
        fastState.blockScan
      );
    const scannedById = {};
    sheetScan.contexts.forEach(context => {
      if (context.eventId) scannedById[context.eventId] = context;
    });
    changedEntries = Object.keys(scannedById).filter(eventId => {
      const previous = previousById[eventId];
      const current = scannedById[eventId];
      if (!previous) return false;
      return previous.rowHash !== current.rowHash ||
        previous.blockKey !== current.blockKey ||
        Boolean(previous.valid) !== Boolean(current.valid);
    });
  }

  const issueCount =
    pendingEntries.length +
    fastState.pendingCreates +
    fastState.invalidExistingEvents +
    missingFromSheet.length +
    missingFromRegistry.length +
    changedEntries.length;
  if (issueCount) {
    throw new Error(
      `${label}前 ${sheet.getName()} 仍有 ${issueCount} 項同步差異` +
      `（待重試 ${pendingEntries.length}、待建立 ${fastState.pendingCreates}、` +
      `無效既有事件 ${fastState.invalidExistingEvents}、索引差異 ` +
      `${missingFromSheet.length + missingFromRegistry.length + changedEntries.length}）；` +
      '請先執行「檢查同步健康」或「同步待處理變更」。'
    );
  }

  return {
    ok: true,
    baseline,
    sheetScan,
    fastState,
    usedFastFingerprint: fingerprintMatches,
    // 排序不會新增或修改 Event ID。以已驗證的 V2 索引做跨表
    // 唯一性基準，只讀目前表單確認內容，避免日常排序掃描全部月份。
    globalIdScan: {
      eventLocations: registryEventLocations,
      duplicateIds: []
    }
  };
}

function buildCurrentCalendarScan_(spreadsheet) {
  assertNoLegacyArchiveState_();
  const contexts = [];
  const sheetRecords = [];
  const structuralActions = [];
  const structuralConflicts = [];
  const fu = getMainTrackingSheet_(spreadsheet);
  const managedSheets = [];
  if (fu) managedSheets.push(fu);
  spreadsheet.getSheets()
    .filter(sheet => isMonthlySheetName_(sheet.getName()))
    .forEach(sheet => managedSheets.push(sheet));
  managedSheets.forEach(sheet => {
    const scan = buildManagedSheetCalendarScan_(sheet);
    scan.contexts.forEach(context => contexts.push(context));
    scan.sheetRecords.forEach(record => sheetRecords.push(record));
    scan.structuralActions.forEach(action => structuralActions.push(action));
    scan.structuralConflicts.forEach(conflict => {
      structuralConflicts.push(conflict);
    });
  });

  const eventLocations = {};
  contexts.forEach(context => {
    if (!context.eventId) return;
    if (!eventLocations[context.eventId]) eventLocations[context.eventId] = [];
    eventLocations[context.eventId].push(context);
  });
  const archiveLocations = getAnnualArchiveEventIdLocations_(spreadsheet);
  archiveLocations.forEach(location => {
    if (!eventLocations[location.eventId]) {
      eventLocations[location.eventId] = [];
    }
    eventLocations[location.eventId].push(location);
  });
  const duplicateIds = Object.keys(eventLocations)
    .filter(eventId => eventLocations[eventId].length > 1);
  return {
    contexts,
    sheetRecords,
    eventLocations,
    duplicateIds,
    structuralActions,
    structuralConflicts,
    archiveLocations
  };
}

function isSystemManagedCalendarEvent_(event) {
  const titleSlots = toCellText_(event && event.summary)
    .split(/\s*\|\s*/);
  if (titleSlots.length !== 3) return false;
  const description = toCellText_(event && event.description);
  return /(?:^|\n)(?:醫院|Hospital|TEL|Plan|GA|原報到時間|Time note):/m
    .test(description);
}

function listConfiguredCalendarEvents_() {
  const calendarId = getConfiguredCalendarId_();
  if (!calendarId) {
    return {
      ok: false,
      status: 'no_calendar_id',
      message: 'Calendar ID 尚未設定。',
      events: []
    };
  }
  if (!isCalendarAdvancedServiceAvailable_()) {
    return {
      ok: false,
      status: 'no_calendar_service',
      message: 'Calendar Advanced Service 尚未啟用。',
      events: []
    };
  }
  const events = [];
  let pageToken = '';
  let pageCount = 0;
  try {
    do {
      const options = {
        maxResults: 2500,
        showDeleted: false,
        singleEvents: true
      };
      if (pageToken) options.pageToken = pageToken;
      const response = Calendar.Events.list(calendarId, options) || {};
      (response.items || []).forEach(event => events.push(event));
      pageToken = toCellText_(response.nextPageToken);
      pageCount++;
      if (pageCount > 100) {
        throw new Error('Calendar 事件分頁超過安全上限。');
      }
    } while (pageToken);
  } catch (err) {
    return {
      ok: false,
      status: 'calendar_list_failed',
      message: err.message || String(err),
      events: []
    };
  }
  return {
    ok: true,
    status: 'listed',
    message: '',
    calendarId,
    events
  };
}

function analyzeLiveCalendarState_(scan) {
  const actions = [];
  const conflicts = [];
  const listed = listConfiguredCalendarEvents_();
  if (!listed.ok) {
    conflicts.push({
      type: listed.status,
      eventId: '',
      context: null,
      previous: null,
      message: listed.message
    });
    return {
      ok: false,
      actions,
      conflicts,
      listedEvents: [],
      unmanagedEvents: []
    };
  }

  const listedById = {};
  listed.events.forEach(event => {
    const eventId = toCellText_(event && event.id);
    if (eventId) listedById[eventId] = event;
  });
  const duplicateIds = {};
  (scan.duplicateIds || []).forEach(eventId => {
    duplicateIds[eventId] = true;
  });

  scan.contexts.forEach(context => {
    if (
      context.archived ||
      !context.eventId ||
      !context.valid ||
      duplicateIds[context.eventId]
    ) {
      return;
    }
    let event = listedById[context.eventId] || null;
    if (!event) {
      try {
        event = Calendar.Events.get(listed.calendarId, context.eventId);
      } catch (err) {
        if (isCalendarNotFoundError_(err)) {
          actions.push({
            type: 'calendar_missing',
            eventId: context.eventId,
            context,
            previous: null
          });
        } else {
          conflicts.push({
            type: 'calendar_read_failed',
            eventId: context.eventId,
            context,
            previous: null,
            message:
              `Calendar 事件回讀失敗：${err.message || String(err)}`
          });
        }
        return;
      }
    }
    const expected = buildCalendarResource_(context);
    if (!calendarEventMatchesResource_(event, expected)) {
      actions.push({
        type: 'calendar_drift',
        eventId: context.eventId,
        context,
        previous: null
      });
    }
  });

  const canonicalIds = {};
  Object.keys(scan.eventLocations || {}).forEach(eventId => {
    canonicalIds[eventId] = true;
  });
  const unmanagedEvents = [];
  listed.events.forEach(event => {
    const eventId = toCellText_(event && event.id);
    if (!eventId || canonicalIds[eventId]) return;
    if (!isSystemManagedCalendarEvent_(event)) {
      unmanagedEvents.push(event);
      return;
    }
    actions.push({
      type: 'calendar_orphan',
      eventId,
      context: null,
      previous: {
        sheetName: 'Calendar',
        row: '',
        kind: 'CALENDAR'
      }
    });
  });
  return {
    ok: conflicts.length === 0,
    actions,
    conflicts,
    listedEvents: listed.events,
    unmanagedEvents
  };
}

function mergeCalendarAnalyses_(registryAnalysis, liveAnalysis) {
  const base = registryAnalysis || {
    actions: [],
    conflicts: [],
    adoptions: [],
    ignoredRemovedSheets: [],
    duplicates: []
  };
  const actions = (base.actions || []).slice();
  const existingEventActions = {};
  actions.forEach(action => {
    if (action.eventId) existingEventActions[action.eventId] = true;
  });
  (liveAnalysis.actions || []).forEach(action => {
    if (action.eventId && existingEventActions[action.eventId]) {
      return;
    }
    actions.push(action);
    if (action.eventId) existingEventActions[action.eventId] = true;
  });
  return {
    ...base,
    actions,
    conflicts: (base.conflicts || []).concat(
      liveAnalysis.conflicts || []
    )
  };
}

function buildCalendarHealthPreview_(spreadsheet, options) {
  const settings = options || {};
  const registryResult = reconcileCalendarRegistry_(spreadsheet, {
    ...settings,
    apply: false
  });
  const scan = registryResult.scan || buildCurrentCalendarScan_(spreadsheet);
  const liveAnalysis = analyzeLiveCalendarState_(scan);
  const registryAnalysis = registryResult.analysis || {
    actions: [],
    conflicts: [],
    adoptions: [],
    ignoredRemovedSheets: [],
    duplicates: scan.duplicateIds || []
  };
  const analysis = mergeCalendarAnalyses_(
    registryAnalysis,
    liveAnalysis
  );
  const pendingQueue = readPendingCalendarQueue_();
  return {
    ...registryResult,
    ok:
      analysis.conflicts.length === 0 &&
      (scan.duplicateIds || []).length === 0,
    scan,
    analysis,
    liveAnalysis,
    pendingQueue,
    results: []
  };
}

function applyLiveCalendarActions_(scan, liveAnalysis, confirmedOrphanIds) {
  const confirmed = {};
  (confirmedOrphanIds || []).forEach(eventId => {
    confirmed[toCellText_(eventId)] = true;
  });
  const results = [];
  (liveAnalysis.actions || []).forEach(action => {
    let result = null;
    if (
      action.type === 'calendar_drift' ||
      action.type === 'calendar_missing'
    ) {
      result = syncManagedContext_(
        action.context,
        scan.eventLocations
      );
    } else if (
      action.type === 'calendar_orphan' &&
      confirmed[action.eventId]
    ) {
      result = deleteCalendarEvent_(action.eventId);
    }
    if (result) results.push({ action, result });
  });
  return results;
}

function getHealthActionKey_(action) {
  const value = action || {};
  const source = value.context || value.previous || {};
  return [
    value.type || '',
    value.eventId || '',
    source.sheetId || '',
    source.row || ''
  ].join('|');
}

function runFullCalendarReconcile_(spreadsheet, options) {
  const settings = options || {};
  const preflight = buildCalendarHealthPreview_(spreadsheet, {
    ...settings,
    apply: false
  });
  if (
    (preflight.analysis.conflicts || []).length ||
    (preflight.scan.duplicateIds || []).length
  ) {
    return preflight;
  }

  const registryResult = reconcileCalendarRegistry_(spreadsheet, {
    ...settings,
    apply: true
  });
  const registryFailures = (registryResult.results || [])
    .filter(item => !item.result || !item.result.ok);
  if (
    registryFailures.length ||
    (registryResult.analysis && registryResult.analysis.conflicts || []).length
  ) {
    return registryResult;
  }

  const scan = buildCurrentCalendarScan_(spreadsheet);
  if (scan.duplicateIds.length) {
    const duplicateAnalysis = {
      actions: [],
      conflicts: scan.duplicateIds.map(eventId => ({
        type: 'duplicate_event_id',
        eventId,
        context: (scan.eventLocations[eventId] || [])[0] || null,
        previous: null,
        message: 'CalendarEventId 重複，已停止 Calendar 全量同步。'
      })),
      adoptions: [],
      ignoredRemovedSheets: [],
      duplicates: scan.duplicateIds
    };
    annotateAnalysisConflicts_(duplicateAnalysis);
    return {
      ok: false,
      scan,
      analysis: duplicateAnalysis,
      results: registryResult.results || []
    };
  }

  const liveAnalysis = analyzeLiveCalendarState_(scan);
  const registryAnalysis = registryResult.analysis || {
    actions: [],
    conflicts: [],
    adoptions: [],
    ignoredRemovedSheets: [],
    duplicates: []
  };
  const analysis = mergeCalendarAnalyses_(
    registryAnalysis,
    liveAnalysis
  );
  if ((liveAnalysis.conflicts || []).length) {
    return {
      ok: false,
      scan,
      analysis,
      liveAnalysis,
      results: registryResult.results || []
    };
  }

  const liveResults = applyLiveCalendarActions_(
    scan,
    liveAnalysis,
    settings.confirmedOrphanIds || []
  );
  const results = (registryResult.results || []).concat(liveResults);
  const refreshed = buildCurrentCalendarScan_(spreadsheet);
  if (!refreshed.duplicateIds.length) {
    const retained = readCalendarRegistryStore_();
    const retainedEntries = retained.ok
      ? retained.entries.filter(entry => {
        return Boolean(
          entry.pendingSync ||
          entry.pendingDelete ||
          entry.pendingResolution
        );
      })
      : [];
    liveResults.forEach(item => {
      if (!item || !item.result || item.result.ok || !item.action) return;
      const context = item.action.context;
      if (!context || !context.eventId) return;
      retainedEntries.push({
        ...toRegistryEntry_(context),
        pendingSync: true,
        pendingError:
          item.result.status || item.result.message || 'sync_failed'
      });
    });
    writeCalendarRegistryStore_(refreshed, retainedEntries);
  }
  const appliedKeys = {};
  results.forEach(item => {
    appliedKeys[getHealthActionKey_(item.action)] = true;
  });
  const pending = (analysis.actions || []).filter(action => {
    return !appliedKeys[getHealthActionKey_(action)];
  });
  return {
    ok:
      results.every(item => item.result && item.result.ok) &&
      analysis.conflicts.length === 0 &&
      refreshed.duplicateIds.length === 0 &&
      pending.length === 0,
    scan: refreshed,
    analysis,
    liveAnalysis,
    results
  };
}

function toRegistryEntry_(context) {
  return {
    eventId: context.eventId,
    sheetId: context.sheetId,
    sheetName: context.sheetName,
    kind: context.kind,
    archived: Boolean(context.archived),
    row: context.row,
    blockKey: context.blockKey,
    rowHash: context.rowHash,
    bindingHash: toCellText_(context.bindingHash),
    valid: Boolean(context.valid),
    invalidReason: context.invalidReason || '',
    pendingSync: false,
    pendingDelete: false,
    pendingResolution: false,
    pendingResolutionReason: '',
    pendingError: ''
  };
}

function toArchiveRegistryEntry_(location) {
  const sourceMonth = toCellText_(location && location.sourceMonth);
  const eventId = toCellText_(location && location.eventId);
  return {
    eventId,
    sheetId: Number(location && location.sheetId),
    sheetName: toCellText_(location && location.sheetName),
    kind: 'ARCHIVE',
    archived: true,
    row: Number(location && location.row),
    sourceMonth,
    blockKey: sourceMonth,
    bindingHash: '',
    rowHash: sha256Hex_(
      JSON.stringify([
        eventId,
        Number(location && location.sheetId),
        Number(location && location.row),
        sourceMonth
      ])
    ),
    valid: true,
    invalidReason: '',
    pendingSync: false,
    pendingDelete: false,
    pendingResolution: false,
    pendingResolutionReason: '',
    pendingError: ''
  };
}

function normalizeCalendarRegistryReasonCode_(value, fallback) {
  const text = toCellText_(value);
  if (!text) return toCellText_(fallback);
  return /^[A-Za-z0-9_.:-]{1,120}$/.test(text)
    ? text
    : toCellText_(fallback) || 'legacy_error_redacted';
}

function normalizeCalendarRegistryEntry_(entry) {
  const value = entry || {};
  return {
    ...value,
    eventId: toCellText_(value.eventId),
    sheetId: Number(value.sheetId),
    sheetName: toCellText_(value.sheetName),
    kind: toCellText_(value.kind),
    archived: Boolean(value.archived),
    row: Number(value.row),
    blockKey: toCellText_(value.blockKey),
    rowHash: toCellText_(value.rowHash),
    bindingHash: toCellText_(value.bindingHash),
    valid: Boolean(value.valid),
    invalidReason: toCellText_(value.invalidReason),
    pendingSync: Boolean(value.pendingSync),
    pendingDelete: Boolean(value.pendingDelete),
    pendingResolution: Boolean(value.pendingResolution),
    pendingResolutionReason: normalizeCalendarRegistryReasonCode_(
      value.pendingResolutionReason,
      value.pendingResolution ? 'pending_resolution' : ''
    ),
    pendingError: normalizeCalendarRegistryReasonCode_(
      value.pendingError,
      value.pendingError ? 'legacy_error_redacted' : ''
    )
  };
}

function getRegistryChunkKeys_(propertyValues) {
  const properties = propertyValues ||
    PropertiesService.getScriptProperties().getProperties();
  return Object.keys(properties)
    .filter(key => key.indexOf(CALENDAR_REGISTRY_CHUNK_PREFIX) === 0);
}

function packRegistryEntries_(entries, baseKey) {
  const chunks = [];
  let current = [];
  (entries || []).forEach(entry => {
    const candidate = current.concat([entry]);
    const serialized = JSON.stringify(candidate);
    if (serialized.length > CALENDAR_REGISTRY_CHUNK_SIZE && current.length) {
      chunks.push(current);
      current = [entry];
    } else {
      current = candidate;
    }
  });
  if (current.length || !chunks.length) chunks.push(current);
  return chunks.map((chunk, index) => ({
    key: `${baseKey}_${index}`,
    value: JSON.stringify(chunk)
  }));
}

function writeCalendarRegistryStore_(
  scan,
  retainedEntries,
  retainedSheetMetadata
) {
  if ((scan.duplicateIds || []).length) {
    throw new Error(
      `CalendarEventId 重複：${scan.duplicateIds.join('、')}；` +
      'V2 索引未寫入。'
    );
  }
  const properties = PropertiesService.getScriptProperties();
  const retainedById = {};
  (retainedEntries || []).forEach(entry => {
    const normalized = normalizeCalendarRegistryEntry_(entry);
    if (normalized.eventId) {
      retainedById[normalized.eventId] = normalized;
    }
  });
  const currentEntries = scan.contexts
    .filter(context => context.eventId)
    .map(toRegistryEntry_)
    .filter(entry => !retainedById[entry.eventId]);
  (scan.archiveLocations || [])
    .map(toArchiveRegistryEntry_)
    .filter(entry => entry.eventId && !retainedById[entry.eventId])
    .forEach(entry => currentEntries.push(entry));
  const currentIds = {};
  currentEntries.forEach(entry => {
    currentIds[entry.eventId] = true;
  });
  Object.keys(retainedById).forEach(eventId => {
    const entry = retainedById[eventId];
    if (!currentIds[entry.eventId]) {
      currentEntries.push(entry);
      currentIds[entry.eventId] = true;
    }
  });

  const entriesBySheet = {};
  currentEntries.forEach(entry => {
    const key = String(entry.sheetId);
    if (!entriesBySheet[key]) entriesBySheet[key] = [];
    entriesBySheet[key].push(entry);
  });
  const sheetRecordsById = {};
  (retainedSheetMetadata || []).forEach(record => {
    if (!record || record.sheetId === undefined) return;
    sheetRecordsById[String(record.sheetId)] = { ...record };
  });
  (scan.sheetRecords || []).forEach(record => {
    if (!record || record.sheetId === undefined) return;
    sheetRecordsById[String(record.sheetId)] = { ...record };
  });
  const sheetIndex = [];
  const nextPropertyValues = {};
  const indexedSheetIds = Array.from(new Set(
    Object.keys(entriesBySheet).concat(Object.keys(sheetRecordsById))
  )).sort((left, right) => Number(left) - Number(right));
  indexedSheetIds.forEach(sheetId => {
    const sheetEntries = entriesBySheet[sheetId] || [];
    const chunks = packRegistryEntries_(
      sheetEntries,
      `${CALENDAR_REGISTRY_CHUNK_PREFIX}${sheetId}`
    );
    if (sheetEntries.length) {
      chunks.forEach(chunk => {
        nextPropertyValues[chunk.key] = chunk.value;
      });
    }
    const record = sheetRecordsById[sheetId] || {};
    const representative = sheetEntries[0] || record;
    sheetIndex.push({
      sheetId: Number(sheetId),
      sheetName: record.sheetName || representative.sheetName || '',
      kind: record.kind || representative.kind || '',
      archived: Boolean(
        record.archived === undefined
          ? representative.archived
          : record.archived
      ),
      keys: sheetEntries.length
        ? chunks.map(chunk => chunk.key)
        : [],
      count: sheetEntries.length,
      fastFingerprint: toCellText_(record.fastFingerprint)
    });
  });

  const index = {
    version: CALENDAR_REGISTRY_VERSION,
    updatedAt: new Date().toISOString(),
    sheets: sheetIndex,
    eventCount: currentEntries.length,
    fingerprint: sha256Hex_(
      JSON.stringify(
        currentEntries
          .map(entry => [
            entry.eventId,
            entry.sheetId,
            entry.row,
             entry.blockKey,
             entry.rowHash,
             entry.bindingHash,
             Boolean(entry.pendingSync),
             Boolean(entry.pendingDelete),
             Boolean(entry.pendingResolution),
             entry.pendingResolutionReason
           ])
          .sort()
      )
    )
  };
  nextPropertyValues[CALENDAR_REGISTRY_INDEX_KEY] = JSON.stringify(index);

  // Script Properties 是遠端服務。一次 setProperties 比逐分頁
  // deleteProperty/setProperty 快很多，也避免排序功能把時間耗在索引 I/O。
  const existingPropertyValues = properties.getProperties();
  properties.setProperties(nextPropertyValues);
  getRegistryChunkKeys_(existingPropertyValues)
    .filter(key => !Object.prototype.hasOwnProperty.call(
      nextPropertyValues,
      key
    ))
    .forEach(key => properties.deleteProperty(key));
  return index;
}

function readCalendarRegistryStore_() {
  const properties = PropertiesService.getScriptProperties();
  const propertyValues = properties.getProperties();
  const index = safeJsonParse_(
    propertyValues[CALENDAR_REGISTRY_INDEX_KEY],
    null
  );
  if (!index || Number(index.version) !== CALENDAR_REGISTRY_VERSION) {
    return {
      ok: false,
      missing: true,
      entries: [],
      index: null,
      message: '尚未建立 CALENDAR_ROW_REGISTRY_V2。'
    };
  }
  const entries = [];
  try {
    (index.sheets || []).forEach(sheetInfo => {
      (sheetInfo.keys || []).forEach(key => {
        const chunk = safeJsonParse_(propertyValues[key], null);
        if (!Array.isArray(chunk)) {
          throw new Error(`索引分段 ${key} 遺失或損壞。`);
        }
        chunk.forEach(entry => {
          entries.push(normalizeCalendarRegistryEntry_(entry));
        });
      });
    });
  } catch (err) {
    return {
      ok: false,
      missing: false,
      entries: [],
      index,
      message: err.message || String(err)
    };
  }
  return { ok: true, missing: false, entries, index, message: '' };
}

function emptyPendingCalendarQueue_() {
  return {
    version: CALENDAR_PENDING_QUEUE_VERSION,
    fullScan: false,
    allowMissingDeletes: false,
    sheets: {},
    reasons: [],
    updatedAt: ''
  };
}

function readPendingCalendarQueue_() {
  const raw = PropertiesService.getScriptProperties()
    .getProperty(CALENDAR_PENDING_QUEUE_PROPERTY);
  const parsed = safeJsonParse_(raw, null);
  if (
    !parsed ||
    Number(parsed.version) !== CALENDAR_PENDING_QUEUE_VERSION
  ) {
    return emptyPendingCalendarQueue_();
  }
  return {
    ...emptyPendingCalendarQueue_(),
    ...parsed,
    sheets: parsed.sheets && typeof parsed.sheets === 'object'
      ? parsed.sheets
      : {},
    reasons: Array.isArray(parsed.reasons) ? parsed.reasons : []
  };
}

function normalizePendingCalendarReason_(reason) {
  const value = toCellText_(reason).toUpperCase();
  return /^[A-Z0-9_:-]{1,64}$/.test(value) ? value : 'UNSPECIFIED';
}

function pendingCalendarQueueHasWork_(queue) {
  const value = queue || emptyPendingCalendarQueue_();
  return Boolean(
    value.fullScan ||
    Object.keys(value.sheets || {}).some(sheetId => {
      return Array.isArray(value.sheets[sheetId]) &&
        value.sheets[sheetId].length > 0;
    })
  );
}

function summarizePendingCalendarQueue_(queue) {
  const value = queue || emptyPendingCalendarQueue_();
  const sheetIds = Object.keys(value.sheets || {}).filter(sheetId => {
    return Array.isArray(value.sheets[sheetId]) &&
      value.sheets[sheetId].length > 0;
  });
  const rowCount = sheetIds.reduce((total, sheetId) => {
    return total + value.sheets[sheetId].length;
  }, 0);
  return {
    hasWork: pendingCalendarQueueHasWork_(value),
    fullScan: Boolean(value.fullScan),
    sheetCount: sheetIds.length,
    rowCount
  };
}

function normalizePendingCalendarRequest_(request) {
  const value = request || {};
  const rows = Array.from(new Set((value.rows || [])
    .map(Number)
    .filter(row => Number.isInteger(row) && row >= 2)))
    .sort((left, right) => left - right);
  return {
    fullScan: Boolean(value.fullScan),
    allowMissingDeletes: Boolean(value.allowMissingDeletes),
    sheetId: Number(value.sheetId) || 0,
    rows,
    reason: normalizePendingCalendarReason_(value.reason)
  };
}

function mergePendingCalendarRequest_(queue, request) {
  const target = queue || emptyPendingCalendarQueue_();
  const value = normalizePendingCalendarRequest_(request);
  target.fullScan = target.fullScan || value.fullScan;
  if (target.fullScan) target.sheets = {};
  target.allowMissingDeletes =
    target.allowMissingDeletes || value.allowMissingDeletes;
  if (value.reason && target.reasons.indexOf(value.reason) === -1) {
    target.reasons.push(value.reason);
  }
  if (!target.fullScan && value.sheetId && value.rows.length) {
    const key = String(value.sheetId);
    const current = Array.isArray(target.sheets[key])
      ? target.sheets[key]
      : [];
    const merged = Array.from(new Set(current.concat(value.rows)))
      .sort((left, right) => left - right);
    const rowCount = Object.values(target.sheets)
      .reduce((total, rows) => total + (rows || []).length, 0) +
      merged.length - current.length;
    if (rowCount > CALENDAR_PENDING_QUEUE_MAX_ROWS) {
      target.fullScan = true;
      target.sheets = {};
    } else {
      target.sheets[key] = merged;
    }
  }
  target.updatedAt = new Date().toISOString();
  return target;
}

function withPendingCalendarQueueLock_(callback) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(CALENDAR_PENDING_QUEUE_LOCK_WAIT_MS)) {
    return { acquired: false, value: null };
  }
  try {
    return { acquired: true, value: callback() };
  } finally {
    lock.releaseLock();
  }
}

function enqueuePendingCalendarSync_(request) {
  const properties = PropertiesService.getScriptProperties();
  const locked = withPendingCalendarQueueLock_(() => {
    const queue = mergePendingCalendarRequest_(
      readPendingCalendarQueue_(),
      request
    );
    properties.setProperty(
      CALENDAR_PENDING_QUEUE_PROPERTY,
      JSON.stringify(queue)
    );
    return queue;
  });
  if (locked.acquired) return locked.value;

  const fallback = mergePendingCalendarRequest_(
    emptyPendingCalendarQueue_(),
    {
      fullScan: true,
      allowMissingDeletes: request && request.allowMissingDeletes,
      reason: 'QUEUE_LOCK_FALLBACK'
    }
  );
  properties.setProperty(
    CALENDAR_PENDING_QUEUE_PROPERTY,
    JSON.stringify(fallback)
  );
  return fallback;
}

function getPendingCalendarRetryTriggers_() {
  return ScriptApp.getProjectTriggers().filter(trigger => {
    return trigger.getHandlerFunction() ===
      CALENDAR_PENDING_RETRY_HANDLER;
  });
}

function clearPendingCalendarRetryTriggers_() {
  const triggers = getPendingCalendarRetryTriggers_();
  triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  return triggers.length;
}

function ensurePendingCalendarRetryScheduled_(spreadsheet) {
  if (!pendingCalendarQueueHasWork_(readPendingCalendarQueue_())) {
    return { ok: true, scheduled: false, empty: true };
  }
  const locked = withPendingCalendarQueueLock_(() => {
    const properties = PropertiesService.getScriptProperties();
    if (
      spreadsheet &&
      typeof spreadsheet.getId === 'function'
    ) {
      properties.setProperty(
        CALENDAR_PENDING_SPREADSHEET_ID_PROPERTY,
        spreadsheet.getId()
      );
    }
    const existing = getPendingCalendarRetryTriggers_();
    if (existing.length) {
      return {
        ok: true,
        scheduled: false,
        existing: existing.length
      };
    }
    ScriptApp.newTrigger(CALENDAR_PENDING_RETRY_HANDLER)
      .timeBased()
      .after(CALENDAR_PENDING_RETRY_DELAY_MS)
      .create();
    return { ok: true, scheduled: true, existing: 0 };
  });
  if (locked.acquired) return locked.value;
  const properties = PropertiesService.getScriptProperties();
  if (spreadsheet && typeof spreadsheet.getId === 'function') {
    properties.setProperty(
      CALENDAR_PENDING_SPREADSHEET_ID_PROPERTY,
      spreadsheet.getId()
    );
  }
  const existing = getPendingCalendarRetryTriggers_();
  if (existing.length) {
    return {
      ok: true,
      scheduled: false,
      existing: existing.length,
      lockFallback: true
    };
  }
  ScriptApp.newTrigger(CALENDAR_PENDING_RETRY_HANDLER)
    .timeBased()
    .after(CALENDAR_PENDING_RETRY_DELAY_MS)
    .create();
  return {
    ok: true,
    scheduled: true,
    existing: 0,
    lockFallback: true
  };
}

function getPendingCalendarSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = toCellText_(
    properties.getProperty(
      CALENDAR_PENDING_SPREADSHEET_ID_PROPERTY
    )
  );
  if (
    spreadsheetId &&
    SpreadsheetApp.openById &&
    typeof SpreadsheetApp.openById === 'function'
  ) {
    return SpreadsheetApp.openById(spreadsheetId);
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('找不到待重試同步所屬的試算表。');
  }
  return active;
}

function takePendingCalendarQueue_() {
  const properties = PropertiesService.getScriptProperties();
  const locked = withPendingCalendarQueueLock_(() => {
    const queue = readPendingCalendarQueue_();
    properties.deleteProperty(CALENDAR_PENDING_QUEUE_PROPERTY);
    return queue;
  });
  return locked.acquired ? locked.value : null;
}

function restorePendingCalendarQueue_(queue) {
  if (!queue) return;
  if (queue.fullScan) {
    enqueuePendingCalendarSync_({
      fullScan: true,
      allowMissingDeletes: queue.allowMissingDeletes,
      reason: 'RETRY_AFTER_FAILURE'
    });
    return;
  }
  Object.keys(queue.sheets || {}).forEach(sheetId => {
    enqueuePendingCalendarSync_({
      sheetId: Number(sheetId),
      rows: queue.sheets[sheetId],
      reason: 'RETRY_AFTER_FAILURE'
    });
  });
}

function applyPendingCalendarQueue_(spreadsheet, queue) {
  if (!queue) return { ok: true, empty: true, results: [] };
  const hasRows = Object.keys(queue.sheets || {}).length > 0;
  if (!queue.fullScan && !hasRows) {
    return { ok: true, empty: true, results: [] };
  }
  if (queue.fullScan) {
    const result = reconcileCalendarRegistry_(spreadsheet, {
      apply: true,
      changeType: queue.allowMissingDeletes
        ? 'REMOVE_ROW'
        : 'QUEUED',
      allowMissingDeletes: Boolean(queue.allowMissingDeletes)
    });
    const retryableFailures = (result.results || []).filter(item => {
      return !item.result || !item.result.ok;
    });
    if (!retryableFailures.length) {
      return {
        ...result,
        ok: true,
        pendingResolutionCount:
          (result.analysis && result.analysis.conflicts || []).length
      };
    }
    return result;
  }
  const results = [];
  Object.keys(queue.sheets).forEach(sheetId => {
    const sheet = getSheetById_(spreadsheet, Number(sheetId));
    if (!sheet) return;
    results.push(
      syncManagedRowsAt_(spreadsheet, sheet, queue.sheets[sheetId])
    );
  });
  return {
    ok: results.every(result => result && result.ok),
    empty: false,
    results
  };
}

function drainPendingCalendarQueue_(spreadsheet) {
  const queue = takePendingCalendarQueue_();
  if (!queue) return { ok: true, empty: true, results: [] };
  try {
    const result = applyPendingCalendarQueue_(spreadsheet, queue);
    if (!result || result.ok === false) restorePendingCalendarQueue_(queue);
    return result;
  } catch (err) {
    restorePendingCalendarQueue_(queue);
    throw err;
  }
}

function annotatePendingCalendarRetryFailure_(
  spreadsheet,
  queue,
  errorMessage
) {
  if (!spreadsheet || !queue || queue.fullScan) return 0;
  let annotated = 0;
  Object.keys(queue.sheets || {}).forEach(sheetId => {
    const sheet = getSheetById_(spreadsheet, Number(sheetId));
    if (!sheet) return;
    const contexts = getManagedContextsForRows_(
      sheet,
      queue.sheets[sheetId]
    );
    contexts.forEach(context => {
      const note = toCellText_(getContextNoteCell_(context).getNote());
      const syncMarker = getSystemNoteBlockMarkers_(
        CALENDAR_SYNC_NOTE_PREFIX
      ).start;
      const conflictMarker = getSystemNoteBlockMarkers_(
        CALENDAR_CONFLICT_NOTE_PREFIX
      ).start;
      if (
        note.indexOf(syncMarker) !== -1 ||
        note.indexOf(conflictMarker) !== -1
      ) {
        return;
      }
      setContextSyncNote_(
        context,
        `自動重試仍未完成${errorMessage
          ? `：${errorMessage}`
          : ''
        }。請先執行「修復選取列同步」；` +
          '若有多列問題，請執行「同步待處理變更」。',
        false
      );
      annotated++;
    });
  });
  return annotated;
}

function retryPendingCalendarSync() {
  clearPendingCalendarRetryTriggers_();
  const queued = readPendingCalendarQueue_();
  if (!pendingCalendarQueueHasWork_(queued)) {
    PropertiesService.getScriptProperties().deleteProperty(
      CALENDAR_PENDING_SPREADSHEET_ID_PROPERTY
    );
    return { ok: true, empty: true, results: [] };
  }

  let spreadsheet;
  try {
    spreadsheet = getPendingCalendarSpreadsheet_();
  } catch (err) {
    console.error(err && err.stack ? err.stack : err);
    return { ok: false, status: 'spreadsheet_unavailable', error: err };
  }

  const attempt = tryWithCalendarSyncLock_(() => {
    const queue = takePendingCalendarQueue_();
    if (!queue) {
      return { ok: false, status: 'queue_lock_busy', queue: queued };
    }
    try {
      const result = applyPendingCalendarQueue_(spreadsheet, queue);
      if (!result || result.ok === false) {
        restorePendingCalendarQueue_(queue);
        annotatePendingCalendarRetryFailure_(
          spreadsheet,
          queue,
          '資料或 Calendar 狀態仍需確認'
        );
        return {
          ok: false,
          status: 'retry_failed',
          queue,
          result
        };
      }
      return { ok: true, status: 'retried', queue, result };
    } catch (err) {
      restorePendingCalendarQueue_(queue);
      annotatePendingCalendarRetryFailure_(
        spreadsheet,
        queue,
        err.message || String(err)
      );
      console.error(err && err.stack ? err.stack : err);
      return {
        ok: false,
        status: 'retry_error',
        queue,
        error: err
      };
    }
  }, CALENDAR_PENDING_RETRY_LOCK_WAIT_MS);

  if (!attempt.acquired) {
    ensurePendingCalendarRetryScheduled_(spreadsheet);
    return {
      ok: false,
      queued: true,
      status: 'retry_lock_busy'
    };
  }

  const remaining = readPendingCalendarQueue_();
  if (!pendingCalendarQueueHasWork_(remaining)) {
    PropertiesService.getScriptProperties().deleteProperty(
      CALENDAR_PENDING_SPREADSHEET_ID_PROPERTY
    );
  } else if (
    attempt.value &&
    (
      attempt.value.ok ||
      attempt.value.status === 'queue_lock_busy'
    )
  ) {
    ensurePendingCalendarRetryScheduled_(spreadsheet);
  }
  return attempt.value;
}

function rebuildCalendarRegistry_(spreadsheet, options) {
  const settings = options || {};
  const scan = buildCurrentCalendarScan_(spreadsheet);
  const structureIssueCount =
    (scan.structuralActions || []).length +
    (scan.structuralConflicts || []).length;
  if (structureIssueCount && !settings.allowConflicts) {
    throw new Error(
      `月份刀表有 ${structureIssueCount} 項日期區塊結構問題；索引未重建。`
    );
  }
  if (scan.duplicateIds.length && !settings.allowConflicts) {
    throw new Error(
      `CalendarEventId 重複：${scan.duplicateIds.join('、')}。索引未重建。`
    );
  }
  const index = writeCalendarRegistryStore_(scan, []);
  const verify = readCalendarRegistryStore_();
  if (!verify.ok || verify.entries.length !== index.eventCount) {
    throw new Error('V2 索引回讀驗證失敗。');
  }
  return {
    ok: true,
    eventCount: index.eventCount,
    duplicateIds: scan.duplicateIds,
    fingerprint: index.fingerprint
  };
}

function refreshCalendarRegistryForSheet_(spreadsheet, sheet, options) {
  const settings = options || {};
  const baseline = settings.baseline || readCalendarRegistryStore_();
  if (!baseline.ok) {
    return rebuildCalendarRegistry_(spreadsheet, {
      allowConflicts: Boolean(settings.allowStructuralActions)
    });
  }
  const baselineIdCounts = {};
  (baseline.entries || []).forEach(entry => {
    const eventId = toCellText_(entry && entry.eventId);
    if (!eventId) return;
    baselineIdCounts[eventId] = (baselineIdCounts[eventId] || 0) + 1;
  });
  const duplicateBaselineIds = Object.keys(baselineIdCounts).filter(eventId => {
    return baselineIdCounts[eventId] > 1;
  });
  if (duplicateBaselineIds.length) {
    throw new Error(
      `V2 索引有 ${duplicateBaselineIds.length} 組重複 ` +
      'CalendarEventId；已停止單表索引刷新。'
    );
  }
  const sheetScan = settings.sheetScan ||
    buildManagedSheetCalendarScan_(sheet);
  if (
    sheetScan.structuralConflicts.length ||
    (
      sheetScan.structuralActions.length &&
      !settings.allowStructuralActions
    )
  ) {
    throw new Error(
      `${sheet.getName()} 有 ` +
      `${sheetScan.structuralActions.length} 項待修復結構、` +
      `${sheetScan.structuralConflicts.length} 項結構衝突；索引未刷新。`
    );
  }
  const globalIds = settings.globalIdScan ||
    buildGlobalEventIdLocationsLightweight_(spreadsheet);
  if (globalIds.duplicateIds.length) {
    throw new Error(
      `CalendarEventId 重複：${globalIds.duplicateIds.join('、')}；索引未刷新。`
    );
  }

  const sheetId = Number(sheet.getSheetId());
  const beforeEntries = baseline.entries.filter(entry => {
    return Number(entry.sheetId) === sheetId;
  });
  const beforeIds = beforeEntries.map(entry => entry.eventId).filter(Boolean);
  const currentIds = sheetScan.contexts
    .map(context => context.eventId)
    .filter(Boolean);
  const currentIdMap = {};
  currentIds.forEach(eventId => {
    currentIdMap[eventId] = true;
  });
  const allowedRemovedIds = {};
  (settings.allowedRemovedEventIds || []).forEach(eventId => {
    if (eventId) allowedRemovedIds[eventId] = true;
  });
  const missingIds = beforeIds.filter(eventId => {
    return !currentIdMap[eventId] && !allowedRemovedIds[eventId];
  });
  const unresolvedMissingIds = settings.allowRemovedEventIds
    ? []
    : missingIds;

  const retainIds = {};
  (settings.retainEventIds || []).forEach(eventId => {
    if (eventId) retainIds[eventId] = true;
  });
  const pendingSyncByEventId = settings.pendingSyncByEventId || {};
  Object.keys(pendingSyncByEventId).forEach(eventId => {
    if (eventId) retainIds[eventId] = true;
  });
  const pendingResolutionByEventId = {
    ...(settings.pendingResolutionByEventId || {})
  };
  unresolvedMissingIds.forEach(eventId => {
    if (!eventId) return;
    retainIds[eventId] = true;
    if (!pendingResolutionByEventId[eventId]) {
      pendingResolutionByEventId[eventId] =
        'missing_event_id_from_sheet';
    }
  });
  Object.keys(pendingResolutionByEventId).forEach(eventId => {
    if (eventId) retainIds[eventId] = true;
  });
  const retainedEntries = baseline.entries.filter(entry => {
    return Number(entry.sheetId) !== sheetId || retainIds[entry.eventId];
  }).map(entry => {
    const pendingError = normalizeCalendarRegistryReasonCode_(
      pendingSyncByEventId[entry.eventId],
      pendingSyncByEventId[entry.eventId] ? 'sync_failed' : ''
    );
    const pendingResolutionReason =
      normalizeCalendarRegistryReasonCode_(
        pendingResolutionByEventId[entry.eventId],
        pendingResolutionByEventId[entry.eventId]
          ? 'pending_resolution'
          : ''
      );
    return {
      ...entry,
      pendingSync: pendingError
        ? true
        : Boolean(entry.pendingSync),
      pendingError: pendingError || entry.pendingError || '',
      pendingResolution: pendingResolutionReason
        ? true
        : Boolean(entry.pendingResolution),
      pendingResolutionReason:
        pendingResolutionReason ||
        entry.pendingResolutionReason ||
        ''
    };
  });
  const retainedIdMap = {};
  retainedEntries.forEach(entry => {
    if (entry && entry.eventId) retainedIdMap[entry.eventId] = true;
  });
  Object.keys(pendingSyncByEventId).forEach(eventId => {
    if (!eventId || retainedIdMap[eventId]) return;
    const context = sheetScan.contexts.find(item => item.eventId === eventId);
    if (!context) return;
    retainedEntries.push({
      ...toRegistryEntry_(context),
      pendingSync: true,
      pendingError: normalizeCalendarRegistryReasonCode_(
        pendingSyncByEventId[eventId],
        'sync_failed'
      )
    });
  });
  const index = writeCalendarRegistryStore_(
    sheetScan,
    retainedEntries,
    baseline.index && baseline.index.sheets
  );
  const verify = readCalendarRegistryStore_();
  if (!verify.ok || verify.entries.length !== index.eventCount) {
    throw new Error('單表 V2 索引回讀驗證失敗。');
  }
  return {
    ok: true,
    eventCount: index.eventCount,
    sheetEventCount: currentIds.length,
    pendingResolutionCount: unresolvedMissingIds.length,
    pendingResolutionIds: unresolvedMissingIds.slice(),
    duplicateIds: [],
    fingerprint: index.fingerprint
  };
}

function assertCalendarRegistrySafeToRebuild_(spreadsheet, actionLabel, options) {
  const settings = options || {};
  const baseline = readCalendarRegistryStore_();
  if (!baseline.ok) return { ok: true, missing: true };
  const health = reconcileCalendarRegistry_(spreadsheet, {
    apply: false,
    changeType: 'REBUILD_PREFLIGHT'
  });
  const repairableStructureTypes = {
    monthly_header_marker_missing: true,
    monthly_stale_header_marker: true
  };
  const ignoredRepairableActions = health.analysis.actions.filter(action => {
    return Boolean(
      settings.allowRepairableStructureActions &&
      action &&
      action.structural &&
      repairableStructureTypes[action.type]
    );
  });
  const ignoredRepairableConflicts = health.analysis.conflicts.filter(
    conflict => {
      return Boolean(
        settings.allowRepairableStructureActions &&
        conflict &&
        conflict.structural &&
        repairableStructureTypes[conflict.type]
      );
    }
  );
  const ignoredFuStoppedTrackingConflicts = health.analysis.conflicts.filter(
    conflict => isIgnorableFuStoppedTrackingConflict_(conflict)
  );
  const pending = health.analysis.actions.length -
    ignoredRepairableActions.length;
  const conflicts = health.analysis.conflicts.length -
    ignoredRepairableConflicts.length -
    ignoredFuStoppedTrackingConflicts.length;
  if (pending || conflicts || health.scan.duplicateIds.length) {
    throw new Error(
      `${actionLabel || '此操作'}前仍有 ${pending} 項待處理、` +
      `${conflicts} 項衝突；請先完成同步健康處理，索引未重建。`
    );
  }
  return {
    ok: true,
    missing: false,
    health,
    ignoredRepairableActions,
    ignoredRepairableConflicts,
    ignoredFuStoppedTrackingConflicts
  };
}

function isIgnorableFuStoppedTrackingConflict_(conflict) {
  const context = conflict && conflict.context;
  const previous = conflict && conflict.previous;
  return Boolean(
    conflict &&
    conflict.type === 'event_id_removed_or_partial_move' &&
    context &&
    context.kind === 'FU' &&
    context.invalidReason === 'no_date' &&
    !context.eventId &&
    (context.chartNo || context.patientName) &&
    previous &&
    previous.eventId &&
    registryIdentityMatchesContext_(previous, context)
  );
}

function isContextAffected_(context, options) {
  if (!options || !options.scope) return true;
  const scope = options.scope;
  if (Number(context.sheetId) !== Number(scope.sheetId)) return false;
  if (!scope.rows || !scope.rows.length) return true;
  return scope.rows.indexOf(context.row) !== -1;
}

function registryIdentityMatchesContext_(entry, context) {
  if (!entry || !context) return false;
  if (
    entry.kind &&
    context.kind &&
    entry.kind !== context.kind
  ) {
    return false;
  }
  const rowHash = toCellText_(entry.rowHash);
  if (rowHash && rowHash === toCellText_(context.rowHash)) {
    return true;
  }
  const bindingHash = toCellText_(entry.bindingHash);
  return Boolean(
    bindingHash &&
    bindingHash === toCellText_(context.bindingHash)
  );
}

function findRegistryIdentityMatches_(entries, context) {
  return (entries || []).filter(entry => {
    return registryIdentityMatchesContext_(entry, context);
  });
}

function analyzeRegistryDifferences_(baselineEntries, scan, options) {
  const settings = options || {};
  const currentById = {};
  scan.contexts.forEach(context => {
    if (!context.eventId) return;
    if (!currentById[context.eventId]) currentById[context.eventId] = [];
    currentById[context.eventId].push(context);
  });
  const baselineGroupsById = {};
  (baselineEntries || []).forEach(entry => {
    const eventId = toCellText_(entry && entry.eventId);
    if (!eventId) return;
    if (!baselineGroupsById[eventId]) baselineGroupsById[eventId] = [];
    baselineGroupsById[eventId].push(entry);
  });
  const baselineById = {};
  const duplicateBaselineIds = {};
  Object.keys(baselineGroupsById).forEach(eventId => {
    const entries = baselineGroupsById[eventId];
    if (entries.length === 1) baselineById[eventId] = entries[0];
    else duplicateBaselineIds[eventId] = true;
  });
  const currentSheetIds = {};
  scan.sheetRecords.forEach(record => {
    currentSheetIds[String(record.sheetId)] = true;
  });
  const noIdContexts = scan.contexts.filter(context => !context.eventId);
  const actions = [];
  const conflicts = [];
  const adoptions = [];
  const ignoredRemovedSheets = [];

  Object.keys(duplicateBaselineIds).forEach(eventId => {
    conflicts.push({
      type: 'duplicate_registry_event_id',
      eventId,
      context: null,
      previous: baselineGroupsById[eventId][0] || null,
      message:
        'V2 索引中同一 CalendarEventId 對應多筆資料；' +
        '已停止自動判斷。'
    });
  });

  Object.keys(currentById).forEach(eventId => {
    const locations = currentById[eventId];
    const allLocations = scan.eventLocations[eventId] || locations;
    if (allLocations.length > 1) {
      locations.forEach(context => {
        conflicts.push({
          type: 'duplicate_event_id',
          eventId,
          context,
          message: buildDuplicateEventIdMessage_(context, allLocations)
        });
      });
      return;
    }
    const context = locations[0];
    if (duplicateBaselineIds[eventId]) return;
    const previous = baselineById[eventId];
    if (!previous) {
      adoptions.push({ type: 'adopt', eventId, context });
      return;
    }
    if (context.archived) return;
    if (
      previous.pendingResolution &&
      !registryIdentityMatchesContext_(previous, context)
    ) {
      conflicts.push({
        type: 'pending_resolution_context_changed',
        eventId,
        context,
        previous,
        message:
          '此 Event ID 原已等待人工確認，且目前列與舊綁定指紋不同；' +
          '未自動更新、刪除或重綁。'
      });
      return;
    }
    const affected = isContextAffected_(context, settings);
    if (context.invalidReason) {
      if (
        affected &&
        (
          settings.changeType === 'EDIT' ||
          settings.allowMissingDeletes
        ) &&
        (
          context.invalidReason === 'no_identity' ||
          (context.kind === 'FU' && context.invalidReason === 'no_date')
        )
      ) {
        actions.push({ type: 'delete_current', eventId, context, previous });
      } else {
        conflicts.push({
          type: context.invalidReason,
          eventId,
          context,
          previous,
          message: context.invalidReason === 'invalid_time'
            ? context.timeInfo.errorMessage
            : context.invalidReason === 'no_block'
              ? '病人列缺少上方有效日期標題；未自動改期或刪除事件。'
              : context.invalidReason === 'monthly_hybrid_date_event_id'
                ? '第一欄是完整日期但該列已有 CalendarEventId；未自動改期或重建事件。'
              : '資料列缺少建立事件所需的日期或病歷號／姓名。'
        });
      }
      return;
    }
    const placementChanged =
      Number(previous.sheetId) !== Number(context.sheetId) ||
      previous.blockKey !== context.blockKey;
    const contentChanged = previous.rowHash !== context.rowHash;
    const pendingSync = Boolean(previous.pendingSync);
    const pendingResolution = Boolean(previous.pendingResolution);
    if (
      affected &&
      settings.changeType === 'REMOVE_ROW' &&
      placementChanged &&
      Number(previous.sheetId) === Number(context.sheetId)
    ) {
      conflicts.push({
        type: 'possible_date_header_deleted',
        eventId,
        context,
        previous,
        message:
          '整列刪除後日期區塊改變，可能是日期標題被刪除；未自動批次改期，請確認區塊後再同步。'
      });
    } else if (
      affected &&
      (placementChanged || contentChanged || pendingSync)
    ) {
      actions.push({
        type: 'update',
        eventId,
        context,
        previous,
        placementChanged,
        contentChanged,
        pendingSync
      });
    } else if (affected && pendingResolution) {
      actions.push({
        type: 'resolve_pending_resolution',
        eventId,
        context,
        previous
      });
    }
  });

  noIdContexts.forEach(context => {
    if (context.archived || !context.valid || !isContextAffected_(context, settings)) {
      return;
    }
    const matchingPrevious = findRegistryIdentityMatches_(
      baselineEntries,
      context
    );
    if (matchingPrevious.length) {
      conflicts.push({
        type: matchingPrevious.length > 1
          ? 'multiple_binding_match'
          : 'event_id_removed_or_partial_move',
        eventId: matchingPrevious.length === 1
          ? matchingPrevious[0].eventId
          : '',
        context,
        previous: matchingPrevious.length === 1
          ? matchingPrevious[0]
          : null,
        message: matchingPrevious.length > 1
          ? '多筆舊索引與此列的內容綁定指紋相同；' +
            '未刪除、未重綁也未建立新事件。'
          : '找到內容綁定指紋相同但 CalendarEventId 遺失的列；' +
            '未刪除舊事件，也未建立新事件。'
      });
    } else {
      actions.push({ type: 'create', eventId: '', context });
    }
  });

  (baselineEntries || []).forEach(entry => {
    if (duplicateBaselineIds[entry.eventId]) return;
    if (currentById[entry.eventId]) return;
    if (!currentSheetIds[String(entry.sheetId)]) {
      ignoredRemovedSheets.push(entry);
      return;
    }
    const matchingNoIds = noIdContexts.filter(context => {
      return registryIdentityMatchesContext_(entry, context);
    });
    if (matchingNoIds.length) {
      if (!conflicts.some(item => item.eventId === entry.eventId)) {
        conflicts.push({
          type: matchingNoIds.length > 1
            ? 'multiple_binding_match'
            : 'event_id_removed_or_partial_move',
          eventId: entry.eventId,
          context: matchingNoIds.length === 1
            ? matchingNoIds[0]
            : null,
          previous: entry,
          message: matchingNoIds.length > 1
            ? '多列同時與舊索引的內容綁定指紋相同；' +
              '未自動刪除或重綁 Calendar 事件。'
            : 'CalendarEventId 已離開原列但內容綁定指紋仍存在；' +
              '未自動刪除 Calendar 事件。'
        });
      }
      return;
    }
    if (entry.pendingResolution && !entry.pendingDelete) {
      conflicts.push({
        type: entry.pendingResolutionReason || 'pending_resolution',
        eventId: entry.eventId,
        context: null,
        previous: entry,
        message:
          '此舊索引已處於待確認狀態；' +
          '不因後續其他列的刪除事件而自動刪除 Calendar。'
      });
      return;
    }
    if (
      settings.changeType === 'REMOVE_ROW' ||
      settings.allowMissingDeletes ||
      entry.pendingDelete
    ) {
      actions.push({
        type: 'delete_missing',
        eventId: entry.eventId,
        context: null,
        previous: entry
      });
    } else {
      conflicts.push({
        type: 'missing_registry_row',
        eventId: entry.eventId,
        context: null,
        previous: entry,
        message:
          '索引中的事件列已消失；因無法確認是否為整列刪除，等待維護工具確認。'
      });
    }
  });

  scan.duplicateIds.forEach(eventId => {
    if (currentById[eventId]) return;
    const locations = scan.eventLocations[eventId] || [];
    const first = locations[0] || {};
    conflicts.push({
      type: 'duplicate_archive_event_id',
      eventId,
      context: null,
      previous: first,
      message: '年度封存表內 CalendarEventId 重複，已停止同步。'
    });
  });

  (scan.structuralActions || [])
    .concat(scan.structuralConflicts || [])
    .forEach(issue => {
      const duplicate = conflicts.some(existing => {
        return existing.context && issue.context &&
          Number(existing.context.sheetId) === Number(issue.context.sheetId) &&
          Number(existing.context.row) === Number(issue.context.row) &&
          (
            existing.type === issue.type ||
            (
              existing.type === 'no_block' &&
              issue.type === 'monthly_orphan_patient'
            )
          );
      });
      if (!duplicate) conflicts.push(issue);
    });

  return {
    actions,
    conflicts,
    adoptions,
    ignoredRemovedSheets,
    duplicates: scan.duplicateIds
  };
}

function deleteCalendarEvent_(eventId) {
  const calendarId = getConfiguredCalendarId_();
  if (!calendarId) {
    return { ok: false, status: 'no_calendar_id', message: 'Calendar ID 尚未設定。' };
  }
  if (!isCalendarAdvancedServiceAvailable_()) {
    return {
      ok: false,
      status: 'no_calendar_service',
      message: 'Calendar Advanced Service 尚未啟用。'
    };
  }
  try {
    Calendar.Events.remove(calendarId, eventId);
    return { ok: true, status: 'deleted' };
  } catch (err) {
    if (isCalendarNotFoundError_(err)) {
      return { ok: true, status: 'already_missing' };
    }
    return {
      ok: false,
      status: 'delete_failed',
      message: err.message || String(err)
    };
  }
}

function syncManagedContext_(context, duplicateLocations, runtime) {
  if (context.archived) return { ok: true, status: 'skipped_archived' };
  if (
    context.eventId &&
    duplicateLocations &&
    duplicateLocations[context.eventId] &&
    duplicateLocations[context.eventId].length > 1
  ) {
    setContextSyncNote_(
      context,
      buildDuplicateEventIdMessage_(
        context,
        duplicateLocations[context.eventId]
      ),
      true
    );
    return { ok: false, status: 'duplicate_event_id' };
  }
  if (context.invalidReason === 'invalid_time') {
    clearContextSystemNotes_(context);
    setSystemNote_(
      context.sheet.getRange(context.row, context.columns.TIME),
      TIME_ERROR_NOTE_PREFIX,
      context.timeInfo.errorMessage
    );
    return { ok: false, status: 'invalid_time' };
  }
  if (
    context.eventId &&
    (
      context.invalidReason === 'no_identity' ||
      (context.kind === 'FU' && context.invalidReason === 'no_date')
    )
  ) {
    return deleteCurrentContextEvent_(context, {
      bestEffort: Boolean(
        context.kind === 'FU' &&
        context.invalidReason === 'no_date' &&
        (context.chartNo || context.patientName)
      )
    });
  }
  if (isCalendarNoopDraftContext_(context)) {
    clearContextSystemNotes_(context);
    return { ok: true, status: 'skipped_incomplete_draft' };
  }
  if (!context.valid) {
    return { ok: false, status: context.invalidReason || 'invalid_row' };
  }
  const settings = runtime || {};
  const calendarId = Object.prototype.hasOwnProperty.call(
    settings,
    'calendarId'
  )
    ? settings.calendarId
    : getConfiguredCalendarId_();
  if (!calendarId) {
    setContextSyncNote_(context, 'Calendar ID 尚未設定。', false);
    return { ok: false, status: 'no_calendar_id' };
  }
  const calendarServiceAvailable = Object.prototype.hasOwnProperty.call(
    settings,
    'calendarServiceAvailable'
  )
    ? settings.calendarServiceAvailable
    : isCalendarAdvancedServiceAvailable_();
  if (!calendarServiceAvailable) {
    setContextSyncNote_(context, 'Calendar Advanced Service 尚未啟用。', false);
    return { ok: false, status: 'no_calendar_service' };
  }

  const resource = buildCalendarResource_(context);
  let event = null;
  let status = '';
  if (context.eventId) {
    try {
      event = Calendar.Events.update(resource, calendarId, context.eventId);
      status = 'updated';
    } catch (err) {
      if (!isCalendarNotFoundError_(err)) {
        setContextSyncNote_(
          context,
          `Calendar 尚未更新至 ${getContextExpectedScheduleText_(context)}；` +
            `更新失敗：${err.message || err}`,
          false
        );
        return { ok: false, status: 'update_failed', error: err };
      }
    }
  }
  if (!event) {
    try {
      event = Calendar.Events.insert(resource, calendarId);
      status = context.eventId ? 'recreated_404' : 'created';
    } catch (err) {
      setContextSyncNote_(
        context,
        `Calendar 尚未建立 ${getContextExpectedScheduleText_(context)} 的事件；` +
          `建立失敗：${err.message || err}`,
        false
      );
      return { ok: false, status: 'create_failed', error: err };
    }
  }

  const eventId = toCellText_(event && event.id);
  if (!eventId) {
    setContextSyncNote_(context, 'Calendar API 未回傳 event ID。', false);
    return { ok: false, status: 'missing_created_event_id' };
  }
  // An editor can move rows while Calendar is responding; the script lock
  // does not lock the sheet UI. Updating an event never needs to rewrite its
  // ID to the row number captured before the API call.
  if (eventId !== context.eventId) {
    context.sheet.getRange(context.row, context.columns.EVENT_ID).setValue(eventId);
  }
  try {
    const verified = Calendar.Events.get(calendarId, eventId);
    if (!calendarEventMatchesResource_(verified, resource)) {
      setContextSyncNote_(
        context,
        `Calendar 回讀內容與工作表預期的 ` +
          `${getContextExpectedScheduleText_(context)} 不一致。`,
        false
      );
      return { ok: false, status: 'verify_mismatch', eventId };
    }
  } catch (err) {
    setContextSyncNote_(
      context,
      `Calendar ${getContextExpectedScheduleText_(context)} 回讀驗證失敗：` +
        `${err.message || err}`,
      false
    );
    return { ok: false, status: 'verify_failed', eventId };
  }
  if (status === 'updated') {
    const currentMatches = buildManagedSheetCalendarScan_(context.sheet)
      .contexts.filter(item => item.eventId === eventId);
    if (currentMatches.length !== 1) {
      return {
        ok: false,
        status: currentMatches.length > 1
          ? 'duplicate_event_id'
          : 'row_changed_during_sync',
        eventId
      };
    }
    const current = currentMatches[0];
    if (
      !current.valid || current.rowHash !== context.rowHash ||
      current.blockKey !== context.blockKey
    ) {
      return { ok: false, status: 'row_changed_during_sync', eventId };
    }
    clearContextSystemNotes_(current);
  } else {
    clearContextSystemNotes_(context);
  }
  return { ok: true, status, eventId };
}

function deleteCurrentContextEvent_(context, options) {
  const settings = options || {};
  if (!context.eventId) return { ok: true, status: 'no_event' };
  const result = deleteCalendarEvent_(context.eventId);
  if (!result.ok) {
    if (settings.bestEffort) {
      context.sheet
        .getRange(context.row, context.columns.EVENT_ID)
        .clearContent();
      clearContextSystemNotes_(context);
      return {
        ok: true,
        status: 'tracking_stopped',
        releaseEventId: true,
        calendarDeleteOk: false,
        calendarDeleteStatus: result.status,
        message: result.message || ''
      };
    }
    if (context.recoveredEventId) {
      context.sheet
        .getRange(context.row, context.columns.EVENT_ID)
        .setValue(context.eventId);
    }
    setContextSyncNote_(context, `刪除 Calendar 事件失敗：${result.message}`, false);
    return result;
  }
  context.sheet.getRange(context.row, context.columns.EVENT_ID).clearContent();
  clearContextSystemNotes_(context);
  return {
    ...result,
    releaseEventId: true
  };
}

function annotateAnalysisConflicts_(analysis) {
  (analysis.conflicts || []).forEach(conflict => {
    if (conflict.context) {
      if (conflict.structural) {
        const markerColumn =
          conflict.context.columns.CHART_NO ||
          conflict.context.columns.TIME;
        setSystemNote_(
          conflict.context.sheet.getRange(
            conflict.context.row,
            markerColumn
          ),
          MONTHLY_STRUCTURE_NOTE_PREFIX,
          conflict.message
        );
        return;
      }
      if (
        conflict.type === 'invalid_time' &&
        conflict.context.columns.TIME
      ) {
        clearContextSystemNotes_(conflict.context);
        setSystemNote_(
          conflict.context.sheet.getRange(
            conflict.context.row,
            conflict.context.columns.TIME
          ),
          TIME_ERROR_NOTE_PREFIX,
          conflict.message
        );
        return;
      }
      setContextSyncNote_(
        conflict.context,
        conflict.message,
        true
      );
    }
  });
}

function reconcileCalendarRegistry_(spreadsheet, options) {
  const settings = options || {};
  const baseline = readCalendarRegistryStore_();
  if (!baseline.ok) {
    if (settings.apply && settings.scope) {
      const initialScan = buildCurrentCalendarScan_(spreadsheet);
      initialScan.contexts
        .filter(context => isContextAffected_(context, settings) && context.valid)
        .forEach(context => {
          syncManagedContext_(context, initialScan.eventLocations);
        });
    }
    const rebuilt = rebuildCalendarRegistry_(spreadsheet, {
      allowConflicts: false
    });
    return {
      ok: true,
      initialized: true,
      actions: [],
      conflicts: [],
      results: [],
      eventCount: rebuilt.eventCount
    };
  }

  const scan = buildCurrentCalendarScan_(spreadsheet);
  const analysis = analyzeRegistryDifferences_(baseline.entries, scan, settings);
  annotateAnalysisConflicts_(analysis);
  if (!settings.apply) {
    return {
      ok: analysis.conflicts.length === 0 && scan.duplicateIds.length === 0,
      scan,
      analysis,
      results: []
    };
  }

  const results = [];
  const resolvedDeleteIds = {};
  const failedDeleteEntries = [];
  analysis.actions.forEach(action => {
    let result;
    if (action.type === 'create' || action.type === 'update') {
      result = syncManagedContext_(action.context, scan.eventLocations);
    } else if (action.type === 'delete_current') {
      result = deleteCurrentContextEvent_(action.context);
      if (result.ok) resolvedDeleteIds[action.eventId] = true;
    } else if (action.type === 'delete_missing') {
      result = deleteCalendarEvent_(action.eventId);
      if (result.ok) {
        resolvedDeleteIds[action.eventId] = true;
      } else {
        failedDeleteEntries.push({
          ...action.previous,
          pendingDelete: true,
          pendingError: result.status || 'delete_failed'
        });
      }
    } else if (action.type === 'resolve_pending_resolution') {
      clearContextSystemNotes_(action.context);
      result = { ok: true, status: 'pending_resolution_cleared' };
    } else {
      result = { ok: true, status: 'ignored' };
    }
    results.push({ action, result });
  });

  const unresolvedEntries = [];
  results.forEach(item => {
    if (!item || !item.result || item.result.ok) return;
    const action = item.action || {};
    if (
      action.type !== 'create' &&
      action.type !== 'update' &&
      action.type !== 'calendar_drift' &&
      action.type !== 'calendar_missing'
    ) {
      return;
    }
    const context = action.context;
    if (!context) return;
    let pendingEntry = action.previous
      ? { ...action.previous }
      : toRegistryEntry_(context);
    if (item.result.eventId) {
      pendingEntry = {
        ...toRegistryEntry_(context),
        eventId: item.result.eventId
      };
    }
    if (!pendingEntry.eventId) return;
    pendingEntry.pendingSync = true;
    pendingEntry.pendingError = normalizeCalendarRegistryReasonCode_(
      item.result.status,
      'sync_failed'
    );
    unresolvedEntries.push(pendingEntry);
  });
  analysis.conflicts.forEach(conflict => {
    if (
      conflict.previous &&
      !resolvedDeleteIds[conflict.previous.eventId]
    ) {
      unresolvedEntries.push({
        ...conflict.previous,
        pendingResolution: true,
        pendingResolutionReason: normalizeCalendarRegistryReasonCode_(
          conflict.type,
          'registry_conflict'
        ),
        pendingError: normalizeCalendarRegistryReasonCode_(
          conflict.previous.pendingError || conflict.type,
          'registry_conflict'
        )
      });
    }
  });
  failedDeleteEntries.forEach(entry => unresolvedEntries.push(entry));
  if (settings.scope) {
    (baseline.entries || []).forEach(entry => {
      const sameSheet =
        Number(entry.sheetId) === Number(settings.scope.sheetId);
      const inRows = !settings.scope.rows ||
        !settings.scope.rows.length ||
        settings.scope.rows.indexOf(entry.row) !== -1;
      if (!sameSheet || !inRows) unresolvedEntries.push(entry);
    });
  }
  const retainedById = {};
  unresolvedEntries.forEach(entry => {
    if (entry && entry.eventId && !resolvedDeleteIds[entry.eventId]) {
      retainedById[entry.eventId] = entry;
    }
  });
  const refreshed = buildCurrentCalendarScan_(spreadsheet);
  writeCalendarRegistryStore_(refreshed, Object.values(retainedById));
  // Pure reorders require no Calendar call, but may carry obsolete failure
  // notes. Only clear notes for unique, unchanged, previously settled IDs.
  const settledById = {};
  baseline.entries.forEach(entry => {
    if (!entry.pendingSync && !entry.pendingDelete && !entry.pendingResolution) {
      settledById[entry.eventId] = entry;
    }
  });
  const blockedIds = new Set((analysis.conflicts || []).map(item => item.eventId));
  clearContextsSystemNotesBatch_(refreshed.contexts.filter(context => {
    const previous = settledById[context.eventId];
    return previous && context.valid && !context.archived &&
      isContextAffected_(context, settings) &&
      !blockedIds.has(context.eventId) && !retainedById[context.eventId] &&
      (refreshed.eventLocations[context.eventId] || []).length === 1 &&
      Number(previous.sheetId) === Number(context.sheetId) &&
      previous.rowHash === context.rowHash &&
      previous.blockKey === context.blockKey;
  }));
  return {
    ok:
      results.every(item => item.result.ok) &&
      analysis.conflicts.length === 0 &&
      refreshed.duplicateIds.length === 0,
    scan: refreshed,
    analysis,
    results
  };
}

function getHealthActionLocation_(action) {
  const source = action.context || action.previous || {};
  return {
    sheetName: source.sheetName || '',
    row: source.row || ''
  };
}

function sanitizeHealthMessage_(message, eventId) {
  const raw = toCellText_(message);
  return eventId
    ? raw.split(toCellText_(eventId)).join('[Event ID 已隱藏]')
    : raw;
}

function describeHealthFailure_(item) {
  const result = item.result || {};
  const messages = {
    duplicate_event_id: 'CalendarEventId 重複，已停止同步。',
    duplicate_archive_event_id: '年度封存表內 CalendarEventId 重複，已停止同步。',
    invalid_time: '時間格式無效，請修正時間後重試。',
    no_calendar_id: 'Calendar ID 尚未設定。',
    no_calendar_service: 'Calendar Advanced Service 尚未啟用。',
    update_failed: 'Calendar API 更新事件失敗，請稍後重試。',
    create_failed: 'Calendar API 建立事件失敗，請稍後重試。',
    missing_created_event_id: 'Calendar API 建立事件後未回傳 Event ID。',
    verify_mismatch: 'Calendar 回讀內容與工作表不一致。',
    verify_failed: 'Calendar 回讀驗證失敗，請稍後重試。',
    row_changed_during_sync: '同步期間資料列已移動或變更，等待依最新位置重試。',
    delete_failed: 'Calendar API 刪除事件失敗，已保留待重試狀態。',
    calendar_list_failed: 'Calendar 事件清單讀取失敗，請稍後重試。',
    calendar_read_failed: 'Calendar 事件回讀失敗，請稍後重試。'
  };
  return messages[result.status] ||
    sanitizeHealthMessage_(result.message || result.status || '同步失敗。');
}

function describePendingHealthAction_(action) {
  if (action && action.message) {
    return sanitizeHealthMessage_(action.message, action.eventId);
  }
  const messages = {
    create: '有效資料列尚無事件，等待建立。',
    update: '工作表內容或所在日期區塊已變更，等待更新事件。',
    delete_current: '資料列已清空必要資料，等待刪除自己的事件。',
    delete_missing: '資料列已消失，等待確認刪除事件。',
    resolve_pending_resolution:
      '原待確認的 Event ID 已回到唯一吻合資料列，等待清除異常狀態。',
    calendar_missing: '工作表事件在 Calendar 中遺失，等待重新建立。',
    calendar_drift: 'Calendar 的內容、日期、時間或顏色與工作表不同，等待更新。',
    calendar_orphan: '系統格式事件已無 FU、月表或年度封存來源，等待確認刪除。'
  };
  return messages[action.type] || '資料與同步索引不同。';
}

function getHealthIssueRows_(result) {
  const rows = [];
  const analysis = result.analysis || { actions: [], conflicts: [] };
  const appliedResults = result.results || [];
  const appliedKeys = {};
  appliedResults.forEach(item => {
    appliedKeys[getHealthActionKey_(item.action)] = true;
  });
  (analysis.actions || [])
    .filter(action => !appliedKeys[getHealthActionKey_(action)])
    .forEach(action => {
      const location = getHealthActionLocation_(action);
      rows.push([
        new Date(),
        '待處理',
        action.type,
        location.sheetName,
        location.row,
        sha256Hex_(action.eventId || ''),
        describePendingHealthAction_(action)
      ]);
    });
  appliedResults
    .filter(item => !item.result || !item.result.ok)
    .forEach(item => {
      const action = item.action || {};
      const location = getHealthActionLocation_(action);
      rows.push([
        new Date(),
        '失敗',
        (item.result && item.result.status) || action.type || 'sync_failed',
        location.sheetName,
        location.row,
        sha256Hex_(action.eventId || ''),
        describeHealthFailure_(item)
      ]);
    });
  (analysis.conflicts || []).forEach(conflict => {
    rows.push([
      new Date(),
      '衝突',
      conflict.type,
      conflict.context
        ? conflict.context.sheetName
        : conflict.previous
          ? conflict.previous.sheetName
          : '',
      conflict.context
        ? conflict.context.row
        : conflict.previous
          ? conflict.previous.row
          : '',
      sha256Hex_(conflict.eventId || ''),
      sanitizeHealthMessage_(conflict.message, conflict.eventId)
    ]);
  });
  const pendingQueue = summarizePendingCalendarQueue_(
    result && result.pendingQueue
  );
  if (pendingQueue.hasWork) {
    rows.push([
      new Date(),
      '待重試',
      'pending_queue',
      '',
      '',
      '',
      pendingQueue.fullScan
        ? '有一筆全系統同步要求已排入重試；可執行「同步待處理變更」。'
        : `有 ${pendingQueue.sheetCount} 張表、${pendingQueue.rowCount} 列` +
          '已排入重試；可執行「同步待處理變更」。'
    ]);
  }
  return rows;
}

function buildHealthAlertSummary_(result, maxItems) {
  const rows = getHealthIssueRows_(result);
  if (!rows.length) {
    return {
      hasIssues: false,
      issueCount: 0,
      displayedCount: 0,
      text: '未發現待處理變更或衝突。'
    };
  }
  const requestedLimit = Number(maxItems);
  const limit = Number.isFinite(requestedLimit) && requestedLimit > 0
    ? Math.floor(requestedLimit)
    : 8;
  const lines = rows.slice(0, limit).map(row => {
    const sheetName = toCellText_(row[3]);
    const rowNumber = Number(row[4]);
    const location = [
      sheetName,
      rowNumber > 0 ? `第 ${rowNumber} 列` : ''
    ].filter(Boolean).join(' ');
    const label = `${toCellText_(row[1])}／${toCellText_(row[2])}`;
    return `${label}${location ? `（${location}）` : ''}：` +
      toSingleLineText_(row[6]);
  });
  if (rows.length > limit) {
    lines.push(`另有 ${rows.length - limit} 項未列出；` +
      '請先處理以上項目後重新檢查。');
  }
  return {
    hasIssues: true,
    issueCount: rows.length,
    displayedCount: Math.min(rows.length, limit),
    text: lines.join('\n')
  };
}

function checkSyncHealth() {
  return runMenuAction_('檢查同步健康', () => {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    assertNoLegacyArchiveState_();
    assertNoActiveAnnualArchiveTransaction_();
    const result = buildCalendarHealthPreview_(spreadsheet, {
      apply: false,
      changeType: 'HEALTH'
    });
    const summary = buildHealthAlertSummary_(result, 10);
    const actionCount = result.analysis.actions.length;
    const conflictCount = result.analysis.conflicts.length;
    const queued = summarizePendingCalendarQueue_(result.pendingQueue);
    SpreadsheetApp.getUi().alert(
      '同步健康檢查',
      `待處理 ${actionCount} 項；衝突 ${conflictCount} 項；` +
        `排隊 ${queued.hasWork ? 1 : 0} 批。` +
        (summary.hasIssues
          ? `\n\n異常摘要：\n${summary.text}`
          : '\n未發現待處理變更或衝突。') +
        '\n本檢查只使用唯讀 Calendar API，不會修改或刪除事件。',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return result;
  });
}

function syncPendingChanges() {
  return runMenuAction_('同步待處理變更', () => {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    assertNoLegacyArchiveState_();
    assertNoActiveAnnualArchiveTransaction_();
    const clearedDraftNotes = clearFuNoopDraftSystemNotes_(spreadsheet);
    const preview = buildCalendarHealthPreview_(spreadsheet, {
      apply: false,
      changeType: 'MANUAL',
      allowMissingDeletes: true
    });
    const deletes = preview.analysis.actions.filter(action => {
      return action.type === 'delete_missing';
    }).length;
    const orphans = preview.analysis.actions.filter(action => {
      return action.type === 'calendar_orphan';
    });
    if (deletes || orphans.length) {
      const response = SpreadsheetApp.getUi().alert(
        '確認刪除遺留事件',
        `偵測到 ${deletes} 筆已刪除資料列的事件，以及 ` +
          `${orphans.length} 筆無 FU、月表或年度封存來源的系統事件。` +
          '\n確認後會刪除上述事件；人工建立、格式不符的事件不在範圍內，' +
          '整張分頁消失的事件也不會在此自動刪除。',
        SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
      );
      if (response !== SpreadsheetApp.getUi().Button.OK) {
        return { ok: false, cancelled: true };
      }
    }
    const confirmedOrphanIds = orphans.map(action => action.eventId);
    const result = withCalendarSyncLock_(() => {
      const queuedBeforeRun = takePendingCalendarQueue_();
      try {
        const applied = runFullCalendarReconcile_(spreadsheet, {
          apply: true,
          changeType: 'MANUAL',
          allowMissingDeletes: true,
          confirmedOrphanIds
        });
        if (!applied.ok) restorePendingCalendarQueue_(queuedBeforeRun);
        return applied;
      } catch (err) {
        restorePendingCalendarQueue_(queuedBeforeRun);
        throw err;
      }
    });
    result.pendingQueue = readPendingCalendarQueue_();
    const clearedResolvedNotes = result.ok
      ? clearResolvedCalendarSystemNotes_(result.scan)
      : 0;
    if (!pendingCalendarQueueHasWork_(result.pendingQueue)) {
      clearPendingCalendarRetryTriggers_();
      PropertiesService.getScriptProperties().deleteProperty(
        CALENDAR_PENDING_SPREADSHEET_ID_PROPERTY
      );
    }
    const summary = buildHealthAlertSummary_(result, 10);
    const failed = result.results.filter(item => !item.result.ok).length;
    SpreadsheetApp.getUi().alert(
      `已處理 ${result.results.length} 項；失敗 ${failed} 項；` +
      `仍有衝突 ${result.analysis.conflicts.length} 項。` +
      (summary.hasIssues
        ? `\n\n未完成摘要：\n${summary.text}`
        : '\n所有可安全處理項目均已完成。') +
      (
        clearedDraftNotes + clearedResolvedNotes
          ? `\n已清除 ${clearedDraftNotes + clearedResolvedNotes} 個已解決的系統 note。`
          : ''
      )
    );
    return result;
  });
}

function getAffectedRowsForEdit_(
  event,
  sheet,
  kind,
  columns,
  precomputedBlockScan
) {
  const rows = [];
  const start = Math.max(2, event.range.getRow());
  const end = event.range.getRow() + event.range.getNumRows() - 1;
  if (kind !== 'MONTHLY') {
    for (let row = start; row <= end; row++) rows.push(row);
    return rows;
  }

  const scan = precomputedBlockScan ||
    scanMonthlyBlocks_(sheet, columns);
  const touchesBlockIdentity = isEditTouchingColumns_(
    event.range,
    columns,
    ['TIME', 'HOSPITAL']
  );
  const patientsByRow = {};
  scan.blocks.forEach(block => {
    block.patientRows.forEach(patient => {
      patientsByRow[patient.row] = patient;
    });
  });
  scan.orphans.forEach(patient => {
    patientsByRow[patient.row] = patient;
  });
  scan.hybrids.forEach(item => {
    patientsByRow[item.row] = item;
  });

  for (let row = start; row <= end; row++) {
    const block = touchesBlockIdentity
      ? scan.blocks.find(item => Number(item.row) === Number(row))
      : null;
    if (block) {
      block.patientRows.forEach(patient => rows.push(patient.row));
      continue;
    }
    if (patientsByRow[row]) {
      rows.push(row);
      continue;
    }
    const values = scan.rowValuesByRow[row] || [];
    const type = classifyMonthlyRowValues_(values, columns).type;
    if (type === MONTHLY_ROW_TYPES.HYBRID_CONFLICT) {
      rows.push(row);
    }
  }
  return Array.from(new Set(rows));
}

function getManagedContextsForRows_(sheet, rows) {
  const wanted = {};
  (rows || []).forEach(row => {
    wanted[Number(row)] = true;
  });
  return buildManagedSheetCalendarScan_(sheet).contexts
    .filter(context => wanted[Number(context.row)]);
}

function isFuLifecycleRecoveryContext_(context) {
  return Boolean(
    context &&
    context.kind === 'FU' &&
    !context.eventId &&
    context.invalidReason === 'no_date' &&
    (context.chartNo || context.patientName) &&
    context.bindingHash
  );
}

function buildFuLifecycleRecoveryResolutions_(
  contexts,
  baseline,
  globalIdScan
) {
  const entries = baseline && baseline.ok
    ? baseline.entries || []
    : [];
  const eventLocations = globalIdScan && globalIdScan.eventLocations
    ? globalIdScan.eventLocations
    : {};
  const candidatesByRow = {};
  const ownersByEventId = {};
  (contexts || []).forEach(context => {
    if (!isFuLifecycleRecoveryContext_(context)) return;
    const candidates = entries.filter(entry => {
      return Boolean(
        entry &&
        entry.eventId &&
        entry.kind === 'FU' &&
        Number(entry.sheetId) === Number(context.sheetId) &&
        toCellText_(entry.bindingHash) &&
        toCellText_(entry.bindingHash) ===
          toCellText_(context.bindingHash)
      );
    });
    candidatesByRow[context.row] = candidates;
    candidates.forEach(entry => {
      if (!ownersByEventId[entry.eventId]) {
        ownersByEventId[entry.eventId] = [];
      }
      ownersByEventId[entry.eventId].push(context.row);
    });
  });

  const byRow = {};
  Object.keys(candidatesByRow).forEach(rowKey => {
    const row = Number(rowKey);
    const candidates = candidatesByRow[row] || [];
    if (!candidates.length) {
      byRow[row] = { status: 'none', candidates: [], reason: '' };
      return;
    }
    const candidateIds = Array.from(new Set(
      candidates.map(entry => entry.eventId).filter(Boolean)
    ));
    if (candidates.length !== 1 || candidateIds.length !== 1) {
      byRow[row] = {
        status: 'pending',
        candidates,
        reason: 'multiple_registry_binding_matches'
      };
      return;
    }
    const eventId = candidateIds[0];
    if ((ownersByEventId[eventId] || []).length !== 1) {
      byRow[row] = {
        status: 'pending',
        candidates,
        reason: 'multiple_sheet_rows_match_binding'
      };
      return;
    }
    if ((eventLocations[eventId] || []).length) {
      byRow[row] = {
        status: 'pending',
        candidates,
        reason: 'event_id_still_present_elsewhere'
      };
      return;
    }
    byRow[row] = {
      status: 'recover',
      candidates,
      eventId,
      reason: 'unique_binding_match'
    };
  });
  return byRow;
}

function getFuLifecyclePendingResolutionMessage_(reason) {
  const messages = {
    multiple_registry_binding_matches:
      '多筆舊索引與此 FU 列相同，無法安全判斷要停止哪一個事件。',
    multiple_sheet_rows_match_binding:
      '多個 FU 列與同一舊事件吻合，未自動刪除或重建事件。',
    event_id_still_present_elsewhere:
      '吻合的 CalendarEventId 仍存在其他資料列，未自動刪除事件。'
  };
  return messages[reason] ||
    'FU 追蹤狀態無法唯一確認，已保留索引等待人工處理。';
}

function filterFuTriggerRowsForCalendar_(sheet, rows, columns, options) {
  const settings = options || {};
  const targetRows = Array.from(new Set((rows || [])
    .map(Number)
    .filter(row => Number.isInteger(row) && row >= 2)))
    .sort((left, right) => left - right);
  if (!targetRows.length) {
    return {
      rows: [],
      skippedRows: [],
      skippedContexts: [],
      recoveryResolutions: {},
      requiresRegistryRefresh: false,
      pendingResolutionByEventId: {}
    };
  }
  const minRow = targetRows[0];
  const maxRow = targetRows[targetRows.length - 1];
  const lastColumn = Math.max(
    getLastHeaderColumn_(sheet),
    ...Object.values(columns)
  );
  const values = sheet.getRange(
    minRow,
    1,
    maxRow - minRow + 1,
    lastColumn
  ).getValues();
  const targetSet = {};
  targetRows.forEach(row => {
    targetSet[row] = true;
  });
  const skippedContexts = [];
  const activeRows = [];
  const contexts = [];
  values.forEach((rowValues, index) => {
    const row = minRow + index;
    if (!targetSet[row]) return;
    const context = buildFuRowContext_(
      sheet,
      row,
      columns,
      rowValues
    );
    contexts.push(context);
  });
  let recoveryResolutions = {};
  let requiresRegistryRefresh = false;
  const pendingResolutionByEventId = {};
  if (settings.recoverFromRegistry) {
    const baseline = settings.baseline || readCalendarRegistryStore_();
    const hasCandidates = baseline.ok && contexts.some(context => {
      if (!isFuLifecycleRecoveryContext_(context)) return false;
      return (baseline.entries || []).some(entry => {
        return entry &&
          entry.eventId &&
          entry.kind === 'FU' &&
          Number(entry.sheetId) === Number(context.sheetId) &&
          toCellText_(entry.bindingHash) ===
            toCellText_(context.bindingHash);
      });
    });
    const needsMissingIdCheck = baseline.ok && contexts.some(context => {
      return isCalendarNoopDraftContext_(context);
    });
    const globalIdScan = hasCandidates || needsMissingIdCheck
      ? settings.globalIdScan || (
        settings.spreadsheet
          ? buildGlobalEventIdLocationsLightweight_(settings.spreadsheet)
          : { eventLocations: {} }
      )
      : { eventLocations: {} };
    recoveryResolutions = buildFuLifecycleRecoveryResolutions_(
      contexts,
      baseline,
      globalIdScan
    );
    if (baseline.ok && needsMissingIdCheck) {
      const sheetId = Number(sheet.getSheetId());
      (baseline.entries || []).forEach(entry => {
        if (
          Number(entry.sheetId) !== sheetId ||
          !entry.eventId ||
          (globalIdScan.eventLocations[entry.eventId] || []).length
        ) {
          return;
        }
        requiresRegistryRefresh = true;
        pendingResolutionByEventId[entry.eventId] =
          'edited_row_content_cleared_or_id_missing';
      });
    }
  }
  contexts.forEach(context => {
    const recovery = recoveryResolutions[context.row];
    if (
      !isCalendarNoopDraftContext_(context) ||
      (recovery && recovery.status !== 'none')
    ) {
      activeRows.push(context.row);
    } else {
      skippedContexts.push(context);
    }
  });
  clearContextsSystemNotesBatch_(skippedContexts);
  return {
    rows: activeRows,
    skippedRows: skippedContexts.map(context => context.row),
    skippedContexts,
    recoveryResolutions,
    requiresRegistryRefresh,
    pendingResolutionByEventId
  };
}

function clearFuNoopDraftSystemNotes_(spreadsheet) {
  const sheet = getMainTrackingSheet_(spreadsheet);
  if (!sheet) return 0;
  const scan = buildManagedSheetCalendarScan_(sheet);
  return clearContextsSystemNotesBatch_(
    scan.contexts.filter(isCalendarNoopDraftContext_)
  );
}

function clearResolvedCalendarSystemNotes_(scan) {
  if (!scan || !Array.isArray(scan.contexts)) return 0;
  return clearContextsSystemNotesBatch_(
    scan.contexts.filter(context => {
      return context.valid || isCalendarNoopDraftContext_(context);
    })
  );
}

function annotateQueuedSyncRows_(sheet, rows, message) {
  if (!message) return 0;
  let count = 0;
  getManagedContextsForRows_(sheet, rows).forEach(context => {
    setContextSyncNote_(context, message, false);
    count++;
  });
  return count;
}

function runTriggerCalendarWork_(spreadsheet, request, callback, noteTarget) {
  try {
    const attempt = tryWithCalendarSyncLock_(() => {
      const primary = callback();
      const pending = drainPendingCalendarQueue_(spreadsheet);
      if (primary && primary.ok === false) {
        enqueuePendingCalendarSync_(request);
      }
      return { queued: false, primary, pending };
    });
    if (attempt.acquired) {
      if (
        (attempt.value.primary && attempt.value.primary.ok === false) ||
        (attempt.value.pending && attempt.value.pending.ok === false)
      ) {
        ensurePendingCalendarRetryScheduled_(spreadsheet);
      }
      return attempt.value;
    }
  } catch (err) {
    enqueuePendingCalendarSync_(request);
    ensurePendingCalendarRetryScheduled_(spreadsheet);
    if (noteTarget && noteTarget.sheet && noteTarget.rows) {
      annotateQueuedSyncRows_(
        noteTarget.sheet,
        noteTarget.rows,
        `自動同步發生未預期錯誤：${err.message || err}；` +
          '工作表資料已保留並排入重試。請執行「修復選取列同步」。'
      );
    }
    console.error(err && err.stack ? err.stack : err);
    return { queued: true, error: err };
  }

  enqueuePendingCalendarSync_(request);
  ensurePendingCalendarRetryScheduled_(spreadsheet);
  return { queued: true, reason: 'calendar_sync_busy' };
}

function processRowChange(e) {
  if (!e || !e.range) return;
  if (isAnnualArchiveTransactionActive_()) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  if (isAnnualArchiveSheet_(sheet)) return;
  const isFu = isMainTrackingSheetName_(sheetName);
  const isMonthly = isMonthlySheetName_(sheetName);
  if (!isFu && !isMonthly) return;
  const kind = isFu ? 'FU' : 'MONTHLY';
  const columns = isFu
    ? getRequiredFuColumns_(sheet)
    : getRequiredMonthlyColumns_(sheet);
  const calendarKeys = isFu ? FU_CALENDAR_KEYS : MONTHLY_CALENDAR_KEYS;
  const wrapKeys = isFu ? FU_WRAP_KEYS : MONTHLY_WRAP_KEYS;
  const standardKeys = isFu
    ? CONFIG.FIELD_KEYS
    : CONFIG.MONTHLY_FIELD_KEYS;
  const touchedKeys = getTouchedManagedKeys_(
    e.range,
    columns,
    standardKeys
  );
  const calendarTouched = isEditTouchingColumns_(
    e.range,
    columns,
    calendarKeys
  );
  const wrapTouched = touchedKeys.some(key => {
    return wrapKeys.indexOf(key) !== -1;
  });
  const headerPresentationTouched = kind === 'MONTHLY' &&
    isEditTouchingColumns_(
      e.range,
      columns,
      ['TIME', 'HOSPITAL', 'CHART_NO']
    );
  const start = Math.max(2, e.range.getRow());
  const end = e.range.getRow() + e.range.getNumRows() - 1;
  if (touchedKeys.length) {
    const rows = [];
    for (let row = start; row <= end; row++) rows.push(row);
    applyTouchedManagedRowsFormatBatch_(
      sheet,
      rows,
      kind,
      columns,
      touchedKeys
    );
  }

  // Time normalization must not wait behind a long Calendar operation.
  normalizeEditedTimes_(e, sheet, kind, columns);
  if (headerPresentationTouched) {
    const lastColumn = Math.max(
      getLastHeaderColumn_(sheet),
      ...Object.values(columns)
    );
    const rowCount = Math.max(0, end - start + 1);
    const values = rowCount
      ? sheet.getRange(start, 1, rowCount, lastColumn).getValues()
      : [];
    values.forEach((rowValues, index) => {
      if (!needsMonthlyHeaderPresentation_(rowValues, columns)) return;
      ensureMonthlyHeaderPresentation_(
        sheet,
        start + index,
        columns,
        rowValues
      );
    });
  }
  if (!calendarTouched) {
    if (wrapTouched && end >= start) {
      sheet.autoResizeRows(start, end - start + 1);
    }
    return;
  }

  const spreadsheet = e.source || SpreadsheetApp.getActiveSpreadsheet();
  const preparedSheetScan = kind === 'MONTHLY'
    ? buildManagedSheetCalendarScan_(sheet)
    : null;
  let rows = getAffectedRowsForEdit_(
    e,
    sheet,
    kind,
    columns,
    preparedSheetScan && preparedSheetScan.blockScan
  );
  let fuFilter = null;
  if (kind === 'FU' && rows.length) {
    fuFilter = filterFuTriggerRowsForCalendar_(
      sheet,
      rows,
      columns,
      {
        recoverFromRegistry: true,
        spreadsheet
      }
    );
    rows = fuFilter.rows;
  }
  if (!rows.length) {
    if (fuFilter && fuFilter.requiresRegistryRefresh) {
      const request = {
        fullScan: true,
        allowMissingDeletes: false,
        reason: 'EDIT_UNKNOWN_ROW_CLEAR'
      };
      const result = runTriggerCalendarWork_(
        spreadsheet,
        request,
        () => ({
          ok: refreshCalendarRegistryForSheet_(
            spreadsheet,
            sheet,
            {
              pendingResolutionByEventId:
                fuFilter.pendingResolutionByEventId
            }
          ).ok,
          status: 'registry_pending_resolution_recorded'
        }),
        null
      );
      if (wrapTouched && end >= start) {
        sheet.autoResizeRows(start, end - start + 1);
      }
      return result;
    }
    if (wrapTouched && end >= start) {
      sheet.autoResizeRows(start, end - start + 1);
    }
    return;
  }
  const request = {
    sheetId: sheet.getSheetId(),
    rows,
    reason: 'EDIT'
  };
  const result = runTriggerCalendarWork_(
    spreadsheet,
    request,
    () => syncManagedRowsAt_(
      spreadsheet,
      sheet,
      rows,
      { sheetScan: preparedSheetScan }
    ),
    { sheet, rows }
  );
  if (wrapTouched && end >= start) {
    sheet.autoResizeRows(start, end - start + 1);
  }
  return result;
}

function processCalendarStructureChange(e) {
  if (!e || !e.source) return;
  if (isAnnualArchiveTransactionActive_()) return;
  const changeType = toCellText_(e.changeType) || 'OTHER';
  if (changeType === 'EDIT') {
    return { ok: true, status: 'handled_by_on_edit' };
  }
  const request = {
    fullScan: true,
    allowMissingDeletes: changeType === 'REMOVE_ROW',
    reason: `CHANGE_${changeType}`
  };
  enqueuePendingCalendarSync_(request);
  return runTriggerCalendarWork_(
    e.source,
    request,
    () => ({
      ok: true,
      status: 'queued_structure_change'
    }),
    null
  );
}

function buildCalendarSyncRuntime_() {
  return {
    calendarId: getConfiguredCalendarId_(),
    calendarServiceAvailable: isCalendarAdvancedServiceAvailable_()
  };
}

function syncPreparedManagedContext_(context, eventLocations, runtime) {
  try {
    return syncManagedContext_(context, eventLocations, runtime);
  } catch (err) {
    setContextSyncNote_(
      context,
      `Calendar 同步發生未預期錯誤：${err.message || err}`,
      false
    );
    return {
      ok: false,
      status: 'unexpected_sync_error',
      error: err
    };
  }
}

function getRegistryRetentionForSync_(context, result) {
  const replacedEventId = Boolean(
    context.eventId &&
    result.eventId &&
    result.eventId !== context.eventId
  );
  const deletedEventId = Boolean(
    context.eventId &&
    result.ok &&
    (
      result.status === 'deleted' ||
      result.status === 'already_missing' ||
      result.releaseEventId
    )
  );
  return {
    retainEventIds:
      !replacedEventId && !result.ok && context.eventId
        ? [context.eventId]
        : [],
    allowedRemovedEventIds:
      replacedEventId || deletedEventId ? [context.eventId] : []
  };
}

function syncManagedRowsAt_(spreadsheet, sheet, rows, options) {
  const settings = options || {};
  const targetRows = Array.from(new Set((rows || [])
    .map(Number)
    .filter(row => row >= 2)));
  if (!targetRows.length) {
    return { ok: true, results: [], registry: null };
  }
  const sheetScan = settings.sheetScan ||
    buildManagedSheetCalendarScan_(sheet);
  const globalIds = resolveGlobalEventIdLocationsForSync_(
    spreadsheet,
    sheet,
    sheetScan
  );
  const contextsByRow = {};
  sheetScan.contexts.forEach(context => {
    contextsByRow[context.row] = context;
  });
  const recoveryResolutions = isMainTrackingSheetName_(sheet.getName())
    ? buildFuLifecycleRecoveryResolutions_(
      targetRows.map(row => contextsByRow[row]).filter(Boolean),
      globalIds.baseline,
      globalIds
    )
    : {};
  const runtime = buildCalendarSyncRuntime_();
  const retainEventIds = [];
  const allowedRemovedEventIds = [];
  const pendingSyncByEventId = {};
  const pendingResolutionByEventId = {};
  const results = targetRows.map(row => {
    const baseContext = contextsByRow[row];
    if (!baseContext) {
      return {
        row,
        context: null,
        result: { ok: false, status: 'row_not_managed' }
      };
    }
    const recovery = recoveryResolutions[row];
    if (recovery && recovery.status === 'pending') {
      (recovery.candidates || []).forEach(entry => {
        if (entry && entry.eventId) {
          pendingResolutionByEventId[entry.eventId] = recovery.reason;
        }
      });
      setContextSyncNote_(
        baseContext,
        getFuLifecyclePendingResolutionMessage_(recovery.reason),
        true
      );
      return {
        row,
        context: baseContext,
        result: {
          ok: true,
          status: 'pending_resolution',
          reason: recovery.reason
        }
      };
    }
    const context = recovery && recovery.status === 'recover'
      ? {
        ...baseContext,
        eventId: recovery.eventId,
        recoveredEventId: true
      }
      : baseContext;
    const result = syncPreparedManagedContext_(
      context,
      globalIds.eventLocations,
      runtime
    );
    const retention = getRegistryRetentionForSync_(context, result);
    retention.retainEventIds.forEach(eventId => retainEventIds.push(eventId));
    retention.allowedRemovedEventIds.forEach(eventId => {
      allowedRemovedEventIds.push(eventId);
    });
    if (!result.ok) {
      const pendingEventId = toCellText_(result.eventId || context.eventId);
      if (pendingEventId) {
        pendingSyncByEventId[pendingEventId] =
          result.status || result.message || 'sync_failed';
      }
    }
    return { row, context, result };
  });
  const registry = globalIds.duplicateIds.length
    ? {
      ok: false,
      deferred: true,
      duplicateIds: globalIds.duplicateIds,
      message:
        '其他資料列仍有重複 CalendarEventId；本次唯一事件已處理，' +
        '索引刷新延後至衝突修復後。'
    }
    : refreshCalendarRegistryForSheet_(
      spreadsheet,
      sheet,
      {
        baseline: globalIds.baseline,
        retainEventIds,
        allowedRemovedEventIds,
        pendingSyncByEventId,
        pendingResolutionByEventId,
        globalIdScan: globalIds
      }
    );
  return {
    ok: results.every(item => item.result.ok),
    results,
    registry
  };
}

function syncPreparedManagedContextAt_(spreadsheet, context, options) {
  const settings = options || {};
  const performance = settings.performance || null;
  const sheetScan = settings.sheetScan ||
    buildManagedSheetCalendarScan_(context.sheet);
  const globalIds = resolveGlobalEventIdLocationsForSync_(
    spreadsheet,
    context.sheet,
    sheetScan
  );
  const result = syncPreparedManagedContext_(
    context,
    globalIds.eventLocations,
    buildCalendarSyncRuntime_()
  );
  if (performance) performance.mark('Calendar');
  const retention = getRegistryRetentionForSync_(context, result);
  const pendingEventId = !result.ok
    ? toCellText_(result.eventId || context.eventId)
    : '';
  const registryOptions = {
    ...retention,
    baseline: globalIds.baseline,
    globalIdScan: globalIds,
    pendingSyncByEventId: pendingEventId
      ? {
        [pendingEventId]:
          result.status || result.message || 'sync_failed'
      }
      : {}
  };
  const registry = globalIds.duplicateIds.length
    ? {
      ok: false,
      deferred: true,
      duplicateIds: globalIds.duplicateIds
    }
    : refreshCalendarRegistryForSheet_(
      spreadsheet,
      context.sheet,
      registryOptions
    );
  if (performance) performance.mark('索引更新');
  return { ...result, registry };
}

function getRepairSeedRows_(sheet, range) {
  const rows = [];
  const start = Math.max(2, range.getRow());
  const end = range.getRow() + range.getNumRows() - 1;
  if (isMainTrackingSheetName_(sheet.getName())) {
    for (let row = start; row <= end; row++) rows.push(row);
    return Array.from(new Set(rows));
  }
  if (!isMonthlySheetName_(sheet.getName())) return [];

  const columns = getRequiredMonthlyColumns_(sheet);
  const scan = scanMonthlyBlocks_(sheet, columns);
  for (let row = start; row <= end; row++) {
    const block = scan.blocks.find(item => Number(item.row) === Number(row));
    if (block) {
      block.patientRows.forEach(patient => rows.push(patient.row));
    } else {
      rows.push(row);
    }
  }
  return Array.from(new Set(rows));
}

function getRepairContextKey_(context) {
  return `${Number(context.sheetId)}:${Number(context.row)}`;
}

function getUniqueCalendarRepairOwnerKey_(event, contexts) {
  if (!event || !isSystemManagedCalendarEvent_(event)) return '';
  const titleSlots = toCellText_(event.summary).split(/\s*\|\s*/);
  if (titleSlots.length !== 3) return '';
  const matches = (contexts || []).filter(context => {
    if (
      !context ||
      !context.valid ||
      toSingleLineText_(context.chartNo) !== toSingleLineText_(titleSlots[0]) ||
      toSingleLineText_(context.patientName) !== toSingleLineText_(titleSlots[1])
    ) {
      return false;
    }
    const expected = buildCalendarResource_(context);
    return calendarEndpointMatches_(event.start, expected.start) &&
      calendarEndpointMatches_(event.end, expected.end);
  });
  return matches.length === 1 ? getRepairContextKey_(matches[0]) : '';
}

function readCalendarEventsForRepair_(calendarId, eventIds) {
  const eventsById = {};
  const missingIds = {};
  const errorsById = {};
  Array.from(new Set((eventIds || []).filter(Boolean))).forEach(eventId => {
    try {
      eventsById[eventId] = Calendar.Events.get(calendarId, eventId);
    } catch (err) {
      if (isCalendarNotFoundError_(err)) {
        missingIds[eventId] = true;
      } else {
        errorsById[eventId] = err.message || String(err);
      }
    }
  });
  return { eventsById, missingIds, errorsById };
}

function getRepairEventFingerprint_(event) {
  if (!event) return '';
  return sha256Hex_(JSON.stringify({
    id: toCellText_(event.id),
    summary: toCellText_(event.summary),
    description: toCellText_(event.description),
    colorId: toCellText_(event.colorId),
    startDate: toCellText_(event.start && event.start.date),
    startDateTime: toCellText_(event.start && event.start.dateTime),
    startTimeZone: toCellText_(event.start && event.start.timeZone),
    endDate: toCellText_(event.end && event.end.date),
    endDateTime: toCellText_(event.end && event.end.dateTime),
    endTimeZone: toCellText_(event.end && event.end.timeZone),
    privateKind: getCalendarPrivateProperty_(
      event,
      CALENDAR_PRIVATE_KIND_KEY
    ),
    privateState: getCalendarPrivateProperty_(
      event,
      CALENDAR_PRIVATE_STATE_KEY
    )
  }));
}

function buildRepairPlanFingerprint_(items) {
  return sha256Hex_(JSON.stringify((items || []).map(item => ({
    sheetId: item.context.sheetId,
    row: item.context.row,
    rowHash: item.context.rowHash,
    blockKey: item.context.blockKey,
    currentEventId: item.currentEventId,
    proposedEventId: item.proposedEventId,
    rawTime: item.rawTime,
    normalizedTime: item.normalizedTime,
    action: item.action,
    currentEventFingerprint: getRepairEventFingerprint_(item.currentEvent),
    proposedEventFingerprint: getRepairEventFingerprint_(item.event)
  })).sort((left, right) => {
    return Number(left.sheetId) - Number(right.sheetId) ||
      Number(left.row) - Number(right.row);
  })));
}

function getCalendarEventScheduleText_(event) {
  if (!event || !event.start) return 'Calendar 無可讀事件';
  if (event.start.date) return `${toCellText_(event.start.date)} 全天`;
  if (!event.start.dateTime) return 'Calendar 時間不完整';
  const start = new Date(event.start.dateTime);
  if (isNaN(start.getTime())) return 'Calendar 時間無效';
  return Utilities.formatDate(
    start,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm'
  );
}

function getRepairPreviewSchedule_(context, normalizedTime) {
  if (!context.date) return '日期區塊無效';
  const date = formatDateKey_(context.date);
  const time = normalizedTime ||
    (
      context.timeInfo && context.timeInfo.parsedTime
        ? context.timeInfo.parsedTime.text
        : '全天'
    );
  const hospital = context.kind === 'MONTHLY' && context.hospital
    ? `／${context.hospital}`
    : '';
  return `${date} ${time}${hospital}`;
}

function describeRepairPlanItem_(item) {
  const location = getContextLocationText_(item.context);
  const schedule = getRepairPreviewSchedule_(
    item.context,
    item.normalizedTime
  );
  const messages = {
    create: '建立缺少的 Calendar 事件',
    recreate: '原事件已不存在，使用原列資料重建',
    rebind: '將錯位的事件 ID 重新連回唯一吻合的既有事件',
    update: '更新既有事件的日期、時間或內容',
    normalize: '正規化報到時間並更新事件',
    clear_note: '事件內容正確，清除過期的系統 note',
    none: '目前不需修復',
    conflict: item.message || '無法形成唯一安全對應'
  };
  const calendarSchedule = item.currentEvent
    ? `；目前 Calendar ${getCalendarEventScheduleText_(item.currentEvent)}`
    : item.currentEventId
      ? '；目前 Event ID 無可讀事件'
      : '；目前尚無 Event ID';
  return `${location}：${messages[item.action] || item.action}；` +
    `Sheet 預期 ${schedule}${calendarSchedule}`;
}

function buildSelectedCalendarRepairPlan_(
  spreadsheet,
  selectedSheetId,
  selectedRows,
  options
) {
  const settings = options || {};
  const scan = buildCurrentCalendarScan_(spreadsheet);
  const baseline = readCalendarRegistryStore_();
  const contextsByKey = {};
  scan.contexts.forEach(context => {
    contextsByKey[getRepairContextKey_(context)] = context;
  });

  const selected = {};
  (selectedRows || []).forEach(row => {
    const key = `${Number(selectedSheetId)}:${Number(row)}`;
    if (contextsByKey[key]) selected[key] = contextsByKey[key];
  });
  Object.values(selected).forEach(context => {
    if (!context.eventId) return;
    (scan.eventLocations[context.eventId] || []).forEach(location => {
      if (location && location.sheet && location.row) {
        selected[getRepairContextKey_(location)] = location;
      }
    });
  });
  const contexts = Object.values(selected);
  if (!contexts.length) {
    return {
      ok: false,
      message: '選取範圍沒有可同步的 FU 或月刀表病人列。',
      items: [],
      fingerprint: ''
    };
  }
  if (contexts.length > CALENDAR_REPAIR_MAX_PREVIEW_ROWS) {
    return {
      ok: false,
      message:
        `一次最多修復 ${CALENDAR_REPAIR_MAX_PREVIEW_ROWS} 列；` +
        `目前選取及相關列共 ${contexts.length} 列。`,
      items: [],
      fingerprint: ''
    };
  }

  const baselineEntries = baseline.ok ? baseline.entries : [];
  const candidateIds = {};
  contexts.forEach(context => {
    if (context.eventId) candidateIds[context.eventId] = true;
    baselineEntries.forEach(entry => {
      if (entry.rowHash === context.rowHash && entry.eventId) {
        candidateIds[entry.eventId] = true;
      }
    });
  });
  const runtime = buildCalendarSyncRuntime_();
  if (!runtime.calendarId) {
    return {
      ok: false,
      message: 'Calendar ID 尚未設定。',
      items: [],
      fingerprint: ''
    };
  }
  if (!runtime.calendarServiceAvailable) {
    return {
      ok: false,
      message: 'Calendar Advanced Service 尚未啟用。',
      items: [],
      fingerprint: ''
    };
  }

  const read = readCalendarEventsForRepair_(
    runtime.calendarId,
    Object.keys(candidateIds)
  );
  const involvedKeys = {};
  contexts.forEach(context => {
    involvedKeys[getRepairContextKey_(context)] = true;
  });
  const eventOwnedOutsideSelection = eventId => {
    return (scan.eventLocations[eventId] || []).some(location => {
      return location && location.sheet &&
        !involvedKeys[getRepairContextKey_(location)];
    });
  };

  const unresolved = contexts.some(context => {
    const signature = getContextCalendarBindingSignature_(context);
    return !Object.values(read.eventsById).some(event => {
      return getEventCalendarBindingSignature_(event) === signature &&
        !eventOwnedOutsideSelection(toCellText_(event.id));
    });
  });
  if (unresolved && settings.allowCalendarList !== false) {
    const listed = listConfiguredCalendarEvents_();
    if (!listed.ok) {
      return {
        ok: false,
        message: `無法讀取 Calendar 修復候選：${listed.message}`,
        items: [],
        fingerprint: ''
      };
    }
    listed.events.forEach(event => {
      const eventId = toCellText_(event && event.id);
      if (
        eventId &&
        isSystemManagedCalendarEvent_(event) &&
        !eventOwnedOutsideSelection(eventId)
      ) {
        read.eventsById[eventId] = event;
      }
    });
  }

  const duplicateOwnerKeys = {};
  contexts.forEach(context => {
    const eventId = toCellText_(context.eventId);
    if (
      eventId &&
      (scan.eventLocations[eventId] || []).length > 1 &&
      !duplicateOwnerKeys[eventId]
    ) {
      duplicateOwnerKeys[eventId] = getUniqueCalendarRepairOwnerKey_(
        read.eventsById[eventId],
        contexts
      );
    }
  });

  const items = contexts.map(context => {
    const signature = getContextCalendarBindingSignature_(context);
    const matches = Object.values(read.eventsById).filter(event => {
      return getEventCalendarBindingSignature_(event) === signature &&
        !eventOwnedOutsideSelection(toCellText_(event.id));
    });
    const currentEventId = toCellText_(context.eventId);
    const rawTime = context.columns.TIME
      ? getRowFieldValue_(context.values, context.columns, 'TIME')
      : '';
    const normalizedTime =
      context.kind === 'MONTHLY' &&
      context.timeInfo &&
      context.timeInfo.parsedTime &&
      toCellText_(rawTime) !== context.timeInfo.parsedTime.text
        ? context.timeInfo.parsedTime.text
        : '';
    const item = {
      context,
      currentEventId,
      proposedEventId: '',
      rawTime: toCellText_(rawTime),
      normalizedTime,
      currentEvent: currentEventId
        ? read.eventsById[currentEventId] || null
        : null,
      event: null,
      action: 'none',
      message: ''
    };
    if (!context.valid) {
      item.action = 'conflict';
      item.message =
        context.invalidReason === 'invalid_time'
          ? context.timeInfo.errorMessage
          : '資料列缺少有效日期區塊或病歷號／姓名。';
      return item;
    }
    if (matches.length > 1) {
      item.action = 'conflict';
      item.message = '找到多個內容相同的 Calendar 事件，無法安全自動選擇。';
      return item;
    }
    if (matches.length === 1) {
      item.event = matches[0];
      item.proposedEventId = toCellText_(matches[0].id);
    }

    if (
      !item.proposedEventId &&
      currentEventId &&
      duplicateOwnerKeys[currentEventId] === getRepairContextKey_(context)
    ) {
      item.event = item.currentEvent;
      item.proposedEventId = currentEventId;
    }

    if (
      currentEventId &&
      read.errorsById[currentEventId] &&
      !item.proposedEventId
    ) {
      item.action = 'conflict';
      item.message =
        `Calendar 事件讀取失敗：${read.errorsById[currentEventId]}`;
      return item;
    }
    if (!item.proposedEventId && currentEventId && read.missingIds[currentEventId]) {
      item.proposedEventId = currentEventId;
      item.action = 'recreate';
      return item;
    }
    if (!item.proposedEventId && currentEventId) {
      item.action = 'conflict';
      item.message =
        '目前 Event ID 對應的事件內容與此列不同，且找不到唯一替代事件。';
      return item;
    }
    if (!item.proposedEventId) {
      item.action = 'create';
      return item;
    }
    if (item.proposedEventId !== currentEventId) {
      item.action = 'rebind';
      return item;
    }
    const expected = buildCalendarResource_(context);
    if (!calendarEventMatchesResource_(item.event, expected)) {
      item.action = normalizedTime ? 'normalize' : 'update';
      return item;
    }
    if (normalizedTime) {
      item.action = 'normalize';
      return item;
    }
    const noteCell = getContextNoteCell_(context);
    const note = toCellText_(noteCell.getNote());
    if (
      note.indexOf(getSystemNoteBlockMarkers_(
        CALENDAR_SYNC_NOTE_PREFIX
      ).start) !== -1 ||
      note.indexOf(getSystemNoteBlockMarkers_(
        CALENDAR_CONFLICT_NOTE_PREFIX
      ).start) !== -1
    ) {
      item.action = 'clear_note';
    }
    return item;
  });

  const proposedOwners = {};
  items.forEach(item => {
    if (!item.proposedEventId || item.action === 'conflict') return;
    if (!proposedOwners[item.proposedEventId]) {
      proposedOwners[item.proposedEventId] = [];
    }
    proposedOwners[item.proposedEventId].push(item);
  });
  Object.keys(proposedOwners).forEach(eventId => {
    const owners = proposedOwners[eventId];
    if (owners.length <= 1) return;
    owners.forEach(item => {
      item.action = 'conflict';
      item.message =
        '同一個 Calendar 事件同時吻合多列，無法安全自動重綁。';
    });
  });

  const fingerprint = buildRepairPlanFingerprint_(items);
  return {
    ok: items.every(item => item.action !== 'conflict'),
    message: '',
    items,
    selectedSheetId: Number(selectedSheetId),
    selectedRows: (selectedRows || []).slice(),
    fingerprint
  };
}

function applySelectedCalendarRepairPlan_(spreadsheet, plan) {
  const fresh = buildSelectedCalendarRepairPlan_(
    spreadsheet,
    plan.selectedSheetId,
    plan.selectedRows,
    { allowCalendarList: true }
  );
  if (!fresh.ok || fresh.fingerprint !== plan.fingerprint) {
    throw new Error(
      '預覽後資料列、索引或 Calendar 狀態已改變；本次未寫入，請重新執行修復預覽。'
    );
  }
  const actionable = fresh.items.filter(item => item.action !== 'none');
  actionable.forEach(item => {
    const context = item.context;
    if (item.normalizedTime && context.columns.TIME) {
      context.sheet
        .getRange(context.row, context.columns.TIME)
        .setNumberFormat('@')
        .setValue(item.normalizedTime);
    }
    if (
      item.action === 'rebind' &&
      item.proposedEventId &&
      context.columns.EVENT_ID
    ) {
      context.sheet
        .getRange(context.row, context.columns.EVENT_ID)
        .setValue(item.proposedEventId);
    }
  });

  const refreshedScan = buildCurrentCalendarScan_(spreadsheet);
  const repairKeys = {};
  actionable.forEach(item => {
    repairKeys[getRepairContextKey_(item.context)] = true;
  });
  const relatedDuplicateIds = refreshedScan.duplicateIds.filter(eventId => {
    return (refreshedScan.eventLocations[eventId] || []).some(location => {
      return location && location.sheet &&
        repairKeys[getRepairContextKey_(location)];
    });
  });
  if (relatedDuplicateIds.length) {
    throw new Error(
      '選取範圍在修復預寫入後仍有重複 CalendarEventId；' +
        'Calendar 尚未更新，請重新檢查。'
    );
  }
  const contextsByKey = {};
  refreshedScan.contexts.forEach(context => {
    contextsByKey[getRepairContextKey_(context)] = context;
  });
  const runtime = buildCalendarSyncRuntime_();
  const results = [];
  const sheetRetention = {};
  actionable.forEach(item => {
    const context = contextsByKey[getRepairContextKey_(item.context)];
    if (!context) {
      results.push({
        item,
        result: { ok: false, status: 'row_not_managed' }
      });
      return;
    }
    if (item.action === 'clear_note') {
      clearContextSystemNotes_(context);
      results.push({
        item,
        result: { ok: true, status: 'notes_cleared' }
      });
      return;
    }
    const result = syncPreparedManagedContext_(
      context,
      refreshedScan.eventLocations,
      runtime
    );
    results.push({ item, context, result });
    const key = String(context.sheetId);
    if (!sheetRetention[key]) {
      sheetRetention[key] = {
        sheet: context.sheet,
        retainEventIds: [],
        allowedRemovedEventIds: [],
        pendingSyncByEventId: {}
      };
    }
    const retention = getRegistryRetentionForSync_(context, result);
    retention.retainEventIds.forEach(eventId => {
      sheetRetention[key].retainEventIds.push(eventId);
    });
    retention.allowedRemovedEventIds.forEach(eventId => {
      sheetRetention[key].allowedRemovedEventIds.push(eventId);
    });
    if (!result.ok) {
      const pendingEventId = toCellText_(result.eventId || context.eventId);
      if (pendingEventId) {
        sheetRetention[key].pendingSyncByEventId[pendingEventId] =
          result.status || result.message || 'sync_failed';
      }
    }
    if (
      item.currentEventId &&
      item.currentEventId !== context.eventId &&
      !refreshedScan.eventLocations[item.currentEventId]
    ) {
      sheetRetention[key].allowedRemovedEventIds.push(item.currentEventId);
    }
  });
  const registryDeferred = refreshedScan.duplicateIds.length > 0;
  if (!registryDeferred) {
    Object.values(sheetRetention).forEach(settings => {
      refreshCalendarRegistryForSheet_(
        spreadsheet,
        settings.sheet,
        {
          retainEventIds: settings.retainEventIds,
          allowedRemovedEventIds: settings.allowedRemovedEventIds,
          pendingSyncByEventId: settings.pendingSyncByEventId
        }
      );
    });
  }
  return {
    ok: results.every(item => item.result && item.result.ok),
    results,
    plan: fresh,
    registry: registryDeferred
      ? {
        ok: false,
        deferred: true,
        duplicateIds: refreshedScan.duplicateIds,
        message:
          '其他未選取資料列仍有重複 CalendarEventId；' +
          '本次修復已完成，但索引刷新延後。'
      }
      : { ok: true, deferred: false }
  };
}

function repairSelectedCalendarRows() {
  return runMenuAction_('修復選取列同步', () => {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    assertNoLegacyArchiveState_();
    assertNoActiveAnnualArchiveTransaction_();
    const sheet = spreadsheet.getActiveSheet();
    if (
      !isMainTrackingSheetName_(sheet.getName()) &&
      !isMonthlySheetName_(sheet.getName())
    ) {
      throw new Error('請先在 FU 或 YYYYMM 月刀表選取要修復的資料列。');
    }
    const range = sheet.getActiveRange();
    if (!range || range.getRow() < 2) {
      throw new Error('請選取第 2 列以後的病人列或刀日標題。');
    }
    const selectedRows = getRepairSeedRows_(sheet, range);
    const preview = buildSelectedCalendarRepairPlan_(
      spreadsheet,
      sheet.getSheetId(),
      selectedRows,
      { allowCalendarList: true }
    );
    if (!preview.items.length) {
      throw new Error(preview.message || '選取範圍沒有可修復資料列。');
    }
    const lines = preview.items.map(describeRepairPlanItem_);
    const conflictCount = preview.items.filter(item => {
      return item.action === 'conflict';
    }).length;
    const message = [
      `相關資料列：${preview.items.length}；無法安全修復：${conflictCount}。`,
      '',
      ...lines,
      '',
      conflictCount
        ? '有歧義，未提供執行選項；請先人工確認資料列。'
        : '按下 OK 後會重新核對 fingerprint，再執行上述修復；不會刪除事件。'
    ].join('\n');
    if (conflictCount) {
      SpreadsheetApp.getUi().alert(
        '修復選取列同步預覽',
        message,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      preview.items.forEach(item => {
        if (item.action === 'conflict') {
          setContextSyncNote_(item.context, item.message, true);
        }
      });
      return { ok: false, preview, conflicts: conflictCount };
    }
    const response = SpreadsheetApp.getUi().alert(
      '修復選取列同步預覽',
      message,
      SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
    );
    if (response !== SpreadsheetApp.getUi().Button.OK) {
      return { ok: false, cancelled: true, preview };
    }
    const result = withCalendarSyncLock_(() => {
      return applySelectedCalendarRepairPlan_(spreadsheet, preview);
    });
    const failed = result.results.filter(item => {
      return !item.result || !item.result.ok;
    }).length;
    SpreadsheetApp.getUi().alert(
      '修復選取列同步',
      `已處理 ${result.results.length} 列；失敗 ${failed} 列。` +
        (
          failed
            ? '\n失敗列已保留詳細 note，可再次執行本工具。'
            : '\nCalendar 已回讀驗證，系統 note 已清除。'
        ),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return result;
  });
}
