/** Sheet-only integrity and formatting. No Calendar engine or registry. */
function parseScheduleDate_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return value.getFullYear() >= 2000 && value.getFullYear() <= 2099
      ? new Date(value.getFullYear(), value.getMonth(), value.getDate()) : null;
  }
  return parseFullDate_(value);
}

function scanScheduleEventBindings_(spreadsheet) {
  const eventLocations = {}, structuralActions = [], structuralConflicts = [];
  spreadsheet.getSheets().forEach(sheet => {
    const name = sheet.getName();
    const kind = isMainTrackingSheetName_(name) ? 'FU' : isMonthlySheetName_(name) ? 'MONTHLY' : '';
    if (!kind) return;
    const columns = kind === 'FU' ? getRequiredFuColumns_(sheet) : getRequiredMonthlyColumns_(sheet);
    const width = Math.max(getLastHeaderColumn_(sheet), ...Object.values(columns));
    const count = Math.max(0, sheet.getLastRow() - 1);
    const values = count ? sheet.getRange(2, 1, count, width).getValues() : [];
    values.forEach((row, index) => {
      const id = toCellText_(getRowFieldValue_(row, columns, 'EVENT_ID'));
      if (id) (eventLocations[id] ||= []).push({ sheetId: sheet.getSheetId(), row: index + 2, eventId: id });
    });
    if (kind === 'MONTHLY') {
      const scan = scanMonthlyBlocksFromValues_(sheet, columns, values, 2, width, count + 1);
      getMonthlyStructureIssues_(sheet, scan).forEach(issue =>
        (issue.severity === 'action' ? structuralActions : structuralConflicts).push(issue));
    }
  });
  getAnnualArchiveEventIdLocations_(spreadsheet).forEach(location =>
    (eventLocations[location.eventId] ||= []).push(location));
  return { eventLocations, structuralActions, structuralConflicts,
    duplicateIds: Object.keys(eventLocations).filter(id => eventLocations[id].length > 1) };
}

function assertScheduleIntegrity_(spreadsheet, actionLabel, options) {
  // The second argument was historically a label in setup; accept an options object too.
  const settings = options || (typeof actionLabel === 'object' ? actionLabel : {});
  const scan = scanScheduleEventBindings_(spreadsheet);
  const identities = scanScheduleIdentities_(spreadsheet);
  const issues = [...identities.conflicts, ...scan.duplicateIds.map(id => `CalendarEventId 重複：${id}`),
    ...scan.structuralConflicts.map(issue => issue.message || '刀日結構衝突')];
  if (!settings.allowRepairableStructureActions && scan.structuralActions.length) issues.push('刀日標題需修復');
  if (issues.length) throw new Error(issues.join('\n'));
  return { ok: true, eventCount: Object.keys(scan.eventLocations).length, assigned: identities.assigned };
}

function buildScheduleRowSnapshot_(sheet, patient, block, columns) {
  const values = patient.values;
  const chartNo = toCellText_(getRowFieldValue_(values, columns, 'CHART_NO'));
  const patientName = toCellText_(getRowFieldValue_(values, columns, 'NAME'));
  const timeInfo = resolveScheduleTime_(getRowFieldValue_(values, columns, 'TIME'));
  const invalidReason = !block ? 'no_block' : !(chartNo || patientName) ? 'no_identity'
    : timeInfo.errorMessage ? 'invalid_time' : '';
  return { sheet, sheetId: sheet.getSheetId(), sheetName: sheet.getName(), row: patient.row,
    chartNo, patientName, timeInfo, valid: !invalidReason, invalidReason,
    columns, values, eventId: toCellText_(getRowFieldValue_(values, columns, 'EVENT_ID')),
    syncId: toCellText_(getRowFieldValue_(values, columns, 'SYNC_ID')),
    blockKey: block ? block.blockKey : '',
    rowHash: sha256Hex_(JSON.stringify(values.map(normalizeSortFingerprintCell_))) };
}

function readMonthlyCancelledRows_(sheet, columns, startRow, rowCount) {
  const firstRow = Number(startRow) || 2, count = Math.max(0, Number(rowCount) || 0);
  const cancelled = {};
  if (count && columns.CHART_NO) {
    sheet.getRange(firstRow, columns.CHART_NO, count, 1).getFontLines().forEach((row, index) => {
      if (toCellText_(row[0]).toLowerCase() === 'line-through') cancelled[firstRow + index] = true;
    });
  }
  return cancelled;
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

function processScheduleEdit(e) {
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
  const wrapKeys = isFu ? FU_WRAP_KEYS : MONTHLY_WRAP_KEYS;
  const standardKeys = isFu
    ? CONFIG.FIELD_KEYS
    : CONFIG.MONTHLY_FIELD_KEYS;
  const touchedKeys = getTouchedManagedKeys_(
    e.range,
    columns,
    standardKeys
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

  // Normalize only the edited rows; no external synchronization is performed.
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
  if (wrapTouched && end >= start) sheet.autoResizeRows(start, end - start + 1);
}

// Installed triggers or old open dialogs may still call these names after deployment.
function processRowChange(e) { return processScheduleEdit(e); }
function processCalendarStructureChange() { return externalCalendarResult_(); }
function retryPendingCalendarSync() { return externalCalendarResult_(); }
function checkSyncHealth() { return externalCalendarResult_(); }
function syncPendingChanges() { return externalCalendarResult_(); }
function repairSelectedCalendarRows() { return externalCalendarResult_(); }
function applySelectedCalendarRepairPlan_() { return externalCalendarResult_(); }
function clearContextEventIds() { return externalCalendarResult_(); }
function repairCalendarBindings() { return externalCalendarResult_(); }
