/**
 * FU／月份快照、月份模板、排序、多月份彙整與封存。
 */
const FU_SOURCE_CONDITION_NOTE_PREFIX = '來源 FU Condition（建立時快照）：';

function parseJsonPropertyStrict_(propertyName, fallback) {
  const raw = PropertiesService.getScriptProperties().getProperty(propertyName);
  if (!raw) return fallback;
  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error(
      `Script Property「${propertyName}」損壞；系統已安全停止，請勿重建索引。`
    );
  }
}

function readLegacyArchiveState_() {
  const value = parseJsonPropertyStrict_(LEGACY_ARCHIVED_MONTHS_PROPERTY, {});
  if (!value || Array.isArray(value) || typeof value !== 'object') {
    throw new Error('舊封存狀態格式無效；系統已安全停止。');
  }
  return value;
}

function assertNoLegacyArchiveState_() {
  const legacy = readLegacyArchiveState_();
  if (Object.keys(legacy).length) {
    throw new Error(
      '偵測到舊版「隱藏月份／可解除」封存狀態。請先人工檢查舊封存資料；' +
      '本版本不會自動重新解讀或刪除。'
    );
  }
  return true;
}

function readAnnualArchiveTransaction_() {
  const value = parseJsonPropertyStrict_(
    ANNUAL_ARCHIVE_TRANSACTION_PROPERTY,
    null
  );
  if (value && (Array.isArray(value) || typeof value !== 'object')) {
    throw new Error('永久封存交易狀態格式無效；刀表作業已安全停止。');
  }
  return value;
}

function writeAnnualArchiveTransaction_(transaction) {
  const properties = PropertiesService.getScriptProperties();
  if (!transaction) {
    properties.deleteProperty(ANNUAL_ARCHIVE_TRANSACTION_PROPERTY);
    return;
  }
  const saved = { ...transaction };
  if (saved.version === 2) { delete saved.summary; delete saved.changes; }
  properties.setProperty(
    ANNUAL_ARCHIVE_TRANSACTION_PROPERTY,
    JSON.stringify(saved)
  );
}

function isAnnualArchiveTransactionActive_() {
  try {
    const transaction = readAnnualArchiveTransaction_();
    return Boolean(transaction && !['preview', 'applied', 'rejected'].includes(transaction.status));
  } catch (err) {
    return true;
  }
}

function assertNoActiveAnnualArchiveTransaction_() {
  const transaction = readAnnualArchiveTransaction_();
  if (transaction && !['preview', 'applied', 'rejected'].includes(transaction.status)) {
    throw new Error(
      `永久封存交易尚未完成（${transaction.operationId || '未知識別碼'}）。` +
      '請從「彙整舊月刀表」繼續完成，期間暫停刀表編輯格式處理。'
    );
  }
  assertNativeArchiveAvailable_('');
  return true;
}

function getSheetById_(spreadsheet, sheetId) {
  return spreadsheet.getSheets().find(sheet => {
    return Number(sheet.getSheetId()) === Number(sheetId);
  }) || null;
}

function getActiveMonthlySheets_(spreadsheet) {
  assertNoLegacyArchiveState_();
  assertNoActiveAnnualArchiveTransaction_();
  return spreadsheet.getSheets()
    .filter(sheet => isMonthlySheetName_(sheet.getName()));
}

function isAnnualArchiveSheet_(sheet) {
  if (!sheet) return false;
  if (/^刀表封存_\d{4}$/.test(sheet.getName())) return true;
  try {
    return toCellText_(sheet.getRange(1, 1).getNote()) ===
      ANNUAL_ARCHIVE_NOTE_MARKER;
  } catch (err) {
    return false;
  }
}

function reorderManagedSheets_(spreadsheet) {
  const visible = spreadsheet.getSheets().filter(sheet => !sheet.isSheetHidden());
  const fu = getMainTrackingSheet_(spreadsheet);
  const ivi = spreadsheet.getSheetByName(CONFIG.SHEET_IVI);
  const months = visible
    .filter(sheet => isMonthlySheetName_(sheet.getName()))
    .sort((left, right) => right.getName().localeCompare(left.getName()));
  const archives = visible
    .filter(isAnnualArchiveSheet_)
    .sort((left, right) => right.getName().localeCompare(left.getName()));
  const fixedIds = {};
  [fu, ivi].concat(months, archives).filter(Boolean).forEach(sheet => {
    fixedIds[String(sheet.getSheetId())] = true;
  });
  const others = visible.filter(sheet => !fixedIds[String(sheet.getSheetId())]);
  const ordered = [fu, ivi]
    .concat(months, archives, others)
    .filter(Boolean);
  const previousActive = spreadsheet.getActiveSheet();
  ordered.forEach((sheet, index) => {
    if (sheet.getIndex && sheet.getIndex() === index + 1) return;
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(index + 1);
  });
  if (previousActive && !previousActive.isSheetHidden()) {
    spreadsheet.setActiveSheet(previousActive);
  } else if (fu) {
    spreadsheet.setActiveSheet(fu);
  }
  return ordered.map(sheet => sheet.getName());
}

function meaningfulMonthlyWidth_(sheet) {
  const width = Math.max(sheet.getMaxColumns(), sheet.getLastColumn());
  const range = sheet.getRange(1, 1, Math.max(1, sheet.getMaxRows()), width);
  const values = range.getValues();
  const formulas = typeof range.getFormulas === 'function' ? range.getFormulas() : [];
  const notes = typeof range.getNotes === 'function' ? range.getNotes() : [];
  let last = 0;
  values.forEach((row, rowIndex) => row.forEach((value, column) => {
    if ((value !== '' && value !== null && value !== undefined) ||
        (formulas[rowIndex] || [])[column] || (notes[rowIndex] || [])[column]) {
      last = Math.max(last, column + 1);
    }
  }));
  return Math.max(1, last);
}

function monthlyHeaderLayout_(sheet) {
  const columns = getRequiredMonthlyColumns_(sheet);
  const width = Math.max(meaningfulMonthlyWidth_(sheet), ...Object.values(columns));
  const values = sheet.getRange(1, 1, 1, width).getValues()[0];
  Object.keys(columns).forEach(key => {
    values[columns[key] - 1] = CONFIG.MONTHLY_FIELD_HEADERS[key];
  });
  return values.map(value => toSingleLineText_(value).normalize('NFKC'));
}

function getMonthlyCreationTemplate_(spreadsheet) {
  const sheets = spreadsheet.getSheets().filter(sheet => isMonthlySheetName_(sheet.getName()))
    .sort((left, right) => right.getName().localeCompare(left.getName()));
  if (!sheets.length) return null;
  const layout = JSON.stringify(monthlyHeaderLayout_(sheets[0]));
  if (sheets.some(sheet => JSON.stringify(monthlyHeaderLayout_(sheet)) !== layout)) {
    throw new Error('月份刀表的臨床、技術或人工欄位位置不同；請先對齊所有月份欄位後再建立新月。');
  }
  return sheets[0];
}

function createOrInitializeMonthlySheet_(spreadsheet, monthName) {
  assertNoLegacyArchiveState_();
  assertNoActiveAnnualArchiveTransaction_();
  if (!/^\d{6}$/.test(monthName) || Number(monthName.slice(4)) < 1 ||
      Number(monthName.slice(4)) > 12) {
    throw new Error('月份格式必須是 YYYYMM，例如 202608。');
  }
  if (listArchivedMonthNames_(spreadsheet).indexOf(monthName) !== -1) {
    throw new Error(
      `${monthName} 已永久封存於年度表；若需要重建，請先從私人整檔備份人工復原。`
    );
  }
  let sheet = spreadsheet.getSheetByName(monthName);
  const created = !sheet;
  const template = created ? getMonthlyCreationTemplate_(spreadsheet) : null;
  if (!sheet) sheet = spreadsheet.insertSheet(monthName);
  if (template) {
    const width = meaningfulMonthlyWidth_(template);
    ensureColumnCapacity_(sheet, width);
    template.getRange(1, 1, 1, width).copyTo(sheet.getRange(1, 1, 1, width));
  }
  const format = applyMonthlyFormatting_(sheet, {
    forceWidths: created,
    initializeTextInputs: created,
    enableDiagnosisSummary: created
  });
  if (template) {
    const width = meaningfulMonthlyWidth_(template);
    // Native header copy preserves formulas, manual headers, notes and formats.
    template.getRange(1, 1, 1, width).copyTo(sheet.getRange(1, 1, 1, width));
    format.columns = ensureManagedHeaders_(sheet, 'MONTHLY');
    if (typeof template.getColumnWidth === 'function') {
      for (let column = 1; column <= width; column++) {
        sheet.setColumnWidth(column, template.getColumnWidth(column));
      }
    }
  }
  const addedBlocks = ensureDefaultMonthlyBlocks_(sheet, format.columns);
  return { sheet, created, addedBlocks, columns: format.columns };
}

function createMonthlySurgerySheet() {
  return runMenuAction_('建立新月刀表', () => {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      '建立新月刀表',
      '請輸入月份（YYYYMM）：',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      return { ok: false, cancelled: true };
    }
    const monthName = toCellText_(response.getResponseText());
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const result = createOrInitializeMonthlySheet_(spreadsheet, monthName);
    reorderManagedSheets_(spreadsheet);
    spreadsheet.setActiveSheet(result.sheet);
    ui.alert(
      `${result.created ? '已建立' : '已更新格式'} ${monthName}；` +
      `補上 ${result.addedBlocks} 個第 2／4 個星期一與星期四的高榮日期區塊。`
    );
    return { ok: true, ...result };
  });
}

function buildSurgeryDateInsertionPreview_(sheet, insertRow, date, hospital = HOSPITAL_KAOH) {
  if (!isMonthlySheetName_(sheet.getName())) {
    throw new Error('請先切換到 YYYYMM 月份刀表。');
  }
  if (insertRow < 2) {
    throw new Error('刀日標題不能插入欄位名稱列。');
  }
  if (formatMonthKey_(date) !== sheet.getName()) {
    throw new Error(`日期必須屬於 ${sheet.getName()}。`);
  }
  if (!CONFIG.HOSPITAL_OPTIONS.includes(hospital)) throw new Error('請選擇高榮或聯醫。');
  const columns = getRequiredMonthlyColumns_(sheet);
  const scan = scanMonthlyBlocks_(sheet, columns);
  const conflicts = getMonthlyStructureIssues_(sheet, scan)
    .filter(issue => issue.severity === 'conflict');
  if (conflicts.length) {
    throw new Error(
      `${sheet.getName()} 仍有 ${conflicts.length} 項月份結構衝突；` +
      '請先修正刀日標記與日期。'
    );
  }
  const nextHeader = scan.blocks
    .filter(block => block.row >= insertRow)
    .sort((left, right) => left.row - right.row)[0] || null;
  const endExclusive = nextHeader ? nextHeader.row : Number.MAX_SAFE_INTEGER;
  const contextsByRow = {};
  scan.blocks.forEach(block => {
    block.patientRows.forEach(patient => {
      contextsByRow[patient.row] = buildScheduleRowSnapshot_(
        sheet,
        patient,
        block,
        columns
      );
    });
  });
  scan.orphans.forEach(patient => {
    contextsByRow[patient.row] = buildScheduleRowSnapshot_(
      sheet,
      patient,
      null,
      columns
    );
  });
  const affectedRows = Object.keys(contextsByRow)
    .map(Number)
    .filter(row => row >= insertRow && row < endExclusive)
    .sort((left, right) => left - right);
  const patientRows = affectedRows.filter(row => {
    const context = contextsByRow[row];
    return Boolean(
      context &&
      (context.chartNo || context.patientName || context.eventId)
    );
  });
  const invalidTimeRows = patientRows.filter(row => {
    const context = contextsByRow[row];
    return context && context.invalidReason === 'invalid_time';
  });
  if (invalidTimeRows.length) {
    throw new Error(
      `第 ${invalidTimeRows.join('、')} 列有無效報到時間；` +
      '請先修正後再插入刀日。'
    );
  }
  const invalidRows = patientRows.filter(row => {
    const context = contextsByRow[row];
    return context &&
      context.invalidReason &&
      context.invalidReason !== 'no_block' &&
      context.invalidReason !== 'invalid_time';
  });
  if (invalidRows.length) {
    throw new Error(
      `第 ${invalidRows.join('、')} 列缺少可安全同步的病歷號／姓名或日期資訊；` +
      '請先修正後再插入刀日。'
    );
  }
  const dateKey = formatDateKey_(date);
  const duplicateRows = scan.blocks
    .filter(block => {
      return block.dateKey === dateKey && block.hospital === hospital;
    })
    .map(block => block.row);
  return {
    columns,
    scan,
    insertRow,
    date,
    dateKey,
    hospital,
    duplicateRows,
    affectedRows,
    patientRows,
    fingerprint: sha256Hex_(JSON.stringify({
      sheetId: sheet.getSheetId(),
      insertRow,
      dateKey,
      hospital,
      duplicateRows,
      affected: affectedRows.map(row => {
        const context = contextsByRow[row];
        return [
          row,
          context && context.eventId,
          context && context.blockKey,
          context && context.rowHash
        ];
      })
    }))
  };
}

function insertSurgeryDateAtSelection() {
  return runMenuAction_('選取列新增刀日', () => {
    assertNoActiveAnnualArchiveTransaction_();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    if (!isMonthlySheetName_(sheet.getName())) {
      throw new Error('請先切換到要新增刀日的 YYYYMM 月份刀表。');
    }
    const activeRange = sheet.getActiveRange();
    const insertRow = activeRange.getRow();
    if (insertRow < 2) {
      throw new Error('請選擇第 2 列以後的位置。');
    }
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      '選取列新增刀日',
      `將在 ${sheet.getName()} 第 ${insertRow} 列前插入刀日標題。\n` +
      '請輸入八位數日期（例如 20260908）：',
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) {
      return { ok: false, cancelled: true };
    }
    const date = parseEightDigitDate_(response.getResponseText());
    if (!date) throw new Error('請輸入有效八位數日期 YYYYMMDD，例如 20260908。');

    const hospitalReply = ui.prompt('選擇院區', '請輸入高榮或聯醫：', ui.ButtonSet.OK_CANCEL);
    if (hospitalReply.getSelectedButton() !== ui.Button.OK) return { ok: false, cancelled: true };
    const hospital = toCellText_(hospitalReply.getResponseText());
    const preview = buildSurgeryDateInsertionPreview_(sheet, insertRow, date, hospital);
    const globalIds = scanScheduleEventBindings_(spreadsheet);
    if (globalIds.duplicateIds.length) {
      throw new Error(
        `CalendarEventId 重複：${globalIds.duplicateIds.join('、')}；未插入刀日。`
      );
    }
    const warnings = [];
    if (preview.duplicateRows.length) {
      warnings.push(
        `${preview.dateKey}／${hospital}已存在於第 ` +
        `${preview.duplicateRows.join('、')} 列；仍會建立獨立區塊。`
      );
    }
    if (preview.affectedRows.length) {
      warnings.push(
        `插入後有 ${preview.affectedRows.length} 個資料列會改歸新刀日；` +
        '日曆由指定主機監控後更新。'
      );
    }
    if (warnings.length) {
      const confirmation = ui.alert(
        '確認新增刀日',
        warnings.join('\n') + '\n\n是否繼續？',
        ui.ButtonSet.OK_CANCEL
      );
      if (confirmation !== ui.Button.OK) {
        return { ok: false, cancelled: true };
      }
    }

    return withSheetOperationLock_(() => {
      const currentSheet = getSheetById_(spreadsheet, sheet.getSheetId());
      if (!currentSheet || currentSheet.getName() !== sheet.getName()) {
        throw new Error('月份分頁已變更，請重新操作。');
      }
      const refreshedPreview = buildSurgeryDateInsertionPreview_(
        currentSheet,
        insertRow,
        date,
        hospital
      );
      if (refreshedPreview.fingerprint !== preview.fingerprint) {
        throw new Error('選取位置附近資料已變更，請重新執行新增刀日。');
      }
      assertScheduleIntegrity_(spreadsheet, { allowRepairableStructureActions: true });
      const columns = refreshedPreview.columns;
      const lastColumn = Math.max(
        getLastHeaderColumn_(currentSheet),
        ...Object.values(columns)
      );
      currentSheet.insertRowBefore(insertRow);
      try {
        const values = Array(lastColumn).fill('');
        setRowFieldValue_(values, columns, 'TIME', date);
        setRowFieldValue_(values, columns, 'HOSPITAL', hospital);
        setRowFieldValue_(
          values,
          columns,
          'CHART_NO',
          MONTHLY_DATE_HEADER_MARKER
        );
        setRowFieldValue_(values, columns, 'SYNC_ID', newScheduleSyncId_());
        currentSheet.getRange(insertRow, 1, 1, lastColumn).setValues([values]);
        applyManagedRowFormat_(
          currentSheet,
          insertRow,
          'MONTHLY',
          columns
        );
        const presentation = ensureMonthlyHeaderPresentation_(
          currentSheet,
          insertRow,
          columns,
          values
        );
        if (!presentation.ok) {
          throw new Error('新刀日標題格式或標記驗證失敗。');
        }
      } catch (err) {
        currentSheet.deleteRow(insertRow);
        throw err;
      }

      const affectedRows = refreshedPreview.affectedRows.map(row => row + 1);
      assertScheduleIntegrity_(spreadsheet, { allowRepairableStructureActions: true });
      spreadsheet.setActiveSheet(currentSheet);
      currentSheet.setActiveRange(
        currentSheet.getRange(insertRow, columns.TIME)
      );
      ui.alert(`已在第 ${insertRow} 列建立 ${preview.dateKey}／${hospital}刀日。\n刀表已保存，日曆由指定主機監控後更新。`);
      return { ok: true, row: insertRow, dateKey: preview.dateKey,
        duplicateRows: preview.duplicateRows, affectedRows };
    });
  });
}

function sortCurrentScheduleSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (isMainTrackingSheetName_(sheet.getName())) return sortFuByDate();
  if (isMonthlySheetName_(sheet.getName())) return sortCurrentMonthlySheet();
  return runMenuAction_('整理目前分頁', () => {
    throw new Error('請先切換至 FU 或 YYYYMM 月刀表。');
  });
}

function sortCurrentMonthlySheet() {
  return runMenuAction_('整理月刀表', () => withSheetOperationLock_(() => {
    assertNoActiveAnnualArchiveTransaction_();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    if (!isMonthlySheetName_(sheet.getName())) throw new Error('請先切換到 YYYYMM 月刀表。');
    assertScheduleIntegrity_(spreadsheet);
    const result = rebuildMonthlySheetSorted_(sheet);
    assertScheduleIntegrity_(spreadsheet);
    SpreadsheetApp.getUi().alert(`${sheet.getName()} ${result.changed ? '已整理' : '原本已完成排序'}：${result.blockCount} 個刀日、${result.patientCount} 筆病人。`);
    return result;
  }));
}

function getFuDateSortKey_(value, originalIndex) {
  const date = parseScheduleDate_(value);
  return {
    group: date ? 0 : 1,
    timestamp: date ? date.getTime() : 0,
    originalIndex
  };
}

function buildFuSortKeysFromValues_(values, columns, firstRow) {
  const invalidRows = [];
  const sortKeys = (values || []).map((rowValues, index) => {
    const raw = getRowFieldValue_(rowValues, columns, 'DATE');
    if (toCellText_(raw) && !parseScheduleDate_(raw)) {
      invalidRows.push((Number(firstRow) || 2) + index);
    }
    const key = getFuDateSortKey_(raw, index);
    return [
      String(key.group),
      String(key.timestamp).padStart(13, '0'),
      String(key.originalIndex).padStart(10, '0')
    ].join('|');
  });
  return { invalidRows, sortKeys };
}

function rebuildFuSheetSorted_(sheet, options) {
  const settings = options || {};
  const precomputedValues = settings.precomputedValues;
  const columns = settings.precomputedColumns ||
    getRequiredFuColumns_(sheet);
  const lastRow = Math.max(1, sheet.getLastRow());
  if (lastRow < 2) return { ok: true, rowCount: 0 };
  const rowCount = lastRow - 1;
  if (
    Array.isArray(precomputedValues) &&
    precomputedValues.length === rowCount
  ) {
    const fastPlan = buildFuSortKeysFromValues_(
      precomputedValues,
      columns,
      2
    );
    if (fastPlan.invalidRows.length) {
      throw new Error(
        `FU 有無效日期：第 ${fastPlan.invalidRows.join('、')} 列；未排序。`
      );
    }
    if (areSortKeysInAscendingOrder_(fastPlan.sortKeys)) {
      return {
        ok: true,
        rowCount,
        changed: false,
        sheetScan: settings.precomputedScan || null,
        strategy: 'precomputed_noop'
      };
    }
  }
  const lastColumn = Math.max(
    sheet.getLastColumn(),
    getLastHeaderColumn_(sheet),
    ...Object.values(columns)
  );
  const before = captureSortableRangeState_(
    sheet,
    2,
    rowCount,
    lastColumn
  );
  const plan = buildFuSortKeysFromValues_(before.values, columns, 2);
  if (plan.invalidRows.length) {
    throw new Error(
      `FU 有無效日期：第 ${plan.invalidRows.join('、')} 列；未排序。`
    );
  }
  const sortKeys = plan.sortKeys;
  const changed = !areSortKeysInAscendingOrder_(sortKeys);
  let postSortValues = before.values;
  if (changed) {
    sortRowsWithTemporaryKeys_(
      sheet,
      2,
      rowCount,
      lastColumn,
      sortKeys,
      () => {
        const written = captureSortableRangeState_(
          sheet,
          2,
          rowCount,
          lastColumn
        );
        if (
          JSON.stringify(written.fingerprints) !==
          JSON.stringify(before.fingerprints)
        ) {
          throw new Error('排序後資料、公式或 notes 核對不一致。');
        }
        postSortValues = written.values;
        return { fingerprints: written.fingerprints };
      }
    );
  }
  if (changed && typeof sheet.autoResizeRows === 'function') {
    sheet.autoResizeRows(2, rowCount);
  }
  return {
    ok: true,
    rowCount,
    changed,
    strategy: 'native_range_sort'
  };
}

function sortFuByDate() {
  return runMenuAction_('整理FU日期', () => withSheetOperationLock_(() => {
    assertNoActiveAnnualArchiveTransaction_();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getMainTrackingSheet_(spreadsheet);
    if (!sheet) throw new Error('找不到 FU。');
    assertScheduleIntegrity_(spreadsheet);
    const result = rebuildFuSheetSorted_(sheet);
    assertScheduleIntegrity_(spreadsheet);
    spreadsheet.setActiveSheet(sheet);
    SpreadsheetApp.getUi().alert(`FU ${result.changed ? '已整理' : '原本已排序'} ${result.rowCount} 列；空白日期置底。`);
    return result;
  }));
}

function getSelectedManagedRow_(expectedKind) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  if (row < 2) throw new Error('請選擇病人資料列，不要選欄位名稱列。');
  if (expectedKind === 'FU' && !isMainTrackingSheetName_(sheet.getName())) {
    throw new Error('請先在 FU 選擇一筆資料列。');
  }
  if (
    expectedKind === 'MONTHLY' &&
    !isMonthlySheetName_(sheet.getName())
  ) {
    throw new Error('請先在月份刀表選擇一筆病人列。');
  }
  return { spreadsheet, sheet, row };
}

function getMonthlySnapshotFromRow_(sheet, row) {
  const columns = getRequiredMonthlyColumns_(sheet);
  const scan = scanMonthlyBlocks_(sheet, columns);
  const block = scan.blocks.find(item => {
    return item.patientRows.some(patient => patient.row === row);
  }) || null;
  if (!block || block.row === row) {
    throw new Error('選取列不是有效日期區塊下的病人列。');
  }
  const patient = block.patientRows.find(item => item.row === row);
  const values = patient.values;
  const chartNo = toCellText_(getRowFieldValue_(values, columns, 'CHART_NO'));
  const patientName = toCellText_(getRowFieldValue_(values, columns, 'NAME'));
  if (!chartNo && !patientName) {
    throw new Error('病歷號與姓名皆空白，無法加入 FU。');
  }
  const condition = buildMonthlyCondition_({
    side: getRowFieldValue_(values, columns, 'SIDE'),
    diagnosis: getRowFieldValue_(values, columns, 'DIAGNOSIS'),
    grade: getRowFieldValue_(values, columns, 'GRADE'),
    procedure: getRowFieldValue_(values, columns, 'PROCEDURE'),
    iol: getRowFieldValue_(values, columns, 'IOL'),
    iolTarget: getRowFieldValue_(values, columns, 'IOL_TARGET'),
    iolFinal: getRowFieldValue_(values, columns, 'IOL_FINAL'),
    axis: getRowFieldValue_(values, columns, 'AXIS')
  }, block.date);
  const snapshot = {
    chartNo,
    patientName,
    tel: toCellText_(getRowFieldValue_(values, columns, 'TEL')),
    condition,
    plan: toCellText_(getRowFieldValue_(values, columns, 'PLAN')),
    memo: toCellText_(getRowFieldValue_(values, columns, 'MEMO'))
  };
  return {
    sheet,
    row,
    columns,
    block,
    values,
    snapshot,
    sourceFingerprint: sha256Hex_(
      JSON.stringify({
        sheetId: sheet.getSheetId(),
        row,
        rowHash: sha256Hex_(JSON.stringify(values)),
        blockKey: block.blockKey,
        headerRow: block.row
      })
    )
  };
}

function findFuSnapshotDuplicates_(
  fuSheet,
  snapshot,
  precomputedColumns,
  precomputedLastColumn
) {
  const columns = precomputedColumns ||
    getRequiredFuColumns_(fuSheet);
  const lastColumn = precomputedLastColumn || Math.max(
    fuSheet.getLastColumn(),
    ...Object.values(columns)
  );
  const matches = [];
  if (fuSheet.getLastRow() < 2) return matches;
  const rows = fuSheet.getRange(
    2,
    1,
    fuSheet.getLastRow() - 1,
    lastColumn
  ).getValues();
  rows.forEach((values, index) => {
    const current = {
      chartNo: toCellText_(getRowFieldValue_(values, columns, 'CHART_NO')),
      patientName: toCellText_(getRowFieldValue_(values, columns, 'NAME')),
      tel: toCellText_(getRowFieldValue_(values, columns, 'TEL')),
      condition: toCellText_(getRowFieldValue_(values, columns, 'COND')),
      plan: toCellText_(getRowFieldValue_(values, columns, 'PLAN')),
      memo: toCellText_(getRowFieldValue_(values, columns, 'MEMO'))
    };
    const exact = JSON.stringify(current) === JSON.stringify(snapshot);
    const possible =
      current.chartNo &&
      current.chartNo === snapshot.chartNo &&
      current.condition === snapshot.condition;
    if (exact || possible) {
      matches.push({
        row: index + 2,
        level: exact ? 'exact' : 'possible'
      });
    }
  });
  return matches;
}

function locateCachedFuSnapshot_(
  fuSheet,
  cached,
  snapshot,
  precomputedColumns,
  precomputedLastColumn
) {
  if (!cached || !cached.row) return null;
  if (cached.row > fuSheet.getLastRow()) return null;
  const columns = precomputedColumns ||
    getRequiredFuColumns_(fuSheet);
  const values = fuSheet.getRange(
    cached.row,
    1,
    1,
    precomputedLastColumn || Math.max(
      fuSheet.getLastColumn(),
      ...Object.values(columns)
    )
  ).getValues()[0];
  const fingerprint = sha256Hex_(JSON.stringify({
    chartNo: toCellText_(getRowFieldValue_(values, columns, 'CHART_NO')),
    patientName: toCellText_(getRowFieldValue_(values, columns, 'NAME')),
    tel: toCellText_(getRowFieldValue_(values, columns, 'TEL')),
    condition: toCellText_(getRowFieldValue_(values, columns, 'COND')),
    plan: toCellText_(getRowFieldValue_(values, columns, 'PLAN')),
    memo: toCellText_(getRowFieldValue_(values, columns, 'MEMO'))
  }));
  return fingerprint === sha256Hex_(JSON.stringify(snapshot))
    ? cached.row
    : null;
}

function addSelectedMonthlyRowToFu() {
  return runMenuAction_('月刀表 => FU', () => {
    const performance = createPerformancePhaseTimer_('月刀表 => FU');
    assertNoActiveAnnualArchiveTransaction_();
    const selected = getSelectedManagedRow_('MONTHLY');
    const source = getMonthlySnapshotFromRow_(selected.sheet, selected.row);
    const fuSheet = getMainTrackingSheet_(selected.spreadsheet);
    if (!fuSheet) throw new Error('找不到 FU／All，請確認追蹤分頁名稱。');
    const columns = getRequiredFuColumns_(fuSheet);
    const lastColumn = Math.max(
      fuSheet.getLastColumn(),
      ...Object.values(columns)
    );
    const cache = CacheService.getUserCache();
    const cacheKey = MONTHLY_TO_FU_CACHE_PREFIX + source.sourceFingerprint;
    const cached = safeJsonParse_(cache.get(cacheKey), null);
    const cachedRow = locateCachedFuSnapshot_(
      fuSheet,
      cached,
      source.snapshot,
      columns,
      lastColumn
    );
    if (cachedRow) {
      performance.mark('前置檢查').mark('Sheet 寫入').log();
      selected.spreadsheet.setActiveSheet(fuSheet);
      fuSheet.setActiveRange(fuSheet.getRange(cachedRow, 1));
      SpreadsheetApp.getUi().alert(
        `這次操作已建立 FU 第 ${cachedRow} 列，已直接定位，未重複新增。`
      );
      return { ok: true, reused: true, row: cachedRow };
    }

    const duplicates = findFuSnapshotDuplicates_(
      fuSheet,
      source.snapshot,
      columns,
      lastColumn
    );
    if (duplicates.length) {
      const message = duplicates.slice(0, 8).map(match => {
        return `FU 第 ${match.row} 列：` +
          (match.level === 'exact' ? '六個快照欄位完全相同' : '病歷號＋Condition 相同');
      }).join('\n');
      const response = SpreadsheetApp.getUi().alert(
        '疑似重複追蹤',
        `${message}\n\n同一病人可以有多個追蹤事項；按「確定」仍建立新列。`,
        SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
      );
      if (response !== SpreadsheetApp.getUi().Button.OK) {
        performance.mark('前置檢查').log();
        return { ok: false, cancelled: true };
      }
    }

    performance.mark('前置檢查');
    const targetRow = Math.max(2, fuSheet.getLastRow() + 1);
    if (targetRow > fuSheet.getMaxRows()) fuSheet.insertRowAfter(fuSheet.getMaxRows());
    const values = Array(lastColumn).fill('');
    setRowFieldValue_(values, columns, 'CHART_NO', source.snapshot.chartNo);
    setRowFieldValue_(values, columns, 'NAME', source.snapshot.patientName);
    setRowFieldValue_(values, columns, 'TEL', source.snapshot.tel);
    setRowFieldValue_(values, columns, 'COND', source.snapshot.condition);
    setRowFieldValue_(values, columns, 'PLAN', source.snapshot.plan);
    setRowFieldValue_(values, columns, 'MEMO', source.snapshot.memo);
    setRowFieldValue_(values, columns, 'SYNC_ID', newScheduleSyncId_());
    fuSheet.getRange(targetRow, 1, 1, lastColumn).setValues([values]);
    applyManagedRowFormat_(fuSheet, targetRow, 'FU', columns);
    cache.put(
      cacheKey,
      JSON.stringify({ row: targetRow }),
      CROSS_SHEET_CACHE_SECONDS
    );
    performance.mark('Sheet 寫入').log();
    selected.spreadsheet.setActiveSheet(fuSheet);
    fuSheet.setActiveRange(fuSheet.getRange(targetRow, columns.CHART_NO));
    SpreadsheetApp.getUi().alert(
      `已建立 FU 第 ${targetRow} 列；醫院、Tag、日期與 CalendarEventId 保持空白。`
    );
    return { ok: true, row: targetRow, duplicates };
  });
}

function parseConditionSuggestion_(condition) {
  const text = toSingleLineText_(condition);
  const warnings = [];
  const suggestion = {
    ga: '',
    side: '',
    diagnosis: '',
    grade: '',
    procedure: '',
    iol: '',
    iolTarget: '',
    iolFinal: '',
    axis: '',
    hospital: HOSPITAL_KAOH
  };
  const sides = text.match(/\b(?:OD|OS|OU)\b/gi) || [];
  const uniqueSides = Array.from(new Set(sides.map(value => value.toUpperCase())));
  if (uniqueSides.length === 1) suggestion.side = uniqueSides[0];
  if (uniqueSides.length > 1) {
    suggestion.side = uniqueSides.indexOf('OU') !== -1 ? 'OU' : '';
    warnings.push('Condition 含多個側別，請確認複合手術內容。');
  }
  const split = text.split(/\bs\/p\b/i);
  const diagnosisPart = toCellText_(split[0]);
  const diagnosisMatch = diagnosisPart.match(/^(.+?)(?:\(([^)]*)\))?(?:\s+(?:OD|OS|OU))?$/i);
  if (diagnosisMatch) {
    suggestion.diagnosis = toCellText_(diagnosisMatch[1]);
    suggestion.grade = toCellText_(diagnosisMatch[2]);
  }
  if (split.length > 1) {
    const procedurePart = split.slice(1).join(' s/p ')
      .replace(/\s+\d{8}\s*$/, '')
      .trim();
    const procedureMatch = procedurePart.match(/^(.+?)(?:\(([^)]*)\))?(?:\s+(?:OD|OS|OU))?$/i);
    if (procedureMatch) {
      suggestion.procedure = toCellText_(procedureMatch[1]);
      suggestion.iol = toCellText_(procedureMatch[2]);
    }
  }
  if (!suggestion.diagnosis && text) {
    warnings.push('無法可靠拆解診斷；原 Condition 會保存在目的列 note。');
  }
  if (/\+.*\b(?:OD|OS)\b/i.test(text)) {
    warnings.push('Condition 可能包含複合術式，請在進階欄位確認。');
  }
  return { suggestion, warnings };
}

function getFuSourceForDialog_(sheet, row) {
  const columns = getRequiredFuColumns_(sheet);
  const lastColumn = Math.max(getLastHeaderColumn_(sheet), ...Object.values(columns));
  const values = sheet.getRange(row, 1, 1, lastColumn).getValues()[0];
  const data = {
    chartNo: toCellText_(getRowFieldValue_(values, columns, 'CHART_NO')),
    patientName: toCellText_(getRowFieldValue_(values, columns, 'NAME')),
    tel: toCellText_(getRowFieldValue_(values, columns, 'TEL')),
    condition: toCellText_(getRowFieldValue_(values, columns, 'COND')),
    plan: toCellText_(getRowFieldValue_(values, columns, 'PLAN')),
    memo: toCellText_(getRowFieldValue_(values, columns, 'MEMO'))
  };
  if (!data.chartNo && !data.patientName) {
    throw new Error('病歷號與姓名皆空白，無法安排至月份刀表。');
  }
  return {
    columns,
    values,
    data,
    fingerprint: sha256Hex_(JSON.stringify({
      sheetId: sheet.getSheetId(),
      row,
      data
    }))
  };
}

function buildFuToMonthlyDialogData_() {
  const selected = getSelectedManagedRow_('FU');
  const source = getFuSourceForDialog_(selected.sheet, selected.row);
  const parsed = parseConditionSuggestion_(source.data.condition);
  const months = getActiveMonthlySheets_(selected.spreadsheet)
    .sort((left, right) => right.getName().localeCompare(left.getName()))
    .map(sheet => sheet.getName());
  return {
    fuRow: selected.row,
    sourceFingerprint: source.fingerprint,
    requestToken: Utilities.getUuid(),
    source: {
      chartNo: source.data.chartNo,
      name: source.data.patientName,
      tel: source.data.tel,
      condition: source.data.condition
    },
    options: {
      hospitals: CONFIG.HOSPITAL_OPTIONS,
      ga: ['', 'GA'],
      sides: ['', 'OD', 'OS', 'OU'],
      diagnoses: CONFIG.MONTHLY_DIAGNOSIS_OPTIONS,
      procedures: CONFIG.MONTHLY_PROCEDURE_OPTIONS
    },
    suggestion: {
      ...parsed.suggestion,
      plan: source.data.plan,
      memo: source.data.memo
    },
    warnings: parsed.warnings,
    requiresConfirmation: parsed.warnings.length > 0,
    months
  };
}

function getFuToMonthlyBlockOptions(monthName) {
  const normalizedMonth = toCellText_(monthName);
  if (
    !/^\d{6}$/.test(normalizedMonth) ||
    Number(normalizedMonth.slice(4)) < 1 ||
    Number(normalizedMonth.slice(4)) > 12
  ) {
    throw new Error('月份格式必須是 YYYYMM，例如 202608。');
  }
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  assertNoActiveAnnualArchiveTransaction_();
  const sheet = spreadsheet.getSheetByName(normalizedMonth);
  if (!sheet) {
    return {
      month: normalizedMonth,
      exists: false,
      blocks: []
    };
  }
  return {
    month: normalizedMonth,
    exists: true,
    blocks: getMonthlyBlockOptions_(sheet)
  };
}

function showFuToMonthlyDialog() {
  return runMenuAction_('FU => 月刀表', () => {
    const data = buildFuToMonthlyDialogData_();
    const template = HtmlService.createTemplateFromFile('fu_to_monthly');
    template.initialDataJson = JSON.stringify(data).replace(/<\//g, '<\\/');
    SpreadsheetApp.getUi().showModalDialog(
      template.evaluate().setWidth(650).setHeight(720),
      'FU => 月刀表'
    );
    return { ok: true };
  });
}

function showIolListDialog() {
  return runMenuAction_('水晶體清單', () => {
    SpreadsheetApp.getUi().showModalDialog(
      HtmlService.createHtmlOutputFromFile('iol_list').setWidth(620).setHeight(620),
      '水晶體清單'
    );
    return { ok: true };
  });
}

function getIolListDates() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const dates = new Set();
  getActiveMonthlySheets_(spreadsheet).forEach(sheet => {
    scanMonthlyBlocks_(sheet, getRequiredMonthlyColumns_(sheet)).blocks.forEach(block => {
      dates.add(formatCompactDate_(block.date));
    });
  });
  return Array.from(dates).sort().reverse();
}

function buildIolListForDate_(sheet, date, includeCancelled) {
  const columns = getRequiredMonthlyColumns_(sheet);
  const scan = scanMonthlyBlocks_(sheet, columns);
  const dateKey = formatDateKey_(date);
  const blocks = scan.blocks.filter(block => formatDateKey_(block.date) === dateKey);
  if (scan.hybrids.length || scan.orphans.length) {
    throw new Error('月表有日期／事件 ID 衝突或孤立病人列，請修復結構後再輸出清單。');
  }
  const patientRows = blocks.flatMap(block => block.patientRows.map(patient => patient.row));
  const firstRow = patientRows.length ? Math.min(...patientRows) : 2;
  const rowCount = patientRows.length ? Math.max(...patientRows) - firstRow + 1 : 0;
  const lastDisplayColumn = Math.max(columns.CHART_NO, columns.NAME, columns.IOL, columns.IOL_FINAL);
  const cancelled = readMonthlyCancelledRows_(sheet, columns, firstRow, rowCount);
  const display = rowCount
    ? sheet.getRange(firstRow, 1, rowCount, lastDisplayColumn).getDisplayValues() : [];
  const lines = [`水晶體清單 ${formatCompactDate_(date)}`];
  let count = 0;
  let excluded = 0;
  blocks.forEach(block => {
    const rows = [];
    block.patientRows.forEach(patient => {
      const values = display[patient.row - firstRow];
      if (!toCellText_(getRowFieldValue_(values, columns, 'NAME')) &&
        !toCellText_(getRowFieldValue_(values, columns, 'CHART_NO'))) return;
      if (cancelled[patient.row] && !includeCancelled) { excluded++; return; }
      const name = toSingleLineText_(getRowFieldValue_(values, columns, 'NAME')) || '姓名未填';
      const iol = toSingleLineText_(getRowFieldValue_(values, columns, 'IOL')) || '未填';
      const final = toSingleLineText_(getRowFieldValue_(values, columns, 'IOL_FINAL')) || '未填';
      rows.push(`${name}${cancelled[patient.row] ? '（已取消）' : ''}｜${iol}｜${final}`);
      count++;
    });
    if (rows.length) lines.push('', `${block.hospital}（第 ${block.row} 列刀日）`, '姓名｜IOL｜IOL Final', ...rows);
  });
  if (!count) lines.push('此刀日沒有符合條件的病人資料。');
  return { date: formatCompactDate_(date), count, excluded, text: lines.join('\n') };
}

function getIolListForDate(dateText, includeCancelled) {
  assertNoActiveAnnualArchiveTransaction_();
  const date = parseEightDigitDate_(dateText);
  if (!date) throw new Error('請選擇有效八位數刀日，例如 20260908。');
  const sheets = getActiveMonthlySheets_(SpreadsheetApp.getActiveSpreadsheet());
  const results = sheets.map(sheet => buildIolListForDate_(sheet, date, includeCancelled === true));
  return { date: formatCompactDate_(date),
    count: results.reduce((sum, result) => sum + result.count, 0),
    excluded: results.reduce((sum, result) => sum + result.excluded, 0),
    text: results.filter(result => result.count || result.excluded).map(result => result.text).join('\n\n') ||
      `水晶體清單 ${formatCompactDate_(date)}\n此刀日沒有符合條件的病人資料。` };
}

function findMonthlyDuplicateCandidates_(
  sheet,
  request,
  date,
  hospital,
  precomputedScan,
  precomputedColumns
) {
  const exactPayload = {
    chartNo: request.chartNo,
    patientName: request.patientName,
    date: formatDateKey_(date),
    hospital,
    time: request.time,
    ga: request.ga,
    side: request.side,
    diagnosis: request.diagnosis,
    grade: request.grade,
    procedure: request.procedure,
    iol: request.iol,
    iolTarget: request.iolTarget,
    iolFinal: request.iolFinal,
    axis: request.axis,
    plan: request.plan,
    memo: request.memo
  };
  const matches = [];
  if (sheet) {
    const columns =
      precomputedColumns || getRequiredMonthlyColumns_(sheet);
    const scan =
      precomputedScan || scanMonthlyBlocks_(sheet, columns);
    scan.blocks.forEach(block => {
      block.patientRows.forEach(patient => {
        const values = patient.values;
        const current = {
          chartNo: toCellText_(getRowFieldValue_(values, columns, 'CHART_NO')),
          patientName: toCellText_(getRowFieldValue_(values, columns, 'NAME')),
          date: block.dateKey,
          hospital: block.hospital,
          time: toCellText_(getRowFieldValue_(values, columns, 'TIME')),
          ga: toCellText_(getRowFieldValue_(values, columns, 'GA')),
          side: toCellText_(getRowFieldValue_(values, columns, 'SIDE')),
          diagnosis: toCellText_(getRowFieldValue_(values, columns, 'DIAGNOSIS')),
          grade: toCellText_(getRowFieldValue_(values, columns, 'GRADE')),
          procedure: toCellText_(getRowFieldValue_(values, columns, 'PROCEDURE')),
          iol: toCellText_(getRowFieldValue_(values, columns, 'IOL')),
          iolTarget: toCellText_(getRowFieldValue_(values, columns, 'IOL_TARGET')),
          iolFinal: toCellText_(getRowFieldValue_(values, columns, 'IOL_FINAL')),
          axis: toCellText_(getRowFieldValue_(values, columns, 'AXIS')),
          plan: toCellText_(getRowFieldValue_(values, columns, 'PLAN')),
          memo: toCellText_(getRowFieldValue_(values, columns, 'MEMO'))
        };
        const exact = JSON.stringify(current) === JSON.stringify(exactPayload);
        const identityMatches = request.chartNo
          ? current.chartNo === request.chartNo
          : current.patientName === request.patientName;
        const possible =
          identityMatches &&
          current.date === exactPayload.date &&
          current.hospital === exactPayload.hospital;
        if (exact || possible) {
          matches.push({
            sheetName: sheet.getName(),
            row: patient.row,
            level: exact ? 'exact' : 'possible',
            reason: exact
              ? '手術日期、醫院及結構化資料完全相同'
              : '同一病人、手術日期與醫院相同'
          });
        }
      });
    });
  }
  return {
    matches,
    fingerprint: sha256Hex_(JSON.stringify(matches))
  };
}

function locateCachedMonthlySchedule_(spreadsheet, cached, identity) {
  if (!cached || !/^[a-f0-9]{32}$/.test(toCellText_(cached.syncId))) {
    throw new Error('重送記錄缺少永久排程 ID，請關閉視窗後重新核對刀表。');
  }
  const matches = [];
  getActiveMonthlySheets_(spreadsheet).forEach(sheet => {
    const columns = getRequiredMonthlyColumns_(sheet);
    const scan = scanMonthlyBlocks_(sheet, columns);
    scan.blocks.forEach(block => block.patientRows.forEach(patient => {
      if (toCellText_(getRowFieldValue_(patient.values, columns, 'SYNC_ID')) !== cached.syncId) return;
      const chartNo = toCellText_(getRowFieldValue_(patient.values, columns, 'CHART_NO'));
      const patientName = toCellText_(getRowFieldValue_(patient.values, columns, 'NAME'));
      if (identity && (identity.chartNo ? chartNo !== identity.chartNo : patientName !== identity.patientName)) {
        throw new Error('永久排程 ID 綁定的病人已變更，請先核對刀表。');
      }
      matches.push({ sheet, row: patient.row, syncId: cached.syncId });
    }));
  });
  if (matches.length > 1) throw new Error('相同請求的排程識別重複，請先核對。');
  return matches[0] || null;
}

function submitFuToMonthlySchedule(payload) {
  const request = payload || {};
  return withSheetOperationLock_(() => {
    const performance = createPerformancePhaseTimer_('FU => 月刀表');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    assertNoActiveAnnualArchiveTransaction_();
    assertScheduleIntegrity_(spreadsheet, { allowRepairableStructureActions: true });
    const fuSheet = getMainTrackingSheet_(spreadsheet);
    if (!fuSheet) throw new Error('找不到 FU。');
    const fuRow = Number(request.fuRow);
    const source = getFuSourceForDialog_(fuSheet, fuRow);
    if (source.fingerprint !== toCellText_(request.sourceFingerprint)) {
      throw new Error('FU 來源列已變更，請關閉視窗後重新開啟。');
    }
    const targetMonth = toCellText_(request.targetMonth);
    const date = parseEightDigitDate_(request.date);
    if (!date) throw new Error('請輸入有效八位數手術日期，例如 20260908。');
    if (formatMonthKey_(date) !== targetMonth) {
      throw new Error('手術日期必須屬於目標月份。');
    }
    const hospital = toCellText_(request.hospital) || HOSPITAL_KAOH;
    if (CONFIG.HOSPITAL_OPTIONS.indexOf(hospital) === -1) {
      throw new Error('醫院必須是高榮或聯醫。');
    }
    const timeInfo = resolveScheduleTime_(request.time);
    if (timeInfo.errorMessage) throw new Error(timeInfo.errorMessage);
    const normalizedTime = timeInfo.parsedTime
      ? timeInfo.parsedTime.text
      : getRawTimeText_(request.time);
    const normalizedRequest = {
      ...request,
      chartNo: source.data.chartNo,
      patientName: source.data.patientName,
      tel: source.data.tel,
      time: normalizedTime,
      hospital
    };

    const requestToken = toCellText_(request.requestToken);
    if (!/^[a-f0-9-]{32,36}$/i.test(requestToken)) throw new Error('請重新開啟排程視窗以取得有效請求識別碼。');
    const cache = CacheService.getUserCache();
    const tokenKey = FU_TO_MONTHLY_CACHE_PREFIX + requestToken;
    const cached = safeJsonParse_(cache.get(tokenKey), null);
    if (cached) {
      const located = locateCachedMonthlySchedule_(spreadsheet, cached, source.data);
      if (located) {
        performance.mark('前置檢查');
        cache.put(
          tokenKey,
          JSON.stringify({ syncId: located.syncId }),
          CROSS_SHEET_CACHE_SECONDS
        );
        performance.log();
        return {
          ok: true,
          reused: true,
          sheetName: located.sheet.getName(),
          row: located.row,
          message:
            `相同 request token 已建立 ${located.sheet.getName()} 第 ${located.row} 列；` +
            '刀表已存在，日曆由指定主機監控。'
        };
      }
    }

    const existingMonth = spreadsheet.getSheetByName(targetMonth);
    let targetColumns = null;
    let targetScan = null;
    if (existingMonth) {
      if (!isMonthlySheetName_(existingMonth.getName())) {
        throw new Error(`${targetMonth} 不是有效月份刀表。`);
      }
      targetColumns = getRequiredMonthlyColumns_(existingMonth);
      targetScan = scanMonthlyBlocks_(existingMonth, targetColumns);
    }
    const duplicate = findMonthlyDuplicateCandidates_(
      existingMonth,
      normalizedRequest,
      date,
      hospital,
      targetScan,
      targetColumns
    );
    if (
      duplicate.matches.length &&
      toCellText_(request.confirmedDuplicateFingerprint) !==
        duplicate.fingerprint
    ) {
      performance.mark('前置檢查').log();
      return {
        ok: false,
        status: 'needs_duplicate_confirmation',
        duplicateFingerprint: duplicate.fingerprint,
        matches: duplicate.matches,
        message: '找到疑似重複手術；確認後仍可建立。'
      };
    }

    performance.mark('前置檢查');
    let monthResult;
    if (existingMonth) {
      monthResult = {
        sheet: existingMonth,
        created: false,
        addedBlocks: 0,
        columns: targetColumns
      };
    } else {
      monthResult = createOrInitializeMonthlySheet_(
        spreadsheet,
        targetMonth
      );
    }
    const sheet = monthResult.sheet;
    const columns = monthResult.columns || getRequiredMonthlyColumns_(sheet);
    if (!targetScan) targetScan = scanMonthlyBlocks_(sheet, columns);
    let block;
    const selectedBlockRow = Number(request.selectedBlockRow);
    if (selectedBlockRow) {
      block = resolveMonthlyBlockRef_(
        sheet,
        {
          headerRow: selectedBlockRow,
          dateKey: formatDateKey_(date),
          hospital
        },
        columns,
        targetScan
      );
      if (!block) {
        throw new Error(
          '所選日期區塊的位置或內容已變更，請關閉視窗後重新選擇。'
        );
      }
    } else {
      block = findOrCreateMonthlyBlock_(
        sheet,
        date,
        hospital,
        columns,
        targetScan
      );
    }
    const row = insertPatientAtBlockEnd_(
      sheet,
      block,
      {
        TIME: normalizedTime,
        CHART_NO: source.data.chartNo,
        NAME: source.data.patientName,
        TEL: source.data.tel,
        GA: toCellText_(request.ga),
        SIDE: toCellText_(request.side),
        DIAGNOSIS: toCellText_(request.diagnosis),
        GRADE: toCellText_(request.grade),
        PROCEDURE: toCellText_(request.procedure),
        PLAN: toCellText_(request.plan),
        IOL: toCellText_(request.iol),
        IOL_TARGET: toCellText_(request.iolTarget),
        IOL_FINAL: toCellText_(request.iolFinal),
        AXIS: toCellText_(request.axis),
        MEMO: toCellText_(request.memo)
      },
      columns,
      targetScan
    );
    if (source.data.condition) {
      sheet.getRange(row, columns.DIAGNOSIS).setNote(
        FU_SOURCE_CONDITION_NOTE_PREFIX + source.data.condition
      );
    }
    const written = sheet.getRange(
      row,
      1,
      1,
      Math.max(getLastHeaderColumn_(sheet), ...Object.values(columns))
    ).getValues()[0];
    if (
      toCellText_(getRowFieldValue_(written, columns, 'CHART_NO')) !==
        source.data.chartNo ||
      toCellText_(getRowFieldValue_(written, columns, 'NAME')) !==
        source.data.patientName
    ) {
      throw new Error('月表列回讀驗證失敗，請核對剛建立的資料列。');
    }
    const createdContext = buildScheduleRowSnapshot_(
      sheet,
      { row, values: written },
      block,
      columns
    );
    if (!createdContext.valid) {
      throw new Error('新建月份資料列無法建立安全重送指紋，請核對剛建立的資料列。');
    }
    performance.mark('Sheet 寫入');
    const cacheRecord = { syncId: createdContext.syncId };
    cache.put(
      tokenKey,
      JSON.stringify(cacheRecord),
      CROSS_SHEET_CACHE_SECONDS
    );
    if (monthResult.created) reorderManagedSheets_(spreadsheet);
    performance.log();
    spreadsheet.setActiveSheet(sheet);
    sheet.setActiveRange(sheet.getRange(row, columns.CHART_NO));
    return {
      ok: true,
      status: 'created',
      sheetName: sheet.getName(),
      row,
      message:
        `已建立 ${sheet.getName()} 第 ${row} 列；` +
        '日曆由指定主機監控後更新。'
    };
  });
}
