/**
 * 永久年度封存。
 *
 * 封存表不參與 Calendar 同步；來源月表只會在私人整檔備份、
 * 暫存年度表及回讀驗證全部成功後刪除。
 */
const ARCHIVE_ROW_TYPES = {
  MONTH: 'MONTH_HEADER',
  DATE: 'DATE_HEADER',
  PATIENT: 'PATIENT'
};
const ARCHIVE_STAGE_PREFIX = '__ARCHIVE_STAGE_';
const ARCHIVE_OLD_PREFIX = '__ARCHIVE_OLD_';

function getAnnualArchiveName_(year) {
  return `${ANNUAL_ARCHIVE_PREFIX}${year}`;
}

function isClosedMonthForArchive_(monthName, currentMonth) {
  return isMonthlySheetName_(monthName) &&
    monthName < (currentMonth || formatMonthKey_(new Date()));
}

function getAnnualArchiveRequiredHeaders_() {
  return CONFIG.MONTHLY_HEADERS
    .filter(Boolean)
    .concat([
      ANNUAL_ARCHIVE_SOURCE_MONTH_HEADER,
      ANNUAL_ARCHIVE_ROW_TYPE_HEADER
    ]);
}

function getAnnualArchiveHeaderInfo_(sheet) {
  const lastColumn = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(
    ANNUAL_ARCHIVE_HEADER_ROW,
    1,
    1,
    lastColumn
  ).getValues()[0].map(toCellText_);
  const occurrences = {};
  headers.forEach((header, index) => {
    if (!header) return;
    if (!occurrences[header]) occurrences[header] = [];
    occurrences[header].push(index + 1);
  });
  const duplicates = Object.keys(occurrences)
    .filter(header => occurrences[header].length > 1);
  if (duplicates.length) {
    throw new Error(
      `${sheet.getName()} 封存欄名重複：${duplicates.join('、')}。`
    );
  }
  const missing = getAnnualArchiveRequiredHeaders_()
    .filter(header => !occurrences[header]);
  if (missing.length) {
    throw new Error(
      `${sheet.getName()} 封存結構缺少：${missing.join('、')}。`
    );
  }
  const columns = {};
  Object.keys(occurrences).forEach(header => {
    columns[header] = occurrences[header][0];
  });
  return {
    headers,
    columns,
    lastColumn,
    eventColumn: columns[CONFIG.MONTHLY_FIELD_HEADERS.EVENT_ID],
    sourceMonthColumn: columns[ANNUAL_ARCHIVE_SOURCE_MONTH_HEADER],
    rowTypeColumn: columns[ANNUAL_ARCHIVE_ROW_TYPE_HEADER]
  };
}

function scanAnnualArchiveStructure_(sheet) {
  const info = getAnnualArchiveHeaderInfo_(sheet);
  const lastRow = sheet.getLastRow();
  if (lastRow <= ANNUAL_ARCHIVE_HEADER_ROW) {
    return {
      info,
      months: [],
      monthRanges: [],
      eventLocations: [],
      lastRow
    };
  }
  const rowCount = lastRow - ANNUAL_ARCHIVE_HEADER_ROW;
  const values = sheet.getRange(
    ANNUAL_ARCHIVE_HEADER_ROW + 1,
    1,
    rowCount,
    info.lastColumn
  ).getValues();
  const months = [];
  const monthRanges = [];
  const eventLocations = [];
  const seenMonths = {};
  let current = null;
  values.forEach((rowValues, index) => {
    const row = ANNUAL_ARCHIVE_HEADER_ROW + 1 + index;
    const sourceMonth = toCellText_(
      rowValues[info.sourceMonthColumn - 1]
    );
    const rowType = toCellText_(rowValues[info.rowTypeColumn - 1]);
    const hasVisibleContent = rowValues.some((value, columnIndex) => {
      const column = columnIndex + 1;
      return column !== info.sourceMonthColumn &&
        column !== info.rowTypeColumn &&
        column !== info.eventColumn &&
        toCellText_(value);
    });
    if (!sourceMonth && !rowType && !hasVisibleContent) return;
    if (
      !isMonthlySheetName_(sourceMonth) ||
      sourceMonth.slice(0, 4) !== sheet.getName().slice(-4)
    ) {
      throw new Error(
        `${sheet.getName()} 第 ${row} 列 ArchiveSourceMonth 無效。`
      );
    }
    if (rowType === ARCHIVE_ROW_TYPES.MONTH) {
      if (seenMonths[sourceMonth]) {
        throw new Error(
          `${sheet.getName()} 的 ${sourceMonth} 出現多個月份區段。`
        );
      }
      seenMonths[sourceMonth] = true;
      current = {
        month: sourceMonth,
        startRow: row,
        endRow: row,
        patientCount: 0,
        blockCount: 0
      };
      months.push(sourceMonth);
      monthRanges.push(current);
    } else {
      if (!current || current.month !== sourceMonth) {
        throw new Error(
          `${sheet.getName()} 第 ${row} 列不在有效月份分隔列之下。`
        );
      }
      if (
        rowType !== ARCHIVE_ROW_TYPES.DATE &&
        rowType !== ARCHIVE_ROW_TYPES.PATIENT
      ) {
        throw new Error(
          `${sheet.getName()} 第 ${row} 列 ArchiveRowType 無效。`
        );
      }
      current.endRow = row;
      if (rowType === ARCHIVE_ROW_TYPES.DATE) current.blockCount++;
      if (rowType === ARCHIVE_ROW_TYPES.PATIENT) current.patientCount++;
    }
    const eventId = toCellText_(rowValues[info.eventColumn - 1]);
    if (eventId) {
      if (rowType !== ARCHIVE_ROW_TYPES.PATIENT) {
        throw new Error(
          `${sheet.getName()} 第 ${row} 列非病人列卻含 CalendarEventId。`
        );
      }
      eventLocations.push({
        eventId,
        sheetId: sheet.getSheetId(),
        sheetName: sheet.getName(),
        row,
        kind: 'ARCHIVE',
        archived: true,
        sourceMonth
      });
    }
  });
  const sorted = months.slice().sort();
  if (JSON.stringify(sorted) !== JSON.stringify(months)) {
    throw new Error(`${sheet.getName()} 月份區段未依時間遞增。`);
  }
  return { info, months, monthRanges, eventLocations, lastRow };
}

function listArchivedMonthNames_(spreadsheet) {
  const result = [];
  spreadsheet.getSheets()
    .filter(isAnnualArchiveSheet_)
    .forEach(sheet => {
      scanAnnualArchiveStructure_(sheet).months.forEach(month => {
        if (result.indexOf(month) !== -1) {
          throw new Error(`永久封存月份 ${month} 出現在多張年度表。`);
        }
        result.push(month);
      });
    });
  return result.sort();
}

function getAnnualArchiveEventIdLocations_(spreadsheet) {
  const locations = [];
  spreadsheet.getSheets()
    .filter(isAnnualArchiveSheet_)
    .forEach(sheet => {
      const info = getAnnualArchiveHeaderInfo_(sheet);
      const rowCount = Math.max(
        0,
        sheet.getLastRow() - ANNUAL_ARCHIVE_HEADER_ROW
      );
      if (!rowCount) return;
      const firstSystemColumn = Math.min(
        info.eventColumn,
        info.sourceMonthColumn,
        info.rowTypeColumn
      );
      const lastSystemColumn = Math.max(
        info.eventColumn,
        info.sourceMonthColumn,
        info.rowTypeColumn
      );
      const systemValues = sheet.getRange(
        ANNUAL_ARCHIVE_HEADER_ROW + 1,
        firstSystemColumn,
        rowCount,
        lastSystemColumn - firstSystemColumn + 1
      ).getValues();
      systemValues.forEach((rowValue, index) => {
        const eventId = toCellText_(
          rowValue[info.eventColumn - firstSystemColumn]
        );
        if (!eventId) return;
        const sourceMonth = toCellText_(
          rowValue[info.sourceMonthColumn - firstSystemColumn]
        );
        const rowType = toCellText_(
          rowValue[info.rowTypeColumn - firstSystemColumn]
        );
        if (
          !isMonthlySheetName_(sourceMonth) ||
          rowType !== ARCHIVE_ROW_TYPES.PATIENT
        ) {
          throw new Error(
            `${sheet.getName()} 第 ${ANNUAL_ARCHIVE_HEADER_ROW + 1 + index} ` +
            '列的封存系統欄無效。'
          );
        }
        locations.push({
          eventId,
          sheetId: sheet.getSheetId(),
          sheetName: sheet.getName(),
          row: ANNUAL_ARCHIVE_HEADER_ROW + 1 + index,
          kind: 'ARCHIVE',
          archived: true,
          sourceMonth
        });
      });
    });
  return locations;
}

function getMonthlyCustomHeaders_(headers) {
  const standard = {};
  CONFIG.MONTHLY_HEADERS.filter(Boolean).forEach(header => {
    standard[header] = true;
  });
  const seen = {};
  return (headers || []).map(toCellText_).filter(header => {
    if (!header || standard[header] || seen[header]) return false;
    if (
      header === ANNUAL_ARCHIVE_SOURCE_MONTH_HEADER ||
      header === ANNUAL_ARCHIVE_ROW_TYPE_HEADER
    ) {
      throw new Error(`月份刀表不可使用保留欄名「${header}」。`);
    }
    seen[header] = true;
    return true;
  });
}

function normalizeArchiveFingerprintValue_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return `DATE:${Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      "yyyy-MM-dd'T'HH:mm:ss"
    )}`;
  }
  return toCellText_(value);
}

function fingerprintArchiveRows_(headers, rows) {
  const system = {};
  system[ANNUAL_ARCHIVE_SOURCE_MONTH_HEADER] = true;
  system[ANNUAL_ARCHIVE_ROW_TYPE_HEADER] = true;
  const canonicalHeaders = (headers || [])
    .map(toCellText_)
    .filter(header => header && !system[header])
    .sort();
  const indexByHeader = {};
  (headers || []).forEach((header, index) => {
    const normalized = toCellText_(header);
    if (normalized) indexByHeader[normalized] = index;
  });
  const normalizedRows = (rows || []).map(item => {
    const snapshot = item.snapshot || item;
    const fields = [];
    canonicalHeaders.forEach(header => {
      const index = indexByHeader[header];
      if (index === undefined) return;
      const value = normalizeArchiveFingerprintValue_(
        (snapshot.values || [])[index]
      );
      const formula = toCellText_((snapshot.formulas || [])[index]);
      const note = toCellText_((snapshot.notes || [])[index]);
      fields.push([
        header,
        value,
        formula,
        note,
        toCellText_((snapshot.backgrounds || [])[index]),
        toCellText_((snapshot.fontColors || [])[index]),
        toCellText_((snapshot.fontFamilies || [])[index]),
        Number((snapshot.fontSizes || [])[index]) || 0,
        toCellText_((snapshot.fontWeights || [])[index]),
        toCellText_((snapshot.numberFormats || [])[index]),
        toCellText_((snapshot.horizontal || [])[index]),
        toCellText_((snapshot.vertical || [])[index]),
        toCellText_((snapshot.wraps || [])[index])
      ]);
    });
    return [item.type || '', fields];
  });
  return sha256Hex_(JSON.stringify(normalizedRows));
}

function buildArchiveSourceSection_(sheet) {
  const month = sheet.getName();
  const columns = getRequiredMonthlyColumns_(sheet);
  const scan = scanMonthlyBlocks_(sheet, columns);
  if (scan.hybrids.length) {
    throw new Error(
      `${month} 有完整日期與 CalendarEventId 並存的混合列：` +
      scan.hybrids.map(item => item.row).join('、')
    );
  }
  if (scan.markerIssues.length) {
    throw new Error(
      `${month} 的刀日標記缺少或含衝突：` +
      scan.markerIssues.map(item => item.row).join('、')
    );
  }
  if (scan.orphans.length) {
    throw new Error(
      `${month} 有缺少上方日期標題的病人列：` +
      scan.orphans.map(item => item.row).join('、')
    );
  }
  scan.blocks.forEach(block => {
    if (formatMonthKey_(block.date) !== month) {
      throw new Error(
        `${month} 第 ${block.row} 列日期 ${block.dateKey} 不屬於該月份。`
      );
    }
  });
  const headers = getCanonicalMonthlyHeaderValues_(sheet, columns);
  const lastColumn = headers.length;
  const snapshots = captureMonthlyRows_(
    sheet,
    Math.max(1, sheet.getLastRow()),
    lastColumn
  );
  const byRow = {};
  snapshots.forEach(snapshot => {
    byRow[snapshot.sourceRow] = snapshot;
  });
  const orderedGroups = buildMonthlySnapshotGroups_(scan, byRow)
    .sort(compareMonthlySnapshotGroups_);
  const rows = [{ type: ARCHIVE_ROW_TYPES.MONTH, snapshot: null }];
  const eventIds = [];
  orderedGroups.forEach(group => {
    rows.push({ type: ARCHIVE_ROW_TYPES.DATE, snapshot: group.header });
    group.patients
      .sort((left, right) => compareMonthlyPatientRows_(left, right, columns))
      .forEach(snapshot => {
        rows.push({ type: ARCHIVE_ROW_TYPES.PATIENT, snapshot });
        const eventId = toCellText_(
          getRowFieldValue_(snapshot.values, columns, 'EVENT_ID')
        );
        if (eventId) eventIds.push(eventId);
      });
  });
  const formulaText = rows
    .filter(item => item.snapshot)
    .flatMap(item => item.snapshot.formulas || [])
    .filter(Boolean)
    .join('\n');
  return {
    month,
    year: month.slice(0, 4),
    sheetId: sheet.getSheetId(),
    headers,
    customHeaders: getMonthlyCustomHeaders_(headers),
    rows,
    eventIds: eventIds.sort(),
    patientCount: rows.filter(item => item.type === ARCHIVE_ROW_TYPES.PATIENT).length,
    blockCount: rows.filter(item => item.type === ARCHIVE_ROW_TYPES.DATE).length,
    sourceFingerprint: fingerprintArchiveRows_(headers, rows),
    formulaText
  };
}

function createArchiveSnapshotShell_(length) {
  const defaultWrap = typeof SpreadsheetApp !== 'undefined' &&
    SpreadsheetApp.WrapStrategy
    ? SpreadsheetApp.WrapStrategy.OVERFLOW
    : null;
  return {
    values: Array(length).fill(''),
    formulas: Array(length).fill(''),
    notes: Array(length).fill(''),
    backgrounds: Array(length).fill('#FFFFFF'),
    fontColors: Array(length).fill('#000000'),
    fontFamilies: Array(length).fill(MANAGED_FONT_FAMILY),
    fontSizes: Array(length).fill(MANAGED_FONT_SIZE),
    fontWeights: Array(length).fill('normal'),
    numberFormats: Array(length).fill('@'),
    horizontal: Array(length).fill('left'),
    vertical: Array(length).fill('top'),
    wraps: Array(length).fill(defaultWrap),
    validations: Array(length).fill(null)
  };
}

function remapArchiveSnapshot_(source, sourceHeaders, outputHeaders, month, type) {
  const target = createArchiveSnapshotShell_(outputHeaders.length);
  const sourceByHeader = {};
  (sourceHeaders || []).forEach((header, index) => {
    const normalized = toCellText_(header);
    if (normalized && sourceByHeader[normalized] === undefined) {
      sourceByHeader[normalized] = index;
    }
  });
  outputHeaders.forEach((header, targetIndex) => {
    const normalized = toCellText_(header);
    if (normalized === ANNUAL_ARCHIVE_SOURCE_MONTH_HEADER) {
      target.values[targetIndex] = month;
      return;
    }
    if (normalized === ANNUAL_ARCHIVE_ROW_TYPE_HEADER) {
      target.values[targetIndex] = type;
      return;
    }
    const sourceIndex = sourceByHeader[normalized];
    if (!normalized || sourceIndex === undefined || !source) return;
    [
      'values',
      'formulas',
      'notes',
      'backgrounds',
      'fontColors',
      'fontFamilies',
      'fontSizes',
      'fontWeights',
      'numberFormats',
      'horizontal',
      'vertical',
      'wraps',
      'validations'
    ].forEach(property => {
      if (source[property]) {
        target[property][targetIndex] = source[property][sourceIndex];
      }
    });
  });
  return target;
}

function getNewAnnualArchiveHeaders_(sections) {
  const eventHeader = CONFIG.MONTHLY_FIELD_HEADERS.EVENT_ID;
  const standardVisible = CONFIG.MONTHLY_HEADERS.filter(header => {
    return header !== eventHeader;
  });
  const custom = [];
  const seen = {};
  (sections || []).forEach(section => {
    section.customHeaders.forEach(header => {
      if (seen[header]) return;
      seen[header] = true;
      custom.push(header);
    });
  });
  return standardVisible.concat(custom, [
    eventHeader,
    ANNUAL_ARCHIVE_SOURCE_MONTH_HEADER,
    ANNUAL_ARCHIVE_ROW_TYPE_HEADER
  ]);
}

function makeInternalArchiveSheetName_(spreadsheet, prefix, year, operationId) {
  const token = toCellText_(operationId).replace(/[^A-Za-z0-9]/g, '').slice(0, 10);
  const base = `${prefix}${year}_${token || 'RUN'}`;
  if (!spreadsheet.getSheetByName(base)) return base;
  let suffix = 2;
  while (spreadsheet.getSheetByName(`${base}_${suffix}`)) suffix++;
  return `${base}_${suffix}`;
}

function initializeAnnualArchiveSheet_(sheet, year, headers, backupName) {
  ensureColumnCapacity_(sheet, headers.length);
  if (ANNUAL_ARCHIVE_HEADER_ROW > sheet.getMaxRows()) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      ANNUAL_ARCHIVE_HEADER_ROW - sheet.getMaxRows()
    );
  }
  sheet.getRange(1, 1)
    .setValue(getAnnualArchiveName_(year))
    .setNote(ANNUAL_ARCHIVE_NOTE_MARKER);
  sheet.getRange(2, 1).setValue(`封存年度：${year}`);
  sheet.getRange(3, 1).setValue(`更新時間：${new Date().toISOString()}`);
  sheet.getRange(4, 1).setValue(
    '歷史靜態封存：人工編輯不回寫月份刀表，也不同步 Calendar。'
  );
  sheet.getRange(5, 1).setValue(`完整備份：${backupName}`);
  sheet.getRange(ANNUAL_ARCHIVE_HEADER_ROW, 1, 1, headers.length)
    .setValues([headers]);
  sheet.getRange(1, 1, ANNUAL_ARCHIVE_HEADER_ROW, headers.length)
    .setFontFamily(MANAGED_FONT_FAMILY)
    .setFontSize(MANAGED_FONT_SIZE)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top')
    .setWrap(true);
  sheet.getRange(ANNUAL_ARCHIVE_HEADER_ROW, 1, 1, headers.length)
    .setBackground(TABLE_HEADER_BACKGROUND)
    .setFontColor('#000000')
    .setFontWeight('bold');
  sheet.setFrozenRows(ANNUAL_ARCHIVE_HEADER_ROW);
}

function ensureAnnualArchiveCustomHeaders_(sheet, sections) {
  let info = getAnnualArchiveHeaderInfo_(sheet);
  const desired = [];
  const seen = {};
  sections.forEach(section => {
    section.customHeaders.forEach(header => {
      if (!info.columns[header] && !seen[header]) {
        seen[header] = true;
        desired.push(header);
      }
    });
  });
  if (!desired.length) return info;
  sheet.insertColumnsBefore(info.eventColumn, desired.length);
  sheet.getRange(
    ANNUAL_ARCHIVE_HEADER_ROW,
    info.eventColumn,
    1,
    desired.length
  ).setValues([desired]);
  if (typeof sheet.showColumns === 'function') {
    sheet.showColumns(info.eventColumn, desired.length);
  }
  desired.forEach((header, index) => {
    sheet.setColumnWidth(info.eventColumn + index, 120);
  });
  info = getAnnualArchiveHeaderInfo_(sheet);
  return info;
}

function removeAnnualArchiveProtections_(sheet) {
  sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
    .filter(protection => {
      return toCellText_(protection.getDescription())
        .indexOf('SURGERY_ANNUAL_ARCHIVE:') === 0;
    })
    .forEach(protection => protection.remove());
}

function protectAnnualArchiveSheet_(sheet) {
  removeAnnualArchiveProtections_(sheet);
  const info = getAnnualArchiveHeaderInfo_(sheet);
  const systemColumns = [
    info.eventColumn,
    info.sourceMonthColumn,
    info.rowTypeColumn
  ].sort((left, right) => left - right);
  systemColumns.forEach(column => {
    sheet.hideColumns(column);
    const protection = sheet.getRange(1, column, sheet.getMaxRows(), 1)
      .protect()
      .setDescription(`SURGERY_ANNUAL_ARCHIVE:SYSTEM:${column}`);
    if (typeof protection.setWarningOnly === 'function') {
      protection.setWarningOnly(false);
    }
  });
  buildVisibleRangesExcludingColumns_(
    sheet,
    1,
    sheet.getMaxRows(),
    info.lastColumn,
    systemColumns
  ).forEach((range, index) => {
    const protection = range.protect()
      .setDescription(`SURGERY_ANNUAL_ARCHIVE:CLINICAL:${index}`);
    if (typeof protection.setWarningOnly === 'function') {
      protection.setWarningOnly(true);
    }
  });
}

function applyAnnualArchiveConditionalFormatRules_(
  sheet,
  headerRow,
  columns,
  lastColumn
) {
  const startRow = headerRow + 1;
  const rowCount = Math.max(1, sheet.getMaxRows() - headerRow);
  const planRanges = buildVisibleRangesExcludingColumns_(
    sheet,
    startRow,
    sheet.getMaxRows(),
    lastColumn,
    [columns.GA, columns.SIDE]
  );
  const rules = [];
  const marker = name => `${CONDITIONAL_FORMAT_MARKER_PREFIX}ARCHIVE_${name}`;
  if (columns.PLAN) {
    const planLetter = columnToLetter_(columns.PLAN);
    const planReference = `$${planLetter}${startRow}`;
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
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND(${redMatch},N("${marker('PLAN_RED')}")=0)`
        )
        .setBackground(PLAN_RED_BACKGROUND)
        .setRanges(planRanges)
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND(NOT(${redMatch}),${greenMatch},` +
          `N("${marker('PLAN_GREEN')}")=0)`
        )
        .setBackground(PLAN_GREEN_BACKGROUND)
        .setRanges(planRanges)
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND(NOT(${redMatch}),NOT(${greenMatch}),${yellowMatch},` +
          `N("${marker('PLAN_YELLOW')}")=0)`
        )
        .setBackground(PLAN_YELLOW_BACKGROUND)
        .setRanges(planRanges)
        .build()
    );
  }
  if (columns.GA) {
    const letter = columnToLetter_(columns.GA);
    rules.unshift(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND($${letter}${startRow}="GA",N("${marker('GA')}")=0)`
        )
        .setBackground(MONTHLY_GA_BACKGROUND)
        .setRanges([sheet.getRange(startRow, columns.GA, rowCount, 1)])
        .build()
    );
  }
  if (columns.SIDE) {
    const letter = columnToLetter_(columns.SIDE);
    rules.unshift(
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND($${letter}${startRow}="OU",N("${marker('OU')}")=0)`
        )
        .setBackground(SIDE_OU_BACKGROUND)
        .setRanges([sheet.getRange(startRow, columns.SIDE, rowCount, 1)])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(
          `=AND($${letter}${startRow}="OD",N("${marker('OD')}")=0)`
        )
        .setBackground(SIDE_OD_BACKGROUND)
        .setRanges([sheet.getRange(startRow, columns.SIDE, rowCount, 1)])
        .build()
    );
  }
  sheet.setConditionalFormatRules(
    rules.concat(removeSystemConditionalFormats_(sheet.getConditionalFormatRules()))
  );
}

function applyAnnualArchiveConditionalFormats_(sheet) {
  const info = getAnnualArchiveHeaderInfo_(sheet);
  const columns = {};
  Object.keys(CONFIG.MONTHLY_FIELD_HEADERS).forEach(key => {
    columns[key] = info.columns[CONFIG.MONTHLY_FIELD_HEADERS[key]] || 0;
  });
  applyAnnualArchiveConditionalFormatRules_(
    sheet,
    ANNUAL_ARCHIVE_HEADER_ROW,
    columns,
    info.lastColumn
  );
}

function applyAnnualArchiveDisplay_(sheet, isNew) {
  const info = getAnnualArchiveHeaderInfo_(sheet);
  sheet.getRange(1, 1, ANNUAL_ARCHIVE_HEADER_ROW, info.lastColumn)
    .setFontFamily(MANAGED_FONT_FAMILY)
    .setFontSize(MANAGED_FONT_SIZE)
    .setHorizontalAlignment('left')
    .setVerticalAlignment('top')
    .setWrap(true);
  sheet.getRange(ANNUAL_ARCHIVE_HEADER_ROW, 1, 1, info.lastColumn)
    .setBackground(TABLE_HEADER_BACKGROUND)
    .setFontColor('#000000')
    .setFontWeight('bold');
  if (isNew) {
    Object.keys(MONTHLY_COLUMN_WIDTHS).forEach(key => {
      const header = CONFIG.MONTHLY_FIELD_HEADERS[key];
      const column = info.columns[header];
      if (column) sheet.setColumnWidth(column, MONTHLY_COLUMN_WIDTHS[key]);
    });
  }
  if (isNew) {
    MONTHLY_WRAP_KEYS.forEach(key => {
      const column = info.columns[CONFIG.MONTHLY_FIELD_HEADERS[key]];
      if (column) sheet.getRange(1, column, sheet.getMaxRows(), 1).setWrap(true);
    });
  }
  applyAnnualArchiveConditionalFormats_(sheet);
  protectAnnualArchiveSheet_(sheet);
  sheet.setFrozenRows(ANNUAL_ARCHIVE_HEADER_ROW);
}

function getArchiveInsertRow_(structure, month) {
  const later = structure.monthRanges.find(item => item.month > month);
  return later ? later.startRow : Math.max(
    ANNUAL_ARCHIVE_HEADER_ROW + 1,
    structure.lastRow + 1
  );
}

function writeArchiveSnapshotsBatch_(sheet, startRow, items, lastColumn) {
  if (!items.length) return;
  const snapshots = items.map(item => item.snapshot);
  const range = sheet.getRange(startRow, 1, snapshots.length, lastColumn);
  range.setValues(snapshots.map(item => item.values.slice(0, lastColumn)));
  range.setNotes(snapshots.map(item => item.notes.slice(0, lastColumn)));
  range.setBackgrounds(
    snapshots.map(item => item.backgrounds.slice(0, lastColumn))
  );
  range.setFontColors(
    snapshots.map(item => item.fontColors.slice(0, lastColumn))
  );
  range.setFontFamilies(
    snapshots.map(item => item.fontFamilies.slice(0, lastColumn))
  );
  range.setFontSizes(
    snapshots.map(item => item.fontSizes.slice(0, lastColumn))
  );
  range.setFontWeights(
    snapshots.map(item => item.fontWeights.slice(0, lastColumn))
  );
  range.setNumberFormats(
    snapshots.map(item => item.numberFormats.slice(0, lastColumn))
  );
  range.setHorizontalAlignments(
    snapshots.map(item => item.horizontal.slice(0, lastColumn))
  );
  range.setVerticalAlignments(
    snapshots.map(item => item.vertical.slice(0, lastColumn))
  );
  range.setWrapStrategies(
    snapshots.map(item => item.wraps.slice(0, lastColumn))
  );
  range.setDataValidations(
    snapshots.map(item => item.validations.slice(0, lastColumn))
  );
  snapshots.forEach((snapshot, rowOffset) => {
    snapshot.formulas.forEach((formula, columnIndex) => {
      if (formula) {
        sheet.getRange(startRow + rowOffset, columnIndex + 1)
          .setFormulaR1C1(formula);
      }
    });
  });
}

function insertArchiveSection_(sheet, section) {
  const structure = scanAnnualArchiveStructure_(sheet);
  if (structure.months.indexOf(section.month) !== -1) {
    throw new Error(`${section.month} 已存在於 ${sheet.getName()}。`);
  }
  const info = structure.info;
  const outputHeaders = info.headers;
  const mappedRows = section.rows.map(item => {
    if (item.type === ARCHIVE_ROW_TYPES.MONTH) {
      const snapshot = createArchiveSnapshotShell_(outputHeaders.length);
      const timeColumn = info.columns[CONFIG.MONTHLY_FIELD_HEADERS.TIME];
      snapshot.values[timeColumn - 1] = section.month;
      snapshot.values[info.sourceMonthColumn - 1] = section.month;
      snapshot.values[info.rowTypeColumn - 1] = item.type;
      snapshot.backgrounds.fill(TABLE_HEADER_BACKGROUND);
      snapshot.fontWeights.fill('bold');
      return { type: item.type, snapshot };
    }
    return {
      type: item.type,
      snapshot: remapArchiveSnapshot_(
        item.snapshot,
        section.headers,
        outputHeaders,
        section.month,
        item.type
      )
    };
  });
  const insertRow = getArchiveInsertRow_(structure, section.month);
  const rowCount = mappedRows.length;
  if (insertRow <= sheet.getLastRow()) {
    sheet.insertRowsBefore(insertRow, rowCount);
  } else if (insertRow + rowCount - 1 > sheet.getMaxRows()) {
    sheet.insertRowsAfter(
      sheet.getMaxRows(),
      insertRow + rowCount - 1 - sheet.getMaxRows()
    );
  }
  writeArchiveSnapshotsBatch_(
    sheet,
    insertRow,
    mappedRows,
    outputHeaders.length
  );
  const expectedFingerprint = fingerprintArchiveRows_(outputHeaders, mappedRows);
  const actualSnapshots = captureArchiveRangeSnapshots_(
    sheet,
    insertRow,
    rowCount,
    outputHeaders.length
  );
  const actualRows = actualSnapshots.map(snapshot => {
    const type = toCellText_(
      snapshot.values[info.rowTypeColumn - 1]
    );
    return { type, snapshot };
  });
  const actualFingerprint = fingerprintArchiveRows_(outputHeaders, actualRows);
  if (expectedFingerprint !== actualFingerprint) {
    throw new Error(`${section.month} 寫入年度封存後回讀驗證失敗。`);
  }
  return {
    month: section.month,
    rowCount,
    patientCount: section.patientCount,
    blockCount: section.blockCount,
    fingerprint: actualFingerprint
  };
}

function captureArchiveRangeSnapshots_(sheet, startRow, rowCount, lastColumn) {
  if (!rowCount) return [];
  const range = sheet.getRange(startRow, 1, rowCount, lastColumn);
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
    sourceRow: startRow + index,
    values: rowValues,
    formulas: formulas[index],
    notes: notes[index],
    backgrounds: backgrounds[index],
    fontColors: fontColors[index],
    fontFamilies: fontFamilies[index],
    fontSizes: fontSizes[index],
    fontWeights: fontWeights[index],
    numberFormats: numberFormats[index],
    horizontal: horizontal[index],
    vertical: vertical[index],
    wraps: wraps[index],
    validations: validations[index]
  }));
}

function fingerprintEventIdSet_(eventIds) {
  return sha256Hex_(JSON.stringify(
    Array.from(new Set((eventIds || []).map(toCellText_).filter(Boolean))).sort()
  ));
}

function assertArchiveSyncHealthy_(spreadsheet) {
  const baseline = readCalendarRegistryStore_();
  if (!baseline.ok) {
    throw new Error('同步索引尚未建立或已損壞；請先執行「安裝／修復系統」。');
  }
  const health = reconcileCalendarRegistry_(spreadsheet, {
    apply: false,
    changeType: 'ARCHIVE_PREFLIGHT'
  });
  const pending = health.analysis.actions.length;
  const conflicts = health.analysis.conflicts.length;
  if (pending || conflicts || health.scan.duplicateIds.length) {
    throw new Error(
      `封存前同步檢查未通過：待處理 ${pending}、衝突 ${conflicts}。` +
      '請先完成同步健康處理。'
    );
  }
  return health;
}

function getArchiveSourceSections_(spreadsheet, monthNames) {
  const currentMonth = formatMonthKey_(new Date());
  const archived = listArchivedMonthNames_(spreadsheet);
  const sections = monthNames.map(month => {
    if (!isClosedMonthForArchive_(month, currentMonth)) {
      throw new Error(`${month} 不是已結束月份，不能永久封存。`);
    }
    if (archived.indexOf(month) !== -1) {
      throw new Error(`${month} 已經永久封存。`);
    }
    const sheet = spreadsheet.getSheetByName(month);
    if (!sheet || !isMonthlySheetName_(sheet.getName())) {
      throw new Error(`找不到月份刀表 ${month}。`);
    }
    return buildArchiveSourceSection_(sheet);
  });
  sections.forEach(section => {
    monthNames.forEach(deletedMonth => {
      if (
        section.formulaText.indexOf(`${deletedMonth}!`) !== -1 ||
        section.formulaText.indexOf(`'${deletedMonth}'!`) !== -1
      ) {
        throw new Error(
          `${section.month} 含直接引用即將刪除分頁 ${deletedMonth} 的公式；` +
          '請先轉為值。'
        );
      }
    });
  });
  return sections;
}

function verifyPrivateBackup_(file) {
  const isPrivate = file.getSharingAccess() === DriveApp.Access.PRIVATE;
  const editors = file.getEditors();
  const viewers = file.getViewers();
  if (!isPrivate || editors.length || viewers.length) {
    throw new Error(
      '整檔備份無法確認為私人檔案；來源月份未刪除。'
    );
  }
  return true;
}

function createPrivateSpreadsheetBackup_(spreadsheet, sections) {
  SpreadsheetApp.flush();
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyyMMdd_HHmmss'
  );
  const name = `手術排程系統_Backup_${timestamp}`;
  const copy = DriveApp.getFileById(spreadsheet.getId()).makeCopy(name);
  try {
    if (typeof copy.setShareableByEditors === 'function') {
      copy.setShareableByEditors(false);
    }
    copy.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
    copy.getEditors().forEach(user => copy.removeEditor(user));
    copy.getViewers().forEach(user => copy.removeViewer(user));
    verifyPrivateBackup_(copy);
    const backupSpreadsheet = SpreadsheetApp.openById(copy.getId());
    const sourceNames = spreadsheet.getSheets().map(sheet => sheet.getName()).sort();
    const backupNames = backupSpreadsheet.getSheets()
      .map(sheet => sheet.getName())
      .sort();
    if (JSON.stringify(sourceNames) !== JSON.stringify(backupNames)) {
      throw new Error('整檔備份分頁清單回讀不一致。');
    }
    sections.forEach(section => {
      const backupSheet = backupSpreadsheet.getSheetByName(section.month);
      if (!backupSheet) {
        throw new Error(`整檔備份缺少 ${section.month}。`);
      }
      const backupSection = buildArchiveSourceSection_(backupSheet);
      if (
        backupSection.sourceFingerprint !== section.sourceFingerprint ||
        fingerprintEventIdSet_(backupSection.eventIds) !==
          fingerprintEventIdSet_(section.eventIds)
      ) {
        throw new Error(`${section.month} 整檔備份回讀驗證失敗。`);
      }
    });
  } catch (err) {
    throw new Error(
      `私人整檔備份驗證失敗；來源月份未刪除。備份檔保留供檢查：` +
      `${err.message || err}`
    );
  }
  return { fileId: copy.getId(), name };
}

function groupArchiveSectionsByYear_(sections) {
  const result = {};
  sections.forEach(section => {
    if (!result[section.year]) result[section.year] = [];
    result[section.year].push(section);
  });
  Object.keys(result).forEach(year => {
    result[year].sort((left, right) => left.month.localeCompare(right.month));
  });
  return result;
}

function prepareAnnualArchiveStages_(spreadsheet, transaction, sections) {
  const grouped = groupArchiveSectionsByYear_(sections);
  const stages = [];
  Object.keys(grouped).sort().forEach(year => {
    const finalName = getAnnualArchiveName_(year);
    const existing = spreadsheet.getSheetByName(finalName);
    let stage;
    let isNew = false;
    if (existing) {
      scanAnnualArchiveStructure_(existing);
      stage = existing.copyTo(spreadsheet);
    } else {
      stage = spreadsheet.insertSheet();
      isNew = true;
    }
    stage.setName(makeInternalArchiveSheetName_(
      spreadsheet,
      ARCHIVE_STAGE_PREFIX,
      year,
      transaction.operationId
    ));
    stage.hideSheet();
    if (isNew) {
      initializeAnnualArchiveSheet_(
        stage,
        year,
        getNewAnnualArchiveHeaders_(grouped[year]),
        transaction.backupName
      );
    } else {
      stage.getRange(5, 1).setValue(`完整備份：${transaction.backupName}`);
      stage.getRange(3, 1).setValue(`更新時間：${new Date().toISOString()}`);
    }
    ensureAnnualArchiveCustomHeaders_(stage, grouped[year]);
    const inserted = grouped[year].map(section => {
      return insertArchiveSection_(stage, section);
    });
    applyAnnualArchiveDisplay_(stage, isNew);
    const verified = scanAnnualArchiveStructure_(stage);
    grouped[year].forEach(section => {
      if (verified.months.indexOf(section.month) === -1) {
        throw new Error(`${section.month} 未出現在年度封存暫存表。`);
      }
    });
    stages.push({
      year,
      finalName,
      stageSheetId: stage.getSheetId(),
      existingSheetId: existing ? existing.getSheetId() : 0,
      oldSheetId: 0,
      months: grouped[year].map(section => section.month),
      inserted,
      committed: false
    });
  });
  return stages;
}

function updateArchiveTransaction_(transaction, changes) {
  Object.keys(changes || {}).forEach(key => {
    transaction[key] = changes[key];
  });
  transaction.updatedAt = new Date().toISOString();
  writeAnnualArchiveTransaction_(transaction);
  return transaction;
}

function commitAnnualArchiveStages_(spreadsheet, transaction) {
  transaction.state = 'COMMITTING_ARCHIVES';
  writeAnnualArchiveTransaction_(transaction);
  transaction.stages.forEach(stageInfo => {
    const stage = getSheetById_(spreadsheet, stageInfo.stageSheetId);
    if (!stage) {
      throw new Error(`${stageInfo.year} 封存暫存表遺失。`);
    }
    const finalByName = spreadsheet.getSheetByName(stageInfo.finalName);
    if (
      finalByName &&
      Number(finalByName.getSheetId()) === Number(stageInfo.stageSheetId)
    ) {
      stageInfo.committed = true;
      return;
    }
    if (finalByName) {
      const oldName = makeInternalArchiveSheetName_(
        spreadsheet,
        ARCHIVE_OLD_PREFIX,
        stageInfo.year,
        transaction.operationId
      );
      stageInfo.oldSheetId = finalByName.getSheetId();
      writeAnnualArchiveTransaction_(transaction);
      finalByName.setName(oldName);
      finalByName.hideSheet();
      writeAnnualArchiveTransaction_(transaction);
    }
    stage.setName(stageInfo.finalName);
    stage.showSheet();
    stage.getRange(1, 1)
      .setValue(stageInfo.finalName)
      .setNote(ANNUAL_ARCHIVE_NOTE_MARKER);
    stageInfo.committed = true;
    writeAnnualArchiveTransaction_(transaction);
  });
  updateArchiveTransaction_(transaction, { state: 'ARCHIVES_COMMITTED' });
}

function deleteArchivedSourceSheets_(spreadsheet, transaction) {
  transaction.state = 'DELETING_SOURCES';
  writeAnnualArchiveTransaction_(transaction);
  transaction.months.forEach(month => {
    if (transaction.deletedMonths.indexOf(month) !== -1) return;
    const sourceId = transaction.sourceSheetIds[month];
    const sheet = getSheetById_(spreadsheet, sourceId);
    if (sheet) {
      if (sheet.getName() !== month || !isMonthlySheetName_(sheet.getName())) {
        throw new Error(`${month} 來源 sheet ID 已指向其他分頁，停止刪除。`);
      }
      spreadsheet.deleteSheet(sheet);
    }
    transaction.deletedMonths.push(month);
    writeAnnualArchiveTransaction_(transaction);
  });
  transaction.stages.forEach(stageInfo => {
    const oldSheet = getSheetById_(spreadsheet, stageInfo.oldSheetId);
    if (oldSheet) spreadsheet.deleteSheet(oldSheet);
  });
  updateArchiveTransaction_(transaction, { state: 'SOURCES_DELETED' });
}

function rollbackArchiveBeforeSourceDeletion_(spreadsheet, transaction) {
  if ((transaction.deletedMonths || []).length) return false;
  (transaction.stages || []).forEach(stageInfo => {
    const stage = getSheetById_(spreadsheet, stageInfo.stageSheetId);
    const oldSheet = getSheetById_(spreadsheet, stageInfo.oldSheetId);
    if (stage) spreadsheet.deleteSheet(stage);
    if (oldSheet) {
      oldSheet.setName(stageInfo.finalName);
      oldSheet.showSheet();
    }
  });
  writeAnnualArchiveTransaction_(null);
  return true;
}

function finalizeAnnualArchiveTransaction_(spreadsheet, transaction) {
  const archivedMonths = listArchivedMonthNames_(spreadsheet);
  transaction.months.forEach(month => {
    if (archivedMonths.indexOf(month) === -1) {
      throw new Error(`${month} 永久封存結果驗證失敗。`);
    }
    if (spreadsheet.getSheetByName(month)) {
      throw new Error(`${month} 來源分頁仍存在。`);
    }
  });
  const scan = buildCurrentCalendarScan_(spreadsheet);
  if (scan.duplicateIds.length) {
    throw new Error(
      `封存後 CalendarEventId 重複：${scan.duplicateIds.join('、')}。`
    );
  }
  const currentFingerprint = fingerprintEventIdSet_(
    Object.keys(scan.eventLocations)
  );
  if (currentFingerprint !== transaction.allEventIdFingerprint) {
    throw new Error('封存前後 CalendarEventId 集合不一致。');
  }
  const registry = rebuildCalendarRegistry_(spreadsheet, {
    allowConflicts: false
  });
  reorderManagedSheets_(spreadsheet);
  writeAnnualArchiveTransaction_(null);
  return {
    ok: true,
    months: transaction.months,
    backupName: transaction.backupName,
    backupFileId: transaction.backupFileId,
    eventCount: registry.eventCount,
    patientCount: transaction.patientCount,
    blockCount: transaction.blockCount,
    message:
      `已永久封存 ${transaction.months.join('、')}，` +
      `共 ${transaction.patientCount} 筆病人、${transaction.blockCount} 個日期區塊；` +
      `原月份分頁已刪除。私人整檔備份：${transaction.backupName}。`
  };
}

function loadTransactionSourceSections_(spreadsheet, transaction) {
  return transaction.months.map(month => {
    const sheet = getSheetById_(spreadsheet, transaction.sourceSheetIds[month]);
    if (!sheet || sheet.getName() !== month) {
      throw new Error(`${month} 來源分頁遺失，無法重建封存暫存表。`);
    }
    const section = buildArchiveSourceSection_(sheet);
    if (section.sourceFingerprint !== transaction.sourceFingerprints[month]) {
      throw new Error(`${month} 在封存中斷後被修改，已停止續跑。`);
    }
    return section;
  });
}

function verifyTransactionBackup_(transaction) {
  if (!transaction.backupFileId) {
    throw new Error('永久封存交易缺少已驗證的私人整檔備份。');
  }
  let file;
  try {
    file = DriveApp.getFileById(transaction.backupFileId);
  } catch (err) {
    throw new Error('永久封存的私人整檔備份已無法存取。');
  }
  verifyPrivateBackup_(file);
  return file;
}

function continueAnnualArchiveTransaction_(spreadsheet, transaction) {
  try {
    let sourceSections = null;
    if (!transaction.backupFileId) {
      sourceSections = loadTransactionSourceSections_(spreadsheet, transaction);
      const backup = createPrivateSpreadsheetBackup_(
        spreadsheet,
        sourceSections
      );
      updateArchiveTransaction_(transaction, {
        backupFileId: backup.fileId,
        backupName: backup.name,
        state: 'BACKUP_VERIFIED'
      });
    }
    verifyTransactionBackup_(transaction);
    if (!transaction.stages || !transaction.stages.length) {
      const sections = sourceSections ||
        loadTransactionSourceSections_(spreadsheet, transaction);
      transaction.stages = prepareAnnualArchiveStages_(
        spreadsheet,
        transaction,
        sections
      );
      updateArchiveTransaction_(transaction, { state: 'STAGES_READY' });
    }
    if (transaction.state === 'STAGES_READY' ||
        transaction.state === 'COMMITTING_ARCHIVES') {
      commitAnnualArchiveStages_(spreadsheet, transaction);
    }
    if (
      transaction.state === 'ARCHIVES_COMMITTED' ||
      transaction.state === 'DELETING_SOURCES'
    ) {
      deleteArchivedSourceSheets_(spreadsheet, transaction);
    }
    if (transaction.state === 'SOURCES_DELETED') {
      return finalizeAnnualArchiveTransaction_(spreadsheet, transaction);
    }
    throw new Error(`未知封存交易階段：${transaction.state}`);
  } catch (err) {
    const rolledBack = rollbackArchiveBeforeSourceDeletion_(
      spreadsheet,
      transaction
    );
    const suffix = rolledBack
      ? '來源月份未刪除；暫存變更已回復，私人備份仍保留。'
      : '交易標記已保留，自動同步維持暫停；請再次執行「彙整舊月刀表」續跑。';
    throw new Error(`${err.message || err}\n${suffix}`);
  }
}

function createAnnualArchiveTransaction_(spreadsheet, request) {
  assertNoLegacyArchiveState_();
  const active = readAnnualArchiveTransaction_();
  if (active) return continueAnnualArchiveTransaction_(spreadsheet, active);
  if (request.confirmedPermanentDeletion !== true) {
    throw new Error('尚未確認永久刪除來源月份分頁。');
  }
  const monthNames = Array.from(new Set(
    (request.months || []).map(toCellText_).filter(Boolean)
  )).sort();
  if (!monthNames.length) throw new Error('請至少選擇一個已結束月份。');
  const health = assertArchiveSyncHealthy_(spreadsheet);
  const sections = getArchiveSourceSections_(spreadsheet, monthNames);
  const sourceSheetIds = {};
  const sourceFingerprints = {};
  sections.forEach(section => {
    sourceSheetIds[section.month] = section.sheetId;
    sourceFingerprints[section.month] = section.sourceFingerprint;
  });
  const transaction = {
    version: 1,
    operationId: toCellText_(request.operationId) || Utilities.getUuid(),
    state: 'PREFLIGHTED',
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    months: monthNames,
    sourceSheetIds,
    sourceFingerprints,
    deletedMonths: [],
    stages: [],
    backupFileId: '',
    backupName: '',
    patientCount: sections.reduce((sum, item) => sum + item.patientCount, 0),
    blockCount: sections.reduce((sum, item) => sum + item.blockCount, 0),
    allEventIdFingerprint: fingerprintEventIdSet_(
      Object.keys(health.scan.eventLocations)
    )
  };
  writeAnnualArchiveTransaction_(transaction);
  try {
    const backup = createPrivateSpreadsheetBackup_(spreadsheet, sections);
    updateArchiveTransaction_(transaction, {
      backupFileId: backup.fileId,
      backupName: backup.name,
      state: 'BACKUP_VERIFIED'
    });
    transaction.stages = prepareAnnualArchiveStages_(
      spreadsheet,
      transaction,
      sections
    );
    updateArchiveTransaction_(transaction, { state: 'STAGES_READY' });
    return continueAnnualArchiveTransaction_(spreadsheet, transaction);
  } catch (err) {
    const rolledBack = rollbackArchiveBeforeSourceDeletion_(
      spreadsheet,
      transaction
    );
    if (!rolledBack) {
      throw err;
    }
    throw new Error(`${err.message || err}\n來源月份未刪除。`);
  }
}

function getRollupDialogData_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  assertNoLegacyArchiveState_();
  const currentMonth = formatMonthKey_(new Date());
  const archived = listArchivedMonthNames_(spreadsheet);
  const months = getActiveMonthlySheets_(spreadsheet)
    .filter(sheet => (
      isClosedMonthForArchive_(sheet.getName(), currentMonth) &&
      archived.indexOf(sheet.getName()) === -1
    ))
    .sort((left, right) => left.getName().localeCompare(right.getName()))
    .map(sheet => ({
      name: sheet.getName(),
      year: sheet.getName().slice(0, 4),
      patientCount: scanMonthlyBlocks_(
        sheet,
        getRequiredMonthlyColumns_(sheet)
      ).blocks.reduce((sum, block) => sum + block.patientRows.length, 0)
    }));
  return {
    operationId: Utilities.getUuid(),
    months,
    defaultMonths: months.length ? [months[0].name] : [],
    currentMonth
  };
}

function showMonthlyRollupDialog() {
  return runMenuAction_('彙整舊月刀表', () => {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const pending = readAnnualArchiveTransaction_();
    if (pending) {
      const response = SpreadsheetApp.getUi().alert(
        '繼續永久封存',
        `偵測到尚未完成的封存交易：${pending.months.join('、')}。\n` +
          '按「確定」會從安全檢查點繼續。',
        SpreadsheetApp.getUi().ButtonSet.OK_CANCEL
      );
      if (response !== SpreadsheetApp.getUi().Button.OK) {
        return { ok: false, cancelled: true };
      }
      const result = withCalendarSyncLock_(() => {
        return continueAnnualArchiveTransaction_(spreadsheet, pending);
      });
      SpreadsheetApp.getUi().alert(result.message);
      return result;
    }
    const data = getRollupDialogData_();
    if (!data.months.length) {
      throw new Error('目前沒有可永久封存的已結束月份。');
    }
    const template = HtmlService.createTemplateFromFile('monthly_rollup');
    template.initialDataJson = JSON.stringify(data).replace(/<\//g, '<\\/');
    SpreadsheetApp.getUi().showModalDialog(
      template.evaluate().setWidth(560).setHeight(620),
      '彙整舊月刀表'
    );
    return { ok: true };
  });
}

function submitMonthlyRollupRequest(payload) {
  return withCalendarSyncLock_(() => {
    const request = payload || {};
    const token = toCellText_(request.operationId);
    if (!token) throw new Error('永久封存 request token 遺失，請重新開啟視窗。');
    const cache = CacheService.getUserCache();
    const key = `ANNUAL_ARCHIVE_${token}`;
    const cached = safeJsonParse_(cache.get(key), null);
    if (cached && cached.ok) return cached;
    const result = createAnnualArchiveTransaction_(
      SpreadsheetApp.getActiveSpreadsheet(),
      request
    );
    cache.put(key, JSON.stringify(result), CROSS_SHEET_CACHE_SECONDS);
    return result;
  });
}
