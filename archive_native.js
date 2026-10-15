/** Native Sheets v4 archive contract shared with the Windows planners. */
const NATIVE_ARCHIVE_NOTE = 'VGHKS_ARCHIVE_CONTRACT_V2';
const NATIVE_ARCHIVE_LOCK_KEY = 'vghks.sheet.archive.lock';
const NATIVE_ARCHIVE_LOCK_ID = 2147483646;
const NATIVE_ARCHIVE_WITNESS = 'vghks.sheet.management';
const NATIVE_CELL_FIELDS = 'userEnteredValue,userEnteredFormat,dataValidation,note,textFormatRuns';

function nativeClone_(value) { return JSON.parse(JSON.stringify(value)); }
function nativeText_(cell) {
  const value = (cell || {}).userEnteredValue || {};
  if ('stringValue' in value) return String(value.stringValue);
  if ('numberValue' in value) return String(value.numberValue);
  if ('boolValue' in value) return String(value.boolValue);
  return String((cell || {}).formattedValue || '');
}
function nativeNorm_(value) { return String(value || '').normalize('NFKC').replace(/\s+/g, ' ').trim(); }
function nativeLiteral_(value) { return { userEnteredValue: { stringValue: String(value) } }; }
function nativeCell_(cell, freeze) {
  const result = {};
  NATIVE_CELL_FIELDS.split(',').forEach(key => {
    if ((cell || {})[key] !== undefined) result[key] = nativeClone_(cell[key]);
  });
  if (freeze && cell.effectiveFormat) {
    const effective = cell.effectiveFormat, format = result.userEnteredFormat || {};
    ['backgroundColor', 'backgroundColorStyle'].forEach(key => {
      if (effective[key]) format[key] = nativeClone_(effective[key]);
    });
    if (format.backgroundColorStyle) delete format.backgroundColor;
    const text = format.textFormat || {};
    ['bold', 'italic', 'strikethrough', 'foregroundColor', 'foregroundColorStyle'].forEach(key => {
      if ((effective.textFormat || {})[key] !== undefined) text[key] = nativeClone_(effective.textFormat[key]);
    });
    if (text.foregroundColorStyle) delete text.foregroundColor;
    if (Object.keys(text).length) format.textFormat = text;
    result.userEnteredFormat = format;
  }
  return result;
}
function nativeHeader_(cell) {
  const text = nativeNorm_(nativeText_(cell));
  const formula = ((cell || {}).userEnteredValue || {}).formulaValue || '';
  if (formula.includes(MONTHLY_DIAGNOSIS_SUMMARY_FORMULA_MARKER) || text === '診斷' || text.startsWith('診斷|')) return '診斷';
  if (['時間', '日期/報到時間'].includes(text)) return CONFIG.MONTHLY_FIELD_HEADERS.TIME;
  if (['日曆eventID', '日曆apiEventID'].includes(text)) return 'CalendarEventId';
  return text;
}
function nativeColumns_(sheet, header) {
  const cols = {};
  ((sheet.rows || [])[header] || []).forEach((cell, index) => {
    const name = nativeHeader_(cell);
    if (name) {
      if (cols[name] !== undefined) throw new Error(`${sheet.properties.title} 有重複欄名 ${name}。`);
      cols[name] = index;
    }
  });
  return cols;
}
function nativeLabels_(sheet, header) {
  const cols = nativeColumns_(sheet, header), rows = sheet.rows || [];
  let width = 0;
  rows.forEach(row => row.forEach((cell, index) => {
    if (Object.keys(cell.userEnteredValue || {}).length || cell.note || (cell.textFormatRuns || []).length) width = Math.max(width, index + 1);
  }));
  const labels = Array.from({ length: width }, (_, index) => nativeHeader_((rows[header] || [])[index] || {}));
  const missing = CONFIG.MONTHLY_HEADERS.filter(Boolean).filter(name => cols[name] === undefined);
  if (missing.length) throw new Error(`${sheet.properties.title} 缺少欄位：${missing.join('、')}。`);
  const blanks = labels.map((name, index) => name ? -1 : index).filter(index => index >= 0);
  if (blanks.length > 1 || blanks.some(index => rows.slice(header + 1).some(row => {
    const cell = row[index] || {}; return Object.keys(cell.userEnteredValue || {}).length || cell.note || (cell.textFormatRuns || []).length;
  }))) throw new Error('未命名人工欄無法安全重映射，請先命名。');
  return { labels, cols };
}
function nativeValue_(row, cols, name) { return nativeText_(row[cols[name]] || {}).trim(); }
function nativeDay_(cell) {
  const value = (cell || {}).effectiveValue || (cell || {}).userEnteredValue || {};
  if (Number.isInteger(value.numberValue) && value.numberValue >= 36526 && value.numberValue <= 73050) {
    return new Date((value.numberValue - 25569) * 86400000).toISOString().slice(0, 10);
  }
  let text = nativeText_(cell).trim();
  if (/^\d{8}$/.test(text)) text = `${text.slice(0, 4)}-${text.slice(4, 6)}-${text.slice(6)}`;
  const parsed = parseFullDate_(text);
  return parsed ? formatDateKey_(parsed) : '';
}
function nativeKind_(row, cols) {
  const cell = row[cols[CONFIG.MONTHLY_FIELD_HEADERS.TIME]] || {};
  const date = nativeDay_(cell), mrn = nativeNorm_(nativeValue_(row, cols, '病歷號'));
  const hospital = nativeNorm_(nativeValue_(row, cols, '醫院'));
  const raw = nativeText_(cell), value = cell.effectiveValue || cell.userEnteredValue || {};
  const numeric = value.numberValue, dateFormat = ['DATE', 'DATE_TIME'].includes(((cell.userEnteredFormat || {}).numberFormat || {}).type);
  const dateHint = typeof numeric === 'number' && (!Number.isFinite(numeric) || numeric >= 36526 || dateFormat && numeric >= 1);
  if (date || mrn.startsWith('◆') || /\d{3,4}[-/.]\d{1,2}[-/.]\d{1,2}/.test(raw) || /^\d{8}$/.test(raw) || dateHint) {
    if (!date || mrn !== '◆ 刀日' || !CONFIG.HOSPITAL_OPTIONS.includes(hospital) || nativeValue_(row, cols, 'CalendarEventId')) throw new Error('刀日標題日期、標記、院區或事件 ID 無效。');
    return { kind: 'DATE_HEADER', date, hospital };
  }
  return { kind: mrn || nativeValue_(row, cols, '姓名') ? 'PATIENT' : '' };
}
function nativeArchiveStructure_(sheet) {
  if ((sheet.rows || []).length < 6) throw new Error('年度封存缺少第六列表頭。');
  const info = nativeLabels_(sheet, 5), cols = info.cols, months = [], ranges = [];
  if (cols.ArchiveSourceMonth === undefined || cols.ArchiveRowType === undefined) throw new Error('封存來源月份／列型態缺少。');
  let month = '', day = false;
  sheet.rows.slice(6).forEach((row, offset) => {
    if (!row.some(cell => cell.userEnteredValue)) return;
    const source = nativeValue_(row, cols, 'ArchiveSourceMonth'), kind = nativeValue_(row, cols, 'ArchiveRowType');
    if (!/^20\d{2}(0[1-9]|1[0-2])$/.test(source) || source.slice(0, 4) !== sheet.properties.title.slice(-4)) throw new Error('封存來源月份無效。');
    if (kind === 'MONTH_HEADER') {
      if (months.includes(source) || (month && source < month) || nativeValue_(row, cols, 'ScheduleSyncId') || nativeValue_(row, cols, 'CalendarEventId')) throw new Error('封存月份順序或識別碼衝突。');
      month = source; day = false; months.push(month); ranges.push({ month, row: offset + 6 });
    } else {
      if (source !== month) throw new Error('封存區段來源月份不一致。');
      const parsed = nativeKind_(row, cols);
      if (kind === 'DATE_HEADER' && parsed.kind === kind && parsed.date.replace(/-/g, '').slice(0, 6) === month) day = true;
      else if (!(kind === 'PATIENT' && parsed.kind === kind && day)) throw new Error('封存刀日／病人列無效。');
      if (!/^[0-9a-f]{32}$/.test(nativeValue_(row, cols, 'ScheduleSyncId'))) throw new Error('封存永久識別碼不完整。');
    }
  });
  return { ...info, months, ranges };
}
function nativeIdentityLocations_(book) {
  const ids = {};
  book.sheets.forEach(sheet => {
    const title = sheet.properties.title, archive = /^刀表封存_\d{4}$/.test(title);
    if (!archive && !/^\d{6}$/.test(title) && !['FU', 'All'].includes(title)) return;
    const header = archive ? 5 : 0, cols = nativeColumns_(sheet, header);
    sheet.rows.slice(header + 1).forEach((row, index) => {
      ['ScheduleSyncId', 'CalendarEventId'].forEach(name => {
        const value = nativeValue_(row, cols, name);
        if (value) {
          if (name === 'ScheduleSyncId' && !/^[0-9a-f]{32}$/.test(value)) throw new Error('永久識別碼格式錯誤。');
          const key = name + ':' + value;
          (ids[key] || (ids[key] = [])).push(`${title}!${index + header + 2}`);
        }
      });
    });
  });
  Object.values(ids).forEach(values => values.sort());
  return ids;
}
function nativeAssertIdentities_(book, allowed) {
  const ids = nativeIdentityLocations_(book);
  Object.keys(ids).forEach(key => {
    if (ids[key].length > 1 && JSON.stringify(ids[key]) !== JSON.stringify((allowed || {})[key])) throw new Error('永久排程或 CalendarEventId 重複，請核對。');
  });
}
function nativeConstantFormula_(formula) {
  const text = formula.replace(/"(?:[^"]|"")*"/g, '').replace(/\b(?:TRUE|FALSE)\b/gi, '1');
  return /^=[0-9.\s+*/^%&<>=()\-]*$/.test(text);
}
function nativeFormulaChecks_(book, months) {
  const sourceIds = book.sheets.filter(sheet => months.includes(sheet.properties.title)).map(sheet => sheet.properties.sheetId);
  if ((book.namedRanges || []).some(item => sourceIds.includes((item.range || {}).sheetId))) throw new Error('命名範圍引用待封存月份，請先人工處理。');
  const references = value => value && typeof value === 'object' && (sourceIds.includes(value.sheetId) || Object.values(value).some(references));
  book.sheets.forEach(sheet => {
    const source = months.includes(sheet.properties.title);
    if ((sheet.charts || []).length && (source || references(sheet.charts))) throw new Error('圖表引用待封存月份，不能安全刪除來源。');
    sheet.rows.forEach((row, index) => row.forEach(cell => {
      const formula = (cell.userEnteredValue || {}).formulaValue;
      if (!formula) return;
      if (/\b(?:INDIRECT|OFFSET|IMPORTRANGE)\s*\(/i.test(formula) || months.some(month => new RegExp(`(?:'${month}'|\\b${month})\\s*!`).test(formula))) throw new Error('公式引用待刪月份或動態位址，來源保留。');
      if (source && index > 0 && !nativeConstantFormula_(formula)) throw new Error('病人資料公式不是純常數運算，請先人工處理。');
    }));
  });
}
function nativeEditable_(sheet) {
  if ((sheet.protectedRanges || []).some(item => item.warningOnly !== true && item.requestingUserCanEdit !== true)) throw new Error('目前帳號無權編輯既有保護範圍；不調整保護設定。');
}
function nativeClean_(value) {
  if (Array.isArray(value)) return value.map(nativeClean_);
  if (value && typeof value === 'object') {
    const result = {};
    Object.keys(value).sort().forEach(key => {
      const item = nativeClean_(value[key]);
      if (item !== null && item !== '' && !(Array.isArray(item) && !item.length) && !(item && typeof item === 'object' && !Object.keys(item).length)) result[key] = item;
    });
    return result;
  }
  return value;
}
function nativeProjection_(sheet) {
  const rows = sheet.rows.map(row => row.map(cell => {
    const out = nativeCell_(cell, false), fmt = out.userEnteredFormat || {};
    if (fmt.backgroundColorStyle) delete fmt.backgroundColor;
    if ((fmt.textFormat || {}).foregroundColorStyle) delete fmt.textFormat.foregroundColor;
    return nativeClean_(out);
  }));
  rows.forEach(row => { while (row.length && !Object.keys(row[row.length - 1]).length) row.pop(); });
  while (rows.length && !rows[rows.length - 1].length) rows.pop();
  const dimensions = (values, size) => {
    const result = (values || []).map(item => ({ pixelSize: item.pixelSize === undefined ? size : item.pixelSize, hiddenByUser: Boolean(item.hiddenByUser) }));
    while (result.length && result[result.length - 1].pixelSize === size && !result[result.length - 1].hiddenByUser) result.pop();
    return result;
  };
  return nativeClean_({ properties: { sheetId: sheet.properties.sheetId, title: sheet.properties.title, frozenRows: sheet.properties.gridProperties.frozenRowCount || 0 },
    rows, row_metadata: dimensions(sheet.row_metadata, 21), column_widths: dimensions(sheet.column_widths, 100),
    merges: sheet.merges || [], conditionalFormats: sheet.conditionalFormats || [], protectedRanges: sheet.protectedRanges || [], charts: sheet.charts || [],
    effects: (sheet.conditionalFormats || []).length ? sheet.rows.map(row => row.map(cell => cell.effectiveFormat || {})) : [] });
}
function nativeSame_(a, b) { return JSON.stringify(nativeProjection_(a)) === JSON.stringify(nativeProjection_(b)); }
function nativeWriteRows_(requests, target, start, rows, width) {
  const grid = target.properties.gridProperties, sheetId = target.properties.sheetId;
  if (start + rows.length > grid.rowCount) { requests.push({ appendDimension: { sheetId, dimension: 'ROWS', length: start + rows.length - grid.rowCount } }); grid.rowCount = start + rows.length; }
  while (target.rows.length < start + rows.length) target.rows.push([]);
  rows.forEach((row, index) => { target.rows[start + index] = nativeClone_(row); });
  for (let offset = 0; offset < rows.length; offset += 200) {
    const part = rows.slice(offset, offset + 200);
    requests.push({ updateCells: { range: { sheetId, startRowIndex: start + offset, endRowIndex: start + offset + part.length, startColumnIndex: 0, endColumnIndex: width },
      rows: part.map(values => ({ values: values.map(cell => nativeCell_(cell, false)) })), fields: NATIVE_CELL_FIELDS } });
  }
}
function nativeCopyDimensions_(requests, target, start, source, count) {
  for (let index = 0; index < Math.min(count, (source.row_metadata || []).length); index++) {
    const props = source.row_metadata[index];
    if (!Object.keys(props).length) continue;
    while (target.row_metadata.length <= start + index) target.row_metadata.push({});
    target.row_metadata[start + index] = nativeClone_(props);
    requests.push({ updateDimensionProperties: { range: { sheetId: target.properties.sheetId, dimension: 'ROWS', startIndex: start + index, endIndex: start + index + 1 }, properties: props, fields: 'pixelSize,hiddenByUser' } });
  }
}
function planNativeArchive_(original, selected, currentMonth, operationId) {
  if (!Array.isArray(selected) || !selected.length || selected.length > 120 || new Set(selected).size !== selected.length || selected.some(month => !/^20\d{2}(0[1-9]|1[0-2])$/.test(month) || month >= currentMonth)) throw new Error('請選擇不重複的已結束月份。');
  nativeAssertIdentities_(original); nativeFormulaChecks_(original, selected);
  const book = nativeClone_(original), requests = [], sources = {}, targets = {}, originalTargets = {}, summaries = [];
  book.sheets.filter(sheet => /^刀表封存_\d{4}$/.test(sheet.properties.title)).forEach(nativeArchiveStructure_);
  selected.slice().sort().forEach(month => {
    const found = original.sheets.find(sheet => sheet.properties.title === month);
    if (!found) throw new Error(`${month} 來源不存在，請核對原封存交易。`);
    const source = nativeClone_(found);
    const last = Math.max(source.rows.length, (source.row_metadata || []).length, ...(source.merges || []).map(range => range.endRowIndex || 0));
    while (source.rows.length < last) source.rows.push([]);
    nativeEditable_(source);
    const sourceInfo = nativeLabels_(source, 0), labels = sourceInfo.labels, sourceCols = sourceInfo.cols;
    if (labels.includes('ArchiveSourceMonth') || labels.includes('ArchiveRowType')) throw new Error('來源月表使用封存專用欄名。');
    sources[month] = nativeClone_(source);
    let day = false, patients = 0, blocks = 0;
    const kinds = source.rows.slice(1).map(row => {
      const parsed = nativeKind_(row, sourceCols);
      if (parsed.kind === 'DATE_HEADER') { if (parsed.date.replace(/-/g, '').slice(0, 6) !== month) throw new Error('刀日不屬於來源月份。'); day = true; blocks++; }
      if (parsed.kind === 'PATIENT') { if (!day) throw new Error('病人列沒有有效刀日。'); patients++; }
      const uid = nativeValue_(row, sourceCols, 'ScheduleSyncId'), event = nativeValue_(row, sourceCols, 'CalendarEventId');
      if (['ScheduleSyncId', 'CalendarEventId'].some(name => (((row[sourceCols[name]] || {}).userEnteredValue || {}).formulaValue))) throw new Error('永久識別欄不能是公式。');
      if (parsed.kind && !/^[0-9a-f]{32}$/.test(uid)) throw new Error('請先由主程式準備永久識別碼。');
      if (!parsed.kind && (uid || event || row.some(cell => cell.userEnteredValue))) throw new Error('有無法分類的人工內容，不能封存。');
      row.forEach(cell => {
        const validation = (cell.dataValidation || {}).condition || {};
        if (['CUSTOM_FORMULA', 'ONE_OF_RANGE'].includes(validation.type) || validation.type !== 'ONE_OF_LIST' && (validation.values || []).some(value => /^\s*=/.test(value.userEnteredValue || ''))) throw new Error('資料驗證引用公式／範圍，不能安全搬列。');
      });
      return parsed.kind;
    });
    if (!blocks) throw new Error('來源月份沒有完整刀日。');
    const title = `刀表封存_${month.slice(0, 4)}`;
    let target = book.sheets.find(sheet => sheet.properties.title === title), migrated = false;
    if (!(title in originalTargets)) originalTargets[title] = target ? nativeClone_(target) : null;
    if (!target) {
      let sheetId;
      do { sheetId = Math.floor(Math.random() * 1073741824) + 1; } while (book.sheets.some(sheet => sheet.properties.sheetId === sheetId));
      const header = labels.map((label, index) => ({ ...nativeCell_(source.rows[0][index] || {}, Boolean((source.conditionalFormats || []).length)), userEnteredValue: { stringValue: label } }));
      header.forEach(cell => { delete cell.textFormatRuns; });
      header.push(nativeLiteral_('ArchiveSourceMonth'), nativeLiteral_('ArchiveRowType'));
      target = { properties: { sheetId, title, gridProperties: { rowCount: 1000, columnCount: header.length, frozenRowCount: 6 } },
        rows: [[{ ...nativeLiteral_(title), note: ANNUAL_ARCHIVE_NOTE_MARKER }], [], [], [], [], header], merges: [], column_widths: [], row_metadata: [] };
      book.sheets.push(target); requests.push({ addSheet: { properties: nativeClone_(target.properties) } });
      nativeWriteRows_(requests, target, 0, nativeClone_(target.rows), header.length); nativeCopyDimensions_(requests, target, 5, source, 1);
    }
    nativeEditable_(target);
    target.row_metadata = target.row_metadata || []; target.column_widths = target.column_widths || []; target.merges = target.merges || [];
    const info = nativeArchiveStructure_(target), cols = info.cols, outputLabels = info.labels;
    if (info.months.includes(month)) throw new Error(`${month} 已封存，不能重複加入。`);
    if ((target.conditionalFormats || []).length) {
      nativeWriteRows_(requests, target, 0, target.rows.map(row => row.map(cell => nativeCell_(cell, true))), outputLabels.length);
      for (let index = target.conditionalFormats.length - 1; index >= 0; index--) requests.push({ deleteConditionalFormatRule: { sheetId: target.properties.sheetId, index } });
      target.conditionalFormats = []; migrated = true;
    }
    const sourceBlanks = labels.map((label, index) => label ? -1 : index).filter(index => index >= 0);
    const targetBlanks = outputLabels.map((label, index) => label ? -1 : index).filter(index => index >= 0);
    if (sourceBlanks.length !== targetBlanks.length) throw new Error('空白分隔欄結構不一致。');
    labels.forEach((label, index) => {
      if (label && cols[label] === undefined) {
        const column = outputLabels.length; outputLabels.push(label); cols[label] = column;
        target.rows[5][column] = { ...nativeCell_(source.rows[0][index], Boolean((source.conditionalFormats || []).length)), userEnteredValue: { stringValue: label } };
        delete target.rows[5][column].textFormatRuns;
        if (column >= target.properties.gridProperties.columnCount) { requests.push({ appendDimension: { sheetId: target.properties.sheetId, dimension: 'COLUMNS', length: column + 1 - target.properties.gridProperties.columnCount } }); target.properties.gridProperties.columnCount = column + 1; }
        requests.push({ updateCells: { range: { sheetId: target.properties.sheetId, startRowIndex: 5, endRowIndex: 6, startColumnIndex: column, endColumnIndex: column + 1 }, rows: [{ values: [target.rows[5][column]] }], fields: NATIVE_CELL_FIELDS } });
      }
    });
    const mapping = labels.map((label, index) => label ? cols[label] : targetBlanks[sourceBlanks.indexOf(index)]), width = outputLabels.length;
    const ignored = ['CalendarEventId', 'ScheduleSyncId', 'ArchiveSourceMonth', 'ArchiveRowType'];
    if (sourceBlanks.length && JSON.stringify(labels.slice(0, sourceBlanks[0]).filter(label => !ignored.includes(label))) !== JSON.stringify(outputLabels.slice(0, targetBlanks[0]).filter(label => !ignored.includes(label)))) throw new Error('空白分隔欄位置不一致。');
    const freeze = Boolean((source.conditionalFormats || []).length), rows = [], separator = Array.from({ length: width }, () => ({}));
    labels.forEach((label, index) => {
      const cell = nativeCell_(source.rows[0][index] || {}, freeze), formula = (cell.userEnteredValue || {}).formulaValue;
      delete cell.textFormatRuns;
      if (formula) {
        if ([CONFIG.MONTHLY_FIELD_HEADERS.TIME, '病歷號', '姓名', 'ScheduleSyncId', 'CalendarEventId'].includes(label)) throw new Error('識別／日期表頭不能是公式。');
        const effective = source.rows[0][index].effectiveValue;
        cell.userEnteredValue = effective && !effective.errorValue ? nativeClone_(effective) : { stringValue: source.rows[0][index].formattedValue || '' };
        cell.note = `${cell.note || ''}\n封存時固定摘要；原始表頭公式：\n${formula}`.trim();
      } else delete cell.userEnteredValue;
      separator[mapping[index]] = cell;
    });
    separator[cols[CONFIG.MONTHLY_FIELD_HEADERS.TIME]].userEnteredValue = { stringValue: month };
    separator[cols[CONFIG.MONTHLY_FIELD_HEADERS.TIME]].note = `${separator[cols[CONFIG.MONTHLY_FIELD_HEADERS.TIME]].note || ''}\n${NATIVE_ARCHIVE_NOTE}`.trim();
    separator[cols.ArchiveSourceMonth] = nativeLiteral_(month); separator[cols.ArchiveRowType] = nativeLiteral_('MONTH_HEADER'); rows.push(separator);
    source.rows.slice(1).forEach((old, index) => {
      const row = Array.from({ length: width }, () => ({}));
      old.forEach((cell, column) => { if (column < labels.length) row[mapping[column]] = nativeCell_(cell, freeze); });
      if (kinds[index]) { row[cols.ArchiveSourceMonth] = nativeLiteral_(month); row[cols.ArchiveRowType] = nativeLiteral_(kinds[index]); }
      rows.push(row);
    });
    let start = target.rows.length;
    const later = info.ranges.find(range => range.month > month), backfill = Boolean(later);
    if (later) {
      start = later.row;
      if ((target.charts || []).length || target.rows.some(row => row.some(cell => { const formula = (cell.userEnteredValue || {}).formulaValue; return formula && !nativeConstantFormula_(formula); }))) throw new Error('既有封存含相依公式或圖表，不能安全補月。');
      if ((target.merges || []).some(range => (range.startRowIndex || 0) < start && range.endRowIndex > start)) throw new Error('補月位置跨越合併儲存格。');
      requests.push({ insertDimension: { range: { sheetId: target.properties.sheetId, dimension: 'ROWS', startIndex: start, endIndex: start + rows.length }, inheritFromBefore: false } });
      requests.push({ updateDimensionProperties: { range: { sheetId: target.properties.sheetId, dimension: 'ROWS', startIndex: start, endIndex: start + rows.length }, properties: { pixelSize: 21, hiddenByUser: false }, fields: 'pixelSize,hiddenByUser' } });
      target.properties.gridProperties.rowCount += rows.length; target.rows.splice(start, 0, ...rows.map(() => []));
      while (target.row_metadata.length < start) target.row_metadata.push({});
      target.row_metadata.splice(start, 0, ...rows.map(() => ({})));
      const ranges = [...(target.merges || [])];
      (target.protectedRanges || []).forEach(value => ranges.push(value.range || {}, ...(value.unprotectedRanges || [])));
      ranges.forEach(range => { if ((range.startRowIndex || 0) >= start) range.startRowIndex = (range.startRowIndex || 0) + rows.length; if ((range.endRowIndex || 0) > start) range.endRowIndex += rows.length; });
    }
    nativeWriteRows_(requests, target, start, rows, width); nativeCopyDimensions_(requests, target, start, source, rows.length);
    (source.merges || []).forEach(merge => {
      const first = merge.startColumnIndex || 0, end = merge.endColumnIndex, columns = mapping.slice(first, end);
      if (!(merge.startRowIndex > 0) || columns.length !== end - first || columns.some((column, index) => column !== columns[0] + index)) throw new Error('合併儲存格跨越表頭或重映射欄，不能安全封存。');
      for (let row = merge.startRowIndex; row < merge.endRowIndex; row++) for (let col = first; col < end; col++) if ((row !== merge.startRowIndex || col !== first) && ((source.rows[row] || [])[col] || {}).userEnteredValue) throw new Error('合併範圍含多個內容，不能安全封存。');
      const range = { ...nativeClone_(merge), sheetId: target.properties.sheetId, startRowIndex: start + merge.startRowIndex, endRowIndex: start + merge.endRowIndex, startColumnIndex: columns[0], endColumnIndex: columns[columns.length - 1] + 1 };
      target.merges.push(range); requests.push({ mergeCells: { range, mergeType: 'MERGE_ALL' } });
    });
    while (target.column_widths.length < width) target.column_widths.push({});
    (source.column_widths || []).slice(0, labels.length).forEach((props, index) => {
      const col = mapping[index];
      if (!Object.keys(target.column_widths[col]).length && Object.keys(props).length) {
        target.column_widths[col] = nativeClone_(props);
        requests.push({ updateDimensionProperties: { range: { sheetId: target.properties.sheetId, dimension: 'COLUMNS', startIndex: col, endIndex: col + 1 }, properties: props, fields: 'pixelSize,hiddenByUser' } });
      }
    });
    ['CalendarEventId', 'ScheduleSyncId', 'ArchiveSourceMonth', 'ArchiveRowType'].forEach(label => {
      const col = cols[label]; target.column_widths[col].hiddenByUser = true;
      requests.push({ updateDimensionProperties: { range: { sheetId: target.properties.sheetId, dimension: 'COLUMNS', startIndex: col, endIndex: col + 1 }, properties: { hiddenByUser: true }, fields: 'hiddenByUser' } });
    });
    targets[title] = nativeClone_(target); summaries.push({ month, target: title, patients, blocks, copied_rows: rows.length, backfill, migrated_existing_format: migrated });
  });
  return { transaction_id: operationId, requests, sources, targets, original_targets: originalTargets, book, summary: { months: summaries, patients: summaries.reduce((sum, item) => sum + item.patients, 0) } };
}
function verifyNativeArchive_(plan, fresh, complete) {
  try {
    const find = title => fresh.sheets.find(sheet => sheet.properties.title === title);
    if (Object.keys(plan.sources).some(title => complete ? Boolean(find(title)) : !find(title) || !nativeSame_(plan.sources[title], find(title)))) return false;
    if (Object.keys(plan.targets).some(title => !find(title) || !nativeSame_(plan.targets[title], find(title)))) return false;
    Object.values(plan.targets).forEach(nativeArchiveStructure_);
    const expected = { sheets: [...Object.values(plan.targets), ...(!complete ? Object.values(plan.sources) : [])] };
    const allowed = nativeIdentityLocations_(expected); nativeAssertIdentities_(fresh, allowed);
    return true;
  } catch (_) { return false; }
}
