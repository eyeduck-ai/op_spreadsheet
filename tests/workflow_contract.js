'use strict';
// Export cells produced by real GAS month/day/patient writes, then read them in
// both Windows programs. Only the native Sheet service is synthetic.
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path');
const { call, evaluate, plain, FakeSheet, makeFakeSpreadsheet, context } = require('./smoke');

function nativeSheet(name, id, rows) {
  const sheet = new FakeSheet(name, id, rows);
  sheet.hideColumns = () => {};
  sheet.insertRowsBefore = (row, count) => {
    sheet.rows.splice(row - 1, 0, ...Array.from({ length: count }, () => []));
    ['notes', 'formulas', 'fontLines', 'validations'].forEach(field => {
      const shifted = {};
      Object.entries(sheet[field]).forEach(([key, value]) => {
        const [oldRow, column] = key.split(':').map(Number);
        shifted[`${oldRow >= row ? oldRow + count : oldRow}:${column}`] = value;
      });
      sheet[field] = shifted;
    });
  };
  sheet.insertRowBefore = row => sheet.insertRowsBefore(row, 1);
  return sheet;
}

function cells(sheet) {
  return {
    name: sheet.getName(), id: sheet.getSheetId(),
    rows: sheet.rows.map(row => {
      const values = row.map(value => value instanceof Date
        ? Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()) / 86400000 + 25569 : value);
      while (values.length && values[values.length - 1] === '') values.pop();
      return values;
    })
  };
}

function workflowContract() {
  const config = plain(evaluate('CONFIG'));
  const labels = ['ScheduleSyncId', '人工備註', ...config.MONTHLY_HEADERS.filter(value =>
    value !== 'ScheduleSyncId' && value !== 'CalendarEventId'), 'CalendarEventId'];
  const template = nativeSheet('202610', 1, [labels]);
  template.setNoteAt(1, 2, '人工表頭註記');
  const sheets = [template], spreadsheet = makeFakeSpreadsheet(sheets);
  spreadsheet.insertSheet = name => {
    const sheet = nativeSheet(name, sheets.length + 1, [[]]);
    sheets.push(sheet);
    return sheet;
  };
  const prototype = Object.getPrototypeOf(template.getRange(1, 1));
  const savedCopy = prototype.copyTo, savedFormat = context.applyMonthlyFormatting_;
  const savedUuid = context.Utilities.getUuid;
  let nextId = 1;
  context.Utilities.getUuid = () => (nextId++).toString(16).padStart(32, '0');
  prototype.copyTo = function(target) {
    for (let offset = 0; offset < this.numColumns; offset++) {
      const column = this.column + offset, out = target.column + offset;
      target.sheet.setValueAt(target.row, out, this.sheet.valueAt(this.row, column));
      target.sheet.setFormulaAt(target.row, out, this.sheet.formulaAt(this.row, column));
      target.sheet.setNoteAt(target.row, out, this.sheet.noteAt(this.row, column));
    }
  };
  // Display formatting is covered by sheet_contract.js; all data writes below
  // execute the deployed functions unchanged, including default day creation.
  context.applyMonthlyFormatting_ = sheet => ({ columns: call('getRequiredMonthlyColumns_', sheet) });
  try {
    const created = call('createOrInitializeMonthlySheet_', spreadsheet, '202611');
    assert.strictEqual(created.created, true);
    assert.strictEqual(created.addedBlocks, 4);
    const sheet = created.sheet, columns = created.columns;
    assert.deepStrictEqual(sheet.rows[0], labels);
    assert.strictEqual(sheet.noteAt(1, 2), '人工表頭註記');
    const date = new Date(2026, 10, 16);
    const high = call('findOrCreateMonthlyBlock_', sheet, date, '高榮', columns);
    const union = call('findOrCreateMonthlyBlock_', sheet, date, '聯醫', columns);
    assert.notStrictEqual(high.row, union.row);
    assert.strictEqual(call('findOrCreateMonthlyBlock_', sheet, date, '高榮', columns).row, high.row);
    const newRow = call('insertPatientAtBlockEnd_', sheet, union,
      { CHART_NO: 'TEST001', NAME: '合成甲', TIME: '08:30', HOSPITAL: '高榮', EVENT_ID: 'must-not-copy' }, columns);
    assert.strictEqual(newRow, union.row + 1);
    assert.strictEqual(sheet.valueAt(newRow, columns.HOSPITAL), '聯醫');
    assert.strictEqual(sheet.valueAt(newRow, columns.EVENT_ID), '');
    const syncId = sheet.valueAt(newRow, columns.SYNC_ID);
    sheet.setValueAt(newRow, columns.EVENT_ID, 'synthetic-original-event');
    sheet.setValueAt(newRow, 2, '人工欄位保留');
    const creation = sheets.map(cells);

    // Empty-display formulas and notes must force insertion of a complete row.
    const blankRow = newRow + 1;
    sheet.setFormulaAt(blankRow, 2, '=IF(TRUE,"","")');
    sheet.setNoteAt(blankRow, 2, '不可覆寫');
    const another = call('insertPatientAtBlockEnd_', sheet, union,
      { CHART_NO: 'TEST002', NAME: '合成乙' }, columns);
    assert.strictEqual(another, blankRow);
    assert.strictEqual(sheet.formulaAt(blankRow + 1, 2), '=IF(TRUE,"","")');
    assert.strictEqual(sheet.noteAt(blankRow + 1, 2), '不可覆寫');
    assert.notStrictEqual(sheet.valueAt(another, columns.SYNC_ID), syncId);

    const destination = call('createOrInitializeMonthlySheet_', spreadsheet, '202612').sheet;
    const targetCols = call('getRequiredMonthlyColumns_', destination);
    const target = call('findOrCreateMonthlyBlock_', destination, new Date(2026, 11, 7), '高榮', targetCols);
    const moved = sheet.rows.splice(newRow - 1, 1)[0];
    destination.rows[target.row] = moved;
    const located = call('locateCachedMonthlySchedule_', spreadsheet, { syncId }, { chartNo: 'TEST001' });
    assert.strictEqual(located.sheet, destination);
    assert.strictEqual(destination.valueAt(located.row, targetCols.EVENT_ID), 'synthetic-original-event');
    assert.strictEqual(destination.valueAt(located.row, targetCols.TIME), '08:30');
    assert.strictEqual(destination.valueAt(located.row, 2), '人工欄位保留');
    assert.throws(() => call('locateCachedMonthlySchedule_', spreadsheet, { syncId }, { chartNo: 'WRONG' }), /病人已變更/);
    assert.strictEqual(call('scanScheduleIdentities_', spreadsheet).conflicts.length, 0);
    return { schema: 1, creation, moved: sheets.map(cells), patient_sync_id: syncId,
      patient_event_id: 'synthetic-original-event', creation_date: '2026-11-16', creation_hospital: '聯醫',
      moved_date: '2026-12-07', moved_hospital: '高榮' };
  } finally {
    prototype.copyTo = savedCopy;
    context.applyMonthlyFormatting_ = savedFormat;
    context.Utilities.getUuid = savedUuid;
  }
}

module.exports = { workflowContract };
if (require.main === module) {
  const file = path.join(__dirname, 'fixtures/workflow_contract.json'), actual = workflowContract();
  if (process.argv.includes('--record')) fs.writeFileSync(file, JSON.stringify(actual, null, 2) + '\n');
  else assert.deepStrictEqual(actual, JSON.parse(fs.readFileSync(file, 'utf8')));
  console.log('Workflow contract passed (new month, date/campus, empty row, notes/formulas, permanent ID retry and cross-month move)');
}
