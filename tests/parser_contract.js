'use strict';
// Actual GAS parsers exported as typed-sheet fixtures for the Windows EXEs.
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path');
const { call, evaluate, plain, FakeSheet, makeFakeSpreadsheet, context } = require('./smoke');

function dateValue(input, type) {
  if (type !== 'serial') return input;
  const value = new Date(1899, 11, 30 + Math.floor(input));
  value.setMilliseconds(Math.round((input - Math.floor(input)) * 86400000));
  return value;
}

function parserContract() {
  const config = plain(evaluate('CONFIG'));
  const headers = config.MONTHLY_HEADERS;
  const sheet = new FakeSheet('202610', 1, [headers]);
  const columns = call('getRequiredMonthlyColumns_', sheet);
  const dates = [
    ['2026-10-11'], ['2026/1/2'], ['2026.10.11'], ['20261011'], ['2026/10/11 Sun'], ['2026/10/11 sUn'],
    ['2026/10/11 Mon'], ['2026/10/11 Sunday'], ['2026/10/11 註記'], ['回診 2026/10/11'], ['115/10/11'],
    ['1999/12/31'], ['2100/1/1'], ['2026/2/29'], ['2024/2/29'], ['08:30'], ['800'],
    [36526, 'serial'], [73050, 'serial'], [36525, 'serial'], [73051, 'serial'], [46306.5, 'serial'],
    [46306 + 1 / 86400000, 'serial'],
    [36526, 'number'], [73050, 'number'], [36525, 'number'], [73051, 'number'],
    [20261011, 'number'], [46306.5, 'number'], [0.25, 'number'], [830, 'number']
  ].map(([input, type = 'text']) => {
    const result = call('parseFullDate_', dateValue(input, type));
    return { input, type, iso: result ? call('formatDateKey_', result) : '' };
  });
  const rows = [
    { date: '2026-10-11', marker: '◆ 刀日', hospital: '高榮' },
    { date: '2026-10-11', marker: '◆　刀日', hospital: '聯醫' },
    { date: '2026-10-11', marker: '◆  刀日', hospital: '聯醫' },
    { date: '2026-10-11', marker: '◆ 刀日', hospital: '' },
    { date: '2026-10-11', marker: '◆ 刀日', hospital: '其他' },
    { date: '2026-10-11', marker: '', hospital: '高榮' },
    { date: '2026-10-11', marker: 'TEST001', hospital: '高榮' },
    { date: '2026-10-11', marker: '◆ 刀日', hospital: '高榮', event: 'event-1' },
    { date: '2026/10/11 Mon', marker: '◆ 刀日', hospital: '高榮' },
    { date: '08:30', marker: '◆ 刀日', hospital: '高榮' },
    { date: '08:30', marker: 'TEST001', hospital: '高榮' },
    { date: '2026-10-12 09:30', marker: 'TEST001', hospital: '高榮' },
    { date: '2026-02-30', marker: 'TEST001', hospital: '高榮' },
    { date: '回診 2026-10-11', marker: 'TEST001', hospital: '高榮' },
    { date: 46306, marker: '◆ 刀日', hospital: '高榮' },
    { date: 46306.5, marker: '◆ 刀日', hospital: '高榮' },
    { date: 830, marker: 'TEST001', hospital: '高榮' },
    { date: '', marker: '', hospital: '' }
  ].map(input => {
    const row = headers.map(header => ({'日期／報到時間': input.date, '醫院': input.hospital,
      '病歷號': input.marker, CalendarEventId: input.event || ''})[header] || '');
    return { ...input, kind: call('classifyMonthlyRowValues_', row, columns).type };
  });
  const headerCases = [
    { header: '時間', key: 'TIME' }, { header: '日期/報到時間', key: 'TIME' },
    { header: '日期／報到時間', key: 'TIME' },
    { header: '診斷｜Cat 1', key: 'DIAGNOSIS' },
    { header: '', formula: '=T(N("SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V2"))&"診斷"', key: 'DIAGNOSIS' },
    { header: '#ERROR!', formula: '=T(N("SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V2"))&"診斷"', key: 'DIAGNOSIS' },
    { header: '', formula: '=IF(TRUE,"診斷","")', key: 'DIAGNOSIS' },
    { header: '診斷人工備註', key: 'DIAGNOSIS' },
    { header: '日曆eventID', key: 'EVENT_ID' }, { header: '日曆apiEventID', key: 'EVENT_ID' }
  ].map(input => {
    const changed = headers.slice(), index = columns[input.key] - 1;
    changed[index] = input.header;
    const sample = new FakeSheet('202610', 2, [changed]);
    if (input.formula) sample.setFormulaAt(1, index + 1, input.formula);
    const info = call('getMonthlyColumnInfo_', sample);
    return { ...input, canonical: info.columns[input.key] ? config.MONTHLY_FIELD_HEADERS[input.key] : '' };
  });
  const times = ['08:30', '8:30', '0830', '830', '08：30', '', '下午', '25:00', '1260', '8', 0, 0.25, 830].map(input => {
    const result = call('resolveScheduleTime_', input);
    return { input, time: result.parsedTime ? result.parsedTime.text : '', note: result.timeNote, invalid: !!result.errorMessage };
  });
  return { contract_version: 1, schema: 1, date_cases: dates, row_cases: rows, header_cases: headerCases, time_cases: times };
}

function templateChecks() {
  const config = plain(evaluate('CONFIG'));
  const reordered = ['ScheduleSyncId', '人工備註', ...config.MONTHLY_HEADERS.filter(value => value !== 'ScheduleSyncId')];
  const older = new FakeSheet('202609', 10, [reordered]);
  const latest = new FakeSheet('202610', 11, [reordered]);
  const diagnosis = call('getRequiredMonthlyColumns_', latest).DIAGNOSIS;
  latest.setFormulaAt(1, diagnosis, '=T(N("SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V2"))&"診斷"');
  latest.setNoteAt(1, 2, '人工表頭註記');
  assert.strictEqual(call('getMonthlyCreationTemplate_', makeFakeSpreadsheet([older, latest])), latest);
  const mismatch = new FakeSheet('202608', 12, [config.MONTHLY_HEADERS]);
  assert.throws(() => call('getMonthlyCreationTemplate_', makeFakeSpreadsheet([latest, mismatch])), /先對齊/);
  const sheets = [older, latest];
  const spreadsheet = makeFakeSpreadsheet(sheets);
  spreadsheet.insertSheet = name => { const sheet = new FakeSheet(name, 13, [[]]); sheets.push(sheet); return sheet; };
  const rangePrototype = Object.getPrototypeOf(latest.getRange(1, 1));
  const savedCopy = rangePrototype.copyTo;
  rangePrototype.copyTo = function(target) {
    for (let col = 0; col < this.numColumns; col++) {
      target.sheet.setValueAt(target.row, target.column + col, this.sheet.valueAt(this.row, this.column + col));
      target.sheet.setFormulaAt(target.row, target.column + col, this.sheet.formulaAt(this.row, this.column + col));
      target.sheet.setNoteAt(target.row, target.column + col, this.sheet.noteAt(this.row, this.column + col));
    }
  };
  const saved = { applyMonthlyFormatting_: context.applyMonthlyFormatting_, ensureManagedHeaders_: context.ensureManagedHeaders_, ensureDefaultMonthlyBlocks_: context.ensureDefaultMonthlyBlocks_ };
  context.applyMonthlyFormatting_ = sheet => ({ columns: call('getRequiredMonthlyColumns_', sheet) });
  context.ensureManagedHeaders_ = sheet => call('getRequiredMonthlyColumns_', sheet);
  context.ensureDefaultMonthlyBlocks_ = () => 0;
  try {
    const result = call('createOrInitializeMonthlySheet_', spreadsheet, '202611');
    assert.deepStrictEqual(result.sheet.rows[0], reordered);
    assert.strictEqual(result.sheet.formulaAt(1, diagnosis), latest.formulaAt(1, diagnosis));
    assert.strictEqual(result.sheet.noteAt(1, 2), '人工表頭註記');
  } finally { Object.assign(context, saved); rangePrototype.copyTo = savedCopy; }
}

module.exports = { parserContract };
if (require.main === module) {
  templateChecks();
  const file = path.join(__dirname, 'fixtures/schema_parsing.json'), actual = parserContract();
  if (process.argv.includes('--record')) fs.writeFileSync(file, JSON.stringify(actual, null, 2) + '\n');
  else assert.deepStrictEqual(actual, JSON.parse(fs.readFileSync(file, 'utf8')));
  console.log('Parser contract passed (strict dates, headers, hospitals, markers, time, new-month schema)');
}
