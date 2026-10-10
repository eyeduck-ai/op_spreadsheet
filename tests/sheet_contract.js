'use strict';
// Export the real GAS sheet model for offline cross-repository validation.
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path');
const { call, evaluate, plain, FakeSheet, context } = require('./smoke');

function sheetContract() {
  const config = plain(evaluate('CONFIG'));
  const labelsFor = (values, headers) => values.map(key => headers[key]);
  const widths = Object.fromEntries(Object.entries(plain(evaluate('MONTHLY_COLUMN_WIDTHS')))
    .map(([key, width]) => [config.MONTHLY_FIELD_HEADERS[key], width]));
  function rulesFor(headers) {
    const sheet = new FakeSheet('202610', 1, [headers, Array(headers.length).fill('')]);
    const columns = call('getRequiredMonthlyColumns_', sheet);
    const saved = context.SpreadsheetApp.newConditionalFormatRule;
    context.SpreadsheetApp.newConditionalFormatRule = () => {
      const state = {};
      const builder = {
        whenFormulaSatisfied: value => { state.formula = value; return builder; },
        setBackground: value => { state.color = value.slice(1); return builder; },
        setBold: value => { state.bold = value; return builder; },
        setRanges: values => { state.ranges = values.map(range => [range.getColumn() - 1,
          range.getColumn() - 1 + range.getNumColumns()]); return builder; },
        build: () => ({ bold: false, ...state })
      };
      return builder;
    };
    try {
      call('applyMonthlyConditionalFormats_', sheet, columns, headers.length);
      return plain(sheet.getConditionalFormatRules());
    } finally { context.SpreadsheetApp.newConditionalFormatRule = saved; }
  }
  const reordered = ['ScheduleSyncId', '人工備註', ...config.MONTHLY_HEADERS.filter(value =>
    value !== 'ScheduleSyncId' && value !== 'CalendarEventId'), 'CalendarEventId'];
  const months = ['202402', '202608', '202609', '202610', '202611', '202701'];
  const standardSheet = new FakeSheet('202610', 1, [config.MONTHLY_HEADERS]);
  return plain({
    schema: 1,
    monthly_headers: config.MONTHLY_HEADERS,
    fu_headers: config.HEADERS,
    system_headers: labelsFor(config.SYSTEM_FIELD_KEYS, config.FIELD_HEADERS),
    hospitals: config.HOSPITAL_OPTIONS,
    options: { '醫院': config.HOSPITAL_OPTIONS, GA: config.MONTHLY_GA_OPTIONS,
      '側別': config.MONTHLY_SIDE_OPTIONS, '診斷': config.MONTHLY_DIAGNOSIS_OPTIONS,
      '術式': config.MONTHLY_PROCEDURE_OPTIONS, Tag: config.FU_TAG_OPTIONS },
    monthly_wrap: labelsFor(plain(evaluate('MONTHLY_WRAP_KEYS')), config.MONTHLY_FIELD_HEADERS).sort(),
    fu_wrap: labelsFor(plain(evaluate('FU_WRAP_KEYS')), config.FIELD_HEADERS).sort(),
    monthly_text: labelsFor(plain(evaluate('MONTHLY_TEXT_KEYS')), config.MONTHLY_FIELD_HEADERS).sort(),
    monthly_widths: widths,
    font: { family: evaluate('MANAGED_FONT_FAMILY'), size: evaluate('MANAGED_FONT_SIZE') },
    blank_rows: evaluate('MONTHLY_TEMPLATE_BLANK_ROWS'),
    header_marker: evaluate('MONTHLY_DATE_HEADER_MARKER'),
    diagnosis_formula: call('buildMonthlyDiagnosisSummaryFormula_', call('getRequiredMonthlyColumns_', standardSheet)),
    default_dates: Object.fromEntries(months.map(month => [month,
      call('getDefaultMonthlyDates_', month).map(date => call('formatDateKey_', date))])),
    rule_cases: [config.MONTHLY_HEADERS, reordered].map(headers => ({ headers, rules: rulesFor(headers) }))
  });
}

module.exports = { sheetContract };
if (require.main === module) {
  const file = path.join(__dirname, 'fixtures/sheet_contract.json');
  const actual = sheetContract();
  if (process.argv.includes('--record')) {
    fs.mkdirSync(path.dirname(file), { recursive: true });
    fs.writeFileSync(file, JSON.stringify(actual, null, 2) + '\n');
  } else assert.deepStrictEqual(actual, JSON.parse(fs.readFileSync(file, 'utf8')));
  console.log('Sheet contract passed (headers, formats, month dates, formula, reordered conditional rules)');
}
