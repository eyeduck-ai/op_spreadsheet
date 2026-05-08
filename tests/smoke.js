const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const codePath = path.join(__dirname, '..', 'code.js');
const code = fs.readFileSync(codePath, 'utf8');
const context = {
  console
};

vm.createContext(context);
vm.runInContext(code, context, { filename: codePath });

function makeHeaderSheet(headers) {
  return {
    getMaxColumns: () => headers.length,
    getLastColumn: () => headers.length,
    getRange: (_row, column, _numRows, numColumns) => ({
      getValues: () => [headers.slice(column - 1, column - 1 + numColumns)]
    })
  };
}

function rowFromObject(headers, values) {
  return headers.map(header => Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '');
}

function testHeaderMappingWithCustomColumn() {
  const headers = ['病歷號', '姓名', '自訂欄', 'TEL', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', '日曆eventID'];
  const info = context.buildFieldColumnInfo_(makeHeaderSheet(headers));

  assert.deepStrictEqual(Object.assign({}, info.columns), {
    CHART_NO: 1,
    NAME: 2,
    TEL: 4,
    TAG: 5,
    COND: 6,
    DATE: 7,
    TIME: 8,
    PLAN: 9,
    MEMO: 10,
    EVENT_ID: 11
  });
}

function testTimeNoteRoundTrip() {
  const description = 'TEL: 0912345678\nPlan: PE/IOL\nTime note: GA';
  const descriptionInfo = context.parseCalendarDescription_(description);

  assert.strictEqual(descriptionInfo.hasTimeNote, true);
  assert.strictEqual(descriptionInfo.timeNote, 'GA');
  assert.strictEqual(context.getSheetTimeValueFromCalendarEvent_({
    isAllDay: true,
    timeText: ''
  }, descriptionInfo), 'GA');
}

function testMergeBlocksTrackedAbsorbedRows() {
  const headers = ['病歷號', '姓名', 'TEL', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', '日曆eventID'];
  const cols = context.buildFieldColumnInfo_(makeHeaderSheet(headers)).columns;
  const data = [
    headers,
    rowFromObject(headers, { '病歷號': '123', '姓名': '王小明', Tag: 'FU', TEL: '0911', Condition: 'A' }),
    rowFromObject(headers, { '病歷號': '123', '姓名': '王小明', Tag: 'FU', TEL: '0922', Condition: 'B', '日曆eventID': 'evt_1' })
  ];
  const result = context.buildMergedPatientRows_(data, cols);

  assert.strictEqual(result.ok, false);
  assert.deepStrictEqual(Array.from(result.blockedRows), [3]);
}

function testMergeCombinesSafeTextFields() {
  const headers = ['病歷號', '姓名', 'TEL', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', '日曆eventID'];
  const cols = context.buildFieldColumnInfo_(makeHeaderSheet(headers)).columns;
  const data = [
    headers,
    rowFromObject(headers, { '病歷號': '123', '姓名': '王小明', Tag: 'FU', TEL: '0911', Condition: 'A', Plan: 'P1' }),
    rowFromObject(headers, { '病歷號': '123', '姓名': '王小明', Tag: 'OPD', TEL: '0922', Condition: 'B', Plan: 'P2', '心得': 'note' })
  ];
  const result = context.buildMergedPatientRows_(data, cols);
  const mergedRow = result.rows[0];

  assert.strictEqual(result.ok, true);
  assert.strictEqual(context.getRowFieldValue_(mergedRow, cols, 'TEL'), '0911\n0922');
  assert.strictEqual(context.getRowFieldValue_(mergedRow, cols, 'COND'), 'A\nB');
  assert.strictEqual(context.getRowFieldValue_(mergedRow, cols, 'PLAN'), 'P1\nP2');
  assert.strictEqual(context.getRowFieldValue_(mergedRow, cols, 'MEMO'), 'note');
  assert.strictEqual(context.getRowFieldValue_(mergedRow, cols, 'TAG'), 'FU\nOPD');
}

function testConditionalFormulaBuilding() {
  const yellowExpression = context.buildKeywordExpression_(9, ['問', '補', '通知']);
  const greenExpression = context.buildKeywordExpression_(9, ['APPLY']);
  const gaExpression = context.buildKeywordExpression_(8, ['GA']);

  assert.ok(yellowExpression.includes('SEARCH("問", $I2)'));
  assert.ok(yellowExpression.includes('SEARCH("補", $I2)'));
  assert.ok(yellowExpression.includes('SEARCH("通知", $I2)'));
  assert.strictEqual(greenExpression, 'ISNUMBER(SEARCH("APPLY", $I2))');
  assert.strictEqual(gaExpression, 'ISNUMBER(SEARCH("GA", $H2))');
}

testHeaderMappingWithCustomColumn();
testTimeNoteRoundTrip();
testMergeBlocksTrackedAbsorbedRows();
testMergeCombinesSafeTextFields();
testConditionalFormulaBuilding();

console.log('Smoke tests passed');
