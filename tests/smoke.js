const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const codePath = path.join(__dirname, '..', 'code.js');
const code = fs.readFileSync(codePath, 'utf8');

function pad2(value) {
  return String(value).padStart(2, '0');
}

const context = {
  console,
  Session: {
    getScriptTimeZone: () => 'Asia/Taipei'
  },
  Utilities: {
    formatDate: date => {
      return `${date.getFullYear()}/${pad2(date.getMonth() + 1)}/${pad2(date.getDate())}`;
    },
    formatString: (format, ...values) => {
      let index = 0;
      return format.replace(/%02d/g, () => pad2(values[index++]));
    }
  }
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

function makeDataSheet(headers, rows) {
  const data = [headers].concat(rows.map(row => row.slice()));
  const notes = {};

  return {
    rows: data,
    notes,
    getMaxColumns: () => headers.length,
    getLastColumn: () => headers.length,
    getLastRow: () => data.length,
    getRange: (row, column, numRows, numColumns) => ({
      getValues: () => data
        .slice(row - 1, row - 1 + numRows)
        .map(dataRow => dataRow.slice(column - 1, column - 1 + numColumns)),
      getCell: (relativeRow, relativeColumn) => {
        const rowIndex = row + relativeRow - 2;
        const columnIndex = column + relativeColumn - 2;
        const noteKey = `${rowIndex}:${columnIndex}`;

        return {
          clearContent: () => {
            data[rowIndex][columnIndex] = '';
          },
          getNote: () => notes[noteKey] || '',
          clearNote: () => {
            notes[noteKey] = '';
          },
          setNote: note => {
            notes[noteKey] = note;
          }
        };
      }
    })
  };
}

function rowFromObject(headers, values) {
  return headers.map(header => Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '');
}

function toPlain(value) {
  return JSON.parse(JSON.stringify(value));
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

function testEventIdColumnLookupWithMovedHeader() {
  const headers = ['病歷號', '日曆eventID', '姓名', 'TEL', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得'];
  assert.strictEqual(context.getEventIdColumnForSheet_(makeHeaderSheet(headers)), 2);
}

function testEventIdColumnLookupMissingHeader() {
  const headers = ['病歷號', '姓名', 'TEL', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得'];
  assert.strictEqual(context.getEventIdColumnForSheet_(makeHeaderSheet(headers)), 0);
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

function makeConditionalFormatSheet(headers) {
  let rules = [];

  return {
    getMaxRows: () => 20,
    getMaxColumns: () => headers.length,
    getLastRow: () => 5,
    getLastColumn: () => headers.length,
    getConditionalFormatRules: () => [],
    setConditionalFormatRules: nextRules => {
      rules = nextRules;
    },
    getCapturedRules: () => rules,
    getRange: (row, column, numRows, numColumns) => ({
      row,
      column,
      numRows,
      numColumns,
      getRow: () => row,
      getColumn: () => column,
      getNumRows: () => numRows,
      getNumColumns: () => numColumns,
      getValues: () => row === 1 ? [headers.slice(column - 1, column - 1 + numColumns)] : []
    })
  };
}

function installConditionalFormatMock() {
  context.SpreadsheetApp = {
    BooleanCriteria: {
      CUSTOM_FORMULA: 'CUSTOM_FORMULA',
      TEXT_EQUAL_TO: 'TEXT_EQUAL_TO'
    },
    newConditionalFormatRule: () => {
      const rule = {
        formula: '',
        background: '',
        fontColor: '',
        ranges: []
      };
      return {
        whenFormulaSatisfied(formula) {
          rule.formula = formula;
          return this;
        },
        setBackground(background) {
          rule.background = background;
          return this;
        },
        setFontColor(fontColor) {
          rule.fontColor = fontColor;
          return this;
        },
        setRanges(ranges) {
          rule.ranges = ranges;
          return this;
        },
        build() {
          return rule;
        }
      };
    }
  };
}

function testConditionalFormattingKeepsTagIndependent() {
  installConditionalFormatMock();
  const headers = ['病歷號', '姓名', 'TEL', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', '日曆eventID'];
  const sheet = makeConditionalFormatSheet(headers);
  const columns = context.buildFieldColumnInfo_(makeHeaderSheet(headers)).columns;

  context.applyConditionalFormatting_(sheet, columns);

  const rules = sheet.getCapturedRules();
  const opRule = rules.find(rule => rule.formula.includes('SYSTEM_CF_OP_RED'));
  const greenRule = rules.find(rule => rule.formula.includes('SYSTEM_CF_PLAN_GREEN"'));
  const yellowRule = rules.find(rule => rule.formula.includes('SYSTEM_CF_PLAN_YELLOW"'));

  assert.ok(opRule);
  assert.ok(!opRule.formula.includes('NOT('));
  assert.deepStrictEqual(Array.from(opRule.ranges.map(range => range.column)), [columns.TAG]);

  [greenRule, yellowRule].forEach(rule => {
    assert.ok(rule);
    rule.ranges.forEach(range => {
      assert.notStrictEqual(range.column, columns.TAG);
      assert.notStrictEqual(range.column, columns.TIME);
      assert.ok(range.column + range.numColumns - 1 < columns.TAG || range.column > columns.TAG);
      assert.ok(range.column + range.numColumns - 1 < columns.TIME || range.column > columns.TIME);
    });
  });
}

function testClearOldEventsUsesFixedOpEventIdColumn() {
  const originalSpreadsheetApp = context.SpreadsheetApp;
  const originalGetConfiguredCalendarId = context.getConfiguredCalendarId_;
  const originalIsCalendarAdvancedServiceAvailable = context.isCalendarAdvancedServiceAvailable_;
  const originalWithCalendarSyncLock = context.withCalendarSyncLock_;
  const originalDeleteCalendarEventByPossibleId = context.deleteCalendarEventByPossibleId_;
  const headers = ['病歷號', '姓名', '日曆eventID', 'TEL'];
  const opSheet = makeDataSheet(headers, [
    ['001', '王小明', 'evt_deleted', '0911'],
    ['002', '陳小美', '', '0922'],
    ['003', '林大明', 'evt_missing', '0933'],
    ['004', '黃小美', 'evt_failed', '0944']
  ]);
  const alerts = [];
  const deleteCalls = [];
  const ui = {
    Button: { OK: 'OK' },
    ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL' },
    alert: (...args) => {
      alerts.push(args);
      return 'OK';
    }
  };

  try {
    context.SpreadsheetApp = {
      getUi: () => ui,
      getActiveSpreadsheet: () => ({
        getSheetByName: name => name === 'OP' ? opSheet : null
      })
    };
    context.getConfiguredCalendarId_ = () => 'calendar_1';
    context.isCalendarAdvancedServiceAvailable_ = () => true;
    context.withCalendarSyncLock_ = callback => {
      callback();
      return true;
    };
    context.deleteCalendarEventByPossibleId_ = (_calendarId, eventId) => {
      deleteCalls.push(eventId);
      if (eventId === 'evt_failed') {
        return { ok: false, message: 'boom' };
      }
      return { ok: true, status: eventId === 'evt_missing' ? 'already_missing' : 'deleted' };
    };

    context.clearCalendarEventsFromSpecifiedColumn();

    assert.deepStrictEqual(deleteCalls, ['evt_deleted', 'evt_missing', 'evt_failed']);
    assert.deepStrictEqual(opSheet.rows, [
      headers,
      ['001', '王小明', '', '0911'],
      ['002', '陳小美', '', '0922'],
      ['003', '林大明', '', '0933'],
      ['004', '黃小美', 'evt_failed', '0944']
    ]);
    assert.ok(alerts[0][1].includes('OP'));
    assert.ok(alerts[0][1].includes('C 欄'));
    assert.strictEqual(alerts[1][0], '清除完成');
  } finally {
    context.SpreadsheetApp = originalSpreadsheetApp;
    context.getConfiguredCalendarId_ = originalGetConfiguredCalendarId;
    context.isCalendarAdvancedServiceAvailable_ = originalIsCalendarAdvancedServiceAvailable;
    context.withCalendarSyncLock_ = originalWithCalendarSyncLock;
    context.deleteCalendarEventByPossibleId_ = originalDeleteCalendarEventByPossibleId;
  }
}

function testClearOldEventsMissingEventIdHeaderAborts() {
  const originalSpreadsheetApp = context.SpreadsheetApp;
  const originalGetConfiguredCalendarId = context.getConfiguredCalendarId_;
  const originalIsCalendarAdvancedServiceAvailable = context.isCalendarAdvancedServiceAvailable_;
  const originalWithCalendarSyncLock = context.withCalendarSyncLock_;
  const alerts = [];
  let lockCalled = false;
  const ui = {
    Button: { OK: 'OK' },
    ButtonSet: { OK: 'OK', OK_CANCEL: 'OK_CANCEL' },
    alert: (...args) => {
      alerts.push(args);
      return 'OK';
    }
  };

  try {
    context.SpreadsheetApp = {
      getUi: () => ui,
      getActiveSpreadsheet: () => ({
        getSheetByName: name => name === 'OP'
          ? makeHeaderSheet(['病歷號', '姓名', 'TEL', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得'])
          : null
      })
    };
    context.getConfiguredCalendarId_ = () => 'calendar_1';
    context.isCalendarAdvancedServiceAvailable_ = () => true;
    context.withCalendarSyncLock_ = () => {
      lockCalled = true;
      return true;
    };

    context.clearCalendarEventsFromSpecifiedColumn();

    assert.strictEqual(lockCalled, false);
    assert.strictEqual(alerts.length, 1);
    assert.ok(alerts[0][0].includes('日曆eventID'));
  } finally {
    context.SpreadsheetApp = originalSpreadsheetApp;
    context.getConfiguredCalendarId_ = originalGetConfiguredCalendarId;
    context.isCalendarAdvancedServiceAvailable_ = originalIsCalendarAdvancedServiceAvailable;
    context.withCalendarSyncLock_ = originalWithCalendarSyncLock;
  }
}

function testSurgeryExportDataForDate() {
  const headers = ['病歷號', '姓名', 'TEL', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', '日曆eventID'];
  const cols = context.buildFieldColumnInfo_(makeHeaderSheet(headers)).columns;
  const targetDate = new Date(2026, 4, 9);
  const data = [
    headers,
    rowFromObject(headers, {
      '病歷號': '001',
      '姓名': '王小明',
      TEL: '0911',
      Tag: 'OP',
      Condition: 'Cataract s/p PE/IOL(B -0.5 20.0)\n術後注意',
      '日期': new Date(2026, 4, 9),
      '時間': '08:30',
      Plan: '帶藥'
    }),
    rowFromObject(headers, {
      '病歷號': '002',
      '姓名': '陳小美',
      TEL: '0922',
      Tag: 'OP',
      Condition: 'Cataract s/p MSICS(A -1.0 21.5)',
      '日期': new Date(2026, 4, 9),
      '時間': '09:30',
      Plan: 'APPLY'
    }),
    rowFromObject(headers, {
      '病歷號': '003',
      '姓名': '非手術',
      Tag: 'FU',
      Condition: 'Cataract s/p PE/IOL(C -0.5 22.0)',
      '日期': new Date(2026, 4, 9)
    }),
    rowFromObject(headers, {
      '病歷號': '004',
      '姓名': '隔日手術',
      Tag: 'OP',
      Condition: 'Cataract s/p PE/IOL(D -0.5 23.0)',
      '日期': new Date(2026, 4, 10)
    })
  ];

  const result = context.buildSurgeryExportDataForDate_(data, cols, targetDate);

  assert.strictEqual(result.dateText, '2026/05/09');
  assert.deepStrictEqual(toPlain(result.iolList), [
    ['陳小美', 'A', '-1.0', '21.5'],
    ['王小明', 'B', '-0.5', '20.0']
  ]);
  assert.deepStrictEqual(toPlain(result.patientList), [
    ['001', '王小明', '0911', '08:30', 'Cataract', 'PE/IOL(B -0.5 20.0)', '術後注意\n帶藥'],
    ['002', '陳小美', '0922', '09:30', 'Cataract', 'MSICS(A -1.0 21.5)', 'APPLY']
  ]);
}

function testUpcomingExportDatesIncludeSevenDays() {
  const dates = context.buildUpcomingExportDates_(new Date(2026, 4, 9, 15, 30), 7)
    .map(date => context.formatExportDate_(date));

  assert.deepStrictEqual(toPlain(dates), [
    '2026/05/09',
    '2026/05/10',
    '2026/05/11',
    '2026/05/12',
    '2026/05/13',
    '2026/05/14',
    '2026/05/15'
  ]);
}

testHeaderMappingWithCustomColumn();
testEventIdColumnLookupWithMovedHeader();
testEventIdColumnLookupMissingHeader();
testTimeNoteRoundTrip();
testMergeBlocksTrackedAbsorbedRows();
testMergeCombinesSafeTextFields();
testConditionalFormulaBuilding();
testConditionalFormattingKeepsTagIndependent();
testClearOldEventsUsesFixedOpEventIdColumn();
testClearOldEventsMissingEventIdHeaderAborts();
testSurgeryExportDataForDate();
testUpcomingExportDatesIncludeSevenDays();

console.log('Smoke tests passed');
