'use strict';

const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const files = [
  'code.js',
  'sheet_model.js',
  'calendar_sync.js',
  'workflows.js',
  'annual_archive.js'
];
const sourceByFile = Object.fromEntries(
  files.map(file => [
    file,
    fs.readFileSync(path.join(root, file), 'utf8')
  ])
);
const scriptPropertyStore = {};
const scriptPropertyCalls = {
  getProperty: 0,
  setProperty: 0,
  setProperties: 0,
  deleteProperty: 0,
  getProperties: 0
};

function digestBytes(value) {
  return Array.from(crypto.createHash('sha256').update(String(value)).digest())
    .map(value => value > 127 ? value - 256 : value);
}

const context = {
  console,
  Date,
  JSON,
  Math,
  Number,
  String,
  Boolean,
  Array,
  Object,
  RegExp,
  Error,
  Set,
  Map,
  Intl,
  Utilities: {
    DigestAlgorithm: { SHA_256: 'SHA_256' },
    Charset: { UTF_8: 'UTF_8' },
    computeDigest: (_algorithm, value) => digestBytes(value),
    formatDate: (date, _timezone, format) => {
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      const hour = String(date.getHours()).padStart(2, '0');
      const minute = String(date.getMinutes()).padStart(2, '0');
      const second = String(date.getSeconds()).padStart(2, '0');
      if (format === 'yyyy-MM-dd') return `${year}-${month}-${day}`;
      if (format === 'yyyyMM') return `${year}${month}`;
      if (format === 'yyyyMMdd') return `${year}${month}${day}`;
      if (format.includes('HH:mm:ss')) {
        return `${year}-${month}-${day}T${hour}:${minute}:${second}+08:00`;
      }
      return `${year}-${month}-${day}`;
    },
    getUuid: () => 'test-request-token'
  },
  Session: {
    getScriptTimeZone: () => 'Asia/Taipei'
  },
  SpreadsheetApp: {
    newDataValidation: () => {
      const state = {};
      const builder = {
        requireValueInList: (values, showDropdown) => {
          state.values = values.slice();
          state.showDropdown = showDropdown;
          return builder;
        },
        setAllowInvalid: value => {
          state.allowInvalid = value;
          return builder;
        },
        build: () => ({ ...state })
      };
      return builder;
    }
  },
  PropertiesService: {
    getScriptProperties: () => ({
      getProperty: key => {
        scriptPropertyCalls.getProperty++;
        return Object.hasOwn(scriptPropertyStore, key)
          ? scriptPropertyStore[key]
          : null;
      },
      setProperty: (key, value) => {
        scriptPropertyCalls.setProperty++;
        scriptPropertyStore[key] = String(value);
      },
      setProperties: values => {
        scriptPropertyCalls.setProperties++;
        Object.keys(values || {}).forEach(key => {
          scriptPropertyStore[key] = String(values[key]);
        });
      },
      deleteProperty: key => {
        scriptPropertyCalls.deleteProperty++;
        delete scriptPropertyStore[key];
      },
      getProperties: () => {
        scriptPropertyCalls.getProperties++;
        return { ...scriptPropertyStore };
      }
    })
  }
};
vm.createContext(context);
files.forEach(file => {
  vm.runInContext(sourceByFile[file], context, { filename: file });
});

function evaluate(expression) {
  return vm.runInContext(expression, context);
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function call(name, ...args) {
  context.__args = args;
  return vm.runInContext(`${name}(...__args)`, context);
}

class FakeRange {
  constructor(sheet, row, column, numRows = 1, numColumns = 1) {
    this.sheet = sheet;
    this.row = row;
    this.column = column;
    this.numRows = numRows;
    this.numColumns = numColumns;
  }

  getValues() {
    const result = [];
    for (let rowOffset = 0; rowOffset < this.numRows; rowOffset++) {
      const row = [];
      for (let columnOffset = 0; columnOffset < this.numColumns; columnOffset++) {
        row.push(
          this.sheet.valueAt(
            this.row + rowOffset,
            this.column + columnOffset
          )
        );
      }
      result.push(row);
    }
    return result;
  }

  getValue() {
    return this.getValues()[0][0];
  }

  getDisplayValues() {
    return this.getValues().map(row => {
      return row.map(value => value === null || value === undefined
        ? ''
        : String(value));
    });
  }

  getFontLines() {
    this.sheet.calls.getFontLines++;
    const result = [];
    for (let rowOffset = 0; rowOffset < this.numRows; rowOffset++) {
      const row = [];
      for (let columnOffset = 0; columnOffset < this.numColumns; columnOffset++) {
        row.push(
          this.sheet.fontLineAt(
            this.row + rowOffset,
            this.column + columnOffset
          )
        );
      }
      result.push(row);
    }
    return result;
  }

  getFormulas() {
    this.sheet.calls.getFormulas++;
    const result = [];
    for (let rowOffset = 0; rowOffset < this.numRows; rowOffset++) {
      const row = [];
      for (let columnOffset = 0; columnOffset < this.numColumns; columnOffset++) {
        row.push(
          this.sheet.formulaAt(
            this.row + rowOffset,
            this.column + columnOffset
          )
        );
      }
      result.push(row);
    }
    return result;
  }

  getFormulasR1C1() {
    return this.getFormulas();
  }

  getFormula() {
    return this.sheet.formulaAt(this.row, this.column);
  }

  setValue(value) {
    this.sheet.setFormulaAt(this.row, this.column, '');
    this.sheet.setValueAt(this.row, this.column, value);
    this.sheet.calls.setValues++;
    return this;
  }

  setValues(values) {
    for (let rowOffset = 0; rowOffset < this.numRows; rowOffset++) {
      for (let columnOffset = 0; columnOffset < this.numColumns; columnOffset++) {
        this.sheet.setFormulaAt(
          this.row + rowOffset,
          this.column + columnOffset,
          ''
        );
        this.sheet.setValueAt(
          this.row + rowOffset,
          this.column + columnOffset,
          (values[rowOffset] || [])[columnOffset] ?? ''
        );
      }
    }
    this.sheet.calls.setValues++;
    return this;
  }

  setFormula(formula) {
    this.sheet.setFormulaAt(this.row, this.column, formula);
    this.sheet.calls.setFormulas++;
    return this;
  }

  setFormulaR1C1(formula) {
    return this.setFormula(formula);
  }

  clearContent() {
    for (let rowOffset = 0; rowOffset < this.numRows; rowOffset++) {
      for (let columnOffset = 0; columnOffset < this.numColumns; columnOffset++) {
        this.sheet.setFormulaAt(
          this.row + rowOffset,
          this.column + columnOffset,
          ''
        );
        this.sheet.setValueAt(
          this.row + rowOffset,
          this.column + columnOffset,
          ''
        );
      }
    }
    return this;
  }

  getNote() {
    return this.sheet.noteAt(this.row, this.column);
  }

  getNotes() {
    this.sheet.calls.getNotes++;
    const result = [];
    for (let rowOffset = 0; rowOffset < this.numRows; rowOffset++) {
      const row = [];
      for (let columnOffset = 0; columnOffset < this.numColumns; columnOffset++) {
        row.push(
          this.sheet.noteAt(
            this.row + rowOffset,
            this.column + columnOffset
          )
        );
      }
      result.push(row);
    }
    return result;
  }

  setNote(value) {
    this.sheet.setNoteAt(this.row, this.column, value);
    return this;
  }

  setNotes(values) {
    for (let rowOffset = 0; rowOffset < this.numRows; rowOffset++) {
      for (let columnOffset = 0; columnOffset < this.numColumns; columnOffset++) {
        this.sheet.setNoteAt(
          this.row + rowOffset,
          this.column + columnOffset,
          (values[rowOffset] || [])[columnOffset] ?? ''
        );
      }
    }
    return this;
  }

  clearNote() {
    this.sheet.setNoteAt(this.row, this.column, '');
    return this;
  }

  setNumberFormat() {
    return this;
  }

  setDataValidation(value) {
    this.sheet.validations[`${this.row}:${this.column}`] = value;
    return this;
  }

  clearDataValidations() {
    delete this.sheet.validations[`${this.row}:${this.column}`];
    return this;
  }

  setFontFamily() {
    return this;
  }

  setFontSize() {
    return this;
  }

  setFontWeight() {
    return this;
  }

  setHorizontalAlignment() {
    return this;
  }

  setVerticalAlignment() {
    return this;
  }

  setWrap() {
    return this;
  }

  sort(spec) {
    this.sheet.sortRange(
      this.row,
      this.column,
      this.numRows,
      this.numColumns,
      spec
    );
    return this;
  }

  getRow() {
    return this.row;
  }

  getColumn() {
    return this.column;
  }

  getNumRows() {
    return this.numRows;
  }

  getNumColumns() {
    return this.numColumns;
  }

  getSheet() {
    return this.sheet;
  }
}

class FakeSheet {
  constructor(name, id, rows) {
    this.name = name;
    this.id = id;
    this.rows = rows.map(row => row.slice());
    this.notes = {};
    this.formulas = {};
    this.fontLines = {};
    this.validations = {};
    this.columnWidths = {};
    this.calls = {
      getRange: 0,
      setValues: 0,
      setFormulas: 0,
      sort: 0,
      insertColumnsAfter: 0,
      deleteColumns: 0,
      autoResizeRows: 0,
      getFormulas: 0,
      getFontLines: 0,
      getNotes: 0,
      setColumnWidth: [],
      rangeLists: [],
      numberFormatRangeLists: []
    };
  }

  getName() {
    return this.name;
  }

  getSheetId() {
    return this.id;
  }

  getLastColumn() {
    return Math.max(1, ...this.rows.map(row => row.length));
  }

  getLastRow() {
    let result = 0;
    this.rows.forEach((row, index) => {
      if (row.some(value => String(value ?? '').trim())) result = index + 1;
    });
    return Math.max(result, 1);
  }

  getMaxRows() {
    return Math.max(this.rows.length, 100);
  }

  getMaxColumns() {
    return this.getLastColumn();
  }

  getConditionalFormatRules() { return this.conditionalRules || []; }

  setConditionalFormatRules(rules) { this.conditionalRules = rules; }

  getRange(row, column, numRows, numColumns) {
    this.calls.getRange++;
    return new FakeRange(this, row, column, numRows, numColumns);
  }

  getRangeList(addresses) {
    this.calls.rangeLists.push(addresses.slice());
    const chain = {
      setFontFamily: () => chain,
      setFontSize: () => chain,
      setFontWeight: () => chain,
      setHorizontalAlignment: () => chain,
      setVerticalAlignment: () => chain,
      setWrap: () => chain,
      setNumberFormat: format => {
        this.calls.numberFormatRangeLists.push({
          addresses: addresses.slice(),
          format
        });
        return chain;
      }
    };
    return chain;
  }

  insertColumnsAfter(afterPosition, howMany) {
    this.calls.insertColumnsAfter++;
    this.rows.forEach(row => {
      while (row.length < afterPosition) row.push('');
      row.splice(afterPosition, 0, ...Array(howMany).fill(''));
    });
    return this;
  }

  deleteColumns(columnPosition, howMany) {
    this.calls.deleteColumns++;
    this.rows.forEach(row => {
      row.splice(columnPosition - 1, howMany);
    });
    return this;
  }

  autoResizeRows() {
    this.calls.autoResizeRows++;
    return this;
  }

  setColumnWidth(column, width) {
    this.columnWidths[column] = width;
    this.calls.setColumnWidth.push({ column, width });
    return this;
  }

  sortRange(row, column, numRows, numColumns, spec) {
    this.calls.sort++;
    const sortSpec = Array.isArray(spec) ? spec[0] : spec;
    const sortColumn = typeof sortSpec === 'number'
      ? sortSpec
      : sortSpec.column;
    const ascending = typeof sortSpec === 'number'
      ? true
      : sortSpec.ascending !== false;
    const records = [];
    for (let offset = 0; offset < numRows; offset++) {
      const sourceRow = row + offset;
      records.push({
        sourceRow,
        values: (this.rows[sourceRow - 1] || []).slice(),
        notes: Object.fromEntries(
          Object.entries(this.notes)
            .filter(([key]) => Number(key.split(':')[0]) === sourceRow)
            .map(([key, value]) => [Number(key.split(':')[1]), value])
        ),
        formulas: Object.fromEntries(
          Object.entries(this.formulas)
            .filter(([key]) => Number(key.split(':')[0]) === sourceRow)
            .map(([key, value]) => [Number(key.split(':')[1]), value])
        ),
        fontLines: Object.fromEntries(
          Object.entries(this.fontLines)
            .filter(([key]) => Number(key.split(':')[0]) === sourceRow)
            .map(([key, value]) => [Number(key.split(':')[1]), value])
        ),
        validations: Object.fromEntries(
          Object.entries(this.validations)
            .filter(([key]) => Number(key.split(':')[0]) === sourceRow)
            .map(([key, value]) => [Number(key.split(':')[1]), value])
        )
      });
    }
    records.sort((left, right) => {
      const leftValue = String(left.values[sortColumn - 1] ?? '');
      const rightValue = String(right.values[sortColumn - 1] ?? '');
      const compared = leftValue.localeCompare(rightValue);
      return ascending ? compared : -compared;
    });
    for (let offset = 0; offset < numRows; offset++) {
      const targetRow = row + offset;
      const record = records[offset];
      this.rows[targetRow - 1] = record.values.slice(
        column - 1,
        column - 1 + numColumns
      );
      Object.keys(this.notes).forEach(key => {
        if (Number(key.split(':')[0]) === targetRow) delete this.notes[key];
      });
      Object.keys(this.formulas).forEach(key => {
        if (Number(key.split(':')[0]) === targetRow) delete this.formulas[key];
      });
      Object.keys(this.fontLines).forEach(key => {
        if (Number(key.split(':')[0]) === targetRow) delete this.fontLines[key];
      });
      Object.keys(this.validations).forEach(key => {
        if (Number(key.split(':')[0]) === targetRow) {
          delete this.validations[key];
        }
      });
      Object.entries(record.notes).forEach(([targetColumn, value]) => {
        this.notes[`${targetRow}:${targetColumn}`] = value;
      });
      Object.entries(record.formulas).forEach(([targetColumn, value]) => {
        this.formulas[`${targetRow}:${targetColumn}`] = value;
      });
      Object.entries(record.fontLines).forEach(([targetColumn, value]) => {
        this.fontLines[`${targetRow}:${targetColumn}`] = value;
      });
      Object.entries(record.validations).forEach(([targetColumn, value]) => {
        this.validations[`${targetRow}:${targetColumn}`] = value;
      });
    }
  }

  valueAt(row, column) {
    return (this.rows[row - 1] || [])[column - 1] ?? '';
  }

  setValueAt(row, column, value) {
    while (this.rows.length < row) this.rows.push([]);
    while (this.rows[row - 1].length < column) this.rows[row - 1].push('');
    this.rows[row - 1][column - 1] = value;
  }

  formulaAt(row, column) {
    return this.formulas[`${row}:${column}`] || '';
  }

  setFormulaAt(row, column, formula) {
    const key = `${row}:${column}`;
    if (formula) this.formulas[key] = formula;
    else delete this.formulas[key];
  }

  fontLineAt(row, column) {
    return this.fontLines[`${row}:${column}`] || 'none';
  }

  setFontLineAt(row, column, value) {
    this.fontLines[`${row}:${column}`] = value;
  }

  noteAt(row, column) {
    return this.notes[`${row}:${column}`] || '';
  }

  setNoteAt(row, column, value) {
    this.notes[`${row}:${column}`] = value;
  }
}

function makeFakeSpreadsheet(sheets, id = 'spreadsheet-test') {
  return {
    getId: () => id,
    getSheets: () => sheets.slice(),
    getSheetByName: name => (
      sheets.find(sheet => sheet.getName() === name) || null
    )
  };
}

function testEightDigitDateInputOnlyInFunctions() {
  ['20260908', 20260908, '20240229'].forEach(value => assert.ok(call('parseEightDigitDate_', value)));
  ['20260229', '20261301', '20260931', '2026/9/8', '2026098'].forEach(value =>
    assert.strictEqual(call('parseEightDigitDate_', value), null));
  assert.ok(!call('processRowChange', { range: new FakeSheet('IVI', 818,
    [['姓名', '下次回診時間'], ['甲', '20260908']]).getRange(2, 2) }));
  assert.ok(!evaluate('processRowChange.toString()').includes('normalizeManagedDateRange_'));
}

function testRestoreDateFormatsPreservesValuesAndTimes() {
  const date = new Date(2026, 8, 8);
  const sheet = new FakeSheet('202609', 819, [['自訂', '時間'],
    ['保留', date], ['保留', '0830'], ['保留', 'PM']]);
  // Model a migrated date and unrelated time/custom formats.
  const formats = { '2:2': 'yyyyMMdd', '3:2': '@', '4:2': '@' };
  const originalGetRange = sheet.getRange.bind(sheet);
  sheet.getRange = (...args) => {
    const range = originalGetRange(...args);
    range.getNumberFormats = () => Array.from({length: range.getNumRows()}, (_, i) =>
      [formats[(range.getRow() + i) + ':' + range.getColumn()] || '']);
    range.setNumberFormat = format => {
      formats[range.getRow() + ':' + range.getColumn()] = format;
      return range;
    };
    return range;
  };
  const before = sheet.rows.map(row => row.slice());
  const result = call('restoreWorksheetDateFormatsInSpreadsheet_', makeFakeSpreadsheet([sheet]));
  assert.strictEqual(result[0].restored, 1);
  assert.strictEqual(formats['2:2'], 'yyyy/m/d ddd');
  assert.strictEqual(formats['3:2'], '@');
  assert.deepStrictEqual(sheet.rows, before);
  assert.strictEqual(call('restoreWorksheetDateFormatsInSpreadsheet_', makeFakeSpreadsheet([sheet])).length, 0);
}

function testPlanMigrationMovesWholeColumnAndIsIdempotent() {
  const sheet = new FakeSheet('202609', 821, [['自訂', 'Plan', 'IOL', 'Axis', 'CalendarEventId'],
    ['保留', '計畫', '+20.0', '90', 'event-keep']]);
  let moves = 0;
  sheet.moveColumns = (range, destination) => {
    moves++;
    const from = range.getColumn() - 1;
    const to = destination - 1 - (from < destination - 1 ? 1 : 0);
    sheet.rows.forEach(row => row.splice(to, 0, row.splice(from, 1)[0]));
  };
  assert.strictEqual(call('moveMonthlyPlanAfterAxis_', sheet).moved, true);
  assert.deepStrictEqual(sheet.rows, [['自訂', 'IOL', 'Axis', 'Plan', 'CalendarEventId'],
    ['保留', '+20.0', '90', '計畫', 'event-keep']]);
  assert.strictEqual(call('moveMonthlyPlanAfterAxis_', sheet).moved, false);
  assert.strictEqual(moves, 1);
  const conflict = new FakeSheet('202608', 822, [['Plan', 'Axis', 'Plan'], ['A', 'B', 'C']]);
  assert.strictEqual(call('moveMonthlyPlanAfterAxis_', conflict).ok, false);
  assert.deepStrictEqual(conflict.rows[1], ['A', 'B', 'C']);
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  assert.strictEqual(headers.indexOf('Plan'), headers.indexOf('Axis') + 1);
}

function testIolListIsReadOnlyAndKeepsAllDateBlocks() {
  const { sheet, columns, api } = makeMonthlyReorderFixture();
  sheet.setValueAt(3, columns.IOL, 'Lens <A>');
  sheet.setValueAt(3, columns.IOL_FINAL, '+20.0');
  sheet.setFontLineAt(4, columns.CHART_NO, 'line-through');
  const header = sheet.rows[1].slice();
  header[columns.HOSPITAL - 1] = '聯醫';
  const patient = sheet.rows[2].slice();
  patient[columns.EVENT_ID - 1] = '';
  patient[columns.NAME - 1] = '第三位';
  sheet.rows.push(header, patient);
  const before = JSON.stringify(sheet.rows);
  const result = call('buildIolListForDate_', sheet, new Date(2026, 8, 10), false);
  assert.strictEqual(result.count, 2);
  assert.strictEqual(result.excluded, 1);
  assert.ok(result.text.includes('測試甲｜Lens <A>｜+20.0'));
  assert.ok(result.text.includes('第三位｜Lens <A>｜+20.0'));
  assert.ok(result.text.includes('20260910'));
  const included = call('buildIolListForDate_', sheet, new Date(2026, 8, 10), true);
  assert.strictEqual(included.count, 3);
  assert.ok(included.text.includes('測試乙（已取消）｜未填｜未填'));
  assert.strictEqual(JSON.stringify(sheet.rows), before);
  assert.strictEqual(api.counts().updateCount, 0);
  assert.strictEqual(api.counts().insertCount, 0);
}

function testIolListReadsOnlyRequestedDateDisplayRows() {
  const { sheet, columns } = makeMonthlyReorderFixture();
  const header = sheet.rows[1].slice();
  header[columns.TIME - 1] = new Date(2026, 8, 11);
  const patient = sheet.rows[2].slice();
  patient[columns.EVENT_ID - 1] = '';
  sheet.rows.push(header, patient);
  const reads = [];
  const getRange = sheet.getRange.bind(sheet);
  sheet.getRange = (...args) => {
    const range = getRange(...args);
    const display = range.getDisplayValues.bind(range);
    range.getDisplayValues = () => { reads.push(args); return display(); };
    return range;
  };
  const result = call('buildIolListForDate_', sheet, new Date(2026, 8, 10), true);
  assert.strictEqual(result.count, 2);
  assert.deepStrictEqual(reads, [[3, 1, 2,
    Math.max(columns.CHART_NO, columns.NAME, columns.IOL, columns.IOL_FINAL)]]);
  reads.length = 0;
  const empty = call('buildIolListForDate_', sheet, new Date(2026, 8, 12), false);
  assert.strictEqual(empty.count, 0);
  assert.deepStrictEqual(reads, []);
}

function testEntropionCaseNormalizationPreservesOtherText() {
  assert.strictEqual(call('normalizeEntropionCase_', 'CATA + ENTROPION OU'), 'CATA + Entropion OU');
  assert.strictEqual(call('normalizeEntropionCase_', 'Entropion'), 'Entropion');
  assert.strictEqual(call('normalizeEntropionCase_', 'ENTROPIONX'), 'ENTROPIONX');
  assert.strictEqual(call('normalizeEntropionCase_', 123), 123);
}

function testPlanMigrationRepairsNativeRefErrorsAndPreservesOtherRules() {
  const sheet = new FakeSheet('202609', 824, [['Axis', 'Plan'], ['90', 'APPLY']]);
  const makeRule = formula => ({
    formula, style: 'green',
    getBooleanCondition: () => ({ getCriteriaValues: () => [formula] }),
    getRanges: () => [sheet.getRange(43, 1, 1, 2)],
    copy: () => ({ whenFormulaSatisfied: updated => ({ build: () => makeRule(updated) }) })
  });
  const custom = makeRule('=ISNUMBER(SEARCH("!",#REF!))');
  sheet.conditionalRules = [custom, makeRule(
    '=AND(ISNUMBER(SEARCH("APPLY",#REF!)),N("SURGERY_SYSTEM_CF_MONTH_PLAN_GREEN")=0)')];
  const result = call('moveMonthlyPlanAfterAxis_', sheet);
  assert.strictEqual(result.moved, false);
  assert.strictEqual(result.repairedRules, 1);
  assert.strictEqual(sheet.conditionalRules[0], custom);
  assert.ok(sheet.conditionalRules[1].formula.includes('SEARCH("APPLY",$B43)'));
  assert.strictEqual(sheet.conditionalRules[1].style, 'green');
  assert.strictEqual(call('moveMonthlyPlanAfterAxis_', sheet).repairedRules, 0);
}

function testVersionAndModuleSplit() {
  assert.strictEqual(evaluate('CONFIG.VERSION'), '2026.09.12.2');
  assert.strictEqual(evaluate('typeof processRowChange'), 'function');
  assert.strictEqual(evaluate('typeof processCalendarStructureChange'), 'function');
  assert.strictEqual(evaluate('typeof createMonthlySurgerySheet'), 'function');
  assert.strictEqual(
    evaluate('typeof migrateMonthlyDiagnosisSummaryHeaders'),
    'function'
  );
  assert.strictEqual(evaluate('typeof rebuildCalendarRegistry_'), 'function');
  assert.ok(sourceByFile['code.js'].includes('sheet_model.js'));
  assert.ok(sourceByFile['code.js'].includes('calendar_sync.js'));
  assert.ok(sourceByFile['code.js'].includes('workflows.js'));
}

function testOnlyCalendarEventIdIsCanonicalSystemField() {
  assert.deepStrictEqual(plain(evaluate('CONFIG.SYSTEM_FIELD_KEYS')), ['EVENT_ID']);
  assert.deepStrictEqual(
    plain(evaluate('CONFIG.MONTHLY_SYSTEM_FIELD_KEYS')),
    ['EVENT_ID']
  );
  assert.strictEqual(
    evaluate('CONFIG.FIELD_HEADERS.EVENT_ID'),
    'CalendarEventId'
  );
  assert.strictEqual(
    plain(evaluate('CONFIG.FIELD_KEYS')).includes('TIME'),
    false
  );
  assert.strictEqual(
    plain(evaluate('CONFIG.HEADERS')).includes('時間'),
    false
  );
  const deployedSource = files.map(file => sourceByFile[file]).join('\n');
  [
    'CalendarSheetWriteUpdated',
    'CalendarSyncState',
    'FUTrackingId',
    'SurgeryTrackingId'
  ].forEach(header => {
    assert.strictEqual(deployedSource.includes(header), false);
  });
}

function testCurrentMenuHasNoCompletedMigrationOrLegacyOutput() {
  const source = sourceByFile['code.js'];
  [
    '完整安裝／批次同步',
    'FU Tag／同步狀態升級',
    '系統識別欄退役',
    'OP-高榮',
    'OP-聯醫',
    '手術清單',
    '反向同步',
    '預覽同步架構升級',
    '執行同步架構升級'
  ].forEach(text => assert.strictEqual(source.includes(text), false));
  [
    '安裝／修復系統',
    '檢查同步健康',
    '同步待處理變更',
    '預覽 FU 追蹤生命週期修復',
    '執行 FU 追蹤生命週期修復',
    '套用所有 FU／月表建議欄寬',
    '啟用／修復月表診斷統計表頭',
    '月刀表 => FU',
    'FU => 月刀表',
    '建立新月刀表',
    '水晶體清單（選刀日／複製）',
    '調整月表 Plan 至 Axis 後方',
    '恢復工作表原日期顯示格式',
    '選取列新增刀日',
    '整理目前分頁',
    '資料遷移與舊版修復',
    '修復選取列同步',
    '彙整舊月刀表'
  ].forEach(text => assert.ok(source.includes(text)));
  [
    '重建同步索引',
    '套用標準格式',
    '彙整並封存月份刀表',
    '解除月份封存',
    '建立月份刀表',
    '在選取列新增刀日',
    '整理目前月份刀表',
    '整理 FU 日期',
    '將月表資料列加入 FU',
    '從 FU 安排至月份刀表',
    '永久彙整舊月份',
    '預覽 FU 移除時間欄',
    '執行 FU 移除時間欄',
    '預覽月份日期標題升級',
    '執行月份日期標題升級',
    '預覽舊備份分頁與空白列清理',
    '執行舊備份分頁與空白列清理'
  ].forEach(text => assert.strictEqual(source.includes(text), false));
  assert.strictEqual(
    evaluate('typeof previewLegacyBackupAndGridCleanup'),
    'undefined'
  );
  assert.strictEqual(
    evaluate('typeof executeLegacyBackupAndGridCleanup'),
    'undefined'
  );
  assert.strictEqual(
    evaluate('typeof previewFuLifecycleRecoveryMigration'),
    'function'
  );
  assert.strictEqual(
    evaluate('typeof executeFuLifecycleRecoveryMigration'),
    'function'
  );
  assert.strictEqual(
    evaluate('typeof removeLegacySyncHealthReportSheet'),
    'undefined'
  );
}

function testMonthlySheetNameRecognition() {
  assert.strictEqual(call('isMonthlySheetName_', '202608'), true);
  assert.strictEqual(call('isMonthlySheetName_', '202613'), false);
  assert.strictEqual(call('isMonthlySheetName_', '刀表彙整_202604-202606'), false);
  assert.strictEqual(call('isMonthlySheetName_', '20268'), false);
}

function testTimeNormalizationVariants() {
  assert.strictEqual(call('parseTimeInput_', '0800').text, '08:00');
  assert.strictEqual(call('parseTimeInput_', 800).text, '08:00');
  assert.strictEqual(call('parseTimeInput_', 830).text, '08:30');
  assert.strictEqual(call('parseTimeInput_', '8:00').text, '08:00');
  assert.strictEqual(call('parseTimeInput_', '23:59').text, '23:59');
  assert.strictEqual(call('parseTimeInput_', 0.5).text, '12:00');
  assert.strictEqual(
    call('parseTimeInput_', new Date(1899, 11, 30, 8, 30)).text,
    '08:30'
  );
  assert.strictEqual(call('resolveCalendarTime_', '').parsedTime, null);
  assert.strictEqual(call('resolveCalendarTime_', 'PM').timeNote, 'PM');
  assert.ok(call('resolveCalendarTime_', '2460').errorMessage.includes('時間範圍錯誤'));
  assert.ok(call('resolveCalendarTime_', '25:00').errorMessage.includes('時間範圍錯誤'));
}

function testMultiRowPasteNormalizesEveryMonthlyTimeCell() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const timeColumn = headers.indexOf('日期／報到時間') + 1;
  const rows = [headers];
  ['0800', 830].forEach(value => {
    const row = Array(headers.length).fill('');
    row[timeColumn - 1] = value;
    rows.push(row);
  });
  const sheet = new FakeSheet('202608', 11, rows);
  const columns = call('getRequiredMonthlyColumns_', sheet);
  sheet.calls.getRange = 0;
  sheet.calls.setValues = 0;
  const result = call(
    'normalizeEditedTimes_',
    {
      range: sheet.getRange(2, timeColumn, 2, 1)
    },
    sheet,
    'MONTHLY',
    columns
  );
  assert.strictEqual(result.normalized, 2);
  assert.strictEqual(sheet.valueAt(2, timeColumn), '08:00');
  assert.strictEqual(sheet.valueAt(3, timeColumn), '08:30');
  assert.strictEqual(sheet.calls.setValues, 1);
  assert.ok(sheet.calls.getRange <= 5);
}

function testCalendarPendingQueueCoalescesWithoutClinicalText() {
  const queue = call('emptyPendingCalendarQueue_');
  call('mergePendingCalendarRequest_', queue, {
    sheetId: 88,
    rows: [9, 4, 9],
    reason: 'EDIT'
  });
  call('mergePendingCalendarRequest_', queue, {
    sheetId: 88,
    rows: [5, 4],
    reason: '虛構病人資料'
  });
  assert.deepStrictEqual(plain(queue.sheets['88']), [4, 5, 9]);
  assert.deepStrictEqual(plain(queue.reasons), ['EDIT', 'UNSPECIFIED']);
  assert.strictEqual(JSON.stringify(queue).includes('虛構病人資料'), false);
  assert.deepStrictEqual(
    plain(call('summarizePendingCalendarQueue_', queue)),
    {
      hasWork: true,
      fullScan: false,
      sheetCount: 1,
      rowCount: 3
    }
  );

  call('mergePendingCalendarRequest_', queue, {
    fullScan: true,
    allowMissingDeletes: true,
    reason: 'CHANGE_REMOVE_ROW'
  });
  assert.strictEqual(queue.fullScan, true);
  assert.strictEqual(queue.allowMissingDeletes, true);
  assert.deepStrictEqual(plain(queue.sheets), {});
}

function testWholeMonthlyTimeNormalizationUsesBoundedBatchReads() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const time = headers.indexOf('日期／報到時間');
  const rows = [headers];
  for (let i = 0; i < 100; i++) {
    const row = headers.map(() => '');
    row[time] = 830;
    rows.push(row);
  }
  const sheet = new FakeSheet('202610', 870, rows);
  const columns = call('getRequiredMonthlyColumns_', sheet);
  sheet.calls.getRange = 0;
  sheet.calls.setValues = 0;
  const result = call('normalizeExistingTimeColumn_', sheet, 'MONTHLY', columns);
  assert.deepStrictEqual(plain(result), { normalized: 100, errors: 0 });
  assert.ok(sheet.calls.getRange <= 5, '100 rows must use bounded batch reads');
  assert.strictEqual(sheet.calls.setValues, 1);
  assert.ok(sheet.rows.slice(1).every(row => row[time] === '08:30'));
  sheet.calls.setValues = 0;
  assert.strictEqual(call('normalizeExistingTimeColumn_', sheet, 'MONTHLY', columns).normalized, 0);
  assert.strictEqual(sheet.calls.setValues, 0);
}

function testFuSnapshotContextDoesNotRereadSheet() {
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const row = headers.map(header => ({ '病歷號': '001', '姓名': '測試',
    '日期': new Date(2026, 9, 1), CalendarEventId: 'event-test' }[header] || ''));
  const sheet = new FakeSheet('FU', 871, [headers, row]);
  const columns = call('getRequiredFuColumns_', sheet);
  const fresh = call('buildFuRowContext_', sheet, 2, columns);
  sheet.calls.getRange = 0;
  for (let i = 0; i < 100; i++) {
    const result = call('buildFuRowContext_', sheet, 2, columns, row);
    assert.strictEqual(result.rowHash, fresh.rowHash);
    assert.strictEqual(result.valid, fresh.valid);
  }
  assert.strictEqual(sheet.calls.getRange, 0);
}

function testUnifiedSortRoutesWithoutChangingSortImplementations() {
  const oldFu = context.sortFuByDate;
  const oldMonthly = context.sortCurrentMonthlySheet;
  const oldActive = context.SpreadsheetApp.getActiveSpreadsheet;
  let name = 'FU';
  try {
    context.sortFuByDate = () => 'fu-sort';
    context.sortCurrentMonthlySheet = () => 'monthly-sort';
    context.SpreadsheetApp.getActiveSpreadsheet = () => ({
      getActiveSheet: () => ({ getName: () => name })
    });
    assert.strictEqual(call('sortCurrentScheduleSheet'), 'fu-sort');
    name = '202610';
    assert.strictEqual(call('sortCurrentScheduleSheet'), 'monthly-sort');
  } finally {
    context.sortFuByDate = oldFu;
    context.sortCurrentMonthlySheet = oldMonthly;
    context.SpreadsheetApp.getActiveSpreadsheet = oldActive;
  }
}

function withTriggerRuntimeMocks(options, callback) {
  const settings = options || {};
  const previousScriptApp = context.ScriptApp;
  const previousLockService = context.LockService;
  const triggers = [];
  const createdDelays = [];
  const makeLock = acquired => ({
    tryLock: () => acquired,
    releaseLock: () => {}
  });
  context.LockService = {
    getDocumentLock: () => makeLock(
      settings.documentLockAcquired !== false
    ),
    getScriptLock: () => makeLock(
      settings.scriptLockAcquired !== false
    )
  };
  context.ScriptApp = {
    getProjectTriggers: () => triggers.slice(),
    deleteTrigger: trigger => {
      const index = triggers.indexOf(trigger);
      if (index >= 0) triggers.splice(index, 1);
    },
    newTrigger: handler => {
      let delay = 0;
      const builder = {
        timeBased: () => builder,
        after: value => {
          delay = value;
          return builder;
        },
        create: () => {
          const trigger = {
            getHandlerFunction: () => handler
          };
          triggers.push(trigger);
          createdDelays.push(delay);
          return trigger;
        }
      };
      return builder;
    }
  };
  try {
    return callback({ triggers, createdDelays });
  } finally {
    if (previousScriptApp === undefined) delete context.ScriptApp;
    else context.ScriptApp = previousScriptApp;
    if (previousLockService === undefined) delete context.LockService;
    else context.LockService = previousLockService;
  }
}

function testFuNoDateDraftSkipsSyncAndClearsLegacyBusyNote() {
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const draft = Array(headers.length).fill('');
  draft[headers.indexOf('病歷號')] = 'X-DRAFT';
  const active = Array(headers.length).fill('');
  active[headers.indexOf('姓名')] = '測試姓名';
  active[headers.indexOf('日期')] = new Date(2026, 7, 10);
  const sheet = new FakeSheet('FU', 70, [headers, draft, active]);
  const columns = call('getRequiredFuColumns_', sheet);
  const noteCell = sheet.getRange(2, columns.CHART_NO);
  noteCell.setNote('人工註記');
  call(
    'setSystemNote_',
    noteCell,
    evaluate('CALENDAR_SYNC_NOTE_PREFIX'),
    '自動同步忙碌，已排入重試。'
  );

  const filtered = call(
    'filterFuTriggerRowsForCalendar_',
    sheet,
    [2, 3],
    columns
  );
  assert.deepStrictEqual(plain(filtered.rows), [3]);
  assert.deepStrictEqual(plain(filtered.skippedRows), [2]);
  assert.strictEqual(noteCell.getNote(), '人工註記');
}

function testIncompleteFuLifecycleSkipsOrDeletesOwnEvent() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const api = installCalendarMock({
    initialEvents: [{
      id: 'event-to-delete',
      summary: 'X001 |  | 追蹤',
      description: '',
      start: { date: '2026-08-10' },
      end: { date: '2026-08-11' }
    }]
  });
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const draft = Array(headers.length).fill('');
  draft[headers.indexOf('病歷號')] = 'X-DRAFT';
  const existing = Array(headers.length).fill('');
  existing[headers.indexOf('病歷號')] = 'X001';
  existing[headers.indexOf('CalendarEventId')] = 'event-to-delete';
  const sheet = new FakeSheet('FU', 71, [headers, draft, existing]);
  const columns = call('getRequiredFuColumns_', sheet);

  const skipped = call(
    'syncManagedContext_',
    call('buildFuRowContext_', sheet, 2, columns),
    {}
  );
  assert.strictEqual(skipped.ok, true);
  assert.strictEqual(skipped.status, 'skipped_incomplete_draft');
  assert.strictEqual(api.counts().insertCount, 0);
  assert.strictEqual(api.counts().removeCount, 0);

  const deleted = call(
    'syncManagedContext_',
    call('buildFuRowContext_', sheet, 3, columns),
    {}
  );
  assert.strictEqual(deleted.ok, true);
  assert.strictEqual(deleted.status, 'deleted');
  assert.strictEqual(api.counts().removeCount, 1);
  assert.strictEqual(
    sheet.valueAt(3, headers.indexOf('CalendarEventId') + 1),
    ''
  );
  assert.deepStrictEqual(
    plain(call(
      'getRegistryRetentionForSync_',
      { eventId: 'event-to-delete' },
      deleted
    ).allowedRemovedEventIds),
    ['event-to-delete']
  );
}

function testBusyTriggerQueuesOnceWithoutWritingBusyNote() {
  clearScriptProperties();
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const row = Array(headers.length).fill('');
  row[headers.indexOf('姓名')] = '測試姓名';
  row[headers.indexOf('日期')] = new Date(2026, 7, 10);
  const sheet = new FakeSheet('FU', 72, [headers, row]);
  const spreadsheet = {
    getId: () => 'spreadsheet-test'
  };

  withTriggerRuntimeMocks(
    { documentLockAcquired: false },
    runtime => {
      let callbackCount = 0;
      const first = call(
        'runTriggerCalendarWork_',
        spreadsheet,
        { sheetId: 72, rows: [2], reason: 'EDIT' },
        () => {
          callbackCount++;
          return { ok: true };
        },
        { sheet, rows: [2] }
      );
      const second = call(
        'runTriggerCalendarWork_',
        spreadsheet,
        { sheetId: 72, rows: [2, 2], reason: 'EDIT' },
        () => {
          callbackCount++;
          return { ok: true };
        },
        { sheet, rows: [2] }
      );
      assert.strictEqual(first.reason, 'calendar_sync_busy');
      assert.strictEqual(second.reason, 'calendar_sync_busy');
      assert.strictEqual(callbackCount, 0);
      assert.deepStrictEqual(
        plain(call('readPendingCalendarQueue_').sheets['72']),
        [2]
      );
      assert.strictEqual(runtime.triggers.length, 1);
      assert.strictEqual(
        runtime.createdDelays[0],
        evaluate('CALENDAR_PENDING_RETRY_DELAY_MS')
      );
      assert.strictEqual(sheet.getRange(2, 2).getNote(), '');
    }
  );
}

function testPendingRetryWorkerDrainsQueueOnce() {
  clearScriptProperties();
  const previousOpenById = context.SpreadsheetApp.openById;
  const previousApply = context.applyPendingCalendarQueue_;
  const spreadsheet = { getId: () => 'spreadsheet-test' };
  context.SpreadsheetApp.openById = () => spreadsheet;
  try {
    withTriggerRuntimeMocks({}, runtime => {
      call('enqueuePendingCalendarSync_', {
        fullScan: true,
        reason: 'CHANGE_INSERT_ROW'
      });
      call('ensurePendingCalendarRetryScheduled_', spreadsheet);
      assert.strictEqual(runtime.triggers.length, 1);
      context.applyPendingCalendarQueue_ = () => ({
        ok: true,
        empty: false,
        results: []
      });
      const result = call('retryPendingCalendarSync');
      assert.strictEqual(result.ok, true);
      assert.strictEqual(runtime.triggers.length, 0);
      assert.strictEqual(
        call(
          'pendingCalendarQueueHasWork_',
          call('readPendingCalendarQueue_')
        ),
        false
      );
      assert.strictEqual(
        scriptPropertyStore.CALENDAR_SYNC_PENDING_SPREADSHEET_ID_V1,
        undefined
      );
    });
  } finally {
    context.applyPendingCalendarQueue_ = previousApply;
    if (previousOpenById === undefined) {
      delete context.SpreadsheetApp.openById;
    } else {
      context.SpreadsheetApp.openById = previousOpenById;
    }
  }
}

function testOnEditNormalizesBeforeShortLockAndAvoidsFullReconcile() {
  const source = sourceByFile['calendar_sync.js'];
  const body = source.slice(
    source.indexOf('function processRowChange'),
    source.indexOf('function processCalendarStructureChange')
  );
  assert.ok(
    body.indexOf('normalizeEditedTimes_') <
      body.indexOf('runTriggerCalendarWork_')
  );
  assert.strictEqual(body.includes('reconcileCalendarRegistry_'), false);
  assert.strictEqual(body.includes('listConfiguredCalendarEvents_'), false);
  assert.ok(body.includes('applyTouchedManagedRowsFormatBatch_'));
  assert.strictEqual(body.includes('applyManagedRowsFormatBatch_'), false);
  assert.ok(body.includes('headerPresentationTouched'));
  assert.ok(body.includes('wrapTouched'));
  assert.ok(source.includes("changeType === 'EDIT'"));
  assert.ok(source.includes("status: 'handled_by_on_edit'"));
  assert.ok(source.includes('.timeBased()'));
  assert.ok(source.includes('.after(CALENDAR_PENDING_RETRY_DELAY_MS)'));
  assert.strictEqual(source.includes('自動同步忙碌'), false);
}

function testFullDateDoesNotCollideWithTime() {
  assert.strictEqual(
    call('formatDateKey_', call('parseFullDate_', '2026/8/10')),
    '2026-08-10'
  );
  assert.strictEqual(call('parseFullDate_', 800), null);
  assert.strictEqual(
    call('parseFullDate_', new Date(1899, 11, 30, 8, 0)),
    null
  );
  assert.strictEqual(
    call('formatDateKey_', call('parseFullDate_', '20260230')),
    ''
  );
  assert.strictEqual(
    call('formatDateKey_', call('parseFullDate_', '2026/8/13 Thu')),
    '2026-08-13'
  );
  assert.strictEqual(call('parseFullDate_', '2026/8/13 Mon'), null);
}

function testMonthlyDateHeaderUsesDateAndBlankEventId() {
  const columns = {
    TIME: 1,
    HOSPITAL: 2,
    CHART_NO: 3,
    NAME: 4,
    TEL: 5,
    GA: 6,
    SIDE: 7,
    DIAGNOSIS: 8,
    GRADE: 9,
    PROCEDURE: 10,
    PLAN: 11,
    IOL: 12,
    IOL_TARGET: 13,
    IOL_FINAL: 14,
    AXIS: 15,
    MEMO: 16,
    SN: 18,
    CDE: 19,
    ENERGY_TIME: 20,
    ENERGY_PERCENT: 21,
    REFRACTION: 22,
    EVENT_ID: 23
  };
  const header = Array(23).fill('');
  header[0] = new Date(2026, 7, 10);
  assert.strictEqual(
    call('classifyMonthlyRowValues_', header, columns).type,
    'DATE_HEADER'
  );
  header[3] = '虛構姓名';
  assert.strictEqual(
    call('classifyMonthlyRowValues_', header, columns).type,
    'DATE_HEADER'
  );
  header[22] = 'event-conflict';
  assert.strictEqual(
    call('classifyMonthlyRowValues_', header, columns).type,
    'HYBRID_CONFLICT'
  );
  header[22] = '';
  header[0] = new Date(1899, 11, 30, 8, 0);
  assert.strictEqual(
    call('classifyMonthlyRowValues_', header, columns).type,
    'PATIENT'
  );
}

function testMonthlyHeaderMarkerPreservesOtherCells() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const row = Array(headers.length).fill('');
  row[columns.TIME - 1] = new Date(2026, 7, 10);
  row[columns.NAME - 1] = '刀日人工備註';
  const sheet = new FakeSheet('202608', 202608, [headers, row]);
  const result = call(
    'ensureMonthlyHeaderPresentation_',
    sheet,
    2,
    columns
  );
  assert.strictEqual(result.ok, true);
  assert.strictEqual(
    sheet.valueAt(2, columns.CHART_NO),
    evaluate('MONTHLY_DATE_HEADER_MARKER')
  );
  assert.strictEqual(sheet.valueAt(2, columns.HOSPITAL), '高榮');
  assert.strictEqual(sheet.valueAt(2, columns.NAME), '刀日人工備註');
  assert.strictEqual(
    sheet.validations[`2:${columns.CHART_NO}`].allowInvalid,
    false
  );

  sheet.setValueAt(2, columns.TIME, '08:00');
  call('ensureMonthlyHeaderPresentation_', sheet, 2, columns);
  assert.strictEqual(sheet.valueAt(2, columns.CHART_NO), '');
  assert.strictEqual(sheet.valueAt(2, columns.NAME), '刀日人工備註');
}

function testMonthlyHeaderMarkerNeverOverwritesExistingChartCell() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const row = Array(headers.length).fill('');
  row[columns.TIME - 1] = new Date(2026, 7, 10);
  row[columns.CHART_NO - 1] = '人工內容';
  const sheet = new FakeSheet('202608', 202609, [headers, row]);
  const result = call(
    'ensureMonthlyHeaderPresentation_',
    sheet,
    2,
    columns
  );
  assert.strictEqual(result.ok, false);
  assert.strictEqual(sheet.valueAt(2, columns.CHART_NO), '人工內容');
  assert.ok(
    sheet.getRange(2, columns.CHART_NO).getNote().includes('未被覆寫')
  );
}

function testMonthlyStructureIssuesReportHybridAndMissingMarker() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const dateHeader = Array(headers.length).fill('');
  dateHeader[columns.TIME - 1] = new Date(2026, 7, 10);
  dateHeader[columns.NAME - 1] = '刀日備註';
  const hybrid = Array(headers.length).fill('');
  hybrid[columns.TIME - 1] = new Date(2026, 7, 13);
  hybrid[columns.CHART_NO - 1] = 'X-HYBRID';
  hybrid[columns.EVENT_ID - 1] = 'event-hybrid';
  const sheet = new FakeSheet('202608', 202610, [
    headers,
    dateHeader,
    hybrid
  ]);
  const scan = call('scanMonthlyBlocks_', sheet, columns);
  const issues = call('getMonthlyStructureIssues_', sheet, scan);
  assert.deepStrictEqual(
    plain(issues.map(item => item.type).sort()),
    ['monthly_header_marker_missing', 'monthly_hybrid_date_event_id']
  );
  assert.strictEqual(scan.blocks.length, 1);
  assert.strictEqual(scan.hybrids.length, 1);
}

function testDuplicateMonthlyBlocksAreIndependentAndSupported() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const marker = evaluate('MONTHLY_DATE_HEADER_MARKER');
  const header1 = Array(headers.length).fill('');
  header1[columns.TIME - 1] = new Date(2026, 7, 10);
  header1[columns.HOSPITAL - 1] = '高榮';
  header1[columns.CHART_NO - 1] = marker;
  header1[columns.NAME - 1] = '第一區塊備註';
  const patient1 = Array(headers.length).fill('');
  patient1[columns.TIME - 1] = '08:00';
  patient1[columns.CHART_NO - 1] = 'TEST-001';
  const header2 = Array(headers.length).fill('');
  header2[columns.TIME - 1] = new Date(2026, 7, 10);
  header2[columns.HOSPITAL - 1] = '高榮';
  header2[columns.CHART_NO - 1] = marker;
  header2[columns.NAME - 1] = '第二區塊備註';
  const patient2 = Array(headers.length).fill('');
  patient2[columns.TIME - 1] = '09:00';
  patient2[columns.CHART_NO - 1] = 'TEST-002';
  const sheet = new FakeSheet('202608', 202611, [
    headers,
    header1,
    patient1,
    header2,
    patient2
  ]);
  const scan = call('scanMonthlyBlocks_', sheet, columns);
  assert.strictEqual(scan.blocks.length, 2);
  assert.strictEqual(scan.duplicateBlocks.length, 1);
  assert.strictEqual(
    plain(call('getMonthlyStructureIssues_', sheet, scan))
      .some(issue => issue.type === 'monthly_duplicate_block'),
    false
  );
  const options = plain(call('getMonthlyBlockOptions_', sheet));
  assert.deepStrictEqual(
    options.map(item => item.headerRow),
    [2, 4]
  );
  assert.ok(options[0].label.includes('第 2 列'));
  assert.ok(options[1].label.includes('第 4 列'));

  const byRow = {
    2: { values: ['header-1'] },
    3: { values: ['patient-1'] },
    4: { values: ['header-2'] },
    5: { values: ['patient-2'] }
  };
  const groups = call('buildMonthlySnapshotGroups_', scan, byRow);
  groups.sort((left, right) => call(
    'compareMonthlySnapshotGroups_',
    left,
    right
  ));
  assert.strictEqual(groups.length, 2);
  assert.deepStrictEqual(
    plain(groups.map(group => group.header.values[0])),
    ['header-1', 'header-2']
  );
  assert.deepStrictEqual(
    plain(groups.map(group => group.patients[0].values[0])),
    ['patient-1', 'patient-2']
  );
}

function testOnEditScopesCalendarRowsToTouchedPatientOrHeaderBlock() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const marker = evaluate('MONTHLY_DATE_HEADER_MARKER');
  const header1 = Array(headers.length).fill('');
  header1[columns.TIME - 1] = new Date(2026, 7, 10);
  header1[columns.HOSPITAL - 1] = '高榮';
  header1[columns.CHART_NO - 1] = marker;
  const patient1 = Array(headers.length).fill('');
  patient1[columns.TIME - 1] = '08:00';
  patient1[columns.CHART_NO - 1] = 'TEST-201';
  const patient2 = Array(headers.length).fill('');
  patient2[columns.TIME - 1] = '09:00';
  patient2[columns.NAME - 1] = '虛構病人甲';
  const header2 = Array(headers.length).fill('');
  header2[columns.TIME - 1] = new Date(2026, 7, 24);
  header2[columns.HOSPITAL - 1] = '高榮';
  header2[columns.CHART_NO - 1] = marker;
  const patient3 = Array(headers.length).fill('');
  patient3[columns.TIME - 1] = '10:00';
  patient3[columns.NAME - 1] = '虛構病人乙';
  const sheet = new FakeSheet('202608', 202614, [
    headers,
    header1,
    patient1,
    patient2,
    header2,
    patient3
  ]);

  const patientRows = call(
    'getAffectedRowsForEdit_',
    { range: sheet.getRange(4, columns.TIME, 1, 1) },
    sheet,
    'MONTHLY',
    columns
  );
  assert.deepStrictEqual(plain(patientRows), [4]);

  const headerRows = call(
    'getAffectedRowsForEdit_',
    { range: sheet.getRange(2, columns.HOSPITAL, 1, 1) },
    sheet,
    'MONTHLY',
    columns
  );
  assert.deepStrictEqual(plain(headerRows), [3, 4]);
}

function testSurgeryDateInsertionPreviewSplitsOnlyRowsBelowSelection() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const marker = evaluate('MONTHLY_DATE_HEADER_MARKER');
  const firstHeader = Array(headers.length).fill('');
  firstHeader[columns.TIME - 1] = new Date(2026, 7, 10);
  firstHeader[columns.HOSPITAL - 1] = '高榮';
  firstHeader[columns.CHART_NO - 1] = marker;
  const firstPatient = Array(headers.length).fill('');
  firstPatient[columns.TIME - 1] = '08:00';
  firstPatient[columns.CHART_NO - 1] = 'TEST-101';
  const secondPatient = Array(headers.length).fill('');
  secondPatient[columns.TIME - 1] = '09:00';
  secondPatient[columns.NAME - 1] = '虛構病人';
  const secondHeader = Array(headers.length).fill('');
  secondHeader[columns.TIME - 1] = new Date(2026, 7, 24);
  secondHeader[columns.HOSPITAL - 1] = '高榮';
  secondHeader[columns.CHART_NO - 1] = marker;
  const sheet = new FakeSheet('202608', 202612, [
    headers,
    firstHeader,
    firstPatient,
    secondPatient,
    secondHeader
  ]);
  const preview = call(
    'buildSurgeryDateInsertionPreview_',
    sheet,
    4,
    new Date(2026, 7, 17)
  );
  assert.deepStrictEqual(plain(preview.affectedRows), [4]);
  assert.deepStrictEqual(plain(preview.calendarRows), [4]);
  assert.deepStrictEqual(plain(preview.duplicateRows), []);
  const beforeHeader = call(
    'buildSurgeryDateInsertionPreview_',
    sheet,
    5,
    new Date(2026, 7, 17)
  );
  assert.deepStrictEqual(plain(beforeHeader.affectedRows), []);
  const duplicate = call(
    'buildSurgeryDateInsertionPreview_',
    sheet,
    4,
    new Date(2026, 7, 10)
  );
  assert.deepStrictEqual(plain(duplicate.duplicateRows), [2]);
  assert.throws(
    () => call(
      'buildSurgeryDateInsertionPreview_',
      sheet,
      4,
      new Date(2026, 8, 1)
    ),
    /日期必須屬於 202608/
  );
}

function testBlockRefRequiresExactHeaderRowAndIdentity() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const marker = evaluate('MONTHLY_DATE_HEADER_MARKER');
  const header1 = Array(headers.length).fill('');
  header1[columns.TIME - 1] = new Date(2026, 7, 10);
  header1[columns.HOSPITAL - 1] = '高榮';
  header1[columns.CHART_NO - 1] = marker;
  const header2 = header1.slice();
  const sheet = new FakeSheet('202608', 202613, [
    headers,
    header1,
    header2
  ]);
  const resolved = call(
    'resolveMonthlyBlockRef_',
    sheet,
    { headerRow: 3, dateKey: '2026-08-10', hospital: '高榮' },
    columns
  );
  assert.strictEqual(resolved.row, 3);
  assert.strictEqual(
    call(
      'resolveMonthlyBlockRef_',
      sheet,
      { headerRow: 3, dateKey: '2026-08-11', hospital: '高榮' },
      columns
    ),
    null
  );
}

function testPerformanceFastPathsAvoidGlobalRegistryRebuilds() {
  const workflows = sourceByFile['workflows.js'];
  const fuSort = workflows.slice(
    workflows.indexOf('function rebuildFuSheetSorted_'),
    workflows.indexOf('function getSelectedManagedRow_')
  );
  assert.ok(fuSort.includes('sortRowsWithTemporaryKeys_'));
  assert.strictEqual(fuSort.includes('writeRowSnapshot_'), false);
  assert.strictEqual(fuSort.includes('applyFuFormatting_'), false);
  assert.strictEqual(fuSort.includes('reconcileCalendarRegistry_'), false);
  const monthlySort = sourceByFile['sheet_model.js'].slice(
    sourceByFile['sheet_model.js'].indexOf(
      'function rebuildMonthlySheetSorted_'
    ),
    sourceByFile['sheet_model.js'].indexOf(
      'function findOrCreateMonthlyBlock_'
    )
  );
  assert.ok(monthlySort.includes('sortRowsWithTemporaryKeys_'));
  assert.strictEqual(monthlySort.includes('writeRowSnapshot_'), false);
  assert.strictEqual(monthlySort.includes('applyMonthlyFormatting_'), false);
  const reorderPreflight = sourceByFile['calendar_sync.js'].slice(
    sourceByFile['calendar_sync.js'].indexOf(
      'function assertManagedSheetRegistrySafeForReorder_'
    ),
    sourceByFile['calendar_sync.js'].indexOf(
      'function buildCurrentCalendarScan_'
    )
  );
  assert.strictEqual(
    reorderPreflight.includes('reconcileCalendarRegistry_'),
    false
  );
  assert.strictEqual(
    reorderPreflight.includes('listConfiguredCalendarEvents_'),
    false
  );
  assert.strictEqual(
    reorderPreflight.includes('buildGlobalEventIdLocationsLightweight_'),
    false
  );
  assert.ok(reorderPreflight.includes('buildManagedSheetFastState_'));
  assert.ok(reorderPreflight.includes('fastFingerprint'));
  const registryWrite = sourceByFile['calendar_sync.js'].slice(
    sourceByFile['calendar_sync.js'].indexOf(
      'function writeCalendarRegistryStore_'
    ),
    sourceByFile['calendar_sync.js'].indexOf(
      'function readCalendarRegistryStore_'
    )
  );
  assert.ok(registryWrite.includes('properties.setProperties('));
  assert.strictEqual(
    registryWrite.includes('properties.setProperty(chunk.key'),
    false
  );
  assert.strictEqual(
    evaluate('typeof assertRegistrySafeIgnoringMonthlyStructure_'),
    'undefined'
  );
  const syncRows = sourceByFile['calendar_sync.js'].slice(
    sourceByFile['calendar_sync.js'].indexOf(
      'function syncManagedRowsAt_'
    ),
    sourceByFile['calendar_sync.js'].indexOf(
      'function getRepairSeedRows_'
    )
  );
  assert.ok(syncRows.includes('globalIdScan: globalIds'));
  assert.ok(syncRows.includes('settings.sheetScan ||'));
  const monthlyToFu = workflows.slice(
    workflows.indexOf('function addSelectedMonthlyRowToFu'),
    workflows.indexOf('function parseConditionSuggestion_')
  );
  assert.strictEqual(monthlyToFu.includes('rebuildCalendarRegistry_'), false);
  const createMonth = workflows.slice(
    workflows.indexOf('function createMonthlySurgerySheet'),
    workflows.indexOf('function buildSurgeryDateInsertionPreview_')
  );
  assert.strictEqual(createMonth.includes('rebuildCalendarRegistry_'), false);
  const duplicateScan = workflows.slice(
    workflows.indexOf('function findMonthlyDuplicateCandidates_'),
    workflows.indexOf('function locateCachedMonthlySchedule_')
  );
  assert.strictEqual(duplicateScan.includes('getActiveMonthlySheets_'), false);
  const dialogBuilder = workflows.slice(
    workflows.indexOf('function buildFuToMonthlyDialogData_'),
    workflows.indexOf('function getFuToMonthlyBlockOptions')
  );
  assert.strictEqual(dialogBuilder.includes('getMonthlyBlockOptions_'), false);
  assert.strictEqual(dialogBuilder.includes('scanMonthlyBlocks_'), false);
  assert.strictEqual(
    evaluate('typeof getFuToMonthlyBlockOptions'),
    'function'
  );
  const fuToMonthlySubmit = workflows.slice(
    workflows.indexOf('function submitFuToMonthlySchedule'),
    workflows.length
  );
  assert.strictEqual(
    fuToMonthlySubmit.includes('getMonthlyContextAtRow_'),
    false
  );
  assert.strictEqual(
    fuToMonthlySubmit.includes('syncManagedRowAt_'),
    false
  );
  assert.ok(fuToMonthlySubmit.includes('syncPreparedManagedContextAt_'));
  const insertPatient = sourceByFile['sheet_model.js'].slice(
    sourceByFile['sheet_model.js'].indexOf(
      'function insertPatientAtBlockEnd_'
    ),
    sourceByFile['sheet_model.js'].indexOf(
      'function getMonthlyBlockOptions_'
    )
  );
  assert.strictEqual(
    (insertPatient.match(/scanMonthlyBlocks_/g) || []).length,
    1
  );
  assert.strictEqual(insertPatient.includes('const refreshed ='), false);
  const rowFormat = sourceByFile['sheet_model.js'].slice(
    sourceByFile['sheet_model.js'].indexOf('function applyManagedRowFormat_'),
    sourceByFile['sheet_model.js'].indexOf(
      'function initializeMainTrackingSheet_'
    )
  );
  assert.ok(rowFormat.includes('getRangeList'));
  assert.ok(sourceByFile['calendar_sync.js'].includes(
    'function refreshCalendarRegistryForSheet_'
  ));
  const defaultBlocks = sourceByFile['sheet_model.js'].slice(
    sourceByFile['sheet_model.js'].indexOf(
      'function ensureDefaultMonthlyBlocks_'
    ),
    sourceByFile['sheet_model.js'].indexOf(
      'function getMonthlyTimeSortKey_'
    )
  );
  assert.strictEqual(
    (defaultBlocks.match(/scanMonthlyBlocks_/g) || []).length,
    1
  );
  assert.ok(defaultBlocks.includes('{ skipExistingCheck: true, scan }'));
  assert.strictEqual(
    evaluate('typeof previewFuRemoveTimeMigration'),
    'undefined'
  );
  assert.strictEqual(
    evaluate('typeof executeMonthlyDateHeaderUpgrade'),
    'undefined'
  );
}

function testFuNativeSortUsesOneSortAndPreservesNotes() {
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const columns = Object.fromEntries(
    Object.entries(plain(evaluate('CONFIG.FIELD_HEADERS')))
      .map(([key, header]) => [key, headers.indexOf(header) + 1])
  );
  const makeRow = (name, date, eventId) => {
    const row = Array(headers.length).fill('');
    row[columns.NAME - 1] = name;
    row[columns.DATE - 1] = date;
    row[columns.EVENT_ID - 1] = eventId;
    return row;
  };
  const sheet = new FakeSheet('FU', 301, [
    headers,
    makeRow('較晚', new Date(2026, 7, 20), 'event-late'),
    makeRow('無日期', '', ''),
    makeRow('同日甲', new Date(2026, 7, 10), 'event-a'),
    makeRow('同日乙', new Date(2026, 7, 10), 'event-b')
  ]);
  sheet.setNoteAt(2, columns.PLAN, '人工 note');
  const result = call('rebuildFuSheetSorted_', sheet);
  assert.strictEqual(result.strategy, 'native_range_sort');
  assert.deepStrictEqual(
    [2, 3, 4, 5].map(row => sheet.valueAt(row, columns.NAME)),
    ['同日甲', '同日乙', '較晚', '無日期']
  );
  assert.strictEqual(sheet.noteAt(4, columns.PLAN), '人工 note');
  assert.strictEqual(sheet.calls.sort, 1);
  assert.strictEqual(sheet.calls.insertColumnsAfter, 1);
  assert.strictEqual(sheet.calls.deleteColumns, 1);
  assert.strictEqual(sheet.calls.setValues, 1);
  assert.strictEqual(sheet.calls.autoResizeRows, 1);
  assert.ok(sheet.calls.getRange < 20);
}

function testFuNativeSortSkipsAlreadySortedRows() {
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const columns = Object.fromEntries(
    Object.entries(plain(evaluate('CONFIG.FIELD_HEADERS')))
      .map(([key, header]) => [key, headers.indexOf(header) + 1])
  );
  const makeRow = (name, date, eventId) => {
    const row = Array(headers.length).fill('');
    row[columns.NAME - 1] = name;
    row[columns.DATE - 1] = date;
    row[columns.EVENT_ID - 1] = eventId;
    return row;
  };
  const sheet = new FakeSheet('FU', 304, [
    headers,
    makeRow('較早', new Date(2026, 7, 10), 'event-early'),
    makeRow('較晚', new Date(2026, 7, 20), 'event-late'),
    makeRow('無日期', '', '')
  ]);
  const fastState = call('buildManagedSheetFastState_', sheet);
  const rangeCallsBefore = sheet.calls.getRange;
  const result = call(
    'rebuildFuSheetSorted_',
    sheet,
    {
      precomputedValues: fastState.values,
      precomputedColumns: fastState.columns
    }
  );
  assert.strictEqual(result.changed, false);
  assert.strictEqual(result.strategy, 'precomputed_noop');
  assert.strictEqual(result.sheetScan, null);
  assert.strictEqual(sheet.calls.getRange, rangeCallsBefore);
  assert.strictEqual(sheet.calls.getFormulas, 0);
  assert.strictEqual(sheet.calls.getNotes, 0);
  assert.strictEqual(sheet.calls.sort, 0);
  assert.strictEqual(sheet.calls.insertColumnsAfter, 0);
  assert.strictEqual(sheet.calls.deleteColumns, 0);
  assert.strictEqual(sheet.calls.setValues, 0);
  assert.strictEqual(sheet.calls.autoResizeRows, 0);
}

function testTemporaryNativeSortRollsBackOnVerificationFailure() {
  const sheet = new FakeSheet('FU', 302, [
    ['欄位'],
    ['第二'],
    ['第一']
  ]);
  assert.throws(
    () => call(
      'sortRowsWithTemporaryKeys_',
      sheet,
      2,
      2,
      1,
      ['1', '0'],
      () => {
        throw new Error('驗證失敗');
      }
    ),
    /驗證失敗/
  );
  assert.deepStrictEqual(
    [sheet.valueAt(2, 1), sheet.valueAt(3, 1)],
    ['第二', '第一']
  );
  assert.strictEqual(sheet.getLastColumn(), 1);
  assert.strictEqual(sheet.calls.sort, 2);
  assert.strictEqual(sheet.calls.deleteColumns, 1);
}

function testMonthlyNativeSortPreservesBlocksAndUsesOneSort() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const marker = evaluate('MONTHLY_DATE_HEADER_MARKER');
  const makeHeader = (date, memo) => {
    const row = Array(headers.length).fill('');
    row[columns.TIME - 1] = date;
    row[columns.HOSPITAL - 1] = '高榮';
    row[columns.CHART_NO - 1] = marker;
    row[columns.NAME - 1] = memo;
    return row;
  };
  const makePatient = (time, name, eventId) => {
    const row = Array(headers.length).fill('');
    row[columns.TIME - 1] = time;
    row[columns.NAME - 1] = name;
    row[columns.EVENT_ID - 1] = eventId;
    return row;
  };
  const sheet = new FakeSheet('202608', 303, [
    headers,
    makeHeader(new Date(2026, 7, 24), '較晚刀日'),
    makePatient('PM', '文字時間', 'event-pm'),
    makePatient('09:00', '早上', 'event-am'),
    Array(headers.length).fill(''),
    Array(headers.length).fill(''),
    makeHeader(new Date(2026, 7, 10), '較早刀日'),
    makePatient('12:00', '中午', 'event-noon'),
    makePatient('08:00', '最早', 'event-first')
  ]);
  sheet.setNoteAt(2, columns.NAME, '刀日人工 note');
  const result = call('rebuildMonthlySheetSorted_', sheet);
  assert.strictEqual(result.strategy, 'native_range_sort');
  assert.strictEqual(result.blockCount, 2);
  assert.strictEqual(result.patientCount, 4);
  assert.strictEqual(
    call('formatDateKey_', sheet.valueAt(2, columns.TIME)),
    '2026-08-10'
  );
  assert.deepStrictEqual(
    [3, 4].map(row => sheet.valueAt(row, columns.NAME)),
    ['最早', '中午']
  );
  assert.strictEqual(
    call('formatDateKey_', sheet.valueAt(10, columns.TIME)),
    '2026-08-24'
  );
  assert.deepStrictEqual(
    [11, 12].map(row => sheet.valueAt(row, columns.NAME)),
    ['早上', '文字時間']
  );
  assert.strictEqual(sheet.noteAt(10, columns.NAME), '刀日人工 note');
  assert.strictEqual(sheet.calls.sort, 1);
  assert.strictEqual(sheet.calls.insertColumnsAfter, 1);
  assert.strictEqual(sheet.calls.deleteColumns, 1);
  assert.strictEqual(sheet.calls.autoResizeRows, 1);
  assert.ok(sheet.calls.getRange < 40);
}

function testMonthlyNativeSortSkipsAlreadySortedSheet() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const marker = evaluate('MONTHLY_DATE_HEADER_MARKER');
  const header = Array(headers.length).fill('');
  header[columns.TIME - 1] = new Date(2026, 7, 10);
  header[columns.HOSPITAL - 1] = '高榮';
  header[columns.CHART_NO - 1] = marker;
  const makePatient = (time, name, eventId) => {
    const row = Array(headers.length).fill('');
    row[columns.TIME - 1] = time;
    row[columns.NAME - 1] = name;
    row[columns.EVENT_ID - 1] = eventId;
    return row;
  };
  const sheet = new FakeSheet('202608', 305, [
    headers,
    header,
    makePatient('08:00', '較早', 'event-early'),
    makePatient('12:00', '較晚', 'event-late'),
    ...Array.from(
      { length: 5 },
      () => Array(headers.length).fill('')
    )
  ]);
  const result = call('rebuildMonthlySheetSorted_', sheet);
  assert.strictEqual(result.changed, false);
  assert.strictEqual(result.sheetScan, null);
  assert.strictEqual(result.strategy, 'precomputed_noop');
  assert.strictEqual(sheet.calls.getFormulas, 0);
  assert.strictEqual(sheet.calls.getNotes, 0);
  assert.strictEqual(sheet.calls.sort, 0);
  assert.strictEqual(sheet.calls.insertColumnsAfter, 0);
  assert.strictEqual(sheet.calls.deleteColumns, 0);
  assert.strictEqual(sheet.calls.autoResizeRows, 0);
}

function testFuToMonthlyReusesPrecomputedMonthScan() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const marker = evaluate('MONTHLY_DATE_HEADER_MARKER');
  const header = Array(headers.length).fill('');
  header[columns.TIME - 1] = new Date(2026, 7, 10);
  header[columns.HOSPITAL - 1] = '高榮';
  header[columns.CHART_NO - 1] = marker;
  const patient = Array(headers.length).fill('');
  patient[columns.TIME - 1] = '08:00';
  patient[columns.CHART_NO - 1] = 'TEST-001';
  patient[columns.NAME - 1] = '測試病人';
  patient[columns.DIAGNOSIS - 1] = 'CATA';
  const sheet = new FakeSheet('202608', 202614, [
    headers,
    header,
    patient
  ]);
  const scan = call('scanMonthlyBlocks_', sheet, columns);
  const rangeReadsAfterScan = sheet.calls.getRange;
  const request = {
    chartNo: 'TEST-001',
    patientName: '測試病人',
    time: '08:00',
    ga: '',
    side: '',
    diagnosis: 'CATA',
    grade: '',
    procedure: '',
    iol: '',
    iolTarget: '',
    iolFinal: '',
    axis: '',
    plan: '',
    memo: ''
  };
  const duplicate = call(
    'findMonthlyDuplicateCandidates_',
    sheet,
    request,
    new Date(2026, 7, 10),
    '高榮',
    scan,
    columns
  );
  const resolved = call(
    'resolveMonthlyBlockRef_',
    sheet,
    { headerRow: 2, dateKey: '2026-08-10', hospital: '高榮' },
    columns,
    scan
  );
  assert.strictEqual(duplicate.matches.length, 1);
  assert.strictEqual(resolved.row, 2);
  assert.strictEqual(sheet.calls.getRange, rangeReadsAfterScan);
}

function testManagedRowFormattingUsesRangeListsAndLeavesCustomColumnAlone() {
  const headers = [
    '自訂欄',
    'Condition',
    '姓名',
    'CalendarEventId',
    '日期',
    'Tag',
    '病歷號',
    'Plan',
    'TEL',
    '醫院',
    '心得'
  ];
  const sheet = new FakeSheet('FU', 2050, [
    headers,
    Array(headers.length).fill('')
  ]);
  const columns = call('getRequiredFuColumns_', sheet);
  const beforeRangeCalls = sheet.calls.getRange;
  call('applyManagedRowFormat_', sheet, 2, 'FU', columns);
  assert.strictEqual(sheet.calls.rangeLists.length, 2);
  const formatted = sheet.calls.rangeLists.flat();
  assert.strictEqual(formatted.includes('A2'), false);
  assert.ok(formatted.includes('B2'));
  assert.ok(formatted.includes('D2'));
  assert.strictEqual(
    sheet.calls.getRange - beforeRangeCalls,
    1,
    'FU 列格式只應額外讀寫 Tag 驗證，不逐格設定字型'
  );
}

function testTouchedRowFormattingOnlyChangesEditedStandardColumns() {
  const headers = ['自訂欄'].concat(plain(evaluate('CONFIG.HEADERS')));
  const sheet = new FakeSheet('FU', 2051, [
    headers,
    Array(headers.length).fill('')
  ]);
  const columns = call('getRequiredFuColumns_', sheet);
  const range = sheet.getRange(2, columns.PLAN, 1, 1);
  const touched = call(
    'getTouchedManagedKeys_',
    range,
    columns,
    plain(evaluate('CONFIG.FIELD_KEYS'))
  );
  assert.deepStrictEqual(plain(touched), ['PLAN']);
  call(
    'applyTouchedManagedRowsFormatBatch_',
    sheet,
    [2],
    'FU',
    columns,
    touched
  );
  const addresses = sheet.calls.rangeLists.flat();
  const expected = `${call('columnToLetter_', columns.PLAN)}2`;
  assert.ok(addresses.every(address => address === expected));
  assert.strictEqual(addresses.includes('A2'), false);
  assert.strictEqual(
    sheet.calls.numberFormatRangeLists.length,
    0
  );
}

function testLightweightEventIndexAndSingleSheetRegistryRefresh() {
  Object.keys(scriptPropertyStore).forEach(key => {
    if (
      key === 'CALENDAR_ROW_REGISTRY_V2_INDEX' ||
      key.startsWith('CALENDAR_ROW_REGISTRY_V2_')
    ) {
      delete scriptPropertyStore[key];
    }
  });
  const fuHeaders = plain(evaluate('CONFIG.HEADERS'));
  const fuRow = Array(fuHeaders.length).fill('');
  fuRow[fuHeaders.indexOf('病歷號')] = 'TEST-FU';
  fuRow[fuHeaders.indexOf('日期')] = new Date(2026, 7, 20);
  fuRow[fuHeaders.indexOf('CalendarEventId')] = 'event-fu';
  const fu = new FakeSheet('FU', 3101, [fuHeaders, fuRow]);

  const monthHeaders = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const header = Array(monthHeaders.length).fill('');
  header[columns.TIME - 1] = new Date(2026, 7, 24);
  header[columns.HOSPITAL - 1] = '高榮';
  header[columns.CHART_NO - 1] = evaluate('MONTHLY_DATE_HEADER_MARKER');
  const patient = Array(monthHeaders.length).fill('');
  patient[columns.TIME - 1] = '08:00';
  patient[columns.NAME - 1] = '虛構病人';
  patient[columns.EVENT_ID - 1] = 'event-month';
  const month = new FakeSheet('202608', 3102, [
    monthHeaders,
    header,
    patient
  ]);
  const spreadsheet = {
    getSheets: () => [fu, month],
    getSheetByName: name => (
      name === 'FU' ? fu : name === '202608' ? month : null
    )
  };
  const scan = call('buildCurrentCalendarScan_', spreadsheet);
  call('writeCalendarRegistryStore_', scan, []);
  const refreshed = call(
    'refreshCalendarRegistryForSheet_',
    spreadsheet,
    month
  );
  assert.strictEqual(refreshed.eventCount, 2);
  assert.strictEqual(refreshed.sheetEventCount, 1);
  call(
    'refreshCalendarRegistryForSheet_',
    spreadsheet,
    month,
    {
      retainEventIds: ['event-month'],
      pendingSyncByEventId: { 'event-month': 'update_failed' }
    }
  );
  const pendingRegistry = call('readCalendarRegistryStore_');
  assert.ok(
    pendingRegistry.index.sheets.every(item => item.fastFingerprint),
    '單表索引刷新應保留其他表的 fastFingerprint'
  );
  const pendingEntry = pendingRegistry.entries.find(entry => {
    return entry.eventId === 'event-month';
  });
  assert.strictEqual(pendingEntry.pendingSync, true);
  assert.strictEqual(pendingEntry.pendingError, 'update_failed');
  fu.setValueAt(
    2,
    fuHeaders.indexOf('CalendarEventId') + 1,
    'event-month'
  );
  const duplicate = call(
    'buildGlobalEventIdLocationsLightweight_',
    spreadsheet
  );
  assert.deepStrictEqual(plain(duplicate.duplicateIds), ['event-month']);
  Object.keys(scriptPropertyStore).forEach(key => {
    if (
      key === 'CALENDAR_ROW_REGISTRY_V2_INDEX' ||
      key.startsWith('CALENDAR_ROW_REGISTRY_V2_')
    ) {
      delete scriptPropertyStore[key];
    }
  });
}

function testSingleSheetSyncUsesVerifiedRegistryBeforeGlobalScan() {
  clearScriptProperties();
  const fuHeaders = plain(evaluate('CONFIG.HEADERS'));
  const fuRow = Array(fuHeaders.length).fill('');
  fuRow[fuHeaders.indexOf('姓名')] = '虛構追蹤';
  fuRow[fuHeaders.indexOf('日期')] = new Date(2026, 7, 20);
  fuRow[fuHeaders.indexOf('CalendarEventId')] = 'event-fu-fast';
  const fu = new FakeSheet('FU', 3201, [fuHeaders, fuRow]);

  const monthHeaders = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const header = Array(monthHeaders.length).fill('');
  header[columns.TIME - 1] = new Date(2026, 7, 24);
  header[columns.HOSPITAL - 1] = '高榮';
  header[columns.CHART_NO - 1] = evaluate('MONTHLY_DATE_HEADER_MARKER');
  const patient = Array(monthHeaders.length).fill('');
  patient[columns.TIME - 1] = '08:00';
  patient[columns.NAME - 1] = '虛構手術';
  patient[columns.EVENT_ID - 1] = 'event-month-fast';
  const month = new FakeSheet('202608', 3202, [
    monthHeaders,
    header,
    patient
  ]);
  let getSheetsCalls = 0;
  const spreadsheet = {
    getSheets: () => {
      getSheetsCalls++;
      return [fu, month];
    },
    getSheetByName: name => (
      name === 'FU' ? fu : name === '202608' ? month : null
    )
  };
  const fullScan = call('buildCurrentCalendarScan_', spreadsheet);
  call('writeCalendarRegistryStore_', fullScan, []);
  const monthScan = call('buildManagedSheetCalendarScan_', month);
  getSheetsCalls = 0;
  const resolved = call(
    'resolveGlobalEventIdLocationsForSync_',
    spreadsheet,
    month,
    monthScan
  );
  assert.strictEqual(resolved.source, 'registry_v2');
  assert.strictEqual(getSheetsCalls, 0);
  assert.deepStrictEqual(
    Object.keys(plain(resolved.eventLocations)).sort(),
    ['event-fu-fast', 'event-month-fast']
  );

  month.setValueAt(3, columns.EVENT_ID, 'event-not-in-registry');
  const changedScan = call('buildManagedSheetCalendarScan_', month);
  const fallback = call(
    'resolveGlobalEventIdLocationsForSync_',
    spreadsheet,
    month,
    changedScan
  );
  assert.strictEqual(fallback.source, 'live_sheet_scan');
  assert.ok(getSheetsCalls > 0);
  clearScriptProperties();
}

function testArchiveEventIdsArePersistedInRegistry() {
  clearScriptProperties();
  const scan = registryScan([
    registryContext({ eventId: 'event-active', sheetId: 41 })
  ], [{
    sheetId: 41,
    sheetName: '202608',
    kind: 'MONTHLY',
    archived: false,
    fastFingerprint: 'fast-active'
  }]);
  scan.archiveLocations = [{
    eventId: 'event-archive',
    sheetId: 42,
    sheetName: '刀表封存_2026',
    row: 9,
    kind: 'ARCHIVE',
    archived: true,
    sourceMonth: '202604'
  }];
  call('writeCalendarRegistryStore_', scan, []);
  const loaded = call('readCalendarRegistryStore_');
  const archived = loaded.entries.find(entry => {
    return entry.eventId === 'event-archive';
  });
  assert.ok(archived);
  assert.strictEqual(archived.archived, true);
  assert.strictEqual(archived.kind, 'ARCHIVE');
  assert.strictEqual(archived.sourceMonth, '202604');
  assert.strictEqual(loaded.index.eventCount, 2);
  clearScriptProperties();
}

function testReorderPreflightIsLocalAndDetectsUnsyncedRowHash() {
  Object.keys(scriptPropertyStore)
    .filter(key => key.includes('CALENDAR_ROW_REGISTRY_V2'))
    .forEach(key => delete scriptPropertyStore[key]);
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const row = Array(headers.length).fill('');
  row[headers.indexOf('姓名')] = '測試姓名';
  row[headers.indexOf('日期')] = new Date(2026, 7, 20);
  row[headers.indexOf('Condition')] = '穩定內容';
  row[headers.indexOf('CalendarEventId')] = 'event-stable';
  const fu = new FakeSheet('FU', 3150, [headers, row]);
  const spreadsheet = {
    getSheets: () => [fu],
    getSheetByName: name => name === 'FU' ? fu : null
  };
  call(
    'writeCalendarRegistryStore_',
    call('buildCurrentCalendarScan_', spreadsheet),
    []
  );
  const safe = call(
    'assertManagedSheetRegistrySafeForReorder_',
    spreadsheet,
    fu,
    '測試排序'
  );
  assert.strictEqual(safe.ok, true);
  assert.strictEqual(safe.usedFastFingerprint, true);
  assert.strictEqual(safe.sheetScan, null);
  assert.ok(safe.fastState.fastFingerprint);
  assert.strictEqual(safe.globalIdScan.duplicateIds.length, 0);
  fu.setValueAt(2, headers.indexOf('Condition') + 1, '尚未同步的新內容');
  assert.throws(
    () => call(
      'assertManagedSheetRegistrySafeForReorder_',
      spreadsheet,
      fu,
      '測試排序'
    ),
    /同步差異/
  );
  Object.keys(scriptPropertyStore)
    .filter(key => key.includes('CALENDAR_ROW_REGISTRY_V2'))
    .forEach(key => delete scriptPropertyStore[key]);
}

function testFastFingerprintMetadataFallsBackSafely() {
  Object.keys(scriptPropertyStore)
    .filter(key => key.includes('CALENDAR_ROW_REGISTRY_V2'))
    .forEach(key => delete scriptPropertyStore[key]);
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const row = Array(headers.length).fill('');
  row[headers.indexOf('姓名')] = '虛構病人';
  row[headers.indexOf('日期')] = new Date(2026, 7, 20);
  row[headers.indexOf('Condition')] = '虛構追蹤';
  row[headers.indexOf('CalendarEventId')] = 'event-fast';
  const fu = new FakeSheet('FU', 3160, [headers, row]);
  const spreadsheet = {
    getSheets: () => [fu],
    getSheetByName: name => name === 'FU' ? fu : null
  };
  const scan = call('buildCurrentCalendarScan_', spreadsheet);
  const index = call('writeCalendarRegistryStore_', scan, []);
  assert.ok(index.sheets[0].fastFingerprint);
  assert.strictEqual(
    JSON.stringify(index).includes('虛構病人'),
    false
  );

  const first = call(
    'assertManagedSheetRegistrySafeForReorder_',
    spreadsheet,
    fu,
    '快速排序'
  );
  assert.strictEqual(first.usedFastFingerprint, true);
  assert.strictEqual(first.sheetScan, null);

  const storedIndex = JSON.parse(
    scriptPropertyStore.CALENDAR_ROW_REGISTRY_V2_INDEX
  );
  delete storedIndex.sheets[0].fastFingerprint;
  scriptPropertyStore.CALENDAR_ROW_REGISTRY_V2_INDEX =
    JSON.stringify(storedIndex);
  const fallback = call(
    'assertManagedSheetRegistrySafeForReorder_',
    spreadsheet,
    fu,
    '舊索引排序'
  );
  assert.strictEqual(fallback.usedFastFingerprint, false);
  assert.strictEqual(fallback.sheetScan.contexts.length, 1);

  call(
    'writeCalendarRegistryStore_',
    scan,
    [{
      ...plain(call('toRegistryEntry_', scan.contexts[0])),
      pendingSync: true,
      pendingError: 'test_pending'
    }]
  );
  assert.throws(
    () => call(
      'assertManagedSheetRegistrySafeForReorder_',
      spreadsheet,
      fu,
      '待重試排序'
    ),
    /待重試 1/
  );
  Object.keys(scriptPropertyStore)
    .filter(key => key.includes('CALENDAR_ROW_REGISTRY_V2'))
    .forEach(key => delete scriptPropertyStore[key]);
}

function testDefaultMonthlyDates() {
  assert.deepStrictEqual(
    plain(
      call('getDefaultMonthlyDates_', '202608')
        .map(date => call('formatDateKey_', date))
    ),
    ['2026-08-10', '2026-08-13', '2026-08-24', '2026-08-27']
  );
  assert.deepStrictEqual(
    plain(
      call('getDefaultMonthlyDates_', '202609')
        .map(date => call('formatDateKey_', date))
    ),
    ['2026-09-10', '2026-09-14', '2026-09-24', '2026-09-28']
  );
  assert.strictEqual(call('getDefaultMonthlyDates_', '202402').length, 4);
  assert.strictEqual(call('getDefaultMonthlyDates_', '202701').length, 4);
}

function testCalendarTitlePreservesSlots() {
  assert.strictEqual(
    call('buildCalendarTitle_', 'X001', '', 'CATA OS'),
    'X001 |  | CATA OS'
  );
  assert.strictEqual(
    call('buildCalendarTitle_', '', '測試姓名', ''),
    ' | 測試姓名 | '
  );
}

function testRepairBindingSignatureIgnoresOnlyTrailingSurgeryDate() {
  const fromSheet = call(
    'buildCalendarBindingSignature_',
    'TEST-301',
    '虛構姓名',
    'CATA OS s/p Phaco-IOL OS 20260813'
  );
  const fromCalendar = call(
    'getEventCalendarBindingSignature_',
    {
      summary:
        'TEST-301 | 虛構姓名 | CATA OS s/p Phaco-IOL OS 20260810'
    }
  );
  assert.strictEqual(fromSheet, fromCalendar);
  assert.notStrictEqual(
    fromSheet,
    call(
      'buildCalendarBindingSignature_',
      'TEST-301',
      '虛構姓名',
      'ERM OS s/p VT OS 20260813'
    )
  );
  assert.strictEqual(
    call(
      'getEventCalendarBindingSignature_',
      { summary: 'invalid title without three slots' }
    ),
    ''
  );
}

function testRepairEventFingerprintIgnoresApiObjectKeyOrder() {
  const first = {
    id: 'event-test',
    summary: 'TEST-302 | 虛構姓名 | CATA OS',
    description: '醫院: 高榮',
    colorId: '10',
    start: {
      dateTime: '2026-08-13T12:00:00+08:00',
      timeZone: 'Asia/Taipei'
    },
    end: {
      dateTime: '2026-08-13T12:30:00+08:00',
      timeZone: 'Asia/Taipei'
    }
  };
  const reordered = {
    colorId: '10',
    description: '醫院: 高榮',
    summary: 'TEST-302 | 虛構姓名 | CATA OS',
    id: 'event-test',
    start: {
      timeZone: 'Asia/Taipei',
      dateTime: '2026-08-13T12:00:00+08:00'
    },
    end: {
      timeZone: 'Asia/Taipei',
      dateTime: '2026-08-13T12:30:00+08:00'
    }
  };
  assert.strictEqual(
    call('getRepairEventFingerprint_', first),
    call('getRepairEventFingerprint_', reordered)
  );
  assert.notStrictEqual(
    call('getRepairEventFingerprint_', first),
    call('getRepairEventFingerprint_', {
      ...first,
      extendedProperties: {
        private: {
          surgerySyncKind: 'MONTHLY',
          surgerySyncState: 'CANCELLED'
        }
      }
    })
  );
}

function testMonthlyConditionComposition() {
  const date = new Date(2026, 7, 10);
  assert.strictEqual(
    call(
      'buildMonthlyCondition_',
      {
        diagnosis: 'CATA',
        grade: 'NS+CO++',
        side: 'OS',
        procedure: 'Phaco-IOL',
        iol: 'EMV',
        iolTarget: 'T-0.1',
        iolFinal: '+19.0',
        axis: '20'
      },
      date
    ),
    'CATA(NS+CO++) OS s/p Phaco-IOL(EMV T-0.1 +19.0 @20) OS 20260810'
  );
  const composite = call(
    'buildMonthlyCondition_',
    {
      diagnosis: 'CATA OS + eyelid mass OD',
      grade: 'NS+',
      side: 'OU',
      procedure: 'Phaco-IOL OS + excision OD',
      iol: 'EMV'
    },
    date
  );
  assert.strictEqual(
    composite,
    'CATA OS + eyelid mass OD(NS+) s/p Phaco-IOL OS + excision OD(EMV) 20260810'
  );
  assert.strictEqual((composite.match(/\bOU\b/g) || []).length, 0);
}

function testCalendarResourceTimedAndAllDay() {
  const base = {
    kind: 'MONTHLY',
    chartNo: 'X001',
    patientName: '測試姓名',
    condition: 'CATA OS',
    hospital: '高榮',
    tel: '',
    plan: '',
    ga: '',
    date: new Date(2026, 7, 10),
    timeInfo: call('resolveCalendarTime_', '08:00')
  };
  const timed = call('buildCalendarResource_', base);
  assert.ok(timed.start.dateTime.includes('08:00:00'));
  assert.ok(timed.end.dateTime.includes('08:30:00'));
  assert.strictEqual(timed.summary, 'X001 | 測試姓名 | CATA OS');
  assert.strictEqual(timed.colorId, '10');

  const ga = call('buildCalendarResource_', {
    ...base,
    ga: 'GA'
  });
  assert.strictEqual(ga.start.date, '2026-08-10');
  assert.strictEqual(ga.end.date, '2026-08-11');
  assert.ok(ga.description.includes('原報到時間: 08:00'));
  assert.strictEqual(ga.colorId, '4');

  const textTime = call('buildCalendarResource_', {
    ...base,
    timeInfo: call('resolveCalendarTime_', 'PM')
  });
  assert.strictEqual(textTime.start.date, '2026-08-10');
  assert.ok(textTime.description.includes('Time note: PM'));
  assert.strictEqual(textTime.colorId, '2');

  ['0800', 800, 830, '8:00'].forEach(value => {
    const normalized = call('buildCalendarResource_', {
      ...base,
      timeInfo: call('resolveCalendarTime_', value)
    });
    assert.strictEqual(normalized.colorId, '10');
  });

  const fu = call('buildCalendarResource_', {
    ...base,
    kind: 'FU'
  });
  assert.strictEqual(fu.colorId, '8');
  assert.strictEqual(fu.start.date, '2026-08-10');
  assert.strictEqual(fu.start.dateTime, undefined);
  const fuAllDay = call('buildCalendarResource_', {
    ...base,
    kind: 'FU',
    timeInfo: call('resolveCalendarTime_', '')
  });
  assert.strictEqual(fuAllDay.colorId, '8');

  assert.strictEqual(
    call(
      'calendarEventMatchesResource_',
      { ...timed, colorId: '2' },
      timed
    ),
    false
  );
}

function testMonthlyStrikethroughControlsCancellationColorAndMarker() {
  const standardHeaders = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const headers = ['自訂欄'].concat(standardHeaders.slice().reverse());
  const sheet = new FakeSheet('202608', 2026081, [
    headers,
    Array(headers.length).fill(''),
    Array(headers.length).fill('')
  ]);
  const columns = call('getRequiredMonthlyColumns_', sheet);
  sheet.setValueAt(2, columns.TIME, new Date(2026, 7, 28));
  sheet.setValueAt(2, columns.HOSPITAL, '高榮');
  sheet.setValueAt(
    2,
    columns.CHART_NO,
    evaluate('MONTHLY_DATE_HEADER_MARKER')
  );
  sheet.setValueAt(3, columns.TIME, '08:00');
  sheet.setValueAt(3, columns.CHART_NO, 'CANCEL-001');
  sheet.setValueAt(3, columns.NAME, '虛構取消個案');
  sheet.setValueAt(3, columns.DIAGNOSIS, 'CATA');
  sheet.setValueAt(3, columns.IOL_TARGET, '+20.0');
  sheet.setValueAt(3, columns.EVENT_ID, 'cancel-event');

  const normalScan = call('buildManagedSheetCalendarScan_', sheet);
  const normal = normalScan.contexts.find(item => item.row === 3);
  assert.ok(normal);
  assert.strictEqual(normal.cancelled, false);
  assert.strictEqual(call('getCalendarColorId_', normal), '10');
  assert.strictEqual(
    call('buildCalendarResource_', normal).extendedProperties,
    undefined
  );

  sheet.setFontLineAt(3, columns.CHART_NO, 'line-through');
  const cancelledScan = call('buildManagedSheetCalendarScan_', sheet);
  const cancelled = cancelledScan.contexts.find(item => item.row === 3);
  const cancelledResource = call('buildCalendarResource_', cancelled);
  assert.strictEqual(cancelled.cancelled, true);
  assert.strictEqual(cancelledResource.colorId, '8');
  assert.deepStrictEqual(
    plain(cancelledResource.extendedProperties.private),
    {
      surgerySyncKind: 'MONTHLY',
      surgerySyncState: 'CANCELLED'
    }
  );
  assert.strictEqual(cancelled.bindingHash, normal.bindingHash);
  assert.notStrictEqual(cancelled.rowHash, normal.rowHash);
  assert.notStrictEqual(
    cancelledScan.fastFingerprint,
    normalScan.fastFingerprint
  );
  assert.strictEqual(
    call('calendarEventMatchesResource_', cancelledResource, cancelledResource),
    true
  );
  assert.strictEqual(
    call(
      'calendarEventMatchesResource_',
      { ...cancelledResource, extendedProperties: undefined },
      cancelledResource
    ),
    false
  );

  sheet.setValueAt(3, columns.GA, 'GA');
  const cancelledGa = call('buildManagedSheetCalendarScan_', sheet)
    .contexts.find(item => item.row === 3);
  assert.strictEqual(call('getCalendarColorId_', cancelledGa), '8');

  sheet.setFontLineAt(3, columns.CHART_NO, 'none');
  const restored = call('buildManagedSheetCalendarScan_', sheet)
    .contexts.find(item => item.row === 3);
  const restoredResource = call('buildCalendarResource_', restored);
  assert.strictEqual(restored.cancelled, false);
  assert.strictEqual(restoredResource.colorId, '4');
  assert.strictEqual(restoredResource.extendedProperties, undefined);
  assert.ok(sheet.calls.getFontLines >= 4);
}

function testFuIgnoresStrikethroughAndFormatQueuesCalendarScanOnly() {
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const row = Array(headers.length).fill('');
  row[headers.indexOf('病歷號')] = 'FU-STRIKE';
  row[headers.indexOf('日期')] = new Date(2026, 7, 28);
  row[headers.indexOf('Condition')] = '追蹤';
  const sheet = new FakeSheet('FU', 2026082, [headers, row]);
  const columns = call('getRequiredFuColumns_', sheet);
  sheet.setFontLineAt(2, columns.CHART_NO, 'line-through');
  const scan = call('buildManagedSheetCalendarScan_', sheet);
  assert.strictEqual(scan.contexts.length, 1);
  assert.strictEqual(scan.contexts[0].cancelled, undefined);
  assert.strictEqual(sheet.calls.getFontLines, 0);
  assert.strictEqual(
    call('buildCalendarResource_', scan.contexts[0]).colorId,
    '8'
  );

  const source = sourceByFile['calendar_sync.js'];
  const body = source.slice(
    source.indexOf('function processCalendarStructureChange'),
    source.indexOf('function buildCalendarSyncRuntime_')
  );
  assert.ok(body.includes('fullScan: true'));
  assert.ok(body.includes('reason: `CHANGE_${changeType}`'));
  assert.strictEqual(body.includes('setNumberFormat'), false);
}

function testArbitraryHeaderOrderUsesNames() {
  const headers = [
    '自訂欄',
    'Condition',
    '姓名',
    'CalendarEventId',
    '日期',
    'Tag',
    '病歷號',
    'Plan',
    'TEL',
    '醫院',
    '心得'
  ];
  const sheet = new FakeSheet('FU', 1, [headers]);
  const info = call('getFuColumnInfo_', sheet);
  assert.strictEqual(info.columns.COND, 2);
  assert.strictEqual(info.columns.NAME, 3);
  assert.strictEqual(info.columns.EVENT_ID, 4);
  assert.strictEqual(info.columns.CHART_NO, 7);
  assert.strictEqual(info.missingKeys.length, 0);
}

function testDuplicateHeadersAreDiagnosed() {
  const headers = [
    '病歷號',
    '姓名',
    'TEL',
    '醫院',
    'Tag',
    'Condition',
    '日期',
    'Plan',
    '心得',
    'CalendarEventId',
    '姓名'
  ];
  const sheet = new FakeSheet('FU', 1, [headers]);
  const info = call('getFuColumnInfo_', sheet);
  assert.ok(info.duplicateMessages.some(message => message.includes('姓名')));
}

function testMonthlyDiagnosisSummaryFormulaDefinition() {
  const definitions = plain(call(
    'getMonthlyDiagnosisSummaryDefinitions_'
  ));
  assert.deepStrictEqual(definitions, [
    {
      label: 'CATA',
      keywords: ['CATA', 'Cataract']
    },
    {
      label: 'Retina',
      keywords: [
        'VH',
        'ERM',
        'Subluxation',
        'Dislocation',
        'RRD',
        'TRD',
        'Subluxated IOL',
        'RD'
      ]
    },
    {
      label: 'Plasty',
      keywords: ['Dermatochalasis', 'Ptosis', 'Dacryocystitis', 'Entropion']
    }
  ]);
  const categories = plain(call('getMonthlyDiagnosisSummaryCategories_'));
  assert.deepStrictEqual(categories, ['CATA', 'Retina', 'Plasty']);
  const patterns = definitions.map(definition => call(
    'buildMonthlyDiagnosisSummaryPattern_',
    definition.keywords
  ));
  assert.deepStrictEqual(patterns, [
    '(^|[^A-Z0-9])(CATA|CATARACT)([^A-Z0-9]|$)',
    '(^|[^A-Z0-9])' +
      '(VH|ERM|SUBLUXATION|DISLOCATION|RRD|TRD|SUBLUXATED IOL|RD)' +
      '([^A-Z0-9]|$)',
    '(^|[^A-Z0-9])' +
      '(DERMATOCHALASIS|PTOSIS|DACRYOCYSTITIS|ENTROPION)' +
      '([^A-Z0-9]|$)'
  ]);
  const cataPattern = new RegExp(patterns[0], 'i');
  const retinaPattern = new RegExp(patterns[1], 'i');
  const plastyPattern = new RegExp(patterns[2], 'i');
  assert.ok(cataPattern.test('cataract'));
  assert.ok(retinaPattern.test('ERM(4)'));
  assert.ok(retinaPattern.test('IOL subluxation with VH'));
  assert.ok(retinaPattern.test('Subluxated IOL'));
  assert.ok(retinaPattern.test('RD'));
  assert.ok(retinaPattern.test('RRD'));
  assert.ok(retinaPattern.test('TRD'));
  assert.ok(retinaPattern.test('IOL dislocation'));
  assert.strictEqual(retinaPattern.test('Dermatochalasis'), false);
  assert.ok(plastyPattern.test('Acute dacryocystitis'));
  assert.ok(plastyPattern.test('entropion'));
  assert.ok(evaluate('CONFIG.MONTHLY_DIAGNOSIS_OPTIONS').includes('Entropion'));
  assert.ok(cataPattern.test('Cataract + ptosis'));
  assert.ok(plastyPattern.test('Cataract + ptosis'));
  const formula = call('buildMonthlyDiagnosisSummaryFormula_', {
    SIDE: 28,
    DIAGNOSIS: 31
  });
  assert.ok(formula.startsWith('=LET('));
  assert.ok(formula.includes('SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V2'));
  assert.ok(formula.includes('TO_TEXT(IFERROR(INDEX($AE:$AE,2):INDEX($AE:$AE,ROWS($AE:$AE)),""))'));
  assert.ok(formula.includes('TO_TEXT(IFERROR(INDEX($AB:$AB,2):INDEX($AB:$AB,ROWS($AB:$AB)),""))'));
  assert.strictEqual(formula.includes('$G$2:$G'), false);
  assert.strictEqual(formula.includes('$H$2:$H'), false);
  assert.ok(formula.includes('IF(sideValues="OU",2'));
  assert.ok(formula.includes('"^(OD|OS)$"),1,0'));
  assert.ok(formula.includes(
    '"(^|[^A-Z0-9])(CATA|CATARACT)([^A-Z0-9]|$)"'
  ));
  assert.ok(formula.includes(
    '"(^|[^A-Z0-9])' +
      '(VH|ERM|SUBLUXATION|DISLOCATION|RRD|TRD|SUBLUXATED IOL|RD)' +
      '([^A-Z0-9]|$)"'
  ));
  assert.ok(formula.includes('TEXTJOIN(" | ",TRUE'));
  assert.ok(formula.includes(
    'managedMarker&"診斷"&IF(summaryText="","","｜"&summaryText)'
  ));
  assert.strictEqual(
    (formula.match(/SUMPRODUCT\(/g) || []).length,
    categories.length
  );
  let previousOutputIndex = -1;
  categories.forEach((category, index) => {
    const variable = `diagnosisCount${index + 1}`;
    assert.ok(formula.includes(
      `IF(${variable}>0,"${category} "&${variable},"")`
    ));
    const outputIndex = formula.indexOf(`"${category} "&${variable}`);
    assert.ok(outputIndex > previousOutputIndex);
    previousOutputIndex = outputIndex;
  });
}

function testMonthlyDiagnosisSummaryHeadersStayCanonical() {
  const standardHeaders = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const diagnosisColumn = standardHeaders.indexOf('診斷') + 1;

  const dynamicHeaders = standardHeaders.slice();
  dynamicHeaders[diagnosisColumn - 1] = '診斷｜CATA 12 | Retina 2';
  const dynamicSheet = new FakeSheet('202608', 2601, [dynamicHeaders]);
  const dynamicInfo = call('getMonthlyColumnInfo_', dynamicSheet);
  assert.strictEqual(dynamicInfo.columns.DIAGNOSIS, diagnosisColumn);
  assert.strictEqual(dynamicInfo.missingKeys.length, 0);

  const duplicateHeaders = dynamicHeaders.concat(['診斷']);
  const duplicateSheet = new FakeSheet('202608', 2602, [duplicateHeaders]);
  const duplicateInfo = call('getMonthlyColumnInfo_', duplicateSheet);
  assert.ok(
    duplicateInfo.duplicateMessages.some(message => {
      return message.includes('診斷') &&
        message.includes('H') &&
        message.includes('X');
    })
  );

  const errorHeaders = standardHeaders.slice();
  errorHeaders[diagnosisColumn - 1] = '#ERROR!';
  const errorSheet = new FakeSheet('202608', 2603, [errorHeaders]);
  errorSheet.setFormulaAt(
    1,
    diagnosisColumn,
    '=T(N("SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V1"))&"診斷"'
  );
  const errorInfo = call('getMonthlyColumnInfo_', errorSheet);
  assert.strictEqual(errorInfo.columns.DIAGNOSIS, diagnosisColumn);
  assert.strictEqual(errorInfo.missingKeys.length, 0);
  assert.strictEqual(errorSheet.calls.getFormulas, 1);

  const errorDuplicateHeaders = errorHeaders.concat(['診斷']);
  const errorDuplicateSheet = new FakeSheet(
    '202608',
    2604,
    [errorDuplicateHeaders]
  );
  errorDuplicateSheet.setFormulaAt(
    1,
    diagnosisColumn,
    call('buildMonthlyDiagnosisSummaryFormula_', monthlyColumns())
  );
  const errorDuplicateInfo = call(
    'getMonthlyColumnInfo_',
    errorDuplicateSheet
  );
  assert.ok(
    errorDuplicateInfo.duplicateMessages.some(message => {
      return message.includes('診斷');
    })
  );
}

function testMonthlyDiagnosisSummaryMigrationIsSafeAndIdempotent() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const diagnosisColumn = columns.DIAGNOSIS;
  const dataRow = Array(headers.length).fill('');
  dataRow[columns.SIDE - 1] = 'OU';
  dataRow[columns.DIAGNOSIS - 1] = 'CATA + ERM';
  dataRow[columns.CHART_NO - 1] = 'SAFE-001';

  const plainSheet = new FakeSheet('202608', 2610, [
    headers.slice(),
    dataRow.slice()
  ]);
  const passive = call(
    'ensureMonthlyDiagnosisSummaryHeader_',
    plainSheet,
    columns,
    { enable: false }
  );
  assert.strictEqual(passive.status, 'plain');
  assert.strictEqual(plainSheet.formulaAt(1, diagnosisColumn), '');

  const expectedFormula = call(
    'buildMonthlyDiagnosisSummaryFormula_',
    columns
  );
  assert.ok(expectedFormula.includes(
    'SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V2'
  ));
  assert.strictEqual(call(
    'isManagedMonthlyDiagnosisSummaryFormula_',
    '=T(N("SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V1"))&"診斷"'
  ), true);
  const currentSheet = new FakeSheet('202609', 2611, [headers.slice()]);
  currentSheet.setFormulaAt(1, diagnosisColumn, expectedFormula);

  const staleSheet = new FakeSheet('202610', 2612, [headers.slice()]);
  staleSheet.setFormulaAt(
    1,
    diagnosisColumn,
    '=T(N("SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V1"))&"診斷"'
  );

  const customSheet = new FakeSheet('202611', 2613, [
    headers.slice(),
    dataRow.slice()
  ]);
  const customFormula = '="人工診斷摘要"';
  customSheet.setFormulaAt(1, diagnosisColumn, customFormula);

  const spreadsheet = makeFakeSpreadsheet([
    plainSheet,
    currentSheet,
    staleSheet,
    customSheet,
    new FakeSheet('備註', 2614, [['不處理']])
  ], 'diagnosis-summary-migration');
  const originalData = plain(plainSheet.rows.slice(1));
  const plan = call(
    'buildMonthlyDiagnosisSummaryMigrationPlan_',
    spreadsheet
  );
  assert.deepStrictEqual(plain(plan.counts), {
    enable: 1,
    refresh: 1,
    enabled: 0,
    refreshed: 0,
    unchanged: 1,
    manual: 1
  });
  const previewText = call(
    'buildMonthlyDiagnosisSummaryMigrationPreviewText_',
    plan
  );
  assert.ok(previewText.includes('病人資料列複製：0'));
  assert.ok(previewText.includes('病人資料列移除：0'));
  assert.ok(previewText.includes('202611：診斷表頭已有非系統公式'));
  const result = call('applyMonthlyDiagnosisSummaryMigrationPlan_', plan);
  assert.deepStrictEqual(plain(result.counts), {
    enable: 0,
    refresh: 0,
    enabled: 1,
    refreshed: 1,
    unchanged: 1,
    manual: 1
  });
  assert.strictEqual(result.dataRowsCopied, 0);
  assert.strictEqual(result.dataRowsRemoved, 0);
  assert.strictEqual(result.calendarTouched, false);
  const resultText = call(
    'buildMonthlyDiagnosisSummaryMigrationResultText_',
    result
  );
  assert.ok(resultText.includes('Calendar 變更：0'));
  assert.strictEqual(
    plainSheet.formulaAt(1, diagnosisColumn),
    expectedFormula
  );
  assert.strictEqual(
    staleSheet.formulaAt(1, diagnosisColumn),
    expectedFormula
  );
  assert.strictEqual(customSheet.formulaAt(1, diagnosisColumn), customFormula);
  assert.deepStrictEqual(plain(plainSheet.rows.slice(1)), originalData);
  assert.deepStrictEqual(plain(customSheet.rows.slice(1)), originalData);

  const setFormulaCalls = plainSheet.calls.setFormulas;
  const secondPlan = call(
    'buildMonthlyDiagnosisSummaryMigrationPlan_',
    spreadsheet
  );
  assert.strictEqual(secondPlan.counts.unchanged, 3);
  assert.strictEqual(secondPlan.counts.manual, 1);
  const secondResult = call(
    'applyMonthlyDiagnosisSummaryMigrationPlan_',
    secondPlan
  );
  assert.strictEqual(secondResult.counts.unchanged, 3);
  assert.strictEqual(secondResult.counts.manual, 1);
  assert.strictEqual(plainSheet.calls.setFormulas, setFormulaCalls);

  const workflowSource = sourceByFile['workflows.js'];
  const createBody = workflowSource.slice(
    workflowSource.indexOf('function createOrInitializeMonthlySheet_'),
    workflowSource.indexOf('function createMonthlySurgerySheet')
  );
  assert.ok(createBody.includes('enableDiagnosisSummary: created'));
}

function testAnnualArchiveCanonicalizesDiagnosisSummaryHeader() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  headers[columns.DIAGNOSIS - 1] = '診斷｜CATA 4 | Retina 2';
  const sheet = new FakeSheet('202608', 2620, [headers]);
  const canonical = plain(call(
    'getCanonicalMonthlyHeaderValues_',
    sheet,
    columns
  ));
  assert.strictEqual(
    canonical[columns.DIAGNOSIS - 1],
    evaluate('CONFIG.MONTHLY_FIELD_HEADERS.DIAGNOSIS')
  );
  assert.strictEqual(
    sheet.valueAt(1, columns.DIAGNOSIS),
    '診斷｜CATA 4 | Retina 2'
  );
  assert.strictEqual(
    plain(call('getMonthlyCustomHeaders_', canonical))
      .includes('診斷｜CATA 4 | Retina 2'),
    false
  );

  const snapshot = {
    values: Array(canonical.length).fill(''),
    formulas: Array(canonical.length).fill(''),
    notes: Array(canonical.length).fill('')
  };
  snapshot.values[columns.DIAGNOSIS - 1] = 'CATA';
  const outputHeaders = plain(evaluate('CONFIG.MONTHLY_HEADERS'))
    .filter(Boolean);
  const mapped = call(
    'remapArchiveSnapshot_',
    snapshot,
    canonical,
    outputHeaders,
    '202608',
    'PATIENT'
  );
  assert.strictEqual(
    mapped.values[outputHeaders.indexOf('診斷')],
    'CATA'
  );
}

function testFuNameOnlyRowIsValid() {
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const row = Array(headers.length).fill('');
  row[headers.indexOf('姓名')] = '測試姓名';
  row[headers.indexOf('Condition')] = '追蹤';
  row[headers.indexOf('日期')] = new Date(2026, 7, 10);
  const sheet = new FakeSheet('FU', 101, [headers, row]);
  const columns = call('getRequiredFuColumns_', sheet);
  const built = call('buildFuRowContext_', sheet, 2, columns);
  assert.strictEqual(built.valid, true);
  assert.strictEqual(built.chartNo, '');
  assert.strictEqual(built.patientName, '測試姓名');
  assert.strictEqual(built.timeInfo.parsedTime, null);
}

function monthlyColumns() {
  const result = {};
  plain(evaluate('CONFIG.MONTHLY_FIELD_KEYS')).forEach(key => {
    const header = evaluate(`CONFIG.MONTHLY_FIELD_HEADERS.${key}`);
    const index = plain(evaluate('CONFIG.MONTHLY_HEADERS')).indexOf(header);
    result[key] = index + 1;
  });
  return result;
}

function testMonthlyNameOnlyRowIsValid() {
  const columns = monthlyColumns();
  const values = Array(23).fill('');
  values[columns.TIME - 1] = '0830';
  values[columns.NAME - 1] = '測試姓名';
  values[columns.DIAGNOSIS - 1] = 'CATA';
  const sheet = new FakeSheet('202608', 202608, [
    plain(evaluate('CONFIG.MONTHLY_HEADERS')),
    values
  ]);
  const patient = { row: 2, values };
  const block = {
    date: new Date(2026, 7, 10),
    dateKey: '2026-08-10',
    hospital: '高榮',
    blockKey: '2026-08-10|高榮'
  };
  const built = call(
    'buildMonthlyRowContext_',
    sheet,
    patient,
    block,
    columns,
    false
  );
  assert.strictEqual(built.valid, true);
  assert.strictEqual(built.chartNo, '');
  assert.strictEqual(built.patientName, '測試姓名');
}

function registryContext(overrides = {}) {
  return {
    eventId: 'event-1',
    sheetId: 1,
    sheetName: '202608',
    kind: 'MONTHLY',
    archived: false,
    row: 5,
    blockKey: '2026-08-10|高榮',
    rowHash: 'row-hash-1',
    bindingHash: 'binding-hash-1',
    valid: true,
    invalidReason: '',
    ...overrides
  };
}

function registryEntry(overrides = {}) {
  return {
    eventId: 'event-1',
    sheetId: 1,
    sheetName: '202608',
    kind: 'MONTHLY',
    archived: false,
    row: 3,
    blockKey: '2026-08-10|高榮',
    rowHash: 'row-hash-1',
    bindingHash: 'binding-hash-1',
    valid: true,
    invalidReason: '',
    pendingDelete: false,
    pendingResolution: false,
    ...overrides
  };
}

function registryScan(contexts, sheetRecords) {
  const eventLocations = {};
  contexts.forEach(item => {
    if (!item.eventId) return;
    if (!eventLocations[item.eventId]) eventLocations[item.eventId] = [];
    eventLocations[item.eventId].push(item);
  });
  return {
    contexts,
    sheetRecords: sheetRecords || [
      { sheetId: 1, sheetName: '202608', kind: 'MONTHLY', archived: false }
    ],
    eventLocations,
    duplicateIds: Object.keys(eventLocations)
      .filter(id => eventLocations[id].length > 1)
  };
}

function analyze(baseline, contexts, options = {}, sheetRecords) {
  return call(
    'analyzeRegistryDifferences_',
    baseline,
    registryScan(contexts, sheetRecords),
    options
  );
}

function testRegistrySameBlockRowReorderNeedsNoApi() {
  const result = analyze(
    [registryEntry()],
    [registryContext({ row: 20 })],
    { changeType: 'OTHER' }
  );
  assert.strictEqual(result.actions.length, 0);
  assert.strictEqual(result.conflicts.length, 0);
}

function testRegistryMoveUpdatesExistingEvent() {
  const result = analyze(
    [registryEntry()],
    [registryContext({ blockKey: '2026-08-24|高榮' })],
    { changeType: 'OTHER' }
  );
  assert.deepStrictEqual(plain(result.actions.map(item => item.type)), ['update']);
  assert.strictEqual(result.actions[0].placementChanged, true);
}

function testRegistryPendingSyncForcesRetryWithoutContentChange() {
  const result = analyze(
    [registryEntry({ pendingSync: true, pendingError: 'update_failed' })],
    [registryContext()],
    { changeType: 'OTHER' }
  );
  assert.deepStrictEqual(
    plain(result.actions.map(item => item.type)),
    ['update']
  );
  assert.strictEqual(result.actions[0].pendingSync, true);
}

function testRegistryCrossMonthMoveUpdatesExistingEvent() {
  const result = analyze(
    [registryEntry()],
    [
      registryContext({
        sheetId: 2,
        sheetName: '202609',
        blockKey: '2026-09-14|高榮'
      })
    ],
    { changeType: 'OTHER' },
    [
      { sheetId: 1, sheetName: '202608', kind: 'MONTHLY', archived: false },
      { sheetId: 2, sheetName: '202609', kind: 'MONTHLY', archived: false }
    ]
  );
  assert.deepStrictEqual(plain(result.actions.map(item => item.type)), ['update']);
}

function testRegistryMissingIdNewRowCreatesEvent() {
  const result = analyze(
    [],
    [registryContext({ eventId: '' })],
    { changeType: 'EDIT' }
  );
  assert.deepStrictEqual(plain(result.actions.map(item => item.type)), ['create']);
}

function testRegistryManualIdClearIsConflictNotDuplicateCreate() {
  const result = analyze(
    [registryEntry()],
    [registryContext({ eventId: '' })],
    { changeType: 'EDIT' }
  );
  assert.strictEqual(result.actions.length, 0);
  assert.ok(
    result.conflicts.some(item => item.type === 'event_id_removed_or_partial_move')
  );
}

function testRegistryNativeRowDeleteDeletesEvent() {
  const result = analyze(
    [registryEntry()],
    [],
    { changeType: 'REMOVE_ROW' }
  );
  assert.deepStrictEqual(
    plain(result.actions.map(item => item.type)),
    ['delete_missing']
  );
}

function testRegistryAmbiguousMissingRowWaitsForConfirmation() {
  const result = analyze(
    [registryEntry()],
    [],
    { changeType: 'OTHER' }
  );
  assert.strictEqual(result.actions.length, 0);
  assert.ok(result.conflicts.some(item => item.type === 'missing_registry_row'));
}

function testRegistryWholeSheetDeletionNeverDeletesEvents() {
  const result = analyze(
    [registryEntry()],
    [],
    { changeType: 'REMOVE_GRID', allowMissingDeletes: true },
    []
  );
  assert.strictEqual(result.actions.length, 0);
  assert.strictEqual(result.ignoredRemovedSheets.length, 1);
}

function testRegistryHeaderDeletionDoesNotBatchReschedule() {
  const result = analyze(
    [registryEntry()],
    [registryContext({ blockKey: '2026-08-24|高榮' })],
    { changeType: 'REMOVE_ROW' }
  );
  assert.strictEqual(result.actions.length, 0);
  assert.ok(
    result.conflicts.some(item => item.type === 'possible_date_header_deleted')
  );
}

function testRegistryDuplicateEventIdIsDiagnostic() {
  const contexts = [
    registryContext({ row: 5 }),
    registryContext({ row: 8 })
  ];
  const result = analyze(
    [registryEntry()],
    contexts,
    { changeType: 'OTHER' }
  );
  assert.strictEqual(result.actions.length, 0);
  assert.strictEqual(
    result.conflicts.filter(item => item.type === 'duplicate_event_id').length,
    2
  );
}

function testActiveAndArchivedDuplicateEventIdIsDiagnostic() {
  const active = registryContext({ eventId: 'shared-history-event' });
  const archived = {
    eventId: 'shared-history-event',
    sheetId: 900,
    sheetName: '刀表封存_2026',
    row: 12,
    kind: 'ARCHIVE',
    archived: true
  };
  const scan = registryScan([active]);
  scan.eventLocations['shared-history-event'].push(archived);
  scan.duplicateIds = ['shared-history-event'];
  const result = call(
    'analyzeRegistryDifferences_',
    [registryEntry({ eventId: 'shared-history-event' })],
    scan,
    { changeType: 'OTHER' }
  );
  assert.ok(result.conflicts.some(item => {
    return item.type === 'duplicate_event_id';
  }));
}

function testArchiveOnlyDuplicateEventIdIsDiagnostic() {
  const archiveLocation = row => ({
    eventId: 'duplicate-archive-event',
    sheetId: 900,
    sheetName: '刀表封存_2026',
    row,
    kind: 'ARCHIVE',
    archived: true
  });
  const scan = registryScan([]);
  scan.eventLocations['duplicate-archive-event'] = [
    archiveLocation(10),
    archiveLocation(20)
  ];
  scan.duplicateIds = ['duplicate-archive-event'];
  const result = call(
    'analyzeRegistryDifferences_',
    [],
    scan,
    { changeType: 'HEALTH' }
  );
  assert.ok(result.conflicts.some(item => {
    return item.type === 'duplicate_archive_event_id';
  }));
}

function testArchivedRowsDoNotSync() {
  const result = analyze(
    [registryEntry({ archived: true })],
    [
      registryContext({
        archived: true,
        blockKey: '2026-08-24|高榮',
        rowHash: 'changed'
      })
    ],
    { changeType: 'OTHER' }
  );
  assert.strictEqual(result.actions.length, 0);
}

function testInvalidTimeIsConflict() {
  const result = analyze(
    [registryEntry()],
    [
      registryContext({
        valid: false,
        invalidReason: 'invalid_time',
        timeInfo: { errorMessage: '時間範圍錯誤' }
      })
    ],
    { changeType: 'EDIT' }
  );
  assert.ok(result.conflicts.some(item => item.type === 'invalid_time'));
}

function testAnnualArchiveNamingAndEligibility() {
  assert.strictEqual(
    call('getAnnualArchiveName_', '2026'),
    '刀表封存_2026'
  );
  assert.strictEqual(call('isClosedMonthForArchive_', '202604', '202607'), true);
  assert.strictEqual(call('isClosedMonthForArchive_', '202607', '202607'), false);
  assert.strictEqual(call('isClosedMonthForArchive_', '202608', '202607'), false);
  assert.strictEqual(call('isClosedMonthForArchive_', '202613', '202701'), false);
}

function testConditionParserProvidesSafeSuggestions() {
  const result = call(
    'parseConditionSuggestion_',
    'CATA(NS++) OS s/p Phaco-IOL(EMV T-0.2 +20.0) OS 20260810'
  );
  assert.strictEqual(result.suggestion.side, 'OS');
  assert.strictEqual(result.suggestion.diagnosis, 'CATA');
  assert.strictEqual(result.suggestion.grade, 'NS++');
  assert.ok(result.suggestion.procedure.includes('Phaco-IOL'));
}

function testTimeSortIsStable() {
  const columns = { TIME: 1 };
  const rows = [
    { values: ['PM'], originalIndex: 1 },
    { values: ['0830'], originalIndex: 2 },
    { values: [''], originalIndex: 3 },
    { values: ['0800'], originalIndex: 4 },
    { values: ['待通知'], originalIndex: 5 }
  ];
  rows.sort((left, right) => call(
    'compareMonthlyPatientRows_',
    left,
    right,
    columns
  ));
  assert.deepStrictEqual(
    rows.map(item => item.values[0]),
    ['0800', '0830', 'PM', '待通知', '']
  );
}

function testFuDateSortKeyUsesAscendingDatesAndBlankLast() {
  const values = [
    { value: '', index: 0 },
    { value: new Date(2026, 7, 20), index: 1 },
    { value: new Date(2026, 7, 10), index: 2 },
    { value: new Date(2026, 7, 10), index: 3 }
  ];
  values.sort((left, right) => {
    const leftKey = call('getFuDateSortKey_', left.value, left.index);
    const rightKey = call('getFuDateSortKey_', right.value, right.index);
    return leftKey.group - rightKey.group ||
      leftKey.timestamp - rightKey.timestamp ||
      leftKey.originalIndex - rightKey.originalIndex;
  });
  assert.deepStrictEqual(values.map(item => item.index), [2, 3, 1, 0]);
}

function testFuTagOptions() {
  assert.deepStrictEqual(
    plain(evaluate('CONFIG.FU_TAG_OPTIONS')),
    [
      'FU',
      'Complication',
      'CATA',
      'Cornea',
      'Retina',
      'Plasty',
      'Neuro',
      'Refraction',
      'Suture IOL',
      'IVI'
    ]
  );
}

function testColumnWidthsAndFonts() {
  assert.strictEqual(evaluate('MANAGED_FONT_FAMILY'), 'Arial');
  assert.strictEqual(evaluate('MANAGED_FONT_SIZE'), 10);
  assert.deepStrictEqual(
    plain(evaluate(`({
      TIME: MONTHLY_COLUMN_WIDTHS.TIME,
      HOSPITAL: MONTHLY_COLUMN_WIDTHS.HOSPITAL,
      CHART_NO: MONTHLY_COLUMN_WIDTHS.CHART_NO,
      NAME: MONTHLY_COLUMN_WIDTHS.NAME,
      TEL: MONTHLY_COLUMN_WIDTHS.TEL,
      GA: MONTHLY_COLUMN_WIDTHS.GA,
      SIDE: MONTHLY_COLUMN_WIDTHS.SIDE,
      GRADE: MONTHLY_COLUMN_WIDTHS.GRADE,
      IOL: MONTHLY_COLUMN_WIDTHS.IOL,
      IOL_TARGET: MONTHLY_COLUMN_WIDTHS.IOL_TARGET,
      IOL_FINAL: MONTHLY_COLUMN_WIDTHS.IOL_FINAL,
      AXIS: MONTHLY_COLUMN_WIDTHS.AXIS
    })`)),
    {
      TIME: 115,
      HOSPITAL: 58,
      CHART_NO: 76,
      NAME: 72,
      TEL: 96,
      GA: 42,
      SIDE: 46,
      GRADE: 82,
      IOL: 96,
      IOL_TARGET: 80,
      IOL_FINAL: 76,
      AXIS: 52
    }
  );
  assert.deepStrictEqual(
    plain(evaluate(`({
      HOSPITAL: FU_COLUMN_WIDTHS.HOSPITAL,
      CHART_NO: FU_COLUMN_WIDTHS.CHART_NO,
      NAME: FU_COLUMN_WIDTHS.NAME,
      TEL: FU_COLUMN_WIDTHS.TEL
    })`)),
    { HOSPITAL: 58, CHART_NO: 76, NAME: 72, TEL: 96 }
  );
  assert.strictEqual(evaluate('MONTHLY_COLUMN_WIDTHS.PLAN'), 260);
  assert.strictEqual(evaluate('FU_COLUMN_WIDTHS.COND'), 320);
  assert.ok(plain(evaluate('MONTHLY_WRAP_KEYS')).includes('GRADE'));
}

function testExplicitColumnWidthsPreserveCustomColumns() {
  const headers = ['自訂補充'].concat(plain(evaluate('CONFIG.HEADERS')));
  const sheet = new FakeSheet('FU', 2400, [
    headers,
    Array(headers.length).fill('')
  ]);
  sheet.columnWidths[1] = 211;
  const columns = call('getRequiredFuColumns_', sheet);
  call('applyColumnWidths_', sheet, columns, plain(evaluate(
    'FU_COLUMN_WIDTHS'
  )));
  assert.strictEqual(sheet.columnWidths[1], 211);
  assert.strictEqual(sheet.columnWidths[columns.HOSPITAL], 58);
  assert.strictEqual(sheet.columnWidths[columns.CHART_NO], 76);
  assert.strictEqual(sheet.columnWidths[columns.NAME], 72);
  assert.strictEqual(sheet.columnWidths[columns.TEL], 96);
  assert.strictEqual(
    sheet.calls.setColumnWidth.some(item => item.column === 1),
    false
  );
}

function testMonthlyClinicalIdentifiersDefaultToPlainText() {
  assert.deepStrictEqual(
    plain(evaluate('MONTHLY_TEXT_KEYS')),
    [
      'TEL',
      'IOL',
      'IOL_TARGET',
      'IOL_FINAL',
      'AXIS',
      'SN',
      'CDE',
      'ENERGY_TIME',
      'ENERGY_PERCENT'
    ]
  );
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const sheet = new FakeSheet('202608', 2401, [
    headers,
    Array(headers.length).fill('')
  ]);
  const columns = call('getRequiredMonthlyColumns_', sheet);
  call('applyManagedRowFormat_', sheet, 2, 'MONTHLY', columns);
  assert.strictEqual(sheet.calls.numberFormatRangeLists.length, 1);
  const textFormat = sheet.calls.numberFormatRangeLists[0];
  assert.strictEqual(textFormat.format, '@');
  assert.strictEqual(textFormat.addresses.length, 9);
  [
    'E2',
    'K2',
    'L2',
    'M2',
    'N2',
    'R2',
    'S2',
    'T2',
    'U2'
  ].forEach(address => assert.ok(textFormat.addresses.includes(address)));

  const source = sourceByFile['sheet_model.js'];
  const monthlyFormattingBody = source.slice(
    source.indexOf('function applyMonthlyFormatting_'),
    source.indexOf('function applyManagedRowFormat_')
  );
  assert.ok(monthlyFormattingBody.includes(
    'applyMonthlyTextNumberFormats_(sheet, columns, 2, dataRowCount)'
  ));
}

function testMonthlyTextFormatsOnlyInitializeSystemCreatedRows() {
  assert.strictEqual(
    evaluate('typeof repairMonthlyTextInputFormats_'),
    'undefined'
  );
  const deployedSource = files.map(file => sourceByFile[file]).join('\n');
  assert.strictEqual(
    deployedSource.includes('repairMonthlyTextInputFormats_'),
    false
  );

  const codeSource = sourceByFile['code.js'];
  const onOpenBody = codeSource.slice(
    codeSource.indexOf('function onOpen()'),
    codeSource.indexOf("const maintenance = ui.createMenu('維護工具')")
  );
  assert.ok(onOpenBody.includes('SpreadsheetApp.getUi()'));
  assert.strictEqual(onOpenBody.includes('getActiveSpreadsheet'), false);

  const calendarSource = sourceByFile['calendar_sync.js'];
  const structureBody = calendarSource.slice(
    calendarSource.indexOf('function processCalendarStructureChange'),
    calendarSource.indexOf('function buildCalendarSyncRuntime_')
  );
  assert.ok(structureBody.includes("allowMissingDeletes: changeType === 'REMOVE_ROW'"));
  assert.strictEqual(structureBody.includes('setNumberFormat'), false);

  const sheetSource = sourceByFile['sheet_model.js'];
  const monthlyFormattingBody = sheetSource.slice(
    sheetSource.indexOf('function applyMonthlyFormatting_'),
    sheetSource.indexOf('function applyManagedRowFormat_')
  );
  assert.ok(monthlyFormattingBody.includes('settings.initializeTextInputs'));
  const touchedRowsBody = sheetSource.slice(
    sheetSource.indexOf('function applyTouchedManagedRowsFormatBatch_'),
    sheetSource.indexOf('function applyTouchedManagedRowsFormat_')
  );
  assert.ok(touchedRowsBody.includes("const textKeys = ['TIME']"));
  assert.strictEqual(touchedRowsBody.includes('MONTHLY_TEXT_KEYS'), false);

  const workflowSource = sourceByFile['workflows.js'];
  const createMonthlyBody = workflowSource.slice(
    workflowSource.indexOf('function createOrInitializeMonthlySheet_'),
    workflowSource.indexOf('function createMonthlySurgerySheet')
  );
  assert.ok(createMonthlyBody.includes('initializeTextInputs: created'));

  const appendBlockBody = sheetSource.slice(
    sheetSource.indexOf('function appendMonthlyBlock_'),
    sheetSource.indexOf('function ensureDefaultMonthlyBlocks_')
  );
  assert.ok(appendBlockBody.includes('applyManagedRowsFormatBatch_'));
  const insertPatientBody = sheetSource.slice(
    sheetSource.indexOf('function insertPatientAtBlockEnd_'),
    sheetSource.indexOf('function getMonthlyBlockOptions_')
  );
  assert.ok(insertPatientBody.includes('insertedBlankRows'));
  assert.ok(insertPatientBody.includes('applyManagedRowsFormatBatch_'));
}

function testMonthlyIolTextInputsHaveNoPostEditRewrite() {
  const sheetSource = sourceByFile['sheet_model.js'];
  const calendarSource = sourceByFile['calendar_sync.js'];
  const onEditBody = calendarSource.slice(
    calendarSource.indexOf('function processRowChange'),
    calendarSource.indexOf('function processCalendarStructureChange')
  );
  assert.strictEqual(
    sheetSource.includes('normalizeMonthlySignedTextInputs_'),
    false
  );
  assert.strictEqual(
    sheetSource.includes('normalizeMonthlySignedTextValue_'),
    false
  );
  assert.strictEqual(onEditBody.includes('IOL_SIGNED'), false);
  assert.strictEqual(
    onEditBody.includes('normalizeMonthlySignedTextInputs_'),
    false
  );
}

function testMonthlyHeaderConditionalFormatUsesExactMarker() {
  const expectedMarker = evaluate('MONTHLY_DATE_HEADER_MARKER');
  assert.strictEqual(
    call('buildMonthlyHeaderConditionalPredicate_', { CHART_NO: 3 }),
    `$C2="${expectedMarker}"`
  );
  assert.strictEqual(
    call('buildMonthlyHeaderConditionalPredicate_', { CHART_NO: 27 }),
    `$AA2="${expectedMarker}"`
  );

  const columns = monthlyColumns();
  ['PM', 'AM', '待通知', ''].forEach((time, index) => {
    const row = Array(plain(evaluate('CONFIG.MONTHLY_HEADERS')).length).fill('');
    row[columns.TIME - 1] = time;
    row[columns.CHART_NO - 1] = `TEST-${index + 1}`;
    row[columns.PLAN - 1] = ['#', '!', 'APPLY', '#'][index];
    assert.strictEqual(
      call('getMonthlyHeaderMarkerState_', row, columns).valid,
      false
    );
  });
  const headerRow = Array(plain(evaluate('CONFIG.MONTHLY_HEADERS')).length).fill('');
  headerRow[columns.CHART_NO - 1] = expectedMarker;
  headerRow[columns.PLAN - 1] = '#!APPLY';
  assert.strictEqual(
    call('getMonthlyHeaderMarkerState_', headerRow, columns).valid,
    true
  );

  const source = sourceByFile['sheet_model.js'];
  const body = source.slice(
    source.indexOf('function applyMonthlyConditionalFormats_'),
    source.indexOf('function applyFuFormatting_')
  );
  assert.ok(body.includes(
    'buildMonthlyHeaderConditionalPredicate_(columns)'
  ));
  assert.strictEqual(body.includes('timeLetter'), false);
  assert.strictEqual(body.includes('eventIdLetter'), false);
  assert.strictEqual(body.includes('YEAR('), false);
  assert.strictEqual(body.includes('MOD('), false);
  assert.strictEqual(body.includes('patientFieldReferences'), false);
  assert.strictEqual(body.includes('COUNTA('), false);
  assert.strictEqual(
    (body.match(/NOT\(\$\{headerPredicate\}\)/g) || []).length,
    6
  );
  assert.ok(body.includes('`=AND(${headerPredicate},`'));
  assert.ok(body.includes('`=AND(NOT(${headerPredicate}),${redMatch},`'));
  assert.ok(body.includes('`=AND(NOT(${headerPredicate}),`'));
  assert.ok(body.includes('.setBold(true)'));
  assert.strictEqual(body.includes(".setFontWeight('bold')"), false);
}

function testPlanConditionalFormatMarkersAndPriority() {
  assert.deepStrictEqual(
    plain(evaluate('CONFIG.PLAN_RED_KEYWORDS')),
    ['!']
  );
  assert.deepStrictEqual(
    plain(evaluate('CONFIG.PLAN_GREEN_KEYWORDS')),
    ['APPLY']
  );
  assert.deepStrictEqual(
    plain(evaluate('CONFIG.PLAN_YELLOW_KEYWORDS')),
    ['#']
  );
  assert.strictEqual(evaluate('PLAN_RED_BACKGROUND'), '#F4CCCC');
  assert.strictEqual(evaluate('PLAN_GREEN_BACKGROUND'), '#D9EAD3');
  assert.strictEqual(evaluate('PLAN_YELLOW_BACKGROUND'), '#FFF2CC');
  assert.strictEqual(
    call('buildPlanContainsAnyFormula_', '$H2', ['!']),
    'ISNUMBER(SEARCH("!",$H2))'
  );

  const sheetSource = sourceByFile['sheet_model.js'];
  assert.ok(sheetSource.includes('FU_PLAN_RED'));
  assert.ok(sheetSource.includes('MONTH_PLAN_RED'));
  assert.ok(sheetSource.includes('.setBackground(PLAN_RED_BACKGROUND)'));
  assert.ok(sheetSource.includes(
    '`=AND(NOT(${redMatch}),${greenMatch},`'
  ));
  assert.ok(sheetSource.includes(
    '`=AND(NOT(${redMatch}),NOT(${greenMatch}),${yellowMatch},`'
  ));

  const archiveSource = sourceByFile['annual_archive.js'];
  assert.ok(archiveSource.includes("marker('PLAN_RED')"));
  assert.ok(archiveSource.includes('.setBackground(PLAN_RED_BACKGROUND)'));
}

function testConditionalFormatCleanupRemovesCurrentAndLegacySystemRules() {
  const body = sourceByFile['sheet_model.js'].slice(
    sourceByFile['sheet_model.js'].indexOf(
      'function removeSystemConditionalFormats_'
    ),
    sourceByFile['sheet_model.js'].indexOf(
      'function buildVisibleRangesExcludingColumns_'
    )
  );
  assert.ok(body.includes('CONDITIONAL_FORMAT_MARKER_PREFIX'));
  assert.ok(body.includes('LEGACY_CONDITIONAL_FORMAT_MARKER_PREFIX'));
}

function testPermanentArchiveIsPrivateTransactionalAndCalendarFree() {
  const source = sourceByFile['annual_archive.js'];
  assert.ok(source.includes('DriveApp.Access.PRIVATE'));
  assert.ok(source.includes('makeCopy(name)'));
  assert.ok(source.includes('spreadsheet.deleteSheet(sheet)'));
  assert.ok(
    sourceByFile['workflows.js'].includes('ANNUAL_ARCHIVE_TRANSACTION_PROPERTY')
  );
  assert.ok(source.includes('ArchiveSourceMonth'));
  assert.ok(source.includes('ArchiveRowType'));
  assert.ok(source.includes('sourceSheetIds'));
  assert.strictEqual(source.includes('Calendar.Events'), false);
  assert.strictEqual(source.includes('hideSheet();\n      archiveMap'), false);
}

function testAnnualArchiveStructureUsesProtectedSystemMetadata() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'))
    .filter(Boolean)
    .concat(['ArchiveSourceMonth', 'ArchiveRowType']);
  const rows = Array.from({ length: 9 }, () => Array(headers.length).fill(''));
  rows[5] = headers.slice();
  const timeColumn = headers.indexOf('日期／報到時間');
  const eventColumn = headers.indexOf('CalendarEventId');
  const sourceColumn = headers.indexOf('ArchiveSourceMonth');
  const typeColumn = headers.indexOf('ArchiveRowType');
  rows[6][timeColumn] = '202604';
  rows[6][sourceColumn] = '202604';
  rows[6][typeColumn] = 'MONTH_HEADER';
  rows[7][timeColumn] = new Date(2026, 3, 13);
  rows[7][sourceColumn] = '202604';
  rows[7][typeColumn] = 'DATE_HEADER';
  rows[8][eventColumn] = 'archive-event-1';
  rows[8][sourceColumn] = '202604';
  rows[8][typeColumn] = 'PATIENT';
  const sheet = new FakeSheet('刀表封存_2026', 2600, rows);
  const result = call('scanAnnualArchiveStructure_', sheet);
  assert.deepStrictEqual(plain(result.months), ['202604']);
  assert.strictEqual(result.monthRanges[0].blockCount, 1);
  assert.strictEqual(result.monthRanges[0].patientCount, 1);
  assert.strictEqual(result.eventLocations[0].eventId, 'archive-event-1');
}

function testLegacyArchiveStateFailsClosed() {
  scriptPropertyStore.ARCHIVED_MONTHLY_SHEETS_V1 =
    JSON.stringify({ 123: { name: '202604' } });
  assert.throws(
    () => call('assertNoLegacyArchiveState_'),
    /舊版/
  );
  delete scriptPropertyStore.ARCHIVED_MONTHLY_SHEETS_V1;
  scriptPropertyStore.ARCHIVED_MONTHLY_SHEETS_V1 = '{broken';
  assert.throws(
    () => call('assertNoLegacyArchiveState_'),
    /損壞/
  );
  delete scriptPropertyStore.ARCHIVED_MONTHLY_SHEETS_V1;
}

function testArchiveTransactionGuardFailsClosed() {
  scriptPropertyStore.ANNUAL_SURGERY_ARCHIVE_TRANSACTION_V1 =
    JSON.stringify({ operationId: 'pending-test', months: ['202604'] });
  assert.strictEqual(call('isAnnualArchiveTransactionActive_'), true);
  assert.throws(
    () => call('assertNoActiveAnnualArchiveTransaction_'),
    /尚未完成/
  );
  delete scriptPropertyStore.ANNUAL_SURGERY_ARCHIVE_TRANSACTION_V1;
}

function testFuScanIsBatchedAndRetryCacheRelocatesByFingerprint() {
  const calendarSource = sourceByFile['calendar_sync.js'];
  const scanBody = calendarSource.slice(
    calendarSource.indexOf('function buildManagedSheetCalendarScan_'),
    calendarSource.indexOf('function buildGlobalEventIdLocationsLightweight_')
  );
  assert.ok(scanBody.includes(
    'sheet.getRange(2, 1, rowCount, lastColumn).getValues()'
  ));
  assert.strictEqual(scanBody.includes('sheet.getRange(row, 1'), false);
  const workflowSource = sourceByFile['workflows.js'];
  assert.ok(workflowSource.includes('function locateCachedMonthlySchedule_'));
  ['rowHash', 'blockKey', 'eventId'].forEach(field => {
    assert.ok(workflowSource.includes(field));
  });
}

function testColumnWidthsRequireExplicitRequestAfterCreation() {
  const source = sourceByFile['sheet_model.js'];
  const body = source.slice(
    source.indexOf('function applyColumnWidths_'),
    source.indexOf('function buildListValidation_')
  );
  assert.strictEqual(body.includes('PropertiesService'), false);
  assert.strictEqual(
    source.includes('COLUMN_WIDTH_PROFILE_PROPERTY_PREFIX'),
    false
  );
  assert.strictEqual(
    source.includes('COLUMN_WIDTH_PROFILE_VERSION'),
    false
  );
  const fuFormat = source.slice(
    source.indexOf('function applyFuFormatting_'),
    source.indexOf('function applyMonthlyFormatting_')
  );
  const monthFormat = source.slice(
    source.indexOf('function applyMonthlyFormatting_'),
    source.indexOf('function applyManagedRowFormat_')
  );
  assert.ok(fuFormat.includes('settings.forceWidths'));
  assert.ok(monthFormat.includes('settings.forceWidths'));
  const initFu = source.slice(
    source.indexOf('function initializeMainTrackingSheet_'),
    source.indexOf('function applyAllRecommendedColumnWidths')
  );
  assert.ok(initFu.includes('forceWidths: created'));
  const workflowSource = sourceByFile['workflows.js'];
  const initMonth = workflowSource.slice(
    workflowSource.indexOf('function createOrInitializeMonthlySheet_'),
    workflowSource.indexOf('function createMonthlySurgerySheet')
  );
  assert.ok(initMonth.includes('forceWidths: created'));
}

function testInstallDoesNotBatchRewriteCalendar() {
  const setupBody = sourceByFile['code.js'].slice(
    sourceByFile['code.js'].indexOf('function setup()'),
    sourceByFile['code.js'].indexOf('function assertNoLegacyOpOnly_')
  );
  assert.strictEqual(setupBody.includes('batchSync'), false);
  assert.strictEqual(setupBody.includes('Calendar.Events'), false);
  assert.ok(setupBody.includes('rebuildCalendarRegistry_'));
  assert.ok(
    setupBody.indexOf('assertCalendarRegistrySafeToRebuild_') <
      setupBody.indexOf('rebuildCalendarRegistry_')
  );
  assert.ok(setupBody.includes('allowRepairableStructureActions: true'));
  assert.ok(setupBody.includes('normalizeTimes: false'));
  const preflightBody = sourceByFile['calendar_sync.js'].slice(
    sourceByFile['calendar_sync.js'].indexOf(
      'function assertCalendarRegistrySafeToRebuild_'
    ),
    sourceByFile['calendar_sync.js'].indexOf('function isContextAffected_')
  );
  assert.ok(preflightBody.includes('ignoredRepairableActions'));
  assert.ok(preflightBody.includes('ignoredRepairableConflicts'));
}

function testSystemNotesPreserveManualNotes() {
  const sheet = new FakeSheet('FU', 40, [['病歷號'], ['X001']]);
  const cell = sheet.getRange(2, 1);
  cell.setNote('人工註記');
  const prefix = evaluate('CALENDAR_SYNC_NOTE_PREFIX');

  call('setSystemNote_', cell, prefix, '第一次失敗');
  assert.ok(cell.getNote().includes('人工註記'));
  assert.ok(cell.getNote().includes('第一次失敗'));

  call('setSystemNote_', cell, prefix, '第二次失敗');
  assert.ok(cell.getNote().includes('人工註記'));
  assert.ok(cell.getNote().includes('第二次失敗'));
  assert.strictEqual(cell.getNote().includes('第一次失敗'), false);

  call('clearSystemNote_', cell, [prefix]);
  assert.strictEqual(cell.getNote(), '人工註記');

  cell.setNote(`${prefix}舊版訊息\n人工後記`);
  call('clearSystemNote_', cell, [prefix]);
  assert.strictEqual(cell.getNote(), '人工後記');
}

function testCalendarErrorsUseVisibleIdentityCell() {
  const sheet = new FakeSheet('FU', 41, [
    ['病歷號', '姓名', '時間', 'CalendarEventId'],
    ['X001', '測試姓名', '', 'event-secret']
  ]);
  const managed = {
    sheet,
    row: 2,
    chartNo: 'X001',
    patientName: '測試姓名',
    columns: {
      CHART_NO: 1,
      NAME: 2,
      TIME: 3,
      EVENT_ID: 4
    }
  };
  sheet.getRange(2, 1).setNote('人工註記');
  sheet.getRange(2, 4).setNote(
    `${evaluate('CALENDAR_SYNC_NOTE_PREFIX')}舊版隱藏錯誤`
  );

  call(
    'setContextSyncNote_',
    managed,
    '更新 Calendar 事件失敗：event-secret',
    false
  );

  const visibleNote = sheet.getRange(2, 1).getNote();
  assert.ok(visibleNote.includes('人工註記'));
  assert.ok(visibleNote.includes('更新 Calendar 事件失敗'));
  assert.ok(visibleNote.includes('檢查同步健康'));
  assert.strictEqual(visibleNote.includes('event-secret'), false);
  assert.strictEqual(sheet.getRange(2, 4).getNote(), '');

  const nameOnly = {
    ...managed,
    chartNo: '',
    patientName: '測試姓名'
  };
  assert.strictEqual(call('getContextNoteCell_', nameOnly).getColumn(), 2);

  sheet.getRange(2, 3).setNote('人工時間註記');
  const invalidTime = {
    ...managed,
    archived: false,
    eventId: '',
    invalidReason: 'invalid_time',
    timeInfo: { errorMessage: '時間範圍錯誤。' }
  };
  const invalidResult = call('syncManagedContext_', invalidTime, {});
  assert.strictEqual(invalidResult.status, 'invalid_time');
  assert.strictEqual(sheet.getRange(2, 1).getNote(), '人工註記');
  assert.ok(sheet.getRange(2, 3).getNote().includes('人工時間註記'));
  assert.ok(sheet.getRange(2, 3).getNote().includes('時間範圍錯誤'));
}

function testHealthRowsRedactIdsAndSeparateFailures() {
  const conflictRows = call('getHealthIssueRows_', {
    analysis: {
      actions: [],
      conflicts: [{
        type: 'duplicate_event_id',
        eventId: 'event-secret',
        context: { sheetName: 'FU', row: 2 },
        message: 'CalendarEventId event-secret 出現在多列。'
      }]
    },
    results: []
  });
  assert.strictEqual(conflictRows[0][1], '衝突');
  assert.strictEqual(conflictRows[0][6].includes('event-secret'), false);

  const failureRows = call('getHealthIssueRows_', {
    analysis: {
      actions: [{
        type: 'update',
        eventId: 'event-secret',
        context: { sheetName: 'FU', row: 2 }
      }],
      conflicts: []
    },
    results: [{
      action: {
        type: 'update',
        eventId: 'event-secret',
        context: { sheetName: 'FU', row: 2 }
      },
      result: { ok: false, status: 'update_failed' }
    }]
  });
  assert.strictEqual(failureRows.length, 1);
  assert.strictEqual(failureRows[0][1], '失敗');
  assert.strictEqual(failureRows[0][2], 'update_failed');

  const healthySummary = call('buildHealthAlertSummary_', {
    analysis: { actions: [], conflicts: [] },
    results: []
  });
  assert.strictEqual(healthySummary.hasIssues, false);
  assert.ok(healthySummary.text.includes('未發現'));

  const appliedAction = {
    type: 'calendar_drift',
    eventId: 'event-drift',
    context: { sheetName: '202608', sheetId: 8, row: 5 }
  };
  const pendingAction = {
    type: 'calendar_orphan',
    eventId: 'event-orphan',
    previous: { sheetName: 'Calendar', row: '' }
  };
  const mixedRows = call('getHealthIssueRows_', {
    analysis: {
      actions: [appliedAction, pendingAction],
      conflicts: []
    },
    results: [{
      action: appliedAction,
      result: { ok: true, status: 'updated' }
    }]
  });
  assert.strictEqual(mixedRows.length, 1);
  assert.strictEqual(mixedRows[0][1], '待處理');
  assert.strictEqual(mixedRows[0][2], 'calendar_orphan');
  assert.ok(mixedRows[0][6].includes('等待確認刪除'));

  const conflictSummary = call('buildHealthAlertSummary_', {
    analysis: {
      actions: [],
      conflicts: [{
        type: 'duplicate_event_id',
        eventId: 'event-secret',
        context: { sheetName: 'FU', row: 2 },
        message: 'CalendarEventId event-secret 出現在多列。'
      }]
    },
    results: []
  }, 10);
  assert.strictEqual(conflictSummary.hasIssues, true);
  assert.ok(conflictSummary.text.includes('FU 第 2 列'));
  assert.strictEqual(conflictSummary.text.includes('event-secret'), false);
}

function testHealthRowsExposePendingQueueWithoutClinicalData() {
  const rows = call('getHealthIssueRows_', {
    analysis: { actions: [], conflicts: [] },
    results: [],
    pendingQueue: {
      version: 1,
      fullScan: false,
      allowMissingDeletes: false,
      sheets: { 88: [4, 9] },
      reasons: ['EDIT'],
      updatedAt: '2026-07-27T00:00:00.000Z'
    }
  });
  assert.strictEqual(rows.length, 1);
  assert.strictEqual(rows[0][1], '待重試');
  assert.strictEqual(rows[0][2], 'pending_queue');
  assert.ok(rows[0][6].includes('1 張表、2 列'));
  assert.strictEqual(rows[0][6].includes('虛構病人'), false);
}

function testHealthReportIsRetired() {
  const calendarSource = sourceByFile['calendar_sync.js'];
  const deployedSource = files.map(file => sourceByFile[file]).join('\n');
  [
    'SYNC_ARCHITECTURE_UPGRADE',
    'previewSyncArchitectureUpgrade',
    'runSyncArchitectureUpgrade',
    'backupLegacySyncProperties_'
  ].forEach(text => assert.strictEqual(calendarSource.includes(text), false));
  assert.strictEqual(deployedSource.includes('writeHealthReportSheet_'), false);
  assert.strictEqual(calendarSource.includes('sheet.showSheet()'), false);
  assert.strictEqual(calendarSource.includes('sheet.hideSheet()'), false);
  assert.strictEqual(deployedSource.includes('同步健康報告'), false);
  assert.strictEqual(
    evaluate('typeof inspectLegacyHealthReportSheet_'),
    'undefined'
  );
  assert.strictEqual(
    evaluate('typeof removeLegacySyncHealthReportSheet'),
    'undefined'
  );
  assert.strictEqual(evaluate('typeof buildHealthAlertSummary_'), 'function');
}
function testRegistryStoresRawIdButNoClinicalPlaintext() {
  const entry = call('toRegistryEntry_', registryContext());
  assert.strictEqual(entry.eventId, 'event-1');
  [
    'chartNo',
    'patientName',
    'condition',
    'plan',
    'tel'
  ].forEach(key => assert.strictEqual(Object.hasOwn(entry, key), false));
}

function clearScriptProperties() {
  Object.keys(scriptPropertyStore).forEach(key => {
    delete scriptPropertyStore[key];
  });
}

function installCalendarMock(options = {}) {
  const events = {};
  (options.initialEvents || []).forEach(event => {
    events[event.id] = JSON.parse(JSON.stringify(event));
  });
  let insertCount = 0;
  let updateCount = 0;
  let removeCount = 0;
  let getCount = 0;
  let listCount = 0;
  context.Calendar = {
    Events: {
      insert(resource) {
        insertCount++;
        const id = `created-${insertCount}`;
        events[id] = { id, ...JSON.parse(JSON.stringify(resource)) };
        return events[id];
      },
      update(resource, _calendarId, eventId) {
        updateCount++;
        if (options.update404 || !events[eventId]) {
          const error = new Error('404 not found');
          throw error;
        }
        events[eventId] = { id: eventId, ...JSON.parse(JSON.stringify(resource)) };
        return events[eventId];
      },
      get(_calendarId, eventId) {
        getCount++;
        if ((options.getErrorIds || []).includes(eventId)) {
          throw new Error(options.getErrorMessage || '503 read failed');
        }
        if (
          (options.get404Ids || []).includes(eventId) ||
          !events[eventId]
        ) {
          throw new Error('404 not found');
        }
        return events[eventId];
      },
      remove(_calendarId, eventId) {
        removeCount++;
        if (options.removeError) throw new Error(options.removeError);
        if (options.remove404 || !events[eventId]) throw new Error('404 not found');
        delete events[eventId];
      },
      list(_calendarId, listOptions = {}) {
        listCount++;
        if (options.listError) throw new Error(options.listError);
        const omitted = new Set(options.omitFromList || []);
        const ids = Object.keys(events)
          .filter(id => !omitted.has(id))
          .sort();
        const pageSize = options.pageSize || ids.length || 1;
        const start = Number(listOptions.pageToken || 0);
        const pageIds = ids.slice(start, start + pageSize);
        const next = start + pageSize < ids.length
          ? String(start + pageSize)
          : '';
        return {
          items: pageIds.map(id => events[id]),
          nextPageToken: next
        };
      }
    }
  };
  return {
    events,
    counts: () => ({
      insertCount,
      updateCount,
      removeCount,
      getCount,
      listCount
    })
  };
}

function liveCalendarContext(overrides = {}) {
  return {
    eventId: 'event-current',
    sheetId: 50,
    sheetName: '202608',
    kind: 'MONTHLY',
    archived: false,
    row: 2,
    blockKey: '2026-08-10|高榮',
    rowHash: 'live-row-hash',
    valid: true,
    invalidReason: '',
    chartNo: 'X001',
    patientName: '測試姓名',
    condition: 'CATA OS',
    hospital: '高榮',
    tel: '',
    plan: '',
    ga: '',
    date: new Date(2026, 7, 10),
    timeInfo: call('resolveCalendarTime_', '08:00'),
    ...overrides
  };
}

function testCalendarRepairOwnerRequiresUniqueIdentityAndSchedule() {
  const owner = liveCalendarContext({
    eventId: 'event-duplicate',
    sheetId: 50,
    row: 30,
    chartNo: 'X001',
    patientName: '測試姓名',
    condition: '最新診斷',
    date: new Date(2026, 7, 13),
    timeInfo: call('resolveCalendarTime_', '12:00')
  });
  const event = plain(call('buildCalendarResource_', owner));
  event.id = 'event-duplicate';
  event.summary = 'X001 | 測試姓名 | 舊診斷';

  assert.strictEqual(
    call('getUniqueCalendarRepairOwnerKey_', event, [owner]),
    '50:30'
  );

  const sameIdentityAndSchedule = {
    ...owner,
    row: 33,
    condition: '另一診斷'
  };
  assert.strictEqual(
    call(
      'getUniqueCalendarRepairOwnerKey_',
      event,
      [owner, sameIdentityAndSchedule]
    ),
    ''
  );

  const wrongSchedule = {
    ...owner,
    row: 34,
    timeInfo: call('resolveCalendarTime_', '12:30')
  };
  assert.strictEqual(
    call('getUniqueCalendarRepairOwnerKey_', event, [wrongSchedule]),
    ''
  );

  const wrongIdentity = {
    ...owner,
    row: 35,
    chartNo: 'X002'
  };
  assert.strictEqual(
    call('getUniqueCalendarRepairOwnerKey_', event, [wrongIdentity]),
    ''
  );
}

function testLiveCalendarAuditFindsColorDriftMissingAndOrphans() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const current = liveCalendarContext();
  const expected = call('buildCalendarResource_', current);
  const archiveEvent = {
    id: 'event-archive',
    summary: 'A001 | 歷史病例 | CATA',
    description: 'TEL: 000',
    colorId: '5',
    start: { date: '2026-01-01' },
    end: { date: '2026-01-02' }
  };
  const orphanEvent = {
    id: 'event-orphan',
    summary: 'O001 | 舊資料 | CATA',
    description: 'Plan: migrated',
    colorId: '5',
    start: { date: '2026-02-01' },
    end: { date: '2026-02-02' }
  };
  const api = installCalendarMock({
    initialEvents: [
      { id: current.eventId, ...expected, colorId: '2' },
      archiveEvent,
      orphanEvent,
      {
        id: 'event-manual',
        summary: '私人行程',
        description: '不屬於系統',
        start: { date: '2026-03-01' },
        end: { date: '2026-03-02' }
      }
    ],
    omitFromList: [current.eventId],
    pageSize: 2
  });
  const missing = liveCalendarContext({
    eventId: 'event-missing',
    row: 3,
    rowHash: 'missing-row-hash'
  });
  const scan = registryScan([current, missing], [{
    sheetId: 50,
    sheetName: '202608',
    kind: 'MONTHLY',
    archived: false
  }]);
  scan.eventLocations['event-archive'] = [{
    eventId: 'event-archive',
    sheetId: 90,
    sheetName: '刀表封存_2026',
    kind: 'ARCHIVE',
    archived: true,
    row: 10
  }];

  const result = call('analyzeLiveCalendarState_', scan);
  assert.deepStrictEqual(
    plain(result.actions.map(item => item.type).sort()),
    ['calendar_drift', 'calendar_missing', 'calendar_orphan']
  );
  assert.strictEqual(
    result.actions.find(item => item.type === 'calendar_orphan').eventId,
    'event-orphan'
  );
  assert.strictEqual(result.unmanagedEvents.length, 1);
  assert.strictEqual(result.conflicts.length, 0);
  assert.strictEqual(api.counts().getCount, 2);
  assert.ok(api.counts().listCount >= 2);
}

function testLiveCalendarAuditFailsClosedAndOrphanDeleteNeedsConfirmation() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  installCalendarMock({ listError: 'temporary list failure' });
  const failed = call(
    'analyzeLiveCalendarState_',
    registryScan([liveCalendarContext()])
  );
  assert.strictEqual(failed.ok, false);
  assert.strictEqual(failed.conflicts[0].type, 'calendar_list_failed');

  const orphan = {
    id: 'event-orphan',
    summary: 'O001 | 舊資料 | CATA',
    description: 'Plan: migrated',
    start: { date: '2026-02-01' },
    end: { date: '2026-02-02' }
  };
  const api = installCalendarMock({ initialEvents: [orphan] });
  const action = {
    type: 'calendar_orphan',
    eventId: orphan.id,
    context: null,
    previous: { sheetName: 'Calendar', row: '' }
  };
  const scan = registryScan([]);
  const analysis = { actions: [action], conflicts: [] };
  assert.strictEqual(
    call('applyLiveCalendarActions_', scan, analysis, []).length,
    0
  );
  assert.ok(api.events[orphan.id]);
  const applied = call(
    'applyLiveCalendarActions_',
    scan,
    analysis,
    [orphan.id]
  );
  assert.strictEqual(applied.length, 1);
  assert.strictEqual(applied[0].result.ok, true);
  assert.strictEqual(api.events[orphan.id], undefined);
}

function testCalendarLifecycleCreateAndVerify() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const api = installCalendarMock();
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const row = Array(headers.length).fill('');
  row[headers.indexOf('姓名')] = '測試姓名';
  row[headers.indexOf('日期')] = new Date(2026, 7, 10);
  row[headers.indexOf('Condition')] = '追蹤';
  const sheet = new FakeSheet('FU', 50, [headers, row]);
  const columns = call('getRequiredFuColumns_', sheet);
  const managed = call('buildFuRowContext_', sheet, 2, columns);
  const result = call('syncManagedContext_', managed, {});
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.status, 'created');
  assert.strictEqual(api.counts().insertCount, 1);
  assert.strictEqual(
    sheet.valueAt(2, headers.indexOf('CalendarEventId') + 1),
    'created-1'
  );
}

function makeMonthlyReorderFixture() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  // Deliberately reorder columns to exercise header-based lookup.
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS')).reverse();
  const makeRow = fields => headers.map(header => fields[header] || '');
  const sheet = new FakeSheet('202609', 909, [
    headers,
    makeRow({ '日期／報到時間': new Date(2026, 8, 10), '醫院': '高榮', '病歷號': '◆ 刀日' }),
    makeRow({ '病歷號': 'TEST-A', '姓名': '測試甲', CalendarEventId: 'event-a' }),
    makeRow({ '病歷號': 'TEST-B', '姓名': '測試乙', CalendarEventId: 'event-b' })
  ]);
  const spreadsheet = makeFakeSpreadsheet([sheet]);
  const columns = call('getRequiredMonthlyColumns_', sheet);
  const scan = call('buildCurrentCalendarScan_', spreadsheet);
  call('writeCalendarRegistryStore_', scan);
  const api = installCalendarMock({ initialEvents: scan.contexts.map(item => ({
    ...plain(call('buildCalendarResource_', item)), id: item.eventId
  })) });
  return { sheet, spreadsheet, columns, scan, api };
}

function testMonthlyMoveDuringCalendarUpdateDoesNotCopyEventId() {
  const { sheet, columns, scan, api } = makeMonthlyReorderFixture();
  const originalUpdate = context.Calendar.Events.update;
  context.Calendar.Events.update = (...args) => {
    const result = originalUpdate(...args);
    [sheet.rows[2], sheet.rows[3]] = [sheet.rows[3], sheet.rows[2]];
    sheet.setNoteAt(3, columns.CHART_NO, '另一列的人工註記');
    call('setSystemNote_', sheet.getRange(3, columns.CHART_NO),
      evaluate('CALENDAR_SYNC_NOTE_PREFIX'), '另一列仍待重試');
    return result;
  };
  try {
    const result = call('syncManagedContext_', scan.contexts[0], scan.eventLocations);
    assert.strictEqual(result.ok, true);
    assert.strictEqual(sheet.valueAt(3, columns.EVENT_ID), 'event-b');
    assert.strictEqual(sheet.valueAt(4, columns.EVENT_ID), 'event-a');
    assert.ok(sheet.noteAt(3, columns.CHART_NO).includes('另一列仍待重試'));
    assert.strictEqual(api.counts().insertCount, 0);
  } finally {
    context.Calendar.Events.update = originalUpdate;
  }
}

function testMonthlyReorderClearsStaleNotesWithoutCalendarCalls() {
  const { sheet, spreadsheet, columns, api } = makeMonthlyReorderFixture();
  [sheet.rows[2], sheet.rows[3]] = [sheet.rows[3], sheet.rows[2]];
  sheet.setNoteAt(4, columns.CHART_NO, '人工備註保留');
  call('setSystemNote_', sheet.getRange(4, columns.CHART_NO),
    evaluate('CALENDAR_CONFLICT_NOTE_PREFIX'), '舊重複警告');
  call('setSystemNote_', sheet.getRange(3, columns.CHART_NO),
    evaluate('CALENDAR_SYNC_NOTE_PREFIX'), '舊同步警告');
  call('reconcileCalendarRegistry_', spreadsheet, { apply: false, changeType: 'HEALTH' });
  assert.ok(sheet.noteAt(4, columns.CHART_NO).includes('舊重複警告'));
  const result = call('reconcileCalendarRegistry_', spreadsheet,
    { apply: true, changeType: 'OTHER' });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(sheet.noteAt(4, columns.CHART_NO), '人工備註保留');
  assert.strictEqual(sheet.noteAt(3, columns.CHART_NO), '');
  assert.strictEqual(api.counts().updateCount, 0);
  assert.strictEqual(api.counts().insertCount, 0);
  assert.strictEqual(api.counts().removeCount, 0);
  assert.strictEqual(call('readCalendarRegistryStore_').entries
    .find(entry => entry.eventId === 'event-a').row, 4);
}

function testMonthlyReorderKeepsActualDuplicateDiagnostics() {
  const { sheet, spreadsheet, columns, api } = makeMonthlyReorderFixture();
  sheet.rows[3][columns.EVENT_ID - 1] = 'event-a';
  call('reconcileCalendarRegistry_', spreadsheet, { apply: false, changeType: 'OTHER' });
  assert.throws(() => call('reconcileCalendarRegistry_', spreadsheet,
    { apply: true, changeType: 'OTHER' }), /重複/);
  [3, 4].forEach(row => assert.ok(sheet.noteAt(row, columns.CHART_NO).includes('重複')));
  assert.strictEqual(api.counts().updateCount, 0);
}

function testMonthlyDateChangesDuringSyncRemainPendingUntilRetried() {
  const { sheet, spreadsheet, columns, api } = makeMonthlyReorderFixture();
  const originalUpdate = context.Calendar.Events.update;
  context.Calendar.Events.update = (...args) => {
    const result = originalUpdate(...args);
    // Move this patient under a new date while the request is in flight.
    const nextHeader = sheet.rows[1].slice();
    nextHeader[columns.TIME - 1] = new Date(2026, 8, 17);
    const moved = sheet.rows.splice(2, 1)[0];
    sheet.rows.push(nextHeader, moved);
    return result;
  };
  try {
    const result = call('syncManagedRowsAt_', spreadsheet, sheet, [3]);
    assert.strictEqual(result.ok, false);
    assert.strictEqual(result.results[0].result.status, 'row_changed_during_sync');
    const pending = call('readCalendarRegistryStore_').entries
      .find(entry => entry.eventId === 'event-a');
    assert.strictEqual(pending.pendingSync, true);
    assert.strictEqual(pending.pendingError, 'row_changed_during_sync');
    assert.strictEqual(sheet.valueAt(5, columns.EVENT_ID), 'event-a');
    assert.strictEqual(sheet.valueAt(3, columns.EVENT_ID), 'event-b');
  } finally {
    context.Calendar.Events.update = originalUpdate;
  }
  const retried = call('reconcileCalendarRegistry_', spreadsheet,
    { apply: true, changeType: 'OTHER' });
  assert.strictEqual(retried.ok, true);
  assert.strictEqual(api.events['event-a'].start.date, '2026-09-17');
  assert.strictEqual(api.events['event-b'].start.date, '2026-09-10');
  assert.strictEqual(api.counts().insertCount, 0);
  assert.strictEqual(call('readCalendarRegistryStore_').entries
    .find(entry => entry.eventId === 'event-a').pendingSync, false);
}

function testMonthlyReorderDoesNotClearFailedPendingSync() {
  const { sheet, spreadsheet, columns, scan } = makeMonthlyReorderFixture();
  call('writeCalendarRegistryStore_', scan, [{
    ...call('toRegistryEntry_', scan.contexts[0]),
    pendingSync: true, pendingError: 'update_failed'
  }]);
  const originalUpdate = context.Calendar.Events.update;
  context.Calendar.Events.update = () => { throw new Error('503 temporary failure'); };
  try {
    const result = call('reconcileCalendarRegistry_', spreadsheet,
      { apply: true, changeType: 'OTHER' });
    assert.strictEqual(result.ok, false);
    assert.ok(sheet.noteAt(3, columns.CHART_NO).includes('更新失敗'));
    assert.strictEqual(call('readCalendarRegistryStore_').entries
      .find(entry => entry.eventId === 'event-a').pendingSync, true);
  } finally {
    context.Calendar.Events.update = originalUpdate;
  }
}

function testCalendarLifecycle404Recreates() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const api = installCalendarMock({ update404: true });
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const row = Array(headers.length).fill('');
  row[headers.indexOf('病歷號')] = 'X404';
  row[headers.indexOf('日期')] = new Date(2026, 7, 10);
  row[headers.indexOf('CalendarEventId')] = 'missing-event';
  const sheet = new FakeSheet('FU', 51, [headers, row]);
  const columns = call('getRequiredFuColumns_', sheet);
  const managed = call('buildFuRowContext_', sheet, 2, columns);
  const result = call('syncManagedContext_', managed, {
    'missing-event': [managed]
  });
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.status, 'recreated_404');
  assert.strictEqual(api.counts().updateCount, 1);
  assert.strictEqual(api.counts().insertCount, 1);
}

function testCalendarLifecycleDelete404ClearsId() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  installCalendarMock({ remove404: true });
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const row = Array(headers.length).fill('');
  row[headers.indexOf('病歷號')] = 'X404';
  row[headers.indexOf('CalendarEventId')] = 'already-gone';
  const sheet = new FakeSheet('FU', 52, [headers, row]);
  const columns = call('getRequiredFuColumns_', sheet);
  const managed = call('buildFuRowContext_', sheet, 2, columns);
  const result = call('deleteCurrentContextEvent_', managed);
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.status, 'already_missing');
  assert.strictEqual(
    sheet.valueAt(2, headers.indexOf('CalendarEventId') + 1),
    ''
  );
}

function testRegistryChunkRoundTrip() {
  clearScriptProperties();
  Object.keys(scriptPropertyCalls).forEach(key => {
    scriptPropertyCalls[key] = 0;
  });
  const contexts = Array.from({ length: 80 }, (_, index) => {
    return registryContext({
      eventId: `event-${index}`,
      row: index + 2,
      rowHash: `hash-${index}`
    });
  });
  const scan = registryScan(contexts);
  const index = call('writeCalendarRegistryStore_', scan, []);
  const loaded = call('readCalendarRegistryStore_');
  assert.strictEqual(index.eventCount, 80);
  assert.strictEqual(loaded.ok, true);
  assert.strictEqual(loaded.entries.length, 80);
  assert.ok(index.sheets[0].keys.length > 1);
  assert.ok(
    loaded.entries.every(entry => !Object.hasOwn(entry, 'patientName'))
  );
  assert.strictEqual(
    scriptPropertyCalls.setProperties,
    1,
    '索引應以一次批次寫入 Script Properties'
  );
  assert.strictEqual(
    scriptPropertyCalls.setProperty,
    0,
    '索引不可逐分段 setProperty'
  );
  assert.strictEqual(
    scriptPropertyCalls.getProperty,
    0,
    '索引回讀不可逐分段 getProperty'
  );
  assert.strictEqual(
    scriptPropertyCalls.getProperties,
    2,
    '索引寫入與回讀各只需一次 getProperties'
  );
}

function testRetainedRegistryEntryOverridesCurrentScopedScan() {
  Object.keys(scriptPropertyStore)
    .filter(key => key.includes('CALENDAR_ROW_REGISTRY_V2'))
    .forEach(key => delete scriptPropertyStore[key]);
  const current = registryContext({
    eventId: 'event-retained',
    rowHash: 'new-unprocessed-hash',
    row: 20
  });
  const retained = registryEntry({
    eventId: 'event-retained',
    rowHash: 'old-baseline-hash',
    row: 3
  });
  call('writeCalendarRegistryStore_', registryScan([current]), [retained]);
  const stored = call('readCalendarRegistryStore_');
  assert.strictEqual(stored.ok, true);
  assert.strictEqual(stored.entries.length, 1);
  assert.strictEqual(stored.entries[0].rowHash, 'old-baseline-hash');
  Object.keys(scriptPropertyStore)
    .filter(key => key.includes('CALENDAR_ROW_REGISTRY_V2'))
    .forEach(key => delete scriptPropertyStore[key]);
}

function testPendingDeleteIsRetried() {
  const result = analyze(
    [registryEntry({ pendingDelete: true })],
    [],
    { changeType: 'OTHER' }
  );
  assert.deepStrictEqual(
    plain(result.actions.map(item => item.type)),
    ['delete_missing']
  );
}

function testRemoveRowQueuePreservesDeletionSemantics() {
  const previous = context.reconcileCalendarRegistry_;
  let captured = null;
  context.reconcileCalendarRegistry_ = (_spreadsheet, options) => {
    captured = options;
    return { ok: true, results: [] };
  };
  try {
    const queue = call('emptyPendingCalendarQueue_');
    queue.fullScan = true;
    queue.allowMissingDeletes = true;
    const result = call('applyPendingCalendarQueue_', {}, queue);
    assert.strictEqual(result.ok, true);
    assert.strictEqual(captured.changeType, 'REMOVE_ROW');
    assert.strictEqual(captured.allowMissingDeletes, true);
  } finally {
    context.reconcileCalendarRegistry_ = previous;
  }
}

function testNativeRemoveRowsDeletesEventsAndUpdatesRegistry() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const sheet = new FakeSheet('FU', 6001, [headers]);
  const spreadsheet = makeFakeSpreadsheet([sheet]);
  const entries = ['removed-a', 'removed-b'].map((eventId, index) => {
    return registryEntry({
      eventId,
      sheetId: 6001,
      sheetName: 'FU',
      kind: 'FU',
      row: index + 2,
      rowHash: `removed-row-${index}`,
      bindingHash: `removed-binding-${index}`
    });
  });
  call(
    'writeCalendarRegistryStore_',
    registryScan([], [{
      sheetId: 6001,
      sheetName: 'FU',
      kind: 'FU',
      archived: false
    }]),
    entries
  );
  const api = installCalendarMock({
    initialEvents: entries.map(entry => ({
      id: entry.eventId,
      summary: 'X |  | 追蹤',
      description: 'Plan: delete row'
    }))
  });
  const queue = call('emptyPendingCalendarQueue_');
  queue.fullScan = true;
  queue.allowMissingDeletes = true;
  const result = call(
    'applyPendingCalendarQueue_',
    spreadsheet,
    queue
  );
  assert.strictEqual(result.ok, true);
  assert.strictEqual(api.counts().removeCount, 2);
  assert.strictEqual(call('readCalendarRegistryStore_').entries.length, 0);
}

function testNativeRemoveRowDeleteFailureKeepsPendingDelete() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const sheet = new FakeSheet('FU', 6002, [headers]);
  const spreadsheet = makeFakeSpreadsheet([sheet]);
  const entry = registryEntry({
    eventId: 'removed-fails',
    sheetId: 6002,
    sheetName: 'FU',
    kind: 'FU',
    row: 2,
    rowHash: 'removed-row-fails',
    bindingHash: 'removed-binding-fails'
  });
  call(
    'writeCalendarRegistryStore_',
    registryScan([], [{
      sheetId: 6002,
      sheetName: 'FU',
      kind: 'FU',
      archived: false
    }]),
    [entry]
  );
  installCalendarMock({
    initialEvents: [{ id: entry.eventId, summary: 'X |  | 追蹤' }],
    removeError: '503 delete unavailable'
  });
  const queue = call('emptyPendingCalendarQueue_');
  queue.fullScan = true;
  queue.allowMissingDeletes = true;
  const result = call(
    'applyPendingCalendarQueue_',
    spreadsheet,
    queue
  );
  assert.strictEqual(result.ok, false);
  const retained = call('readCalendarRegistryStore_').entries[0];
  assert.strictEqual(retained.eventId, 'removed-fails');
  assert.strictEqual(retained.pendingDelete, true);
}

function testRegistryDeleteNeverUsesShiftedRowNumberAsIdentity() {
  const result = analyze(
    [registryEntry({ row: 3, rowHash: 'old', bindingHash: 'old-binding' })],
    [registryContext({
      eventId: '',
      row: 3,
      rowHash: 'different',
      bindingHash: 'different-binding'
    })],
    { changeType: 'REMOVE_ROW' }
  );
  assert.ok(result.actions.some(item => item.type === 'delete_missing'));
  assert.ok(result.actions.some(item => item.type === 'create'));
  assert.strictEqual(
    result.conflicts.some(item => {
      return item.type === 'event_id_removed_or_partial_move';
    }),
    false
  );
}

function testRegistryBindingMatchPreventsUnsafeRowDelete() {
  const result = analyze(
    [registryEntry({ row: 3, rowHash: 'old', bindingHash: 'same-binding' })],
    [registryContext({
      eventId: '',
      row: 30,
      rowHash: 'changed-after-move',
      bindingHash: 'same-binding'
    })],
    { changeType: 'REMOVE_ROW' }
  );
  assert.strictEqual(
    result.actions.some(item => item.type === 'delete_missing'),
    false
  );
  assert.ok(result.conflicts.some(item => {
    return item.type === 'event_id_removed_or_partial_move';
  }));
}

function testPendingResolutionIsNotDeletedByLaterUnrelatedRemoveRow() {
  const result = analyze(
    [registryEntry({
      pendingResolution: true,
      pendingResolutionReason: 'multiple_binding_match'
    })],
    [],
    { changeType: 'REMOVE_ROW', allowMissingDeletes: true }
  );
  assert.strictEqual(
    result.actions.some(item => item.type === 'delete_missing'),
    false
  );
  assert.ok(result.conflicts.some(item => {
    return item.type === 'multiple_binding_match';
  }));
}

function testPendingResolutionClearsWhenSameBindingIdReturns() {
  const result = analyze(
    [registryEntry({
      pendingResolution: true,
      pendingResolutionReason: 'missing_event_id_from_sheet'
    })],
    [registryContext()],
    { changeType: 'HEALTH' }
  );
  assert.deepStrictEqual(
    plain(result.actions.map(item => item.type)),
    ['resolve_pending_resolution']
  );
  assert.strictEqual(result.conflicts.length, 0);
}

function testSingleSheetRefreshIsolatesMissingRegistryEntry() {
  clearScriptProperties();
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const row = Array(headers.length).fill('');
  row[headers.indexOf('病歷號')] = 'CURRENT';
  row[headers.indexOf('日期')] = new Date(2026, 7, 20);
  row[headers.indexOf('CalendarEventId')] = 'event-current';
  const sheet = new FakeSheet('FU', 6101, [headers, row]);
  const spreadsheet = makeFakeSpreadsheet([sheet]);
  const current = call(
    'buildFuRowContext_',
    sheet,
    2,
    call('getRequiredFuColumns_', sheet)
  );
  const stale = registryEntry({
    eventId: 'event-stale',
    sheetId: 6101,
    sheetName: 'FU',
    kind: 'FU',
    row: 80,
    rowHash: 'stale-row',
    bindingHash: 'stale-binding'
  });
  call(
    'writeCalendarRegistryStore_',
    registryScan(
      [current],
      [{ sheetId: 6101, sheetName: 'FU', kind: 'FU', archived: false }]
    ),
    [stale]
  );
  const refreshed = call(
    'refreshCalendarRegistryForSheet_',
    spreadsheet,
    sheet
  );
  assert.strictEqual(refreshed.ok, true);
  assert.strictEqual(refreshed.pendingResolutionCount, 1);
  const loaded = call('readCalendarRegistryStore_');
  const retained = loaded.entries.find(entry => {
    return entry.eventId === 'event-stale';
  });
  assert.strictEqual(retained.pendingResolution, true);
  assert.strictEqual(
    retained.pendingResolutionReason,
    'missing_event_id_from_sheet'
  );
  assert.ok(loaded.entries.some(entry => {
    return entry.eventId === 'event-current';
  }));
}

function testPendingResolutionDoesNotBlockUnrelatedFuCreate() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const fixture = makeFuLifecycleFixture({
    sheetId: 6102,
    date: new Date(2026, 7, 20),
    chartNo: 'NEW-ROW'
  });
  seedFuLifecycleRegistry(fixture, [
    makeFuRegistryEntry(fixture, 'event-stale-other', {
      row: 90,
      rowHash: 'stale-other-row',
      bindingHash: 'stale-other-binding'
    })
  ]);
  const api = installCalendarMock();
  const result = call(
    'syncManagedRowsAt_',
    fixture.spreadsheet,
    fixture.sheet,
    [2]
  );
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.results[0].result.status, 'created');
  assert.strictEqual(api.counts().insertCount, 1);
  const loaded = call('readCalendarRegistryStore_');
  assert.ok(loaded.entries.some(entry => {
    return entry.eventId === 'created-1' && !entry.pendingResolution;
  }));
  assert.ok(loaded.entries.some(entry => {
    return entry.eventId === 'event-stale-other' && entry.pendingResolution;
  }));
}

function makeFuLifecycleFixture(options = {}) {
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const row = Array(headers.length).fill('');
  row[headers.indexOf('病歷號')] = options.chartNo || 'LIFE-001';
  row[headers.indexOf('姓名')] = options.patientName || '';
  row[headers.indexOf('Condition')] = options.condition || '追蹤';
  row[headers.indexOf('日期')] = options.date || '';
  row[headers.indexOf('CalendarEventId')] = options.eventId || '';
  const sheet = new FakeSheet(
    'FU',
    options.sheetId === undefined ? 6201 : options.sheetId,
    [headers, row]
  );
  const spreadsheet = makeFakeSpreadsheet(
    [sheet],
    options.spreadsheetId || 'fu-lifecycle-test'
  );
  const columns = call('getRequiredFuColumns_', sheet);
  const contextValue = call('buildFuRowContext_', sheet, 2, columns);
  return { headers, row, sheet, spreadsheet, columns, context: contextValue };
}

function seedFuLifecycleRegistry(fixture, entries) {
  call(
    'writeCalendarRegistryStore_',
    registryScan(
      [],
      [{
        sheetId: fixture.sheet.getSheetId(),
        sheetName: 'FU',
        kind: 'FU',
        archived: false,
        fastFingerprint: 'fixture'
      }]
    ),
    entries
  );
}

function makeFuRegistryEntry(fixture, eventId, overrides = {}) {
  return {
    ...registryEntry({
      eventId,
      sheetId: fixture.sheet.getSheetId(),
      sheetName: 'FU',
      kind: 'FU',
      row: 2,
      blockKey: '2026-07-01',
      rowHash: 'historical-row-hash',
      bindingHash: fixture.context.bindingHash,
      ...overrides
    })
  };
}

function testFuDateAndIdClearUsesUniqueBindingAndDeletesEvent() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const fixture = makeFuLifecycleFixture();
  seedFuLifecycleRegistry(fixture, [
    makeFuRegistryEntry(fixture, 'event-old')
  ]);
  const api = installCalendarMock({
    initialEvents: [{
      id: 'event-old',
      summary: 'LIFE-001 |  | 追蹤',
      description: 'Plan: lifecycle',
      colorId: '8',
      start: { date: '2026-07-01' },
      end: { date: '2026-07-02' }
    }]
  });
  const result = call(
    'syncManagedRowsAt_',
    fixture.spreadsheet,
    fixture.sheet,
    [2]
  );
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.results[0].result.status, 'deleted');
  assert.strictEqual(api.counts().removeCount, 1);
  assert.strictEqual(api.counts().insertCount, 0);
  assert.strictEqual(
    fixture.sheet.valueAt(2, fixture.columns.EVENT_ID),
    ''
  );
  assert.strictEqual(fixture.sheet.valueAt(2, fixture.columns.CHART_NO), 'LIFE-001');
  const loaded = call('readCalendarRegistryStore_');
  assert.strictEqual(
    loaded.entries.some(entry => entry.eventId === 'event-old'),
    false
  );
}

function testFuRecoveredDeleteFailureStillStopsTracking() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const fixture = makeFuLifecycleFixture({ sheetId: 6202 });
  seedFuLifecycleRegistry(fixture, [
    makeFuRegistryEntry(fixture, 'event-delete-fails')
  ]);
  installCalendarMock({
    initialEvents: [{
      id: 'event-delete-fails',
      summary: 'LIFE-001 |  | 追蹤',
      description: 'Plan: lifecycle',
      colorId: '8',
      start: { date: '2026-07-01' },
      end: { date: '2026-07-02' }
    }],
    removeError: '503 temporary delete failure'
  });
  const result = call(
    'syncManagedRowsAt_',
    fixture.spreadsheet,
    fixture.sheet,
    [2]
  );
  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.results[0].result.status, 'tracking_stopped');
  assert.strictEqual(
    result.results[0].result.calendarDeleteStatus,
    'delete_failed'
  );
  assert.strictEqual(
    fixture.sheet.valueAt(2, fixture.columns.EVENT_ID),
    ''
  );
  assert.strictEqual(
    fixture.sheet.noteAt(2, fixture.columns.CHART_NO),
    ''
  );
  const loaded = call('readCalendarRegistryStore_');
  const pending = loaded.entries.find(entry => {
    return entry.eventId === 'event-delete-fails';
  });
  assert.strictEqual(pending, undefined);
}

function testInstallPreflightIgnoresUniqueFuStoppedTrackingConflict() {
  const fixture = makeFuLifecycleFixture({ sheetId: 6204 });
  const previous = makeFuRegistryEntry(fixture, 'event-stopped');
  const conflict = {
    type: 'event_id_removed_or_partial_move',
    eventId: previous.eventId,
    context: fixture.context,
    previous
  };
  assert.strictEqual(
    call('isIgnorableFuStoppedTrackingConflict_', conflict),
    true
  );
  assert.strictEqual(
    call('isIgnorableFuStoppedTrackingConflict_', {
      ...conflict,
      type: 'multiple_binding_match'
    }),
    false
  );
  const dated = makeFuLifecycleFixture({
    sheetId: 6205,
    date: new Date(2026, 7, 20)
  });
  assert.strictEqual(
    call('isIgnorableFuStoppedTrackingConflict_', {
      ...conflict,
      context: dated.context
    }),
    false
  );
}

function testFuStopTrackingContextAllowsNonBindingChanges() {
  const fixture = makeFuLifecycleFixture({ sheetId: 6206 });
  const item = {
    action: 'stop_tracking',
    sheetId: fixture.sheet.getSheetId(),
    row: 2,
    rowHash: 'stale-non-binding-row-hash',
    bindingHash: fixture.context.bindingHash
  };
  const recovered = call(
    'getFuLifecycleMigrationContext_',
    fixture.spreadsheet,
    item
  );
  assert.ok(recovered);
  assert.strictEqual(recovered.row, 2);
  assert.strictEqual(
    call(
      'getFuLifecycleMigrationContext_',
      fixture.spreadsheet,
      { ...item, action: 'rebind' }
    ),
    null
  );

  const stillBound = makeFuLifecycleFixture({
    sheetId: 6207,
    eventId: 'same-preview-event'
  });
  const stillBoundItem = {
    action: 'stop_tracking',
    sheetId: stillBound.sheet.getSheetId(),
    row: 2,
    rowHash: 'stale-non-binding-row-hash',
    bindingHash: stillBound.context.bindingHash,
    eventId: 'same-preview-event'
  };
  assert.ok(call(
    'getFuLifecycleMigrationContext_',
    stillBound.spreadsheet,
    stillBoundItem
  ));
  assert.ok(call(
    'getFuLifecycleMigrationContext_',
    stillBound.spreadsheet,
    {
      ...stillBoundItem,
      bindingHash: 'event-binding-does-not-match-stale-row'
    }
  ));
  assert.strictEqual(
    call(
      'getFuLifecycleMigrationContext_',
      stillBound.spreadsheet,
      { ...stillBoundItem, eventId: 'different-event' }
    ),
    null
  );

  const zeroSheetId = makeFuLifecycleFixture({
    sheetId: 0,
    eventId: 'zero-sheet-event'
  });
  assert.ok(call(
    'getFuLifecycleMigrationContext_',
    zeroSheetId.spreadsheet,
    {
      action: 'stop_tracking',
      sheetId: 0,
      row: 2,
      rowHash: zeroSheetId.context.rowHash,
      bindingHash: zeroSheetId.context.bindingHash,
      eventId: 'zero-sheet-event'
    }
  ));
}

function testFuAmbiguousBindingNeverCallsCalendar() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const fixture = makeFuLifecycleFixture({ sheetId: 6203 });
  seedFuLifecycleRegistry(fixture, [
    makeFuRegistryEntry(fixture, 'event-a'),
    makeFuRegistryEntry(fixture, 'event-b')
  ]);
  const api = installCalendarMock({
    initialEvents: [
      { id: 'event-a', summary: 'A', description: '' },
      { id: 'event-b', summary: 'B', description: '' }
    ]
  });
  const result = call(
    'syncManagedRowsAt_',
    fixture.spreadsheet,
    fixture.sheet,
    [2]
  );
  assert.strictEqual(result.ok, true);
  assert.strictEqual(
    result.results[0].result.status,
    'pending_resolution'
  );
  assert.strictEqual(api.counts().removeCount, 0);
  assert.strictEqual(api.counts().insertCount, 0);
  const loaded = call('readCalendarRegistryStore_');
  assert.strictEqual(
    loaded.entries.filter(entry => entry.pendingResolution).length,
    2
  );
}

function testRegistryRoundTripIncludesBindingAndResolutionFingerprint() {
  clearScriptProperties();
  const scan = registryScan([registryContext()]);
  const first = call('writeCalendarRegistryStore_', scan, []);
  const retained = registryEntry({
    pendingResolution: true,
    pendingResolutionReason: 'ambiguous_binding',
    pendingError: '虛構病患敏感錯誤內容'
  });
  const second = call('writeCalendarRegistryStore_', registryScan([]), [retained]);
  assert.notStrictEqual(first.fingerprint, second.fingerprint);
  const loaded = call('readCalendarRegistryStore_');
  assert.strictEqual(loaded.entries[0].bindingHash, 'binding-hash-1');
  assert.strictEqual(loaded.entries[0].pendingResolution, true);
  assert.strictEqual(
    loaded.entries[0].pendingResolutionReason,
    'ambiguous_binding'
  );
  assert.strictEqual(
    loaded.entries[0].pendingError,
    'legacy_error_redacted'
  );
  assert.strictEqual(
    JSON.stringify(scriptPropertyStore).includes('虛構病患敏感錯誤內容'),
    false
  );
  assert.strictEqual(Object.hasOwn(loaded.entries[0], 'patientName'), false);
}

function testFuLifecycleMigrationPreviewAndApplyStopTracking() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const fixture = makeFuLifecycleFixture({ sheetId: 6301 });
  seedFuLifecycleRegistry(fixture, [
    makeFuRegistryEntry(fixture, 'migration-event', { bindingHash: '' }),
    registryEntry({
      eventId: 'unrelated-monthly-event',
      sheetId: 9999,
      sheetName: '202607',
      kind: 'MONTHLY',
      row: 5,
      rowHash: 'unrelated-monthly-row',
      bindingHash: 'unrelated-monthly-binding'
    })
  ]);
  const api = installCalendarMock({
    initialEvents: [{
      id: 'migration-event',
      summary: 'LIFE-001 |  | 追蹤',
      description: 'Plan: legacy migration',
      colorId: '8',
      start: { date: '2026-07-01' },
      end: { date: '2026-07-02' }
    }]
  });
  const preview = call(
    'buildFuLifecycleRecoveryMigrationPlan_',
    fixture.spreadsheet
  );
  assert.strictEqual(preview.ok, true);
  assert.strictEqual(preview.counts.stopTracking, 1);
  assert.strictEqual(preview.items[0].action, 'stop_tracking');
  assert.strictEqual(JSON.stringify(preview).includes('LIFE-001'), false);
  const result = call(
    'applyFuLifecycleRecoveryMigrationPlan_',
    fixture.spreadsheet,
    preview,
    { confirmedOrphanIds: [] }
  );
  assert.strictEqual(result.ok, true);
  assert.strictEqual(api.counts().removeCount, 1);
  assert.strictEqual(api.events['migration-event'], undefined);
  assert.strictEqual(fixture.sheet.valueAt(2, fixture.columns.CHART_NO), 'LIFE-001');
  const remaining = call('readCalendarRegistryStore_').entries;
  assert.deepStrictEqual(
    plain(remaining.map(entry => entry.eventId)),
    ['unrelated-monthly-event']
  );
  assert.ok(result.backup.partCount >= 1);
  assert.ok(scriptPropertyStore[result.backup.manifestKey]);
}

function testFuLifecycleMigrationRepairsTwoVerifiedDuplicateGroups() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const makeRow = values => {
    const row = Array(headers.length).fill('');
    Object.entries(values).forEach(([header, value]) => {
      row[headers.indexOf(header)] = value;
    });
    return row;
  };
  const rows = [
    headers,
    makeRow({
      病歷號: 'DUP-A',
      姓名: '虛構甲',
      Condition: '追蹤甲',
      日期: new Date(2026, 7, 28),
      CalendarEventId: 'keep-a'
    }),
    makeRow({
      病歷號: 'DUP-B',
      姓名: '虛構乙',
      Condition: '追蹤乙',
      日期: new Date(2026, 9, 12),
      CalendarEventId: 'keep-b'
    }),
    makeRow({
      病歷號: 'HISTORY-ONLY',
      姓名: '特殊案例保留',
      Condition: '不再追蹤',
      CalendarEventId: 'stale-b'
    })
  ];
  const sheet = new FakeSheet('FU', 6401, rows);
  const spreadsheet = makeFakeSpreadsheet(
    [sheet],
    'fu-duplicate-lifecycle-test'
  );
  const scan = call('buildCurrentCalendarScan_', spreadsheet);
  const keepAContext = scan.contexts.find(item => item.eventId === 'keep-a');
  const keepBContext = scan.contexts.find(item => item.eventId === 'keep-b');
  const staleContext = scan.contexts.find(item => item.eventId === 'stale-b');
  call('writeCalendarRegistryStore_', scan, []);

  const keepA = {
    id: 'keep-a',
    ...plain(call('buildCalendarResource_', keepAContext))
  };
  const keepB = {
    id: 'keep-b',
    ...plain(call('buildCalendarResource_', keepBContext))
  };
  const api = installCalendarMock({
    initialEvents: [
      keepA,
      keepB,
      {
        ...keepA,
        id: 'orphan-a',
        description: ''
      },
      {
        ...keepB,
        id: 'stale-b',
        description: 'Plan: 舊資料略有差異'
      }
    ]
  });
  const preview = call(
    'buildFuLifecycleRecoveryMigrationPlan_',
    spreadsheet
  );
  assert.strictEqual(preview.ok, true);
  assert.strictEqual(preview.counts.stopTracking, 1);
  assert.strictEqual(preview.counts.orphanDelete, 1);
  assert.strictEqual(preview.counts.manual, 0);
  assert.strictEqual(preview.counts.actionable, 2);
  const stop = preview.items.find(item => item.eventId === 'stale-b');
  const orphan = preview.items.find(item => item.eventId === 'orphan-a');
  assert.strictEqual(stop.action, 'stop_tracking');
  assert.strictEqual(stop.row, staleContext.row);
  assert.ok(stop.duplicateGroupHash);
  assert.strictEqual(orphan.action, 'orphan_delete');
  assert.strictEqual(orphan.reason, 'fu_duplicate_without_data_source');
  assert.ok(orphan.duplicateGroupHash);
  assert.strictEqual(JSON.stringify(preview).includes('虛構甲'), false);

  const before = sheet.rows[staleContext.row - 1].slice();
  const result = call(
    'applyFuLifecycleRecoveryMigrationPlan_',
    spreadsheet,
    preview,
    { confirmedOrphanIds: ['orphan-a'] }
  );
  assert.strictEqual(result.ok, true);
  assert.strictEqual(api.counts().removeCount, 2);
  assert.ok(api.events['keep-a']);
  assert.ok(api.events['keep-b']);
  assert.strictEqual(api.events['orphan-a'], undefined);
  assert.strictEqual(api.events['stale-b'], undefined);
  const eventIdColumn = headers.indexOf('CalendarEventId');
  const after = sheet.rows[staleContext.row - 1].slice();
  assert.strictEqual(after[eventIdColumn], '');
  before.forEach((value, index) => {
    if (index !== eventIdColumn) assert.strictEqual(after[index], value);
  });
  assert.deepStrictEqual(
    plain(call('readCalendarRegistryStore_').entries.map(entry => {
      return entry.eventId;
    }).sort()),
    ['keep-a', 'keep-b']
  );
}

function testFuLifecycleMigrationIgnoresCancelledMonthlyGrayEvent() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const fixture = makeFuLifecycleFixture({ sheetId: 6402 });
  seedFuLifecycleRegistry(fixture, []);
  installCalendarMock({
    initialEvents: [{
      id: 'cancelled-monthly',
      summary: 'LIFE-001 |  | 追蹤',
      description: 'Plan: 月表取消',
      colorId: '8',
      start: { date: '2026-07-01' },
      end: { date: '2026-07-02' },
      extendedProperties: {
        private: {
          surgerySyncKind: 'MONTHLY',
          surgerySyncState: 'CANCELLED'
        }
      }
    }]
  });
  const preview = call(
    'buildFuLifecycleRecoveryMigrationPlan_',
    fixture.spreadsheet
  );
  assert.strictEqual(preview.ok, true);
  assert.strictEqual(preview.items.length, 0);
  assert.strictEqual(preview.counts.orphanDelete, 0);
  assert.strictEqual(
    call(
      'isFuLifecycleMigrationManagedEvent_',
      context.Calendar.Events.get('calendar@example.test', 'cancelled-monthly'),
      true
    ),
    false
  );
}

function testFuLifecycleMigrationRecreatesValid404Event() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const fixture = makeFuLifecycleFixture({
    sheetId: 6304,
    date: new Date(2026, 7, 20),
    eventId: 'missing-valid-event'
  });
  seedFuLifecycleRegistry(fixture, [
    makeFuRegistryEntry(fixture, 'missing-valid-event', {
      rowHash: fixture.context.rowHash,
      bindingHash: fixture.context.bindingHash
    })
  ]);
  const api = installCalendarMock();
  const preview = call(
    'buildFuLifecycleRecoveryMigrationPlan_',
    fixture.spreadsheet
  );
  assert.strictEqual(preview.ok, true);
  assert.strictEqual(preview.counts.recreateMissing, 1);
  const result = call(
    'applyFuLifecycleRecoveryMigrationPlan_',
    fixture.spreadsheet,
    preview,
    { confirmedOrphanIds: [] }
  );
  assert.strictEqual(result.ok, true);
  assert.strictEqual(api.counts().updateCount, 1);
  assert.strictEqual(api.counts().insertCount, 1);
  assert.strictEqual(
    fixture.sheet.valueAt(2, fixture.columns.EVENT_ID),
    'created-1'
  );
  const ids = call('readCalendarRegistryStore_').entries.map(entry => {
    return entry.eventId;
  });
  assert.deepStrictEqual(plain(ids), ['created-1']);
}

function testFuLifecycleMigrationPlansRebind404OrphanAndManual() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const fixture = makeFuLifecycleFixture({
    sheetId: 6302,
    date: new Date(2026, 7, 20)
  });
  seedFuLifecycleRegistry(fixture, [
    makeFuRegistryEntry(fixture, 'event-missing', {
      row: 80,
      bindingHash: 'unmatched-old-binding'
    }),
    makeFuRegistryEntry(fixture, 'event-read-fails', {
      row: 81,
      bindingHash: 'unmatched-error-binding'
    })
  ]);
  installCalendarMock({
    initialEvents: [
      {
        id: 'event-rebind',
        summary: 'LIFE-001 |  | 追蹤',
        description: 'Plan: rebind',
        colorId: '8',
        start: { date: '2026-08-20' },
        end: { date: '2026-08-21' }
      },
      {
        id: 'event-orphan',
        summary: 'ORPHAN |  | 追蹤',
        description: 'Plan: orphan',
        colorId: '8',
        start: { date: '2026-08-01' },
        end: { date: '2026-08-02' }
      }
    ],
    getErrorIds: ['event-read-fails']
  });
  const preview = call(
    'buildFuLifecycleRecoveryMigrationPlan_',
    fixture.spreadsheet
  );
  assert.strictEqual(preview.ok, true);
  assert.strictEqual(preview.counts.rebind, 1);
  assert.strictEqual(preview.counts.removeMissing, 1);
  assert.strictEqual(preview.counts.orphanDelete, 1);
  assert.strictEqual(preview.counts.manual, 1);
  const originalFingerprint = preview.fingerprint;
  fixture.sheet.setValueAt(2, fixture.columns.COND, '追蹤已變更');
  const changed = call(
    'buildFuLifecycleRecoveryMigrationPlan_',
    fixture.spreadsheet
  );
  assert.notStrictEqual(changed.fingerprint, originalFingerprint);
}

function testFuLifecycleMigrationAmbiguityIsFailClosed() {
  clearScriptProperties();
  scriptPropertyStore.CALENDAR_ID = 'calendar@example.test';
  const fixture = makeFuLifecycleFixture({
    sheetId: 6303,
    date: new Date(2026, 7, 20),
    eventId: 'ambiguous-keeper'
  });
  seedFuLifecycleRegistry(fixture, [
    makeFuRegistryEntry(fixture, 'ambiguous-keeper', {
      rowHash: fixture.context.rowHash,
      bindingHash: fixture.context.bindingHash
    })
  ]);
  const makeEvent = id => ({
    id,
    summary: 'LIFE-001 |  | 追蹤',
    description: 'Plan: ambiguous',
    colorId: '8',
    start: { date: '2026-08-20' },
    end: { date: '2026-08-21' }
  });
  const api = installCalendarMock({
    initialEvents: [
      makeEvent('ambiguous-keeper'),
      makeEvent('ambiguous-a'),
      makeEvent('ambiguous-b')
    ]
  });
  const preview = call(
    'buildFuLifecycleRecoveryMigrationPlan_',
    fixture.spreadsheet
  );
  assert.strictEqual(preview.ok, true);
  assert.strictEqual(preview.counts.manual, 3);
  assert.strictEqual(preview.counts.stopTracking, 0);
  assert.strictEqual(preview.counts.orphanDelete, 0);
  assert.strictEqual(api.counts().removeCount, 0);
}

function testHtmlUsesCollapsedAdvancedAreaAndArchiveLanguage() {
  const fuHtml = fs.readFileSync(path.join(root, 'fu_to_monthly.html'), 'utf8');
  const rollupHtml = fs.readFileSync(path.join(root, 'monthly_rollup.html'), 'utf8');
  assert.ok(fuHtml.includes('<details class="panel">'));
  assert.ok(fuHtml.includes('進階手術資料'));
  assert.ok(fuHtml.includes('selectedBlockRow'));
  assert.ok(fuHtml.includes('headerRow: block.headerRow'));
  assert.ok(fuHtml.includes('.getFuToMonthlyBlockOptions(month)'));
  assert.ok(rollupHtml.includes('彙整舊月刀表'));
  assert.ok(rollupHtml.includes('建立私人備份並永久封存'));
  assert.ok(rollupHtml.includes('一般選單不提供還原'));
  assert.ok(rollupHtml.includes('confirmedPermanentDeletion: true'));
}

function testClaspIncludesAllRuntimeModules() {
  const claspIgnore = fs.readFileSync(path.join(root, '.claspignore'), 'utf8');
  files.forEach(file => assert.ok(claspIgnore.includes(`!${file}`)));
  assert.ok(claspIgnore.includes('!annual_archive.js'));
  assert.ok(claspIgnore.includes('!fu_to_monthly.html'));
  assert.ok(claspIgnore.includes('!monthly_rollup.html'));
}

const tests = [
  testVersionAndModuleSplit,
  testEightDigitDateInputOnlyInFunctions,
  testRestoreDateFormatsPreservesValuesAndTimes,
  testPlanMigrationMovesWholeColumnAndIsIdempotent,
  testIolListIsReadOnlyAndKeepsAllDateBlocks,
  testIolListReadsOnlyRequestedDateDisplayRows,
  testEntropionCaseNormalizationPreservesOtherText,
  testPlanMigrationRepairsNativeRefErrorsAndPreservesOtherRules,
  testOnlyCalendarEventIdIsCanonicalSystemField,
  testCurrentMenuHasNoCompletedMigrationOrLegacyOutput,
  testMonthlySheetNameRecognition,
  testTimeNormalizationVariants,
  testMultiRowPasteNormalizesEveryMonthlyTimeCell,
  testWholeMonthlyTimeNormalizationUsesBoundedBatchReads,
  testFuSnapshotContextDoesNotRereadSheet,
  testUnifiedSortRoutesWithoutChangingSortImplementations,
  testCalendarPendingQueueCoalescesWithoutClinicalText,
  testFuNoDateDraftSkipsSyncAndClearsLegacyBusyNote,
  testIncompleteFuLifecycleSkipsOrDeletesOwnEvent,
  testBusyTriggerQueuesOnceWithoutWritingBusyNote,
  testPendingRetryWorkerDrainsQueueOnce,
  testOnEditNormalizesBeforeShortLockAndAvoidsFullReconcile,
  testFullDateDoesNotCollideWithTime,
  testMonthlyDateHeaderUsesDateAndBlankEventId,
  testMonthlyHeaderMarkerPreservesOtherCells,
  testMonthlyHeaderMarkerNeverOverwritesExistingChartCell,
  testMonthlyStructureIssuesReportHybridAndMissingMarker,
  testDuplicateMonthlyBlocksAreIndependentAndSupported,
  testOnEditScopesCalendarRowsToTouchedPatientOrHeaderBlock,
  testSurgeryDateInsertionPreviewSplitsOnlyRowsBelowSelection,
  testBlockRefRequiresExactHeaderRowAndIdentity,
  testPerformanceFastPathsAvoidGlobalRegistryRebuilds,
  testFuNativeSortUsesOneSortAndPreservesNotes,
  testFuNativeSortSkipsAlreadySortedRows,
  testTemporaryNativeSortRollsBackOnVerificationFailure,
  testMonthlyNativeSortPreservesBlocksAndUsesOneSort,
  testMonthlyNativeSortSkipsAlreadySortedSheet,
  testFuToMonthlyReusesPrecomputedMonthScan,
  testManagedRowFormattingUsesRangeListsAndLeavesCustomColumnAlone,
  testTouchedRowFormattingOnlyChangesEditedStandardColumns,
  testLightweightEventIndexAndSingleSheetRegistryRefresh,
  testSingleSheetSyncUsesVerifiedRegistryBeforeGlobalScan,
  testArchiveEventIdsArePersistedInRegistry,
  testReorderPreflightIsLocalAndDetectsUnsyncedRowHash,
  testFastFingerprintMetadataFallsBackSafely,
  testDefaultMonthlyDates,
  testCalendarTitlePreservesSlots,
  testRepairBindingSignatureIgnoresOnlyTrailingSurgeryDate,
  testRepairEventFingerprintIgnoresApiObjectKeyOrder,
  testMonthlyConditionComposition,
  testCalendarResourceTimedAndAllDay,
  testMonthlyStrikethroughControlsCancellationColorAndMarker,
  testFuIgnoresStrikethroughAndFormatQueuesCalendarScanOnly,
  testArbitraryHeaderOrderUsesNames,
  testDuplicateHeadersAreDiagnosed,
  testMonthlyDiagnosisSummaryFormulaDefinition,
  testMonthlyDiagnosisSummaryHeadersStayCanonical,
  testMonthlyDiagnosisSummaryMigrationIsSafeAndIdempotent,
  testAnnualArchiveCanonicalizesDiagnosisSummaryHeader,
  testFuNameOnlyRowIsValid,
  testMonthlyNameOnlyRowIsValid,
  testRegistrySameBlockRowReorderNeedsNoApi,
  testRegistryMoveUpdatesExistingEvent,
  testRegistryPendingSyncForcesRetryWithoutContentChange,
  testRegistryCrossMonthMoveUpdatesExistingEvent,
  testRegistryMissingIdNewRowCreatesEvent,
  testRegistryManualIdClearIsConflictNotDuplicateCreate,
  testRegistryNativeRowDeleteDeletesEvent,
  testRegistryDeleteNeverUsesShiftedRowNumberAsIdentity,
  testRegistryBindingMatchPreventsUnsafeRowDelete,
  testPendingResolutionIsNotDeletedByLaterUnrelatedRemoveRow,
  testPendingResolutionClearsWhenSameBindingIdReturns,
  testRegistryAmbiguousMissingRowWaitsForConfirmation,
  testRegistryWholeSheetDeletionNeverDeletesEvents,
  testRegistryHeaderDeletionDoesNotBatchReschedule,
  testRegistryDuplicateEventIdIsDiagnostic,
  testActiveAndArchivedDuplicateEventIdIsDiagnostic,
  testArchiveOnlyDuplicateEventIdIsDiagnostic,
  testArchivedRowsDoNotSync,
  testInvalidTimeIsConflict,
  testAnnualArchiveNamingAndEligibility,
  testConditionParserProvidesSafeSuggestions,
  testTimeSortIsStable,
  testFuDateSortKeyUsesAscendingDatesAndBlankLast,
  testFuTagOptions,
  testColumnWidthsAndFonts,
  testExplicitColumnWidthsPreserveCustomColumns,
  testMonthlyClinicalIdentifiersDefaultToPlainText,
  testMonthlyTextFormatsOnlyInitializeSystemCreatedRows,
  testMonthlyIolTextInputsHaveNoPostEditRewrite,
  testMonthlyHeaderConditionalFormatUsesExactMarker,
  testPlanConditionalFormatMarkersAndPriority,
  testConditionalFormatCleanupRemovesCurrentAndLegacySystemRules,
  testPermanentArchiveIsPrivateTransactionalAndCalendarFree,
  testAnnualArchiveStructureUsesProtectedSystemMetadata,
  testLegacyArchiveStateFailsClosed,
  testArchiveTransactionGuardFailsClosed,
  testFuScanIsBatchedAndRetryCacheRelocatesByFingerprint,
  testColumnWidthsRequireExplicitRequestAfterCreation,
  testInstallDoesNotBatchRewriteCalendar,
  testSystemNotesPreserveManualNotes,
  testCalendarErrorsUseVisibleIdentityCell,
  testHealthRowsRedactIdsAndSeparateFailures,
  testHealthRowsExposePendingQueueWithoutClinicalData,
  testHealthReportIsRetired,
  testRegistryStoresRawIdButNoClinicalPlaintext,
  testCalendarRepairOwnerRequiresUniqueIdentityAndSchedule,
  testLiveCalendarAuditFindsColorDriftMissingAndOrphans,
  testLiveCalendarAuditFailsClosedAndOrphanDeleteNeedsConfirmation,
  testCalendarLifecycleCreateAndVerify,
  testMonthlyMoveDuringCalendarUpdateDoesNotCopyEventId,
  testMonthlyReorderClearsStaleNotesWithoutCalendarCalls,
  testMonthlyReorderKeepsActualDuplicateDiagnostics,
  testMonthlyDateChangesDuringSyncRemainPendingUntilRetried,
  testMonthlyReorderDoesNotClearFailedPendingSync,
  testCalendarLifecycle404Recreates,
  testCalendarLifecycleDelete404ClearsId,
  testRegistryChunkRoundTrip,
  testRetainedRegistryEntryOverridesCurrentScopedScan,
  testPendingDeleteIsRetried,
  testRemoveRowQueuePreservesDeletionSemantics,
  testNativeRemoveRowsDeletesEventsAndUpdatesRegistry,
  testNativeRemoveRowDeleteFailureKeepsPendingDelete,
  testSingleSheetRefreshIsolatesMissingRegistryEntry,
  testPendingResolutionDoesNotBlockUnrelatedFuCreate,
  testFuDateAndIdClearUsesUniqueBindingAndDeletesEvent,
  testFuRecoveredDeleteFailureStillStopsTracking,
  testInstallPreflightIgnoresUniqueFuStoppedTrackingConflict,
  testFuStopTrackingContextAllowsNonBindingChanges,
  testFuAmbiguousBindingNeverCallsCalendar,
  testRegistryRoundTripIncludesBindingAndResolutionFingerprint,
  testFuLifecycleMigrationPreviewAndApplyStopTracking,
  testFuLifecycleMigrationRepairsTwoVerifiedDuplicateGroups,
  testFuLifecycleMigrationIgnoresCancelledMonthlyGrayEvent,
  testFuLifecycleMigrationRecreatesValid404Event,
  testFuLifecycleMigrationPlansRebind404OrphanAndManual,
  testFuLifecycleMigrationAmbiguityIsFailClosed,
  testHtmlUsesCollapsedAdvancedAreaAndArchiveLanguage,
  testClaspIncludesAllRuntimeModules
];

tests.forEach(test => test());
console.log(`smoke tests passed (${tests.length})`);
