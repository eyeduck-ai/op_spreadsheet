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
  'schedule_integrity.js',
  'workflows.js',
  'annual_archive.js',
  'archive_native.js',
  'archive_transaction.js',
  'schedule_identity.js'
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
  Sheets: { Spreadsheets: { get: () => ({ developerMetadata: [] }) } },
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
    getUuid: () => crypto.randomUUID()
  },
  Session: {
    getScriptTimeZone: () => 'Asia/Taipei'
  },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({ getId: () => 'SYNTHETIC-SPREADSHEET' }),
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

function testIolListIsReadOnlyAndKeepsAllDateBlocks() {
  const { sheet, columns } = makeMonthlyReorderFixture();
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

function testVersionAndModuleSplit() {
  assert.strictEqual(evaluate('CONFIG.VERSION'), '2026.10.19');
  assert.strictEqual(evaluate('typeof processRowChange'), 'function');
  assert.strictEqual(evaluate('typeof processCalendarStructureChange'), 'function');
  assert.strictEqual(evaluate('typeof createMonthlySurgerySheet'), 'function');
  assert.strictEqual(
    evaluate('typeof migrateMonthlyDiagnosisSummaryHeaders'),
    'undefined'
  );
  assert.strictEqual(evaluate('typeof assertScheduleIntegrity_'), 'function');
  assert.ok(sourceByFile['code.js'].includes('sheet_model.js'));
  assert.ok(sourceByFile['code.js'].includes('schedule_integrity.js'));
  assert.ok(sourceByFile['code.js'].includes('workflows.js'));
}

function testOnlyCalendarEventIdIsCanonicalSystemField() {
  assert.deepStrictEqual(plain(evaluate('CONFIG.SYSTEM_FIELD_KEYS')), ['EVENT_ID', 'SYNC_ID']);
  assert.deepStrictEqual(
    plain(evaluate('CONFIG.MONTHLY_SYSTEM_FIELD_KEYS')),
    ['EVENT_ID', 'SYNC_ID']
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
    '修復刀表格式',
    '套用所有 FU／月表建議欄寬',
    '月刀表 => FU',
    'FU => 月刀表',
    '建立新月刀表',
    '水晶體清單（選刀日／複製）',
    '選取列新增刀日',
    '整理目前分頁',
    '彙整舊月刀表'
  ].forEach(text => assert.ok(source.includes(text)));
  [
    '交由本系統同步日曆',
    '遷移固定排程識別欄',
    '檢查排程識別欄',
    '顯示日曆交接狀態',
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
    '預覽 FU 追蹤生命週期修復',
    '執行 FU 追蹤生命週期修復',
    '調整月表 Plan 至 Axis 後方',
    '恢復工作表原日期顯示格式',
    '統一 Entropion 大小寫',
    '啟用／修復月表診斷統計表頭',
    '資料遷移與舊版修復',
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
    'undefined'
  );
  assert.strictEqual(
    evaluate('typeof executeFuLifecycleRecoveryMigration'),
    'undefined'
  );
  assert.strictEqual(evaluate('typeof migrateMonthlyPlanOrder'), 'undefined');
  assert.strictEqual(evaluate('typeof restoreWorksheetDateFormats'), 'undefined');
  assert.strictEqual(evaluate('typeof migrateEntropionCase'), 'undefined');
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
  assert.strictEqual(call('resolveScheduleTime_', '').parsedTime, null);
  assert.strictEqual(call('resolveScheduleTime_', 'PM').timeNote, 'PM');
  assert.ok(call('resolveScheduleTime_', '2460').errorMessage.includes('時間範圍錯誤'));
  assert.ok(call('resolveScheduleTime_', '25:00').errorMessage.includes('時間範圍錯誤'));
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
  const event = { range: sheet.getRange(2, columns.TIME, 100, 1) };
  const result = call('normalizeEditedTimes_', event, sheet, 'MONTHLY', columns);
  assert.strictEqual(result.normalized, 100);
  assert.strictEqual(result.errors, 0);
  assert.ok(sheet.calls.getRange <= 5, '100 rows must use bounded batch reads');
  assert.strictEqual(sheet.calls.setValues, 1);
  assert.ok(sheet.rows.slice(1).every(row => row[time] === '08:30'));
  sheet.calls.setValues = 0;
  assert.strictEqual(call('normalizeEditedTimes_', event, sheet, 'MONTHLY', columns).normalized, 0);
  assert.strictEqual(sheet.calls.setValues, 0);
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
  header[1] = '高榮';
  header[2] = '◆ 刀日';
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
  header[2] = '';
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
  assert.strictEqual(result.ok, false);
  assert.strictEqual(sheet.valueAt(2, columns.CHART_NO), '');
  assert.strictEqual(sheet.valueAt(2, columns.HOSPITAL), '');
  assert.strictEqual(sheet.valueAt(2, columns.NAME), '刀日人工備註');
  sheet.setValueAt(2, columns.CHART_NO, '◆ 刀日');
  sheet.setValueAt(2, columns.HOSPITAL, '聯醫');
  assert.strictEqual(call('ensureMonthlyHeaderPresentation_', sheet, 2, columns).ok, true);
  assert.strictEqual(
    sheet.validations[`2:${columns.CHART_NO}`].allowInvalid,
    false
  );

  sheet.setValueAt(2, columns.TIME, '08:00');
  call('ensureMonthlyHeaderPresentation_', sheet, 2, columns);
  assert.strictEqual(sheet.valueAt(2, columns.CHART_NO), '◆ 刀日');
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
    ['monthly_hybrid_date_event_id', 'monthly_invalid_date_header']
  );
  assert.strictEqual(scan.blocks.length, 0);
  assert.strictEqual(scan.hybrids.length, 2);
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
  assert.deepStrictEqual(plain(preview.patientRows), [4]);
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
  const fastState = { values: sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues(), columns };
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
  assert.ok(!result.sheetScan);
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
  assert.ok(!result.sheetScan);
  assert.strictEqual(result.strategy, 'precomputed_noop');
  assert.strictEqual(sheet.calls.getFormulas, 1); // Header formulas must always be checked for duplicate fields.
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
    '心得',
    'ScheduleSyncId'
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
    '心得',
    'ScheduleSyncId'
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
  const categories = definitions.map(definition => definition.label);
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

function columnToLetterForTest(column) {
  return call('columnToLetter_', column);
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
        message.includes(columnToLetterForTest(standardHeaders.length + 1));
    })
  );

  const errorHeaders = standardHeaders.slice();
  errorHeaders[diagnosisColumn - 1] = '#ERROR!';
  const errorSheet = new FakeSheet('202608', 2603, [errorHeaders]);
  errorSheet.setFormulaAt(
    1,
    diagnosisColumn,
    '=T(N("SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V2"))&"診斷"'
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

function testMonthlyDiagnosisSummaryCreationAndManagedRefresh() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS'));
  const columns = monthlyColumns();
  const diagnosisColumn = columns.DIAGNOSIS;
  const dataRow = Array(headers.length).fill('');
  dataRow[columns.SIDE - 1] = 'OU';
  dataRow[columns.DIAGNOSIS - 1] = 'CATA + ERM';
  dataRow[columns.CHART_NO - 1] = 'SAFE-001';

  const sheet = new FakeSheet('202608', 2610, [
    headers.slice(),
    dataRow.slice()
  ]);
  const originalData = plain(sheet.rows.slice(1));
  const passive = call(
    'ensureMonthlyDiagnosisSummaryHeader_',
    sheet,
    columns,
    { enable: false }
  );
  assert.strictEqual(passive.status, 'plain');
  assert.strictEqual(sheet.formulaAt(1, diagnosisColumn), '');

  const expectedFormula = call(
    'buildMonthlyDiagnosisSummaryFormula_',
    columns
  );
  const enabled = call(
    'ensureMonthlyDiagnosisSummaryHeader_',
    sheet,
    columns,
    { enable: true }
  );
  assert.strictEqual(enabled.status, 'enabled');
  assert.strictEqual(sheet.formulaAt(1, diagnosisColumn), expectedFormula);
  assert.deepStrictEqual(plain(sheet.rows.slice(1)), originalData);

  const staleSheet = new FakeSheet('202609', 2611, [headers.slice()]);
  staleSheet.setFormulaAt(
    1,
    diagnosisColumn,
    '=T(N("SURGERY_MONTHLY_DIAGNOSIS_SUMMARY_V2"))&"診斷"'
  );
  const refreshed = call(
    'ensureMonthlyDiagnosisSummaryHeader_',
    staleSheet,
    columns
  );
  assert.strictEqual(refreshed.status, 'refreshed');
  assert.strictEqual(
    staleSheet.formulaAt(1, diagnosisColumn),
    expectedFormula
  );

  const customSheet = new FakeSheet('202610', 2612, [headers.slice()]);
  const customFormula = '="人工診斷摘要"';
  customSheet.setFormulaAt(1, diagnosisColumn, customFormula);
  const conflict = call(
    'ensureMonthlyDiagnosisSummaryHeader_',
    customSheet,
    columns
  );
  assert.strictEqual(conflict.status, 'conflict');
  assert.strictEqual(customSheet.formulaAt(1, diagnosisColumn), customFormula);

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


function monthlyColumns() {
  const result = {};
  plain(evaluate('CONFIG.MONTHLY_FIELD_KEYS')).forEach(key => {
    const header = evaluate(`CONFIG.MONTHLY_FIELD_HEADERS.${key}`);
    const index = plain(evaluate('CONFIG.MONTHLY_HEADERS')).indexOf(header);
    result[key] = index + 1;
  });
  return result;
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

  assert.strictEqual(call('processCalendarStructureChange'), undefined);

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
  const integritySource = sourceByFile['schedule_integrity.js'];
  const onEditBody = integritySource.slice(
    integritySource.indexOf('function processScheduleEdit'),
    integritySource.indexOf('function processRowChange')
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
    source.indexOf('function repairMainTrackingFormatting_'),
    source.indexOf('function applyAllRecommendedColumnWidths')
  );
  assert.ok(initFu.includes('forceWidths: false'));
  assert.strictEqual(initFu.includes('insertSheet'), false);
  const workflowSource = sourceByFile['workflows.js'];
  const initMonth = workflowSource.slice(
    workflowSource.indexOf('function createOrInitializeMonthlySheet_'),
    workflowSource.indexOf('function createMonthlySurgerySheet')
  );
  assert.ok(initMonth.includes('forceWidths: created'));
}

function testFormatRepairDoesNotInitializeOrSynchronize() {
  const repair = evaluate('repairScheduleFormatting.toString()');
  assert.ok(repair.includes('assertScheduleIntegrity_'));
  assert.strictEqual(/Registry|Calendar\.|syncManaged|reorderManagedSheets_|newScheduleSyncId_/.test(repair), false);
  assert.strictEqual(evaluate('typeof setup'), 'undefined');
  assert.strictEqual(evaluate('typeof installOrRepairSystem'), 'undefined');
  const install = evaluate('repairScheduleEditTrigger_.toString()');
  assert.ok(install.includes("newTrigger('processScheduleEdit')"));
  assert.strictEqual(install.includes('.onChange()'), false);
  assert.strictEqual(install.includes('.timeBased()'), false);
}

function testSystemNotesPreserveManualNotes() {
  const sheet = new FakeSheet('FU', 40, [['病歷號'], ['X001']]);
  const cell = sheet.getRange(2, 1);
  cell.setNote('人工註記');
  const prefix = evaluate('MONTHLY_STRUCTURE_NOTE_PREFIX');

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


function makeMonthlyReorderFixture() {
  const headers = plain(evaluate('CONFIG.MONTHLY_HEADERS')).reverse();
  const makeRow = fields => headers.map(header => fields[header] || '');
  const sheet = new FakeSheet('202609', 909, [headers,
    makeRow({ '日期／報到時間': new Date(2026, 8, 10), '醫院': '高榮', '病歷號': '◆ 刀日' }),
    makeRow({ '病歷號': 'TEST-A', '姓名': '測試甲', CalendarEventId: 'event-a' }),
    makeRow({ '病歷號': 'TEST-B', '姓名': '測試乙', CalendarEventId: 'event-b' })]);
  return { sheet, spreadsheet: makeFakeSpreadsheet([sheet]), columns: call('getRequiredMonthlyColumns_', sheet) };
}


function testHtmlUsesCollapsedAdvancedAreaAndArchiveLanguage() {
  const fuHtml = fs.readFileSync(path.join(root, 'fu_to_monthly.html'), 'utf8');
  const rollupHtml = fs.readFileSync(path.join(root, 'monthly_rollup.html'), 'utf8');
  assert.ok(fuHtml.includes('<details class="panel">'));
  assert.ok(fuHtml.includes('進階手術資料'));
  assert.ok(fuHtml.includes('selectedBlockRow'));
  assert.ok(fuHtml.includes('headerRow: block.headerRow'));
  assert.ok(fuHtml.includes('.getFuToMonthlyBlockOptions(month)'));
  assert.strictEqual(/calendarStatus|同步 Calendar|重試 Calendar/.test(fuHtml), false);
  assert.ok(rollupHtml.includes('彙整舊月刀表'));
  assert.ok(rollupHtml.includes('建立私人備份並複製核對'));
  assert.ok(rollupHtml.includes('確認封存後移除原月份'));
  assert.ok(rollupHtml.includes('一般選單不提供還原'));
  assert.ok(rollupHtml.includes("confirmedPermanentDeletion: action === 'finalize' && confirmDeletion.checked"));
}

function testClaspIncludesAllRuntimeModules() {
  const claspIgnore = fs.readFileSync(path.join(root, '.claspignore'), 'utf8');
  files.forEach(file => assert.ok(claspIgnore.includes(`!${file}`)));
  assert.ok(claspIgnore.includes('!annual_archive.js'));
  assert.ok(claspIgnore.includes('!fu_to_monthly.html'));
  assert.ok(claspIgnore.includes('!monthly_rollup.html'));
}


function testExternalOwnerPreventsCalendarMutationsAndKeepsEditFormatting() {
  const body = evaluate('processScheduleEdit.toString()');
  assert.ok(body.includes('normalizeEditedTimes_'));
  assert.ok(body.includes('applyTouchedManagedRowsFormatBatch_'));
  assert.ok(body.includes('ensureMonthlyHeaderPresentation_'));
  assert.strictEqual(/syncManaged|Calendar|Registry/.test(body), false);
  assert.strictEqual(sourceByFile['code.js'].includes("'handoverCalendarToApplication'"), false);
}


function prepareIdentitySheet(sheet) {
  const hidden = [];
  const protections = [];
  sheet.hideColumns = column => hidden.push(column);
  sheet.getProtections = () => protections;
  const getRange = sheet.getRange.bind(sheet);
  sheet.getRange = (...args) => {
    const range = getRange(...args);
    range.protect = () => {
      const protection = { description: '', warningOnly: true,
        setDescription(value) { this.description = value; return this; },
        getDescription() { return this.description; },
        setWarningOnly(value) { this.warningOnly = value; return this; } };
      protections.push(protection);
      return protection;
    };
    return range;
  };
  context.SpreadsheetApp.ProtectionType = { RANGE: 'RANGE' };
  return { hidden, protections };
}




function testScheduleIdentitySurvivesNativeSortAndWholeDateChanges() {
  const { sheet, columns } = makeMonthlyReorderFixture();
  sheet.rows.slice(1).forEach((row, index) => {
    if (row.some(value => value !== '')) sheet.setValueAt(index + 2, columns.SYNC_ID, (index + 1).toString(16).padStart(32, '0'));
  });
  const before = {};
  sheet.rows.slice(1).forEach(row => {
    const chart = row[columns.CHART_NO - 1];
    if (chart && chart !== '◆ 刀日') before[chart] = row[columns.SYNC_ID - 1];
  });
  sheet.setValueAt(2, columns.TIME, new Date(2026, 10, 18));
  call('rebuildMonthlySheetSorted_', sheet);
  sheet.rows.slice(1).forEach(row => {
    const chart = row[columns.CHART_NO - 1];
    if (chart && chart !== '◆ 刀日') assert.strictEqual(row[columns.SYNC_ID - 1], before[chart]);
  });
  const scan = call('scanMonthlyBlocks_', sheet, columns);
  assert.ok(scan.blocks.every(item => item.blockKey.startsWith('2026-11-18|')));
}

function testFormattingOnlyDeploymentHasNoCalendarApiAndLegacyCallsAreSafe() {
  const deployed = files.map(file => sourceByFile[file]).join('\n');
  assert.strictEqual(/\bCalendar\.(?:Events|Calendars|CalendarList)\b/.test(deployed), false);
  assert.strictEqual(/\bCalendarApp\b/.test(deployed), false);
  assert.strictEqual(/\|\|=|&&=|\?\?=/.test(deployed), false, 'Apps Script parser rejects logical assignment');
  assert.strictEqual(evaluate('typeof buildCalendarResource_'), 'undefined');
  assert.strictEqual(evaluate('typeof rebuildCalendarRegistry_'), 'undefined');
  assert.strictEqual(evaluate('typeof enqueuePendingCalendarSync_'), 'undefined');
  plain(evaluate('RETIRED_TRIGGER_HANDLERS')).filter(name => name !== 'processRowChange')
    .forEach(name => assert.strictEqual(call(name), undefined));
  const manifest = JSON.parse(fs.readFileSync(path.join(root, 'appsscript.json'), 'utf8'));
  assert.deepStrictEqual(manifest.dependencies, { enabledAdvancedServices: [{ userSymbol: 'Sheets', version: 'v4', serviceId: 'sheets' }] });
  assert.strictEqual(deployed.includes('vghks.calendar.control'), false);
  assert.strictEqual(evaluate('typeof migrateScheduleSyncIds_'), 'undefined');
  assert.strictEqual(sourceByFile['code.js'].includes("'migrateScheduleSyncIds'"), false);
  assert.strictEqual(fs.existsSync(path.join(root, 'calendar_sync.js')), false);
}


function testNewPatientInsertionPreservesManualBlankRowAndGetsNewIdentity() {
  const { sheet, columns } = makeMonthlyReorderFixture();
  const blank = Array(sheet.rows[0].length).fill('');
  sheet.rows.push(blank.slice(), blank.slice());
  sheet.setNoteAt(5, columns.MEMO, '人工空白列備註');
  const before = sheet.rows.length;
  sheet.insertRowBefore = row => {
    sheet.rows.splice(row - 1, 0, Array(sheet.rows[0].length).fill(''));
    const shifted = {};
    Object.entries(sheet.notes).forEach(([key, value]) => {
      const [oldRow, column] = key.split(':').map(Number);
      shifted[`${oldRow >= row ? oldRow + 1 : oldRow}:${column}`] = value;
    });
    sheet.notes = shifted;
  };
  const scan = call('scanMonthlyBlocks_', sheet, columns);
  const row = call('insertPatientAtBlockEnd_', sheet, scan.blocks[0],
    { CHART_NO: 'SYN-NEW', NAME: '新增虛構病人' }, columns, scan);
  assert.strictEqual(row, 5);
  assert.strictEqual(sheet.rows.length, before + 1);
  assert.strictEqual(sheet.valueAt(row, columns.CHART_NO), 'SYN-NEW');
  assert.match(sheet.valueAt(row, columns.SYNC_ID), /^[a-f0-9]{32}$/);
  assert.strictEqual(sheet.valueAt(row, columns.EVENT_ID), '');
  assert.strictEqual(sheet.noteAt(6, columns.MEMO), '人工空白列備註');
  assert.strictEqual(sheet.valueAt(6, columns.SYNC_ID), '');
}


function testDuplicateEventBindingStopsSheetOperations() {
  const { sheet, spreadsheet, columns } = makeMonthlyReorderFixture();
  sheet.setValueAt(4, columns.EVENT_ID, sheet.valueAt(3, columns.EVENT_ID));
  const before = JSON.stringify(sheet.rows);
  assert.throws(() => call('assertScheduleIntegrity_', spreadsheet), /CalendarEventId 重複/);
  assert.strictEqual(JSON.stringify(sheet.rows), before);
}

function testUnnamedTailDataSurvivesMonthlySorting() {
  const { sheet, columns } = makeMonthlyReorderFixture();
  const tail = sheet.getLastColumn() + 2;
  sheet.setValueAt(5, tail, '人工尾欄資料');
  const scan = call('scanMonthlyBlocks_', sheet, columns);
  assert.strictEqual(scan.blocks[0].patientRows.length, 3);
  call('rebuildMonthlySheetSorted_', sheet);
  assert.strictEqual(sheet.rows.filter(row => row[tail - 1] === '人工尾欄資料').length, 1);
  assert.strictEqual(sheet.valueAt(1, tail), '');
  assert.strictEqual(call('scanMonthlyBlocks_', sheet, columns).blocks[0].patientRows.length, 3);
}

function testFuRequestCacheUsesPermanentIdentityAfterMoveAndCalendarReplacement() {
  const { sheet, spreadsheet, columns } = makeMonthlyReorderFixture();
  const syncId = 'd'.repeat(32);
  sheet.setValueAt(3, columns.SYNC_ID, syncId);
  const cached = { syncId };
  [sheet.rows[2], sheet.rows[3]] = [sheet.rows[3], sheet.rows[2]];
  sheet.setValueAt(4, columns.EVENT_ID, 'replacement-event');
  sheet.setValueAt(2, columns.TIME, new Date(2026, 9, 20));
  const found = call('locateCachedMonthlySchedule_', spreadsheet, cached);
  assert.strictEqual(found.row, 4);
  assert.strictEqual(found.syncId, syncId);
  assert.strictEqual(sheet.valueAt(found.row, columns.EVENT_ID), 'replacement-event');
  assert.throws(() => call('locateCachedMonthlySchedule_', spreadsheet, { eventId: 'replacement-event' }), /永久排程 ID/);
  assert.throws(() => call('locateCachedMonthlySchedule_', spreadsheet, cached, { chartNo: 'WRONG' }), /病人已變更/);
  sheet.rows.push(sheet.rows[3].slice());
  assert.throws(() => call('locateCachedMonthlySchedule_', spreadsheet, cached), /排程識別重複/);
}

function testSurgeryDatePreviewSelectsCampusAndDetectsChange() {
  const { sheet } = makeMonthlyReorderFixture();
  const date = new Date(2026, 8, 10);
  const high = call('buildSurgeryDateInsertionPreview_', sheet, 4, date, '高榮');
  const union = call('buildSurgeryDateInsertionPreview_', sheet, 4, date, '聯醫');
  assert.deepStrictEqual(plain(high.duplicateRows), [2]);
  assert.deepStrictEqual(plain(union.duplicateRows), []);
  assert.strictEqual(union.hospital, '聯醫');
  assert.notStrictEqual(high.fingerprint, union.fingerprint);
  assert.throws(() => call('buildSurgeryDateInsertionPreview_', sheet, 4, date, '未知'), /高榮或聯醫/);
}


function testInsertDayWritesSelectedCampusAndRejectsDuplicateIdentityBeforeWriting() {
  const { sheet, spreadsheet, columns } = makeMonthlyReorderFixture();
  sheet.setValueAt(4, columns.MEMO, '人工保留');
  sheet.getActiveRange = () => sheet.getRange(4, 1);
  sheet.insertRowBefore = row => sheet.rows.splice(row - 1, 0, Array(sheet.rows[0].length).fill(''));
  sheet.setActiveRange = () => {};
  spreadsheet.getActiveSheet = () => sheet;
  spreadsheet.setActiveSheet = () => {};
  const saved = { ...context.SpreadsheetApp };
  const functions = ['runMenuAction_', 'withSheetOperationLock_'];
  const originals = functions.map(name => evaluate(name));
  let prompts, alerts = [];
  context.SpreadsheetApp.getActiveSpreadsheet = () => spreadsheet;
  context.SpreadsheetApp.getUi = () => ({ Button: { OK: 'OK' }, ButtonSet: { OK_CANCEL: 'OK_CANCEL' },
    prompt: () => { const answer = prompts.shift(); return { getSelectedButton: () => 'OK', getResponseText: () => answer }; },
    alert: text => { alerts.push(text); return 'OK'; } });
  context.runMenuAction_ = (_label, action) => action();
  context.withSheetOperationLock_ = action => action();
  try {
    prompts = ['20260911', '聯醫'];
    const result = call('insertSurgeryDateAtSelection');
    assert.strictEqual(result.ok, true);
    assert.strictEqual(sheet.valueAt(4, columns.HOSPITAL), '聯醫');
    assert.strictEqual(sheet.valueAt(4, columns.CHART_NO), '◆ 刀日');
    assert.match(sheet.valueAt(4, columns.SYNC_ID), /^[a-f0-9]{32}$/);
    assert.strictEqual(sheet.valueAt(5, columns.EVENT_ID), 'event-b');
    assert.strictEqual(sheet.valueAt(5, columns.MEMO), '人工保留');
    assert.ok(alerts.some(text => text.includes('2026-09-11／聯醫刀日')));
    sheet.setValueAt(3, columns.SYNC_ID, sheet.valueAt(4, columns.SYNC_ID));
    const before = JSON.stringify(sheet.rows);
    prompts = ['20260912', '高榮'];
    assert.throws(() => call('insertSurgeryDateAtSelection'), /識別碼重複/);
    assert.strictEqual(JSON.stringify(sheet.rows), before);
  } finally {
    context.SpreadsheetApp = saved;
    functions.forEach((name, index) => { context[name] = originals[index]; });
  }
}

function testTriggerRepairRemovesLegacyWritersAndPreservesOtherTriggersAndHandover() {
  const saved = { ...context.SpreadsheetApp }, oldScript = context.ScriptApp;
  const trigger = name => ({ getHandlerFunction: () => name });
  let triggers = ['processRowChange', 'retryPendingCalendarSync', 'processMonthlyStructureChange', 'manualOtherTrigger'].map(trigger);
  context.SpreadsheetApp.getActiveSpreadsheet = () => makeFakeSpreadsheet([]);
  context.ScriptApp = {
    getProjectTriggers: () => triggers.slice(),
    deleteTrigger: value => { triggers = triggers.filter(item => item !== value); },
    newTrigger: name => {
      const builder = { forSpreadsheet: () => builder, onEdit: () => builder,
        create: () => triggers.push(trigger(name)) };
      return builder;
    }
  };
  scriptPropertyStore.CALENDAR_ROW_REGISTRY_V2_CHUNK_1 = 'legacy';
  scriptPropertyStore.CALENDAR_SYNC_PENDING_V2 = 'legacy';
  scriptPropertyStore.CALENDAR_OWNER_APPLICATION_V1 = 'preserved';
  try {
    const result = call('repairScheduleEditTrigger_');
    assert.strictEqual(result.onEdit, 'processScheduleEdit');
    assert.deepStrictEqual(triggers.map(value => value.getHandlerFunction()).sort(), ['manualOtherTrigger', 'processScheduleEdit']);
    assert.strictEqual(scriptPropertyStore.CALENDAR_ROW_REGISTRY_V2_CHUNK_1, 'legacy');
    assert.strictEqual(scriptPropertyStore.CALENDAR_SYNC_PENDING_V2, 'legacy');
    assert.strictEqual(scriptPropertyStore.CALENDAR_OWNER_APPLICATION_V1, 'preserved');
  } finally {
    context.SpreadsheetApp = saved;
    context.ScriptApp = oldScript;
    delete scriptPropertyStore.CALENDAR_OWNER_APPLICATION_V1;
    delete scriptPropertyStore.CALENDAR_ROW_REGISTRY_V2_CHUNK_1;
    delete scriptPropertyStore.CALENDAR_SYNC_PENDING_V2;
  }
}


function testObsoleteCalendarMenusAndInternalScaffoldingAreRemoved() {
  ['promptAndSaveCalendarId', 'migrateScheduleSyncIds', 'checkScheduleSyncIds',
    'handoverCalendarToApplication', 'showCalendarHandover', 'restoreGasCalendarOwner',
    'checkSyncHealth', 'syncPendingChanges', 'repairSelectedCalendarRows',
    'applySelectedCalendarRepairPlan_', 'clearContextEventIds', 'repairCalendarBindings',
    'isExternalCalendarOwner_', 'externalCalendarMenu_', 'externalCalendarResult_',
    'removeAnnualArchiveProtections_', 'hideAndProtectSystemColumns_',
    'protectAnnualArchiveSheet_', 'isSpreadsheetFormulaErrorHeader_',
    'initializeMainTrackingSheet_', 'initializeAllMonthlySheets_', 'normalizeExistingTimeColumn_']
    .forEach(name => assert.strictEqual(evaluate(`typeof ${name}`), 'undefined'));
}

function testIdentityScanDiagnosesExistingDuplicatesWithoutMigration() {
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const first = new FakeSheet('FU', 902, [headers, ['SYN-1', '甲']]);
  const second = new FakeSheet('All', 903, [headers, ['SYN-2', '乙']]);
  const column = headers.indexOf('ScheduleSyncId') + 1;
  first.setValueAt(2, column, 'a'.repeat(32));
  second.setValueAt(2, column, 'a'.repeat(32));
  const before = JSON.stringify([first.rows, second.rows]);
  const result = call('scanScheduleIdentities_', makeFakeSpreadsheet([first, second]));
  assert.ok(result.conflicts.some(value => value.includes('識別碼重複')));
  assert.strictEqual(JSON.stringify([first.rows, second.rows]), before);
}

function testSystemColumnsAreHiddenWithoutCreatingOrChangingProtections() {
  const headers = plain(evaluate('CONFIG.HEADERS'));
  const sheet = new FakeSheet('FU', 907, [headers, ['SYN-1', '甲']]);
  const protection = prepareIdentitySheet(sheet);
  const columns = call('getRequiredFuColumns_', sheet);
  call('hideSystemColumns_', sheet, columns, ['EVENT_ID', 'SYNC_ID']);
  assert.strictEqual(protection.protections.length, 0);
  assert.deepStrictEqual(protection.hidden, [columns.EVENT_ID, columns.SYNC_ID]);
  const sync = { description: 'SURGERY_SYSTEM_COLUMN:SYNC_ID', warningOnly: false,
    setRange() { throw new Error('Owner range must remain unchanged'); },
    setWarningOnly() { throw new Error('Owner policy must remain unchanged'); },
    remove() { throw new Error('Owner protection must remain unchanged'); } };
  protection.protections.push(sync);
  call('hideSystemColumns_', sheet, columns, ['EVENT_ID', 'SYNC_ID']);
  assert.strictEqual(protection.protections.length, 1);
  assert.strictEqual(protection.protections[0], sync);
  assert.strictEqual(sync.warningOnly, false);
}

function testAnnualArchiveOnlyHidesTechnicalColumnsAndPreservesProtections() {
  const headers = call('getNewAnnualArchiveHeaders_', []);
  const sheet = new FakeSheet('年度封存_2025', 908, [[], [], [], [], [], headers]);
  const state = prepareIdentitySheet(sheet);
  const owner = { description: 'SURGERY_ANNUAL_ARCHIVE:SYSTEM:1', warningOnly: false,
    remove() { throw new Error('Owner protection must remain unchanged'); } };
  state.protections.push(owner);
  call('hideAnnualArchiveSystemColumns_', sheet);
  const info = call('getAnnualArchiveHeaderInfo_', sheet);
  assert.deepStrictEqual(state.hidden, [info.eventColumn, info.syncColumn,
    info.sourceMonthColumn, info.rowTypeColumn].sort((left, right) => left - right));
  assert.deepStrictEqual(state.protections, [owner]);
  files.forEach(file => assert.strictEqual(/\.protect\s*\(|\.setWarningOnly\s*\(|\.getProtections\s*\(/.test(sourceByFile[file]), false));
}

function testMissingHeadersPointToTheAppropriateRepairOwner() {
  ['FU', 'MONTHLY'].forEach(kind => {
    const headers = plain(evaluate(kind === 'FU' ? 'CONFIG.HEADERS' : 'CONFIG.MONTHLY_HEADERS'));
    const fn = kind === 'FU' ? 'getRequiredFuColumns_' : 'getRequiredMonthlyColumns_';
    const name = kind === 'FU' ? 'FU' : '202610';
    const technical = new FakeSheet(name, 920, [headers.filter(value => value !== 'ScheduleSyncId')]);
    assert.throws(() => call(fn, technical), error => /檢查並準備刀表/.test(error.message) && !/修復刀表格式/.test(error.message));
    const clinical = new FakeSheet(name, 921, [headers.filter(value => value !== '姓名')]);
    assert.throws(() => call(fn, clinical), error => /補回第一列表頭/.test(error.message) && !/檢查並準備刀表/.test(error.message));
  });
}

function testFormatRepairOnlyTouchesExistingSheetsWithoutAssigningIds() {
  const savedSpreadsheetApp = { ...context.SpreadsheetApp };
  const savedScriptApp = context.ScriptApp;
  const savedUuid = context.Utilities.getUuid;
  const menus = [];
  let inserted = 0, triggersCreated = 0;
  const menu = { addItem: (_label, fn) => { menus.push(fn); return menu; },
    addSeparator: () => menu, addSubMenu: () => menu, addToUi: () => menu };
  const ui = { ButtonSet: { OK: 'OK' }, createMenu: () => menu, alert: () => {} };
  const conditional = () => { const builder = {
    whenFormulaSatisfied: () => builder, setBackground: () => builder,
    setBold: () => builder, setRanges: () => builder, build: () => ({})
  }; return builder; };
  context.SpreadsheetApp.getUi = () => ui;
  context.SpreadsheetApp.newConditionalFormatRule = conditional;
  context.ScriptApp = { getProjectTriggers: () => [], newTrigger: () => {
    const builder = { forSpreadsheet: () => builder, onEdit: () => builder,
      create: () => { triggersCreated++; } }; return builder;
  } };
  context.Utilities.getUuid = () => { throw new Error('Repair must not assign identities'); };
  try {
    [false, true].forEach(includeFu => {
      const headers = [...plain(evaluate('CONFIG.MONTHLY_HEADERS')), '人工欄'];
      const row = fields => headers.map(header => fields[header] || '');
      const month = new FakeSheet('202610', 922, [headers,
        row({ '日期／報到時間': new Date(2026, 9, 12), '醫院': '高榮', '病歷號': '◆ 刀日' }),
        row({ '日期／報到時間': 830, '病歷號': 'SYN-REPAIR', '姓名': '合成病人', '人工欄': '保留' })]);
      const sheets = [month];
      if (includeFu) sheets.push(new FakeSheet('FU', 923,
        [plain(evaluate('CONFIG.HEADERS')), ['SYN-FU', '合成追蹤']]));
      month.setFormulaAt(3, headers.length, '=1+1');
      month.setNoteAt(3, headers.length, '人工註記');
      const before = JSON.stringify(sheets.map(sheet => sheet.rows));
      const protections = sheets.map(sheet => {
        const state = prepareIdentitySheet(sheet);
        state.protections.push({ warningOnly: false, description: 'owner' });
        const getRange = sheet.getRange.bind(sheet);
        sheet.getRange = (...args) => {
          const range = getRange(...args);
          range.setBackground = () => range;
          range.setFontColor = () => range;
          return range;
        };
        return state;
      });
      const spreadsheet = makeFakeSpreadsheet(sheets);
      spreadsheet.insertSheet = () => { inserted++; throw new Error('Repair must not create sheets'); };
      context.SpreadsheetApp.getActiveSpreadsheet = () => spreadsheet;
      const result = call('repairScheduleFormatting');
      assert.strictEqual(result.ok, true);
      assert.strictEqual(Boolean(result.fuResult.skipped), !includeFu);
      assert.strictEqual(result.monthlyResult.count, 1);
      assert.strictEqual(JSON.stringify(sheets.map(sheet => sheet.rows)), before);
      assert.strictEqual(month.formulaAt(3, headers.length), '=1+1');
      assert.strictEqual(month.noteAt(3, headers.length), '人工註記');
      assert.ok(month.getConditionalFormatRules().length > 0, 'Existing month formatting is repaired');
      assert.ok(protections.every(state => state.hidden.length === 2 && state.protections.length === 1 && !state.protections[0].warningOnly));
    });
    assert.strictEqual(inserted, 0);
    assert.strictEqual(triggersCreated, 2);
    menus.forEach(name => assert.strictEqual(evaluate(`typeof ${name}`), 'function'));
  } finally {
    context.SpreadsheetApp = savedSpreadsheetApp;
    context.ScriptApp = savedScriptApp;
    context.Utilities.getUuid = savedUuid;
  }
}

const tests = [
  testMissingHeadersPointToTheAppropriateRepairOwner,
  testFormatRepairOnlyTouchesExistingSheetsWithoutAssigningIds,
  testSystemColumnsAreHiddenWithoutCreatingOrChangingProtections,
  testAnnualArchiveOnlyHidesTechnicalColumnsAndPreservesProtections,
  testObsoleteCalendarMenusAndInternalScaffoldingAreRemoved,
  testIdentityScanDiagnosesExistingDuplicatesWithoutMigration,
  testInsertDayWritesSelectedCampusAndRejectsDuplicateIdentityBeforeWriting,
  testTriggerRepairRemovesLegacyWritersAndPreservesOtherTriggersAndHandover,
  testVersionAndModuleSplit,
  testEightDigitDateInputOnlyInFunctions,
  testIolListIsReadOnlyAndKeepsAllDateBlocks,
  testIolListReadsOnlyRequestedDateDisplayRows,
  testOnlyCalendarEventIdIsCanonicalSystemField,
  testCurrentMenuHasNoCompletedMigrationOrLegacyOutput,
  testMonthlySheetNameRecognition,
  testTimeNormalizationVariants,
  testMultiRowPasteNormalizesEveryMonthlyTimeCell,
  testWholeMonthlyTimeNormalizationUsesBoundedBatchReads,
  testUnifiedSortRoutesWithoutChangingSortImplementations,
  testFullDateDoesNotCollideWithTime,
  testMonthlyDateHeaderUsesDateAndBlankEventId,
  testMonthlyHeaderMarkerPreservesOtherCells,
  testMonthlyHeaderMarkerNeverOverwritesExistingChartCell,
  testMonthlyStructureIssuesReportHybridAndMissingMarker,
  testDuplicateMonthlyBlocksAreIndependentAndSupported,
  testSurgeryDateInsertionPreviewSplitsOnlyRowsBelowSelection,
  testBlockRefRequiresExactHeaderRowAndIdentity,
  testFuNativeSortUsesOneSortAndPreservesNotes,
  testFuNativeSortSkipsAlreadySortedRows,
  testTemporaryNativeSortRollsBackOnVerificationFailure,
  testMonthlyNativeSortPreservesBlocksAndUsesOneSort,
  testMonthlyNativeSortSkipsAlreadySortedSheet,
  testFuToMonthlyReusesPrecomputedMonthScan,
  testManagedRowFormattingUsesRangeListsAndLeavesCustomColumnAlone,
  testTouchedRowFormattingOnlyChangesEditedStandardColumns,
  testDefaultMonthlyDates,
  testMonthlyConditionComposition,
  testArbitraryHeaderOrderUsesNames,
  testDuplicateHeadersAreDiagnosed,
  testMonthlyDiagnosisSummaryFormulaDefinition,
  testMonthlyDiagnosisSummaryHeadersStayCanonical,
  testMonthlyDiagnosisSummaryCreationAndManagedRefresh,
  testAnnualArchiveCanonicalizesDiagnosisSummaryHeader,
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
  testColumnWidthsRequireExplicitRequestAfterCreation,
  testFormatRepairDoesNotInitializeOrSynchronize,
  testSystemNotesPreserveManualNotes,
  testHtmlUsesCollapsedAdvancedAreaAndArchiveLanguage,
  testClaspIncludesAllRuntimeModules,
  testNewPatientInsertionPreservesManualBlankRowAndGetsNewIdentity,
  testScheduleIdentitySurvivesNativeSortAndWholeDateChanges,
  testFormattingOnlyDeploymentHasNoCalendarApiAndLegacyCallsAreSafe,
  testExternalOwnerPreventsCalendarMutationsAndKeepsEditFormatting,
  testDuplicateEventBindingStopsSheetOperations,
  testUnnamedTailDataSurvivesMonthlySorting,
  testFuRequestCacheUsesPermanentIdentityAfterMoveAndCalendarReplacement,
  testSurgeryDatePreviewSelectsCampusAndDetectsChange,
];

module.exports = { call, FakeSheet, evaluate, plain, makeFakeSpreadsheet, context };
if (require.main === module) {
  tests.forEach(test => test());
  console.log(`smoke tests passed (${tests.length})`);
}
