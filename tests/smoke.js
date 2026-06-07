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
    getScriptTimeZone: () => 'Asia/Taipei',
    getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' })
  },
  Utilities: {
    formatDate: (date, _timezone, format) => {
      const yyyy = date.getFullYear();
      const mm = pad2(date.getMonth() + 1);
      const dd = pad2(date.getDate());
      const hh = pad2(date.getHours());
      const min = pad2(date.getMinutes());

      if (format === 'yyyy-MM-dd') {
        return `${yyyy}-${mm}-${dd}`;
      }

      if (format === 'HH:mm') {
        return `${hh}:${min}`;
      }

      if (format === "yyyy-MM-dd'T'HH:mm:ssXXX") {
        return `${yyyy}-${mm}-${dd}T${hh}:${min}:00+08:00`;
      }

      return `${yyyy}/${mm}/${dd}`;
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

function makeDataSheet(headers, rows, sheetName) {
  const data = [headers].concat(rows.map(row => row.slice()));
  const notes = {};
  const hiddenColumns = [];
  const protections = [];
  let conditionalRules = [];

  return {
    rows: data,
    notes,
    hiddenColumns,
    protections,
    getName: () => sheetName || '',
    getMaxRows: () => data.length,
    getMaxColumns: () => data[0].length,
    getLastColumn: () => data[0].length,
    getLastRow: () => data.length,
    insertColumnsAfter: (column, count) => {
      data.forEach(row => {
        for (let i = 0; i < count; i++) {
          row.splice(column, 0, '');
        }
      });
    },
    insertColumnBefore: column => {
      data.forEach(row => row.splice(column - 1, 0, ''));
    },
    insertRowsAfter: (row, count) => {
      for (let i = 0; i < count; i++) {
        data.splice(row, 0, new Array(data[0].length).fill(''));
      }
    },
    insertRowAfter: row => {
      data.splice(row, 0, new Array(data[0].length).fill(''));
    },
    deleteColumn: column => {
      data.forEach(row => row.splice(column - 1, 1));
    },
    hideColumns: column => {
      hiddenColumns.push(column);
    },
    getProtections: () => protections,
    getConditionalFormatRules: () => conditionalRules,
    setConditionalFormatRules: nextRules => {
      conditionalRules = nextRules;
    },
    clear: () => {
      data.forEach(row => {
        for (let i = 0; i < row.length; i++) {
          row[i] = '';
        }
      });
    },
    autoResizeColumns: () => {},
    getRange: (row, column, numRows, numColumns) => {
      const height = numRows || 1;
      const width = numColumns || 1;
      const range = {
        getValues: () => data
          .slice(row - 1, row - 1 + height)
          .map(dataRow => dataRow.slice(column - 1, column - 1 + width)),
        setValues: values => {
          values.forEach((valueRow, rowOffset) => {
            while (data.length <= row - 1 + rowOffset) {
              data.push(new Array(data[0].length).fill(''));
            }
            valueRow.forEach((value, columnOffset) => {
              while (data[row - 1 + rowOffset].length <= column - 1 + columnOffset) {
                data[row - 1 + rowOffset].push('');
              }
              data[row - 1 + rowOffset][column - 1 + columnOffset] = value;
            });
          });
          return range;
        },
        getValue: () => data[row - 1] && data[row - 1][column - 1],
        setValue: value => {
          while (data.length <= row - 1) {
            data.push(new Array(data[0].length).fill(''));
          }
          while (data[row - 1].length <= column - 1) {
            data[row - 1].push('');
          }
          data[row - 1][column - 1] = value;
          return range;
        },
        clearContent: () => {
          for (let rowOffset = 0; rowOffset < height; rowOffset++) {
            while (data.length <= row - 1 + rowOffset) {
              data.push(new Array(data[0].length).fill(''));
            }
            for (let columnOffset = 0; columnOffset < width; columnOffset++) {
              while (data[row - 1 + rowOffset].length <= column - 1 + columnOffset) {
                data[row - 1 + rowOffset].push('');
              }
              data[row - 1 + rowOffset][column - 1 + columnOffset] = '';
            }
          }
          return range;
        },
        copyTo: targetRange => {
          targetRange.setValues(range.getValues());
          return range;
        },
        setNumberFormat: () => range,
        setDataValidation: () => range,
        setHorizontalAlignment: () => range,
        setVerticalAlignment: () => range,
        setFontWeight: () => range,
        setBackground: () => range,
        getNote: () => notes[`${row - 1}:${column - 1}`] || '',
        clearNote: () => {
          notes[`${row - 1}:${column - 1}`] = '';
          return range;
        },
        setNote: note => {
          notes[`${row - 1}:${column - 1}`] = note;
          return range;
        },
        protect: () => {
          const protection = {
            description: '',
            editors: [],
            removedEditors: [],
            warningOnly: null,
            domainEdit: true,
            range,
            getDescription: () => protection.description,
            setDescription: description => {
              protection.description = description;
              return protection;
            },
            setWarningOnly: value => {
              protection.warningOnly = value;
              return protection;
            },
            addEditor: editor => {
              if (protection.editors.indexOf(editor) === -1) {
                protection.editors.push(editor);
              }
              return protection;
            },
            getEditors: () => protection.editors.slice(),
            removeEditors: editors => {
              protection.removedEditors = protection.removedEditors.concat(editors);
              protection.editors = protection.editors.filter(editor => editors.indexOf(editor) === -1);
              return protection;
            },
            canDomainEdit: () => protection.domainEdit,
            setDomainEdit: value => {
              protection.domainEdit = value;
              return protection;
            },
            remove: () => {
              const index = protections.indexOf(protection);
              if (index !== -1) protections.splice(index, 1);
            }
          };
          protections.push(protection);
          return protection;
        },
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
      };

      return range;
    }
  };
}

function rowFromObject(headers, values) {
  return headers.map(header => Object.prototype.hasOwnProperty.call(values, header) ? values[header] : '');
}

function toPlain(value) {
  return JSON.parse(JSON.stringify(value));
}

function testHeaderMappingWithCustomColumn() {
  const headers = ['病歷號', '姓名', '自訂欄', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
  const info = context.buildFieldColumnInfo_(makeHeaderSheet(headers));

  assert.deepStrictEqual(Object.assign({}, info.columns), {
    CHART_NO: 1,
    NAME: 2,
    TEL: 4,
    HOSPITAL: 5,
    TAG: 6,
    COND: 7,
    DATE: 8,
    TIME: 9,
    PLAN: 10,
    MEMO: 11,
    EVENT_ID: 12,
    SHEET_WRITE_UPDATED: 13
  });
}

function testEventIdColumnLookupWithMovedHeader() {
  const headers = ['病歷號', 'CalendarEventId', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarSheetWriteUpdated'];
  assert.strictEqual(context.getEventIdColumnForSheet_(makeHeaderSheet(headers)), 2);
}

function testLegacyEventIdColumnLookupFallback() {
  const headers = ['病歷號', '日曆eventID', '姓名', 'TEL', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得'];
  assert.strictEqual(context.getEventIdColumnForSheet_(makeHeaderSheet(headers)), 2);
}

function testEventIdColumnLookupMissingHeader() {
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得'];
  assert.strictEqual(context.getEventIdColumnForSheet_(makeHeaderSheet(headers)), 0);
}

function testLegacyEventIdMigrationToEnglishHeader() {
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', '日曆eventID'];
  const sheet = makeDataSheet(headers, [
    rowFromObject(headers, { '病歷號': '001', '姓名': '王小明', '日曆eventID': 'evt_1' }),
    rowFromObject(headers, { '病歷號': '002', '姓名': '陳小美' })
  ]);

  const headerResult = context.ensureHeaders_(sheet);
  const migrationResult = context.migrateLegacyEventIdColumns_(sheet);
  const info = context.buildFieldColumnInfo_(sheet);

  assert.ok(headerResult.addedHeaders.includes('CalendarEventId'));
  assert.ok(headerResult.addedHeaders.includes('CalendarSheetWriteUpdated'));
  assert.ok(migrationResult.message.includes('補回 1 格'));
  assert.strictEqual(sheet.rows[0].includes('日曆eventID'), false);
  assert.strictEqual(context.getRowFieldValue_(sheet.rows[1], info.columns, 'EVENT_ID'), 'evt_1');
  assert.ok(info.columns.SHEET_WRITE_UPDATED);
}

function testHospitalColumnInsertedBeforeTagPreservesData() {
  const headers = ['病歷號', '姓名', 'TEL', 'Tag', 'Condition'];
  const sheet = makeDataSheet(headers, [
    rowFromObject(headers, { '病歷號': '001', '姓名': '王小明', TEL: '0911', Tag: 'OP', Condition: 'Cataract' })
  ]);

  const result = context.ensureHeaders_(sheet);
  const hospitalIndex = sheet.rows[0].indexOf('醫院');
  const tagIndex = sheet.rows[0].indexOf('Tag');

  assert.ok(result.addedHeaders.includes('醫院'));
  assert.strictEqual(hospitalIndex + 1, tagIndex);
  assert.strictEqual(sheet.rows[1][hospitalIndex], '');
  assert.strictEqual(sheet.rows[1][tagIndex], 'OP');
}

function testHospitalColumnMovedBeforeTagPreservesData() {
  const headers = ['病歷號', '姓名', 'TEL', 'Tag', '醫院', 'Condition'];
  const sheet = makeDataSheet(headers, [
    rowFromObject(headers, { '病歷號': '001', '姓名': '王小明', TEL: '0911', '醫院': '聯醫', Tag: 'OP', Condition: 'Cataract' })
  ]);

  const result = context.ensureHeaders_(sheet);
  const hospitalIndex = sheet.rows[0].indexOf('醫院');
  const tagIndex = sheet.rows[0].indexOf('Tag');

  assert.ok(result.migrationMessages.some(message => message.includes('移到')));
  assert.strictEqual(hospitalIndex + 1, tagIndex);
  assert.strictEqual(sheet.rows[1][hospitalIndex], '聯醫');
  assert.strictEqual(sheet.rows[1][tagIndex], 'OP');
}

function testDuplicateHospitalHeadersAreReportedForManualReview() {
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', '醫院', 'Condition'];
  const sheet = makeDataSheet(headers, [
    rowFromObject(headers, { '病歷號': '001', '姓名': '王小明', TEL: '0911', '醫院': '高榮', Tag: 'OP', Condition: 'Cataract' })
  ]);

  const result = context.ensureHeaders_(sheet);

  assert.ok(result.migrationMessages.some(message => message.includes('請手動確認')));
  assert.strictEqual(sheet.rows[0].filter(header => header === '醫院').length, 2);
}

function testHospitalClassificationDefaultsToKaohsiungVeterans() {
  assert.strictEqual(context.classifyHospital_(''), '高榮');
  assert.strictEqual(context.classifyHospital_('高榮'), '高榮');
  assert.strictEqual(context.classifyHospital_('其他'), '高榮');
  assert.strictEqual(context.classifyHospital_('聯醫'), '聯醫');
}

function testMissingAllWithLegacyOpPromptsManualRename() {
  const spreadsheet = {
    getSheetByName: name => name === 'OP' ? {} : null
  };
  const message = context.buildMissingMainSheetMessage_(spreadsheet, '初始化');

  assert.ok(message.includes('OP'));
  assert.ok(message.includes('All'));
  assert.ok(message.includes('手動'));
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

function testCalendarTitleKeepsBlankPatientNameSlot() {
  const title = context.buildCalendarTitle_('001', '', 'Cataract');
  const titleInfo = context.parseCalendarTitle_(title);

  assert.strictEqual(title, '001 |  | Cataract');
  assert.strictEqual(titleInfo.ok, true);
  assert.strictEqual(titleInfo.chartNo, '001');
  assert.strictEqual(titleInfo.patientName, '');
  assert.strictEqual(titleInfo.condition, 'Cataract');
}

function testSheetOriginatedCalendarEchoIsSkipped() {
  const originalApplyRowDataFormats = context.applyRowDataFormats_;
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
  const sheet = makeDataSheet(headers, [
    rowFromObject(headers, {
      '病歷號': '001',
      '姓名': '王小明',
      TEL: '0911',
      Tag: 'OP',
      Condition: 'Cataract',
      '日期': new Date(2026, 4, 9),
      '時間': '',
      Plan: 'P1',
      CalendarEventId: 'evt_1',
      CalendarSheetWriteUpdated: '2026-05-09T00:00:01.000Z'
    })
  ]);

  try {
    context.applyRowDataFormats_ = () => {};
    const rowIndex = context.buildCalendarEventRowIndex_(sheet);
    const status = context.processCalendarApiEventChange_(sheet, rowIndex, {
      id: 'evt_1',
      updated: '2026-05-09T00:00:01.000Z',
      summary: '001 | 王小明 | Cataract',
      description: 'TEL: 0911\nPlan: P1',
      colorId: '5',
      start: {
        date: '2026-05-09'
      },
      end: {
        date: '2026-05-10'
      }
    }, 'calendar_1');

    assert.strictEqual(status, 'skipped_sheet_echo');
    assert.strictEqual(context.getRowFieldValue_(sheet.rows[1], rowIndex.columns, 'TIME'), '');
  } finally {
    context.applyRowDataFormats_ = originalApplyRowDataFormats;
  }
}

function testSameTimestampDifferentCalendarContentIsNotSkipped() {
  const originalApplyRowDataFormats = context.applyRowDataFormats_;
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
  const sheet = makeDataSheet(headers, [
    rowFromObject(headers, {
      '病歷號': '001',
      '姓名': '王小明',
      TEL: '0911',
      Tag: 'OP',
      Condition: 'Cataract',
      '日期': new Date(2026, 4, 9),
      '時間': '',
      Plan: 'P1',
      CalendarEventId: 'evt_1',
      CalendarSheetWriteUpdated: '2026-05-09T00:00:01.000Z'
    })
  ]);

  try {
    context.applyRowDataFormats_ = () => {};
    const rowIndex = context.buildCalendarEventRowIndex_(sheet);
    const status = context.processCalendarApiEventChange_(sheet, rowIndex, {
      id: 'evt_1',
      updated: '2026-05-09T00:00:01.000Z',
      summary: '001 | 王小明 | Retina',
      description: 'TEL: 0911\nPlan: P2',
      colorId: '11',
      start: {
        date: '2026-05-09'
      },
      end: {
        date: '2026-05-10'
      }
    }, 'calendar_1');

    assert.strictEqual(status, 'updated');
    assert.strictEqual(context.getRowFieldValue_(sheet.rows[1], rowIndex.columns, 'COND'), 'Retina');
    assert.strictEqual(context.getRowFieldValue_(sheet.rows[1], rowIndex.columns, 'PLAN'), 'P2');
  } finally {
    context.applyRowDataFormats_ = originalApplyRowDataFormats;
  }
}

function testOlderAllDayEchoAfterTimedSheetUpdateIsSkipped() {
  const originalApplyRowDataFormats = context.applyRowDataFormats_;
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
  const sheet = makeDataSheet(headers, [
    rowFromObject(headers, {
      '病歷號': '001',
      '姓名': '王小明',
      TEL: '0911',
      Tag: 'OP',
      Condition: 'Cataract',
      '日期': new Date(2026, 4, 9),
      '時間': '08:30',
      Plan: 'P1',
      CalendarEventId: 'evt_1',
      CalendarSheetWriteUpdated: '2026-05-09T00:00:02.000Z'
    })
  ]);

  try {
    context.applyRowDataFormats_ = () => {};
    const rowIndex = context.buildCalendarEventRowIndex_(sheet);
    const status = context.processCalendarApiEventChange_(sheet, rowIndex, {
      id: 'evt_1',
      updated: '2026-05-09T00:00:01.000Z',
      summary: '001 | 王小明 | Cataract',
      description: 'TEL: 0911\nPlan: P1',
      start: {
        date: '2026-05-09'
      }
    }, 'calendar_1');

    assert.strictEqual(status, 'skipped_sheet_echo');
    assert.strictEqual(context.getRowFieldValue_(sheet.rows[1], rowIndex.columns, 'TIME'), '08:30');
  } finally {
    context.applyRowDataFormats_ = originalApplyRowDataFormats;
  }
}

function testNewerCalendarAllDayUpdateClearsTime() {
  const originalApplyRowDataFormats = context.applyRowDataFormats_;
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
  const sheet = makeDataSheet(headers, [
    rowFromObject(headers, {
      '病歷號': '001',
      '姓名': '王小明',
      TEL: '0911',
      Tag: 'OP',
      Condition: 'Cataract',
      '日期': new Date(2026, 4, 9),
      '時間': '08:30',
      Plan: 'P1',
      CalendarEventId: 'evt_1',
      CalendarSheetWriteUpdated: '2026-05-09T00:00:01.000Z'
    })
  ]);

  try {
    context.applyRowDataFormats_ = () => {};
    const rowIndex = context.buildCalendarEventRowIndex_(sheet);
    const status = context.processCalendarApiEventChange_(sheet, rowIndex, {
      id: 'evt_1',
      updated: '2026-05-09T00:00:02.000Z',
      summary: '001 | 王小明 | Cataract',
      description: 'TEL: 0911\nPlan: P1',
      colorId: '11',
      start: {
        date: '2026-05-09'
      }
    }, 'calendar_1');

    assert.strictEqual(status, 'updated');
    assert.strictEqual(context.getRowFieldValue_(sheet.rows[1], rowIndex.columns, 'TIME'), '');
  } finally {
    context.applyRowDataFormats_ = originalApplyRowDataFormats;
  }
}

function testMergeBlocksTrackedAbsorbedRows() {
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
  const cols = context.buildFieldColumnInfo_(makeHeaderSheet(headers)).columns;
  const data = [
    headers,
    rowFromObject(headers, { '病歷號': '123', '姓名': '王小明', Tag: 'FU', TEL: '0911', Condition: 'A' }),
    rowFromObject(headers, { '病歷號': '123', '姓名': '王小明', Tag: 'FU', TEL: '0922', Condition: 'B', 'CalendarEventId': 'evt_1' })
  ];
  const result = context.buildMergedPatientRows_(data, cols);

  assert.strictEqual(result.ok, false);
  assert.deepStrictEqual(Array.from(result.blockedRows), [3]);
}

function testMergeCombinesSafeTextFields() {
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
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
    newDataValidation: () => ({
      requireValueInList() {
        return this;
      },
      setAllowInvalid() {
        return this;
      },
      build() {
        return {};
      }
    }),
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

function makeSpreadsheet(sheets) {
  return {
    sheets,
    getSheetByName: name => sheets[name] || null,
    insertSheet: name => {
      const sheet = makeDataSheet([''], [], name);
      sheets[name] = sheet;
      return sheet;
    }
  };
}

function testConditionalFormattingKeepsTagIndependent() {
  installConditionalFormatMock();
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
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

function testSystemColumnsAreHiddenAndProtected() {
  const originalSpreadsheetApp = context.SpreadsheetApp;
  const originalPropertiesService = context.PropertiesService;
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
  const sheet = makeDataSheet(headers, []);
  const columns = context.buildFieldColumnInfo_(makeHeaderSheet(headers)).columns;

  try {
    context.SpreadsheetApp = {
      ProtectionType: {
        RANGE: 'RANGE'
      }
    };
    context.PropertiesService = {
      getScriptProperties: () => ({
        getProperty: key => key === 'SYSTEM_COLUMN_PROTECTION_EDITORS'
          ? 'maintainer@example.com; owner@example.com'
          : ''
      })
    };

    context.hideSystemColumns_(sheet, columns);
    const result = context.protectSystemColumns_(sheet, columns);
    const secondResult = context.protectSystemColumns_(sheet, columns);

    assert.strictEqual(result.ok, true);
    assert.strictEqual(secondResult.ok, true);
    assert.deepStrictEqual(sheet.hiddenColumns, [columns.EVENT_ID, columns.SHEET_WRITE_UPDATED]);
    assert.deepStrictEqual(
      sheet.protections.map(protection => protection.description),
      ['OP_SPREADSHEET_SYSTEM_COLUMN:EVENT_ID', 'OP_SPREADSHEET_SYSTEM_COLUMN:SHEET_WRITE_UPDATED']
    );
    assert.ok(sheet.protections.every(protection => protection.warningOnly === false));
    assert.ok(sheet.protections.every(protection => protection.domainEdit === false));
    assert.ok(sheet.protections.every(protection => protection.editors.some(editor => editor.getEmail && editor.getEmail() === 'owner@example.com')));
    assert.ok(sheet.protections.every(protection => protection.editors.includes('maintainer@example.com')));
    assert.ok(sheet.protections.every(protection => protection.removedEditors.every(editor => !(editor.getEmail && editor.getEmail() === 'owner@example.com'))));
  } finally {
    context.SpreadsheetApp = originalSpreadsheetApp;
    context.PropertiesService = originalPropertiesService;
  }
}

function testHospitalMirrorSheetsFilterRowsFromAll() {
  const originalSpreadsheetApp = context.SpreadsheetApp;
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
  const allSheet = makeDataSheet(headers, [
    rowFromObject(headers, { '病歷號': '001', '姓名': '高榮空白', Tag: 'OP' }),
    rowFromObject(headers, { '病歷號': '002', '姓名': '聯醫病人', '醫院': '聯醫', Tag: 'OP' }),
    rowFromObject(headers, { '病歷號': '003', '姓名': '其他醫院', '醫院': '其他', Tag: 'OP' }),
    rowFromObject(headers, { '病歷號': '004', '姓名': '非手術', '醫院': '聯醫', Tag: 'FU' })
  ], 'All');
  const spreadsheet = makeSpreadsheet({ All: allSheet });

  try {
    installConditionalFormatMock();
    context.SpreadsheetApp.getActiveSpreadsheet = () => spreadsheet;

    const result = context.refreshHospitalMirrorSheets_();

    assert.strictEqual(result.ok, true);
    assert.deepStrictEqual(
      spreadsheet.sheets['OP-高榮'].rows.slice(1).map(row => row[0]),
      ['001', '003']
    );
    assert.deepStrictEqual(
      spreadsheet.sheets['OP-聯醫'].rows.slice(1).map(row => row[0]),
      ['002']
    );
  } finally {
    context.SpreadsheetApp = originalSpreadsheetApp;
  }
}

function testProcessCalendarChangeIsNoOpWhenReverseSyncDisabled() {
  const originalWithCalendarSyncLock = context.withCalendarSyncLock_;
  const originalSyncCalendarChangesToSheet = context.syncCalendarChangesToSheet_;
  const originalScheduleUpcomingWeekExport = context.scheduleUpcomingWeekExport_;
  let lockCalled = false;
  let syncCalled = false;
  let exportCalled = false;

  try {
    context.withCalendarSyncLock_ = () => {
      lockCalled = true;
      return true;
    };
    context.syncCalendarChangesToSheet_ = () => {
      syncCalled = true;
      return { ok: true, processedCount: 1 };
    };
    context.scheduleUpcomingWeekExport_ = () => {
      exportCalled = true;
    };

    const result = context.processCalendarChange({ calendarId: 'calendar_1' });

    assert.strictEqual(result.status, 'disabled');
    assert.strictEqual(lockCalled, false);
    assert.strictEqual(syncCalled, false);
    assert.strictEqual(exportCalled, false);
  } finally {
    context.withCalendarSyncLock_ = originalWithCalendarSyncLock;
    context.syncCalendarChangesToSheet_ = originalSyncCalendarChangesToSheet;
    context.scheduleUpcomingWeekExport_ = originalScheduleUpcomingWeekExport;
  }
}

function testSyncCalendarChangesToSheetIsDisabled() {
  const result = context.syncCalendarChangesToSheet_('calendar_1');

  assert.strictEqual(result.ok, true);
  assert.strictEqual(result.status, 'disabled');
  assert.strictEqual(result.processedCount, 0);
}

function testDisableCalendarReverseSyncDeletesOnlyReverseTriggersAndTokens() {
  const originalScriptApp = context.ScriptApp;
  const originalPropertiesService = context.PropertiesService;
  const deletedHandlers = [];
  const properties = {
    CALENDAR_SYNC_TOKEN_calendar_1: 'token_1',
    CALENDAR_SYNC_TOKEN_calendar_2: 'token_2',
    CALENDAR_ID: 'calendar_1'
  };
  const triggers = [
    { getHandlerFunction: () => 'processCalendarChange' },
    { getHandlerFunction: () => 'processRowChange' },
    { getHandlerFunction: () => 'runPendingUpcomingWeekExport' }
  ];

  try {
    context.ScriptApp = {
      getProjectTriggers: () => triggers,
      deleteTrigger: trigger => {
        deletedHandlers.push(trigger.getHandlerFunction());
      }
    };
    context.PropertiesService = {
      getScriptProperties: () => ({
        getKeys: () => Object.keys(properties),
        deleteProperty: key => {
          delete properties[key];
        }
      })
    };

    const result = context.disableCalendarReverseSync_(false);

    assert.strictEqual(result.ok, true);
    assert.strictEqual(result.deletedTriggerCount, 1);
    assert.strictEqual(result.deletedTokenCount, 2);
    assert.deepStrictEqual(deletedHandlers, ['processCalendarChange']);
    assert.deepStrictEqual(Object.assign({}, properties), { CALENDAR_ID: 'calendar_1' });
  } finally {
    context.ScriptApp = originalScriptApp;
    context.PropertiesService = originalPropertiesService;
  }
}

function testDuplicateRowClearsSystemTrackingFields() {
  const originalSpreadsheetApp = context.SpreadsheetApp;
  const originalRefreshHospitalMirrorSheetsSafely = context.refreshHospitalMirrorSheetsSafely_;
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
  const sheet = makeDataSheet(headers, [
    rowFromObject(headers, {
      '病歷號': '001',
      '姓名': '王小明',
      '日期': new Date(2026, 4, 9),
      '時間': '08:30',
      CalendarEventId: 'evt_1',
      CalendarSheetWriteUpdated: '2026-05-09T00:00:01.000Z'
    })
  ]);

  sheet.getName = () => 'All';
  sheet.getActiveCell = () => ({ getRow: () => 2 });

  try {
    context.SpreadsheetApp = {
      getUi: () => ({
        ButtonSet: { OK: 'OK' },
        alert: () => {}
      }),
      getActiveSpreadsheet: () => ({
        getActiveSheet: () => sheet
      })
    };
    context.refreshHospitalMirrorSheetsSafely_ = () => ({ ok: true });

    context.duplicateRow();

    const columns = context.buildFieldColumnInfo_(sheet).columns;
    assert.strictEqual(context.getRowFieldValue_(sheet.rows[2], columns, 'DATE'), '');
    assert.strictEqual(context.getRowFieldValue_(sheet.rows[2], columns, 'TIME'), '');
    assert.strictEqual(context.getRowFieldValue_(sheet.rows[2], columns, 'EVENT_ID'), '');
    assert.strictEqual(context.getRowFieldValue_(sheet.rows[2], columns, 'SHEET_WRITE_UPDATED'), '');
  } finally {
    context.SpreadsheetApp = originalSpreadsheetApp;
    context.refreshHospitalMirrorSheetsSafely_ = originalRefreshHospitalMirrorSheetsSafely;
  }
}

function testDuplicateCalendarEventIdIsDiagnosedAndSkipped() {
  const originalApplyRowDataFormats = context.applyRowDataFormats_;
  const originalConsoleWarn = context.console.warn;
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
  const sheet = makeDataSheet(headers, [
    rowFromObject(headers, {
      '病歷號': '001',
      '姓名': '王小明',
      TEL: '0911',
      Tag: 'OP',
      Condition: 'Cataract',
      '日期': new Date(2026, 4, 9),
      '時間': '',
      Plan: 'P1',
      CalendarEventId: 'evt_1'
    }),
    rowFromObject(headers, {
      '病歷號': '002',
      '姓名': '陳小美',
      TEL: '0922',
      Tag: 'OP',
      Condition: 'Retina',
      '日期': new Date(2026, 4, 10),
      '時間': '',
      Plan: 'P2',
      CalendarEventId: 'evt_1'
    })
  ]);

  try {
    context.applyRowDataFormats_ = () => {};
    context.console.warn = () => {};
    const rowIndex = context.buildCalendarEventRowIndex_(sheet);
    const status = context.processCalendarApiEventChange_(sheet, rowIndex, {
      id: 'evt_1',
      updated: '2026-05-09T00:00:02.000Z',
      summary: '001 | 王小明 | Cataract',
      description: 'TEL: 0911\nPlan: P1',
      colorId: '11',
      start: {
        date: '2026-05-09'
      },
      end: {
        date: '2026-05-10'
      }
    }, 'calendar_1');

    assert.strictEqual(status, 'skipped_duplicate_event_id');
    assert.deepStrictEqual(Array.from(rowIndex.duplicateEventIds.evt_1), [2, 3]);
    assert.ok(sheet.getRange(2, rowIndex.columns.EVENT_ID).getNote().includes('CalendarEventId 重複'));
    assert.ok(sheet.getRange(3, rowIndex.columns.EVENT_ID).getNote().includes('第 2, 3 列'));
  } finally {
    context.applyRowDataFormats_ = originalApplyRowDataFormats;
    context.console.warn = originalConsoleWarn;
  }
}

function testClearOldEventsUsesFixedOpEventIdColumn() {
  const originalSpreadsheetApp = context.SpreadsheetApp;
  const originalRefreshHospitalMirrorSheetsSafely = context.refreshHospitalMirrorSheetsSafely_;
  const originalGetConfiguredCalendarId = context.getConfiguredCalendarId_;
  const originalIsCalendarAdvancedServiceAvailable = context.isCalendarAdvancedServiceAvailable_;
  const originalWithCalendarSyncLock = context.withCalendarSyncLock_;
  const originalDeleteCalendarEventByPossibleId = context.deleteCalendarEventByPossibleId_;
  const headers = ['病歷號', '姓名', 'CalendarEventId', 'CalendarSheetWriteUpdated', 'TEL'];
  const opSheet = makeDataSheet(headers, [
    ['001', '王小明', 'evt_deleted', '2026-05-09T00:00:01.000Z', '0911'],
    ['002', '陳小美', '', '2026-05-09T00:00:02.000Z', '0922'],
    ['003', '林大明', 'evt_missing', '2026-05-09T00:00:03.000Z', '0933'],
    ['004', '黃小美', 'evt_failed', '2026-05-09T00:00:04.000Z', '0944']
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
        getSheetByName: name => name === 'All' ? opSheet : null
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
    context.refreshHospitalMirrorSheetsSafely_ = () => ({ ok: true });

    context.clearCalendarEventsFromSpecifiedColumn();

    assert.deepStrictEqual(deleteCalls, ['evt_deleted', 'evt_missing', 'evt_failed']);
    assert.deepStrictEqual(opSheet.rows, [
      headers,
      ['001', '王小明', '', '', '0911'],
      ['002', '陳小美', '', '2026-05-09T00:00:02.000Z', '0922'],
      ['003', '林大明', '', '', '0933'],
      ['004', '黃小美', 'evt_failed', '2026-05-09T00:00:04.000Z', '0944']
    ]);
    assert.ok(alerts[0][1].includes('All'));
    assert.ok(alerts[0][1].includes('C 欄'));
    assert.strictEqual(alerts[1][0], '清除完成');
  } finally {
    context.SpreadsheetApp = originalSpreadsheetApp;
    context.getConfiguredCalendarId_ = originalGetConfiguredCalendarId;
    context.isCalendarAdvancedServiceAvailable_ = originalIsCalendarAdvancedServiceAvailable;
    context.withCalendarSyncLock_ = originalWithCalendarSyncLock;
    context.deleteCalendarEventByPossibleId_ = originalDeleteCalendarEventByPossibleId;
    context.refreshHospitalMirrorSheetsSafely_ = originalRefreshHospitalMirrorSheetsSafely;
  }
}

function testClearOldEventsSupportsLegacyEventIdColumn() {
  const originalSpreadsheetApp = context.SpreadsheetApp;
  const originalRefreshHospitalMirrorSheetsSafely = context.refreshHospitalMirrorSheetsSafely_;
  const originalGetConfiguredCalendarId = context.getConfiguredCalendarId_;
  const originalIsCalendarAdvancedServiceAvailable = context.isCalendarAdvancedServiceAvailable_;
  const originalWithCalendarSyncLock = context.withCalendarSyncLock_;
  const originalDeleteCalendarEventByPossibleId = context.deleteCalendarEventByPossibleId_;
  const headers = ['病歷號', '姓名', '日曆eventID', 'TEL'];
  const opSheet = makeDataSheet(headers, [
    ['001', '王小明', 'evt_deleted', '0911']
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
        getSheetByName: name => name === 'All' ? opSheet : null
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
      return { ok: true, status: 'deleted' };
    };
    context.refreshHospitalMirrorSheetsSafely_ = () => ({ ok: true });

    context.clearCalendarEventsFromSpecifiedColumn();

    assert.deepStrictEqual(deleteCalls, ['evt_deleted']);
    assert.strictEqual(opSheet.rows[1][2], '');
    assert.ok(alerts[0][1].includes('C 欄'));
  } finally {
    context.SpreadsheetApp = originalSpreadsheetApp;
    context.getConfiguredCalendarId_ = originalGetConfiguredCalendarId;
    context.isCalendarAdvancedServiceAvailable_ = originalIsCalendarAdvancedServiceAvailable;
    context.withCalendarSyncLock_ = originalWithCalendarSyncLock;
    context.deleteCalendarEventByPossibleId_ = originalDeleteCalendarEventByPossibleId;
    context.refreshHospitalMirrorSheetsSafely_ = originalRefreshHospitalMirrorSheetsSafely;
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
        getSheetByName: name => name === 'All'
          ? makeHeaderSheet(['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得'])
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
    assert.ok(alerts[0][0].includes('CalendarEventId'));
  } finally {
    context.SpreadsheetApp = originalSpreadsheetApp;
    context.getConfiguredCalendarId_ = originalGetConfiguredCalendarId;
    context.isCalendarAdvancedServiceAvailable_ = originalIsCalendarAdvancedServiceAvailable;
    context.withCalendarSyncLock_ = originalWithCalendarSyncLock;
  }
}

function testSurgeryExportDataForDate() {
  const headers = ['病歷號', '姓名', 'TEL', '醫院', 'Tag', 'Condition', '日期', '時間', 'Plan', '心得', 'CalendarEventId', 'CalendarSheetWriteUpdated'];
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
      '醫院': '聯醫',
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
    ['陳小美', 'A', '21.5', '-1.0'],
    ['王小明', 'B', '20.0', '-0.5']
  ]);
  assert.deepStrictEqual(toPlain(result.patientList), [
    ['001', '王小明', '0911', '08:30', 'Cataract', 'PE/IOL(B -0.5 20.0)', '術後注意\n帶藥'],
    ['002', '陳小美', '0922', '09:30', 'Cataract', 'MSICS(A -1.0 21.5)', 'APPLY']
  ]);
  assert.deepStrictEqual(toPlain(result.hospitalGroups.map(group => ({
    hospital: group.hospital,
    patientList: group.patientList,
    iolList: group.iolList
  }))), [
    {
      hospital: '高榮',
      patientList: [
        ['001', '王小明', '0911', '08:30', 'Cataract', 'PE/IOL(B -0.5 20.0)', '術後注意\n帶藥']
      ],
      iolList: [
        ['王小明', 'B', '20.0', '-0.5']
      ]
    },
    {
      hospital: '聯醫',
      patientList: [
        ['002', '陳小美', '0922', '09:30', 'Cataract', 'MSICS(A -1.0 21.5)', 'APPLY']
      ],
      iolList: [
        ['陳小美', 'A', '21.5', '-1.0']
      ]
    }
  ]);
}

function testSplitExportWritesSurgeryAndIolSheetsByHospital() {
  const surgerySheet = makeDataSheet([''], [], '手術清單');
  const iolSheet = makeDataSheet([''], [], '水晶體清單');
  const exportData = {
    dateLabel: '2026/05/09(Fri)',
    hospitalGroups: [
      {
        hospital: '高榮',
        patientList: [['001', '王小明', '0911', '08:30', 'Cataract', 'PE/IOL', '']],
        iolList: []
      },
      {
        hospital: '聯醫',
        patientList: [['002', '陳小美', '0922', '09:30', 'Cataract', 'MSICS', 'APPLY']],
        iolList: [['陳小美', 'A', '21.5', '-1.0']]
      }
    ]
  };

  context.writeSplitExportData_(surgerySheet, iolSheet, exportData, {
    surgery: 1,
    iol: 1
  });

  assert.strictEqual(surgerySheet.rows[0][0], '[ 2026/05/09(Fri) 高榮 手術清單 ]');
  assert.deepStrictEqual(surgerySheet.rows[1].slice(0, 7), ['病歷號', '姓名', 'TEL', '時間', '疾病', '術式', '補充說明']);
  assert.strictEqual(surgerySheet.rows[2][1], '王小明');
  assert.strictEqual(surgerySheet.rows[4][0], '[ 2026/05/09(Fri) 聯醫 手術清單 ]');
  assert.strictEqual(iolSheet.rows[0][0], '[ 2026/05/09(Fri) 高榮 水晶體清單 ]');
  assert.strictEqual(iolSheet.rows[1][0], '本日無水晶體資料');
  assert.strictEqual(iolSheet.rows[3][0], '[ 2026/05/09(Fri) 聯醫 水晶體清單 ]');
  assert.deepStrictEqual(iolSheet.rows[5].slice(0, 4), ['陳小美', 'A', '21.5', '-1.0']);
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
testLegacyEventIdColumnLookupFallback();
testEventIdColumnLookupMissingHeader();
testLegacyEventIdMigrationToEnglishHeader();
testHospitalColumnInsertedBeforeTagPreservesData();
testHospitalColumnMovedBeforeTagPreservesData();
testDuplicateHospitalHeadersAreReportedForManualReview();
testHospitalClassificationDefaultsToKaohsiungVeterans();
testMissingAllWithLegacyOpPromptsManualRename();
testTimeNoteRoundTrip();
testCalendarTitleKeepsBlankPatientNameSlot();
testSheetOriginatedCalendarEchoIsSkipped();
testSameTimestampDifferentCalendarContentIsNotSkipped();
testOlderAllDayEchoAfterTimedSheetUpdateIsSkipped();
testNewerCalendarAllDayUpdateClearsTime();
testMergeBlocksTrackedAbsorbedRows();
testMergeCombinesSafeTextFields();
testConditionalFormulaBuilding();
testConditionalFormattingKeepsTagIndependent();
testSystemColumnsAreHiddenAndProtected();
testHospitalMirrorSheetsFilterRowsFromAll();
testProcessCalendarChangeIsNoOpWhenReverseSyncDisabled();
testSyncCalendarChangesToSheetIsDisabled();
testDisableCalendarReverseSyncDeletesOnlyReverseTriggersAndTokens();
testDuplicateRowClearsSystemTrackingFields();
testDuplicateCalendarEventIdIsDiagnosedAndSkipped();
testClearOldEventsUsesFixedOpEventIdColumn();
testClearOldEventsSupportsLegacyEventIdColumn();
testClearOldEventsMissingEventIdHeaderAborts();
testSurgeryExportDataForDate();
testSplitExportWritesSurgeryAndIolSheetsByHospital();
testUpcomingExportDatesIncludeSevenDays();

console.log('Smoke tests passed');
