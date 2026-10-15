'use strict';
const assert = require('node:assert/strict'), fs = require('node:fs');
const { call, evaluate, plain, context } = require('./smoke');
const { execute, clone } = require('./native_book');

function fixture(month, id, procedure, side, model) {
  const labels = [...plain(evaluate('CONFIG.MONTHLY_HEADERS')), '人工備註'];
  const columns = Object.fromEntries(labels.map((label, index) => [label, index]));
  const literal = value => ({ userEnteredValue: { stringValue: value } });
  const day = Array.from({ length: labels.length }, () => ({}));
  day[columns['日期／報到時間']] = literal(`${month.slice(0,4)}-${month.slice(4)}-15`);
  day[columns['醫院']] = literal('高榮'); day[columns['病歷號']] = literal('◆ 刀日');
  day[columns.ScheduleSyncId] = literal((id * 100 + 1).toString(16).padStart(32, '0'));
  const patient = Array.from({ length: labels.length }, () => ({}));
  for (const [key, value] of Object.entries({ '病歷號': `TEST${id}`, '姓名': '合成病人', '側別': side, '術式': procedure, IOL: model, Plan: '保留 +20.0',
    ScheduleSyncId: (id * 100 + 2).toString(16).padStart(32, '0'), CalendarEventId: `synthetic-event-${id}` })) patient[columns[key]] = literal(value);
  patient[columns['人工備註']] = { ...literal('人工富文字'), note: '完整保留註記', textFormatRuns: [{ startIndex: 0, format: { italic: true } }] };
  const cancelled = clone(patient); cancelled[columns['病歷號']] = { ...literal(`CANCEL${id}`), effectiveFormat: { textFormat: { strikethrough: true } } };
  cancelled[columns.ScheduleSyncId] = literal((id * 100 + 3).toString(16).padStart(32, '0'));
  cancelled[columns.CalendarEventId] = literal(`synthetic-cancel-${id}`);
  return { properties: { sheetId: id, title: month, gridProperties: { rowCount: 100, columnCount: labels.length, frozenRowCount: 1 } },
    rows: [labels.map(literal), day, patient, cancelled],
    column_widths: labels.map(() => ({ pixelSize: 110 })), row_metadata: [{ pixelSize: 35 }, {}, { pixelSize: 57, hiddenByUser: true }],
    merges: [{ sheetId: id, startRowIndex: 2, endRowIndex: 3, startColumnIndex: columns.Plan, endColumnIndex: columns.Plan + 2 }],
    conditionalFormats: [{ ranges: [{ sheetId: id, startRowIndex: 1 }], booleanRule: { condition: { type: 'CUSTOM_FORMULA', values: [{ userEnteredValue: '=FALSE' }] }, format: { textFormat: { strikethrough: true } } } }] };
}
function archiveContract() {
  let random = 1; context.Math = Object.create(Math); context.Math.random = () => (random++ % 1000) / 1000;
  let book = { properties: { timeZone: 'Asia/Taipei' }, developerMetadata: [], sheets: [fixture('202609', 101, 'LenSx-Phaco-IOL', 'OU', 'SN60WF'), fixture('202606', 102, 'VT+MP', 'OD', '')] };
  const before = clone(book), first = plain(call('planNativeArchive_', book, ['202609'], '202610', 'GAS-SEPTEMBER'));
  book = execute(book, first.requests); assert.equal(call('verifyNativeArchive_', first, book, false), true);
  book = execute(book, [{ deleteSheet: { sheetId: 101 } }]); assert.equal(call('verifyNativeArchive_', first, book, true), true);
  const afterSeptember = clone(book), annual = book.sheets.find(sheet => sheet.properties.title === '刀表封存_2026'), annualId = annual.properties.sheetId;
  const second = plain(call('planNativeArchive_', book, ['202606'], '202610', 'GAS-JUNE'));
  assert.equal(second.summary.months[0].backfill, true);
  book = execute(book, second.requests); assert.equal(call('verifyNativeArchive_', second, book, false), true);
  const during = clone(book); book = execute(book, [{ deleteSheet: { sheetId: 102 } }]);
  assert.equal(call('verifyNativeArchive_', second, book, true), true);
  assert.equal(book.sheets[0].properties.sheetId, annualId);
  const info = plain(call('nativeArchiveStructure_', book.sheets[0]));
  assert.deepEqual(info.months, ['202606', '202609']);
  const cancelled = book.sheets[0].rows.find(row => (row[info.cols['病歷號']]?.userEnteredValue?.stringValue || '').startsWith('CANCEL101'));
  assert.equal(cancelled[info.cols['病歷號']].userEnteredFormat.textFormat.strikethrough, true);
  assert.ok(book.sheets[0].rows.some(row => row[info.cols['人工備註']]?.textFormatRuns?.length));
  assert.equal(book.sheets[0].conditionalFormats?.length || 0, 0);
  const lock = plain(call('nativeLockRequest_', 'ONE', false));
  assert.throws(() => execute(execute(before, [lock]), [plain(call('nativeLockRequest_', 'TWO', false))]), /duplicate metadata/);
  assert.throws(() => call('planNativeArchive_', before, ['202610'], '202610', 'BAD'), /已結束/);
  const crossYear = clone(before); crossYear.sheets.push(fixture('202512', 103, 'MSICS-IOL', 'OS', 'DIB00'));
  const yearly = plain(call('planNativeArchive_', crossYear, ['202512','202606'], '202610', 'YEARS'));
  assert.deepEqual(Object.keys(yearly.targets).sort(), ['刀表封存_2025','刀表封存_2026']);
  assert.equal(call('verifyNativeArchive_', yearly, execute(crossYear, yearly.requests), false), true);
  const altered = clone(during); altered.sheets.find(sheet => sheet.properties.title === '202606').rows[2][3].userEnteredValue.stringValue = '合成異動';
  assert.equal(call('verifyNativeArchive_', second, altered, false), false);
  return { schema: 2, lock: { id: evaluate('NATIVE_ARCHIVE_LOCK_ID'), key: evaluate('NATIVE_ARCHIVE_LOCK_KEY') }, before,
    after_september: afterSeptember, during_backfill: during, after_backfill: book };
}
function verifyWindowsInput(filename) {
  const input = JSON.parse(fs.readFileSync(filename, 'utf8'));
  let book = execute(input.book, input.requests);
  book = execute(book, input.source_ids.map(sheetId => ({ deleteSheet: { sheetId } })));
  book.sheets.filter(sheet => /^刀表封存_/.test(sheet.properties.title)).forEach(sheet => call('nativeArchiveStructure_', sheet));
  const month = '202605'; book.sheets.push(fixture(month, 501, 'MSICS-IOL', 'OS', 'DIB00'));
  const plan = plain(call('planNativeArchive_', book, [month], '202610', 'WINDOWS-TO-GAS'));
  assert.equal(call('verifyNativeArchive_', plan, execute(book, plan.requests), false), true);
}
function runtimeChecks() {
  const saved = { read: context.readNativeArchiveBook_, drive: context.DriveApp, active: context.SpreadsheetApp.getActiveSpreadsheet,
    blob: context.Utilities.newBlob, gzip: context.Utilities.gzip, ungzip: context.Utilities.ungzip,
    api: context.Sheets, lock: context.withSheetOperationLock_ };
  const files = new Map(); let nextFile = 1, copyWrites = 0, deleteWrites = 0, loseCopy = true, loseDelete = true;
  let book = { properties: { timeZone: 'Asia/Taipei' }, developerMetadata: [], sheets: [fixture('202609', 201, 'LenSx-Phaco-IOL', 'OU', 'SN60WF')] };
  const file = (name, content, snapshot) => {
    const id = `private-${nextFile++}`;
    const result = { getId: () => id, getName: () => name, getBlob: () => ({ getDataAsString: () => content }),
      getSharingAccess: () => 'PRIVATE', getEditors: () => [], getViewers: () => [], setSharing() {}, setShareableByEditors() {},
      makeCopy: copyName => file(copyName, '', clone(book)), snapshot };
    files.set(id, result); return result;
  };
  const sourceFile = file('source', ''); files.set('GAS-TEST', sourceFile);
  context.SpreadsheetApp.getActiveSpreadsheet = () => ({ getId: () => 'GAS-TEST' });
  context.Utilities.newBlob = (content, mime, name) => ({ getDataAsString: () => content, getName: () => name });
  context.Utilities.gzip = blob => ({ getDataAsString: blob.getDataAsString, getName: () => blob.getName() + '.gz' });
  context.Utilities.ungzip = blob => blob;
  context.DriveApp = { Access: { PRIVATE: 'PRIVATE' }, Permission: { NONE: 'NONE' }, createFile: blob => file(blob.getName(), blob.getDataAsString()), getFileById: id => files.get(id) };
  context.readNativeArchiveBook_ = id => clone(id === 'GAS-TEST' ? book : files.get(id).snapshot);
  context.Sheets = { Spreadsheets: { get: () => ({ developerMetadata: clone(book.developerMetadata) }), batchUpdate: ({ requests }) => {
    book = execute(book, plain(requests));
    if (requests.some(request => request.updateCells)) { copyWrites++; if (loseCopy) { loseCopy = false; throw new Error('synthetic copy response lost'); } }
    if (requests.some(request => request.deleteSheet)) { deleteWrites++; if (loseDelete) { loseDelete = false; throw new Error('synthetic delete response lost'); } }
    return {};
  } } };
  context.withSheetOperationLock_ = callback => callback();
  call('writeAnnualArchiveTransaction_', null);
  try {
    const preview = plain(call('submitMonthlyRollupRequest', { months: ['202609'], confirmedPermanentDeletion: true }));
    assert.equal(preview.status, 'preview'); assert.equal(copyWrites, 0); assert.equal(deleteWrites, 0);
    assert.equal(book.developerMetadata.length, 0);
    const copy = plain(call('submitMonthlyRollupRequest', { action: 'copy', operationId: preview.operationId }));
    assert.equal(copy.status, 'uncertain'); assert.equal(copyWrites, 1); assert.equal(book.sheets.length, 2);
    assert.throws(() => call('assertNativeArchiveAvailable_', '', book), /未完成封存/);
    const checked = plain(call('submitMonthlyRollupRequest', { action: 'check', operationId: preview.operationId }));
    assert.equal(checked.status, 'copied'); assert.equal(copyWrites, 1); assert.equal(deleteWrites, 0);
    plain(call('submitMonthlyRollupRequest', { action: 'finalize', operationId: preview.operationId, confirmedPermanentDeletion: false }));
    assert.equal(deleteWrites, 0);
    // Check recovers the copied state; a rejected confirmation never deletes.
    call('submitMonthlyRollupRequest', { action: 'check', operationId: preview.operationId });
    const removed = plain(call('submitMonthlyRollupRequest', { action: 'finalize', operationId: preview.operationId, confirmedPermanentDeletion: true }));
    assert.equal(removed.status, 'uncertain'); assert.equal(deleteWrites, 1); assert.equal(book.sheets.length, 1);
    const final = plain(call('submitMonthlyRollupRequest', { action: 'check', operationId: preview.operationId }));
    assert.equal(final.status, 'applied'); assert.equal(deleteWrites, 1); assert.equal(copyWrites, 1);
    assert.equal(book.developerMetadata.length, 0);
    book.sheets.push(fixture('202606', 203, 'VT+MP', 'OD', ''));
    const unsubmitted = plain(call('submitMonthlyRollupRequest', { action: 'preview', months: ['202606'] }));
    const record = plain(call('readAnnualArchiveTransaction_')); record.status = 'locked'; call('writeAnnualArchiveTransaction_', record);
    book = execute(book, [plain(call('nativeLockRequest_', unsubmitted.operationId, false))]);
    book.sheets.find(sheet => sheet.properties.title === '202606').rows[2][3].userEnteredValue.stringValue = '合成來源改動';
    const cancelled = plain(call('submitMonthlyRollupRequest', { action: 'cancel', operationId: unsubmitted.operationId }));
    assert.equal(cancelled.status, 'rejected'); assert.equal(book.developerMetadata.length, 0);
    assert.ok(book.sheets.some(sheet => sheet.properties.title === '202606'));
    assert.equal(copyWrites, 1); assert.equal(deleteWrites, 1);
  } finally {
    call('writeAnnualArchiveTransaction_', null);
    context.readNativeArchiveBook_ = saved.read; context.DriveApp = saved.drive; context.Sheets = saved.api;
    context.Utilities.newBlob = saved.blob; context.Utilities.gzip = saved.gzip; context.Utilities.ungzip = saved.ungzip;
    context.SpreadsheetApp.getActiveSpreadsheet = saved.active; context.withSheetOperationLock_ = saved.lock;
  }
}
module.exports = { archiveContract, fixture };
if (require.main === module) {
  if (process.argv[2] === '--verify-input') { verifyWindowsInput(process.argv[3]); console.log('Windows to GAS archive checks passed'); }
  else { const result = archiveContract(); runtimeChecks(); if (process.argv.includes('--json')) process.stdout.write(JSON.stringify(result)); else console.log('Native archive contract checks passed (native cells, backfill, lock, two phases, restart receipts)'); }
}
