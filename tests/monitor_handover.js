'use strict';
// Execute the real GAS row classifiers; no Google services or patient data.
const assert = require('node:assert/strict');
const fixture = require('./fixtures/monitor_handover.json');
const { call, evaluate, plain, FakeSheet, makeFakeSpreadsheet } = require('./smoke');
const config = plain(evaluate('CONFIG'));
const headers = ['ScheduleSyncId', '人工備註', ...config.MONTHLY_HEADERS.filter(value =>
  value !== 'ScheduleSyncId' && value !== 'CalendarEventId'), 'CalendarEventId'];
const row = values => headers.map(header => values[header] || '');
const day = (date, hospital, uid) => row({'日期／報到時間': date, '醫院': hospital,
  '病歷號': '◆ 刀日', ScheduleSyncId: uid});
const source = new FakeSheet('202610', 1, [headers,
  day(new Date(2026, 9, 12), '高榮', 'a'.repeat(32)),
  row({'病歷號': 'TEST001', '姓名': '合成甲', '醫院': '聯醫', '日期／報到時間': '0830',
    ScheduleSyncId: fixture.patient_sync_id, CalendarEventId: fixture.patient_event_id}),
  day(new Date(2026, 9, 12), '聯醫', 'c'.repeat(32)),
  row({'病歷號': 'TEST002', '姓名': '合成乙', ScheduleSyncId: 'd'.repeat(32)})]);
const sourceColumns = call('getRequiredMonthlyColumns_', source);
const blocks = call('scanMonthlyBlocks_', source, sourceColumns);
assert.deepStrictEqual(plain(blocks.blocks.map(block => block.blockKey)), fixture.same_day_blocks);
assert.deepStrictEqual(plain(blocks.blocks.map(block => block.patientRows.length)), [1, 1]);
assert.strictEqual(blocks.blocks[0].hospital, '高榮'); // Patient cell may still say 聯醫.
const before = call('scanScheduleIdentities_', makeFakeSpreadsheet([source]));
assert.strictEqual(before.conflicts.length, 0);
assert.strictEqual(before.assigned, 4);
assert.strictEqual(before.sheets[0].rows[0].type, 'DATE_HEADER');
assert.strictEqual(source.rows[1][sourceColumns.EVENT_ID - 1], '');
assert.strictEqual(fixture.day_headers_create_events, false);

// Whole-row cut/paste retains the two hidden IDs under the new header.
const moved = source.rows.splice(2, 1)[0];
const target = new FakeSheet('202611', 2, [headers,
  day(new Date(2026, 10, 2), fixture.moved_hospital, 'e'.repeat(32)), moved]);
const targetColumns = call('getRequiredMonthlyColumns_', target);
const targetScan = call('scanMonthlyBlocks_', target, targetColumns);
assert.strictEqual(targetScan.blocks[0].dateKey, fixture.moved_day);
assert.strictEqual(targetScan.blocks[0].hospital, fixture.moved_hospital);
assert.strictEqual(target.rows[2][targetColumns.SYNC_ID - 1], fixture.patient_sync_id);
assert.strictEqual(target.rows[2][targetColumns.EVENT_ID - 1], fixture.patient_event_id);
assert.strictEqual(call('scanScheduleIdentities_', makeFakeSpreadsheet([source, target])).conflicts.length, 0);
call('rebuildMonthlySheetSorted_', target);
assert.strictEqual(target.rows[2][targetColumns.SYNC_ID - 1], fixture.patient_sync_id);
assert.strictEqual(target.rows[2][targetColumns.EVENT_ID - 1], fixture.patient_event_id);

// Visible-cell-only cut/paste leaves an orphan ID and an unidentified patient.
const visible = target.rows[2].slice();
visible[targetColumns.SYNC_ID - 1] = '';
visible[targetColumns.EVENT_ID - 1] = '';
target.rows[2] = row({ScheduleSyncId: fixture.patient_sync_id, CalendarEventId: fixture.patient_event_id});
target.rows.push(visible);
const broken = call('scanScheduleIdentities_', makeFakeSpreadsheet([source, target]));
assert.ok(broken.conflicts.some(message => message.includes('只剩識別碼')));
assert.strictEqual(broken.missing, 1);
// Private main-EXE baselines cannot be sorted or interpreted as a schedule.
const state = new FakeSheet(fixture.main_state_sheet, 3, [['schemaVersion', 'ScheduleSyncId'], [1, fixture.patient_sync_id]]);
assert.strictEqual(call('getScheduleIdentitySheetInfo_', state), null);
assert.strictEqual(call('isMonthlySheetName_', fixture.main_state_sheet), false);
console.log('Monitor handover contract passed (hospital/date, whole-row move, hidden IDs, visible-only conflicts, private state)');
