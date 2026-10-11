'use strict';
const assert = require('node:assert/strict');
const { summarize, safeFailureReason } = require('../tools/verify_protections');
const result = summarize({ namedRanges: [{ namedRangeId: 'private', range: { sheetId: 1, startColumnIndex: 9, endColumnIndex: 10 } }],
  sheets: [{ properties: { sheetId: 1, title: 'FU' }, protectedRanges: [
    { namedRangeId: 'private', requestingUserCanEdit: true, editors: { users: ['secret@example.invalid'] } },
    { range: { sheetId: 1, startColumnIndex: 9, endColumnIndex: 10 } },
    { range: { sheetId: 1, startColumnIndex: 9, endColumnIndex: 10 }, warningOnly: true },
    { range: { sheetId: 1, startColumnIndex: 8, endColumnIndex: 9 } }
  ] }, { properties: { sheetId: 2, title: '202604' }, protectedRanges: [{ range: { sheetId: 2 }, requestingUserCanEdit: true }] },
  { properties: { sheetId: 3, title: 'Private' } }] });
assert.deepEqual(result, [
  { sheet: 'FU', legacy_column: 'J', strict_count: 2, strict_editable_count: 1, strict_not_editable_count: 1, warning_count: 1, unresolved_count: 0 },
  { sheet: '202604', legacy_column: 'W', strict_count: 1, strict_editable_count: 1, strict_not_editable_count: 0, warning_count: 0, unresolved_count: 0 }
]);
assert.equal(JSON.stringify(result).includes('secret'), false);
assert.equal(JSON.stringify(result).includes('private'), false);
assert.equal(safeFailureReason({ error: { details: [{ reason: 'ACCESS_TOKEN_SCOPE_INSUFFICIENT' }] } }), 'insufficient_oauth_scope');
assert.equal(safeFailureReason({ error: { details: [{ reason: 'SERVICE_DISABLED' }] } }), 'sheets_api_disabled');
console.log('Read-only protection metadata summary passed');
