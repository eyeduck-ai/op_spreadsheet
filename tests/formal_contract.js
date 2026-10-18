'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { evaluate, plain } = require('./smoke');
const root = path.resolve(__dirname, '..');
const spec = JSON.parse(fs.readFileSync(path.join(root, 'contracts/sheet-contract.json'), 'utf8'));
const actual = plain(evaluate('CONFIG'));
delete actual.VERSION;
assert.deepStrictEqual(actual, spec.gas_config);
assert.deepStrictEqual(plain(evaluate('LEGACY_EVENT_ID_HEADERS')), spec.event_aliases);
assert.deepStrictEqual(plain(evaluate('LEGACY_MONTHLY_TIME_HEADERS')), spec.date_aliases);
assert.strictEqual(evaluate('MONTHLY_DATE_HEADER_MARKER'), spec.day_marker);
assert.strictEqual(evaluate('MONTHLY_DIAGNOSIS_SUMMARY_FORMULA_MARKER'), spec.diagnosis_formula_marker);
for (const [script, fn, file] of [
  ['sheet_contract', 'sheetContract', 'sheet_contract.json'],
  ['parser_contract', 'parserContract', 'schema_parsing.json'],
  ['workflow_contract', 'workflowContract', 'workflow_contract.json'],
  ['archive_contract', 'archiveContract', 'gas_archives.json'],
]) {
  assert.deepStrictEqual(require('./' + script)[fn](),
    JSON.parse(fs.readFileSync(path.join(root, 'contracts/cases', file), 'utf8')));
}
console.log('Formal contract 1.0.0 passed: GAS fields, aliases, parsing, workflows and native archive exports');
