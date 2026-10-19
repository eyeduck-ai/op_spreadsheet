'use strict';
const assert = require('node:assert/strict');
const { sourceVersion, main } = require('../tools/verify_bound_parent');

async function tests() {
  assert.strictEqual(sourceVersion('const CONFIG = { "VERSION": "2026.10.19.2" };'), '2026.10.19.2');
  assert.strictEqual(sourceVersion("const CONFIG = { VERSION: '2026.10.15' };"), '2026.10.15');
  assert.throws(() => sourceVersion('const CONFIG = {};'), /version unavailable/);
  const env = {
    CLASP_JSON: JSON.stringify({ scriptId: 'synthetic-script' }),
    CLASPRC_JSON: JSON.stringify({ access_token: 'synthetic-access' }),
    EXPECTED_SPREADSHEET_ID: 'synthetic-sheet'
  };
  for (const [parent, status, expected] of [['synthetic-sheet', 200, 0], ['other-sheet', 200, 1], ['', 403, 1]]) {
    const output = [];
    const code = await main(env, async url => {
      assert.ok(url.startsWith('https://script.googleapis.com/v1/projects/'));
      return { ok: status === 200, status, json: async () => ({ parentId: parent }) };
    }, value => output.push(value));
    assert.strictEqual(code, expected);
    const report = JSON.parse(output[0]);
    assert.strictEqual(report.verified, expected === 0);
    assert.ok(!/synthetic-(?:access|script|sheet)|other-sheet/.test(output[0]));
  }
  const output = [];
  assert.strictEqual(await main({}, async () => { throw new Error('Must not call API'); }, value => output.push(value)), 1);
  assert.strictEqual(JSON.parse(output[0]).verified, false);
  console.log('Deployment verification passed (generated CONFIG, exact binding, HTTP errors and missing credentials fail closed)');
}
tests().catch(error => { console.error(error); process.exitCode = 1; });
