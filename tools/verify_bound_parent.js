'use strict';
// CI-only read-only check. Never print credentials, scriptId or API bodies.
const fs = require('node:fs');
const path = require('node:path');
function sourceVersion(source) {
  const match = source.match(/["']?VERSION["']?\s*:\s*["']([^"']+)["']/);
  if (!match) throw new Error('Source version unavailable');
  return match[1];
}
async function verify(env = process.env, request = fetch) {
  const version = sourceVersion(fs.readFileSync(path.join(__dirname, '..', 'code.js'), 'utf8'));
  const report = (parentMatch, reason = '') => ({
    version, parent_match: parentMatch, verified: parentMatch === true, ...(reason ? { reason } : {})
  });
  try {
    const clasp = JSON.parse(env.CLASP_JSON || '{}');
    const credentials = JSON.parse(env.CLASPRC_JSON || '{}');
    const oauth = credentials.tokens?.default || credentials.token || credentials;
    const settings = credentials.oauth2ClientSettings || {};
    const clientId = oauth.client_id || oauth.clientId || settings.clientId;
    const clientSecret = oauth.client_secret || oauth.clientSecret || settings.clientSecret;
    let access = oauth.access_token;
    if (oauth.refresh_token && clientId && clientSecret) {
      const response = await request('https://oauth2.googleapis.com/token', {
        method: 'POST', signal: AbortSignal.timeout(15000),
        body: new URLSearchParams({ grant_type: 'refresh_token', refresh_token: oauth.refresh_token,
          client_id: clientId, client_secret: clientSecret })
      });
      if (!response.ok) return report(null, `OAuth HTTP ${response.status}`);
      access = (await response.json()).access_token;
    }
    if (!access || !clasp.scriptId || !env.EXPECTED_SPREADSHEET_ID) {
      return report(null, 'Required CI credential metadata unavailable');
    }
    const response = await request(`https://script.googleapis.com/v1/projects/${encodeURIComponent(clasp.scriptId)}`, {
      headers: { Authorization: `Bearer ${access}` }, signal: AbortSignal.timeout(15000)
    });
    if (!response.ok) return report(null, `Apps Script HTTP ${response.status}`);
    const project = await response.json();
    return report(project.parentId === env.EXPECTED_SPREADSHEET_ID);
  } catch { return report(null, 'Parent verification unavailable'); }
}
async function main(env = process.env, request = fetch, write = console.log) {
  try {
    const result = await verify(env, request);
    write(JSON.stringify(result));
    return result.verified ? 0 : 1;
  } catch {
    write(JSON.stringify({ verified: false, parent_match: null, reason: 'Source verification unavailable' }));
    return 1;
  }
}
module.exports = { sourceVersion, verify, main };
if (require.main === module) main().then(code => { process.exitCode = code; });
