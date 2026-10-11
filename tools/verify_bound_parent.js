'use strict';
// CI-only read-only check. Never print credentials, scriptId or API bodies.
const fs = require('node:fs');
const version = fs.readFileSync('code.js', 'utf8').match(/VERSION:\s*'([^']+)'/)[1];
const report = (parentMatch, reason = '') => console.log(JSON.stringify({
  version, parent_match: parentMatch, verified: parentMatch !== null, ...(reason ? { reason } : {})
}));
async function verify() {
  try {
    const clasp = JSON.parse(process.env.CLASP_JSON || '{}');
    const credentials = JSON.parse(process.env.CLASPRC_JSON || '{}');
    const oauth = credentials.tokens?.default || credentials.token || credentials;
    const settings = credentials.oauth2ClientSettings || {};
    const clientId = oauth.client_id || oauth.clientId || settings.clientId;
    const clientSecret = oauth.client_secret || oauth.clientSecret || settings.clientSecret;
    let access = oauth.access_token;
    if (oauth.refresh_token && clientId && clientSecret) {
      const response = await fetch('https://oauth2.googleapis.com/token', {
        method: 'POST', signal: AbortSignal.timeout(15000),
        body: new URLSearchParams({ grant_type: 'refresh_token', refresh_token: oauth.refresh_token,
          client_id: clientId, client_secret: clientSecret })
      });
      if (!response.ok) return report(null, `OAuth HTTP ${response.status}`);
      access = (await response.json()).access_token;
    }
    if (!access || !clasp.scriptId || !process.env.EXPECTED_SPREADSHEET_ID) {
      return report(null, 'Required CI credential metadata unavailable');
    }
    const response = await fetch(`https://script.googleapis.com/v1/projects/${encodeURIComponent(clasp.scriptId)}`, {
      headers: { Authorization: `Bearer ${access}` }, signal: AbortSignal.timeout(15000)
    });
    if (!response.ok) return report(null, `Apps Script HTTP ${response.status}`);
    const project = await response.json();
    report(project.parentId === process.env.EXPECTED_SPREADSHEET_ID);
  } catch { report(null, 'Parent verification unavailable'); }
}
void verify();
