'use strict';
// CI-only metadata GET. Never print credentials, identities, editors or API bodies.
const fs = require('node:fs');
const version = fs.readFileSync('code.js', 'utf8').match(/VERSION:\s*'([^']+)'/)[1];

function summarize(book) {
  const named = new Map((book.namedRanges || []).map(value => [value.namedRangeId, value.range]));
  return (book.sheets || []).filter(sheet => sheet.properties.title === 'FU' || /^\d{6}$/.test(sheet.properties.title)).map(sheet => {
    const column = sheet.properties.title === 'FU' ? 9 : 22;
    const result = { sheet: sheet.properties.title, legacy_column: column === 9 ? 'J' : 'W', strict_count: 0,
      strict_editable_count: 0, strict_not_editable_count: 0, warning_count: 0, unresolved_count: 0 };
    for (const protection of sheet.protectedRanges || []) {
      const range = protection.range || named.get(protection.namedRangeId);
      if (!range) { result.unresolved_count += 1; continue; }
      if (range.sheetId !== sheet.properties.sheetId || (range.startColumnIndex || 0) > column ||
          (range.endColumnIndex !== undefined && range.endColumnIndex <= column)) continue;
      if (protection.warningOnly === true) { result.warning_count += 1; continue; }
      result.strict_count += 1;
      result[protection.requestingUserCanEdit === true ? 'strict_editable_count' : 'strict_not_editable_count'] += 1;
    }
    return result;
  });
}

function safeFailureReason(body) {
  const reasons = (body?.error?.details || []).map(detail => detail.reason);
  if (reasons.includes('ACCESS_TOKEN_SCOPE_INSUFFICIENT') ||
      (body?.error?.errors || []).some(error => error.reason === 'insufficientPermissions') ||
      /insufficient authentication scopes/i.test(body?.error?.message || '')) return 'insufficient_oauth_scope';
  if (reasons.includes('SERVICE_DISABLED')) return 'sheets_api_disabled';
  return 'sheets_access_denied_or_unavailable';
}

async function verify() {
  const report = values => console.log(JSON.stringify({ version, check: 'technical_protections', read_only: true, ...values }));
  try {
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
      if (!response.ok) return report({ verified: false, reason: 'oauth_refresh_unavailable', http_status: response.status });
      access = (await response.json()).access_token;
    }
    if (!access || !process.env.EXPECTED_SPREADSHEET_ID) return report({ verified: false, reason: 'required_ci_metadata_unavailable' });
    const fields = 'namedRanges(namedRangeId,range),sheets(properties(sheetId,title),protectedRanges(range,namedRangeId,warningOnly,requestingUserCanEdit))';
    const url = new URL(`https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(process.env.EXPECTED_SPREADSHEET_ID)}`);
    url.searchParams.set('fields', fields);
    const response = await fetch(url, { headers: { Authorization: `Bearer ${access}` }, signal: AbortSignal.timeout(15000) });
    if (!response.ok) {
      const body = await response.json().catch(() => ({}));
      return report({ verified: false, sheets_access: false, http_status: response.status, reason: safeFailureReason(body) });
    }
    report({ verified: true, sheets_access: true, technical_protections: summarize(await response.json()) });
  } catch { report({ verified: false, reason: 'protection_verification_unavailable' }); }
}
module.exports = { summarize, safeFailureReason };
if (require.main === module) void verify();
