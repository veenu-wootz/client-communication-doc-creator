/**
 * glideWrite.js — write the stored file URL back into the Glide row.
 *
 * OPTIONAL, AND LAST IN THE CHAIN (PLAN.md T7). It only runs when an upload
 * actually produced a URL — there is nothing to write otherwise — and when the
 * Glide credentials are present. Missing either is a skip, never an error.
 *
 * ── NOT YET WIRED ────────────────────────────────────────────────────────────
 * Glide has two API surfaces and we do not yet know which the Strike app uses:
 *
 *   'v2'      PATCH https://api.glideapps.com/tables/{table}/rows/{rowId}
 *             Authorization: Bearer <token>            (Big Tables)
 *   'mutate'  POST  https://api.glideapp.io/api/function/mutateTables
 *             { appID, mutations: [{ kind: 'set-columns-in-row', ... }] }
 *
 * Both are implemented below and chosen by GLIDE_API. When you send the key,
 * table/row identifiers and target column, set the env vars and it starts
 * working — there is no other code to change.
 *
 * Structure follows qualityinspectionreport/appsheetRows.js: env-keyed, POST,
 * throw on a non-2xx so the caller can log it as non-fatal.
 */

const fetch = require('node-fetch');

const API = () => (process.env.GLIDE_API || 'v2').toLowerCase();

function isConfigured(writeback = {}) {
  if (!process.env.GLIDE_TOKEN) return false;
  const table = writeback.table || process.env.GLIDE_TABLE;
  const column = writeback.column || process.env.GLIDE_COLUMN;
  if (!table || !column) return false;
  if (API() === 'mutate' && !process.env.GLIDE_APP_ID) return false;
  return true;
}

function describeMissing(writeback = {}) {
  const missing = [];
  if (!process.env.GLIDE_TOKEN) missing.push('GLIDE_TOKEN');
  if (!(writeback.table || process.env.GLIDE_TABLE)) missing.push('table');
  if (!(writeback.column || process.env.GLIDE_COLUMN)) missing.push('column');
  if (API() === 'mutate' && !process.env.GLIDE_APP_ID) missing.push('GLIDE_APP_ID');
  return missing;
}

/**
 * @param {object} writeback  { rowId, table, column } from the payload
 * @param {string} fileUrl    the stored deck URL
 */
async function writeDeckUrl(writeback = {}, fileUrl) {
  if (!fileUrl) {
    console.log('  glide: skipped — no file URL to write (upload did not run)');
    return { skipped: true, reason: 'no file url' };
  }
  if (!writeback.rowId) {
    console.log('  glide: skipped — payload carried no row id');
    return { skipped: true, reason: 'no row id' };
  }
  if (!isConfigured(writeback)) {
    const reason = `glide not configured (${describeMissing(writeback).join(', ')})`;
    console.log(`  glide: skipped — ${reason}`);
    return { skipped: true, reason };
  }

  const table = writeback.table || process.env.GLIDE_TABLE;
  const column = writeback.column || process.env.GLIDE_COLUMN;
  const token = process.env.GLIDE_TOKEN;

  let url, options;

  if (API() === 'mutate') {
    url = 'https://api.glideapp.io/api/function/mutateTables';
    options = {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({
        appID: process.env.GLIDE_APP_ID,
        mutations: [{
          kind: 'set-columns-in-row',
          tableName: table,
          rowID: writeback.rowId,
          columnValues: { [column]: fileUrl },
        }],
      }),
    };
  } else {
    url = `https://api.glideapps.com/tables/${encodeURIComponent(table)}/rows/${encodeURIComponent(writeback.rowId)}`;
    options = {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${token}` },
      body: JSON.stringify({ [column]: fileUrl }),
    };
  }

  const res = await fetch(url, { ...options, timeout: 15000 });
  if (!res.ok) throw new Error(`Glide API ${res.status}: ${(await res.text()).slice(0, 300)}`);

  console.log(`  glide → row ${writeback.rowId}.${column} updated`);
  return { skipped: false };
}

module.exports = { writeDeckUrl, isConfigured, describeMissing };
