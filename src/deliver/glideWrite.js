/**
 * glideWrite.js — write the generated deck's details back into the Glide row.
 *
 * OPTIONAL, like every delivery step: a missing token is a logged skip, never an
 * error. The email attachment stays the guaranteed path.
 *
 * Uses Glide's mutateTables function with a set-columns-in-row mutation. Column
 * ids are Glide's internal names (`QZRyl` and friends) — opaque but stable, and
 * overridable by env so a different table needs no code change.
 *
 * The token is read from the environment and never hardcoded: one of this
 * repo's two remotes is public.
 */

const ENDPOINT = 'https://api.glideapp.io/api/function/mutateTables';

const APP_ID = () => process.env.GLIDE_APP_ID || 'ARzoymvBNIgO6RcvRk7l';
const TABLE  = (override) => override
  || process.env.GLIDE_TABLE
  || 'native-table-24696dcc-caaf-4bf8-a015-1e9ef394aa1b';

/**
 * canonical name → Glide's internal column id.
 *
 * `YU0gy` is "RFQ Version Used". Do NOT use `QZRyl` here: that is the app's own
 * "RFQ Folder Drive Item ID" column, written by Strike, and this service has no
 * business touching it.
 */
const COLUMNS = () => ({
  version:     process.env.GLIDE_COL_VERSION      || 'YU0gy',
  fileId:      process.env.GLIDE_COL_FILE_ID      || 'RcHZF',
  fileLink:    process.env.GLIDE_COL_FILE_LINK    || 'g0KAH',
  generatedOn: process.env.GLIDE_COL_GENERATED_ON || 'TjmbZ',
  generatedBy: process.env.GLIDE_COL_GENERATED_BY || 'wVeBR',
});

const isConfigured = () => Boolean(process.env.GLIDE_TOKEN);

/**
 * @param {object} writeback  { rowId, version, table } from the payload
 * @param {object} values     { fileId, fileLink, generatedOn, generatedBy }
 * @returns {Promise<{skipped: boolean, reason?: string, written?: string[]}>}
 */
async function writeBackToGlide(writeback = {}, values = {}) {
  if (!isConfigured()) {
    const reason = 'GLIDE_TOKEN not set';
    console.log(`  glide: skipped — ${reason}`);
    return { skipped: true, reason };
  }
  if (!writeback.rowId) {
    console.log('  glide: skipped — payload carried no rfq_row_id');
    return { skipped: true, reason: 'no rfq_row_id' };
  }

  const col = COLUMNS();

  // Every column is written every time, blank included. These five are written
  // only by this service, never by hand, so a missing value must clear the cell
  // rather than leave the previous run's answer sitting there — a stale version
  // beside a fresh file is worse than an empty one.
  const blank = (v) => (v === undefined || v === null ? '' : String(v));
  const columnValues = {
    [col.version]: blank(writeback.version),
    [col.fileId]: blank(values.fileId),
    [col.fileLink]: blank(values.fileLink),
    [col.generatedOn]: blank(values.generatedOn),
    [col.generatedBy]: blank(values.generatedBy),
  };

  const res = await fetch(ENDPOINT, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${process.env.GLIDE_TOKEN}`,
    },
    body: JSON.stringify({
      appID: APP_ID(),
      mutations: [{
        kind: 'set-columns-in-row',
        tableName: TABLE(writeback.table),
        columnValues,
        rowID: writeback.rowId,
      }],
    }),
  });

  const text = await res.text();
  if (!res.ok) throw new Error(`Glide API ${res.status}: ${text.slice(0, 300)}`);

  console.log(`  glide → row ${writeback.rowId}: ${Object.keys(columnValues).length} column(s) written`);
  return { skipped: false, written: Object.keys(columnValues) };
}

module.exports = { writeBackToGlide, isConfigured, COLUMNS, ENDPOINT };
