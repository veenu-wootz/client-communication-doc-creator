/**
 * glide.test.js — the write-back to the Glide row.
 *
 * The live API is never called here: these tests intercept fetch and assert the
 * request shape. Hitting the real endpoint would mutate a production RFQ row on
 * every test run.
 */

const test = require('node:test');
const assert = require('node:assert');

const { writeBackToGlide, isConfigured, ENDPOINT } = require('../src/deliver/glideWrite');

/** Run fn with a stubbed fetch, returning the request body it would have sent. */
async function capture(fn) {
  const real = globalThis.fetch;
  let sent = null;
  globalThis.fetch = async (url, init) => {
    sent = { url: String(url), headers: init.headers, body: JSON.parse(init.body) };
    return { ok: true, status: 200, text: async () => '{"ok":true}' };
  };
  try { await fn(); } finally { globalThis.fetch = real; }
  return sent;
}

/** Set env for one call and restore afterwards. */
async function withEnv(vars, fn) {
  const saved = {};
  for (const [k, v] of Object.entries(vars)) {
    saved[k] = process.env[k];
    if (v === undefined) delete process.env[k]; else process.env[k] = v;
  }
  try { return await fn(); } finally {
    for (const [k, v] of Object.entries(saved)) {
      if (v === undefined) delete process.env[k]; else process.env[k] = v;
    }
  }
}

const VALUES = {
  fileId: '01DDVW3ITEST',
  fileLink: 'https://example.sharepoint.com/x.pptx',
  generatedOn: '2026-09-05T15:04:13.408Z',
  generatedBy: 'vinay.jadon@wootz.work',
};

test('the mutation carries every column, keyed by Glide\'s internal ids', async () => {
  const sent = await withEnv({ GLIDE_TOKEN: 'test-token' }, () =>
    capture(() => writeBackToGlide({ rowID: undefined, rowId: 'ROW1', version: 'V1' }, VALUES)));

  assert.strictEqual(sent.url, ENDPOINT);
  assert.strictEqual(sent.headers.Authorization, 'Bearer test-token');

  const m = sent.body.mutations[0];
  assert.strictEqual(m.kind, 'set-columns-in-row');
  assert.strictEqual(m.rowID, 'ROW1');
  assert.deepStrictEqual(m.columnValues, {
    QZRyl: 'V1',
    RcHZF: VALUES.fileId,
    g0KAH: VALUES.fileLink,
    TjmbZ: VALUES.generatedOn,
    wVeBR: VALUES.generatedBy,
  });
});

test('a value we do not have is omitted, not written as empty', async () => {
  // Sending "" would blank a field the sender had filled in by hand.
  const sent = await withEnv({ GLIDE_TOKEN: 'test-token' }, () =>
    capture(() => writeBackToGlide({ rowId: 'ROW1', version: '' },
      { fileId: 'ID', fileLink: null, generatedOn: '2026-09-05T00:00:00Z', generatedBy: '' })));

  const keys = Object.keys(sent.body.mutations[0].columnValues);
  assert.deepStrictEqual(keys.sort(), ['RcHZF', 'TjmbZ'].sort(),
    'only the columns with real values are sent');
});

test('column ids, app and table can be overridden without touching code', async () => {
  const sent = await withEnv({
    GLIDE_TOKEN: 'test-token',
    GLIDE_APP_ID: 'OTHER_APP',
    GLIDE_TABLE: 'other-table',
    GLIDE_COL_FILE_ID: 'zzz9',
  }, () => capture(() => writeBackToGlide({ rowId: 'ROW1' }, { fileId: 'ID' })));

  assert.strictEqual(sent.body.appID, 'OTHER_APP');
  assert.strictEqual(sent.body.mutations[0].tableName, 'other-table');
  assert.deepStrictEqual(sent.body.mutations[0].columnValues, { zzz9: 'ID' });
});

test('a table named in the payload beats the env default', async () => {
  const sent = await withEnv({ GLIDE_TOKEN: 'test-token', GLIDE_TABLE: 'env-table' }, () =>
    capture(() => writeBackToGlide({ rowId: 'ROW1', table: 'payload-table' }, { fileId: 'ID' })));
  assert.strictEqual(sent.body.mutations[0].tableName, 'payload-table');
});

test('no token, no row, or nothing to write are skips — never errors', async () => {
  await withEnv({ GLIDE_TOKEN: undefined }, async () => {
    assert.strictEqual(isConfigured(), false);
    const r = await writeBackToGlide({ rowId: 'ROW1' }, VALUES);
    assert.strictEqual(r.skipped, true);
    assert.match(r.reason, /GLIDE_TOKEN/);
  });

  await withEnv({ GLIDE_TOKEN: 'test-token' }, async () => {
    const noRow = await writeBackToGlide({}, VALUES);
    assert.strictEqual(noRow.skipped, true);
    assert.match(noRow.reason, /rfq_row_id/);

    const noValues = await writeBackToGlide({ rowId: 'ROW1' }, {});
    assert.strictEqual(noValues.skipped, true);
  });
});

test('an API error is raised so the caller can log it, not swallowed', async () => {
  const real = globalThis.fetch;
  globalThis.fetch = async () => ({ ok: false, status: 401, text: async () => 'unauthorized' });
  try {
    await withEnv({ GLIDE_TOKEN: 'bad' }, async () => {
      await assert.rejects(
        () => writeBackToGlide({ rowId: 'ROW1' }, VALUES),
        /Glide API 401/,
      );
    });
  } finally { globalThis.fetch = real; }
});
