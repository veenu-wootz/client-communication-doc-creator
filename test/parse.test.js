/**
 * parse.test.js — the Strike/Glide payload parser.
 *
 * These are written against shapes seen in a real Glide export, not against
 * what a clean webhook would ideally send. The first test is the one that
 * matters most: before it existed, this payload produced a cover-only deck
 * with every query silently missing.
 */

const test = require('node:test');
const assert = require('node:assert');

const {
  parseStrikePayload, cleanRichText, extractJsonObjects, undoubleQuotes,
} = require('../src/input/parseStrikePayload');

/** Glide joins repeating rows into one string inside a one-element array. */
const glideItems = (rows) => [rows.map((r) => JSON.stringify(r)).join(', ')];

test('rows joined into a single string are all recovered', () => {
  const body = {
    report_title: 'Sample Report',
    reference_name: 'PPT Creator',
    items: glideItems([
      { item_name: 'HEX NUT', description: 'First row' },
      { description: 'Second row', images: ['https://x/a.jpg', 'https://x/b.jpg'] },
      { description: 'Third row', images: [] },
      { description: 'Fourth row', images: ['https://x/c.jpg'] },
      { description: 'Fifth row' },
    ]),
  };

  const r = parseStrikePayload(body);
  assert.strictEqual(r.items.length, 5, 'every row must survive the join');
  assert.strictEqual(r.items[0].item_name, 'HEX NUT');
  assert.strictEqual(r.items[1].images.length, 2);
  assert.deepStrictEqual(r.items.map((i) => i.body),
    ['First row', 'Second row', 'Third row', 'Fourth row', 'Fifth row']);
});

test('a description containing commas is not torn in half', () => {
  const body = { items: glideItems([
    { description: 'Pieces received, inspected, and released, then shipped' },
    { description: 'Second row' },
  ]) };
  const r = parseStrikePayload(body);
  assert.strictEqual(r.items.length, 2, 'splitting on ", " would have broken this');
  assert.match(r.items[0].body, /released, then shipped$/);
});

test('a row with raw newlines inside its JSON is repaired, not dropped', () => {
  const body = { items: ['{"description":"line one\nline two"}, {"description":"second row"}'] };
  const r = parseStrikePayload(body);
  assert.strictEqual(r.items.length, 2, 'a malformed row must not take its neighbours with it');
  assert.match(r.items[0].body, /line one\nline two/);
});

test('a CSV-quoted body copied out of a Glide cell still parses', () => {
  const inner = JSON.stringify({ report_title: 'Sample Report', items: glideItems([{ description: 'A row' }]) });
  const csv = `"${inner.replace(/"/g, '""')}"`;
  const r = parseStrikePayload(csv);
  assert.strictEqual(r.document.report_title, 'Sample Report');
  assert.strictEqual(r.items.length, 1);
});

test('markdown emphasis is stripped, and every word kept', () => {
  const body = { items: glideItems([
    { description: 'Confirm the **tolerance stack-up** and the __surface finish__ before Monday.' },
  ]) };
  const r = parseStrikePayload(body);
  assert.strictEqual(r.items[0].body,
    'Confirm the tolerance stack-up and the surface finish before Monday.');
});

test("Glide's list padding collapses but the line breaks stay", () => {
  const raw = 'There are three:\n\n1.  Point one\n    \n2.  Point two\n    \n3.  Point three';
  const r = parseStrikePayload({ items: glideItems([{ description: raw }]) });
  const lines = r.items[0].body.split('\n').filter((l) => l.trim());
  assert.deepStrictEqual(lines,
    ['There are three:', '1.  Point one', '2.  Point two', '3.  Point three']);
  assert.ok(!/\n{3,}/.test(r.items[0].body), 'no runs of blank lines');
});

test('images arrive as bare URL strings', () => {
  const r = parseStrikePayload({ items: glideItems([
    { description: 'x', images: ['https://storage.googleapis.com/a.jpg', 'https://storage.googleapis.com/b.jpg'] },
  ]) });
  assert.deepStrictEqual(r.items[0].images.map((i) => i.path),
    ['https://storage.googleapis.com/a.jpg', 'https://storage.googleapis.com/b.jpg']);
});

test('report_title and reference_name are read, with the old names still accepted', () => {
  const now = parseStrikePayload({ report_title: 'New', reference_name: 'REF-1', items: glideItems([{ description: 'x' }]) });
  assert.strictEqual(now.document.report_title, 'New');
  assert.strictEqual(now.document.reference_name, 'REF-1');

  // part_number still maps through to reference_name. project_name no longer
  // maps to report_title — they are separate fields now (filename vs display).
  const legacy = parseStrikePayload({ part_number: 'BRK-1', items: glideItems([{ description: 'x' }]) });
  assert.strictEqual(legacy.document.reference_name, 'BRK-1', 'part_number still maps through');
});

test('a payload with no readable rows says so instead of going quiet', () => {
  const r = parseStrikePayload({ report_title: 'Sample', items: [] });
  assert.strictEqual(r.items.length, 0);
  assert.ok(r.warnings.some((w) => /no query or update rows/.test(w)),
    'an empty result must be reported, not silently shipped as a cover-only deck');
});

test('needs_response reads Glide booleans in every form it sends', () => {
  const r = parseStrikePayload({ items: glideItems([
    { description: 'a', needs_response: true },
    { description: 'b', needs_response: 'Yes' },
    { description: 'c', needs_response: 'false' },
    { description: 'd', needs_response: 0 },
    { description: 'e' },
  ]) });
  assert.deepStrictEqual(r.items.map((i) => i.needs_response),
    [true, true, false, false, undefined], 'the last is left for enrichment to decide');
});

test('helpers behave on their own', () => {
  assert.strictEqual(extractJsonObjects('{"a":1}, {"a":2}').length, 2);
  assert.strictEqual(undoubleQuotes('"{""a"":1}"'), '{"a":1}');
  assert.strictEqual(cleanRichText('**bold**  \n\n\n\nnext'), 'bold\n\nnext');
});

test('the OneDrive destination is read from the payload, with aliases', () => {
  const canonical = parseStrikePayload({
    report_title: 'T', drive_id: 'b!abc', folder_item_id: '01FOLDER',
    items: glideItems([{ description: 'x' }]),
  });
  assert.deepStrictEqual(canonical.storage, { driveId: 'b!abc', folderId: '01FOLDER' });

  const aliased = parseStrikePayload({
    report_title: 'T', driveId: 'b!abc', drive_item_id: '01FOLDER',
    items: glideItems([{ description: 'x' }]),
  });
  assert.deepStrictEqual(aliased.storage, { driveId: 'b!abc', folderId: '01FOLDER' });

  const absent = parseStrikePayload({ report_title: 'T', items: glideItems([{ description: 'x' }]) });
  assert.deepStrictEqual(absent.storage, { driveId: '', folderId: '' },
    'no destination in the payload falls through to the env default downstream');
});

test('report_title and project_name are independent — neither aliases the other', () => {
  // One names the file, the other is displayed on the deck. A payload that
  // sends both gets both, each used for its own purpose.
  const both = parseStrikePayload({
    report_title: 'Deck heading', project_name: 'Folder name',
    items: glideItems([{ description: 'x' }]),
  });
  assert.strictEqual(both.document.report_title, 'Deck heading');
  assert.strictEqual(both.document.project_name, 'Folder name');

  // Sending one must NOT populate the other.
  const onlyProject = parseStrikePayload({
    project_name: 'Test - 2 1014', items: glideItems([{ description: 'x' }]),
  });
  assert.strictEqual(onlyProject.document.project_name, 'Test - 2 1014');
  assert.strictEqual(onlyProject.document.report_title, null,
    'project_name must not leak into the displayed title');

  const onlyReport = parseStrikePayload({
    report_title: 'Bracket Assembly', items: glideItems([{ description: 'x' }]),
  });
  assert.strictEqual(onlyReport.document.report_title, 'Bracket Assembly');
  assert.strictEqual(onlyReport.document.project_name, null);
});

test('rfq_row_id and rfq_version are read for the Glide write-back', () => {
  const r = parseStrikePayload({
    rfq_row_id: '9vtmhesRQxSBUP144VEsYg', rfq_version: 'V1',
    items: glideItems([{ description: 'x' }]),
  });
  assert.strictEqual(r.writeback.rowId, '9vtmhesRQxSBUP144VEsYg');
  assert.strictEqual(r.writeback.version, 'V1');

  // The older row_id name still maps through.
  const legacy = parseStrikePayload({ row_id: 'ROW9', items: glideItems([{ description: 'x' }]) });
  assert.strictEqual(legacy.writeback.rowId, 'ROW9');

  const absent = parseStrikePayload({ items: glideItems([{ description: 'x' }]) });
  assert.deepStrictEqual([absent.writeback.rowId, absent.writeback.version], ['', '']);
});
