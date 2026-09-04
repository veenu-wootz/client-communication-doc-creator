/**
 * enrich.test.js — the fallback chain, which has to work with no API key at all.
 * The provider calls themselves are not exercised here; they cost money and need
 * a network. What matters is that nothing downstream ever sees an undefined
 * needs_response.
 */

const test = require('node:test');
const assert = require('node:assert');

const { enrich, guessNeedsResponse, guessTitle, isFormattingOnly } = require('../src/input/enrich');

const noProvider = (fn) => async () => {
  const saved = process.env.ENRICH_PROVIDER;
  process.env.ENRICH_PROVIDER = 'off';
  try { await fn(); } finally {
    if (saved === undefined) delete process.env.ENRICH_PROVIDER;
    else process.env.ENRICH_PROVIDER = saved;
  }
};

test('a direct question is read as needing a response', () => {
  assert.strictEqual(guessNeedsResponse('Sheet 3 says Ra 0.8, drawing says Ra 1.6. Which governs?'), true);
  assert.strictEqual(guessNeedsResponse('Please confirm the fillet radius before we cut steel.'), true);
  assert.strictEqual(guessNeedsResponse('Kindly advise on the datum.'), true);
});

test('a statement of progress is not', () => {
  assert.strictEqual(guessNeedsResponse('Fixture design is complete and released to the shop floor.'), false);
  assert.strictEqual(guessNeedsResponse('Trials shifted to 22 Sep following the drawing hold.'), false);
  assert.strictEqual(guessNeedsResponse(''), false);
});

test('a title falls back to the opening of the body, cut on a word boundary', () => {
  const t = guessTitle({ body: 'Drawing shows R2.5 at the flange base but the GD&T callout references R4 and we need a decision.' });
  assert.ok(t.length <= 60);
  assert.ok(!t.endsWith(' '), 'no trailing space');
  assert.ok(!/\S$/.test(t) === false, 'cut on a word boundary');
  assert.strictEqual(guessTitle({ body: '', title: '' }), 'Untitled');
});

test('with no provider, every row still ends up fully labelled', noProvider(async () => {
  const { items, provider } = await enrich([
    { id: 'a', title: '', body: 'Which radius governs?', images: [] },
    { id: 'b', title: 'Fixture ready', body: 'Design complete.', images: [] },
    { id: 'c', title: '', body: '', images: [{ path: 'x.png' }] },
  ]);

  assert.strictEqual(provider, null);
  for (const it of items) {
    assert.strictEqual(typeof it.needs_response, 'boolean', `${it.id}: needs_response must be resolved`);
    assert.ok(it.title && it.title.length > 0, `${it.id}: title must be resolved`);
  }
  assert.strictEqual(items[0].needs_response, true);
  assert.strictEqual(items[1].needs_response, false);
  assert.strictEqual(items[1].title, 'Fixture ready', 'a title the sender typed survives');
}));

test('an explicit flag from the form is honoured when no model runs', noProvider(async () => {
  const { items } = await enrich([
    { id: 'a', title: 'Notice', body: 'Fixture design complete.', needs_response: true, images: [] },
  ]);
  assert.strictEqual(items[0].needs_response, true, 'the sender overrode the heuristic');
}));

test('an empty document enriches to nothing without calling out', noProvider(async () => {
  const { items, warnings } = await enrich([]);
  assert.deepStrictEqual(items, []);
  assert.deepStrictEqual(warnings, []);
}));

test('a reformatted body is accepted only when nothing but whitespace changed', () => {
  const original = '1 Sheet Thickness 2 Laser Marking Details 3 Radius of the corners';

  assert.strictEqual(
    isFormattingOnly(original, '1 Sheet Thickness\n2 Laser Marking Details\n3 Radius of the corners'),
    true, 'line breaks alone are the whole point');

  assert.strictEqual(isFormattingOnly(original, '1 Sheet Thickness\n2 Laser Marking\n3 Radius of the corners'),
    false, 'a dropped word must be rejected');
  assert.strictEqual(isFormattingOnly(original, '- 1 Sheet Thickness\n- 2 Laser Marking Details\n- 3 Radius of the corners'),
    false, 'added bullet characters must be rejected');
  assert.strictEqual(isFormattingOnly(original, '2 Laser Marking Details\n1 Sheet Thickness\n3 Radius of the corners'),
    false, 'reordering must be rejected');
  assert.strictEqual(isFormattingOnly(original, ''), false, 'an empty repair must be rejected');
});

test('with no provider the body is passed through untouched', noProvider(async () => {
  const original = '1 Sheet Thickness 2 Laser Marking Details 3 Radius of the corners';
  const { items } = await enrich([{ id: 'a', title: '', body: original, images: [] }]);
  assert.strictEqual(items[0].body, original, 'no AI means the sender\'s own text, unchanged');
}));
