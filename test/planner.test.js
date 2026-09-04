/**
 * planner.test.js — Tier 1. The acceptance tests of §14, run against the slide
 * plan rather than a .pptx. Fast, and it is where nearly all the logic lives.
 */

const test = require('node:test');
const assert = require('node:assert');

const { planSlides } = require('../src/plan/planner');
const C = require('../src/plan/constants');
const { fixtures, doc, query, update, square, words } = require('./fixtures');

const plan = (name) => planSlides(fixtures[name]);
const items = (p) => p.slides.filter((s) => s.kind === 'item');
const queryNumbers = (p) => items(p).filter((s) => s.queryNo && !s.continued).map((s) => s.queryNo);
const allText = (p) => JSON.stringify(p);

/** Every line of body text a plan renders for one item, joined. */
function bodyTextOf(p, itemId) {
  return items(p)
    .filter((s) => s.itemId === itemId)
    .flatMap((s) => (s.body ? s.body.lines : []))
    .join(' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function imageCountOf(p, itemId) {
  return items(p).filter((s) => s.itemId === itemId).reduce((n, s) => n + s.images.length, 0);
}

// ── §14.1 / §14.2 — numbering ────────────────────────────────

test('1 · numbers skip unflagged items, and queries come before updates', () => {
  const p = plan('01-flagged-unflagged-flagged');
  assert.deepStrictEqual(queryNumbers(p), [1, 2]);

  // Sections: both queries first in the sender's own order, then the update.
  assert.deepStrictEqual(items(p).map((s) => s.itemId), ['a', 'c', 'b']);
  assert.strictEqual(items(p).find((s) => s.itemId === 'b').queryNo, null);
});

test('2 · ten alternating items number Q1–Q5, gapless', () => {
  const p = plan('02-alternating-ten');
  assert.deepStrictEqual(queryNumbers(p), [1, 2, 3, 4, 5]);
  assert.strictEqual(p.meta.queryCount, 5);
  assert.strictEqual(p.meta.updateCount, 5);
});

// ── §14.3–§14.6 — layout selection ───────────────────────────

test('3 · a single upright image uses IMAGE_SIDE', () => {
  assert.strictEqual(items(plan('03-one-upright-image'))[0].layout, 'IMAGE_SIDE');
});

test('4 · a wide image is stacked, and far wider than a side column would allow', () => {
  const s = items(plan('04-one-wide-image'))[0];
  assert.strictEqual(s.layout, 'IMAGE_STACKED');

  // §9.3 exists so a wide drawing is not squeezed into a half-width column.
  // It cannot always reach the full content width — a 2.5-aspect image needs
  // 4.85in of height and a reply-bearing slide only has 4.05in — so the test is
  // that it comfortably beats the column it was rescued from.
  const sideColumn = (C.CONTENT_W - C.GUTTER) * C.SIDE_IMAGE_FRACTION;
  assert.ok(s.images[0].box.w > sideColumn * 1.3,
    `wide image ${s.images[0].box.w.toFixed(2)}in should beat the ${sideColumn.toFixed(2)}in side column`);
  assert.ok(s.images[0].box.w <= C.CONTENT_W + 0.01, 'and stay inside the content width');
});

test('5 · three images grid two-over-one, equal widths, bottom cell centred', () => {
  const p = plan('05-three-images');
  const grid = items(p).find((s) => s.layout === 'IMAGE_GRID');
  assert.ok(grid, 'expected a grid slide');
  assert.strictEqual(grid.images.length, 3);
  const [a, b, c] = grid.images.map((i) => i.box);
  assert.ok(Math.abs(a.y - b.y) < 0.01, 'top row shares a baseline');
  assert.ok(c.y > a.y, 'third image sits on a second row');
  const centre = C.CONTENT_X + C.CONTENT_W / 2;
  assert.ok(Math.abs((c.x + c.w / 2) - centre) < 0.02, 'bottom cell is horizontally centred');
});

test('6 · six images all render — nothing is dropped (PLAN.md §5.1)', () => {
  const p = plan('06-six-images');
  assert.strictEqual(imageCountOf(p, 'a'), 6);
  assert.ok(p.meta.warnings.some((w) => /6 images/.test(w)), 'exceeding the form cap is logged');
});

// ── §14.7 / §14.8 — overflow and the reply box ───────────────

test('7 · long body with four images: every word and image survives', () => {
  const p = plan('07-long-body-4-images');
  const slides = items(p).filter((s) => s.itemId === 'a');
  assert.ok(slides.length > 1, 'expected the item to span several slides');
  assert.strictEqual(imageCountOf(p, 'a'), 4);

  const rendered = bodyTextOf(p, 'a');
  const source = words(900).replace(/\s+/g, ' ').trim();
  assert.strictEqual(rendered, source, 'body text must survive intact across slides');

  assert.ok(slides[0].body && slides[0].body.lines.length > 0, 'body starts on the first slide');
  assert.ok(slides[0].images.length >= 1, 'the first slide keeps visual context');
});

test('8 · a multi-slide query carries exactly one correctly numbered reply box', () => {
  const p = plan('07-long-body-4-images');
  const boxes = items(p).filter((s) => s.itemId === 'a' && s.replyBox);
  assert.strictEqual(boxes.length, 1);
  assert.match(boxes[0].replyBox.label, /Q1$/);
  const slides = items(p).filter((s) => s.itemId === 'a');
  assert.strictEqual(slides[slides.length - 1].replyBox !== null, true, 'reply box is on the last slide');
});

test('9 · an unflagged item is labelled Update, with no number and no reply box', () => {
  const p = plan('09-unflagged-only');
  for (const s of items(p)) {
    assert.strictEqual(s.replyBox, null);
    assert.strictEqual(s.queryNo, null);
    assert.strictEqual(s.header.chip.text, 'Update', 'no slide is left unlabelled');
    assert.strictEqual(s.header.chip.tone, 'update');
  }
  assert.ok(!/Q\d/.test(JSON.stringify(items(p))), 'no query number text on an update slide');
});

// ── §14.10–§14.13 — contents slide ───────────────────────────

test('10 · two queries: no contents slide, instruction on the cover', () => {
  const p = plan('10-two-queries');
  assert.strictEqual(p.meta.hasContents, false);
  assert.ok(!p.slides.some((s) => s.kind === 'contents'));
  assert.ok(p.slides[0].instruction, 'cover carries the reply instruction');
});

test('11 · four queries: a summary listing exactly four query entries', () => {
  const p = plan('11-four-queries');
  const contents = p.slides.filter((s) => s.kind === 'contents');
  assert.strictEqual(contents.length, 1);
  assert.strictEqual(contents[0].rows.filter((r) => r.type === 'entry').length, 4);
  assert.strictEqual(contents[0].heading.text, 'Summary');
  assert.match(contents[0].rows[0].text, /^QUERIES — 4 need your response$/);
  assert.strictEqual(p.slides[0].instruction, null, 'instruction moves to the summary');
});

test('12 · the summary lists both sections, and keeps the count in the label', () => {
  const c = plan('12-four-queries-two-updates').slides.find((s) => s.kind === 'contents');
  const sections = c.rows.filter((r) => r.type === 'section').map((r) => r.text);
  assert.deepStrictEqual(sections, ['QUERIES — 4 need your response', 'OTHER UPDATES']);

  const entries = c.rows.filter((r) => r.type === 'entry');
  assert.strictEqual(entries.filter((e) => e.n).length, 4, 'four numbered queries');
  assert.ok(entries.some((e) => /Fixture/.test(e.text)), 'updates are named, not just counted');
  assert.ok(entries.some((e) => /Timeline/.test(e.text)));

  // The queries section comes first, so the count is the first thing read.
  assert.strictEqual(c.rows[0].text, 'QUERIES — 4 need your response');
});

test('13 · zero flagged items: no contents, no numbers, no instruction', () => {
  const p = plan('13-zero-flagged');
  assert.strictEqual(p.meta.queryCount, 0);
  assert.ok(!p.slides.some((s) => s.kind === 'contents'));
  assert.strictEqual(p.slides[0].instruction, null);
  assert.ok(!items(p).some((s) => s.replyBox));
});

// ── §14.14–§14.16 — degenerate input ─────────────────────────

test('14 · an empty items array emits the cover alone', () => {
  const p = plan('14-empty-items');
  assert.strictEqual(p.meta.totalSlides, 1);
  assert.strictEqual(p.slides[0].kind, 'cover');
});

test('15 · a cover with every optional field null composes without gaps', () => {
  const cover = plan('15-cover-all-null').slides[0];
  assert.strictEqual(cover.variant, 'B');
  assert.deepStrictEqual(cover.blocks.map((b) => b.kind), ['title']);
  assert.ok(cover.footer.text.includes('Priya Nair'));
});

test('16 · a broken image renders a placeholder and warns, never crashes', () => {
  const p = plan('16-broken-image');
  const s = items(p)[0];
  assert.strictEqual(s.images.length, 1);
  assert.strictEqual(s.images[0].missing, true);
});

// ── §14.17 / §14.18 — grouping ───────────────────────────────

test('17 · five text-only updates group at most three per slide, in order', () => {
  const p = plan('17-five-updates');
  const grouped = p.slides.filter((s) => s.kind === 'grouped');
  assert.ok(grouped.length >= 1);
  for (const g of grouped) assert.ok(g.blocks.length <= C.MAX_GROUPED_UPDATES);
  const seen = p.slides.flatMap((s) =>
    s.kind === 'grouped' ? s.blocks.map((b) => b.itemId) : s.kind === 'item' ? [s.itemId] : []);
  assert.deepStrictEqual([...new Set(seen)], ['u0', 'u1', 'u2', 'u3', 'u4']);
});

test('18 · an update between two queries is never grouped', () => {
  const p = plan('18-update-between-queries');
  assert.strictEqual(p.slides.filter((s) => s.kind === 'grouped').length, 0);
});

// ── §14.19 — determinism ─────────────────────────────────────

test('19 · the same input plans identically twice', () => {
  for (const name of Object.keys(fixtures)) {
    const a = planSlides(fixtures[name]);
    const b = planSlides(fixtures[name]);
    assert.deepStrictEqual(a, b, `${name} is not deterministic`);
  }
});

// ── Rules that exist because we flow instead of truncating ───

test('20 · no ellipsis is ever introduced into rendered content', () => {
  for (const name of Object.keys(fixtures)) {
    const p = planSlides(fixtures[name]);
    for (const s of p.slides) {
      if (s.kind === 'cover' || s.kind === 'contents') continue;
      const text = JSON.stringify([s.body, s.images, s.header, s.blocks]);
      assert.ok(!text.includes('…'), `${name}: ellipsis found on slide ${s.page}`);
    }
  }
});

test('21 · forty queries paginate the contents with no entry shortened', () => {
  const p = plan('22-forty-queries');
  const contents = p.slides.filter((s) => s.kind === 'contents');
  assert.ok(contents.length > 1, 'expected the contents list to paginate');
  const entries = contents.flatMap((s) => s.rows.filter((r) => r.type === 'entry'));
  assert.strictEqual(entries.length, 40);
  assert.deepStrictEqual(entries.map((e) => e.n), Array.from({ length: 40 }, (_, i) => i + 1));
  for (const e of entries) assert.ok(!e.text.includes('…'), 'summary entries are never truncated');
  assert.strictEqual(contents[0].heading.text, 'Summary');
  assert.strictEqual(contents.at(-1).tail.at(-1).kind, 'instruction');
});

test('22 · an empty body with one image uses the stacked variant', () => {
  const s = items(plan('23-empty-body-with-image'))[0];
  assert.strictEqual(s.layout, 'IMAGE_STACKED');
  assert.strictEqual(s.body, null);
});

test('23 · an item with no body and no images renders the header alone', () => {
  const p = plan('24-no-body-no-images');
  const s = items(p)[0];
  assert.strictEqual(s.body, null);
  assert.strictEqual(s.images.length, 0);
  assert.ok(s.replyBox, 'still a query, so it still owes a reply');
});

test('24 · a cover with a product photo uses variant A and reserves a photo zone', () => {
  const cover = plan('21-cover-with-photo').slides[0];
  assert.strictEqual(cover.variant, 'A');
  assert.ok(cover.photo.zone.w > 3, 'photo zone has real width');
  assert.ok(cover.photo.zone.x > C.CONTENT_X + C.CONTENT_W * 0.5, 'photo sits in the right column');
});

// ── Geometry invariants across every fixture ─────────────────

test('25 · every planned box sits inside the canvas and within the margins', () => {
  const EPS = 0.02;
  for (const name of Object.keys(fixtures)) {
    const p = planSlides(fixtures[name]);
    for (const s of p.slides) {
      for (const { box, what } of collectBoxes(s)) {
        assert.ok(box.x >= C.MARGIN - EPS, `${name} p${s.page} ${what}: x=${box.x.toFixed(2)} left of margin`);
        assert.ok(box.y >= 0.3 - EPS, `${name} p${s.page} ${what}: y=${box.y.toFixed(2)} above the top`);
        assert.ok(box.x + box.w <= C.SLIDE_W - C.MARGIN + EPS,
          `${name} p${s.page} ${what}: right edge ${(box.x + box.w).toFixed(2)} past the margin`);
        assert.ok(box.y + box.h <= C.SLIDE_H - 0.1 + EPS,
          `${name} p${s.page} ${what}: bottom ${(box.y + box.h).toFixed(2)} off the slide`);
      }
    }
  }
});

test('26 · content never collides with the reply box', () => {
  for (const name of Object.keys(fixtures)) {
    const p = planSlides(fixtures[name]);
    for (const s of p.slides) {
      if (s.kind !== 'item' || !s.replyBox) continue;
      const top = s.replyBox.box.y;
      if (s.body) assert.ok(s.body.box.y + s.body.box.h <= top + 0.02, `${name} p${s.page}: body overlaps the reply box`);
      for (const im of s.images) {
        const bottom = im.caption ? im.caption.box.y + im.caption.box.h : im.box.y + im.box.h;
        assert.ok(bottom <= top + 0.02, `${name} p${s.page}: image overlaps the reply box`);
      }
    }
  }
});

/** Every positioned box on a slide, with a label for failure messages. */
function collectBoxes(s) {
  const out = [];
  const add = (box, what) => { if (box && Number.isFinite(box.x)) out.push({ box, what }); };

  if (s.logo) add(s.logo.box, 'logo');
  if (s.footer) add(s.footer.box, 'footer');
  if (s.footer && s.footer.page) add(s.footer.page.box, 'page');

  if (s.kind === 'cover') {
    for (const b of s.blocks) { add(b.box, b.kind); add(b.labelBox, 'label'); add(b.valueBox, 'value'); }
    if (s.instruction) add(s.instruction.box, 'instruction');
    if (s.photo) add(s.photo.zone, 'photo');
  } else if (s.kind === 'contents') {
    if (s.heading) add(s.heading.box, 'heading');
    for (const r of s.rows) add(r.box, `${r.type}${r.n ? ` Q${r.n}` : ''}`);
    for (const t of s.tail || []) add(t.box, t.kind);
  } else if (s.kind === 'grouped') {
    if (s.chip) add(s.chip.box, 'update chip');
    for (const b of s.blocks) { add(b.title.box, 'block title'); if (b.body) add(b.body.box, 'block body'); }
  } else if (s.kind === 'item') {
    if (s.header.chip) add(s.header.chip.box, 'chip');
    add(s.header.title.box, 'header title');
    if (s.body) add(s.body.box, 'body');
    for (const im of s.images) { add(im.box, 'image'); if (im.caption) add(im.caption.box, 'caption'); }
    if (s.replyBox) { add(s.replyBox.box, 'reply box'); add(s.replyBox.labelBox, 'reply label'); }
  }
  return out;
}

// ── Fixes from the first real-content review ─────────────────

test('27 · the contents footer never prints underneath the logo', () => {
  for (const name of Object.keys(fixtures)) {
    const p = planSlides(fixtures[name]);
    for (const s of p.slides.filter((x) => x.kind === 'contents')) {
      if (!s.logo || !s.footer.text) continue;
      const logoRight = s.logo.box.x + s.logo.box.w;
      assert.ok(s.footer.box.x >= logoRight,
        `${name}: contents footer starts at ${s.footer.box.x.toFixed(2)}, logo ends at ${logoRight.toFixed(2)}`);
    }
  }
});

test('28 · the contents slide drops the author, the cover keeps it', () => {
  const p = planSlides(fixtures['11-four-queries']);
  assert.ok(p.slides[0].footer.text.includes('Priya Nair'), 'cover names the author');
  const contents = p.slides.find((s) => s.kind === 'contents');
  assert.ok(!contents.footer.text.includes('Priya Nair'), 'contents does not — the logo covers it');
  assert.ok(contents.footer.text.includes('2026'), 'contents keeps the date');
});

test('29 · a low-resolution image is not blown up past its own detail', () => {
  const p = planSlides({
    document: { project_name: 'Test', created_by: 'X', created_at: '2026-09-03' },
    items: [{ id: 'a', title: 'Small screenshot', needs_response: true, body: 'Please confirm.',
      images: [{ path: 'tiny.png', width: 200, height: 150, aspect: 200 / 150, ok: true }] }],
  });
  const im = items(p)[0].images[0];
  const dpi = 200 / im.box.w;
  assert.ok(dpi >= C.MIN_RENDER_DPI - 1,
    `a 200px image rendered ${im.box.w.toFixed(2)}in wide is only ${dpi.toFixed(0)} DPI`);
});

test('30 · a high-resolution image still fills the space available', () => {
  const p = planSlides({
    document: { project_name: 'Test', created_by: 'X', created_at: '2026-09-03' },
    items: [{ id: 'a', title: 'Drawing', needs_response: true, body: 'Please confirm.',
      images: [{ path: 'big.png', width: 1600, height: 1200, aspect: 1600 / 1200, ok: true }] }],
  });
  const im = items(p)[0].images[0];
  assert.ok(im.box.w > 4, `expected a large render, got ${im.box.w.toFixed(2)}in`);
});


test('32 · item_name appears above the header and names the SKU on contents', () => {
  const p = plan('25-multiple-skus');

  const first = items(p).find((s) => s.itemId === 'a');
  assert.ok(first.header.eyebrow, 'an item with a SKU gets an eyebrow line');
  assert.strictEqual(first.header.eyebrow.text, 'HEX NUT M12');
  assert.ok(first.header.title.box.y > first.header.eyebrow.box.y,
    'the title sits below the SKU, not on top of it');

  const bare = items(p).find((s) => s.itemId === 'd');
  assert.strictEqual(bare.header.eyebrow, null, 'no SKU means no eyebrow and no wasted height');
  assert.ok(bare.header.title.box.y < first.header.title.box.y);

  const contents = p.slides.find((s) => s.kind === 'contents');
  const entries = contents.rows.filter((r) => r.type === 'entry');
  assert.match(entries[0].text, /^Q1 · HEX NUT M12 — /);
  assert.match(entries[2].text, /^Q3 · Tolerance stack-up$/, 'no SKU, no separator');
});

test('33 · an item_name never displaces the query number', () => {
  const p = plan('25-multiple-skus');
  for (const s of items(p)) {
    if (!s.queryNo) continue;
    assert.ok(s.header.chip, `Q${s.queryNo} must keep its chip alongside the SKU`);
    assert.strictEqual(s.header.chip.text, `Q${s.queryNo}`);
  }
});

test('34 · a numbered list in the body keeps one point per line', () => {
  const p = plan('26-numbered-list-body');
  const s = items(p)[0];
  const numbered = s.body.lines.filter((l) => /^\d+\./.test(l.trim()));
  assert.strictEqual(numbered.length, 4, 'each point renders on its own line');
});

test('35 · the cover carries one reference field, not part plus PO', () => {
  const cover = plan('01-flagged-unflagged-flagged').slides[0];
  const labels = cover.blocks.filter((b) => b.kind === 'field').map((b) => b.label);
  assert.deepStrictEqual(labels, ['Reference', 'Attention']);
});

// ── Sections (PLAN.md §5.8) ──────────────────────────────────

test('36 · every item slide is labelled — a number or an Update chip', () => {
  for (const name of Object.keys(fixtures)) {
    const p = planSlides(fixtures[name]);
    for (const s of p.slides) {
      if (s.kind === 'item') {
        assert.ok(s.header.chip, `${name} p${s.page}: an unlabelled slide leaves the reader guessing`);
        assert.strictEqual(s.header.chip.tone, s.queryNo ? 'query' : 'update');
        if (s.queryNo) assert.strictEqual(s.header.chip.text, `Q${s.queryNo}`);
        else assert.strictEqual(s.header.chip.text, 'Update');
      }
      if (s.kind === 'grouped') {
        assert.ok(s.chip, `${name} p${s.page}: grouped updates need labelling too`);
        assert.strictEqual(s.chip.text, 'Update');
      }
    }
  }
});

test('37 · queries come first, updates after, order kept inside each section', () => {
  const p = plan('02-alternating-ten');
  const ids = items(p).filter((s) => !s.continued).map((s) => s.itemId);
  const grouped = p.slides.filter((s) => s.kind === 'grouped').flatMap((s) => s.blocks.map((b) => b.itemId));
  const seen = [...ids, ...grouped];

  const queries = seen.filter((id) => id.startsWith('q'));
  const updates = seen.filter((id) => id.startsWith('u'));

  assert.deepStrictEqual(queries, ['q0', 'q2', 'q4', 'q6', 'q8'], 'queries keep the sender order');
  assert.deepStrictEqual(updates, ['u1', 'u3', 'u5', 'u7', 'u9'], 'so do updates');

  const lastQuery = seen.lastIndexOf(queries.at(-1));
  const firstUpdate = seen.indexOf(updates[0]);
  assert.ok(lastQuery < firstUpdate, 'no update appears before the final query');
});

test('38 · grouping updates together shortens the deck', () => {
  // Interleaved in the input, the five updates could never group under the old
  // adjacency rule. Sectioning makes them contiguous, so they share slides.
  const p = plan('02-alternating-ten');
  const grouped = p.slides.filter((s) => s.kind === 'grouped');
  assert.ok(grouped.length > 0, 'expected updates to share slides once contiguous');
  for (const g of grouped) assert.ok(g.blocks.length <= C.MAX_GROUPED_UPDATES);
});

test('39 · a deck of only updates still labels every slide and owes nothing', () => {
  const p = plan('13-zero-flagged');
  assert.strictEqual(p.meta.queryCount, 0);
  assert.ok(!p.slides.some((s) => s.kind === 'contents'), 'nothing is owed, so no summary');
  for (const s of p.slides) {
    if (s.kind === 'item') assert.strictEqual(s.header.chip.text, 'Update');
    if (s.kind === 'grouped') assert.strictEqual(s.chip.text, 'Update');
  }
});

test('40 · the summary threshold counts every item, not just the queries', () => {
  const build = (q, u) => planSlides({
    document: doc(),
    items: [
      ...Array.from({ length: q }, (_, i) => query(`q${i}`, `Question ${i + 1}`, 'Please confirm.')),
      ...Array.from({ length: u }, (_, i) => update(`u${i}`, `Update ${i + 1}`, 'Progress note.')),
    ],
  });
  const hasSummary = (p) => p.slides.some((s) => s.kind === 'contents');

  assert.strictEqual(hasSummary(build(2, 0)), false, 'two items is too few to need a summary');
  assert.strictEqual(hasSummary(build(1, 2)), true, 'one query plus two updates is three items');
  assert.strictEqual(hasSummary(build(0, 5)), true, 'a deck of updates still gets a summary');
});

test('41 · a summary with nothing owed does not ask for a query number', () => {
  const p = plan('17-five-updates');
  const sum = p.slides.find((s) => s.kind === 'contents');
  assert.ok(sum, 'five updates is past the threshold');
  assert.deepStrictEqual(sum.rows.filter((r) => r.type === 'section').map((r) => r.text), ['OTHER UPDATES']);
  assert.deepStrictEqual(sum.tail, [], 'no queries, so no "reply quoting the query number"');
});

// ── Second review round: summary furniture and continuation ──

test('42 · the summary carries the same footer as every other slide', () => {
  const p = plan('12-four-queries-two-updates');
  const summary = p.slides.find((s) => s.kind === 'contents');
  const item = items(p)[0];

  assert.strictEqual(summary.footer.text, item.footer.text,
    'the summary should not have a footer of its own kind');
  assert.ok(summary.footer.page, 'and it carries the page indicator');
  assert.strictEqual(summary.pageLabel, `Page ${summary.page}`);
  assert.strictEqual(p.slides[0].pageLabel, undefined, 'the cover stays unnumbered');
});

test('43 · the summary heading sits level with the title on other slides', () => {
  const summary = plan('12-four-queries-two-updates').slides.find((s) => s.kind === 'contents');
  assert.strictEqual(summary.heading.box.y, C.HEADER_Y,
    'anchored at HEADER_Y, not an inch lower at CONTENT_Y');
});

test('44 · a section label that follows a list gets room above it', () => {
  const summary = plan('12-four-queries-two-updates').slides.find((s) => s.kind === 'contents');
  const rows = summary.rows;
  const updates = rows.findIndex((r) => r.type === 'section' && /OTHER UPDATES/.test(r.text));
  assert.ok(updates > 0, 'expected an updates section after the queries');

  const lastQuery = rows[updates - 1];
  const gapBeforeSection = rows[updates].box.y - (lastQuery.box.y + lastQuery.box.h);
  const entryGap = rows[updates + 2].box.y - (rows[updates + 1].box.y + rows[updates + 1].box.h);

  assert.ok(gapBeforeSection > entryGap * 2,
    `a section break (${gapBeforeSection.toFixed(2)}in) should read wider than a line break (${entryGap.toFixed(2)}in)`);

  // The first section needs no lead — it sits directly under the heading.
  assert.strictEqual(rows[0].lead, 0);
});

test('45 · a slide that carries on says so at its foot; the last one does not', () => {
  const p = plan('07-long-body-4-images');
  const slides = items(p).filter((s) => s.itemId === 'a');
  assert.ok(slides.length > 1, 'expected a multi-slide item');

  for (const s of slides.slice(0, -1)) {
    assert.ok(s.continuesNote, `p${s.page} continues, so it must say so`);
    assert.match(s.continuesNote.text, /[Cc]ontinued/);
  }
  assert.strictEqual(slides.at(-1).continuesNote, null, 'the final slide has nothing to continue to');
});

test('46 · the continuation note never overlaps content, the reply box, or the footer', () => {
  for (const name of Object.keys(fixtures)) {
    const p = planSlides(fixtures[name]);
    for (const s of p.slides) {
      if (s.kind !== 'item' || !s.continuesNote) continue;
      const note = s.continuesNote.box;

      if (s.body) {
        assert.ok(s.body.box.y + s.body.box.h <= note.y + 0.02,
          `${name} p${s.page}: body runs into the continuation note`);
      }
      for (const im of s.images) {
        const bottom = im.caption ? im.caption.box.y + im.caption.box.h : im.box.y + im.box.h;
        assert.ok(bottom <= note.y + 0.02, `${name} p${s.page}: image runs into the continuation note`);
      }
      assert.ok(note.y + note.h <= C.FOOTER_Y + 0.02,
        `${name} p${s.page}: continuation note collides with the footer`);
      assert.strictEqual(s.replyBox, null, 'a continuing slide never holds the reply box');
    }
  }
});
