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
const numbers = (p) => items(p).filter((s) => !s.continued).map((s) => s.number);
const templates = (p) => p.slides.filter((s) => s.kind === 'template');
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

test('1 · every item is numbered, in the order the sender wrote it', () => {
  const p = plan('01-flagged-unflagged-flagged');
  // Flat model: the update between two queries is numbered like everything
  // else and stays where it was put.
  assert.deepStrictEqual(numbers(p), [1, 2, 3]);
  assert.deepStrictEqual(items(p).map((s) => s.itemId), ['a', 'b', 'c']);
});

test('2 · ten items number 1–10, gapless, regardless of needs_response', () => {
  const p = plan('02-alternating-ten');
  assert.deepStrictEqual(numbers(p), [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
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

test('5 · three wide images grid two-over-one, equal widths, bottom cell centred', () => {
  const p = plan('05b-three-wide-images');
  const grid = items(p).find((s) => s.layout === 'IMAGE_GRID');
  assert.ok(grid, 'expected a grid slide');
  assert.strictEqual(grid.images.length, 3);
  const [a, b, c] = grid.images.map((i) => i.box);
  assert.ok(Math.abs(a.y - b.y) < 0.01, 'top row shares a baseline');
  assert.ok(c.y > a.y, 'third image sits on a second row');
  const centre = C.CONTENT_X + C.CONTENT_W / 2;
  assert.ok(Math.abs((c.x + c.w / 2) - centre) < 0.02, 'bottom cell is horizontally centred');
});

test('5b · images too small to read in a grid spread over slides instead', () => {
  // A square drawing in a 2-column grid renders about 1.8in across — no
  // dimension on it can be read. Legibility beats packing (spec §9.6).
  const p = plan('05-three-images');
  const slides = items(p);
  assert.ok(slides.length > 1, 'expected the item to spread rather than grid');
  for (const s of slides) {
    for (const im of s.images) {
      assert.ok(im.box.w >= C.MIN_IMAGE_W,
        `p${s.page}: image renders ${im.box.w.toFixed(2)}in, below the ${C.MIN_IMAGE_W}in floor`);
    }
  }
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

test('8 · a multi-slide item carries exactly one correctly numbered reply box', () => {
  const p = plan('07-long-body-4-images');
  const boxes = items(p).filter((s) => s.itemId === 'a' && s.replyBox);
  assert.strictEqual(boxes.length, 1);
  assert.match(boxes[0].replyBox.label, /1$/);
  const slides = items(p).filter((s) => s.itemId === 'a');
  assert.strictEqual(slides[slides.length - 1].replyBox !== null, true, 'reply box is on the last slide');
});

test('9 · an update is numbered and gets a reply box like anything else', () => {
  const p = plan('09-unflagged-only');
  const slides = items(p);
  assert.strictEqual(slides[0].number, 1);
  assert.strictEqual(slides[0].header.chip.text, '1');
  assert.ok(slides.at(-1).replyBox, 'reply box on every item, updates included');
  assert.ok(!/\bQ\d/.test(JSON.stringify(slides)), 'no Q prefix survives anywhere');
});

// ── §14.10–§14.13 — contents slide ───────────────────────────

test('10 · two queries: no contents slide, instruction on the cover', () => {
  const p = plan('10-two-queries');
  assert.strictEqual(p.meta.hasContents, false);
  assert.ok(!p.slides.some((s) => s.kind === 'contents'));
  assert.ok(p.slides[0].instruction, 'cover carries the reply instruction');
});

test('11 · four items: one summary slide listing all four', () => {
  const p = plan('11-four-queries');
  const contents = p.slides.filter((s) => s.kind === 'contents');
  assert.strictEqual(contents.length, 1);
  assert.strictEqual(contents[0].heading.text, 'Summary');
  assert.strictEqual(contents[0].list.entries.length, 4);
  assert.strictEqual(contents[0].list.startAt, 1);
  assert.strictEqual(p.slides[0].instruction, null, 'instruction moves to the summary');
});

test('12 · the summary is one flat numbered list, with no counts to go stale', () => {
  const c = plan('12-four-queries-two-updates').slides.find((s) => s.kind === 'contents');
  assert.strictEqual(c.list.entries.length, 6, 'queries and updates in one list');
  assert.deepStrictEqual(c.list.entries.map((e) => e.n), [1, 2, 3, 4, 5, 6]);
  assert.ok(c.list.entries.some((e) => /Fixture/.test(e.text)));

  // A hardcoded "4 need a reply" would go stale the moment a slide is added.
  const text = JSON.stringify(c);
  assert.ok(!/need your response|QUERIES|OTHER UPDATES/.test(text), 'no section labels, no counts');
});

test('13 · a deck of updates is numbered and answerable like any other', () => {
  const p = plan('13-zero-flagged');
  assert.strictEqual(p.meta.queryCount, 0, 'still counted for the email summary');
  assert.deepStrictEqual(numbers(p), [1, 2]);
  assert.ok(items(p).some((s) => s.replyBox), 'updates take replies too now');
});

// ── §14.14–§14.16 — degenerate input ─────────────────────────

test('14 · an empty items array emits the cover plus the blank templates', () => {
  const p = plan('14-empty-items');
  assert.strictEqual(p.slides[0].kind, 'cover');
  assert.strictEqual(items(p).length, 0);
  assert.strictEqual(templates(p).length, 2, 'the sender still gets something to start from');
});

test('15 · a cover with empty fields still shows every label, with a rule to write on', () => {
  const cover = plan('15-cover-all-null').slides[0];
  assert.strictEqual(cover.variant, 'B');

  const labels = cover.blocks.filter((b) => b.kind === 'field').map((b) => b.label);
  assert.ok(labels.includes('Reference number'), 'the label renders even with no value');
  assert.ok(labels.includes('Additional information'));

  const empty = cover.blocks.filter((b) => b.kind === 'field' && !b.filled);
  assert.ok(empty.length > 0);
  for (const b of empty) assert.ok(b.rule, 'an empty field gets a hairline rule, not a filled box');

  const prepared = cover.blocks.find((b) => b.label === 'Prepared by');
  assert.ok(prepared.lines.join(' ').includes('Priya Nair'), 'author moved out of the footer');
});

test('16 · a broken image renders a placeholder and warns, never crashes', () => {
  const p = plan('16-broken-image');
  const s = items(p)[0];
  assert.strictEqual(s.images.length, 1);
  assert.strictEqual(s.images[0].missing, true);
});

// ── §14.17 / §14.18 — grouping ───────────────────────────────

test('17 · five updates each get their own slide, numbered in order', () => {
  // Grouping is gone: uniform slides matter more than a shorter deck, because
  // every slide has to be duplicable by hand (PLAN.md §5.14).
  const p = plan('17-five-updates');
  assert.ok(!p.slides.some((s) => s.kind === 'grouped'), 'no grouped slides remain');
  assert.deepStrictEqual(items(p).map((s) => s.itemId), ['u0', 'u1', 'u2', 'u3', 'u4']);
  assert.deepStrictEqual(numbers(p), [1, 2, 3, 4, 5]);
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
  const entries = contents.flatMap((s) => s.list.entries);
  assert.strictEqual(entries.length, 40);
  assert.deepStrictEqual(entries.map((e) => e.n), Array.from({ length: 40 }, (_, i) => i + 1));
  for (const e of entries) assert.ok(!e.text.includes('…'), 'summary entries are never truncated');
  assert.strictEqual(contents[0].heading.text, 'Summary');
  assert.strictEqual(contents.at(-1).tail.at(-1).kind, 'instruction');

  // Each paginated slide restarts PowerPoint's own numbering where the last left
  // off, so the visible numbers stay continuous across the break.
  let expected = 1;
  for (const c of contents) {
    assert.strictEqual(c.list.startAt, expected, 'numbering must continue across summary slides');
    expected += c.list.entries.length;
  }
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
    if (s.list) add(s.list.box, 'summary list');
    for (const t of s.tail || []) add(t.box, t.kind);
  } else if (s.kind === 'template') {
    add(s.header.chip.box, 'chip');
    add(s.titleBox, 'title placeholder');
    add(s.bodyBox, 'body placeholder');
    for (const b of s.images) add(b, 'image placeholder');
    add(s.replyBox.box, 'reply box');
  } else if (s.kind === 'item') {
    if (s.header.chip) add(s.header.chip.box, 'chip');
    add(s.header.title.box, 'header title');
    if (s.body) add(s.body.box, 'body');
    if (s.placeholder) add(s.placeholder.box, 'image placeholder');
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

test('28 · the author sits on the cover body, not in any footer', () => {
  const p = planSlides(fixtures['11-four-queries']);
  const cover = p.slides[0];
  assert.strictEqual(cover.footer, null, 'the cover footer is gone entirely');

  const prepared = cover.blocks.find((b) => b.label === 'Prepared by');
  assert.ok(prepared.lines.join(' ').includes('Priya Nair'));

  const contents = p.slides.find((s) => s.kind === 'contents');
  assert.ok(!contents.footer.text.includes('Priya Nair'), 'not repeated on the summary');
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
  const entries = contents.list.entries;
  // The number comes from PowerPoint's bullet, so the text is the title alone.
  assert.match(entries[0].text, /^HEX NUT M12 — /);
  assert.strictEqual(entries[3].text, 'Tolerance stack-up', 'no SKU, no separator');
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

test('35 · the cover is a template: every field label always present', () => {
  const cover = plan('01-flagged-unflagged-flagged').slides[0];
  const labels = cover.blocks.filter((b) => b.kind === 'field').map((b) => b.label);
  assert.deepStrictEqual(labels,
    ['Reference number', 'Attention', 'Additional information', 'Prepared by']);
});

// ── Sections (PLAN.md §5.8) ──────────────────────────────────

test('36 · every item slide carries a number chip, identical in style', () => {
  for (const name of Object.keys(fixtures)) {
    const p = planSlides(fixtures[name]);
    for (const s of p.slides.filter((x) => x.kind === 'item')) {
      assert.ok(s.header.chip, `${name} p${s.page}: an unlabelled slide reads as an oversight`);
      assert.strictEqual(s.header.chip.text, String(s.number));
      assert.strictEqual(s.header.chip.size, C.CHIP.size,
        'one chip style everywhere, so a hand-added slide matches');
    }
  }
});

test('37 · input order is preserved exactly — nothing is regrouped', () => {
  const p = plan('02-alternating-ten');
  const ids = items(p).filter((s) => !s.continued).map((s) => s.itemId);
  assert.deepStrictEqual(ids,
    ['q0', 'u1', 'q2', 'u3', 'q4', 'u5', 'q6', 'u7', 'q8', 'u9'],
    'queries and updates stay interleaved as written');
});

test('38 · every item ends with a reply box, whatever it is', () => {
  for (const name of Object.keys(fixtures)) {
    const p = planSlides(fixtures[name]);
    const byItem = new Map();
    for (const s of p.slides.filter((x) => x.kind === 'item')) {
      if (!byItem.has(s.itemId)) byItem.set(s.itemId, []);
      byItem.get(s.itemId).push(s);
    }
    for (const [id, slides] of byItem) {
      const boxes = slides.filter((s) => s.replyBox);
      assert.strictEqual(boxes.length, 1, `${name}/${id}: exactly one reply box per item`);
      assert.strictEqual(boxes[0], slides.at(-1), `${name}/${id}: and it is on the last slide`);
      assert.strictEqual(boxes[0].replyBox.label, `Your response — ${boxes[0].number}`);
    }
  }
});

test('39 · an item with no images reserves the column; one with images does not', () => {
  const bare = items(plan('24-no-body-no-images'))[0];
  assert.strictEqual(bare.layout, 'IMAGE_PLACEHOLDER');
  assert.ok(bare.placeholder, 'a sender can see where a picture would go');

  // The adaptive layout is untouched wherever images actually exist.
  const withImage = items(plan('03-one-upright-image'))[0];
  assert.strictEqual(withImage.layout, 'IMAGE_SIDE');
  assert.strictEqual(withImage.placeholder, null);
  assert.strictEqual(items(plan('04-one-wide-image'))[0].layout, 'IMAGE_STACKED');
  assert.ok(items(plan('05b-three-wide-images')).some((s) => s.layout === 'IMAGE_GRID'));
});

test('40 · a continuation slide that ran out of images gets no placeholder', () => {
  // The item has pictures, just not on this slide — offering to add one would
  // tell the reader something untrue.
  const p = plan('07-long-body-4-images');
  for (const s of items(p).filter((x) => x.continued)) {
    assert.strictEqual(s.placeholder, null, `p${s.page} must not offer an image slot`);
  }
});

test('41 · two numbered template slides close every deck', () => {
  for (const name of Object.keys(fixtures)) {
    const p = planSlides(fixtures[name]);
    const t = templates(p);
    assert.strictEqual(t.length, 2, `${name}: one single-image and one multi-image template`);
    // One nudge per generated item slide, but a template is a worked example —
    // the second shows the multi-image arrangement.
    assert.deepStrictEqual(t.map((x) => x.variant), ['single', 'multi']);
    assert.strictEqual(t[0].images.length, 1);
    assert.strictEqual(t[1].images.length, 3);

    // Numbered on from the last real item, so a copied template needs no renumbering.
    const last = items(p).filter((s) => !s.continued).length;
    assert.deepStrictEqual(t.map((x) => x.number), [last + 1, last + 2], `${name}: numbering continues`);
    for (const x of t) assert.ok(x.replyBox, 'templates carry the same reply box as real slides');

    assert.strictEqual(p.slides.at(-1).kind, 'template', 'and they sit at the very end');
  }
});

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

test('44 · the summary list is one block, so Enter continues the numbering', () => {
  // The whole point of the flat summary: separate boxes per row made it
  // impossible for a sender to add a line (PLAN.md §5.14).
  const c = plan('12-four-queries-two-updates').slides.find((x) => x.kind === 'contents');
  assert.ok(c.list, 'a single list, not a box per entry');
  assert.strictEqual(c.rows, undefined, 'the per-row model is gone');
  assert.ok(c.list.indent > 0, 'room is left for the bullet PowerPoint draws');
  assert.strictEqual(c.list.entries.length, 6);
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

test('47 · the response area is a full-size box, not a strip at its top', () => {
  // A one-line label meant clicking the reply area gave a sliver to type in.
  for (const name of Object.keys(fixtures)) {
    const p = planSlides(fixtures[name]);
    for (const s of p.slides) {
      if (!s.replyBox) continue;
      const { box, labelBox } = s.replyBox;
      assert.ok(labelBox.h >= box.h - C.REPLY_INSET * 2 - 0.01,
        `${name} p${s.page}: reply text box is ${labelBox.h.toFixed(2)}in inside a ${box.h.toFixed(2)}in area`);
      assert.ok(labelBox.y + labelBox.h <= box.y + box.h + 0.01, 'and stays within it');
    }
  }
});

test('48 · no image anywhere renders too small to read, unless it is alone', () => {
  // The one exception: a single image always goes on its slide at whatever size
  // it comes out, because there is nothing left to reduce.
  for (const name of Object.keys(fixtures)) {
    const p = planSlides(fixtures[name]);
    for (const s of p.slides.filter((x) => x.kind === 'item')) {
      if (s.images.length < 2) continue;
      for (const im of s.images) {
        assert.ok(im.box.w >= C.MIN_IMAGE_W - 0.01,
          `${name} p${s.page}: ${s.images.length} images, one rendering ${im.box.w.toFixed(2)}in wide`);
      }
    }
  }
});
