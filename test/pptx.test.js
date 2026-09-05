/**
 * pptx.test.js — Tier 2. Assertions against the real OOXML inside the .pptx.
 *
 * A .pptx is a zip of XML. Reading the slide XML back is the only way to prove
 * what actually landed on the slide, as opposed to what the planner intended.
 */

const test = require('node:test');
const assert = require('node:assert');
const JSZip = require('jszip');
const { XMLParser } = require('fast-xml-parser');

const { planSlides } = require('../src/plan/planner');
const { renderPptx } = require('../src/render/renderPptx');
const C = require('../src/plan/constants');
const { fixtures } = require('./fixtures');

const EMU = 914400;
const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@' });

/** Open a rendered deck and return its slide XML, in order. */
async function openDeck(fixtureName) {
  const plan = planSlides(fixtures[fixtureName]);
  const buf = await renderPptx(plan);
  const zip = await JSZip.loadAsync(buf);

  const names = Object.keys(zip.files)
    .filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => Number(a.match(/(\d+)/)[1]) - Number(b.match(/(\d+)/)[1]));

  const slides = [];
  for (const n of names) slides.push(await zip.file(n).async('string'));

  const presentation = await zip.file('ppt/presentation.xml').async('string');
  return { plan, zip, slides, presentation, buf };
}

/** Walk a parsed node, collecting every value stored under `key`. */
function collect(node, key, out = []) {
  if (node === null || typeof node !== 'object') return out;
  for (const [k, v] of Object.entries(node)) {
    if (k === key) (Array.isArray(v) ? v : [v]).forEach((x) => out.push(x));
    if (typeof v === 'object') collect(v, key, out);
  }
  return out;
}

/** Every positioned shape on a slide, in inches. */
function shapesOf(slideXml) {
  const doc = parser.parse(slideXml);
  const boxes = [];
  for (const xfrm of collect(doc, 'a:xfrm')) {
    const off = xfrm['a:off'];
    const ext = xfrm['a:ext'];
    if (!off || !ext) continue;
    boxes.push({
      x: Number(off['@x']) / EMU, y: Number(off['@y']) / EMU,
      w: Number(ext['@cx']) / EMU, h: Number(ext['@cy']) / EMU,
    });
  }
  return boxes;
}

/** All text on a slide, joined. */
function textOf(slideXml) {
  const doc = parser.parse(slideXml);
  return collect(doc, 'a:t').map((t) => (typeof t === 'object' ? t['#text'] ?? '' : String(t))).join(' ');
}

function geometriesOf(slideXml) {
  const doc = parser.parse(slideXml);
  return collect(doc, 'a:prstGeom').map((g) => g['@prst']);
}

// ── The canvas ───────────────────────────────────────────────

test('the presentation canvas is 13.333 x 7.5in, not the 10 x 5.625in default', async () => {
  const { presentation } = await openDeck('01-flagged-unflagged-flagged');
  const doc = parser.parse(presentation);
  const size = collect(doc, 'p:sldSz')[0];
  assert.ok(Math.abs(Number(size['@cx']) / EMU - C.SLIDE_W) < 0.01, `width ${Number(size['@cx']) / EMU}`);
  assert.ok(Math.abs(Number(size['@cy']) / EMU - C.SLIDE_H) < 0.01, `height ${Number(size['@cy']) / EMU}`);
});

test('slide count matches the plan for every fixture', async () => {
  for (const name of Object.keys(fixtures)) {
    const { plan, slides } = await openDeck(name);
    assert.strictEqual(slides.length, plan.meta.totalSlides, `${name}: slide count`);
  }
});

// ── Geometry ─────────────────────────────────────────────────

test('every rendered shape sits inside the canvas', async () => {
  const EPS = 0.02;
  for (const name of Object.keys(fixtures)) {
    const { slides } = await openDeck(name);
    slides.forEach((xml, i) => {
      for (const b of shapesOf(xml)) {
        assert.ok(b.x >= -EPS && b.y >= -EPS, `${name} slide ${i + 1}: shape at ${b.x.toFixed(2)},${b.y.toFixed(2)}`);
        assert.ok(b.x + b.w <= C.SLIDE_W + EPS, `${name} slide ${i + 1}: right edge ${(b.x + b.w).toFixed(2)} off canvas`);
        assert.ok(b.y + b.h <= C.SLIDE_H + EPS, `${name} slide ${i + 1}: bottom ${(b.y + b.h).toFixed(2)} off canvas`);
      }
    });
  }
});

test('shapes respect the margins, with the footer band and cover logo excepted', async () => {
  const EPS = 0.02;
  for (const name of Object.keys(fixtures)) {
    const { slides } = await openDeck(name);
    slides.forEach((xml, i) => {
      for (const b of shapesOf(xml)) {
        const isFooterBand = b.y >= C.FOOTER_Y - 0.05;
        const isTopLogo = b.y < C.MARGIN;          // the cover logo, by design
        if (isFooterBand || isTopLogo) continue;
        assert.ok(b.x >= C.MARGIN - EPS, `${name} slide ${i + 1}: x=${b.x.toFixed(2)} inside left margin`);
        assert.ok(b.x + b.w <= C.SLIDE_W - C.MARGIN + EPS,
          `${name} slide ${i + 1}: right edge ${(b.x + b.w).toFixed(2)} inside right margin`);
      }
    });
  }
});

// ── Numbering and the reply box ──────────────────────────────

test('exactly one reply box per item, on its last slide, correctly numbered', async () => {
  for (const name of Object.keys(fixtures)) {
    const { plan, slides } = await openDeck(name);
    const byItem = new Map();
    plan.slides.forEach((s, i) => {
      if (s.kind !== 'item') return;
      if (!byItem.has(s.number)) byItem.set(s.number, []);
      byItem.get(s.number).push(i);
    });

    assert.ok(byItem.size > 0 || fixtures[name].items.length === 0,
      `${name}: expected numbered items — this assertion must not pass vacuously`);

    for (const [num, idxs] of byItem) {
      const hits = idxs.filter((i) => textOf(slides[i]).includes(`Your response — ${num}`));
      assert.strictEqual(hits.length, 1, `${name}: item ${num} should have exactly one reply box`);
      assert.strictEqual(hits[0], idxs[idxs.length - 1], `${name}: item ${num} reply box is not on the last slide`);
      assert.ok(geometriesOf(slides[hits[0]]).includes('roundRect'), `${name}: reply box must be a roundRect`);
    }
  }
});

test('the Q prefix is gone from every deck', async () => {
  for (const name of Object.keys(fixtures)) {
    const { slides } = await openDeck(name);
    slides.forEach((xml, i) => {
      assert.ok(!/\bQ\d+\b/.test(textOf(xml)),
        `${name} slide ${i + 1}: a Q-number survived the flat renumbering`);
    });
  }
});

test('every slide of a multi-slide item repeats its number', async () => {
  const { plan, slides } = await openDeck('07-long-body-4-images');
  let checked = 0;
  plan.slides.forEach((s, i) => {
    if (s.kind !== 'item') return;
    assert.ok(textOf(slides[i]).includes(String(s.number)), `slide ${i + 1} missing its number`);
    checked += 1;
  });
  assert.ok(checked > 1, 'expected a multi-slide item to actually be checked');
});

test('the summary renders as one numbered list, not a box per entry', async () => {
  const { plan, slides, zip } = await openDeck('12-four-queries-two-updates');
  const i = plan.slides.findIndex((s) => s.kind === 'contents');
  assert.ok(i >= 0);

  // buAutoNum is PowerPoint's own numbering; its presence is what makes Enter
  // continue the list for whoever edits the deck.
  const xml = await zip.file(`ppt/slides/slide${i + 1}.xml`).async('string');
  assert.ok(xml.includes('buAutoNum'), 'summary must use real auto-numbering');

  const text = textOf(slides[i]);
  for (const title of ['Question 1', 'Fixture ready', 'Timeline']) {
    assert.ok(text.includes(title), `summary missing "${title}"`);
  }
});

// ── Content preservation ─────────────────────────────────────

test('no ellipsis is ever written into a deck', async () => {
  for (const name of Object.keys(fixtures)) {
    const { slides } = await openDeck(name);
    slides.forEach((xml, i) => {
      const t = textOf(xml);
      assert.ok(!t.includes('…'), `${name} slide ${i + 1}: ellipsis in "${t.slice(0, 80)}"`);
      assert.ok(!t.includes('...'), `${name} slide ${i + 1}: ellipsis in "${t.slice(0, 80)}"`);
    });
  }
});

test('the renderer emits every body line the planner placed', async () => {
  // Paired with planner.test.js §7, which proves the split across slides is
  // lossless. Together: the planner loses nothing, and the renderer draws all
  // of what the planner produced.
  for (const name of Object.keys(fixtures)) {
    const { plan, slides } = await openDeck(name);
    plan.slides.forEach((s, i) => {
      if (s.kind !== 'item' || !s.body) return;
      const rendered = textOf(slides[i]).replace(/\s+/g, ' ');
      for (const line of s.body.lines) {
        if (!line.trim()) continue;
        assert.ok(rendered.includes(line.trim()),
          `${name} slide ${i + 1}: body line missing from the deck — "${line.slice(0, 40)}"`);
      }
    });
  }
});

test('footers carry the page number and no total', async () => {
  const { plan, slides } = await openDeck('11-four-queries');
  plan.slides.forEach((s, i) => {
    if (s.kind !== 'item' && s.kind !== 'template') return;
    const t = textOf(slides[i]);
    assert.ok(t.includes(`Page ${s.page}`), `slide ${i + 1} missing "Page ${s.page}"`);
    assert.ok(!/Page \d+ of/.test(t), 'footer must not carry a total');
  });
});

// ── Determinism ──────────────────────────────────────────────

test('rendering the same input twice produces identical slides', async () => {
  for (const name of ['01-flagged-unflagged-flagged', '07-long-body-4-images', '12-four-queries-two-updates']) {
    const a = await openDeck(name);
    const b = await openDeck(name);
    assert.deepStrictEqual(a.slides, b.slides, `${name}: slide XML differs between runs`);
  }
});

test('the deck is a valid zip carrying the parts PowerPoint requires', async () => {
  const { zip } = await openDeck('01-flagged-unflagged-flagged');
  for (const part of ['[Content_Types].xml', 'ppt/presentation.xml', '_rels/.rels', 'ppt/slides/slide1.xml']) {
    assert.ok(zip.file(part), `missing required part ${part}`);
  }
  for (const name of Object.keys(zip.files)) {
    if (!name.endsWith('.xml')) continue;
    const xml = await zip.file(name).async('string');
    assert.doesNotThrow(() => parser.parse(xml), `${name} is not well-formed XML`);
  }
});

test('every picture placeholder is a real one, so it can be clicked and vanishes if unused', async () => {
  // pptxgenjs cannot emit type="pic" (its PLACEHOLDER_TYPES table is empty), so
  // the attribute is injected after write. If a library upgrade breaks that,
  // this test is what catches it — otherwise the deck would quietly ship
  // placeholders nobody can click.
  for (const name of ['01-flagged-unflagged-flagged', '24-no-body-no-images', '21-cover-with-photo']) {
    const { plan, zip } = await openDeck(name);

    const expected = new Set(plan.slides.map((s) => s.masterName).filter(Boolean));
    assert.ok(expected.size > 0, `${name}: expected at least one picture placeholder`);

    let typed = 0;
    for (const f of Object.keys(zip.files).filter((n) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n))) {
      const xml = await zip.file(f).async('string');
      if (!/name="PIC_[^"]*"/.test(xml)) continue;
      assert.match(xml, /<p:ph[^>]*type="pic"/, `${name}: ${f} is not a picture placeholder`);
      typed += 1;
    }
    assert.strictEqual(typed, expected.size, `${name}: every needed layout must be promoted`);
  }
});

test('a slide with a real image gets no placeholder layout', async () => {
  const { plan } = await openDeck('03-one-upright-image');
  const item = plan.slides.find((s) => s.kind === 'item');
  assert.strictEqual(item.masterName, undefined, 'an image is present, so nothing is offered');
});

test('the summary numbers every paragraph, with none left unbulleted', async () => {
  const { plan, zip } = await openDeck('12-four-queries-two-updates');
  const i = plan.slides.findIndex((s) => s.kind === 'contents');
  const xml = await zip.file(`ppt/slides/slide${i + 1}.xml`).async('string');

  const starts = [...xml.matchAll(/buAutoNum[^>]*startAt="(\d+)"/g)].map((m) => Number(m[1]));
  assert.deepStrictEqual(starts, [1, 2, 3, 4, 5, 6],
    'each entry states its own number — pptxgenjs always writes startAt, so omitting it renders 1,1,1');
});
