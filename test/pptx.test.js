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

test('exactly one reply box per query, on its last slide, correctly numbered', async () => {
  for (const name of Object.keys(fixtures)) {
    const { plan, slides } = await openDeck(name);
    const flagged = new Map();
    plan.slides.forEach((s, i) => {
      if (s.kind === 'item' && s.queryNo) {
        if (!flagged.has(s.queryNo)) flagged.set(s.queryNo, []);
        flagged.get(s.queryNo).push(i);
      }
    });

    for (const [qNo, idxs] of flagged) {
      const hits = idxs.filter((i) => textOf(slides[i]).includes(`Your response — Q${qNo}`));
      assert.strictEqual(hits.length, 1, `${name}: Q${qNo} should have exactly one reply box`);
      assert.strictEqual(hits[0], idxs[idxs.length - 1], `${name}: Q${qNo} reply box is not on the last slide`);
      assert.ok(geometriesOf(slides[hits[0]]).includes('roundRect'), `${name}: reply box must be a roundRect`);
    }
  }
});

test('no query number appears anywhere on an unflagged item slide', async () => {
  for (const name of Object.keys(fixtures)) {
    const { plan, slides } = await openDeck(name);
    plan.slides.forEach((s, i) => {
      if (s.kind !== 'item' || s.queryNo) return;
      assert.ok(!/\bQ\d+\b/.test(textOf(slides[i])), `${name} slide ${i + 1}: update slide carries a query number`);
    });
  }
});

test('every query slide repeats its number in the header chip', async () => {
  const { plan, slides } = await openDeck('07-long-body-4-images');
  plan.slides.forEach((s, i) => {
    if (s.kind === 'item' && s.queryNo) {
      assert.ok(textOf(slides[i]).includes(`Q${s.queryNo}`), `slide ${i + 1} missing its Q number`);
    }
  });
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
    if (s.kind !== 'item' && s.kind !== 'grouped') return;
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
