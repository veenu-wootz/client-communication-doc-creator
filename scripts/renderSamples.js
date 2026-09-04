/**
 * renderSamples.js — render every fixture to out/ for a real look.
 *
 * Tier 3 of the test plan, and deliberately manual: open one in PowerPoint,
 * Save-as-PDF, and check fidelity by eye. No rendering engine can honestly tell
 * you what PowerPoint will do except PowerPoint.
 *
 *   npm run sample
 *
 * Also writes each fixture as JSON, so you can POST one at the server:
 *   curl -X POST localhost:3000/generate -H 'Content-Type: application/json' \
 *        -d @out/01-flagged-unflagged-flagged.json
 */

const fs = require('fs');
const path = require('path');
const sharp = require('sharp');

const { planSlides } = require('../src/plan/planner');
const { renderPptx } = require('../src/render/renderPptx');
const { fixtures } = require('../test/fixtures');

const OUT = path.join(__dirname, '..', 'out');
const LOGO_URL = process.env.LOGO_URL
  || 'https://res.cloudinary.com/dbwg6zz3l/image/upload/w_300,f_png,q_90/v1773643264/Black_Yellow_kq9kef.png';

/** A drawing-ish placeholder so the layouts can be judged at real proportions. */
async function drawingPng(w, h, label) {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}">
    <rect width="100%" height="100%" fill="#ffffff"/>
    <rect x="6" y="6" width="${w - 12}" height="${h - 12}" fill="none" stroke="#9aa3ad" stroke-width="3"/>
    <line x1="${w * 0.15}" y1="${h * 0.25}" x2="${w * 0.85}" y2="${h * 0.25}" stroke="#374151" stroke-width="4"/>
    <line x1="${w * 0.15}" y1="${h * 0.25}" x2="${w * 0.15}" y2="${h * 0.75}" stroke="#374151" stroke-width="4"/>
    <circle cx="${w * 0.62}" cy="${h * 0.58}" r="${Math.min(w, h) * 0.16}" fill="none" stroke="#374151" stroke-width="4"/>
    <text x="${w / 2}" y="${h * 0.92}" font-family="Arial" font-size="${Math.max(14, Math.min(w, h) * 0.09)}"
          fill="#6B7280" text-anchor="middle">${label}</text>
  </svg>`;
  return sharp(Buffer.from(svg)).png().toBuffer();
}

const toDataUri = (buf) => `image/png;base64,${buf.toString('base64')}`;

async function fetchLogo() {
  try {
    const fetch = require('node-fetch');
    const res = await fetch(LOGO_URL, { timeout: 10000 });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buf = Buffer.from(await res.arrayBuffer());
    const meta = await sharp(buf).metadata();
    return { data: toDataUri(buf), width: meta.width, height: meta.height,
             aspect: meta.width / meta.height, ok: true, path: 'logo.png' };
  } catch (e) {
    console.warn(`  logo unavailable (${e.message}) — rendering without it`);
    return null;
  }
}

/** Give every image in a fixture real pixels, so the deck looks like a deck. */
async function materialise(fixture, logo) {
  const doc = { ...fixture.document, logo: logo || null };

  if (doc.product_photo && doc.product_photo.ok !== false) {
    const p = doc.product_photo;
    doc.product_photo = { ...p, data: toDataUri(await drawingPng(p.width, p.height, 'product')) };
  }

  const items = [];
  for (const item of fixture.items) {
    const images = [];
    for (const im of item.images || []) {
      if (im.ok === false) { images.push(im); continue; }   // keep the broken-path case broken
      images.push({ ...im, data: toDataUri(await drawingPng(im.width, im.height, path.basename(im.path, '.png'))) });
    }
    items.push({ ...item, images });
  }

  return { document: doc, items };
}

async function main() {
  fs.mkdirSync(OUT, { recursive: true });
  console.log('Fetching logo...');
  const logo = await fetchLogo();

  const names = Object.keys(fixtures).sort();
  let totalSlides = 0;

  for (const name of names) {
    const doc = await materialise(fixtures[name], logo);
    const plan = planSlides(doc);
    const buf = await renderPptx(plan);

    fs.writeFileSync(path.join(OUT, `${name}.pptx`), buf);
    fs.writeFileSync(path.join(OUT, `${name}.json`), JSON.stringify(doc, null, 2));

    totalSlides += plan.meta.totalSlides;
    const warn = plan.meta.warnings.length ? `  ⚠ ${plan.meta.warnings.length}` : '';
    console.log(
      `  ${name.padEnd(32)} ${String(plan.meta.totalSlides).padStart(2)} slides` +
      `  ${String(plan.meta.queryCount).padStart(2)}Q ${String(plan.meta.updateCount).padStart(2)}U` +
      `  ${(buf.length / 1024).toFixed(0).padStart(4)} KB${warn}`);
    for (const w of plan.meta.warnings) console.log(`      ${w}`);
  }

  console.log(`\n${names.length} decks, ${totalSlides} slides → ${OUT}`);
}

main().catch((e) => { console.error(e); process.exit(1); });
