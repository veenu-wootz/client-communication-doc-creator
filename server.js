/**
 * server.js — Strike/Glide webhook in, formatted deck out.
 *
 *   POST /generate   parse → enrich → fetch images → plan → render → deliver
 *   GET  /           health
 *   POST /preview    same pipeline, returns the .pptx directly (no delivery)
 *
 * Responds to the webhook immediately and finishes the work in the background,
 * the way qualityinspectionreport/server.js does — Strike and Glide both time
 * out and retry otherwise, which would generate the deck twice.
 */

require('dotenv').config();

const path = require('path');
const express = require('express');

const { parseStrikePayload } = require('./src/input/parseStrikePayload');
const { enrich } = require('./src/input/enrich');
const { prepareImages } = require('./src/media/prepareImages');
const { planSlides } = require('./src/plan/planner');
const { renderPptx } = require('./src/render/renderPptx');
const { sendDeckEmail } = require('./src/deliver/sendEmail');
const { uploadDeck } = require('./src/deliver/upload');
const { uploadToOneDrive } = require('./src/deliver/oneDriveUpload');
const { writeBackToGlide } = require('./src/deliver/glideWrite');

const app = express();
const PORT = process.env.PORT || 3000;

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// Local test form — a single static HTML file, no build step. Open
// http://localhost:PORT/test.html and it fills the same JSON shape Strike/Glide
// would send, calling /preview (download) or /generate (email) on this server.
app.use(express.static(path.join(__dirname, 'public')));

app.use(express.json({ limit: '50mb' }));

const safe = (s) => String(s || 'document').replace(/[^a-zA-Z0-9\-_. ]/g, '_').trim();

/**
 * "05 Sep 2026 2.48 PM" in IST, for the filename.
 *
 * Generation time, not the submission's created_at — the filename records when
 * the deck was produced. A dot separates the time because a colon is illegal
 * in OneDrive/SharePoint filenames.
 */
function istStamp(d = new Date()) {
  const parts = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Kolkata',
    day: '2-digit', month: 'numeric', year: 'numeric',
    hour: 'numeric', minute: '2-digit', hour12: true,
  }).formatToParts(d).reduce((a, p) => ({ ...a, [p.type]: p.value }), {});

  // Same three-letter months the deck's own footer uses, so the filename and
  // the slides never disagree about how a date is written.
  const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const month = MONTHS[Number(parts.month) - 1];

  return `${parts.day} ${month} ${parts.year} ${parts.hour}.${parts.minute} ${parts.dayPeriod.toUpperCase()}`;
}

/** Everything from raw webhook body to a rendered deck. Shared by both routes. */
async function build(body) {
  // One instant for the whole run, so the filename, the Glide row and anything
  // else that records "when" cannot disagree by a second.
  const generatedAt = new Date();
  const parsed = parseStrikePayload(body);
  console.log(`  parsed: ${parsed.items.length} rows, report "${parsed.document.report_title || '(none)'}"`);

  const enriched = await enrich(parsed.items);
  const prepared = await prepareImages({ document: parsed.document, items: enriched.items });

  const plan = planSlides(prepared.document);
  plan.meta.warnings.push(...(parsed.warnings || []), ...enriched.warnings, ...prepared.warnings);

  const buffer = await renderPptx(plan);
  const project = parsed.document.project_name || parsed.document.report_title;
  const filename = `Queries - ${safe(project)} - ${istStamp(generatedAt)}.pptx`;

  return { parsed, plan, buffer, filename, generatedAt };
}

app.get('/', (req, res) => {
  res.json({ status: 'ok', service: 'Client Communication Doc Creator', version: '1.0.0' });
});

app.post('/generate', async (req, res) => {
  const started = Date.now();

  // Acknowledge before doing the work, so the no-code platform does not retry.
  res.json({ success: true, message: 'Received, generating' });

  try {
    console.log('\n━━━━━━ /generate ━━━━━━');
    const { parsed, plan, buffer, filename, generatedAt } = await build(req.body);
    console.log(`  planned ${plan.meta.totalSlides} slides — ${plan.meta.itemCount} points`);
    console.log(`  rendered ${(buffer.length / 1024).toFixed(0)} KB`);

    // Optional, in order. Each one skipping is normal, not an error.
    // OneDrive first — it's the configured store and puts the deck where the
    // team already works. S3 stays as a fallback for whoever configures it.
    let upload = await uploadToOneDrive(buffer, filename, parsed.storage);
    if (!upload.url) upload = await uploadDeck(buffer, filename);

    try {
      await writeBackToGlide(parsed.writeback, {
        fileId: upload.itemId,
        fileLink: upload.url,
        // ISO for Glide's date column — a display string like "05 Sep 2026
        // 7.10 PM" would not sort or filter there. The readable form is in the
        // filename, where a person actually reads it.
        generatedOn: generatedAt.toISOString(),
        // No fallback: these columns record what this run actually had. Falling
        // back to created_by would put a person's NAME in a column meant for an
        // email address, which reads as data rather than as a gap.
        generatedBy: parsed.delivery.to,
      });
    } catch (e) {
      console.warn(`  glide write failed (non-fatal): ${e.message}`);
    }

    // The guaranteed delivery path — always last, always attempted.
    await sendDeckEmail(plan, buffer, filename, parsed.delivery, upload.url);

    for (const w of plan.meta.warnings) console.log(`  ⚠ ${w}`);
    console.log(`✓ done in ${((Date.now() - started) / 1000).toFixed(1)}s — ${filename}\n`);
  } catch (err) {
    // The response already went out above; never try to send a second one.
    console.error('✗ generate failed:', err);
  }
});

/** Same pipeline, deck returned inline. For testing without SMTP or S3. */
app.post('/preview', async (req, res) => {
  try {
    const { plan, buffer, filename } = await build(req.body);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.setHeader('X-Slide-Count', String(plan.meta.totalSlides));
    res.setHeader('X-Item-Count', String(plan.meta.itemCount));
    res.send(buffer);
  } catch (err) {
    console.error('✗ preview failed:', err);
    if (!res.headersSent) res.status(500).json({ error: err.message });
  }
});

// Only listen when run directly, so tests can import build() without a port.
if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`\nClient Communication Doc Creator on port ${PORT}`);
    const note = (ok, label) => console.log(`  ${ok ? '✓' : '·'} ${label}${ok ? '' : ' (will be skipped)'}`);
    note(Boolean(process.env.SMTP_USER && process.env.SMTP_PASSWORD), 'email');
    note(require('./src/deliver/upload').isConfigured(), 'S3 upload');
    note(Boolean(process.env.GLIDE_TOKEN), 'Glide write-back');
    note(Boolean(process.env.ANTHROPIC_API_KEY || process.env.OPENAI_API_KEY), 'LLM enrichment');
    console.log('');
  });
}

module.exports = { app, build, istStamp };
