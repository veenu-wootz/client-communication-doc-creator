/**
 * sendTestEmail.js — a minimal, real send through the exact production path.
 *
 * Deliberately bypasses Strike parsing and AI enrichment — this exists to
 * isolate ONE thing: how the SMTP handoff itself behaves (timing, auth,
 * acceptance), not the rest of the pipeline. Uses planSlides → renderPptx →
 * sendDeckEmail exactly as server.js does for /generate.
 *
 * Reads credentials from .env — never prints SMTP_PASSWORD or logs it in any
 * form. Run with: node scripts/sendTestEmail.js
 */

require('dotenv').config();

const { planSlides } = require('../src/plan/planner');
const { renderPptx } = require('../src/render/renderPptx');
const { sendDeckEmail } = require('../src/deliver/sendEmail');

const doc = {
  document: {
    report_title: 'SMTP timing test',
    reference_name: 'local-test',
    created_by: 'Local test script',
    created_at: new Date().toISOString().slice(0, 10),
    logo: null,          // skip the network fetch — isolate the SMTP step only
    product_photo: null,
  },
  items: [
    { id: 'a', title: 'Delivery check', body: 'If this landed, note the time.', needs_response: true, images: [] },
    { id: 'b', title: 'Sent at', body: new Date().toISOString(), needs_response: false, images: [] },
  ],
};

async function main() {
  const to = process.env.TEST_RECIPIENT || 'vinay.jadon@wootz.work';
  console.log(`SMTP_SERVICE=${process.env.SMTP_SERVICE || '(unset)'}  SMTP_USER=${process.env.SMTP_USER || '(unset)'}  SMTP_FROM_NAME=${process.env.SMTP_FROM_NAME || '(default: Wootz)'}`);
  console.log(`sending to ${to}\n`);

  const plan = planSlides(doc);
  const buffer = await renderPptx(plan);
  console.log(`rendered ${(buffer.length / 1024).toFixed(0)} KB\n`);

  const t0 = Date.now();
  const result = await sendDeckEmail(plan, buffer, 'smtp-test.pptx', { to }, null);
  const elapsed = ((Date.now() - t0) / 1000).toFixed(2);

  if (result.skipped) {
    console.log(`✗ skipped: ${result.reason}`);
  } else {
    console.log(`✓ SMTP accepted in ${elapsed}s — messageId: ${result.messageId}`);
    console.log(`  now check the inbox and note the actual arrival time.`);
  }
}

main().catch((e) => {
  // Nodemailer error objects never include the password, only auth failure
  // codes/reasons — safe to print in full.
  console.error('✗ failed:', e.message);
  process.exit(1);
});
