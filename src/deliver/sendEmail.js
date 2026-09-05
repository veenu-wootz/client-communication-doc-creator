/**
 * sendEmail.js — send the finished deck to the person who created it.
 *
 * Adapted from qualityinspectionreport/sendEmail.js: same lazy transporter and
 * the same office365/gmail switch, which are known to work against the Wootz
 * mailbox.
 *
 * It goes to the SENDER, not the client. They forward it from their own
 * mailbox, so the client's reply lands in their inbox in their own thread — and
 * so no automated send can ever reach a customer with the wrong addressee typed
 * into a form (PLAN.md P6).
 */

const nodemailer = require('nodemailer');

let transporter;

function getTransporter() {
  if (!transporter) {
    const service = (process.env.SMTP_SERVICE || 'office365').toLowerCase();
    if (service === 'gmail') {
      transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASSWORD },
      });
    } else {
      transporter = nodemailer.createTransport({
        host: 'smtp.office365.com', port: 587, secure: false,
        auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASSWORD },
        tls: { ciphers: 'SSLv3' },
      });
    }
  }
  return transporter;
}

const isConfigured = () => Boolean(process.env.SMTP_USER && process.env.SMTP_PASSWORD);

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

function buildHtml(plan, fileUrl, fileNote = null) {
  const d = plan.document;
  // Every item, in order — the deck no longer separates queries from updates.
  const points = plan.slides
    .filter((s) => s.kind === 'item' && !s.continued)
    .map((s) => `<li style="margin:4px 0;"><b>${s.number}</b> &middot; ${esc(s.title)}</li>`)
    .join('');

  const count = plan.meta.itemCount || 0;
  const owed = count === 0
    ? 'This document has no points in it — check the submission.'
    : count === 1
      ? 'It puts <b>1 point</b> to the client.'
      : `It puts <b>${count} points</b> to the client.`;

  const row = (label, value) => value
    ? `<tr><td style="padding:6px 0;color:#6B7280;width:130px;">${label}</td><td style="padding:6px 0;">${esc(value)}</td></tr>`
    : '';

  return `
    <div style="font-family:Arial,sans-serif;max-width:620px;margin:0 auto;">
      <div style="background:#0F4C5C;padding:20px 24px;border-radius:8px 8px 0 0;">
        <h2 style="color:#fff;margin:0;font-size:18px;">PPT ready to send</h2>
        <p style="color:#cfe0e4;margin:4px 0 0;font-size:13px;">${esc(d.report_title || 'Untitled report')}</p>
      </div>
      <div style="background:#F5F6F8;padding:22px 24px;border:1px solid #D9DCE1;border-top:none;">
        <table style="width:100%;border-collapse:collapse;font-size:13px;">
          ${row('Reference', d.reference_name)}
          ${row('Attention', d.addressee)}
          ${row('Prepared by', d.created_by)}
          <tr><td style="padding:6px 0;color:#6B7280;">Contents</td><td style="padding:6px 0;">
            ${count} ${count === 1 ? 'point' : 'points'}, ${plan.meta.totalSlides} slides</td></tr>
        </table>
      </div>
      <div style="background:#fff;padding:18px 24px;border:1px solid #D9DCE1;border-top:none;${fileUrl ? '' : 'border-radius:0 0 8px 8px;'}">
        <p style="margin:0 0 10px;font-size:13px;color:#111827;">
          The deck is attached. ${owed} Forward it to your client from your own mailbox so their
          reply comes back to you.
        </p>
        ${points ? `<ul style="margin:10px 0 0;padding-left:20px;font-size:13px;color:#111827;">${points}</ul>` : ''}
      </div>
      ${fileUrl ? `<div style="background:#fff;padding:0 24px 18px;border:1px solid #D9DCE1;border-top:none;border-radius:0 0 8px 8px;">
        <p style="margin:0;font-size:12px;color:#6B7280;">Also stored at <a href="${esc(fileUrl)}" style="color:#0F4C5C;">this link</a>.</p>
      </div>` : ''}
      ${!fileUrl && fileNote ? `<div style="background:#FFF6E5;padding:12px 24px;border:1px solid #E8D9B5;border-top:none;border-radius:0 0 8px 8px;">
        <p style="margin:0;font-size:12px;color:#7A5B00;"><b>Not filed to OneDrive.</b> ${esc(fileNote)} The deck is attached, so nothing is lost \u2014 but it was not saved to the project folder.</p>
      </div>` : ''}
    </div>`;
}

/**
 * @param {object} plan       the slide plan (for the summary)
 * @param {Buffer} pptxBuffer the deck
 * @param {string} filename
 * @param {object} delivery   { to, cc, bcc }
 * @param {string|null} fileUrl
 */
async function sendDeckEmail(plan, pptxBuffer, filename, delivery, fileUrl = null, fileNote = null) {
  if (!isConfigured()) {
    console.warn('  email: SMTP_USER/SMTP_PASSWORD not set — skipping send');
    return { skipped: true, reason: 'smtp not configured' };
  }

  const to = (delivery && delivery.to) || process.env.FALLBACK_EMAIL || '';
  const list = (v) => String(v || '').split(',').map((s) => s.trim()).filter(Boolean).join(', ');
  const cc = list(delivery && delivery.cc);
  const bcc = list([delivery && delivery.bcc, process.env.INTERNAL_EMAILS].filter(Boolean).join(','));

  if (!to && !bcc) {
    console.warn('  email: no recipient on the payload and no FALLBACK_EMAIL — skipping send');
    return { skipped: true, reason: 'no recipient' };
  }

  const d = plan.document;
  const subject = `PPT — ${d.report_title || 'Untitled'}`
    + (d.reference_name ? ` | ${d.reference_name}` : '')
    + (plan.meta.itemCount ? ` | ${plan.meta.itemCount} ${plan.meta.itemCount === 1 ? 'point' : 'points'}` : '');

  const info = await getTransporter().sendMail({
    from: `"Wootz" <${process.env.SMTP_USER}>`,
    to: to || undefined,
    cc: cc || undefined,
    bcc: bcc || undefined,
    subject,
    text: `${d.report_title || 'Untitled'} — ${plan.meta.itemCount} points.\n`
        + `The deck is attached. Forward it to your client from your own mailbox so their reply comes back to you.`,
    html: buildHtml(plan, fileUrl, fileNote),
    attachments: [{
      filename,
      content: pptxBuffer,
      contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    }],
  });

  console.log(`  email → to:${to}${cc ? ` cc:${cc}` : ''}${bcc ? ` bcc:${bcc}` : ''} (${info.messageId})`);
  return { skipped: false, messageId: info.messageId };
}

module.exports = { sendDeckEmail, isConfigured, buildHtml };
