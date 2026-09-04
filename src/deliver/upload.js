/**
 * upload.js — put the deck on S3 and return a permanent URL.
 *
 * OPTIONAL BY DESIGN (PLAN.md T6). Infra may not be available, so a missing
 * credential is a skip with a log line, never an error. The email attachment is
 * the guaranteed delivery path; this is an enhancement on top of it.
 *
 * Adapted from qualityinspectionreport/awsUpload.js, with the content type
 * generalised so it can carry a .pptx as happily as a .pdf.
 */

const { S3Client, PutObjectCommand } = require('@aws-sdk/client-s3');

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

const REQUIRED = ['AWS_REGION', 'AWS_S3_BUCKET_NAME', 'AWS_ACCESS_KEY_ID', 'AWS_SECRET_ACCESS_KEY'];

/** All four env vars, or we do not attempt it. */
function isConfigured() {
  return REQUIRED.every((k) => Boolean(process.env[k]));
}

function missingVars() {
  return REQUIRED.filter((k) => !process.env[k]);
}

let client;
function getClient() {
  if (!client) {
    client = new S3Client({
      region: process.env.AWS_REGION,
      credentials: {
        accessKeyId: process.env.AWS_ACCESS_KEY_ID,
        secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
      },
    });
  }
  return client;
}

/**
 * @returns {Promise<{url: string|null, skipped: boolean, reason?: string}>}
 */
async function uploadDeck(buffer, filename, contentType = PPTX_MIME) {
  if (!isConfigured()) {
    const reason = `S3 not configured (${missingVars().join(', ')})`;
    console.log(`  upload: skipped — ${reason}`);
    return { url: null, skipped: true, reason };
  }

  const bucket = process.env.AWS_S3_BUCKET_NAME;
  const region = process.env.AWS_REGION;
  const key = filename.replace(/[^a-zA-Z0-9\-_.]/g, '_');

  try {
    await getClient().send(new PutObjectCommand({
      Bucket: bucket, Key: key, Body: buffer, ContentType: contentType,
    }));
    const url = `https://${bucket}.s3.${region}.amazonaws.com/${encodeURIComponent(key)}`;
    console.log(`  upload → ${url}`);
    return { url, skipped: false };
  } catch (e) {
    // Non-fatal: the deck is already attached to the email.
    console.warn(`  upload failed (non-fatal): ${e.message}`);
    return { url: null, skipped: true, reason: e.message };
  }
}

module.exports = { uploadDeck, isConfigured, missingVars, PPTX_MIME };
