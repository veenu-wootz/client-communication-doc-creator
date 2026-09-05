/**
 * checkGraphAccess.js — does this app registration actually have what we need?
 *
 * Fetches a token via client credentials and reads back the permissions the
 * token actually carries, rather than trusting an assumption about what the
 * app was granted. Then, if a drive is configured, proves read access to the
 * target folder.
 *
 * Never prints the client secret or the access token itself — only the
 * non-sensitive claims that answer the question.
 *
 *   node scripts/checkGraphAccess.js
 */

require('dotenv').config();

const TENANT = process.env.GRAPH_TENANT_ID;
const CLIENT = process.env.GRAPH_CLIENT_ID;
const SECRET = process.env.GRAPH_CLIENT_SECRET;
const DRIVE = process.env.GRAPH_DRIVE_ID;
const FOLDER = process.env.GRAPH_FOLDER_ITEM_ID;

const NEEDED = 'Files.ReadWrite.All';

/** Read the claims out of a JWT without verifying it — we only want to look. */
function claims(token) {
  const part = token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
  return JSON.parse(Buffer.from(part, 'base64').toString('utf8'));
}

async function getToken() {
  const res = await fetch(`https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: CLIENT,
      client_secret: SECRET,
      grant_type: 'client_credentials',
      scope: 'https://graph.microsoft.com/.default',
    }),
  });

  const body = await res.json();
  if (!res.ok) {
    // Microsoft's error descriptions are informative and carry no secret.
    throw new Error(`${body.error}: ${(body.error_description || '').split('\n')[0]}`);
  }
  return body.access_token;
}

async function main() {
  const missing = ['GRAPH_TENANT_ID', 'GRAPH_CLIENT_ID', 'GRAPH_CLIENT_SECRET']
    .filter((k) => !process.env[k]);
  if (missing.length) {
    console.log(`not configured yet — fill in: ${missing.join(', ')}`);
    return;
  }

  console.log('requesting a token...\n');
  const token = await getToken();
  const c = claims(token);

  const roles = c.roles || [];
  console.log(`app id     : ${c.appid || '(not in token)'}`);
  console.log(`tenant     : ${c.tid}`);
  console.log(`audience   : ${c.aud}`);
  console.log(`expires    : ${new Date(c.exp * 1000).toISOString()}`);
  console.log(`\npermissions this app actually has (${roles.length}):`);
  if (roles.length === 0) {
    console.log('  (none — no application permissions granted, or consent was never given)');
  } else {
    for (const r of roles.sort()) console.log(`  - ${r}`);
  }

  const canWriteFiles = roles.some((r) => r === NEEDED || r === 'Sites.ReadWrite.All');
  console.log(`\n${canWriteFiles ? '✓' : '✗'} ${NEEDED} — ${canWriteFiles
    ? 'present, uploads will work'
    : `MISSING. The admin needs to add it as an Application permission and grant consent.`}`);

  if (roles.includes('Mail.Send')) {
    console.log('✓ Mail.Send — also present, so the email path could move to Graph too');
  }

  if (!DRIVE || !FOLDER) {
    console.log('\nGRAPH_DRIVE_ID / GRAPH_FOLDER_ITEM_ID not set — skipping the folder check.');
    return;
  }

  console.log('\nchecking the target folder...');
  const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${DRIVE}/items/${FOLDER}`, {
    headers: { Authorization: `Bearer ${token}` },
  });
  const item = await res.json();
  if (!res.ok) {
    console.log(`✗ cannot read the folder — ${item.error?.code}: ${item.error?.message}`);
    return;
  }
  console.log(`✓ folder reachable: "${item.name}"  (${item.folder?.childCount ?? '?'} items inside)`);
  console.log(`  ${item.webUrl}`);
}

main().catch((e) => {
  console.error('✗ failed:', e.message);
  process.exit(1);
});
