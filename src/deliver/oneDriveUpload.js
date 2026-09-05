/**
 * oneDriveUpload.js — put the finished deck in a OneDrive/SharePoint folder.
 *
 * OPTIONAL, like every other delivery step: missing credentials are a logged
 * skip, never an error. The email attachment remains the guaranteed path.
 *
 * There is no Graph API for authoring PowerPoint content — Excel has a workbook
 * API, PowerPoint has no equivalent. So the deck is generated here and the
 * finished file is uploaded, which is the only real approach rather than a
 * workaround.
 *
 * Auth is the OAuth2 client-credentials flow. Tokens last about an hour and are
 * cached in module scope, so a burst of submissions doesn't re-authenticate
 * for each one.
 */

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

// Credentials come from the environment; the destination comes from the
// payload so each report can land in its own project folder, falling back to
// the env default when the submission doesn't name one.
const CREDENTIALS = ['GRAPH_TENANT_ID', 'GRAPH_CLIENT_ID', 'GRAPH_CLIENT_SECRET'];

const isConfigured = () => CREDENTIALS.every((k) => Boolean(process.env[k]));
const missingVars = () => CREDENTIALS.filter((k) => !process.env[k]);

/** Payload destination wins; env vars are the default. */
function resolveDestination(storage = {}) {
  return {
    driveId: storage.driveId || process.env.GRAPH_DRIVE_ID || '',
    folderId: storage.folderId || process.env.GRAPH_FOLDER_ITEM_ID || '',
  };
}

let cached = { token: null, expiresAt: 0 };

async function getToken() {
  // 60s of slack so a token can't expire mid-upload.
  if (cached.token && Date.now() < cached.expiresAt - 60_000) return cached.token;

  const res = await fetch(
    `https://login.microsoftonline.com/${process.env.GRAPH_TENANT_ID}/oauth2/v2.0/token`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: new URLSearchParams({
        client_id: process.env.GRAPH_CLIENT_ID,
        client_secret: process.env.GRAPH_CLIENT_SECRET,
        grant_type: 'client_credentials',
        scope: 'https://graph.microsoft.com/.default',
      }),
    },
  );

  const body = await res.json();
  if (!res.ok) throw new Error(`token request failed — ${body.error}: ${(body.error_description || '').split('\n')[0]}`);

  cached = { token: body.access_token, expiresAt: Date.now() + body.expires_in * 1000 };
  return cached.token;
}

/**
 * @param {Buffer} buffer    the .pptx
 * @param {string} filename  as it should appear in the folder
 * @param {object} storage   { driveId, folderId } from the payload; optional
 * @returns {Promise<{url: string|null, itemId: string|null, skipped: boolean, reason?: string}>}
 */
async function uploadToOneDrive(buffer, filename, storage = {}, contentType = PPTX_MIME) {
  if (!isConfigured()) {
    const reason = `OneDrive not configured (${missingVars().join(', ')})`;
    console.log(`  onedrive: skipped — ${reason}`);
    return { url: null, itemId: null, skipped: true, reason };
  }

  const { driveId: drive, folderId: folder } = resolveDestination(storage);
  if (!drive || !folder) {
    const reason = 'no destination — payload carried no drive_id/folder_item_id and no env default is set';
    console.log(`  onedrive: skipped — ${reason}`);
    return { url: null, itemId: null, skipped: true, reason };
  }

  try {
    const token = await getToken();

    // conflictBehavior=rename rather than the default replace: two submissions
    // for the same project on the same day would otherwise silently overwrite
    // each other, and a client deck lost that way is hard to notice.
    const url = `https://graph.microsoft.com/v1.0/drives/${drive}/items/${folder}:/`
      + `${encodeURIComponent(filename)}:/content?%40microsoft.graph.conflictBehavior=rename`;

    const res = await fetch(url, {
      method: 'PUT',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': contentType },
      body: buffer,
    });

    const item = await res.json();
    if (!res.ok) {
      throw new Error(`${item.error?.code || res.status}: ${item.error?.message || 'upload rejected'}`);
    }

    console.log(`  onedrive → ${item.name}${storage.folderId ? ' (folder from payload)' : ' (folder from env default)'}`);
    return { url: item.webUrl, itemId: item.id, skipped: false };
  } catch (e) {
    // Non-fatal: the deck is already attached to the email.
    console.warn(`  onedrive upload failed (non-fatal): ${e.message}`);
    return { url: null, itemId: null, skipped: true, reason: e.message };
  }
}

module.exports = { uploadToOneDrive, isConfigured, missingVars, resolveDestination, PPTX_MIME };
