/**
 * fetchImage.js — get bytes for one image, whatever form the payload gave us.
 *
 * Accepts an https URL, a data: URI, or a local path. Deliberately knows
 * nothing about Glide, Strike, or any other platform: if a URL ever needs
 * reshaping for a vendor, that belongs in parseStrikePayload.js, which is the
 * one file allowed to know where the payload came from.
 *
 * Never throws. An image that cannot be read comes back with ok:false and the
 * planner renders a labelled placeholder in its place (spec §13).
 */

const fs = require('fs');
const fetch = require('node-fetch');
const sharp = require('sharp');
const { IMAGE_MAX_LONG_EDGE } = require('../plan/constants');

const FETCH_TIMEOUT_MS = 15000;

/** Read metadata and downscale if the source is larger than we need. */
async function normalise(buffer, label) {
  const meta = await sharp(buffer).metadata();
  const longEdge = Math.max(meta.width || 0, meta.height || 0);

  let out = buffer;
  let { width, height } = meta;

  if (longEdge > IMAGE_MAX_LONG_EDGE) {
    const resized = sharp(buffer).resize({
      width: meta.width >= meta.height ? IMAGE_MAX_LONG_EDGE : undefined,
      height: meta.height > meta.width ? IMAGE_MAX_LONG_EDGE : undefined,
      fit: 'inside',
      withoutEnlargement: true,
    });
    out = await resized.png().toBuffer();
    const after = await sharp(out).metadata();
    width = after.width;
    height = after.height;
  } else if (meta.format !== 'png' && meta.format !== 'jpeg') {
    // §13: convert what we can to PNG rather than treat it as unreadable.
    out = await sharp(buffer).png().toBuffer();
  }

  const mime = out === buffer && meta.format === 'jpeg' ? 'image/jpeg' : 'image/png';

  return {
    ok: true,
    label,
    width, height,
    aspect: width && height ? width / height : 1,
    buffer: out,
    // pptxgenjs wants "image/png;base64,..." — note: no "data:" prefix.
    data: `${mime};base64,${out.toString('base64')}`,
  };
}

const failed = (label, reason) => ({ ok: false, label, reason, width: 0, height: 0, aspect: 1, buffer: null, data: null });

/**
 * @param {string|object} source  URL, data URI, path, or {path|url|src, caption}
 * @returns {Promise<object>} always resolves
 */
async function fetchImage(source) {
  const ref = typeof source === 'string' ? { path: source } : (source || {});
  const url = ref.path || ref.url || ref.src || '';
  const label = ref.caption || String(url).split('/').pop() || 'image';

  if (!url || !String(url).trim()) return failed(label, 'no path');

  try {
    if (String(url).startsWith('data:')) {
      const b64 = String(url).slice(String(url).indexOf(',') + 1);
      return await normalise(Buffer.from(b64, 'base64'), label);
    }

    if (/^https?:\/\//i.test(url)) {
      const res = await fetch(url, { timeout: FETCH_TIMEOUT_MS });
      if (!res.ok) return failed(label, `HTTP ${res.status}`);
      return await normalise(Buffer.from(await res.arrayBuffer()), label);
    }

    if (fs.existsSync(url)) {
      return await normalise(fs.readFileSync(url), label);
    }

    return failed(label, 'not found');
  } catch (e) {
    console.warn(`  image unreadable (${label}): ${e.message}`);
    return failed(label, e.message);
  }
}

module.exports = { fetchImage };
