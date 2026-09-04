/**
 * parseStrikePayload.js — Strike/Glide JSON → the canonical document shape.
 *
 * THIS IS THE ONLY FILE THAT KNOWS WHERE THE PAYLOAD CAME FROM. Vendor field
 * names, vendor quoting, vendor markdown, vendor date formats: all here.
 *
 * Real Glide payloads are messier than a clean webhook body:
 *   - `items` arrives as an ARRAY CONTAINING ONE STRING, and that string holds
 *     several JSON objects joined by ", " — not a JSON array
 *   - copied out of a Glide cell, the whole body can be CSV-quoted (`""` for `"`)
 *   - descriptions carry markdown (`**bold**`) from Glide's rich text editor
 *
 * All three are handled below. Nothing is mandatory: every field degrades to a
 * documented fallback rather than rejecting the submission.
 */

/** canonical name → source keys to try, in order. Case/spacing insensitive. */
const FIELD_MAP = {
  // One free-text label instead of separate part/PO fields — it can carry a
  // part number, a PO number, or a name, whichever the sender has.
  report_title:       ['report_title', 'reportTitle', 'Report Title', 'project_name', 'projectName', 'Project Name', 'project', 'title'],
  reference_name:     ['reference_name', 'referenceName', 'Reference Name', 'reference', 'part_number', 'partNumber', 'Part Number'],
  addressee:          ['addressee', 'addressed_to', 'addressedTo', 'Addressed to', 'attention', 'contact_name'],
  additional_details: ['additional_details', 'additional_information', 'Additional Information', 'additionalInfo', 'notes', 'remarks'],
  created_by:         ['created_by', 'createdBy', 'Created By', 'user', 'owner'],
  created_at:         ['created_at', 'createdAt', 'Created At', 'date', 'submitted_at', 'timestamp'],
  product_photo:      ['product_photo', 'productPhoto', 'Product Photo', 'cover_photo'],
  logo:               ['logo', 'logo_url'],
};

const ITEM_MAP = {
  // Which SKU or part this row is about — one document often spans several.
  item_name:      ['item_name', 'itemName', 'Item Name', 'item', 'sku', 'SKU', 'part_name', 'partName'],
  title:          ['title', 'Title', 'heading', 'subject'],
  body:           ['body', 'description', 'Description', 'details', 'text', 'query', 'content'],
  needs_response: ['needs_response', 'needsResponse', 'Needs Response', 'needs_reply', 'is_query', 'requires_response'],
};

const ITEM_ARRAY_KEYS = ['items', 'queries', 'sections', 'rows', 'entries', 'Query/Update', 'queryUpdates'];
const IMAGE_KEYS = ['images', 'photos', 'Photo', 'photo', 'attachments', 'files'];
const EMAIL_KEYS = ['to_email', 'toEmail', 'email', 'created_by_email', 'exported_by', 'user_email', 'sender_email'];

const norm = (k) => String(k).toLowerCase().replace(/[\s_\-/]/g, '');

/** Case- and separator-insensitive lookup across a list of aliases. */
function pick(obj, aliases) {
  if (!obj || typeof obj !== 'object') return undefined;
  const index = new Map(Object.keys(obj).map((k) => [norm(k), k]));
  for (const alias of aliases) {
    const hit = index.get(norm(alias));
    if (hit !== undefined) {
      const v = obj[hit];
      if (v !== null && v !== undefined && String(v).trim() !== '') return v;
    }
  }
  return undefined;
}

const str = (v) => (v === null || v === undefined ? '' : String(v).trim());

/** Glide sends booleans as true/false, "true"/"false", "Yes"/"No", or 1/0. */
function toBool(v) {
  if (v === undefined || v === null || v === '') return undefined;
  if (typeof v === 'boolean') return v;
  const s = String(v).trim().toLowerCase();
  if (['true', 'yes', 'y', '1', 'query', 'question'].includes(s)) return true;
  if (['false', 'no', 'n', '0', 'update', 'info'].includes(s)) return false;
  return undefined;
}

/** DD/MM/YYYY and DD-MM-YYYY → ISO. Anything else is passed through. */
function toIsoDate(v) {
  const s = str(v);
  if (!s) return '';
  const dmy = s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  if (dmy) return `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;
  return s;
}

/** Undo CSV-style quoting, as picked up when copying out of a Glide cell. */
function undoubleQuotes(s) {
  let t = String(s).trim();
  if (t.startsWith('"') && t.endsWith('"')) t = t.slice(1, -1);
  return t.replace(/""/g, '"');
}

/**
 * Pull every top-level {...} out of a string, honouring quotes and escapes.
 *
 * Glide joins repeating rows into one string — `{...}, {...}, {...}` — which is
 * not valid JSON. Splitting on ", " would tear apart any description containing
 * a comma, so this walks brace depth instead.
 */
function extractJsonObjects(text, onFailure) {
  const out = [];
  let depth = 0, start = -1, inString = false, escaped = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inString) {
      if (escaped) escaped = false;
      else if (ch === '\\') escaped = true;
      else if (ch === '"') inString = false;
      continue;
    }
    if (ch === '"') { inString = true; continue; }
    if (ch === '{') { if (depth === 0) start = i; depth += 1; continue; }
    if (ch === '}') {
      depth -= 1;
      if (depth === 0 && start >= 0) {
        const chunk = text.slice(start, i + 1);
        const row = parseObjectChunk(chunk);
        if (row) out.push(row);
        else if (onFailure) onFailure(chunk);   // never drop a row in silence
        start = -1;
      }
    }
  }
  return out;
}

/**
 * Parse one {...} chunk, repairing the one malformation seen in the wild:
 * raw newlines and tabs sitting inside a string value. JSON forbids literal
 * control characters there, so a row containing them would otherwise be thrown
 * away — losing a query without a word of warning.
 */
function parseObjectChunk(chunk) {
  try { return JSON.parse(chunk); } catch { /* try repairing it */ }

  let repaired = '';
  let inString = false, escaped = false;
  for (const ch of chunk) {
    if (inString) {
      if (escaped) { escaped = false; repaired += ch; continue; }
      if (ch === '\\') { escaped = true; repaired += ch; continue; }
      if (ch === '"') { inString = false; repaired += ch; continue; }
      if (ch === '\n') { repaired += '\\n'; continue; }
      if (ch === '\r') { repaired += '\\r'; continue; }
      if (ch === '\t') { repaired += '\\t'; continue; }
      repaired += ch;
      continue;
    }
    if (ch === '"') inString = true;
    repaired += ch;
  }

  try { return JSON.parse(repaired); } catch { return null; }
}

/** Whatever shape the rows arrived in, return a list of row objects. */
function coerceItemRows(value, warnings = []) {
  if (!value) return [];
  const rows = [];

  for (const entry of (Array.isArray(value) ? value : [value])) {
    if (!entry) continue;
    if (typeof entry === 'object') { rows.push(entry); continue; }
    if (typeof entry !== 'string') continue;

    const text = entry.trim();
    if (!text) continue;

    let parsed = null;
    try { parsed = JSON.parse(text); } catch { /* not plain JSON — keep going */ }

    if (Array.isArray(parsed)) {
      rows.push(...parsed.filter((r) => r && typeof r === 'object'));
      continue;
    }
    if (parsed && typeof parsed === 'object') { rows.push(parsed); continue; }

    const found = extractJsonObjects(text, (chunk) => {
      warnings.push(`a row could not be read and was skipped: ${chunk.slice(0, 80)}…`);
    });
    if (found.length) { rows.push(...found); continue; }

    // Not JSON at all — treat the string as the description itself.
    rows.push({ description: text });
  }

  return rows;
}

/**
 * Tidy rich text out of Glide's editor.
 *
 * Strips markdown emphasis markers, which would otherwise print as literal
 * asterisks in a client-facing document, and collapses the blank-line padding
 * its list formatting leaves behind. Only formatting characters are removed —
 * every word survives.
 */
function cleanRichText(value) {
  return String(value ?? '')
    .replace(/\r\n?/g, '\n')
    .replace(/[ \t]+$/gm, '')          // trailing spaces the editor leaves on list lines
    .replace(/\*\*(.+?)\*\*/g, '$1')   // **bold**
    .replace(/__(.+?)__/g, '$1')       // __bold__
    .replace(/\n{3,}/g, '\n\n')        // runs of blank lines
    .trim();
}

/** Images arrive as a list of URLs, a list of objects, a comma-joined string, or photo1..photoN. */
function extractImages(row) {
  const found = [];
  const listed = pick(row, IMAGE_KEYS);

  if (Array.isArray(listed)) {
    for (const im of listed) {
      if (!im) continue;
      if (typeof im === 'string') found.push({ path: im.trim(), caption: '' });
      else found.push({ path: im.path || im.url || im.src || '', caption: str(im.caption) });
    }
  } else if (typeof listed === 'string' && listed.trim()) {
    const text = listed.trim();
    let parsed = null;
    try { parsed = JSON.parse(text); } catch { /* plain string */ }
    if (Array.isArray(parsed)) {
      for (const im of parsed) {
        if (typeof im === 'string') found.push({ path: im.trim(), caption: '' });
        else if (im) found.push({ path: im.path || im.url || im.src || '', caption: str(im.caption) });
      }
    } else {
      for (const part of text.split(',')) {
        if (part.trim()) found.push({ path: part.trim(), caption: '' });
      }
    }
  }

  for (let i = 1; i <= 8; i++) {
    const v = pick(row, [`photo${i}`, `photo_${i}`, `Photo ${i}`, `image${i}`, `image_${i}`]);
    if (v) found.push({ path: str(v), caption: str(pick(row, [`caption${i}`, `caption_${i}`])) });
  }

  return found.filter((im) => im.path);
}

/** Rows under any of the known keys, else flat description_N keys. */
function extractItems(body, warnings = []) {
  for (const key of ITEM_ARRAY_KEYS) {
    const raw = pick(body, [key]);
    const rows = coerceItemRows(raw, warnings);
    if (rows.length) return rows;
  }

  const flat = [];
  for (let i = 1; i <= 50; i++) {
    const desc = pick(body, [`description_${i}`, `description${i}`, `Description ${i}`, `query_${i}`]);
    if (desc === undefined) continue;
    const row = { description: desc };
    for (const [k, v] of Object.entries(body)) {
      const m = norm(k).match(/^(photo|image|caption|title|itemname|needsresponse)(\d+)?[_]?(\d+)?$/);
      if (m && (m[3] === String(i) || m[2] === String(i))) row[k] = v;
    }
    flat.push(row);
  }
  return flat;
}

/** The body itself may arrive as a JSON string, possibly CSV-quoted. */
function coerceBody(body) {
  if (body && typeof body === 'object') return body;
  if (typeof body !== 'string') return {};
  for (const candidate of [body.trim(), undoubleQuotes(body)]) {
    try {
      const o = JSON.parse(candidate);
      if (o && typeof o === 'object') return o;
    } catch { /* try the next form */ }
  }
  return {};
}

/**
 * @param {object|string} body  raw webhook body
 * @returns {{document: object, items: object[], delivery: object, writeback: object}}
 */
function parseStrikePayload(body) {
  const warnings = [];
  const src = coerceBody(body);
  const meta = pick(src, ['document', 'meta', 'metadata', 'Sample', 'row']) || src;

  const document = {};
  for (const [canonical, aliases] of Object.entries(FIELD_MAP)) {
    const v = pick(meta, aliases) ?? pick(src, aliases);
    document[canonical] = canonical === 'created_at' ? toIsoDate(v) : (v ?? null);
  }
  if (!document.created_at) document.created_at = new Date().toISOString().slice(0, 10);
  if (typeof document.product_photo === 'string') document.product_photo = { path: document.product_photo };

  const items = extractItems(src, warnings).map((row, i) => ({
    id: str(pick(row, ['id', 'row_id', 'rowId'])) || `itm_${String(i + 1).padStart(3, '0')}`,
    item_name: str(pick(row, ITEM_MAP.item_name)),
    title: cleanRichText(str(pick(row, ITEM_MAP.title))),
    body: cleanRichText(str(pick(row, ITEM_MAP.body))),
    needs_response: toBool(pick(row, ITEM_MAP.needs_response)),   // undefined → enrichment decides
    images: extractImages(row),
  })).filter((it) => it.body || it.title || it.images.length);

  if (!items.length && src && Object.keys(src).length) {
    warnings.push('no query or update rows were found in the payload — check the items field');
  }

  return {
    document,
    items,
    warnings,
    delivery: {
      to: str(pick(src, EMAIL_KEYS) ?? pick(meta, EMAIL_KEYS)),
      cc: str(pick(src, ['cc', 'cc_email', 'ccEmail'])),
      bcc: str(pick(src, ['bcc', 'bcc_email', 'bccEmail'])),
    },
    writeback: {
      rowId: str(pick(src, ['row_id', 'rowId', 'rowID', 'glide_row_id'])),
      table: str(pick(src, ['table', 'table_name', 'tableName', 'glide_table'])),
      column: str(pick(src, ['column', 'column_id', 'columnId', 'target_column'])),
    },
  };
}

module.exports = {
  parseStrikePayload, FIELD_MAP, ITEM_MAP,
  toBool, toIsoDate, cleanRichText, extractJsonObjects, coerceItemRows, undoubleQuotes,
};
