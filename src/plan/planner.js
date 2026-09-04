/**
 * planner.js — document JSON in, complete slide plan out.
 *
 * PURE: no I/O, no clock, no randomness. Every box is resolved to inches here so
 * the renderer only has to draw. This is where §6–§13 of the format spec live,
 * and it is what the acceptance tests exercise.
 *
 * Governing rule (PLAN.md §5.1): content is never truncated. Anything that does
 * not fit flows onto another slide.
 */

const C = require('./constants');
const { wrap, heightOf, splitToFit } = require('./textFit');
const { planHeader, contentBox, layoutItemSlide } = require('./layout');

const MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const TITLE_GAP  = 0.35;
const FIELD_GAP  = 0.20;
const BLOCK_GAP  = 0.28;
const HEADING_GAP = 0.30;
const ENTRY_GAP  = 0.14;
const SECTION_GAP = 0.22;

/** ISO or loose date → "03 Sep 2026". UTC, so output never depends on the host. */
function formatDate(value) {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);
  return `${String(d.getUTCDate()).padStart(2, '0')} ${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`;
}

const clean = (v) => (typeof v === 'string' ? v.trim() : v ? String(v).trim() : '');

/**
 * Title fallbacks — never leaves an item unlabelled.
 *
 * Only the first line is used. Bodies routinely open with a lead-in and then a
 * numbered list, so taking a flat slice of characters produced headers with
 * line breaks and half a list item in them. Cuts on a word boundary and adds no
 * ellipsis: a header is a label, not truncated content.
 */
function resolveTitle(item) {
  const t = clean(item.title);
  if (t) return t.split(/\n/)[0].trim();

  const b = clean(item.body);
  if (!b) return C.UNTITLED;

  const firstLine = b.split(/\n/).map((l) => l.trim()).find(Boolean) || '';
  if (!firstLine) return C.UNTITLED;

  const trimmed = firstLine.replace(/[\s:;,.\-–—]+$/, '');
  if (trimmed.length <= C.TITLE_FALLBACK_CHARS) return trimmed || C.UNTITLED;

  const cut = trimmed.slice(0, C.TITLE_FALLBACK_CHARS);
  const space = cut.lastIndexOf(' ');
  return (space > 20 ? cut.slice(0, space) : cut).replace(/[\s:;,.\-–—]+$/, '') || C.UNTITLED;
}

/** Normalise input and assign gapless query numbers over flagged items. */
function normalize(doc) {
  const warnings = [];
  const d = doc.document || {};
  const items = (doc.items || []).map((raw, i) => {
    const images = (raw.images || []).map((im) => ({
      path: im.path || im.src || im.url || '',
      caption: clean(im.caption),
      width: im.width || 0,
      height: im.height || 0,
      aspect: im.aspect || (im.width && im.height ? im.width / im.height : 1),
      ok: im.ok !== false,
      data: im.data || null,
    }));
    if (images.length > C.MAX_IMAGES_PER_ITEM) {
      warnings.push(`item ${raw.id || i}: ${images.length} images (form cap is ${C.MAX_IMAGES_PER_ITEM}) — all rendered across additional slides`);
    }
    return {
      id: raw.id || `itm_${String(i + 1).padStart(3, '0')}`,
      item_name: clean(raw.item_name),
      title: resolveTitle(raw),
      body: clean(raw.body ?? raw.description),
      needs_response: raw.needs_response === true,
      images,
    };
  });

  let n = 0;
  for (const it of items) if (it.needs_response) it.queryNo = ++n;

  return {
    document: {
      report_title: clean(d.report_title ?? d.project_name),
      reference_name: clean(d.reference_name ?? d.part_number),
      addressee: clean(d.addressee),
      additional_details: clean(d.additional_details),
      created_by: clean(d.created_by),
      created_at: clean(d.created_at),
      product_photo: d.product_photo || null,
      logo: d.logo || null,
    },
    items,
    queryCount: n,
    updateCount: items.length - n,
    warnings,
  };
}

// ── Shared furniture ─────────────────────────────────────────

function logoBox(logo, height, rightAlign) {
  if (!logo || logo.ok === false) return null;
  const aspect = logo.aspect || 3;
  const w = Math.min(height * aspect, rightAlign ? 3 : C.FOOTER_LOGO_SLOT_W);
  return {
    image: logo,
    box: rightAlign
      ? { x: C.SLIDE_W - C.MARGIN - w, y: 0.35, w, h: height }
      : { x: C.CONTENT_X, y: C.FOOTER_Y + 0.06, w, h: height },
  };
}

/**
 * Footer for the cover and contents slides.
 *
 * `hasLogo` shifts the text clear of the bottom-left logo. Without it the text
 * box starts at CONTENT_X, exactly where the logo sits, and the two print on
 * top of each other — item slides already offset for this, the contents slide
 * did not.
 *
 * `showPreparedBy` carries the author on the cover only. On the contents slide
 * the logo already says who sent it, so the name is noise.
 */
function coverFooter(d, { hasLogo = false, showPreparedBy = true } = {}) {
  const parts = [];
  if (showPreparedBy && d.created_by) parts.push(`Prepared by ${d.created_by}`);
  if (d.created_at) parts.push(formatDate(d.created_at));

  const inset = hasLogo ? C.FOOTER_LOGO_SLOT_W + 0.1 : 0;
  return {
    text: parts.join(' · '),
    box: { x: C.CONTENT_X + inset, y: C.FOOTER_Y, w: C.CONTENT_W - inset, h: C.FOOTER_H },
  };
}

/** Item footer: project · part · date on the left, "Page N" on the right. */
function itemFooter(d, hasLogo) {
  const parts = [d.report_title, d.reference_name, formatDate(d.created_at)].filter(Boolean);
  const x = hasLogo ? C.CONTENT_X + C.FOOTER_LOGO_SLOT_W + 0.1 : C.CONTENT_X;
  return {
    text: parts.join(' · '),
    box: { x, y: C.FOOTER_Y, w: C.SLIDE_W - C.MARGIN - x - 1.2, h: C.FOOTER_H },
    page: { box: { x: C.SLIDE_W - C.MARGIN - 1.1, y: C.FOOTER_Y, w: 1.1, h: C.FOOTER_H } },
  };
}

// ── Cover (§7.1) ─────────────────────────────────────────────

function planCover(n, showInstruction) {
  const d = n.document;
  const hasPhoto = Boolean(d.product_photo && d.product_photo.ok !== false);
  const variant = hasPhoto ? 'A' : 'B';

  const textW = hasPhoto ? C.CONTENT_W * 0.55 : C.CONTENT_W * 0.70;
  const blocks = [];
  let y = C.CONTENT_Y;

  if (d.report_title) {
    const size = C.TYPE.coverTitle.size;
    const lines = wrap(d.report_title, textW, size, true);
    const h = heightOf(lines.length, size);
    blocks.push({ kind: 'title', lines, size, box: { x: C.CONTENT_X, y, w: textW, h } });
    y += h + TITLE_GAP;
  }

  // One free-text reference instead of separate part/PO fields — it carries a
  // part number, a PO number, or a name, whichever the sender has.
  const fields = [
    ['Reference', d.reference_name],
    ['Attention', d.addressee],
  ].filter(([, v]) => v);

  for (const [label, value] of fields) {
    const labelH = heightOf(1, C.TYPE.coverLabel.size);
    const vLines = wrap(value, textW, C.TYPE.coverValue.size);
    const valueH = heightOf(vLines.length, C.TYPE.coverValue.size);
    blocks.push({
      kind: 'field',
      label,
      lines: vLines,
      labelBox: { x: C.CONTENT_X, y, w: textW, h: labelH },
      valueBox: { x: C.CONTENT_X, y: y + labelH, w: textW, h: valueH },
    });
    y += labelH + valueH + FIELD_GAP;
  }

  if (d.additional_details) {
    // Rendered in full — the spec's 240-char truncation is dropped (PLAN.md §5.1).
    const size = C.TYPE.coverValue.size;
    const lines = wrap(d.additional_details, textW, size);
    const h = heightOf(lines.length, size);
    blocks.push({ kind: 'note', lines, size, box: { x: C.CONTENT_X, y: y + 0.06, w: textW, h } });
    y += h + BLOCK_GAP;
  }

  let instruction = null;
  if (showInstruction) {
    const size = C.TYPE.body.size;
    const lines = wrap(C.REPLY_INSTRUCTION, textW, size);
    instruction = {
      text: C.REPLY_INSTRUCTION, lines, size,
      box: { x: C.CONTENT_X, y: y + 0.1, w: textW, h: heightOf(lines.length, size) },
    };
  }

  const photoW = C.CONTENT_W - textW - C.GUTTER;
  const photoZone = { x: C.CONTENT_X + textW + C.GUTTER, y: C.CONTENT_Y, w: photoW, h: C.CONTENT_BOTTOM - C.CONTENT_Y };

  return {
    kind: 'cover',
    variant,
    blocks,
    instruction,
    photo: hasPhoto ? { image: d.product_photo, zone: photoZone } : null,
    logo: logoBox(d.logo, C.COVER_LOGO_H, true),
    footer: coverFooter(d, { hasLogo: false, showPreparedBy: true }),
  };
}

// ── Summary (§7.2, reworked) ─────────────────────────────────

/**
 * The summary slide, grouped into sections.
 *
 * The spec listed only queries here, on the grounds that naming updates dilutes
 * the "how many answers are owed" signal. In practice the opposite happened: a
 * client flicking through hit unnumbered slides that the summary never
 * mentioned, and had to infer the convention with nobody to ask. Both sections
 * are listed now, and the count moves into the QUERIES label so the signal
 * survives (PLAN.md §5.8).
 *
 * Rows — section labels and entries alike — are laid out as one flat list, so
 * pagination stays the simple "fill until full, then start another slide".
 */
function planSummary(n) {
  const entrySize = C.TYPE.contentsEntry.size;
  const labelSize = C.TYPE.sectionLabel.size;
  const area = { x: C.CONTENT_X, y: C.CONTENT_Y, w: C.CONTENT_W, h: C.CONTENT_BOTTOM - C.CONTENT_Y };

  const label = (text) => ({
    type: 'section', text, size: labelSize,
    lines: wrap(text, area.w, labelSize, true),
    h: heightOf(1, labelSize) + SECTION_GAP,
  });
  const entry = (text, queryNo) => {
    const lines = wrap(text, area.w, entrySize);   // wraps, never truncated
    return { type: 'entry', n: queryNo || null, text, lines, size: entrySize,
             h: heightOf(lines.length, entrySize) + ENTRY_GAP };
  };

  const queries = n.items.filter((i) => i.needs_response);
  const updates = n.items.filter((i) => !i.needs_response);

  const rows = [];
  if (queries.length) {
    rows.push(label(C.SECTION_QUERIES(queries.length)));
    for (const it of queries) {
      rows.push(entry(it.item_name
        ? `Q${it.queryNo} · ${it.item_name} — ${it.title}`
        : `Q${it.queryNo} · ${it.title}`, it.queryNo));
    }
  }
  if (updates.length) {
    rows.push(label(C.SECTION_UPDATES));
    for (const it of updates) {
      rows.push(entry(it.item_name ? `· ${it.item_name} — ${it.title}` : `· ${it.title}`));
    }
  }

  const tailLines = queries.length
    ? [{ kind: 'instruction', text: C.REPLY_INSTRUCTION }]
    : [];
  const tailH = tailLines.length ? heightOf(1, C.TYPE.body.size) + ENTRY_GAP + 0.18 : 0;

  const slides = [];
  let idx = 0;

  while (idx < rows.length) {
    const isFirst = slides.length === 0;
    let cursor = area.y;
    let headingBlock = null;

    if (isFirst) {
      const size = C.TYPE.contentsHead.size;
      const lines = wrap(C.SUMMARY_HEADING, area.w, size, true);
      const h = heightOf(lines.length, size);
      headingBlock = { text: C.SUMMARY_HEADING, lines, size, box: { x: area.x, y: cursor, w: area.w, h } };
      cursor += h + HEADING_GAP;
    }

    const bottom = area.y + area.h * C.FILL_TARGET;

    // Try to close the list here, leaving room for the reply instruction.
    let take = 0, probe = cursor;
    while (idx + take < rows.length && take < C.CONTENTS_MAX_ENTRIES) {
      if (probe + rows[idx + take].h > bottom - tailH) break;
      probe += rows[idx + take].h; take += 1;
    }

    if (idx + take < rows.length) {
      // Not closing here — reclaim the tail's space for more rows.
      take = 0; probe = cursor;
      while (idx + take < rows.length && take < C.CONTENTS_MAX_ENTRIES) {
        if (probe + rows[idx + take].h > bottom) break;
        probe += rows[idx + take].h; take += 1;
      }
      if (take === 0) take = 1;   // always make progress
    }

    // Never strand a section label at the foot of a slide.
    if (take > 1 && idx + take < rows.length && rows[idx + take - 1].type === 'section') take -= 1;

    const placed = [];
    for (let k = 0; k < take; k++) {
      const r = rows[idx + k];
      const gap = r.type === 'section' ? SECTION_GAP : ENTRY_GAP;
      placed.push({ ...r, box: { x: area.x, y: cursor, w: area.w, h: r.h - gap } });
      cursor += r.h;
    }
    idx += take;

    let tail = null;
    if (idx >= rows.length) {
      tail = [];
      let ty = cursor + 0.18;
      for (const t of tailLines) {
        const size = C.TYPE.body.size;
        const h = heightOf(1, size);
        tail.push({ ...t, size, box: { x: area.x, y: ty, w: area.w, h } });
        ty += h + ENTRY_GAP;
      }
    }

    const summaryLogo = logoBox(n.document.logo, C.FOOTER_LOGO_H, false);
    slides.push({ kind: 'contents', heading: headingBlock, rows: placed, tail,
      logo: summaryLogo,
      footer: coverFooter(n.document, { hasLogo: Boolean(summaryLogo), showPreparedBy: false }) });
  }

  return slides;
}

// ── Item slides (§8–§11) ─────────────────────────────────────

function buildItemSlide(item, attempt, { continued, isLast }) {
  const header = planHeader({
    queryNo: item.queryNo || null, title: item.title, continued, itemName: item.item_name,
  });
  const flagged = item.needs_response;

  return {
    kind: 'item',
    itemId: item.id,
    queryNo: item.queryNo || null,
    itemName: item.item_name,
    title: item.title,
    continued,
    header,
    layout: attempt.layout,
    body: attempt.body,
    images: attempt.images,
    replyBox: flagged && isLast
      ? {
          label: C.REPLY_LABEL(item.queryNo),
          box: { x: C.CONTENT_X, y: C.REPLY_Y, w: C.CONTENT_W, h: C.REPLY_H },
          labelBox: {
            x: C.CONTENT_X + C.REPLY_INSET, y: C.REPLY_Y + C.REPLY_INSET,
            w: C.CONTENT_W - C.REPLY_INSET * 2, h: heightOf(1, C.TYPE.replyLabel.size),
          },
        }
      : null,
  };
}

/**
 * One item → one or more slides.
 *
 * Each pass first asks "can this be the last slide?" — laid out with the reply
 * box, which shrinks the content area. If everything fits, it is the last slide.
 * If not, the slide is re-laid out at full height and the remainder carries on.
 * Terminates because every pass consumes at least one image or one line.
 */
function planItem(item, warnings) {
  const slides = [];
  let body = item.body;
  let images = item.images.slice();
  let continued = false;
  let guard = 0;

  for (;;) {
    if (++guard > 200) {
      warnings.push(`item ${item.id}: layout guard tripped after ${guard} slides`);
      break;
    }

    const headerH = planHeader({
      queryNo: item.queryNo || null, title: item.title, continued, itemName: item.item_name,
    }).height;
    const flagged = item.needs_response;

    const trial = layoutItemSlide({ body, images, headerH, hasReply: flagged });
    if (trial.bodyRemainder === '' && trial.imagesUsed === images.length) {
      slides.push(buildItemSlide(item, trial, { continued, isLast: true }));
      break;
    }

    let chosen = layoutItemSlide({ body, images, headerH, hasReply: false });

    // If the full-height layout would swallow everything, there'd be nothing
    // left for the reply box's own slide. Keep the trial's split instead so the
    // remainder carries the reply box on a final slide.
    if (flagged && chosen.bodyRemainder === '' && chosen.imagesUsed === images.length) {
      chosen = trial;
    }

    slides.push(buildItemSlide(item, chosen, { continued, isLast: false }));

    const beforeBody = body.length;
    const beforeImgs = images.length;
    body = chosen.bodyRemainder;
    images = images.slice(chosen.imagesUsed);

    if (body.length === beforeBody && images.length === beforeImgs) {
      warnings.push(`item ${item.id}: no layout progress — remaining content dropped to avoid a loop`);
      break;
    }
    continued = true;
  }

  return slides;
}

/** §12 — consecutive text-only updates may share a slide. */
function tryGroup(group, doc) {
  const area = { x: C.CONTENT_X, y: C.CONTENT_Y, w: C.CONTENT_W, h: C.CONTENT_BOTTOM - C.CONTENT_Y };
  const blocks = [];
  let y = area.y;

  for (const it of group) {
    const tSize = C.TYPE.groupedTitle.size;
    const tLines = wrap(it.item_name ? `${it.item_name} — ${it.title}` : it.title, area.w, tSize, true);
    const tH = heightOf(tLines.length, tSize);
    const bSize = C.TYPE.body.size;
    const bLines = it.body ? wrap(it.body, area.w, bSize) : [];
    const bH = heightOf(bLines.length, bSize);

    blocks.push({
      itemId: it.id,
      itemName: it.item_name,
      title: { text: it.title, lines: tLines, size: tSize, box: { x: area.x, y, w: area.w, h: tH } },
      body: bLines.length
        ? { lines: bLines, size: bSize, box: { x: area.x, y: y + tH + 0.08, w: area.w, h: bH } }
        : null,
    });
    y += tH + (bH ? bH + 0.08 : 0) + C.GUTTER;
  }

  const used = y - area.y - C.GUTTER;
  if (used > area.h * C.FILL_TARGET) return null;

  return {
    kind: 'grouped',
    chip: {
      text: C.UPDATE_CHIP_LABEL, tone: 'update', size: C.CHIP.update.size,
      box: { x: C.CONTENT_X, y: C.HEADER_Y, w: 0.87, h: C.HEADER_H },
    },
    blocks,
    logo: logoBox(doc.logo, C.FOOTER_LOGO_H, false),
  };
}

/**
 * Item slides, grouped into sections: everything owing an answer first, then
 * the rest. Order within each section is the sender's own.
 *
 * This deliberately overrides the spec's "item order is preserved absolutely".
 * That rule existed to protect numbering clarity, but in a real deck it worked
 * against it — a client hit numbered and unnumbered slides in alternation and
 * had to infer the convention unaided. Grouping serves the same goal better
 * (PLAN.md §5.8). A bonus: updates become contiguous, so §12 grouping actually
 * fires and text-only updates share slides.
 */
function planItemSlides(n, warnings) {
  const ordered = [
    ...n.items.filter((it) => it.needs_response),
    ...n.items.filter((it) => !it.needs_response),
  ];

  const slides = [];
  let i = 0;

  while (i < ordered.length) {
    const it = ordered[i];
    const groupable = (x) => !x.needs_response && x.images.length === 0;

    if (groupable(it)) {
      let span = 1;
      while (span < C.MAX_GROUPED_UPDATES && i + span < ordered.length && groupable(ordered[i + span])) span++;

      let placed = false;
      for (let size = span; size >= 2 && !placed; size--) {
        const g = tryGroup(ordered.slice(i, i + size), n.document);
        if (g) { slides.push(g); i += size; placed = true; }
      }
      if (placed) continue;
    }

    slides.push(...planItem(it, warnings));
    i += 1;
  }

  return slides;
}

// ── Entry point ──────────────────────────────────────────────

function planSlides(doc) {
  const n = normalize(doc);
  const warnings = n.warnings.slice();

  const showContents = n.items.length >= C.INCLUDE_SUMMARY_MIN_ITEMS;
  const instructionOnCover = !showContents && n.queryCount >= 1;

  const slides = [planCover(n, instructionOnCover)];
  if (showContents) slides.push(...planSummary(n));
  slides.push(...planItemSlides(n, warnings));

  // Page numbers and footers, once the sequence is final.
  slides.forEach((s, i) => {
    s.page = i + 1;
    if (s.kind === 'item' || s.kind === 'grouped') {
      s.logo = s.logo || logoBox(n.document.logo, C.FOOTER_LOGO_H, false);
      s.footer = itemFooter(n.document, Boolean(s.logo));
      s.pageLabel = `Page ${i + 1}`;
    }
  });

  return {
    meta: {
      totalSlides: slides.length,
      queryCount: n.queryCount,
      updateCount: n.updateCount,
      hasContents: showContents,
      warnings,
    },
    document: n.document,
    slides,
  };
}

module.exports = { planSlides, normalize, formatDate, resolveTitle };
