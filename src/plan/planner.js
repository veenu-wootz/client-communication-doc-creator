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
const SUMMARY_BULLET_INDENT = 0.32;   // room PowerPoint's own number takes

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

/** Normalise input and number every item sequentially, in input order. */
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

  // Every item is numbered, in the order the sender wrote it. needs_response is
  // still carried — enrichment sets it and the grouped-queries-format branch
  // needs it — but it no longer affects numbering, ordering, or layout.
  items.forEach((it, i) => { it.number = i + 1; });

  return {
    document: {
      // No cross-filling. parseStrikePayload is the single place that decides
      // what each field holds; a fallback here quietly re-merged report_title
      // and project_name after they were deliberately separated.
      report_title: clean(d.report_title),
      reference_name: clean(d.reference_name),
      project_name: clean(d.project_name),
      addressee: clean(d.addressee),
      additional_details: clean(d.additional_details),
      created_by: clean(d.created_by),
      created_at: clean(d.created_at),
      product_photo: d.product_photo || null,
      logo: d.logo || null,
    },
    items,
    itemCount: items.length,
    // Still counted for the notification email's summary line, though nothing
    // in the layout distinguishes them any more.
    queryCount: items.filter((i) => i.needs_response).length,
    updateCount: items.filter((i) => !i.needs_response).length,
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

  // The column is reserved either way: with a photo it holds one, without it
  // holds a placeholder so the sender can add one. The cover is a template.
  const textW = C.CONTENT_W * 0.55;
  const blocks = [];
  let y = C.CONTENT_Y;

  if (d.report_title) {
    const size = C.TYPE.coverTitle.size;
    const lines = wrap(d.report_title, textW, size, true);
    const h = heightOf(lines.length, size);
    blocks.push({ kind: 'title', lines, size, box: { x: C.CONTENT_X, y, w: textW, h } });
    y += h + TITLE_GAP;
  }

  /**
   * A cover field. The label ALWAYS renders, value or not — the cover is a
   * template now, so an absent value shows a hairline rule to write on rather
   * than vanishing. A rule reads as a form line, so nothing has to be deleted
   * if it is left blank; a filled box would have to be (PLAN.md §5.14).
   */
  const field = (label, value, emptyLines = 1) => {
    const labelH = heightOf(1, C.TYPE.coverLabel.size);
    const size = C.TYPE.coverValue.size;
    const filled = Boolean(value);
    const vLines = filled ? wrap(value, textW, size) : [];
    // An empty field reserves real writing room rather than a single line —
    // free text needs somewhere to go, and the rule sits at the foot of it.
    const valueH = filled ? heightOf(vLines.length, size) : heightOf(emptyLines, size);

    blocks.push({
      kind: 'field',
      label,
      filled,
      lines: vLines,
      labelBox: { x: C.CONTENT_X, y, w: textW, h: labelH },
      valueBox: { x: C.CONTENT_X, y: y + labelH, w: textW, h: valueH },
      // Sits on the value's baseline, so typed text lands on the line.
      rule: filled ? null : { x: C.CONTENT_X, y: y + labelH + valueH, w: textW },
    });
    y += labelH + valueH + FIELD_GAP;
  };

  field('Reference number', d.reference_name);
  if (d.addressee) field('Attention', d.addressee);
  field('Additional information', d.additional_details, C.COVER_NOTE_EMPTY_LINES);

  // Prepared by moves out of the footer and becomes a field like the others —
  // it is information about the document, and at 9pt in the footer nobody read it.
  const prepared = [d.created_by, d.created_at ? formatDate(d.created_at) : '']
    .filter(Boolean).join('  ·  ');
  field('Prepared by', prepared);

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
    photoPlaceholder: hasPhoto ? null : { box: photoZone },
    logo: logoBox(d.logo, C.COVER_LOGO_H, true),
    footer: null,   // prepared-by now lives in the body; the logo carries the rest
  };
}

// ── Summary ──────────────────────────────────────────────────

/**
 * The summary slide: the deck's points, listed in order.
 *
 * Rendered as ONE text box per slide with PowerPoint's own numbered bullets, so
 * a sender can click into the list and press Enter to add a point — the number
 * appears by itself. Separate boxes per row made that impossible, which was the
 * single biggest obstacle to anyone adopting the deck as a template.
 *
 * No "3 need a reply" count: it would go stale the moment a slide is added by
 * hand, and a confidently wrong count is worse than none (PLAN.md §5.14).
 *
 * Pagination survives — a forty-item list still cannot overflow its box — and
 * each continuation slide restarts PowerPoint's numbering at the right value.
 */
function planSummary(n) {
  const entrySize = C.TYPE.contentsEntry.size;
  const area = { x: C.CONTENT_X, y: C.HEADER_Y, w: C.CONTENT_W, h: C.CONTENT_BOTTOM - C.HEADER_Y };

  // The number is rendered by PowerPoint's bullet, so the text is the title
  // alone. Width is discounted to leave room for the bullet's own indent.
  const textW = area.w - SUMMARY_BULLET_INDENT;
  const built = n.items.map((it) => {
    const text = it.item_name ? `${it.item_name} — ${it.title}` : it.title;
    const lines = wrap(text, textW, entrySize);
    return { n: it.number, text, lines, h: heightOf(lines.length, entrySize) + ENTRY_GAP };
  });

  const tailH = heightOf(1, C.TYPE.body.size) + ENTRY_GAP + 0.18;

  const slides = [];
  let idx = 0;

  while (idx < built.length) {
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
    while (idx + take < built.length && take < C.CONTENTS_MAX_ENTRIES) {
      if (probe + built[idx + take].h > bottom - tailH) break;
      probe += built[idx + take].h; take += 1;
    }

    if (idx + take < built.length) {
      // Not closing here — reclaim the tail's space for more entries.
      take = 0; probe = cursor;
      while (idx + take < built.length && take < C.CONTENTS_MAX_ENTRIES) {
        if (probe + built[idx + take].h > bottom) break;
        probe += built[idx + take].h; take += 1;
      }
      if (take === 0) take = 1;   // always make progress
    }

    const entries = built.slice(idx, idx + take);
    const listH = entries.reduce((a, e) => a + e.h, 0);

    // One box for the whole slice. numberStartAt keeps the count running across
    // paginated summary slides.
    const list = {
      startAt: entries[0].n,
      entries: entries.map((e) => ({ n: e.n, text: e.text, lines: e.lines })),
      size: entrySize,
      indent: SUMMARY_BULLET_INDENT,
      box: { x: area.x, y: cursor, w: area.w, h: listH },
    };
    cursor += listH;
    idx += take;

    let tail = null;
    if (idx >= built.length) {
      const size = C.TYPE.body.size;
      tail = [{
        kind: 'instruction', text: C.REPLY_INSTRUCTION, size,
        box: { x: area.x, y: cursor + 0.18, w: area.w, h: heightOf(1, size) },
      }];
    }

    const summaryLogo = logoBox(n.document.logo, C.FOOTER_LOGO_H, false);
    slides.push({ kind: 'contents', heading: headingBlock, list, tail,
      logo: summaryLogo, footer: itemFooter(n.document, Boolean(summaryLogo)) });
  }

  return slides;
}

// ── Item slides (§8–§11) ─────────────────────────────────────

function buildItemSlide(item, attempt, { continued, isLast }) {
  const header = planHeader({
    number: item.number, title: item.title, continued, itemName: item.item_name,
  });

  return {
    kind: 'item',
    itemId: item.id,
    number: item.number,
    itemName: item.item_name,
    title: item.title,
    continued,
    header,
    layout: attempt.layout,
    body: attempt.body,
    images: attempt.images,
    placeholder: attempt.placeholder || null,
    // Told at the foot of the slide, not only in the next slide's header — a
    // reader needs to know there is more before they turn the page, not after.
    continuesNote: isLast ? null : {
      text: C.CONTINUES_LABEL,
      // Directly below whatever content area this slide actually used — the
      // strip layoutItemSlide reserved when it was told the slide continues.
      box: {
        x: C.CONTENT_X,
        y: attempt.content.y + attempt.content.h,
        w: C.CONTENT_W,
        h: C.CONTINUES_H,
      },
    },
    // On every item, not only those flagged as needing a response. Uniformity is
    // what lets a sender duplicate any slide and have it look right (PLAN.md §5.14).
    replyBox: isLast
      ? {
          label: C.REPLY_LABEL(item.number),
          box: { x: C.CONTENT_X, y: C.REPLY_Y, w: C.CONTENT_W, h: C.REPLY_H },
          labelBox: {
            x: C.CONTENT_X + C.REPLY_INSET, y: C.REPLY_Y + C.REPLY_INSET,
            w: C.CONTENT_W - C.REPLY_INSET * 2,
            // Fills the box, so clicking anywhere in the reply area lands the
            // cursor in a text box the full size of the space provided.
            h: C.REPLY_H - C.REPLY_INSET * 2,
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
      number: item.number, title: item.title, continued, itemName: item.item_name,
    }).height;
    // Reserve the image column only when the item has no pictures at all.
    const wantPlaceholder = item.images.length === 0;

    const trial = layoutItemSlide({ body, images, headerH, hasReply: true, wantPlaceholder });
    if (trial.bodyRemainder === '' && trial.imagesUsed === images.length) {
      slides.push(buildItemSlide(item, trial, { continued, isLast: true }));
      break;
    }

    // Not the last slide, so it carries on overleaf and must say so — which
    // costs a strip at the foot, hence the re-layout rather than reusing trial.
    let chosen = layoutItemSlide({ body, images, headerH, hasReply: false, continues: true, wantPlaceholder });

    // If the full-height layout would swallow everything, there'd be nothing
    // left for the reply box's own slide. Keep a split instead so the remainder
    // carries the reply box on a final slide.
    if (chosen.bodyRemainder === '' && chosen.imagesUsed === images.length) {
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

/**
 * Item slides, in the order the sender wrote them.
 *
 * No partitioning, no reordering, no grouping. Queries and updates are one
 * numbered stream now, so a slide inserted by hand lands where it was put and
 * carries the next number — which section-ordering and grouping both prevented
 * (PLAN.md §5.14).
 */
function planItemSlides(n, warnings) {
  const slides = [];
  for (const item of n.items) slides.push(...planItem(item, warnings));
  return slides;
}

// ── Template slides ──────────────────────────────────────────

/**
 * Two blank slides appended for the sender to copy.
 *
 * Their boxes use the same geometry as generated slides, so a copied template
 * filled in by hand is indistinguishable from one the generator produced. The
 * renderer wires these to real PowerPoint placeholders, which show a
 * "click to add" prompt while editing and print nothing if left untouched —
 * so shipping them costs nothing when a sender ignores them.
 */
function planTemplates(n, startNumber) {
  const headerH = C.HEADER_H;
  const c = contentBox(headerH, true);
  const colW = c.w - C.GUTTER;

  const chrome = (number) => {
    const header = planHeader({ number, title: '', continued: false });
    return {
      kind: 'template',
      number,
      header,
      replyBox: {
        label: C.REPLY_LABEL(number),
        box: { x: C.CONTENT_X, y: C.REPLY_Y, w: C.CONTENT_W, h: C.REPLY_H },
        labelBox: {
          x: C.CONTENT_X + C.REPLY_INSET, y: C.REPLY_Y + C.REPLY_INSET,
          w: C.CONTENT_W - C.REPLY_INSET * 2,
          // Fills the box rather than sitting as a one-line strip at the top, so
          // clicking anywhere in the reply area lands the cursor in a text box
          // the full size of the space provided.
          h: C.REPLY_H - C.REPLY_INSET * 2,
        },
      },
      logo: logoBox(n.document.logo, C.FOOTER_LOGO_H, false),
    };
  };

  // A — one picture beside the text, matching IMAGE_SIDE.
  const imgW = colW * C.SIDE_IMAGE_FRACTION;
  const single = {
    ...chrome(startNumber),
    variant: 'single',
    titleBox: { ...c, y: C.HEADER_Y, h: C.HEADER_H },
    images: [{ x: c.x, y: c.y, w: imgW, h: c.h }],
    bodyBox: { x: c.x + imgW + C.GUTTER, y: c.y, w: colW - imgW, h: c.h },
  };

  // B — three pictures in the grid, body above, matching IMAGE_GRID. The
  // one-nudge-per-page rule is about generated item slides; a template is a
  // worked example, so it shows the multi-image arrangement.
  const bodyH = heightOf(C.GRID_BODY_MAX_LINES, C.TYPE.body.size);
  const gridY = c.y + bodyH + C.GUTTER;
  const gridH = c.h - (gridY - c.y);
  const cellW = (c.w - C.GUTTER) / 2;
  const rowH = (gridH - C.GUTTER) / 2;
  const multi = {
    ...chrome(startNumber + 1),
    variant: 'multi',
    titleBox: { ...c, y: C.HEADER_Y, h: C.HEADER_H },
    bodyBox: { x: c.x, y: c.y, w: c.w, h: bodyH },
    images: [
      { x: c.x, y: gridY, w: cellW, h: rowH },
      { x: c.x + cellW + C.GUTTER, y: gridY, w: cellW, h: rowH },
      { x: c.x + (c.w - cellW) / 2, y: gridY + rowH + C.GUTTER, w: cellW, h: rowH },
    ],
  };

  return [single, multi];
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
  slides.push(...planTemplates(n, n.items.length + 1));

  // Page numbers and footers, once the sequence is final.
  slides.forEach((s, i) => {
    s.page = i + 1;
    if (s.kind === 'item' || s.kind === 'template') {
      s.logo = s.logo || logoBox(n.document.logo, C.FOOTER_LOGO_H, false);
      s.footer = itemFooter(n.document, Boolean(s.logo));
    }
    if (s.kind !== 'cover') s.pageLabel = `Page ${i + 1}`;
  });

  return {
    meta: {
      totalSlides: slides.length,
      itemCount: n.itemCount,
      // Still surfaced for the notification email, though the deck itself no
      // longer distinguishes the two.
      queryCount: n.queryCount,
      updateCount: n.updateCount,
      hasContents: showContents,
      warnings,
    },
    document: n.document,
    slides,
  };
}

module.exports = { planSlides, normalize, formatDate, resolveTitle, planTemplates };
