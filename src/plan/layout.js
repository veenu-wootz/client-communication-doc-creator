/**
 * layout.js — geometry for a single item slide.
 *
 * Pure. Given the content still to place and the space available, decide which
 * §9 layout applies and where every box goes. Reports what it could not place;
 * the planner puts that on the next slide. Nothing is ever discarded here.
 */

const C = require('./constants');
const { wrap, splitToFit, heightOf, measure } = require('./textFit');

const CAPTION_GAP = 0.06;
const HEADER_GAP  = 0.25;   // between the header block and the content area
const CHIP_PAD_X  = 0.17;   // horizontal padding inside the Q-number chip
const CHIP_GAP    = 0.18;   // between the chip and the title

/**
 * Contain-fit: scale to fit the box, preserving aspect. Never stretches.
 *
 * imgW/imgH are pixels and boxW/boxH are inches, so `scale` is an
 * inches-per-pixel factor — an oversized image is always shrunk to fit, which
 * is the common case and has always worked.
 *
 * The floor is the other direction: a small image would otherwise be blown up
 * to fill whatever box it is given. A 200px screenshot across a 5in column
 * lands at ~40 DPI, and an unreadable drawing defeats the point of sending it.
 * MIN_RENDER_DPI caps the physical size so a low-resolution source renders
 * smaller and sharp rather than large and blurred.
 */
function fitBox(imgW, imgH, boxW, boxH) {
  if (!imgW || !imgH || boxW <= 0 || boxH <= 0) return { w: Math.max(0, boxW), h: Math.max(0, boxH) };
  const contain = Math.min(boxW / imgW, boxH / imgH);
  const scale = Math.min(contain, 1 / C.MIN_RENDER_DPI);
  return { w: imgW * scale, h: imgH * scale };
}

/**
 * The header. Wraps rather than truncates (PLAN.md §5.1); if two lines at the
 * header size is not enough, the size steps down before a third line is
 * allowed, and the content area moves down to make room. The Q-number is never
 * shrunk or wrapped.
 */
function planHeader({ queryNo, title, continued, itemName }) {
  const text = `${title}${continued ? ' (continued)' : ''}`;
  const baseSize = C.TYPE.slideHeader.size;

  // Every item slide is labelled. A query shows its number; an update says so.
  // Without this an unnumbered slide reads as an oversight to a client who has
  // never seen the format and has nobody to ask.
  const tone = queryNo ? 'query' : 'update';
  const chipText = queryNo ? `Q${queryNo}` : C.UPDATE_CHIP_LABEL;
  const chipSize = C.CHIP[tone].size;

  // item_name sits above the header as a muted line. One document often covers
  // several SKUs, so the reader needs to know which part a query is about
  // before they read the question.
  const name = (itemName || '').trim();
  const eyebrow = name
    ? {
        text: name,
        size: C.TYPE.eyebrow.size,
        box: { x: C.CONTENT_X, y: C.HEADER_Y, w: C.CONTENT_W, h: C.HEADER_EYEBROW_H },
      }
    : null;
  const topY = C.HEADER_Y + (eyebrow ? C.HEADER_EYEBROW_H : 0);

  const chipW = measure(chipText, chipSize, true) + CHIP_PAD_X * 2;
  const titleX = C.CONTENT_X + chipW + CHIP_GAP;
  const titleW = C.CONTENT_W - chipW - CHIP_GAP;

  let size = baseSize;
  let lines = wrap(text, titleW, size, true);
  while (lines.length > C.HEADER_MAX_LINES && size > 16) {
    size -= 2;
    lines = wrap(text, titleW, size, true);
  }

  const titleH = lines.length <= 1 ? C.HEADER_H
    : lines.length === 2 ? C.HEADER_H_2LINE
    : heightOf(lines.length, size) + 0.33;

  return {
    eyebrow,
    chip: {
      text: chipText, tone, size: chipSize,
      box: { x: C.CONTENT_X, y: topY, w: chipW, h: C.HEADER_H },
    },
    title: { text, lines, size, box: { x: titleX, y: topY, w: titleW, h: titleH } },
    height: titleH + (eyebrow ? C.HEADER_EYEBROW_H : 0),
  };
}

/** The area available for content, given the header height and reply box. */
function contentBox(headerH, hasReply) {
  const y = C.HEADER_Y + headerH + HEADER_GAP;
  const bottom = hasReply ? C.REPLY_Y - C.REPLY_GAP : C.CONTENT_BOTTOM;
  return { x: C.CONTENT_X, y, w: C.CONTENT_W, h: Math.max(0.5, bottom - y) };
}

/** §9 layout selection. `hasBody` false uses the stacked variants (§13). */
function selectLayout(images, hasBody) {
  const n = images.length;
  if (n === 0) return 'TEXT_ONLY';
  const isWide = (im) => (im.aspect || 1) >= C.WIDE_ASPECT_THRESHOLD;

  if (n === 1) return (!hasBody || isWide(images[0])) ? 'IMAGE_STACKED' : 'IMAGE_SIDE';
  if (n === 2) return (!hasBody || images.some(isWide)) ? 'IMAGE_STACKED_PAIR' : 'IMAGE_SIDE_PAIR';
  return 'IMAGE_GRID';
}

/** Place one image plus its caption inside a cell, top-aligned, centred. */
function imageInCell(img, cell) {
  const capLines = img.caption ? wrap(img.caption, cell.w, C.TYPE.caption.size) : [];
  const capH = capLines.length ? heightOf(capLines.length, C.TYPE.caption.size) + CAPTION_GAP : 0;
  const { w, h } = fitBox(img.width, img.height, cell.w, Math.max(0.2, cell.h - capH));

  return {
    image: img,
    missing: !img.ok,
    box: { x: cell.x + (cell.w - w) / 2, y: cell.y, w, h },
    caption: capLines.length
      ? {
          text: img.caption,
          lines: capLines,
          box: { x: cell.x, y: cell.y + h + CAPTION_GAP, w: cell.w, h: capH },
        }
      : null,
    usedH: h + capH,
  };
}

/**
 * Try one concrete layout at one body size.
 *
 * Returns null when the arrangement is not viable — the cell an image would go
 * into is narrower than MIN_IMAGE_W, or the grid body needs more room than the
 * grid allows. A null tells the caller to try with fewer images on this slide.
 *
 * The floor applies to the CELL, not to the rendered image: a square photo in a
 * wide cell is short because of its own aspect, not because the layout squeezed
 * it, and testing the rendered width would make the 3-4 image grid unreachable
 * for anything but wide drawings.
 */
function tryLayout(layoutName, images, body, size, c) {
  const hasBody = Boolean(body && body.trim());
  const placed = [];
  let bodyBox = null;
  let bodySizedToText = false;
  let cellW = c.w;   // width of the cell each image is laid into

  if (layoutName === 'TEXT_ONLY') {
    bodyBox = { ...c };

  } else if (layoutName === 'IMAGE_SIDE' || layoutName === 'IMAGE_SIDE_PAIR') {
    const colW = c.w - C.GUTTER;
    const pair = layoutName === 'IMAGE_SIDE_PAIR';
    // A fixed split, deliberately. Widening the image column was tried and
    // reverted: IMAGE_SIDE only receives images with aspect < WIDE_ASPECT_THRESHOLD,
    // and in a column this tall those are already height-bound, so a wider
    // column changes nothing for anything below about 1.3 aspect.
    const imgW = colW * (pair ? C.SIDE_PAIR_IMAGE_FRACTION : C.SIDE_IMAGE_FRACTION);
    const bodyW = colW - imgW;

    cellW = pair ? (imgW - C.GUTTER) / 2 : imgW;
    if (images.length > 1 && cellW < C.MIN_IMAGE_W) return null;

    images.forEach((im, i) => {
      placed.push(imageInCell(im, { x: c.x + i * (cellW + C.GUTTER), y: c.y, w: cellW, h: c.h }));
    });
    bodyBox = { x: c.x + imgW + C.GUTTER, y: c.y, w: bodyW, h: c.h };

  } else if (layoutName === 'IMAGE_STACKED' || layoutName === 'IMAGE_STACKED_PAIR') {
    // Give the image zone whatever the body does not need, between a floor and
    // a ceiling. A wide drawing capped at a flat fraction ends up barely wider
    // than the side column, which defeats the point of the stacked layout.
    let zoneH = c.h;
    if (hasBody) {
      const bodyH = heightOf(wrap(body, c.w, size).length, size);
      const minBodyH = heightOf(2, size);
      const want = c.h - bodyH - C.GUTTER;
      const ceiling = Math.min(c.h * C.IMAGE_ZONE_MAX, c.h - minBodyH - C.GUTTER);
      zoneH = Math.max(c.h * C.IMAGE_ZONE_FRACTION, Math.min(want, ceiling));
    }

    if (layoutName === 'IMAGE_STACKED') {
      placed.push(imageInCell(images[0], { x: c.x, y: c.y, w: c.w, h: zoneH }));
    } else {
      const cellH = (zoneH - C.GUTTER) / 2;
      images.forEach((im, i) => {
        placed.push(imageInCell(im, { x: c.x, y: c.y + i * (cellH + C.GUTTER), w: c.w, h: cellH }));
      });
    }

    const usedH = Math.max(...placed.map((p) => p.box.y + p.usedH)) - c.y;
    const rest = c.h - usedH - C.GUTTER;
    if (hasBody && rest < heightOf(1, size)) return null;
    bodyBox = hasBody ? { x: c.x, y: c.y + usedH + C.GUTTER, w: c.w, h: rest } : null;

  } else if (layoutName === 'IMAGE_GRID') {
    // Two rows only. A fifth image belongs on the next slide, not on a third
    // row that would run past the content area.
    if (images.length > C.GRID_MAX_CELLS) return null;

    // Body sits above the grid. If it needs more than GRID_BODY_MAX_LINES the
    // body wins the slide and the caller retries with fewer images, rather than
    // the body being cut (PLAN.md §5.1).
    const lines = hasBody ? wrap(body, c.w, size) : [];
    if (lines.length > C.GRID_BODY_MAX_LINES) return null;

    const bodyH = lines.length ? heightOf(lines.length, size) : 0;
    const gridY = c.y + (bodyH ? bodyH + C.GUTTER : 0);
    const gridH = c.h - (gridY - c.y);
    cellW = (c.w - C.GUTTER) / 2;
    const rowH = (gridH - C.GUTTER) / 2;

    if (cellW < C.MIN_IMAGE_W || rowH < 0.6) return null;

    // 3 images: two on the top row, one centred beneath at the same cell width.
    const slots = images.length === 3
      ? [
          { x: c.x, y: gridY },
          { x: c.x + cellW + C.GUTTER, y: gridY },
          { x: c.x + (c.w - cellW) / 2, y: gridY + rowH + C.GUTTER },
        ]
      : images.map((_, i) => ({
          x: c.x + (i % 2) * (cellW + C.GUTTER),
          y: gridY + Math.floor(i / 2) * (rowH + C.GUTTER),
        }));

    images.forEach((im, i) => {
      placed.push(imageInCell(im, { x: slots[i].x, y: slots[i].y, w: cellW, h: rowH }));
    });

    bodyBox = bodyH ? { x: c.x, y: c.y, w: c.w, h: bodyH } : null;
    bodySizedToText = true;
  }

  // Split the body against whatever box the layout left for it.
  let fitted = [], remainder = '';
  if (hasBody) {
    if (!bodyBox) return null;
    const s = splitToFit(body, bodyBox.w, bodyBox.h, size, false,
      bodySizedToText ? { fill: 1 } : undefined);
    fitted = s.fitted;
    remainder = s.remainder;
    if (fitted.length === 0) return null;   // no room at all — try fewer images
  }

  return {
    layout: layoutName,
    images: placed,
    body: hasBody ? { lines: fitted, size, box: bodyBox } : null,
    bodyRemainder: remainder,
  };
}

/**
 * Lay out one item slide.
 *
 * Prefers to place every remaining image; steps the image count down until the
 * whole body fits alongside them. Never goes below one image while images
 * remain, so a query always has visual context on its own slide (spec §11).
 * Whatever does not fit is returned as a remainder, never dropped.
 */
function layoutItemSlide({ body, images, headerH, hasReply }) {
  const c = contentBox(headerH, hasReply);
  const hasBody = Boolean(body && body.trim());
  const sizes = [];
  for (let s = C.BODY_SIZE_MAX; s >= C.BODY_SIZE_MIN; s -= C.BODY_SIZE_STEP) sizes.push(s);

  if (images.length === 0) {
    for (const size of sizes) {
      const att = tryLayout('TEXT_ONLY', [], body, size, c);
      if (att && att.bodyRemainder === '') return { ...att, imagesUsed: 0, content: c };
    }
    const att = tryLayout('TEXT_ONLY', [], body, C.BODY_SIZE_MIN, c)
      || { layout: 'TEXT_ONLY', images: [], body: null, bodyRemainder: '' };
    return { ...att, imagesUsed: 0, content: c };
  }

  let fallback = null;
  for (let n = images.length; n >= 1; n--) {
    const subset = images.slice(0, n);
    const name = selectLayout(subset, hasBody);
    for (const size of sizes) {
      const att = tryLayout(name, subset, body, size, c);
      if (!att) continue;
      if (att.bodyRemainder === '') return { ...att, imagesUsed: n, content: c };
      fallback = { ...att, imagesUsed: n, content: c };   // keep the last (fewest images,
    }                                                     // smallest type = most body room)
  }

  if (fallback) return fallback;

  // Nothing was viable — place the first image alone and carry the body.
  const solo = tryLayout('IMAGE_STACKED', images.slice(0, 1), '', C.BODY_SIZE_MIN, c);
  return { ...solo, bodyRemainder: body || '', imagesUsed: 1, content: c };
}

module.exports = { fitBox, planHeader, contentBox, selectLayout, tryLayout, layoutItemSlide };
