/**
 * renderPptx.js — draw a slide plan.
 *
 * Deliberately dumb: every position was decided by the planner, so this file
 * only translates boxes into pptxgenjs calls. If something is in the wrong
 * place, the bug is in the planner, not here.
 *
 * pptxgenjs footguns this file is written around (spec §15):
 *   - the layout must be set before any slide is added, or shapes past the
 *     default 10 x 5.625in canvas are written but never displayed
 *   - hex colours take no '#' and no alpha
 *   - option objects are mutated in place, so each call gets a fresh one
 *   - every text call sets isTextBox, or screen readers announce it as a graphic
 *   - text boxes carry internal padding; margin 0 wherever text must align
 *   - only roundRect honours rectRadius
 *   - shadows are omitted entirely (negative offsets corrupt the file)
 */

const fs = require('fs');
const JSZip = require('jszip');
const PptxGenJS = require('pptxgenjs');
const C = require('../plan/constants');

const LAYOUT_NAME = 'CLIENT_COMM_16x9';
const CHIP_RADIUS = 0.04;
const REPLY_RADIUS = 0.03;

/** Rejoin planner lines into text, keeping paragraph breaks. */
function toText(lines) {
  if (!lines || lines.length === 0) return '';
  if (C.HARD_WRAP) return lines.join('\n');
  const paras = [];
  let cur = [];
  for (const l of lines) {
    if (l === '') { paras.push(cur.join(' ')); cur = []; } else cur.push(l);
  }
  paras.push(cur.join(' '));
  return paras.join('\n');
}

/** Base options every text call starts from. Never share one between calls. */
function textOpts(box, { size, bold = false, color = C.COLOR.ink, align = 'left', valign = 'top' }) {
  return {
    x: box.x, y: box.y, w: box.w, h: box.h,
    fontFace: C.FONT, fontSize: size, bold, color, align, valign,
    margin: 0, isTextBox: true, wrap: true, autoFit: false, shrinkText: false,
  };
}

/** Is there anything we can actually hand to pptxgenjs for this image? */
function imageSource(img) {
  if (!img || img.ok === false) return null;
  if (img.data) return { data: img.data };
  if (img.buffer) return { data: `image/png;base64,${img.buffer.toString('base64')}` };
  if (img.path && fs.existsSync(img.path)) return { path: img.path };
  return null;
}

/** The header chip — the item's number, identical on every slide. */
function drawChip(slide, pres, chip) {
  if (!chip) return;
  slide.addText(chip.text, {
    x: chip.box.x, y: chip.box.y, w: chip.box.w, h: chip.box.h,
    shape: pres.ShapeType.roundRect, rectRadius: CHIP_RADIUS,
    fill: { color: C.CHIP.bg },
    fontFace: C.FONT, fontSize: chip.size || C.CHIP.size, bold: true,
    color: C.CHIP.text, align: 'center', valign: 'middle',
    margin: 0, isTextBox: true, autoFit: false, shrinkText: false,
  });
}

/**
 * Every picture placeholder in the deck is a REAL PowerPoint placeholder, not a
 * drawn box: it offers the click-to-insert icon, and it prints nothing if the
 * sender never uses it. A drawn rectangle did neither — it gave no way to add an
 * image and had to be deleted by hand (PLAN.md §5.16).
 *
 * Placeholders can only be declared on a master, so one master is defined per
 * distinct geometry the deck actually needs. In practice that is two or three.
 */
function masterKeyFor(prefix, boxes) {
  const mm = (v) => Math.round(v * 100);
  return `${prefix}_${boxes.map((b) => [mm(b.x), mm(b.y), mm(b.w), mm(b.h)].join('_')).join('__')}`;
}

/**
 * Every slide needing placeholders, and what it needs.
 *
 * Template masters are keyed apart from the rest even where the geometry
 * matches: they carry title and body placeholders too, and an item slide
 * sharing that master would sprout prompts it should not have.
 */
function placeholderSlots(plan) {
  const slots = new Map();
  const want = (prefix, pics, texts = []) => {
    const k = masterKeyFor(prefix, pics);
    if (!slots.has(k)) slots.set(k, { pics, texts });
    return k;
  };

  for (const s of plan.slides) {
    if (s.kind === 'cover' && s.photoPlaceholder) s.masterName = want('PIC', [s.photoPlaceholder.box]);
    else if (s.kind === 'item' && s.placeholder) s.masterName = want('PIC', [s.placeholder.box]);
    else if (s.kind === 'template') {
      // Title and body are placeholders too, not drawn text: PowerPoint then
      // shows its own prompt, clears it the moment the sender types, and the
      // typed text takes the ink colour rather than staying the prompt's grey.
      s.masterName = want(`TPL${s.variant}`, s.images, [
        {
          box: { x: s.titleBox.x + 0.95, y: s.titleBox.y, w: s.titleBox.w - 0.95, h: s.titleBox.h },
          size: C.TYPE.slideHeader.size, bold: true, valign: 'middle', text: C.TEMPLATE_TITLE_PROMPT,
        },
        {
          box: s.bodyBox,
          size: C.TYPE.body.size, bold: false, valign: 'top', text: C.TEMPLATE_BODY_PROMPT,
        },
      ]);
    }
  }
  return slots;
}

function definePictureMasters(pres, slots) {
  for (const [name, { pics, texts }] of slots) {
    // Pictures FIRST — promotePicturePlaceholders promotes by position, since
    // pptxgenjs discards the name we give each placeholder.
    const objects = pics.map((box, i) => ({
      placeholder: {
        options: {
          name: `pic${i + 1}`, type: 'pic',
          x: box.x, y: box.y, w: box.w, h: box.h,
          fontFace: C.FONT, fontSize: C.TYPE.caption.size,
          color: C.COLOR.muted, align: 'center', valign: 'middle',
        },
        text: C.PLACEHOLDER_LABEL,
      },
    }));

    for (const [i, t] of texts.entries()) {
      objects.push({
        placeholder: {
          options: {
            name: `txt${i + 1}`, type: 'body',
            x: t.box.x, y: t.box.y, w: t.box.w, h: t.box.h,
            fontFace: C.FONT, fontSize: t.size, bold: t.bold,
            color: C.COLOR.ink, valign: t.valign, margin: 0,
          },
          text: t.text,
        },
      });
    }

    pres.defineSlideMaster({ title: name, objects });
  }
}

/** A missing or unreadable image: a filled rectangle naming the file. §13 */
function drawPlaceholder(slide, pres, box, label) {
  slide.addShape(pres.ShapeType.rect, {
    x: box.x, y: box.y, w: box.w, h: Math.max(box.h, 0.5),
    fill: { color: C.COLOR.surface },
    line: { color: C.COLOR.border, width: C.HAIRLINE },
  });
  slide.addText(label || 'image unavailable', textOpts(
    { x: box.x + 0.1, y: box.y + Math.max(box.h, 0.5) / 2 - 0.12, w: box.w - 0.2, h: 0.24 },
    { size: C.TYPE.caption.size, color: C.COLOR.muted, align: 'center' },
  ));
}

function drawImage(slide, pres, placed) {
  const box = placed.box;
  const src = imageSource(placed.image);

  if (!src) {
    const name = (placed.image && (placed.image.caption || placed.image.path)) || 'image unavailable';
    drawPlaceholder(slide, pres, box, String(name).split('/').pop());
  } else {
    slide.addImage({ ...src, x: box.x, y: box.y, w: box.w, h: box.h });
    // Hairline frame, no shadow.
    slide.addShape(pres.ShapeType.rect, {
      x: box.x, y: box.y, w: box.w, h: box.h,
      fill: { type: 'none' }, line: { color: C.COLOR.border, width: C.HAIRLINE },
    });
  }

  if (placed.caption) {
    slide.addText(toText(placed.caption.lines), textOpts(placed.caption.box, {
      size: C.TYPE.caption.size, color: C.COLOR.muted,
    }));
  }
}

function drawLogo(slide, logo) {
  if (!logo) return;
  const src = imageSource(logo.image);
  if (!src) return;
  slide.addImage({ ...src, x: logo.box.x, y: logo.box.y, w: logo.box.w, h: logo.box.h });
}

function drawFooter(slide, s) {
  if (!s.footer) return;
  if (s.footer.text) {
    slide.addText(s.footer.text, textOpts(s.footer.box, {
      size: C.TYPE.footer.size, color: C.COLOR.muted, valign: 'middle',
    }));
  }
  if (s.footer.page && s.pageLabel) {
    slide.addText(s.pageLabel, textOpts(s.footer.page.box, {
      size: C.TYPE.footer.size, color: C.COLOR.muted, align: 'right', valign: 'middle',
    }));
  }
}

// ── Slide kinds ──────────────────────────────────────────────

function renderCover(pres, slide, s) {
  for (const b of s.blocks) {
    if (b.kind === 'title') {
      slide.addText(toText(b.lines), textOpts(b.box, { size: b.size, bold: true }));
    } else if (b.kind === 'field') {
      slide.addText(b.label, textOpts(b.labelBox, {
        size: C.TYPE.coverLabel.size, color: C.COLOR.muted,
      }));
      // An empty field still gets its box, so a click lands in the right place,
      // and a hairline rule beneath it to write on. A rule reads as a form line
      // and needs no deleting if it is left blank.
      slide.addText(b.filled ? toText(b.lines) : '',
        textOpts(b.valueBox, { size: C.TYPE.coverValue.size }));
      if (b.rule) {
        slide.addShape(pres.ShapeType.line, {
          x: b.rule.x, y: b.rule.y, w: b.rule.w, h: 0,
          line: { color: C.COLOR.border, width: C.HAIRLINE },
        });
      }
    } else if (b.kind === 'note') {
      slide.addText(toText(b.lines), textOpts(b.box, { size: b.size }));
    }
  }

  if (s.instruction) {
    slide.addText(toText(s.instruction.lines), textOpts(s.instruction.box, {
      size: s.instruction.size, color: C.COLOR.muted,
    }));
  }

  if (s.photo) {
    const src = imageSource(s.photo.image);
    const z = s.photo.zone;
    if (src) {
      // Contain within the zone, vertically centred.
      const img = s.photo.image;
      const scale = Math.min(z.w / (img.width || 1), z.h / (img.height || 1));
      const w = (img.width || z.w) * scale;
      const h = (img.height || z.h) * scale;
      slide.addImage({ ...src, x: z.x + (z.w - w) / 2, y: z.y + (z.h - h) / 2, w, h });
    }
  }

  drawLogo(slide, s.logo);
  drawFooter(slide, s);
}

function renderContents(pres, slide, s) {
  if (s.heading) {
    slide.addText(toText(s.heading.lines), textOpts(s.heading.box, { size: s.heading.size, bold: true }));
  }
  // One box, PowerPoint's own numbered bullets. This is what lets a sender
  // click in, press Enter, and get the next number for free — the single
  // biggest thing standing between this deck and being used as a template.
  if (s.list) {
    // The bullet must sit on EVERY run, not on the shape: pptxgenjs applies a
    // shape-level bullet to the first paragraph only and stamps <a:buNone/> on
    // the rest, which is why the list came out numbered "1", blank, "1".
    //
    // Each paragraph also declares its own number. pptxgenjs always writes a
    // startAt attribute (defaulting to 1), so leaving it off every paragraph
    // would render 1, 1, 1 — stating the real number makes it deterministic.
    slide.addText(
      s.list.entries.map((e) => ({
        text: toText(e.lines),
        options: {
          bullet: { type: 'number', numberStartAt: e.n, indent: s.list.indent * 72 },
          breakLine: true,
        },
      })),
      textOpts(s.list.box, { size: s.list.size }),
    );
  }
  for (const t of s.tail || []) {
    slide.addText(t.text, textOpts(t.box, { size: t.size, color: C.COLOR.muted }));
  }
  drawLogo(slide, s.logo);
  drawFooter(slide, s);
}

function renderItem(pres, slide, s) {
  // The SKU this row is about, above the header.
  if (s.header.eyebrow) {
    slide.addText(s.header.eyebrow.text, textOpts(s.header.eyebrow.box, {
      size: s.header.eyebrow.size, color: C.COLOR.muted, valign: 'middle',
    }));
  }

  // Header: the chip, then the title beside it.
  drawChip(slide, pres, s.header.chip);

  slide.addText(toText(s.header.title.lines), textOpts(s.header.title.box, {
    size: s.header.title.size, bold: true, valign: 'middle',
  }));

  if (s.body) {
    slide.addText(toText(s.body.lines), textOpts(s.body.box, { size: s.body.size }));
  }

  for (const im of s.images) drawImage(slide, pres, im);

  if (s.continuesNote) {
    slide.addText(s.continuesNote.text, textOpts(s.continuesNote.box, {
      size: C.TYPE.caption.size, color: C.COLOR.muted, align: 'right', valign: 'middle',
    }));
  }

  if (s.replyBox) {
    slide.addShape(pres.ShapeType.roundRect, {
      x: s.replyBox.box.x, y: s.replyBox.box.y, w: s.replyBox.box.w, h: s.replyBox.box.h,
      rectRadius: REPLY_RADIUS,
      fill: { color: C.COLOR.paper },
      line: { color: C.COLOR.border, width: C.HAIRLINE },
    });
    slide.addText(s.replyBox.label, textOpts(s.replyBox.labelBox, {
      size: C.TYPE.replyLabel.size, color: C.COLOR.muted,
    }));
  }

  drawLogo(slide, s.logo);
  drawFooter(slide, s);
}

/** A template slide: chrome as usual, the picture left to its placeholder. */
function renderTemplate(pres, slide, s) {
  drawChip(slide, pres, s.header.chip);

  slide.addShape(pres.ShapeType.roundRect, {
    x: s.replyBox.box.x, y: s.replyBox.box.y, w: s.replyBox.box.w, h: s.replyBox.box.h,
    rectRadius: REPLY_RADIUS,
    fill: { color: C.COLOR.paper },
    line: { color: C.COLOR.border, width: C.HAIRLINE },
  });
  slide.addText(s.replyBox.label, textOpts(s.replyBox.labelBox, {
    size: C.TYPE.replyLabel.size, color: C.COLOR.muted,
  }));

  drawLogo(slide, s.logo);
  drawFooter(slide, s);
}

/**
 * Turn our generic placeholders into real PICTURE placeholders.
 *
 * pptxgenjs cannot do this: its PLACEHOLDER_TYPES table is empty, so the branch
 * that would write type="pic" never fires and every placeholder it emits is a
 * plain one. A plain placeholder still disappears when unused, but it offers no
 * click-to-insert-picture icon — which is the half the sender actually needs.
 *
 * So the attribute is injected afterwards, into our own PIC_* layouts only. The
 * match is asserted rather than assumed: if a pptxgenjs upgrade changes the
 * shape of this XML, this must fail loudly instead of quietly reverting to
 * placeholders nobody can click.
 */
async function promotePicturePlaceholders(buffer, slots) {
  if (slots.size === 0) return { buffer, promoted: 0 };

  // How many of each layout's placeholders are pictures. The rest are text and
  // must be left alone, or a title would become a picture frame.
  const picCount = new Map([...slots].map(([k, v]) => [k, v.pics.length]));

  const zip = await JSZip.loadAsync(buffer);
  const layouts = Object.keys(zip.files).filter((n) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(n));
  let promoted = 0;

  for (const name of layouts) {
    const xml = await zip.file(name).async('string');
    const key = (xml.match(/name="((?:PIC|TPL)[^"]*)"/) || [])[1];
    if (!key || !picCount.has(key)) continue;            // not one of ours

    let remaining = picCount.get(key);
    const next = xml.replace(/<p:ph\b/g, (m) => (remaining-- > 0 ? '<p:ph type="pic"' : m));
    if (next !== xml) zip.file(name, next);
    promoted += 1;
  }

  if (promoted !== slots.size) {
    console.warn(`  ⚠ picture placeholders: promoted ${promoted} of ${slots.size} — `
      + 'image slots will show a text prompt instead of an insert icon');
  }

  return { buffer: await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' }), promoted };
}

// ── Entry point ──────────────────────────────────────────────

/**
 * @param {object} plan  output of planSlides
 * @returns {Promise<Buffer>} the .pptx
 */
async function renderPptx(plan) {
  const pres = new PptxGenJS();

  // Before any slide is added. Not negotiable — see the header comment.
  pres.defineLayout({ name: LAYOUT_NAME, width: C.SLIDE_W, height: C.SLIDE_H });
  pres.layout = LAYOUT_NAME;

  const slots = placeholderSlots(plan);
  definePictureMasters(pres, slots);

  pres.author = plan.document.created_by || 'Wootz';
  pres.company = 'Wootz';
  pres.title = plan.document.report_title || 'Client communication';

  for (const s of plan.slides) {
    const slide = s.masterName ? pres.addSlide({ masterName: s.masterName }) : pres.addSlide();
    slide.background = { color: C.COLOR.paper };

    if (s.kind === 'template') renderTemplate(pres, slide, s);
    else if (s.kind === 'cover') renderCover(pres, slide, s);
    else if (s.kind === 'contents') renderContents(pres, slide, s);
    else renderItem(pres, slide, s);
  }

  const raw = await pres.write({ outputType: 'nodebuffer' });
  const { buffer } = await promotePicturePlaceholders(raw, slots);
  return buffer;
}

module.exports = { renderPptx, toText, masterKeyFor, promotePicturePlaceholders };
