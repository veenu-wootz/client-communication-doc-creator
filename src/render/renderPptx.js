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
 * The reserved image column on a slide that has no picture.
 *
 * Outline only, no fill: if the sender never drops an image in, this reads as a
 * deliberately reserved frame rather than an unfinished box. Drawn rather than a
 * native PowerPoint placeholder because its position is computed per slide, and
 * native placeholders live at fixed positions on a master (PLAN.md §5.14).
 */
function drawImagePlaceholder(slide, pres, ph) {
  slide.addShape(pres.ShapeType.rect, {
    x: ph.box.x, y: ph.box.y, w: ph.box.w, h: ph.box.h,
    fill: { type: 'none' },
    line: { color: C.COLOR.border, width: C.HAIRLINE, dashType: C.PLACEHOLDER_DASH },
  });
  slide.addText(C.PLACEHOLDER_LABEL, textOpts(
    { x: ph.box.x, y: ph.box.y + ph.box.h / 2 - 0.14, w: ph.box.w, h: 0.28 },
    { size: C.TYPE.caption.size, color: C.COLOR.muted, align: 'center' },
  ));
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
    slide.addText(
      s.list.entries.map((e) => ({ text: toText(e.lines), options: { breakLine: true } })),
      {
        ...textOpts(s.list.box, { size: s.list.size }),
        bullet: { type: 'number', numberStartAt: s.list.startAt, indent: s.list.indent * 72 },
      },
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

  if (s.placeholder) drawImagePlaceholder(slide, pres, s.placeholder);
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

/**
 * Define one master per template variant.
 *
 * Native placeholders are what make these safe to ship: PowerPoint shows a
 * "click to add" prompt while editing and renders nothing at all if the sender
 * never touches them, unlike a drawn box which would print empty.
 */
function defineTemplateMasters(pres, templates) {
  for (const t of templates) {
    const objects = [
      {
        placeholder: {
          options: {
            name: 'title', type: 'title',
            x: t.titleBox.x + 0.95, y: t.titleBox.y,
            w: t.titleBox.w - 0.95, h: t.titleBox.h,
            fontFace: C.FONT, fontSize: C.TYPE.slideHeader.size, bold: true, color: C.COLOR.ink,
            margin: 0, valign: 'middle',
          },
          text: C.TEMPLATE_TITLE_PROMPT,
        },
      },
      {
        placeholder: {
          options: {
            name: 'body', type: 'body',
            x: t.bodyBox.x, y: t.bodyBox.y, w: t.bodyBox.w, h: t.bodyBox.h,
            fontFace: C.FONT, fontSize: C.TYPE.body.size, color: C.COLOR.ink,
            margin: 0, valign: 'top',
          },
          text: C.TEMPLATE_BODY_PROMPT,
        },
      },
    ];

    t.images.forEach((box, i) => {
      objects.push({
        placeholder: {
          options: {
            name: `pic${i + 1}`, type: 'pic',
            x: box.x, y: box.y, w: box.w, h: box.h,
            fontFace: C.FONT, fontSize: C.TYPE.caption.size, color: C.COLOR.muted, align: 'center',
          },
          text: C.TEMPLATE_IMAGE_PROMPT,
        },
      });
    });

    pres.defineSlideMaster({ title: `${C.TEMPLATE_MASTER}_${t.variant}`, objects });
  }
}

/** A template slide: chrome drawn as usual, content left to the placeholders. */
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

  defineTemplateMasters(pres, plan.slides.filter((s) => s.kind === 'template'));

  pres.author = plan.document.created_by || 'Wootz';
  pres.company = 'Wootz';
  pres.title = plan.document.report_title || 'Client communication';

  for (const s of plan.slides) {
    const slide = s.kind === 'template'
      ? pres.addSlide({ masterName: `${C.TEMPLATE_MASTER}_${s.variant}` })
      : pres.addSlide();
    slide.background = { color: C.COLOR.paper };

    if (s.kind === 'template') renderTemplate(pres, slide, s);
    else if (s.kind === 'cover') renderCover(pres, slide, s);
    else if (s.kind === 'contents') renderContents(pres, slide, s);
    else renderItem(pres, slide, s);
  }

  return pres.write({ outputType: 'nodebuffer' });
}

module.exports = { renderPptx, toText };
