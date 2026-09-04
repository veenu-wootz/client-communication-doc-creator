/**
 * constants.js — the single source of truth for geometry, type, and thresholds.
 *
 * Every number the layout depends on lives here. Nothing downstream may use a
 * magic number. All distances are in INCHES; all type sizes in POINTS.
 *
 * See PLAN.md §5 for where these knowingly differ from query-doc-format-spec.md.
 */

// ── Canvas ───────────────────────────────────────────────────
// The pptxgenjs default canvas is 10 x 5.625in. The layout MUST be set to these
// dimensions before any slide is added — coordinates past the canvas edge are
// written rather than clamped, so shapes silently vanish.
const SLIDE_W = 13.333;
const SLIDE_H = 7.5;
const MARGIN  = 0.6;

const CONTENT_X = 0.6;
const CONTENT_W = 12.133;          // SLIDE_W - 2*MARGIN

// ── Header ───────────────────────────────────────────────────
const HEADER_Y        = 0.55;
const HEADER_H        = 0.70;      // one line
const HEADER_H_2LINE  = 1.15;      // wrapped to two lines
const HEADER_MAX_LINES = 2;        // we wrap rather than truncate (PLAN.md §5.1)
// Height taken by the item_name line above the header, when one is present.
const HEADER_EYEBROW_H = 0.24;

// ── Content area ─────────────────────────────────────────────
const CONTENT_Y        = 1.50;     // top of content when the header is one line
const CONTENT_Y_2LINE  = 1.95;     // top of content when the header wrapped
const CONTENT_BOTTOM   = 6.95;     // bottom of the content area on a plain slide

// ── Reply box ────────────────────────────────────────────────
const REPLY_Y   = 5.85;
const REPLY_H   = 1.10;
const REPLY_GAP = 0.20;            // clearance between content and the reply box
const REPLY_INSET = 0.12;          // label inset inside the box

// ── Footer ───────────────────────────────────────────────────
// The footer sits inside MARGIN. It and the logo are the two documented
// exceptions to "no element outside MARGIN" (PLAN.md §5.2).
const FOOTER_Y     = 7.05;
const FOOTER_H     = 0.30;
const FOOTER_LOGO_H = 0.20;
const FOOTER_LOGO_SLOT_W = 0.75;   // reserved width so footer text never collides
const COVER_LOGO_H  = 0.45;

const GUTTER = 0.30;               // between content blocks

// ── Layout thresholds ────────────────────────────────────────
const WIDE_ASPECT_THRESHOLD = 1.4; // aspect >= this is "wide", else "upright"
const MIN_IMAGE_W           = 3.0; // below this an image is not legible; reflow instead
const MAX_IMAGES_PER_ITEM   = 4;   // the Strike form's own cap
const IMAGE_MAX_LONG_EDGE   = 1600; // downscale before embedding, to control file size
// Never render an image so large that it drops below this effective resolution.
// A 200px-wide screenshot stretched across a 5in column is ~40 DPI and unreadable,
// which defeats the point of showing the drawing at all.
const MIN_RENDER_DPI        = 96;

// Floor for an image zone in the stacked layouts, as a fraction of the content
// area. The zone grows beyond this when the body is short — a wide drawing is
// the exact thing the client needs to read (spec §9.3).
const IMAGE_ZONE_FRACTION = 0.55;
// Ceiling on the same zone, so the body always keeps a share of the slide.
const IMAGE_ZONE_MAX = 0.80;
// Width of the image column in the side layouts. Two images need more room than
// one, or each cell falls below MIN_IMAGE_W and the pair layout never fires.
const SIDE_IMAGE_FRACTION = 0.45;
const SIDE_PAIR_IMAGE_FRACTION = 0.60;
// In the grid layouts the body sits above the images; beyond this many lines the
// body wins the slide and the images move on rather than the body being cut.
const GRID_BODY_MAX_LINES = 3;
// The grid is two rows. More images than this go to another slide rather than
// growing a third row off the bottom of the content area.
const GRID_MAX_CELLS = 4;

// The summary appears once the deck has this many items in total — queries and
// updates alike. Counting only queries meant a document of 2 queries and 6
// updates got no summary at all, even though the summary now names updates too.
const INCLUDE_SUMMARY_MIN_ITEMS = 3;
const CONTENTS_MAX_ENTRIES         = 12;  // then paginate — never trim (PLAN.md §5.1)
const MAX_GROUPED_UPDATES          = 3;

// ── Text fitting (PLAN.md §4) ────────────────────────────────
// We plan to fill this fraction of a box, then break to the next slide. Being
// conservative costs whitespace, never content, because nothing is ever
// truncated. This is the one number to tune if slides look too empty or too full.
const FILL_TARGET = 0.85;

// Emit our own line breaks instead of letting PowerPoint wrap. Stronger
// no-overflow guarantee, worse copy-paste. Off unless real decks show spill.
const HARD_WRAP = false;

const LINE_HEIGHT = 1.2;           // multiple of font size

// ── Type scale ───────────────────────────────────────────────
const FONT = 'Arial';              // PLAN.md T3 — do not change without reading it

const TYPE = {
  eyebrow:       { size: 11, bold: false },   // item_name above the slide header
  sectionLabel:  { size: 12, bold: true  },   // "QUERIES — 3 need your response"
  coverTitle:    { size: 32, bold: true  },
  coverValue:    { size: 14, bold: false },
  coverLabel:    { size: 10, bold: false },
  contentsHead:  { size: 24, bold: true  },
  contentsEntry: { size: 15, bold: false },
  slideHeader:   { size: 22, bold: true  },
  body:          { size: 16, bold: false },
  groupedTitle:  { size: 16, bold: true  },
  caption:       { size: 10, bold: false },
  replyLabel:    { size: 10, bold: false },
  footer:        { size:  9, bold: false },
};

const BODY_SIZE_MAX  = 16;
const BODY_SIZE_MIN  = 14;         // hard floor; below this we add a slide instead
const BODY_SIZE_STEP = 1;

// ── Palette ──────────────────────────────────────────────────
// Hex WITHOUT '#' and WITHOUT alpha — both corrupt the file.
const COLOR = {
  ink:      '111827',   // primary text
  muted:    '6B7280',   // labels, captions, footer
  border:   'D9DCE1',   // hairlines, image frames, reply box
  surface:  'F5F6F8',   // subtle block fills, missing-image placeholder
  accent:   '0F4C5C',   // query number chip text
  accentBg: 'E3EDEF',   // query number chip background
  paper:    'FFFFFF',   // slide background
};

const HAIRLINE = 0.75;             // pt, for image frames and the reply box

// Every item slide carries a chip so none is ever ambiguous: a query shows its
// number loudly, an update says so quietly. The query number is the thing that
// matters, so the update chip is deliberately smaller and greyer.
const CHIP = {
  query:  { text: COLOR.accent, bg: COLOR.accentBg, size: 22 },
  update: { text: COLOR.muted,  bg: COLOR.surface,  size: 12 },
};
const UPDATE_CHIP_LABEL = 'Update';

// ── Copy ─────────────────────────────────────────────────────
const REPLY_INSTRUCTION = 'Please reply quoting the query number.';

// The summary slide groups the deck: everything owing an answer, then the rest.
const SUMMARY_HEADING = 'Summary';
const SECTION_QUERIES = (n) =>
  (n === 1 ? 'QUERIES — 1 needs your response' : `QUERIES — ${n} need your response`);
const SECTION_UPDATES = 'OTHER UPDATES';
const REPLY_LABEL       = (n) => `Your response — Q${n}`;
const UNTITLED          = 'Untitled';
const TITLE_FALLBACK_CHARS = 60;   // first N chars of body when no title exists

module.exports = {
  SLIDE_W, SLIDE_H, MARGIN, CONTENT_X, CONTENT_W,
  HEADER_Y, HEADER_H, HEADER_H_2LINE, HEADER_MAX_LINES, HEADER_EYEBROW_H,
  CONTENT_Y, CONTENT_Y_2LINE, CONTENT_BOTTOM,
  REPLY_Y, REPLY_H, REPLY_GAP, REPLY_INSET,
  FOOTER_Y, FOOTER_H, FOOTER_LOGO_H, FOOTER_LOGO_SLOT_W, COVER_LOGO_H,
  GUTTER,
  WIDE_ASPECT_THRESHOLD, MIN_IMAGE_W, MAX_IMAGES_PER_ITEM, IMAGE_MAX_LONG_EDGE,
  IMAGE_ZONE_FRACTION, IMAGE_ZONE_MAX, SIDE_IMAGE_FRACTION, SIDE_PAIR_IMAGE_FRACTION,
  MIN_RENDER_DPI,
  GRID_BODY_MAX_LINES, GRID_MAX_CELLS,
  INCLUDE_SUMMARY_MIN_ITEMS, CONTENTS_MAX_ENTRIES, MAX_GROUPED_UPDATES,
  FILL_TARGET, HARD_WRAP, LINE_HEIGHT,
  FONT, TYPE, BODY_SIZE_MAX, BODY_SIZE_MIN, BODY_SIZE_STEP,
  COLOR, HAIRLINE, CHIP, UPDATE_CHIP_LABEL,
  REPLY_INSTRUCTION, SUMMARY_HEADING, SECTION_QUERIES, SECTION_UPDATES,
  REPLY_LABEL, UNTITLED, TITLE_FALLBACK_CHARS,
};
