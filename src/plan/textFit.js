/**
 * textFit.js — how much text fits where.
 *
 * Deliberately simple. See PLAN.md §4 for why this is a character-width table
 * rather than real font metrics: precision only mattered while a wrong estimate
 * meant losing text, and nothing is ever truncated here. A wrong estimate now
 * costs whitespace, which is the cheap failure.
 *
 * Widths are in em (fraction of font size) for Arial, measured from the
 * standard face. They do not need to be exact — FILL_TARGET absorbs the error.
 */

const { LINE_HEIGHT, FILL_TARGET } = require('./constants');

// ── Character widths in em ───────────────────────────────────
const W_NARROW  = 0.28;   // i l j t f I . , ; : ' ` | ! [ ] ( ) { } / \ -
const W_WIDE    = 0.83;   // m w M W
const W_UPPER   = 0.67;   // A-Z and other caps-width glyphs
const W_DIGIT   = 0.56;   // 0-9 and most punctuation-with-width
const W_DEFAULT = 0.52;   // lowercase
const W_SPACE   = 0.28;

const NARROW_SET = new Set("iljtfI.,;:'`|!()[]{}/\\-".split(''));
const WIDE_SET   = new Set('mwMW'.split(''));

/** Width of one character, in em. */
function charWidth(ch) {
  if (ch === ' ') return W_SPACE;
  if (NARROW_SET.has(ch)) return W_NARROW;
  if (WIDE_SET.has(ch)) return W_WIDE;
  if (ch >= '0' && ch <= '9') return W_DIGIT;
  if (ch >= 'A' && ch <= 'Z') return W_UPPER;
  return W_DEFAULT;
}

/**
 * Width of a string in inches at a given point size.
 * Bold text runs a little wider; 1.05 is close enough for Arial.
 */
function measure(text, sizePt, bold = false) {
  let em = 0;
  for (const ch of String(text ?? '')) em += charWidth(ch);
  return (em * sizePt * (bold ? 1.05 : 1)) / 72;
}

/** Height of n lines at a given point size, in inches. */
function heightOf(lineCount, sizePt) {
  return (lineCount * sizePt * LINE_HEIGHT) / 72;
}

/** How many whole lines fit in a box of the given height. */
function linesThatFit(heightIn, sizePt) {
  return Math.max(0, Math.floor(heightIn / ((sizePt * LINE_HEIGHT) / 72)));
}

/**
 * Break a long word that cannot fit on a line of its own.
 * Returns the number of characters that fit within widthIn.
 */
function charsThatFit(word, widthIn, sizePt, bold) {
  let em = 0;
  const limitEm = (widthIn * 72) / (sizePt * (bold ? 1.05 : 1));
  let i = 0;
  for (const ch of word) {
    const w = charWidth(ch);
    if (em + w > limitEm && i > 0) break;
    em += w;
    i += 1;
  }
  return i;
}

/**
 * Greedy word wrap. Returns an array of lines.
 * Preserves paragraph breaks in the source text as separate lines.
 */
function wrap(text, widthIn, sizePt, bold = false) {
  const source = String(text ?? '');
  if (!source.trim()) return [];

  const lines = [];

  for (const paragraph of source.split(/\r?\n/)) {
    if (!paragraph.trim()) {
      lines.push('');
      continue;
    }

    let line = '';
    const words = paragraph.trim().split(/\s+/);

    for (let word of words) {
      // A single word wider than the whole box: break it by character rather
      // than let it overflow. Rare, but part numbers and URLs do this.
      while (measure(word, sizePt, bold) > widthIn) {
        if (line) { lines.push(line); line = ''; }
        const n = charsThatFit(word, widthIn, sizePt, bold);
        if (n === 0) break;
        lines.push(word.slice(0, n));
        word = word.slice(n);
      }

      const candidate = line ? `${line} ${word}` : word;
      if (measure(candidate, sizePt, bold) <= widthIn) {
        line = candidate;
      } else {
        if (line) lines.push(line);
        line = word;
      }
    }

    if (line) lines.push(line);
  }

  return lines;
}

/**
 * Split wrapped lines into what fits a box and what is left over.
 *
 * `FILL_TARGET` keeps us clear of the bottom edge — the whole point of the
 * conservative approach. The remainder is never discarded; the planner puts it
 * on the next slide.
 *
 * Pass `fill: 1` when the box was sized from this very text (the grid body
 * does this). Discounting a box against the content that defined it always
 * reports a false overflow.
 *
 * Returns { fitted: string[], remainder: string }.
 */
function splitToFit(text, widthIn, heightIn, sizePt, bold = false, { fill = FILL_TARGET } = {}) {
  const lines = wrap(text, widthIn, sizePt, bold);
  const capacity = linesThatFit(heightIn * fill, sizePt);

  if (lines.length <= capacity) {
    return { fitted: lines, remainder: '' };
  }

  return {
    fitted: lines.slice(0, capacity),
    remainder: lines.slice(capacity).join(' ').trim(),
  };
}

/** Does the whole text fit the box at this size? */
function fits(text, widthIn, heightIn, sizePt, bold = false, { fill = FILL_TARGET } = {}) {
  return wrap(text, widthIn, sizePt, bold).length <= linesThatFit(heightIn * fill, sizePt);
}

module.exports = {
  measure, heightOf, linesThatFit, wrap, splitToFit, fits, charWidth,
};
