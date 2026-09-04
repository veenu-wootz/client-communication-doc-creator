/**
 * enrich.js — fill in each row's title and needs_response.
 *
 * needs_response is the single field the whole document format rests on: it is
 * what turns a row into a numbered Q with a reply box. So this runs a fallback
 * chain rather than ever failing:
 *
 *   LLM  →  the value the sender supplied in Strike  →  keyword heuristic
 *
 * Enrichment failing is never fatal. A deck built from heuristics is worse than
 * one built from the model; a deck that never got generated is worse than both.
 */

const { TITLE_FALLBACK_CHARS, UNTITLED } = require('../plan/constants');

/**
 * 'ai-first'     — the model's answer wins, Strike's value is the fallback
 * 'strike-first' — a value the sender explicitly set wins, the model fills gaps
 */
const ENRICH_PRECEDENCE = process.env.ENRICH_PRECEDENCE || 'ai-first';

const PROVIDERS = {
  anthropic: () => require('./providers/anthropic'),
  openai: () => require('./providers/openai'),
};

/** Explicit choice, else whichever key is actually present. */
function selectProvider() {
  const named = (process.env.ENRICH_PROVIDER || '').toLowerCase().trim();
  if (named === 'off') return null;
  if (named && PROVIDERS[named]) return PROVIDERS[named]();
  if (named) {
    console.warn(`  ENRICH_PROVIDER="${named}" is not a known provider — falling back to key detection`);
  }
  for (const load of Object.values(PROVIDERS)) {
    const p = load();
    if (p.isConfigured()) return p;
  }
  return null;
}

const ASKING = /\b(confirm|clarify|which|please advise|advise|kindly|revert|approve|approval|shall we|can you|could you|let us know|awaiting|pending your|your input|acceptable)\b/i;

/** Last resort when there is no model and the sender set nothing. */
function guessNeedsResponse(text) {
  const s = String(text || '');
  if (!s.trim()) return false;
  if (s.includes('?')) return true;
  return ASKING.test(s);
}

/**
 * Last resort title: the opening LINE of the body, cut on a word boundary.
 *
 * Bodies commonly open with a lead-in and then a numbered list, so slicing raw
 * characters produced titles with line breaks and half a list item inside them.
 */
function guessTitle(item) {
  const body = String(item.body || '').trim();
  if (!body) return item.title || UNTITLED;

  const firstLine = body.split(/\n/).map((l) => l.trim()).find(Boolean) || '';
  if (!firstLine) return item.title || UNTITLED;

  const tidy = (t) => t.replace(/[\s:;,.\-–—]+$/, '');
  if (firstLine.length <= TITLE_FALLBACK_CHARS) return tidy(firstLine) || UNTITLED;

  const cut = firstLine.slice(0, TITLE_FALLBACK_CHARS);
  const space = cut.lastIndexOf(' ');
  return tidy(space > 20 ? cut.slice(0, space) : cut) || UNTITLED;
}

const clean = (v) => (typeof v === 'string' ? v.trim() : '');

/** Collapse every run of whitespace, so only wording is compared. */
const wordsOnly = (v) => String(v ?? '').replace(/\s+/g, ' ').trim();

/**
 * Accept a reformatted body only if it is the original text with whitespace
 * moved around — nothing added, dropped, reordered, or reworded.
 *
 * This is what makes AI reformatting safe to run on the sender's own words: the
 * worst case is that the repair is rejected and the raw text renders as before.
 */
function isFormattingOnly(original, candidate) {
  if (typeof candidate !== 'string' || !candidate.trim()) return false;
  return wordsOnly(original) === wordsOnly(candidate);
}

/**
 * @param {object[]} items  parsed rows; title may be '', needs_response may be undefined
 * @returns {Promise<{items: object[], warnings: string[], provider: string|null}>}
 */
async function enrich(items) {
  const warnings = [];
  if (!items || items.length === 0) return { items: [], warnings, provider: null };

  const provider = selectProvider();
  let derived = null;

  if (provider) {
    try {
      const started = Date.now();
      const out = await provider.deriveItemMeta(items);
      if (out.length !== items.length) {
        warnings.push(`enrichment returned ${out.length} labels for ${items.length} rows — using it only where it lines up`);
      }
      derived = out;
      console.log(`  enriched ${items.length} rows via ${provider.name}/${provider.model} in ${Date.now() - started}ms`);
    } catch (e) {
      warnings.push(`enrichment via ${provider.name} failed (${e.message}) — using the form's own values`);
      console.warn(`  enrichment failed: ${e.message}`);
    }
  } else {
    warnings.push('no enrichment provider configured — using the form\'s own values');
  }

  const aiFirst = ENRICH_PRECEDENCE === 'ai-first';

  const resolved = items.map((item, i) => {
    const ai = derived && derived[i] ? derived[i] : null;
    const supplied = clean(item.title);
    const suppliedFlag = typeof item.needs_response === 'boolean' ? item.needs_response : undefined;

    const aiTitle = ai ? clean(ai.title) : '';
    const aiFlag = ai && typeof ai.needs_response === 'boolean' ? ai.needs_response : undefined;

    const title = (aiFirst ? (aiTitle || supplied) : (supplied || aiTitle)) || guessTitle(item);

    // A repair is used only when it survives the whitespace-only check.
    let body = item.body;
    if (ai && typeof ai.body === 'string' && ai.body.trim() && ai.body !== item.body) {
      if (isFormattingOnly(item.body, ai.body)) {
        body = ai.body;
      } else {
        warnings.push(`item ${item.id}: reformatted body altered the wording — kept the original text`);
      }
    }

    let needs;
    if (aiFirst) needs = aiFlag ?? suppliedFlag;
    else needs = suppliedFlag ?? aiFlag;
    if (needs === undefined) needs = guessNeedsResponse(item.body);

    return { ...item, title, body, needs_response: needs };
  });

  return { items: resolved, warnings, provider: provider ? provider.name : null };
}

module.exports = {
  enrich, guessNeedsResponse, guessTitle, selectProvider, isFormattingOnly, ENRICH_PRECEDENCE,
};
