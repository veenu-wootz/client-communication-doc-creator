/**
 * schema.js — the one prompt and the one output shape, shared by both providers.
 *
 * Only the transport differs between providers; keeping the instruction here
 * means the two can never drift apart in what they actually ask for.
 */

const SYSTEM = `You label rows from a supplier-to-client communication form.

For each row you receive a free-text description written by an engineer or
sales person at a manufacturing supplier. Return two things per row:

1. "title" — a short noun phrase naming the subject, 2 to 6 words, no trailing
   punctuation, no "Query about" or "Update on" prefix. It becomes a slide
   heading. Use the row's own terminology: part features, drawing references,
   process names. Examples: "Fillet radius conflict", "Surface finish on face B",
   "Phase 2 trial dates".

2. "body" — a formatting-only repair, or null.

   Engineers often type several numbered points into one box without pressing
   Enter, so the description arrives as "1 Sheet Thickness 2 Laser Marking
   Details 3 Radius of the corners". Rendered as a paragraph it is unreadable,
   and it hides the correspondence between those numbers and the numbered
   callouts on the attached drawing.

   When the description reads as a run-on list like that, return the SAME text
   with a line break before each point. Otherwise return null.

   This is checked mechanically against the original, ignoring whitespace. Add,
   drop, reorder, or reword anything at all and the repair is discarded and the
   sender's raw text is used instead. Insert line breaks; change nothing else.
   Do not add numbering, bullets, or punctuation that was not already there.

3. "needs_response" — true when the sender is asking the client something and is
   waiting on an answer; false when they are only informing the client.

   Asking: a direct question, a request to confirm or clarify, two conflicting
   specifications presented for a decision, a choice offered to the client.
   Informing: progress, completion, schedule changes, dispatch notes.

   When a row genuinely does both, it is asking. When it is genuinely ambiguous,
   mark it as asking. A row wrongly marked as informing becomes a question the
   client never sees and never answers, which is the failure this whole document
   exists to prevent. A row wrongly marked as asking is only a visible surplus.

Return one entry per input row, in the same order. Never merge or drop rows.`;

const ITEM_SCHEMA = {
  type: 'object',
  properties: {
    items: {
      type: 'array',
      items: {
        type: 'object',
        properties: {
          title: { type: 'string', description: 'Short noun phrase, 2-6 words' },
          body: {
            type: ['string', 'null'],
            description: 'The same text with line breaks inserted before each point, or null to leave it alone',
          },
          needs_response: { type: 'boolean', description: 'True if the sender is waiting on an answer' },
        },
        // Strict schemas want every property listed; `body` carries null when
        // there is nothing to repair, which also keeps the response small.
        required: ['title', 'body', 'needs_response'],
        additionalProperties: false,
      },
    },
  },
  required: ['items'],
  additionalProperties: false,
};

/** The rows, numbered, as the user turn. */
function buildUserMessage(items) {
  return items
    .map((it, i) => {
      const parts = [`Row ${i + 1}:`];
      if (it.title) parts.push(`  existing title: ${it.title}`);
      parts.push(`  description: ${it.body || '(no description — images only)'}`);
      if (it.images && it.images.length) parts.push(`  attachments: ${it.images.length}`);
      return parts.join('\n');
    })
    .join('\n\n');
}

module.exports = { SYSTEM, ITEM_SCHEMA, buildUserMessage };
