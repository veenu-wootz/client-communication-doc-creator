/**
 * anthropic.js — enrichment via the Claude API.
 *
 * One batched call for the whole document, not one per row. A strict tool
 * schema is what guarantees the shape comes back parseable; the caller still
 * has a fallback chain if it does not.
 */

const Anthropic = require('@anthropic-ai/sdk');
const { SYSTEM, ITEM_SCHEMA, buildUserMessage } = require('./schema');

const MODEL = process.env.ANTHROPIC_MODEL || 'claude-opus-5';
const TOOL_NAME = 'emit_labels';

let client;
const getClient = () => (client = client || new Anthropic());

async function deriveItemMeta(items) {
  const res = await getClient().messages.create({
    model: MODEL,
    max_tokens: 16000,
    // Labelling short text needs no deep reasoning; low effort keeps the
    // webhook snappy. Adaptive thinking stays on — disabling it on Opus 5 has
    // documented failure modes around tool calls.
    output_config: { effort: 'low' },
    system: SYSTEM,
    tools: [{
      name: TOOL_NAME,
      description: 'Return one label per input row, in order.',
      input_schema: ITEM_SCHEMA,
      strict: true,
    }],
    tool_choice: { type: 'tool', name: TOOL_NAME },
    messages: [{ role: 'user', content: buildUserMessage(items) }],
  });

  const call = res.content.find((b) => b.type === 'tool_use' && b.name === TOOL_NAME);
  if (!call) throw new Error(`no ${TOOL_NAME} call in response (stop_reason: ${res.stop_reason})`);

  const out = typeof call.input === 'string' ? JSON.parse(call.input) : call.input;
  if (!out || !Array.isArray(out.items)) throw new Error('tool call carried no items array');
  return out.items;
}

const isConfigured = () => Boolean(process.env.ANTHROPIC_API_KEY);

module.exports = { deriveItemMeta, isConfigured, name: 'anthropic', model: MODEL };
