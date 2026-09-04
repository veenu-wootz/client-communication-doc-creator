/**
 * openai.js — the same enrichment via the OpenAI API.
 *
 * Deliberately mirrors providers/anthropic.js: same prompt, same schema, same
 * return shape. Only the transport differs.
 */

const OpenAI = require('openai');
const { SYSTEM, ITEM_SCHEMA, buildUserMessage } = require('./schema');

const MODEL = process.env.OPENAI_MODEL || 'gpt-4.1-mini';

let client;
const getClient = () => (client = client || new OpenAI());

async function deriveItemMeta(items) {
  const res = await getClient().chat.completions.create({
    model: MODEL,
    messages: [
      { role: 'system', content: SYSTEM },
      { role: 'user', content: buildUserMessage(items) },
    ],
    response_format: {
      type: 'json_schema',
      json_schema: { name: 'item_labels', strict: true, schema: ITEM_SCHEMA },
    },
  });

  const text = res.choices?.[0]?.message?.content;
  if (!text) throw new Error('empty completion');

  const out = JSON.parse(text);
  if (!out || !Array.isArray(out.items)) throw new Error('completion carried no items array');
  return out.items;
}

const isConfigured = () => Boolean(process.env.OPENAI_API_KEY);

module.exports = { deriveItemMeta, isConfigured, name: 'openai', model: MODEL };
