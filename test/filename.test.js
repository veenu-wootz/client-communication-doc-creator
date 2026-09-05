/**
 * filename.test.js — the name the deck is filed under.
 *
 * Format: "Queries - {project} - {DD Mon YYYY H.MM AM/PM}" in IST. The stamp is
 * generation time, not the submission's created_at.
 */

const test = require('node:test');
const assert = require('node:assert');

const { istStamp } = require('../server');

test('the stamp is IST, not UTC', () => {
  // 09:18 UTC is 14:48 IST — a +5:30 offset, so this catches a timezone slip.
  assert.strictEqual(istStamp(new Date('2026-09-05T09:18:33.374Z')), '05 Sep 2026 2.48 PM');
});

test('it crosses midnight into the next IST day correctly', () => {
  // 19:00 UTC on the 4th is 00:30 IST on the 5th.
  assert.strictEqual(istStamp(new Date('2026-09-04T19:00:00Z')), '05 Sep 2026 12.30 AM');
});

test('months are the same three letters the deck footer uses', () => {
  const { formatDate } = require('../src/plan/planner');
  for (const [iso, month] of [['2026-09-05', 'Sep'], ['2026-01-05', 'Jan'], ['2026-12-05', 'Dec']]) {
    assert.ok(istStamp(new Date(`${iso}T06:00:00Z`)).includes(month));
    assert.ok(formatDate(iso).includes(month), 'footer and filename must not disagree');
  }
});

test('it carries no character OneDrive rejects in a filename', () => {
  const stamp = istStamp(new Date('2026-09-05T09:18:33.374Z'));
  for (const bad of ['"', '*', ':', '<', '>', '?', '/', '\\', '|']) {
    assert.ok(!stamp.includes(bad), `stamp must not contain ${bad}`);
  }
});
