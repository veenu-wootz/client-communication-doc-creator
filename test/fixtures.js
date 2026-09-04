/**
 * fixtures.js — the acceptance-test inputs.
 *
 * Numbered to match §14 of query-doc-format-spec.md, plus a few that exist only
 * because we flow content instead of truncating it (PLAN.md §5.1).
 *
 * Images carry width/height directly so the planner can be tested without any
 * network or file access — that is the point of keeping it pure.
 */

const LOGO = { path: 'logo.png', width: 900, height: 300, aspect: 3, ok: true };

const img = (name, w, h, caption) => ({
  path: `media/${name}.png`, width: w, height: h, aspect: w / h, caption: caption || '', ok: true,
});

const upright = (n, c) => img(n, 800, 1000, c);   // aspect 0.8
const wide    = (n, c) => img(n, 2000, 800, c);   // aspect 2.5
const square  = (n, c) => img(n, 1000, 1000, c);  // aspect 1.0

const words = (n) =>
  Array.from({ length: n }, (_, i) =>
    ['tolerance', 'datum', 'fixture', 'weld', 'radius', 'flange', 'bracket', 'gauge'][i % 8]).join(' ');

const doc = (over = {}) => ({
  report_title: 'Bracket Assembly — Phase 2',
  reference_name: 'BRK-4471-B',
  addressee: 'Mr. R. Iyer, Meridian Auto',
  additional_details: 'Following the design review on 28 Aug.',
  created_by: 'Priya Nair',
  created_at: '2026-09-03',
  product_photo: null,
  logo: LOGO,
  ...over,
});

const query = (id, title, body, images = []) => ({ id, title, body, needs_response: true, images });
const update = (id, title, body, images = []) => ({ id, title, body, needs_response: false, images });
const named = (item, item_name) => ({ ...item, item_name });

const FILLET = 'Drawing shows R2.5 at the flange base but the GD&T callout references R4. Confirm which governs.';

const fixtures = {
  '01-flagged-unflagged-flagged': {
    document: doc(),
    items: [
      query('a', 'Fillet radius conflict', FILLET),
      update('b', 'Fixture design complete', 'Fixture design is complete and released to the shop floor.'),
      query('c', 'Surface finish on face B', 'Sheet 3 says Ra 0.8, the drawing says Ra 1.6. Which governs?'),
    ],
  },

  '02-alternating-ten': {
    document: doc(),
    items: Array.from({ length: 10 }, (_, i) =>
      i % 2 === 0
        ? query(`q${i}`, `Question ${i / 2 + 1}`, `Please confirm point ${i / 2 + 1} on the drawing.`)
        : update(`u${i}`, `Update ${Math.ceil(i / 2)}`, 'Progress note for the current phase.')),
  },

  '03-one-upright-image': {
    document: doc(),
    items: [query('a', 'Fillet radius', FILLET, [upright('flange', 'Sheet 2, detail A')])],
  },

  '04-one-wide-image': {
    document: doc(),
    items: [query('a', 'Tolerance stack', 'Datum B is unmarked on this view. Please confirm the stack-up basis.', [wide('assembly', 'Wide assembly drawing')])],
  },

  '05-three-images': {
    document: doc(),
    items: [query('a', 'Weld positions', 'Which two are correct?', [square('w1'), square('w2'), square('w3')])],
  },

  '06-six-images': {
    document: doc(),
    items: [query('a', 'Six views', 'All six views attached for reference.',
      [square('v1'), square('v2'), square('v3'), square('v4'), square('v5'), square('v6')])],
  },

  '07-long-body-4-images': {
    document: doc(),
    items: [query('a', 'Weld position on bracket', words(900),
      [square('p1', 'primary'), square('p2'), square('p3'), square('p4')])],
  },

  '09-unflagged-only': {
    document: doc(),
    items: [update('a', 'Phase 2 timeline', 'Trials shifted to 22 Sep following the drawing hold.')],
  },

  '10-two-queries': {
    document: doc(),
    items: [query('a', 'First', 'Please confirm A.'), query('b', 'Second', 'Please confirm B.')],
  },

  '11-four-queries': {
    document: doc(),
    items: Array.from({ length: 4 }, (_, i) => query(`q${i}`, `Question ${i + 1}`, `Please confirm item ${i + 1}.`)),
  },

  '12-four-queries-two-updates': {
    document: doc(),
    items: [
      ...Array.from({ length: 4 }, (_, i) => query(`q${i}`, `Question ${i + 1}`, `Please confirm item ${i + 1}.`)),
      update('u1', 'Fixture ready', 'Fixture design complete.'),
      update('u2', 'Timeline', 'Trials shifted to 22 Sep.'),
    ],
  },

  '13-zero-flagged': {
    document: doc(),
    items: [
      update('a', 'Fixture design complete', 'Released to the shop floor.'),
      update('b', 'Phase 2 timeline', 'Trials shifted to 22 Sep.'),
    ],
  },

  '14-empty-items': { document: doc(), items: [] },

  '15-cover-all-null': {
    document: {
      report_title: 'Bracket Assembly', reference_name: null,
      addressee: null, additional_details: null, created_by: 'Priya Nair',
      created_at: '2026-09-03', product_photo: null, logo: LOGO,
    },
    items: [query('a', 'Only question', 'Please confirm the revision.')],
  },

  '16-broken-image': {
    document: doc(),
    items: [query('a', 'Missing drawing', 'The referenced sheet did not load.',
      [{ path: 'media/gone.png', width: 0, height: 0, aspect: 1, ok: false, caption: 'gone.png' }])],
  },

  '17-five-updates': {
    document: doc(),
    items: Array.from({ length: 5 }, (_, i) =>
      update(`u${i}`, `Update ${i + 1}`, `Short progress note number ${i + 1}.`)),
  },

  '18-update-between-queries': {
    document: doc(),
    items: [
      query('a', 'First question', 'Please confirm A.'),
      update('b', 'Interleaved update', 'This update sits between two queries.'),
      query('c', 'Second question', 'Please confirm B.'),
    ],
  },

  '21-cover-with-photo': {
    document: doc({ product_photo: { path: 'media/product.png', width: 1200, height: 900, aspect: 4 / 3, ok: true } }),
    items: [query('a', 'Fillet radius', FILLET, [upright('flange')])],
  },

  '22-forty-queries': {
    document: doc(),
    items: Array.from({ length: 40 }, (_, i) =>
      query(`q${i}`, `Question ${i + 1} about a reasonably long descriptive title that keeps going`, `Please confirm item ${i + 1}.`)),
  },

  '23-empty-body-with-image': {
    document: doc(),
    items: [query('a', 'Image only', '', [upright('drawing', 'The only content')])],
  },

  '25-multiple-skus': {
    document: doc(),
    items: [
      named(query('a', 'Thread pitch out of spec', '4% measured at 1.75 mm against 1.75 plus or minus 0.02 mm. Confirm disposition.'), 'HEX NUT M12'),
      named(query('b', 'Sand casting alternative', 'Open to sand casting instead of investment casting? Around 30% cheaper, Ra 6.3 to Ra 12.5.'), 'BRACKET BRK-2210'),
      named(update('c', 'Pilot batch on the bench', '47 of 50 passed pressure and flow. 3 failed seal integrity.'), 'PUMP ASSY'),
      query('d', 'Tolerance stack-up', 'Drawing shows 0.05 mm, assembly implies 0.02 mm. Which governs?'),
    ],
  },

  '26-numbered-list-body': {
    document: doc(),
    items: [query('a', 'Queries on XSJ-956E/C',
      'There are four queries:\n\n1.  Sheet thickness\n2.  Laser marking details, lettering standard, depth, font\n3.  Radius of the corners\n4.  Centre to centre distances of the holes',
      [upright('drawing', 'XSJ-956E/C')])],
  },

  '24-no-body-no-images': {
    document: doc(),
    items: [query('a', 'Header alone', '')],
  },
};

module.exports = { fixtures, LOGO, upright, wide, square, words, doc, query, update, named };
