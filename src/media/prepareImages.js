/**
 * prepareImages.js — resolve every image in a document before planning.
 *
 * The planner is pure and does no I/O, so all fetching and measuring happens
 * here. Everything runs in parallel; one bad image never blocks or fails the
 * rest, it just arrives with ok:false.
 */

const { fetchImage } = require('./fetchImage');

const LOGO_URL = () => process.env.LOGO_URL
  || 'https://res.cloudinary.com/dbwg6zz3l/image/upload/w_300,f_png,q_90/v1773643264/Black_Yellow_kq9kef.png';

/**
 * @param {object} doc  canonical document JSON (images still unresolved)
 * @returns {Promise<{document: object, warnings: string[]}>}
 */
async function prepareImages(doc) {
  const warnings = [];
  const d = doc.document || {};

  const jobs = [];
  const push = (source, apply) => jobs.push(fetchImage(source).then(apply));

  const document = { ...d };

  push(d.logo || LOGO_URL(), (r) => {
    document.logo = r.ok ? r : null;
    if (!r.ok) warnings.push(`logo unavailable (${r.reason}) — omitted`);
  });

  if (d.product_photo) {
    push(d.product_photo, (r) => {
      document.product_photo = r.ok ? r : null;
      if (!r.ok) warnings.push(`product photo unavailable (${r.reason}) — cover uses its no-photo variant`);
    });
  } else {
    document.product_photo = null;
  }

  const items = (doc.items || []).map((item) => {
    const images = (item.images || []).map((im) => {
      const caption = typeof im === 'string' ? '' : (im.caption || '');
      const slot = { caption };
      push(im, (r) => {
        Object.assign(slot, r, { caption });
        if (!r.ok) warnings.push(`item ${item.id}: image "${r.label}" unreadable (${r.reason}) — placeholder rendered`);
      });
      return slot;
    });
    return { ...item, images };
  });

  await Promise.all(jobs);
  return { document: { document, items }, warnings };
}

module.exports = { prepareImages };
