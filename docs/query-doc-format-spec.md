# Query Document — Format Specification

**Purpose of this document:** define the output format of a PowerPoint generator precisely
enough that an implementation written from it is deterministic and correct without
guesswork. Every rule here is normative. Where a threshold is given, it must be a named
constant in code, not a magic number.

---

## 1. Context

### 1.1 What this produces

A generator that takes structured content and emits a `.pptx` (and PDF export) used by a
manufacturing supplier to communicate with a client during quotation and project
execution.

### 1.2 Who uses it

The sender is a project POC or RFQ POC at the supplier. They are technical, time-poor,
and today produce these communications as loose email text, marked-up screenshots,
ad-hoc slides, or Excel sheets — with no consistent format.

The recipient is an engineer, designer, or purchase contact at the client company. **The
recipient has no relationship with this product.** They will not install anything, sign
in, or learn a convention. They receive the deck as an email attachment and reply by
email.

### 1.3 What the document is for

Two things travel in one document:

- **Queries** — the sender needs an answer. Examples: a drawing shows R2.5 but the
  callout says R4; a surface finish is specified twice with different values.
- **Updates** — the sender is informing, not asking. Examples: fixture design complete,
  trial date moved.

These are entangled in practice. The same image often carries both — "here's the fixture
layout we completed, and the datum on face C is unclear." They are therefore **not two
sections**. They are one ordered stream of items, some of which are flagged as needing a
response.

### 1.4 The one thing this format exists to do

Make it unambiguous to the recipient **how many answers are owed and which is which.**

Query numbering is the entire mechanism. A recipient who replies "Q1: R4 governs, Q3: use
Rev B" has given the sender attributable answers instead of prose to be untangled by hand.
Every rule about ordering, numbering, and slide breaks exists to protect this. Where a
layout preference conflicts with numbering clarity, numbering wins.

### 1.5 Design register — read this before styling anything

This is **not a pitch deck.** General presentation-design advice ("vary the layouts",
"every slide needs a visual element", "don't create text-only slides", "pick a bold
palette") is actively wrong here and must be ignored.

This document is a technical instrument closer to a construction RFI sheet or an
inspection report. Its virtues are consistency, legibility, and predictability. Every
query slide should look like every other query slide. Repetition is the goal, because it
lets the recipient learn the format once and then scan. A drawing must be large enough
that a dimension is readable; nothing decorative may compete with it.

---

## 2. Non-goals — explicitly out of scope

Do not build these. Do not add them speculatively.

| Out of scope | Note |
|---|---|
| Fillable form fields in the PPTX | The reply box is a drawn rectangle, never an input |
| Parsing returned decks | Answers come back by email, handled elsewhere |
| Round / revision markers | Date only, for now |
| Prompting the user to split a long query | Overflow must resolve automatically |
| AI calls inside the renderer | All AI-derived content arrives in the input |
| Tables | Reserve the schema field; do not render |
| Editing generated output | Regeneration only |
| Client-side response capture | Not this component |

---

## 3. Core invariants

These must hold for every document the generator produces. Each maps to an acceptance
test in §14.

1. **Item order is preserved absolutely.** Items render in input array order. The
   generator never reorders, regroups across non-adjacent items, or sorts.
2. **Query numbers are sequential, global, and gapless.** `Q1, Q2, Q3…` assigned across
   flagged items in array order. Never restart, never skip, never scope to a section.
3. **Unflagged items are never numbered.** Numbering skips over them without consuming a
   number.
4. **One query is never confusable with another.** A query may span multiple slides, but
   every slide carrying it repeats its number, and no slide carries content from two
   different queries.
5. **Exactly one reply box per query**, on the last slide of that query.
6. **No text is ever clipped, overflowed, or spilled outside its container.**
7. **The generator never crashes on bad input.** Missing images, absent fields, and empty
   strings degrade gracefully.
8. **Same input produces the same output.** No randomness, no time-dependent values other
   than fields supplied in the input.

---

## 4. Input data model

The renderer is a **pure function of this JSON.** It performs no network calls, no AI
calls, and no content derivation. Titles are produced upstream and arrive here.

```jsonc
{
  "document": {
    "project_name": "Bracket Assembly — Phase 2",  // required, string
    "part_number": "BRK-4471-B",                   // optional, string|null
    "po_number": "PO-2026-0912",                   // optional, string|null
    "addressee": "Mr. R. Iyer, Meridian Auto",     // optional, string|null
    "additional_details": "Following the design review on 28 Aug.", // optional, string|null
    "created_by": "Priya Nair",                    // required, string
    "created_at": "2026-09-03",                    // required, ISO 8601 date
    "product_photo": "media/bracket.png",          // optional, path|null
    "logo": "assets/wootz-logo.png"                // required, path
  },
  "items": [
    {
      "id": "itm_001",                             // required, unique, stable
      "title": "Fillet radius conflict",           // required, string, AI-derived upstream
      "body": "Drawing shows R2.5 at the flange base but the GD&T callout references R4. Confirm which governs.",
      "needs_response": true,                      // required, boolean
      "images": [                                  // optional, 0-4 items
        { "path": "media/flange.png", "caption": "Sheet 2, detail A" }
      ],
      "table": null                                // reserved, always null in v1
    }
  ]
}
```

### 4.1 Field rules

| Field | Rule |
|---|---|
| `items` | May be empty. May contain zero flagged items. |
| `items[].id` | Used for logging and error messages only; never rendered. |
| `items[].title` | Required. If empty string, fall back to first 60 chars of `body`. If both empty, use `"Untitled"`. |
| `items[].body` | May be empty when images carry the content. |
| `items[].needs_response` | Inferred upstream, user-correctable upstream. The renderer treats it as authoritative. |
| `items[].images` | Hard cap of 4. If more than 4 are supplied, render the first 4 and emit a warning to the log — do not fail. |
| `images[].caption` | Optional. Rendered under the image at caption size. |
| `additional_details` | Freeform. **Truncate at 240 characters** with an ellipsis when rendering. |
| `product_photo` | When null, the cover uses its no-photo variant (§7.1). |

Internal field names above are for the schema and API only. **No field label text is ever
printed on a slide** except where §7 explicitly specifies it.

---

## 5. Geometry and type system

### 5.1 Canvas

Slide size is **13.333in × 7.5in** (16:9 wide). Set the presentation layout to this
**before adding any slide** — the library default is 10in × 5.625in, and coordinates past
the canvas edge are written rather than clamped, so shapes silently vanish.

Named constants, all in inches:

```
SLIDE_W            = 13.333
SLIDE_H            = 7.5
MARGIN             = 0.6

CONTENT_X          = 0.6
CONTENT_W          = 12.133          // SLIDE_W - 2*MARGIN

HEADER_Y           = 0.55
HEADER_H           = 0.70

CONTENT_Y          = 1.50
CONTENT_H_PLAIN    = 5.35            // no reply box
CONTENT_H_REPLY    = 4.05            // reply box present

REPLY_Y            = 5.75
REPLY_H            = 1.10

FOOTER_Y           = 6.95
FOOTER_H           = 0.35

GUTTER             = 0.30            // between content blocks
```

Do not place any element outside `MARGIN`. Do not place any element between
`CONTENT_Y + CONTENT_H_*` and `FOOTER_Y` other than the reply box.

### 5.2 Type scale

Body text uses a **safe font that renders true-to-width**: Calibri or Arial. Do not use
Aptos. Do not use fonts outside the safe list — overflow measurement becomes unreliable
and this format depends on accurate fit calculation.

| Role | Font | Size | Weight | Color |
|---|---|---|---|---|
| Cover — project name | Calibri | 32pt | bold | `ink` |
| Cover — field value | Calibri | 14pt | regular | `ink` |
| Cover — field label | Calibri | 10pt | regular | `muted` |
| Contents — heading | Calibri | 24pt | bold | `ink` |
| Contents — list entry | Calibri | 15pt | regular | `ink` |
| Slide header | Calibri | 22pt | bold | `ink` |
| Body | Calibri | 16pt → floor **14pt** | regular | `ink` |
| Image caption | Calibri | 10pt | regular | `muted` |
| Reply box label | Calibri | 10pt | regular | `muted` |
| Footer | Calibri | 9pt | regular | `muted` |

Body text shrinks in **1pt steps from 16pt to a hard floor of 14pt**. Below 14pt the
cascade moves to the next stage (§11); it never shrinks further.

All body text is **left-aligned**. Only the cover project name and contents heading may be
centered, and only in the no-photo cover variant.

### 5.3 Palette

Restrained and neutral. This is a client-facing technical document; color carries no
decorative role.

```
ink        = 111827   // primary text
muted      = 6B7280   // labels, captions, footer
border     = D9DCE1   // hairlines, image frames, reply box
surface    = F5F6F8   // subtle block fills
accent     = 0F4C5C   // query number chip only
accent_bg  = E3EDEF   // query number chip background
paper      = FFFFFF   // slide background
```

Write hex **without `#` and without alpha** — both corrupt the file. For translucency use
the library's dedicated transparency option.

`accent` and `accent_bg` must be single constants, replaceable in one place when brand
colors are confirmed.

### 5.4 Prohibited styling

- No accent line or rule beneath any title
- No color bars, edge stripes, or sidebar bands
- No gradient fills
- No cream or beige backgrounds — the slide background is white
- No decorative iconography
- No drop shadows except optionally a hairline frame on images

---

## 6. Slide sequence

```
1.  Cover                              (always)
2.  Contents                           (conditional — see below)
3.  Item slides, in input array order  (one or more slides per item)
```

```
INCLUDE_CONTENTS_MIN_QUERIES = 3
```

The contents slide is included **iff `query_count >= INCLUDE_CONTENTS_MIN_QUERIES`.**

The reply instruction line (§7.2) appears on the contents slide when present; otherwise it
appears on the cover, provided `query_count >= 1`. When `query_count == 0` it does not
appear at all.

---

## 7. Fixed slides

### 7.1 Cover

Two variants, selected by whether `product_photo` resolves to a readable image.

**Variant A — with product photo.** Left column (55% of content width) holds the text
block; right column (40%, with a `GUTTER` between) holds the photo, scaled to fit its box
preserving aspect ratio, vertically centered.

**Variant B — no product photo.** Single column, text block left-aligned at `CONTENT_X`,
occupying at most 70% of content width. Generous whitespace to the right. Do not stretch
the text to fill the space, and do not substitute a placeholder graphic.

Content of the text block, in order, omitting any block whose value is null or empty:

1. `project_name` — project name size
2. A field group, each as a `muted` label above an `ink` value:
   - `Part number` → `part_number`
   - `PO number` → `po_number`
   - `Attention` → `addressee`
3. `additional_details` as a plain paragraph at field-value size, no label, truncated per §4.1
4. Reply instruction line, if it belongs here per §6

The logo sits in the **top-right corner**, within `MARGIN`, height 0.45in, aspect
preserved.

The cover footer carries `created_by` and `created_at`, formatted as
`Prepared by {created_by} · {DD Mon YYYY}` at footer size in `muted`.

**The cover must compose correctly when `product_photo`, `part_number`, `po_number`,
`addressee`, and `additional_details` are all absent simultaneously.** This will occur on
real first documents. The result should be a clean project name, prepared-by line, and
logo — not a broken layout with gaps where blocks were expected.

### 7.2 Contents

Heading: `"{n} questions for you"`, or `"1 question for you"` when `n == 1`.

Below it, one line per flagged item in order:

```
Q1 · Fillet radius conflict
Q2 · Surface finish on face B
Q3 · Tooling cost basis
```

Titles truncate to a single line with an ellipsis; never wrap a contents entry.

If the document also contains unflagged items, add one line in `muted` beneath the list:
`"plus {m} project updates"` (`"plus 1 project update"` when `m == 1`).

**Updates are never itemised on the contents slide.** The slide's single job is to state
how many answers are owed; listing updates dilutes that signal and duplicates content
appearing in full later.

Last line, at body size in `muted`:

> Please reply quoting the query number.

---

## 8. Item slides — header and footer

### 8.1 Header

| Case | Header text |
|---|---|
| Flagged item, first slide | `Q{n} · {title}` |
| Flagged item, continuation slide | `Q{n} · {title} (continued)` |
| Unflagged item | `{title}` |

The `Q{n}` portion renders as a chip: `accent` text on an `accent_bg` rounded rectangle,
with the title following in `ink` at header size. The chip is part of the header at the
top-left — **not** a corner label at top-right, which readers miss and which drifts on
export.

The header occupies one line. If `Q{n} · {title}` exceeds the header width at header size,
truncate the **title** with an ellipsis. The number is never truncated or shrunk.

### 8.2 Footer

Every item slide carries, at footer size in `muted`:

```
{project_name} · {part_number} · {DD Mon YYYY}          Page {i} of {total}
```

Project name and part number truncate from the left group if the line would collide with
the page indicator. Omit null segments and their separators. The cover and contents slides
use the cover footer of §7.1 instead.

---

## 9. Layout selection

Deterministic. Given an item, compute `image_count` and, for each image, `aspect = width /
height`.

```
WIDE_ASPECT_THRESHOLD = 1.4
```

An image is **wide** when `aspect >= WIDE_ASPECT_THRESHOLD`, otherwise **upright**.

| Condition | Layout |
|---|---|
| `image_count == 0` | `TEXT_ONLY` |
| `image_count == 1` and upright | `IMAGE_SIDE` |
| `image_count == 1` and wide | `IMAGE_STACKED` |
| `image_count == 2` and both upright | `IMAGE_SIDE_PAIR` |
| `image_count == 2` and either wide | `IMAGE_STACKED_PAIR` |
| `image_count` in `{3, 4}` | `IMAGE_GRID` |

### 9.1 `TEXT_ONLY`

Body occupies the full content width at the top of the content area. Set body size to
16pt. Do **not** vertically center and do **not** enlarge type to fill the slide — a
sparse slide is correct here and reads as deliberate.

### 9.2 `IMAGE_SIDE`

Two columns with a `GUTTER` between: image column 45% of `CONTENT_W`, body column 55%.
Image scaled to fit its box preserving aspect, top-aligned. Body top-aligned.

### 9.3 `IMAGE_STACKED`

Image full content width at the top of the content area, height capped at 55% of available
content height, aspect preserved, horizontally centered. Body below it, full width, with a
`GUTTER` between.

A wide drawing placed in a half-width column becomes unreadable at the exact dimension
being asked about. This case is why the aspect rule exists.

### 9.4 `IMAGE_SIDE_PAIR`

As `IMAGE_SIDE`, but the image column is split into two equal cells side by side with a
`GUTTER` between. Captions beneath each.

### 9.5 `IMAGE_STACKED_PAIR`

As `IMAGE_STACKED`, but the two images stack vertically in the image zone, each capped at
half the zone height.

### 9.6 `IMAGE_GRID`

Body moves **above** the images, full content width, capped at 3 lines (truncate with an
ellipsis; the remainder flows to a continuation slide per §11). The image grid takes the
full content width below.

- **3 images:** two cells on the top row, one on the bottom row. The bottom cell has the
  same width as one top cell and is **horizontally centered.**
- **4 images:** 2 × 2.

```
MIN_IMAGE_W = 3.0
```

If a grid cell would render an image narrower than `MIN_IMAGE_W`, the image is not usable
at that size — trigger the overflow cascade (§11) rather than shipping an illegible
drawing.

### 9.7 Image rendering rules

- Preserve aspect ratio always. Never stretch or distort.
- Fit within the cell (contain), do not crop.
- Hairline `border` frame around each image, no shadow.
- Downscale source images to a maximum long edge of 1600px before embedding, to control
  file size.
- Caption, when present, sits directly beneath the image at caption size in `muted`, left
  aligned to the image, truncated to one line.

---

## 10. Reply box

Rendered **iff `needs_response == true`**, on the **last slide of that item only.**

- Position: `CONTENT_X, REPLY_Y`, width `CONTENT_W`, height `REPLY_H`
- A rounded rectangle, `border` stroke at hairline weight, `paper` fill
- Label in the top-left of the box, inset 0.12in, at reply-label size in `muted`:
  **`Your response — Q{n}`**
- The box is otherwise empty

The number in the label is mandatory. When a query spans two slides, the repeated number
is what prevents the recipient answering against the wrong query.

This is a **visual placeholder**, not an input control. Do not create a PowerPoint form
field, text placeholder, or editable content control. Its purpose is to signal that
something is owed — the same function an empty box serves on a paper RFI sheet. It works
identically for a recipient reading a PDF on a phone who will reply by email.

Unflagged items never receive a reply box. When a reply box is present, the content area
uses `CONTENT_H_REPLY`; otherwise `CONTENT_H_PLAIN`.

---

## 11. Overflow cascade

Content must never be clipped. When an item's content does not fit its content area, apply
these stages **in order**, stopping at the first that succeeds.

**Stage 1 — shrink text.** Reduce body size in 1pt steps from 16pt to the 14pt floor.

**Stage 2 — shrink images.** Scale images down toward `MIN_IMAGE_W`, reclaiming the space
for text. Never below `MIN_IMAGE_W`.

**Stage 3 — continuation slide.** Move images to a continuation slide.

Stage 3 rules:

- **Body text never moves.** It always stays on the item's first slide with the header,
  because the recipient reads the question first.
- **At least one image stays on the first slide**, so the question has visual context.
- Remaining images move to the continuation slide in their original order, laid out per §9
  for their own count.
- The continuation header is `Q{n} · {title} (continued)` for a flagged item, or
  `{title} (continued)` for an unflagged one.
- The reply box moves to the continuation slide — it always sits on the item's last slide.
- An item may produce at most **3 slides**. Beyond that, drop the surplus images and log a
  warning; do not produce a fourth slide and do not fail.

**Stage 4 — body still does not fit at 14pt with minimum images.** Reduce the body to what
fits at 14pt, and continue the remaining body text onto the continuation slide **above**
the images. This is the only case in which body text appears on a continuation slide.

Since prompting the user to split an over-long query is out of scope, this cascade must
terminate for any input. Verify termination explicitly.

---

## 12. Grouping unflagged items

```
MAX_GROUPED_UPDATES = 3
```

Consecutive unflagged items may share one slide when **all** of the following hold:

1. Every item in the group has `image_count == 0`
2. They are **adjacent in the input array** — never group across an intervening item
3. The group size is at most `MAX_GROUPED_UPDATES`
4. Their combined rendered height fits `CONTENT_H_PLAIN`

Each grouped item renders as its own titled block, title at section-header size, body
beneath, separated by `GUTTER`. The slide header is omitted for grouped slides; the blocks
carry their own titles.

**Flagged items are never grouped, with anything, ever.** An item with images is never
grouped. Grouping never reorders.

---

## 13. Edge cases

| Case | Behavior |
|---|---|
| `items` is empty | Emit cover only. Do not fail. |
| Zero flagged items | No contents slide, no numbers, no reply boxes, no reply instruction. |
| All items flagged | Normal. No update line on contents. |
| Image path missing or unreadable | Render a `surface`-filled rectangle with a hairline `border` and the filename in `muted` at caption size. Log a warning. Never crash, never skip silently. |
| Unsupported image format | Convert to PNG where possible; otherwise treat as unreadable. |
| Image over 4 per item | Render first 4, log a warning. |
| `title` empty | First 60 chars of `body`; if that is also empty, `"Untitled"`. |
| `body` empty, images present | Layout renders images only; do not leave a blank text column — use the stacked variant instead of the side variant. |
| `body` empty and no images | Render the header alone. Valid. |
| Title too long for header | Truncate title with ellipsis; never truncate the number. |
| `additional_details` over 240 chars | Truncate with ellipsis. |
| Logo missing | Omit it; do not substitute a placeholder. |
| Duplicate image paths in one item | Render both. Not an error. |
| Non-ASCII text | Must render correctly. Do not transliterate. |
| Very large source image | Downscale to 1600px long edge. |

---

## 14. Acceptance tests

Each maps to an invariant in §3. Implement these as automated checks.

| # | Test | Expected |
|---|---|---|
| 1 | 3 items: flagged, unflagged, flagged | Numbers are Q1 and Q2; middle item unnumbered; slide order matches input |
| 2 | 10 items alternating flagged/unflagged | Numbers Q1–Q5, gapless, in order |
| 3 | Item with 1 image, aspect 0.8 | `IMAGE_SIDE` |
| 4 | Item with 1 image, aspect 2.5 | `IMAGE_STACKED`, image full content width |
| 5 | Item with 3 images | Grid 2-over-1, bottom cell centered, equal cell widths |
| 6 | Item with 6 images | First 4 rendered, warning logged, no failure |
| 7 | Flagged item, 900-word body, 4 images | Multiple slides; body on slide 1; reply box on last slide only; no clipping |
| 8 | Any flagged item spanning 2 slides | Exactly one reply box; its label contains the correct `Q{n}` |
| 9 | Unflagged item | No reply box, no number, anywhere on its slides |
| 10 | 2 queries only | No contents slide; reply instruction on cover |
| 11 | 4 queries | Contents slide present, listing exactly 4 entries |
| 12 | 4 queries + 2 updates | Contents lists 4 queries and the line `plus 2 project updates` |
| 13 | Zero flagged items | No contents slide, no reply instruction, no numbers |
| 14 | Empty `items` array | Cover only, exits cleanly |
| 15 | Cover with all optional fields null | Renders cleanly, no gaps or orphaned labels |
| 16 | Broken image path | Placeholder rectangle, warning logged, no crash |
| 17 | 5 consecutive text-only updates | Grouped max 3 per slide, order preserved |
| 18 | Text-only update between two flagged items | Not grouped with anything |
| 19 | Same input rendered twice | Byte-identical slide count, order, and numbering |
| 20 | Every generated deck | Passes PPTX schema validation; no text overflows any container |

---

## 15. Implementation notes

Written against `pptxgenjs`. These are known failure modes that produce corrupt or silently
wrong output.

- **Set the presentation layout to 13.333in × 7.5in before adding any slide.** The default
  canvas is 10in × 5.625in. Out-of-canvas coordinates are written rather than clamped, so
  elements simply do not appear.
- **Hex colors take no `#` and no alpha channel.** Both corrupt the file. Use the library's
  transparency option for translucency.
- **Option objects are mutated in place** on first use. Build a fresh options object for
  every `add*` call; never share one across two calls.
- **Every text call needs the text-box flag set**, or screen readers announce the text as a
  graphic.
- **Text boxes carry built-in internal padding.** Set the text margin to 0 wherever text
  must align with an image edge, box edge, or another text block at the same x.
- **Rounded-corner radius applies only to the rounded-rectangle shape**, not the plain
  rectangle. The reply box must use the rounded-rectangle shape.
- **One presentation instance per output file.** Never reuse.
- **Shadow offsets must be non-negative** — negative values corrupt the file. This format
  uses no shadows, so simply omit them.
- **Never reorder the children of the presentation element** in generated XML.
- Run PPTX schema validation on every generated deck as part of the test suite, not just
  manually.

### 15.1 Rendering pipeline

1. Validate input against the §4 schema. Reject malformed input with a clear error before
   any rendering begins.
2. Assign query numbers across flagged items in array order.
3. Resolve and preprocess images: read dimensions, compute aspect, downscale, mark
   unreadable ones.
4. Plan slides — select layouts, run the overflow cascade, apply update grouping. **This
   stage produces a complete slide plan as data, before any drawing occurs.** Keep it a
   pure function; it is where every rule in §9 through §12 lives, and it is what the
   acceptance tests should exercise directly.
5. Render the plan.
6. Validate the output file.

Separating planning from rendering matters more than any other structural decision here.
Nearly all the logic in this specification is layout planning, and it must be testable
without opening a `.pptx`.
