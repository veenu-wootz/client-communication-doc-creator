# Client Communication Doc Creator — Decisions

**This is the living decision record.** Conceptual, product, and technical decisions all live
here. When something changes, edit this file — don't start a second doc. Anything marked
OPEN is not yet decided.

Related: [`docs/query-doc-format-spec.md`](docs/query-doc-format-spec.md) is the format specification. Where this file and the
spec disagree, **this file wins** — the deviations are listed in §5 with reasons.

Last updated: 2026-09-05 · built through M5, plus two review rounds against real content; Glide endpoint unwired pending credentials

---

## 1. The problem

Sales/RFQ people (at RFQ stage) and Operations people (at manufacturing stage) need to ask
clients questions and give them updates. The real unit of progress is:

> *"I have information to communicate"* → *"I got the response from the client"*

Three observed failure patterns:

1. **Formatting costs time.** Everyone invents their own format — loose email text, marked-up
   screenshots, ad-hoc PPTs, Excel. Compiling something presentable is the work.
2. **Replies come back tangled.** Clients answer several unrelated things in one mail, so the
   response has to be untangled by hand before it's usable.
3. **Nothing is recorded formally.** Tracking is hard; sometimes even the sender has to dig
   through mail to find what they sent.

**Scope for v1:** pattern 1 directly. Pattern 2 indirectly — numbered queries make replies
attributable, so a client answering "Q1: R4 governs, Q3: use Rev B" has given structured
answers. Pattern 3 is **out of scope**, though writing the file URL back to the Glide row
lays groundwork.

---

## 2. Product decisions

| # | Decision | Reasoning |
|---|---|---|
| P1 | One ordered stream of items, not separate "queries" and "updates" sections | They're entangled in practice — the same image often carries both. Splitting them would reorder the sender's narrative. |
| P2 | Query numbering (`Q1, Q2, Q3…`) is the core mechanism | The document's one job is making it unambiguous how many answers are owed and which is which. Where any layout preference conflicts with numbering clarity, numbering wins. |
| P3 | Updates are never numbered and never itemised on the contents slide | Listing them dilutes the "how many answers are owed" signal. |
| P4 | The reply box is a drawn rectangle, never a form field | The recipient has no relationship with this product. They will not install, sign in, or learn a convention — they read a deck and reply by email. The box signals that something is owed, the way it does on a paper RFI sheet. |
| P5 | **Never lose content.** No truncation, ellipsis, clipping, or dropped images anywhere | Context matters to the client. When content doesn't fit, it flows to another slide the way a person would do it by hand. The sender must never have to fix the deck before sending it. |
| P6 | Email goes to the **sender**, who forwards it | The client's reply has to land in the sender's inbox, in their own thread. An automated send to a client address is also a wrong-recipient accident waiting to happen. |
| P7 | Output is `.pptx` only | The sender can tweak a slide before sending, and PowerPoint's own Save-as-PDF covers the PDF case. No server-side conversion, no paid service. |
| P8 | Regeneration, not editing | If something's wrong, fix the form and resubmit. |

**Explicit non-goals (v1):** fillable form fields, parsing returned decks, revision markers,
prompting the user to split a long query, AI inside the renderer, tables, editing generated
output, client-side response capture.

---

## 3. Technical decisions

| # | Decision | Reasoning |
|---|---|---|
| T1 | **pptxgenjs** for PPTX generation | Pure JS, MIT, no native deps, no paid tier. Runs unchanged on Render or AWS. |
| T2 | **Plan/render split.** A pure `planSlides(document) → SlidePlan` with every box precomputed in inches; the renderer just draws it | Nearly all the logic is layout planning, and it must be testable without opening a `.pptx`. This matters more than any other structural choice here. |
| T3 | **Arial**, everywhere | The only professional sans-serif guaranteed on both Windows and macOS with no Office install and no font embedding — the client sees what you see. Calibri only reaches a Mac via Office; Helvetica is Mac-only; Segoe UI is Windows-only; Aptos is unevenly deployed. Embedding a brand font risks substitution on a document whose job is to be read correctly by a stranger. |
| T4 | **Simple character-width fitting**, not real font metrics | See §4. Chosen for leanness after the no-truncation rule made precision unnecessary. |
| T5 | Text boxes are sized to the **full remaining content area**, not to measured height | Estimation error then appears as whitespace inside a box that has room, never as text spilling over an edge. |
| T6 | **S3 upload is optional** — runs only if all AWS env vars are present, otherwise skipped with a log line | Infra may not be available. The email attachment is the guaranteed delivery path; S3 is an enhancement. |
| T7 | **Glide write-back is optional**, and only attempted if an S3 URL exists | Nothing to write without a URL. Code ships dormant until credentials arrive. |
| T8 | **LLM enrichment behind a one-method interface**, Anthropic and OpenAI adapters both shipped | `deriveItemMeta(items) → [{title, needs_response}]`. Provider picked by `ENRICH_PROVIDER`, defaulting to whichever API key is in the environment. |
| T9 | Enrichment failure is **never fatal** | Falls back: Strike-supplied value → keyword heuristic → `needs_response: false` and title from the first 60 chars. |
| T10 | `parseStrikePayload.js` is the **only** file that knows we're on Glide | Vendor-specific URL shaping or field naming goes there and nowhere else. Notably, the sibling project's `appsheetFileUrl()` (reconstructing URLs from AppSheet file IDs) is **not** ported. |
| T11 | `node:test` for tests, no framework | Zero install, works on Node 26. |
| T12 | Webhook ACKs immediately, then processes in background | Prevents Strike/Glide workflow timeouts and retry loops. A `/preview` route runs the same pipeline synchronously and returns the deck, for testing without SMTP or S3. |
| T13 | The cover carries the prominent logo top-right; every other slide carries the small logo in the footer | Two logos on the cover reads as a mistake. Every page has the mark, which is what was asked for. |
| T15 | **AI may repair body formatting**, guarded by a whitespace-only check | Senders type numbered points into one box without pressing Enter, so lists arrive as run-on prose. The enrichment call returns a line-broken version; it is used only if it matches the original word-for-word ignoring whitespace, so a model that rewrites anything gets discarded and the sender's raw text renders instead. Optional like the rest — no key means the text passes through untouched. |
| T16 | **Images have a resolution floor** (`MIN_RENDER_DPI = 96`) | Contain-fit sizes an image to its box regardless of how few pixels it has, so a 200px screenshot was being stretched across a 5in column at ~40 DPI. A drawing nobody can read defeats the point of attaching it, so a low-resolution source now renders smaller and sharp. Downscaling an oversized image already worked correctly and is unchanged. |
| T17 | The image column in the side layouts stays a **fixed** split | Widening it when the body is short was tried and reverted. `IMAGE_SIDE` only receives images with aspect < 1.4, and in a 4.15in-tall column those are already height-bound — measured gain was 0.00in at aspect 1.0 and 0.07in at 1.3. Inert complexity, so it went. |
| T14 | Fixtures are a module (`test/fixtures.js`), not loose JSON files | They share builders, so they cannot drift apart. `npm run sample` writes each one out as JSON too, so any of them can be POSTed at the running service. |

### Reused from the sibling `qualityinspectionreport` project

(Internal, not vendored here — referenced for provenance.)

`awsUpload.js` (near-verbatim, plus an `isConfigured()` guard) · `sendEmail.js`'s
`getTransporter()` lazy singleton and office365/gmail switch · `server.js`'s Express +
immediate-ACK shape · `generateQIR.js`'s image fetch-and-encode and its contain-fit math ·
`appsheetRows.js` as the structural template for `glideWrite.js` · the Cloudinary Wootz logo
URL.

**Bug not to port:** the sibling `server.js` catch block calls `res.status(500).json()` after
`res.json()` has already responded, which throws `ERR_HTTP_HEADERS_SENT`. Guarded here with
`res.headersSent`.

---

## 4. How text fitting works

The mechanism, and why it's this one rather than something more precise:

- A **character-width table** (~10 lines, no dependency): narrow glyphs (`iljt.,;'`) ≈ 0.28em,
  wide (`mwMW`) ≈ 0.83em, caps and digits ≈ 0.56em, default lowercase ≈ 0.50em. Better than a
  flat average on real content like part numbers and dimensions.
- A conservative **`FILL_TARGET = 0.85`** — we plan to fill 85% of a box, then break to the
  next slide.
- Body text steps **16 → 14pt** before breaking to a new slide.

**Why not real font metrics.** Parsing glyph advances out of a bundled TTF is more accurate,
but the accuracy has a long tail — kerning, ligatures, bold vs regular tables, non-ASCII
fallback, and PowerPoint's own line-breaking differing from ours. Each a small fix, and there
are many. That precision only mattered while a wrong estimate meant *losing text*; under P5 a
wrong estimate costs whitespace instead.

**The tradeoff, stated plainly:** a long item may break to a second slide slightly earlier
than it strictly needed to. One constant to tune.

**Escape hatch:** `HARD_WRAP` (off) emits our own line breaks instead of letting PowerPoint
wrap — a stronger no-overflow guarantee at the cost of copy-paste quality. It exists so that
if real decks ever show spill, the fix is a flag flip rather than a redesign.

---

## 5. Deviations from the format spec

Section numbers below refer to [`docs/query-doc-format-spec.md`](docs/query-doc-format-spec.md).

The spec is the format authority; these are the places we knowingly differ, and why.

### 5.1 Every truncation rule becomes a flow rule (per P5)

The spec reached for truncation to guarantee the layout cascade terminates. We'd rather spend
a slide than a sentence.

| Spec | Here |
|---|---|
| §11 item capped at 3 slides, surplus images dropped | No cap. An item takes as many slides as its content needs. |
| §11 Stage 4 truncates body at the font floor | Body continues onto the next slide, above that slide's images. |
| §7.2 contents entries truncate to one line | Entries wrap to two lines; the list paginates to another slide when full. |
| §8.1 header title truncates with ellipsis | Header wraps to at most two lines; the content area starts lower on that slide. The `Q{n}` chip is never wrapped or shrunk. |
| §9.6 grid body capped at 3 lines with ellipsis | Body takes what it needs; the remainder flows onward. |
| §9.7 caption truncated to one line | Caption wraps to two lines. |
| §4.1 `additional_details` truncated at 240 chars | Rendered in full. |
| §13 more than 4 images per item | All render, across as many slides as needed. (The form caps at 4 anyway.) |

Termination is still guaranteed: content is finite, and every continuation slide consumes at
least one image or one line of body.

### 5.2 Other deviations

| Spec | Here | Why |
|---|---|---|
| §8.2 footer reads `Page {i} of {total}` | `Page {i}` only | Requested. |
| §5.4 "no decorative iconography" | Wootz logo on the cover **and** every slide footer | A company mark is identification, not decoration. Requested. |
| §5.2 lists Calibri or Arial | Arial, fixed | T3. |
| §4 marks several fields required | Nothing is required; every field has a documented fallback | Requested — the form shouldn't force values. §7.1 already demands the cover compose with all optional blocks absent. |
| §14 #19 "byte-identical" output | Determinism asserted on the slide plan and slide XML, excluding `docProps/` | A `.pptx` embeds a save-timestamp, so raw bytes always differ even when every slide is identical. The intent — a generator that doesn't quietly vary run to run — is fully tested. |
| §5.1 "no element outside `MARGIN`" | Footer and logo are documented exceptions | The spec's own `FOOTER_Y + FOOTER_H = 7.30` already sits inside the bottom margin. |

### 5.3 Gaps the spec left open, and how they're closed

| Gap | Resolution |
|---|---|
| Contents slide has no entry cap and would overflow | Paginates to another contents slide. Heading on the first, reply instruction on the last. |
| §11 Stage 4 has no terminating guard | Moot — no cap, and termination follows from finite content (§5.1). |
| Page numbering undefined for cover/contents | Cover and contents consume numbers but show the cover footer with no page indicator, so item slides visibly start at 2 or 3. |
| pptxgenjs has no text-measurement API, and its `autoFit` has no shrink-on-overflow mode | §4. |

### 5.4 Layout calls made while building

Three judgment calls the spec did not settle. Each was forced by geometry, and each is one
constant away from being reversed.

**The `MIN_IMAGE_W` floor applies to the grid *cell*, not the rendered image.** Read
literally, spec §9.6 tests the rendered width — but on a 7.5in slide a two-row grid gives
rows about 1.6in tall, so any square or upright image renders under 3.0in wide and the 3–4
image grid becomes unreachable for anything but wide drawings. That contradicts the layout
in the third reference screenshot. The floor now tests the cell the image is placed into: a
square photo is short because of its own aspect, not because the layout squeezed it.
*Consequence:* a 3–4 image query fits on one slide, with images around 1.6in tall. If that
reads too small in a real deck, reverting is a one-line change in `tryLayout` and those items
then spread over more slides with bigger images.

**The image column in `IMAGE_SIDE_PAIR` is 60% wide, not the 45% of `IMAGE_SIDE`.** At 45%
each of the two cells is 2.5in — below `MIN_IMAGE_W`, so the pair layout could never fire and
two upright images always split across slides. Two images need more room than one.

**The stacked image zone grows past `IMAGE_ZONE_FRACTION` when the body is short**, up to
`IMAGE_ZONE_MAX` (80%). A flat 55% cap left a wide drawing only marginally wider than the
side column it was rescued from, which defeats the point of §9.3. A 2.5-aspect drawing now
renders about 7.5in wide instead of 5.6in. It still cannot reach the full 12.1in content
width on a reply-bearing slide — that needs 4.85in of height and only 4.05in exists — and
that is geometry, not a bug.

### 5.5 Changes from the first real-content review (2026-09-04)

Running a real RFQ through the pipeline surfaced four things the fixtures never
would have. All are applied.

| Finding | Fix |
|---|---|
| On the contents slide the logo and the "Prepared by…" line printed on top of each other | `coverFooter()` was not logo-aware, so its text box started at `CONTENT_X` — exactly where the bottom-left logo sits. Item slides already offset for this; the contents slide now does the same. |
| The author's name on the contents slide is redundant | The logo already says who sent it. The contents footer now carries the date only; the cover keeps the full "Prepared by …" line. |
| "4 questions for you" reads as though the questions themselves follow | It is an index, not the questions. Heading is now "N questions in this document" (`CONTENTS_HEADING` in constants). |
| Numbered points typed without line breaks rendered as one run-on paragraph, hiding the match between the text's numbers and the drawing's callouts | T15. Note the renderer already preserved real line breaks correctly — verified — so this only affects text that arrives with none. |

Two spacing changes came out of the same review: the footer moved down 0.10in
(`FOOTER_Y` 6.95 → 7.05, `FOOTER_H` 0.35 → 0.30) with the reply box and content
bottom following it, which hands that 0.10in to the content zone rather than
leaving it as dead margin at the bottom of the slide.

**On the callout-number correlation.** The drawings in the real document carry
numbered callouts (1, 2, 3…) that line up with the numbered points in the query
text — that is how the sender relates a question to a place on the drawing. It
is handled purely as text formatting (T15). Query numbering is untouched: those
in-text numbers are the sender's own, and never become `Q` numbers.

### 5.6 Field model and the real Glide payload (2026-09-04)

**Fields.** `project_name` → `report_title`; `part_number` → `reference_name`;
`po_number` dropped. One free-text reference carries a part number, a PO number,
or a name — fewer fields for the sender to fill, and the cover shows a single
`Reference` line. A new per-row `item_name` names the SKU a query is about,
because one document routinely spans several: it renders as a muted line above
the slide header and prefixes the contents entry. The old names still map
through as aliases, so existing payloads keep working.

**What the real Glide export actually looks like**, and what each part forced:

| Glide sends | Handling |
|---|---|
| `items` as an **array containing one string**, with the rows joined by `", "` | This is the one that mattered. The old parser read fields off a string, got nothing, and filtered every row away — a cover-only deck with all five queries **silently gone**. Rows are now recovered with a brace-depth scanner that respects quotes and escapes, so a description containing a comma is not torn in half. |
| CSV-doubled quotes when copied out of a cell | Un-doubled and re-parsed as a fallback. |
| Markdown (`**bold**`) from the rich text editor | Emphasis markers stripped — they would otherwise print as literal asterisks to the client. Only the markers go; every word stays. |
| List padding (`\n    \n` between points) | Trailing spaces dropped, runs of blank lines collapsed. The line breaks themselves survive, so numbered lists render one point per line with no AI involved. |
| Images as bare URL strings | Already supported; also handles a JSON array arriving as a string. |

A row that cannot be parsed is now **reported as a warning instead of dropped**,
and a payload yielding no rows says so. Silent loss was the failure mode worth
engineering against: a deck that is quietly missing four of five questions looks
perfectly fine to the person sending it.

**Title fallbacks take the first line, not the first 60 characters.** Bodies open
with a lead-in and then a numbered list, so a flat character slice produced
headers containing line breaks and half a list item — and appended an ellipsis,
violating the no-truncation rule the rest of the system is tested against. Both
fallbacks (`resolveTitle` in the planner, `guessTitle` in enrichment) now take
the first line and cut on a word boundary with no ellipsis.

### 5.8 Sections, and labelling every slide (2026-09-04)

**The problem, seen in a real deck.** With queries numbered and updates left
bare, a client flicking through hit `Q1`, then an unnumbered slide, then `Q2`,
then another unnumbered one — and had to infer the convention with nobody to
ask. The format was only self-explanatory to someone who had it explained.

**Two changes, both deliberate overrides of the spec.**

*Every item slide now carries a chip.* A query shows its number loudly
(`accent` on `accent_bg`, 22pt); an update says `Update` quietly (`muted` on
`surface`, 12pt). Grouped update slides carry it too. No slide is ever
ambiguous. This overrides §8.1, which gave unflagged items a bare title.

*Items are grouped into sections* — everything owing an answer first, then the
rest, with the sender's order kept inside each. This overrides invariant §3.1,
"item order is preserved absolutely". That rule existed to protect numbering
clarity; in practice it worked against it. Grouping serves the original goal
better than the original rule did.

| Cost | Assessment |
|---|---|
| The sender's narrative order is lost across sections | Accepted. For a client-facing document, "what must I answer" beats "the order it was written in". Order within each section is untouched. |
| A query row may contain updates inside it, and lands in the queries section | Accepted by the user as a manageable trade. Splitting a row is the fancy path and stays out: a row's images belong to the row as a whole, so there is no honest way to divide them. |
| **Bonus:** updates become contiguous | §12 grouping finally fires. On the real payload the deck went from 9 slides to 8; on the alternating fixture, five updates now share slides instead of taking one each. |

**The summary slide, reworked.** It was headed "N questions in this document"
and mentioned updates only as "plus 2 project updates". It now reads:

```
Summary

QUERIES — 3 need your response
Q1 · HEX NUT M12 — Thread pitch out of spec
...
OTHER UPDATES
· Pilot batch on the bench
...

Please reply quoting the query number.
```

Updates are named rather than counted, per §5.6's finding that a client should
never meet a slide the summary did not mention. The count moves into the
QUERIES label so the "how many answers are owed" signal survives — it is the
first line read, above everything else.

Section labels and entries lay out as one flat list of rows, so pagination stays
the same "fill until full" loop, with a guard against stranding a section label
at the foot of a slide.

**The summary threshold counts every item, not just queries**
(`INCLUDE_SUMMARY_MIN_ITEMS = 3`). Gating on queries alone meant a document of
two queries and six updates got no summary at all, even though the summary now
names updates. A deck with no queries gets a summary but no "reply quoting the
query number" line — nothing is owed.

**Sub-numbering inside a row (1.1, 1.2) was considered and dropped.** Senders
already number their own points and those numbers render as written, so a client
can reply "Q1, point 2" unaided. Having AI renumber them means rewriting the
sender's text, which the guardrail exists to prevent.

### 5.9 Second review round (2026-09-04)

| Feedback | Change |
|---|---|
| The summary's footer differed from every other slide's | It now uses the same `itemFooter` — report title · reference · date, with `Page N`. The cover keeps its own "Prepared by …" footer and stays unnumbered. |
| QUERIES and OTHER UPDATES sat too close together | A section label that follows a list gets `SECTION_LEAD` (0.34in) above it; the first one, sitting under the heading, gets none. |
| The summary heading looked oddly low | It was anchored at `CONTENT_Y` (1.50in) while every other slide puts its title at `HEADER_Y` (0.55in) — almost an inch lower, which read as a mistake. Now at `HEADER_Y`. It was never going to "auto-move"; the position was fixed. |
| A reader could not tell a query continued overleaf until they had already turned the page | A slide that carries on now says so at its foot ("Continued on the next slide →"). `contentBox` reserves `CONTINUES_H` for it when the planner knows the slide continues, so it never collides with content, the reply box, or the footer. |
| Email heading | Subject is now `PPT — {report_title} | …` and the body heading `PPT ready to send`. |

**A latent bug the continuation work exposed.** `imageInCell` counted
`CAPTION_GAP` twice — once reserving space in the cell, once positioning the
caption box — so a captioned image overflowed its cell by exactly that gap. It
went unnoticed while nothing sat directly beneath the content area; the moment
the continuation note did, the geometry test caught it. Fixed by giving the
caption box the text height only, since the gap is already spent on placement.

### 5.11 SMTP_FROM_NAME is a no-op on Gmail relay (2026-09-04)

Confirmed by testing: setting `SMTP_FROM_NAME` while `SMTP_SERVICE=gmail` has no
effect. Two different values (`Wootz.Strike`, then `Wootz.Work`) both sent
correctly per our own code — verified by printing the loaded env var
immediately before the send — and both arrived displaying as `"Wootz"`
regardless. Gmail's SMTP relay overrides the `From` display name with the
sending Google Account's own "Send mail as" name, independent of the header
the application sends; this is Gmail-side anti-spoofing enforcement, not
something nodemailer or this codebase controls.

The variable still works correctly on `office365`, where the mismatch it was
built to fix (PLAN.md §5.6) actually applies. Documented in `.env.example` so
this isn't rediscovered.

### 5.12 OneDrive as the file store (2026-09-05)

Decks are filed to OneDrive/SharePoint via Microsoft Graph. There is no Graph
API for authoring PowerPoint content — Excel has a workbook API, PowerPoint has
no equivalent — so generating the file here and uploading it is the only real
approach, not a workaround.

| Decision | Reasoning |
|---|---|
| **OneDrive is primary, S3 is the fallback** | S3 was never configured in practice (every run logged "upload: skipped") and needs an AWS account. OneDrive puts the deck where the team already works. S3 code stays, costs nothing, and runs only if OneDrive produces no URL. |
| **The destination comes from the payload**, `GRAPH_*` env vars are the fallback | Strike sends `drive_id` and `folder_item_id` per submission, so each report lands in its own project folder. Verified end-to-end with the env defaults deliberately unset. |
| `conflictBehavior=rename` | Two submissions for the same project on the same day would otherwise silently overwrite each other, and a client deck lost that way is hard to notice. |
| Simple `PUT`, no upload session | Decks run 57–900 KB, far below the threshold where chunked upload is needed. |

**Scope of the credentials, measured not assumed.** The app registration carries
`Sites.Selected`, not `Files.ReadWrite.All`. Tenant-wide site search returns
`accessDenied`, so it cannot roam the tenant — but `Sites.Selected` is scoped to
a **site**, not a folder, and the granted site is a personal OneDrive whose root
lists 94 items. So the app can write anywhere in that drive; the folder id is an
address, not a boundary. This scope pre-existed this project (another service
writes PDFs with the same registration). Narrowing it would mean a dedicated
SharePoint site for tool output — noted in §8 rather than assumed.

The returned `webUrl` opens the deck in PowerPoint Online in edit mode, so the
sender can adjust a slide before forwarding. It is access-controlled, not a
public link — fine for internal use, but the URL written back to Glide only
opens for people with access to that drive.

### 5.13 Filename and project_name (2026-09-05)

Decks are filed as **`Queries - {project_name} - {DD Mon YYYY H.MM AM/PM}.pptx`**,
the timestamp being generation time in IST — not the submission's `created_at`,
since the name records when the deck was produced.

`project_name` and `report_title` are **independent fields and do not alias each
other**: `report_title` is what the deck displays, `project_name` is what the
file is named. Whatever the payload sends for each is used for that purpose and
nothing else.

The consequence is that Strike must send **both**. A payload carrying only
`project_name` leaves `report_title` null, and the cover then renders without a
heading — which composes cleanly (§7.1 requires it) but is not what anyone
wants. The filename keeps a `project_name || report_title` fallback, since a
file called "Queries - document - …" would be worse than borrowing the title.

Two details worth keeping: the time separator is a dot because a colon is
illegal in OneDrive/SharePoint filenames, and months are formatted with the same
three-letter list `formatDate` uses, so the filename and the deck's own footer
can't disagree. `en-GB` renders September as "Sept", which is why the month is
mapped explicitly rather than taken from the locale.

### 5.14 Flat format — from finished document to adoptable template (2026-09-05)

The generator was correct but hard to *adopt*. A sender who wanted to add a slide
had to renumber `Q1..Qn` by hand, could not press Enter in the summary because
every entry was its own text box, and had no way to match the generated styling.
The segregated format was excellent for a document nobody touches.

**What changed, and why each.**

| Change | Reason |
|---|---|
| One sequential stream `1..N`, input order, no sections | Attributability never needed the split — "3: R4 governs" is as unambiguous as "Q3". Sections meant a hand-inserted slide landed in the wrong place. |
| Summary is **one text box with `bullet: {type:'number'}`** | The single biggest adoption blocker. Real PowerPoint auto-numbering: Enter adds the next point. Pagination is kept, with `numberStartAt` continuing the count across slides. |
| No "3 need a reply" count | It goes stale the moment a slide is added by hand, and a confidently wrong count is worse than none. |
| Reply box on **every** item | Uniformity: any slide can be duplicated and still look right. |
| Grouping removed | Same reason — every item gets its own slide so they are interchangeable. |
| Image placeholder when an item has no picture | Shows where one goes. Only when the ITEM has none: a continuation slide that merely ran out would be telling the reader something untrue. |
| Two numbered template slides appended | Built on a slide master with **native placeholders**, which prompt while editing and print nothing if untouched — safe to ship unused. |
| Cover is a template | Field labels always render; an empty value gets a hairline rule to write on rather than vanishing. A rule needs no deleting; a filled box would. |
| Prepared-by moved out of the footer into a cover field | At 9pt in the footer nobody read it. |
| `coverLabel 10 → 12`, `coverValue 14 → 18` | The cover reads as a form, not fine print. Title stays 32 so the hierarchy holds. |

**Costs accepted knowingly.** Decks are longer — a reply box everywhere costs
content height on every slide, and no grouping means five updates take five
slides. `REPLY_H` dropped `1.10 → 0.85` and `REPLY_GAP` `0.20 → 0.15` to return
0.30in per slide, which offsets part of it. And an unused per-slide image
placeholder *does* print, unlike the template ones: it has to be drawn by the
planner to follow the adaptive layout, and native placeholders live at fixed
positions on a master. It is outline-only in the border colour to keep that quiet.

**What did not change: the adaptive layout.** `selectLayout`, the overflow
cascade, the aspect rules and the grid all behave exactly as before wherever
images exist. Only the no-image case was given a reserved column.

`needs_response` is still carried — enrichment sets it and the
`grouped-queries-format` branch needs it — but nothing in the layout reads it.

**The previous format is preserved** on branch `grouped-queries-format` in the
personal repo, for the eventual switch back once the format is established.

### 5.15 Glide write-back (2026-09-05)

After a deck is generated the RFQ row is updated with five columns, addressed by
Glide's internal column ids via `mutateTables` / `set-columns-in-row`:

| Column | id | Value |
|---|---|---|
| RFQ Version Used | `YU0gy` | `rfq_version` from the payload, echoed back |
| Query PPT File ID | `RcHZF` | the OneDrive **DriveItem id** |
| Query PPT File Link | `g0KAH` | the deck's `webUrl` |
| Query PPT Generated On | `TjmbZ` | ISO 8601 UTC |
| Query PPT Generated By | `wVeBR` | the sender's email |

**Why the DriveItem id and not the SharePoint GUID.** Both were tested against
the live tenant and *both* work — `GET /drives/{drive}/items/{id}` accepts either,
and the DriveItem id also downloads via `/content`. The DriveItem id wins because
the upload response hands it to us directly, whereas the GUID has to be parsed
out of a `webUrl` query string or fetched with an extra call. Fetching a file
later needs the drive id alongside it, which Strike already holds.

**Generated On is ISO, not the filename's display string.** The Glide column is a
date, and `"05 Sep 2026 8.34 PM"` would neither sort nor filter there. The
readable form stays in the filename, where a person actually reads it. One
`generatedAt` instant is taken at the top of `build()` and used for both, so they
cannot disagree.

**`QZRyl` must never be written.** It is the app's own *RFQ Folder Drive Item
ID*, populated by Strike. The version was mapped to it by mistake, so runs wrote
`V1`…`V5` over real folder ids and a later run blanked one. The version column
is `YU0gy`; a test asserts nothing this service sends ever names `QZRyl`.

**All five columns are written every time, blank included.** They are written
only by this service and never by hand, so omitting a blank left the previous
run's answer beside a freshly generated file — a stale version reads as fact.
Nothing is derived into them either: `generatedBy` used to fall back to
`created_by`, which would put a person's *name* in a column meant for an email
address. A gap is now visibly a gap.

Verified end to end against the live table: five columns written, then read back
with `queryTables` to confirm they actually landed — a 200 alone proves nothing,
since a wrong column id is a silent no-op.

### 5.16 Placeholders and summary numbering, corrected (2026-09-05)

Two defects found by looking at a real deck rather than at the plan data.

**The summary was not numbering.** It rendered "1.", then an unnumbered line,
then "1." again. Cause, from the XML: pptxgenjs applies a *shape-level* `bullet`
to the first paragraph only and writes `<a:buNone/>` on every one after it. The
bullet has to sit on each run. A second trap sat behind it — pptxgenjs always
emits a `startAt` attribute, defaulting to `1`, with no way to omit it, so a list
whose paragraphs all said `startAt="1"` would render 1, 1, 1. Each paragraph now
declares its own number, which is deterministic.

**Placeholders are now real PowerPoint placeholders everywhere** — on the cover
when no product photo is supplied, on any item slide without a picture, and on
both templates. The drawn dashed box they replace failed on both counts the
sender cared about: it offered no way to insert an image, and it had to be
deleted by hand if unused.

This needed a workaround. `defineSlideMaster` accepts a placeholder `type`, but
pptxgenjs's `PLACEHOLDER_TYPES` table is **empty**, so the branch that would emit
`type="pic"` never fires and every placeholder it produces is a generic one — it
vanishes when unused, but shows no insert icon. The attribute is therefore
injected into our own `PIC_*` layouts after `write()`, and the injection is
asserted rather than assumed: a pptxgenjs upgrade that changes this XML must fail
loudly rather than quietly ship placeholders nobody can click. `jszip` moved to
`dependencies` for it.

Placeholders live on masters, which are fixed-position, so one master is defined
per distinct geometry the deck actually needs — typically three, deduplicated by
`masterKeyFor`. The adaptive layout is untouched: a slide with a real image gets
no master and behaves exactly as before.

**One placeholder per generated item slide**, since a second nudge on the same
page adds no information. The templates are worked examples rather than generated
output, so they keep their original shapes: one with a single picture beside the
text, one with the three-image grid.

**The reply area is a full-size text box.** Its label previously occupied a
one-line strip at the top of the box, so clicking the reply area gave a sliver to
type into rather than the space the box appeared to offer.

Decks also got markedly smaller (110-180 KB down to 41-68 KB), since the drawn
rectangles and their labels are gone.

### 5.17 The legibility floor, corrected (2026-09-05)

A real deck put three drawings on one slide, each rendering about 1.3in wide —
too small to read a dimension on. Three hypotheses were checked and all three
were wrong: placeholders were not attached to any slide that has images,
continuation and the cascade were both working, and measuring the same input on
both branches showed images marginally **larger** on the flat format
(1.26/2.54/3.19in) than on the grouped one (1.16/2.33/2.92in). The flat format
did not cause it.

The cause was older than either: the viability floor tested the **cell** an image
sits in rather than the image as **rendered**. A 2-column grid cell is 5.9in wide
and always passes, while a portrait drawing inside it comes out 1.26in across.
That check was changed to the cell during the first build, to stop the 3–4 image
grid being unreachable for square images — the wrong trade, since it bought a
grid nobody can read.

The floor now measures what the reader actually sees. The same item spreads over
three slides at 3.64in, 6.30in and 7.91in — three to six times larger — with the
continuation notes that implies. A lone image is exempt: it goes on its slide at
whatever size it comes out, because there is nothing left to reduce.

The grid is not gone, only earned: it fires when images are wide enough to stay
legible in it (aspect of roughly 1.7 or more), which the wide-image fixture
covers. Square drawings now pair up two per slide at 3.4in instead.

### 5.18 Known consequence, not a bug

A long body pushes an item across several slides — 900 words with four images plans to five
item slides, one image each. That is the no-truncation rule working as intended: the
alternative is cutting the sender's text. Don't "fix" it by adding a slide cap.

`GRID_MAX_CELLS = 4` exists because the grid is two rows. A fifth image goes to the next
slide rather than growing a third row off the bottom of the content area — which is exactly
what it did before the cap, and the geometry test caught it.

---

## 6. Pipeline

```
POST /generate
  → parseStrikePayload   Strike/Glide JSON → canonical document JSON (tolerant)
  → enrich               one LLM call: title + needs_response for all items
  → prepareImages        fetch, downscale to 1600px, read dimensions/aspect
  → planSlides           PURE. Emits SlidePlan + warnings.
  → renderPptx           SlidePlan → pptxgenjs → Buffer
  → sendEmail            always — deck goes to the sender, attached
  → uploadToS3           only if AWS env vars present
  → writeToGlide         only if the upload produced a URL
```

---

## 7. Testing

**Tier 1 — planner.** Fast, no `.pptx`. The spec's §14 cases adjusted for the flow rules,
plus: a 900-word body with 4 images places every word and every image; a 40-query document
paginates its contents with no entry shortened.

**Tier 2 — PPTX structure.** Unzip, parse slide XML, assert: slide count matches the plan;
every shape lies inside the canvas and within `MARGIN` (footer and logo whitelisted); chip
text, footers, `Page {i}` and `Your response — Q{n}` labels correct; exactly one reply box per
flagged item and none on unflagged ones; no `Q\d+` on an unflagged item's slides; **no
ellipsis character anywhere in the deck**; and the concatenated body text across an item's
slides equals the input exactly.

**Tier 3 — export check.** `npm run sample` renders every fixture to `out/`. Open one in
PowerPoint, Save-as-PDF, confirm fidelity by eye. Deliberately manual — a real rendering
engine is the only honest way to verify what PowerPoint will do.

---

## 8. Open items

| Item | Status |
|---|---|
| Glide API key, table/row identifiers, target column | **OPEN** — both Glide API surfaces are implemented in `glideWrite.js` and selected by `GLIDE_API`. Setting the env vars is the whole wiring. |
| Real Strike payload JSON | **OPEN** — `parseStrikePayload.js` opens with a declarative field-name map; editing the alias lists is the whole integration. |
| AWS S3 credentials | **OPEN** — optional by design (T6); absence is a logged skip, not an error. |
| Which LLM provider is the default | Resolved by environment — whichever key is present. |
| Does the 3–4 image grid read well at ~1.6in tall? | **Needs eyes on a real deck** — see §5.4. `npm run sample` then open `05-three-images.pptx`. |
| Should tool output live in its own SharePoint site? | **OPEN** — `Sites.Selected` grants at site level, so the app can currently write anywhere in the target OneDrive (94 top-level folders, including customer documents). A dedicated site would scope it to exactly what the tool needs. One-time admin ask. |
| Titles without enrichment are weak | **Expected** — with no API key the fallback yields "There are five queries" rather than "Crankshaft casting and tolerance queries". The structure is right, the wording is not. Enrichment is what fixes it. |
| Is the reply box worth 1.10in? | **OPEN** — it plus its gap take 1.30in of every query slide. Dropping it to ~0.85in would hand 0.25in (~6%) to the image. Not changed unilaterally: the box has to look like somewhere you would write. |
| Does PowerPoint's Save-as-PDF hold the layout? | **Needs a manual check** — no rendering engine here can answer it. |

## 9. Status

Built and passing 94 tests: the planner and its geometry, the rendered OOXML, and the
enrichment fallback chain. `npm run sample` renders 21 fixtures to 114 slides.

The whole pipeline runs with **no credentials at all** — no LLM key, no SMTP, no S3, no
Glide. Each unconfigured integration logs a skip and the deck is still produced. That is the
property to preserve: nothing optional may become load-bearing.
