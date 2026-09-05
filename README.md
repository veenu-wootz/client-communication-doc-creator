# Client Communication Doc Creator

Turns a Strike (Glide) form submission into a consistently formatted `.pptx` that a
supplier's RFQ or Operations POC emails to their client — replacing the ad-hoc mix of
loose email text, marked-up screenshots and hand-built slides that costs them time today.

The document's one job is to make it unambiguous **how many answers are owed and which is
which**: queries are numbered `Q1, Q2, Q3…` gaplessly, each gets a reply box, and a contents
slide states the count. Updates travel in the same stream but are never numbered.

**`PLAN.md` is the decision record** — product, format, and technical decisions, and every
place this knowingly differs from [`docs/query-doc-format-spec.md`](docs/query-doc-format-spec.md). Read it before changing
layout behaviour.

## Quick start

```bash
npm install
cp .env.example .env      # nothing in it is required to see output
npm test                  # 44 tests: planner, PPTX structure, enrichment fallbacks
npm run sample            # renders every fixture to out/
open out/07-long-body-4-images.pptx
```

Then open a deck in PowerPoint and Save-as-PDF to confirm export fidelity. That check is
deliberately manual — no rendering engine can honestly tell you what PowerPoint will do.

## Running the service

```bash
node server.js
curl -X POST localhost:3000/preview -H 'Content-Type: application/json' \
     -d @out/01-flagged-unflagged-flagged.json --output deck.pptx
```

| Route | Purpose |
|---|---|
| `GET /` | Health check |
| `POST /generate` | The webhook. ACKs immediately, then generates and delivers in the background. |
| `POST /preview` | Same pipeline, returns the `.pptx` inline. For testing without SMTP or S3. |

**Everything optional is genuinely optional.** With no API keys, no SMTP, no S3 and no Glide
credentials, `/preview` still returns a correct deck. Each integration that is not configured
logs a skip line and the pipeline continues. The startup banner shows which are live.

## How it fits together

```
POST /generate
  → parseStrikePayload   Strike/Glide JSON → canonical document      ← the only vendor-aware file
  → enrich               one batched LLM call: title + needs_response
  → prepareImages        fetch, downscale to 1600px, measure aspect
  → planSlides           PURE. All layout logic. Emits a complete slide plan.
  → renderPptx           draws the plan — no decisions of its own
  → uploadDeck           optional (S3)
  → writeDeckUrl         optional (Glide), and only if the upload produced a URL
  → sendDeckEmail        to the sender, who forwards it to the client
```

The planner/renderer split is the important one: nearly all the logic is layout planning, and
keeping it a pure function means the acceptance tests exercise it directly without opening a
`.pptx`. If something lands in the wrong place, the bug is in `src/plan/`, not `src/render/`.

## Two rules that explain most of the code

**Never lose content.** No truncation, ellipsis, clipping or dropped images anywhere. When
content does not fit, it flows onto another slide. A test asserts no ellipsis character ever
reaches a deck. The sender should never have to fix the file before sending it.

**Stay lean.** Text fitting uses a small character-width table and one conservative
`FILL_TARGET`, not real font metrics — see `PLAN.md` §4 for why that trade is the right one
here. The cost is that a long item may break to a second slide slightly sooner than it had
to. One constant tunes it.

## Deployment

Node only — no native rendering, no LibreOffice, no Docker needed. Runs on Render or AWS with
`npm install` and `node server.js`. `sharp` ships prebuilt binaries for both.

## Tests

| File | Covers |
|---|---|
| `test/planner.test.js` | The §14 acceptance cases against the slide plan: numbering, layout selection, grid geometry, reply boxes, contents pagination, grouping, determinism, and that every box sits within the margins |
| `test/pptx.test.js` | The rendered OOXML: canvas size, shape bounds, reply-box uniqueness, no query numbers on update slides, no ellipsis, well-formed XML, byte-stable slides across runs |
| `test/enrich.test.js` | The fallback chain, with no API key present |
| `test/filename.test.js` | The IST timestamp in the filename — timezone offset, midnight rollover, month format matching the deck footer, and no characters OneDrive rejects |
| `test/parse.test.js` | The Strike/Glide payload parser against shapes seen in a real export — rows joined into one string, CSV quoting, markdown, malformed rows |
