# Portable final-artifact provenance

## Identity is separate from storage

The logical identity is `(corpus_id, document_id, revision)`, bound to a final
artifact SHA-256. A claim occurrence also binds the accepted draft's SHA-256 and
canon/plan fingerprints. Physical paths live in the manifest or optional resource
bindings, not canonical business facts. Moving unchanged artifacts and rebinding
storage does not change their logical provenance.

Record exporter associations during authoring; resolve exact mappings in the
explicit QA pass **after the final OfficeCLI save and saved-cache verification**.
Reopening and saving, changing prose, rearranging slides, editing formulas, or
rewriting properties can invalidate the hash and locator even if the visible
value looks unchanged. Reindex, remap, reverify, and repeat affected QA. Do not
write artifact hashes back into the frozen canon.

## Exporter handoff

For new execution-version-2 runs, generation persists private associations under
`build\export-bindings` and completes with a manifest/generation receipt. It does
not require detailed corpus occurrence mapping or certify complete provenance.
`officecli_runner.py qa` resolves and verifies those associations, creates the
index/mappings/provenance/coverage, and prepares additional reviews.
`finalize-qa` performs an independent strict audit after actual reviews.
Do not duplicate these full-corpus passes by also invoking each CLI below.
Existing version-1 execution and direct provenance commands retain the sequence:

1. Validate the draft against the frozen canon and plan.
2. Author/save with OfficeCLI, using checked standalone batches or flushing
   run-owned residents. Reopen/check saved formulas and capture required HTML
   renders. Index only final saved bytes; no Microsoft Office step is required.
3. Record actual final artifact hashes in `build\manifest.json`.
4. Run `index-artifacts` with canon, plan, manifest, root, and an explicit
   `build\artifact-index.json` output path.
5. Associate **every `(section_index, claim_id)` use** with an exact occurrence
   in the indexed final content. Export-time tracking can help, but final bytes
   must still be verified. Never choose the first matching token blindly.
6. Save the associations as `build\claim-locations.json`; run
   `verify-provenance` and `coverage`, then the complete `audit`.

The CLI calls appear in `SKILL.md`. Complete-corpus commands require the entire
manifest. For a pilot, `artifact_containers(path, format)` and
`check_draft(canon, plan, draft)` inspect an individual file/draft without a
pretend complete manifest. Mapping remains the Office backend/agent's
responsibility; the helper does not automatically understand natural-language
claim placement.

## Exporter bindings for new runs

Put a complete `export_bindings` list on each job while authoring, rather than
calculating final text offsets by hand:

```json
{
  "document_id": "doc-0001",
  "layout_file": "build/layouts/press-release-v1.json",
  "batch_file": "build/scripts/doc-0001.json",
  "render_targets": [{"id": "all", "pages": 1}],
  "export_bindings": [
    {"section_index": 0, "claim_id": "revenue", "command_index": 2,
     "fragment": "Revenue reached $12.0 million",
     "rendering": {"scale": 1000000, "decimals": 1, "prefix": "$", "suffix": " million"}}
  ]
}
```

This is an illustrative binding, not a complete document: every section/claim use
must be represented. Choose exactly one target form per binding:

- `command_index`: zero-based index in the actual accepted atomic batch. For
  `set`, the target is its exact `path`. For `add`, the helper uses the successful
  indexed OfficeCLI result's exact output path, not the command's parent or a
  first text match. A corrective authoring batch must retain matching binding
  indices or supply a new affected revision.
- `target`: exact saved OfficeCLI DOM target for template/merge authoring. Proven
  automatic translations are Word `/body/p[@paraId=...]`, presentation
  `/slide[N]/shape[@id=...]`, and workbook `/<sheet>/<cell>`. Word paragraph IDs,
  slide relationship order/shape IDs, and workbook sheet/cell identity are read
  from the actual saved package.
- `locator`: an exact supported final indexed locator, including table cells,
  notes, header/footer paragraphs, ranges and chart data. This preserves the full
  locator contract when a backend has an explicit target association but its
  OfficeCLI DOM translation has not been proven. No manually computed span is
  needed. Do not guess table/notes DOM paths.

`fragment` optionally identifies grounded text inside that target. The helper
normalizes it, requires exactly one occurrence of the rendered claim inside the
fragment, and derives the final span. If the fragment/value appears more than
once in the bound target, supply an explicit zero-based `occurrence_index` local
to that target, or a more specific fragment. Missing/ambiguous targets or values
fail; there is no global first-token search. `rendering` uses the existing exact
numeric rules below. Derived cells/ranges automatically copy actual saved
formula evidence for independent verification.

For templates, bind the exact output target associated with the template's
content slot; the helper is not a prose/template generator and does not infer
natural-language slot semantics. If a template target cannot be resolved, supply
an exact indexed association or report the blocker instead of inventing one.

Associations are schema-checked during generation and bind accepted job,
artifact/draft hashes, commands and successful author result. Exact final-file
resolution and all-use/physical-span verification occur during QA. Alternatively,
a job may still supply explicit `occurrences`; strict final-byte checks apply
unchanged. Bindings never enter OfficeCLI command props, visible text, hidden
cells, notes or Office properties. Repeated sections still need distinct
physical occurrences; exporter automation does not relax semantic review.

## Supported locators

Copy the locator object exactly from the final artifact index. Indices are
zero-based; Office IDs and package parts are strings. Relationship resolution,
not guesses about `slide1.xml` or `sheet1.xml`, identifies content.

| Format/container | Locator fields in addition to `kind` |
|---|---|
| DOCX `paragraph` | `part`, `index`; paragraphs include referenced header/footer stories and paragraphs within tables |
| DOCX `table_cell` | `part`, `table`, `row`, `column` |
| PPTX `shape` | `part`, `slide_id`, `shape_id` |
| PPTX `slide_cell` | `part`, `slide_id`, `shape_id`, `row`, `column` |
| PPTX `notes_shape` | `part`, `slide_id`, `shape_id` |
| XLSX `cell` | `part`, `sheet_id`, `sheet`, `cell` |
| XLSX `range` | `part`, `sheet_id`, `sheet`, `range` such as `B2:B3` |
| `chart_point` | Chart `part`, `series`, `component`, `point`, plus owner context: Word `paragraph`, PPTX `slide_id`/`shape_id`, or XLSX `sheet_id`/`sheet` |

Charts support saved series-name/category/value/x/y/bubble-size cache points
(`tx`, `cat`, `val`, `xVal`, `yVal`, `bubbleSize`) and associated formulas when
present. Chart indices are version-bound, not persistent semantic IDs.

`chart_point` locates saved *data* inside a chart part (series/component/point
provenance for content claims). It is a separate, independent concept from
`chart_object_inventory()`'s chart *object* capture identity (sheet/worksheet
part/drawing part/shape ID/chart relationship/part/anchor, used to locate and
capture the chart as a rendered visual region). A chart can satisfy one without
the other: data-claim provenance does not imply the object is visually
captured, and a captured/visible chart object does not imply every cached
series point has been claim-bound. See `dependencies.md` for the supported
chart-object capture/coverage contract and `contracts.md` for the
`render_targets` chart entry schema.

Paragraph/shape/cell text concatenates split runs and normalizes whitespace.
An occurrence's `start`/`end` are offsets into that normalized text, with an
exclusive end. Table/range text joins indexed cell content in row order.
Every cell of a range must exist in the index. XLSX values are saved scalars or
formula caches, **not formatted Excel display strings**; styles, currency
labels, and units still require OfficeCLI visual and semantic QA.
Numeric cells use a normalized scalar representation: integral values retain
exact integer precision, and fractional values use the shortest round-tripping
binary64 representation. For example, numeric XML `2.1480000000000001` indexes
as `2.148`, the same binary64 value. The index retains the original XML digits in
`raw_numeric_value`; string cells are not normalized numerically. This is not
tolerance-based matching or business rounding. Different scalar values still
fail exact occurrence matching; OfficeCLI saved-cache expected-value QA remains required.

## Occurrence mapping contract

The header contains `schema_version: 2`, `corpus_id`, `canon_sha256`, and
`plan_sha256`. Every document has exactly the following fields:

```json
{
  "document_id": "doc-0001",
  "revision": 1,
  "artifact_sha256": "<actual final file SHA-256>",
  "draft_sha256": "<actual accepted draft JSON SHA-256>",
  "occurrences": [
    {
      "claim_id": "revenue",
      "section_index": 0,
      "locator": {
        "kind": "paragraph",
        "part": "word/document.xml",
        "index": 2
      },
      "start": 8,
      "end": 12,
      "expected_text": "12.0",
      "rendering": {"scale": 1000000, "decimals": 1}
    }
  ]
}
```

This illustrative occurrence assumes the final indexed paragraph begins
`Revenue 12.0` and the canonical claim is `12000000`; offsets must be taken
from the actual final index. There is no automatic business-text citation
insertion. Generation IDs belong in build metadata, not Office properties,
hidden cells, notes, or visible business prose.

Each occurrence requires `claim_id`, `section_index`, `locator`, `start`, `end`,
and `expected_text`. Optional fields are `rendering` and `formula`.
`expected_text` must equal both the located excerpt and grounded rendered value,
not merely contain a similar number. All uses must be covered; one location
cannot satisfy repeated sections using the same claim.
Paragraph/table-cell, shape/table-cell, and cell/range aliases cannot disguise
reuse of the same underlying text span. The index's internal text segments
support this check; copy locators, not internal segments, into mappings.

Rendering supports positive numeric `scale`, `decimals` from 0-12,
boolean `grouping`, and string `prefix`/`suffix`. Rounding is half-up.
Without rendering, normalized string representations must match exactly.
For numeric Excel caches, match the saved scalar rather than styling it as a
currency string. Review the visible display separately.

For derived cell/range claims, `formula` is required and must equal the saved
formula (a list in row order for ranges). All mapped derived cells need nonempty
saved formulas and caches. A missing cache is an error, not a request to invent
one. The helper's arithmetic derivation receipt and saved formula evidence are
distinct checks: formula dependency equivalence and business meaning still
need OfficeCLI saved-cache verification and semantic review.

Excel shared-formula followers retain structured saved evidence rather than
invented expanded formulas: `kind: "shared"`, `index`, `base_cell`,
`base_formula`, `range`, and `cell`. Copy that exact object into `formula`.
The reader verifies the shared base and range; relative-reference expansion
and calculation require a separately proven OfficeCLI capability. The default
runner rejects shared formulas rather than inventing expansion. Unsupported array/spill
followers without their own formula evidence cannot be mapped as derived
claims; use a supported formula representation or report an export blocker.

`provenance.json` adds record IDs and, for statement claims, source-version IDs,
speakers, statement kinds, and qualifiers. Derived occurrences list input claim
IDs. `audit` recomputes this receipt from the final files and mappings, and
rejects stale, incomplete, ambiguous, reused, or forged evidence.

## Optional storage bindings

`build\resource-bindings.json` has `schema_version: 2`, `corpus_id`, and a
`bindings` list. Each entry identifies document, revision, artifact hash,
and exactly one storage kind:

```json
{
  "document_id": "doc-0001",
  "revision": 1,
  "artifact_sha256": "<same final file hash>",
  "kind": "sharepoint",
  "drive_id": "<deployment drive ID>",
  "item_id": "<deployment item ID>"
}
```

A `local` entry instead has a safe relative `path`, with no drive/item IDs.
At most one binding per document/storage kind is allowed. This is an optional
reference layer, not an uploader or proof that a remote item exists. Deployment
must verify remote bytes and assign real ACLs separately.

## Coverage, recipient views, and limits

Coverage distinguishes the fixed plan from delivered artifact/claim/source/
viewpoint/condition coverage. It does not score a downstream agent or measure
all possible enterprise topics. Condition evidence remains structural plus
honest semantic condition review.

`access-view` requires completed strict additional QA, audits first and selects documents for a canonical principal
or `public` at `--as-of`. It omits full canon fingerprints, condition
annotations, and inaccessible source identifiers. The manifest does not copy
files or create secure isolation. Enforce actual ACLs in the delivery system
and keep canon/private build metadata out of the recipient's input.

These locators are reproducible within the exact hashed artifact, not promised
to survive editing. The reader is not full OOXML conformance validation,
semantic entailment, OCR, hidden-content analysis, or visual certification.
Unmodeled footnotes, Word floating text boxes, images/SmartArt, embedded
objects, and unsupported chart cache representations need a supported
text/cell/shape representation or an explicit export blocker. There is no
silent alternative engine or generated evidence fallback.
