---
name: company-content-generator
description: Generate realistic, coherent company content across multiple years in DOCX, PPTX, and XLSX, grounded in canonical structured data. Use for fictional company document corpora, press releases, product presentations, meeting transcripts, product specifications, financial reports, HR policies, and additional user-selected content types.
---

# Company content generator

Create a company's evolving document corpus, not independent plausible documents.
The company and its people are invented, but each deliverable must read as a
normal business document authored inside that company. Facts come from a
single versioned core info set: `canon.json`, including organization and glossary records.
Writers know only what was available at their publication date.

Read `references\contracts.md` and `references\content-types.md` before generating.
Read `references\dependencies.md` for external sources, installation, and preflight.
The Python helper at `scripts\company.py` uses only the standard library. Resolve
paths relative to this skill's directory; commands below assume the project root.
When installed elsewhere, use the resolved absolute script path.

## Dependencies and preflight

The workflow supports two backends: local Microsoft Office automation and the
pinned LibreOffice container. Both consume the same validated canon, plan,
date-filtered contexts, and grounded drafts. Read
`references\dependencies.md` for backend requirements and limitations.

At the start of every new run, before creating the output directory, canon,
plan, contexts, or drafts:

1. Determine the formats needed by the selected document types.
2. On Windows, run the native Office COM preflight for only those formats.
   It must establish usable application connections and exercise authoring,
   rendering, and Excel recalculation where selected; application registration
   or detection alone is insufficient.
3. Recommend **Native Office** when every requested format passes its native
   probe. Otherwise recommend the **LibreOffice container** and explain the
   missing or unavailable native capability. Verify Docker Engine and Compose
   v2 if the container is selected.
4. Ask the user to choose between the available backends, marking the
   recommendation. Do not auto-select or silently switch after a failure.
   Record the user's selection, recommendation, requested formats, and probe
   evidence in `build\preflight.json`; on resume, keep that run's backend.
5. If the chosen backend is unavailable, stop before generation and report
   the blocker. Do not fall back to the other backend. If neither backend is
   available, do not create a partial corpus.
6. Use only prepared local canon, plan, date-filtered contexts, and
   claim-checked drafts as exporter inputs.
   Do not retrieve real employee/company data or upload corpus content to external services.

After freezing the canon/plan and validating grounded drafts, export one
planned item per selected format as a pilot before bulk export. Verify exact
planned filenames and draft content. Inspect every Word/PowerPoint/LibreOffice
rendered preview for readability, clipping, tables, charts, and pagination.
For XLSX, verify formulas, cached values, expected results, and formula errors
using the selected backend's actual calculation engine. A draft without
formulas is not grounds for inventing calculations.

Record the selected backend and its actual application/runtime and package
versions, connection/probe evidence, pilot IDs and hashes, render evidence,
and limitations in `build\preflight.json`. Successful conversion is not itself
visual review. LibreOffice is not Microsoft Office and does not establish
Office fidelity; report its use accurately. Preserve accepted inputs and
outputs after a failure, report the completed count, and use a new output
revision for retries.

## Required intake

When invoked to generate content, first inspect the user's request and supplied
context. Reuse explicitly supplied values; examples and defaults in this skill
are not user requirements. Ask the user, using the host's structured question
tool where available, for every missing item below before generation:

- Company name and industry.
- Content volume: approximate document count and the historical period
  (start year and number of years, unless supplied as a date range).
- Target content types: offer press releases, product presentations, meeting
  transcripts, specs, quarterly/annual reports, HR policies, board decks, and
  financial workbooks. Allow multiple selections and additional free-text types.
  Preserve additional specified types rather than silently replacing them.
- Output location: ask where the corpus should be generated unless an explicit
  path was already supplied. Confirm the resolved location; never silently use
  the working directory, Downloads, or a hardcoded `corpus` folder.

Collect only the missing fields in a focused interaction. Do not repeatedly ask
for values already provided. If the user is unavailable, generation is blocked;
do not invent a company name, industry, volume, type selection, or output path.
Maintaining this skill is not a request to generate a corpus and needs no intake.

The seed may default to 42. Content depth defaults to **medium**; accept an explicit
**light**, **medium**, or **deep** setting from the user's request and record it as
`content_depth`. Mention the default during intake; do not require an extra
question when the user has not specified a depth. See Content depth below.
An existing canon is supported after copying it into the chosen output layout
and validating its embedded organization, glossary, and business records.
Map familiar natural-language types to built-in IDs. For additional types, define their
formats, canonical anchor kinds, audience, and required genre sections using the
`additional_types` request contract; add the needed canon records before planning.
Do not require edits to the installed skill just to support a custom type.

Treat the approximate count as an exact planning target, excluding supporting
data and build artifacts. Use modern `.pptx`, `.docx`, and `.xlsx`. Select formats
appropriate to the requested types; a press-release-only corpus may be DOCX-only.
Use all three formats when the requested type mix supports them, unless the user
specifies a narrower format selection. Never add unrequested types to force a
format mix. Explain coverage warnings rather than increasing the target.
Do not access real employee/customer/internal-company data to fill gaps.

## Output layout

The chosen output root must contain exactly:

```text
<output>\
  canon.json
  documents\
    <realistically named final Office documents>
  build\
    request.json
    plan.json
    manifest.json
    qa.json
    preflight.json
    depth-targets.json
    drafts\
    contexts\
    scripts\
    renders\
```

`canon.json` is the entire **core info set** and the only root-level file. It
contains structured business history, referenced people and departments,
effective-dated role/reporting history, and glossary definitions. Do not generate
redundant standalone org-chart or glossary files. Derive bounded, date-filtered
`org_chart` and `glossary` views in writer contexts directly from canon records.
Final documents belong exclusively
in `documents`. Everything else generated or used for the run goes under
`build`: scripts, dependencies/lockfiles, metadata, drafts, prompts, logs,
contexts, temporary files, extracted text, charts/images, and QA renders.
An installed shared tool may live outside the corpus; run-specific artifacts
from that tool must still be directed into `build`.

Prepare bootstrap input in the session workspace if needed; `init` copies the
request into `build`. Use explicit paths or run renderers from `build` so their
default output does not pollute the root or `documents`. Do not delete/rearrange
an existing corpus without approval; create a new directory/revision instead.
Older three-file layouts and plans without depth settings are legacy runs:
preserve their files and receipts. This skill change does not migrate, delete
files from, or re-certify completed corpora.

## Bounded organization grounding

The org chart is a reference graph, not a complete employee directory. Include
only named people/roles directly used by canonical roles, ownership, policies,
meetings, quotes, contacts, or other canon relationships that support documents.
Include referenced departments and managers needed to interpret those roles.
Never create anonymous staff, one node per employee, or additional names merely
to match financial headcount. The named-person count need not equal headcount.

Preserve stable person/department IDs and full effective-dated role history,
including title changes and reporting changes. A role assignment itself is a
canonical reference; unused person/department catalog entries without references
are not part of the org chart or writer context. A referenced contact may have
no internal role; do not invent one just to complete a hierarchy.

All named people, speakers, owners, approvers, departments, and role titles must
be grounded in the canon's organization records alongside business and glossary records.
If a new person is needed, add their canonical identity and relevant role or
relationship references, and revalidate the canon before
freezing or create a new revision. Do not append people to a frozen chart or
introduce them only in prose. Writers use the context's date-filtered `org_chart`
and its `active_roles`, never the unfiltered full-history root file.

## Content depth

Use request field `content_depth: "light" | "medium" | "deep"` (default **medium**).
This controls substantive coverage, not just length, font size, or page count.
It does not alter the requested document count, years, types, or format mix.
The planner copies the setting into every document; contexts carry it in
`document.content_depth`, drafts repeat it, and QA verifies that same setting.
Changing depth after freeze requires a new canon/plan revision and fresh QA.

| Depth | Content expectation |
|---|---|
| **Light** | Concise but complete genre: essential facts, context, rationale, implications, and actionable next steps. Never a few generic sentences or a bare financial table. |
| **Medium** | Substantive business document: multiple grounded supporting details, operating drivers, comparisons where available, tradeoffs, risks, and practical follow-through. This is the default. |
| **Deep** | Detailed working/reference document: richer source granularity, reconciled analyses, alternatives, dependencies, detailed examples, implementation controls, and appendices where useful. |

Follow the genre-specific coverage and measurable targets in
`references\content-types.md`. Before freezing, expand the canon with enough
public-safe details, approved quotes, requirements, meeting positions, financial
breakdowns, historical comparisons, and assumptions to support the selected depth.
The seeded baseline is not sufficient merely because it validates. More depth
requires richer structured source facts, not unsupported elaboration.

Do not pad with repeated paragraphs, generic risks, empty sheets, duplicated
charts, one-sentence slides, or arbitrary extra names. Deep does not license
inventing actuals, forecasts, certifications, customer results, or legal claims.
When requested analysis requires unavailable facts, enrich the canon before
freeze or report a blocker; never silently downgrade depth.

Set concrete per-document genre targets in `build\depth-targets.json` before drafting,
using the playbook's ranges and useful analysis modules. Justify any adaptation
to industry/audience/format there rather than silently weakening the targets.
Pilots must meet those targets before bulk export. Review both substantive
coverage and measured delivered content in a separate `depth_review`, not
only section presence or length. Keep depth labels and targets in build metadata,
not in the company's document prose.

## Authentic business content

Follow `references\content-types.md` for each genre's structure, voice, detail,
and acceptance criteria. Write complete business content, not descriptions of
what a document would contain, outlines, canonical record dumps, generic
summaries, or the same prose wrapped in three Office formats. A press release
must read as a publishable press release; a product presentation must present
customer problems, value, capabilities, use cases, and next steps.

Never label final content as synthetic, fake, fictional, simulated, sample/test
material, or AI-generated. Do not discuss the generation process or add such
labels in titles, body text, speaker dialogue, footers, speaker notes, comments,
watermarks, charts, workbook tabs/cells, or Office properties. Write in the
company's ordinary voice, not an outside narrator's account of an invented
company. Existing canon flags and build provenance are private inputs, not copy
to export. Ordinary technical terms such as "synthetic aperture" are allowed
when grounded; they must not be used as a self-description of the corpus.

Quotes and dialogue use named canonical speakers and effective roles without
generation disclaimers. Facts, obligations, decisions, contacts, quotations,
requirements, benefits, and metrics needed by a genre must exist in the canon.
Extend the canon before drafting; do not compensate for missing facts with
placeholders, lorem ipsum, ungrounded claims, or explanation of the missing data.
Complete a separate genre review, not merely a citation or ZIP-structure check.

## Realistic document naming

Every primary Office file must have a realistic business title and a descriptive
filename, not `doc-0001.docx`, `report.xlsx`, or a content-type label plus raw
record IDs. The planner provides both `title` and `filename` from canonical
topics; review them for natural wording before freezing the plan.

- Name the actual subject: reporting period, named product/version, acquisition
  target and stage, policy/version, or meeting agenda. Do not invent a launch,
  acquisition close, decision, or business claim just to make a catchy title.
- Distinguish announcement/planning from completion/launch in titles as well as
  body text. Financial titles use the reporting period, not publication year.
- Use the planned title for the document heading and Office title property.
  Use the exact planned `filename` under `documents`; manifests must point to
  that file. Keep opaque document and record IDs in private build metadata or
  sidecars, not in the human-readable title or filename.
- Filenames start with the publication date and use descriptive words. Include
  real policy/product versions where relevant; use business-like follow-up,
  supplement, revision, or edition suffixes when necessary for uniqueness.
  Avoid vague `final-final`, `output`, `untitled`, or arbitrary numeric filenames.
- Preserve safe, case-insensitively unique filenames. The helper removes unsafe
  characters and limits filenames to 140 characters while preserving the
  beginning and subject-bearing end. Do not rename a frozen output without
  creating a new plan revision and regenerating affected drafts/QA receipts.

Examples: `Lumenvale Systems - Q1 2022 Operating Results`,
`Lumenvale Systems - Beacon 1.0 Product Specification`, and
`Lumenvale Systems Completes Acquisition of Sableworks Analytics`.
A corresponding filename is
`2022-04-15 - Lumenvale Systems - Q1 2022 Operating Results.xlsx`.
Draft JSON may still use stable ID-based filenames for machine bookkeeping.

## Workflow

Commands below use `$outputRoot`, the resolved path explicitly supplied by the
user, and `request.json`, the bootstrap input in the session workspace. Substitute
the resolved absolute script path when the skill is installed outside this
project. Do not run examples with an unset path or assume the example is intake.

1. **Create the core info set before prose.** Write `request.json` using the documented
   contract, then initialize a new, dedicated output directory:

   ```powershell
   python .github\skills\company-content-generator\scripts\company.py init --config request.json --out "$outputRoot"
   ```

   This produces a seeded baseline, not an industry-complete simulation. Adapt
   root `canon.json`, including its organization and glossary records, to the brief before planning.
   Expand structured records for any
   additional facts required by the requested content. Preserve financial
   reconciliation, stable IDs, references, effective dates, and availability
   dates. Do not invent extra facts while drafting.

2. **Validate and freeze.** Run `validate`, then `plan`. Inspect all plan warnings.
   Dates cover the requested fiscal years plus a 90-day publication tail for
   final-quarter and annual reporting. Repeated topics represent distinct
   follow-ups, not duplicated text. Review the descriptive titles and filenames
   against their anchors before freezing. A small target may not cover every year.

   ```powershell
   python .github\skills\company-content-generator\scripts\company.py validate --canon "$outputRoot\canon.json"
   python .github\skills\company-content-generator\scripts\company.py plan --canon "$outputRoot\canon.json" --out "$outputRoot\build\plan.json"
   ```

   Once any prose is generated, never change the canon or plan in place.
   A correction requires a new revision and regeneration of affected documents.
   The plan contains the canon's SHA-256 fingerprint, covering the source records
   of the single core artifact, including organization, glossary, and depth setting.

3. **Generate in publication order, in bounded batches.** Start with one pilot
   document per format. For each plan item, retrieve its date-filtered context:

   ```powershell
   python .github\skills\company-content-generator\scripts\company.py context --canon "$outputRoot\canon.json" --plan "$outputRoot\build\plan.json" --doc doc-0001 --out "$outputRoot\build\contexts\doc-0001.json"
   ```

   Use this context as the only factual input. Never give a writer the unfiltered
   canon: it includes future actuals, launches, promotions, and acquisitions.
   The context includes date-filtered `org_chart` and `glossary` views.
   `active_roles` identifies the roles effective on that date. Distinguish
   announced plans from completed events, forecasts from actuals, and period end
   from publication date. If a document needs a new substantive fact, stop,
   extend the canon, and replan before writing.

4. **Write grounded drafts.** Save one JSON draft per document under `build\drafts`,
   using the contract in `references\contracts.md`. Each factual statement has
   an exact record/path/value claim; each claim ID occurs in the section's
   `claim_ids`. Specific numbers, people, dates, policy obligations, named
   initiatives, decisions, owners, and deadlines all need canonical provenance.
   Harmless connective prose can be uncited but cannot introduce new facts.
   Copy the plan's `content_depth` into the draft and meet its genre-specific
   depth targets with source-grounded detail.
   Include a semantic `role` on each section to cover the plan's required genre
   sections. Do not expose claim receipts in the business prose. Follow the
   content type's voice, audience, and confidentiality.

   A source reference does not prove a paraphrase is correct. Review every
   section against its claims, verifying units, rounding, tense, negation,
   time-of-knowledge, and author authority. Do not reinterpret a revenue value
   as profit or an announced acquisition as closed.

5. **Export through the selected backend.** Native Office runs use the
   repository's Windows PowerShell COM exporter for the requested formats.
   Container runs use `python3 -m office_backend` with the prepared `--canon`,
   `--plan`, `--contexts`, `--drafts`, and `--root` paths as documented in
   `prototypes\libreoffice-backend\README.md`. If the selected backend, contract,
   authoring operation, render, or calculation is unavailable, stop and report
   the blocker; do not switch engines or generate substitute drafts.
   Do not deliver renamed text files or just draft JSON.
   See `references\content-types.md` for layouts and format expectations.
   Put deliverables under `documents` using the exact planned filenames. Use
   normal company branding, a business title, appropriate publication/effective
   dates, and business confidentiality marks only where applicable. Opaque IDs,
   generation provenance, and claim receipts stay in `build`, not visible prose
   or Office properties.

   For financial XLSX, link derived totals and margins to source cells using
   grounded formula structures and explicitly state currency and units. Use
   Excel recalculation for Native Office or LibreOffice Calc for the container,
   preserve formula caches, and check expected values and formula errors. Do
   not fabricate workbook calculations when the draft lacks grounded inputs.
   For presentations, use charts, timelines, and meaningful layouts; render
   each slide through PowerPoint or LibreOffice Impress and inspect every
   preview. Render DOCX through Word or LibreOffice Writer and verify
   pagination, tables, headers, and footer labeling.

6. **Audit and resume safely.** Validate drafts before rendering. Write
   `build\manifest.json` only for completed documents with paths and SHA-256 hashes.
   Audit checks the exact target, mixed formats, source claims, planned
   descriptive filenames, safe file paths, file hashes, and basic Office ZIP/XML
   integrity:

   ```powershell
   python .github\skills\company-content-generator\scripts\company.py check --canon "$outputRoot\canon.json" --plan "$outputRoot\build\plan.json" --drafts "$outputRoot\build\drafts"
   python .github\skills\company-content-generator\scripts\company.py audit --canon "$outputRoot\canon.json" --plan "$outputRoot\build\plan.json" --drafts "$outputRoot\build\drafts" --manifest "$outputRoot\build\manifest.json" --root "$outputRoot"
   ```

   This is **structural and claim-receipt validation**, not semantic, visual, or
   recalculation certification. Complete those reviews separately and write
   results to `build\qa.json` using the documented contract. Audit requires passed
   semantic, genre, depth, and visual reviews for every document, plus
   recalculation evidence from the selected backend for XLSX. Structural genre-section and
   self-description scans do not independently prove authentic writing or
   exhaustive absence of disclaimers.
   Check receipts are honest evidence, never prefilled success flags.

   On resume, recheck the selected backend and its operations, verify
   canon/plan/context/draft hashes and backend evidence, and
   skip only verified documents. A selected-document pilot or partial corpus
   must be reported as partial. Retry in a new output revision without changing
   IDs, dates, facts, accepted drafts, or unrelated files.

## Consistency rules

- The canon is the source of truth. Do not rely on memory from earlier drafts.
- IDs remain stable despite renaming, reorganizations, and product revisions.
- `known_on` controls visibility; `effective_on` controls state. An announced
  future milestone is visible but must not be described as already effective.
- Org relationships and titles come from the canon's effective role snapshot. Do not
  introduce meetings attended by nonexistent people or use their future titles.
- Quarterly actuals are unavailable before their `known_on`; annual totals must
  equal quarterly totals. Currency is USD in the baseline. Headcount is a
  point-in-time measure, not a value to sum across quarters.
- M&A announcements and closes are separate records. Do not consolidate a
  target before the close or assume an announcement guarantees completion.
- Product specifications cite version-specific requirements and acceptance
  criteria. Add these records before writing specs; the baseline milestone is
  not a complete specification.
- HR policy revisions are effective-dated. Never portray fictional policies as
  legal advice or claim real regulatory compliance.
- Terms use the glossary's canonical definitions. Record abbreviations and
  synonyms centrally; do not give the same acronym multiple meanings.
- Fictional risk, disagreement, and changed forecasts are allowed only as
  explicitly recorded facts, not uncontrolled contradictions.
- The seeded helper is reproducible; model-generated prose is not guaranteed
  byte-for-byte reproducible. Preserve accepted drafts and artifact hashes.

## Completion

Deliver root `canon.json`, final documents, and all build inputs/receipts in the
required layout. Verify each document reads as its actual target genre, has a
realistic name, and contains no generation self-labeling.
Report the selected content depth, actual counts by year, type, and format, coverage warnings, and any
remaining blockers. Never claim the corpus is complete merely because a plan
or draft set exists.
