---
name: company-content-generator
description: Generate realistic, coherent company content across multiple years in DOCX, PPTX, and XLSX, grounded in canonical facts and attributed statements. Supports viewpoints, source histories, portable provenance, and optional corpus imperfections for general-purpose enterprise content.
---

# Company content generator

Create a company's evolving document corpus, not independent plausible documents.
The company and its people are invented, but each deliverable must read as a
normal business document authored inside that company. Facts come from a
single versioned core info set: `canon.json`, including organization and glossary records.
The canon distinguishes business facts from assertions by fictional sources:
faithfully reporting an opinion or an erroneous historical assertion does not
make it an objective fact. This is a general-purpose content generator, not a
decision or agent-evaluation engine.
Writers know only what was available at their publication date.

Read `references\contracts.md` and `references\content-types.md` before generating.
Read `references\dependencies.md` for external sources, installation, and preflight.
Read `references\provenance.md` before preparing exporter bindings and optional QA.
Read `references\use-case-layering.md` only when discussing an optional
customer/use-case layer; do not implement or invoke its illustrative evaluator.
The Python helper at `scripts\company.py` uses only the standard library. Resolve
paths relative to this skill's directory; commands below assume the project root.
When installed elsewhere, use the resolved absolute script path.

## Dependencies and preflight

**[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) is the sole document
authoring/HTML-export/formula engine. Microsoft Office is not required.**
Python 3.10+ provides business, claim, package/hash, provenance, and coverage
checks, not document authoring. Screenshot pixels never come from OfficeCLI
itself: a small pinned **Playwright-for-Python + matching managed Chromium**
adapter (`scripts\playwright_capture.py`) turns OfficeCLI's own static HTML
export into bound, checked images. OfficeCLI never does screenshotting; the
browser adapter never authors, edits, or exports Office documents.

Before creating the output directory, canon, plan, or drafts, follow
`references\dependencies.md`: detect OfficeCLI/skill versions, check official
releases once, perform authorized installation/update when missing or required,
verify versioned release checksums and installed version, then run tiny actual
save/reopen/check/HTML-export/Playwright-capture probes for selected formats.
Configure only authorized skill/MCP targets. Preserve usable prior binaries;
do not execute downloaded installers through Invoke-Expression or silently
update mid-run. Pin OfficeCLI executable/version/hash **and** Playwright
version/managed-Chromium executable/version/hash in `build\preflight.json`
(`capture_tool` identity); revalidate changed capabilities on resume.
Maintaining this skill and running helper tests do not require Microsoft
Office or generation intake.

Force `view <file> html --out <path>` for HTML export and the pinned Playwright
adapter for pixels. Never use OfficeCLI's own `view screenshot --render html`
image path, native/auto rendering, `refresh`, native page counts, COM,
LibreOffice/OpenOffice, cloud Office, or standalone OOXML authoring libraries
as a substitute. Suppress per-command updates and use checked atomic standalone
batches. Inspect runtime help for unfamiliar operations instead of guessing.
OfficeCLI layout is not native Office fidelity certification; Playwright
capture proof is Windows-pinned-version-only and makes no cross-platform or
pixel-identity guarantee. DOCX output from this skill uses explicit, deliberate
page breaks (never automatic natural-flow pagination): every authored page's
content must fit within its physical rendered page, each capture rejects a
rendered page/box-overflow or page-count mismatch against the explicit
authored break plan, and a one-page document needs no superfluous break.
Page-dependent fields, unproven formula functions, and unsupported
drawing/chart objects (anything outside the proven column/line
`twoCellAnchor` chart capture path) stop the bounded workflow explicitly;
no fallback was used.

Required OfficeCLI/Playwright-capture-adapter capability loss blocks
generation/resume; preserve existing outputs and report the actual completed
count. Do not retrieve real workplace data, upload run content, modify
personal skills, or install tools without authorization.

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
- Additional QA: explicitly ask during intake, "After generation and its mandatory
  checks/pilot reviews, should additional corpus QA run?" Offer **No additional QA**
  (selected by default), **Sampled**, **Balanced**, and **Full**.
  Reuse an explicitly supplied choice. Do not silently choose for an unavailable
  user. Record the answer in execution metadata, not the business canon.

Collect only the missing fields in a focused interaction. Do not repeatedly ask
for values already provided. If the user is unavailable, generation is blocked;
do not invent a company name, industry, volume, type selection, output path, or QA answer.
Maintaining this skill is not a request to generate a corpus and needs no intake.

The seed may default to 42. Content depth defaults to **medium**; accept an explicit
**light**, **medium**, or **deep** setting from the user's request and record it as
`content_depth`. Mention the default during intake; do not require an extra
question when the user has not specified a depth. See Content depth below.
Additional corpus QA is **opt-in**, not a prerequisite for generation completion.
When requested, accept **Sampled** (selected
complete content and visual reviews), **Balanced** (all content,
selected visual reviews), or **Full** (all content and visual reviews).
Review scope is independent of light/medium/deep content depth. Generation always
checks all files, claims, date/access, required sections, measurable depth,
structure/issues, hashes and saved formulas, and fully reviews family pilots.
Detailed final-artifact provenance, coverage and additional corpus reviews run
only in the separate QA pass. If selected at intake, invoke that pass automatically
after verified generation; otherwise stop without running it. A later explicit QA
request uses accepted files without reauthoring or repeating generation intake.
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
    generation.json
    qa.json                (only after additional QA reviews)
    qa-state.json          (additional QA)
    preflight.json
    execution-plan.json
    review-plan.json        (additional QA)
    review-work.json        (additional QA; pending gates are not passes)
    runner-state.json
    machine\
    qa-machine\            (additional QA; generation receipts stay unchanged)
    document-state\
    export-bindings\
    timings.json           (agent-recorded phase spans)
    tool-results\
    reviews\
    pilot-reviews\         (preserved generation reviews when additional QA starts)
    depth-targets.json
    artifact-index.json    (additional QA)
    claim-locations.json   (additional QA)
    provenance.json        (additional QA)
    coverage.json          (additional QA)
    resource-bindings.json  (optional)
    views\                 (requested simulated delivery manifests)
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
New runs and helper validation require **schema version 2**. Do not validate or
silently upgrade version-1 corpora. The repository's semiconductor example uses
schema version 2; its generation receipt does not certify additional corpus QA.

## General-purpose source and viewpoint modeling

Keep canonical actuals and business invariants intact. Add `source_version`
and `statement` records for attributed content, and `viewpoint` records for
requested author/audience/focus combinations. Statements distinguish fact,
opinion, assumption, hypothesis, recommendation, risk, and decision; qualifiers
express uncertainty, dissent, and unresolved matters. Backward `corrects`,
`retracts`, and `supersedes` links preserve source history without rewriting it.
Approval/status and authority are scoped source properties, not universal
truth or trust scores.

Add only the named people, sources, and structured details needed by the brief.
Writers use the planned viewpoint and canon-derived role, not stereotypes or a
requirement that every perspective recommend a decision. Required source
statements must be represented as attributed value claims.

Author knowledge, author access, and intended reader access are different.
Use effective-dated fictional groups, knowledge grants, and explicit releases
when broader publication is intended. `context` and `check` share the same
projection rules; dependencies and derived claims cannot bypass those rules.
The helper filters structured inputs; it does not isolate the agent's entire
memory or enforce real file/service permissions.

## Optional corpus conditions

Default to clean, consistent content. Enable gaps, conflicts, noise, or stale
sources only when requested, using `corpus_conditions` and canonical `condition`
instances from the contracts. Do not ask users to configure these for ordinary
generation. Resolve explicitly requested viewpoints and controls during intake;
explain any incompatible budget, genre, date, or visibility requirements.

Distinguish information unknown to an author from known information intentionally
undocumented, and from a source deliberately outside the collection. Unknown
records are withheld from writer context; undocumented records remain knowable
but are suppressed from final claims. Follow the context's suppression
instructions without exporting their IDs or generator rationale.

Intentional conflicts are linked source assertions, never inconsistent actuals.
Staleness is relative to a declared reference date/newer source. Noise consists
of complete, grounded adjacent material, not padding or corrupt files.
Condition instances can require multiple documents; document count stays fixed.
Enrich missing canonical instances before planning. Missing exports, accidental
contradictions, and missing genre prerequisites remain errors.

Generator-only condition labels stay private in `build`/the canon. Do not tell
business authors that a gap/conflict is a test, prescribe expected agent
behavior, or invent a conclusion. A real author may acknowledge uncertainty or
missing information only when that acknowledgement is itself grounded.

## Bounded organization grounding

The org chart is a reference graph, not a complete employee directory. Include
only named people/roles directly used by canonical roles, ownership, policies,
meetings, quotes, contacts, or other canon relationships that support documents.
Include referenced departments and managers needed to interpret those roles.
Never create anonymous staff, one node per employee, or additional names merely
to match financial headcount. The named-person count need not equal headcount.
Canonical viewpoint authors are legitimate referenced identities; they need
not be internal employees. Access-group membership alone does not expand the
bounded chart or create employment roles.

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
Depth never changes condition counts, access rules, provenance completeness, or
whether imperfections appear. A deep corpus may be clean; a light corpus may
have an explicitly requested disagreement.

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
Extend the canon before drafting; do not compensate for missing prerequisites
with placeholders, lorem ipsum, or ungrounded claims. A deliberately configured
unknown or undocumented fact must not be silently filled to satisfy a depth
target. Explain an unknown in prose only when the author has a grounded basis
to recognize it; otherwise leave it unstated.
Complete the required genre reviews under the selected QA policy, not merely
a citation or ZIP-structure check.

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

3. **Freeze execution scope; draft in bounded batches.** For new runs, set
   `build\execution-plan.json` using `generation_execution_plan` (execution version 2)
   and the contract reference. Supply the explicit `post_generation_qa` answer.
   Bind layout revisions/hashes, family pilots, flagged exceptions,
   and a runtime budget only if supplied. Estimate content/visual review counts
   and confirm pilot/selected-QA image-delivery capacity before bulk work. Deterministic sampling
   chooses the first and largest normalized draft per genre/format/layout family,
   plus pilots, condition/source-history carriers and affected exceptions.
   Use genuinely shared, content-independent layout files: separate document
   prose/data from layout identity. A family is `(genre, format, layout revision,
   layout hash)`. Do not generate a new layout per document unless required.
   Use `family_summary` to report family sizes, pilots, selection reasons and
   singleton/all-selected warnings; never fake layout reuse to shrink sampling.
   Prepare the bounded date-filtered contexts together:

   ```powershell
   python .github\skills\company-content-generator\scripts\company.py contexts --canon "$outputRoot\canon.json" --plan "$outputRoot\build\plan.json" --out "$outputRoot\build\contexts"
   ```

   Each writer uses only its document's context as factual input. Never give a writer the unfiltered
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

   A source reference does not prove a paraphrase is correct. For required
   content reviews, inspect every section against its claims, verifying units, rounding, tense, negation,
   time-of-knowledge, and author authority. Do not reinterpret a revenue value
   as profit or an announced acquisition as closed.

5. **Execute the shared OfficeCLI runner.** Follow installed OfficeCLI guidance
   and runtime schema help. Provide checked batch or template/merge jobs,
   frozen layout files, exact exporter claim associations, grounded expected
   formula bindings and screenshot targets as described in `contracts.md`.
   The runner is not a prose generator or genre-authoring framework.
   Record `export_bindings` while building jobs. Bind each section/claim use to
   a checked batch command, exact template target or indexed locator; the helper
   derives final normalized spans during QA. Do not reconstruct offsets through
   interactive searches or globally match identical tokens after export.
   Complete one actual planned pilot per layout family and review it fully
   before bulk use. Freeze accepted layouts, not repeated prose. Stress grounded
   long titles/text, historical comparisons and numeric/formula/layout edge cases.
   A missing required capability blocks; separate backend development from
   generation rather than improvising an unbounded replacement pipeline.
   See `references\content-types.md` for layouts and format expectations.
   Put deliverables under `documents` using the exact planned filenames. Use
   normal company branding, a business title, appropriate publication/effective
   dates, and business confidentiality marks only where applicable. Opaque IDs,
   generation provenance, and claim receipts stay in `build`, not visible prose
   or Office properties.

   For XLSX, write inputs before dependent formulas, state units/currency, save,
   reopen and verify every cache against independently grounded derived claims.
   At most one corrective known-order OfficeCLI formula rewrite is allowed.
   Missing/stale/error/unsupported caches block; no raw cache injection or
   deferred calculation on open. During generation capture complete pilots and
   required corrective checks with OfficeCLI HTML; defer other policy-selected
   captures to QA. Always declare/validate target coverage and supported objects.
   Check all schema/issues and explicit chart caches. Hash/index final saved
   bytes; later edits invalidate affected renders, mappings and reviews.

   ```powershell
   python .github\skills\company-content-generator\scripts\officecli_runner.py run --root "$outputRoot" --jobs build\jobs-pilots.json
   # Inspect every pilot and persist actual reviews before the bulk command.
   python .github\skills\company-content-generator\scripts\officecli_runner.py run --root "$outputRoot" --jobs build\jobs.json
   python .github\skills\company-content-generator\scripts\officecli_runner.py assemble --root "$outputRoot" --jobs build\jobs.json
   ```

   Persist actual pilot reviews under `build\reviews`; the runner never writes
   editorial/visual passed receipts. Use one grouped review per document.
   A failed sample blocks the family and requires explicit flagged-family
   expansion. Undelivered/unreadable images cannot be accepted. Contact sheets
   are useful only when legible; inspect detailed crops as necessary.
   One concrete corrective retry per logical operation is persisted across
   resume; a second failure stops with evidence. Record phase active/wait time.
   Honor user-supplied budgets at safe checkpoints, without inventing a budget
   or promising a fixed completion time.
   Bulk execution automatically uses up to four independent document workers,
   bounded by CPUs/eligible jobs; `--workers 1` is the explicit serial override.
   One owning capture lane reuses one browser. Never run competing runner
   invocations against the same corpus. A known failure pauses its family;
   retain findings, correct/review explicitly, and expand affected QA scope.

6. **Complete generation independently.** `assemble` verifies the entire planned
   set, generation machine evidence and actual pilot/corrective reviews, then
   writes `build\manifest.json` and `build\generation.json`. It does not require
   full provenance/coverage, corpus renders, or `qa.json`. For a standalone recheck:

   ```powershell
   python .github\skills\company-content-generator\scripts\company.py audit-generation --canon "$outputRoot\canon.json" --plan "$outputRoot\build\plan.json" --drafts "$outputRoot\build\drafts" --manifest "$outputRoot\build\manifest.json" --root "$outputRoot"
   ```

   If intake declined additional QA, deliver now: **generation complete;
   additional corpus QA not requested**. Do not certify unreviewed content,
   layout, exact final-artifact provenance or complete coverage.

7. **Run additional QA only when selected.** Automatically invoke the separate
   pass after successful generation if intake selected an intensity. A later
   explicit request may use `--policy Sampled`, `--policy Balanced`
   or `--policy Full` without changing canon/drafts/Office files:

   ```powershell
   python .github\skills\company-content-generator\scripts\officecli_runner.py qa --root "$outputRoot" --jobs build\jobs.json
   # Perform required reviews in review-work.json; persist actual per-document reviews.
   python .github\skills\company-content-generator\scripts\officecli_runner.py finalize-qa --root "$outputRoot"
   ```

   `qa` verifies accepted bytes, resolves private exporter bindings, writes exact
   mappings/provenance/coverage, determines review scope, and creates only missing
   selected captures. It writes separate `qa-machine` evidence and a bound
   `review-plan.json`; `review-work.json` preserves verified existing reviews
   and marks missing gates pending. Complete actual grouped reviews under
   `build\reviews`; `finalize-qa` collects them and performs the independent strict
   audit before marking QA complete. Do not rerun separate indexing, provenance,
   coverage or audit commands redundantly when this orchestrated pass did them.
   Required content reviews need
   semantic, genre and `depth_review`; selected visuals need every required
   delivered image. Unselected gates are `not_selected`, never passed.
   Documents carrying configured conditions also require `condition_review`.
   The final audit recomputes selection, provenance and coverage from final bytes
   and drafts in a fresh validation context; stale/forged receipts fail. It checks
   evidence, not independent semantic or native-rendering certification.
   Structural genre-section and self-description scans do not independently
   prove authentic writing or exhaustive absence of disclaimers.
   Check receipts are honest evidence, never prefilled success flags.
   QA failures preserve generated files and leave QA blocked/pending. Editing
   accepted bytes requires an explicit affected revision, never a silent repair.

8. **Measure and resume safely.** Use the runner's `phase --action start --name ...`
   and `phase --action end --id <returned-id>` around actual agent-owned
   canon/planning, drafting, editorial and visual work (`--pass qa` for QA).
   Use `--parent` for nested spans and `--document` for document-specific work.
   End parent spans only after children. Omitted work is unknown, not zero.
   `summary --root "$outputRoot"` reports wall time separately from overlapping
   worker/subprocess/capture activity, queue spans, reuse counters, gaps and
   generation/QA milestones. Never infer where hours went without measurements.

   On resume, recheck pinned OfficeCLI identity and required capabilities,
   verify canon/plan/draft/job/layout/final hashes and skip
   only verified documents. A partial corpus must be reported as partial.
   Retry failed documents without changing IDs, dates, facts, or unrelated files.
   Existing execution-version-1 runs and legacy schema-v2 QA without execution
   sidecars retain their strict contracts. No automatic downgrade, migration,
   existing corpus regeneration, or re-certification is part of this workflow.

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
  explicitly recorded statements with correct attribution and status.
  Deliberate factual conflicts require requested condition instances; accidental
  contradictions and inconsistent canonical actuals remain errors.
- The seeded helper is reproducible; model-generated prose is not guaranteed
  byte-for-byte reproducible. Preserve accepted drafts and artifact hashes.

## Completion

Deliver root `canon.json`, final documents, and the actual build inputs/receipts.
Report generation completion separately from additional QA: not requested,
pending/awaiting reviews, blocked, or complete. Mandatory generation checks and
accepted pilot/corrective reviews are required even with no additional QA.
Report content depth, selected QA intensity if any, actual reviewed/total content
and visual counts (including pilot work), counts by year/type/format, warnings,
timing evidence, OfficeCLI limitations and blockers. Do not claim unselected
documents were reviewed, infer full coverage from the plan, or claim generation
complete from drafts alone. A selected QA pass is not complete until strict audit
passes. Verified final-artifact provenance and coverage are delivered under
`build` only when that pass produces them.
If a recipient view is
requested, generate it with `access-view --principal <canonical-person-id-or-public>
--as-of <YYYY-MM-DD>` and the normal canon/plan/draft/manifest/root arguments.
Completed strict QA/audit is a prerequisite; generation-only output cannot
produce a certified recipient view. The output must be an explicit path under `build\views`. These manifests are
simulated selections, not ACLs. Do not distribute the full canon or private
build tree as an agent-visible corpus.

No customer decision question, independent answer key, evaluation partition,
expected disposition, or decision metric is required to use this skill.
