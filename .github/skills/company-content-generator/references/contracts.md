# Data contracts

All JSON is UTF-8. Dates are ISO `YYYY-MM-DD`, IDs are stable strings, financial
amounts are integer USD, and version 2 uses calendar fiscal years. Python 3.10+
is required. Invalid input returns a nonzero exit code with an explicit error.
New canon, plan, context, draft, manifest, QA, and derived receipts use
`schema_version: 2`. Version-1 validation and automatic migration are not
supported. Preserve historical files; create a new run for this contract.

`corpus_id` is a logical namespace, not an output path. The helper uses a supplied
alphanumeric/underscore/hyphen namespace or derives one deterministically from
the normalized initialization request. Read the resolved value from the canon.
Canon/plan fingerprints and document revisions distinguish later content.
Temporal resolution is ISO calendar dates; this contract does not promise
sub-day knowledge/publication isolation.

## Request

```json
{
  "company_name": "Lumenvale Systems",
  "industry": "industrial analytics software",
  "start_year": 2022,
  "years": 3,
  "documents": 45,
  "seed": 42,
  "content_depth": "medium",
  "document_types": [
    "meeting_transcript", "press_release", "product_presentation", "product_spec",
    "quarterly_report", "annual_report", "hr_policy",
    "board_deck", "financial_model"
  ]
}
```

For the skill's interactive workflow, company name, industry, volume, target
types, and output path must be supplied by the user/context or collected before
generation. The example is not a fallback for missing user requirements.
The low-level helper's `document_types` default is all built-in types; the skill
must write the user's actual selection explicitly. A selected type can support
several formats. Optional `formats` restricts output to supported Office formats;
otherwise use the formats supported by the selected types. Every selected type
must support at least one selected format and the target must cover those formats.
A one-document, DOCX-only press-release run is valid.
The seed is an integer; years and counts must
be positive integers. The period must leave room for the publication tail in
Python's supported date range. Very large targets cost model time and disk space:
the helper never initiates AI calls or spawns workers.

`content_depth` is optional and accepts exactly `light`, `medium`, or `deep`.
Omission resolves to `medium`; `init` writes that resolved value into the
canonical request and `build\request.json` without mutating the input config.
Depth applies to every requested document, not the document count or format mix.
Depth does not enable imperfections, change access rules, or weaken provenance.
`references\content-types.md` defines substantive coverage and measurable
planning targets for each genre. Record run-specific targets and justified
adaptations in `build\depth-targets.json` before drafting. The helper validates
the setting and its propagation, not word counts, sheet counts, or analytical
quality; those require an honest depth review of delivered content.

Optional request fields:

| Field | Meaning |
|---|---|
| `corpus_id` | Stable logical namespace, otherwise deterministically derived |
| `viewpoints` | Distinct canonical viewpoint IDs, each requiring at least one compatible planned document |
| `corpus_conditions` | Explicit condition requests; omitted means clean generation |

A condition request contains `id`, `kind`, and positive integer `count`.
Kinds are `unknown`, `undocumented`, `omitted_source`, `conflict`, `noise`,
and `stale`. Optional exact selectors are `subject_id`, `document_type`,
`viewpoint_id`, and `period`. Each canonical condition instance must repeat its
request's selectors. These are instance counts, not document counts or ratios.
There is no arbitrary selector language, difficulty enum, expected behavior,
decision question, disposition, or scoring field.

```json
{
  "id": "release-disagreement",
  "kind": "conflict",
  "count": 1,
  "subject_id": "product-beacon"
}
```

This is one element of a complete request's `corpus_conditions` list. Enrich the
canon with the matching sources, statements, and condition instance before
`plan`. `init` creates the ordinary baseline, not fabricated case evidence.
An explicit requirement that cannot fit the selected formats, dates, visibility,
or fixed document budget fails before freeze. The helper does not increase the
target or silently drop required viewpoints/sources.

Additional types are per-request definitions, not installed-skill code edits:

```json
{
  "document_types": ["integration_plan"],
  "formats": ["docx"],
  "additional_types": {
    "integration_plan": {
      "formats": ["docx"],
      "anchor_kinds": ["integration"],
      "audience": "internal",
      "required_sections": ["scope", "workstreams", "milestones", "risks"]
    }
  }
}
```

The snippet extends a complete request. Add matching `integration` records with
canonical facts and a natural `data.title` before planning. Missing anchors cause
an explicit error rather than a fallback to another content type.

## Output layout and core info set

The output root has exactly `canon.json`, `documents`, and `build`.
`canon.json` is the single core info set on which content is grounded.
`init` writes it, creates both directories, and puts the resolved request at
`build\request.json`. Organization records (`person`, `department`, `role`) and
term definitions (`glossary`) remain in the canon's `records`; they are not
removed or flattened. No standalone org-chart or glossary files are generated.
Writer contexts derive both views directly from date-filtered canon records.

### Derived organization view

The in-memory bounded organization projection has this shape:

```json
{
  "schema_version": 2,
  "people": [],
  "departments": [],
  "roles": []
}
```

Each collection contains complete canonical records using the same IDs,
`known_on`, `effective_on`, `data`, and `refs` as the canon. People records provide
names; department records provide names; roles provide `person_id`,
`department_id`, `title`, and `reports_to`. Keep all historical assignments,
not just the latest title.

Scope is deliberately bounded: include people/departments referenced by canonical
roles or other records (owners, meeting participants, contacts, quote speakers,
etc.) and their role history. A role assignment counts as a canonical reference,
and managers are included through reporting references. Unused person/department
definitions without relationships are excluded. Referenced contacts without an
internal role remain people entries; do not fabricate an assignment.

The chart is not a full employee roster and must not be expanded to match
financial headcount. Add needed identities and their canonical relationships
before drafting, so they are in the core set first. Draft claims cannot cite
people/departments outside that chart. Unknown extension anchor kinds still need
their reference IDs declared in `refs`.

The writer index derives this projection directly from the canon; there is
no independently editable competing copy. Changing a role or term changes the
canon hash and requires a new revision after freeze. Canon validation still
checks IDs, references, knowability, effective roles, and reporting cycles.

All derived or run-specific artifacts go under `build`, including plan, manifest,
QA/preflight receipts, contexts, drafts, renderer scripts, dependency files,
logs, prompts, images, and temporary renders. Final deliverables alone go under
`documents`. CLI commands reject derived paths outside `build`; audit also checks
the clean root and excludes unmanifested artifacts from `documents`.

Legacy corpora with standalone org/glossary files or plans lacking
`content_depth` do not satisfy the new contract. Preserve their frozen sources,
artifacts, and QA. Migration is a separate, explicitly requested new
run/revision, not automatic file deletion during validation.

## Canon

`schema_version: 2`, `corpus_id`, `request`, and `records` are required. Each record has:

| Field | Meaning |
|---|---|
| `id` | Globally unique record ID, e.g. `financial-2022-q1` |
| `kind` | `company`, `department`, `person`, `role`, `product`, `milestone`, `acquisition`, `financial`, `glossary`, `policy`, or an extension kind |
| `known_on` | First date the fact is knowable to an internal author |
| `effective_on` | Date the event/state takes effect |
| `data` | Structured factual object; no prose-only source of truth |
| `refs` | IDs of dependencies; all must already be knowable |
| `access_groups` | Optional distinct group IDs, declared in `refs`; absence/empty means unrestricted at the record level |
| `knowledge_restricted` | Optional boolean; true requires a timely grant to the writer |

Record shape example:

```json
{
  "id": "milestone-beacon-launch",
  "kind": "milestone",
  "known_on": "2022-08-15",
  "effective_on": "2022-08-15",
  "refs": ["product-beacon"],
  "data": {"product_id": "product-beacon", "version": "1.0", "status": "launched"}
}
```

The baseline generates a company, departments, six fictional people with roles,
a CFO transition when the period has more than one year, two products with
milestones, an announced and subsequently closed acquisition, four quarterly
financial records and one annual aggregate per fiscal year, glossary terms, and
an effective-dated leave policy. It intentionally does not claim to simulate a
full balance sheet, cash flow, consolidated M&A accounting, employment law, or
complete product specifications.

For quarterly/annual `financial`, include `year`, `quarter` (1-4 or `null` for
annual), `period_end`, `revenue_usd`, `cogs_usd`, `opex_usd`,
`operating_income_usd`, and `headcount`. Revenue/costs/headcount are nonnegative;
income may be negative. `revenue - cogs - opex = operating_income`. Annual
financial amounts sum quarters and annual headcount equals Q4 headcount.

Roles include `person_id`, `department_id`, `title`, and `reports_to` (person ID
or `null`). New role records supersede old assignments from their effective
date. One person's two roles cannot start on the same date. Every active manager
must have an active role and the reporting graph must be acyclic. A future
effective date is allowed; do not insert knowledge of a later departure into
an earlier record.

Milestones include `product_id`, `version`, and `status`; acquisitions include
`acquisition_id`, `target`, `status`, and `price_usd`; policies include
`policy_id`, `version`, `owner_id`, and structured obligations. Reference all
foreign IDs using `refs`.

For meeting transcripts, add a `meeting` extension record with participants,
agenda, discussion positions, decisions, action owners, and due dates.
For specs add `requirement` records with product/version, requirement IDs,
functional/nonfunctional constraints, and acceptance criteria. Arbitrary
extension kinds are permitted, but their business invariants require manual
review. Do not assume the helper validates every extension field.

Internal knowability is not public availability. Record `public_on` inside
`data` if it differs; public documents must use a reviewed, public-safe subset
of internal context. The helper's date filtering is not access control.
Keep sensitive fictional internal information out of public press releases.

### Sources, statements, and controls

Use the same record envelope (`id`, `kind`, `known_on`, `effective_on`, `data`,
`refs`) for the following kinds. Every foreign ID must be declared in `refs`;
dependencies cannot be future or self references.

| Kind | Required `data` fields | Optional `data` fields / rules |
|---|---|---|
| `source_version` | `source_id`, positive `version`, descriptive `title`, `author_id` (person), `audience`, `authority_scope`, `status`, selected/supported `document_type`, `published_on` | `viewpoint_id`; `supersedes` an earlier version of the same source. Status: `draft`, `approved`, `reference`, `retracted`. Publication cannot precede knowability. |
| `statement` | `subject_id`, `predicate`, `value`, `kind`, `source_version_id`, `speaker_id` | `units`, `period`, `supports[]`, `qualifiers[]`, `corrects`, `retracts`. The statement must be known by its immutable source's publication date. |
| `viewpoint` | `name`, `author_id`, `audience` | `focus_ids[]`, all referenced and visible to the planned author |
| `access_group` | Nonempty `members[]` of canonical people | Membership becomes usable only when both known and effective |
| `knowledge_grant` | `record_id`, `principal_id` (person) | Both known/effective dates must be no later than the writer's publication date |
| `release` | `record_id`, `groups[]` | Optional boolean `public`; at least one group or public release required, effective/known as of publication |
| `condition` | `request_id`, `kind`, `subject_id` | See instance contract below; generator-only, not writer-visible |

Statement kinds: `fact`, `opinion`, `assumption`, `hypothesis`, `recommendation`,
`risk`, `decision`. Qualifiers: `uncertain`, `dissent`, `unresolved`.
Corrections/retractions are backward relationships within the same
subject/predicate/period/units scope, not exclusive statement
types. Linked source/statement histories cannot be cyclic.

An attributed `fact` is what that source asserts, not an independently
certified objective truth. Differing factual assertions in the same
subject/predicate/period/units scope require an explicit requested conflict
unless connected by correction/retraction/source supersession. Opinions may
differ without being treated as failed actuals reconciliation.

```json
{
  "id": "position-engineering",
  "kind": "statement",
  "known_on": "2022-02-01",
  "effective_on": "2022-02-01",
  "refs": ["product-beacon", "readiness-source-v1", "person-riley"],
  "data": {
    "subject_id": "product-beacon",
    "predicate": "release_readiness",
    "value": "Additional load testing is needed",
    "kind": "opinion",
    "source_version_id": "readiness-source-v1",
    "speaker_id": "person-riley",
    "qualifiers": ["uncertain", "unresolved"]
  }
}
```

This record extends a canon containing the referenced source and people. A
source version links to its author and subject through `refs`; its statements
link back to the source, avoiding a source -> statement -> source dependency
cycle. New approved/corrected sources are new records; do not overwrite old
assertions or embed later corrections into old writer contexts.

Controls (`viewpoint`, `access_group`, `knowledge_grant`, `release`, `condition`)
are private canonical metadata. Writer records omit them and strip their
reference IDs/access-policy fields. The context supplies the selected
viewpoint's name, author, audience, and visible focus separately. A source's
authority/status is preserved but never converted into a universal trust score.

### Condition instance contract

Each `condition` record corresponds to exactly one requested instance:

- `request_id`, `kind`, and `subject_id` are required.
- `statement_ids[]`, `source_version_ids[]`, and `blocked_record_ids[]` default
  to empty lists; all IDs are unique and referenced.
- Optional `document_type`, `viewpoint_id`, and `period` restrict compatible
  carrier documents and must agree with corresponding request selectors.
  Period tokens match canonical topic years (`2022`), financial quarters
  (`2022-Q1`), or an immutable source statement's exact `period` string.
  Source-period scope remains available to the planner even when assertions
  are deliberately suppressed; it does not disclose their values to writers.
- `conflict` needs at least two statements with differing values in the same
  subject/predicate/period/units scope. Its subject/period must match those
  statements. Their source versions are mandatory planned artifacts.
- `unknown`/`undocumented` need blocked record IDs. Optional source-version IDs
  choose carriers; otherwise carriers match the subject's anchor/dependencies.
  Unknown records are excluded from author knowledge. Undocumented records
  remain writer-visible but are suppressed from final claims and derivations.
- `omitted_source` needs source-version IDs that must not be planned/exported.
  Other subject-related documents carry the condition; no missing file is
  counted as a successful omission.
- `noise` needs source-version IDs for complete adjacent material. Relevance
  is relative to the configured subject, not an assertion that the document
  has no value in any use case.
- `stale` needs source-version IDs plus `as_of_date`; each old source must have
  a published superseding version by that reference date.

There must be exactly the requested number of canonical instances before
planning. Mandatory sources and viewpoints consume the existing budget.
Compatible instances may share documents; incompatible gap/conflict overlap
fails. Deliberate controls never excuse broken business actuals, accidental
unregistered conflicts, failed exports, or missing genre prerequisites.

Structured absence checks do not prove that every semantic paraphrase or
inference is absent. Configured carriers therefore also need semantic
`condition_review` in QA.

## Plan and context

Plan fields: `schema_version: 2`, `corpus_id`, `canon_sha256`, `documents`, `warnings`.
Each document includes `id`, `type`, `format`, `title`, `filename`, `publication_date`,
`audience`, `anchor_ids`, `revision`, `content_depth`, and `required_sections`. The count is exactly the requested
target. Formats are allocated evenly. Types and fiscal-year coverage are
balanced deterministically, preferring less-used anchors; inspect warnings when
the budget cannot cover all requested years or types.

Every document also carries `author_id`, nullable `viewpoint_id`, `access_groups`,
nullable `source_version_id`, `required_statement_ids`, `condition_ids`,
`blocked_record_ids`, and `suppressed_record_ids`. Required source statements
must be knowable and represented as value claims; their publication date and
genre come from the source. Requested sources/viewpoints are allocated before
ordinary balancing, so explicit controls may override even distribution.
Missing selected-format coverage remains an error.

Default intended access is public for public genres, `group-board` for board
content, and `group-internal` otherwise. A source may explicitly specify its
own group list. Public documents cannot have restricted intended access.
Access groups are delivery metadata, not filesystem/service permissions.

`title` is a natural, subject-specific business title, not a content-type label
with internal IDs. `filename` is the required export basename, including its
Office extension. For example:

```json
{
  "title": "Lumenvale Systems - Q1 2022 Operating Results",
  "filename": "2022-04-15 - Lumenvale Systems - Q1 2022 Operating Results.xlsx"
}
```

Titles describe canonical periods, products/versions, acquisition stages, or
policies. Repeated anchors get follow-up or supplement suffixes. Filenames
start with the publication date, contain only letters, digits, spaces, hyphens,
underscores, and periods, and are at most 140 characters. Long names retain
their beginning and end. Case-insensitive collisions get an `Edition N` suffix;
the planner never uses opaque document IDs as display-name disambiguators.
`content_depth` must equal the canon request's resolved setting (default medium).
The depth is included in the plan fingerprint; a missing or changed setting
requires a new revision, not an in-place update of accepted drafts.
Names must be reviewed for semantic accuracy alongside the draft's claims.

Existing plans without `filename`, or with the old ID-based titles, must be
replanned in a new revision. Do not mutate accepted drafts or rename files
behind an existing manifest: changing the plan invalidates its draft/QA hashes.

Context includes the sanitized plan item, canon and plan fingerprints, company
record, records available to the author at publication, and `active_roles`.
It also includes date-filtered `glossary` and `org_chart` views. `org_chart`
contains `people`, `departments`, visible role history in `roles`, and an
`active_roles` snapshot based on both knowability and effective dates. The
top-level `active_roles` is retained for compatibility and matches that snapshot.
People/departments omitted from the bounded chart are also omitted from writer
records and are unavailable to draft claims.
Do not copy future states from an unfiltered canon into a draft.

Projection uses dates, group access, explicit releases, author knowledge grants,
the bounded organization, and dependencies. A missing business dependency
removes its dependent record too. Conditions and full access policies are not
writer inputs. The context's document omits condition IDs, unknown record IDs,
and access groups; `suppressed_record_ids` instructs the writer not to state
otherwise known information. These IDs and all other scaffolding stay private.

## Drafts and grounded claims

Each draft file under `build\drafts` is named `<document-id>.json`. Required fields:

```json
{
  "schema_version": 2,
  "corpus_id": "<from canon>",
  "document_id": "doc-0001",
  "canon_sha256": "<from plan>",
  "plan_sha256": "<from context>",
  "title": "<exact plan title>",
  "publication_date": "<exact plan date>",
  "content_depth": "<exact plan content_depth>",
  "claims": [
    {
      "id": "claim-1",
      "record_id": "company",
      "path": ["data", "name"],
      "value": "Lumenvale Systems"
    }
  ],
  "sections": [
    {
      "role": "<one of the plan's required_sections>",
      "heading": "Overview",
      "text": "Lumenvale Systems ...",
      "claim_ids": ["claim-1"]
    }
  ]
}
```

Paths are arrays of object keys or list indexes relative to the complete
canonical record. No JSON Pointer escaping is needed. Claims must exactly equal
the referenced value and must be knowable on the publication date. Claims need
unique IDs; all must be used in a section. Every anchor must have at least one
claim. Sections may add format-specific structures (tables, charts, slide
layouts, speaker turns) but retain `role`, `heading`, `text`, and `claim_ids` for review.
The checker verifies receipts, not the truth of natural-language text.
Claims for required statements must use `path: ["data", "value"]`.
Optional `statement_kind` must match the referenced statement's kind (ordinary
business records are facts). A section's optional `speaker_id` must be visible;
unless `reported_speech: true`, its attributed statement claims must belong to
that speaker. Reported speech still needs faithful semantic attribution review.
Claims cannot reference unknown/suppressed records or their dependent records.
Direct claims must also be accessible/released to the intended readers, not
just to the author.

Multi-input derived claims replace `record_id`/`path` with `derivation`:

```json
{
  "id": "margin",
  "value": "25.0",
  "units": "percent",
  "derivation": {
    "operation": "ratio",
    "inputs": ["income", "revenue"],
    "rounding": 1
  }
}
```

Both input claims must already be grounded and have units. Operations: `sum`,
`difference` (two inputs), `product`, and `ratio` (two inputs). Sum/difference
need matching input/output units; ratios need equal input units and `ratio` or
`percent` output. Rounding is 0-12 decimal places, half-up. Cyclic references,
unsupported operations, invalid results, and division by zero fail. This small
receipt evaluator does not execute arbitrary expressions or replace OfficeCLI
formula evaluation.
Derived and primary claim shapes cannot be mixed. Primary units must match
explicit source units; `_usd` financial fields require `USD`. Product output
units still need semantic dimensional review: the evaluator is not a general
unit-conversion or dimensional-algebra engine. Paths cannot expose stripped
access policies or private dependency references from writer contexts.
Draft `content_depth` must exactly match the planned setting. Narrative length
and analytical richness are assessed against the run's depth targets during QA,
not inferred from the number of JSON sections or claims.
Every required genre role must have a nonempty section; extra roles are allowed.
Do not export role IDs, claim receipts, or draft scaffolding as business prose.
Final documents have no synthetic/fake/fictional/AI-generation labeling.
The helper scans common self-description phrases in drafts and Office XML;
manual semantic/genre review must check the entire document and its properties.

## Manifest and QA

The manifest lives at `build\manifest.json`. Paths are relative to `--root`, and may
not escape that directory, including through symbolic links.

```json
{
  "schema_version": 2,
  "corpus_id": "<from canon>",
  "canon_sha256": "<from plan>",
  "plan_sha256": "<from context>",
  "documents": [
    {
      "document_id": "doc-0001",
      "path": "documents/2022-02-01 - Lumenvale Systems - Beacon 1.0 Release Plan - Leadership Meeting Transcript.docx",
      "sha256": "<SHA-256 of actual bytes>"
    }
  ]
}
```

JSON paths may use `/` portably; Windows shell commands use `\`.
Only primary Office files count toward the requested target. The manifest must
contain exactly one entry per plan document. Package audit verifies CRCs,
parsable XML, and mandatory parts, not full OOXML schema conformance.
Each artifact's basename must exactly match the corresponding plan `filename`;
a generic filename is rejected even when the bytes and hash are valid.

### Execution and review scope (new runs)

Business `schema_version` stays **2**. New execution sidecars separately use
`execution_version: 1` / `review_version: 1`. Create
`build\execution-plan.json` through `company.execution_plan(canon, plan, layouts,
pilot_ids, qa_policy="FullContentSampledVisual", revision=1, flagged_ids=[],
budget_seconds=None)`. Its header is `receipt_header(canon, plan)`; every layout
entry contains `document_id`, nonempty `layout_revision`, and `layout_sha256`
for a frozen layout file under `build`. `required_features` maps every selected
format to tested feature names (default `basic`, `html_screenshots`); additional
features require explicit bound capability probes, not a version-only claim.
Every family needs a pilot. Optional
budget is positive user-supplied seconds, never an inferred time promise.

| Policy | Complete content review | Complete visual review |
|---|---|---|
| `Sampled` | Deterministic sample + mandatory documents | Same selection |
| `FullContentSampledVisual` (default) | Every document | Deterministic sample + mandatory documents |
| `Full` | Every document | Every document |

Sampling takes first publication/date-ID and largest normalized draft-section
content per `(type, format, layout revision, layout hash)`, with stable date/ID
ties. Always include pilots, condition carriers, both ends of source
supersession/correction/retraction relationships, and flagged exceptions.
`flagged_ids` and corrected mechanical failures expand **the entire affected
family** for content and visual review. A failed sample blocks the family;
record it in a new explicit execution revision before resuming. Do not erase
findings or reselect an easier sample. Policy changes require an explicit
execution revision, recomputed review plan and QA; unchanged valid bytes/evidence
can be reused without reauthoring. Depth remains independently light/medium/deep.

Every policy checks every draft/file for canonical claims, dates/access,
required genre sections, measurable depth, structure/self-labeling, schema/issues,
title/filename/hash, supported saved formula values, provenance and coverage.
These mechanical guards do not certify semantic entailment or substantive depth.

`make_review_plan` recomputes `build\review-plan.json`, bound to canon/plan,
execution fingerprint, preflight file hash, per-document draft/artifact/machine
hashes, normalized sizes, selection reasons, and `content_required` /
`visual_required`. Audit recomputes and compares the entire plan, not policy
strings or claimed counts. Nonselected gates must be exactly
`{"status": "not_selected"}`, not successful reviews.

The runner writes `build\machine\<document-id>.json` for complete mechanical
checks. It binds final artifact/draft hashes, OfficeCLI identity, the
Playwright-Chromium capture-tool identity, a hash-bound capability snapshot,
tool-result files, measured depth targets, and actual HTML-export +
Playwright-capture render targets/results/images. Every required capture
covers all pages/slides or all populated worksheet regions; cropped or
partial replacement is not a full review. Content/condition review evidence
and visual evidence must be written by an actual reviewer, not the runner.

Successful read-only inspections may run again when an explicitly revised
artifact has a new hash. These rechecks do not consume corrective retries or
expand review by themselves. Failed, timed-out, and interrupted inspection
attempts remain counted across resume and hash changes; a second failure
blocks further attempts. Mutating operations retain their two-attempt limit.

`build\qa.json` records honest review evidence. For new runs it adds
`execution_version: 1`, `execution_sha256` (execution object's fingerprint),
`review_plan_sha256` (review-plan object's fingerprint), and `qa_policy`.
Every document also binds `machine_sha256`; this example is for a selected
DOCX with no configured condition:

```json
{
  "schema_version": 2,
  "corpus_id": "<from canon>",
  "canon_sha256": "<from plan>",
  "plan_sha256": "<from context>",
  "execution_version": 1,
  "execution_sha256": "<execution object fingerprint>",
  "review_plan_sha256": "<review-plan object fingerprint>",
  "qa_policy": "FullContentSampledVisual",
  "documents": [
    {
      "document_id": "doc-0001",
      "artifact_sha256": "<same as manifest>",
      "draft_sha256": "<SHA-256 of draft JSON file>",
      "machine_sha256": "<SHA-256 of machine receipt file>",
      "semantic_review": {
        "status": "passed",
        "evidence": "Reviewed every section against date-filtered claims; no unsupported facts."
      },
      "genre_review": {
        "status": "passed",
        "evidence": "Reviewed the complete document against its genre profile; natural business voice, no placeholders or generation self-labeling."
      },
      "visual_review": {
        "status": "passed",
        "images_delivered": true,
        "render_ids": ["all"],
        "evidence": "Rendered pages and reviewed pagination, tables, footers, and business branding."
      },
      "depth_review": {
        "status": "passed",
        "content_depth": "medium",
        "evidence": "Describe measured delivered content and reviewed substantive coverage against the medium genre targets; never prefill this example as a pass."
      },
      "condition_review": {"status": "not_selected"}
    }
  ]
}
```

New-run XLSX machine evidence requires every saved formula to match a grounded
expected derived value and an explicit `chart_cache_stale` scan, not a
recalculation-on-open flag. Selected visuals bind every render ID and require
`images_delivered: true`; missing/undisplayed/unreadable images never count.
Every selected content review requires `depth_review` with passed status, nonempty evidence,
and `content_depth` equal to the plan. Evidence should cite actual delivered
coverage, measurements, and useful analyses against `build\depth-targets.json`;
length alone or a larger collection of sheets is not evidence of depth.
Changing a plan, draft, or output invalidates its QA receipt. The audit checks receipt
presence and hashes; it cannot independently establish that a human or AI
actually performed the claimed review.

Configured carrier documents also require passed `condition_review` with
nonempty evidence describing actual gap/conflict/noise/staleness inspection.
Final audit recomputes `build\provenance.json` from `claim-locations.json` and
final files, and recomputes `build\coverage.json`. Merely setting a passed flag
cannot satisfy these checks. See `provenance.md` for the mapping/binding schemas.
All per-section claim uses require distinct final-artifact occurrences.

Coverage reports planned/delivered document/type/format counts, grounded
record/entity/topic/subject references, periods, viewpoints, statement kinds,
qualifiers, source versions, claim/occurrence totals, and configured conditions.
Publication years, canonical topic years, and covered reporting/statement
periods are separate dimensions; a report published the next year is not
misclassified as that year's actuals. Period counts are document coverage and
can include more than one period per document.
The denominator is the requested plan and grounded claims, not all possible
enterprise knowledge. It measures structural corpus integrity, not semantic
depth or an evaluated agent's performance.
Period coverage describes a document's modeled scope, not a guarantee that a
suppressed source assertion was delivered; statement and condition coverage
report that distinction separately.

`access-view` first audits the corpus and returns only eligible document
identities, revisions/hashes, filenames, and filtered occurrence metadata at
the requested date/principal. Future/restricted documents and private build
annotations are omitted. Its output belongs under `build\views`; it does not
copy artifacts, enforce ACLs, or upload content. Distribute selected evidence,
not the entire canon/build tree, through a separate deployment tool.

Legacy schema-v2 QA with **no execution sidecar** retains the existing
all-document semantic/genre/depth/visual gates, plus its XLSX `recalculation`
and carrier condition receipts. That compatibility path does not recertify
native fidelity or silently upgrade/downgrade old evidence. New sidecars cannot
bypass enforcement by omitting version/policy fields. No example migration.

Generation passes authorized OfficeCLI dependency/setup and selected-format
capability probes before corpus creation; see `dependencies.md`. Audit binds
passed probe logs/images/tool identity, not a version-only assertion. Keep
immutable capability snapshots so a verified later update does not falsely
rewrite the authoring engine for accepted historical bytes.

### Thin OfficeCLI runner job contract

The standard-library `scripts\officecli_runner.py` is a mechanical runner, not
an authoring framework. Store a JSON array of jobs under `build`; each job has:

- `document_id` and `layout_file` (a build-relative frozen layout/template file).
- Either `batch_file` (a build-relative nonempty atomic DOM command array) **or**
  `template` plus `data_file` for OfficeCLI merge; never force-overwrite.
  Authoring batches support checked add/set/remove/move/swap, no raw cache edits.
- `expected_formulas` for XLSX: every `{sheet, cell, formula, value, claim_id}`.
  `claim_id` binds an independently checked draft `derivation`; bindings are in
  dependency order, and include every formula. Unsupported/shared/array forms
  stop rather than inventing values. Empty bindings are valid only for no formulas.
- `render_targets`: `[{"id": "all"}]` for DOCX/PPTX; for XLSX,
  `[{"id": "results", "type": "range", "range": "Results!A1:F40"}, ...]`,
  covering all populated cells. Declare targets even if currently unselected
  so expansion can render accepted bytes later. Standard column/line chart
  objects (`twoCellAnchor`-only) are a supported, separately identified
  target type: `{"id": "...", "type": "chart", "sheet": "...", "shape_id": "..."}`,
  resolved against a freshly recomputed `chart_object_inventory()`, never a
  caller-trusted `expected_dom_path`. `verify_render_targets()` fails closed
  if the package contains any unsupported drawing/chart object, or if any
  supported chart object lacks a matching target (range-only targets cannot
  satisfy chart coverage). A target's own `expected_dom_path`, if supplied,
  must match the current inventory exactly — editing/reordering the workbook
  invalidates a stale path rather than silently matching the wrong object.
- `occurrences`: exact exporter-associated entries from the occurrence contract
  in `provenance.md`, never global first-token matches. `assemble` verifies the
  entire real planned set before writing complete-corpus receipts.

Run pilots first with a subset job file. After actually reviewing each pilot,
write its hash-bound review entry to `build\reviews\<document-id>.json`; all five
review gates are present (`condition_review` is `not_selected` if no carrier).
Only then run family siblings. Draft validation/sample sizing requires the
complete frozen draft set; it does not give writers unfiltered canonical data.

Checkpoints bind canon/plan/draft/job/source/final hashes. Verified completed jobs
skip mechanical work; changed authoring inputs stop and require an explicit
affected revision, preserving accepted files. Policy/exception changes rebuild
only missing review evidence. `assemble` writes manifest/index/mappings/provenance/
coverage/review-plan, **never** editorial/visual passes or a completed QA audit.
Run the ordinary `audit` after required reviews populate `qa.json`.

After an initial atomic authoring failure, an approved `correction_file` and
nonempty `correction_reason` may be supplied without changing the original job.
They are bound in retry evidence, not used to overwrite a completed artifact.
For externally corrected read/render failures, `retry_reasons` maps operation
names (e.g. `render-all`) to concrete corrective actions. One corrective retry is
persisted across resume; a second failure blocks. Missing/stale numeric caches
get at most one automatic known-order formula rewrite through OfficeCLI,
followed by fresh cache/schema/issues/hash checks. No substitute calculator.
Phase logs distinguish active execution from between-invocation waits.

## Depth targets

`build\depth-targets.json` is a planning/review sidecar, not an additional source
of business facts. Set targets after planning but before drafting and pilot
review. Cover every planned document, including custom types:

```json
{
  "schema_version": 2,
  "plan_sha256": "<from plan>",
  "documents": [
    {
      "document_id": "<planned press-release ID>",
      "content_depth": "medium",
      "measure": "delivered_words",
      "target_min": 450,
      "target_max": 700,
      "coverage": [
        "Supporting announcement facts and practical stakeholder implications",
        "Implementation timing, constraints, and meaningful background"
      ],
      "adaptation_rationale": null
    }
  ]
}
```

Use genre-appropriate measures such as delivered words, discussion words,
meaningful slides, distinct requirements, or useful analysis modules.
`adaptation_rationale` records a justified change to the playbook's targets;
it is not permission to drop required genre sections or unsupported analysis.
Preserve accepted targets with the run revision. A plan change invalidates
their binding. The runner/audit measure saved content using `delivered_words`,
`delivered_characters`, `delivered_containers`, `delivered_slides` or
`delivered_sheets`. These are quantity guards, not meaningful-slide/sheet or
substantive-analysis certification. Genre-specific qualitative targets remain
part of actual content/pilot review; unsupported required measures need a
bounded checked adapter before generation, not silent weakening.
Word/character/container measures deduplicate paragraph/table-cell and other
physical-text aliases; two locators for the same text cannot inflate depth.
