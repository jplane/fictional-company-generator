# Data contracts

All JSON is UTF-8. Dates are ISO `YYYY-MM-DD`, IDs are stable strings, financial
amounts are integer USD, and version 1 uses calendar fiscal years. Python 3.10+
is required. Invalid input returns a nonzero exit code with an explicit error.

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
`references\content-types.md` defines substantive coverage and measurable
planning targets for each genre. Record run-specific targets and justified
adaptations in `build\depth-targets.json` before drafting. The helper validates
the setting and its propagation, not word counts, sheet counts, or analytical
quality; those require an honest depth review of delivered content.

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
  "schema_version": 1,
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

`schema_version`, `request`, and `records` are required. Each record has:

| Field | Meaning |
|---|---|
| `id` | Globally unique record ID, e.g. `financial-2022-q1` |
| `kind` | `company`, `department`, `person`, `role`, `product`, `milestone`, `acquisition`, `financial`, `glossary`, `policy`, or an extension kind |
| `known_on` | First date the fact is knowable to an internal author |
| `effective_on` | Date the event/state takes effect |
| `data` | Structured factual object; no prose-only source of truth |
| `refs` | IDs of dependencies; all must already be knowable |

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

## Plan and context

Plan fields: `schema_version`, `canon_sha256`, `documents`, `warnings`.
Each document includes `id`, `type`, `format`, `title`, `filename`, `publication_date`,
`audience`, `anchor_ids`, `revision`, `content_depth`, and `required_sections`. The count is exactly the requested
target. Formats are allocated evenly. Types and fiscal-year coverage are
balanced deterministically, preferring less-used anchors; inspect warnings when
the budget cannot cover all requested years or types.

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

Context includes the complete plan item, canon and plan fingerprints, company
record, all records with `known_on <= publication_date`, and `active_roles`.
It also includes date-filtered `glossary` and `org_chart` views. `org_chart`
contains `people`, `departments`, visible role history in `roles`, and an
`active_roles` snapshot based on both knowability and effective dates. The
top-level `active_roles` is retained for compatibility and matches that snapshot.
People/departments omitted from the bounded chart are also omitted from writer
records and are unavailable to draft claims.
Do not copy future states from an unfiltered canon into a draft.

## Drafts and grounded claims

Each draft file under `build\drafts` is named `<document-id>.json`. Required fields:

```json
{
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
Draft `content_depth` must exactly match the planned setting. Narrative length
and analytical richness are assessed against the run's depth targets during QA,
not inferred from the number of JSON sections or claims.
Every required genre role must have a nonempty section; extra roles are allowed.
Do not export role IDs, claim receipts, or draft scaffolding as business prose.
Final documents have no synthetic/fake/fictional/AI-generation labeling.
The helper scans common self-description phrases in drafts and Office XML;
manual semantic/genre review must check the entire document and its properties.

### Backend export structures

The container CLI consumes the actual canon, frozen plan, per-document
date-filtered context, and grounded draft. It never creates replacement canon
or prose. The supported structures below are rendered; unknown top-level,
section, slide, table, or chart fields stop export with an explicit error.

A section may contain simple tables with `title`, `headers`, and `rows`. DOCX
renders each table with its section. XLSX renders section content and tables
on a named worksheet. An XLSX table cell may be an explicit formula object; its
optional expected value is checked after LibreOffice Calc recalculation:

```json
{"formula": "=SUM(B2:B4)", "expected_value": 17}
```

Formula expressions may not link to an external workbook. A draft without
formula cells remains a narrative/source workbook; the backend does not invent
calculations or claim formula results were verified.

PPTX drafts may add top-level `slides`, with exactly one slide per grounded
section. Each slide has a matching `heading` and complete section `text` as a
string or list of strings in `body`. Supported optional fields are `subtitle`,
`badge`, `notes`, `layout`, `chart`, and `tables`. Layouts are `cover`, `cards`,
`grid`, `statement`, `capabilities`, `workflow`, `stats`, `chart`, `closing`,
and `content`. A simple chart has text `labels`, finite numeric `values`, and
optional `kind` (`bar`, `column`, `line`, `pie`, or `doughnut`) and `title`.
Speaker notes must repeat material present in the cited section. Unsupported
layout/data shapes or mismatches between slides and sections are rejected,
never dropped silently.

`claims` and `measurements` are private receipts and are not rendered.
`requirement_rows` render as a traceability table in DOCX or XLSX; structured
`model` data render into an XLSX source-model worksheet. IDs, claim receipts,
hashes, and generation provenance are not placed in business-facing content.

## Manifest and QA

The manifest lives at `build\manifest.json`. Paths are relative to `--root`, and may
not escape that directory, including through symbolic links.

```json
{
  "schema_version": 1,
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
The container writes `build\manifest.json` only after the complete plan is
exported. An invocation restricted with `--document` writes
`build\backend-pilot-manifest.json` and `build\backend-pilot-evidence.json`,
explicitly scoped to selected plan items; these are not complete-corpus
receipts and cannot satisfy the full-plan audit.

`build\qa.json` records honest review evidence:

```json
{
  "schema_version": 1,
  "canon_sha256": "<from plan>",
  "plan_sha256": "<from context>",
  "documents": [
    {
      "document_id": "doc-0001",
      "artifact_sha256": "<same as manifest>",
      "draft_sha256": "<SHA-256 of draft JSON file>",
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
        "evidence": "Rendered pages and reviewed pagination, tables, footers, and business branding."
      },
      "depth_review": {
        "status": "passed",
        "content_depth": "medium",
        "evidence": "Describe measured delivered content and reviewed substantive coverage against the medium genre targets; never prefill this example as a pass."
      }
    }
  ]
}
```

For XLSX also require `recalculation` with `status: "passed"` and nonempty
`evidence` identifying the selected calculation application/runtime (Microsoft
Excel or LibreOffice Calc), actual recalculation operation, formula count,
cached-result checks, and formula-error scan. If the source draft has no
formulas, state this explicitly and do not claim expected formula values were
verified. DOCX/PPTX visual-review evidence must identify the selected rendering
application/runtime (Word/PowerPoint or LibreOffice Writer/Impress) and the
rendered pages/slides inspected. LibreOffice is not Microsoft Office and does
not certify Office fidelity.
Every document requires `depth_review` with passed status, nonempty evidence,
and `content_depth` equal to the plan. Evidence should cite actual delivered
coverage, measurements, and useful analyses against `build\depth-targets.json`;
length alone or a larger collection of sheets is not evidence of depth.
Changing a plan, draft, or output invalidates its QA receipt. The audit checks receipt
presence and hashes; it cannot independently establish that a human or AI
actually performed the claimed review.

Before any corpus artifacts are created, pass the selected-format capability
check and record the recommended and user-selected backend in
`build\preflight.json`. For Native Office, record the selected applications,
versions, COM connection/operation evidence, implementation source, source
SHA-256 hashes, and Git revision when available. If the checkout has no `.git`
entry, record `revision: null` with
`revision_status: "not-in-git-checkout"` and rely on the hashes; Git metadata
is not required to pass preflight. If `.git` exists but Git is unavailable,
record `revision: null` with `revision_status: "git-unavailable"`. For the
container, record Docker/Compose versions, image digest and repository
revision, container Python/package and LibreOffice versions, helper-mount and
smoke-test evidence. Record actual pilot results for either backend. State
plainly that LibreOffice pagination, formatting, and calculation may differ
from Office; do not claim Office fidelity. The preflight receipt is not
independently enforced by the Python audit. A missing application, image, or
capability is a stop condition; never fall back to another engine.

## Depth targets

`build\depth-targets.json` is a planning/review sidecar, not an additional source
of business facts. Set targets after planning but before drafting and pilot
review. Cover every planned document, including custom types:

```json
{
  "schema_version": 1,
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
their binding. The helper does not parse this sidecar or independently measure
content; the skill's pilot and depth-review workflow must verify coverage,
targets, and actual delivered measurements honestly.
