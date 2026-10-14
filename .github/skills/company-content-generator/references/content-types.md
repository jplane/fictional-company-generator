# Content-type playbook

Write the actual genre, not a summary about the genre. Public communications
use an outward-facing company voice; technical documents use precise,
testable language; meetings use natural speaker turns; presentations tell a
visual, audience-specific story. Do not reuse identical prose across formats.
No deliverable identifies itself, its company, or its speakers as synthetic,
fake, fictional, simulated, or AI-generated. Keep provenance in `build`.

Use subject-specific business names, not generic type names or internal IDs.
Reports name their fiscal period; specs name the product and version; press
releases name the target/product and distinguish planned/announced from
launched/closed; policies name their subject and revision; meetings and board
briefings name the actual agenda topic. Follow the plan's title and filename.

| Type | Default formats | Expected content |
|---|---|---|
| `meeting_transcript` | DOCX | Meeting date, named participants and their current roles, agenda, timestamped speaker turns, decisions, action owners, due dates. Natural discussion, not a narrated event summary. |
| `press_release` | DOCX | Headline, dateline, news-led opening, substantive announcement, attributed executive quote, company boilerplate, and media contact. Distinguish announced from closed M&A. |
| `product_presentation` | PPTX | Customer problem, clear value proposition, version-specific capabilities, workflows/use cases, evidence from canon, and concrete next steps. Customer-facing structure, not a board report. |
| `product_spec` | DOCX, XLSX | Product/version, goals, scoped requirements, constraints, interfaces, acceptance criteria, dependencies, risks. XLSX is a traceability matrix, not prose stuffed into cells. |
| `quarterly_report` | DOCX, PPTX, XLSX | Quarter actuals, canonical KPIs, grounded product/org developments, risks, labeled outlook. Do not invent comparative periods that do not exist. |
| `annual_report` | DOCX, PPTX, XLSX | Reconciled annual financials, strategy and events, operating review, year-end organization, risk factors. Year-end state differs from publication-date state. |
| `hr_policy` | DOCX | Policy name/version, owner, effective date, scope, eligibility, obligations, exceptions, review process. Write direct administrative rules, not advice about policy writing. |
| `board_deck` | PPTX | Date-specific executive overview, financial charts, product timeline, M&A status, risks, canonically recorded decisions requested. Each slide has a meaningful visual. |
| `financial_model` | XLSX | Source actuals, explicit units, separate canonical forecast assumptions if requested, formula-linked income and margins, notes and provenance. Recalculate and scan errors. |

## Genre requirements and editorial acceptance

Each draft section has a machine-readable `role`. The planner supplies the
required roles below; sections may combine into natural paragraphs, slides, or
worksheets at export. Do not export role IDs, claims, or generic scaffolding
headings just because they exist in draft JSON.

| Type | Required section roles |
|---|---|
| `press_release` | `dateline_and_lead`, `announcement_details`, `attributed_quote`, `company_boilerplate`, `media_contact` |
| `product_presentation` | `customer_problem`, `value_proposition`, `capabilities`, `use_cases`, `next_steps` |
| `meeting_transcript` | `participants`, `discussion`, `decisions`, `actions` |
| `product_spec` | `scope`, `requirements`, `acceptance_criteria` |
| `quarterly_report` | `financial_results`, `operating_review`, `outlook_and_risks` |
| `annual_report` | `financial_results`, `business_review`, `organization`, `risks` |
| `hr_policy` | `scope`, `policy_rules`, `administration` |
| `board_deck` | `executive_summary`, `performance`, `risks`, `decisions_requested` |
| `financial_model` | `source_actuals`, `calculations`, `assumptions_and_notes` |

## Content depth: light, medium, deep

Read the planned `content_depth` before authoring. Default **medium** means a
substantive business document, not the minimum number of sentences necessary
to fill the required section roles. All depths must satisfy the complete genre.
Depth controls source granularity, reasoning, useful analyses, and practical
detail; it does not change the corpus's requested document count.
It also does not change source validity, attribution, access boundaries,
provenance completeness, or configured condition counts. Deep is not a
synonym for difficult, conflicting, or noisy.

### Richness of sources and viewpoints

| Capability | Light | Medium | Deep |
|---|---|---|---|
| Statement semantics | Clearly distinguish the included facts, opinions, assumptions, and proposals | Explain their basis and qualifications | Develop dependencies, competing interpretations, and uncertainty where relevant |
| Viewpoints | Authentic author/audience framing | Substantive differing priorities | Connected, source-grounded viewpoints across documents where requested |
| Source history | Correct current/historical status | Explain meaningful changes/corrections | Richer version history and longitudinal analysis |
| Configured gaps/conflicts | Faithful concise omissions or differing assertions | Context and consequences known to the author | Richer development without automatically resolving conflicts or filling gaps |
| Provenance and visibility | Complete for included substantive claims | Same guarantees | Same guarantees |

Do not force every semantic category into every genre. A deep press release is
still a press release, not an epistemic inventory. Statements and opinions
remain attributed; a source's confidence or approval is not proof of truth.
Use natural language, not visible `FACT`/`OPINION` generation labels.

Deliberate unknowns, undocumented information, omitted sources, and stale
versions must be configured before freeze. They do not permit broken genre
sections, empty placeholders, false actuals, or random missing exports.
The quality review distinguishes faithful source disagreement from an
accidental authoring error and verifies that added depth did not undo a gap.

### Measurable planning targets

Use these ranges to set concrete per-document targets in
`build\depth-targets.json` before drafting. They are editorial targets, not
automatic proof of quality. Record and justify adaptations to the real scope,
industry, audience, or format before writing; do not rationalize thin documents
after export. Measure delivered content, excluding repeated headers, boilerplate
scaffolding, and private claims. Approved quotes and business boilerplate count
as normal press-release content.

| Genre / measure | Light | Medium (default) | Deep |
|---|---|---|---|
| Press release / words | 200-350 | 450-700 | 800-1,200 |
| Product presentation / meaningful slides | 5-7 | 9-13 | 15-22 |
| Meeting transcript / discussion words | 600-1,000 | 1,500-2,500 | 3,000-5,000 |
| Product specification / distinct testable requirements | 6-10 | 12-20 | 24-40 |
| Quarterly report / narrative words (DOCX) or slides (PPTX) | 600-1,000 / 5-7 | 1,500-2,500 / 9-13 | 3,000-5,000 / 15-22 |
| Annual report / narrative words (DOCX) or slides (PPTX) | 1,000-1,600 / 7-10 | 2,500-4,000 / 13-18 | 5,000-8,000 / 22-30 |
| HR policy / words | 400-700 | 900-1,500 | 1,800-3,000 |
| Board deck / meaningful slides | 6-8 | 10-15 | 18-25 |
| Financial model and financial XLSX reports / useful analysis modules | 3 | 6 | 9 |

XLSX specifications use the requirement targets, not financial-module counts.
DOCX specifications add explanatory scope, interface, verification, and release
control detail around their requirements. PPTX depth uses concise visuals and
explanatory notes rather than paragraphs pasted onto slides. Notes may deepen
the explanation, but cannot substitute for missing visible story/analyses.
Do not stretch a short release into an artificial multi-page report.

A **useful analysis module** answers a distinct business question with a
source-linked table/model, formulas where applicable, a reconciliation or
validation check, and an interpretation. A sheet, chart, repeated view, or
single SUM formula is not automatically a module. Several modules may share
a sheet; do not create empty or duplicate sheets to reach a quota.

### Substantive coverage by genre

Medium includes light's coverage; deep includes medium's coverage. Enrich the
canon before freeze to supply all detail needed by the selected profile.

| Genre | Light | Medium adds | Deep adds |
|---|---|---|---|
| Press release | Complete news lead, concrete scope/status/timing, rationale, approved quote, boilerplate, contact | Several supporting facts, practical implications for customers/dealers/employees, implementation or availability details, meaningful background | Detailed rollout/transition mechanics, limitations/dependencies, richer approved evidence and background, more than one perspective where approved and useful |
| Product presentation | Problem, value, version-specific capabilities, one workflow, next step | Multiple distinct use cases, adoption requirements, constraints, supporting evidence, differentiated capability visuals | End-to-end workflows, integration/implementation detail, alternatives and tradeoffs, objections/FAQ, grounded roadmap and evidence appendix |
| Meeting transcript | Timestamped back-and-forth, questions and responses, decisions, owned/due actions | Substantive disagreement or alternatives, financial/operating implications, clarifications, decision rationale, dependencies | Multiple connected agenda threads, probing follow-up, evidence-based alternatives, unresolved items/escalations, detailed commitments and controls |
| Product specification | Numbered requirements with criteria, scope, interfaces, owners, release gates | Functional/nonfunctional groupings, tolerances and edge cases, methods, dependency mapping, traceability | Detailed interface/configuration constraints, failure cases, verification matrix, change control, unresolved risks and release readiness criteria |
| Quarterly report | Results, operating drivers, risks, explicitly labeled outlook | Available prior-period comparisons, product/customer/cost breakdowns, quantitative driver bridges, mitigation progress | Multiple reconciled operating cuts, variance/root-cause analysis, cash/working-capital detail where grounded, dependencies and supported outlook alternatives |
| Annual report | Quarter-reconciled annual results, operating story, year-end organization, risks | Business/product performance, grounded strategy outcomes, capital/resource allocation, trend tables and year-specific risk discussion | Detailed multi-year analysis where available, unit economics/operating drivers, governance and execution controls, forward dependencies and substantial appendix |
| HR policy | Applicability, eligibility, obligations, exceptions, owner/effective date, administration | Step-by-step process, responsibilities, time limits, records, escalation, grounded examples | Exception paths, tricky scenarios, cross-policy dependencies, review/change controls, FAQ and administrator guidance |
| Board deck | Executive overview, performance, risks, explicit decisions requested | Actual-versus-plan where canonically available, quantified drivers, strategic options, resource tradeoffs, milestone accountability | Decision papers and alternatives, grounded scenario analysis, dependencies, trigger thresholds, detailed financial/operating appendix |
| Financial workbook | Income bridge, margins/unit economics, period/reconciliation checks, source and methodology notes | Six distinct modules including available-period trends, product/channel or cost detail, operating drivers, liquidity/funding or another relevant grounded analysis | Nine distinct modules using richer source granularity, variance bridges, working-capital/funding detail, sensitivities/scenarios only if canon supports them, explicit model limitations and controls |

For a medium financial workbook, useful modules might be income reconciliation,
margin analysis, quarterly trends, revenue mix, cost drivers, and liquidity.
For deep, add a relevant unit-economics model, working-capital bridge, and
approved-assumption sensitivity. These are examples, not permission to invent
cash balances, transactions, targets, or forecast assumptions. Choose modules
that answer the company's actual questions and reconcile their inputs.

Actual breakdowns must reconcile to canonical totals; point-in-time balances
are not quarterly flows. A liquidity snapshot is not a cash-flow model.
Comparisons use only periods knowable by publication. Forecasts, budgets,
scenarios, and illustrative assumptions must have explicit canonical status,
units, approval/availability dates, and sensitivity boundaries. Do not silently
derive a full balance sheet or financing model from revenue and operating income.

### Source readiness and depth review

Before freezing, identify missing sources for each planned genre and depth.
Populate structured supporting details, quoted positions, numbered requirements,
operating breakdowns, actual-versus-plan inputs, or approved assumptions in the
canon as needed. Do not merely lengthen prose around a thin baseline.
Preserve public/internal boundaries and all known/effective dates.

Review each pilot's actual content against the chosen targets before bulk
generation. Every final document requires a separate `depth_review` receipt
with the selected depth, actual measurements, and specific coverage evidence.
Reviewers must reject padding, repeated historical summaries, token charts,
unnavigable prose-filled cells, empty tabs, unsupported numbers, or generic
detail that could belong to any company. Meeting words must be genuine dialogue;
requirement counts must be distinct controls; workbook modules must be usable.
The helper checks metadata and receipt presence, not these substantive judgments.

### Press releases

Use a newsworthy headline and optional factual subheadline. Lead with a dated
city/company announcement covering who, what, when, and why it matters. Follow
with substantive details and scope, an approved quote attributed to a named
person in their effective role, background, an "About <company>" boilerplate,
and a media contact. A quote is a paragraph in the news narrative, not a heading
that says "attributed_quote". Do not open with "This document describes..." or
list canonical facts as a report.

Record the dateline location, news rationale, quote wording/speaker/approval,
boilerplate, and contact details in the canon before writing. Use invented
contact details, not real people's addresses; email/URLs may use reserved
example domains. Do not fabricate product benefits, transaction close status,
customer endorsements, or regulatory claims. If the genre needs facts absent
from the baseline, enrich the canon rather than emitting placeholders.

### Product presentations

Structure a persuasive but grounded story for a specified audience: the
customer's problem, impact, the product's value, capabilities, an example
workflow or use case, adoption/deployment considerations, and a next step.
Use concise slide text and explanatory speaker notes when appropriate. Use
charts, screenshots/mockups, capability cards, or process visuals that depict
canonical product behavior; do not turn every slide into a wall of prose.
Ground capabilities, benefits, personas, deployment constraints, and evidence
in version-specific canon records. Do not invent performance benchmarks or
customer results. Clearly distinguish available capabilities from roadmap plans
using normal product terminology, without generation disclaimers.

### Meetings, technical documents, and policies

Meeting transcripts contain believable back-and-forth: questions, responses,
clarifications, tradeoffs, and explicit resolutions. Ground speaker positions,
decisions, assignments, and deadlines in a meeting record. Each participant
uses a role effective on the meeting date. A list of minutes is not a transcript.

Specifications contain numbered requirements, interfaces/constraints,
dependencies, out-of-scope items, and measurable acceptance criteria.
Use traceable requirement records; a launch milestone is not a specification.
XLSX specs are requirements/traceability matrices with usable columns, not a
paragraph spread across arbitrary cells.

Policies use direct, consistent employee-facing language with applicability,
eligibility, obligations, exceptions, administration, ownership, and effective
dates. Do not print "fictional policy" or generation caveats. Do not invent legal
compliance guarantees. Administrative obligations must be recorded in the canon.

### Financial reports, board decks, and workbooks

Reports explain actual financial performance with explicit periods and units,
substantive operating context, and canonically recorded outlook/risks. Annual
reports reconcile quarters and distinguish year-end organization from later
publication-date changes. Do not invent missing comparative periods.

Board decks emphasize decisions, operating performance, status, and risk in an
executive visual style. Product sales narratives and board narratives are not
interchangeable. Requested decisions and supporting rationale need canon records.

Workbooks contain meaningful source, calculation, and notes/assumptions sheets
with navigable layouts. Formulas derive totals/margins from source cells and
remain updateable. Put methodology and normal business-source notes in the
workbook; raw generation claims/hashes belong in `build`. Evaluate/save with
OfficeCLI and reopen every formula cache against independently grounded
expected values, not just formula syntax or a calculation-on-open flag.

## Planning behavior

The planner uses financial records for reporting anchors, milestone/acquisition
records for press releases, milestones for specs/product presentations, policy revisions for HR, and
financial/milestone/acquisition records for board decks and meetings.
These anchors identify the topic; they do not alone provide enough content for
a full meeting transcript or spec. Supply structured extension records before
drafting. Include related visible facts, without reciting every historical fact.

Annual and quarterly documents are dated on the applicable financial record's
availability date. Other items use their anchor's knowability date or a
deterministic follow-up date. A document's author can know facts up to publication,
but year-end org descriptions use the role state at year end, not future hires.

## Additional user-selected types

Do not silently discard additional types or substitute a generic report.
Define them in the request's `additional_types` object, with supported formats,
canonical anchor kinds, audience, and `required_sections`, then include their
IDs in `document_types`. Add the required source records to the canon before
planning. Extension anchors with unfamiliar kinds need a human-readable
`data.title`. No installed-skill code edit is required.

Before drafting an additional type, define its purpose, reader, voice, normal
structure, expected depth, source-fact needs, and editorial acceptance criteria
in `build\type-profiles.json`. Examples include integration plans, sales
enablement, incident postmortems, onboarding packs, and customer support reports.
Validate business-specific invariants manually or with a reviewed build helper.
Never introduce financial totals, entities, or decisions in a type template.

## QA beyond automated checks

Compare people/titles, numbers/units, dates, product versions, policy revisions,
and term definitions across documents for the same period. Check tense against
both known and effective dates. Review M&A states and public/internal audience
boundaries. Trace every substantive assertion, including speaker dialogue, to
a visible canonical record.

Ground named people, speakers, owners, approvers, departments, and reporting
relationships in the context's date-filtered organization view. Ground business
facts and terminology in canon records. This single core info set precedes prose;
do not invent a person/title in a draft or expand the org chart to cover unnamed
employees merely because a financial report gives total headcount.

Perform a `genre_review` for every document selected for content review: does this read as a real instance
of its requested type, for the correct audience, rather than a generic summary,
template, outline, or data dump? Separately verify selected depth against the
genre's targets and substantive coverage, using actual delivered measurements.
Verify no generation/provenance self-labeling
appears anywhere, including notes, comments, hidden worksheet content, or Office
properties. Required section roles and automated phrase scans are structural
guards, not proof of convincing prose or exhaustive coverage.

Apply `Sampled`, `FullContentSampledVisual` (default), or `Full` independently of
content depth. All files retain structural, claim/date/access, measurable-depth,
issue/schema, hash, cache, provenance and coverage checks. Selected content
reviews cover the complete document. Selected visual reviews use OfficeCLI's
HTML export captured through the pinned Playwright-Chromium adapter for every
page/slide or relevant worksheet region; unselected gates are
`not_selected`. Always review pilots, condition/source-history carriers and
flagged exceptions. Failed samples require affected-family expansion.
Review XLSX column widths, readability, chart ranges, supported chart objects
(title/legend/axis/series pixels, not just that an SVG exists), formula
caches, and error cells. Structural chart-object capture (correct identity,
unclipped/unoccluded bounds) does not by itself certify every visual/semantic
detail; it is reviewed like any other selected visual under the active QA
policy. Inspect representative
cross-format documents about the same event together. A clean ZIP/XML audit is
not proof of coherent prose, readable layout, or calculated formulas.
Use OfficeCLI for document authoring/checking/HTML export, and the pinned
Playwright-Chromium adapter for capture; never OfficeCLI's own
`view screenshot --render` path, native/auto rendering, `refresh`, or native
page counts. Hash/index final saved bytes. No Microsoft Office dependency
remains; HTML previews do not certify native layout fidelity. An unavailable
required OfficeCLI/Playwright-capture capability blocks.
Do not count missing/undisplayed/unreadable images as reviewed or repeat full
maintenance checks for unchanged batches. See `dependencies.md` for setup,
supported capabilities, bounded retry and review-capacity planning.
