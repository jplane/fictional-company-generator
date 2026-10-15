# Fictional Company Content Generator

An AI skill for creating realistic, consistent documents for a fictional company
over time. It produces DOCX, PPTX, and XLSX content grounded in a shared company
history, financials, organization, and glossary.

**Key capabilities:** shared company history and date-grounded Office documents;
configurable genres and **light/medium/deep** detail; attributed statements,
viewpoints, and evolving sources; portable claim provenance, coverage reporting,
and simulated access views through optional additional QA. Gaps, conflicts, noise,
and stale sources are **opt-in**; clean, consistent content remains the default.
Generation and QA use
[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) for document operations,
with pinned Playwright-Chromium captures and no Microsoft Office dependency.

## Capabilities

| Capability | What it provides |
|---|---|
| Coherent company history | One versioned canon for business actuals, financial reconciliation, products, acquisitions, effective-dated roles, and terminology |
| Time-aware writing | Publication-date knowledge and effective-date state keep future results, promotions, and announced-but-uncompleted events distinct |
| Authentic Office content | Genre-specific prose, realistic filenames, presentations, and workbooks with saved, independently checked formula results; deliberate DOCX page breaks |
| Configurable scope and depth | An exact planned document count, requested years/types/formats, light/medium/deep substance, and custom genres without editing the skill |
| Perspectives and source history | Attributed facts, opinions, assumptions, risks, recommendations, and decisions; source versions preserve corrections, retractions, and supersession |
| Optional corpus imperfections | Explicit unknowns, undocumented facts, omitted sources, conflicting assertions, adjacent noise, and stale sources without corrupting canonical actuals |
| Mandatory generation checks | Every-file mechanical checks, fully reviewed layout-family pilots, corrective review expansion, and hash-bound completion/resume evidence |
| Opt-in additional QA | Selected content/visual reviews, final-artifact claim locations and hashes, provenance and coverage reports, and audited simulated recipient-view manifests |

Author knowledge, author access, and intended reader access are modeled
separately. These are fictional visibility rules, not real service permissions
or ACL enforcement. The skill is a general-purpose content generator: it does
not require a decision question or produce an independent evaluation answer key.

## Install

In this repository, the skill is already project-scoped at
`.github/skills/company-content-generator`.

To use it elsewhere, copy that **entire folder**, including `references` and
`scripts`, to one of these locations:

| Scope | Destination |
|---|---|
| Another project | `<project>/.github/skills/company-content-generator` |
| Your Copilot CLI sessions | `$HOME/.copilot/skills/company-content-generator` |

Do not overwrite an existing installation without reviewing it. Start a new
Copilot session in the target project after installation.

**Requirements:** Python 3.10+, reviewed installed OfficeCLI, and the pinned
Playwright 1.63.0 + matching managed Chromium capture adapter (installed only
into the active virtual environment, never a system Python). Dependency setup
explicitly includes authorized OfficeCLI installation/update, official
versioned assets, SHA-256/version verification, scoped skill/MCP setup, and
post-update capability probes. No tools are installed automatically. Missing
required capabilities stop the workflow; native Office and alternative
document-engine fallbacks are not used.

See `.github/skills/company-content-generator/references/dependencies.md` for
setup/update, capability checks and limitations. These tools are not bundled.
OfficeCLI's HTML export plus Playwright-Chromium capture does not certify
native Office layout fidelity; unsupported fields/formulas/drawing objects
remain explicit blockers (standard column/line XLSX charts are a supported,
identity-bound capture path — see `dependencies.md`).

## Use

Ask Copilot to use **company-content-generator** and describe the company and
content you want. No sample request file or manually authored JSON is required
for normal use. The skill collects missing required details before generation.
Choose a new output directory; existing corpora are not silently overwritten
or migrated.

### Options you can specify

| Option | What to provide |
|---|---|
| Company | Company name and industry |
| Period | Start year and number of years, or a historical date range |
| Volume | Approximate number of final documents |
| Content types | Press releases, product presentations, meeting transcripts, product specs, quarterly/annual reports, HR policies, board decks, financial workbooks, or additional types you describe |
| Formats | DOCX, PPTX, XLSX, or a narrower selection appropriate to the content types |
| Depth | **Light**, **medium** (default), or **deep** |
| Additional QA | Asked explicitly during intake: **No additional QA** (default), **Sampled**, **Balanced**, or **Full**; independent of depth |
| Runtime budget | Optional user-supplied budget, enforced at safe checkpoints |
| Output | Explicit path to a new output directory |
| Grounding | Existing canonical company data or details about products, financials, acquisitions, and organizational changes |
| Seed | Optional integer for reproducible structured data; defaults to `42` |
| Corpus controls | Optional viewpoints, explicit gaps/conflicts/noise/staleness, and fictional visibility/knowledge boundaries; independent of depth |

### Sample prompts

**Semiconductor company corpus (the included example's scope)**

> Use company-content-generator to create 50 documents for SunnySemi LLC,
> a fictional, wildly successful semiconductor company based in Florida,
> covering 2020 through 2025. Include press releases, product presentations,
> meeting transcripts, product specifications, board decks, and financial
> workbooks. Use DOCX, PPTX, and XLSX, medium depth, and seed 42. Generate
> the output in /company-content/sunnysemi-2020-2025. No additional QA.

This describes the example's scope, not a byte-for-byte reproduction command.
Use a new output directory rather than overwriting the included corpus.

**Focused content**

> Use company-content-generator to create 8 DOCX press releases for Northstar
> Marine, a struggling, fictional sailboat manufacturer, covering 2023 through 2024.
> Use light depth and focus on product launches and company announcements.
> Generate the output in C:\CompanyContent\Northstar-press-releases.
> No additional QA.

**Custom types and richer analysis**

> Use company-content-generator to create 20 documents for Cedarline Robotics,
> a fictional warehouse automation company, covering 2021 through 2025.
> Include acquisition integration plans, customer onboarding guides, board
> presentations, and financial workbooks. Use deep depth and seed 17.
> Establish the company history, acquisition milestones, product details, and
> financial breakdowns in the canon before drafting. Generate the output in
> C:\CompanyContent\Cedarline-2021-2025. Run Balanced additional QA.

**Viewpoints, changing sources, and intentional disagreement**

> Use company-content-generator to create 18 documents for GulfLight Chips,
> a fictional Florida semiconductor designer, covering 2023 through 2025.
> Include meeting transcripts, product specifications, and board decks at
> medium depth. Include engineering and finance viewpoints on launch
> readiness, with an explicitly attributed disagreement about schedule
> assumptions and a later source that corrects an earlier estimate.
> Keep actual financial results consistent and distinguish assumptions
> from approved decisions. Preserve the source history without giving
> earlier authors future knowledge. Generate the output in
> C:\CompanyContent\GulfLight-source-history. Run Sampled additional QA.

**Explicit gaps and stale material**

> Use company-content-generator to create 12 medium-depth documents for
> Pelican Devices, a fictional sensor manufacturer, covering 2024 through
> 2025. Include meeting transcripts, product presentations, and board decks.
> Deliberately leave one supplier qualification result unknown to the
> authors, omit one otherwise known supporting source from the collection,
> and include an older planning deck that is stale as of 2025-12-31 relative
> to a newer approved source. Keep the total at 12 documents, preserve
> canonical actuals, and do not label these conditions in the business prose.
> Generate the output in C:\CompanyContent\Pelican-2024-2025.
> Run Full additional QA.

**Additional QA and a simulated public view of an existing corpus**

> Use company-content-generator to run Full additional QA on the existing
> corpus at examples/semiconductor-company without regenerating accepted
> documents. After QA and its strict audit pass, create a simulated public
> recipient-view manifest as of 2025-12-31 under build/views. Do not treat
> the manifest as real ACL enforcement or include the private canon/build
> tree as recipient-visible content.

## How generation works

1. **Build the canon.** Establish and validate shared company facts, financials,
   events, people, and terminology before writing documents.
2. **Set depth and plan.** Light, medium, or deep determines substantive coverage,
   not just length. Enrich the canon accordingly, then freeze a dated document
   plan with realistic titles and per-document targets.
3. **Draft and export.** Write from date-filtered facts so documents cannot
   reference future developments. Validate claims and execute the shared thin
   OfficeCLI runner: checked atomic batches/template merges, saved formula checks,
   mandatory pilot HTML captures, hash-bound checkpoints and exporter-recorded
   claim associations. Shared content-independent layouts amortize pilot work.
   Bulk authoring uses up to four workers bounded by CPUs/eligible documents,
   with one reusable capture browser and an explicit `--workers 1` serial override.
4. **Complete generation.** Fully review family pilots before bulk use and run
   mechanical checks on every file. Generation completion writes a manifest and
   generation receipt, not fabricated corpus reviews or full provenance.
   Allow one concrete corrective retry, then stop with evidence instead of
   developing a new backend in-flight.
5. **Run additional QA only if requested.** Intake explicitly asks whether to run
   this separate pass after generation, defaulting to no additional QA. Sampled
   reviews selected complete documents; Balanced reviews all
   content and sampled visuals; Full reviews all content and visuals. Mandatory
   selections and affected-family expansion remain. The pass resolves recorded
   exporter associations, verifies final-file provenance/coverage, captures missing
   selected visuals and collects actual reviews before an independent final audit.
   It can also be explicitly invoked later without regenerating files.

Generation and QA have distinct completion states. No additional QA means files
are generated with mandatory checks and pilot reviews, not that every document
has been editorially/visually reviewed or its final-artifact provenance certified.
Existing corpora retain their original strict rules; no upgrade/migration is
performed. Automated structural checks do not replace editorial/visual review.

For a new-contract corpus, the runner supports:

**Windows PowerShell**

```powershell
python .github\skills\company-content-generator\scripts\officecli_runner.py qa --root "C:\CompanyContent\SunnySemi-2020-2025" --jobs build\jobs.json --policy Sampled
# Complete the required actual reviews under build\reviews, then:
python .github\skills\company-content-generator\scripts\officecli_runner.py finalize-qa --root "C:\CompanyContent\SunnySemi-2020-2025"
python .github\skills\company-content-generator\scripts\officecli_runner.py summary --root "C:\CompanyContent\SunnySemi-2020-2025"
```

**macOS/Linux (bash/zsh; active virtual environment)**

```sh
outputRoot="$HOME/CompanyContent/SunnySemi-2020-2025"
python .github/skills/company-content-generator/scripts/officecli_runner.py qa --root "$outputRoot" --jobs build/jobs.json --policy Sampled
# Complete the required actual reviews under build/reviews, then:
python .github/skills/company-content-generator/scripts/officecli_runner.py finalize-qa --root "$outputRoot"
python .github/skills/company-content-generator/scripts/officecli_runner.py summary --root "$outputRoot"
```

The skill records agent-owned planning/drafting/review spans using the runner's
`phase` command. Timing summaries distinguish wall time, overlapping worker time,
capture activity and between-invocation gaps; missing agent timing is unknown,
not zero. No throughput guarantee is inferred from synthetic tests.

## Included example

[`examples/semiconductor-company`](examples/semiconductor-company) showcases
**SunnySemi LLC**, a fictional, wildly successful Florida semiconductor company
headquartered in Tampa with an engineering site in Sarasota. Its history follows
SunCore inference ASICs, GulfSwitch power controllers, an acquisition, leadership
changes, and rapid financial growth over fiscal years **2020-2025**.

The example contains **50 medium-depth documents**:

| Content type | Count | Format |
|---|---|---|
| Press releases | 12 | DOCX |
| Product specifications | 6 | DOCX |
| Meeting transcripts | 6 | DOCX |
| Product presentations | 8 | PPTX |
| Board decks | 6 | PPTX |
| Financial workbooks | 12 | XLSX |
| **Total** | **50** | **24 DOCX, 14 PPTX, 12 XLSX** |

Publication dates extend through **2026-01-15** for the FY2025 financial
workbook.

`canon.json` holds the shared company facts; `documents/` contains the final
Office files. `build/` preserves the request, accepted plan, drafts, generation
scripts, exporter bindings, reviews, and machine evidence. Its `manifest.json`
and `generation.json` record completed schema-version-2 generation.
All 50 documents received mandatory pilot/corrective content and visual reviews
in this run. **Additional corpus QA was not requested**: the example does not
claim certified final-artifact provenance, complete coverage, or native Office
layout fidelity. Regenerable logs, contexts, intermediate plan variants, and
renders are excluded from Git.

The example is for exploration, not a required skill dependency or a default
configuration for new runs.

## Detailed guidance

The skill's `SKILL.md` contains the generation workflow and output conventions.
Its `references/content-types.md` covers genres and depth; `references/contracts.md`
covers structured inputs and helper usage. `references/provenance.md` explains
portable artifact locations and storage bindings; `references/use-case-layering.md`
shows how customer-specific decision/evaluation features can be layered above
the general-purpose generator without implementing them in the core.
Those files are the authoritative
instructions rather than duplicated documentation here.

## Maintenance

`python -m unittest discover -s tests -q` runs the standard-library tests from
an active virtual environment without installing OfficeCLI or Playwright.
The `Cross-platform Python tests` GitHub Actions workflow runs the same command
on Windows, macOS, and Ubuntu with Python 3.10 and 3.14 for pushes, pull requests,
and manual dispatch. Its optional real-tool integrations are deliberately
skipped: a green Python matrix does not certify document rendering.
See `references/dependencies.md` inside the skill for opt-in real-tool checks.
