# Fictional Company Content Generator

An AI skill for creating realistic, consistent documents for a fictional company
over time. It produces DOCX, PPTX, and XLSX content grounded in a shared company
history, financials, organization, and glossary.

**Key capabilities:** shared company history and date-grounded Office documents;
configurable genres and **light/medium/deep** detail; attributed statements,
viewpoints, and evolving sources; portable claim provenance, coverage reporting,
and simulated access views. Gaps, conflicts, noise, and stale sources are
**opt-in**; clean, consistent content remains the default. Generation and QA use
[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) for document operations,
with no Microsoft Office dependency.

## Install

In this repository, the skill is already project-scoped at
`.github\skills\company-content-generator`.

To use it elsewhere, copy that **entire folder**, including `references` and
`scripts`, to one of these locations:

| Scope | Destination |
|---|---|
| Another project | `<project>\.github\skills\company-content-generator` |
| Your Copilot CLI sessions | `$HOME\.copilot\skills\company-content-generator` |

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

See `.github\skills\company-content-generator\references\dependencies.md` for
setup/update, capability checks and limitations. These tools are not bundled.
OfficeCLI's HTML export plus Playwright-Chromium capture does not certify
native Office layout fidelity; unsupported fields/formulas/drawing coverage
remain explicit blockers.

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
| QA policy | **Sampled**, **FullContentSampledVisual** (default), or **Full**; independent of depth |
| Runtime budget | Optional user-supplied budget, enforced at safe checkpoints |
| Output | Explicit path to a new output directory |
| Grounding | Existing canonical company data or details about products, financials, acquisitions, and organizational changes |
| Seed | Optional integer for reproducible structured data; defaults to `42` |
| Corpus controls | Optional viewpoints, explicit gaps/conflicts/noise/staleness, and fictional visibility/knowledge boundaries; independent of depth |

### Sample prompts

**Mixed company corpus**

> Use company-content-generator to create about 45 documents for Lumenvale
> Systems, a fictional industrial analytics software company, covering 2022
> through 2024. Include press releases, product presentations, meeting
> transcripts, product specs, quarterly and annual reports, HR policies,
> and financial workbooks. Use DOCX, PPTX, and XLSX with medium depth.
> Generate the output in C:\CompanyContent\Lumenvale-2022-2024.

**Focused content**

> Use company-content-generator to create 8 DOCX press releases for Northstar
> Marine, a struggling, fictional sailboat manufacturer, covering 2023 through 2024.
> Use light depth and focus on product launches and company announcements.
> Generate the output in C:\CompanyContent\Northstar-press-releases.

**Custom types and richer analysis**

> Use company-content-generator to create 20 documents for Cedarline Robotics,
> a fictional warehouse automation company, covering 2021 through 2025.
> Include acquisition integration plans, customer onboarding guides, board
> presentations, and financial workbooks. Use deep depth and seed 17.
> Ground the content in company history and financial data previously outlined in this chat.

## How generation works

1. **Build the canon.** Establish and validate shared company facts, financials,
   events, people, and terminology before writing documents.
2. **Set depth and plan.** Light, medium, or deep determines substantive coverage,
   not just length. Enrich the canon accordingly, then freeze a dated document
   plan with realistic titles and per-document targets.
3. **Draft and export.** Write from date-filtered facts so documents cannot
   reference future developments. Validate claims and execute the shared thin
   OfficeCLI runner: checked atomic batches/template merges, saved formula checks,
   forced HTML renders, hash-bound checkpoints and exact exporter associations.
4. **Review under the selected policy.** Fully review family pilots before bulk
   use. Sampled reviews selected complete documents; the default reviews all
   content with sampled visuals; Full reviews all content and visuals. Deterministic
   family samples always include carriers/exceptions. A failed sample blocks its
   family and expands review. All policies check every file automatically.
   Allow one concrete corrective retry, then stop with evidence instead of
   developing a new backend in-flight.
5. **Audit delivery.** Verify final files, hashes, and review records. Keep
   provenance in `build\`; automated structural checks do not replace editorial
   and visual review.

## Included example

This is a historical **prior-contract** example, not a template for new runs or
a corpus validated by the current helper. New runs use schema version 2; there
is no automatic migration or version-1 validation.

`examples\asteroid-mining` showcases **AstroDig LLC**, a fictional aerospace
engineering and asteroid resource technology company. It covers fiscal years
2013-2026 with 150 medium-depth documents: 50 each in DOCX, PPTX, and XLSX,
including reports, press releases, presentations, specs, and meeting transcripts.

`canon.json` holds the shared company facts; `documents\` contains the final
Office files. The curated `build\` folder preserves the request, accepted plan,
drafts, generation scripts, and QA/provenance records for inspecting how the
example was produced. Regenerable logs, contexts, intermediate plans, and renders
are excluded from Git. The example is for exploration, not a required skill
dependency or a default configuration for new runs.

## Detailed guidance

The skill's `SKILL.md` contains the generation workflow and output conventions.
Its `references\content-types.md` covers genres and depth; `references\contracts.md`
covers structured inputs and helper usage. `references\provenance.md` explains
portable artifact locations and storage bindings; `references\use-case-layering.md`
shows how customer-specific decision/evaluation features can be layered above
the general-purpose generator without implementing them in the core.
Those files are the authoritative
instructions rather than duplicated documentation here.
