# Fictional Company Content Generator

An AI skill for creating realistic, consistent documents for a fictional company
over time. It produces DOCX, PPTX, and XLSX content grounded in a shared company
history, financials, organization, and glossary.

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

**Requirements:** Python 3.10+ and one selected document backend:
locally installed Microsoft Office applications accessed through the repository's
COM automation, or Docker Engine with Docker Compose v2 for the pinned
LibreOffice container. Each new run checks selected-format native capabilities,
recommends Native Office when all probes pass and the container otherwise, then
asks you to choose. The selected backend is recorded for the run and is never
silently changed after a failure.

See `.github\skills\company-content-generator\references\dependencies.md` for
backend selection, setup, and prerequisite checks. LibreOffice is not
Microsoft Office and does not establish Microsoft Office rendering or
calculation fidelity.

`prototypes\libreoffice-backend` contains the selectable LibreOffice container
backend. It consumes the skill's prepared inputs and exports DOCX, XLSX, and
PPTX through the pinned LibreOffice toolchain. LibreOffice is not Microsoft
Office and does not establish Office fidelity; see that directory's README for
setup, contract limits, and testing.

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
| Backend | Native Office or the LibreOffice container; the skill recommends one after checking local capabilities |
| Depth | **Light**, **medium** (default), or **deep** |
| Output | Explicit path to a new output directory |
| Grounding | Existing canonical company data or details about products, financials, acquisitions, and organizational changes |
| Seed | Optional integer for reproducible structured data; defaults to `42` |

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
   reference future developments. Validate claims and export through the
   selected native Office applications or the pinned LibreOffice container.
4. **Review, fix, and repeat.** Verify pilots before bulk generation. Check
   factual consistency, genre, depth, and rendered layouts; recalculate Excel
   formulas. Correct issues and repeat affected QA until the checks pass.
5. **Audit delivery.** Verify final files, hashes, and review records. Keep
   provenance in `build\`; automated structural checks do not replace editorial
   and visual review.

## Included example

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
covers structured inputs and helper usage. Those files are the authoritative
instructions rather than duplicated documentation here.
