# Fictional Company Generator

An AI skill that creates realistic documents for a fictional company over several
years. It produces Word (DOCX), PowerPoint (PPTX), and Excel (XLSX) files that share
the same company history, financial results, people, and terminology.

**Key capabilities:** choose the document types and **light/medium/deep** level
of detail, show different viewpoints, and track who said what as sources change.
Optional quality assurance (QA) checks can trace statements back to their sources,
report how much of the content has been checked, and show which files a fictional
reader would be allowed to see. Missing information, disagreements, unrelated
material, and outdated sources are included **only if you ask for them**.
By default, the content is clean and consistent.

The skill uses [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) to work with
documents and a fixed version of Playwright with Chromium to capture their
appearance. Microsoft Office is not required.

## Capabilities

| Capability | What it provides |
|---|---|
| Coherent company history | A shared, versioned record of company results, financial totals, products, acquisitions, roles and when they take effect, and terminology |
| Time-aware writing | Each document uses only what its author could know on its publication date and distinguishes announced changes from completed ones |
| Authentic Office content | Writing suited to each document type, realistic filenames, presentations, Excel files with saved and independently checked formula results, and planned Word page breaks |
| Configurable scope and depth | A fixed document count, your choice of years, types, and formats, and light/medium/deep detail; you can request new document types without editing the skill |
| Perspectives and source history | Clear attribution for facts, opinions, assumptions, risks, recommendations, and decisions, with a history of corrected, withdrawn, and replaced sources |
| Optional corpus imperfections | Unknown or undocumented facts, missing sources, disagreements, related but irrelevant material, and outdated sources, while keeping the shared company results unchanged |
| Mandatory generation checks | Basic checks on every file, full reviews of sample files for each layout, wider reviews when fixes are needed, and file hashes that tie completion and resume records to the files checked |
| Opt-in additional QA | Content and appearance reviews, records of where statements appear in final files, file hashes, source-tracing and coverage reports, and checked file lists for simulated reader access |

The skill tracks what each author knows, what they can access, and what the
intended reader can access as separate rules. These rules are fictional: they do
not set real permissions or enforce access control lists (ACLs).

## Install

The skill can be used as-is from a clone of this repo.

To use it elsewhere, copy that **entire folder**, including `references` and
`scripts`, to one of these locations:

| Scope | Destination |
|---|---|
| Another project | `<project>/.github/skills/company-content-generator` |
| All your Copilot CLI sessions | `$HOME/.copilot/skills/company-content-generator` |

**Requirements:**

- [Python](https://www.python.org/downloads/) 3.10 or later.
- An [OfficeCLI](https://github.com/iOfficeAI/OfficeCLI) installation you have reviewed.
- [Playwright](https://playwright.dev/python/) 1.63.0 with the matching managed
  [Chromium browser](https://playwright.dev/python/docs/browsers) capture adapter.
- A [Python virtual environment](https://docs.python.org/3/library/venv.html) is
  recommended to keep dependencies separate from other projects.

Required tools are not bundled. The skill can help install or update them with
your approval.

## Use

Ask Copilot to use **company-content-generator** and describe the company and
content you want. For normal use, you do not need a request file or hand-written
JSON. The skill asks for any missing required details before it starts.
Choose a new output directory.

### Options you can specify

| Option | What to provide |
|---|---|
| Company | Company name and industry |
| Period | Start year and number of years, or a date range in the past |
| Volume | Approximate number of final documents |
| Content types | Press releases, product presentations, meeting transcripts, product specs, quarterly/annual reports, HR policies, board decks, financial workbooks, or additional types you describe |
| Formats | DOCX, PPTX, XLSX, or just the formats that fit your document types |
| Depth | **Light**, **medium** (default), or **deep** |
| Additional QA | The skill asks you to choose **No additional QA** (default), **Sampled**, **Balanced**, or **Full**; this is separate from depth |
| Runtime budget | An optional budget you set for the run, enforced at safe stopping points |
| Output | The path to a new output directory |
| Grounding | Existing shared company data, or details about products, finances, acquisitions, and organizational changes |
| Seed | An optional integer that makes the structured data repeatable; defaults to `42` |
| Corpus controls | Optional viewpoints, missing information, disagreements, irrelevant or outdated material, and rules for what fictional authors and readers can know or see; separate from depth |

### Sample prompts

**Semiconductor company corpus (what the included example covers)**

> Use company-content-generator to create 50 documents for SunnySemi LLC,
> a fictional, wildly successful semiconductor company based in Florida,
> covering 2020 through 2025. Include press releases, product presentations,
> meeting transcripts, product specifications, board decks, and financial
> workbooks. Use DOCX, PPTX, and XLSX, medium depth, and seed 42. Generate
> the output in /company-content/sunnysemi-2020-2025. No additional QA.

**Focused content**

> Use company-content-generator to create 8 DOCX press releases for Northstar
> Marine, a struggling, fictional sailboat manufacturer, covering 2023 through 2024.
> Use light depth and focus on product launches and company announcements.
> No additional QA.

**Custom types and richer analysis**

> Use company-content-generator to create 20 documents for Cedarline Robotics,
> a fictional warehouse automation company, covering 2021 through 2025.
> Include acquisition integration plans, customer onboarding guides, board
> presentations, and financial workbooks. Use deep depth and seed 17.
> Record the company history, acquisition milestones, product details, and
> financial breakdowns in the shared company record before drafting. Generate
> the output in C:\CompanyContent\Cedarline-2021-2025. Run Balanced additional QA.

**Viewpoints, changing sources, and intentional disagreement**

> Use company-content-generator to create 18 documents for GulfLight Chips,
> a fictional Florida semiconductor designer, covering 2023 through 2025.
> Include meeting transcripts, product specifications, and board decks at
> medium depth. Include engineering and finance viewpoints on launch
> readiness, clearly naming who disagrees about schedule
> assumptions and a later source that corrects an earlier estimate.
> Keep actual financial results consistent and separate assumptions
> from approved decisions. Keep the earlier sources, but do not give
> their authors knowledge of later events. Run Sampled additional QA.

**Explicit gaps and stale material**

> Use company-content-generator to create 12 medium-depth documents for
> Pelican Devices, a fictional sensor manufacturer, covering 2024 through
> 2025. Include meeting transcripts, product presentations, and board decks.
> Deliberately leave one supplier qualification result unknown to the
> authors, omit one otherwise known supporting source from the collection,
> and include an older planning deck that is out of date as of 2025-12-31
> compared with a newer approved source. Keep the total at 12 documents,
> leave the shared company results unchanged, and do not add labels explaining
> these intentional gaps or outdated information in the documents themselves.
> Generate the output in ~/repos/pelican-2024-2025.
> Run Full additional QA.

**Additional QA and a simulated public view of an existing corpus**

> Use company-content-generator to run Full additional QA on the existing
> document collection at examples/semiconductor-company without regenerating
> accepted documents. After QA passes its strict audit, create a file list
> under build/views showing what a fictional public reader could see as of
> 2025-12-31. This list must not enforce real access permissions or expose
> the private company record and build files to that reader.

## How generation works

1. **Build the canon.** Create and check the shared company record, called the
   canon. It holds the facts, finances, events, people, and terminology used by
   all documents.
2. **Set depth and plan.** Depth level (light/medium/deep) controls how much substance
   each document covers, not just its length, and the supporting detail required
   in the canon. Shared company facts remain consistent across depths; deeper
   content may require enriching the canon before drafting. Add the needed
   detail, then finalize a plan with dates, realistic titles, and targets for
   each document.
3. **Draft and export.** Use only facts available on each document's date.
   Check statements, then create the files with the shared OfficeCLI runner.
   It checks all-or-nothing batches and template merges, verifies saved formula
   results, captures sample HTML layouts, and records links between statements
   and exported content. File hashes tie checkpoints to the files they describe.
   Documents can share layouts, so each layout's sample review can be reused.
   Bulk generation uses up to four workers, limited by available CPUs and
   eligible documents, and reuses one browser for captures. Use `--workers 1`
   to process documents one at a time.
4. **Complete generation.** Fully review a sample of each layout before using
   it for the rest of the collection, and run basic checks on every file.
   Save a file list and generation receipt. These records do not claim that
   all documents were reviewed or all statements traced to sources.
   If a problem occurs, allow one retry with a specific fix, then stop and
   report the evidence rather than building a new document engine mid-run.
5. **Run additional QA only if requested.** The skill asks whether you want
   this separate pass; the default is no additional QA. Sampled reviews selected
   documents in full. Balanced reviews all content and a sample of appearances.
   Full reviews all content and appearances. Required review selections still
   apply, and problems expand the review to other files using the affected
   layout. The pass checks recorded links between statements and exported
   content, verifies source-tracing and coverage records against the final
   files, takes any missing selected captures, and collects completed reviews
   before an independent final audit. You can also request it later without
   regenerating files.

Generation and additional QA are marked complete separately. Without additional
QA, files still receive the required checks and layout sample reviews. That does
not mean every document's writing and appearance have been reviewed or every
statement in the final files has been traced to its source. Automated file
checks do not replace those reviews.

The skill uses the runner's `phase` command to record time spent by the agent
on planning, drafting, and reviewing. Timing summaries separate elapsed time,
work done in parallel, capture activity, and gaps between commands.

## Included example

[`examples/semiconductor-company`](examples/semiconductor-company) contains
**SunnySemi LLC**, a fictional, wildly successful Florida semiconductor company
headquartered in Tampa with an engineering site in Sarasota. Its history follows
SunCore inference ASICs, GulfSwitch power controllers, an acquisition, leadership
changes, and rapid financial growth during fiscal years **2020-2025**.

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

The latest publication date is **2026-01-15**, for the FY2025 financial workbook.

`canon.json` holds the shared company facts, and `documents/` contains the final
Office files. `build/` keeps the request, approved plan, drafts, generation
scripts, links between statements and exported content, reviews, and automated
check results. Its `manifest.json` and `generation.json` record completed
generation using schema version 2.

In this run, the required sample and follow-up reviews covered the content and
appearance of all 50 documents. **Additional corpus QA was not requested**.
