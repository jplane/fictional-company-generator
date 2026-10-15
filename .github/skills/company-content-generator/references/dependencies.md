# OfficeCLI dependency, setup, and capability boundaries

## Required components

Use **[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI)** as the sole document
engine for authoring, extraction, schema/issues checks, supported formula
evaluation, and static HTML export. Microsoft Office is **not required**.
Python 3.10+ performs business/claim validation, read-only package inspection,
hashing, provenance, coverage, and orchestration; it is not a second authoring
or calculation engine. Screenshot pixels never come from OfficeCLI: a thin
pinned **Playwright-for-Python 1.63.0** adapter
(`scripts\playwright_capture.py`) drives a matching **managed Chromium** build
to capture OfficeCLI's own HTML export. OfficeCLI does not bundle that browser
and no longer performs the screenshotting itself. Probe actual image
production; an executable/version check is not sufficient.

Prefer the installed `officecli` skill and runtime help. Resolve conflicting
companion instructions in favor of this pipeline; do not modify personal skills.
No LibreOffice, OpenOffice, cloud Office, COM adapter, standalone OOXML authoring
library, PDF/export plugin, or native rendering/calculation fallback is used.
OfficeCLI's editable charts and shapes are allowed when their required
capabilities and visual coverage have actually been proven.

## Authorized install/update setup

Dependency setup includes **installation and updates**, not just detection:

1. Locate `officecli`, inspect `--version`, the installed skill, source/license,
   and required runtime help. Check the official
   [release source](https://github.com/iOfficeAI/OfficeCLI/releases) once during
   setup. Install if absent; update if required capabilities are missing or the
   user requests an update. Offer compatible available updates without forcing
   a working installation to change. Disclose an unavailable release lookup;
   never infer "up to date" from a failed lookup.
2. Obtain authorization if not already covered. State the target/version,
   destination, and proposed PATH, skill, or MCP changes. Binary replacement
   does not authorize overwriting personal skill customizations or registering
   every detected client. Loading this skill is not installation approval.
3. Obtain the official **versioned release asset** for the host OS/architecture
   and its matching published SHA-256. Verify the downloaded bytes **before
   execution**, then verify `--version`. Missing/mismatched checksum or version
   blocks required setup. Do not pipe a downloaded installer to
   `Invoke-Expression`, invent an `update` command, or accept an unpinned latest
   download as verified. Preserve the prior usable binary until the replacement
   verifies; do not silently fall back after required update failure.
4. Configure the selected agent's OfficeCLI skill separately. Consult
   `officecli install --help` for the installed release before approved
   selected-agent setup. Bare `officecli install` performs broad binary/skill/MCP
   setup and is **not** an unscoped default. MCP registration is optional and
   separately authorized; the checked CLI runner does not require it.
5. Record source/version/checksum, setup actions and blockers. Reprobe required
   capabilities after installation/update. Pin the executable version and SHA-256
   for generation. Never update silently in the document loop. The runner's
   `setup_action` and `verify_release` helpers support decisions and verification;
   they deliberately do not download, install, or change PATH.

Review exact external revisions and licenses before enabling them. Keep content
local; do not upload drafts/canon or retrieve real workplace data. Additional
acquisition requires approval. Shared tools may remain outside the corpus;
run-specific inputs, logs, scripts, and dependencies belong under `build`.

## Playwright capture adapter setup

The adapter is a thin, **optional** dependency: importing
`scripts\playwright_capture.py` and running the default unit suite never
require Playwright to be installed (the import is lazy, inside
`CaptureEngine.__init__`). It is only required to actually capture pixels
(preflight probes and real renders).

1. If an existing testing/session venv already has the pinned Playwright
   dependency and its matching managed Chromium build working, **no
   additional install is required** — reuse it as-is.
2. Otherwise, install the pinned `playwright==1.63.0` package with `pip`
   scoped to the active virtual environment (never into a global/system
   interpreter), then run `playwright install chromium` to fetch the matching
   managed build from the official Playwright CDN. Do not hardcode the
   current user's name, session ID, or any machine-specific path into this
   repository; `PLAYWRIGHT_BROWSERS_PATH` (if set at all) is a per-process
   environment variable for a run, never a committed/global setting.
3. Honor whatever package index/feed is already configured for the active
   Python environment (`pip config`, `PIP_INDEX_URL`, a `pip.ini`/`pip.conf`,
   etc.). Only when a genuinely missing dependency actually requires a fresh
   install and no feed is otherwise configured does the current user's
   preferred feed, `https://packagefeedproxy.microsoft.io/pypi/simple/`,
   apply — this is a personal preference for this consumer, not a value to
   force on every generic consumer of this skill.
4. Record the resulting identity in the capability receipt: Playwright
   version, browser channel, managed Chromium version, its executable path,
   and its SHA-256 — the same evidentiary treatment as the OfficeCLI
   executable/version/checksum above. Treat drift in either identity as a
   capability change requiring revalidation, not a silent pass-through.

## Preflight, once per capability identity


**Before creating the output directory, canon, plan, or drafts**, verify Python
and perform tiny session-owned probes for the selected formats with the runner:

```powershell
python .github\skills\company-content-generator\scripts\officecli_runner.py preflight --formats docx pptx xlsx --out "$sessionWorkspace\officecli-preflight.json"
```

Resolve `$sessionWorkspace` explicitly; select only the requested formats.
Copy the verified receipt and referenced probe images into `build` after
initialization. Never mark preflight passed from a missing tool, skipped test,
capture-adapter-presence claim, or undisplayed image. The automatic probe
proves readable image bytes, not that a reviewer has seen them. The Playwright
adapter must actually be available (installed, matching Chromium reachable) to
deliver those images for subsequent required visual reviews.

The built-in probes prove basic authoring/save/reopen, extraction, schema/issues,
HTML export + Playwright-Chromium capture, and ordered negative/fractional
cross-sheet formulas. Additional required features need bounded, explicit
probes and actual family pilots; the default probe is not blanket certification
of every OfficeCLI feature.
Do not repeatedly rediscover unchanged help/source or rerun the full maintenance
suite per batch. Revalidate changed binaries, backends, or required capabilities.

Set per-process `OFFICECLI_SKIP_UPDATE=1` and
`OFFICECLI_NO_AUTO_RESIDENT=1` for checked standalone execution. The shared runner
does this without modifying global settings. If a separately approved workflow
uses residents, flush/close only run-owned files before hashing or other tools
read disk bytes. Never quit unrelated sessions or kill processes by name.

If OfficeCLI, the Playwright capture adapter, or a required operation is
unavailable, report **"Generation stopped: required OfficeCLI/Playwright
capability unavailable: <capability>. No fallback was used."** Preserve
existing outputs and report the actual completed count. Maintaining this skill
and running helper tests are not corpus generation.

## Supported save/check/render path

- Use checked **atomic** batches or template/merge operations, not best-effort
  success. Check every item, errors, caveats, truncation, and actual saved bytes.
  Text replacement is not proof that numeric/formula cells retained their types.
  Never force-overwrite accepted artifacts.
- HTML export uses `view <file> html --out <path>`; pixels come from the pinned
  Playwright-Chromium adapter, **never** OfficeCLI's own
  `view screenshot --render html/auto/native` image path.
  The adapter captures one scoped image per `div.page[data-page='N']`,
  visible `div.slide-container[data-slide='N']`, or bounded
  `td[data-path='/Sheet/A1']` region — never a whole-page/-document capture
  used as a stand-in for per-unit review. Review every selected page/slide;
  contact sheets are an optional triage aid only, never a replacement for
  viewing the full page/slide image. XLSX ranges must cover all relevant
  cells and layouts, including merged-anchor cells and non-default/active
  sheets (activate the sheet tab before capturing it). The adapter scrolls
  offscreen ranges into view and enlarges the viewport when a declared range
  exceeds it, preserving font sizes and rejecting partial range captures.
  XLSX chart objects
  (standard column and line charts anchored with a `twoCellAnchor`) are a
  supported, independently proven capture path: `chart_object_inventory()`
  enumerates every chart-bearing graphic frame in the same document order as
  the verified OfficeCLI HTML-export contract
  (`WorksheetDrawing.Descendants<GraphicFrame>()`, OfficeCLI commit
  `b65272a18aef891941451705ff93e8a287f3eb72`,
  `ExcelHandler.HtmlPreview.Charts.cs:41-48,227-231` and
  `ExcelHandler.Helpers.Chart.cs:347-372`), keyed by worksheet part + drawing
  part + source shape ID + chart relationship/part + anchor — never by title
  or start cell alone, so duplicate titles/start anchors stay independently
  addressable. `verify_render_targets()` fails closed on any XLSX package
  containing an **unsupported** drawing/chart object (pie/3-D/combo/other
  anchor forms, orphan drawing parts) unless every supported chart object
  also has a matching render target; it never silently omits, substitutes,
  or shifts a supported chart's index around an unsupported one. The capture
  adapter's `capture_xlsx_chart()` locates the exact
  `div.chart-container[data-path='/{sheet}/chart[N]']` (exact-attribute
  match, never interpolated/fragile partial selectors), requires exactly one
  visible `svg[viewBox]` chart body (excluding decorative legend-swatch
  icons, which carry no `viewBox`), and reports real bounds plus
  clip/occlusion against the document and neighboring objects — a fully or
  partially overlapping pair of charts is reported as occluded, never
  accepted as a clean layout by hiding or ignoring the other object. Column
  and line support is an initial, extensible boundary (see `company.py`'s
  `SUPPORTED_CHART_KINDS`), not a permanent ban on other chart types;
  ordinary per-series colors/fonts are not an exhaustively whitelisted style
  set, and any other chart feature (3-D, combo, unproven styles) must not be
  claimed as supported by the column/line label. A user-required unsupported
  chart type must be reported as a specific blocker, never silently omitted
  or substituted. Source-order identity is bound to the current package
  state only: editing, reordering, or regenerating a workbook invalidates a
  previously captured chart target's `expected_dom_path` and the mapping
  must be recomputed from a fresh inventory, not reused.
- DOCX output from this skill uses **explicit, deliberate page breaks**, never
  automatic natural-flow pagination; a one-page document needs no superfluous
  break. Every authored page's content must fit its physical rendered page:
  the capture adapter rejects a page-count mismatch against the explicit
  break plan and rejects any page whose DOM content overflows its rendered
  box (clipped/dropped/shrunk-unreadably text is a failure, not an accepted
  partial page). Redistribute content within the planned page budget instead
  of writing a custom paginator or accepting an overflowing capture.
- PPTX layouts with long/wrapping titles must reserve the wrapped title's
  full height and keep body content clear of it, with a readable minimum
  font size (no truncation or over-shrinking). This is an ordinary layout
  fix per affected frozen layout, not a universal overlap/collision
  checker — intentional overlays elsewhere are not flagged. Run a long-title
  stress pilot on a layout before bulk use.
- Do not use `refresh` or `stats --page-count`: they can invoke native Word.
  Page-dependent TOCs/cross-references and other unproven field capabilities are
  outside the bounded default path. Required unsupported features block.
- There is no verified standalone `recalc` command. Write source inputs before
  dependent formulas, save, reopen, and check **every** formula against a grounded
  independent expected result. OfficeCLI may remove stale caches rather than
  refresh them; dependency propagation is not guaranteed. One known-order
  corrective formula rewrite is allowed, then recheck final caches and issues.
  Missing, stale, unsupported, error, or unassessed caches block delivery.
  Never inject raw XML caches, invent values, build a substitute calculator,
  accept tolerance-based provenance, or rely on Excel calculating on open.
- Scan all schema/content/structure issues without truncation, and request
  `chart_cache_stale` explicitly when applicable; the default issue scan omits it.

OfficeCLI HTML layout is **not native Office fidelity certification**. Automated
checks cannot certify authentic prose, semantic entailment, substantive depth,
or that someone actually reviewed an image. These limitations remain visible
at delivery under every QA policy.

## Pilots and efficient execution

Pilot each `(genre, format, layout revision/hash)` family before bulk use. Pilots
count toward the document target and need complete content and visual review.
Include grounded stress cases: long titles/text, negative/fractional values,
versions, historical source comparisons, formula dependency ordering, and
worksheet regions. Freeze accepted layouts/commands, not repetitive prose.

Estimate review workload and image-delivery capacity up front. Use one grouped
review of content, genre, depth, conditions and selected renders per document.
The skill explicitly asks during intake whether to run additional QA, defaulting
the choice to no additional QA. Mandatory generation checks and full pilot reviews
still apply. When selected, Sampled/Balanced/Full control the
separate post-generation pass independently of content depth. Do not capture or
review every bulk document simply because a QA policy is recorded for later use.
A failed sample
blocks its family; record the exception and expand affected-sibling review.
Missing/undelivered/unreadable images cannot count as reviewed.

Use `scripts\officecli_runner.py` for mechanics, bounded stages, durable
checkpoints, pinned tool identity, phase timing, and at most one concrete
corrective retry per logical operation. A user-supplied runtime budget stops at
a safe checkpoint; no invented budget or completion promise. Preserve verified
outputs on resume. Changed facts/drafts/layouts require explicit affected
revisions and fresh mappings/reviews, not silent replacement.

New execution-version-2 runs separate generation completion (`assemble`) from
additional QA (`qa`, actual reviews, then `finalize-qa`). Exporter-recorded private
bindings remove manual offset reconstruction; verification still uses actual
saved bytes. Reuse the helper's operation-scoped validated snapshots rather than
repeating standalone index/provenance/coverage commands in the orchestrated pass.
Final independent audit revalidates bytes in a fresh context.

Reuse content-independent layouts, with document-specific prose/data in separate
jobs. `family_summary` reports singleton families and all-selected visual samples;
do not introduce a unique layout hash for each document merely because its text
differs. Existing browser reuse is retained, not reimplemented per document.
Automatic authoring/check execution uses up to four workers, bounded by available
CPUs and eligible jobs, with `--workers 1` for explicit serial operation. One
capture lane owns synchronous Playwright. A run lock and document-owned atomic
checkpoints prevent competing invocations and aggregate-checkpoint races.
Required failures are surfaced, not silently retried through another backend.

Record agent-owned planning/drafting/editorial/visual spans with the runner's
`phase` command, and inspect `summary`. Distinguish wall time, overlapping worker
activity, capture time and between-invocation gaps; missing phase data is unknown.
Optimize measured repeated work instead of claiming a timing bottleneck from
document counts alone. Existing corpora keep their previous contract and are not
migrated or re-certified by these changes.

Separate backend development from generation. A missing capability produces
evidence and a blocker, not an open-ended framework-building detour. Delegate
review only when authorized. Do not regenerate a completed corpus merely to
rewrite historical engine evidence.

## Maintenance checks

`python -m unittest discover -s tests -q` runs standard-library tests without
Microsoft Office. `RUN_OFFICECLI_TESTS=1` enables actual OfficeCLI integrations.
Missing tools skip optional maintenance tests explicitly; skips never count as
passed generation preflight. Tiny fixtures are not a new company corpus or
editorial/visual certification.
