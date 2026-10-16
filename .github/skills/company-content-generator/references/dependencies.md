# OfficeCLI dependency, setup, and capability boundaries

## Required components

Use **[OfficeCLI](https://github.com/iOfficeAI/OfficeCLI)** as the sole document
engine for authoring, extraction, schema/issues checks, supported formula
evaluation, and static HTML export. Microsoft Office is **not required**.
Python 3.10+ performs business/claim validation, read-only package inspection,
hashing, provenance, coverage, and orchestration; it is not a second authoring
or calculation engine. Screenshot pixels never come from OfficeCLI: a thin
pinned **Playwright-for-Python 1.63.0** adapter
(`scripts/playwright_capture.py`) drives a matching **managed Chromium** build
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

Choose the official asset for both **OS and architecture**, not just its name.
OfficeCLI publishes separate Windows x64/arm64, macOS x64/arm64 (Apple Silicon),
and Linux x64/arm64 assets, with separate Alpine/musl variants. Asset availability
is release-specific: check the chosen version's release and `SHA256SUMS`. A Linux
OfficeCLI asset alone does not establish that the managed Chromium adapter
supports that distribution. Keep executable permission, PATH, and macOS
Gatekeeper/quarantine troubleshooting explicit and scoped; do not disable
platform protections or substitute an unverified binary.

## uv environment and native paths

Use **[uv](https://docs.astral.sh/uv/)** and Python **3.10+**. Resolve the installed
skill directory containing `pyproject.toml` and `uv.lock`; run generation examples
from that directory. uv creates and manages its `.venv` automatically, without
manual creation or activation. Do not use `--active` to modify an unrelated
environment. For a read-only installation, set process-local
`UV_PROJECT_ENVIRONMENT` to an approved, dedicated writable directory (under
session storage before initialization or `build` afterward), then keep that
environment selection consistent for setup and execution.

Obtain authorization before installing uv, Python, packages, or browsers.
Set process-local `UV_PYTHON_DOWNLOADS=never` unless a Python download was
explicitly authorized; use `--python` with an approved installed interpreter
when needed. uv does not replace OfficeCLI or install Chromium/system libraries
as part of ordinary helper execution.

The base project has no dependencies. The `capture` extra pins
`playwright==1.63.0`; its transitive dependencies are fixed by `uv.lock`.
Use `--locked` to reject stale metadata rather than changing the lock during
generation. Update the lock explicitly with `uv lock` only as authorized
maintenance, and revalidate changed package/browser identities and capabilities.
Once setup is complete, reuse that environment and avoid downloads in the
document loop; `--offline` can enforce cached-only uv execution.

Examples are command recipes using `/` and quoted `<placeholders>`, not a
universal shell script. Replace placeholders with resolved absolute paths,
use OS-specific separators for native CLI arguments, and adapt variable,
quoting, and environment-assignment syntax to the actual shell. New persisted
corpus-relative references still use `/` on every OS; do not rewrite stored
evidence. Never assume a Windows drive path is meaningful on POSIX. The helper
does not expand a literal `~` in JSON.
Choose a short output root; complete-path limits and external-tool long-path
support remain filesystem/host-specific. Do not change registry or global Git
settings as a workaround.

## Playwright capture adapter setup

The adapter is a thin, **optional** dependency: importing
`scripts/playwright_capture.py` and running the default unit suite never
require Playwright to be installed (the import is lazy, inside
`CaptureEngine.__init__`). It is only required to actually capture pixels
(preflight probes and real renders).

1. If the selected uv environment already has the locked capture dependencies
   and matching managed Chromium working, **no additional install is required**.
2. Otherwise, synchronize the locked `capture` extra, then use that same
   environment to fetch the matching managed build from the official Playwright
   CDN. Do not hardcode a user name, session ID, or machine-specific path.
   `PLAYWRIGHT_BROWSERS_PATH`, if needed, is process-local, never a global setting.
3. Honor the consumer's approved package feed through uv configuration or
   `UV_DEFAULT_INDEX`. uv does not read `pip.ini`, `pip.conf`, or `PIP_INDEX_URL`;
   translate an approved pip-only feed explicitly rather than silently switching
   indexes. Never commit credentials or force a consumer-specific feed into the
   portable project. The distributed lock uses public PyPI artifact URLs;
   changing the registry can invalidate `--locked` even when versions match.
   If policy requires a different registry, obtain approval to regenerate a
   mirror-specific lock in a dedicated writable copy of the skill before
   preflight; do not silently rewrite a shared installation or use `--frozen`
   to bypass an inconsistent lock. Keep that selected copy and feed consistent
   for the run, and review its resolved versions and registry URLs.
4. Record the resulting identity in the capability receipt: Playwright
   version, browser channel, managed Chromium version, its executable path,
   and its SHA-256 — the same evidentiary treatment as the OfficeCLI
   executable/version/checksum above. Treat drift in either identity as a
   capability change requiring revalidation, not a silent pass-through.

After authorization and only when dependencies are genuinely missing, run from
the installed skill directory:

```text
uv sync --locked --extra capture
uv run --locked --extra capture python -m playwright install chromium
```

### Linux libraries, sandboxing, and fonts

Consult Playwright's official
[system requirements](https://playwright.dev/python/docs/intro#system-requirements)
and [browser setup](https://playwright.dev/python/docs/browsers#install-system-dependencies)
for the pinned release. The current upstream guide lists Windows 11+/Server
2019+/WSL, macOS 14+, Debian 12/13, and Ubuntu 22.04/24.04/26.04 (Linux x64 or
arm64); this is upstream guidance, **not** certification of every host by this
skill or a promise that a future platform is supported by the pinned version.
Minimal Linux/container images can lack Chromium's shared libraries even after
the Python package and browser download succeed.

`uv run --locked --extra capture python -m playwright install-deps chromium`
installs host system packages and
may require elevated privileges. It is a **separate, explicitly authorized
machine change**, not an automatic generation or unit-test step. Review the
needed packages and the host administrator's approved installation procedure.
Do not blindly use `--with-deps` or assume all Linux distributions are supported.

The adapter launches sandboxed Chromium. Use an appropriate non-root account
and a host/container configuration that permits Chromium's sandbox; inspect
actual launch errors and follow upstream/admin guidance. Do not add
`--no-sandbox`, disable sandboxing, or weaken host protections to make a probe
appear successful. An unavailable sandbox/browser is an explicit blocker.

Arial and Calibri are not guaranteed on Windows, macOS, or Linux. For **new**
layouts, select fonts actually available to Chromium on that host with suitable
licenses, and record the chosen typography in the frozen layout. A font request
in HTML and `document.fonts.ready` do not prove the requested font was installed;
fallback fonts can alter wrapping, overflow, and slide geometry. Check actual
preflight images and representative/long-title family pilots. Host, font, or
browser changes require fresh capability/layout review; do not silently install
fonts, substitute them in historical outputs, or claim identical pixels.
Font inventory is not currently enforced or fingerprinted by the adapter.

## Preflight, once per capability identity


**Before creating the output directory, canon, plan, or drafts**, verify Python
and perform tiny session-owned probes for the selected formats with the runner:

```text
uv run --locked --extra capture scripts/officecli_runner.py preflight --formats docx pptx xlsx --out "<session-workspace>/officecli-preflight.json"
```

Resolve `<session-workspace>` explicitly; select only the requested formats.
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

Use `scripts/officecli_runner.py` for mechanics, bounded stages, durable
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

From the repository root, select the skill's locked base project:

```text
uv run --locked --project .github/skills/company-content-generator python -m unittest discover -s tests -q
```

This runs standard-library tests without Microsoft Office or capture packages.
`RUN_OFFICECLI_TESTS=1` enables actual OfficeCLI integrations.
Missing tools skip optional maintenance tests explicitly; skips never count as
passed generation preflight. Tiny fixtures are not a new company corpus or
editorial/visual certification.

The repository's `Cross-platform Python tests` GitHub Actions workflow runs it
on Windows, macOS, and Ubuntu with Python 3.10/3.14. It installs no application
dependencies and deliberately skips real OfficeCLI/browser integrations.

To opt in locally **after** authorized setup and actual tool availability, set
`RUN_OFFICECLI_TESTS=1` for that process using the current shell's syntax, run:

```text
uv run --locked --extra capture --project .github/skills/company-content-generator python -m unittest discover -s tests -q
```

Then restore the prior environment setting. `--project` selects the environment
without changing the repository working directory used for test discovery.
From another directory, use absolute project and test-directory paths. Plain
Python remains supported for callers managing their own compatible environment;
uv is the primary documented setup and execution workflow.

The separate `SkillSpector static scan` workflow scans the distributable skill
with a pinned scanner and its upstream lockfile, `--no-llm`, and
`--fail-on-incomplete`. It preserves the native risk-score threshold (>50) and
errors rather than accepting all findings into a baseline. JSON reports and
revision metadata are retained even on failure. Tracing and model calls are
disabled; OSV.dev may receive dependency names/versions, not source contents.
Scanner dependencies are separate from this skill's runtime.

The included render proof remains Windows-pinned-version-only. A green Python
CI matrix is not real macOS/Linux capture proof, native Office fidelity, or
cross-platform pixel identity. New hosts still need real preflight and pilots.
