# Microsoft Office dependencies and setup

## Hard prerequisite: local Microsoft Office

Corpus generation requires locally installed, licensed, usable Microsoft Office.
Only skills and/or MCPs that actually automate those applications are eligible.
Office-compatible file output, an "Office" product name, or a cloud Microsoft
365 account is not proof of a local Microsoft Office backend.

| Selected format | Required application | Required operations |
|---|---|---|
| DOCX | Microsoft Word | Author/save, extract content, paginate, render/export QA pages |
| PPTX | Microsoft PowerPoint | Author/save, extract slides/notes, render/export every slide |
| XLSX | Microsoft Excel | Author/save, recalculate, save/reopen formula caches, inspect errors and layout |

For every selected format, the backend must also associate grounded draft
section/claim uses with final artifact containers. It can record locators during
export or use the helper's final-file index afterward. See `provenance.md`.
This is an export handoff requirement, not a requirement to insert generation
IDs into business text or Office properties.

Verify only the applications needed by the selected formats. A Word-only run
does not require Excel or PowerPoint. Verify application usability through the
selected backend before creating a corpus directory, canon, plan, or drafts.
Repeat the check before resuming generation.

If Office is absent, inaccessible, unlicensed, or cannot be verified, notify the
user and stop:

> Generation stopped: local Microsoft Office is required. Missing or unavailable:
> \<applications/capabilities\>. No fallback was used.

If Office is installed but no compliant skill/MCP can drive it, explain that
blocker and stop as well. Do not create a canon/planning-only partial corpus,
install replacements, or switch to cloud-only document operations. On failure
mid-run, preserve accepted artifacts and report the actual completed count.
Skill maintenance and standalone deterministic helper tests are exempt: they
do not generate an Office corpus.

## Eligible and excluded tooling

A reviewed skill or MCP may use Windows COM/Office interop, or another verified
local automation route, to control actual Microsoft Office applications.
Microsoft documents Word/Excel/PowerPoint application automation; see the
official references below. On other operating systems, verify an actual native
Office connection rather than assuming the Windows probe applies.

There is no default named external document skill. Review the implementation,
not its name. A `docx`, `pptx`, or `xlsx` skill qualifies only when its selected
workflow really uses Microsoft Office for authoring, rendering, and calculation.
Reject instructions that require an incompatible backend; do not load them
and silently override their mandatory rendering or calculation steps.

**Never use LibreOffice, OpenOffice, `soffice`, standalone OOXML generators,
third-party rendering/calculation engines, or cloud-only Office access as a
fallback.** Packages such as npm `docx`/`pptxgenjs` and Python `openpyxl` are
not Microsoft Office authoring engines and are not selected export backends.
The independent `iOfficeAI/OfficeCLI` engine advertises operation without an
Office installation; its previously listed skill bundle is not an eligible
replacement. Do not install it for this workflow.

Create diagrams/charts and QA exports in Microsoft Office rather than requiring
generic diagram/PDF-generation skills. Read-only JSON, ZIP/XML, hash, extracted
text, and image/PDF inspection tools may support QA of Office-created outputs.
They may not substitute for Word/PowerPoint rendering or Excel recalculation.
The canon helper requires only Python 3.10+ and its standard library.

## Discover and obtain a compliant skill or MCP

Use the host's skill/MCP discovery UI to inspect available candidates and their
actual source paths. Already installed skills are not implicitly eligible.
Inspect the complete instructions, referenced scripts, runtime dependencies,
licenses, and exact revision. Confirm the candidate connects to local Microsoft
Office and covers every selected capability, including QA.

For an external candidate:

1. Obtain its repository/package source from the publisher or an
   organization-approved catalog. Do not guess a repository, MCP tool signature,
   or capability from its name. No external bundle is automatically installed.
2. Download into a unique review directory outside discoverable skill paths.
   Pin an exact commit/release and review all code, license/notice files,
   permissions, local application operations, and network behavior.
3. Reject candidates that generate documents independently of Microsoft Office,
   mandate substitute engines, upload content to third parties, or require
   unrelated permissions. A remote/cloud-only MCP cannot bypass local Office.
4. Obtain authorization before installing or enabling external code. Copy the
   complete reviewed skill folder, including helpers/templates/notices, into
   the chosen project's `.github\skills` or the user's `.copilot\skills`.
   Do not overwrite an existing skill or copy only `SKILL.md`.
5. For an MCP, use the host's supported registration procedure for that exact
   reviewed package/version. Restrict its scope to the required local Office
   operations. Confirm the loaded source/version and actual connection after
   reload; a project skill may shadow a personal one.

The following optional download pattern does not execute downloaded code.
Supply the actual reviewed source and commit first; never run the placeholders:

```powershell
$sourceRepository = "<approved publisher repository URL>"
$approvedCommit = "<reviewed full commit ID>"
if ($sourceRepository.StartsWith("<") -or $approvedCommit.StartsWith("<")) {
    throw "Provide the approved source and exact reviewed commit first"
}
$reviewRoot = Join-Path $env:TEMP ("office-automation-review-" + [guid]::NewGuid().ToString("N"))
git clone --no-checkout $sourceRepository $reviewRoot
if ($LASTEXITCODE -ne 0) { throw "Source download failed" }
git -C $reviewRoot checkout --detach $approvedCommit
if ($LASTEXITCODE -ne 0) { throw "Cannot check out the approved revision" }
Get-ChildItem -LiteralPath $reviewRoot -Force
```

Review the checkout before copying any skill or enabling an MCP. Retain
applicable licenses/notices. Record the source, exact commit/release, permission
review, installation scope, and adaptations. Re-review upgrades. Do not pipe
downloaded installers into `Invoke-Expression`, grant blanket shell permissions,
or automatically install Microsoft Office. If Office is missing, notify and stop;
the user can install/activate it through Microsoft or their organization's
approved channel before requesting another run.

## Local Office availability and connection probes

On Windows, an initial read-only COM registration check can identify missing
applications without opening documents or touching the user's running sessions.
Set `$selectedFormats` from the actual request, not from this example:

```powershell
$selectedFormats = @("docx", "pptx", "xlsx")
$programIds = @{
    docx = "Word.Application"
    pptx = "PowerPoint.Application"
    xlsx = "Excel.Application"
}
$missing = @()
foreach ($format in $selectedFormats) {
    if (-not $programIds.ContainsKey($format)) {
        throw "Unsupported format: $format"
    }
    if ($null -eq [type]::GetTypeFromProgID($programIds[$format])) {
        $missing += $programIds[$format]
    }
}
if ($missing.Count -gt 0) {
    throw ("Generation stopped: local Microsoft Office is required. Missing or unavailable: " +
        ($missing -join ", ") + ". No fallback was used.")
}
```

Registration is only an initial check: stale registration, activation, licensing,
dialogs, or automation permissions can still prevent use. The selected skill/MCP
must establish a real application connection and report the actual application
version before generation starts. A successful registration probe alone must
never mark preflight passed. A failed/inconclusive connection is a stop condition,
not permission to try a different document engine.

Use an isolated automation session when supported. Never quit or kill the
user's existing Word, Excel, or PowerPoint sessions. Clean up only instances
and documents owned by the run, and release automation handles. Do not alter
global macro/trust/security settings, accept unexpected prompts automatically,
or terminate applications by process name.

Probe only the selected compliant backend's actual dependencies. Install
authorized missing automation packages only after a failed probe or manifest
change; do not install standalone document-generator stacks. Keep run-specific
scripts, dependency manifests/locks, logs, and exports under `build`.
Shared installed Office/automation tools may remain outside the corpus.

## Mandatory Microsoft Office pilots and QA

After the connection checks pass and the canon/plan are frozen, use one actual
planned item per selected format as a pilot. Pilots count toward the requested
target. No bulk generation until every selected pilot passes.

| Pilot | Passing evidence |
|---|---|
| DOCX | Authored/saved with Word; valid package; extracted content matches the grounded draft; genre/depth review; Word-rendered pages inspected for pagination, tables, and branding |
| PPTX | Authored/saved with PowerPoint; valid package; slide text/notes match the draft; genre/depth review; every PowerPoint-rendered slide checked for overflow/overlap |
| XLSX | Authored/saved and recalculated with Microsoft Excel; formulas retained; saved cached values readable after reopen; zero formula errors; expected totals/margins and layout checked |

An API success response or recalculation exit code is not sufficient evidence.
Inspect results, errors, expected values, and rendered content. Flush/save/close
run-owned documents before hashing or delivery. Do not alter accepted artifacts
after recording their hashes.

Record application names/versions, local connection/probe evidence, selected
skill/MCP sources and exact revisions, permission review, pilot IDs/hashes,
QA evidence, and blockers in `build\preflight.json`. This is an agent-maintained
workflow receipt: the Python package/claim audit does not detect installed Office
or certify which application created a file. Do not claim otherwise.
Continue to provide the contract's `build\qa.json` and `build\manifest.json`,
including semantic, genre, depth, visual, and Excel recalculation receipts.
For new schema-version-2 runs, also produce occurrence mappings, verified
`build\provenance.json`, and recomputed `build\coverage.json`. A matching token
does not certify correct attribution or meaning; verify these in the pilot and
semantic review. Configured condition carriers require `condition_review`.
Single-file pilots may use the helper's read-only `artifact_containers` and
`check_draft` functions before a complete-corpus manifest exists.
Recheck Office availability on resume and after backend/version changes.

Inspect body text, notes, comments, hidden workbook content, and properties for
generation self-labeling. Never prefill passed flags. If any required Microsoft
Office operation fails, notify the user and stop without a fallback.

The helper uses only standard-library JSON, hash, ZIP, and XML inspection.
These read-only checks do not author, render, or recalculate Office documents.
Opt-in integration tests use temporary, test-owned Office documents and must
not attach to or terminate a user's existing application session.

For skill maintenance, deterministic tests do not require Office:
`python -m unittest discover -s tests -q`. To additionally check real Office
save/reopen containers, run the same suite with
`$env:RUN_OFFICE_PROVENANCE_TESTS = "1"`. Each integration test skips explicitly
if its application is already running or absent; automation/save/reopen
failures are failures, not fallback successes. These tiny fixtures are not a
generated corpus or a visual/semantic quality certification.

## Authoritative references

Microsoft Office automation and independent-backend exclusion were checked on
October 9, 2026. Recheck the exact candidate revision before adoption.

```text
Microsoft Office automation concepts:
https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/understanding-automation
Microsoft Word application automation:
https://learn.microsoft.com/en-us/office/vba/api/word.application
Microsoft Excel controlling Office applications:
https://learn.microsoft.com/en-us/office/vba/excel/concepts/working-with-other-applications/controlling-one-microsoft-office-application-from-another
Microsoft PowerPoint application object:
https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application
Excluded independent OfficeCLI source (not an installation recommendation):
https://github.com/iOfficeAI/OfficeCLI
```
