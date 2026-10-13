# Document backend selection and setup

The prepared canon, plan, contexts, and grounded drafts are shared inputs for
two document-export backends. Select exactly one backend for each new run:

| Backend | Requirements | Authoring, calculation, and rendering |
|---|---|---|
| Native Office | Windows, plus usable local Microsoft Office applications for the selected formats | PowerShell COM automation in Word, Excel, and PowerPoint |
| LibreOffice container | Docker Engine and Docker Compose v2 | Pinned Python authoring packages and LibreOffice Writer, Calc, and Impress |

The container is an independent backend, not Microsoft Office. Its output may
differ in pagination, formatting, or calculation. Do not describe its
rendering or recalculation as Microsoft Office QA or claim Office fidelity.

## Per-run capability check and recommendation

Before creating the output directory, canon, plan, contexts, or drafts:

1. Determine the requested formats. Probe only the matching native Office
   applications: Word for DOCX, PowerPoint for PPTX, and Excel for XLSX.
2. On Windows, use the first-party COM preflight in
   `scripts\office_com.ps1 -Command preflight -Formats <selected-formats>`.
   A passing probe must connect to each selected application, author and
   reopen a small representative file, report the application version, and
   exercise the selected format's essential render/calculation operation.
   COM registration alone is not proof of installed, activated, usable Office.
3. If every requested native format passes, recommend **Native Office**. If
   any requested format is absent or its probe fails or is inconclusive,
   recommend **LibreOffice container**. Explain the missing native capability.
4. Ask the user to choose a backend for the new run, marking the recommendation.
   If the container is recommended, verify Docker Engine and Compose v2 are
   available before accepting that choice. A recommendation is not a silent
   backend switch.
5. Record the selected backend, recommendation, selected formats, capability
   results, application/runtime versions, and any limitations in
   `build\preflight.json`. On resume, use that run's recorded backend. If it
   becomes unavailable, stop and preserve files; never change backend
   automatically after a failure.

If neither backend is usable, stop before corpus generation and report the
specific blockers. Do not produce a partial corpus or install Office,
Docker, or document-generation dependencies without authorization.

## Native Microsoft Office

Native Office means locally installed desktop Word, Excel, and PowerPoint
controlled through Windows PowerShell STA COM automation. Web subscriptions,
Office file associations, and the presence of a COM registration are not
enough. The preflight must establish a real application connection and
exercise the operations required by the selected formats.

| Format | Required native operations |
|---|---|
| DOCX | Word author/save, extract, paginate, and PDF render |
| PPTX | PowerPoint author/save, extract slide text/notes, and render every slide |
| XLSX | Excel author/save, calculate, reopen cached values, scan formula errors, and render |

Run the probe into a new temporary directory. Keep the emitted receipt with
the run's preflight evidence; do not use it as a substitute for the actual
planned-document pilot.

```powershell
$probe = Join-Path $env:TEMP ("company-office-probe-" + [guid]::NewGuid().ToString("N"))
& .github\skills\company-content-generator\scripts\office_com.ps1 `
    -Command preflight `
    -Formats @("docx", "pptx", "xlsx") `
    -ProbeDirectory $probe
```

Use only the selected formats in `-Formats`. The script must launch isolated,
run-owned Office instances and stop only those processes. Never close or
terminate the user's existing Office sessions, alter global security settings,
or accept unexpected prompts automatically. A failed/inconclusive native
preflight means Native Office is unavailable for this run; if the user chooses
it anyway, stop rather than routing to the container.

Use the prepared-artifact native exporter for each planned document. It must
write the exact planned filename, verify extracted content, create Office
render evidence under `build`, and refuse to overwrite outputs. For XLSX,
preserve supplied formulas, recalculate in Excel, reopen to verify cached
values, and scan every used cell for formula errors. Never invent formulas
when the grounded draft supplies none.

## Pinned LibreOffice container

The container exports only the prepared local canon, frozen plan,
date-filtered contexts, and validated grounded drafts. It does not generate
prose, seed a different canon, or select the document mix. Its Dockerfile,
locks, CLI, package-index option, supported draft structures, and test command
are documented in `prototypes\libreoffice-backend\README.md`.

The standard invocation is:

```powershell
$env:CORPUS_ROOT = $outputRoot
docker compose -f prototypes\libreoffice-backend\compose.yaml run --build --rm backend
```

Use an approved HTTPS `PIP_INDEX_URL` only if public PyPI is unavailable. The
default is public PyPI. The setting must not contain credentials, tokens, query
parameters, or fragments; the selected index is exclusive, hash verification
remains enabled, and there is no second-index fallback:

```powershell
$env:PIP_INDEX_URL = 'https://packagefeedproxy.microsoft.io/pypi/simple/'
docker compose -f prototypes\libreoffice-backend\compose.yaml build backend
```

Do not place credentials in the URL, build arguments, command lines, source
files, logs, or image environment. Rebuild after changing the host variable.
If Docker, Compose, the pinned image, the contract helper, or a required
Writer/Calc/Impress operation is unavailable, stop; do not use the other
backend unless the user starts a new run and selects it.

The exporter checks package integrity, content extraction, PDF output, and
page/slide previews. LibreOffice recalculates XLSX and checks retained formulas,
cached values where supplied, and formula-error cells. A workbook with no
grounded formulas is reported as having no formula results to verify. Review
every generated preview separately; successful conversion is not a visual,
semantic, genre, or depth review.

## Shared preflight, QA, and data handling

For either backend, keep run-specific scripts, logs, temporary files, Office
outputs, PDFs, previews, and receipts under the selected corpus's `build`
directory. Final documents alone belong in `documents`. Keep canon, plan,
contexts, and drafts unchanged after export; an accepted output is never
overwritten during a retry.

`build\preflight.json` is an agent-maintained receipt. It must identify:

- the chosen and recommended backend, selected formats, and timestamp;
- native application or container/runtime versions and probe evidence;
- backend implementation source/revision and relevant package/image versions;
- the planned pilot IDs, output hashes, render evidence, and limitations.

Do not mark preflight passed based only on a version string, executable
detection, COM registration, Docker availability, or successful conversion.
After freezing and validating the plan and drafts, complete one real planned
pilot per selected format through the chosen backend. Inspect visual output;
check XLSX formulas, cached values, expected results, and formula errors.
Record honest semantic, genre, depth, visual, and recalculation reviews in
`build\qa.json`. The structural Python audit does not certify which application
created a file or prove those reviews occurred.

Neither backend may retrieve real employee/customer/company data or upload
canon, contexts, drafts, documents, or previews to an external service.
Do not switch backends mid-run to recover from an error. Preserve accepted
artifacts, report the blocker and completed count, and retry only in a new
output revision after resolving it.
