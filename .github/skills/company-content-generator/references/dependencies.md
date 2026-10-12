# Container backend dependencies and setup

## Required backend

Corpus document export runs in the repository's pinned LibreOffice backend
container. The host does not need Microsoft Office, Python, LibreOffice, or a
third-party document skill. Docker Engine and the Docker Compose v2 plugin are
required for the standard local workflow. An interactive Dev Container using
the same pinned `Dockerfile` is also supported when it can read the prepared
corpus paths.

| Format | Container authoring and QA |
|---|---|
| DOCX | `python-docx`; structural extraction, LibreOffice Writer PDF export, PDF text/page validation, PNG page previews |
| PPTX | pinned `python-pptx`; package/text checks, LibreOffice Impress PDF export, slide/page count and PNG preview for every slide |
| XLSX | `openpyxl`; LibreOffice Calc save/recalculation, retained formula/cache and error scan, PDF export and PNG page previews |

Dependencies are locked in `prototypes\libreoffice-backend\apt-packages.lock`
and `requirements-pptx.lock`. The Debian base image is digest-pinned; APT
resolves against the pinned Debian snapshot; the Python PowerPoint wheel uses
an exact version and SHA-256 hash. The pip index defaults to public PyPI and
can be changed at build time with the host `PIP_INDEX_URL` environment
variable. For the Microsoft package proxy, set it to
`https://packagefeedproxy.microsoft.io/pypi/simple/`. Only use an approved
HTTPS package index URL without embedded credentials, tokens, query parameters,
or fragments. The selected index is used exclusively; locked hashes remain
enforced and there is no fallback to a second index. Runtime export must use
this image built from the current repository revision. Do not substitute a host
installation, cloud converter, Office engine, or another renderer after a
failure.

## Preflight

Before creating the chosen corpus output directory, canon, plan, contexts, or
drafts:

1. Confirm the user has Docker Engine and Compose v2 available, inspect the
   repository Dockerfile and dependency locks, and set `PIP_INDEX_URL` to the
   organization's approved HTTPS package index only when the default public
   PyPI route is unavailable. Never include credentials or tokens in this
   build argument.
2. Build the pinned image and run its focused smoke suite, which checks the
   non-root container user, contract helper mount, formats, PDF previews,
   formula recalculation/error handling, and no-overwrite behavior.
3. Verify the container imports `docx`, `openpyxl`, `pptx`, and `pypdf`, and
   that `soffice --version` reports the pinned LibreOffice runtime.
4. If Docker, the pinned build, package imports, helper mount, or required
   Writer/Calc/Impress operations are unavailable, stop before corpus
   generation. Report the exact missing capability; do not switch engines.
5. Keep all run-specific authored files, calculated workbooks, PDFs, previews,
   logs, manifests, and engine receipts under the selected corpus's `build`.
   The source canon, plan, contexts, and drafts are read-only workflow inputs.

The normal full-corpus invocation uses the prepared corpus root:

```powershell
$env:CORPUS_ROOT = $outputRoot
docker compose -f prototypes\libreoffice-backend\compose.yaml run --build --rm backend
```

The service calls the Python CLI with explicit paths:

```text
python3 -m office_backend
  --canon /corpus/canon.json
  --plan /corpus/build/plan.json
  --contexts /corpus/build/contexts
  --drafts /corpus/build/drafts
  --root /corpus
```

Use repeatable `--document <planned-id>` options for a bounded selected-item
pilot. All required inputs must exist and match the plan; the CLI stages exact
copies only when the specified output root is new. It rejects symlinks,
unexpected output-root entries, unplanned IDs, context/draft fingerprint
mismatches, unsupported draft structures, and any attempt to overwrite a
document or evidence output.

If the CLI or renderer fails, preserve source artifacts, report the exact
error and partial generated count, and use a new output root for a retry after
the issue is resolved. Do not delete or overwrite accepted documents to retry.

## Pinned backend QA and honest evidence

Each planned output is authored with its specified basename under `documents`.
The container verifies the repository claim contract, package integrity,
extracted section text, PDF structure/content, page count, and rasterizes every
page under `build\renders\<document-id>`. For XLSX, LibreOffice Calc reopens
and saves the workbook, and the backend checks formula expressions, cached
values where present, and formula-error cells. A workbook with no supplied
formulas is reported as containing no formula results to verify; the backend
does not invent financial calculations.

The backend writes `build\manifest.json` only for a complete plan. A
`--document` run writes an explicitly partial pilot manifest/evidence file
instead. The separate engine evidence records the backend/runtime version,
source context/draft hashes, artifact hashes, PDF checks, preview paths, and
formula outcomes. It is machine evidence only, not a semantic, genre, depth,
or completed visual review.

Inspect every rendered page/slide preview. Do not mark a visual review passed
from successful conversion or image creation alone. Complete separate
semantic, genre, and depth reviews using the skill's content playbook and
`build\depth-targets.json`; review page layout, table/chart readability,
clipping, and formula outputs. Write honest `build\qa.json` receipts only after
those reviews. Record `LibreOffice` and the exact version in visual and XLSX
recalculation evidence. Never label this as Microsoft Office rendering,
Microsoft Excel recalculation, or Office-compatible output.

LibreOffice rendering and calculation may differ from Microsoft Office.
Record that limitation in `build\preflight.json` and do not claim Office
fidelity. The output container uses no remote document service and does not
upload canon, context, draft, or document content.
