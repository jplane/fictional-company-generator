# Container document backend

This prototype exports **prepared** company-corpus artifacts. The agent still
handles intake, creates and validates the canon and plan, derives date-filtered
contexts, and writes grounded drafts under the company-content-generator skill.
The Python CLI does not seed a canon, select document types, call an AI model,
or replace supplied draft content. It validates the supplied context/draft
contract, authors each planned output, recalculates XLSX files, renders all
formats to PDF, and saves page previews and engine evidence under `build/`.

The existing content-generation workflow invokes this backend only after the
skill has prepared and validated the corpus inputs. LibreOffice is a pinned
container renderer/calculation engine, not Microsoft Office and not an Office
fidelity certification.

## Inputs and CLI

The host needs Docker Engine and Docker Compose v2. Python, Office, LibreOffice,
and document-generation packages are supplied by the pinned container.

For a corpus already prepared by the skill, set `CORPUS_ROOT` to that corpus
directory. The default Compose command uses:

```text
python3 -m office_backend \
  --canon /corpus/canon.json \
  --plan /corpus/build/plan.json \
  --contexts /corpus/build/contexts \
  --drafts /corpus/build/drafts \
  --root /corpus
```

PowerShell example:

```powershell
$env:CORPUS_ROOT = 'C:\Users\you\Documents\Jimworks-v2'
docker compose -f prototypes\libreoffice-backend\compose.yaml run --build --rm backend
```

The equivalent CLI arguments are required: `--canon`, `--plan`, `--contexts`,
`--drafts`, and `--root` (or `--output-root`). Normally they point to the
prepared files and corpus root inside the container. If the output root is new,
the CLI safely stages the canon, plan, request, and selected contexts/drafts
there without changing the source inputs.

To run a bounded two-document pilot against an existing corpus from a read-only
source mount, create an empty output directory first and pass the selected IDs:

```powershell
New-Item -ItemType Directory -Force 'C:\temp\jimworks-pilot' | Out-Null
docker run --rm `
  --mount 'type=bind,source=C:\Users\you\Documents\Jimworks-v2,target=/inputs,readonly' `
  --mount 'type=bind,source=C:\temp\jimworks-pilot,target=/output' `
  --mount 'type=bind,source=C:\path\to\fictional-company-generator,target=/workspace,readonly' `
  --mount 'type=bind,source=C:\path\to\fictional-company-generator\.github\skills\company-content-generator\scripts\company.py,target=/app/company_contract.py,readonly' `
  --workdir /workspace/prototypes/libreoffice-backend `
  jplane-libreoffice-backend `
  python3 -m office_backend `
  --canon /inputs/canon.json `
  --plan /inputs/build/plan.json `
  --contexts /inputs/build/contexts `
  --drafts /inputs/build/drafts `
  --root /output/jimworks-pilot `
  --document doc-0003 `
  --document doc-0004
```

The image name in this example can instead be any tag built from this
directory's Dockerfile. `--document` may be repeated. Without it, all planned
documents and all matching contexts/drafts are required. Every provided
context is recomputed from the canon and plan through the repository contract
helper, and every selected draft is checked for matching IDs, fingerprints,
claims, required sections, and supported export structures.

The output layout remains:

```text
<root>\
  canon.json
  documents\
    <exact planned filenames>
  build\
    request.json
    plan.json
    contexts\
    drafts\
    manifest.json                 # only when the complete plan is exported
    backend-pilot-manifest.json   # selected documents only
    backend-evidence.json
    backend-pilot-evidence.json
    backend-export.log
    renders\<document-id>\
      document.pdf
      doc-<id>-page-<n>.png
```

The CLI refuses symlinked output roots/directories, plan mismatches, missing or
unplanned selected IDs, unsupported draft structures, and existing document,
manifest, evidence, log, or render outputs. It does not overwrite inputs or
create a success-shaped manifest for partial exports. A selected-document run
gets an explicitly partial pilot manifest and evidence receipt; only an
all-documents run writes the normal `build/manifest.json`. Semantic, genre,
depth, and human visual review remain required before `build/qa.json` may be
marked passed and the corpus audited.

## Dev Container and tests

For interactive development, open the repository root in VS Code with the Dev
Containers extension and choose **Reopen in Container**. It builds the same
pinned runtime, runs as the unprivileged `app` user, and mounts the repository
contract helper read-only at `/app/company_contract.py`. Mount the prepared
corpus at `/corpus` before invoking the CLI with the arguments above.

The Compose image runs the same CLI. Prepare a disposable corpus root for
smoke tests; unit tests build their own canonical inputs in temporary folders:

```powershell
New-Item -ItemType Directory -Force prototypes\libreoffice-backend\output | Out-Null
docker compose -f prototypes\libreoffice-backend\compose.yaml run --build --rm backend `
  python3 -m unittest discover -s tests -v
```

The test-only `run_demo` helper exercises a small product-spec fixture; it is
not reachable through the production CLI. Its generated fixture is not
evidence of a user's supplied documents.

## Supported content structures

DOCX exports every section's heading and text, section tables, and optional
requirement rows. XLSX exports an overview, a worksheet for each section,
section tables, optional requirement rows, and structured `model` data. An
XLSX table cell may contain an explicit formula object:

```json
{"formula": "=SUM(B2:B4)", "expected_value": 17}
```

LibreOffice must preserve the formula and cached result; the backend rejects
formula errors and external-workbook links. Workbooks without formulas are
recalculated and reported as having no formula results to verify; the exporter
does not invent spreadsheet calculations.

PPTX exports one slide per grounded section, preserving section text and
speaker notes. It supports the built-in slide layouts `cover`, `cards`, `grid`,
`statement`, `capabilities`, `workflow`, `stats`, `chart`, `closing`, and
`content`, optional confidentiality badges/subtitles, tables (up to 6 columns
and 10 rows per slide), and simple charts with text labels and finite numeric
values. Slides must map one-to-one to sections and contain the section's full
text. Unknown fields, unsupported cell/object values, unsupported layouts, and
unmapped slide content cause an explicit error rather than being silently
discarded.

## Pinned runtime and limitations

- Debian 12.9 slim is pinned by image digest; APT packages resolve against the
  Debian snapshot dated 2025-02-03.
- Python 3.11.2, python-docx 0.8.11, openpyxl 3.0.9, python-pptx 1.0.2
  (wheel hash-pinned), pypdf 3.4.1, LibreOffice Writer/Calc/Impress 7.4.7,
  Liberation fonts, and Poppler are pinned.
- DOCX/XLSX/PPTX package integrity, extracted text, PDF pages, and PNG previews
  are checked. XLSX formula expressions, cached values, and error cells are
  checked after LibreOffice recalculation. The agent must still review the
  preview images and conduct separate semantic, genre, and depth reviews.
- LibreOffice pagination, formatting, and formula support can differ from
  Microsoft Office. This container makes **no Office fidelity claim**.
- Macros, add-ins, pivot tables, complex templates, and unsupported structured
  slide/table/chart objects are not covered. The CLI does not generate prose,
  accept arbitrary Office content, or change a frozen plan/draft.
