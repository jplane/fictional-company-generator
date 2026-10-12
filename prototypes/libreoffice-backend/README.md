# LibreOffice backend proof of concept

This isolated prototype exercises a bounded slice of the company's date-filtered
context/draft contract for exactly one planned DOCX product specification and one
planned XLSX traceability matrix. It uses the repository's standard-library
`company.py` helper through a read-only bind mount, a seeded test canon, and
four small requirement records. It never reads `examples/` corpora or any
user/company files. The Compose batch runtime does not mount the rest of the
repository; the Dev Container mounts the repository as an editable workspace.
The existing company-content workflow remains Microsoft-Office-backed and does
not invoke this prototype.

## Quick start

The host needs Docker Engine and the Docker Compose v2 plugin; it does not need
Python, Office, LibreOffice, or a package installer.

Create the output directory first so Docker does not create a root-owned host
directory. In PowerShell from the repository root:

```powershell
New-Item -ItemType Directory -Force prototypes\libreoffice-backend\output | Out-Null
```

On a POSIX shell, use `mkdir -p prototypes/libreoffice-backend/output`.
Then build the image and run the contract pilot:

```text
docker compose -f prototypes/libreoffice-backend/compose.yaml run --build --rm backend
```

It writes `canon.json`, the two planned documents under `documents/`, and
contract contexts, drafts, recalculation evidence, PDFs, and PNG page previews
under `build/` in `prototypes/libreoffice-backend/output/`. The pilot validates
both drafts against the repository helper, renders DOCX and XLSX to PDF,
checks extracted PDF text and pages, rasterizes every page for visual review,
and checks XLSX formulas, cached values, and error cells.

Run the smoke suite in an isolated temporary directory:

```text
docker compose -f prototypes/libreoffice-backend/compose.yaml run --build --rm backend python3 -m unittest discover -s tests -v
```

The test suite includes a real container non-root check, date-filtered contract
pilot, DOCX/XLSX PDF checks, cached-formula assertions, a deliberate formula
error, and checks that existing output files/directories are not overwritten.
The persisted `output/` directory is ignored by Git. Pilot execution refuses
any nonempty output directory. For another run, select a new child path:

```text
docker compose -f prototypes/libreoffice-backend/compose.yaml run --rm backend python3 -m office_backend --output /output/contract-pilot-2
```

On a native Linux host, match the container UID/GID to the directory owner when
needed:

```sh
POC_UID="$(id -u)" POC_GID="$(id -g)" docker compose -f prototypes/libreoffice-backend/compose.yaml run --build --rm backend
```

The default container identity is unprivileged UID/GID `10001`; Compose permits
overriding it with `POC_UID` and `POC_GID` for bind-mount permission differences.
In this Compose batch runtime, only the output directory is writable. The source
helper mount is read-only, and the rest of the repository is not mounted.

## Developer setup (Dev Container)

For interactive development, open the repository root in VS Code with the
Dev Containers extension and choose **Reopen in Container**. The repository-root
`.devcontainer/devcontainer.json` builds this same pinned image and opens
`prototypes/libreoffice-backend` as the workspace. The container runs as the
unprivileged `app` user; the Python extension is enabled, the repository is
mounted for editing, and the contract helper is mounted read-only at
`/app/company_contract.py`.

From the Dev Container's terminal, run the smoke suite and pilot from the
prototype workspace directory:

```sh
python3 -m unittest discover -s tests -v
python3 -m office_backend --output output/devcontainer-pilot
```

Generated files stay under the Git-ignored `output/` directory. Choose a new
output subdirectory for each pilot run; existing output is never overwritten.
The Compose quick start above remains available for a headless, narrowly
mounted batch run without mounting the repository workspace.

## Pinned runtime

- Debian 12.9 base image pinned by manifest digest:
  `debian:12.9-slim@sha256:12c396bd585df7ec21d5679bb6a83d4878bc4415ce926c9e5ea6426d23c60bdc`.
- Python 3.11.2, `python-docx` 0.8.11, `openpyxl` 3.0.9, `pypdf` 3.4.1,
  LibreOffice Writer/Calc 7.4.7, Liberation fonts, and Poppler tools are
  exact-pinned in `apt-packages.lock`, resolved against the Debian snapshot
  dated 2025-02-03.
- The prototype uses the repository helper for the canon, plan, contexts,
  claims, and draft checks. These are test-only artifacts, not a generated
  production corpus.

## Scope and limitations

This is contract/engine smoke evidence, not a production backend, content
approval, Microsoft Office QA receipt, or Office-compatibility claim. The
pilot has two basic requirements per specification—below the skill's
light-depth specification target of 6–10 distinct requirements—and deliberately
does not claim a depth review. The PDFs receive structural/text checks and page
previews, but no side-by-side visual comparison with Microsoft Word or Excel.
LibreOffice may paginate, format, or calculate differently from Office;
outcomes may also vary with fonts and supported functions.

The prototype does not cover PPTX, macros, add-ins, pivot tables, complex
templates, or Excel-specific functions. It remains disconnected from the
production content-generation workflow, whose Microsoft Office prerequisites
and backend selection are unchanged.
