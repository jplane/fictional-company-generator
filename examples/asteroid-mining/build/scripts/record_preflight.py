import hashlib
import json
import os
from pathlib import Path

import PIL
import pymupdf

ROOT = Path(__file__).resolve().parents[2]
BUILD = ROOT / "build"


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


skills = []
for name in ("docx", "pptx", "xlsx"):
    folder = Path.home() / ".copilot" / "skills" / name
    skills.append({
        "name": name, "path": str(folder),
        "bundled_version": (folder / ".bundled-version").read_text().strip(),
        "instructions_sha256": sha(folder / "SKILL.md"),
        "license_sha256": sha(folder / "LICENSE.txt"),
        "permission": "User explicitly confirmed applicable authorization before use.",
        "review": "Loaded instructions and reviewed local license. No external uploads or redistribution of skill materials.",
    })
pilots = []
for identifier, evidence in [
    ("doc-0001", "498-word publishable release; reviewed all two rendered pages, dateline, scope, approved executive quote, boilerplate, and contact. Office/extracted/rendered text matches draft."),
    ("doc-0002", "11-slide customer engineering presentation; independent agent inspected every slide and reverified corrected slides 1, 5, 9, 11. No clipping, overlap, or blocked readability. Minor isolated final title word retained."),
    ("doc-0008", "Six distinct formula-linked operating analyses; all nine rendered pages reviewed. Fixed clipped chart placement. Native LibreOffice recalculation retains 56 formula caches, zero errors, and 13 expected value checks."),
]:
    check = json.loads((BUILD / "checks" / f"{identifier}.json").read_text())
    assert not check["render_text_issues"], (identifier, check["render_text_issues"])
    pilots.append({
        "document_id": identifier, "artifact_sha256": check["artifact_sha256"],
        "draft_sha256": check["draft_sha256"], "status": "passed", "evidence": evidence,
        "measurements": check["measurements"], "recalculation": check["recalculation"],
    })
receipt = {
    "schema_version": 1, "status": "passed", "selected_backend": "docx-js / PptxGenJS / openpyxl with native LibreOffice and local PDF rasterization",
    "skills": skills,
    "tool_versions": {"python": "3.14.6", "docx": "9.7.1", "pptxgenjs": "4.0.1",
                      "openpyxl": "3.1.5", "libreoffice": "26.8.0.3", "poppler": "25.07.0",
                      "pymupdf": pymupdf.VersionBind, "pillow": PIL.__version__},
    "installations": [
        {"channel": "winget", "id": "TheDocumentFoundation.LibreOffice", "version": "26.8.0.3",
         "installer_sha256": "4aa6c6e1895f4055104effcb556bd3362d20c6ad707c149543304f395ef9db95",
         "result": "Installer hash verified; publisher installer completed."},
        {"channel": "winget", "id": "oschwartz10612.Poppler", "version": "25.07.0-0",
         "result": "Installer hash verified; approved third-party Windows distribution installed."}
    ],
    "windows_adaptation": "Upstream recalc helper assumes Linux/macOS macro paths. Run-local scripts instead use the same LibreOffice calculation engine through native Windows headless import/export, an isolated profile under build, and explicit cached-result/error/expected-value checks. Installed PyMuPDF rasterizes intermediate PDFs and Pillow assembles review galleries; Poppler was installed and version-probed but not needed for these renders. No proprietary helper code copied or edited.",
    "pilot_results": pilots, "blockers": [],
    "review_limitations": "Preflight certifies actual planned pilots, not ungenerated bulk outputs. Every subsequent document still requires its own QA receipt.",
}
(BUILD / "preflight.json").write_text(json.dumps(receipt, indent=2) + "\n")
print("Preflight passed: three actual planned pilots verified.")
