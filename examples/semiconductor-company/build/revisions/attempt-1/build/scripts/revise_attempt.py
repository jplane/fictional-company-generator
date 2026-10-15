import json
import shutil
import sys
from pathlib import Path

root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(root.parents[1] / ".github" / "skills" / "company-content-generator" / "scripts"))
import company

build = root / "build"
archive = build / "revisions" / "attempt-1"
if archive.exists():
    raise SystemExit("Initial attempt already preserved; do not reset retry evidence.")
archive.mkdir(parents=True)
shutil.copy2(root / "canon.json", archive / "canon.json")
shutil.move(str(root / "documents"), archive / "documents")
(root / "documents").mkdir()
(archive / "build").mkdir()
for child in list(build.iterdir()):
    if child.name == "revisions":
        continue
    if child.name == "scripts":
        shutil.copytree(child, archive / "build" / child.name)
    else:
        shutil.move(str(child), archive / "build" / child.name)
for name in ["preflight.json", "request.json", "setup-evidence.json"]:
    shutil.copy2(archive / "build" / name, build / name)
for child in (archive / "build").glob("officecli-preflight*"):
    shutil.copy2(child, build / child.name)
canon = company.load(root / "canon.json")
plan = company.load(archive / "build" / "plan.json")
for doc in plan["documents"]:
    doc["revision"] = 2
company.validate_plan(canon, plan)
def write(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2) + "\n", encoding="utf-8")
write(build / "plan.json", plan)
for doc in plan["documents"]:
    write(build / "contexts" / f"{doc['id']}.json", company.get_context(canon, plan, doc["id"]))
write(build / "revision-lineage.json", {
    "revision": 2, "business_canon_unchanged": True, "user_approved_preservation": True,
    "prior_attempt": "build/revisions/attempt-1",
    "corrections": [
        "Define Word paragraph styles before assigning them.",
        "Increase slide label and footer text boxes to the exact schema issue minima.",
        "Select each XLSX worksheet tab explicitly; use unmerged corner cells.",
    ],
    "retry_policy": "This is the single corrective revision. Repeated operation failures stop the run.",
    "scope": "All six affected families require complete corrective content and visual review.",
    "additional_corpus_qa": "not_requested",
})
print("Preserved the first attempt; prepared explicit document revision 2 with unchanged business facts.")
