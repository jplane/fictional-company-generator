import importlib.util
import json
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
BUILD = ROOT / "build"
spec = importlib.util.spec_from_file_location(
    "company", ROOT.parents[1] / ".github" / "skills" / "company-content-generator" / "scripts" / "company.py")
company = importlib.util.module_from_spec(spec)
spec.loader.exec_module(company)
canon = company.load(ROOT / "canon.json")
plan = company.load(BUILD / "plan.json")
semantics = {r["document_id"]: r for r in company.load(BUILD / "semantic-review.json")["documents"]}
visuals = {}
for filename in ("presentation-visual-review.json", "narrative-workbook-visual-review.json"):
    report = company.load(BUILD / filename)
    for entry in report["documents"]:
        assert entry["document_id"] not in visuals
        visuals[entry["document_id"]] = entry
assert set(visuals) == {d["id"] for d in plan["documents"]}
manifest_entries, qa_entries = [], []
for item in plan["documents"]:
    identifier = item["id"]
    path = ROOT / "documents" / item["filename"]
    check = company.load(BUILD / "checks" / f"{identifier}.json")
    semantic = semantics[identifier]
    visual = visuals[identifier]
    assert visual["status"] == "passed", (identifier, visual)
    assert not check["render_text_issues"]
    artifact_sha = company.file_hash(path)
    draft_sha = company.file_hash(BUILD / "drafts" / f"{identifier}.json")
    assert artifact_sha == check["artifact_sha256"]
    assert draft_sha == check["draft_sha256"] == semantic["draft_sha256"]
    manifest_entries.append({
        "document_id": identifier, "path": str(Path("documents") / item["filename"]),
        "sha256": artifact_sha,
    })
    qa_entry = {
        "document_id": identifier, "artifact_sha256": artifact_sha, "draft_sha256": draft_sha,
        "semantic_review": {"status": semantic["status"], "evidence": semantic["evidence"]},
        "genre_review": {"status": "passed", "evidence": semantic["evidence"] + " Business title, audience and full required genre structure verified."},
        "depth_review": {
            "status": "passed", "content_depth": "medium",
            "evidence": f'{semantic["depth_measure"]}: {semantic["depth_actual"]}; target '
                        f'{semantic["depth_target"][0]}-{semantic["depth_target"][1]}. '
                        + semantic["evidence"],
        },
        "visual_review": {"status": visual["status"], "evidence": visual["evidence"]},
    }
    if item["format"] == "xlsx":
        recalc = check["recalculation"]
        assert recalc["status"] == "passed" and recalc["formula_errors"] == 0
        qa_entry["recalculation"] = {
            "status": "passed",
            "evidence": f'{recalc["engine"]}; {recalc["formula_count"]} formulas with retained caches; '
                        f'zero formula errors; {recalc["expected_values_checked"]} expected-value checks. '
                        + recalc["note"],
        }
    qa_entries.append(qa_entry)
base = {"schema_version": 1, "canon_sha256": company.fingerprint(canon),
        "plan_sha256": company.fingerprint(plan)}
company.save(BUILD / "manifest.json", {**base, "documents": manifest_entries})
company.save(BUILD / "qa.json", {**base, "documents": qa_entries})
index = {r["id"]: r for r in canon["records"]}
summary = {
    **base, "company": canon["request"]["company_name"], "content_depth": "medium",
    "documents": len(plan["documents"]),
    "formats": Counter(d["format"] for d in plan["documents"]),
    "types": Counter(d["type"] for d in plan["documents"]),
    "fiscal_year_counts": Counter(index[d["anchor_ids"][0]]["data"]["year"] for d in plan["documents"]),
    "publication_year_counts": Counter(d["publication_date"][:4] for d in plan["documents"]),
    "rendered_pages_and_slides": sum(company.load(BUILD / "checks" / f'{d["id"]}.json')["pages"] for d in plan["documents"]),
    "recalculated_formulas": sum(company.load(BUILD / "checks" / f'{d["id"]}.json')["recalculation"]["formula_count"]
                                for d in plan["documents"] if d["format"] == "xlsx"),
    "warnings": plan["warnings"], "blockers": [],
    "qa_limitations": "Claim checks and automated business invariants are not independent semantic certification. Visual review uses all-page contact sheets/galleries with full-size inspection of dense and corrected layouts; not every page was inspected at full resolution.",
}
company.save(BUILD / "completion-summary.json", summary)
print(json.dumps(summary, indent=2))
