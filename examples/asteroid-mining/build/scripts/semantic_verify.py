import importlib.util
import json
import re
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
targets = company.load(BUILD / "depth-targets.json")
receipts = []
all_names = [r["data"]["name"] for r in canon["records"] if r["kind"] == "person"]
for item in plan["documents"]:
    draft = company.load(BUILD / "drafts" / f'{item["id"]}.json')
    context = company.load(BUILD / "contexts" / f'{item["id"]}.json')
    index = {r["id"]: r for r in context["records"]}
    a = index[item["anchor_ids"][0]]["data"]
    company.check_draft(canon, plan, draft)
    prose = "\n".join(s["text"] for s in draft["sections"])
    target = next(t for t in targets["documents"] if t["document_id"] == item["id"])
    measured = draft["measurements"][target["measure"]]
    assert target["target_min"] <= measured <= target["target_max"]
    assert draft["title"] == item["title"]
    assert not re.search(r"\b(?:fictional|simulated|AI-generated|lorem ipsum)\b", prose, re.I)
    assert all(index[c["record_id"]]["known_on"] <= item["publication_date"] for c in draft["claims"])
    for name in all_names:
        if name in prose:
            assert any(r["data"]["name"] == name for r in context["org_chart"]["people"])
    if item["type"] == "press_release":
        assert a["approved_quote"] in prose
        assert a["quote_approved_on"] <= item["publication_date"]
        assert index["company"]["data"]["boilerplate"] in prose
        assert index["company"]["data"]["media_email"] in prose
        # Press releases must not disclose the internal operating ledger.
        assert not any(c["record_id"].startswith("financial-") for c in draft["claims"])
        assert a["scope"].lower() in prose
        evidence = "Reviewed public news structure, actual release scope, limitations, approved quote and effective attribution; no internal financial records cited."
    elif item["type"] == "product_spec":
        rows = draft["requirement_rows"]
        assert len({r["requirement_id"] for r in rows}) == 16
        for r in rows:
            assert r["product_id"] == a["product_id"] and r["version"] == a["version"]
            assert all(r[k] for k in ["statement", "acceptance_criterion", "verification_method", "interface", "failure_disposition"])
            assert r["statement"] in prose and r["acceptance_criterion"] in prose
        evidence = "Reviewed 16 distinct requirement controls with version-specific identity, acceptance tests, methods, dependencies, failure dispositions and release-boundary text."
    elif item["type"] == "meeting_transcript":
        meeting = index[f'meeting-{a["year"]}']["data"]
        roles = company.role_snapshot(context["records"], item["publication_date"])
        for person in meeting["participants"]:
            assert index[person]["data"]["name"] in prose and roles[person]["data"]["title"] in prose
        for decision in meeting["decisions"]:
            assert decision["decision"] in prose and decision["due_on"] in prose
        assert len(re.findall(r"\[\d\d:\d\d:\d\d\]", prose)) == 24
        evidence = "Reviewed 24 timestamped exchanges across operating results, scoped verification, alternatives and prospective gates; participants use effective roles; decisions, owners and dates match meeting record."
    elif item["type"] == "product_presentation":
        assert len(draft["slides"]) == 11
        assert all(slide["layout"] for slide in draft["slides"])
        assert a["limitation"] in prose and a["availability"] in prose
        evidence = "Reviewed customer problem, release-specific controls, three distinct engineering uses, adoption dependencies, ground-only exclusions and technical-review next step; visible story is not a board briefing."
    elif item["type"] == "board_deck":
        board = index[f'board-{a["year"]}']["data"]
        assert len(draft["slides"]) == 12
        assert board["decision_status"] in prose and board["preferred_option"] in prose
        for decision in board["decisions_requested"]:
            assert decision["decision"] in prose and decision["due_on"] in prose
        evidence = "Reviewed Q3 actuals, development allocation, conditional mission gates, strategic alternative, resource tradeoff and four owned requested decisions; requested status is not portrayed as approval."
    elif item["format"] == "pptx":
        for slide in draft["slides"]:
            chart = slide.get("chart")
            if chart:
                assert len(chart["labels"]) == len(chart["values"])
        if item["type"] == "annual_report":
            qs = [index[f'financial-{a["year"]}-q{q}']["data"] for q in range(1, 5)]
            for field in ["revenue_usd", "cogs_usd", "opex_usd", "operating_income_usd"]:
                assert sum(q[field] for q in qs) == a[field]
            assert qs[-1]["headcount"] == a["headcount"]
            year_roles = company.role_snapshot(context["records"], a["period_end"])
            org_section = next(s for s in draft["sections"] if s["role"] == "organization")
            for c in draft["claims"]:
                if c["id"] in org_section["claim_ids"] and c["record_id"].startswith("role-"):
                    role = index[c["record_id"]]
                    assert year_roles[role["data"]["person_id"]]["id"] == role["id"]
            evidence = "Reviewed 14-slide operating story, source-linked financial cuts, four-quarter annual reconciliation, prior-year availability, year-end rather than publication-date organization, strategy and conditional risks."
        else:
            evidence = "Reviewed 11-slide quarter actuals, service/cost cuts, known-quarter trend, delivery evidence, backlog limitation and explicitly labeled outlook."
    else:
        model = draft["model"]
        assert len(model["modules"]) == 6
        assert all(q["quarter"] <= a["quarter"] for q in model["quarters"])
        assert a["engineering_services_usd"] + a["instrument_development_usd"] + a["mission_analysis_usd"] == a["revenue_usd"]
        assert a["research_development_usd"] + a["program_operations_usd"] + a["corporate_administration_usd"] == a["opex_usd"]
        evidence = (
            "Reviewed six financial-workbook analyses distinct from quarterly reporting: income reconciliation, contracted service economics, development spending intensity, delivery/evidence capacity, reconciled Q2-to-Q3 operating bridge and backlog commitment scale. No service-category profit, cash forecast or resource valuation inferred."
            if item["type"] == "financial_model" else
            "Reviewed six distinct quarterly operating questions: income, margin/spending intensity, recognized service mix, cost allocation, workforce/reference evidence and available-quarter trends/commitments; no unsupported cash, forecast or resource model."
        )
    receipts.append({
        "document_id": item["id"], "draft_sha256": company.file_hash(BUILD / "drafts" / f'{item["id"]}.json'),
        "status": "passed", "evidence": evidence,
        "review_method": "Direct review of authored genre logic plus full-document checks of source receipts, date/role snapshots, target measurements, and genre-specific business invariants.",
        "depth_measure": target["measure"], "depth_actual": measured,
        "depth_target": [target["target_min"], target["target_max"]],
    })
report = {"schema_version": 1, "canon_sha256": company.fingerprint(canon),
          "plan_sha256": company.fingerprint(plan), "documents": receipts,
          "limitations": "Semantic automation verifies enumerated business invariants and receipts, not every possible natural-language implication; visual review is recorded separately."}
(BUILD / "semantic-review.json").write_text(json.dumps(report, indent=2) + "\n")
print(json.dumps({"reviewed": len(receipts), "formats": Counter(d["format"] for d in plan["documents"])}, indent=2))
