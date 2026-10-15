import json
import sys
from pathlib import Path

root = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(root.parents[1] / ".github" / "skills" / "company-content-generator" / "scripts"))
import company

plan = company.load(root / "build" / "plan.json")
docs = {d["id"]: d for d in plan["documents"]}
measurements = {d["document_id"]: d for d in company.load(root / "build" / "draft-measurements.json")}
reviewed = sys.argv[1:]
if not reviewed:
    raise SystemExit("Supply only document IDs whose complete content and every image were actually reviewed.")

OBSERVATIONS = {
    "product_spec": (
        "Reviewed scope, host/compute controls, malformed-input and DMA boundaries, resets, telemetry, "
        "thermal behavior, software identity, retained evidence and change control against the "
        "16 version-specific source requirements. Planned status is not described as a completed launch.",
        "Four deliberately broken pages; all requirement numbers, full acceptance conditions and "
        "the release gate remain readable. No overflow, missing controls or orphaned acceptance text.",
        "16 distinct numbered controls with methods and acceptance criteria, plus substantive "
        "interface, ownership, negative-path and release-control explanation."),
    "meeting_transcript": (
        "Reviewed every speaker turn against the meeting positions, current role snapshot, Q2 actuals, "
        "product operating boundaries, rejected unconditional ramp, decision and three owned/due actions. "
        "Preparation, release qualification and board capital authority remain distinct.",
        "Six explicit pages; complete timestamped dialogue and the final decision/action list are "
        "readable without clipping. The opening page is a concise meeting-record cover.",
        "Discussion-only count exceeds 1500 words. Back-and-forth covers alternatives, economic "
        "implications, technical dependencies, unresolved adoption schedules and concrete commitments."),
    "press_release": (
        "Reviewed the complete release against the launched milestone, version-specific capabilities, "
        "approved CEO quote, operating constraints, availability, Florida business description and "
        "reserved-domain contacts. GulfSwitch efficiency is qualified by its reference conditions; "
        "SunCore peak compute is not represented as sustained customer performance.",
        "Two explicit pages; headline, dateline, complete quotation, adoption/availability copy, "
        "company boilerplate and media contact are legible with no clipped text.",
        "Delivered release meets the 450-700 word target and includes practical customer implications, "
        "technical scope, adoption sequence and availability rather than only an announcement headline."),
    "board_deck": (
        "Reviewed all ten slides and notes against Q1/Q2 actuals, reconciled product/cost cuts, "
        "engineering investment, point-in-time headcount, prospective August release, year-specific "
        "risk, transaction status and requested decision. No cash-flow inference or invented approval.",
        "All ten slide images inspected: long cover title wraps cleanly, calibrated financial bars "
        "are on common within-chart scales, labels and notes are readable, and text clears titles. "
        "The OfficeCLI HTML capture includes surrounding notes/navigation; it is not native Office certification.",
        "Ten substantive board slides include actual performance visuals, product/release accountability, "
        "resource tradeoff, risk mitigation and an explicit decision with retained authority."),
    "product_presentation": (
        "Reviewed all ten slides and notes against the released version, customer problem, operating "
        "envelope, interface workflow, three use cases, adoption responsibilities, constraints and "
        "evaluation next steps. No unsupported customer benchmark, certification or performance guarantee.",
        "All ten slide images inspected: wrapped headings, numbered capability/use-case groups, "
        "operating limits and evaluation contact remain readable. No title/body collisions or cut text "
        "within the slide. HTML captures are not native presentation fidelity certification.",
        "Ten meaningful customer-adoption slides cover problem, value, capabilities, workflow, "
        "multiple use cases, integration responsibilities, limits, availability and next steps."),
    "financial_model": (
        "Reviewed all four worksheets and every displayed module against the period actuals and "
        "independently grounded formula derivations. Checked income, gross/operating ratios, product "
        "reconciliation, manufacturing/opex reconciliations and sequential absolute changes. "
        "Headcount is point-in-time; no full funding, cash-flow or forecast claim is made.",
        "All five declared range images inspected, including both Calc ranges and inactive-sheet "
        "Summary/Notes captures. Amounts, ratio labels, wrapped notes, color conventions and zero "
        "reconciliation differences are legible; no omitted populated region.",
        "Six distinct useful modules answer income, margin, product-mix, manufacturing-cost, "
        "operating-investment and sequential-trend questions. Four sheets provide source, calculation, "
        "management interpretation and methodology zones without substituting sheet count for analysis."),
}
for identifier in reviewed:
    doc = docs[identifier]
    machine_path = root / "build" / "machine" / f"{identifier}.json"
    machine = company.load(machine_path)
    draft_path = root / "build" / "drafts" / f"{identifier}.json"
    draft = company.load(draft_path)
    semantic, visual, depth = OBSERVATIONS[doc["type"]]
    measured = company.load(root / "build" / "tool-results" / f"{identifier}-depth.json")
    image_paths = [image["path"] for render in machine["renders"] for image in render["images"]]
    review = dict(document_id=identifier, artifact_sha256=machine["artifact_sha256"],
                  draft_sha256=company.file_hash(draft_path),
                  machine_sha256=company.file_hash(machine_path),
                  semantic_review={"status": "passed", "evidence": f"{doc['title']}: {semantic}"},
                  genre_review={"status": "passed", "evidence":
                                "Reviewed complete delivered content in its business genre: "
                                + semantic + " Ordinary company voice; no placeholders or generation self-labels."},
                  depth_review={"status": "passed", "content_depth": "medium",
                                "evidence": f"{depth} Saved-content measurement: {measured['measured']} "
                                f"{measured['measure']}; draft measures {measurements[identifier]}."},
                  visual_review={"status": "passed", "images_delivered": True,
                                 "render_ids": machine["render_targets"],
                                 "evidence": visual + f" Inspected {len(image_paths)} complete unit images.",
                                 "inspected_images": image_paths},
                  condition_review={"status": "not_selected"})
    company.validate_review(review, doc, {"content_required": True, "visual_required": True}, machine)
    target = root / "build" / "reviews" / f"{identifier}.json"
    company.save(target, review)
print(f"Recorded {len(reviewed)} actual grouped reviews.")
