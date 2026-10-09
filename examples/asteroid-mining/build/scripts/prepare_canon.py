import calendar
import importlib.util
import json
from collections import Counter
from datetime import date, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
if (ROOT / "build" / "plan.json").exists():
    raise RuntimeError("Canon and plan are frozen. Use a new output revision rather than rebuilding this corpus.")
HELPER = ROOT.parents[1] / ".github" / "skills" / "company-content-generator" / "scripts" / "company.py"
spec = importlib.util.spec_from_file_location("company", HELPER)
company = importlib.util.module_from_spec(spec)
spec.loader.exec_module(company)
request = json.loads((ROOT / "build" / "request.json").read_text())
records = []


def add(identifier, kind, when, data, refs=(), effective=None):
    record = dict(id=identifier, kind=kind, known_on=when, effective_on=effective or when,
                  data=data, refs=list(dict.fromkeys(refs)))
    records.append(record)
    return record


add("company", "company", "2013-01-01", {
    "name": "AstroDig LLC", "industry": request["industry"], "currency": "USD",
    "fiscal_year_end": "12-31", "founded_on": "2013-01-01", "headquarters": "Boulder, Colorado",
    "mission": "Develop robotic prospecting and extraction systems for asteroid resources.",
    "business_model": "Earn aerospace engineering, instrument-development, and mission-analysis revenue while developing asteroid resource technology.",
    "mining_goal_year": 2032, "goal_status": "aspirational and contingent on flight validation, funding, launch access, and mission authorization",
    "resource_focus": "water-bearing material for potential in-space use; no commercial extraction revenue assumed",
    "media_email": "media@astrodig.example", "media_contact_id": "person-jordan",
    "boilerplate": "AstroDig LLC is an aerospace engineering company based in Boulder, Colorado. Founded in 2013, the company develops mission-analysis tools, robotic prospecting instruments, and controlled extraction test systems. Its long-term objective is an asteroid resource demonstration by 2032, subject to technical, financial, and mission-readiness gates.",
}, ["person-jordan"])

people = [
    ("avery", "Avery Rowan", "Chief Executive Officer", "executive", None),
    ("morgan", "Morgan Vale", "Chief Financial Officer", "finance", "avery"),
    ("riley", "Riley Arden", "Chief Technology Officer", "engineering", "avery"),
    ("jordan", "Jordan Ellis", "Director of Communications", "operations", "avery"),
    ("casey", "Casey Marlow", "Finance Director", "finance", "morgan"),
    ("quinn", "Quinn Mercer", "Director of Mission Systems", "engineering", "riley"),
    ("sasha", "Sasha Linden", "Head of Verification", "engineering", "riley"),
    ("devon", "Devon Hale", "Director of Program Operations", "operations", "avery"),
]
for key, name in [("executive", "Executive Leadership"), ("finance", "Finance"),
                  ("engineering", "Mission Systems Engineering"), ("operations", "Program Operations")]:
    add(f"dept-{key}", "department", "2013-01-01", {"name": name})
for key, name, title, dept, manager in people:
    add(f"person-{key}", "person", "2013-01-01", {"name": name})
    refs = [f"person-{key}", f"dept-{dept}"] + ([f"person-{manager}"] if manager else [])
    add(f"role-{key}-2013", "role", "2013-01-01",
        {"person_id": f"person-{key}", "department_id": f"dept-{dept}",
         "title": title, "reports_to": f"person-{manager}" if manager else None}, refs)
add("role-casey-2020", "role", "2020-01-01",
    {"person_id": "person-casey", "department_id": "dept-finance",
     "title": "Chief Financial Officer", "reports_to": "person-avery"},
    ["person-casey", "dept-finance", "person-avery"])
add("role-morgan-2020", "role", "2020-01-01",
    {"person_id": "person-morgan", "department_id": "dept-executive",
     "title": "Chief Strategy Officer", "reports_to": "person-avery"},
    ["person-morgan", "dept-executive", "person-avery"])
add("role-quinn-2023", "role", "2023-01-01",
    {"person_id": "person-quinn", "department_id": "dept-engineering",
     "title": "Vice President of Mission Systems", "reports_to": "person-riley"},
    ["person-quinn", "dept-engineering", "person-riley"])
for key, name, purpose in [
    ("scoutmap", "ScoutMap", "Offline mission analysis and target-screening software for aerospace engineering teams."),
    ("regolithworks", "RegolithWorks", "Ground-test instrumentation and closed-chamber extraction benches; not an operational asteroid mine.")]:
    add(f"product-{key}", "product", "2013-01-01",
        {"name": name, "owner_id": "person-quinn", "purpose": purpose},
        ["person-quinn"])

phases = [
    ("Engineering foundation", "mission-analysis feasibility", "reference-data consistency",
     "Establish traceable engineering baselines before committing to hardware.", "No hardware flight heritage exists."),
    ("Target-screening discipline", "trajectory screening", "input uncertainty",
     "Separate reachable targets from targets with verified resource potential.", "Catalog rankings do not establish recoverable resources."),
    ("Resource-characterization studies", "spectral planning", "ambiguous composition",
     "Carry competing composition interpretations into mission trade studies.", "Remote signatures are not direct assays."),
    ("Closed-chamber bench commissioning", "thermal-vacuum commissioning", "chamber contamination",
     "Establish repeatable blank runs before interpreting recovered material.", "Ground chambers cannot reproduce every small-body condition."),
    ("Capture and containment iteration", "capture-cartridge testing", "material escape",
     "Prefer contained recovery over an open collection path.", "Capture efficiency is limited to controlled test materials."),
    ("Autonomy-in-the-loop development", "autonomy rehearsal", "unsafe state transitions",
     "Make abort behavior deterministic before expanding automated sequences.", "Hardware-in-the-loop results are not flight qualification."),
    ("Integrated payload architecture", "payload interface integration", "mass and power growth",
     "Control interface revisions before combining independently tested subsystems.", "Integrated ground prototypes remain development hardware."),
    ("Funding reset and program focus", "reduced-scope integration", "schedule compression",
     "Preserve verification work while deferring optional feature expansion.", "The revised integration scope does not accelerate the mining date."),
    ("Navigation and contact rehearsal", "contact-sequence rehearsal", "navigation uncertainty",
     "Treat loss of contact confidence as a hold condition rather than a reason to continue.", "No asteroid-contact demonstration has been completed."),
    ("Volatile-handling controls", "sealed volatile-handling trials", "sample contamination",
     "Record every custody handoff and chamber condition with the recovered sample.", "Ground recovery does not prove in-space storage feasibility."),
    ("Mission concept consolidation", "prospecting mission concept", "launch access",
     "Size the first mission around information acquisition rather than commercial output.", "Launch services are not booked."),
    ("Flight-interface engineering", "flight-interface development", "supplier qualification",
     "Freeze external interfaces only after verified supplier configuration evidence.", "Flight-interface work is not a launch approval."),
    ("Prospecting architecture review", "prospecting payload review", "resource uncertainty",
     "Require direct characterization evidence before authorizing extraction-scale expansion.", "No mining reserve or commercial resource claim is made."),
    ("Demonstration readiness planning", "2032 demonstration gate review", "funding and flight dependencies",
     "Retain the 2032 goal while making each intermediate commitment conditional.", "Mining has not begun and no mission approval is represented."),
]
revenues = [600000, 1200000, 2200000, 3600000, 5600000, 8200000, 11500000,
            10800000, 15200000, 21000000, 28000000, 36500000, 47000000, 59500000]
headcounts = [8, 14, 22, 32, 45, 60, 77, 74, 93, 118, 145, 172, 204, 238]

req_defs = {
    "regolithworks": [
        ("Containment", "The test cartridge shall remain closed throughout the collection sequence.", "Inspect seal position before and after every run; any open state fails the run.", "mechanical inspection", "capture-cartridge", "Seal wear can invalidate collection evidence."),
        ("Pressure envelope", "The chamber controller shall inhibit a sequence outside the approved pressure envelope.", "Inject out-of-envelope pressure and observe inhibition before actuator enable.", "fault injection", "pressure-sensor", "A failed sensor must not be interpreted as nominal pressure."),
        ("Thermal limit", "The heater controller shall stop heating when the configured temperature limit is reached.", "Command heating at the configured limit and verify shutdown within one control cycle.", "controller test", "temperature-sensor", "Temperature overshoot requires a documented disposition."),
        ("Power interlock", "Actuator power shall require an independent enabled interlock.", "Remove interlock permission and verify zero actuator-enable commands.", "electrical test", "power-interface", "Software permission alone is insufficient."),
        ("Emergency stop", "An emergency-stop request shall supersede an active extraction command.", "Issue emergency stop during each sequence state and verify a latched stop.", "state-transition test", "safety-controller", "Reset shall not resume the interrupted sequence."),
        ("Sample custody", "Every sample container shall carry a unique identifier linked to its run.", "Reconcile all collected containers to run identifiers with no duplicates.", "records audit", "custody-register", "Unidentified material is excluded from reported recovery."),
        ("Blank control", "Each material campaign shall include an empty-cartridge blank run.", "Verify blank-run completion and contamination disposition before material runs.", "campaign review", "chamber-cleaning", "A contaminated blank invalidates subsequent interpretation."),
        ("Mass accounting", "The run record shall include input, recovered, retained, and unaccounted mass.", "Confirm the four quantities reconcile within the approved balance tolerance.", "balance reconciliation", "calibrated-balance", "Unaccounted mass is not recovered product."),
        ("Time synchronization", "Controller and instrument events shall share a synchronized run clock.", "Verify clock offset is no greater than 100 milliseconds in the acceptance run.", "timestamp comparison", "run-clock", "Misaligned events can reverse apparent causality."),
        ("Telemetry continuity", "The controller shall flag any telemetry gap longer than two seconds.", "Inject a three-second gap and verify a flagged run-quality event.", "communications test", "telemetry-recorder", "A flagged gap requires review before accepting results."),
        ("Configuration control", "The run header shall record controller, cartridge, and sensor revisions.", "Verify complete revision fields on every accepted run.", "configuration audit", "revision-register", "Mixed revisions prevent comparable campaign results."),
        ("Calibration freshness", "Mass and temperature measurements shall use current approved calibrations.", "Reject an expired calibration during pre-run checks.", "negative test", "calibration-register", "Calibration drift can be mistaken for extraction improvement."),
        ("Recovery restart", "Restart after an abort shall require a new pre-run checklist.", "Attempt direct resume after abort and verify rejection.", "fault recovery test", "checklist-register", "An interrupted run shall retain its original failed status."),
        ("Data export", "Exported run data shall include units, channel names, and quality flags.", "Reimport a completed run and compare all channel metadata and flags.", "round-trip test", "data-schema", "Numeric values without units are not acceptable evidence."),
        ("Operator authority", "Only an authorized operator shall acknowledge a safety hold.", "Attempt acknowledgment with a read-only account and verify denial.", "access-control test", "operator-register", "A hold is not cleared by changing the view."),
        ("Release evidence", "A development release shall include a signed verification matrix.", "Verify every requirement has a method, result, owner, and disposition.", "release review", "verification-matrix", "Open critical failures block release to the next test campaign."),
    ],
    "scoutmap": [
        ("Dataset identity", "Each analysis shall record the input catalog name, revision, and ingestion date.", "Open a saved analysis and verify that all three identifiers are retained.", "metadata inspection", "catalog-interface", "A changed catalog can alter ranking independently of algorithm changes."),
        ("Units", "The importer shall reject mixed units that have not been explicitly mapped.", "Submit an unmapped mixed-unit dataset and verify an actionable rejection.", "negative import test", "unit-map", "Implicit conversion is not permitted."),
        ("Missing data", "A missing required input shall produce an excluded-target reason.", "Remove a required field and verify exclusion with the field name.", "negative input test", "input-schema", "A missing value shall not be treated as zero."),
        ("Screening constraints", "Each run shall retain the constraint set used for target screening.", "Reopen a run and compare every saved constraint to its exported record.", "round-trip test", "constraint-register", "Unsaved edits shall not alter accepted results."),
        ("Uncertainty", "A target ranking shall display its input-uncertainty classification.", "Verify classifications remain attached to all targets after sorting.", "display test", "uncertainty-model", "Ranking position alone is not resource evidence."),
        ("Composition alternatives", "Analysts shall be able to retain more than one composition interpretation.", "Save two interpretations and verify both remain separately addressable.", "scenario test", "interpretation-register", "Alternative interpretations shall not be averaged silently."),
        ("Reproducibility", "Identical inputs and configuration shall produce identical ordered outputs.", "Compare two runs byte-for-byte after excluding creation timestamps.", "repeat-run test", "solver-configuration", "Nondeterministic ordering blocks comparison acceptance."),
        ("Version boundary", "An analysis shall record the solver and scoring-model versions.", "Inspect a completed export and confirm both version identifiers.", "export inspection", "version-register", "A version upgrade shall not rewrite historical analyses."),
        ("Feasibility labels", "Screening results shall distinguish feasible, excluded, and unresolved targets.", "Exercise all three states and verify separate labels and reasons.", "state coverage", "screening-engine", "Unresolved shall not be presented as feasible."),
        ("Sensitivity", "Analysts shall be able to compare approved constraint variations.", "Run two named constraint sets and verify a labeled comparison.", "comparison test", "constraint-register", "Sensitivity outputs are not mission commitments."),
        ("Audit log", "Accepted analyses shall retain author, reviewer, and acceptance time.", "Verify an accepted analysis has all three fields and a review disposition.", "records audit", "review-register", "Unreviewed runs retain draft analysis status."),
        ("Export schema", "The exporter shall produce a versioned tabular schema.", "Validate exported columns against the registered schema version.", "schema validation", "export-schema", "A changed column name requires a schema revision."),
        ("Range validation", "Out-of-range values shall be rejected with the target identifier.", "Inject an out-of-range input and verify the identifier in the error.", "negative test", "input-schema", "The importer shall not silently clamp data."),
        ("Offline operation", "Accepted analysis workflows shall run without an external network dependency.", "Disconnect network access and complete a reference analysis.", "offline test", "local-runtime", "An optional catalog update is not part of acceptance."),
        ("Comparison integrity", "A comparison shall identify changes in both inputs and configuration.", "Change one input and one setting and verify both change classes.", "change-detection test", "comparison-engine", "An apparent improvement requires a traceable changed variable."),
        ("Release gate", "A development release shall include reviewed reference-case results.", "Verify reference cases, expected outcomes, discrepancies, and approvals.", "release review", "reference-suite", "Open discrepancies require disposition before release."),
    ],
}

for term, definition in [
    ("NEA", "Near-Earth asteroid; used here as a target category, not as a statement of resource ownership."),
    ("TRL", "Technology readiness level; AstroDig does not assign a level without an approved assessment."),
    ("Ground development release", "A configuration approved for a defined ground-test or analysis campaign; not flight qualification."),
    ("Recovery fraction", "Recovered test-material mass divided by input test-material mass for a specified controlled run."),
    ("Operating income", "Recognized revenue less cost of goods sold and operating expenses."),
    ("Contracted backlog", "Remaining contracted engineering work at a stated date; not recognized revenue or a guarantee of future collection."),
    ("2032 demonstration", "A prospective robotic asteroid-resource demonstration, contingent on preceding technical and business gates."),
]:
    add("term-" + str(len(records)), "glossary", "2013-01-01", {"term": term, "definition": definition})

for i, phase in enumerate(phases):
    year = 2013 + i
    theme, campaign, risk, rationale, limitation = phase
    version = f"{i + 1}.0"
    year_refs = ["person-avery", "person-riley", "person-quinn", "person-sasha", "person-devon"]
    annual_revenue = revenues[i]
    annual_cogs = annual_revenue * 38 // 100
    annual_opex = annual_revenue * 65 // 100 + 1400000 + i * 720000
    quarters = []
    for q, fraction in enumerate([20, 23, 26, 31], 1):
        rev = annual_revenue * fraction // 100
        cogs = annual_cogs * fraction // 100
        opex = annual_opex * fraction // 100
        eng = rev * (75 - min(i, 13)) // 100
        inst = rev * (15 + min(i, 13)) // 100
        analysis = rev - eng - inst
        rd = opex * 72 // 100
        ops = opex * 18 // 100
        admin = opex - rd - ops
        end = date(year, q * 3, calendar.monthrange(year, q * 3)[1])
        known = (end + timedelta(days=15)).isoformat()
        hc = max(8, headcounts[i] - (4 - q) * max(1, headcounts[i] // 18))
        d = dict(year=year, quarter=q, period_end=end.isoformat(), revenue_usd=rev,
                 cogs_usd=cogs, opex_usd=opex, operating_income_usd=rev - cogs - opex,
                 headcount=hc, engineering_services_usd=eng, instrument_development_usd=inst,
                 mission_analysis_usd=analysis, research_development_usd=rd,
                 program_operations_usd=ops, corporate_administration_usd=admin,
                 contracted_backlog_usd=rev * (5 + q), completed_work_packages=6 + i * 3 + q,
                 accepted_reference_runs=12 + i * 7 + q * 3,
                 rejected_reference_runs=2 + (i + q) % 5,
                 development_focus=campaign, principal_risk=risk,
                 management_interpretation=rationale, limitation=limitation,
                 revenue_policy="Recognize engineering and development revenue against accepted contracted work; do not recognize projected asteroid sales.",
                 cost_driver="Research and verification capacity is maintained ahead of mission deployment.",
                 outlook="Continue the approved ground-development scope; no asteroid commodity revenue is included in the operating outlook.")
        quarters.append(d)
        add(f"financial-{year}-q{q}", "financial", known, d)
    annual = dict(year=year, quarter=None, period_end=f"{year}-12-31",
                  headcount=quarters[-1]["headcount"])
    for field in ["revenue_usd", "cogs_usd", "opex_usd", "operating_income_usd",
                  "engineering_services_usd", "instrument_development_usd", "mission_analysis_usd",
                  "research_development_usd", "program_operations_usd", "corporate_administration_usd",
                  "completed_work_packages", "accepted_reference_runs", "rejected_reference_runs"]:
        annual[field] = sum(q[field] for q in quarters)
    annual.update({k: quarters[-1][k] for k in ["contracted_backlog_usd", "development_focus",
                  "principal_risk", "management_interpretation", "limitation", "revenue_policy", "cost_driver", "outlook"]})
    add(f"financial-{year}-annual", "financial", f"{year+1}-02-15", annual,
        [f"financial-{year}-q{q}" for q in range(1, 5)])
    known = f"{year}-02-01"
    gates = [
        {"year": 2028, "gate": "Prospecting spacecraft design review", "status": "planned", "dependency": "verified payload interfaces and approved funding"},
        {"year": 2030, "gate": "Direct asteroid characterization", "status": "planned", "dependency": "launch access, mission authorization, and spacecraft performance"},
        {"year": 2032, "gate": "Resource extraction demonstration", "status": "goal", "dependency": "characterization evidence, extraction readiness, and funded mission operations"},
    ]
    add(f"strategy-{year}", "strategy", known, {
        "year": year, "title": theme, "campaign": campaign, "decision_rationale": rationale,
        "principal_risk": risk, "limitation": limitation, "gates": gates,
        "revenue_source": "engineering services, instrument development, and mission analysis",
        "mining_revenue_status": "none", "flight_status": "no completed asteroid mining mission",
        "customer_personas": ["mission-design teams", "instrument-development teams", "verification managers"],
        "capital_priority": "verification capability and controlled development hardware",
        "alternative": "Expand optional functionality before closing verification evidence",
        "alternative_disposition": "Defer optional scope until reference evidence is accepted",
        "release_decision": "Authorize the defined ground-development campaign only",
        "release_authority_id": "person-riley",
        "scope_owner_id": "person-quinn", "verification_owner_id": "person-sasha",
        "operations_owner_id": "person-devon",
        "control": "Stop expansion when a critical verification discrepancy remains open.",
        "quarterly_priorities": [
            "Reconcile requirements and preserve the reference configuration.",
            "Complete controlled reference trials and record discrepancies.",
            "Review campaign evidence and close owned corrective actions.",
            "Consolidate accepted results into the next design baseline.",
        ],
    }, year_refs)
    for product in ("regolithworks", "scoutmap"):
        name = "RegolithWorks" if product == "regolithworks" else "ScoutMap"
        release_date = f"{year}-08-15"
        capability = campaign if product == "regolithworks" else f"{campaign} analysis support"
        d = {
            "product_id": f"product-{product}", "version": version, "status": "ground_development_release",
            "year": year, "title": f"{name} {version} Ground Development Release",
            "campaign": capability, "public_on": release_date, "location": "Boulder, Colorado",
            "scope": "Controlled ground-test use" if product == "regolithworks" else "Offline analysis and review use",
            "not_in_scope": "Flight qualification, asteroid resource certification, commercial mining, or mission approval",
            "release_rationale": rationale, "principal_risk": risk, "limitation": limitation,
            "approved_quote": f"Our {theme.lower()} work is intended to make the next decision more evidence-based. We are releasing a defined ground-development configuration, not declaring an asteroid mining capability. The 2032 objective remains subject to the evidence and resources needed at each gate.",
            "quote_speaker_id": "person-avery", "quote_approved_on": release_date,
            "availability": "Available to contracted engineering teams through a scoped technical review.",
            "adoption_requirements": ["approved input or test-material specification", "named technical reviewer", "configuration-controlled reference case"],
            "next_step": "Request a technical review through media@astrodig.example; define the reference case and acceptance evidence before campaign access.",
            "workflow": ["Register configuration and inputs", "Run approved reference checks", "Execute bounded campaign", "Review evidence and discrepancies", "Accept or hold the next configuration"],
            "use_cases": [
                {"name": "Mission trade study", "task": "Compare candidate operating assumptions without asserting resource recoverability.", "user": "mission-design teams"},
                {"name": "Payload verification campaign", "task": "Link a controlled test sequence to acceptance evidence and configuration identity.", "user": "instrument-development teams"},
                {"name": "Readiness review", "task": "Identify open discrepancies and distinguish accepted evidence from unresolved dependencies.", "user": "verification managers"},
            ],
            "release_controls": ["critical discrepancies block progression", "ground and flight status remain separate", "changes require reviewed evidence"],
        }
        add(f"milestone-{product}-{year}", "milestone", release_date, d,
            [f"product-{product}", "person-avery", f"strategy-{year}"])
        for j, (group, statement, acceptance, method, interface, failure) in enumerate(req_defs[product], 1):
            add(f"requirement-{product}-{year}-{j:02}", "requirement", release_date,
                {"product_id": f"product-{product}", "version": version,
                 "requirement_id": f"{'RW' if product == 'regolithworks' else 'SM'}-{j:03}",
                 "group": group, "statement": statement, "acceptance_criterion": acceptance,
                 "verification_method": method, "interface": interface, "failure_disposition": failure,
                 "owner_id": "person-sasha", "status": "required for development release",
                 "traceability": f"{name} {version} controlled acceptance baseline"},
                [f"product-{product}", f"milestone-{product}-{year}", "person-sasha"])
    finance_id = "person-morgan" if year < 2020 else "person-casey"
    meeting_when = f"{year}-10-15"
    due = f"{year}-11-15"
    quarter_id = f"financial-{year}-q3"
    decisions = [
        {"topic": "Scope", "decision": f"Retain {campaign} as the current development priority.", "owner_id": "person-quinn", "due_on": due},
        {"topic": "Verification", "decision": "Close critical discrepancies before adding optional campaign scope.", "owner_id": "person-sasha", "due_on": due},
        {"topic": "Finance", "decision": "Maintain service delivery and report development spending separately from any prospective mining economics.", "owner_id": finance_id, "due_on": due},
        {"topic": "2032 goal", "decision": "Retain the conditional demonstration goal; do not represent future milestones as approved missions.", "owner_id": "person-avery", "due_on": due},
    ]
    positions = [
        {"speaker_id": "person-avery", "position": rationale, "concern": "A calendar target must not replace an evidence-based gate."},
        {"speaker_id": finance_id, "position": "Engineering revenue supports a portion of development spending, but operating losses remain visible.", "concern": "Backlog is not cash and expected asteroid sales are not operating revenue."},
        {"speaker_id": "person-riley", "position": "Prefer a bounded, repeatable reference campaign over a broader unverified feature set.", "concern": risk},
        {"speaker_id": "person-quinn", "position": f"Keep the current product scope focused on {campaign}.", "concern": limitation},
        {"speaker_id": "person-sasha", "position": "Keep rejected trials in the evidence record and require owned discrepancy dispositions.", "concern": "Removing failed runs would distort readiness evidence."},
        {"speaker_id": "person-devon", "position": "Sequence service commitments and campaign work against the same named resources.", "concern": "A booked engineering commitment cannot also be counted as free development capacity."},
    ]
    add(f"meeting-{year}", "meeting", meeting_when,
        {"title": f"{theme} - Q3 Leadership Review", "year": year, "meeting_on": meeting_when,
         "participants": [p["speaker_id"] for p in positions], "positions": positions,
         "agenda": ["Q3 operating results and delivery capacity", f"{campaign} scope and verification", "Strategic alternatives and resource allocation", "2032 goal dependencies and action closure"],
         "decisions": decisions, "alternative": "Expand optional functionality before verification closure",
         "alternative_outcome": "Deferred to protect reference evidence and contracted delivery",
         "financial_record_id": quarter_id, "strategy_record_id": f"strategy-{year}",
         "product_milestone_id": f"milestone-regolithworks-{year}"},
        [p["speaker_id"] for p in positions] + [quarter_id, f"strategy-{year}", f"milestone-regolithworks-{year}"])
    add(f"board-{year}", "board_review", meeting_when,
        {"title": f"{theme} - Capital and Readiness Review", "year": year,
         "decisions_requested": decisions, "strategy_record_id": f"strategy-{year}",
         "financial_record_id": quarter_id, "alternative": "Expand optional functionality before verification closure",
         "preferred_option": "Maintain bounded development and contracted delivery",
         "tradeoff": "Defer optional feature breadth to preserve verification quality and service capacity",
         "decision_status": "requested; not represented as approved in the board briefing"},
        [quarter_id, f"strategy-{year}"] + [d["owner_id"] for d in decisions])

canon = {"schema_version": 1, "request": request, "records": records}
company.validate_canon(canon)
(ROOT / "canon.json").write_text(json.dumps(canon, indent=2) + "\n", encoding="utf-8")
baseline = company.create_plan(canon)
company.save(ROOT / "build" / "baseline-plan.json", baseline)
index = {r["id"]: r for r in records}
documents = []


def document(kind, fmt, anchor_id, title, when=None):
    anchor = index[anchor_id]
    when = when or anchor["known_on"]
    documents.append(dict(id=f"doc-{len(documents)+1:04}", type=kind, format=fmt,
                          title=f"AstroDig LLC - {title}",
                          filename=company.document_filename(f"AstroDig LLC - {title}", when, fmt),
                          publication_date=when, audience=company.TYPES[kind][2],
                          anchor_ids=[anchor_id], revision=1, content_depth="medium",
                          required_sections=company.CONTENT_SECTIONS[kind]))


for year in range(2013, 2027):
    i = year - 2013
    version = f"{i+1}.0"
    main = f"milestone-regolithworks-{year}"
    software = f"milestone-scoutmap-{year}"
    theme = phases[i][0]
    document("press_release", "docx", main, f"Releases RegolithWorks {version} for {phases[i][1].title()}")
    document("product_presentation", "pptx", main, f"RegolithWorks {version} - Ground Campaign Overview")
    document("product_spec", "docx", main, f"RegolithWorks {version} - Controlled Test Specification")
    document("product_spec", "docx" if i < 8 else "xlsx", software,
             f"ScoutMap {version} - Analysis and Verification Specification")
    document("meeting_transcript", "docx", f"financial-{year}-q3",
             f"{theme} - Q3 Leadership Review Transcript")
    document("board_deck", "pptx", f"financial-{year}-q3",
             f"{theme} - Capital and Readiness Briefing")
    document("annual_report", "pptx", f"financial-{year}-annual",
             f"{year} Annual Business and Financial Review")
    for q in (1, 2, 3):
        # The first eight Q3 reports are visual briefings; the rest are analytical workbooks.
        fmt = "pptx" if q == 3 and i < 8 else "xlsx"
        document("quarterly_report", fmt, f"financial-{year}-q{q}",
                 f"Q{q} {year} Operating Results")
    if year >= 2017:
        document("financial_model", "xlsx", f"financial-{year}-q3",
                 f"Q3 {year} Service Economics and Development Spending")

documents.sort(key=lambda d: (d["publication_date"], d["id"]))
plan = dict(schema_version=1, canon_sha256=company.fingerprint(canon), documents=documents,
            warnings=[
                "Fiscal-year 2013-2026 coverage includes a reporting publication tail through February 15, 2027.",
                "Full-year 2026 and later publication dates belong to the requested invented history; they are not real-world actuals available on October 8, 2026.",
                "Q4 results appear in annual reporting rather than separate Q4 report deliverables.",
                "The 2032 mining demonstration remains prospective; no completed mining mission, resource reserve, or commodity-sales revenue is claimed.",
            ])
company.validate_plan(canon, plan)
company.save(ROOT / "build" / "plan.json", plan)
targets = []
for d in documents:
    kind, fmt = d["type"], d["format"]
    measure, lower, upper = {
        "press_release": ("delivered_words", 450, 700),
        "meeting_transcript": ("discussion_words", 1500, 2500),
        "product_spec": ("distinct_testable_requirements", 16, 20),
        "product_presentation": ("meaningful_slides", 10, 13),
        "annual_report": ("meaningful_slides", 14, 18),
        "board_deck": ("meaningful_slides", 12, 15),
        "quarterly_report": ("meaningful_slides", 10, 13) if fmt == "pptx" else ("useful_analysis_modules", 6, 6),
        "financial_model": ("useful_analysis_modules", 6, 6),
    }[kind]
    targets.append(dict(document_id=d["id"], content_depth="medium", measure=measure,
                        target_min=lower, target_max=upper,
                        coverage=company.CONTENT_SECTIONS[kind],
                        adaptation_rationale=None))
company.save(ROOT / "build" / "depth-targets.json",
             dict(schema_version=1, plan_sha256=company.fingerprint(plan), documents=targets))
company.save(ROOT / "build" / "source-readiness.json", {
    "schema_version": 1, "canon_sha256": company.fingerprint(canon),
    "review": "Business totals, service and cost breakdowns reconcile; all requirements have methods and acceptance criteria; meetings and board requests have named owners and dates.",
    "financial_scope": "Income and operating snapshots only; no balance sheet, cash-flow model, invented mining revenue, or unsupported cash/runway assumptions.",
    "business_milestones": "Ground development releases only; future mission gates retain planned or goal status.",
    "organization": "Eight referenced people; effective-dated finance and mission-systems transitions. Named chart size is not financial headcount.",
    "public_review": "Public release facts and quotes are available on release date; internal financial data is excluded from press releases.",
})
for d in documents:
    company.save(ROOT / "build" / "contexts" / f"{d['id']}.json",
                 company.get_context(canon, plan, d["id"]))
print(json.dumps({"records": len(records), "documents": len(documents),
                  "formats": Counter(d["format"] for d in documents),
                  "types": Counter(d["type"] for d in documents),
                  "fiscal_years": Counter(index[d["anchor_ids"][0]]["data"].get("year") for d in documents),
                  "warnings": plan["warnings"]}, indent=2))
