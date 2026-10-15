import copy
import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
REPO = ROOT.parents[1]
sys.path.insert(0, str(REPO / ".github" / "skills" / "company-content-generator" / "scripts"))
import company


def save(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")


canon = company.seed_canon(company.load(ROOT / "build" / "request.json"))
records = canon["records"]
index = {r["id"]: r for r in records}


def add(identifier, kind, when, data, refs=(), effective=None):
    record = dict(id=identifier, kind=kind, known_on=when,
                  effective_on=effective or when, data=data, refs=list(refs))
    records.append(record)
    index[identifier] = record
    return record


index["company"]["data"].update(
    headquarters="Tampa, Florida", dateline="TAMPA, Fla.",
    founded_year=2016, operating_model="fabless semiconductor designer",
    engineering_site="Sarasota, Florida",
    website="https://sunnysemi.example",
    boilerplate="SunnySemi LLC is a Tampa, Florida-based fabless semiconductor designer. "
    "Its SunCore edge-inference processors and GulfSwitch power-management devices serve "
    "equipment manufacturers building industrial, infrastructure and embedded systems. "
    "The company designs silicon and reference platforms in Florida and works with "
    "external wafer-fabrication, assembly and test partners. Its Sarasota engineering "
    "site supports package characterization, application validation and customer bring-up. "
    "SunnySemi combines version-controlled silicon specifications with evaluation hardware "
    "and application support so customers can qualify components against their own "
    "system requirements. Product information is available at https://sunnysemi.example.",
    media_contact="Jordan Ellis | Corporate Communications | media@sunnysemi.example",
    finance_basis="USD; calendar fiscal years; recognized revenue, not bookings; "
    "operating income excludes no additional adjustments and is not net income or cash flow.",
)
index["product-beacon"]["data"].update(name="SunCore", category="edge-inference processor")
index["product-prism"]["data"].update(name="GulfSwitch", category="power-management controller")
for r in records:
    if r["kind"] == "acquisition":
        r["data"].update(target="CoralKey Packaging", price_usd=85_000_000)
index["person-jordan"]["data"]["communications_email"] = "media@sunnysemi.example"
add("relationship-media", "relationship", "2020-01-01",
    {"person_id": "person-jordan", "responsibility": "corporate media contact"},
    ["person-jordan", "company"])
index["glossary-arr"]["data"] = {
    "term": "TOPS", "definition": "Trillion integer operations per second at the stated precision; "
    "a silicon peak, not application throughput."}
for term, definition in [
    ("Fabless", "A company that designs semiconductors and contracts external wafer fabrication."),
    ("INT8", "Signed eight-bit integer arithmetic used by the SunCore peak-throughput specification."),
    ("PMBus", "A digital power-management command interface; support is limited to listed commands."),
    ("Qualification", "Documented verification against release criteria; not a regulatory certification."),
]:
    add("glossary-" + term.lower(), "glossary", "2020-01-01",
        {"term": term, "definition": definition})

THEMES = [
    ("Launch discipline", "industrial inspection",
     "a single qualified substrate route", "qualify a second substrate route before increasing release allocations",
     "keep first-customer bring-up separate from the production release gate",
     "evaluation-kit shipments and controlled production orders"),
    ("Allocation discipline", "warehouse automation",
     "long wafer-start commitments", "match wafer reservations to accepted customer schedules",
     "allocate constrained starts to qualified designs rather than unverified demand",
     "schedule-backed allocation and repeatable factory test"),
    ("Package readiness", "distributed infrastructure",
     "package-test capacity becoming the binding constraint", "reserve package-test slots before accepting accelerated ramps",
     "release a capacity increment only after package and test readiness are demonstrated together",
     "package characterization and documented interface compatibility"),
    ("Software maturity", "machine-vision equipment",
     "compiler regressions across model updates", "freeze the supported operator set and retain regression evidence",
     "treat compiler qualification as part of the hardware delivery promise",
     "model-porting support and reproducible application validation"),
    ("Thermal accountability", "embedded industrial computing",
     "customer enclosures with inadequate thermal headroom", "require board-level thermal evidence before ramp approval",
     "prefer a verified operating point to an unsupported peak-performance claim",
     "customer thermal reviews and power-aware deployment"),
    ("Profitable scale", "high-density edge systems",
     "working-capital exposure from accelerated production commitments", "stage capacity commitments against qualification and accepted schedules",
     "protect delivery reliability and operating discipline while scaling profitable demand",
     "release-controlled adoption and milestone-based production support"),
]
ANNUAL_REVENUE = [1_200_000_000, 2_400_000_000, 4_500_000_000,
                  8_000_000_000, 14_000_000_000, 23_000_000_000]
HEADCOUNT = [620, 980, 1550, 2450, 3700, 5200]
TOPS = [12, 24, 48, 96, 160, 256]
WATTS = [12, 15, 18, 22, 28, 35]
PROCESS = [12, 12, 7, 7, 5, 5]
VOLTS = [48, 48, 60, 60, 80, 100]
AMPS = [10, 12, 15, 18, 20, 24]
EFFICIENCY = [94, 95, 95, 96, 96, 97]

for offset, year in enumerate(range(2020, 2026)):
    total = ANNUAL_REVENUE[offset]
    for quarter, share in enumerate([20, 23, 26, 31], 1):
        r = index[f"financial-{year}-q{quarter}"]
        revenue = total * share // 100
        cogs = revenue * (46 - offset * 2) // 100
        opex = revenue * (28 - offset) // 100
        r["data"].update(revenue_usd=revenue, cogs_usd=cogs, opex_usd=opex,
                         operating_income_usd=revenue-cogs-opex,
                         headcount=HEADCOUNT[offset] - (4-quarter) * (20+offset*10),
                         suncore_revenue_usd=revenue * (60+offset*2) // 100,
                         gulfswitch_revenue_usd=revenue - revenue * (60+offset*2) // 100,
                         wafer_cost_usd=cogs * 65 // 100,
                         packaging_cost_usd=cogs * 25 // 100,
                         logistics_cost_usd=cogs - cogs*65//100 - cogs*25//100,
                         engineering_opex_usd=opex * 70 // 100,
                         commercial_opex_usd=opex - opex * 70 // 100)
    annual = next(r for r in records if r["kind"] == "financial" and
                  r["data"]["year"] == year and r["data"]["quarter"] is None)
    for field in ("revenue_usd", "cogs_usd", "opex_usd", "operating_income_usd",
                  "suncore_revenue_usd", "gulfswitch_revenue_usd", "wafer_cost_usd",
                  "packaging_cost_usd", "logistics_cost_usd", "engineering_opex_usd",
                  "commercial_opex_usd"):
        annual["data"][field] = sum(index[f"financial-{year}-q{q}"]["data"][field]
                                    for q in range(1, 5))
    annual["data"]["headcount"] = HEADCOUNT[offset]
    theme, market, risk, mitigation, principle, support = THEMES[offset]
    add(f"operating-{year}", "operating", f"{year}-01-15",
        dict(year=year, theme=theme, customer_market=market, risk=risk,
             mitigation=mitigation, operating_principle=principle, adoption_support=support,
             strategy="Own differentiated silicon architecture and application validation; "
             "use external manufacturing rather than build a wafer fab.",
             capacity_option="Stage commitments against qualification evidence.",
             alternative="Commit the entire proposed capacity increment immediately.",
             tradeoff="Staging limits irreversible commitments but can defer some shipment opportunities.",
             board_request="Approve milestone-based capacity releases; retain existing commercial "
             "approval limits and require Finance and Engineering sign-off before each increment.",
             board_status="Requested; no board approval is recorded in this briefing.",
             accountability="Quinn Mercer owns customer qualification; Riley Arden owns release evidence; "
             "the effective CFO owns capacity-commitment review.",
             liquidity_note="No cash balance, debt balance, free-cash-flow measure or forecast is included.",
             price_note="Pricing and contractual lead times are supplied through customer-specific quotations."),
        ["company", "person-quinn", "person-riley"])
    for product, short in [("beacon", "SunCore"), ("prism", "GulfSwitch")]:
        version = f"{offset+1}.0"
        pid = "product-" + product
        when = f"{year}-02-01"
        if product == "beacon":
            metric = f"{TOPS[offset]} INT8 TOPS"
            operating_point = f"{WATTS[offset]} W board design budget"
            capability = f"{PROCESS[offset]} nm inference ASIC with PCIe Gen3 x4 host interface"
            problem = "Moving industrial inference off a general-purpose host competes with "
            problem += "system power, response-time and software-integration constraints."
            workflow = "Capture frames on the host, prepare tensors, submit inference over PCIe, "
            workflow += "and return results to the host application."
            applications = [market, "local equipment anomaly detection", "on-device visual inspection"]
            constraints = "Peak INT8 throughput is not sustained application throughput. "
            constraints += "Customers must validate their model, host bandwidth and thermal design."
            requirements = [
                ("Host link", "Enumerate as a PCIe Gen3 x4 endpoint.",
                 "Complete cold-boot enumeration on the reference host in 100 consecutive boots."),
                ("Arithmetic", "Support signed INT8 multiply-accumulate inference.",
                 "Match the published INT8 golden vectors exactly for supported operators."),
                ("Peak compute", f"Expose a peak of {TOPS[offset]} INT8 TOPS at the rated clock.",
                 "Check the architectural operation-count model against the rated configuration."),
                ("Power envelope", f"Use a {WATTS[offset]} W reference-board design budget.",
                 "Measure board input power on the release workload without exceeding the budget."),
                ("Model interface", "Accept the release compiler's versioned binary model package.",
                 "Load a package built with the matching toolchain and reject an incompatible major version."),
                ("Unsupported operators", "Reject models containing operators outside the supported list.",
                 "Return an operator diagnostic before device execution for every negative test."),
                ("Tensor validation", "Validate tensor shape, element type and buffer bounds.",
                 "Reject malformed dimensions and undersized buffers without DMA outside the allocation."),
                ("DMA protection", "Limit DMA transactions to driver-registered buffers.",
                 "Run the invalid-buffer suite with no unauthorized host-memory access."),
                ("Reset recovery", "Recover the command interface after a host-issued device reset.",
                 "Execute the release smoke workload after 100 repeated reset cycles."),
                ("Telemetry", "Report device temperature and command completion status.",
                 "Compare temperature readout with the calibrated reference within 3 degrees Celsius."),
                ("Thermal response", "Throttle work admission before the board's configured thermal limit.",
                 "Raise ambient temperature and verify throttling and recovery at configured thresholds."),
                ("Error reporting", "Return distinct invalid-model, timeout and device-fault status codes.",
                 "Inject each documented fault and check the driver status and recovery path."),
                ("Driver package", "Deliver a Linux x86-64 reference driver with a version manifest.",
                 "Install on the supported reference image and run the complete release smoke suite."),
                ("Reproducibility", "Identify silicon revision, firmware and compiler in each diagnostic bundle.",
                 "Recreate an issue report using only the recorded release identifiers and input tensor."),
                ("Release evidence", "Retain functional, power and thermal verification results.",
                 "Block production release when any required acceptance result is absent or failing."),
                ("Change control", "Review interface-affecting changes before issuing revised packages.",
                 "Record compatibility impact and obtain Product and Engineering sign-off."),
            ]
        else:
            metric = f"{VOLTS[offset]} V maximum rated input"
            operating_point = f"{AMPS[offset]} A reference-stage output"
            capability = "digitally supervised multiphase power control with a documented PMBus subset"
            problem = "High-density embedded boards need controlled power sequencing and observable "
            problem += "fault behavior without losing sight of board-level thermal limits."
            workflow = "Configure the power stage, sequence rails, observe telemetry, and "
            workflow += "handle protection events through the documented host command interface."
            applications = [market, "industrial controller power stages", "embedded compute carrier boards"]
            constraints = "Efficiency applies only to the stated reference-stage operating point. "
            constraints += "External switches, inductors, cooling and layout remain customer design responsibilities."
            requirements = []
        detail = dict(product_id=pid, version=version, year=year, metric=metric,
                      operating_point=operating_point, capability=capability,
                      customer_problem=problem, workflow=workflow, use_cases=applications,
                      constraints=constraints, support=support, market=market,
                      release_date=f"{year}-08-15", evaluation_step="Request the version-matched evaluation "
                      "kit and interface pack from applications@sunnysemi.example.",
                      integration_step="Agree the target workload, host interface, thermal boundary and "
                      "qualification evidence before scheduling production adoption.",
                      support_email="applications@sunnysemi.example",
                      quote_speaker="person-avery",
                      quote=f"Our priority is to make {short} {version} useful in real equipment, "
                      "not just impressive in a specification table. Customers need clear operating "
                      "boundaries, traceable release evidence and an engineering team that helps "
                      "them qualify the complete design. That is how we intend to earn repeat programs.",
                      quote_approved=True,
                      availability="Production orders and version-matched evaluation kits are available "
                      "from the launch date, subject to quotation and allocation.",
                      qualification="Evaluate the reference board, confirm interface behavior, "
                      "complete workload and thermal validation, then approve the production configuration.",
                      exclusions="No regulatory certification, customer benchmark or customer endorsement is asserted.",
                      node_nm=PROCESS[offset] if product == "beacon" else 180,
                      peak_tops=TOPS[offset] if product == "beacon" else 0,
                      power_w=WATTS[offset] if product == "beacon" else 0,
                      input_max_v=VOLTS[offset] if product == "prism" else 0,
                      output_a=AMPS[offset] if product == "prism" else 0,
                      efficiency_percent=EFFICIENCY[offset] if product == "prism" else 0,
                      efficiency_conditions="48 V input, 12 V output, half rated reference-stage load, "
                      "25 degrees Celsius ambient; excludes host and enclosure losses.")
        add(f"release-{product}-{year}", "release_detail", when, detail,
            [pid, "person-avery", f"operating-{year}"])
        for number, (topic, requirement, test) in enumerate(requirements, 1):
            add(f"req-{product}-{year}-{number:02d}", "requirement", when,
                dict(product_id=pid, version=version, requirement_id=f"SC-{offset+1}-{number:02d}",
                     topic=topic, requirement=requirement, acceptance=test,
                     group="Functional" if number <= 9 else "Verification and control",
                     owner_id="person-riley", method="Reference-board verification"),
                [pid, "person-riley"])
    add(f"meeting-{year}", "meeting", f"{year}-07-25",
        dict(title=f"{year} Production Readiness and {theme}",
             date=f"{year}-07-25", participants=["person-avery", "person-riley", "person-quinn",
                "person-morgan" if year == 2020 else "person-casey"],
             location="Tampa, Florida",
             turn_times=[f"{9 + (i*2)//60:02d}:{(i*2)%60:02d}" for i in range(36)],
             agenda=[theme, "SunCore release evidence", "GulfSwitch adoption", "Capacity commitment controls"],
             positions={
                 "chief_executive": principle,
                 "engineering": "Do not substitute peak specifications for application qualification.",
                 "product": f"Support {market} with {support}.",
                 "finance": f"Control {risk}; {mitigation}."},
             decision=f"Proceed toward the August release subject to evidence gates; {mitigation}.",
             rejected_alternative="An unconditional production ramp before qualification is complete.",
             unresolved="Customer-specific adoption schedules remain subject to accepted quotations.",
             actions=[
                 {"owner_id": "person-riley", "task": "Complete version-matched acceptance evidence",
                  "due": f"{year}-08-05"},
                 {"owner_id": "person-quinn", "task": "Confirm customer evaluation and qualification plans",
                  "due": f"{year}-08-08"},
                 {"owner_id": "person-morgan" if year == 2020 else "person-casey",
                  "task": "Review staged manufacturing commitments against accepted schedules",
                  "due": f"{year}-08-10"}],
             governance="The meeting authorizes preparation, not a board-level capital approval. "
             "Failing release evidence must be escalated to the CEO before production release."),
        ["company", "person-avery", "person-riley", "person-quinn",
         "person-morgan" if year == 2020 else "person-casey", f"operating-{year}"])

company.validate_canon(canon)
save(ROOT / "canon.json", canon)
plan = company.create_plan(canon)
specifications = []
for year in range(2020, 2026):
    specifications.extend([
        ("press_release", "docx", f"milestone-beacon-{year}-launched", f"{year}-08-15",
         f"SunnySemi LLC Launches SunCore {year-2019}.0 for {THEMES[year-2020][1].title()}"),
        ("press_release", "docx", f"milestone-prism-{year}-launched", f"{year}-08-15",
         f"SunnySemi LLC Launches GulfSwitch {year-2019}.0 Power Control"),
        ("product_presentation", "pptx", f"milestone-beacon-{year}-launched", f"{year}-08-20",
         f"SunnySemi LLC - SunCore {year-2019}.0 - Edge Inference Adoption"),
        ("product_spec", "docx", f"milestone-beacon-{year}-planned", f"{year}-02-01",
         f"SunnySemi LLC - SunCore {year-2019}.0 Product Specification"),
        ("meeting_transcript", "docx", f"milestone-beacon-{year}-planned", f"{year}-07-25",
         f"SunnySemi LLC - {year} Production Readiness and {THEMES[year-2020][0]}"),
        ("board_deck", "pptx", f"financial-{year}-q2", f"{year}-07-20",
         f"SunnySemi LLC - {year} Midyear Performance and Capacity Decisions"),
        ("financial_model", "xlsx", f"financial-{year}-q2", f"{year}-07-15",
         f"SunnySemi LLC - H1 {year} Operating Economics"),
        ("financial_model", "xlsx", f"financial-{year}-q4", f"{year+1}-01-15",
         f"SunnySemi LLC - FY {year} Operating Economics"),
    ])
for year in (2022, 2025):
    specifications.append(("product_presentation", "pptx",
                           f"milestone-prism-{year}-launched", f"{year}-08-22",
                           f"SunnySemi LLC - GulfSwitch {year-2019}.0 - Power Stage Adoption"))
for doc, (kind, fmt, anchor, when, title) in zip(plan["documents"], specifications):
    doc.update(type=kind, format=fmt, anchor_ids=[anchor], publication_date=when,
               title=title, filename=company.document_filename(title, when, fmt),
               audience=company.type_registry(canon["request"])[kind][2],
               author_id="person-avery", viewpoint_id=None, source_version_id=None,
               access_groups=company.document_access_groups(
                   index[anchor], company.type_registry(canon["request"])[kind][2]),
               required_sections=list(company.required_sections(canon["request"], kind)),
               required_statement_ids=[], condition_ids=[], blocked_record_ids=[],
               suppressed_record_ids=[])
plan["documents"].sort(key=lambda d: (d["publication_date"], d["id"]))
plan["warnings"] = []
company.validate_plan(canon, plan)
save(ROOT / "build" / "plan.json", plan)
for doc in plan["documents"]:
    context = company.get_context(canon, plan, doc["id"])
    save(ROOT / "build" / "contexts" / f"{doc['id']}.json", context)
save(ROOT / "build" / "setup-evidence.json", {
    "officecli_source": "https://github.com/iOfficeAI/OfficeCLI",
    "officecli_release": "v1.0.157", "license": "Apache-2.0",
    "installed_sha256": "bd64eba587e56cc71709cdd1f38ae9ac77cd65a9c8ea744b0996897c0557e343",
    "official_checksum_matched": True, "installation_changes": [],
    "preflight_images_inspected": ["docx", "pptx", "xlsx", "xlsx-chart-column", "xlsx-chart-line"],
    "output_location_confirmed": str(ROOT),
    "publication_tail": "2026-01-15 for FY 2025 actuals",
    "post_generation_qa": {"enabled": False, "policy": None},
    "business_scenario": "Extraordinarily successful Florida-based fabless semiconductor company",
})
print(json.dumps({"records": len(records), "documents": len(plan["documents"]),
                  "formats": dict(Counter(d["format"] for d in plan["documents"])),
                  "types": dict(Counter(d["type"] for d in plan["documents"]))}, indent=2))
