import json
import re
import sys
from collections import Counter
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT.parents[1] / ".github" / "skills" / "company-content-generator" / "scripts"))
import company


def save(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=True) + "\n", encoding="utf-8")


canon = company.load(ROOT / "canon.json")
plan = company.load(ROOT / "build" / "plan.json")
company.validate_plan(canon, plan)
targets = []
for doc in plan["documents"]:
    kind = doc["type"]
    measure, low, high = {
        "press_release": ("delivered_words", 450, 700),
        "meeting_transcript": ("delivered_words", 1650, 2600),
        "product_spec": ("delivered_words", 850, 1900),
        "product_presentation": ("delivered_slides", 10, 10),
        "board_deck": ("delivered_slides", 10, 10),
        "financial_model": ("delivered_sheets", 4, 4),
    }[kind]
    targets.append(dict(document_id=doc["id"], content_depth="medium", measure=measure,
                        target_min=low, target_max=high,
                        coverage={
                            "press_release": ["Approved quote, technical operating boundaries, adoption and availability"],
                            "meeting_transcript": ["At least 1500 discussion words; alternatives, dependencies, decisions and owned actions"],
                            "product_spec": ["16 distinct numbered testable requirements; methods, interfaces and release controls"],
                            "product_presentation": ["10 substantive slides; use cases, workflow, constraints, adoption and next steps"],
                            "board_deck": ["10 substantive slides; financial visuals, resource tradeoff, risk and decision requested"],
                            "financial_model": ["Six useful modules: income, margins, product mix, manufacturing cost, operating investment and period trend"],
                        }[kind],
                        adaptation_rationale=(
                            "Four navigable sheets implement six distinct analytical modules; sheet count "
                            "is only a mechanical guard. Module usefulness is separately reviewed."
                            if kind == "financial_model" else
                            "DOCX words guard explanatory substance; all 16 distinct requirement controls "
                            "are separately counted and reviewed." if kind == "product_spec" else
                            "Delivered words include metadata; the discussion-only count is also recorded "
                            "and must reach 1500." if kind == "meeting_transcript" else None)))
save(ROOT / "build" / "depth-targets.json",
     dict(schema_version=2, plan_sha256=company.fingerprint(plan), documents=targets))


class Writer:
    def __init__(self, doc):
        self.ctx = company.load(ROOT / "build" / "contexts" / f"{doc['id']}.json")
        self.doc = self.ctx["document"]
        self.records = {r["id"]: r for r in self.ctx["records"]}
        self.roles = {r["data"]["person_id"]: r for r in self.ctx["active_roles"]}
        self.sections, self.claims, self.renderings = [], [], {}
        self.current = None
        self.year = int(re.search(r"20\d\d", doc["anchor_ids"][0]).group())
        self.offset = self.year - 2020
        self.anchor = doc["anchor_ids"][0]
        self.product = "prism" if "prism" in self.anchor else "beacon"
        self.detail = f"release-{self.product}-{self.year}"
        self.ops = f"operating-{self.year}"
        self.meeting = f"meeting-{self.year}"

    def sec(self, role, heading, page=False):
        self.current = dict(role=role, heading=heading, text="", claim_ids=[], page_break=page)

    def v(self, record, *path, rendering=None, units=None):
        fullpath = ["data", *path]
        value = company.resolve(self.records[record], fullpath)
        identifier = f"c{len(self.claims)+1:04d}"
        claim = dict(id=identifier, record_id=record, path=fullpath, value=value)
        if units or (isinstance(path[-1], str) and path[-1].endswith("_usd")):
            claim["units"] = units or "USD"
        elif path[-1] == "headcount":
            claim["units"] = "people"
        self.claims.append(claim)
        self.current["claim_ids"].append(identifier)
        if rendering:
            self.renderings[identifier] = rendering
            return company.display_value(claim, rendering)
        return str(value)

    def money(self, record, key):
        return self.v(record, key, rendering={"scale": 1_000_000, "decimals": 1,
                                             "prefix": "$", "suffix": " million"})

    def person(self, identifier, role=False):
        name = self.v(identifier, "name")
        if role:
            name += ", " + self.v(self.roles[identifier]["id"], "title")
        return name

    def done(self, text, **kwargs):
        self.current.update(text=text, **kwargs)
        self.sections.append(self.current)
        self.current = None

    def finish(self):
        return dict(schema_version=2, corpus_id=self.ctx["corpus_id"],
                    document_id=self.doc["id"], canon_sha256=self.ctx["canon_sha256"],
                    plan_sha256=self.ctx["plan_sha256"], title=self.doc["title"],
                    publication_date=self.doc["publication_date"], content_depth="medium",
                    claims=self.claims, sections=self.sections)


def press(b):
    b.sec("dateline_and_lead", "For immediate release")
    name = b.v("product-"+b.product, "name")
    version = b.v(b.detail, "version")
    date = b.v(b.detail, "release_date")
    status = b.v(b.anchor, "status")
    market = b.v(b.detail, "market")
    b.done(f"{b.v('company', 'dateline')} - {b.v('company', 'name')} today "
           f"announced that {name} {version} has {status}, effective {date}, expanding its "
           f"component portfolio for {market}. The release brings a version-specific "
           "silicon platform together with an evaluation path and clear operating boundaries. "
           "It is intended for equipment teams that need to move from a promising component "
           "specification to a qualified production design without treating those two steps "
           "as interchangeable. Customers can now begin adoption discussions around the "
           "released configuration rather than an unspecified future product.")
    b.sec("announcement_details", "A defined operating envelope")
    metric = b.v(b.detail, "metric")
    point = b.v(b.detail, "operating_point")
    capability = b.v(b.detail, "capability")
    extra = ""
    if b.product == "prism":
        extra = (f" Reference-stage efficiency is {b.v(b.detail, 'efficiency_percent')} percent "
                 f"under these conditions: {b.v(b.detail, 'efficiency_conditions')}")
    b.done(f"{name} {version} combines {capability} with {metric} and a {point}. "
           f"{b.v(b.detail, 'constraints')}{extra} The published boundaries are intended "
           "to help customers determine whether the part fits a proposed system before they "
           "commit engineering resources. Component characteristics do not remove the need "
           "to verify the surrounding board, software and enclosure. The adoption discussion "
           "therefore starts with a defined application and operating point, not a general "
           "promise that every workload will achieve the same result.")
    b.sec("attributed_quote", "Engineering for repeat programs")
    quote = b.v(b.detail, "quote")
    speaker = b.person("person-avery", True)
    b.done(f'"{quote}" said {speaker}. The release approach connects technical '
           "capability with the evidence needed to introduce a component into equipment. "
           "It gives design teams a practical basis for discussing integration responsibilities "
           "and identifying work that remains on the customer side of the interface.")
    b.sec("announcement_details", "From evaluation to adoption", page=True)
    uses = [b.v(b.detail, "use_cases", i) for i in range(3)]
    workflow = b.v(b.detail, "workflow")
    b.done(f"Application discussions include {uses[0]}, {uses[1]} and {uses[2]}. "
           f"The reference workflow is straightforward: {workflow} "
           f"SunnySemi's adoption support emphasizes {b.v(b.detail, 'support')}. "
           "Those activities give equipment manufacturers a way to compare the component's "
           "documented behavior with the requirements of their own design. They are not "
           "a substitute for customer acceptance, and they do not imply an endorsement "
           "from any equipment manufacturer. Integration scope and the production configuration "
           "should be agreed before a customer commits to a ramp.")
    b.sec("announcement_details", "Availability and customer qualification")
    b.done(f"{b.v(b.detail, 'availability')} {b.v(b.detail, 'evaluation_step')} "
           f"{b.v(b.detail, 'integration_step')} {b.v(b.ops, 'price_note')} "
           "This sequencing allows the engineering discussion to address the intended "
           "application before commercial commitments are finalized. Customers should "
           "use the matching release material when evaluating the component and should "
           "not assume that a later revision has identical interface behavior.")
    b.sec("company_boilerplate", "About SunnySemi LLC")
    b.done(b.v("company", "boilerplate"))
    b.sec("media_contact", "Media contact")
    b.done(b.v("company", "media_contact"))


def specification(b):
    b.sec("scope", "Product scope and release basis")
    name, version = b.v("product-beacon", "name"), b.v(b.detail, "version")
    b.done(f"{name} {version} is a {b.v(b.detail, 'capability')}. "
           f"This specification supports the {b.v(b.anchor, 'status')} release targeted for "
           f"{b.v(b.anchor, 'target_date')}; it is not evidence that production release "
           "has already occurred. The design envelope is "
           f"{b.v(b.detail, 'metric')} with a {b.v(b.detail, 'operating_point')}. "
           f"{b.v(b.detail, 'constraints')} Engineering must retain the acceptance "
           "evidence for the exact silicon, firmware, driver and compiler configuration "
           "being released. Interface compatibility is evaluated against that configuration, "
           "not inferred from a shared product family name. This document controls the "
           "requirements and verification path; commercial allocation remains a separate process.")
    b.sec("scope", "Interfaces, dependencies and use boundaries")
    b.done(f"{b.v(b.detail, 'workflow')} The qualification sequence is: "
           f"{b.v(b.detail, 'qualification')} "
           f"{b.v(b.detail, 'integration_step')} "
           "The host owns capture, tensor preparation and the application response to inference "
           "results. The accelerator interface must reject malformed inputs rather than permit "
           "unbounded device activity. Customer model coverage is limited by the release operator "
           "list. A model that compiles successfully still needs application-level verification. "
           "Reference-board results do not establish customer-enclosure thermal performance. "
           "Production configuration changes must be reviewed for their effect on compatibility, "
           "power and the existing verification evidence.")
    b.sec("scope", "Verification ownership and document navigation")
    b.done(f"{b.person('person-riley', True)} owns engineering evidence; "
           f"{b.person('person-quinn', True)} owns the product acceptance boundary. "
           "The following requirements are grouped into host and compute behavior, fault "
           "handling and observability, then release and compatibility control. Each numbered "
           "requirement includes a concrete acceptance condition. The release checklist must "
           "identify the method, configuration and result for every control; a missing result "
           "is not an implied pass. Failures require corrective evidence before the release "
           "gate can close. Technical exclusions remain explicit: "
           f"{b.v(b.detail, 'exclusions')}")
    for number in range(1, 17):
        rid = f"req-beacon-{b.year}-{number:02d}"
        b.sec("requirements", f"{number:02d} | {b.records[rid]['data']['topic']}",
              page=number in (1, 7, 13))
        code = b.v(rid, "requirement_id")
        requirement = b.v(rid, "requirement")
        test = b.v(rid, "acceptance")
        method = b.v(rid, "method")
        b.done(f"{code}: {requirement} Acceptance: {test} "
               f"Method: {method}. Evidence must identify the tested configuration and "
               "show the observed outcome, including any exception that prevents acceptance.")
    b.sec("acceptance_criteria", "Release gate and change control")
    b.done("Acceptance requires all numbered controls to pass for the version-matched "
           "configuration. A functional pass cannot compensate for missing power or thermal "
           "evidence. Negative-path results must be retained alongside nominal-path results "
           "so that failure containment is auditable. Engineering and Product review "
           "interface-affecting changes before distributing a revised package. An unresolved "
           "compatibility issue blocks acceptance of the affected configuration rather than "
           "being carried silently into customer qualification. Customer-specific constraints "
           "remain outside a component-level claim of universal suitability. "
           f"{b.v(b.detail, 'evaluation_step')}")


def transcript(b):
    b.sec("participants", "Meeting record and participants")
    people = [b.person(p, True) for p in b.records[b.meeting]["data"]["participants"]]
    b.done(f"{b.v(b.meeting, 'date')} | {b.v(b.meeting, 'location')}\n"
           + "; ".join(people) + ".\nAgenda: "
           + "; ".join(b.v(b.meeting, "agenda", i) for i in range(4)) + ".\n"
           + f"Release status: {b.v(b.anchor, 'status')}; target {b.v(b.anchor, 'target_date')}.")
    ceo, engineering, product, finance = b.records[b.meeting]["data"]["participants"]
    q2 = f"financial-{b.year}-q2"
    sun, gulf = f"release-beacon-{b.year}", f"release-prism-{b.year}"
    exchanges = [
        (ceo, "Before we discuss allocation, I want us to separate the commercial opportunity "
         "from the evidence needed to release a component. Our operating emphasis is "
         "{theme}. A growing business can create pressure to treat preparation as approval. "
         "What exactly do we need to preserve as we move toward the release gate?"),
        (engineering, "The engineering position is simple: {engpos} For SunCore we are "
         "working against {sunmetric} and {sunpower}. Those boundaries help us define "
         "verification, but they do not tell us how every customer model will behave. "
         "I would not present a peak hardware number as a guaranteed application result."),
        (product, "That distinction matters for {market}. Our support emphasis is "
         "{support}. I want the customer conversation to start with the intended workload, "
         "the host interface and the thermal boundary. If those remain open, a kit shipment "
         "is useful evaluation activity, but it is not a completed production qualification."),
        (finance, "Q2 recognized revenue was {revenue}, with operating income of {income}. "
         "Those are strong actuals, but they are not cash balances or a forecast. My concern "
         "is {risk}. Before we commit to a faster ramp, we should understand whether "
         "accepted schedules and release evidence support the manufacturing exposure."),
        (ceo, "Then let us make the alternative explicit. We could approve an unconditional "
         "ramp and assume qualification will catch up, or preserve the evidence gates "
         "while preparation continues. The latter can leave some shipment opportunities "
         "waiting. Why is that the right tradeoff for this release rather than an "
         "unnecessary delay to the commercial team?"),
        (engineering, "Because a missing result is not a small administrative gap. A system "
         "can enumerate correctly and still fail under a malformed tensor, a repeated reset "
         "or a thermal transition. The release controls cover those paths. We need the "
         "exact silicon and software combination identified, otherwise a passing result "
         "from another configuration is not the evidence we are discussing."),
        (product, "I agree with the gate, but I do not want customers left with an abstract "
         "instruction to qualify. The evaluation sequence is {qualification} We can use "
         "that sequence to explain what evidence belongs to SunnySemi and what depends "
         "on the customer's board or application. Clear ownership should reduce avoidable "
         "back-and-forth even when we cannot promise a schedule."),
        (finance, "The financial distinction is similar. Revenue is recognized revenue, "
         "not an order-book proxy, and operating income is not liquidity. We can assess "
         "the reported economics without pretending we have a full funding model here. "
         "My recommendation is to {mitigation}. That keeps the capacity discussion tied "
         "to evidence rather than enthusiasm about demand."),
        (ceo, "I hear two different gates: a technical gate for the release and a commercial "
         "gate for a production commitment. They should inform each other without becoming "
         "the same approval. Could a customer complete the relevant interface work but "
         "still be unable to adopt the configuration because the thermal design is "
         "not ready? I want that boundary explained plainly."),
        (engineering, "Yes. The device interface can behave correctly while the surrounding "
         "system remains unsuitable at the intended operating point. Our technical "
         "boundary is {sunconstraint} The customer must verify the board, cooling and "
         "enclosure. We should preserve the configuration in the evidence package so "
         "that a later change does not silently inherit the earlier acceptance result."),
        (product, "For GulfSwitch the same principle appears in a different form. The "
         "reference boundary is {gulfmetric} with {gulfpoint}. External components and "
         "layout affect the power stage. I would frame adoption around the documented "
         "reference conditions, then ask the customer to validate the intended design. "
         "An efficiency headline without those conditions would create the wrong expectation."),
        (finance, "That is why I prefer staged commitments to one large unconditional "
         "reservation. We can prepare for successful qualification without treating "
         "every discussion as an accepted production schedule. Staging has a cost in "
         "flexibility and may defer some opportunities. The benefit is that we do not "
         "make the largest irreversible commitment before the relevant uncertainties narrow."),
        (ceo, "Let us test the proposal against this year's operating issue. The risk is "
         "{risk}. If we describe the mitigation as {mitigation}, what should the "
         "review actually ask for? I am not asking for a new numerical hurdle today. "
         "I am asking how the existing evidence and accepted schedules reach the same decision."),
        (engineering, "The review should start with the version-matched acceptance pack. "
         "It should show the nominal and negative-path results, the configuration and "
         "any unresolved exception. A technical pass is not a prediction that every "
         "customer will succeed. It is evidence that the released configuration meets "
         "its own stated controls and can enter the next adoption discussion honestly."),
        (product, "Then the customer side should show the agreed evaluation boundary "
         "and the qualification work still open. For {market}, the support emphasis "
         "remains {support}. I would rather communicate an explicit dependency than "
         "imply an adoption date that has not been accepted. That gives applications "
         "support a concrete conversation instead of a generic promise to accelerate."),
        (finance, "And Finance should review the manufacturing commitment against "
         "accepted schedules, not substitute total recognized revenue for readiness. "
         "The actuals tell us the business is scaling profitably; they do not identify "
         "which configuration has cleared its release controls. We should document "
         "that distinction when we discuss a capacity increment with the effective approvers."),
        (ceo, "I support the direction, but I want the cost acknowledged. A gate can "
         "become a reason to wait indefinitely if no one owns the remaining work. "
         "The decision therefore needs named actions and due dates, not only a risk "
         "description. Which evidence will Engineering deliver, and which customer "
         "commitments will Product bring back for the review?"),
        (engineering, "Engineering's assigned task is {engtask}, due {engdue}. "
         "That means retaining the actual verification results and making unresolved "
         "exceptions visible. I will not close the evidence pack by changing a failed "
         "control into an assumption. If the required result is absent, the affected "
         "configuration remains blocked until the issue is resolved or escalated."),
        (product, "Product's assigned task is {producttask}, due {productdue}. "
         "The point is to make the handoff specific enough that Engineering and Finance "
         "are discussing the same adoption boundary. A request for an evaluation kit "
         "does not by itself define a production schedule. We need to preserve that "
         "distinction in both the customer discussion and the internal review."),
        (finance, "Finance's assigned task is {financetask}, due {financedue}. "
         "I will use the reported actuals as the economic baseline and keep any "
         "unapproved forecast out of that baseline. The commitment discussion should "
         "identify what is being staged and what evidence supports it, rather than "
         "treat the entire proposed ramp as a single indivisible choice."),
        (ceo, "Before we record the decision, I want to return to customer communication. "
         "The operating principle is to {principle}. That should be visible in the "
         "release message without publishing our internal capacity discussion. We "
         "can describe operating boundaries and the qualification sequence publicly. "
         "We should not turn a preparation meeting into an announcement of board approval."),
        (engineering, "Agreed. The technical copy should distinguish the silicon peak "
         "from sustained application behavior and identify the matching release "
         "configuration. It should not add a customer benchmark or certification that "
         "we have not established. The verification requirements give us a concrete "
         "basis for what we can say; customer-specific system results remain outside "
         "that component-level statement."),
        (product, "That also gives us a practical next step: {evaluation} "
         "After that, the applications discussion can define the workload and host "
         "boundary. I do not want a release presentation to imply that evaluation "
         "alone completes adoption. The customer still needs to approve the production "
         "configuration against its own equipment requirements before a ramp is treated as ready."),
        (finance, "I support proceeding with preparation under those conditions. The "
         "alternative we are rejecting is {rejected} We are not rejecting growth; "
         "we are controlling the order in which we incur commitments. That is "
         "consistent with the economics and with the technical dependencies described "
         "today. Any broader approval still has to go through its own authority."),
        (ceo, "The decision is {decision} The distinction between preparation, "
         "qualification and a production commitment should remain in the record. "
         "We have not approved a new forecast or waived a release control. If "
         "an action uncovers a dependency that changes readiness, bring it back "
         "rather than silently carrying it into the release package."),
        (engineering, "I will keep the release evidence tied to the exact configuration "
         "and will make any exception explicit. The governance boundary is "
         "{governance} That protects the purpose of the gate: we should be able "
         "to explain what passed, what did not and why the production configuration "
         "is eligible for the next step."),
        (product, "The remaining open item is {unresolved} I will reflect that "
         "in adoption discussions and avoid presenting an unaccepted customer "
         "schedule as a commitment. The useful follow-through is a qualification "
         "plan with clear responsibilities, not a reassuring sentence that hides "
         "the dependency. We can be ambitious about adoption and precise about its conditions."),
        (finance, "Then the record is clear enough for the next review. The "
         "reported economics remain actuals; the capacity proposal remains gated; "
         "and the named actions identify what evidence comes next. I will keep "
         "the manufacturing-commitment discussion separate from any unsupported "
         "cash-flow conclusion. That should let us act quickly when the evidence "
         "is ready without having to reconstruct today's assumptions."),
    ]
    for i in range(0, len(exchanges), 2):
        b.sec("discussion", f"Discussion | {i//2+1:02d}", page=i in (0, 6, 12, 18, 24))
        turns = []
        for j in (i, i+1):
            speaker, template = exchanges[j]
            values = {}
            for key in re.findall(r"\{(\w+)\}", template):
                bindings = {
                    "theme": (b.ops, "theme"), "market": (b.ops, "customer_market"),
                    "support": (b.ops, "adoption_support"), "risk": (b.ops, "risk"),
                    "mitigation": (b.ops, "mitigation"), "principle": (b.ops, "operating_principle"),
                    "engpos": (b.meeting, "positions", "engineering"),
                    "sunmetric": (sun, "metric"), "sunpower": (sun, "operating_point"),
                    "sunconstraint": (sun, "constraints"), "qualification": (sun, "qualification"),
                    "gulfmetric": (gulf, "metric"), "gulfpoint": (gulf, "operating_point"),
                    "evaluation": (sun, "evaluation_step"), "decision": (b.meeting, "decision"),
                    "rejected": (b.meeting, "rejected_alternative"),
                    "governance": (b.meeting, "governance"), "unresolved": (b.meeting, "unresolved"),
                    "engtask": (b.meeting, "actions", 0, "task"),
                    "engdue": (b.meeting, "actions", 0, "due"),
                    "producttask": (b.meeting, "actions", 1, "task"),
                    "productdue": (b.meeting, "actions", 1, "due"),
                    "financetask": (b.meeting, "actions", 2, "task"),
                    "financedue": (b.meeting, "actions", 2, "due"),
                }
                values[key] = b.money(q2, "revenue_usd" if key == "revenue" else "operating_income_usd") if key in (
                    "revenue", "income") else b.v(*bindings[key])
            stamp = b.v(b.meeting, "turn_times", j)
            turns.append(f"{stamp} | {b.person(speaker)}: " + template.format(**values))
        b.done("\n\n".join(turns))
    b.sec("decisions", "Decision and authority boundary")
    b.done(b.v(b.meeting, "decision") + "\n" + b.v(b.meeting, "governance"))
    b.sec("actions", "Owned follow-through")
    b.done("\n".join(f"{b.person(a['owner_id'])}: "
                     f"{b.v(b.meeting, 'actions', i, 'task')} - due "
                     f"{b.v(b.meeting, 'actions', i, 'due')}."
                     for i, a in enumerate(b.records[b.meeting]["data"]["actions"])))


def presentation(b):
    name = b.records["product-"+b.product]["data"]["name"]
    version = b.records[b.detail]["data"]["version"]
    b.sec("value_proposition", b.doc["title"])
    b.done(f"{b.v('product-'+b.product, 'name')} {b.v(b.detail, 'version')}\n"
           f"{b.v(b.anchor, 'status')} | {b.v(b.detail, 'release_date')}\n"
           f"{b.v('company', 'headquarters')}\n"
           f"{b.person('person-quinn', True)}",
           visual=["Released silicon", "Evaluation", "Qualification"])
    b.sec("customer_problem", "Start with the equipment boundary")
    b.done(b.v(b.detail, "customer_problem") + "\n" + b.v(b.detail, "constraints"),
           visual=["Workload", "Power", "Integration"])
    b.sec("value_proposition", "A component and a qualification path")
    b.done(b.v(b.detail, "capability") + "\n" + b.v(b.detail, "qualification"),
           visual=["Reference platform", "Acceptance evidence", "Customer configuration"])
    b.sec("capabilities", "The published operating envelope")
    b.done(b.v(b.detail, "metric") + "\n" + b.v(b.detail, "operating_point") +
           "\n" + b.v(b.detail, "constraints"),
           visual=[b.records[b.detail]["data"]["metric"], b.records[b.detail]["data"]["operating_point"]])
    b.sec("capabilities", "What the interface enables")
    b.done(b.v(b.detail, "workflow") + "\n" +
           "Keep the host boundary and the version-matched component configuration explicit.",
           visual=["Host interface", "Component execution", "Status and results"])
    b.sec("use_cases", "Application paths to evaluate")
    cases = [b.v(b.detail, "use_cases", i) for i in range(3)]
    b.done("\n".join(cases) + "\nValidate each use case against the intended system, not the product label.",
           visual=cases)
    b.sec("use_cases", "Integration is a shared responsibility")
    b.done(b.v(b.detail, "integration_step") + "\n" + b.v(b.detail, "support"),
           visual=["Customer workload", "SunnySemi evidence", "Joint qualification"])
    b.sec("capabilities", "Boundaries that must stay visible")
    extra = ""
    if b.product == "prism":
        extra = (f"\n{b.v(b.detail, 'efficiency_percent')} percent reference-stage efficiency: "
                 + b.v(b.detail, "efficiency_conditions"))
    b.done(b.v(b.detail, "constraints") + extra + "\n" + b.v(b.detail, "exclusions"),
           visual=["Component scope", "Board scope", "System scope"])
    b.sec("next_steps", "Plan adoption around release evidence")
    b.done(b.v(b.detail, "availability") + "\n" + b.v(b.ops, "price_note"),
           visual=["Evaluate", "Agree configuration", "Quote and allocate"])
    b.sec("next_steps", "Begin with a version-matched evaluation")
    b.done(b.v(b.detail, "evaluation_step") + "\n" + b.v(b.detail, "integration_step"),
           visual=["Evaluation kit", "Interface pack", "Qualification plan"])


def board(b):
    q1, q2 = f"financial-{b.year}-q1", f"financial-{b.year}-q2"
    sun, gulf = f"release-beacon-{b.year}", f"release-prism-{b.year}"
    b.sec("executive_summary", b.doc["title"])
    b.done(f"{b.v(b.ops, 'theme')}\n{b.v('company', 'headquarters')}\n"
           f"{b.person('person-avery', True)}\n"
           "Midyear actuals and a gated capacity decision; confidential board material.",
           visual=["Actual performance", "Release readiness", "Capacity authority"])
    b.sec("performance", "Recognized revenue and operating income")
    b.done(f"Q1 revenue: {b.money(q1, 'revenue_usd')}\n"
           f"Q2 revenue: {b.money(q2, 'revenue_usd')}\n"
           f"Q2 operating income: {b.money(q2, 'operating_income_usd')}\n"
           "Revenue and operating income are actuals, not bookings or cash flow.",
           bars=[("Q1 revenue", b.records[q1]["data"]["revenue_usd"]),
                 ("Q2 revenue", b.records[q2]["data"]["revenue_usd"])])
    b.sec("performance", "The Q2 operating-income bridge")
    b.done(f"Revenue {b.money(q2, 'revenue_usd')}\n"
           f"Less COGS {b.money(q2, 'cogs_usd')}\n"
           f"Less operating expense {b.money(q2, 'opex_usd')}\n"
           f"Operating income {b.money(q2, 'operating_income_usd')}",
           bars=[(k, b.records[q2]["data"][v]) for k, v in [
               ("Revenue", "revenue_usd"), ("COGS", "cogs_usd"),
               ("Opex", "opex_usd"), ("Operating income", "operating_income_usd")]])
    b.sec("performance", "Two silicon franchises fund the strategy")
    b.done(f"SunCore recognized revenue: {b.money(q2, 'suncore_revenue_usd')}\n"
           f"GulfSwitch recognized revenue: {b.money(q2, 'gulfswitch_revenue_usd')}\n"
           + b.v(b.ops, "strategy"),
           bars=[("SunCore", b.records[q2]["data"]["suncore_revenue_usd"]),
                 ("GulfSwitch", b.records[q2]["data"]["gulfswitch_revenue_usd"])])
    b.sec("performance", "Manufacturing spend has distinct drivers")
    b.done(f"Wafer cost {b.money(q2, 'wafer_cost_usd')}\n"
           f"Package and test {b.money(q2, 'packaging_cost_usd')}\n"
           f"Logistics {b.money(q2, 'logistics_cost_usd')}\n"
           "These components reconcile to reported COGS.",
           bars=[(k, b.records[q2]["data"][v]) for k, v in [
               ("Wafer", "wafer_cost_usd"), ("Package/test", "packaging_cost_usd"),
               ("Logistics", "logistics_cost_usd")]])
    b.sec("performance", "Engineering investment and release accountability")
    b.done(f"Engineering opex: {b.money(q2, 'engineering_opex_usd')}\n"
           f"Commercial and administrative opex: {b.money(q2, 'commercial_opex_usd')}\n"
           f"Q2 ending headcount: {b.v(q2, 'headcount')}\n"
           + b.v(b.ops, "accountability"),
           visual=["Engineering evidence", "Customer qualification", "Finance sign-off"])
    b.sec("executive_summary", "The August release gate is still ahead")
    b.done(f"SunCore {b.v(sun, 'version')}: {b.v(sun, 'metric')}\n"
           f"GulfSwitch {b.v(gulf, 'version')}: {b.v(gulf, 'metric')}\n"
           f"Target release date: {b.v(sun, 'release_date')}\n"
           "Planned at this briefing date; completion is not asserted.",
           visual=["Specification", "Verification", "Production release"])
    b.sec("risks", "Concentrate controls on the binding risk")
    b.done("Risk: " + b.v(b.ops, "risk") + "\nMitigation: " + b.v(b.ops, "mitigation") +
           "\n" + b.v(b.ops, "operating_principle"),
           visual=["Identify dependency", "Retain evidence", "Gate commitment"])
    b.sec("risks", "Capacity options and transaction context")
    transaction = "acquisition-sable-announced" if b.year == 2020 else "acquisition-sable-closed"
    b.done(f"Preferred: {b.v(b.ops, 'capacity_option')}\n"
           f"Alternative: {b.v(b.ops, 'alternative')}\n"
           f"{b.v(b.ops, 'tradeoff')}\n"
           f"{b.v(transaction, 'target')}: {b.v(transaction, 'status')}, "
           f"{b.money(transaction, 'price_usd')}.",
           visual=["Stage commitments", "Commit immediately"])
    b.sec("decisions_requested", "Decision requested and authority retained")
    b.done(b.v(b.ops, "board_request") + "\n" + b.v(b.ops, "board_status") +
           "\n" + b.v(b.ops, "liquidity_note"),
           visual=["Board direction", "Engineering evidence", "Finance release"])


class Workbook:
    def __init__(self, b):
        self.b, self.commands, self.bindings, self.expected = b, [], [], []
        self.cells = {}
        self.extents = {}

    def cmd(self, command, path, **kwargs):
        self.commands.append(dict(command=command, path=path, **kwargs))
        return len(self.commands)-1

    def cell(self, sheet, cell, text, claim=None, formula=None, header=False):
        props = {"value": str(text), "font.name": "Calibri", "font.size": "11pt",
                 "alignment.wrapText": "true", "alignment.vertical": "center"}
        if header:
            props.update(fill="123B4A", **{"font.color": "FFFFFF", "font.bold": "true"})
        elif formula:
            props.pop("value")
            props.update(formula=formula, **{"font.color": "008000", "numberformat": "#,##0.00;(#,##0.00);-"})
        elif isinstance(text, (int, float)):
            props.update(**{"font.color": "0000FF", "numberformat": "#,##0;(#,##0);-"})
        index = self.cmd("set", f"/{sheet}/{cell}", props=props)
        if claim:
            self.bindings.append(dict(section_index=len(self.b.sections), claim_id=claim,
                                      command_index=index))
        self.cells[sheet, cell] = text
        return index

    def input(self, sheet, cell, record, field):
        value = self.b.v(record, field)
        claim = self.b.claims[-1]
        raw = claim["value"]
        self.cell(sheet, cell, raw, claim=claim["id"])
        return claim["id"], raw

    def derive(self, sheet, cell, formula, operation, inputs, units="USD", rounding=12):
        claims = {c["id"]: c for c in self.b.claims}
        values = [Decimal(str(claims[i]["value"])) for i in inputs]
        if operation == "sum":
            value = sum(values)
        elif operation == "difference":
            value = values[0]-values[1]
        elif operation == "ratio":
            value = values[0]/values[1]
        else:
            raise ValueError(operation)
        value = value.quantize(Decimal(1).scaleb(-rounding), rounding=ROUND_HALF_UP)
        value = int(value) if value == value.to_integral_value() else str(value.normalize())
        identifier = f"c{len(self.b.claims)+1:04d}"
        self.b.claims.append(dict(id=identifier, value=value, units=units,
                                 derivation=dict(operation=operation, inputs=inputs, rounding=rounding)))
        self.b.current["claim_ids"].append(identifier)
        self.cell(sheet, cell, value, claim=identifier, formula=formula)
        self.expected.append(dict(sheet=sheet, cell=cell, formula=formula,
                                  value=value, claim_id=identifier))
        return identifier, value


def financial(b):
    w = Workbook(b)
    w.cmd("set", "/", props={"title": b.doc["title"], "author": "SunnySemi LLC"})
    w.cmd("set", "/Sheet1", props={"name": "Inputs"})
    for sheet in ("Calc", "Summary", "Notes"):
        w.cmd("add", "/", type="sheet", props={"name": sheet})
    for sheet in ("Inputs", "Calc", "Summary", "Notes"):
        for col, width in [("A", 40), ("B", 24), ("C", 24), ("D", 24), ("E", 24)]:
            w.cmd("set", f"/{sheet}/col[{col}]", props={"width": str(width)})
        for row in range(1, 37):
            w.cmd("set", f"/{sheet}/row[{row}]", props={"height": "27"})
    fields = ["revenue_usd", "cogs_usd", "opex_usd", "operating_income_usd",
              "suncore_revenue_usd", "gulfswitch_revenue_usd", "wafer_cost_usd",
              "packaging_cost_usd", "logistics_cost_usd", "engineering_opex_usd",
              "commercial_opex_usd", "headcount"]
    labels = ["Recognized revenue", "Cost of goods sold", "Operating expense", "Operating income",
              "SunCore revenue", "GulfSwitch revenue", "Wafer cost", "Package and test cost",
              "Logistics cost", "Engineering opex", "Commercial/admin opex", "Ending headcount"]
    quarters = 2 if "-q2" in b.anchor else 4
    b.sec("source_actuals", "Period actuals")
    w.cell("Inputs", "A1", b.doc["title"], header=True)
    w.cmd("set", "/Inputs/A1", props={"merge": "A1:E2"})
    w.cell("Inputs", "A3", "USD unless noted; headcount is point-in-time.", header=True)
    w.cmd("set", "/Inputs/A3", props={"merge": "A3:E3"})
    for q in range(1, quarters+1):
        w.cell("Inputs", f"{chr(65+q)}4", f"Q{q} {b.year}", header=True)
    primary = {}
    for row, (field, label) in enumerate(zip(fields, labels), 5):
        w.cell("Inputs", f"A{row}", label)
        for q in range(1, quarters+1):
            primary[q, field] = w.input("Inputs", f"{chr(65+q)}{row}",
                                        f"financial-{b.year}-q{q}", field)
    b.done("Source actuals for available quarters. Revenue, costs and operating income are USD; "
           "ending headcount is not a flow. Segment, manufacturing and operating-expense "
           "breakdowns reconcile to their reported totals.")
    b.sec("calculations", "Six operating-economics modules")
    calc_labels = [
        "1 | Income reconciliation", "Gross profit", "Calculated operating income",
        "Income reconciliation difference", "2 | Margins", "Gross margin (ratio)",
        "Operating margin (ratio)", "3 | Product mix", "Combined product revenue",
        "Revenue mix difference", "SunCore share (ratio)", "4 | Manufacturing costs",
        "Combined manufacturing costs", "Manufacturing cost difference", "Wafer share (ratio)",
        "5 | Operating investment", "Combined operating expenses", "Opex difference",
        "Engineering share (ratio)", "6 | Sequential actuals", "Revenue change vs prior quarter",
        "Operating-income change", "Ending headcount change",
    ]
    w.cell("Calc", "A1", "Operating economics | formula-linked actuals", header=True)
    w.cmd("set", "/Calc/A1", props={"merge": "A1:E2"})
    module_rows = {4, 8, 11, 15, 19, 23}
    for row, label in enumerate(calc_labels, 4):
        w.cell("Calc", f"A{row}", label, header=row in module_rows)
    output = {}
    for q in range(1, quarters+1):
        col = chr(65+q)
        w.cell("Calc", f"{col}3", f"Q{q} {b.year}", header=True)
        def p(field):
            return primary[q, field][0]
        output[q, "gross"] = w.derive("Calc", f"{col}5", f"Inputs!{col}5-Inputs!{col}6",
                                    "difference", [p("revenue_usd"), p("cogs_usd")])[0]
        output[q, "income"] = w.derive("Calc", f"{col}6", f"{col}5-Inputs!{col}7",
                                     "difference", [output[q, "gross"], p("opex_usd")])[0]
        w.derive("Calc", f"{col}7", f"{col}6-Inputs!{col}8",
                 "difference", [output[q, "income"], p("operating_income_usd")])
        output[q, "grossmargin"] = w.derive("Calc", f"{col}9", f"{col}5/Inputs!{col}5",
                                          "ratio", [output[q, "gross"], p("revenue_usd")], "ratio")[0]
        output[q, "margin"] = w.derive("Calc", f"{col}10", f"{col}6/Inputs!{col}5",
                                     "ratio", [output[q, "income"], p("revenue_usd")], "ratio")[0]
        mix = w.derive("Calc", f"{col}12", f"Inputs!{col}9+Inputs!{col}10", "sum",
                       [p("suncore_revenue_usd"), p("gulfswitch_revenue_usd")])[0]
        w.derive("Calc", f"{col}13", f"{col}12-Inputs!{col}5", "difference", [mix, p("revenue_usd")])
        w.derive("Calc", f"{col}14", f"Inputs!{col}9/Inputs!{col}5", "ratio",
                 [p("suncore_revenue_usd"), p("revenue_usd")], "ratio")
        cost = w.derive("Calc", f"{col}16", f"Inputs!{col}11+Inputs!{col}12+Inputs!{col}13",
                        "sum", [p("wafer_cost_usd"), p("packaging_cost_usd"), p("logistics_cost_usd")])[0]
        w.derive("Calc", f"{col}17", f"{col}16-Inputs!{col}6", "difference", [cost, p("cogs_usd")])
        w.derive("Calc", f"{col}18", f"Inputs!{col}11/Inputs!{col}6", "ratio",
                 [p("wafer_cost_usd"), p("cogs_usd")], "ratio")
        opex = w.derive("Calc", f"{col}20", f"Inputs!{col}14+Inputs!{col}15", "sum",
                        [p("engineering_opex_usd"), p("commercial_opex_usd")])[0]
        w.derive("Calc", f"{col}21", f"{col}20-Inputs!{col}7", "difference", [opex, p("opex_usd")])
        w.derive("Calc", f"{col}22", f"Inputs!{col}14/Inputs!{col}7", "ratio",
                 [p("engineering_opex_usd"), p("opex_usd")], "ratio")
        if q == 1:
            for row in (24, 25, 26):
                w.cell("Calc", f"{col}{row}", "No prior quarter in scope")
        else:
            prev = chr(64+q)
            for row, field, source_row, units in [
                (24, "revenue_usd", 5, "USD"), (25, "operating_income_usd", 8, "USD"),
                (26, "headcount", 16, "people")]:
                w.derive("Calc", f"{col}{row}", f"Inputs!{col}{source_row}-Inputs!{prev}{source_row}",
                         "difference", [p(field), primary[q-1, field][0]], units)
    b.done("Income reconciliation derives gross profit and operating income, then checks the reported result. "
           "Margins distinguish gross and operating economics. Product mix reconciles the two product "
           "franchises. Manufacturing costs reconcile wafer, package/test and logistics expense. "
           "Operating investment separates engineering from commercial and administrative expense. "
           "Sequential actuals compare revenue, income and point-in-time headcount without treating "
           "the latter as a quarterly flow.")
    b.sec("calculations", "Management interpretation")
    w.cell("Summary", "A1", "Management view | available-period actuals", header=True)
    w.cmd("set", "/Summary/A1", props={"merge": "A1:E2"})
    rows = [("Recognized revenue", "revenue_usd", 5), ("COGS", "cogs_usd", 6),
            ("Operating expense", "opex_usd", 7), ("Operating income", "operating_income_usd", 8),
            ("SunCore revenue", "suncore_revenue_usd", 9), ("GulfSwitch revenue", "gulfswitch_revenue_usd", 10)]
    totals = {}
    for row, (label, field, source_row) in enumerate(rows, 4):
        w.cell("Summary", f"A{row}", label)
        inputs = [primary[q, field][0] for q in range(1, quarters+1)]
        formula = "+".join(f"Inputs!{chr(65+q)}{source_row}" for q in range(1, quarters+1))
        totals[field] = w.derive("Summary", f"B{row}", formula, "sum", inputs)[0]
    w.cell("Summary", "A11", "Operating margin (ratio)")
    w.derive("Summary", "B11", "B7/B4", "ratio",
             [totals["operating_income_usd"], totals["revenue_usd"]], "ratio")
    interpretations = [
        ("Income", "Reconciliation differences on Calc must be zero; investigate any divergence."),
        ("Margins", "Gross margin excludes operating expense; operating margin includes it."),
        ("Product mix", "Product totals reconcile to revenue; a mix shift is not a customer-count measure."),
        ("Manufacturing", "Wafer, package/test and logistics are separate cost drivers; do not double count."),
        ("Investment", "Engineering is an operating expense here, not a capital expenditure or cash-flow proxy."),
        ("Trend", "Sequential changes are absolute movements, not forecasts or annualized growth rates."),
    ]
    for row, (label, text) in enumerate(interpretations, 14):
        w.cell("Summary", f"A{row}", label)
        w.cell("Summary", f"B{row}", text)
        w.cmd("set", f"/Summary/B{row}", props={"merge": f"B{row}:E{row}"})
        w.cmd("set", f"/Summary/row[{row}]", props={"height": "42"})
    w.cell("Summary", "E20", "USD basis")
    b.done("Six modules answer different questions; the management view summarizes actuals and interpretation. "
           "The model contains no forecast, cash balance, debt balance or valuation.")
    b.sec("assumptions_and_notes", "Basis, controls and limitations")
    notes = [
        ("Basis", b.v("company", "finance_basis")),
        ("Operating emphasis", b.v(b.ops, "theme")),
        ("Risk", b.v(b.ops, "risk")),
        ("Mitigation", b.v(b.ops, "mitigation")),
        ("Strategy", b.v(b.ops, "strategy")),
        ("Liquidity", b.v(b.ops, "liquidity_note")),
        ("Availability", "Only quarters knowable on the publication date are included."),
        ("Formula controls", "Inputs precede formulas; saved formula values are checked after reopening."),
        ("Color convention", "Blue numbers are source inputs. Green numbers are formula-linked outputs."),
        ("Scope", "This is an operating-economics workbook, not a balance sheet or cash-flow model."),
    ]
    w.cell("Notes", "A1", "Sources, methodology and limitations", header=True)
    w.cmd("set", "/Notes/A1", props={"merge": "A1:E2"})
    for row, (label, text) in enumerate(notes, 4):
        w.cell("Notes", f"A{row}", label)
        # Bind each source note to its exact authored target rather than a global token match.
        claim = b.current["claim_ids"][row-4] if row-4 < 6 else None
        w.cell("Notes", f"B{row}", text, claim=claim)
        w.cmd("set", f"/Notes/B{row}", props={"merge": f"B{row}:E{row}"})
        w.cmd("set", f"/Notes/row[{row}]", props={"height": "48"})
    w.cell("Notes", "E14", "End of notes")
    b.done("\n".join(f"{label}: {text}" for label, text in notes))
    return w


NAVY, TEAL, GOLD, INK = "123B4A", "087E8B", "E5AB37", "172F38"


def word_job(b):
    commands = [
        {"command": "set", "path": "/", "props": {"title": b.doc["title"], "author": "SunnySemi LLC",
                                                     "docDefaults.font": "Calibri", "docDefaults.fontSize": "11pt"}},
        {"command": "set", "path": "/section[1]", "props": {
            "pageWidth": "8.5in", "pageHeight": "11in", "marginTop": "0.7in",
            "marginBottom": "0.7in", "marginLeft": "0.8in", "marginRight": "0.8in"}},
    ]
    for name, size in (("Title", 21), ("Heading1", 15), ("Heading2", 14)):
        commands.append(dict(command="add", path="/styles", type="style",
                             props={"id": name, "name": name, "type": "paragraph",
                                    "font": "Calibri", "size": f"{size}pt", "bold": "true"}))
    bindings = []
    def paragraph(text, size=11, bold=False, color=INK, page=False, style=None):
        props = dict(text=text, font="Calibri", size=f"{size}pt", color=color,
                     spaceAfter="6pt", lineSpacing="1.15x", firstLineIndent="440")
        if bold:
            props["bold"] = "true"
        if page:
            props["pageBreakBefore"] = "true"
        if style:
            props["style"] = style
        commands.append(dict(command="add", path="/body", type="paragraph", props=props))
        return len(commands)-1
    paragraph("SUNNYSEMI LLC | " + ("NEWSROOM" if b.doc["type"] == "press_release" else "INTERNAL"), 10, True, TEAL)
    paragraph(b.doc["title"], 21, True, NAVY, style="Title")
    paragraph(b.doc["publication_date"] + " | " + b.ctx["company"]["data"]["headquarters"], 10, color=TEAL)
    for si, section in enumerate(b.sections):
        if section["page_break"]:
            paragraph("SunnySemi LLC | " + section["heading"], 15, True, NAVY, page=True, style="Heading1")
        elif b.doc["type"] != "press_release" or section["role"] in ("company_boilerplate", "media_contact"):
            paragraph(section["heading"], 14, True, NAVY, style="Heading2")
        ci = paragraph(section["text"])
        for claim in section["claim_ids"]:
            binding = dict(section_index=si, claim_id=claim, command_index=ci, occurrence_index=0)
            if claim in b.renderings:
                binding["rendering"] = b.renderings[claim]
            bindings.append(binding)
    return commands, bindings, [{"id": "all", "pages": 1+sum(s["page_break"] for s in b.sections)}]


def deck_job(b):
    commands = [dict(command="set", path="/", props={
        "title": b.doc["title"], "author": "SunnySemi LLC",
        "slideWidth": "33.867cm", "slideHeight": "19.05cm"})]
    bindings = []
    def shape(slide, text, x, y, width, height, size=20, fill="none", color=INK, bold=False):
        props = dict(text=text, x=f"{x}cm", y=f"{y}cm", width=f"{width}cm",
                     height=f"{height}cm", size=f"{size}pt", font="Calibri",
                     fill=fill, color=color, geometry="rect", valign="center", align="left")
        if bold:
            props["bold"] = "true"
        commands.append(dict(command="add", path=f"/slide[{slide}]", type="shape", props=props))
        return len(commands)-1
    for si, section in enumerate(b.sections):
        slide = si+1
        dark = si in (0, 9)
        commands.append(dict(command="add", path="/", type="slide",
                             props={"background": NAVY if dark else "F8FBFB"}))
        text_color = "FFFFFF" if dark else INK
        shape(slide, f"SUNNYSEMI LLC   |   {b.doc['publication_date']}", 1.4, 0.4, 29, 0.95,
              11, color="BEE5E6" if dark else TEAL)
        shape(slide, section["heading"], 1.4, 1.4, 30.8, 4.2, 36, color=text_color, bold=True)
        body = shape(slide, section["text"], 1.4, 6.1, 17.2, 10.8, 20, color=text_color)
        for claim in section["claim_ids"]:
            item = dict(section_index=si, claim_id=claim, command_index=body, occurrence_index=0)
            if claim in b.renderings:
                item["rendering"] = b.renderings[claim]
            bindings.append(item)
        if "bars" in section:
            bars = section["bars"]
            maximum = max(value for _, value in bars)
            step = 9.7 / len(bars)
            for n, (label, value) in enumerate(bars):
                y = 6.3+n*step
                shape(slide, label, 20.1, y, 11.9, 1.0, 14, color=text_color)
                shape(slide, "", 20.1, y+1.0, 11.9*value/maximum, 0.65,
                      fill=TEAL if n % 2 == 0 else GOLD)
            shape(slide, "USD | common scale within this chart", 20.1, 16.7, 11.9, 0.9, 10, color=text_color)
        else:
            visual = section["visual"]
            step = 10.1/len(visual)
            for n, label in enumerate(visual):
                y = 6.3+n*step
                shape(slide, str(n+1), 20.1, y, 1.5, 1.5, 22, TEAL, "FFFFFF", True)
                shape(slide, label, 22.0, y, 9.8, min(2.8, step-0.2), 20, color=text_color)
        shape(slide, f"{slide:02d}  |  " + ("BOARD CONFIDENTIAL" if b.doc["type"] == "board_deck" else "PRODUCT BRIEFING"),
              1.4, 17.9, 30.8, 0.9, 10, color="BEE5E6" if dark else TEAL)
        commands.append(dict(command="add", path=f"/slide[{slide}]", type="notes",
                             props={"text": section["text"] + " Keep the stated operating and authority boundaries explicit."}))
    return commands, bindings, [{"id": "all", "slides": len(b.sections)}]


jobs, layouts, drafts, measurements = [], [], {}, []
for doc in plan["documents"]:
    b = Writer(doc)
    kind = doc["type"]
    w = None
    if kind == "financial_model":
        w = financial(b)
    else:
        {"press_release": press, "product_spec": specification, "meeting_transcript": transcript,
         "product_presentation": presentation, "board_deck": board}[kind](b)
    draft = b.finish()
    company.check_draft(canon, plan, draft)
    draft_path = ROOT / "build" / "drafts" / f"{doc['id']}.json"
    save(draft_path, draft)
    drafts[doc["id"]] = draft
    if w:
        commands, bindings = w.commands, w.bindings
        render_targets = [
            {"id": "inputs", "type": "range", "range": "Inputs!A1:E16", "sheet_index": 0},
            {"id": "calc-top", "type": "range", "range": "Calc!A1:E14", "sheet_index": 1},
            {"id": "calc-bottom", "type": "range", "range": "Calc!A15:E26", "sheet_index": 1},
            {"id": "summary", "type": "range", "range": "Summary!A1:E20", "sheet_index": 2},
            {"id": "notes", "type": "range", "range": "Notes!A1:E14", "sheet_index": 3},
        ]
    elif doc["format"] == "docx":
        commands, bindings, render_targets = word_job(b)
    else:
        commands, bindings, render_targets = deck_job(b)
    company.validate_export_bindings(draft, bindings)
    batch = f"build/scripts/{doc['id']}-batch.json"
    layout = f"build/layouts/{kind}-v1.json"
    layout_path = ROOT / layout
    if not layout_path.exists():
        save(layout_path, dict(revision="1", genre=kind, format=doc["format"],
                               colors={"navy": NAVY, "teal": TEAL, "accent": GOLD},
                               typography="Calibri; explicit type hierarchy",
                               pagination="Explicit authored page breaks for DOCX",
                               motif="Numbered teal blocks and calibrated financial bars for PPTX",
                               geometry="US Letter 0.8in side margins; PPTX 33.867 x 19.05cm; XLSX four navigable sheets",
                               engine="OfficeCLI atomic DOM batches"))
    save(ROOT / batch, commands)
    job = dict(document_id=doc["id"], layout_file=layout, batch_file=batch,
               render_targets=render_targets, export_bindings=bindings)
    if w:
        job["expected_formulas"] = w.expected
    jobs.append(job)
    layouts.append(dict(document_id=doc["id"], layout_revision="1",
                        layout_sha256=company.file_hash(layout_path)))
    discussion = sum(len(re.findall(r"\b[\w'-]+\b", s["text"])) for s in b.sections if s["role"] == "discussion")
    if kind == "meeting_transcript" and discussion < 1500:
        raise ValueError(f"{doc['id']}: discussion only {discussion} words")
    measurements.append(dict(document_id=doc["id"], type=kind,
                             draft_words=sum(len(re.findall(r"\b[\w'-]+\b", s["text"])) for s in b.sections),
                             discussion_words=discussion, requirements=16 if kind == "product_spec" else None,
                             analysis_modules=6 if w else None))
pilots = []
for kind in canon["request"]["document_types"]:
    candidates = [d for d in plan["documents"] if d["type"] == kind]
    pilot = max(candidates, key=lambda d: (
        sum(len(s["text"]) for s in drafts[d["id"]]["sections"]), len(d["title"]), d["publication_date"]))
    pilots.append(pilot["id"])
execution = company.generation_execution_plan(
    canon, plan, layouts, pilots, post_generation_qa={"enabled": False, "policy": None},
    revision=max(d["revision"] for d in plan["documents"]), flagged_ids=pilots)
save(ROOT / "build" / "execution-plan.json", execution)
save(ROOT / "build" / "jobs.json", jobs)
save(ROOT / "build" / "jobs-pilots.json", [j for j in jobs if j["document_id"] in pilots])
save(ROOT / "build" / "draft-measurements.json", measurements)
save(ROOT / "build" / "family-review-capacity.json", {
    "families": [{"type": kind, "documents": sum(d["type"] == kind for d in plan["documents"]),
                  "pilot_id": next(i for i in pilots if next(d for d in plan["documents"] if d["id"] == i)["type"] == kind)}
                 for kind in canon["request"]["document_types"]],
    "pilot_image_capacity": "All per-page, per-slide and per-range pilot images will be delivered and inspected.",
    "additional_qa": "not_requested", "singleton_families": [], "all_selected_warning": False})
print(json.dumps({"drafts": len(drafts), "pilots": pilots, "measurements": measurements}, indent=2))
