import importlib.util
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
spec = importlib.util.spec_from_file_location(
    "company", ROOT.parents[1] / ".github" / "skills" / "company-content-generator" / "scripts" / "company.py")
company = importlib.util.module_from_spec(spec)
spec.loader.exec_module(company)
plan = company.load(ROOT / "build" / "plan.json")


def money(value):
    return f"${value:,.0f}"


def pct(value):
    return f"{value:.1%}"


class Writer:
    def __init__(self, context):
        self.context = context
        self.doc = context["document"]
        self.index = {r["id"]: r for r in context["records"]}
        self.anchor = self.index[self.doc["anchor_ids"][0]]
        self.claims = {}
        self.sections = []
        self.slides = []
        self.year = self.anchor["data"].get("year")

    def data(self, identifier):
        return self.index[identifier]["data"]

    def name(self, identifier):
        return self.data(identifier)["name"]

    def role(self, identifier, as_of=None):
        snapshot = company.role_snapshot(self.context["records"], as_of or self.doc["publication_date"])
        return snapshot[identifier]["data"]["title"]

    def section(self, role, heading, text, sources, **extra):
        claim_ids = []
        for identifier, field in sources:
            path = ["data", field] if field else ["data"]
            key = (identifier, field)
            if key not in self.claims:
                self.claims[key] = {
                    "id": f"claim-{len(self.claims)+1}", "record_id": identifier,
                    "path": path, "value": company.resolve(self.index[identifier], path)}
            claim_ids.append(self.claims[key]["id"])
        self.sections.append(dict(role=role, heading=heading or "News", text=text,
                                  claim_ids=list(dict.fromkeys(claim_ids)), **extra))

    def slide(self, role, heading, body, sources, **extra):
        self.section(role, heading, "\n\n".join(body), sources)
        self.slides.append(dict(heading=heading, body=body, **extra))

    def requirements(self):
        a = self.anchor["data"]
        return [r for r in self.context["records"] if r["kind"] == "requirement"
                and r["data"]["product_id"] == a["product_id"]
                and r["data"]["version"] == a["version"]]

    def release(self):
        a = self.anchor["data"]
        product = self.name(a["product_id"])
        strategy = self.data(f"strategy-{a['year']}")
        src = [(self.anchor["id"], None), ("company", None), ("person-avery", "name"),
               ("person-jordan", "name")]
        self.section("dateline_and_lead", "", (
            f"BOULDER, COLORADO - {self.doc['publication_date']} - AstroDig LLC today released "
            f"{product} {a['version']} for {a['campaign']}. The configuration is intended for "
            f"{a['scope'].lower()}, providing contracted engineering teams a bounded way to organize "
            "reference work, review evidence, and decide whether the next development step is justified. "
            "The release is part of the company's staged aerospace development program, not an announcement "
            "of an operational asteroid mine. AstroDig continues to pursue a resource demonstration by "
            "2032, with intermediate commitments dependent on technical evidence and available resources."
        ), src)
        self.section("announcement_details", "", (
            f"The current priority is {strategy['title'].lower()}. {a['release_rationale']} "
            "The release gives teams a defined configuration against which to record inputs, conduct "
            "reference checks, and review discrepancies. That boundary matters when test methods, "
            "material assumptions, or engineering interfaces change: a result should be attributed "
            "to the configuration actually evaluated, rather than to the newest product name.\n\n"
            f"Campaign access requires an {a['adoption_requirements'][0]}, a "
            f"{a['adoption_requirements'][1]}, and a {a['adoption_requirements'][2]}. "
            "Teams first register the configuration and inputs, then run the approved reference checks. "
            "They execute the bounded campaign only within the reviewed scope, examine the evidence "
            "and discrepancies, and either accept or hold the next configuration. Critical discrepancies "
            "block progression; an unresolved issue is not converted into accepted evidence by changing "
            "the campaign schedule.\n\n"
            "The configuration supports mission trade studies, payload verification campaigns, and "
            "readiness reviews. Mission-design teams can compare candidate operating assumptions without "
            "asserting resource recoverability. Instrument-development teams can connect a controlled "
            "sequence to acceptance evidence and configuration identity. Verification managers can "
            "separate accepted results from dependencies that still require resolution. Those are "
            "engineering uses, not customer production results or claims of commercial extraction.\n\n"
            f"The principal development risk is {a['principal_risk']}. {a['limitation']} "
            f"The release excludes {a['not_in_scope'].lower()}. These exclusions preserve the distinction "
            "between useful ground-development work and evidence sufficient for a flight or resource "
            "decision. A ground release does not close the later spacecraft, characterization, or "
            "extraction gates."
        ), src + [(f"strategy-{a['year']}", None)])
        self.section("attributed_quote", "", (
            f'{self.name("person-avery")}, {self.role("person-avery")} of AstroDig LLC, '
            f'said: "{a["approved_quote"]}"'
        ), src + [(r["id"], "title") for r in self.context["active_roles"]
                  if r["data"]["person_id"] == "person-avery"])
        self.section("announcement_details", "", (
            f'{a["availability"]} {a["next_step"]} The initial review is intended to establish '
            "the question being tested, the configuration boundary, and the evidence required for "
            "acceptance. It does not promise flight access, certify a resource, or substitute a product "
            "release for a mission-readiness decision."
        ), src)
        self.section("company_boilerplate", "About AstroDig LLC",
                     self.data("company")["boilerplate"], [("company", "boilerplate")])
        self.section("media_contact", "Media contact",
                     f'{self.name("person-jordan")}\n{self.role("person-jordan")}\n'
                     f'{self.data("company")["media_email"]}',
                     [("person-jordan", "name"), ("company", "media_email")] +
                     [(r["id"], "title") for r in self.context["active_roles"]
                      if r["data"]["person_id"] == "person-jordan"])

    def specification(self):
        a = self.anchor["data"]
        reqs = self.requirements()
        product = self.name(a["product_id"])
        src = [(self.anchor["id"], None), (a["product_id"], None)]
        self.section("scope", "Purpose and controlled scope",
                     f'{product} {a["version"]} supports {a["campaign"]}. This baseline is limited to '
                     f'{a["scope"].lower()}. {a["release_rationale"]} The specification controls the '
                     "relationship between the accepted configuration, its inputs or test materials, "
                     "and the verification evidence used for a development decision. It does not "
                     f'authorize {a["not_in_scope"].lower()}. {a["limitation"]}', src)
        self.section("scope", "Configuration, ownership, and dependencies",
                     f'{self.name("person-quinn")} owns the product scope; '
                     f'{self.name("person-sasha")} owns verification. Teams shall register configuration '
                     "and inputs, complete reference checks, execute the bounded campaign, and review "
                     "evidence before accepting a subsequent configuration. Adoption requires "
                     f'{", ".join(a["adoption_requirements"])}. A dependency named below is part of '
                     "the requirement boundary. Replacing that dependency without review would change "
                     "the basis on which the acceptance result is interpreted.",
                     src + [("person-quinn", "name"), ("person-sasha", "name")])
        for r in reqs:
            d = r["data"]
            self.section("requirements", f'{d["requirement_id"]} - {d["group"]}',
                         f'{d["statement"]}\n\nAcceptance: {d["acceptance_criterion"]}\n\n'
                         f'Verification method: {d["verification_method"]}. Dependency: {d["interface"]}. '
                         f'Verification owner: {self.name(d["owner_id"])}.\n\n'
                         f'Disposition: {d["failure_disposition"]} The acceptance record shall identify '
                         "the evaluated configuration and retain the observed outcome. An unsuccessful "
                         "verification result is evidence requiring disposition, not permission to "
                         "replace the criterion after the run. Closure belongs in the signed verification "
                         "matrix before a development release advances to the next campaign.",
                         [(r["id"], None), (d["owner_id"], "name"), (self.anchor["id"], "release_controls")])
        self.section("acceptance_criteria", "Acceptance matrix and release gate",
                     "Each requirement above is a distinct testable control, not an optional checklist "
                     "item. The verification matrix shall retain the requirement identifier, method, "
                     "observed result, responsible reviewer, and discrepancy disposition. Critical "
                     "discrepancies block progression. Ground-development release status and flight "
                     "status remain separate throughout review. Any change requires reviewed evidence "
                     "before it is incorporated into the next accepted configuration.\n\n"
                     f'The current campaign risk is {a["principal_risk"]}. {a["limitation"]} '
                     "The release gate accepts evidence for the defined ground scope only. It does "
                     "not establish resource recoverability, mission approval, or a commercial output "
                     "commitment. This distinction shall remain visible in the acceptance record and "
                     "in technical communication about the release.", src)
        self.requirement_rows = [r["data"] for r in reqs]

    def product_deck(self):
        a = self.anchor["data"]
        s = self.data(f"strategy-{self.year}")
        src = [(self.anchor["id"], None), (f"strategy-{self.year}", None), ("company", None)]
        self.slide("customer_problem", f'RegolithWorks {a["version"]}',
                   [a["campaign"].title(), "Defined ground evidence before the next mission decision."], src,
                   layout="cover", badge="GROUND DEVELOPMENT")
        self.slide("customer_problem", "The decision problem",
                   ["Changing inputs and configurations make results difficult to compare.",
                    f'Current development risk: {a["principal_risk"]}.',
                    a["limitation"]], src, layout="cards")
        self.slide("value_proposition", "A controlled campaign boundary",
                   [a["release_rationale"], f'Scope: {a["scope"]}.',
                    "Accepted evidence and unresolved dependencies stay separate."], src, layout="statement")
        reqs = self.requirements()
        # The milestone is already RegolithWorks; these are version-specific requirements.
        for start, title in [(0, "Containment and safe control"), (5, "Evidence and configuration"), (10, "Recovery and acceptance")]:
            subset = reqs[start:start + 5]
            self.slide("capabilities", title,
                       [f'{r["data"]["group"]}: {r["data"]["statement"]}' for r in subset],
                       src + [(r["id"], None) for r in subset], layout="grid" if start == 5 else "capabilities")
        self.slide("use_cases", "From inputs to a reviewable result",
                   a["workflow"], src, layout="workflow")
        self.slide("use_cases", "Three engineering uses",
                   [f'{u["name"]}: {u["task"]}' for u in a["use_cases"]], src, layout="cards")
        self.slide("capabilities", "What teams must bring",
                   a["adoption_requirements"], src, layout="cards")
        self.slide("capabilities", "Where the evidence stops",
                   [a["limitation"], f'Not in scope: {a["not_in_scope"]}.',
                    "A ground-development release is not flight qualification."], src, layout="statement")
        self.slide("next_steps", "Define the reference case",
                   [a["availability"], a["next_step"]],
                   src, layout="closing")

    def financial_sources(self):
        a = self.anchor["data"]
        year = a["year"]
        strategy = self.data(f"strategy-{year}")
        current = [r for r in self.context["records"] if r["kind"] == "financial"
                   and r["data"]["year"] == year and r["data"]["quarter"] is not None]
        if a["quarter"] is not None:
            current = [r for r in current if r["data"]["quarter"] <= a["quarter"]]
        prior = [r for r in self.context["records"] if r["kind"] == "financial"
                 and r["data"]["year"] == year - 1 and r["data"]["quarter"] == a["quarter"]]
        return a, strategy, sorted(current, key=lambda r: r["data"]["quarter"]), prior

    def financial_deck(self):
        a, s, qs, prior = self.financial_sources()
        annual = self.doc["type"] == "annual_report"
        period = str(a["year"]) if annual else f'Q{a["quarter"]} {a["year"]}'
        roles = ("financial_results", "business_review" if annual else "operating_review",
                 "risks" if annual else "outlook_and_risks")
        src = [(self.anchor["id"], None), (f"strategy-{self.year}", None), ("company", None)]
        self.slide(roles[0], f"{period} Operating Review",
                   [s["title"], "USD; calendar fiscal year; engineering and development operations."],
                   src, layout="cover", badge="INTERNAL")
        self.slide(roles[0], "Operating result at a glance",
                   [f'Revenue: {money(a["revenue_usd"])}',
                    f'Operating result: {money(a["operating_income_usd"])}',
                    f'Period-end headcount: {a["headcount"]}'], src, layout="stats")
        self.slide(roles[0], "Revenue to operating result",
                   [a["cost_driver"], "Development spending remains distinct from prospective mining economics."],
                   src, layout="chart",
                   chart={"labels": ["Revenue", "COGS", "Opex", "Operating result"],
                          "values": [a[k] for k in ["revenue_usd", "cogs_usd", "opex_usd", "operating_income_usd"]]})
        self.slide(roles[0], "Where recognized revenue comes from",
                   [a["revenue_policy"]], src, layout="chart",
                   chart={"labels": ["Engineering", "Instruments", "Mission analysis"],
                          "values": [a[k] for k in ["engineering_services_usd", "instrument_development_usd", "mission_analysis_usd"]]})
        self.slide(roles[0], "Development resources dominate spending",
                   ["Research and development is the largest operating-expense category.",
                    "Program operations and administration remain separately identifiable."], src, layout="chart",
                   chart={"labels": ["R&D", "Program operations", "Administration"],
                          "values": [a[k] for k in ["research_development_usd", "program_operations_usd", "corporate_administration_usd"]]})
        self.slide(roles[0], "Quarterly recognition profile",
                   ["Quarterly revenue reflects accepted contracted work, not asteroid sales.",
                    "Only periods available at publication are included."],
                   src + [(r["id"], None) for r in qs], layout="chart",
                   chart={"labels": [f'Q{r["data"]["quarter"]}' for r in qs],
                          "values": [r["data"]["revenue_usd"] for r in qs]})
        self.slide(roles[1], "Delivery and reference evidence",
                   [f'Completed work packages: {a["completed_work_packages"]}.',
                    f'Accepted reference runs: {a["accepted_reference_runs"]}.',
                    f'Rejected reference runs retained: {a["rejected_reference_runs"]}.'],
                   src, layout="stats")
        self.slide(roles[1], "Commitments are not collections",
                   [f'Contracted backlog at period end: {money(a["contracted_backlog_usd"])}.',
                    self.data(next(r["id"] for r in self.context["glossary"]
                                   if r["data"]["term"] == "Contracted backlog"))["definition"],
                    "No cash-flow or liquidity conclusion is derived from this snapshot."],
                   src + [(r["id"], "definition") for r in self.context["glossary"]
                          if r["data"]["term"] == "Contracted backlog"], layout="statement")
        self.slide(roles[1], "Development priority and tradeoff",
                   [s["decision_rationale"], f'Alternative: {s["alternative"]}.',
                    f'Disposition: {s["alternative_disposition"]}.'], src, layout="cards")
        self.slide(roles[2], "Risk, response, and evidence boundary",
                   [f'Principal risk: {s["principal_risk"]}.', s["control"], s["limitation"]],
                   src, layout="cards")
        if annual:
            self.slide(roles[0], "Annual reconciliation",
                       [f'Four quarters sum to revenue of {money(a["revenue_usd"])}.',
                        f'COGS: {money(a["cogs_usd"])}; opex: {money(a["opex_usd"])}.',
                        f'Operating result: {money(a["operating_income_usd"])}; headcount equals the Q4 snapshot.'],
                       src + [(r["id"], None) for r in qs], layout="statement")
            comparison = (
                [f'Prior-year revenue: {money(prior[0]["data"]["revenue_usd"])}.',
                 f'Current-year revenue: {money(a["revenue_usd"])}.',
                 f'Change: {money(a["revenue_usd"] - prior[0]["data"]["revenue_usd"])}; '
                 f'{pct(a["revenue_usd"] / prior[0]["data"]["revenue_usd"] - 1)}.']
                if prior else ["2013 is the founding reporting year.", "No prior-year company actuals are available.", "The four-quarter profile establishes the comparison baseline."])
            self.slide(roles[0], "Available-year comparison", comparison,
                       src + [(r["id"], None) for r in prior], layout="stats")
            snap = company.role_snapshot(self.context["records"], a["period_end"])
            ids = ["person-avery", "person-riley", "person-casey" if a["year"] >= 2020 else "person-morgan", "person-quinn"]
            self.slide("organization", "Year-end accountable leadership",
                       [f'{self.name(p)} - {snap[p]["data"]["title"]}' for p in ids],
                       src + [(p, "name") for p in ids] + [(snap[p]["id"], "title") for p in ids],
                       layout="capabilities")
            self.slide(roles[2], "The mission gates remain prospective",
                       [f'{g["year"]}: {g["gate"]}; {g["status"]}; depends on {g["dependency"]}.'
                        for g in s["gates"]], src, layout="workflow")
        else:
            self.slide(roles[2], "Labeled operating outlook",
                       [a["outlook"], "No commercial asteroid extraction revenue is assumed.",
                        "The 2032 demonstration remains conditional."], src, layout="closing")

    def board_deck(self):
        a, s, qs, prior = self.financial_sources()
        b = self.data(f"board-{self.year}")
        src = [(self.anchor["id"], None), (f"strategy-{self.year}", None),
               (f"board-{self.year}", None), ("company", None)]
        self.slide("executive_summary", "Capital and Readiness",
                   [s["title"], f'Q3 {self.year}; requested decisions, not recorded approvals.'],
                   src, layout="cover", badge="BOARD CONFIDENTIAL")
        self.slide("executive_summary", "The decision in front of the board",
                   [b["preferred_option"], b["tradeoff"], s["decision_rationale"]], src, layout="statement")
        self.slide("performance", "Services fund only part of development",
                   [f'Q3 revenue: {money(a["revenue_usd"])}.',
                    f'Q3 operating result: {money(a["operating_income_usd"])}.',
                    "Operating losses remain visible; expected asteroid sales are excluded."], src, layout="stats")
        self.slide("performance", "Income and spending bridge",
                   [a["cost_driver"]], src, layout="chart",
                   chart={"labels": ["Revenue", "COGS", "Opex", "Operating result"],
                          "values": [a[k] for k in ["revenue_usd", "cogs_usd", "opex_usd", "operating_income_usd"]]})
        self.slide("performance", "Available-quarter performance",
                   ["Accepted contracted work drives recognition; the series is not a mining forecast."],
                   src + [(r["id"], None) for r in qs], layout="chart",
                   chart={"labels": [f'Q{r["data"]["quarter"]}' for r in qs],
                          "values": [r["data"]["revenue_usd"] for r in qs]})
        self.slide("performance", "Capacity and commitments",
                   [f'Period-end headcount: {a["headcount"]}.',
                    f'Contracted backlog: {money(a["contracted_backlog_usd"])}.',
                    f'Completed work packages: {a["completed_work_packages"]}.'], src, layout="stats")
        self.slide("performance", "Current ground-development scope",
                   [s["campaign"], s["release_decision"], s["limitation"]], src, layout="statement")
        self.slide("risks", "Technical risk and stop condition",
                   [s["principal_risk"], s["control"], "Retain failed trials and owned discrepancy dispositions."],
                   src + [(f"meeting-{self.year}", None)], layout="cards")
        self.slide("risks", "Strategy is a sequence of conditional gates",
                   [f'{g["year"]}: {g["gate"]}; depends on {g["dependency"]}.' for g in s["gates"]],
                   src, layout="workflow")
        self.slide("decisions_requested", "Preferred scope versus broader features",
                   [f'Preferred: {b["preferred_option"]}.', f'Alternative: {b["alternative"]}.',
                    f'Tradeoff: {b["tradeoff"]}.'], src, layout="cards")
        self.slide("decisions_requested", "Requested actions and accountable owners",
                   [f'{d["topic"]}: {d["decision"]} Owner: {self.name(d["owner_id"])}; due {d["due_on"]}.'
                    for d in b["decisions_requested"]],
                   src + [(d["owner_id"], "name") for d in b["decisions_requested"]],
                   layout="capabilities")
        self.slide("decisions_requested", "Approve bounded development, not a mission",
                   [b["decision_status"], "No mission authorization or resource reserve is represented.",
                    "Keep the 2032 objective conditional on evidence, funding, and access."],
                   src, layout="closing")

    def workbook(self):
        a, s, qs, prior = self.financial_sources()
        src = [(self.anchor["id"], None), (f"strategy-{self.year}", None), ("company", None)]
        model = self.doc["type"] == "financial_model"
        roles = ["source_actuals", "calculations", "assumptions_and_notes"] if model else [
            "financial_results", "operating_review", "outlook_and_risks"]
        self.section(roles[0], "Period and recognized source actuals",
                     f'USD; Q{a["quarter"]} {a["year"]}; period ended {a["period_end"]}. '
                     f'Revenue {money(a["revenue_usd"])}; COGS {money(a["cogs_usd"])}; '
                     f'opex {money(a["opex_usd"])}; operating result {money(a["operating_income_usd"])}. '
                     f'{a["revenue_policy"]}', src)
        interpretations = [
            ("Income reconciliation", "Does recognized revenue cover delivery and operating costs?",
             "Operating result remains negative after both delivery costs and development spending; the reconciliation isolates the reported loss rather than presenting revenue as profit."),
            ("Margins and spending intensity", "How much revenue remains after delivery costs, and how much is consumed by operating spending?",
             "Gross margin and operating margin answer different questions. A positive delivery margin does not imply that the development program is self-financing."),
            ("Service revenue mix", "Which contracted activities produce recognized revenue?",
             "Engineering services, instrument development, and mission analysis sum to recognized revenue. No asteroid commodity-sales category is introduced."),
            ("Operating cost allocation", "Where is development capacity being funded?",
             "Research and development, program operations, and corporate administration reconcile to operating expenses. The allocation makes the development priority visible without treating costs as capitalized assets."),
            ("Workforce and delivery productivity", "What output is supported by the period-end workforce snapshot?",
             "Revenue and completed work packages per period-end employee are descriptive operating ratios, not labor-efficiency guarantees. Headcount is a point-in-time measure and is not summed across quarters."),
            ("Available-period trends and commitments", "How do available quarter results and contracted commitments relate?",
             "Quarter trends use only known periods. Backlog is a remaining-work snapshot, not cash or recognized revenue; rejected runs remain visible alongside accepted evidence."),
        ]
        if model:
            interpretations = [
                interpretations[0],
                ("Contracted service economics", "How do the three service categories reconcile, and what share comes from engineering?",
                 "Recognized service categories reconcile to revenue. Engineering share describes the revenue base; no category-level profit is inferred because delivery costs are not allocated by service category."),
                ("Development spending intensity", "How does research spending relate to revenue, workforce and total operating expense?",
                 "Research spending per period-end employee and relative to revenue exposes the development load. These ratios are descriptive; no capitalization, cash burn or return on investment is inferred."),
                ("Delivery capacity and evidence quality", "How do delivered work and accepted or rejected trials relate to the workforce snapshot?",
                 "Completed work packages and reference acceptance are separate operating measures. Rejected trials remain in the denominator, and workforce ratios do not claim unused capacity."),
                ("Quarter-to-quarter operating bridge", "Which recognized revenue and cost movements explain the change from Q2 to Q3?",
                 "The bridge reconciles revenue change, delivery-cost change and operating-expense change to the reported operating-result movement. It is an actual-period bridge, not a forecast or a causal attribution."),
                ("Backlog commitment scale", "How large is remaining contracted work relative to current-quarter revenue and period-end headcount?",
                 "Backlog coverage and backlog per employee describe commitment scale at a point in time. They are not collection forecasts, cash balances, resource reserves or proof of available delivery capacity."),
            ]
        for title, question, interpretation in interpretations:
            self.section(roles[1], title, f'{question}\n\n{interpretation}', src)
        self.section(roles[2], "Methodology and limitations",
                     f'{s["decision_rationale"]} Principal risk: {s["principal_risk"]}. '
                     f'{s["limitation"]} {a["outlook"]} No forecast, balance sheet, liquidity estimate, '
                     "resource valuation, or cash-flow conclusion is supplied. Formula-linked analyses "
                     "use recognized USD amounts and period-end operating snapshots; changing source "
                     "inputs updates the calculations without changing the stated business scope.",
                     src + [(r["id"], None) for r in qs])
        self.model = {"anchor": a, "quarters": [r["data"] for r in qs],
                      "source_date": self.anchor["known_on"], "modules": interpretations,
                      "strategy": s}

    def transcript(self):
        a = self.anchor["data"]
        m = self.data(f"meeting-{self.year}")
        s = self.data(f"strategy-{self.year}")
        milestone = self.data(m["product_milestone_id"])
        participants = m["participants"]
        finance = participants[1]
        src = [(self.anchor["id"], None), (f"meeting-{self.year}", None),
               (f"strategy-{self.year}", None), (m["product_milestone_id"], None), ("company", None)]
        src += [(p, "name") for p in participants]
        src += [(r["id"], "title") for r in self.context["active_roles"]
                if r["data"]["person_id"] in participants]
        self.section("participants", "Meeting details and participants",
                     f'Date: {m["meeting_on"]}\nAgenda: ' + "; ".join(m["agenda"]) + "\n" +
                     "\n".join(f'{self.name(p)} - {self.role(p)}' for p in participants), src)
        # Each agenda thread moves through question, alternative, clarification, and resolution.
        exchanges = [
            ("avery", f"Let us start with the operating record rather than the mission ambition. Q3 revenue is {money(a['revenue_usd'])}, and the operating result is {money(a['operating_income_usd'])}. We need to understand what the current services activity supports and what it does not. I do not want the 2032 objective to turn a development cost into a presumed future sale. What is the cleanest distinction for this review?"),
            (finance.replace("person-", ""), f"The recognized revenue comes from engineering services, instrument development, and mission analysis. COGS is {money(a['cogs_usd'])}; operating expenses are {money(a['opex_usd'])}. Revenue less those two cost categories gives the reported operating loss. The service activity supports part of the program, but not all of it. We should retain that reconciliation whenever we discuss development capacity. Expected asteroid sales are not included in these results."),
            ("devon", f"The other distinction is between delivered work and remaining commitments. We completed {a['completed_work_packages']} work packages in the quarter and have {money(a['contracted_backlog_usd'])} of contracted backlog at period end. That backlog is work still to be performed, not money already collected. If we call the same people available for a development campaign and committed to a service package, we have counted their capacity twice. The schedule needs one resource view."),
            ("quinn", f"That is relevant to {s['campaign']}. I would like broader functionality, but it would compete with the current campaign and with contracted delivery. The alternative on the table is to expand optional functionality before verification closure. The benefit is more breadth; the problem is that breadth does not resolve the evidence gaps in the reference configuration. I would rather keep the campaign boundary intelligible than make the release appear more mature by adding unfinished scope."),
            ("riley", f"I agree with keeping a bounded reference configuration, but I want to avoid equating a narrow scope with a completed technical case. The principal risk remains {s['principal_risk']}. We must distinguish what has been tested from what has merely been retained in the design. A service-funded development activity is still a development activity. The release decision authorizes the ground campaign only; it does not authorize a spacecraft, an extraction mission, or a resource claim."),
            ("avery", "Then the finance presentation should show the service contribution and the operating loss together. We can discuss resource allocation without a cash-flow or mining-economics conclusion that the record does not support. I also want the remaining-work snapshot labeled clearly. It is useful for sequencing commitments, but it should not be allowed to answer a different question about liquidity. Does anyone disagree with keeping that boundary explicit in the review?"),
            (finance.replace("person-", ""), "No. The distinction is also important when someone asks whether development is self-financing. A positive gross margin on accepted work is not the same as a positive operating result after research and verification spending. I will keep the expense allocation visible rather than netting it against a prospective mission benefit. The action is to maintain service delivery and report development spending separately from any prospective mining economics."),
            ("sasha", f"On the evidence side, we recorded {a['accepted_reference_runs']} accepted reference runs and {a['rejected_reference_runs']} rejected runs. A rejected run belongs in the record. Removing it would make the campaign look cleaner without changing the underlying result. We should ask whether the discrepancy has an owner and a disposition, not whether it is convenient to include in a readiness discussion. Accepted evidence and unresolved dependencies need separate treatment."),
            ("riley", f"The current release is RegolithWorks {milestone['version']}, and its scope is {milestone['scope'].lower()}. The workflow starts with registered configuration and inputs, then reference checks, the bounded campaign, and evidence review. The final step is acceptance or a hold on the next configuration. That last distinction is substantive: reaching the review meeting is not itself an acceptance result. A critical discrepancy remains a stop condition even when the next campaign is desirable."),
            ("quinn", f"The scope priority should remain {s['campaign']}. {s['decision_rationale']} I am not proposing that the reference campaign answer every future mission question. I am proposing that it answer its own question well enough to support the next design decision. That means controlling the inputs and configuration, and identifying which interpretation is justified by the evidence. If we change scope halfway through, we need to explain what comparison is still valid."),
            ("sasha", "The acceptance matrix provides that discipline. Each requirement needs a method, a result, an owner, and a disposition. The dependency named with the requirement is part of the verification boundary, not incidental paperwork. If the dependency changes, the reviewer must determine whether the old result still applies. We should not clear a failed criterion by silently changing the criterion after the run. That would erase the reason the hold existed."),
            ("devon", "I want to connect that to campaign scheduling. A hold should remain visible in the resource sequence rather than being hidden behind an assumed continuation date. Contracted engineering commitments still need delivery attention while a development discrepancy is being resolved. We can sequence both against the named resources, but we cannot treat the verification work as optional simply because the service schedule is already full. The tradeoff needs an explicit scope decision."),
            ("avery", "The broader-feature alternative is therefore deferred, not quietly carried as additional work. That is a real choice: it limits breadth while protecting reference evidence and contracted delivery. I want the review to say that plainly. We are not deciding that optional capability has no value; we are deciding that its value does not justify advancing it before the current acceptance basis is closed. The bounded campaign remains the development priority."),
            ("riley", f"The evidence boundary also needs to travel with the release description. {s['limitation']} No technical communication should imply that a ground development release has established flight qualification. The product configuration can be useful to an engineering team without proving a mission-level outcome. The test criterion, the system boundary, and the remaining dependencies all matter. We should preserve those distinctions when explaining progress to people who were not in the campaign."),
            ("quinn", "For adoption, a team needs an approved input or test-material specification, a named technical reviewer, and a configuration-controlled reference case. Those are not administrative add-ons to the release. They define what the engineering evidence will mean. If a team wants a different campaign question, we should first review the boundary and the acceptance evidence rather than granting access on the assumption that the current configuration answers the new question."),
            ("sasha", f"I will own closure of critical discrepancies before optional campaign scope expands. The due date in this review is {m['decisions'][1]['due_on']}. Closure means a documented disposition in the verification record, not the disappearance of a failed run from a summary. If evidence still requires a hold, the hold remains. That is the control we need to retain as the configuration moves from reference work into the next bounded campaign."),
            ("avery", "Now let us separate the long-term goal from the actions we can authorize here. The company is pursuing an asteroid resource demonstration by 2032. That objective is contingent on flight validation, funding, launch access, and mission authorization. It is not a recorded approval of those dependencies. I want to keep the goal useful for strategy while preventing it from being interpreted as a promise that the intermediate gates have already been passed."),
            ("riley", "The planned 2028 gate is a prospecting spacecraft design review, dependent on verified payload interfaces and approved funding. The planned 2030 gate is direct asteroid characterization, dependent on launch access, mission authorization, and spacecraft performance. The 2032 extraction demonstration is a goal dependent on characterization evidence, extraction readiness, and funded mission operations. Each later gate needs evidence from the preceding work. A date on the sequence cannot substitute for that evidence."),
            (finance.replace("person-", ""), "That sequence is why I do not want prospective resource economics inserted into the operating record. The current business earns engineering and development revenue. The strategic goal may guide priorities, but it does not create recognized commodity revenue, a cash balance, or a resource valuation. We can compare the current cost allocation and available-quarter results. We cannot derive a full funding plan merely by renaming an operating loss as an investment."),
            ("devon", "From operations, the same conditional language needs to appear wherever the mission sequence is discussed. Planned launch access is not booked launch access. A development release is not mission authorization. If a dependency is unresolved, the next commitment must preserve that unresolved state. Otherwise a program schedule can appear internally consistent while depending on a decision no one has made. The action list should therefore name the work we own, not imply external approvals."),
            ("quinn", f"For this year, the practical interpretation is {s['decision_rationale'].lower()} The current campaign can contribute evidence toward that interpretation, but it cannot close all later gates. I am comfortable retaining the 2032 goal while making the product boundary explicit. The scope action I own is to retain {s['campaign']} as the development priority. Optional breadth remains deferred until the reference evidence is accepted, and any changed boundary must be reviewed."),
            ("sasha", "I would add that an accepted ground result must retain its stated limitations when it is reused in a readiness argument. We should not summarize an accepted trial as though it directly established a resource or flight conclusion. The reference configuration, the evaluated input, and the acceptance method remain part of the evidence. A future reviewer needs enough context to know what was proved and which dependencies still require a different test or decision."),
            ("avery", f"Then we have four actions, all due {m['decisions'][0]['due_on']}: retain the scoped development priority, close critical discrepancies before optional expansion, maintain service delivery with distinct development-spending reporting, and retain the conditional demonstration goal. Those actions do not approve a mission. They make the current work accountable. I will own the strategic wording so that the 2032 objective stays connected to evidence and resources rather than becoming a substitute for either."),
            (finance.replace("person-", ""), "That is a workable resolution. The next review can examine the action dispositions against the same source distinctions: recognized work, delivery costs, operating spending, remaining commitments, and accepted or rejected reference evidence. If one of those records changes, the interpretation should change with it. We should not carry a favorable conclusion forward simply because it appeared in the previous review. The owned actions give us a basis for checking progress without inventing certainty."),
        ]
        dialogue = []
        for number, (key, text) in enumerate(exchanges):
            stamp = f"{(number * 2) // 60:02}:{(number * 2) % 60:02}:00"
            dialogue.append(f"[{stamp}] {self.name('person-' + key)}: {text}")
        self.section("discussion", "Discussion", "\n\n".join(dialogue), src)
        self.section("decisions", "Resolutions",
                     "\n\n".join(f'{d["topic"]}: {d["decision"]}' for d in m["decisions"]), src)
        self.section("actions", "Owned actions",
                     "\n\n".join(f'{self.name(d["owner_id"])} - {d["decision"]} Due: {d["due_on"]}.'
                                 for d in m["decisions"]), src)

    def build(self):
        kind = self.doc["type"]
        if kind == "press_release":
            self.release()
        elif kind == "product_spec":
            self.specification()
        elif kind == "product_presentation":
            self.product_deck()
        elif kind == "meeting_transcript":
            self.transcript()
        elif kind == "board_deck":
            self.board_deck()
        elif self.doc["format"] == "pptx":
            self.financial_deck()
        else:
            self.workbook()
        result = {
            "document_id": self.doc["id"], "canon_sha256": self.context["canon_sha256"],
            "plan_sha256": self.context["plan_sha256"], "title": self.doc["title"],
            "publication_date": self.doc["publication_date"], "content_depth": "medium",
            "claims": list(self.claims.values()), "sections": self.sections,
            "slides": self.slides,
        }
        if hasattr(self, "model"):
            result["model"] = self.model
        if hasattr(self, "requirement_rows"):
            result["requirement_rows"] = self.requirement_rows
        words = len(re.findall(r"\b[\w'-]+\b", "\n".join(s["text"] for s in self.sections)))
        discussion_words = sum(len(re.findall(r"\b[\w'-]+\b", s["text"]))
                               for s in self.sections if s["role"] == "discussion")
        result["measurements"] = dict(delivered_words=words, discussion_words=discussion_words,
                                      meaningful_slides=len(self.slides),
                                      distinct_testable_requirements=len(getattr(self, "requirement_rows", [])),
                                      useful_analysis_modules=6 if hasattr(self, "model") else 0)
        target = next(t for t in company.load(ROOT / "build" / "depth-targets.json")["documents"]
                      if t["document_id"] == self.doc["id"])
        actual = result["measurements"][target["measure"]]
        if not target["target_min"] <= actual <= target["target_max"]:
            raise ValueError(f'{self.doc["id"]}: {target["measure"]}={actual}, expected '
                             f'{target["target_min"]}-{target["target_max"]}')
        company.check_draft(company.load(ROOT / "canon.json"), plan, result)
        return result


selected = set(sys.argv[1:])
for d in plan["documents"]:
    if selected and d["id"] not in selected:
        continue
    path = ROOT / "build" / "drafts" / f'{d["id"]}.json'
    if path.exists():
        continue
    context = company.load(ROOT / "build" / "contexts" / f'{d["id"]}.json')
    result = Writer(context).build()
    company.save(path, result)
    print(d["id"], d["type"], result["measurements"])
