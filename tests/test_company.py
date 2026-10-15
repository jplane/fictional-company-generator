import copy
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / ".github" / "skills" / "company-content-generator" / "scripts" / "company.py"
SPEC = importlib.util.spec_from_file_location("company", SCRIPT)
company = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(company)


class CompanyTests(unittest.TestCase):
    def setUp(self):
        self.request = {
            "company_name": "Lumenvale Systems", "industry": "analytics software",
            "start_year": 2022, "years": 3, "documents": 45, "seed": 42,
        }
        self.canon = company.seed_canon(self.request)
        self.plan = company.create_plan(self.canon)

    def record(self, identifier):
        return next(r for r in self.canon["records"] if r["id"] == identifier)

    def draft(self, doc):
        record = self.record(doc["anchor_ids"][0])
        key = next(iter(record["data"]))
        return {
            "schema_version": 2,
            "corpus_id": self.canon["corpus_id"],
            "document_id": doc["id"], "canon_sha256": company.fingerprint(self.canon),
            "plan_sha256": company.fingerprint(self.plan),
            "title": doc["title"], "publication_date": doc["publication_date"],
            "content_depth": doc["content_depth"],
            "claims": [{"id": "c1", "record_id": record["id"],
                        "path": ["data", key], "value": record["data"][key]}],
            "sections": [{"role": role, "heading": role.replace("_", " ").title(),
                          "text": f"{record['data'][key]} Operating review and agreed next steps.",
                          "claim_ids": ["c1"]} for role in doc["required_sections"]],
        }

    def test_seed_and_plan_determinism(self):
        self.assertEqual(self.canon, company.seed_canon(self.request))
        self.assertEqual(self.plan, company.create_plan(self.canon))
        changed = dict(self.request, seed=43)
        self.assertNotEqual(self.canon, company.seed_canon(changed))

    def test_financials_reconcile(self):
        index = company.validate_canon(self.canon)
        for year in (2022, 2023, 2024):
            annual = index[f"financial-{year}-annual"]["data"]
            self.assertEqual(annual["revenue_usd"], sum(
                index[f"financial-{year}-q{q}"]["data"]["revenue_usd"] for q in range(1, 5)))
            self.assertEqual(annual["headcount"], index[f"financial-{year}-q4"]["data"]["headcount"])

    def test_reject_bad_income(self):
        self.record("financial-2022-q1")["data"]["operating_income_usd"] += 1
        with self.assertRaisesRegex(company.Invalid, "income does not reconcile"):
            company.validate_canon(self.canon)

    def test_reject_bad_annual_total(self):
        annual = self.record("financial-2022-annual")["data"]
        annual["revenue_usd"] += 1
        annual["operating_income_usd"] += 1
        with self.assertRaisesRegex(company.Invalid, "does not sum quarters"):
            company.validate_canon(self.canon)

    def test_reject_premature_actuals(self):
        self.record("financial-2022-q1")["known_on"] = "2022-03-31"
        with self.assertRaisesRegex(company.Invalid, "premature actuals"):
            company.validate_canon(self.canon)

    def test_reject_premature_completed_event(self):
        for identifier in ("milestone-beacon-2022-launched", "acquisition-sable-closed"):
            with self.subTest(identifier=identifier):
                record = self.record(identifier)
                original = record["known_on"]
                record["known_on"] = "2022-08-01"
                with self.assertRaisesRegex(company.Invalid, "future .*completed"):
                    company.validate_canon(self.canon)
                record["known_on"] = original

    def test_reject_acquisition_close_without_announcement(self):
        self.canon["records"].remove(self.record("acquisition-sable-announced"))
        with self.assertRaisesRegex(company.Invalid, "close precedes announcement"):
            company.validate_canon(self.canon)

    def test_reject_duplicate_financial_period(self):
        duplicate = copy.deepcopy(self.record("financial-2022-q1"))
        duplicate["id"] = "duplicate-quarter"
        self.canon["records"].append(duplicate)
        with self.assertRaisesRegex(company.Invalid, "duplicate financial period"):
            company.validate_canon(self.canon)

    def test_reject_missing_quarter(self):
        self.canon["records"].remove(self.record("financial-2022-q1"))
        with self.assertRaises(company.Invalid):
            company.validate_canon(self.canon)

    def test_reject_duplicate_ids(self):
        self.canon["records"].append(copy.deepcopy(self.record("company")))
        with self.assertRaisesRegex(company.Invalid, "Duplicate record"):
            company.validate_canon(self.canon)

    def test_reject_dangling_and_future_refs(self):
        role = self.record("role-casey-cfo")
        role["refs"].append("nonexistent")
        with self.assertRaisesRegex(company.Invalid, "invalid ref"):
            company.validate_canon(self.canon)
        role["refs"].pop()
        role["refs"].append("financial-2024-annual")
        with self.assertRaisesRegex(company.Invalid, "not yet knowable"):
            company.validate_canon(self.canon)

    def test_reject_role_cycle(self):
        role = self.record("role-avery-initial")
        role["data"]["reports_to"] = "person-morgan"
        role["refs"].append("person-morgan")
        with self.assertRaisesRegex(company.Invalid, "reporting cycle"):
            company.validate_canon(self.canon)

    def test_reject_ambiguous_roles(self):
        role = copy.deepcopy(self.record("role-casey-cfo"))
        role["id"] = "duplicate-assignment"
        self.canon["records"].append(role)
        with self.assertRaisesRegex(company.Invalid, "ambiguous role"):
            company.validate_canon(self.canon)

    def test_role_transition(self):
        before = company.role_snapshot(self.canon["records"], "2023-06-30")
        after = company.role_snapshot(self.canon["records"], "2023-07-01")
        self.assertEqual(before["person-casey"]["data"]["title"], "Finance Director")
        self.assertEqual(after["person-casey"]["data"]["title"], "CFO")
        self.assertEqual(after["person-morgan"]["data"]["title"], "Chief Strategy Officer")

    def add_unused_org_entries(self):
        self.canon["records"].extend([
            {"id": "person-observer", "kind": "person",
             "known_on": "2022-01-01", "effective_on": "2022-01-01",
             "data": {"name": "Finley Brooks", "title": "Advisory Profile"}, "refs": []},
            {"id": "dept-outreach", "kind": "department",
             "known_on": "2022-01-01", "effective_on": "2022-01-01",
             "data": {"name": "Outreach"}, "refs": []},
        ])

    def test_org_chart_is_bounded_and_preserves_role_history(self):
        chart = company.org_chart_document(self.canon)
        self.assertEqual(chart["schema_version"], 2)
        self.assertEqual(len(chart["people"]), 6)
        self.assertEqual(len(chart["departments"]), 4)
        self.assertEqual(len(chart["roles"]), 8)
        self.assertLess(len(chart["people"]), self.record("financial-2022-q1")["data"]["headcount"])
        roles = {r["id"]: r for r in chart["roles"]}
        self.assertEqual(roles["role-casey-initial"]["data"]["title"], "Finance Director")
        self.assertEqual(roles["role-casey-cfo"]["data"]["title"], "CFO")
        self.assertEqual(roles["role-casey-cfo"]["effective_on"], "2023-07-01")
        self.assertEqual(roles["role-casey-cfo"]["data"]["reports_to"], "person-avery")
        self.assertEqual(roles["role-morgan-initial"]["data"]["title"], "CFO")
        self.assertEqual(roles["role-morgan-strategy"]["data"]["title"], "Chief Strategy Officer")
        for collection in ("people", "departments", "roles"):
            for record in chart[collection]:
                self.assertEqual(record, self.record(record["id"]))

    def test_unused_people_and_departments_are_excluded_from_writer_grounding(self):
        self.add_unused_org_entries()
        company.validate_canon(self.canon)
        self.plan = company.create_plan(self.canon)
        chart = company.org_chart_document(self.canon)
        self.assertNotIn("person-observer", {r["id"] for r in chart["people"]})
        self.assertNotIn("dept-outreach", {r["id"] for r in chart["departments"]})
        context = company.get_context(self.canon, self.plan, self.plan["documents"][0]["id"])
        self.assertNotIn("person-observer", {r["id"] for r in context["records"]})
        self.assertNotIn("dept-outreach", {r["id"] for r in context["records"]})

    def test_reference_adds_named_contact_without_fabricating_a_role(self):
        self.add_unused_org_entries()
        self.record("company")["data"]["media_contact_id"] = "person-observer"
        self.record("company")["refs"].extend(["person-observer", "dept-outreach"])
        company.validate_canon(self.canon)
        chart = company.org_chart_document(self.canon)
        self.assertIn("person-observer", {r["id"] for r in chart["people"]})
        self.assertIn("dept-outreach", {r["id"] for r in chart["departments"]})
        self.assertEqual(len(chart["roles"]), 8)
        self.assertFalse(any(r["data"]["person_id"] == "person-observer" for r in chart["roles"]))

    def test_org_context_filters_future_identities_and_role_assignments(self):
        self.add_unused_org_entries()
        observer = self.record("person-observer")
        observer["known_on"] = observer["effective_on"] = "2024-01-01"
        self.canon["records"].append({
            "id": "contact-outreach", "kind": "contact",
            "known_on": "2024-01-01", "effective_on": "2024-01-01",
            "data": {"person_id": "person-observer"}, "refs": ["person-observer"],
        })
        self.plan = company.create_plan(self.canon)
        doc = self.plan["documents"][0]
        for when, title in (("2023-06-30", "Finance Director"), ("2023-07-01", "CFO"),
                            ("2024-01-01", "CFO")):
            with self.subTest(when=when):
                doc["publication_date"] = when
                doc["filename"] = company.document_filename(doc["title"], when, doc["format"])
                context = company.get_context(self.canon, self.plan, doc["id"])
                chart = context["org_chart"]
                self.assertEqual(chart["active_roles"], context["active_roles"])
                active = {r["data"]["person_id"]: r for r in chart["active_roles"]}
                self.assertEqual(active["person-casey"]["data"]["title"], title)
                self.assertTrue(all(r["known_on"] <= when for key in
                                    ("people", "departments", "roles") for r in chart[key]))
                self.assertEqual("person-observer" in {r["id"] for r in chart["people"]},
                                 when >= "2024-01-01")
                self.assertEqual("role-casey-cfo" in {r["id"] for r in chart["roles"]},
                                 when >= "2023-07-01")

    def test_announced_future_role_is_not_active_until_effective(self):
        self.record("role-casey-cfo")["known_on"] = "2023-06-01"
        self.plan = company.create_plan(self.canon)
        doc = self.plan["documents"][0]
        doc["publication_date"] = "2023-06-30"
        doc["filename"] = company.document_filename(doc["title"], doc["publication_date"], doc["format"])
        chart = company.get_context(self.canon, self.plan, doc["id"])["org_chart"]
        self.assertIn("role-casey-cfo", {r["id"] for r in chart["roles"]})
        active = {r["data"]["person_id"]: r for r in chart["active_roles"]}
        self.assertEqual(active["person-casey"]["data"]["title"], "Finance Director")

    def test_claim_requires_person_in_core_org_chart(self):
        self.add_unused_org_entries()
        self.plan = company.create_plan(self.canon)
        doc = self.plan["documents"][0]
        draft = self.draft(doc)
        draft["claims"].append({"id": "c2", "record_id": "person-observer",
                                "path": ["data", "name"], "value": "Finley Brooks"})
        draft["sections"][0]["claim_ids"].append("c2")
        with self.assertRaisesRegex(company.Invalid, "missing from the bounded canonical org graph"):
            company.check_draft(self.canon, self.plan, draft)
        with tempfile.TemporaryDirectory() as directory:
            for planned in self.plan["documents"]:
                company.save(Path(directory) / f"{planned['id']}.json",
                             draft if planned["id"] == doc["id"] else self.draft(planned))
            with self.assertRaisesRegex(company.Invalid, "missing from the bounded canonical org graph"):
                company.check_drafts(self.canon, self.plan, directory)
        self.record("company")["refs"].append("person-observer")
        self.plan = company.create_plan(self.canon)
        draft["canon_sha256"] = company.fingerprint(self.canon)
        draft["plan_sha256"] = company.fingerprint(self.plan)
        company.check_draft(self.canon, self.plan, draft)

    def test_custom_person_anchors_cannot_use_out_of_core_identities(self):
        self.add_unused_org_entries()
        self.canon["request"].update({
            "documents": 1, "document_types": ["advisory_profile"],
            "additional_types": {"advisory_profile": {
                "formats": ["docx"], "anchor_kinds": ["person"],
                "audience": "internal", "required_sections": ["background"],
            }},
        })
        for person in company.org_chart_document(self.canon)["people"]:
            person["data"]["title"] = person["data"]["name"] + " Advisory Profile"
        pool = company.candidates(self.canon, "advisory_profile")
        self.assertNotIn("person-observer", {r["id"] for r in pool})
        self.plan = company.create_plan(self.canon)
        self.assertNotIn("person-observer", self.plan["documents"][0]["anchor_ids"])
        company.check_draft(self.canon, self.plan, self.draft(self.plan["documents"][0]))

    def test_exact_count_and_balanced_formats(self):
        documents = company.validate_plan(self.canon, self.plan)
        self.assertEqual(len(documents), 45)
        for fmt in company.FORMATS:
            self.assertEqual(sum(d["format"] == fmt for d in documents.values()), 15)
        self.assertEqual({d["type"] for d in documents.values()}, set(company.TYPES))
        self.assertEqual(self.plan["documents"], sorted(
            self.plan["documents"], key=lambda d: (d["publication_date"], d["id"])))
        self.assertFalse(any("fiscal years" in w for w in self.plan["warnings"]))

    def test_realistic_names_for_every_content_type(self):
        index = {r["id"]: r for r in self.canon["records"]}
        cases = [
            ("quarterly_report", "financial-2022-q1", "Q1 2022 Operating Results"),
            ("annual_report", "financial-2022-annual", "2022 Annual Business and Financial Review"),
            ("financial_model", "financial-2022-q1", "Q1 2022 Financial Performance Workbook"),
            ("product_spec", "milestone-beacon-2022-planned", "Beacon 1.0 Product Specification"),
            ("product_presentation", "milestone-beacon-2022-planned", "Beacon 1.0 - Product Overview"),
            ("hr_policy", "policy-leave-v1", "Paid Leave Policy - Version 1"),
            ("meeting_transcript", "financial-2022-q1", "Q1 2022 Operating Results - Leadership Meeting Transcript"),
            ("board_deck", "milestone-beacon-2022-planned", "Beacon 1.0 Release Plan - Board Briefing"),
            ("press_release", "acquisition-sable-announced",
             "Lumenvale Systems Announces Acquisition of Sableworks Analytics"),
            ("press_release", "acquisition-sable-closed",
             "Lumenvale Systems Completes Acquisition of Sableworks Analytics"),
            ("press_release", "milestone-beacon-2022-planned",
             "Lumenvale Systems Announces Plans for Beacon 1.0"),
            ("press_release", "milestone-beacon-2022-launched",
             "Lumenvale Systems Launches Beacon 1.0"),
        ]
        for kind, identifier, expected in cases:
            with self.subTest(kind=kind, anchor=identifier):
                title = company.document_title("Lumenvale Systems", kind, index[identifier], index)
                prefix = "" if kind == "press_release" else "Lumenvale Systems - "
                self.assertEqual(title, prefix + expected)

    def test_plan_names_hide_internal_ids_and_preserve_dates(self):
        for doc in self.plan["documents"]:
            self.assertNotIn(doc["id"], doc["title"])
            self.assertNotIn(doc["id"], doc["filename"])
            self.assertNotIn(doc["anchor_ids"][0], doc["title"])
            self.assertTrue(doc["filename"].startswith(doc["publication_date"] + " - "))
            self.assertTrue(doc["filename"].endswith("." + doc["format"]))

    def test_named_policy_uses_canonical_name_without_duplicate_label(self):
        record = self.record("policy-leave-v1")
        record["data"]["name"] = "Paid Time Off Policy"
        index = company.validate_canon(self.canon)
        self.assertEqual(
            company.document_title("Lumenvale Systems", "hr_policy", record, index),
            "Lumenvale Systems - Paid Time Off Policy - Version 1")
        record["data"]["name"] = 123
        with self.assertRaisesRegex(company.Invalid, "policy name"):
            company.validate_canon(self.canon)

    def test_filename_safety_and_length_preserve_subject(self):
        title = ('A Very Long Company ' * 30) + ' - Beacon 1.0: Product/Specification?'
        filename = company.document_filename(title, "2022-02-01", "docx", revision=2)
        self.assertLessEqual(len(filename), company.MAX_FILENAME_LENGTH)
        self.assertIn("Beacon 1.0", filename)
        self.assertTrue(filename.endswith(" - v2.docx"))
        self.assertTrue(all(c.isalnum() or c in " -_." for c in filename))
        self.assertTrue(company.document_filename("CON", "2022-02-01", "docx").startswith("2022-"))

    def test_filename_collisions_are_deterministically_disambiguated(self):
        self.canon["request"]["documents"] = 200
        self.record("product-prism")["data"]["name"] = self.record("product-beacon")["data"]["name"]
        plan = company.create_plan(self.canon)
        company.validate_plan(self.canon, plan)
        filenames = [d["filename"].casefold() for d in plan["documents"]]
        self.assertEqual(len(filenames), len(set(filenames)))
        self.assertTrue(any("Edition 2" in d["filename"] for d in plan["documents"]))
        self.assertEqual(plan, company.create_plan(self.canon))

    def test_reject_generic_or_unsafe_names(self):
        for field, value in (
            ("title", "Report"), ("title", "doc-0001"),
            ("filename", "doc-0001.docx"),
            ("filename", "2022-01-01 - Report.docx"),
            ("filename", "2022-01-01 - bad:name.docx"),
        ):
            plan = copy.deepcopy(self.plan)
            plan["documents"][0][field] = value
            with self.subTest(field=field, value=value), self.assertRaises(company.Invalid):
                company.validate_plan(self.canon, plan)

    def test_reject_duplicate_filenames_ignoring_case(self):
        plan = copy.deepcopy(self.plan)
        doc = plan["documents"][0]
        second = next(d for d in plan["documents"][1:] if d["format"] == doc["format"])
        second["publication_date"] = doc["publication_date"]
        second["anchor_ids"] = list(doc["anchor_ids"])
        second["type"] = doc["type"]
        second["required_sections"] = list(doc["required_sections"])
        second["filename"] = Path(doc["filename"]).stem.upper() + Path(doc["filename"]).suffix
        with self.assertRaisesRegex(company.Invalid, "duplicate filename"):
            company.validate_plan(self.canon, plan)

    def test_longer_history_covers_all_years_with_sufficient_budget(self):
        canon = company.seed_canon(dict(self.request, years=10, documents=60))
        plan = company.create_plan(canon)
        company.validate_plan(canon, plan)
        self.assertFalse(any("fiscal years" in w for w in plan["warnings"]))

    def test_repeated_anchors_stay_within_publication_tail(self):
        canon = company.seed_canon(dict(self.request, years=1, documents=200))
        plan = company.create_plan(canon)
        company.validate_plan(canon, plan)
        self.assertTrue(any("Repeated anchors" in w for w in plan["warnings"]))
        _, _, tail = company.period_bounds(canon["request"])
        self.assertTrue(all(company.iso(d["publication_date"]) <= tail for d in plan["documents"]))

    def test_minimal_count_warns_about_missing_types(self):
        self.canon = company.seed_canon(dict(self.request, documents=3))
        plan = company.create_plan(self.canon)
        self.assertEqual(len(plan["documents"]), 3)
        self.assertEqual({d["format"] for d in plan["documents"]}, set(company.FORMATS))
        self.assertTrue(any("selected types" in w for w in plan["warnings"]))

    def test_single_year(self):
        canon = company.seed_canon(dict(self.request, years=1))
        company.validate_plan(canon, company.create_plan(canon))
        self.assertFalse(any(r["id"] == "role-casey-cfo" for r in canon["records"]))

    def test_invalid_request(self):
        for key, value in (("years", 0), ("years", True), ("documents", 2),
                           ("seed", "42"), ("start_year", 9999), ("years", 10000)):
            with self.subTest(key=key, value=value), self.assertRaises(company.Invalid):
                company.seed_canon(dict(self.request, **{key: value}))
        with self.assertRaisesRegex(company.Invalid, "formats"):
            company.seed_canon(dict(self.request, document_types=["meeting_transcript"], formats=["pptx"]))
        with self.assertRaisesRegex(company.Invalid, "Unknown request"):
            company.seed_canon(dict(self.request, document_count=10))

    def test_default_depth_is_explicit_without_mutating_input(self):
        original = copy.deepcopy(self.request)
        canon = company.seed_canon(self.request)
        self.assertEqual(self.request, original)
        self.assertNotIn("content_depth", self.request)
        self.assertEqual(canon["request"]["content_depth"], "medium")

    def test_depth_modes_propagate_without_changing_count_or_allocation(self):
        allocation = [(d["id"], d["type"], d["format"], d["anchor_ids"])
                      for d in self.plan["documents"]]
        for depth in company.CONTENT_DEPTHS:
            with self.subTest(depth=depth):
                self.canon = company.seed_canon(dict(self.request, content_depth=depth))
                self.plan = company.create_plan(self.canon)
                self.assertEqual([(d["id"], d["type"], d["format"], d["anchor_ids"])
                                  for d in self.plan["documents"]], allocation)
                for doc in self.plan["documents"]:
                    self.assertEqual(doc["content_depth"], depth)
                    company.check_draft(self.canon, self.plan, self.draft(doc))
                context = company.get_context(self.canon, self.plan, self.plan["documents"][0]["id"])
                self.assertEqual(context["document"]["content_depth"], depth)

    def test_invalid_content_depth(self):
        for depth in ("", "DEEP", "normal", None, True, 1, [], {}):
            with self.subTest(depth=depth), self.assertRaisesRegex(company.Invalid, "content_depth"):
                company.seed_canon(dict(self.request, content_depth=depth))

    def test_depth_change_invalidates_frozen_plan(self):
        self.canon["request"]["content_depth"] = "deep"
        with self.assertRaisesRegex(company.Invalid, "Canon changed"):
            company.validate_plan(self.canon, self.plan)

    def test_missing_or_mismatched_plan_depth_requires_revision(self):
        for depth in (None, "light", "DEEP"):
            with self.subTest(depth=depth):
                plan = copy.deepcopy(self.plan)
                if depth is None:
                    plan["documents"][0].pop("content_depth")
                else:
                    plan["documents"][0]["content_depth"] = depth
                with self.assertRaisesRegex(company.Invalid, "content_depth; replan"):
                    company.validate_plan(self.canon, plan)

    def test_missing_or_mismatched_draft_depth_is_rejected(self):
        doc = self.plan["documents"][0]
        for depth in (None, "light"):
            with self.subTest(depth=depth):
                draft = self.draft(doc)
                if depth is None:
                    draft.pop("content_depth")
                else:
                    draft["content_depth"] = depth
                with self.assertRaisesRegex(company.Invalid, "wrong draft content_depth"):
                    company.check_draft(self.canon, self.plan, draft)

    def test_selected_types(self):
        canon = company.seed_canon(dict(self.request, document_types=[
            "meeting_transcript", "board_deck", "financial_model"]))
        plan = company.create_plan(canon)
        self.assertEqual({d["type"] for d in plan["documents"]}, {
            "meeting_transcript", "board_deck", "financial_model"})

    def test_single_press_release_and_product_presentation_runs(self):
        for kind, fmt in (("press_release", "docx"), ("product_presentation", "pptx")):
            with self.subTest(kind=kind):
                canon = company.seed_canon(dict(self.request, documents=1, document_types=[kind]))
                plan = company.create_plan(canon)
                company.validate_plan(canon, plan)
                self.assertEqual(len(plan["documents"]), 1)
                self.assertEqual(plan["documents"][0]["type"], kind)
                self.assertEqual(plan["documents"][0]["format"], fmt)

    def test_selected_type_mix_does_not_add_unrequested_formats(self):
        canon = company.seed_canon(dict(self.request, documents=6,
                                       document_types=["press_release", "product_presentation"]))
        plan = company.create_plan(canon)
        company.validate_plan(canon, plan)
        self.assertEqual({d["format"] for d in plan["documents"]}, {"docx", "pptx"})
        canon = company.seed_canon(dict(self.request, documents=1,
                                       document_types=["quarterly_report"], formats=["docx"]))
        self.assertEqual(company.create_plan(canon)["documents"][0]["format"], "docx")

    def test_additional_content_type_without_registry_mutation(self):
        request = dict(self.request, documents=1, document_types=["integration_plan"],
                       additional_types={"integration_plan": {
                           "formats": ["docx"], "anchor_kinds": ["integration"],
                           "audience": "internal", "required_sections": ["scope", "workstreams", "risks"],
                       }})
        canon = company.seed_canon(request)
        with self.assertRaisesRegex(company.Invalid, "No available canonical anchors"):
            company.create_plan(canon)
        canon["records"].append({
            "id": "integration-sable", "kind": "integration",
            "known_on": "2022-09-20", "effective_on": "2022-09-20",
            "refs": ["acquisition-sable-closed"],
            "data": {"title": "Sableworks Analytics Post-Acquisition Integration"},
        })
        plan = company.create_plan(canon)
        company.validate_plan(canon, plan)
        doc = plan["documents"][0]
        self.assertEqual(doc["type"], "integration_plan")
        self.assertIn("Post-Acquisition Integration", doc["title"])
        self.assertEqual(doc["required_sections"], ["scope", "workstreams", "risks"])
        self.assertNotIn("integration_plan", company.TYPES)

    def test_invalid_additional_type_profiles(self):
        profile = {"formats": ["docx"], "anchor_kinds": ["milestone"], "audience": "internal",
                   "required_sections": ["scope"]}
        for field, value in (("formats", ["txt"]), ("anchor_kinds", []),
                             ("required_sections", []), ("audience", "")):
            with self.subTest(field=field), self.assertRaises(company.Invalid):
                company.seed_canon(dict(
                    self.request, document_types=["custom_type"],
                    additional_types={"custom_type": dict(profile, **{field: value})}))
        with self.assertRaisesRegex(company.Invalid, "duplicate additional type"):
            company.seed_canon(dict(self.request, additional_types={"press_release": profile}))

    def test_stale_canon_invalidates_plan(self):
        self.record("company")["data"]["industry"] = "changed"
        with self.assertRaisesRegex(company.Invalid, "Canon changed"):
            company.validate_plan(self.canon, self.plan)

    def test_context_has_no_future_facts(self):
        doc = next(d for d in self.plan["documents"] if d["publication_date"] < "2022-08-15")
        context = company.get_context(self.canon, self.plan, doc["id"])
        self.assertTrue(all(r["known_on"] <= doc["publication_date"] for r in context["records"]))
        self.assertNotIn("financial-2024-annual", {r["id"] for r in context["records"]})
        self.assertNotIn("role-casey-cfo", {r["id"] for r in context["active_roles"]})

    def test_glossary_context_is_derived_and_date_filtered(self):
        self.canon["records"].append({
            "id": "glossary-later", "kind": "glossary",
            "known_on": "2024-01-01", "effective_on": "2024-01-01",
            "data": {"term": "Later term", "definition": "Known only in 2024."}, "refs": [],
        })
        self.plan = company.create_plan(self.canon)
        for doc in self.plan["documents"]:
            context = company.get_context(self.canon, self.plan, doc["id"])
            self.assertEqual(context["glossary"], [
                r for r in self.canon["records"] if r["kind"] == "glossary"
                and r["known_on"] <= doc["publication_date"]
            ])
            for record in context["glossary"]:
                self.assertEqual(record, self.record(record["id"]))

    def test_planned_future_event_is_visible_not_effective(self):
        doc = next(d for d in self.plan["documents"] if d["publication_date"] == "2022-02-01")
        context = company.get_context(self.canon, self.plan, doc["id"])
        planned = next(r for r in context["records"] if r["id"] == "milestone-beacon-2022-planned")
        self.assertGreater(planned["effective_on"], doc["publication_date"])
        self.assertEqual(planned["data"]["status"], "planned")

    def test_valid_claim_receipts(self):
        for doc in self.plan["documents"]:
            company.check_draft(self.canon, self.plan, self.draft(doc))

    def test_reject_fabricated_value(self):
        draft = self.draft(self.plan["documents"][0])
        draft["claims"][0]["value"] = "invented"
        with self.assertRaisesRegex(company.Invalid, "value differs"):
            company.check_draft(self.canon, self.plan, draft)

    def test_reject_future_claim(self):
        draft = self.draft(self.plan["documents"][0])
        draft["claims"][0] = {
            "id": "c1", "record_id": "financial-2024-annual", "path": ["data", "year"], "value": 2024}
        with self.assertRaisesRegex(company.Invalid, "future record"):
            company.check_draft(self.canon, self.plan, draft)

    def test_reject_unused_claim_and_unknown_section_claim(self):
        draft = self.draft(self.plan["documents"][0])
        for section in draft["sections"]:
            section["claim_ids"] = []
        with self.assertRaisesRegex(company.Invalid, "unused claims"):
            company.check_draft(self.canon, self.plan, draft)
        draft["sections"][0]["claim_ids"] = ["unknown"]
        with self.assertRaisesRegex(company.Invalid, "unknown section claim"):
            company.check_draft(self.canon, self.plan, draft)

    def test_reject_wrong_draft_metadata(self):
        for field, value in (("title", "wrong"), ("publication_date", "2000-01-01")):
            draft = self.draft(self.plan["documents"][0])
            draft[field] = value
            with self.subTest(field=field), self.assertRaises(company.Invalid):
                company.check_draft(self.canon, self.plan, draft)

    def test_reject_missing_genre_sections(self):
        doc = self.plan["documents"][0]
        draft = self.draft(doc)
        draft["sections"].pop()
        with self.assertRaisesRegex(company.Invalid, "missing genre sections"):
            company.check_draft(self.canon, self.plan, draft)
        draft = self.draft(doc)
        draft["sections"][0].pop("role")
        with self.assertRaisesRegex(company.Invalid, "genre role"):
            company.check_draft(self.canon, self.plan, draft)

    def test_reject_self_labeling_but_allow_grounded_technical_terms(self):
        phrases = [
            "This document is synthetic.", "This company is fictional.",
            "This is a fake company.", "AI-generated press release.",
            "Written by AI.", "For demonstration purposes only.", "SYNTHETIC",
        ]
        for phrase in phrases:
            draft = self.draft(self.plan["documents"][0])
            draft["sections"][0]["text"] = phrase
            with self.subTest(phrase=phrase), self.assertRaisesRegex(company.Invalid, "self-identify"):
                company.check_draft(self.canon, self.plan, draft)
        company.validate_content_voice("Synthetic aperture radar improves resolution.", "technical term")
        self.assertNotIn("synthetic", self.draft(self.plan["documents"][0]))
        self.assertNotIn("fictional", self.record("company")["data"])
        for record in self.canon["records"]:
            if record["kind"] == "policy":
                self.assertEqual(record["data"]["scope"], "All full-time employees")

    def test_reject_disclosure_in_exported_office_text_and_properties(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "content.docx"
            for part, xml in (
                ("word/document.xml", "<document><p><r>This document is </r><r>synthetic.</r></p></document>"),
                ("docProps/custom.xml", '<Properties><property name="synthetic">true</property></Properties>'),
            ):
                with zipfile.ZipFile(path, "w") as package:
                    for required in ("[Content_Types].xml", "_rels/.rels", "word/document.xml"):
                        package.writestr(required, xml if required == part else "<fixture/>")
                    if part != "word/document.xml":
                        package.writestr(part, xml)
                with self.subTest(part=part), self.assertRaises(company.Invalid):
                    company.audit_package(path, "docx")

    def test_plan_change_invalidates_drafts(self):
        draft = self.draft(self.plan["documents"][0])
        self.plan["documents"][0]["revision"] += 1
        with self.assertRaisesRegex(company.Invalid, "plan_sha256"):
            company.check_draft(self.canon, self.plan, draft)

    def test_claim_path_and_type_safety(self):
        self.assertEqual(company.resolve({"data": {"items": [7]}}, ["data", "items", 0]), 7)
        for path in ([], ["data", "missing"], ["data", "items", -1], ["data", "items", True]):
            with self.subTest(path=path), self.assertRaises(company.Invalid):
                company.resolve({"data": {"items": [7]}}, path)
        draft = self.draft(self.plan["documents"][0])
        if type(draft["claims"][0]["value"]) is int:
            draft["claims"][0]["value"] = float(draft["claims"][0]["value"])
            with self.assertRaisesRegex(company.Invalid, "value differs"):
                company.check_draft(self.canon, self.plan, draft)

    def test_future_anchor_is_rejected(self):
        self.plan["documents"][0]["anchor_ids"] = ["financial-2024-q4"]
        with self.assertRaisesRegex(company.Invalid, "future anchor"):
            company.validate_plan(self.canon, self.plan)

    def test_structural_audit_and_receipt_invalidation(self):
        self.canon = company.seed_canon(dict(self.request, documents=3))
        self.plan = company.create_plan(self.canon)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            drafts = root / "build" / "drafts"
            manifest, qa = self.create_audit_fixture(root, drafts)
            self.assertEqual(company.audit(self.canon, self.plan, drafts, manifest, root)["documents"], 3)
            qa["documents"][0]["artifact_sha256"] = "stale"
            (root / "build" / "qa.json").write_text(json.dumps(qa), encoding="utf-8")
            with self.assertRaisesRegex(company.Invalid, "stale artifact QA"):
                company.audit(self.canon, self.plan, drafts, manifest, root)

    def test_audit_rejects_modified_draft(self):
        self.canon = company.seed_canon(dict(self.request, documents=3))
        self.plan = company.create_plan(self.canon)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            drafts = root / "build" / "drafts"
            manifest, _ = self.create_audit_fixture(root, drafts)
            path = drafts / f"{self.plan['documents'][0]['id']}.json"
            draft = company.load(path)
            draft["sections"][0]["text"] = "Changed test fixture."
            path.write_text(json.dumps(draft), encoding="utf-8")
            with self.assertRaisesRegex(company.Invalid, "stale draft QA"):
                company.audit(self.canon, self.plan, drafts, manifest, root)

    def test_audit_rejects_modified_artifact(self):
        self.canon = company.seed_canon(dict(self.request, documents=3))
        self.plan = company.create_plan(self.canon)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest, _ = self.create_audit_fixture(root, root / "build" / "drafts")
            (root / manifest["documents"][0]["path"]).write_bytes(b"changed")
            with self.assertRaisesRegex(company.Invalid, "artifact hash mismatch"):
                company.audit(self.canon, self.plan, root / "build" / "drafts", manifest, root)

    def create_audit_fixture(self, root, drafts, include_titles=False):
        company.save(root / "canon.json", self.canon)
        (root / "documents").mkdir()
        manifest = {"schema_version": 2, "corpus_id": self.canon["corpus_id"],
                    "canon_sha256": company.fingerprint(self.canon),
                    "plan_sha256": company.fingerprint(self.plan), "documents": []}
        qa = {"schema_version": 2, "corpus_id": self.canon["corpus_id"],
              "canon_sha256": company.fingerprint(self.canon),
              "plan_sha256": company.fingerprint(self.plan), "documents": []}
        for doc in self.plan["documents"]:
            draft_path = drafts / f"{doc['id']}.json"
            company.save(draft_path, self.draft(doc))
            fmt = doc["format"]
            path = root / "documents" / doc["filename"]
            # These content/relationship fixtures exercise read-only checks, not Office rendering.
            draft = self.draft(doc)
            bodies = [escape(section["text"]) for section in draft["sections"]]
            if include_titles:
                bodies[0] += escape(" " + doc["title"])
            w, p, a, x, r = (company.NS[key] for key in ("w", "p", "a", "x", "r"))
            rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
            parts = {"[Content_Types].xml": "<Types/>", "_rels/.rels": "<Relationships/>"}
            if include_titles:
                parts["docProps/core.xml"] = (
                    '<core xmlns:dc="http://purl.org/dc/elements/1.1/">'
                    f"<dc:title>{escape(doc['title'])}</dc:title></core>")
            if fmt == "docx":
                parts["word/document.xml"] = (
                    f'<w:document xmlns:w="{w}"><w:body>'
                    + "".join(f"<w:p><w:r><w:t>{body}</w:t></w:r></w:p>" for body in bodies)
                    + "</w:body></w:document>")
            elif fmt == "pptx":
                parts["ppt/presentation.xml"] = (
                    f'<p:presentation xmlns:p="{p}" xmlns:r="{r}">'
                    '<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>')
                parts["ppt/_rels/presentation.xml.rels"] = (
                    f'<Relationships xmlns="{rel_ns}"><Relationship Id="rId1" '
                    f'Type="{r}/slide" Target="slides/slide1.xml"/></Relationships>')
                parts["ppt/slides/slide1.xml"] = (
                    f'<p:sld xmlns:p="{p}" xmlns:a="{a}"><p:cSld><p:spTree>'
                    + "".join(f'<p:sp><p:nvSpPr><p:cNvPr id="{i + 1}"/></p:nvSpPr>'
                              f"<p:txBody><a:p><a:r><a:t>{body}</a:t></a:r></a:p></p:txBody></p:sp>"
                              for i, body in enumerate(bodies))
                    + "</p:spTree></p:cSld></p:sld>")
            else:
                parts["xl/workbook.xml"] = (
                    f'<workbook xmlns="{x}" xmlns:r="{r}"><sheets>'
                    '<sheet name="Results" sheetId="1" r:id="rId1"/></sheets></workbook>')
                parts["xl/_rels/workbook.xml.rels"] = (
                    f'<Relationships xmlns="{rel_ns}"><Relationship Id="rId1" '
                    f'Type="{r}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
                parts["xl/worksheets/sheet1.xml"] = (
                    f'<worksheet xmlns="{x}"><sheetData>'
                    + "".join(f'<row r="{i + 1}"><c r="A{i + 1}" t="inlineStr"><is>'
                              f"<t>{body}</t></is></c></row>" for i, body in enumerate(bodies))
                    + "</sheetData></worksheet>")
            with zipfile.ZipFile(path, "w") as package:
                for part, content in parts.items():
                    package.writestr(part, content)
            digest = company.file_hash(path)
            manifest["documents"].append({
                "document_id": doc["id"], "path": str(path.relative_to(root)), "sha256": digest})
            review = {
                "document_id": doc["id"], "artifact_sha256": digest,
                "draft_sha256": company.file_hash(draft_path),
                "semantic_review": {"status": "passed", "evidence": "Test fixture only"},
                "genre_review": {"status": "passed", "evidence": "Test fixture only"},
                "depth_review": {"status": "passed", "content_depth": doc["content_depth"],
                                 "evidence": "Test fixture only"},
                "visual_review": {"status": "passed", "evidence": "Test fixture only"},
            }
            if fmt == "xlsx":
                review["recalculation"] = {"status": "passed", "evidence": "Test fixture only"}
            qa["documents"].append(review)
        (root / "build" / "qa.json").write_text(json.dumps(qa), encoding="utf-8")
        indexed = company.index_artifacts(self.canon, self.plan, manifest, root)
        locations = dict(company.receipt_header(self.canon, self.plan), documents=[])
        for doc in indexed["documents"]:
            draft_path = drafts / f"{doc['document_id']}.json"
            draft = company.load(draft_path)
            value = company.display_value(draft["claims"][0])
            locations["documents"].append({
                "document_id": doc["document_id"], "revision": doc["revision"],
                "artifact_sha256": doc["artifact_sha256"],
                "draft_sha256": company.file_hash(draft_path),
                "occurrences": [{"claim_id": "c1", "section_index": i,
                                 "locator": entry["locator"], "start": 0, "end": len(value),
                                 "expected_text": value}
                                for i, entry in enumerate(doc["containers"])]})
        company.save(root / "build" / "claim-locations.json", locations)
        provenance = company.verify_provenance(self.canon, self.plan, drafts, manifest, root, locations)
        company.save(root / "build" / "provenance.json", provenance)
        company.save(root / "build" / "coverage.json",
                     company.coverage_report(self.canon, self.plan, drafts, provenance))
        return manifest, qa

    def test_audit_rejects_path_escape(self):
        self.canon = company.seed_canon(dict(self.request, documents=3))
        self.plan = company.create_plan(self.canon)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest, _ = self.create_audit_fixture(root, root / "build" / "drafts")
            manifest["documents"][0]["path"] = "../escaped.docx"
            with self.assertRaisesRegex(company.Invalid, "escaping artifact path"):
                company.audit(self.canon, self.plan, root / "build" / "drafts", manifest, root)

    def test_audit_rejects_generic_artifact_filename(self):
        self.canon = company.seed_canon(dict(self.request, documents=3))
        self.plan = company.create_plan(self.canon)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest, _ = self.create_audit_fixture(root, root / "build" / "drafts")
            entry = manifest["documents"][0]
            original = root / entry["path"]
            entry["path"] = str(Path("documents") / f"{entry['document_id']}{original.suffix}")
            original.rename(root / entry["path"])
            with self.assertRaisesRegex(company.Invalid, "descriptive planned filename"):
                company.audit(self.canon, self.plan, root / "build" / "drafts", manifest, root)
    def test_audit_rejects_missing_recalculation(self):
        self.canon = company.seed_canon(dict(self.request, documents=3))
        self.plan = company.create_plan(self.canon)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest, qa = self.create_audit_fixture(root, root / "build" / "drafts")
            for review in qa["documents"]:
                review.pop("recalculation", None)
            (root / "build" / "qa.json").write_text(json.dumps(qa), encoding="utf-8")
            with self.assertRaisesRegex(company.Invalid, "recalculation"):
                company.audit(self.canon, self.plan, root / "build" / "drafts", manifest, root)

    def test_audit_rejects_missing_genre_review(self):
        self.canon = company.seed_canon(dict(self.request, documents=3))
        self.plan = company.create_plan(self.canon)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            drafts = root / "build" / "drafts"
            manifest, qa = self.create_audit_fixture(root, drafts)
            qa["documents"][0].pop("genre_review")
            (root / "build" / "qa.json").write_text(json.dumps(qa), encoding="utf-8")
            with self.assertRaisesRegex(company.Invalid, "genre_review"):
                company.audit(self.canon, self.plan, drafts, manifest, root)

    def test_depth_review_is_required_passed_and_bound_to_plan(self):
        self.canon = company.seed_canon(dict(self.request, documents=3))
        self.plan = company.create_plan(self.canon)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            drafts = root / "build" / "drafts"
            manifest, original = self.create_audit_fixture(root, drafts)
            for mutation in ("missing", "failed", "empty_evidence", "wrong_depth", "missing_depth"):
                with self.subTest(mutation=mutation):
                    qa = copy.deepcopy(original)
                    review = qa["documents"][0]
                    if mutation == "missing":
                        review.pop("depth_review")
                    elif mutation == "failed":
                        review["depth_review"]["status"] = "failed"
                    elif mutation == "empty_evidence":
                        review["depth_review"]["evidence"] = " "
                    elif mutation == "wrong_depth":
                        review["depth_review"]["content_depth"] = "light"
                    else:
                        review["depth_review"].pop("content_depth")
                    (root / "build" / "qa.json").write_text(json.dumps(qa), encoding="utf-8")
                    with self.assertRaisesRegex(company.Invalid, "depth_review"):
                        company.audit(self.canon, self.plan, drafts, manifest, root)

    def test_audit_passes_for_each_depth_without_sidecar_files(self):
        for depth in company.CONTENT_DEPTHS:
            with self.subTest(depth=depth), tempfile.TemporaryDirectory() as directory:
                self.canon = company.seed_canon(dict(self.request, documents=3, content_depth=depth))
                self.plan = company.create_plan(self.canon)
                root = Path(directory)
                drafts = root / "build" / "drafts"
                manifest, _ = self.create_audit_fixture(root, drafts)
                self.assertEqual(company.audit(self.canon, self.plan, drafts, manifest, root)["documents"], 3)
                self.assertEqual({p.name for p in root.iterdir()}, {"canon.json", "documents", "build"})

    def test_layout_rejects_root_pollution_and_canon_drift(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "output"
            company.initialize_output(self.canon, root)
            company.validate_output_layout(self.canon, root)
            stray = root / "renderer.py"
            stray.write_text("print('stray')", encoding="utf-8")
            with self.assertRaisesRegex(company.Invalid, "Output root"):
                company.validate_output_layout(self.canon, root)
            stray.unlink()
            canon = company.load(root / "canon.json")
            next(r for r in canon["records"] if r["kind"] == "glossary")["data"]["definition"] = "Changed definition"
            (root / "canon.json").write_text(json.dumps(canon), encoding="utf-8")
            with self.assertRaisesRegex(company.Invalid, "canon.json differs"):
                company.validate_output_layout(self.canon, root)

    def test_derived_paths_must_be_in_build(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            company.require_build_path(root / "build" / "drafts" / "doc-0001.json", root)
            for path in (root / "plan.json", root / "documents" / "renderer.py",
                         root / "build" / ".." / "context.json"):
                with self.subTest(path=path), self.assertRaisesRegex(company.Invalid, "build directory"):
                    company.require_build_path(path, root)

    def test_single_canon_layout_rejects_legacy_sidecars_without_deleting_them(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "output"
            company.initialize_output(self.canon, root)
            for filename in ("org-chart.json", "glossary.json"):
                with self.subTest(filename=filename):
                    path = root / filename
                    path.write_text('{"legacy": true}', encoding="utf-8")
                    with self.assertRaisesRegex(company.Invalid, "legacy layouts require a new run/revision"):
                        company.validate_output_layout(self.canon, root)
                    self.assertEqual(path.read_text(encoding="utf-8"), '{"legacy": true}')
                    path.unlink()
            company.validate_output_layout(self.canon, root)

    def test_audit_rejects_canonical_role_drift(self):
        self.canon = company.seed_canon(dict(self.request, documents=3))
        self.plan = company.create_plan(self.canon)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            drafts = root / "build" / "drafts"
            manifest, _ = self.create_audit_fixture(root, drafts)
            canon = company.load(root / "canon.json")
            next(r for r in canon["records"] if r["kind"] == "role")["data"]["title"] = "Changed title"
            (root / "canon.json").write_text(json.dumps(canon), encoding="utf-8")
            with self.assertRaisesRegex(company.Invalid, "canon.json differs"):
                company.audit(self.canon, self.plan, drafts, manifest, root)

    def test_audit_rejects_unmanifested_documents_and_artifacts_in_build(self):
        self.canon = company.seed_canon(dict(self.request, documents=3))
        self.plan = company.create_plan(self.canon)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            drafts = root / "build" / "drafts"
            manifest, _ = self.create_audit_fixture(root, drafts)
            stray = root / "documents" / "render.png"
            stray.write_bytes(b"render")
            with self.assertRaisesRegex(company.Invalid, "manifested final artifacts"):
                company.audit(self.canon, self.plan, drafts, manifest, root)
            stray.unlink()
            entry = manifest["documents"][0]
            original = root / entry["path"]
            moved = root / "build" / original.name
            original.rename(moved)
            entry["path"] = str(moved.relative_to(root))
            with self.assertRaisesRegex(company.Invalid, "final artifacts must be in documents"):
                company.audit(self.canon, self.plan, drafts, manifest, root)

    def test_package_rejects_fake_and_malformed_files(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "fake.docx"
            path.write_text("Not a Word file", encoding="utf-8")
            with self.assertRaisesRegex(company.Invalid, "invalid Office package"):
                company.audit_package(path, "docx")
            with zipfile.ZipFile(path, "w") as package:
                for part in ("[Content_Types].xml", "_rels/.rels", "word/document.xml"):
                    package.writestr(part, "<broken>")
            with self.assertRaisesRegex(company.Invalid, "invalid Office package"):
                company.audit_package(path, "docx")

    def test_save_never_overwrites(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "data.json"
            company.save(path, {"original": True})
            with self.assertRaisesRegex(company.Invalid, "Refusing to overwrite"):
                company.save(path, {"original": False})
            self.assertEqual(company.load(path), {"original": True})

    def test_cli_init_validate_plan_context_and_error_exit(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = root / "request.json"
            company.save(config, self.request)
            out = root / "corpus"

            def run(*args):
                return subprocess.run(
                    [sys.executable, str(SCRIPT), *map(str, args)],
                    capture_output=True, text=True, check=False)

            for args in (
                ("init", "--config", config, "--out", out),
                ("validate", "--canon", out / "canon.json"),
                ("plan", "--canon", out / "canon.json", "--out", out / "build" / "plan.json"),
                ("context", "--canon", out / "canon.json", "--plan", out / "build" / "plan.json",
                 "--doc", "doc-0001", "--out", out / "build" / "contexts" / "doc-0001.json"),
            ):
                result = run(*args)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertIsInstance(json.loads(result.stdout), dict)
            self.assertTrue((out / "build" / "contexts" / "doc-0001.json").is_file())
            self.assertEqual({p.name for p in out.iterdir()},
                             {"canon.json", "documents", "build"})
            context = company.load(out / "build" / "contexts" / "doc-0001.json")
            self.assertIn("org_chart", context)
            self.assertIn("glossary", context)
            self.assertEqual(context["document"]["content_depth"], "medium")
            self.assertEqual(company.load(out / "build" / "request.json")["content_depth"], "medium")
            result = run("plan", "--canon", out / "canon.json", "--out", out / "bad-plan.json")
            self.assertEqual(result.returncode, 1)
            self.assertIn("build directory", result.stderr)
            self.assertFalse((out / "bad-plan.json").exists())
            result = run("init", "--config", config, "--out", out)
            self.assertEqual(result.returncode, 1)
            self.assertIn("already exists", result.stderr)
            config.write_text("{}", encoding="utf-8")
            result = run("init", "--config", config, "--out", root / "invalid")
            self.assertEqual(result.returncode, 1)
            self.assertIn("error:", result.stderr)

    def test_cli_explicit_depth_modes_and_invalid_setting(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for depth in company.CONTENT_DEPTHS:
                with self.subTest(depth=depth):
                    config = root / f"{depth}.json"
                    company.save(config, dict(self.request, documents=1,
                                              document_types=["press_release"], content_depth=depth))
                    out = root / depth
                    commands = [
                        ("init", "--config", config, "--out", out),
                        ("plan", "--canon", out / "canon.json", "--out", out / "build" / "plan.json"),
                        ("context", "--canon", out / "canon.json", "--plan", out / "build" / "plan.json",
                         "--doc", "doc-0001", "--out", out / "build" / "contexts" / "doc-0001.json"),
                    ]
                    for args in commands:
                        result = subprocess.run(
                            [sys.executable, str(SCRIPT), *map(str, args)],
                            capture_output=True, text=True, check=False)
                        self.assertEqual(result.returncode, 0, result.stderr)
                    self.assertEqual(company.load(out / "build" / "request.json")["content_depth"], depth)
                    self.assertEqual(company.load(out / "build" / "contexts" / "doc-0001.json")
                                     ["document"]["content_depth"], depth)
            invalid = root / "invalid.json"
            company.save(invalid, dict(self.request, content_depth="maximum"))
            out = root / "invalid-output"
            result = subprocess.run(
                [sys.executable, str(SCRIPT), "init", "--config", str(invalid), "--out", str(out)],
                capture_output=True, text=True, check=False)
            self.assertEqual(result.returncode, 1)
            self.assertIn("content_depth", result.stderr)
            self.assertFalse(out.exists())

    def test_skill_identity_and_intake_instructions(self):
        skill = SCRIPT.parents[1] / "SKILL.md"
        instructions = skill.read_text(encoding="utf-8")
        self.assertIn("name: company-content-generator", instructions)
        self.assertIn("Company name and industry", instructions)
        self.assertIn("Content volume", instructions)
        self.assertIn("Target content types", instructions)
        self.assertIn("Output location", instructions)
        self.assertIn("additional free-text types", instructions)
        self.assertIn("generation is blocked", instructions)
        self.assertIn("only root-level file", instructions)
        self.assertIn("Content depth defaults to **medium**", instructions)
        self.assertIn("`depth_review`", instructions)
        self.assertIn("core info set", instructions)
        self.assertIn("named-person count need not equal headcount", instructions)
        self.assertFalse((ROOT / ".github" / "skills" / "synthetic-company" / "SKILL.md").exists())

    def test_skill_requires_officecli_not_native_office(self):
        instructions = (SCRIPT.parents[1] / "SKILL.md").read_text(encoding="utf-8")
        instructions = " ".join(instructions.split())
        self.assertIn("Microsoft Office is not required", instructions)
        self.assertIn("Before creating the output directory, canon, plan, or drafts", instructions)
        self.assertIn("https://github.com/iOfficeAI/OfficeCLI", instructions)
        self.assertIn("verify versioned release checksums", instructions)
        self.assertIn("Maintaining this skill and running helper tests", instructions)
        for excluded in ("LibreOffice", "OpenOffice", "standalone OOXML", "COM"):
            self.assertIn(excluded, instructions)
        self.assertNotIn("Canon/planning work can proceed without Office", instructions)
        self.assertNotIn("`docx`, `pptx`, and `xlsx` by default", instructions)

    def test_setup_guide_covers_authorized_install_updates_and_probes(self):
        guide = (SCRIPT.parents[1] / "references" / "dependencies.md").read_text(encoding="utf-8")
        guide = " ".join(guide.split())
        for requirement in ("sole document", "Installation", "versioned release asset", "SHA-256",
                            "before execution", "authorization", "Preserve the prior usable binary",
                            "MCP registration is optional", "--render html", "OFFICECLI_SKIP_UPDATE=1"):
            self.assertIn(requirement.lower(), guide.lower())
        for old_command in ("gh skill install iOfficeAI/OfficeCLI", "git clone --filter=blob:none",
                            "soffice --version", "Get-Command soffice", "pdftoppm -v",
                            "import openpyxl", "require.resolve('pptxgenjs')"):
            self.assertNotIn(old_command, guide)
        readme = (ROOT / "README.md").read_text(encoding="utf-8")
        self.assertIn("No tools are installed automatically", readme)

    def test_skill_bounds_execution_and_independent_review_scope(self):
        instructions = (SCRIPT.parents[1] / "SKILL.md").read_text(encoding="utf-8")
        for requirement in ("Balanced", "execution-plan.json", "review-plan.json",
                            "corrective retry", "Undelivered/unreadable", "not_selected",
                            "OfficeCLI runner", "independent of light/medium/deep"):
            self.assertIn(requirement, instructions)

    def test_no_conflicting_native_requirements_remain_in_maintained_guidance(self):
        references = SCRIPT.parents[1] / "references"
        contracts = (references / "contracts.md").read_text(encoding="utf-8")
        self.assertIn("Balanced", contracts)
        for path in [SCRIPT.parents[1] / "SKILL.md", ROOT / "README.md", *references.glob("*.md")]:
            wording = " ".join(path.read_text(encoding="utf-8").split())
            for forbidden in ("Microsoft Office is a hard prerequisite", "recalculate in Excel",
                              "Recalculate workbooks in Microsoft Excel", "Word.Application",
                              "PowerPoint.Application", "Excel.Application", "mandatory native"):
                self.assertNotIn(forbidden, wording, str(path))


if __name__ == "__main__":
    unittest.main()
