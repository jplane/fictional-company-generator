import copy
import unittest

import test_company as baseline

company = baseline.company


class ContentModelTests(unittest.TestCase):
    def setUp(self):
        self.fixture = baseline.CompanyTests()
        self.fixture.request = {
            "company_name": "Lumenvale Systems", "industry": "analytics software",
            "start_year": 2022, "years": 3, "documents": 2, "seed": 42,
            "document_types": ["meeting_transcript"], "formats": ["docx"],
        }
        self.canon = company.seed_canon(self.fixture.request)
        self.fixture.canon = self.canon

    def add(self, identifier, kind, data, refs=(), when="2022-02-01", **extra):
        record = {"id": identifier, "kind": kind, "known_on": when, "effective_on": when,
                  "data": data, "refs": list(refs), **extra}
        self.canon["records"].append(record)
        return record

    def source(self, identifier, when="2022-02-01", supersedes=None, **extra):
        data = {"source_id": "series-" + identifier, "version": 1,
                "title": "Product Readiness Discussion", "author_id": "person-avery",
                "audience": "internal", "authority_scope": "Product release planning",
                "status": "draft", "document_type": "meeting_transcript", "published_on": when,
                **extra}
        refs = ["person-avery", "product-beacon"]
        if supersedes:
            previous = self.record(supersedes)
            data.update(source_id=previous["data"]["source_id"],
                        version=previous["data"]["version"] + 1, supersedes=supersedes)
            refs.append(supersedes)
        return self.add(identifier, "source_version", data, refs, when)

    def statement(self, identifier, source, value, kind="fact", when="2022-02-01", **extra):
        data = {"subject_id": "product-beacon", "predicate": "release_date", "value": value,
                "kind": kind, "source_version_id": source, "speaker_id": "person-avery", **extra}
        refs = ["product-beacon", source, "person-avery"]
        refs.extend(extra.get("supports", []))
        refs.extend(extra[key] for key in ("corrects", "retracts") if key in extra)
        return self.add(identifier, "statement", data, refs, when)

    def condition(self, kind, statements=(), sources=(), blocked=(), **extra):
        request_id = "request-" + kind
        self.canon["request"].setdefault("corpus_conditions", []).append(
            {"id": request_id, "kind": kind, "count": 1})
        data = {"request_id": request_id, "kind": kind, "subject_id": "product-beacon",
                "statement_ids": list(statements), "source_version_ids": list(sources),
                "blocked_record_ids": list(blocked), **extra}
        return self.add("condition-" + kind, "condition", data,
                        dict.fromkeys(["product-beacon", *statements, *sources, *blocked]),
                        when="2022-03-01")

    def record(self, identifier):
        return next(item for item in self.canon["records"] if item["id"] == identifier)

    def plan(self):
        self.fixture.plan = company.create_plan(self.canon)
        return self.fixture.plan

    def test_new_version_rejects_old_canon_and_draft(self):
        old = copy.deepcopy(self.canon)
        old["schema_version"] = 1
        with self.assertRaisesRegex(company.Invalid, "create a new run"):
            company.validate_canon(old)
        plan = self.plan()
        draft = self.fixture.draft(plan["documents"][0])
        draft["schema_version"] = 1
        with self.assertRaisesRegex(company.Invalid, "schema_version 2"):
            company.check_draft(self.canon, plan, draft)

    def test_namespace_and_depth_are_independent_of_output_paths(self):
        self.assertEqual(self.canon["corpus_id"], company.seed_canon(self.fixture.request)["corpus_id"])
        custom = company.seed_canon(dict(self.fixture.request, corpus_id="my-corpus"))
        self.assertEqual(custom["corpus_id"], "my-corpus")
        with self.assertRaisesRegex(company.Invalid, "namespace"):
            company.seed_canon(dict(self.fixture.request, corpus_id="C:\\private"))

    def test_conditions_off_by_default_and_not_implied_by_depth(self):
        for depth in company.CONTENT_DEPTHS:
            canon = company.seed_canon(dict(self.fixture.request, content_depth=depth))
            plan = company.create_plan(canon)
            self.assertFalse(any(doc["condition_ids"] for doc in plan["documents"]))
            self.assertEqual(len(plan["documents"]), 2)

    def test_missing_condition_instances_fail_before_planning(self):
        self.canon["request"]["corpus_conditions"] = [{"id": "g", "kind": "unknown", "count": 1}]
        with self.assertRaisesRegex(company.Invalid, "enrich the canon"):
            self.plan()

    def test_unrequested_condition_and_unknown_fields_are_rejected(self):
        condition = self.condition("unknown", blocked=["product-prism"])
        self.canon["request"]["corpus_conditions"] = []
        with self.assertRaisesRegex(company.Invalid, "not explicitly requested"):
            self.plan()
        self.canon["request"]["corpus_conditions"] = [{"id": "request-unknown", "kind": "unknown", "count": 1}]
        condition["data"]["expected_agent_behavior"] = "abstain"
        with self.assertRaisesRegex(company.Invalid, "unknown condition fields"):
            self.plan()

    def test_controlled_conflict_preserves_canonical_financials(self):
        self.source("source-a")
        self.source("source-b")
        self.statement("statement-a", "source-a", "2022-08-15")
        self.statement("statement-b", "source-b", "2022-09-15")
        self.condition("conflict", statements=["statement-a", "statement-b"])
        original = self.record("financial-2022-q1")["data"]["revenue_usd"]
        for depth in company.CONTENT_DEPTHS:
            self.canon["request"]["content_depth"] = depth
            plan = self.plan()
            self.assertEqual(len(plan["documents"]), 2)
            self.assertEqual({doc["source_version_id"] for doc in plan["documents"]},
                             {"source-a", "source-b"})
            self.assertTrue(all(doc["condition_ids"] == ["condition-conflict"] for doc in plan["documents"]))
            self.assertEqual(self.record("financial-2022-q1")["data"]["revenue_usd"], original)
            context = company.get_context(self.canon, plan, plan["documents"][0]["id"])
            self.assertNotIn("condition_ids", context["document"])
            self.assertNotIn("condition", {record["kind"] for record in context["records"]})

    def test_accidental_conflicts_and_wrong_scope_fail(self):
        self.source("source-a")
        self.source("source-b")
        self.statement("statement-a", "source-a", "2022-08-15")
        self.statement("statement-b", "source-b", "2022-09-15")
        with self.assertRaisesRegex(company.Invalid, "unregistered conflicting"):
            self.plan()
        self.condition("conflict", statements=["statement-a", "statement-b"])
        self.record("statement-b")["data"]["period"] = "2023"
        with self.assertRaisesRegex(company.Invalid, "same scope"):
            self.plan()

    def test_required_sources_do_not_increase_budget(self):
        self.source("source-a")
        self.source("source-b")
        self.statement("statement-a", "source-a", "August")
        self.statement("statement-b", "source-b", "September")
        self.condition("conflict", statements=["statement-a", "statement-b"])
        self.canon["request"]["documents"] = 1
        with self.assertRaisesRegex(company.Invalid, "fixed document budget"):
            self.plan()

    def test_gap_and_conflict_overlap_is_rejected(self):
        self.source("source-a")
        self.source("source-b")
        self.statement("statement-a", "source-a", "August")
        self.statement("statement-b", "source-b", "September")
        self.condition("conflict", statements=["statement-a", "statement-b"])
        self.condition("unknown", sources=["source-a"], blocked=["statement-a"])
        with self.assertRaisesRegex(company.Invalid, "overlap"):
            self.plan()

    def test_unknown_and_undocumented_have_different_writer_knowledge(self):
        for kind in ("unknown", "undocumented"):
            with self.subTest(kind=kind):
                self.setUp()
                self.source("source-a")
                self.statement("statement-a", "source-a", "August")
                self.condition(kind, sources=["source-a"], blocked=["statement-a"])
                plan = self.plan()
                doc = next(doc for doc in plan["documents"] if doc["source_version_id"] == "source-a")
                context = company.get_context(self.canon, plan, doc["id"])
                ids = {record["id"] for record in context["records"]}
                self.assertEqual("statement-a" in ids, kind == "undocumented")
                draft = self.fixture.draft(doc)
                draft["claims"].append({"id": "extra", "record_id": "statement-a",
                                        "path": ["data", "value"], "value": "August"})
                draft["sections"][0]["claim_ids"].append("extra")
                with self.assertRaisesRegex(company.Invalid, "unknown or future"):
                    company.check_draft(self.canon, plan, draft)

    def test_omitted_source_is_not_a_missing_export(self):
        self.source("source-a")
        self.condition("omitted_source", sources=["source-a"])
        plan = self.plan()
        self.assertEqual(len(plan["documents"]), 2)
        self.assertNotIn("source-a", {doc["source_version_id"] for doc in plan["documents"]})
        self.assertTrue(any("condition-omitted_source" in doc["condition_ids"] for doc in plan["documents"]))

    def test_omission_carriers_preserve_period_and_viewpoint_selectors(self):
        self.source("source-a")
        condition = self.condition("omitted_source", sources=["source-a"], period="2099")
        with self.assertRaisesRegex(company.Invalid, "no compatible carrier"):
            self.plan()
        condition["data"].pop("period")
        self.add("view-unselected", "viewpoint",
                 {"name": "Engineering", "author_id": "person-riley", "audience": "internal"},
                 ["person-riley"])
        condition["data"]["viewpoint_id"] = "view-unselected"
        condition["refs"].append("view-unselected")
        with self.assertRaisesRegex(company.Invalid, "no compatible carrier"):
            self.plan()

    def test_period_selector_survives_suppression_of_source_assertions(self):
        self.source("source-a")
        self.statement("statement-a", "source-a", "August", period="2022-Q1")
        self.condition("undocumented", sources=["source-a"], blocked=["statement-a"], period="2022-Q1")
        plan = self.plan()
        doc = next(doc for doc in plan["documents"] if doc["source_version_id"] == "source-a")
        self.assertEqual(doc["required_statement_ids"], [])
        self.assertIn("condition-undocumented", doc["condition_ids"])

    def test_noise_and_staleness_require_real_sources(self):
        self.source("source-a")
        self.source("source-b", when="2022-03-01", supersedes="source-a", status="approved")
        self.condition("stale", sources=["source-a"], as_of_date="2022-03-01")
        self.condition("noise", sources=["source-b"])
        plan = self.plan()
        self.assertEqual({doc["source_version_id"] for doc in plan["documents"]}, {"source-a", "source-b"})
        old = next(doc for doc in plan["documents"] if doc["source_version_id"] == "source-a")
        context = company.get_context(self.canon, plan, old["id"])
        self.assertNotIn("source-b", {record["id"] for record in context["records"]})
        self.record("condition-stale")["data"]["as_of_date"] = "2022-02-01"
        with self.assertRaisesRegex(company.Invalid, "published superseding"):
            self.plan()

    def test_corrections_are_backward_and_do_not_rewrite_old_sources(self):
        self.source("source-a")
        self.statement("statement-a", "source-a", "August")
        self.source("source-b", when="2022-03-01", supersedes="source-a")
        self.statement("statement-b", "source-b", "September", when="2022-03-01", corrects="statement-a")
        company.validate_canon(self.canon)
        self.assertEqual(self.record("statement-a")["data"]["value"], "August")
        self.record("statement-b")["known_on"] = "2022-04-01"
        with self.assertRaisesRegex(company.Invalid, "postdates"):
            company.validate_canon(self.canon)

    def test_opinion_qualifiers_and_required_statement_claims(self):
        self.source("source-a")
        self.statement("statement-a", "source-a", "Proceed cautiously", kind="opinion",
                       qualifiers=["uncertain", "dissent", "unresolved"])
        self.condition("noise", sources=["source-a"])
        plan = self.plan()
        doc = next(doc for doc in plan["documents"] if doc["source_version_id"] == "source-a")
        draft = self.fixture.draft(doc)
        with self.assertRaisesRegex(company.Invalid, "required source assertions"):
            company.check_draft(self.canon, plan, draft)
        draft["claims"].append({"id": "position", "record_id": "statement-a",
                                "path": ["data", "value"], "value": "Proceed cautiously",
                                "statement_kind": "opinion"})
        draft["sections"][0]["claim_ids"].append("position")
        company.check_draft(self.canon, plan, draft)
        draft["claims"][-1]["statement_kind"] = "fact"
        with self.assertRaisesRegex(company.Invalid, "misclassifies"):
            company.check_draft(self.canon, plan, draft)
        draft["claims"][-1]["statement_kind"] = "opinion"
        draft["sections"][0]["speaker_id"] = "person-morgan"
        with self.assertRaisesRegex(company.Invalid, "wrong speaker"):
            company.check_draft(self.canon, plan, draft)
        draft["sections"][0]["reported_speech"] = True
        company.check_draft(self.canon, plan, draft)

    def test_transitive_corrections_and_source_supersession(self):
        self.source("source-a")
        self.statement("statement-a", "source-a", "August")
        self.source("source-b", when="2022-03-01", supersedes="source-a")
        self.statement("statement-b", "source-b", "September", when="2022-03-01",
                       corrects="statement-a")
        self.source("source-c", when="2022-04-01", supersedes="source-b")
        self.statement("statement-c", "source-c", "October", when="2022-04-01",
                       corrects="statement-b")
        company.validate_canon(self.canon)
        for ref in ("statement-b", "statement-c"):
            record = self.record(ref)
            record["refs"].remove(record["data"].pop("corrects"))
        company.validate_canon(self.canon)

    def test_lifecycle_cycles_and_scope_mismatch_fail(self):
        self.source("source-a")
        first = self.statement("statement-a", "source-a", "August")
        second = self.statement("statement-b", "source-a", "September", corrects="statement-a")
        first["data"]["corrects"] = "statement-b"
        first["refs"].append("statement-b")
        with self.assertRaisesRegex(company.Invalid, "cyclic"):
            company.validate_canon(self.canon)
        first["data"].pop("corrects")
        first["refs"].remove("statement-b")
        second["data"]["predicate"] = "support_capacity"
        with self.assertRaisesRegex(company.Invalid, "same scope"):
            company.validate_canon(self.canon)

    def test_source_access_defaults_and_explicit_groups(self):
        source = self.source("source-a")
        self.condition("noise", sources=["source-a"])
        doc = next(doc for doc in self.plan()["documents"] if doc["source_version_id"] == "source-a")
        self.assertEqual(doc["access_groups"], ["group-internal"])
        source["access_groups"] = ["group-board"]
        source["refs"].append("group-board")
        doc = next(doc for doc in self.plan()["documents"] if doc["source_version_id"] == "source-a")
        self.assertEqual(doc["access_groups"], ["group-board"])
        draft = self.fixture.draft(doc)
        draft["claims"].append({"id": "policy", "record_id": "source-a",
                                "path": ["access_groups"], "value": ["group-board"]})
        draft["sections"][0]["claim_ids"].append("policy")
        with self.assertRaisesRegex(company.Invalid, "Invalid claim path"):
            company.check_draft(self.canon, self.fixture.plan, draft)

    def test_viewpoint_coverage_and_author_knowledge_are_explicit(self):
        self.add("view-finance", "viewpoint", {"name": "Finance", "author_id": "person-morgan",
                 "audience": "internal", "focus_ids": ["product-beacon"]},
                 ["person-morgan", "product-beacon"])
        self.add("view-engineering", "viewpoint", {"name": "Engineering", "author_id": "person-riley",
                 "audience": "internal"}, ["person-riley"])
        self.canon["request"]["viewpoints"] = ["view-finance", "view-engineering"]
        plan = self.plan()
        self.assertEqual({doc["author_id"] for doc in plan["documents"]}, {"person-morgan", "person-riley"})
        self.assertEqual({doc["viewpoint_id"] for doc in plan["documents"]}, {"view-finance", "view-engineering"})
        self.add("private-note", "business_note", {"title": "Readiness risk", "value": "Limited capacity"},
                 ["product-beacon"], knowledge_restricted=True)
        plan = self.plan()
        doc = plan["documents"][0]
        self.assertNotIn("private-note", {r["id"] for r in company.get_context(self.canon, plan, doc["id"])["records"]})
        self.add("grant-note", "knowledge_grant", {"record_id": "private-note",
                 "principal_id": doc["author_id"]}, ["private-note", doc["author_id"]])
        plan = self.plan()
        self.assertIn("private-note", {r["id"] for r in company.get_context(self.canon, plan, doc["id"])["records"]})

    def test_viewpoint_author_is_a_bounded_identity_without_invented_roles(self):
        self.add("person-partner-author", "person", {"name": "Sage Vale"})
        self.add("view-partner", "viewpoint",
                 {"name": "Partner", "author_id": "person-partner-author", "audience": "internal"},
                 ["person-partner-author"])
        self.canon["request"]["viewpoints"] = ["view-partner"]
        plan = self.plan()
        doc = next(doc for doc in plan["documents"] if doc["viewpoint_id"] == "view-partner")
        context = company.get_context(self.canon, plan, doc["id"])
        self.assertIn("person-partner-author", {person["id"] for person in context["org_chart"]["people"]})
        self.assertFalse(any(role["data"]["person_id"] == "person-partner-author"
                             for role in context["active_roles"]))

    def test_restricted_dependencies_and_release_to_public(self):
        self.add("private-note", "business_note", {"title": "Release detail", "value": "Limited capacity"},
                 ["product-beacon", "group-board"], access_groups=["group-board"])
        self.add("dependent-note", "business_note", {"title": "Related detail"}, ["private-note"])
        index = company.validate_canon(self.canon)
        doc = {"publication_date": "2022-02-01", "author_id": "person-riley"}
        visible = company.project_records(index, doc)
        self.assertNotIn("private-note", visible)
        self.assertNotIn("dependent-note", visible)
        self.add("release-note", "release", {"record_id": "private-note", "groups": [], "public": True},
                 ["private-note"])
        index = company.validate_canon(self.canon)
        self.assertTrue(company.accessible(self.record("private-note"), index, "public", "2022-02-01"))

    def test_claims_require_release_to_intended_readers(self):
        self.canon["request"]["document_types"] = ["press_release"]
        self.source("source-a", document_type="press_release", audience="public")
        statement = self.statement("statement-a", "source-a", "August")
        statement["access_groups"] = ["group-board"]
        statement["refs"].append("group-board")
        self.condition("noise", sources=["source-a"])

        def draft_for_source():
            plan = self.plan()
            doc = next(doc for doc in plan["documents"] if doc["source_version_id"] == "source-a")
            draft = self.fixture.draft(doc)
            draft["claims"].append({"id": "position", "record_id": "statement-a",
                                    "path": ["data", "value"], "value": "August"})
            draft["sections"][0]["claim_ids"].append("position")
            return plan, draft

        plan, draft = draft_for_source()
        with self.assertRaisesRegex(company.Invalid, "intended"):
            company.check_draft(self.canon, plan, draft)
        self.add("release-position", "release",
                 {"record_id": "statement-a", "groups": [], "public": True}, ["statement-a"])
        plan, draft = draft_for_source()
        company.check_draft(self.canon, plan, draft)

    def test_supported_derivations_reject_cycles_units_and_zero(self):
        claims = {
            "a": {"id": "a", "value": 100, "units": "USD"},
            "b": {"id": "b", "value": 40, "units": "USD"},
            "c": {"id": "c", "value": "60.00", "units": "USD",
                  "derivation": {"operation": "difference", "inputs": ["a", "b"], "rounding": 2}},
        }
        self.assertEqual(company.evaluate_claim("c", claims), "60.00")
        claims["c"]["record_id"] = "private-record"
        with self.assertRaisesRegex(company.Invalid, "cannot also be a primary"):
            company.evaluate_claim("c", claims)
        claims["c"].pop("record_id")
        claims["c"]["derivation"]["inputs"] = ["c", "b"]
        with self.assertRaisesRegex(company.Invalid, "cyclic"):
            company.evaluate_claim("c", claims)
        claims["c"].update(units="percent", value="40.0")
        claims["c"]["derivation"] = {"operation": "ratio", "inputs": ["b", "a"], "rounding": 1}
        self.assertEqual(company.evaluate_claim("c", claims), "40.0")
        claims["a"]["value"] = 0
        with self.assertRaisesRegex(company.Invalid, "division by zero"):
            company.evaluate_claim("c", claims)

    def test_claim_units_cannot_relabel_source_values(self):
        self.source("source-a")
        self.statement("statement-a", "source-a", 120, units="USD")
        self.condition("noise", sources=["source-a"])
        plan = self.plan()
        doc = next(doc for doc in plan["documents"] if doc["source_version_id"] == "source-a")
        draft = self.fixture.draft(doc)
        draft["claims"].append({"id": "cost", "record_id": "statement-a",
                                "path": ["data", "value"], "value": 120, "units": "EUR"})
        draft["sections"][0]["claim_ids"].append("cost")
        with self.assertRaisesRegex(company.Invalid, "units differ"):
            company.check_draft(self.canon, plan, draft)
        draft["claims"][-1]["units"] = "USD"
        company.check_draft(self.canon, plan, draft)


if __name__ == "__main__":
    unittest.main()
