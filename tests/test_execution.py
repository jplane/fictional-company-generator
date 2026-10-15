import base64
import copy
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch
from xml.sax.saxutils import escape

import test_company as baseline

company = baseline.company
sys.modules["company"] = company
CAPTURE_SPEC = importlib.util.spec_from_file_location(
    "playwright_capture", baseline.SCRIPT.with_name("playwright_capture.py"))
playwright_capture = importlib.util.module_from_spec(CAPTURE_SPEC)
CAPTURE_SPEC.loader.exec_module(playwright_capture)
sys.modules["playwright_capture"] = playwright_capture
SPEC = importlib.util.spec_from_file_location("officecli_runner", baseline.SCRIPT.with_name("officecli_runner.py"))
runner = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(runner)
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aCfcAAAAASUVORK5CYII=")
TOOL = {"version": "test-fixture", "sha256": "f" * 64}


CAPTURE_TOOL = {"engine": "playwright-chromium", "playwright_version": "1.63.0", "channel": "chromium",
                "headless": True, "chromium_sandbox": True, "browser_version": "153.0.8010.12",
                "executable_path": "fixture-chromium", "executable_sha256": "e" * 64}


def log(operation, arguments=None):
    data = {"count": 0, "errors" if operation == "validate" else "issues": []}
    if operation in ("author", "formula_correction"):
        data = {"summary": {"total": 1, "executed": 1, "succeeded": 1, "failed": 0, "skipped": 0}}
    if operation == "html_export":
        data = "fixture.html"
    return {"operation": operation, "arguments": arguments or [], "tool": TOOL, "returncode": 0,
            "stdout": json.dumps({"success": True, "data": data}), "stderr": "",
            "active_seconds": 0.01, "status": "passed", "evidence": "Mocked test evidence only"}


def capture_log(fmt, image_sha256):
    unit = {"docx": "pages", "pptx": "slides", "xlsx": "region"}[fmt]
    entry = {"operation": "capture", "status": "passed", "tool": CAPTURE_TOOL,
              "source_html_sha256": "a" * 64, "unit": unit, "image_sha256": image_sha256}
    if unit in ("pages", "slides"):
        entry["rendered_count"] = 1
    else:
        entry.update(start_cell="A1", end_cell="A1")
    return entry


def capability(root, formats):
    image = root / "build" / "probe.png"
    image.parent.mkdir(parents=True, exist_ok=True)
    image.write_bytes(PNG)
    image_sha256 = company.file_hash(image)
    result = {"preflight_version": 2, "engine": "OfficeCLI", "tool": TOOL, "capture_tool": CAPTURE_TOOL,
              "status": "passed", "formats": sorted(formats), "probes": []}
    for fmt in sorted(formats):
        operations = ["create", "author", "validate", "issues", "extract", "html_export"]
        if fmt == "xlsx":
            operations.append("chart_cache_stale")
        logs = [log(op) for op in operations]
        result["probes"].append({"format": fmt, "logs": logs, "formula_count": 2 if fmt == "xlsx" else 0,
                                "features": ["basic", "html_screenshots", "ordered_cross_sheet_formulas"],
                                "image": "probe.png", "image_sha256": image_sha256,
                                "capture": capture_log(fmt, image_sha256)})
    runner.write_json(root / "build" / "preflight.json", result)
    return result


class ExecutionTests(unittest.TestCase):
    def setUp(self):
        self.fixture = baseline.CompanyTests()
        self.fixture.setUp()
        self.canon = company.seed_canon(dict(self.fixture.request, documents=9))
        self.plan = company.create_plan(self.canon)
        self.fixture.canon, self.fixture.plan = self.canon, self.plan

    def execution(self, policy=company.DEFAULT_QA_POLICY):
        layouts = [{"document_id": d["id"], "layout_revision": "v1", "layout_sha256": "a" * 64}
                   for d in self.plan["documents"]]
        seen, pilots = set(), []
        for doc in self.plan["documents"]:
            family = (doc["type"], doc["format"])
            if family not in seen:
                pilots.append(doc["id"])
                seen.add(family)
        return company.execution_plan(self.canon, self.plan, layouts, pilots, policy)

    def enable(self, root, manifest, qa, execution):
        drafts = root / "build" / "drafts"
        files = company.manifest_files(self.canon, self.plan, manifest, root)
        probe = capability(root, {d["format"] for d in self.plan["documents"]})
        cap_path = root / "build" / "capabilities" / "fixture.json"
        runner.write_json(cap_path, probe)
        targets = dict(schema_version=2, plan_sha256=company.fingerprint(self.plan), documents=[
            {"document_id": d["id"], "content_depth": d["content_depth"],
             "measure": "delivered_words", "target_min": 1, "target_max": 100000}
            for d in self.plan["documents"]])
        runner.write_json(root / "build" / "depth-targets.json", targets)
        for doc in self.plan["documents"]:
            identifier, path = doc["id"], files[doc["id"]]
            digest = company.file_hash(path)
            checks = []
            operations = ["validate", "issues", "extract", "depth", "document"]
            if doc["format"] == "xlsx":
                operations += ["formulas", "chart_cache_stale"]
            for op in operations:
                result = dict(log(op), artifact_sha256=digest)
                if op == "formulas":
                    result["bindings"] = []
                if op == "depth":
                    result.update(target_sha256=company.file_hash(root / "build" / "depth-targets.json"),
                        measure="delivered_words",
                        measured=company.depth_measure(company.artifact_containers(path, doc["format"]),
                                                       "delivered_words"))
                result_path = root / "build" / "tool-results" / f"{identifier}-{op}.json"
                runner.write_json(result_path, result)
                checks.append({"operation": op, "path": str(result_path.relative_to(root)),
                               "sha256": company.file_hash(result_path)})
            html_result_path = root / "build" / "tool-results" / f"{identifier}-html_export.json"
            runner.write_json(html_result_path, dict(log("html_export"), artifact_sha256=digest))
            checks.append({"operation": "html_export", "path": str(html_result_path.relative_to(root)),
                           "sha256": company.file_hash(html_result_path)})
            image = root / "build" / "renders" / f"{identifier}.png"
            image.parent.mkdir(exist_ok=True)
            image.write_bytes(PNG)
            image_ref = {"path": str(image.relative_to(root)), "sha256": company.file_hash(image)}
            if doc["format"] == "xlsx":
                region = f"Results!A1:A{len(self.fixture.draft(doc)['sections'])}"
                definitions = [{"id": "all", "range": region}]
                capture = capture_log("xlsx", image_ref["sha256"])
            elif doc["format"] == "docx":
                definitions = [{"id": "all", "pages": 1}]
                capture = capture_log("docx", image_ref["sha256"])
                image_ref.update(page=1, overflow=False)
            else:
                definitions = [{"id": "all", "slides": 1}]
                capture = capture_log("pptx", image_ref["sha256"])
                image_ref.update(slide=1)
            result_path = root / "build" / "tool-results" / f"{identifier}-render-all.json"
            runner.write_json(result_path, dict(capture, artifact_sha256=digest))
            machine = dict(company.receipt_header(self.canon, self.plan), engine="OfficeCLI", tool=TOOL,
                document_id=identifier, artifact_sha256=digest, draft_sha256=company.file_hash(drafts / f"{identifier}.json"),
                status="passed", checks=checks, render_targets=["all"], target_definitions=definitions,
                capability={"path": str(cap_path.relative_to(root)), "sha256": company.file_hash(cap_path)},
                renders=[{"id": "all", "images": [image_ref],
                    "renderer": "OfficeCLI-HTML+Playwright-Chromium", "artifact_sha256": digest,
                    "result": {"path": str(result_path.relative_to(root)), "sha256": company.file_hash(result_path)}}])
            runner.write_json(root / "build" / "machine" / f"{identifier}.json", machine)
        runner.write_json(root / "build" / "execution-plan.json", execution)
        selection = company.make_review_plan(self.canon, self.plan, drafts, execution, root, files)
        runner.write_json(root / "build" / "review-plan.json", selection)
        qa.update(execution_version=1, execution_sha256=company.fingerprint(execution),
                  review_plan_sha256=company.fingerprint(selection), qa_policy=execution["qa_policy"])
        by_id = {e["document_id"]: e for e in selection["documents"]}
        for review in qa["documents"]:
            entry = by_id[review["document_id"]]
            review["machine_sha256"] = entry["machine_sha256"]
            review["condition_review"] = {"status": "not_selected"}
            for gate in ("semantic_review", "genre_review", "depth_review", "visual_review"):
                required = entry["visual_required"] if gate == "visual_review" else entry["content_required"]
                if not required:
                    review[gate] = {"status": "not_selected"}
                elif gate == "visual_review":
                    review[gate].update(images_delivered=True, render_ids=["all"])
        runner.write_json(root / "build" / "qa.json", qa)
        return selection

    def test_all_policies_audit_and_legacy_remains_full(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            drafts = root / "build" / "drafts"
            manifest, qa = self.fixture.create_audit_fixture(root, drafts, include_titles=True)
            result = company.audit(self.canon, self.plan, drafts, manifest, root)
            self.assertEqual(result["qa_policy"], "legacy-all-document")
            for policy in company.QA_POLICIES:
                current = copy.deepcopy(qa)
                selection = self.enable(root, manifest, current, self.execution(policy))
                result = company.audit(self.canon, self.plan, drafts, manifest, root)
                self.assertEqual(result["qa_policy"], policy)
                self.assertEqual(result["review_coverage"]["content"],
                                 sum(e["content_required"] for e in selection["documents"]))
                self.assertEqual(result["review_coverage"]["visual"],
                                 sum(e["visual_required"] for e in selection["documents"]))

    def test_default_policy_and_deterministic_family_sampling(self):
        self.canon = company.seed_canon(dict(self.fixture.request, documents=12,
                                             document_types=["press_release"], formats=["docx"]))
        self.plan = company.create_plan(self.canon)
        self.fixture.canon, self.fixture.plan = self.canon, self.plan
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            drafts = root / "build" / "drafts"
            self.fixture.create_audit_fixture(root, drafts)
            execution = self.execution("Sampled")
            first = self.plan["documents"][0]["id"]
            last = self.plan["documents"][-1]["id"]
            draft = company.load(drafts / f"{last}.json")
            draft["sections"][0]["text"] += " Additional grounded context." * 20
            runner.write_json(drafts / f"{last}.json", draft)
            selected = company.review_selection(self.canon, self.plan, drafts, execution)
            self.assertEqual({e["document_id"] for e in selected if e["visual_required"]}, {first, last})
            self.assertEqual(selected, company.review_selection(self.canon, self.plan, drafts, execution))
            self.assertEqual(self.execution()["qa_policy"], "FullContentSampledVisual")
            execution["flagged_ids"] = [last]
            expanded = company.review_selection(self.canon, self.plan, drafts, execution)
            self.assertTrue(all(e["visual_required"] and e["content_required"] for e in expanded))
            execution["flagged_ids"] = []
            expanded = company.review_selection(self.canon, self.plan, drafts, execution,
                                                {last: {"corrected_failure": True}})
            self.assertTrue(all(e["visual_required"] for e in expanded))

    def test_selection_and_receipt_tampering_fail(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            drafts = root / "build" / "drafts"
            manifest, qa = self.fixture.create_audit_fixture(root, drafts, include_titles=True)
            self.enable(root, manifest, qa, self.execution())
            original = copy.deepcopy(qa)
            for mutation, message in [
                (lambda q: q.update(qa_policy="Sampled"), "policy/execution"),
                (lambda q: q["documents"][0]["visual_review"].update(images_delivered=False), "undelivered"),
                (lambda q: q["documents"][0]["visual_review"].update(render_ids=[]), "incomplete"),
                (lambda q: q["documents"][0].update(machine_sha256="0" * 64), "machine evidence"),
            ]:
                changed = copy.deepcopy(original)
                mutation(changed)
                runner.write_json(root / "build" / "qa.json", changed)
                with self.assertRaisesRegex(company.Invalid, message):
                    company.audit(self.canon, self.plan, drafts, manifest, root)
            runner.write_json(root / "build" / "qa.json", original)
            reviews = company.load(root / "build" / "review-plan.json")
            reviews["documents"][0]["visual_required"] = False
            runner.write_json(root / "build" / "review-plan.json", reviews)
            with self.assertRaisesRegex(company.Invalid, "selection"):
                company.audit(self.canon, self.plan, drafts, manifest, root)

    def test_hash_invalidation_and_no_legacy_downgrade(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            drafts = root / "build" / "drafts"
            manifest, qa = self.fixture.create_audit_fixture(root, drafts, include_titles=True)
            self.enable(root, manifest, qa, self.execution())
            del qa["execution_version"]
            runner.write_json(root / "build" / "qa.json", qa)
            with self.assertRaisesRegex(company.Invalid, "policy/execution"):
                company.audit(self.canon, self.plan, drafts, manifest, root)
            image = root / "build" / "renders" / f"{self.plan['documents'][0]['id']}.png"
            image.write_bytes(b"not an image")
            with self.assertRaisesRegex(company.Invalid, "render evidence"):
                company.audit(self.canon, self.plan, drafts, manifest, root)

    def test_execution_rejects_missing_pilots_ids_policy_and_budget(self):
        execution = self.execution()
        for key, value in (("pilot_ids", []), ("qa_policy", "Standard"),
                           ("budget_seconds", -1), ("revision", 0), ("execution_version", 2)):
            changed = dict(execution, **{key: value})
            with self.assertRaises(company.Invalid):
                company.validate_execution(self.canon, self.plan, changed)
        changed = copy.deepcopy(execution)
        changed["documents"].append(changed["documents"][0])
        with self.assertRaisesRegex(company.Invalid, "duplicate"):
            company.validate_execution(self.canon, self.plan, changed)

    def test_nonselection_is_not_a_pass_and_depth_stays_independent(self):
        doc = self.plan["documents"][0]
        selected = {"content_required": False, "visual_required": False}
        review = {g: {"status": "not_selected"} for g in (
            "semantic_review", "genre_review", "depth_review", "visual_review", "condition_review")}
        company.validate_review(review, doc, selected, {"render_targets": []})
        review["visual_review"] = {"status": "passed", "evidence": "No review occurred"}
        with self.assertRaisesRegex(company.Invalid, "not_selected"):
            company.validate_review(review, doc, selected, {"render_targets": []})

    def test_source_history_and_condition_carriers_are_mandatory(self):
        from test_content_model import ContentModelTests
        model = ContentModelTests()
        model.setUp()
        model.canon["request"]["documents"] = 8
        model.source("source-a")
        model.statement("statement-a", "source-a", "August")
        model.source("source-b", when="2022-03-01", supersedes="source-a")
        model.statement("statement-b", "source-b", "September", when="2022-03-01", corrects="statement-a")
        model.condition("stale", sources=["source-a"], as_of_date="2022-03-01")
        model.condition("noise", sources=["source-b"])
        self.canon, self.plan = model.canon, model.plan()
        self.fixture.canon, self.fixture.plan = self.canon, self.plan
        with tempfile.TemporaryDirectory() as directory:
            drafts = Path(directory)
            for doc in self.plan["documents"]:
                draft = self.fixture.draft(doc)
                for statement in doc["required_statement_ids"]:
                    value = model.record(statement)["data"]["value"]
                    draft["claims"].append({"id": statement, "record_id": statement,
                                            "path": ["data", "value"], "value": value})
                    draft["sections"][0]["claim_ids"].append(statement)
                    draft["sections"][0]["text"] += " " + value
                company.save(drafts / f"{doc['id']}.json", draft)
            selected = company.review_selection(self.canon, self.plan, drafts, self.execution("Sampled"))
            by_id = {d["id"]: d for d in self.plan["documents"]}
            carriers = [e for e in selected if by_id[e["document_id"]]["source_version_id"]]
            self.assertEqual(len(carriers), 2)
            self.assertTrue(all("source_history" in e["reasons"] and e["content_required"]
                                and e["visual_required"] for e in carriers))
            old = next(e for e in carriers if by_id[e["document_id"]]["source_version_id"] == "source-a")
            self.assertIn("condition", old["reasons"])

    def test_policy_selection_does_not_change_with_content_depth(self):
        results = []
        for depth in company.CONTENT_DEPTHS:
            self.canon = company.seed_canon(dict(self.fixture.request, documents=12,
                document_types=["press_release"], formats=["docx"], content_depth=depth))
            self.plan = company.create_plan(self.canon)
            self.fixture.canon, self.fixture.plan = self.canon, self.plan
            with tempfile.TemporaryDirectory() as directory:
                drafts = Path(directory)
                for doc in self.plan["documents"]:
                    company.save(drafts / f"{doc['id']}.json", self.fixture.draft(doc))
                selected = company.review_selection(self.canon, self.plan, drafts, self.execution("Sampled"))
                results.append({e["document_id"] for e in selected if e["visual_required"]})
        self.assertEqual(results[0], results[1])
        self.assertEqual(results[1], results[2])


class RunnerPrimitiveTests(unittest.TestCase):
    def test_delivered_measurements_do_not_count_locator_aliases_twice(self):
        first = company.joined_container({"kind": "paragraph"}, [({"paragraph": 1}, "Two words")])
        alias = company.joined_container({"kind": "table_cell"}, [({"paragraph": 1}, "Two words")])
        self.assertEqual(company.depth_measure([first, alias], "delivered_words"), 2)
        self.assertEqual(company.depth_measure([first, alias], "delivered_containers"), 1)
        distinct = company.joined_container({"kind": "paragraph"}, [({"paragraph": 2}, "Two words")])
        self.assertEqual(company.depth_measure([first, distinct], "delivered_words"), 4)

    def test_unavailable_capture_adapter_preflight_keeps_failure_evidence(self):
        """When the optional Playwright capture adapter is missing, every completed
        OfficeCLI probe operation (create/author/validate/issues/extract/html_export)
        stays recorded, and the receipt is blocked rather than silently passing without
        real renderer proof."""
        class FixtureCLI:
            tool = TOOL
            timeout = 1

            def run(self, arguments, operation):
                return log(operation, arguments)

        def missing_dependency():
            raise playwright_capture.CaptureError(
                "Playwright capture adapter unavailable; install the pinned Playwright "
                "dependency and its matching managed Chromium build before capturing")

        with tempfile.TemporaryDirectory() as directory:
            out = Path(directory) / "preflight.json"
            # Deterministically simulate the "Playwright is not installed" precondition by
            # forcing the adapter's own lazy-import guard to fail, regardless of whether
            # Playwright actually happens to be installed in the environment running this
            # test. This keeps the real CaptureEngine construction path and its genuine
            # CaptureError/missing-dependency semantics intact rather than widening a catch,
            # skipping the test, or weakening the preflight's error handling.
            with patch.object(playwright_capture, "_import_playwright", side_effect=missing_dependency):
                with self.assertRaisesRegex(playwright_capture.CaptureError,
                                             "Playwright capture adapter unavailable"):
                    runner.preflight(FixtureCLI(), ["docx"], out)
            receipt = company.load(out)
            self.assertEqual(receipt["status"], "blocked")
            self.assertIn("Playwright capture adapter unavailable", receipt["blocker"])
            logged = {entry["operation"] for entry in receipt["probes"][0]["logs"]}
            self.assertEqual(logged, {"create", "author", "validate", "issues", "extract", "html_export"})
            self.assertNotIn("capture", receipt["probes"][0])

    def test_complete_response_rejects_failure_caveats_and_truncation(self):
        for operation in ("validate", "issues", "chart_cache_stale", "author"):
            good = json.loads(log(operation)["stdout"])
            runner.response(json.dumps(good), operation)
            for change in ({"success": False}, {"warnings": ["unsupported"]}, {"truncated": True}):
                with self.assertRaises(company.Invalid):
                    runner.response(json.dumps(dict(good, **change)), operation)
        with self.assertRaisesRegex(company.Invalid, "complete JSON"):
            runner.response('{"success":', "issues")
        with self.assertRaisesRegex(company.Invalid, "incomplete scan"):
            runner.response('{"success":true,"data":{"count":1,"issues":[]}}', "issues")

    def test_setup_decisions_and_release_verification(self):
        self.assertEqual(runner.setup_action(None, "1.0", False), "authorization_required")
        self.assertEqual(runner.setup_action(None, "1.0", False, authorized=True), "install")
        self.assertEqual(runner.setup_action("1.0", "2.0", True), "offer_update")
        self.assertEqual(runner.setup_action("1.0", "1.0", True), "use_installed")
        self.assertEqual(runner.setup_action("1.0", None, True), "update_lookup_unavailable")
        self.assertEqual(runner.setup_action("1.0", "2.0", False, authorized=True), "update")
        self.assertEqual(runner.setup_action("1.0", "2.0", True, True, True), "update")
        with self.assertRaises(company.Invalid):
            runner.setup_action(None, None, False, authorized=True)
        with tempfile.TemporaryDirectory() as directory:
            asset = Path(directory) / "officecli"
            asset.write_bytes(b"fixture binary")
            digest = company.file_hash(asset)
            runner.verify_release(asset, digest, "1.0", "1.0")
            for checksum, version in ((None, "1.0"), ("0" * 64, "1.0"), (digest, "2.0")):
                with self.assertRaises(company.Invalid):
                    runner.verify_release(asset, checksum, version, "1.0")

    def test_required_features_cannot_be_claimed_from_basic_probes(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            receipt = capability(root, {"docx"})
            company.validate_required_features({"required_features": {"docx": ["basic"]}}, receipt)
            with self.assertRaisesRegex(company.Invalid, "explicit capability probes"):
                company.validate_required_features({"required_features": {"docx": ["page_fields"]}}, receipt)

    def test_engine_drift_and_forbidden_native_operations(self):
        with tempfile.TemporaryDirectory() as directory:
            binary = Path(directory) / "officecli.exe"
            binary.write_bytes(b"first version")
            version = subprocess.CompletedProcess([], 0, "1.0\n", "")
            with patch.object(runner.shutil, "which", return_value=str(binary)), \
                    patch.object(runner.subprocess, "run", return_value=version):
                cli = runner.OfficeCLI()
                for args in (["refresh", "a.docx"], ["view", "a.docx", "stats", "--page-count"],
                             ["view", "a.docx", "screenshot"], ["view", "a.docx", "screenshot", "--render", "auto"]):
                    with self.assertRaises(company.Invalid):
                        cli.run(args, "test")
                binary.write_bytes(b"changed version")
                with self.assertRaisesRegex(company.Invalid, "drift"):
                    cli.run(["view", "a.docx", "text"], "test")

    def test_retry_timeout_and_budget_are_persistent(self):
        class FakeCLI:
            timeout = 1
            calls = 0

            def run(self, arguments, operation):
                self.calls += 1
                return dict(log(operation), returncode=1, stderr="Failure")

        with tempfile.TemporaryDirectory() as directory:
            engine = object.__new__(runner.Runner)
            engine.state = {"documents": {}}
            engine.state_path = Path(directory) / "state.json"
            engine.cli = FakeCLI()
            state = {"attempts": {}, "operations": [], "retry_reasons": {"issues": "Fixed browser configuration"}}
            for _ in range(2):
                with self.assertRaises(company.Invalid):
                    engine.operation(state, [], "issues")
            with self.assertRaisesRegex(company.Invalid, "already consumed"):
                engine.operation(state, [], "issues")
            self.assertEqual(engine.cli.calls, 2)
            self.assertEqual(company.load(engine.state_path)["documents"], {})
            with patch.object(engine.cli, "run", side_effect=subprocess.TimeoutExpired("officecli", 1)):
                with self.assertRaisesRegex(company.Invalid, "timed out"):
                    engine.operation(state, [], "render")
            self.assertEqual(state["operations"][-1]["status"], "timeout")
            engine.execution = {"budget_seconds": 1}
            engine.state["documents"]["doc"] = {"operations": [{"active_seconds": 2}]}
            with self.assertRaisesRegex(company.Invalid, "budget reached"):
                engine.boundary()

    def test_successful_inspections_recheck_new_hashes_without_retry_or_expansion(self):
        with tempfile.TemporaryDirectory() as directory:
            engine = object.__new__(runner.Runner)
            engine.state = {"documents": {}}
            engine.state_path = Path(directory) / "state.json"
            engine.execution = {}
            engine.cli = type("FakeCLI", (), {"tool": TOOL})()
            with patch.object(engine.cli, "run", create=True,
                              side_effect=lambda args, op: log(op)) as run:
                state = {"attempts": {}, "operations": []}
                for digest in ("a" * 64, "b" * 64, "c" * 64):
                    result = engine.inspection(state, [], "validate", digest)
                    self.assertEqual(result["artifact_sha256"], digest)
                    self.assertFalse(result["corrective"])
                engine.inspection(state, [], "validate", "c" * 64)
                self.assertEqual(run.call_count, 3)
                self.assertEqual(state["attempts"]["validate"], 3)
                self.assertFalse(engine.corrected_failure(state))
                state["attempts"]["author"] = 2
                self.assertTrue(engine.corrected_failure(state))

    def test_inspection_failures_remain_bounded_across_hashes_and_resume(self):
        with tempfile.TemporaryDirectory() as directory:
            engine = object.__new__(runner.Runner)
            engine.state_path = Path(directory) / "state.json"
            engine.state = {"documents": {"doc": {"attempts": {}, "operations": []}}}
            engine.execution = {}
            engine.cli = type("FakeCLI", (), {"tool": TOOL})()
            failure = dict(log("issues"), returncode=1, stderr="Failure")
            with patch.object(engine.cli, "run", create=True, return_value=failure) as run:
                state = engine.state["documents"]["doc"]
                with self.assertRaises(company.Invalid):
                    engine.inspection(state, [], "issues", "a" * 64)
                self.assertEqual(state["operations"][-1]["artifact_sha256"], "a" * 64)
                with self.assertRaisesRegex(company.Invalid, "concrete correction"):
                    engine.inspection(state, [], "issues", "b" * 64)
                self.assertEqual(run.call_count, 1)
                state["retry_reasons"] = {"issues": "Corrected explicit layout"}
                with self.assertRaises(company.Invalid):
                    engine.inspection(state, [], "issues", "b" * 64)
                engine.state = company.load(engine.state_path)
                state = engine.state["documents"]["doc"]
                with self.assertRaisesRegex(company.Invalid, "already consumed"):
                    engine.inspection(state, [], "issues", "c" * 64)
                self.assertEqual(run.call_count, 2)
                self.assertTrue(engine.corrected_failure(state))

    def test_successful_rechecks_do_not_erase_an_earlier_failed_inspection(self):
        with tempfile.TemporaryDirectory() as directory:
            engine = object.__new__(runner.Runner)
            engine.state_path = Path(directory) / "state.json"
            engine.state = {"documents": {}}
            engine.execution = {}
            engine.cli = type("FakeCLI", (), {"tool": TOOL})()
            state = {"attempts": {}, "operations": [],
                     "retry_reasons": {"issues": "Corrected explicit layout"}}
            failure = dict(log("issues"), returncode=1, stderr="Failure")
            results = [failure, log("issues"), log("issues"), failure]
            with patch.object(engine.cli, "run", create=True, side_effect=results) as run:
                with self.assertRaises(company.Invalid):
                    engine.inspection(state, [], "issues", "a" * 64)
                self.assertTrue(engine.inspection(state, [], "issues", "b" * 64)["corrective"])
                self.assertFalse(engine.inspection(state, [], "issues", "c" * 64)["corrective"])
                with self.assertRaises(company.Invalid):
                    engine.inspection(state, [], "issues", "d" * 64)
                with self.assertRaisesRegex(company.Invalid, "already consumed"):
                    engine.inspection(state, [], "issues", "e" * 64)
                self.assertEqual(run.call_count, 4)
                self.assertTrue(engine.corrected_failure(state))

    def test_cache_correction_is_grounded_ordered_and_only_once(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "cache.xlsx"
            path.write_bytes(b"Mock package; formula reads are patched")
            engine = object.__new__(runner.Runner)
            engine.cli = FakeCLI()
            engine.state, engine.state_path = {"documents": {}}, Path(directory) / "state.json"
            bindings = [{"sheet": "Sheet1", "cell": "A2", "formula": "A1*2", "value": 2, "claim_id": "total"}]
            draft = {"claims": [{"id": "total", "value": 2, "derivation": {"operation": "sum"}}]}
            state = {"attempts": {}, "operations": []}
            good = {"formula_count": 1, "evidence": "Checked mock cache"}
            with patch.object(runner, "formula_values", side_effect=[
                    company.Invalid("formula has no saved cached value"), good]), \
                    patch.object(engine.cli, "run", return_value=log("formula_correction")):
                result = engine.formulas(state, path, {"expected_formulas": bindings}, draft)
            self.assertEqual(result["bindings"], bindings)
            self.assertEqual(state["formula_retry"]["status"], "passed")
            self.assertEqual(state["attempts"]["formula_correction"], 1)
            with patch.object(runner, "formula_values", side_effect=company.Invalid("stale/incorrect formula cache")):
                with self.assertRaisesRegex(company.Invalid, "already consumed"):
                    engine.formulas(state, path, {"expected_formulas": bindings}, draft)


class FakeCLI:
    tool = TOOL
    timeout = 5

    def __init__(self):
        self.calls = []

    def run(self, args, operation):
        self.calls.append((operation, args))
        if args[0] == "create":
            self.write_doc(Path(args[1]), [], "")
        if args[0] == "batch":
            commands = company.load(args[args.index("--input") + 1])
            paragraphs = [c["props"]["text"] for c in commands if c["command"] == "add"]
            title = next(c["props"]["title"] for c in commands if c["command"] == "set")
            self.write_doc(Path(args[1]), paragraphs, title)
        if "html" in args:
            Path(args[args.index("--out") + 1]).write_text(
                "<div class='page' data-page='1'>fixture</div>", encoding="utf-8")
        return log(operation, args)

    @staticmethod
    def write_doc(path, paragraphs, title):
        with zipfile.ZipFile(path, "w") as package:
            package.writestr("[Content_Types].xml", "<Types/>")
            package.writestr("_rels/.rels", "<Relationships/>")
            package.writestr("docProps/core.xml", '<core xmlns:dc="http://purl.org/dc/elements/1.1/">'
                              f"<dc:title>{escape(title)}</dc:title></core>")
            package.writestr("word/document.xml", f'<w:document xmlns:w="{company.NS["w"]}"><w:body>'
                              + "".join(f"<w:p><w:r><w:t>{escape(p)}</w:t></w:r></w:p>" for p in paragraphs)
                              + "</w:body></w:document>")


class FakeCaptureEngine:
    """Stand-in for playwright_capture.CaptureEngine: same interface/identity
    contract, no real browser. Used only where a focused unit test needs the
    Runner's render path exercised without the optional Playwright dependency."""
    identity = CAPTURE_TOOL

    def capture_docx_pages(self, html_path, out_dir, expected_pages, viewport=None):
        return self._pages(out_dir, expected_pages, "pages", "page")

    def capture_pptx_slides(self, html_path, out_dir, expected_slides, viewport=None):
        return self._pages(out_dir, expected_slides, "slides", "slide")

    def capture_xlsx_region(self, html_path, out_png, sheet_name, start_cell, end_cell, sheet_index=None, viewport=None):
        out_png = Path(out_png)
        out_png.parent.mkdir(parents=True, exist_ok=True)
        out_png.write_bytes(PNG)
        return {"sheet": sheet_name, "start_cell": start_cell, "end_cell": end_cell,
                "path": str(out_png), "sha256": company.file_hash(out_png)}

    @staticmethod
    def _pages(out_dir, count, list_key, unit_key):
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        entries = []
        for number in range(1, count + 1):
            png = out_dir / f"{unit_key}-{number}.png"
            png.write_bytes(PNG)
            entry = {unit_key: number, "path": str(png), "sha256": company.file_hash(png)}
            if unit_key == "page":
                entry["overflow"] = False
            entries.append(entry)
        return {f"rendered_{list_key}": count, list_key: entries}

    def close(self):
        pass


class RunnerFlowTests(unittest.TestCase):
    def test_checked_run_assemble_audit_resume_and_output_protection(self):
        fixture = baseline.CompanyTests()
        fixture.setUp()
        fixture.canon = company.seed_canon(dict(fixture.request, documents=1,
            document_types=["press_release"], formats=["docx"]))
        fixture.plan = company.create_plan(fixture.canon)
        doc = fixture.plan["documents"][0]
        draft = fixture.draft(doc)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            build = root / "build"
            company.save(root / "canon.json", fixture.canon)
            company.save(build / "plan.json", fixture.plan)
            company.save(build / "drafts" / f"{doc['id']}.json", draft)
            layout = build / "layout.json"
            company.save(layout, {"font": "Arial", "revision": 1})
            execution = company.execution_plan(fixture.canon, fixture.plan, [
                {"document_id": doc["id"], "layout_revision": "v1", "layout_sha256": company.file_hash(layout)}],
                [doc["id"]])
            company.save(build / "execution-plan.json", execution)
            capability(root, {"docx"})
            company.save(build / "depth-targets.json", {
                "plan_sha256": company.fingerprint(fixture.plan), "documents": [
                    {"document_id": doc["id"], "content_depth": "medium", "measure": "delivered_words",
                     "target_min": 1, "target_max": 10000}]})
            commands = [{"command": "set", "path": "/", "props": {"title": doc["title"]}},
                        {"command": "add", "parent": "/body", "type": "paragraph", "props": {"text": doc["title"]}}]
            commands += [{"command": "add", "parent": "/body", "type": "paragraph", "props": {"text": s["text"]}}
                         for s in draft["sections"]]
            company.save(build / "batch.json", commands)
            value = company.display_value(draft["claims"][0])
            job = {"document_id": doc["id"], "layout_file": str(Path("build") / "layout.json"),
                   "batch_file": str(Path("build") / "batch.json"), "render_targets": [{"id": "all", "pages": 1}],
                   "occurrences": [
                       {"claim_id": "c1", "section_index": i, "start": 0, "end": len(value), "expected_text": value,
                        "locator": {"kind": "paragraph", "part": "word/document.xml", "index": i + 1}}
                       for i in range(len(draft["sections"]))]}
            cli = FakeCLI()
            with patch.object(runner.playwright_capture, "CaptureEngine", return_value=FakeCaptureEngine()):
                engine = runner.Runner(root, cli)
                machine = engine.run_job(job)
                count = len(cli.calls)
                runner.Runner(root, cli).run_job(job)
                self.assertEqual(len(cli.calls), count)
                selection = engine.assemble([job])
            self.assertFalse((build / "qa.json").exists())
            review = {"document_id": doc["id"], "artifact_sha256": machine["artifact_sha256"],
                      "draft_sha256": machine["draft_sha256"],
                      "machine_sha256": selection["documents"][0]["machine_sha256"],
                      "condition_review": {"status": "not_selected"}}
            for gate in ("semantic_review", "genre_review", "depth_review", "visual_review"):
                review[gate] = {"status": "passed", "evidence": "Mock review for test fixture only"}
            review["depth_review"]["content_depth"] = "medium"
            review["visual_review"].update(images_delivered=True, render_ids=["all"])
            qa = dict(company.receipt_header(fixture.canon, fixture.plan), execution_version=1,
                      qa_policy=execution["qa_policy"], execution_sha256=company.fingerprint(execution),
                      review_plan_sha256=company.fingerprint(selection), documents=[review])
            company.save(build / "qa.json", qa)
            result = company.audit(fixture.canon, fixture.plan, build / "drafts",
                                   company.load(build / "manifest.json"), root)
            self.assertEqual(result["documents"], 1)
            with self.assertRaisesRegex(company.Invalid, "Complete real job"):
                engine.assemble([])
            runner.write_json(build / "batch.json", commands + [commands[-1]])
            with self.assertRaisesRegex(company.Invalid, "changed inputs"):
                runner.Runner(root, cli).run_job(job)
            runner.write_json(build / "batch.json", commands)
            output = root / "documents" / doc["filename"]
            output.write_bytes(b"external edit")
            with self.assertRaisesRegex(company.Invalid, "refusing overwrite"):
                runner.Runner(root, cli).run_job(job)


if __name__ == "__main__":
    unittest.main()
