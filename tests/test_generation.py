import copy
import json
import os
import subprocess
import sys
import tempfile
import threading
import unittest
from pathlib import Path
from unittest.mock import patch

import test_execution as execution

company = execution.company
runner = execution.runner


class GenerationFixture:
    def __init__(self, root, count=6, policy=None):
        self.root = Path(root)
        self.build = self.root / "build"
        baseline = execution.baseline.CompanyTests()
        baseline.setUp()
        baseline.canon = company.seed_canon(dict(baseline.request, documents=count,
            document_types=["press_release"], formats=["docx"]))
        baseline.plan = company.create_plan(baseline.canon)
        self.canon, self.plan = baseline.canon, baseline.plan
        company.save(self.root / "canon.json", self.canon)
        company.save(self.build / "plan.json", self.plan)
        company.save(self.build / "layout.json", {"font": "Arial", "revision": 1})
        layouts = [{"document_id": d["id"], "layout_revision": "v1",
                    "layout_sha256": company.file_hash(self.build / "layout.json")}
                   for d in self.plan["documents"]]
        self.pilot = self.plan["documents"][0]["id"]
        self.execution = company.generation_execution_plan(
            self.canon, self.plan, layouts, [self.pilot],
            {"enabled": policy is not None, "policy": policy})
        company.save(self.build / "execution-plan.json", self.execution)
        execution.capability(self.root, {"docx"})
        company.save(self.build / "depth-targets.json", {
            "plan_sha256": company.fingerprint(self.plan), "documents": [
                {"document_id": d["id"], "content_depth": "medium", "measure": "delivered_words",
                 "target_min": 1, "target_max": 10000} for d in self.plan["documents"]]})
        self.jobs = []
        for doc in self.plan["documents"]:
            draft = baseline.draft(doc)
            company.save(self.build / "drafts" / f"{doc['id']}.json", draft)
            commands = [
                {"command": "set", "path": "/", "props": {"title": doc["title"]}},
                {"command": "add", "parent": "/body", "type": "paragraph", "props": {"text": doc["title"]}}]
            commands += [{"command": "add", "parent": "/body", "type": "paragraph",
                          "props": {"text": s["text"]}} for s in draft["sections"]]
            batch_name = f"{doc['id']}-batch.json"
            company.save(self.build / batch_name, commands)
            self.jobs.append({"document_id": doc["id"], "layout_file": "build\\layout.json",
                "batch_file": str(Path("build") / batch_name),
                "render_targets": [{"id": "all", "pages": 1}],
                "export_bindings": [{"claim_id": ref, "section_index": i,
                    "locator": {"kind": "paragraph", "part": "word/document.xml", "index": i + 1}}
                    for i, section in enumerate(draft["sections"]) for ref in section["claim_ids"]]})
        self.cli = execution.FakeCLI()

    def engine(self):
        return runner.GenerationRunner(self.root, self.cli)

    def review(self, identifier, directory="machine"):
        machine = company.load(self.build / directory / f"{identifier}.json")
        review = {"document_id": identifier, "artifact_sha256": machine["artifact_sha256"],
                  "draft_sha256": machine["draft_sha256"],
                  "machine_sha256": company.file_hash(self.build / directory / f"{identifier}.json"),
                  "generation_machine_sha256": company.file_hash(self.build / "machine" / f"{identifier}.json"),
                  "condition_review": {"status": "not_selected"}}
        for gate in ("semantic_review", "genre_review", "depth_review", "visual_review"):
            review[gate] = {"status": "passed", "evidence": "Mock review evidence for unit tests only"}
        review["depth_review"]["content_depth"] = "medium"
        review["visual_review"].update(images_delivered=True, render_ids=machine["render_targets"])
        runner.write_json(self.build / "reviews" / f"{identifier}.json", review)
        return review

    def generate(self):
        engine = self.engine()
        try:
            self.assert_passed(engine.run_jobs(self.jobs[:1]))
            self.review(self.pilot)
            self.assert_passed(engine.run_jobs(self.jobs))
            return engine.assemble(self.jobs)
        finally:
            engine.close()

    @staticmethod
    def assert_passed(result):
        if result["status"] != "passed":
            raise AssertionError(result)

    def review_qa_work(self):
        work = company.load(self.build / "review-work.json")
        for review in work["documents"]:
            machine = company.load(self.build / "qa-machine" / f"{review['document_id']}.json")
            for gate in ("semantic_review", "genre_review", "depth_review", "visual_review", "condition_review"):
                if review[gate]["status"] == "pending":
                    review[gate] = {"status": "passed", "evidence": "Mocked actual review for test fixture only"}
                    if gate == "depth_review":
                        review[gate]["content_depth"] = "medium"
                    if gate == "visual_review":
                        review[gate].update(images_delivered=True, render_ids=machine["render_targets"])
            runner.write_json(self.build / "reviews" / f"{review['document_id']}.json", review)


class GenerationTests(unittest.TestCase):
    def setUp(self):
        self.capture = patch.object(runner.playwright_capture, "CaptureEngine", execution.FakeCaptureEngine)
        self.capture.start()
        self.addCleanup(self.capture.stop)

    def test_generation_only_then_later_qa_all_policies_and_unchanged_bytes(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory)
            result = fixture.generate()
            self.assertEqual(result["qa_status"], "not_requested")
            self.assertFalse((fixture.build / "qa.json").exists())
            self.assertFalse((fixture.build / "provenance.json").exists())
            self.assertFalse((fixture.build / "review-plan.json").exists())
            self.assertEqual(sum(op == "html_export" for op, _ in fixture.cli.calls), 1)
            before = {p.name: company.file_hash(p) for p in (fixture.root / "documents").iterdir()}
            generation_machines = {p.name: company.file_hash(p) for p in (fixture.build / "machine").iterdir()}
            with self.assertRaisesRegex(company.Invalid, "Additional QA incomplete"):
                company.audit(fixture.canon, fixture.plan, fixture.build / "drafts",
                              company.load(fixture.build / "manifest.json"), fixture.root)
            for policy in company.QA_POLICIES:
                engine = fixture.engine()
                try:
                    prepared = engine.prepare_qa(fixture.jobs, policy)
                    self.assertEqual(prepared["qa_status"], "awaiting_reviews")
                    fixture.review_qa_work()
                    audited = engine.finalize_qa()
                    self.assertEqual(audited["qa_status"], "complete")
                    self.assertEqual(audited["qa_policy"], policy)
                finally:
                    engine.close()
            self.assertEqual(before, {p.name: company.file_hash(p) for p in (fixture.root / "documents").iterdir()})
            self.assertEqual(generation_machines,
                             {p.name: company.file_hash(p) for p in (fixture.build / "machine").iterdir()})
            self.assertTrue((fixture.build / "provenance.json").exists())
            self.assertEqual(sum(op == "author" for op, _ in fixture.cli.calls), len(fixture.jobs))

    def test_intake_contract_requires_explicit_choice_and_no_version_downgrade(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=1)
            for choice in (None, {}, {"enabled": False, "policy": "Full"},
                           {"enabled": True, "policy": None}, {"enabled": 1, "policy": "Full"}):
                changed = copy.deepcopy(fixture.execution)
                changed["post_generation_qa"] = choice
                with self.assertRaisesRegex(company.Invalid, "Explicit post_generation_qa"):
                    company.validate_execution(fixture.canon, fixture.plan, changed)
            changed = copy.deepcopy(fixture.execution)
            changed.pop("execution_version")
            with self.assertRaisesRegex(company.Invalid, "Unsupported execution"):
                company.validate_execution(fixture.canon, fixture.plan, changed)
            changed = copy.deepcopy(fixture.execution)
            changed["post_generation_qa"] = {"enabled": True, "policy": "Full"}
            self.assertEqual(company.generation_identity(changed), company.generation_identity(fixture.execution))

    def test_pilots_gate_siblings_and_hash_bound_resume(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=4)
            engine = fixture.engine()
            try:
                result = engine.run_jobs(fixture.jobs)
                self.assertEqual(result["status"], "awaiting_reviews")
                self.assertEqual(len(list((fixture.root / "documents").iterdir())), 1)
                fixture.review(fixture.pilot)
                self.assertEqual(engine.run_jobs(fixture.jobs)["status"], "passed")
                calls = len(fixture.cli.calls)
                self.assertEqual(engine.run_jobs(fixture.jobs)["status"], "passed")
                self.assertEqual(len(fixture.cli.calls), calls)
                engine.assemble(fixture.jobs)
                job = fixture.jobs[-1]
                changed = copy.deepcopy(job)
                changed["export_bindings"][0]["locator"]["index"] = 0
                with self.assertRaisesRegex(company.Invalid, "association changed"):
                    engine.run_jobs([changed])
                output = fixture.root / "documents" / fixture.plan["documents"][-1]["filename"]
                output.write_bytes(b"tampered")
                with self.assertRaisesRegex(company.Invalid, "refusing overwrite"):
                    engine.run_jobs([job])
            finally:
                engine.close()

    def test_explicit_binding_offsets_and_ambiguity(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=1)
            fixture.generate()
            job = fixture.jobs[0]
            draft = company.load(fixture.build / "drafts" / f"{fixture.pilot}.json")
            artifact = fixture.root / "documents" / fixture.plan["documents"][0]["filename"]
            bindings = job["export_bindings"]
            resolved = company.resolve_export_bindings(artifact, "docx", draft, bindings)
            self.assertEqual(len(resolved), len(bindings))
            self.assertTrue(all(r["expected_text"] for r in resolved))
            containers = company.artifact_containers(artifact, "docx")
            text = containers[1]["text"]
            fake = copy.deepcopy(containers)
            fake[1]["text"] = text + " " + text
            with patch.object(company, "artifact_containers", return_value=fake):
                with self.assertRaisesRegex(company.Invalid, "ambiguous"):
                    company.resolve_export_bindings(artifact, "docx", draft, bindings)
                explicit = copy.deepcopy(bindings)
                explicit[0]["occurrence_index"] = 1
                result = company.resolve_export_bindings(artifact, "docx", draft, explicit)
                self.assertGreater(result[0]["start"], resolved[0]["start"])

    def test_cache_reuses_artifact_parsing_but_new_audit_detects_tampering(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=1)
            fixture.generate()
            path = fixture.root / "documents" / fixture.plan["documents"][0]["filename"]
            with company.validation_session() as cache:
                first = company.artifact_containers(path, "docx")
                second = company.artifact_containers(path, "docx")
                self.assertIs(first, second)
                self.assertEqual(cache["counts"]["artifact_computations"], 1)
                self.assertEqual(cache["counts"]["artifact_hits"], 1)
            path.write_bytes(b"tampered")
            with self.assertRaises(company.Invalid):
                company.audit_generation(fixture.canon, fixture.plan, fixture.build / "drafts",
                                         company.load(fixture.build / "manifest.json"), fixture.root)

    def test_run_lock_worker_caps_and_phase_summary(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=1)
            with runner.run_lock(directory):
                with self.assertRaisesRegex(company.Invalid, "owned by another"):
                    with runner.run_lock(directory):
                        self.fail("second lock accepted")
            parent = runner.phase_event(directory, "start", "drafting")
            child = runner.phase_event(directory, "start", "draft-document", parent=parent["id"])
            with self.assertRaisesRegex(company.Invalid, "child phases"):
                runner.phase_event(directory, "end", event_id=parent["id"])
            runner.phase_event(directory, "end", event_id=child["id"])
            runner.phase_event(directory, "end", event_id=parent["id"])
            summary = runner.timing_summary(directory)
            self.assertEqual(len(summary["leaf_agent_spans"]), 1)
            self.assertFalse(summary["unknown_agent_phases"])
            with patch.object(os, "cpu_count", return_value=8):
                self.assertEqual(runner.GenerationRunner.worker_count(range(100)), 4)
                self.assertEqual(runner.GenerationRunner.worker_count(range(100), 1), 1)
                self.assertEqual(runner.GenerationRunner.worker_count(range(2)), 2)
            for invalid in (0, 5, True):
                with self.assertRaises(company.Invalid):
                    runner.GenerationRunner.worker_count(range(100), invalid)

    def test_hundred_job_batch_overlaps_and_avoids_bulk_capture(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=100)
            engine = fixture.engine()
            try:
                engine.run_jobs(fixture.jobs[:1])
                fixture.review(fixture.pilot)
                original = fixture.cli.run
                barrier = threading.Barrier(4, timeout=20)
                lock = threading.Lock()
                first_four = []
                def overlapping(args, operation):
                    if operation == "create":
                        with lock:
                            first_four.append(args[1])
                            participate = len(first_four) <= 4
                        if participate:
                            barrier.wait()
                    return original(args, operation)
                fixture.cli.run = overlapping
                with patch.object(os, "cpu_count", return_value=8):
                    result = engine.run_jobs(fixture.jobs[1:])
                self.assertEqual(result["workers"], 4)
                self.assertEqual(result["status"], "passed")
                self.assertEqual(len(first_four), 99)
                self.assertEqual(sum(op == "html_export" for op, _ in fixture.cli.calls), 1)
                self.assertEqual(sum(op == "author" for op, _ in fixture.cli.calls), 100)
                self.assertEqual(len(list((fixture.build / "document-state").glob("*.json"))), 100)
                engine.assemble(fixture.jobs)
                summary = runner.timing_summary(directory)
                self.assertGreater(summary["validation_counts"]["artifact_hits"], 0)
                self.assertEqual(summary["milestones"]["generation"]["documents"], 100)
            finally:
                engine.close()

    def test_cli_phase_and_summary_do_not_require_officecli(self):
        with tempfile.TemporaryDirectory() as directory:
            GenerationFixture(directory, count=1)
            script = str(execution.baseline.SCRIPT.with_name("officecli_runner.py"))
            result = subprocess.run([sys.executable, script, "phase", "--root", directory,
                "--action", "start", "--name", "canon"], capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            event = json.loads(result.stdout)
            result = subprocess.run([sys.executable, script, "phase", "--root", directory,
                "--action", "end", "--id", event["id"]], capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            result = subprocess.run([sys.executable, script, "summary", "--root", directory],
                                    capture_output=True, text=True)
            self.assertEqual(result.returncode, 0, result.stderr)
            self.assertFalse(json.loads(result.stdout)["unknown_agent_phases"])

    def test_failed_checks_require_a_concrete_retry_and_expand_generation_family(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=3)
            original = fixture.cli.run
            failures = []
            def fail_once(args, operation):
                if operation == "issues" and not failures:
                    failures.append(operation)
                    return dict(execution.log(operation, args), returncode=1, stderr="fixture failure")
                return original(args, operation)
            fixture.cli.run = fail_once
            engine = fixture.engine()
            try:
                with self.assertRaisesRegex(company.Invalid, "fixture failure"):
                    engine.run_jobs(fixture.jobs[:1])
            finally:
                engine.close()
            engine = fixture.engine()
            try:
                with self.assertRaisesRegex(company.Invalid, "concrete correction"):
                    engine.run_jobs(fixture.jobs[:1])
            finally:
                engine.close()
            fixture.jobs[0]["retry_reasons"] = {"issues": "Corrected the fixture's external issue"}
            engine = fixture.engine()
            try:
                result = engine.run_jobs(fixture.jobs[:1])
                self.assertEqual(result["status"], "awaiting_reviews")
                fixture.review(fixture.pilot)
                engine.run_jobs(fixture.jobs)
                for doc in fixture.plan["documents"]:
                    fixture.review(doc["id"])
                self.assertEqual(engine.assemble(fixture.jobs)["documents"], 3)
                for doc in fixture.plan["documents"]:
                    machine = company.load(fixture.build / "machine" / f"{doc['id']}.json")
                    self.assertTrue(machine["renders"])
            finally:
                engine.close()

    def test_qa_capture_retry_does_not_modify_generation_and_expands_qa_only(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=4)
            fixture.generate()
            before = {p.name: company.file_hash(p) for p in (fixture.build / "machine").iterdir()}
            original = execution.FakeCaptureEngine.capture_docx_pages
            failed = []
            def fail_once(engine, *args, **kwargs):
                if not failed:
                    failed.append(True)
                    raise runner.playwright_capture.CaptureError("fixture capture failure")
                return original(engine, *args, **kwargs)
            with patch.object(execution.FakeCaptureEngine, "capture_docx_pages", fail_once):
                engine = fixture.engine()
                try:
                    # Full requires a previously uncaptured sibling.
                    with self.assertRaisesRegex(company.Invalid, "fixture capture failure"):
                        engine.prepare_qa(fixture.jobs, "Full", workers=1)
                finally:
                    engine.close()
            for job in fixture.jobs:
                job["retry_reasons"] = {"render-all": "Repaired fixture browser capture"}
            engine = fixture.engine()
            try:
                engine.prepare_qa(fixture.jobs, workers=1)
                self.assertTrue(engine.execution["qa_flagged_ids"])
                fixture.review_qa_work()
                self.assertEqual(engine.finalize_qa()["qa_status"], "complete")
            finally:
                engine.close()
            self.assertEqual(before, {p.name: company.file_hash(p) for p in (fixture.build / "machine").iterdir()})
            result = company.audit_generation(fixture.canon, fixture.plan, fixture.build / "drafts",
                company.load(fixture.build / "manifest.json"), fixture.root)
            self.assertEqual(result["pilot_and_corrective_reviews"], 1)

    def test_budget_and_stale_snapshot_stop_before_authoring(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=1)
            changed = copy.deepcopy(fixture.execution)
            changed["budget_seconds"] = 1
            runner.write_json(fixture.build / "execution-plan.json", changed)
            engine = fixture.engine()
            engine.started -= 2
            try:
                with self.assertRaisesRegex(company.Invalid, "budget reached"):
                    engine.run_jobs(fixture.jobs)
                self.assertFalse(fixture.cli.calls)
            finally:
                engine.close()
            first = fixture.engine()
            second = fixture.engine()
            try:
                first.checkpoint()
                with self.assertRaisesRegex(company.Invalid, "checkpoint changed"):
                    second.run_jobs(fixture.jobs)
            finally:
                first.close()
                second.close()

    def test_skill_requires_explicit_optional_qa_and_timing(self):
        wording = (execution.baseline.SCRIPT.parents[1] / "SKILL.md").read_text(encoding="utf-8")
        for expected in ("Additional QA:", "No additional QA", "during intake", "Sampled",
                         "FullContentSampledVisual", "Full", "generation_execution_plan",
                         "finalize-qa", "--workers 1", "phase --action start", "unknown, not zero"):
            self.assertIn(expected, wording)

    def test_recorded_review_failure_is_preserved_and_expands_the_entire_qa_family(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=4)
            fixture.generate()
            engine = fixture.engine()
            try:
                engine.prepare_qa(fixture.jobs, "Sampled")
                work = company.load(fixture.build / "review-work.json")
                failed = work["documents"][0]
                failed["visual_review"] = {"status": "failed", "evidence": "Fixture text is not legible"}
                runner.write_json(fixture.build / "reviews" / f"{failed['document_id']}.json", failed)
                with self.assertRaisesRegex(company.Invalid, "Review failures recorded"):
                    engine.finalize_qa()
                self.assertEqual(len(list((fixture.build / "qa-findings").glob("*.json"))), 1)
                engine.prepare_qa(fixture.jobs)
                plan = company.load(fixture.build / "review-plan.json")
                self.assertTrue(all(d["content_required"] and d["visual_required"] for d in plan["documents"]))
                fixture.review_qa_work()
                self.assertEqual(engine.finalize_qa()["qa_status"], "complete")
                self.assertEqual(len(list((fixture.build / "qa-findings").glob("*.json"))), 1)
            finally:
                engine.close()

    def test_content_review_is_reused_when_only_visual_scope_increases(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=4)
            fixture.generate()
            engine = fixture.engine()
            try:
                engine.prepare_qa(fixture.jobs, "FullContentSampledVisual")
                fixture.review_qa_work()
                engine.finalize_qa()
                engine.prepare_qa(fixture.jobs, "Full")
                work = company.load(fixture.build / "review-work.json")
                self.assertTrue(all(d["semantic_review"]["status"] == "passed" for d in work["documents"]))
                self.assertTrue(all(d["depth_review"]["status"] == "passed" for d in work["documents"]))
                self.assertTrue(any(d["visual_review"]["status"] == "pending" for d in work["documents"]))
            finally:
                engine.close()

    def test_invalid_capture_evidence_consumes_the_persistent_retry_budget(self):
        class DriftCapture(execution.FakeCaptureEngine):
            identity = dict(execution.CAPTURE_TOOL, browser_version="unapproved-drift")
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=1)
            with patch.object(runner.playwright_capture, "CaptureEngine", DriftCapture):
                for attempt in range(3):
                    if attempt:
                        fixture.jobs[0]["retry_reasons"] = {"render-all": "Attempted to repair pinned capture identity"}
                    engine = fixture.engine()
                    try:
                        message = "identity" if attempt < 2 else "corrective retry already consumed"
                        with self.assertRaisesRegex(company.Invalid, message):
                            engine.run_jobs(fixture.jobs)
                    finally:
                        engine.close()
            state = company.load(fixture.build / "document-state" / f"{fixture.pilot}.json")
            captures = state["documents"][fixture.pilot]["captures"]
            self.assertEqual(len(captures), 2)
            self.assertTrue(all(e["status"] == "blocked" for e in captures))

    def test_browser_is_reused_and_all_browser_operations_stay_on_its_owning_thread(self):
        owners, calls, closed = [], [], []
        class TrackingCapture(execution.FakeCaptureEngine):
            def __init__(self):
                self.owner = threading.get_ident()
                owners.append(self.owner)
            def capture_docx_pages(self, *args, **kwargs):
                calls.append((self.owner, threading.get_ident()))
                return super().capture_docx_pages(*args, **kwargs)
            def close(self):
                closed.append((self.owner, threading.get_ident()))
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=5)
            with patch.object(runner.playwright_capture, "CaptureEngine", TrackingCapture):
                fixture.generate()
                engine = fixture.engine()
                try:
                    engine.prepare_qa(fixture.jobs, "Full")
                finally:
                    engine.close()
            self.assertEqual(len(owners), 2)
            self.assertEqual(len(calls), 5)
            self.assertEqual(len(closed), 2)
            self.assertTrue(all(owner == actual for owner, actual in calls + closed))

    def test_resume_recovers_interrupted_document_spans_without_duplicating_completed_spans(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=1)
            fixture.generate()
            path = fixture.build / "document-state" / f"{fixture.pilot}.json"
            state = company.load(path)
            state["phases"].append({"id": "interrupted-fixture", "phase": "document",
                                    "status": "running", "active_seconds": 1})
            runner.write_json(path, state)
            engine = fixture.engine()
            try:
                recovered = [event for event in engine.state["phases"] if event["id"] == "interrupted-fixture"]
                self.assertEqual(len(recovered), 1)
                self.assertEqual(recovered[0]["status"], "interrupted")
                self.assertEqual(len(engine.state["phases"]), len({e["id"] for e in engine.state["phases"]}))
            finally:
                engine.close()

    def test_known_incomplete_failure_holds_siblings_even_when_failed_job_is_omitted(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture = GenerationFixture(directory, count=4)
            engine = fixture.engine()
            try:
                engine.run_jobs(fixture.jobs[:1])
                fixture.review(fixture.pilot)
                original = fixture.cli.run
                def fail_issues(args, operation):
                    if operation == "issues":
                        return dict(execution.log(operation, args), returncode=1, stderr="known fixture failure")
                    return original(args, operation)
                fixture.cli.run = fail_issues
                with self.assertRaisesRegex(company.Invalid, "known fixture failure"):
                    engine.run_jobs(fixture.jobs[1:2])
            finally:
                engine.close()
            fixture.cli.run = original
            engine = fixture.engine()
            try:
                calls = len(fixture.cli.calls)
                result = engine.run_jobs(fixture.jobs[2:])
                self.assertEqual(result["status"], "awaiting_corrections")
                self.assertEqual(result["awaiting_corrections"], [fixture.jobs[1]["document_id"]])
                self.assertEqual(len(fixture.cli.calls), calls)
            finally:
                engine.close()


@unittest.skipUnless(os.environ.get("RUN_OFFICECLI_TESTS") == "1",
                     "Set RUN_OFFICECLI_TESTS=1 for actual exporter association tests")
class ExporterIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not runner.shutil.which("officecli"):
            raise unittest.SkipTest("OfficeCLI unavailable; not a passed capability probe")

    @staticmethod
    def commands(fmt, text):
        if fmt == "docx":
            return [{"command": "add", "parent": "/body", "type": "paragraph", "props": {"text": text}}]
        if fmt == "pptx":
            return [{"command": "add", "parent": "/", "type": "slide", "props": {}},
                    {"command": "add", "parent": "/slide[1]", "type": "shape",
                     "props": {"text": text, "x": "1cm", "y": "1cm", "width": "10cm", "height": "2cm"}}]
        return [{"command": "set", "path": "/Sheet1/A1", "props": {"value": text}}]

    def test_batch_target_resolution_from_actual_word_and_presentation_results(self):
        cli = runner.OfficeCLI()
        for fmt in ("docx", "pptx", "xlsx"):
            with self.subTest(format=fmt), tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / f"binding.{fmt}"
                runner.checked(cli.run(["create", str(path)], "create"))
                commands = self.commands(fmt, "Bound claim")
                author = runner.checked(cli.run(["batch", str(path), "--commands", json.dumps(commands),
                                                 "--stop-on-error"], "author"))
                draft = {"sections": [{"claim_ids": ["c1"]}], "claims": [{"id": "c1", "value": "Bound claim"}]}
                bindings = [{"claim_id": "c1", "section_index": 0, "command_index": len(commands) - 1}]
                result = company.resolve_export_bindings(path, fmt, draft, bindings, commands, author)
                self.assertEqual(result[0]["expected_text"], "Bound claim")
                # Template handoff can reuse the proven exact saved locator without hand-authored offsets.
                template = [{"claim_id": "c1", "section_index": 0, "locator": result[0]["locator"]}]
                self.assertEqual(company.resolve_export_bindings(path, fmt, draft, template), result)

    def test_actual_template_merge_preserves_bound_targets(self):
        cli = runner.OfficeCLI()
        for fmt in ("docx", "pptx", "xlsx"):
            with self.subTest(format=fmt), tempfile.TemporaryDirectory() as directory:
                source, output = Path(directory) / f"template.{fmt}", Path(directory) / f"merged.{fmt}"
                runner.checked(cli.run(["create", str(source)], "create"))
                commands = self.commands(fmt, "{{claim}}")
                author = runner.checked(cli.run(["batch", str(source), "--commands", json.dumps(commands),
                                                 "--stop-on-error"], "author"))
                draft = {"sections": [{"claim_ids": ["c1"]}], "claims": [{"id": "c1", "value": "{{claim}}"}]}
                binding = [{"claim_id": "c1", "section_index": 0, "command_index": len(commands) - 1}]
                target = company.resolve_export_bindings(source, fmt, draft, binding, commands, author)[0]["locator"]
                runner.checked(cli.run(["merge", str(source), str(output),
                                        "--data", json.dumps({"claim": "Bound claim"})], "merge"))
                draft["claims"][0]["value"] = "Bound claim"
                binding = [{"claim_id": "c1", "section_index": 0, "locator": target}]
                self.assertEqual(company.resolve_export_bindings(output, fmt, draft, binding)[0]["expected_text"],
                                 "Bound claim")

    def test_actual_standalone_officecli_supports_independent_concurrent_jobs(self):
        cli = runner.OfficeCLI()
        with tempfile.TemporaryDirectory() as directory:
            barrier = threading.Barrier(4, timeout=30)
            def author(number):
                fmt = company.FORMATS[number % 3]
                path = Path(directory) / f"concurrent-{number}.{fmt}"
                barrier.wait()
                runner.checked(cli.run(["create", str(path)], "create"))
                commands = self.commands(fmt, f"Bound claim {number}")
                runner.checked(cli.run(["batch", str(path), "--commands", json.dumps(commands),
                                        "--stop-on-error"], "author"))
                containers = company.artifact_containers(path, fmt)
                return any(c["text"] == f"Bound claim {number}" for c in containers)
            with runner.ThreadPoolExecutor(max_workers=4) as pool:
                self.assertTrue(all(pool.map(author, range(4))))


if __name__ == "__main__":
    unittest.main()
