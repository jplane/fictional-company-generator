import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path

from test_execution import company, runner

FIXTURES = Path(__file__).resolve().parent.parent / "build" / "fixtures"


def _playwright_available():
    try:
        import playwright  # noqa: F401
    except ImportError:
        return False
    return True


@unittest.skipUnless(os.environ.get("RUN_OFFICECLI_TESTS") == "1",
                     "Set RUN_OFFICECLI_TESTS=1 for actual OfficeCLI-only capability fixtures")
class OfficeCLIIntegrationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if not shutil.which("officecli"):
            raise unittest.SkipTest("OfficeCLI executable unavailable; this is not a passed preflight")
        cls.cli = runner.OfficeCLI(timeout=120)

    def probe(self, fmt):
        with tempfile.TemporaryDirectory(prefix="officecli-maintenance-") as directory:
            root = Path(directory)
            out = root / "build" / "preflight.json"
            try:
                receipt = runner.preflight(self.cli, [fmt], out)
            except company.Invalid as exc:
                if "capture adapter unavailable" in str(exc) or "Chromium" in str(exc):
                    self.assertNotEqual(company.load(out)["status"], "passed")
                    self.skipTest(f"Required Playwright capture capability unavailable; not a passed preflight: {exc}")
                raise
            company.validate_preflight(receipt, root)
            self.assertEqual(receipt["status"], "passed")
            self.assertEqual(receipt["formats"], [fmt])
            self.assertEqual(receipt["capture_tool"]["engine"], "playwright-chromium")
            for probe in receipt["probes"]:
                operations = [log["operation"] for log in probe["logs"]]
                self.assertIn("html_export", operations)
                self.assertNotIn("screenshot", operations)
                for result in probe["logs"]:
                    self.assertNotIn("screenshot", result["arguments"])
                    self.assertNotIn("--render", result["arguments"])
                self.assertEqual(probe["capture"]["status"], "passed")
                self.assertEqual(probe["capture"]["tool"]["engine"], "playwright-chromium")
            self.assertTrue((out.parent / receipt["probes"][0]["image"]).is_file())

    def test_docx_html_export_extract_and_playwright_capture(self):
        self.probe("docx")

    def test_pptx_html_export_extract_and_playwright_capture(self):
        self.probe("pptx")

    def test_xlsx_html_export_and_ordered_cross_sheet_formula_caches(self):
        self.probe("xlsx")

    def test_saved_formula_changes_stale_caches_and_unsupported_functions_block(self):
        with tempfile.TemporaryDirectory(prefix="officecli-formulas-") as directory:
            path = Path(directory) / "probe.xlsx"
            runner.checked(self.cli.run(["create", str(path)], "create"))
            runner.checked(self.cli.run(["batch", str(path), "--commands",
                json.dumps(runner.fixture_commands("xlsx")), "--stop-on-error"], "author"))
            bindings = [
                {"sheet": "Sheet1", "cell": "A2", "formula": "A1*2", "value": "-4.296"},
                {"sheet": "Results", "cell": "A1", "formula": "Sheet1!A2*2", "value": "-8.592"}]
            runner.formula_values(path, bindings)
            with self.assertRaisesRegex(company.Invalid, "coverage"):
                runner.formula_values(path, [])
            wrong = [dict(bindings[0], value="-4.296001"), bindings[1]]
            with self.assertRaisesRegex(company.Invalid, "stale/incorrect"):
                runner.formula_values(path, wrong)
            runner.checked(self.cli.run(["batch", str(path), "--commands", json.dumps([
                {"command": "set", "path": "/Sheet1/A1", "props": {"value": "3.25"}}]),
                "--stop-on-error"], "author"))
            # A sweep may drop or refresh caches; old expected business values must fail either way.
            with self.assertRaises(company.Invalid):
                runner.formula_values(path, bindings)
            runner.checked(self.cli.run(["batch", str(path), "--commands", json.dumps([
                {"command": "set", "path": "/Sheet1/A2", "props": {"formula": "A1*2"}},
                {"command": "set", "path": "/Results/A1", "props": {"formula": "Sheet1!A2*2"}}]),
                "--stop-on-error"], "formula_correction"))
            runner.formula_values(path, [dict(bindings[0], value="6.5"), dict(bindings[1], value="13")])
            runner.checked(self.cli.run(["batch", str(path), "--commands", json.dumps([
                {"command": "set", "path": "/Sheet1/A2",
                 "props": {"formula": "UNSUPPORTED_CAPABILITY_PROBE(A1)"}}]),
                "--stop-on-error"], "author"))
            with self.assertRaises(company.Invalid):
                runner.formula_values(path, [
                    dict(bindings[0], formula="UNSUPPORTED_CAPABILITY_PROBE(A1)", value="6.5"),
                    dict(bindings[1], value="13")])


@unittest.skipUnless(os.environ.get("RUN_OFFICECLI_TESTS") == "1",
                     "Set RUN_OFFICECLI_TESTS=1 for actual OfficeCLI+Playwright render fixtures")
@unittest.skipUnless(_playwright_available(),
                     "Playwright not installed in this venv; not a passed render proof")
class PlaywrightCaptureIntegrationTests(unittest.TestCase):
    """Exercises the real production capture_target()/CaptureEngine path (not
    unit mocks) against prior-art-grounded fixtures copied into this repo's
    owned build/fixtures/ location. Requires the pinned Playwright + managed
    Chromium build to actually be importable/launchable in this venv."""

    @classmethod
    def setUpClass(cls):
        if not shutil.which("officecli"):
            raise unittest.SkipTest("OfficeCLI executable unavailable; this is not a passed preflight")
        for name in ("docx1_pagebreaks.docx", "docx2_naturalflow.docx",
                     "pptx1_4slides_corrected.pptx", "xlsx1_regions.xlsx"):
            if not (FIXTURES / name).is_file():
                raise unittest.SkipTest(f"Missing owned fixture: {name}")
        cls.cli = runner.OfficeCLI(timeout=120)
        import playwright_capture
        try:
            cls.engine = playwright_capture.CaptureEngine()
        except playwright_capture.CaptureError as exc:
            raise unittest.SkipTest(f"Managed Chromium unavailable; not a passed render proof: {exc}")

    @classmethod
    def tearDownClass(cls):
        cls.engine.close()

    def _export_html(self, fixture_name, tmp_dir):
        src = FIXTURES / fixture_name
        html = Path(tmp_dir) / (src.stem + ".html")
        runner.checked(self.cli.run(runner.html_export_arguments(src, html), "html_export"))
        return html

    def test_docx_explicit_pagebreaks_accepted(self):
        with tempfile.TemporaryDirectory(prefix="docx-pagebreaks-") as tmp:
            html = self._export_html("docx1_pagebreaks.docx", tmp)
            result = self.engine.capture_docx_pages(html, Path(tmp) / "pages", expected_pages=3)
            self.assertEqual(result["rendered_pages"], 3)
            self.assertEqual([p["overflow"] for p in result["pages"]], [False, False, False])
            for page in result["pages"]:
                self.assertTrue(Path(page["path"]).is_file())

    def test_docx_natural_flow_overflow_rejected(self):
        import playwright_capture
        with tempfile.TemporaryDirectory(prefix="docx-naturalflow-") as tmp:
            html = self._export_html("docx2_naturalflow.docx", tmp)
            # Whatever explicit page count a caller might assume for this
            # known-bad fixture, the natural-flow layout must never be
            # silently accepted as a passed multi-page capture.
            with self.assertRaises(playwright_capture.CaptureError):
                self.engine.capture_docx_pages(html, Path(tmp) / "pages", expected_pages=1)

    def test_pptx_corrected_title_body_fixture_accepted(self):
        with tempfile.TemporaryDirectory(prefix="pptx-corrected-") as tmp:
            html = self._export_html("pptx1_4slides_corrected.pptx", tmp)
            result = self.engine.capture_pptx_slides(html, Path(tmp) / "slides", expected_slides=4)
            self.assertEqual(result["rendered_slides"], 4)
            self.assertEqual(len(result["slides"]), 4)
            for slide in result["slides"]:
                self.assertTrue(Path(slide["path"]).is_file())

    def test_xlsx_region_merge_offset_wide_and_second_sheet(self):
        with tempfile.TemporaryDirectory(prefix="xlsx-regions-") as tmp:
            html = self._export_html("xlsx1_regions.xlsx", tmp)
            cases = [
                ("basic", "Sheet1", "A1", "C4", None),
                ("negative_fraction", "Sheet1", "E8", "H13", None),
                ("wide_14_cols", "Sheet1", "A18", "N23", None),
                ("merged_wrapped", "Sheet1", "A28", "F34", None),
                ("second_sheet", "Operations 2026", "B3", "E8", 1),
            ]
            for name, sheet, start, end, sheet_index in cases:
                out_png = Path(tmp) / f"{name}.png"
                region = self.engine.capture_xlsx_region(
                    html, out_png, sheet, start, end, sheet_index=sheet_index)
                self.assertTrue(out_png.is_file(), f"{name}: missing region image")
                self.assertGreater(region["clip"]["width"], 0, f"{name}: zero-width region")
                self.assertGreater(region["clip"]["height"], 0, f"{name}: zero-height region")


if __name__ == "__main__":
    unittest.main()
