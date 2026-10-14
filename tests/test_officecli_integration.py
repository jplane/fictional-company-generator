import json
import os
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path

from test_execution import company, runner


def _build(cli, path, commands):
    """Author a self-contained fixture with maintained OfficeCLI create/batch
    operations only; never raw OOXML authoring. Failure here is a hard error,
    not a skip, once the toolchain is confirmed available."""
    runner.checked(cli.run(["create", str(path)], "create"))
    runner.checked(cli.run(["batch", str(path), "--commands", json.dumps(commands),
                           "--stop-on-error"], "author"))
    return path


def _docx_three_pages_commands():
    pages = ["Page one of a maintained three-page capability fixture.",
             "Page two of a maintained three-page capability fixture.",
             "Page three of a maintained three-page capability fixture."]
    commands = []
    for index, text in enumerate(pages):
        if index:
            commands.append({"command": "add", "parent": "/body", "type": "pagebreak", "props": {}})
        commands.append({"command": "add", "parent": "/body", "type": "paragraph",
                         "props": {"text": text, "firstLineIndent": "0.5cm"}})
    return commands


def _docx_natural_flow_overflow_commands():
    long_text = ("Overflowing natural-flow paragraph text repeated to exercise the "
                "no-explicit-page-break rejection path. ") * 150
    return [{"command": "add", "parent": "/body", "type": "paragraph",
            "props": {"text": long_text, "firstLineIndent": "0.5cm"}}]


def _pptx_four_slides_commands():
    long_title = "A deliberately long corrected slide title spanning the full placeholder width"
    commands = []
    for number in range(1, 5):
        commands.append({"command": "add", "parent": "/", "type": "slide",
                         "props": {"title": f"{long_title} ({number})",
                                   "text": f"Corrected body content for slide {number}."}})
    commands.append({"command": "add", "parent": "/slide[4]", "type": "table",
                     "props": {"rows": "2", "cols": "2", "data": "H1,H2;R1,R2",
                               "x": "1in", "y": "3in"}})
    return commands


def _xlsx_five_regions_commands():
    return [
        {"command": "set", "path": "/Sheet1/A1", "props": {"value": "Basic"}},
        {"command": "set", "path": "/Sheet1/B2", "props": {"value": "10"}},
        {"command": "set", "path": "/Sheet1/C4", "props": {"value": "20"}},
        {"command": "set", "path": "/Sheet1/E8", "props": {"value": "-1.5"}},
        {"command": "set", "path": "/Sheet1/H13", "props": {"value": "-2.25"}},
        {"command": "set", "path": "/Sheet1/A18", "props": {"value": "Wide"}},
        {"command": "set", "path": "/Sheet1/N23", "props": {"value": "End"}},
        {"command": "set", "path": "/Sheet1/A28", "props": {"value": "Region start"}},
        {"command": "set", "path": "/Sheet1/B30:D31", "props": {"merge": "true"}},
        {"command": "set", "path": "/Sheet1/B30", "props": {"value": "Merged and wrapped text inside region"}},
        {"command": "set", "path": "/Sheet1/F34", "props": {"value": "Region end"}},
        {"command": "add", "parent": "/", "type": "sheet", "props": {"name": "Operations 2026"}},
        {"command": "set", "path": "/Operations 2026/B3", "props": {"value": "Ops"}},
        {"command": "set", "path": "/Operations 2026/E8", "props": {"value": "99"}},
    ]


def _xlsx_chart_matrix_commands():
    """Column + line charts covering: standard positive cases, duplicate
    titles with distinct source data, a chart anchored beyond the emitted
    cell grid, an offscreen/scroll-required chart, and a second sheet whose
    name contains spaces."""
    return [
        {"command": "set", "path": "/Sheet1/A1", "props": {"value": "Quarter"}},
        {"command": "set", "path": "/Sheet1/B1", "props": {"value": "Revenue"}},
        {"command": "set", "path": "/Sheet1/A2", "props": {"value": "Q1"}},
        {"command": "set", "path": "/Sheet1/B2", "props": {"value": "100"}},
        {"command": "set", "path": "/Sheet1/A3", "props": {"value": "Q2"}},
        {"command": "set", "path": "/Sheet1/B3", "props": {"value": "150"}},
        {"command": "add", "parent": "/Sheet1", "type": "chart",
         "props": {"dataRange": "Sheet1!B2:B3", "categories": "Sheet1!A2:A3",
                   "chartType": "column", "anchor": "D2:J18", "title": "Quarterly Summary"}},
        {"command": "add", "parent": "/Sheet1", "type": "chart",
         "props": {"dataRange": "Sheet1!B2:B3", "categories": "Sheet1!A2:A3",
                   "chartType": "line", "anchor": "D20:J36", "title": "Quarterly Summary"}},
        {"command": "add", "parent": "/Sheet1", "type": "chart",
         "props": {"dataRange": "Sheet1!B2:B3", "categories": "Sheet1!A2:A3",
                   "chartType": "column", "anchor": "D60:J76", "title": "Offscreen Chart"}},
        {"command": "add", "parent": "/", "type": "sheet", "props": {"name": "Region Two"}},
        {"command": "set", "path": "/Region Two/A1", "props": {"value": "Cat"}},
        {"command": "set", "path": "/Region Two/B1", "props": {"value": "Val"}},
        {"command": "set", "path": "/Region Two/A2", "props": {"value": "X"}},
        {"command": "set", "path": "/Region Two/B2", "props": {"value": "5"}},
        {"command": "set", "path": "/Region Two/A3", "props": {"value": "Y"}},
        {"command": "set", "path": "/Region Two/B3", "props": {"value": "9"}},
        {"command": "add", "parent": "/Region Two", "type": "chart",
         "props": {"dataRange": "Region Two!B2:B3", "categories": "Region Two!A2:A3",
                   "chartType": "line", "anchor": "A40:F56", "title": "Region Two Trend"}},
    ]


def _xlsx_overlapping_charts_commands():
    """Two charts sharing the identical start/end anchor: identity must stay
    correct and distinct, but the capture must report the resulting occlusion
    rather than fabricate an unobscured layout."""
    return [
        {"command": "set", "path": "/Sheet1/A1", "props": {"value": "Quarter"}},
        {"command": "set", "path": "/Sheet1/B1", "props": {"value": "Revenue"}},
        {"command": "set", "path": "/Sheet1/A2", "props": {"value": "Q1"}},
        {"command": "set", "path": "/Sheet1/B2", "props": {"value": "100"}},
        {"command": "set", "path": "/Sheet1/A3", "props": {"value": "Q2"}},
        {"command": "set", "path": "/Sheet1/B3", "props": {"value": "150"}},
        {"command": "add", "parent": "/Sheet1", "type": "chart",
         "props": {"dataRange": "Sheet1!B2:B3", "categories": "Sheet1!A2:A3",
                   "chartType": "column", "anchor": "D2:J18", "title": "Front"}},
        {"command": "add", "parent": "/Sheet1", "type": "chart",
         "props": {"dataRange": "Sheet1!B2:B3", "categories": "Sheet1!A2:A3",
                   "chartType": "line", "anchor": "D2:J18", "title": "Back"}},
    ]


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
    unit mocks), building every fixture into a TemporaryDirectory with
    maintained OfficeCLI create/batch operations. Requires the pinned
    Playwright + managed Chromium build to actually be importable/launchable
    in this venv; once enabled with the toolchain present, fixture build or
    export/capture failure fails the test, it does not skip."""

    @classmethod
    def setUpClass(cls):
        if not shutil.which("officecli"):
            raise unittest.SkipTest("OfficeCLI executable unavailable; this is not a passed preflight")
        cls.cli = runner.OfficeCLI(timeout=120)
        import playwright_capture
        try:
            cls.engine = playwright_capture.CaptureEngine()
        except playwright_capture.CaptureError as exc:
            raise unittest.SkipTest(f"Managed Chromium unavailable; not a passed render proof: {exc}")
        cls._stack = tempfile.TemporaryDirectory(prefix="officecli-chart-fixtures-")
        cls.fixtures = Path(cls._stack.name)
        cls.docx_pagebreaks = _build(cls.cli, cls.fixtures / "docx1_pagebreaks.docx",
                                     _docx_three_pages_commands())
        cls.docx_naturalflow = _build(cls.cli, cls.fixtures / "docx2_naturalflow.docx",
                                      _docx_natural_flow_overflow_commands())
        cls.pptx_corrected = _build(cls.cli, cls.fixtures / "pptx1_4slides_corrected.pptx",
                                    _pptx_four_slides_commands())
        cls.xlsx_regions = _build(cls.cli, cls.fixtures / "xlsx1_regions.xlsx",
                                  _xlsx_five_regions_commands())
        cls.xlsx_chart_matrix = _build(cls.cli, cls.fixtures / "xlsx2_chart_matrix.xlsx",
                                       _xlsx_chart_matrix_commands())
        cls.xlsx_overlap = _build(cls.cli, cls.fixtures / "xlsx3_overlap.xlsx",
                                  _xlsx_overlapping_charts_commands())

    @classmethod
    def tearDownClass(cls):
        cls.engine.close()
        cls._stack.cleanup()

    def _export_html(self, fixture_path, tmp_dir):
        html = Path(tmp_dir) / (fixture_path.stem + ".html")
        runner.checked(self.cli.run(runner.html_export_arguments(fixture_path, html), "html_export"))
        return html

    def _inventory(self, fixture_path):
        with zipfile.ZipFile(fixture_path) as package:
            return company.chart_object_inventory(package)

    def test_docx_explicit_pagebreaks_accepted(self):
        with tempfile.TemporaryDirectory(prefix="docx-pagebreaks-") as tmp:
            html = self._export_html(self.docx_pagebreaks, tmp)
            result = self.engine.capture_docx_pages(html, Path(tmp) / "pages", expected_pages=3)
            self.assertEqual(result["rendered_pages"], 3)
            self.assertEqual([p["overflow"] for p in result["pages"]], [False, False, False])
            for page in result["pages"]:
                self.assertTrue(Path(page["path"]).is_file())

    def test_docx_natural_flow_overflow_rejected(self):
        import playwright_capture
        with tempfile.TemporaryDirectory(prefix="docx-naturalflow-") as tmp:
            html = self._export_html(self.docx_naturalflow, tmp)
            # Whatever explicit page count a caller might assume for this
            # known-bad fixture, the natural-flow layout must never be
            # silently accepted as a passed multi-page capture.
            with self.assertRaises(playwright_capture.CaptureError):
                self.engine.capture_docx_pages(html, Path(tmp) / "pages", expected_pages=2)

    def test_pptx_corrected_title_body_fixture_accepted(self):
        with tempfile.TemporaryDirectory(prefix="pptx-corrected-") as tmp:
            html = self._export_html(self.pptx_corrected, tmp)
            result = self.engine.capture_pptx_slides(html, Path(tmp) / "slides", expected_slides=4)
            self.assertEqual(result["rendered_slides"], 4)
            self.assertEqual(len(result["slides"]), 4)
            for slide in result["slides"]:
                self.assertTrue(Path(slide["path"]).is_file())

    def test_xlsx_region_merge_offset_wide_and_second_sheet(self):
        with tempfile.TemporaryDirectory(prefix="xlsx-regions-") as tmp:
            html = self._export_html(self.xlsx_regions, tmp)
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

    def test_xlsx_chart_matrix_standard_duplicate_titles_beyond_grid_offscreen_and_second_sheet(self):
        """Positive capability matrix: standard column+line charts, duplicate
        titles with distinct data, a chart anchored beyond the emitted cell
        grid, an offscreen/scroll-required chart, and a second sheet whose
        name contains spaces -- all addressed by real inventory identity and
        captured as complete, unclipped, unoccluded chart objects."""
        inventory = self._inventory(self.xlsx_chart_matrix)
        supported = {(o["sheet"], o["shape_id"]): o for o in inventory if o["status"] == "supported"}
        self.assertEqual(len(supported), 4)
        with tempfile.TemporaryDirectory(prefix="xlsx-chart-matrix-") as tmp:
            html = self._export_html(self.xlsx_chart_matrix, tmp)
            # Full render-target coverage must verify against the fresh fixture.
            targets = [{"id": "cells", "type": "range", "range": "Sheet1!A1:B3"},
                       {"id": "cells2", "type": "range", "range": "Region Two!A1:B3"}] + [
                {"id": f"chart-{sheet.replace(' ', '_')}-{shape_id}", "type": "chart",
                 "sheet": sheet, "shape_id": shape_id}
                for sheet, shape_id in supported]
            company.verify_render_targets(self.xlsx_chart_matrix, targets)
            seen_types = set()
            for (sheet, shape_id), obj in supported.items():
                sheet_index = 0 if sheet == "Sheet1" else 1
                out_png = Path(tmp) / f"{sheet.replace(' ', '_')}-{shape_id}.png"
                result = self.engine.capture_xlsx_chart(
                    html, out_png, sheet, obj["expected_dom_path"], shape_id, sheet_index=sheet_index)
                self.assertEqual(result["sheet"], sheet)
                self.assertEqual(result["shape_id"], shape_id)
                self.assertEqual(result["dom_path"], obj["expected_dom_path"])
                self.assertFalse(result["clipped"], f"{sheet}/{shape_id}: unexpected clipping")
                self.assertFalse(result["occluded"], f"{sheet}/{shape_id}: unexpected occlusion")
                self.assertGreater(result["bounds"]["width"], 0)
                self.assertGreater(result["bounds"]["height"], 0)
                self.assertTrue(out_png.is_file())
                seen_types.add(obj["chart_type"])
            self.assertEqual(seen_types, {"column", "line"})

    def test_xlsx_chart_reordering_invalidates_mapping_and_fresh_hash_resolves(self):
        """Removing an earlier chart shifts a later chart's export index
        while its source shape ID is preserved (a real reorder, not a
        relabel). A target pinned to the stale expected_dom_path captured
        before the edit must be rejected, not silently matched; the fresh
        recomputed inventory/hash must resolve the new mapping instead."""
        before = {o["shape_id"]: o for o in self._inventory(self.xlsx_chart_matrix)
                 if o["status"] == "supported" and o["sheet"] == "Sheet1"}
        # shape_id "3" is the line chart, originally the second chart on Sheet1.
        stale_path = before["3"]["expected_dom_path"]
        self.assertEqual(stale_path, "/Sheet1/chart[2]")
        with tempfile.TemporaryDirectory(prefix="xlsx-chart-reorder-") as tmp:
            reordered = Path(tmp) / "reordered.xlsx"
            shutil.copyfile(self.xlsx_chart_matrix, reordered)
            # Removing the first (column) chart shifts the line chart (same
            # shape_id "3") from chart[2] down to chart[1].
            runner.checked(self.cli.run(["batch", str(reordered), "--commands",
                json.dumps([{"command": "remove", "path": "/Sheet1/chart[1]"}]),
                "--stop-on-error"], "author"))
            after = {o["shape_id"]: o for o in self._inventory(reordered) if o["status"] == "supported"}
            self.assertEqual(after["3"]["expected_dom_path"], "/Sheet1/chart[1]")
            self.assertNotEqual(after["3"]["expected_dom_path"], stale_path)
            with self.assertRaisesRegex(company.Invalid, "Stale chart target DOM path"):
                company.verify_render_targets(reordered, [
                    {"id": "cells", "range": "Sheet1!A1:B3"},
                    {"id": "cells2", "range": "Region Two!A1:B3"},
                    {"id": "chart-3", "type": "chart", "sheet": "Sheet1", "shape_id": "3",
                     "expected_dom_path": stale_path},
                    {"id": "chart-4", "type": "chart", "sheet": "Sheet1", "shape_id": "4"},
                    {"id": "chart-rt2", "type": "chart", "sheet": "Region Two", "shape_id": "2"}])
            # The freshly recomputed mapping (no stale expected_dom_path
            # supplied) must resolve and accept the edited package cleanly.
            company.verify_render_targets(reordered, [
                {"id": "cells", "range": "Sheet1!A1:B3"},
                {"id": "cells2", "range": "Region Two!A1:B3"},
                {"id": "chart-3", "type": "chart", "sheet": "Sheet1", "shape_id": "3"},
                {"id": "chart-4", "type": "chart", "sheet": "Sheet1", "shape_id": "4"},
                {"id": "chart-rt2", "type": "chart", "sheet": "Region Two", "shape_id": "2"}])

    def test_xlsx_overlapping_charts_identity_correct_but_occlusion_reported(self):
        """Two charts sharing an identical start/end anchor must still be
        addressed correctly and independently, but their capture must report
        the resulting occlusion rather than claim an unobscured layout."""
        inventory = self._inventory(self.xlsx_overlap)
        supported = {o["shape_id"]: o for o in inventory if o["status"] == "supported"}
        self.assertEqual(len(supported), 2)
        with tempfile.TemporaryDirectory(prefix="xlsx-overlap-") as tmp:
            html = self._export_html(self.xlsx_overlap, tmp)
            occluded_flags = []
            for shape_id, obj in supported.items():
                out_png = Path(tmp) / f"overlap-{shape_id}.png"
                result = self.engine.capture_xlsx_chart(
                    html, out_png, "Sheet1", obj["expected_dom_path"], shape_id, sheet_index=0)
                self.assertEqual(result["shape_id"], shape_id)
                self.assertEqual(result["dom_path"], obj["expected_dom_path"])
                occluded_flags.append(result["occluded"])
            # Identity resolved both objects distinctly (asserted above); at
            # least the back object must be reported as occluded by the front
            # one -- a fully-overlapping pair is never a clean, accepted layout.
            self.assertTrue(any(occluded_flags), "fully overlapping charts must report occlusion")

    def test_xlsx_chart_missing_target_rejected(self):
        """A supported chart with no matching render target must block, not
        silently pass with partial coverage."""
        with self.assertRaisesRegex(company.Invalid, "Incomplete chart-object render coverage"):
            company.verify_render_targets(self.xlsx_chart_matrix, [
                {"id": "cells", "type": "range", "range": "Sheet1!A1:B3"},
                {"id": "cells2", "type": "range", "range": "Region Two!A1:B3"}])

    def test_xlsx_chart_wrong_shape_id_rejected(self):
        """A chart target naming a sheet but a nonexistent shape_id must be
        reported as not matching any supported source object."""
        with self.assertRaisesRegex(company.Invalid, "does not match a supported source chart object"):
            company.verify_render_targets(self.xlsx_chart_matrix, [
                {"id": "cells", "type": "range", "range": "Sheet1!A1:B3"},
                {"id": "cells2", "type": "range", "range": "Region Two!A1:B3"},
                {"id": "chart-bad", "type": "chart", "sheet": "Sheet1", "shape_id": "999"}])


if __name__ == "__main__":
    unittest.main()
