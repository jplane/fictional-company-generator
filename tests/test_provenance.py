import copy
import json
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

import test_company as baseline

company = baseline.company


class ProvenanceTests(unittest.TestCase):
    def bundle(self, root, extra_records=()):
        fixture = baseline.CompanyTests()
        fixture.setUp()
        fixture.canon = company.seed_canon(dict(fixture.request, documents=3))
        fixture.canon["records"].extend(extra_records)
        fixture.plan = company.create_plan(fixture.canon)
        drafts = root / "build" / "drafts"
        manifest, qa = fixture.create_audit_fixture(root, drafts)
        return fixture, drafts, manifest, qa

    def package(self, root, fmt, parts):
        path = root / ("content." + fmt)
        base = {"[Content_Types].xml": "<Types/>", "_rels/.rels": "<Relationships/>"}
        with zipfile.ZipFile(path, "w") as package:
            for part, content in dict(base, **parts).items():
                package.writestr(part, content)
        return path

    def test_hash_bound_audit_and_portable_namespace(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "original"
            root.mkdir()
            fixture, _, manifest, _ = self.bundle(root)
            moved = Path(directory) / "moved"
            shutil.move(str(root), moved)
            self.assertEqual(company.audit(fixture.canon, fixture.plan, moved / "build" / "drafts",
                                          manifest, moved)["documents"], 3)
            before = company.load(moved / "build" / "provenance.json")
            bindings = {"schema_version": 2, "corpus_id": fixture.canon["corpus_id"],
                        "bindings": [{"document_id": fixture.plan["documents"][0]["id"],
                                      "revision": 1,
                                      "artifact_sha256": manifest["documents"][0]["sha256"],
                                      "kind": "sharepoint", "drive_id": "example-drive",
                                      "item_id": "example-item"}]}
            company.save(moved / "build" / "resource-bindings.json", bindings)
            company.audit(fixture.canon, fixture.plan, moved / "build" / "drafts", manifest, moved)
            self.assertEqual(before, company.load(moved / "build" / "provenance.json"))
            bindings["bindings"][0]["artifact_sha256"] = "wrong"
            with self.assertRaisesRegex(company.Invalid, "Stale binding"):
                company.validate_bindings(fixture.canon, fixture.plan, manifest, bindings)

    def test_missing_claim_occurrence_and_stale_mapping_fail(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture, drafts, manifest, _ = self.bundle(Path(directory))
            locations = company.load(Path(directory) / "build" / "claim-locations.json")
            locations["documents"][0]["occurrences"].pop()
            with self.assertRaisesRegex(company.Invalid, "coverage incomplete"):
                company.verify_provenance(fixture.canon, fixture.plan, drafts, manifest, directory, locations)
            locations = company.load(Path(directory) / "build" / "claim-locations.json")
            locations["documents"][0]["artifact_sha256"] = "stale"
            with self.assertRaisesRegex(company.Invalid, "stale artifact/draft"):
                company.verify_provenance(fixture.canon, fixture.plan, drafts, manifest, directory, locations)

    def test_repeated_sections_cannot_reuse_one_occurrence(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture, drafts, manifest, _ = self.bundle(Path(directory))
            locations = company.load(Path(directory) / "build" / "claim-locations.json")
            occurrences = locations["documents"][0]["occurrences"]
            occurrences[1].update({key: occurrences[0][key] for key in ("locator", "start", "end")})
            with self.assertRaisesRegex(company.Invalid, "duplicate occurrence"):
                company.verify_provenance(fixture.canon, fixture.plan, drafts, manifest, directory, locations)

    def test_paragraph_and_table_aliases_cannot_reuse_physical_text(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            fixture, drafts, manifest, _ = self.bundle(root)
            doc = next(doc for doc in fixture.plan["documents"] if doc["format"] == "docx")
            entry = next(item for item in manifest["documents"] if item["document_id"] == doc["id"])
            path = root / entry["path"]
            with zipfile.ZipFile(path) as package:
                parts = {name: package.read(name) for name in package.namelist()}
            locations = company.load(root / "build" / "claim-locations.json")
            mapping = next(item for item in locations["documents"] if item["document_id"] == doc["id"])
            value = mapping["occurrences"][0]["expected_text"]
            w = company.NS["w"]
            parts["word/document.xml"] = (
                f'<w:document xmlns:w="{w}"><w:body><w:tbl><w:tr><w:tc>'
                f'<w:p><w:r><w:t>{escape(value)}</w:t></w:r></w:p>'
                '</w:tc></w:tr></w:tbl></w:body></w:document>')
            with zipfile.ZipFile(path, "w") as package:
                for name, content in parts.items():
                    package.writestr(name, content)
            entry["sha256"] = company.file_hash(path)
            mapping["artifact_sha256"] = entry["sha256"]
            mapping["occurrences"][1]["locator"] = {
                "kind": "table_cell", "part": "word/document.xml", "table": 0, "row": 0, "column": 0}
            with self.assertRaisesRegex(company.Invalid, "overlapping container aliases"):
                company.verify_provenance(fixture.canon, fixture.plan, drafts, manifest, root, locations)

    def test_span_revision_and_exact_numeric_matching(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture, drafts, manifest, _ = self.bundle(Path(directory))
            original = company.load(Path(directory) / "build" / "claim-locations.json")
            for mutation, error in (("span", "outside"), ("revision", "revision"), ("text", "does not match")):
                with self.subTest(mutation=mutation):
                    locations = copy.deepcopy(original)
                    doc = locations["documents"][0]
                    if mutation == "span":
                        doc["occurrences"][0]["end"] = 100000
                    elif mutation == "revision":
                        doc["revision"] = 99
                    else:
                        doc["occurrences"][0]["expected_text"] = "wrong"
                    with self.assertRaisesRegex(company.Invalid, error):
                        company.verify_provenance(fixture.canon, fixture.plan, drafts, manifest, directory, locations)
            self.assertEqual(company.display_value({"value": 12000000},
                             {"scale": 1000000, "decimals": 1}), "12.0")
            self.assertEqual(company.display_value({"value": 12345},
                             {"decimals": 0, "grouping": True, "prefix": "$"}), "$12,345")

    def test_coverage_is_recomputed_not_trusted(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture, drafts, manifest, _ = self.bundle(Path(directory))
            coverage = company.load(Path(directory) / "build" / "coverage.json")
            self.assertEqual(coverage["delivered"]["unverified_claims"], 0)
            self.assertEqual(coverage["delivered"]["documents"], 3)
            self.assertEqual(coverage["delivered"]["claims"], 3)
            coverage["delivered"]["documents"] = 999
            (Path(directory) / "build" / "coverage.json").write_text(json.dumps(coverage), encoding="utf-8")
            with self.assertRaisesRegex(company.Invalid, "incorrect coverage"):
                company.audit(fixture.canon, fixture.plan, drafts, manifest, directory)

    def test_forged_provenance_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture, drafts, manifest, _ = self.bundle(Path(directory))
            receipt = company.load(Path(directory) / "build" / "provenance.json")
            receipt["documents"][0]["occurrences"][0]["record_id"] = "imaginary"
            (Path(directory) / "build" / "provenance.json").write_text(json.dumps(receipt), encoding="utf-8")
            with self.assertRaisesRegex(company.Invalid, "forged provenance"):
                company.audit(fixture.canon, fixture.plan, drafts, manifest, directory)

    def test_access_view_respects_dates_and_omits_private_build_metadata(self):
        with tempfile.TemporaryDirectory() as directory:
            fixture, _, manifest, _ = self.bundle(Path(directory))
            receipt = company.load(Path(directory) / "build" / "provenance.json")
            early = company.access_view(fixture.canon, fixture.plan, manifest, receipt,
                                        "person-avery", "2022-01-01")
            self.assertFalse(early["documents"])
            view = company.access_view(fixture.canon, fixture.plan, manifest, receipt,
                                       "person-avery", "2026-01-01")
            self.assertEqual(len(view["documents"]), 3)
            wording = json.dumps(view)
            self.assertNotIn("condition_ids", wording)
            self.assertNotIn("canon_sha256", wording)
            self.assertNotIn("record_ids", wording)
            public = company.access_view(fixture.canon, fixture.plan, manifest, receipt,
                                         "public", "2026-01-01")
            selected = {item["document_id"] for item in public["documents"]}
            self.assertEqual(selected, {doc["id"] for doc in fixture.plan["documents"]
                                        if doc["audience"] == "public"})

    def test_word_split_runs_tables_and_referenced_header(self):
        w, r = company.NS["w"], company.NS["r"]
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        parts = {
            "word/document.xml": f'<w:document xmlns:w="{w}" xmlns:r="{r}"><w:body>'
                '<w:p><w:r><w:t>Lumen</w:t></w:r><w:r><w:t>vale</w:t></w:r></w:p>'
                '<w:tbl><w:tr><w:tc><w:p><w:r><w:t>12</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
                '<w:sectPr><w:headerReference r:id="head"/></w:sectPr></w:body></w:document>',
            "word/_rels/document.xml.rels": f'<Relationships xmlns="{rel_ns}"><Relationship Id="head" '
                f'Type="{r}/header" Target="header9.xml"/></Relationships>',
            "word/header9.xml": f'<w:hdr xmlns:w="{w}"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>',
        }
        with tempfile.TemporaryDirectory() as directory:
            entries = company.artifact_containers(self.package(Path(directory), "docx", parts), "docx")
            self.assertEqual(entries[0]["text"], "Lumenvale")
            self.assertTrue(any(item["locator"]["kind"] == "table_cell" and item["text"] == "12"
                                for item in entries))
            self.assertTrue(any(item["locator"]["part"] == "word/header9.xml" for item in entries))
            with self.assertRaisesRegex(company.Invalid, "missing or ambiguous"):
                company.located_container(entries, {"kind": "paragraph", "part": "word/document.xml", "index": 9})

    def test_presentation_relationship_order_tables_notes_and_charts(self):
        p, a, r, c = (company.NS[key] for key in ("p", "a", "r", "c"))
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        def shape(value, identifier):
            return f'<p:sp><p:nvSpPr><p:cNvPr id="{identifier}"/></p:nvSpPr><p:txBody>' \
                   f'<a:p><a:r><a:t>{value}</a:t></a:r></a:p></p:txBody></p:sp>'
        parts = {
            "ppt/presentation.xml": f'<p:presentation xmlns:p="{p}" xmlns:r="{r}"><p:sldIdLst>'
                '<p:sldId id="900" r:id="first"/><p:sldId id="256" r:id="second"/></p:sldIdLst></p:presentation>',
            "ppt/_rels/presentation.xml.rels": f'<Relationships xmlns="{rel_ns}">'
                f'<Relationship Id="first" Type="{r}/slide" Target="slides/slide9.xml"/>'
                f'<Relationship Id="second" Type="{r}/slide" Target="slides/slide2.xml"/></Relationships>',
            "ppt/slides/slide9.xml": f'<p:sld xmlns:p="{p}" xmlns:a="{a}" xmlns:c="{c}" xmlns:r="{r}">'
                f'<p:cSld><p:spTree>{shape("First", 1)}<p:graphicFrame><p:nvGraphicFramePr>'
                '<p:cNvPr id="2"/></p:nvGraphicFramePr><a:graphic><a:graphicData><a:tbl><a:tr>'
                '<a:tc><a:txBody><a:p><a:r><a:t>Cell</a:t></a:r></a:p></a:txBody></a:tc>'
                '</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
                '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3"/></p:nvGraphicFramePr>'
                '<a:graphic><a:graphicData><c:chart r:id="chart"/></a:graphicData></a:graphic>'
                '</p:graphicFrame></p:spTree></p:cSld></p:sld>',
            "ppt/slides/slide2.xml": f'<p:sld xmlns:p="{p}" xmlns:a="{a}"><p:cSld>'
                f'<p:spTree>{shape("Second", 1)}</p:spTree></p:cSld></p:sld>',
            "ppt/slides/_rels/slide9.xml.rels": f'<Relationships xmlns="{rel_ns}">'
                f'<Relationship Id="notes" Type="{r}/notesSlide" Target="../notesSlides/notesSlide1.xml"/>'
                f'<Relationship Id="chart" Type="{r}/chart" Target="../charts/chart1.xml"/></Relationships>',
            "ppt/notesSlides/notesSlide1.xml": f'<p:notes xmlns:p="{p}" xmlns:a="{a}"><p:cSld>'
                f'<p:spTree>{shape("Speaker note", 1)}</p:spTree></p:cSld></p:notes>',
            "ppt/charts/chart1.xml": f'<c:chartSpace xmlns:c="{c}"><c:chart><c:plotArea><c:barChart>'
                '<c:ser><c:val><c:numRef><c:f>Results!B2</c:f><c:numCache>'
                '<c:pt idx="0"><c:v>12</c:v></c:pt></c:numCache></c:numRef></c:val>'
                '</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>',
        }
        with tempfile.TemporaryDirectory() as directory:
            entries = company.artifact_containers(self.package(Path(directory), "pptx", parts), "pptx")
            self.assertEqual(entries[0]["text"], "First")
            self.assertEqual(entries[0]["locator"]["slide_id"], "900")
            self.assertTrue(any(entry["locator"]["kind"] == "slide_cell" for entry in entries))
            self.assertTrue(any(entry["locator"]["kind"] == "notes_shape" for entry in entries))
            chart = next(entry for entry in entries if entry["locator"]["kind"] == "chart_point")
            self.assertEqual(chart["text"], "12")
            self.assertEqual(chart["formula"], "Results!B2")

    def test_workbook_relationships_strings_formula_caches_and_ranges(self):
        x, r = company.NS["x"], company.NS["r"]
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        parts = {
            "xl/workbook.xml": f'<workbook xmlns="{x}" xmlns:r="{r}"><sheets>'
                '<sheet name="Actuals" sheetId="8" r:id="sheet"/></sheets></workbook>',
            "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{rel_ns}">'
                f'<Relationship Id="sheet" Type="{r}/worksheet" Target="worksheets/sheet7.xml"/>'
                f'<Relationship Id="strings" Type="{r}/sharedStrings" Target="strings.xml"/></Relationships>',
            "xl/strings.xml": f'<sst xmlns="{x}"><si><r><t>Lumen</t></r><r><t>vale</t></r></si></sst>',
            "xl/worksheets/sheet7.xml": f'<worksheet xmlns="{x}"><sheetData><row r="1">'
                '<c r="A1" t="s"><v>0</v></c><c r="B1"><f>SUM(B2:B3)</f><v>12</v></c>'
                '</row><row r="2"><c r="A2" t="inlineStr"><is><t>Source</t></is></c>'
                '<c r="B2"><v>5</v></c></row><row r="3"><c r="B3"><v>7</v></c></row>'
                '</sheetData></worksheet>',
        }
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            entries = company.artifact_containers(path, "xlsx")
            self.assertEqual(entries[0]["text"], "Lumenvale")
            self.assertEqual(entries[1]["formula"], "SUM(B2:B3)")
            locator = {"kind": "range", "part": "xl/worksheets/sheet7.xml", "sheet_id": "8",
                       "sheet": "Actuals", "range": "B2:B3"}
            self.assertEqual(company.located_container(entries, locator)["text"], "5 7")
            original = parts["xl/worksheets/sheet7.xml"]
            parts["xl/worksheets/sheet7.xml"] = original.replace(
                "<f>SUM(B2:B3)</f>", '<f t="shared" si="0" ref="B1:B2">SUM(B2:B3)</f>').replace(
                '<c r="B2"><v>5</v>', '<c r="B2"><f t="shared" si="0"/><v>5</v>')
            shared = company.artifact_containers(self.package(Path(directory), "xlsx", parts), "xlsx")
            follower = next(entry for entry in shared if entry["locator"]["cell"] == "B2")
            self.assertEqual(follower["formula"], {
                "kind": "shared", "index": "0", "base_cell": "B1",
                "base_formula": "SUM(B2:B3)", "range": "B1:B2", "cell": "B2"})
            parts["xl/worksheets/sheet7.xml"] = parts["xl/worksheets/sheet7.xml"].replace(
                '<f t="shared" si="0"/>', '<f t="shared" si="9"/>')
            with self.assertRaisesRegex(company.Invalid, "shared formula base missing"):
                company.artifact_containers(self.package(Path(directory), "xlsx", parts), "xlsx")
            parts["xl/worksheets/sheet7.xml"] = original
            parts["xl/worksheets/sheet7.xml"] = parts["xl/worksheets/sheet7.xml"].replace(
                "<f>SUM(B2:B3)</f><v>12</v>", "<f>SUM(B2:B3)</f>")
            with self.assertRaisesRegex(company.Invalid, "cached value"):
                company.artifact_containers(self.package(Path(directory), "xlsx", parts), "xlsx")

    def test_external_required_relationship_is_rejected(self):
        p, r = company.NS["p"], company.NS["r"]
        parts = {"ppt/presentation.xml": f'<p:presentation xmlns:p="{p}" xmlns:r="{r}"><p:sldIdLst>'
                 '<p:sldId id="256" r:id="s"/></p:sldIdLst></p:presentation>',
                 "ppt/_rels/presentation.xml.rels": '<Relationships><Relationship Id="s" '
                 f'Type="{r}/slide" Target="https://example.invalid/slide.xml" TargetMode="External"/>'
                 '</Relationships>', "ppt/slides/slide1.xml": "<sld/>"}
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaisesRegex(company.Invalid, "external"):
                company.artifact_containers(self.package(Path(directory), "pptx", parts), "pptx")

    def test_derived_cells_require_real_saved_formulas(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            metric = {"id": "metric-source", "kind": "business_note", "known_on": "2022-01-01",
                      "effective_on": "2022-01-01", "data": {"value": 5}, "refs": ["company"]}
            fixture, drafts, manifest, _ = self.bundle(root, [metric])
            doc = next(doc for doc in fixture.plan["documents"] if doc["format"] == "xlsx")
            draft_path = drafts / (doc["id"] + ".json")
            draft = company.load(draft_path)
            value = 5
            anchor_claim = draft["claims"][0]
            draft["claims"] = [anchor_claim,
                {"id": "a", "record_id": "metric-source", "path": ["data", "value"],
                 "value": value, "units": "count"},
                {"id": "b", "record_id": "metric-source", "path": ["data", "value"],
                 "value": value, "units": "count"},
                {"id": "total", "value": value * 2, "units": "count",
                 "derivation": {"operation": "sum", "inputs": ["a", "b"], "rounding": 0}},
            ]
            for section in draft["sections"]:
                section["claim_ids"] = ["c1", "a", "b", "total"]
            draft_path.write_text(json.dumps(draft), encoding="utf-8")
            locations = company.load(root / "build" / "claim-locations.json")
            mapping = next(item for item in locations["documents"] if item["document_id"] == doc["id"])
            entry = next(item for item in manifest["documents"] if item["document_id"] == doc["id"])
            path = root / entry["path"]
            with zipfile.ZipFile(path) as package:
                parts = {name: package.read(name) for name in package.namelist()}
            x = company.NS["x"]
            for formula in (None, "A1+B1"):
                rows, occurrences = [], []
                for number in range(len(draft["sections"])):
                    cells = []
                    for column, claim_id, scalar in (
                            ("A", "c1", company.display_value(anchor_claim)),
                            ("B", "a", value), ("C", "b", value), ("D", "total", value * 2)):
                        address = column + str(number + 1)
                        saved_formula = (f"<f>B{number + 1}+C{number + 1}</f>"
                                         if formula and column == "D" else "")
                        if column == "A":
                            cells.append(f'<c r="{address}" t="inlineStr"><is><t>'
                                         f'{escape(scalar)}</t></is></c>')
                        else:
                            cells.append(f'<c r="{address}">{saved_formula}<v>{scalar}</v></c>')
                        occurrence = {
                            "claim_id": claim_id, "section_index": number,
                            "locator": {"kind": "cell", "part": "xl/worksheets/sheet1.xml",
                                        "sheet_id": "1", "sheet": "Results", "cell": address},
                            "start": 0, "end": len(str(scalar)), "expected_text": str(scalar),
                        }
                        if column == "D":
                            occurrence["formula"] = f"B{number + 1}+C{number + 1}" if formula else None
                        occurrences.append(occurrence)
                    rows.append(f'<row r="{number + 1}">{"".join(cells)}</row>')
                parts["xl/worksheets/sheet1.xml"] = (
                    f'<worksheet xmlns="{x}"><sheetData>{"".join(rows)}</sheetData></worksheet>')
                with zipfile.ZipFile(path, "w") as package:
                    for name, content in parts.items():
                        package.writestr(name, content)
                entry["sha256"] = company.file_hash(path)
                mapping.update(artifact_sha256=entry["sha256"], draft_sha256=company.file_hash(draft_path),
                               occurrences=occurrences)
                if formula is None:
                    with self.assertRaisesRegex(company.Invalid, "saved formula evidence"):
                        company.verify_provenance(fixture.canon, fixture.plan, drafts, manifest, root, locations)
                else:
                    company.verify_provenance(fixture.canon, fixture.plan, drafts, manifest, root, locations)
                    mapping["occurrences"][-1]["formula"] = "1+1"
                    with self.assertRaisesRegex(company.Invalid, "saved formula differs"):
                        company.verify_provenance(fixture.canon, fixture.plan, drafts, manifest, root, locations)

    def test_new_cli_commands_and_errors(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            fixture, drafts, manifest, _ = self.bundle(root)
            company.save(root / "build" / "plan.json", fixture.plan)
            company.save(root / "build" / "manifest.json", manifest)
            base = ["--canon", str(root / "canon.json"), "--plan", str(root / "build" / "plan.json"),
                    "--manifest", str(root / "build" / "manifest.json"), "--root", str(root)]
            commands = [
                ["index-artifacts", *base, "--out", str(root / "build" / "new-index.json")],
                ["verify-provenance", *base, "--drafts", str(drafts), "--locations",
                 str(root / "build" / "claim-locations.json"), "--out", str(root / "build" / "new-proof.json")],
                ["coverage", *base, "--drafts", str(drafts), "--locations",
                 str(root / "build" / "claim-locations.json"), "--out", str(root / "build" / "new-coverage.json")],
                ["access-view", *base, "--drafts", str(drafts), "--principal", "person-avery",
                 "--as-of", "2026-01-01", "--out", str(root / "build" / "views" / "avery.json")],
            ]
            for args in commands:
                result = subprocess.run([sys.executable, str(baseline.SCRIPT), *args],
                                        capture_output=True, text=True, check=False)
                self.assertEqual(result.returncode, 0, result.stderr)
                self.assertEqual(json.loads(result.stdout)["schema_version"], 2)
            args = commands[0][:-2] + ["--out", str(root / "documents" / "bad.json")]
            result = subprocess.run([sys.executable, str(baseline.SCRIPT), *args],
                                    capture_output=True, text=True, check=False)
            self.assertEqual(result.returncode, 1)
            self.assertIn("build directory", result.stderr)
            self.assertFalse((root / "documents" / "bad.json").exists())


if __name__ == "__main__":
    unittest.main()
