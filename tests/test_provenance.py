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

    def _xlsx_chart_parts(self, cNvPr_id, chart_body, anchor_tag="twoCellAnchor",
                          anchor_extra='<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>'
                          '<xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                          '<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff>'
                          '<xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'):
        return self._xlsx_charts_package([
            {"cNvPr_id": cNvPr_id, "chart_body": chart_body, "anchor_tag": anchor_tag,
             "anchor_extra": anchor_extra}])

    def _xlsx_charts_package(self, charts, broken_chart_rel=False):
        """Build a minimal XLSX part set with N chart-bearing graphic frames in
        one drawing part, each with its own chart relationship/part. `charts`
        is a list of dicts with cNvPr_id/chart_body and optional anchor_tag/
        anchor_extra (defaults to a standard twoCellAnchor)."""
        default_anchor = ('<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff>'
                          '<xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                          '<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff>'
                          '<xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>')
        x, r, a, c, xdr = (company.NS[k] for k in ("x", "r", "a", "c", "xdr"))
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        anchors_xml, chart_rels, chart_parts = [], [], {}
        for index, chart in enumerate(charts, start=1):
            anchor_tag = chart.get("anchor_tag", "twoCellAnchor")
            anchor_extra = chart.get("anchor_extra", default_anchor)
            rel_id = f"chartrel{index}"
            anchors_xml.append(
                f'<xdr:{anchor_tag}>{anchor_extra}<xdr:graphicFrame><xdr:nvGraphicFramePr>'
                f'<xdr:cNvPr id="{chart["cNvPr_id"]}" name="Chart"/><xdr:cNvGraphicFramePr/>'
                '</xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>'
                f'<a:graphic><a:graphicData uri="{c}"><c:chart xmlns:c="{c}" r:id="{rel_id}"/>'
                f'</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:{anchor_tag}>')
            if not (broken_chart_rel and index == len(charts)):
                chart_rels.append(f'<Relationship Id="{rel_id}" Type="{r}/chart" '
                                  f'Target="charts/chart{index}.xml"/>')
                chart_parts[f"xl/drawings/charts/chart{index}.xml"] = chart["chart_body"]
        return {
            "xl/workbook.xml": f'<workbook xmlns="{x}" xmlns:r="{r}"><sheets>'
                '<sheet name="Data" sheetId="1" r:id="sheet"/></sheets></workbook>',
            "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{rel_ns}">'
                f'<Relationship Id="sheet" Type="{r}/worksheet" Target="worksheets/sheet1.xml"/>'
                '</Relationships>',
            "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{x}"><sheetData><row r="1">'
                '<c r="A1"><v>1</v></c></row></sheetData><drawing xmlns:r="' + r + '" r:id="drawing"/></worksheet>',
            "xl/worksheets/_rels/sheet1.xml.rels": f'<Relationships xmlns="{rel_ns}">'
                f'<Relationship Id="drawing" Type="{r}/drawing" Target="../drawings/drawing1.xml"/>'
                '</Relationships>',
            "xl/drawings/drawing1.xml": f'<xdr:wsDr xmlns:xdr="{xdr}" xmlns:a="{a}" xmlns:r="{r}">'
                + "".join(anchors_xml) + '</xdr:wsDr>',
            "xl/drawings/_rels/drawing1.xml.rels": f'<Relationships xmlns="{rel_ns}">'
                + "".join(chart_rels) + '</Relationships>',
            **chart_parts,
        }

    def _column_chart_body(self, value="5"):
        c = company.NS["c"]
        return (f'<c:chartSpace xmlns:c="{c}"><c:chart><c:title><c:tx><c:rich><a:p '
                'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                '<a:r><a:t>Quarterly Summary</a:t></a:r></a:p></c:rich></c:tx></c:title>'
                '<c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:val><c:numRef>'
                f'<c:f>Data!B2</c:f><c:numCache><c:pt idx="0"><c:v>{value}</c:v></c:pt>'
                '</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea>'
                '</c:chart></c:chartSpace>')

    def test_xlsx_render_targets_reject_unreferenced_orphan_drawing_part(self):
        """An xl/drawings part not referenced by any worksheet is a malformed/
        suspicious package and must still block, with a specific diagnostic
        (not the old blanket 'any drawing part exists' message)."""
        x, r = company.NS["x"], company.NS["r"]
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        parts = {
            "xl/workbook.xml": f'<workbook xmlns="{x}" xmlns:r="{r}"><sheets>'
                '<sheet name="Data" sheetId="1" r:id="sheet"/></sheets></workbook>',
            "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{rel_ns}">'
                f'<Relationship Id="sheet" Type="{r}/worksheet" Target="worksheets/sheet1.xml"/>'
                '</Relationships>',
            "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{x}"><sheetData><row r="1">'
                '<c r="A1"><v>1</v></c></row></sheetData></worksheet>',
            "xl/drawings/drawing1.xml": '<xdr:wsDr xmlns:xdr='
                '"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
        }
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            with self.assertRaisesRegex(company.Invalid, "not referenced by any worksheet"):
                company.verify_render_targets(path, [{"id": "all", "range": "Data!A1:A1"}])

    def test_xlsx_render_targets_reject_unsupported_chart_type(self):
        """A referenced, well-formed chart object of an unsupported type (pie)
        must be explicitly reported, not silently substituted or omitted."""
        c = company.NS["c"]
        pie = (f'<c:chartSpace xmlns:c="{c}"><c:chart><c:plotArea><c:pieChart>'
               '<c:ser><c:val><c:numRef><c:f>Data!B2</c:f><c:numCache>'
               '<c:pt idx="0"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val>'
               '</c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>')
        parts = self._xlsx_chart_parts("2", pie)
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            with self.assertRaisesRegex(company.Invalid, "Unsupported chart type: pieChart"):
                company.verify_render_targets(path, [{"id": "all", "range": "Data!A1:A1"}])

    def test_xlsx_render_targets_accept_supported_column_chart_with_full_coverage(self):
        """A supported column chart is accepted once both its cell range and
        its specific source chart object (sheet + shape_id) are targeted."""
        c = company.NS["c"]
        column = (f'<c:chartSpace xmlns:c="{c}"><c:chart><c:plotArea><c:barChart>'
                  '<c:barDir val="col"/><c:ser><c:val><c:numRef><c:f>Data!B2</c:f><c:numCache>'
                  '<c:pt idx="0"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val>'
                  '</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>')
        parts = self._xlsx_chart_parts("2", column)
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            company.verify_render_targets(path, [
                {"id": "cells", "type": "range", "range": "Data!A1:A1"},
                {"id": "chart-2", "type": "chart", "sheet": "Data", "shape_id": "2"}])
            with self.assertRaisesRegex(company.Invalid, "Incomplete chart-object render coverage"):
                company.verify_render_targets(path, [{"id": "cells", "type": "range", "range": "Data!A1:A1"}])
            with self.assertRaisesRegex(company.Invalid, "does not match a supported source chart object"):
                company.verify_render_targets(path, [
                    {"id": "cells", "type": "range", "range": "Data!A1:A1"},
                    {"id": "chart-9", "type": "chart", "sheet": "Data", "shape_id": "9"}])

    def test_chart_inventory_duplicate_titles_distinct_source_objects_stay_addressable(self):
        """Two column charts carrying the identical <c:title> text but distinct
        shape IDs/data must be inventoried as two separate supported objects
        with distinct export paths; identity never depends on title text."""
        parts = self._xlsx_charts_package([
            {"cNvPr_id": "2", "chart_body": self._column_chart_body("100")},
            {"cNvPr_id": "3", "chart_body": self._column_chart_body("150")}])
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            with zipfile.ZipFile(path) as package:
                inventory = company.chart_object_inventory(package)
            supported = {o["shape_id"]: o for o in inventory if o["status"] == "supported"}
            self.assertEqual(set(supported), {"2", "3"})
            self.assertEqual(supported["2"]["expected_dom_path"], "/Data/chart[1]")
            self.assertEqual(supported["3"]["expected_dom_path"], "/Data/chart[2]")
            company.verify_render_targets(path, [
                {"id": "cells", "range": "Data!A1:A1"},
                {"id": "chart-2", "type": "chart", "sheet": "Data", "shape_id": "2"},
                {"id": "chart-3", "type": "chart", "sheet": "Data", "shape_id": "3"}])

    def test_chart_inventory_identical_start_anchor_disambiguated_by_shape_id(self):
        """Two charts sharing the exact same twoCellAnchor from/to coordinates
        must still resolve to two distinct, independently addressable source
        objects (identity is never collapsed onto the shared anchor)."""
        parts = self._xlsx_charts_package([
            {"cNvPr_id": "2", "chart_body": self._column_chart_body("100")},
            {"cNvPr_id": "3", "chart_body": self._column_chart_body("150")}])
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            with zipfile.ZipFile(path) as package:
                inventory = company.chart_object_inventory(package)
            anchors = {o["shape_id"]: o["two_cell_anchor"] for o in inventory}
            self.assertEqual(anchors["2"], anchors["3"])
            self.assertNotEqual(
                next(o for o in inventory if o["shape_id"] == "2")["expected_dom_path"],
                next(o for o in inventory if o["shape_id"] == "3")["expected_dom_path"])

    def test_chart_inventory_rejects_duplicate_source_shape_ids(self):
        """A malformed drawing part with two graphic frames sharing one shape
        ID is an ambiguous source and must fail closed, not pick one silently."""
        parts = self._xlsx_charts_package([
            {"cNvPr_id": "2", "chart_body": self._column_chart_body("100")},
            {"cNvPr_id": "2", "chart_body": self._column_chart_body("150")}])
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            with zipfile.ZipFile(path) as package:
                with self.assertRaisesRegex(company.Invalid, "duplicate source shape ID"):
                    company.chart_object_inventory(package)

    def test_chart_inventory_rejects_broken_chart_relationship(self):
        """A chart reference whose relationship ID is missing/unresolvable in
        the drawing part's rels must fail closed, not skip the object."""
        parts = self._xlsx_charts_package(
            [{"cNvPr_id": "2", "chart_body": self._column_chart_body("100")}],
            broken_chart_rel=True)
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            with zipfile.ZipFile(path) as package:
                with self.assertRaisesRegex(company.Invalid, "Missing, external, or wrong chart relationship"):
                    company.chart_object_inventory(package)

    def test_chart_inventory_reordering_invalidates_stale_target_mapping(self):
        """Inserting a new chart ahead of an existing one shifts the later
        chart's export index; a target built from the stale mapping must be
        rejected rather than silently matching the wrong rendered object."""
        c = company.NS["c"]
        line = (f'<c:chartSpace xmlns:c="{c}"><c:chart><c:plotArea><c:lineChart>'
               '<c:ser><c:val><c:numRef><c:f>Data!B2</c:f><c:numCache>'
               '<c:pt idx="0"><c:v>20</c:v></c:pt></c:numCache></c:numRef></c:val>'
               '</c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>')
        parts = self._xlsx_charts_package([{"cNvPr_id": "2", "chart_body": line}])
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            with zipfile.ZipFile(path) as package:
                before = company.chart_object_inventory(package)
            self.assertEqual(before[0]["expected_dom_path"], "/Data/chart[1]")
            stale_target = {"id": "chart-2", "type": "chart", "sheet": "Data",
                            "shape_id": "2", "expected_dom_path": before[0]["expected_dom_path"]}
            company.verify_render_targets(path, [{"id": "cells", "range": "Data!A1:A1"}, stale_target])

            reordered_parts = self._xlsx_charts_package([
                {"cNvPr_id": "9", "chart_body": self._column_chart_body("5")},
                {"cNvPr_id": "2", "chart_body": line}])
            reordered_path = self.package(Path(directory), "xlsx", reordered_parts)
            with zipfile.ZipFile(reordered_path) as package:
                after = company.chart_object_inventory(package)
            moved = next(o for o in after if o["shape_id"] == "2")
            self.assertEqual(moved["expected_dom_path"], "/Data/chart[2]")
            self.assertNotEqual(moved["expected_dom_path"], before[0]["expected_dom_path"])
            with self.assertRaisesRegex(company.Invalid, "Stale chart target DOM path"):
                company.verify_render_targets(reordered_path, [
                    {"id": "cells", "range": "Data!A1:A1"},
                    stale_target,
                    {"id": "chart-9", "type": "chart", "sheet": "Data", "shape_id": "9"}])
        """An xl/drawings part not referenced by any worksheet is a malformed/
        suspicious package and must still block, with a specific diagnostic
        (not the old blanket 'any drawing part exists' message)."""
        x, r = company.NS["x"], company.NS["r"]
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        parts = {
            "xl/workbook.xml": f'<workbook xmlns="{x}" xmlns:r="{r}"><sheets>'
                '<sheet name="Data" sheetId="1" r:id="sheet"/></sheets></workbook>',
            "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{rel_ns}">'
                f'<Relationship Id="sheet" Type="{r}/worksheet" Target="worksheets/sheet1.xml"/>'
                '</Relationships>',
            "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{x}"><sheetData><row r="1">'
                '<c r="A1"><v>1</v></c></row></sheetData></worksheet>',
            "xl/drawings/drawing1.xml": '<xdr:wsDr xmlns:xdr='
                '"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
        }
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            with self.assertRaisesRegex(company.Invalid, "not referenced by any worksheet"):
                company.verify_render_targets(path, [{"id": "all", "range": "Data!A1:A1"}])

    def test_xlsx_render_targets_reject_unsupported_chart_type(self):
        """A referenced, well-formed chart object of an unsupported type (pie)
        must be explicitly reported, not silently substituted or omitted."""
        c = company.NS["c"]
        pie = (f'<c:chartSpace xmlns:c="{c}"><c:chart><c:plotArea><c:pieChart>'
               '<c:ser><c:val><c:numRef><c:f>Data!B2</c:f><c:numCache>'
               '<c:pt idx="0"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val>'
               '</c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>')
        parts = self._xlsx_chart_parts("2", pie)
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            with self.assertRaisesRegex(company.Invalid, "Unsupported chart type: pieChart"):
                company.verify_render_targets(path, [{"id": "all", "range": "Data!A1:A1"}])

    def test_xlsx_render_targets_accept_supported_column_chart_with_full_coverage(self):
        """A supported column chart is accepted once both its cell range and
        its specific source chart object (sheet + shape_id) are targeted."""
        c = company.NS["c"]
        column = (f'<c:chartSpace xmlns:c="{c}"><c:chart><c:plotArea><c:barChart>'
                  '<c:barDir val="col"/><c:ser><c:val><c:numRef><c:f>Data!B2</c:f><c:numCache>'
                  '<c:pt idx="0"><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val>'
                  '</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>')
        parts = self._xlsx_chart_parts("2", column)
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            company.verify_render_targets(path, [
                {"id": "cells", "type": "range", "range": "Data!A1:A1"},
                {"id": "chart-2", "type": "chart", "sheet": "Data", "shape_id": "2"}])
            with self.assertRaisesRegex(company.Invalid, "Incomplete chart-object render coverage"):
                company.verify_render_targets(path, [{"id": "cells", "type": "range", "range": "Data!A1:A1"}])
            with self.assertRaisesRegex(company.Invalid, "does not match a supported source chart object"):
                company.verify_render_targets(path, [
                    {"id": "cells", "type": "range", "range": "Data!A1:A1"},
                    {"id": "chart-9", "type": "chart", "sheet": "Data", "shape_id": "9"}])

    def test_xlsx_render_targets_reject_any_drawings_part(self):
        """verify_render_targets must keep blocking any XLSX package that contains
        an xl/drawings/*.xml part (i.e. any chart/drawing object), regardless of
        whether the declared render-target cell ranges otherwise look covered.
        Region/range coverage of plain cells is not proof of chart-object pixel
        coverage; see references/dependencies.md capability-probe notes."""
        x, r = company.NS["x"], company.NS["r"]
        parts = {
            "xl/workbook.xml": f'<workbook xmlns="{x}" xmlns:r="{r}"><sheets>'
                '<sheet name="Data" sheetId="1" r:id="sheet"/></sheets></workbook>',
            "xl/_rels/workbook.xml.rels": '<Relationships>'
                f'<Relationship Id="sheet" Type="{r}/worksheet" Target="worksheets/sheet1.xml"/>'
                '</Relationships>',
            "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{x}"><sheetData><row r="1">'
                '<c r="A1"><v>1</v></c></row></sheetData></worksheet>',
            "xl/drawings/drawing1.xml": '<xdr:wsDr xmlns:xdr='
                '"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
        }
        with tempfile.TemporaryDirectory() as directory:
            path = self.package(Path(directory), "xlsx", parts)
            with self.assertRaisesRegex(company.Invalid, "not referenced by any worksheet"):
                company.verify_render_targets(path, [{"id": "all", "range": "Data!A1:A1"}])


    def test_numeric_cache_round_trip_normalization_retains_raw_evidence(self):
        x, r = company.NS["x"], company.NS["r"]
        parts = {
            "xl/workbook.xml": f'<workbook xmlns="{x}" xmlns:r="{r}"><sheets>'
                '<sheet name="Ratios" sheetId="1" r:id="sheet"/></sheets></workbook>',
            "xl/_rels/workbook.xml.rels": '<Relationships>'
                f'<Relationship Id="sheet" Type="{r}/worksheet" Target="worksheets/sheet1.xml"/>'
                '</Relationships>',
            "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{x}"><sheetData><row r="1">'
                '<c r="A1"><f>ROUND(2148/1000,6)</f><v>2.1480000000000001</v></c>'
                '<c r="B1" t="n"><v>2.148001</v></c>'
                '<c r="C1" t="str"><v>2.1480000000000001</v></c>'
                '<c r="D1"><v>9007199254740993</v></c>'
                '<c r="E1"><v>0.000001</v></c>'
                '<c r="F1"><v>12.0</v></c>'
                '</row></sheetData></worksheet>',
        }
        with tempfile.TemporaryDirectory() as directory:
            entries = company.artifact_containers(self.package(Path(directory), "xlsx", parts), "xlsx")
            self.assertEqual([entry["text"] for entry in entries],
                             ["2.148", "2.148001", "2.1480000000000001",
                              "9007199254740993", "1e-06", "12"])
            self.assertEqual(entries[0]["raw_numeric_value"], "2.1480000000000001")
            self.assertEqual(entries[0]["formula"], "ROUND(2148/1000,6)")
            self.assertNotIn("raw_numeric_value", entries[2])
            original = parts["xl/worksheets/sheet1.xml"]
            for invalid in ("NaN", "Infinity", "not-a-number", "1e-9999"):
                parts["xl/worksheets/sheet1.xml"] = original.replace(
                    "<v>2.1480000000000001</v>", f"<v>{invalid}</v>", 1)
                with self.subTest(invalid=invalid), self.assertRaises(company.Invalid):
                    company.artifact_containers(self.package(Path(directory), "xlsx", parts), "xlsx")

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
