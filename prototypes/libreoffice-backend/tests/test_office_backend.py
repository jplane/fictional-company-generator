import json
import os
import pwd
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from office_backend import (
    BackendError,
    _company_contract,
    inspect_recalculated_workbook,
    recalculate_xlsx,
    run_demo,
    validate_pdf,
)


class LibreOfficeBackendSmokeTests(unittest.TestCase):
    def test_container_runs_as_non_root(self):
        self.assertGreater(os.geteuid(), 0)
        self.assertEqual(pwd.getpwuid(os.geteuid()).pw_shell, "/bin/bash")

    def test_contract_helper_mount_is_read_only(self):
        helper = Path("/app/company_contract.py")
        self.assertTrue(helper.is_file())
        with self.assertRaises(OSError):
            with helper.open("r+b"):
                pass

    def test_contract_helper_loads_with_workspace_mounted_source(self):
        with patch(
            "office_backend.__file__",
            "/workspace/prototypes/libreoffice-backend/office_backend.py",
        ):
            contract = _company_contract()
        self.assertEqual(Path(contract.__file__), Path("/app/company_contract.py"))

    def test_root_execution_is_rejected(self):
        with patch("office_backend.os.geteuid", return_value=0):
            with self.assertRaisesRegex(BackendError, "refuses to run as root"):
                run_demo(Path("/tmp/root-run-must-not-create"))

    def test_date_filtered_docx_xlsx_contract_pilot(self):
        with tempfile.TemporaryDirectory() as temporary_dir:
            output = Path(temporary_dir) / "pilot"
            results = run_demo(output)

            canon = json.loads((output / "canon.json").read_text(encoding="utf-8"))
            plan = json.loads(
                (output / "build" / "plan.json").read_text(encoding="utf-8")
            )
            self.assertEqual(
                {document["format"] for document in plan["documents"]},
                {"docx", "xlsx"},
            )
            self.assertEqual(len(plan["documents"]), 2)
            for document in plan["documents"]:
                context = json.loads(
                    (
                        output
                        / "build"
                        / "contexts"
                        / f"{document['id']}.json"
                    ).read_text(encoding="utf-8")
                )
                draft = json.loads(
                    (
                        output
                        / "build"
                        / "drafts"
                        / f"{document['id']}.json"
                    ).read_text(encoding="utf-8")
                )
                self.assertEqual(context["document"], document)
                self.assertEqual(draft["document_id"], document["id"])
                self.assertEqual(draft["canon_sha256"], plan["canon_sha256"])
                visible_ids = {record["id"] for record in context["records"]}
                anchor = next(
                    record
                    for record in canon["records"]
                    if record["id"] == document["anchor_ids"][0]
                )
                future_launches = [
                    record
                    for record in canon["records"]
                    if record["kind"] == "milestone"
                    and record["data"]["product_id"]
                    == anchor["data"]["product_id"]
                    and record["data"]["status"] == "launched"
                    and record["known_on"] > document["publication_date"]
                ]
                self.assertTrue(future_launches)
                self.assertTrue(
                    all(record["id"] not in visible_ids for record in future_launches)
                )

            for document_result in results["documents"]:
                self.assertTrue((output / document_result["artifact"]).is_file())
                self.assertGreater(document_result["pdf_validation"]["pages"], 0)
                for image in document_result["visual_pages"]:
                    self.assertTrue((output / image).read_bytes().startswith(b"\x89PNG"))
            self.assertEqual(
                results["contract_validation"]["checked_drafts"], 2
            )
            self.assertTrue(results["contract_validation"]["office_package_integrity"])
            self.assertEqual(results["office_compatibility_qa"], "not performed")

            workbook_path = output / "documents" / next(
                item["filename"] for item in plan["documents"] if item["format"] == "xlsx"
            )
            recalculated = output / "build" / "renders" / (
                "recalculated-" + next(
                    item["filename"]
                    for item in plan["documents"]
                    if item["format"] == "xlsx"
                )
            )
            formulas_book = load_workbook(workbook_path, data_only=False)
            cached_book = load_workbook(recalculated, data_only=True)
            formulas = {
                f"{sheet.title}!{cell.coordinate}": cell.value
                for sheet in formulas_book.worksheets
                for row in sheet.iter_rows()
                for cell in row
                if cell.data_type == "f"
            }
            cached = {
                f"{sheet.title}!{cell.coordinate}": cell.value
                for sheet in cached_book.worksheets
                for row in sheet.iter_rows()
                for cell in row
                if cell.data_type == "f" or cell.value is not None
            }
            formulas_book.close()
            cached_book.close()
            self.assertEqual(
                formulas,
                {
                    "Validation!B3": "=COUNTA(Requirements!A3:A4)",
                    "Validation!B4": "=COUNTA('Acceptance Criteria'!B3:B6)",
                },
            )
            self.assertEqual(cached["Validation!B3"], 2)
            self.assertEqual(cached["Validation!B4"], 4)

    def test_formula_errors_are_reported_after_recalculation(self):
        with tempfile.TemporaryDirectory() as temporary_dir:
            root = Path(temporary_dir)
            source = root / "broken.xlsx"
            recalculated = root / "recalculated" / "broken.xlsx"
            workbook = Workbook()
            workbook.active["A1"] = "=1/0"
            workbook.save(source)
            workbook.close()

            recalculate_xlsx(source, recalculated)
            with self.assertRaisesRegex(BackendError, "Formula errors found"):
                inspect_recalculated_workbook(recalculated)

    def test_existing_output_directory_is_never_overwritten(self):
        with tempfile.TemporaryDirectory() as temporary_dir:
            output = Path(temporary_dir) / "existing"
            output.mkdir()
            sentinel = output / "keep.txt"
            sentinel.write_text("preserve me", encoding="utf-8")

            with self.assertRaisesRegex(BackendError, "non-empty output directory"):
                run_demo(output)

            self.assertEqual(sentinel.read_text(encoding="utf-8"), "preserve me")
            self.assertEqual(list(output.iterdir()), [sentinel])

    def test_recalculation_refuses_to_overwrite_destination(self):
        with tempfile.TemporaryDirectory() as temporary_dir:
            root = Path(temporary_dir)
            source = root / "source.xlsx"
            destination = root / "existing.xlsx"
            workbook = Workbook()
            workbook.active["A1"] = "=1+1"
            workbook.save(source)
            workbook.close()
            destination.write_bytes(b"preserve me")

            with self.assertRaisesRegex(BackendError, "Refusing to overwrite"):
                recalculate_xlsx(source, destination)

            self.assertEqual(destination.read_bytes(), b"preserve me")

    def test_malformed_pdf_is_rejected(self):
        with tempfile.TemporaryDirectory() as temporary_dir:
            path = Path(temporary_dir) / "invalid.pdf"
            path.write_bytes(b"%PDF-1.4\nnot a complete PDF")

            with self.assertRaisesRegex(BackendError, "header or end marker"):
                validate_pdf(path, ("Required text",))


if __name__ == "__main__":
    unittest.main()
