import os
import shutil
import subprocess
import tempfile
import unittest
from pathlib import Path

import test_company as baseline

company = baseline.company


OFFICE_SCRIPT = r"""
$ErrorActionPreference = 'Stop'
$format = $env:PROVENANCE_FORMAT
$processName = @{docx='WINWORD'; pptx='POWERPNT'; xlsx='EXCEL'}[$format]
$progId = @{docx='Word.Application'; pptx='PowerPoint.Application'; xlsx='Excel.Application'}[$format]
if (Get-Process -Name $processName -ErrorAction SilentlyContinue) {
    Write-Output "SKIP: $processName already running; existing sessions are not used"
    exit 77
}
if ($null -eq [Type]::GetTypeFromProgID($progId)) {
    Write-Output "SKIP: $progId is not installed"
    exit 77
}
$path = Join-Path $env:PROVENANCE_ROOT "content.$format"
$app = $null
$document = $null
try {
    $app = New-Object -ComObject $progId
    if ($format -eq 'docx') {
        $app.Visible = $false
        $document = $app.Documents.Add()
        $document.Content.Text = "Revenue 12.`r"
        $range = $document.Range($document.Content.End - 1, $document.Content.End - 1)
        $table = $document.Tables.Add($range, 1, 1)
        $table.Cell(1, 1).Range.Text = '12'
        $document.Sections.Item(1).Headers.Item(1).Range.Text = 'Operating results'
        $document.SaveAs2([string]$path, 16)
        $document.Close(0)
        $document = $app.Documents.Open($path, $false, $true)
        if ($document.Tables.Count -ne 1) { throw 'Word reopen lost the table' }
        $document.Close(0)
        $document = $null
    } elseif ($format -eq 'pptx') {
        $document = $app.Presentations.Add(0)
        $slide = $document.Slides.Add(1, 12)
        $shape = $slide.Shapes.AddTextbox(1, 36, 36, 480, 60)
        $shape.TextFrame.TextRange.Text = 'Revenue 12'
        $table = $slide.Shapes.AddTable(1, 1, 36, 120, 300, 40)
        $table.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = '12'
        $note = $slide.NotesPage.Shapes.AddTextbox(1, 36, 120, 300, 40)
        $note.TextFrame.TextRange.Text = 'Revenue reflects saved actuals'
        $document.SaveAs($path, 24)
        $document.Close()
        $document = $app.Presentations.Open($path, -1, 0, 0)
        if ($document.Slides.Count -ne 1) { throw 'PowerPoint reopen lost the slide' }
        $document.Close()
        $document = $null
    } else {
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $document = $app.Workbooks.Add()
        $sheet = $document.Worksheets.Item(1)
        $sheet.Name = 'Actuals'
        $sheet.Range('A1').Value2 = 'Revenue'
        $sheet.Range('A2').Value2 = 5
        $sheet.Range('A3').Value2 = 7
        $sheet.Range('A4').Value2 = 9
        $sheet.Range('B1').Formula = '=SUM(A2:A3)'
        $sheet.Range('B1:B2').FillDown()
        $chart = $sheet.ChartObjects().Add(200, 20, 300, 180)
        $chart.Chart.SetSourceData($sheet.Range('A2:A3'))
        $app.CalculateFullRebuild()
        if ($sheet.Range('B1').Value2 -ne 12) { throw 'Excel calculated the wrong result' }
        $document.SaveAs($path, 51)
        $document.Close($false)
        $document = $app.Workbooks.Open($path, 0, $true)
        if ($document.Worksheets.Item(1).Range('B1').Value2 -ne 12) {
            throw 'Excel reopen lost the calculated result'
        }
        $document.Close($false)
        $document = $null
    }
    Write-Output "Verified $progId $($app.Version): save and reopen"
} finally {
    if ($null -ne $document) {
        if ($format -eq 'docx') { $document.Close(0) }
        elseif ($format -eq 'pptx') { $document.Close() }
        else { $document.Close($false) }
    }
    if ($null -ne $app) {
        $remaining = if ($format -eq 'docx') { $app.Documents.Count }
                     elseif ($format -eq 'pptx') { $app.Presentations.Count }
                     else { $app.Workbooks.Count }
        if ($remaining -eq 0) { $app.Quit() }
        else { throw 'Another document appeared; refusing to quit the application' }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
"""


@unittest.skipUnless(os.environ.get("RUN_OFFICE_PROVENANCE_TESTS") == "1",
                     "Set RUN_OFFICE_PROVENANCE_TESTS=1 for local Microsoft Office integration")
class OfficeProvenanceTests(unittest.TestCase):
    def create(self, root, fmt):
        if os.name != "nt" or not shutil.which("powershell.exe"):
            self.skipTest("Windows PowerShell and local Office COM are required")
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-Command", OFFICE_SCRIPT],
            env=dict(os.environ, PROVENANCE_ROOT=str(root), PROVENANCE_FORMAT=fmt),
            capture_output=True, text=True, timeout=120,
        )
        if result.returncode == 77:
            self.skipTest(result.stdout.strip())
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        path = root / ("content." + fmt)
        self.assertTrue(path.is_file())
        return company.artifact_containers(path, fmt)

    def test_word_saved_paragraph_table_and_header(self):
        with tempfile.TemporaryDirectory() as directory:
            entries = self.create(Path(directory), "docx")
            self.assertTrue(any(entry["locator"]["kind"] == "paragraph"
                                and entry["text"] == "Revenue 12." for entry in entries))
            self.assertTrue(any(entry["locator"]["kind"] == "table_cell"
                                and entry["text"] == "12" for entry in entries))
            self.assertTrue(any("header" in entry["locator"]["part"]
                                and entry["text"] == "Operating results" for entry in entries))

    def test_powerpoint_saved_shape_table_and_notes(self):
        with tempfile.TemporaryDirectory() as directory:
            entries = self.create(Path(directory), "pptx")
            self.assertTrue(any(entry["locator"]["kind"] == "shape"
                                and entry["text"] == "Revenue 12" for entry in entries))
            self.assertTrue(any(entry["locator"]["kind"] == "slide_cell"
                                and entry["text"] == "12" for entry in entries))
            self.assertTrue(any(entry["locator"]["kind"] == "notes_shape"
                                and entry["text"] == "Revenue reflects saved actuals" for entry in entries))

    def test_excel_saved_formula_cache_range_and_chart(self):
        with tempfile.TemporaryDirectory() as directory:
            entries = self.create(Path(directory), "xlsx")
            cell = next(entry for entry in entries if entry["locator"].get("cell") == "B1"
                        and entry["locator"].get("sheet") == "Actuals")
            self.assertEqual(cell["text"], "12")
            self.assertEqual(cell["formula"], "SUM(A2:A3)")
            follower = next(entry for entry in entries if entry["locator"].get("cell") == "B2"
                            and entry["locator"].get("sheet") == "Actuals")
            self.assertEqual(follower["text"], "16")
            self.assertTrue(follower["formula"])
            locator = dict(cell["locator"], kind="range", range="A2:A3")
            locator.pop("cell")
            self.assertEqual(company.located_container(entries, locator)["text"], "5 7")
            self.assertTrue(any(entry["locator"]["kind"] == "chart_point"
                                and entry["text"] in ("5", "7") for entry in entries))


if __name__ == "__main__":
    unittest.main()
