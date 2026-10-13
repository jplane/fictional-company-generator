$ErrorActionPreference = "Stop"

if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    throw "Excel COM automation requires a single-threaded apartment (STA)."
}

$root = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_PROBE_DIR)
$log = Join-Path $root "preflight.log"
$baselinePids = @()
if (-not [string]::IsNullOrWhiteSpace($env:JIMWORKS_OFFICE_BASELINE_PIDS)) {
    $baselinePids = @($env:JIMWORKS_OFFICE_BASELINE_PIDS -split "," | ForEach-Object { [int]$_ })
}
$app = $null
$workbook = $null
$worksheet = $null
$reopenedWorkbook = $null
$reopenedWorksheet = $null
$usedRange = $null
$cell = $null
$failure = $null
$cleanupErrors = @()

try {
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Starting isolated Excel connection."
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.EnableEvents = $false
    $name = [string]$app.Name
    $version = [string]$app.Version
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Excel connected: $name $version."

    $workbook = $app.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Cells.Item(1, 1).Value2 = 2
    $worksheet.Cells.Item(1, 2).Value2 = 3
    $worksheet.Cells.Item(1, 3).Formula = "=SUM(A1:B1)"
    $app.CalculateFullRebuild()
    if ([double]$worksheet.Cells.Item(1, 3).Value2 -ne 5) {
        throw "Excel did not calculate the probe formula to the expected value."
    }
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Excel formula recalculated."

    $xlsxPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_DOCX_PATH)
    $pdfPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_PDF_PATH)
    $workbook.SaveAs($xlsxPath, 51)
    $workbook.Save()
    $workbook.Close($false)
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet)
    $worksheet = $null
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    $workbook = $null
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Excel saved XLSX."

    $reopenedWorkbook = $app.Workbooks.Open($xlsxPath, 0, $true)
    $reopenedWorksheet = $reopenedWorkbook.Worksheets.Item(1)
    $formula = [string]$reopenedWorksheet.Cells.Item(1, 3).Formula
    $cachedValue = [double]$reopenedWorksheet.Cells.Item(1, 3).Value2
    $displayValue = [string]$reopenedWorksheet.Cells.Item(1, 3).Text
    if ($formula -ne "=SUM(A1:B1)" -or $cachedValue -ne 5 -or $displayValue -ne "5") {
        throw "Excel formula or cached value did not survive save and reopen."
    }

    $usedRange = $reopenedWorksheet.UsedRange
    $formulaErrors = @()
    for ($row = 1; $row -le $usedRange.Rows.Count; $row++) {
        for ($column = 1; $column -le $usedRange.Columns.Count; $column++) {
            $cell = $usedRange.Cells.Item($row, $column)
            if ([string]$cell.Text -match "^#(REF!|DIV/0!|VALUE!|N/A|NAME\?|NUM!|NULL!)$") {
                $formulaErrors += [string]$cell.Address($false, $false)
            }
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
            $cell = $null
        }
    }
    if ($formulaErrors.Count -gt 0) {
        throw "Excel formula errors found: $($formulaErrors -join ', ')"
    }

    $reopenedWorkbook.ExportAsFixedFormat(0, $pdfPath)
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Excel rendered PDF."
    $newProcesses = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $script:baselinePids -notcontains $_.Id })
    if ($newProcesses.Count -ne 1) {
        throw "Could not verify an isolated Excel process after rendering."
    }
    $processId = [int]$newProcesses[0].Id
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Verified isolated Excel PID $processId."
    $xlsx = Get-Item -LiteralPath $xlsxPath
    $pdf = Get-Item -LiteralPath $pdfPath
    if ($xlsx.Length -lt 1 -or $pdf.Length -lt 1) {
        throw "Excel did not create non-empty XLSX and PDF files."
    }
    $evidence = [ordered]@{
        application = $name
        version = $version
        process_id = $processId
        xlsx = [ordered]@{ path = $xlsx.FullName; bytes = $xlsx.Length }
        formula = $formula
        reopened_cached_value = $cachedValue
        formula_errors = $formulaErrors.Count
        rendered_pdf = [ordered]@{ path = $pdf.FullName; bytes = $pdf.Length }
    }
}
catch {
    $failure = $_.Exception.ToString()
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Excel preflight failed: $failure"
}
finally {
    if ($null -ne $cell) {
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($cell)
        } catch {
            $cleanupErrors += "Excel cell release failed: $($_.Exception.Message)"
        }
    }
    if ($null -ne $usedRange) {
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($usedRange)
        } catch {
            $cleanupErrors += "Excel used-range release failed: $($_.Exception.Message)"
        }
    }
    if ($null -ne $reopenedWorksheet) {
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reopenedWorksheet)
        } catch {
            $cleanupErrors += "Excel reopened worksheet release failed: $($_.Exception.Message)"
        }
    }
    if ($null -ne $reopenedWorkbook) {
        try {
            $reopenedWorkbook.Close($false)
        } catch {
            $cleanupErrors += "Excel reopened workbook close failed: $($_.Exception.Message)"
        }
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($reopenedWorkbook)
        } catch {
            $cleanupErrors += "Excel reopened workbook release failed: $($_.Exception.Message)"
        }
    }
    if ($null -ne $worksheet) {
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($worksheet)
        } catch {
            $cleanupErrors += "Excel worksheet release failed: $($_.Exception.Message)"
        }
    }
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        } catch {
            $cleanupErrors += "Excel workbook close failed: $($_.Exception.Message)"
        }
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
        } catch {
            $cleanupErrors += "Excel workbook release failed: $($_.Exception.Message)"
        }
    }
    if ($null -ne $app) {
        try {
            $app.Quit()
            Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Excel application quit."
        } catch {
            $cleanupErrors += "Excel application quit failed: $($_.Exception.Message)"
        }
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
        } catch {
            $cleanupErrors += "Excel application release failed: $($_.Exception.Message)"
        }
    }
}

if ($null -ne $failure -or $cleanupErrors.Count -gt 0) {
    $messages = @()
    if ($null -ne $failure) { $messages += $failure }
    $messages += $cleanupErrors
    [Console]::Error.WriteLine($messages -join "; ")
    exit 1
}

$resultPath = Join-Path $root "excel-result.json"
[IO.File]::WriteAllText(
    $resultPath,
    ($evidence | ConvertTo-Json -Depth 6),
    [Text.UTF8Encoding]::new($false)
)
$evidence | ConvertTo-Json -Depth 6
