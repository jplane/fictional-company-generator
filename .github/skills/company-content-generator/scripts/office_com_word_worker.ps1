$ErrorActionPreference = "Stop"

if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    throw "Word COM automation requires a single-threaded apartment (STA)."
}

$root = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_PROBE_DIR)
$log = Join-Path $root "preflight.log"
$baselinePids = @()
if (-not [string]::IsNullOrWhiteSpace($env:JIMWORKS_OFFICE_BASELINE_PIDS)) {
    $baselinePids = @($env:JIMWORKS_OFFICE_BASELINE_PIDS -split "," | ForEach-Object { [int]$_ })
}
$app = $null
$document = $null
$failure = $null
$cleanupErrors = @()

function Close-WordDocument {
    param([Parameter(Mandatory = $true)][object]$Document)
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            $Document.Close(0)
            return
        } catch {
            if ($_.Exception.HResult -ne -2147418111 -or $attempt -eq 5) {
                throw
            }
            Start-Sleep -Milliseconds 250
        }
    }
}

try {
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Starting isolated Word connection."
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0
    $name = [string]$app.Name
    $version = [string]$app.Version
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Word connected: $name $version."

    $document = $app.Documents.Add()
    $document.Content.Text = "Jimworks Office automation probe"
    $docxPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_DOCX_PATH)
    $pdfPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_PDF_PATH)
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Word save started."
    $document.SaveAs2($docxPath, 16)
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Word saved DOCX."
    Start-Sleep -Milliseconds 500
    Close-WordDocument -Document $document
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
    $document = $app.Documents.Open($docxPath, $false, $true, $false)
    $extractedText = ([string]$document.Content.Text).Trim([char[]]@(13, 10, 7))
    if ($extractedText -ne "Jimworks Office automation probe") {
        throw "Word content extraction after reopen did not match the probe text: '$extractedText'."
    }
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Word reopened DOCX and extracted content."
    $document.Repaginate()
    $pages = [int]$document.ComputeStatistics(2)
    if ($pages -lt 1) {
        throw "Word returned an invalid page count."
    }
    $document.ExportAsFixedFormat($pdfPath, 17)
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Word rendered PDF."

    $newProcesses = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
        Where-Object { $script:baselinePids -notcontains $_.Id })
    if ($newProcesses.Count -ne 1) {
        throw "Could not verify an isolated Word process after rendering."
    }
    $processId = [int]$newProcesses[0].Id
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Verified isolated Word PID $processId."

    $docx = Get-Item -LiteralPath $docxPath
    $pdf = Get-Item -LiteralPath $pdfPath
    if ($docx.Length -lt 1 -or $pdf.Length -lt 1) {
        throw "Word did not create non-empty DOCX and PDF files."
    }
    $evidence = [ordered]@{
        application = $name
        version = $version
        process_id = $processId
        docx = [ordered]@{ path = $docx.FullName; bytes = $docx.Length }
        extracted_text_match = $true
        pages = $pages
        rendered_pdf = [ordered]@{ path = $pdf.FullName; bytes = $pdf.Length }
    }
}
catch {
    $failure = @(
        $_.Exception.ToString()
        $_.InvocationInfo.PositionMessage
        $_.ScriptStackTrace
    ) -join "`n"
    Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Word preflight failed: $failure"
}
finally {
    if ($null -ne $document) {
        try {
            Close-WordDocument -Document $document
            Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Word document closed."
        } catch {
            $cleanupErrors += "Word document close failed: $($_.Exception.Message)"
        }
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($document)
        } catch {
            $cleanupErrors += "Word document release failed: $($_.Exception.Message)"
        }
    }
    if ($null -ne $app) {
        try {
            $app.Quit()
            Add-Content -LiteralPath $log -Value "$(Get-Date -Format o) Word application quit."
        } catch {
            $cleanupErrors += "Word application quit failed: $($_.Exception.Message)"
        }
        try {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($app)
        } catch {
            $cleanupErrors += "Word application release failed: $($_.Exception.Message)"
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

$resultPath = Join-Path $root "word-result.json"
[IO.File]::WriteAllText(
    $resultPath,
    ($evidence | ConvertTo-Json -Depth 6),
    [Text.UTF8Encoding]::new($false)
)
$evidence | ConvertTo-Json -Depth 6
