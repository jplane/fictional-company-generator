$ErrorActionPreference = "Stop"

if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    throw "PowerPoint COM automation requires a single-threaded apartment (STA)."
}

$root = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_PROBE_DIR)
$pptxPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_PPTX_PATH)
$pdfPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_PDF_PATH)
$baselinePids = @()
if (-not [string]::IsNullOrWhiteSpace($env:JIMWORKS_OFFICE_BASELINE_PIDS)) {
    $baselinePids = @($env:JIMWORKS_OFFICE_BASELINE_PIDS -split "," | ForEach-Object { [int]$_ })
}
$logPath = Join-Path $root "preflight.log"
$recordPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_PROCESS_RECORD_PATH)
$previewDirectory = Join-Path $root "powerpoint-previews"
if (Test-Path -LiteralPath $previewDirectory) {
    throw "PowerPoint preview directory already exists: $previewDirectory"
}
[void][IO.Directory]::CreateDirectory($previewDirectory)

$app = $null
$presentations = $null
$presentation = $null
$slides = $null
$slide = $null
$shapes = $null
$shape = $null
$textRange = $null
$notesPage = $null
$notesPlaceholder = $null
$failure = $null
$cleanupErrors = @()
$processId = $null
$evidence = $null

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Register-RunOwnedOfficeProcess {
    param(
        [Parameter(Mandatory = $true)][string]$ProcessName,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][int[]]$BaselinePids,
        [Parameter(Mandatory = $true)][string]$RecordPath
    )
    for ($attempt = 0; $attempt -lt 50; $attempt++) {
        $newProcesses = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
            Where-Object { $BaselinePids -notcontains $_.Id })
        if ($newProcesses.Count -gt 1) {
            throw "Could not verify a unique run-owned $ProcessName process."
        }
        if ($newProcesses.Count -eq 1) {
            $startTicks = $newProcesses[0].StartTime.ToUniversalTime().Ticks
            $recordStream = [IO.File]::Open($RecordPath, [IO.FileMode]::CreateNew, [IO.FileAccess]::Write, [IO.FileShare]::None)
            try {
                $recordBytes = [Text.Encoding]::ASCII.GetBytes("$($newProcesses[0].Id),$startTicks")
                $recordStream.Write($recordBytes, 0, $recordBytes.Length)
            } finally {
                $recordStream.Dispose()
            }
            return [int]$newProcesses[0].Id
        }
        Start-Sleep -Milliseconds 100
    }
    throw "Could not verify a run-owned $ProcessName process."
}

try {
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Starting isolated PowerPoint connection."
    $app = New-Object -ComObject PowerPoint.Application
    $processId = Register-RunOwnedOfficeProcess -ProcessName POWERPNT -BaselinePids $baselinePids -RecordPath $recordPath
    $app.Visible = -1
    $app.WindowState = 2
    $presentations = $app.Presentations
    $presentation = $presentations.Add()
    $slides = $presentation.Slides
    $slide = $slides.Add(1, 12)
    $shapes = $slide.Shapes
    $shape = $shapes.AddTextbox(1, 72, 72, 600, 96)
    $textRange = $shape.TextFrame.TextRange
    $textRange.Text = "Native PowerPoint automation probe"
    $notesPage = $slide.NotesPage
    $notesPlaceholder = $notesPage.Shapes.Placeholders.Item(2)
    $notesPlaceholder.TextFrame.TextRange.Text = "Notes for the PowerPoint automation probe"
    $presentation.SaveAs($pptxPath, 24)
    $presentation.Close()
    foreach ($object in @($notesPlaceholder, $notesPage, $textRange, $shape, $shapes, $slide, $slides, $presentation)) {
        if ($null -ne $object) {
            Release-ComObject $object
        }
    }
    $notesPlaceholder = $null
    $notesPage = $null
    $textRange = $null
    $shape = $null
    $shapes = $null
    $slide = $null
    $slides = $null
    $presentation = $presentations.Open($pptxPath, -1, 0, 0)
    $slides = $presentation.Slides
    $slide = $slides.Item(1)
    $textRange = $slide.Shapes.Item(1).TextFrame.TextRange
    $notesPage = $slide.NotesPage
    $notesPlaceholder = $notesPage.Shapes.Placeholders.Item(2)
    $extractedText = [string]$textRange.Text
    $extractedNotes = [string]$notesPlaceholder.TextFrame.TextRange.Text
    if ($extractedText -ne "Native PowerPoint automation probe" -or
        -not $extractedNotes.Contains("Notes for the PowerPoint automation probe")) {
        throw "PowerPoint content or notes extraction after reopen did not match the probe text."
    }
    $presentation.SaveAs($pdfPath, 32)
    $presentation.Export($previewDirectory, "PNG", 1280, 720)

    if ($null -eq (Get-Process -Id $processId -ErrorAction SilentlyContinue)) {
        throw "The run-owned PowerPoint process exited before preflight verification."
    }
    $pptx = Get-Item -LiteralPath $pptxPath
    $pdf = Get-Item -LiteralPath $pdfPath
    $previews = @(Get-ChildItem -LiteralPath $previewDirectory -File -Filter "*.png")
    if ($pptx.Length -lt 1 -or $pdf.Length -lt 1 -or $previews.Count -ne 1 -or $previews[0].Length -lt 1) {
        throw "PowerPoint did not create the expected PPTX, PDF, and slide preview."
    }
    $evidence = [ordered]@{
        application = [string]$app.Name
        version = [string]$app.Version
        process_id = $processId
        pptx = [ordered]@{ path = $pptx.FullName; bytes = $pptx.Length }
        extracted_text_match = $true
        slide_count = [int]$slides.Count
        slides_rendered = $previews.Count
        preview_paths = @($previews | ForEach-Object { $_.FullName })
        rendered_pdf = [ordered]@{ path = $pdf.FullName; bytes = $pdf.Length }
    }
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) PowerPoint saved, extracted, and rendered one slide."
}
catch {
    $failure = $_.Exception.ToString()
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) PowerPoint preflight failed: $failure"
}
finally {
    if ($null -ne $presentation) {
        try {
            $presentation.Close()
            Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) PowerPoint presentation closed."
        } catch {
            $cleanupErrors += "PowerPoint presentation close failed: $($_.Exception.Message)"
        }
    }
    foreach ($object in @($notesPlaceholder, $notesPage, $textRange, $shape, $shapes, $slide, $slides, $presentation, $presentations)) {
        if ($null -ne $object) {
            try {
                Release-ComObject $object
            } catch {
                $cleanupErrors += "PowerPoint COM object release failed: $($_.Exception.Message)"
            }
        }
    }
    if ($null -ne $app -and $null -ne $processId) {
        try {
            $app.Quit()
            Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) PowerPoint application quit."
        } catch {
            $cleanupErrors += "PowerPoint application quit failed: $($_.Exception.Message)"
        }
    } elseif ($null -ne $app) {
        $cleanupErrors += "PowerPoint process identity was not verified before quit; launcher cleanup is required."
    }
    if ($null -ne $app) {
        try {
            Release-ComObject $app
        } catch {
            $cleanupErrors += "PowerPoint application release failed: $($_.Exception.Message)"
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($null -ne $failure -or $cleanupErrors.Count -gt 0) {
    $messages = @()
    if ($null -ne $failure) { $messages += $failure }
    $messages += $cleanupErrors
    [Console]::Error.WriteLine($messages -join "; ")
    exit 1
}

[IO.File]::WriteAllText(
    (Join-Path $root "powerpoint-result.json"),
    ($evidence | ConvertTo-Json -Depth 6),
    [Text.UTF8Encoding]::new($false)
)
$evidence | ConvertTo-Json -Depth 6
