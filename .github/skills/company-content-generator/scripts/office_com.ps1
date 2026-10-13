[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("preflight")]
    [string]$Command,

    [Parameter(Mandatory = $true)]
    [string]$ProbeDirectory,

    [ValidateSet("docx", "pptx", "xlsx")]
    [string[]]$Formats = @("docx", "pptx", "xlsx"),

    [ValidateRange(15, 300)]
    [int]$TimeoutSeconds = 90
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Stop-TrackedOfficeProcess {
    param(
        [Parameter(Mandatory = $true)][string]$ProcessName,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][int[]]$BaselinePids,
        [Parameter(Mandatory = $true)][datetime]$StartedAt
    )

    $newProcesses = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
        Where-Object {
            $BaselinePids -notcontains $_.Id -and
            $_.StartTime -ge $StartedAt.AddSeconds(-5) -and
            $_.MainWindowHandle -eq 0
        })
    if ($newProcesses.Count -eq 0) {
        return
    }
    if ($newProcesses.Count -ne 1) {
        throw "Refusing to stop Office processes because the run-owned $ProcessName process is ambiguous."
    }

    Stop-Process -Id $newProcesses[0].Id -ErrorAction SilentlyContinue
    if ($null -ne (Get-Process -Id $newProcesses[0].Id -ErrorAction SilentlyContinue)) {
        throw "Could not stop run-owned $ProcessName process $($newProcesses[0].Id)."
    }
}

function Invoke-OfficeWorker {
    param(
        [Parameter(Mandatory = $true)][ValidateSet("word", "powerpoint", "excel")][string]$Mode,
        [Parameter(Mandatory = $true)][string]$ProcessName
    )

    $baselinePids = @(Get-Process -Name $ProcessName -ErrorAction SilentlyContinue |
        ForEach-Object { $_.Id })
    $oldMode = $env:JIMWORKS_OFFICE_MODE
    $oldDirectory = $env:JIMWORKS_OFFICE_PROBE_DIR
    $oldBaseline = $env:JIMWORKS_OFFICE_BASELINE_PIDS
    $oldDocxPath = $env:JIMWORKS_OFFICE_DOCX_PATH
    $oldPptxPath = $env:JIMWORKS_OFFICE_PPTX_PATH
    $oldPdfPath = $env:JIMWORKS_OFFICE_PDF_PATH
    $env:JIMWORKS_OFFICE_MODE = $Mode
    $env:JIMWORKS_OFFICE_PROBE_DIR = $script:resolvedDirectory
    $env:JIMWORKS_OFFICE_BASELINE_PIDS = $baselinePids -join ","
    $extension = switch ($Mode) {
        "word" { ".docx" }
        "powerpoint" { ".pptx" }
        "excel" { ".xlsx" }
    }
    $probePath = Join-Path $script:resolvedDirectory "$Mode-probe$extension"
    if ($Mode -eq "powerpoint") {
        $env:JIMWORKS_OFFICE_PPTX_PATH = $probePath
    } else {
        $env:JIMWORKS_OFFICE_DOCX_PATH = $probePath
    }
    $env:JIMWORKS_OFFICE_PDF_PATH = Join-Path $script:resolvedDirectory "$Mode-probe.pdf"

    $workerPath = Join-Path $PSScriptRoot "office_com_${Mode}_worker.ps1"
    $workerSource = Get-Content -LiteralPath $workerPath -Raw
    $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($workerSource))
    if ($encodedCommand.Length -gt 30000) {
        throw "The encoded $Mode worker exceeds the safe Windows command-line limit."
    }
    $startedAt = Get-Date
    $worker = $null
    $runOwnedProcessId = $null

    try {
        $worker = Start-Process `
            -FilePath "$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe" `
            -ArgumentList "-NoProfile -STA -OutputFormat Text -EncodedCommand $encodedCommand" `
            -PassThru

        if (-not $worker.WaitForExit($TimeoutSeconds * 1000)) {
            Stop-Process -Id $worker.Id
            Stop-TrackedOfficeProcess -ProcessName $ProcessName -BaselinePids $baselinePids -StartedAt $startedAt
            throw "$Mode Office preflight exceeded $TimeoutSeconds seconds; run-owned processes were stopped."
        }

        $worker.Refresh()
        if ($worker.ExitCode -ne 0) {
            $errorText = (Get-Content -LiteralPath (Join-Path $script:resolvedDirectory "preflight.log") -Raw).Trim()
            Stop-TrackedOfficeProcess -ProcessName $ProcessName -BaselinePids $baselinePids -StartedAt $startedAt
            throw "$Mode Office preflight failed with exit code $($worker.ExitCode): $errorText"
        }

        $resultPath = Join-Path $script:resolvedDirectory "$Mode-result.json"
        if (-not (Test-Path -LiteralPath $resultPath -PathType Leaf)) {
            Stop-TrackedOfficeProcess -ProcessName $ProcessName -BaselinePids $baselinePids -StartedAt $startedAt
            throw "$Mode Office worker exited without writing its evidence file."
        }
        $result = Get-Content -LiteralPath $resultPath -Raw | ConvertFrom-Json
        $runOwnedProcessId = [int]$result.process_id
        if ($runOwnedProcessId -le 0) {
            throw "Could not verify an isolated $ProcessName application process."
        }
        for ($attempt = 1; $attempt -le 50; $attempt++) {
            if ($null -eq (Get-Process -Id $runOwnedProcessId -ErrorAction SilentlyContinue)) {
                break
            }
            Start-Sleep -Milliseconds 200
        }
        if ($null -ne (Get-Process -Id $runOwnedProcessId -ErrorAction SilentlyContinue)) {
            Stop-TrackedOfficeProcess -ProcessName $ProcessName -BaselinePids $baselinePids -StartedAt $startedAt
            throw "$Mode Office process did not exit after a successful worker run."
        }

        $result
    }
    finally {
        if ($null -ne $worker -and -not $worker.HasExited) {
            Stop-Process -Id $worker.Id
        }
        $env:JIMWORKS_OFFICE_MODE = $oldMode
        $env:JIMWORKS_OFFICE_PROBE_DIR = $oldDirectory
        $env:JIMWORKS_OFFICE_BASELINE_PIDS = $oldBaseline
        $env:JIMWORKS_OFFICE_DOCX_PATH = $oldDocxPath
        $env:JIMWORKS_OFFICE_PPTX_PATH = $oldPptxPath
        $env:JIMWORKS_OFFICE_PDF_PATH = $oldPdfPath
    }
}

$script:resolvedDirectory = [IO.Path]::GetFullPath($ProbeDirectory)
if (Test-Path -LiteralPath $script:resolvedDirectory) {
    throw "Probe directory already exists; refusing to overwrite it: $script:resolvedDirectory"
}
[void][IO.Directory]::CreateDirectory($script:resolvedDirectory)

$formatWorkers = @{
    docx = @{ mode = "word"; process = "WINWORD"; app = "Microsoft Word" }
    pptx = @{ mode = "powerpoint"; process = "POWERPNT"; app = "Microsoft PowerPoint" }
    xlsx = @{ mode = "excel"; process = "EXCEL"; app = "Microsoft Excel" }
}
$selectedFormats = @($Formats | Select-Object -Unique)
if ($selectedFormats.Count -ne $Formats.Count) {
    throw "Formats must not contain duplicates."
}
if ($selectedFormats.Count -eq 0) {
    throw "At least one document format must be selected."
}
$applicationEvidence = [ordered]@{}
$workerSources = @()
$exporterWorkerSources = @()
$sourceHashes = [ordered]@{
    launcher = (Get-FileHash -LiteralPath $PSCommandPath -Algorithm SHA256).Hash
    exporter = (Get-FileHash -LiteralPath (Join-Path $PSScriptRoot "office_com_export.ps1") -Algorithm SHA256).Hash
}
foreach ($format in $selectedFormats) {
    $workerInfo = $formatWorkers[$format]
    $evidence = Invoke-OfficeWorker -Mode $workerInfo.mode -ProcessName $workerInfo.process
    if ($evidence.application -ne $workerInfo.app -or
        [string]::IsNullOrWhiteSpace([string]$evidence.version)) {
        throw "$($workerInfo.app) preflight evidence is incomplete."
    }
    switch ($format) {
        "docx" {
            if (-not $evidence.extracted_text_match -or [int]$evidence.pages -lt 1 -or
                [int]$evidence.docx.bytes -lt 1 -or [int]$evidence.rendered_pdf.bytes -lt 1) {
                throw "Word preflight did not verify authoring, extraction, pagination, and rendering."
            }
        }
        "pptx" {
            if (-not $evidence.extracted_text_match -or [int]$evidence.slide_count -lt 1 -or
                [int]$evidence.slides_rendered -ne [int]$evidence.slide_count -or
                [int]$evidence.pptx.bytes -lt 1 -or [int]$evidence.rendered_pdf.bytes -lt 1) {
                throw "PowerPoint preflight did not verify authoring, extraction, and rendering of every slide."
            }
        }
        "xlsx" {
            if ($evidence.formula -ne "=SUM(A1:B1)" -or
                [double]$evidence.reopened_cached_value -ne 5 -or
                [int]$evidence.formula_errors -ne 0 -or
                [int]$evidence.xlsx.bytes -lt 1 -or [int]$evidence.rendered_pdf.bytes -lt 1) {
                throw "Excel preflight did not verify authoring, recalculation, cached values, and rendering."
            }
        }
    }
    $applicationEvidence[$format] = $evidence
    $workerPath = Join-Path $PSScriptRoot "office_com_$($workerInfo.mode)_worker.ps1"
    $workerSources += ".github/skills/company-content-generator/scripts/office_com_$($workerInfo.mode)_worker.ps1"
    $sourceHashes["$($workerInfo.mode)_probe_worker"] = (Get-FileHash -LiteralPath $workerPath -Algorithm SHA256).Hash
    $exportWorkerName = switch ($format) {
        "docx" { "office_com_word_export_worker.ps1" }
        "pptx" { "office_com_powerpoint_export_worker.ps1" }
        "xlsx" { "office_com_excel_generic_export_worker.ps1" }
    }
    $exportWorkerPath = Join-Path $PSScriptRoot $exportWorkerName
    if (-not (Test-Path -LiteralPath $exportWorkerPath -PathType Leaf)) {
        throw "Native Office exporter is missing for $($format): $exportWorkerPath"
    }
    $exporterWorkerSources += ".github/skills/company-content-generator/scripts/$exportWorkerName"
    $sourceHashes["$($format)_export_worker"] = (Get-FileHash -LiteralPath $exportWorkerPath -Algorithm SHA256).Hash
}

$repositoryRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\..\..\.."))
$repositoryHead = (& git -C $repositoryRoot rev-parse HEAD 2>$null | Out-String).Trim()
if ($LASTEXITCODE -ne 0) {
    throw "Could not record the repository revision for the Office backend."
}

$preflight = [ordered]@{
    schema_version = 1
    status = "passed"
    recorded_at = (Get-Date).ToString("o")
    backend = [ordered]@{
        name = "Windows PowerShell STA with local Microsoft Office COM automation"
        source = ".github/skills/company-content-generator/scripts/office_com.ps1"
        selected_formats = $selectedFormats
        worker_source = $workerSources
        exporter_worker_source = $exporterWorkerSources
        revision = $repositoryHead
        source_sha256 = $sourceHashes
        provenance = "First-party local script; no external automation package or network service."
    }
    applications = $applicationEvidence
    network_access = "none"
}

$receiptPath = Join-Path $script:resolvedDirectory "preflight.json"
[IO.File]::WriteAllText(
    $receiptPath,
    ($preflight | ConvertTo-Json -Depth 8),
    [Text.UTF8Encoding]::new($false)
)
$preflight | ConvertTo-Json -Depth 8
