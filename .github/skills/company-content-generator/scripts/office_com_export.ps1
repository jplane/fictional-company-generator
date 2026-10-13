[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet("word", "excel", "powerpoint")]
    [string]$Mode,

    [Parameter(Mandatory = $true)]
    [string]$ContextPath,

    [Parameter(Mandatory = $true)]
    [string]$DraftPath,

    [Parameter(Mandatory = $true)]
    [string]$OutputPath,

    [Parameter(Mandatory = $true)]
    [string]$PdfPath,

    [string]$PreviewDirectory,

    [Parameter(Mandatory = $true)]
    [string]$ExtractedTextPath,

    [Parameter(Mandatory = $true)]
    [string]$ResultPath,

    [Parameter(Mandatory = $true)]
    [string]$LogPath,

    [ValidateRange(30, 600)]
    [int]$TimeoutSeconds = 180
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
    $runOwnedId = [int]$newProcesses[0].Id
    Stop-Process -Id $runOwnedId -ErrorAction SilentlyContinue
    if ($null -ne (Get-Process -Id $runOwnedId -ErrorAction SilentlyContinue)) {
        throw "Could not stop run-owned $ProcessName process $runOwnedId."
    }
}

$resolvedDraft = [IO.Path]::GetFullPath($DraftPath)
$resolvedContext = [IO.Path]::GetFullPath($ContextPath)
$resolvedOutput = [IO.Path]::GetFullPath($OutputPath)
$resolvedPdf = [IO.Path]::GetFullPath($PdfPath)
$resolvedPreviewDirectory = if ([string]::IsNullOrWhiteSpace($PreviewDirectory)) {
    $null
} else {
    [IO.Path]::GetFullPath($PreviewDirectory)
}
$resolvedText = [IO.Path]::GetFullPath($ExtractedTextPath)
$resolvedResult = [IO.Path]::GetFullPath($ResultPath)
$resolvedLog = [IO.Path]::GetFullPath($LogPath)
foreach ($parentPath in @(
    [IO.Path]::GetDirectoryName($resolvedOutput),
    [IO.Path]::GetDirectoryName($resolvedPdf),
    [IO.Path]::GetDirectoryName($resolvedText),
    [IO.Path]::GetDirectoryName($resolvedResult),
    [IO.Path]::GetDirectoryName($resolvedLog)
)) {
    [void][IO.Directory]::CreateDirectory($parentPath)
}
foreach ($output in @($resolvedOutput, $resolvedPdf, $resolvedText, $resolvedResult)) {
    if (Test-Path -LiteralPath $output) {
        throw "Refusing to overwrite an existing Office output: $output"
    }
}
if (-not (Test-Path -LiteralPath $resolvedDraft -PathType Leaf)) {
    throw "Draft does not exist: $resolvedDraft"
}
if (-not (Test-Path -LiteralPath $resolvedContext -PathType Leaf)) {
    throw "Date-filtered context does not exist: $resolvedContext"
}
if ($Mode -eq "powerpoint" -and $null -eq $resolvedPreviewDirectory) {
    throw "PowerPoint export requires a PreviewDirectory for per-slide render evidence."
}

$processName = switch ($Mode) {
    "word" { "WINWORD" }
    "excel" { "EXCEL" }
    "powerpoint" { "POWERPNT" }
}
$baselinePids = @(Get-Process -Name $processName -ErrorAction SilentlyContinue |
    ForEach-Object { [int]$_.Id })
$environmentNames = @(
    "JIMWORKS_OFFICE_EXPORT_CONTEXT_PATH",
    "JIMWORKS_OFFICE_EXPORT_DRAFT_PATH",
    "JIMWORKS_OFFICE_EXPORT_OUTPUT_PATH",
    "JIMWORKS_OFFICE_EXPORT_PDF_PATH",
    "JIMWORKS_OFFICE_EXPORT_PREVIEW_DIRECTORY",
    "JIMWORKS_OFFICE_EXPORT_TEXT_PATH",
    "JIMWORKS_OFFICE_EXPORT_RESULT_PATH",
    "JIMWORKS_OFFICE_EXPORT_LOG_PATH",
    "JIMWORKS_OFFICE_EXPORT_BASELINE_PIDS"
)
$oldEnvironment = @{}
foreach ($name in $environmentNames) {
    $oldEnvironment[$name] = [Environment]::GetEnvironmentVariable($name, "Process")
}
$env:JIMWORKS_OFFICE_EXPORT_CONTEXT_PATH = $resolvedContext
$env:JIMWORKS_OFFICE_EXPORT_DRAFT_PATH = $resolvedDraft
$env:JIMWORKS_OFFICE_EXPORT_OUTPUT_PATH = $resolvedOutput
$env:JIMWORKS_OFFICE_EXPORT_PDF_PATH = $resolvedPdf
$env:JIMWORKS_OFFICE_EXPORT_PREVIEW_DIRECTORY = $resolvedPreviewDirectory
$env:JIMWORKS_OFFICE_EXPORT_TEXT_PATH = $resolvedText
$env:JIMWORKS_OFFICE_EXPORT_RESULT_PATH = $resolvedResult
$env:JIMWORKS_OFFICE_EXPORT_LOG_PATH = $resolvedLog
$env:JIMWORKS_OFFICE_EXPORT_BASELINE_PIDS = $baselinePids -join ","

$workerFile = switch ($Mode) {
    "word" { "office_com_word_export_worker.ps1" }
    "excel" { "office_com_excel_generic_export_worker.ps1" }
    "powerpoint" { "office_com_powerpoint_export_worker.ps1" }
}
$workerPath = Join-Path $PSScriptRoot $workerFile
if (-not (Test-Path -LiteralPath $workerPath -PathType Leaf)) {
    throw "Office export worker is missing: $workerPath"
}
$startedAt = Get-Date
$worker = $null
try {
    Add-Content -LiteralPath $resolvedLog -Value "$(Get-Date -Format o) Starting $Mode Office export."
    $worker = Start-Process `
        -FilePath "$env:windir\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -ArgumentList "-NoProfile -STA -OutputFormat Text -File `"$workerPath`"" `
        -PassThru `
        -WindowStyle Hidden

    if (-not $worker.WaitForExit($TimeoutSeconds * 1000)) {
        Stop-Process -Id $worker.Id
        Stop-TrackedOfficeProcess -ProcessName $processName -BaselinePids $baselinePids -StartedAt $startedAt
        throw "$Mode Office export exceeded $TimeoutSeconds seconds; run-owned processes were stopped."
    }

    $worker.Refresh()
    if ($worker.ExitCode -ne 0) {
        $logText = if (Test-Path -LiteralPath $resolvedLog) {
            (Get-Content -LiteralPath $resolvedLog -Raw).Trim()
        } else {
            "No worker log was written."
        }
        Stop-TrackedOfficeProcess -ProcessName $processName -BaselinePids $baselinePids -StartedAt $startedAt
        throw "$Mode Office export failed with exit code $($worker.ExitCode): $logText"
    }
    if (-not (Test-Path -LiteralPath $resolvedResult -PathType Leaf)) {
        Stop-TrackedOfficeProcess -ProcessName $processName -BaselinePids $baselinePids -StartedAt $startedAt
        throw "$Mode Office worker exited without writing its evidence file."
    }

    $result = Get-Content -LiteralPath $resolvedResult -Raw | ConvertFrom-Json
    $runOwnedProcessId = [int]$result.process_id
    if ($runOwnedProcessId -le 0) {
        throw "Could not verify an isolated $processName application process."
    }
    for ($attempt = 1; $attempt -le 50; $attempt++) {
        if ($null -eq (Get-Process -Id $runOwnedProcessId -ErrorAction SilentlyContinue)) {
            break
        }
        Start-Sleep -Milliseconds 200
    }
    if ($null -ne (Get-Process -Id $runOwnedProcessId -ErrorAction SilentlyContinue)) {
        Stop-TrackedOfficeProcess -ProcessName $processName -BaselinePids $baselinePids -StartedAt $startedAt
        throw "$Mode Office process did not exit after a successful export."
    }
    $result | ConvertTo-Json -Depth 10
}
finally {
    if ($null -ne $worker -and -not $worker.HasExited) {
        Stop-Process -Id $worker.Id
    }
    foreach ($name in $environmentNames) {
        [Environment]::SetEnvironmentVariable($name, $oldEnvironment[$name], "Process")
    }
}
