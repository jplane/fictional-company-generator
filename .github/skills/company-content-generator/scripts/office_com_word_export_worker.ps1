$ErrorActionPreference = "Stop"

if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    throw "Word COM automation requires a single-threaded apartment (STA)."
}

$draftPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_DRAFT_PATH)
$contextPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_CONTEXT_PATH)
$outputPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_OUTPUT_PATH)
$pdfPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_PDF_PATH)
$textPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_TEXT_PATH)
$resultPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_RESULT_PATH)
$logPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_LOG_PATH)
$baselinePids = @()
if (-not [string]::IsNullOrWhiteSpace($env:JIMWORKS_OFFICE_EXPORT_BASELINE_PIDS)) {
    $baselinePids = @($env:JIMWORKS_OFFICE_EXPORT_BASELINE_PIDS -split "," | ForEach-Object { [int]$_ })
}

foreach ($path in @($outputPath, $pdfPath, $textPath, $resultPath)) {
    if (Test-Path -LiteralPath $path) {
        throw "Refusing to overwrite existing output: $path"
    }
}
$draft = Get-Content -LiteralPath $draftPath -Raw | ConvertFrom-Json
$context = Get-Content -LiteralPath $contextPath -Raw | ConvertFrom-Json
$companyName = [string]$context.company.data.name
if ([string]::IsNullOrWhiteSpace($companyName)) {
    throw "Date-filtered context does not contain a company name."
}
$requirementRows = @()
if ($null -ne $draft.requirement_rows) {
    $requirementRows = @($draft.requirement_rows)
}
$app = $null
$document = $null
$documentProperties = $null
$selection = $null
$section = $null
$headerRange = $null
$footerRange = $null
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

function Get-OptionalArray {
    param([object]$Object, [string]$Name)
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return @() }
    @($property.Value)
}

function Set-OfficeDocumentProperty {
    param(
        [Parameter(Mandatory = $true)][object]$Properties,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Value
    )
    $property = $null
    try {
        $objectType = [object].GetMethod("GetType").Invoke($Properties, $null)
        $itemFlags = [Reflection.BindingFlags]::Instance -bor
            [Reflection.BindingFlags]::Public -bor
            [Reflection.BindingFlags]::GetProperty -bor
            [Reflection.BindingFlags]::InvokeMethod
        $property = $objectType.InvokeMember("Item", $itemFlags, $null, $Properties, @($Name))
        $propertyType = [object].GetMethod("GetType").Invoke($property, $null)
        $setFlags = [Reflection.BindingFlags]::Instance -bor
            [Reflection.BindingFlags]::Public -bor
            [Reflection.BindingFlags]::SetProperty
        [void]$propertyType.InvokeMember("Value", $setFlags, $null, $property, @($Value))
        $getFlags = [Reflection.BindingFlags]::Instance -bor
            [Reflection.BindingFlags]::Public -bor
            [Reflection.BindingFlags]::GetProperty
        $actual = [string]$propertyType.InvokeMember("Value", $getFlags, $null, $property, $null)
        if ($actual -ne $Value) {
            throw "Office did not retain the $Name document property."
        }
    } finally {
        Release-ComObject $property
    }
}

function Add-Paragraph {
    param(
        [Parameter(Mandatory = $true)][object]$Selection,
        [Parameter(Mandatory = $true)][object]$Document,
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][string]$Style
    )
    $Selection.Style = $Style
    if ($Style -eq "Normal") {
        $Selection.ParagraphFormat.KeepTogether = $true
    } else {
        $Selection.ParagraphFormat.KeepTogether = $false
    }
    $Selection.TypeText($Text)
    $Selection.TypeParagraph()
}

function Add-WordTable {
    param(
        [Parameter(Mandatory = $true)][object]$Selection,
        [Parameter(Mandatory = $true)][object]$Document,
        [Parameter(Mandatory = $true)][object]$Specification,
        [switch]$AddTrailingParagraph
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$Specification.title)) {
        Add-Paragraph -Selection $Selection -Document $Document -Text ([string]$Specification.title) -Style "Heading 2"
    }
    $rowCount = 1 + $Specification.rows.Count
    $columnCount = $Specification.headers.Count
    $table = $null
    $range = $null
    $headerRow = $null
    $rowObject = $null
    $rowRange = $null
    try {
        $table = $Document.Tables.Add($Selection.Range, $rowCount, $columnCount)
        $table.Style = "Table Grid"
        $range = $table.Range
        $range.Font.Name = "Aptos"
        $range.Font.Size = 8
        $range.ParagraphFormat.SpaceAfter = 0
        $range.Cells.VerticalAlignment = 0
        for ($column = 1; $column -le $columnCount; $column++) {
            $cell = $table.Cell(1, $column).Range
            try {
                $cell.Text = [string]$Specification.headers[$column - 1]
            } finally {
                Release-ComObject $cell
            }
        }
        for ($row = 0; $row -lt $Specification.rows.Count; $row++) {
            for ($column = 0; $column -lt $columnCount; $column++) {
                $cell = $table.Cell($row + 2, $column + 1).Range
                try {
                    $cell.Text = [string]$Specification.rows[$row][$column]
                } finally {
                    Release-ComObject $cell
                }
            }
        }
        $headerRow = $table.Rows.Item(1).Range
        $headerRow.Bold = $true
        $headerRow.Shading.BackgroundPatternColor = 15132390
        $headerRow.Rows.HeadingFormat = $true
        for ($row = 1; $row -le $rowCount; $row++) {
            $rowObject = $table.Rows.Item($row)
            try {
                $rowObject.AllowBreakAcrossPages = $false
                $rowRange = $rowObject.Range
                $rowRange.ParagraphFormat.KeepWithNext = $row -lt $rowCount
            } finally {
                Release-ComObject $rowRange
                Release-ComObject $rowObject
                $rowRange = $null
                $rowObject = $null
            }
        }
        $table.AutoFitBehavior(2)
        $selection.SetRange($table.Range.End, $table.Range.End)
        if ($AddTrailingParagraph) {
            $selection.TypeParagraph()
        }
    } finally {
        Release-ComObject $headerRow
        Release-ComObject $range
        Release-ComObject $table
        Release-ComObject $rowRange
        Release-ComObject $rowObject
    }
}

try {
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Starting Word authoring for $($draft.document_id)."
    $app = New-Object -ComObject Word.Application
    $app.Visible = $false
    $app.DisplayAlerts = 0
    $app.ScreenUpdating = $false
    $app.Options.CheckSpellingAsYouType = $false
    $app.Options.CheckGrammarAsYouType = $false
    $document = $app.Documents.Add()
    $document.PageSetup.TopMargin = $app.InchesToPoints(0.7)
    $document.PageSetup.BottomMargin = $app.InchesToPoints(0.65)
    $document.PageSetup.LeftMargin = $app.InchesToPoints(0.7)
    $document.PageSetup.RightMargin = $app.InchesToPoints(0.7)
    $documentProperties = $document.BuiltinDocumentProperties
    Set-OfficeDocumentProperty -Properties $documentProperties -Name "Title" -Value ([string]$draft.title)
    Set-OfficeDocumentProperty -Properties $documentProperties -Name "Subject" -Value ([string]$draft.publication_date)
    Set-OfficeDocumentProperty -Properties $documentProperties -Name "Author" -Value $companyName
    Release-ComObject $documentProperties
    $documentProperties = $null
    $selection = $app.Selection

    $selection.Style = "Title"
    $selection.Font.Name = "Aptos Display"
    $selection.Font.Size = 24
    $selection.Font.Color = 5197615
    $selection.TypeText([string]$draft.title)
    $selection.TypeParagraph()
    Add-Paragraph -Selection $selection -Document $document -Text "Published $($draft.publication_date)" -Style "Subtitle"

    $section = $document.Sections.Item(1)
    $headerRange = $section.Headers.Item(1).Range
    $headerRange.Text = "$companyName  |  $($draft.title)"
    $headerRange.Font.Name = "Aptos"
    $headerRange.Font.Size = 8
    $headerRange.Font.Bold = $true
    $headerRange.ParagraphFormat.Alignment = 2
    $footerRange = $section.Footers.Item(1).Range
    $footerRange.Text = "$companyName  |  $($draft.publication_date)"
    $footerRange.Font.Name = "Aptos"
    $footerRange.Font.Size = 8
    $footerRange.ParagraphFormat.Alignment = 1

    $sectionIndex = 0
    foreach ($contentSection in $draft.sections) {
        $tableSpecs = @(Get-OptionalArray -Object $contentSection -Name "tables")
        Add-Paragraph -Selection $selection -Document $document -Text ([string]$contentSection.heading) -Style "Heading 1"
        $paragraphs = ([string]$contentSection.text) -split "\r?\n\s*\r?\n"
        foreach ($paragraph in $paragraphs) {
            $trimmed = $paragraph.Trim()
            if ($trimmed.Length -gt 0) {
                Add-Paragraph -Selection $selection -Document $document -Text $trimmed -Style "Normal"
            }
        }
        $tableIndex = 0
        foreach ($tableSpec in $tableSpecs) {
            $addTrailingParagraph = $sectionIndex -lt ($draft.sections.Count - 1) -or
                $tableIndex -lt ($tableSpecs.Count - 1)
            Add-WordTable -Selection $selection -Document $document `
                -Specification $tableSpec -AddTrailingParagraph:$addTrailingParagraph
            $tableIndex++
        }
        $sectionIndex++
    }
    if ($requirementRows.Count -gt 0) {
        Add-Paragraph -Selection $selection -Document $document -Text "Requirement traceability" -Style "Heading 1"
        $headers = @($requirementRows | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique)
        $rows = @()
        foreach ($requirement in $requirementRows) {
            $row = @()
            foreach ($header in $headers) {
                $value = $requirement.$header
                if ($null -eq $value) {
                    $row += ""
                } elseif ($value -is [System.Collections.IDictionary] -or $value -is [System.Management.Automation.PSCustomObject] -or $value -is [array]) {
                    $row += ($value | ConvertTo-Json -Depth 20 -Compress)
                } else {
                    $row += [string]$value
                }
            }
            $rows += ,$row
        }
        Add-WordTable -Selection $selection -Document $document `
            -Specification ([pscustomobject]@{
                title = ""
                headers = $headers
                rows = $rows
            }) -AddTrailingParagraph
    }

    $document.SaveAs2($outputPath, 16)
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Word saved DOCX."
    $document.Close(0)
    foreach ($object in @($footerRange, $headerRange, $section, $selection, $documentProperties, $document)) {
        Release-ComObject $object
    }
    $footerRange = $null
    $headerRange = $null
    $section = $null
    $selection = $null
    $documentProperties = $null
    $document = $app.Documents.Open($outputPath, $false, $true, $false)
    $document.Repaginate()
    $pages = [int]$document.ComputeStatistics(2)
    if ($pages -lt 1) {
        throw "Word returned an invalid page count."
    }
    $document.ExportAsFixedFormat($pdfPath, 17)
    $extractedText = [string]$document.Content.Text
    [IO.File]::WriteAllText($textPath, $extractedText, [Text.UTF8Encoding]::new($false))

    $expected = @([string]$draft.title, [string]$draft.publication_date, $companyName)
    foreach ($contentSection in $draft.sections) {
        $tableSpecs = @(Get-OptionalArray -Object $contentSection -Name "tables")
        $expected += [string]$contentSection.heading
        foreach ($paragraph in ([string]$contentSection.text -split "\r?\n\s*\r?\n")) {
            $trimmed = $paragraph.Trim()
            if ($trimmed.Length -gt 0) {
                $expected += $trimmed
            }
        }
        foreach ($tableSpec in $tableSpecs) {
            $expected += @($tableSpec.headers | ForEach-Object { [string]$_ })
            foreach ($row in $tableSpec.rows) {
                $expected += @($row | ForEach-Object { [string]$_ })
            }
        }
    }
    if ($requirementRows.Count -gt 0) {
        $expected += "Requirement traceability"
        foreach ($requirement in $requirementRows) {
            foreach ($value in $requirement.PSObject.Properties.Value) {
                if ($null -ne $value -and $value -isnot [array] -and
                    $value -isnot [System.Collections.IDictionary] -and
                    $value -isnot [System.Management.Automation.PSCustomObject]) {
                    $expected += [string]$value
                }
            }
        }
    }
    $missingText = @($expected | Where-Object { -not $extractedText.Contains($_) })
    if ($missingText.Count -gt 0) {
        throw "Word extraction did not contain expected title/section text: $($missingText -join '; ')"
    }
    $processes = @(Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
        Where-Object { $baselinePids -notcontains $_.Id })
    if ($processes.Count -ne 1) {
        throw "Could not verify a unique run-owned Word process after rendering."
    }
    $processId = [int]$processes[0].Id
    $docxInfo = Get-Item -LiteralPath $outputPath
    $pdfInfo = Get-Item -LiteralPath $pdfPath
    if ($docxInfo.Length -lt 1 -or $pdfInfo.Length -lt 1) {
        throw "Word did not create non-empty DOCX and PDF files."
    }
    $evidence = [ordered]@{
        document_id = [string]$draft.document_id
        application = [string]$app.Name
        version = [string]$app.Version
        process_id = $processId
        docx = [ordered]@{ path = $docxInfo.FullName; bytes = $docxInfo.Length }
        extracted_text = [ordered]@{ path = $textPath; expected_strings = $expected.Count; matched = $expected.Count }
        pages = $pages
        rendered_pdf = [ordered]@{ path = $pdfInfo.FullName; bytes = $pdfInfo.Length }
        narrative_words = [int](($draft.sections |
            ForEach-Object { [regex]::Matches([string]$_.text, "\b[\w’'-]+\b").Count } |
            Measure-Object -Sum).Sum)
    }
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Word rendered $pages pages and extracted expected content."
}
catch {
    $failure = @(
        $_.Exception.ToString()
        $_.InvocationInfo.PositionMessage
        $_.ScriptStackTrace
    ) -join "`n"
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Word export failed: $failure"
}
finally {
    if ($null -ne $document) {
        try {
            $document.Close(0)
            Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Word document closed."
        } catch {
            $cleanupErrors += "Word document close failed: $($_.Exception.Message)"
        }
    }
    foreach ($object in @($footerRange, $headerRange, $section, $selection, $documentProperties, $document)) {
        if ($null -ne $object) {
            try {
                Release-ComObject $object
            } catch {
                $cleanupErrors += "Word COM object release failed: $($_.Exception.Message)"
            }
        }
    }
    if ($null -ne $app) {
        if ($null -ne $processId) {
            try {
                $app.Quit()
                Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Word application quit."
            } catch {
                $cleanupErrors += "Word application quit failed: $($_.Exception.Message)"
            }
        } else {
            $cleanupErrors += "Word process identity was not verified before quit; launcher cleanup is required."
        }
        try {
            Release-ComObject $app
        } catch {
            $cleanupErrors += "Word application release failed: $($_.Exception.Message)"
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
    $resultPath,
    ($evidence | ConvertTo-Json -Depth 10),
    [Text.UTF8Encoding]::new($false)
)
$evidence | ConvertTo-Json -Depth 10
