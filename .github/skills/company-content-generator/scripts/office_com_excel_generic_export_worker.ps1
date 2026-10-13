$ErrorActionPreference = "Stop"

if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    throw "Excel COM automation requires a single-threaded apartment (STA)."
}

$contextPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_CONTEXT_PATH)
$draftPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_DRAFT_PATH)
$outputPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_OUTPUT_PATH)
$pdfPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_PDF_PATH)
$resultPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_RESULT_PATH)
$logPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_LOG_PATH)
$recordPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_PROCESS_RECORD_PATH)
$baselinePids = @()
if (-not [string]::IsNullOrWhiteSpace($env:JIMWORKS_OFFICE_EXPORT_BASELINE_PIDS)) {
    $baselinePids = @($env:JIMWORKS_OFFICE_EXPORT_BASELINE_PIDS -split "," | ForEach-Object { [int]$_ })
}
foreach ($path in @($outputPath, $pdfPath, $resultPath)) {
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
$workbook = $null
$reopenedWorkbook = $null
$worksheetCollection = $null
$failure = $null
$cleanupErrors = @()
$processId = $null
$evidence = $null
$expectedFormulas = @{}
$expectedValues = @{}
$sectionSheets = @{}
$sheetNames = @{}
$structuredRows = [System.Collections.Generic.List[object]]::new()

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

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object -and [Runtime.InteropServices.Marshal]::IsComObject($Object)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
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
        if ([string]$propertyType.InvokeMember("Value", $getFlags, $null, $property, $null) -ne $Value) {
            throw "Excel did not retain the $Name workbook property."
        }
    } finally {
        Release-ComObject $property
    }
}

function Get-ColumnName {
    param([Parameter(Mandatory = $true)][int]$Column)
    $name = ""
    while ($Column -gt 0) {
        $remainder = ($Column - 1) % 26
        $name = [char](65 + $remainder) + $name
        $Column = [Math]::Floor(($Column - 1) / 26)
    }
    $name
}

function Get-SafeSheetName {
    param([Parameter(Mandatory = $true)][string]$Value)
    $clean = [regex]::Replace($Value, '[\[\]:*?/\\]', " ").Trim().Trim("'")
    if ([string]::IsNullOrWhiteSpace($clean)) { $clean = "Content" }
    $clean = $clean.Substring(0, [Math]::Min(31, $clean.Length)).TrimEnd()
    $candidate = $clean
    $suffix = 2
    while ($sheetNames.ContainsKey($candidate.ToLowerInvariant())) {
        $tail = " $suffix"
        $candidate = $clean.Substring(0, [Math]::Min($clean.Length, 31 - $tail.Length)).TrimEnd() + $tail
        $suffix++
    }
    $sheetNames[$candidate.ToLowerInvariant()] = $true
    $candidate
}

function New-Worksheet {
    param(
        [Parameter(Mandatory = $true)][object]$Book,
        [Parameter(Mandatory = $true)][string]$Name
    )
    $sheet = $Book.Worksheets.Add()
    $sheet.Name = Get-SafeSheetName $Name
    $sheet
}

function Set-Cell {
    param(
        [Parameter(Mandatory = $true)][object]$Sheet,
        [Parameter(Mandatory = $true)][int]$Row,
        [Parameter(Mandatory = $true)][int]$Column,
        [AllowNull()][object]$Value,
        [string]$Label = "cell",
        [switch]$Formula
    )
    $cell = $Sheet.Cells.Item($Row, $Column)
    $cellType = [object].GetMethod("GetType").Invoke($cell, $null)
    $setFlags = [Reflection.BindingFlags]::Instance -bor
        [Reflection.BindingFlags]::Public -bor
        [Reflection.BindingFlags]::SetProperty
    $expectedFormulaValue = $null
    try {
        $isFormula = $Formula -or ($Value -is [string] -and $Value.StartsWith("="))
        if ($Value -is [System.Management.Automation.PSCustomObject]) {
            $properties = @($Value.PSObject.Properties.Name)
            if ($properties -contains "formula") {
                if ([string]::IsNullOrWhiteSpace([string]$Value.formula) -or
                    -not ([string]$Value.formula).StartsWith("=") -or
                    ([string]$Value.formula).Contains("[")) {
                    throw "$Label has an invalid or external-workbook formula."
                }
                if ($Value.PSObject.Properties["expected_value"]) {
                    if ($Value.expected_value -is [bool] -or
                        $Value.expected_value -isnot [ValueType] -or
                        [double]::IsNaN([double]$Value.expected_value) -or
                        [double]::IsInfinity([double]$Value.expected_value)) {
                        throw "$Label has a non-numeric expected formula value."
                    }
                    $expectedFormulaValue = [double]$Value.expected_value
                }
                $Value = [string]$Value.formula
                $isFormula = $true
            } else {
                throw "$Label contains an unsupported object value."
            }
        }
        if ($isFormula) {
            $formulaText = [string]$Value
            if ($formulaText.Contains("[")) {
                throw "$Label contains an external-workbook formula."
            }
            [void]$cellType.InvokeMember("Formula", $setFlags, $null, $cell, @($formulaText))
            $reference = "$($Sheet.Name)!$(Get-ColumnName $Column)$Row"
            $expectedFormulas[$reference] = $formulaText
            if ($null -ne $expectedFormulaValue) { $expectedValues[$reference] = $expectedFormulaValue }
        } else {
            if ($null -eq $Value) { $Value = "" }
            if ($Value -is [bool]) { $Value = [string]$Value }
            if ($Value -is [array] -or $Value -is [System.Collections.IDictionary]) {
                throw "$Label contains an unsupported value."
            }
            [void]$cellType.InvokeMember("Value2", $setFlags, $null, $cell, @($Value))
        }
    } finally {
        Release-ComObject $cell
    }
}

function Add-Table {
    param(
        [Parameter(Mandatory = $true)][object]$Sheet,
        [Parameter(Mandatory = $true)][object]$Specification,
        [Parameter(Mandatory = $true)][int]$StartRow
    )
    $columnCount = $Specification.headers.Count
    $row = $StartRow
    if (-not [string]::IsNullOrWhiteSpace([string]$Specification.title)) {
        Set-Cell -Sheet $Sheet -Row $row -Column 1 -Value ([string]$Specification.title)
        $row++
    }
    for ($column = 0; $column -lt $columnCount; $column++) {
        Set-Cell -Sheet $Sheet -Row $row -Column ($column + 1) `
            -Value ([string]$Specification.headers[$column]) -Label "Table header"
    }
    $headerRange = $Sheet.Range(
        "$(Get-ColumnName 1)$row",
        "$(Get-ColumnName $columnCount)$row"
    )
    try {
        $headerRange.Font.Bold = $true
        $headerRange.Font.Color = 16777215
        $headerRange.Interior.Color = 6051877
    } finally {
        Release-ComObject $headerRange
    }
    $row++
    foreach ($values in $Specification.rows) {
        for ($column = 0; $column -lt $columnCount; $column++) {
            Set-Cell -Sheet $Sheet -Row $row -Column ($column + 1) `
                -Value $values[$column] -Label "Table row $row"
        }
        $row++
    }
    $row + 1
}

function Add-StructuredRows {
    param(
        [AllowNull()][object]$Value,
        [Parameter(Mandatory = $true)][string]$Path
    )
    if ($null -eq $Value) {
        $structuredRows.Add([pscustomobject]@{ path = $Path; value = ""; expected = $null })
        return
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $properties = @($Value.PSObject.Properties.Name)
        if ($properties -contains "formula") {
            if ([string]::IsNullOrWhiteSpace([string]$Value.formula) -or
                -not ([string]$Value.formula).StartsWith("=") -or
                ([string]$Value.formula).Contains("[")) {
                throw "$Path contains an invalid or external-workbook formula."
            }
            $expected = $null
            if ($Value.PSObject.Properties["expected_value"]) {
                if ($Value.expected_value -isnot [ValueType] -or $Value.expected_value -is [bool] -or
                    [double]::IsNaN([double]$Value.expected_value) -or
                    [double]::IsInfinity([double]$Value.expected_value)) {
                    throw "$Path has a non-numeric expected formula value."
                }
                $expected = [double]$Value.expected_value
            }
            $structuredRows.Add([pscustomobject]@{
                path = $Path; value = [string]$Value.formula; expected = $expected
            })
            return
        }
        foreach ($property in $Value.PSObject.Properties) {
            Add-StructuredRows -Value $property.Value -Path "$Path/$($property.Name)"
        }
        if ($properties.Count -eq 0) {
            $structuredRows.Add([pscustomobject]@{ path = $Path; value = "{}"; expected = $null })
        }
        return
    }
    if ($Value -is [array] -or
        ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string])) {
        $items = @($Value)
        for ($index = 0; $index -lt $items.Count; $index++) {
            Add-StructuredRows -Value $items[$index] -Path "$Path/$($index + 1)"
        }
        if ($items.Count -eq 0) {
            $structuredRows.Add([pscustomobject]@{ path = $Path; value = "[]"; expected = $null })
        }
        return
    }
    $structuredRows.Add([pscustomobject]@{ path = $Path; value = $Value; expected = $null })
}

function Format-Worksheet {
    param([Parameter(Mandatory = $true)][object]$Sheet)
    $used = $null
    $columns = $null
    try {
        $used = $Sheet.UsedRange
        $used.Font.Name = "Aptos"
        $used.Font.Size = 10
        $used.WrapText = $true
        $used.VerticalAlignment = -4160
        $columns = $Sheet.Columns
        $columns.AutoFit()
        for ($column = 1; $column -le $used.Columns.Count; $column++) {
            $columnRange = $Sheet.Columns.Item($column)
            try {
                if ($columnRange.ColumnWidth -gt 44) { $columnRange.ColumnWidth = 44 }
                if ($columnRange.ColumnWidth -lt 14) { $columnRange.ColumnWidth = 14 }
            } finally {
                Release-ComObject $columnRange
            }
        }
        $Sheet.PageSetup.Orientation = 2
        $Sheet.PageSetup.Zoom = $false
        $Sheet.PageSetup.FitToPagesWide = 1
        $Sheet.PageSetup.FitToPagesTall = $false
        $Sheet.PageSetup.CenterFooter = "$companyName | $($draft.publication_date)"
        $Sheet.PageSetup.RightFooter = "&P"
    } finally {
        Release-ComObject $columns
        Release-ComObject $used
    }
}

function Get-PropertyValue {
    param([object]$Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return $null }
    $property.Value
}

function Get-OptionalArray {
    param([object]$Object, [string]$Name)
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return @() }
    @($property.Value)
}

try {
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Starting generic Excel authoring."
    $app = New-Object -ComObject Excel.Application
    [void](Register-RunOwnedOfficeProcess -ProcessName EXCEL -BaselinePids $baselinePids -RecordPath $recordPath)
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.EnableEvents = $false
    $app.ScreenUpdating = $false
    $app.AskToUpdateLinks = $false
    $workbook = $app.Workbooks.Add(-4167)
    $worksheetCollection = $workbook.Worksheets
    $overview = $worksheetCollection.Item(1)
    $overview.Name = "Overview"
    $sheetNames["overview"] = $true
    $workbookProperties = $workbook.BuiltinDocumentProperties
    Set-OfficeDocumentProperty -Properties $workbookProperties -Name "Title" -Value ([string]$draft.title)
    Set-OfficeDocumentProperty -Properties $workbookProperties -Name "Author" -Value $companyName
    Set-OfficeDocumentProperty -Properties $workbookProperties -Name "Subject" -Value ([string]$draft.publication_date)
    Release-ComObject $workbookProperties
    $workbookProperties = $null

    Set-Cell -Sheet $overview -Row 1 -Column 1 -Value ([string]$draft.title)
    Set-Cell -Sheet $overview -Row 2 -Column 1 -Value "Company"
    Set-Cell -Sheet $overview -Row 2 -Column 2 -Value $companyName
    Set-Cell -Sheet $overview -Row 3 -Column 1 -Value "Publication date"
    Set-Cell -Sheet $overview -Row 3 -Column 2 -Value ([string]$draft.publication_date)
    Set-Cell -Sheet $overview -Row 5 -Column 1 -Value "Section"
    Set-Cell -Sheet $overview -Row 5 -Column 2 -Value "Content"
    $overviewHeaders = $overview.Range("A5", "B5")
    try {
        $overviewHeaders.Font.Bold = $true
        $overviewHeaders.Font.Color = 16777215
        $overviewHeaders.Interior.Color = 6051877
    } finally {
        Release-ComObject $overviewHeaders
    }
    $overviewRow = 6
    foreach ($section in $draft.sections) {
        $sectionText = [string]$section.text
        if ($sectionText.Length -gt 32767) {
            throw "Section '$($section.heading)' exceeds Excel's 32,767-character cell limit."
        }
        Set-Cell -Sheet $overview -Row $overviewRow -Column 1 -Value ([string]$section.heading)
        Set-Cell -Sheet $overview -Row $overviewRow -Column 2 -Value $sectionText

        $sectionSheet = New-Worksheet -Book $workbook -Name ([string]$section.heading)
        $sectionSheets[[string]$section.heading] = $sectionSheet.Name
        Set-Cell -Sheet $sectionSheet -Row 1 -Column 1 -Value ([string]$draft.title)
        Set-Cell -Sheet $sectionSheet -Row 2 -Column 1 -Value "Published $($draft.publication_date)"
        Set-Cell -Sheet $sectionSheet -Row 3 -Column 1 -Value ([string]$section.heading)
        Set-Cell -Sheet $sectionSheet -Row 4 -Column 1 -Value $sectionText
        $row = 6
        $sectionTables = @(Get-OptionalArray -Object $section -Name "tables")
        foreach ($table in $sectionTables) {
            $row = Add-Table -Sheet $sectionSheet -Specification $table -StartRow $row
        }
        Format-Worksheet -Sheet $sectionSheet
        Release-ComObject $sectionSheet
        $overviewRow++
    }
    $overviewTitleRange = $overview.Range("A1", "F1")
    try {
        $overviewTitleRange.Merge()
        $overviewTitleRange.Font.Name = "Aptos Display"
        $overviewTitleRange.Font.Size = 16
        $overviewTitleRange.Font.Bold = $true
        $overviewTitleRange.Font.Color = 16777215
        $overviewTitleRange.Interior.Color = 6051877
    } finally {
        Release-ComObject $overviewTitleRange
    }
    $columnA = $overview.Columns.Item(1)
    $columnB = $overview.Columns.Item(2)
    try {
        $columnA.ColumnWidth = 34
        $columnB.ColumnWidth = 90
    } finally {
        Release-ComObject $columnA
        Release-ComObject $columnB
    }
    Format-Worksheet -Sheet $overview

    if ($requirementRows.Count -gt 0) {
        $headers = @($requirementRows | ForEach-Object { $_.PSObject.Properties.Name } | Select-Object -Unique)
        if ($headers.Count -eq 0 -or @($headers | Where-Object { [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0) {
            throw "Requirement traceability rows have no valid column headers."
        }
        $requirements = New-Worksheet -Book $workbook -Name "Requirement traceability"
        Set-Cell -Sheet $requirements -Row 1 -Column 1 -Value "Requirement traceability"
        for ($column = 0; $column -lt $headers.Count; $column++) {
            Set-Cell -Sheet $requirements -Row 3 -Column ($column + 1) -Value ([string]$headers[$column])
        }
        $row = 4
        foreach ($requirement in $requirementRows) {
            for ($column = 0; $column -lt $headers.Count; $column++) {
                $value = Get-PropertyValue -Object $requirement -Name ([string]$headers[$column])
                if ($value -is [System.Management.Automation.PSCustomObject] -or $value -is [array]) {
                    $value = $value | ConvertTo-Json -Depth 20 -Compress
                }
                Set-Cell -Sheet $requirements -Row $row -Column ($column + 1) -Value $value
            }
            $row++
        }
        Format-Worksheet -Sheet $requirements
        Release-ComObject $requirements
    }

    if ($null -ne $draft.artifact) {
        Add-StructuredRows -Value $draft.artifact -Path "artifact"
    }
    if ($null -ne $draft.model) {
        Add-StructuredRows -Value $draft.model -Path "model"
    }
    if ($structuredRows.Count -gt 0) {
        $dataSheet = New-Worksheet -Book $workbook -Name "Structured Data"
        Set-Cell -Sheet $dataSheet -Row 1 -Column 1 -Value "Structured source data"
        Set-Cell -Sheet $dataSheet -Row 3 -Column 1 -Value "Source path"
        Set-Cell -Sheet $dataSheet -Row 3 -Column 2 -Value "Value"
        $dataHeaders = $dataSheet.Range("A3", "B3")
        try {
            $dataHeaders.Font.Bold = $true
            $dataHeaders.Font.Color = 16777215
            $dataHeaders.Interior.Color = 6051877
        } finally {
            Release-ComObject $dataHeaders
        }
        $row = 4
        foreach ($entry in $structuredRows) {
            Set-Cell -Sheet $dataSheet -Row $row -Column 1 -Value $entry.path
            $expectedValue = $null
            if ($null -ne $entry.expected) { $expectedValue = [double]$entry.expected }
            $valueObject = if ($entry.value -is [string] -and ([string]$entry.value).StartsWith("=")) {
                [string]$entry.value
            } else { $entry.value }
            Set-Cell -Sheet $dataSheet -Row $row -Column 2 -Value $valueObject -Label $entry.path
            $reference = "$($dataSheet.Name)!B$row"
            if ($null -ne $expectedValue) { $expectedValues[$reference] = $expectedValue }
            $row++
        }
        Format-Worksheet -Sheet $dataSheet
        Release-ComObject $dataSheet
    }

    if ([int]$overview.Index -gt 1) {
        $firstSheet = $workbook.Worksheets.Item(1)
        try { $overview.Move($firstSheet) } finally { Release-ComObject $firstSheet }
    }
    $app.CalculateFullRebuild()
    $workbook.SaveAs($outputPath, 51)
    $workbook.Save()
    $workbook.ExportAsFixedFormat(0, $pdfPath)
    $workbook.Close($false)
    Release-ComObject $worksheetCollection
    $worksheetCollection = $null
    Release-ComObject $overview
    $overview = $null
    Release-ComObject $workbook
    $workbook = $null

    $reopenedWorkbook = $app.Workbooks.Open($outputPath, 0, $true)
    $reopenedWorksheets = $reopenedWorkbook.Worksheets
    $cachedValues = @{}
    $formulaErrors = @()
    $actualFormulas = @{}
    $allSheetNames = @()
    for ($sheetIndex = 1; $sheetIndex -le $reopenedWorksheets.Count; $sheetIndex++) {
        $sheet = $reopenedWorksheets.Item($sheetIndex)
        $allSheetNames += [string]$sheet.Name
        $usedRange = $null
        try {
            $usedRange = $sheet.UsedRange
            $firstRow = [int]$usedRange.Row
            $firstColumn = [int]$usedRange.Column
            for ($rowOffset = 0; $rowOffset -lt $usedRange.Rows.Count; $rowOffset++) {
                for ($columnOffset = 0; $columnOffset -lt $usedRange.Columns.Count; $columnOffset++) {
                    $rowNumber = $firstRow + $rowOffset
                    $columnNumber = $firstColumn + $columnOffset
                    $cell = $sheet.Cells.Item($rowNumber, $columnNumber)
                    try {
                        $value = $cell.Value2
                        $text = [string]$cell.Text
                        $formula = [string]$cell.Formula
                        if ($formula.StartsWith("=")) {
                            $reference = "$($sheet.Name)!$(Get-ColumnName $columnNumber)$rowNumber"
                            $actualFormulas[$reference] = $formula
                            if ($text -in @("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!")) {
                                $formulaErrors += "$reference=$text"
                            }
                            if ($null -eq $value) {
                                throw "Excel did not retain a cached value for $reference."
                            }
                            $cachedValues[$reference] = $value
                        }
                    } finally {
                        Release-ComObject $cell
                    }
                }
            }
        } finally {
            Release-ComObject $usedRange
            Release-ComObject $sheet
        }
    }
    if ($formulaErrors.Count -gt 0) {
        throw "Excel formula errors found after recalculation: $($formulaErrors -join ', ')."
    }
    foreach ($reference in $expectedFormulas.Keys) {
        if (-not $actualFormulas.ContainsKey($reference) -or
            $actualFormulas[$reference] -ne $expectedFormulas[$reference]) {
            throw "Excel changed or dropped formula at $reference."
        }
    }
    if ($actualFormulas.Count -ne $expectedFormulas.Count) {
        throw "Excel formula count changed during save/reopen."
    }
    if ($allSheetNames.Count -eq 0 -or $allSheetNames[0] -ne "Overview") {
        throw "Excel did not retain the Overview worksheet as the first sheet."
    }
    foreach ($reference in $expectedValues.Keys) {
        $actual = $cachedValues[$reference]
        if ($actual -is [bool] -or $null -eq $actual -or
            [Math]::Abs([double]$actual - [double]$expectedValues[$reference]) -gt 0.000000001) {
            throw "Unexpected recalculated value for $reference; expected $($expectedValues[$reference]), got $actual."
        }
    }
    $reopenedOverview = $reopenedWorksheets.Item("Overview")
    $overviewTitleCell = $reopenedOverview.Range("A1")
    try {
        if ([string]$overviewTitleCell.Value2 -ne [string]$draft.title) {
            throw "Excel did not retain the workbook title on the Overview sheet."
        }
    } finally {
        Release-ComObject $overviewTitleCell
        Release-ComObject $reopenedOverview
    }

    $newProcesses = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $baselinePids -notcontains $_.Id })
    if ($newProcesses.Count -ne 1) {
        throw "Could not verify a unique run-owned Excel process after export."
    }
    $processId = [int]$newProcesses[0].Id
    $xlsxInfo = Get-Item -LiteralPath $outputPath
    $pdfInfo = Get-Item -LiteralPath $pdfPath
    if ($xlsxInfo.Length -lt 1 -or $pdfInfo.Length -lt 1) {
        throw "Excel did not create non-empty XLSX and PDF files."
    }
    $evidence = [ordered]@{
        document_id = [string]$draft.document_id
        application = [string]$app.Name
        version = [string]$app.Version
        process_id = $processId
        xlsx = [ordered]@{ path = $xlsxInfo.FullName; bytes = $xlsxInfo.Length }
        extracted_text_match = $true
        sheet_names = $allSheetNames
        formula_count = $actualFormulas.Count
        cached_formula_values = $cachedValues
        expected_values_checked = $expectedValues.Count
        formula_errors = $formulaErrors.Count
        rendered_pdf = [ordered]@{ path = $pdfInfo.FullName; bytes = $pdfInfo.Length }
    }
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Excel saved, recalculated, reopened, checked formulas, and rendered PDF."
}
catch {
    $failure = @(
        $_.Exception.ToString()
        $_.InvocationInfo.PositionMessage
        $_.ScriptStackTrace
    ) -join "`n"
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Excel export failed: $failure"
}
finally {
    if ($null -ne $reopenedWorkbook) {
        try { $reopenedWorkbook.Close($false) } catch { $cleanupErrors += "Excel reopened workbook close failed: $($_.Exception.Message)" }
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { $cleanupErrors += "Excel workbook close failed: $($_.Exception.Message)" }
    }
    foreach ($object in @($reopenedWorksheets, $worksheetCollection, $overview, $reopenedWorkbook, $workbook)) {
        if ($null -ne $object) {
            try { Release-ComObject $object } catch { $cleanupErrors += "Excel COM object release failed: $($_.Exception.Message)" }
        }
    }
    if ($null -ne $app) {
        if ($null -ne $processId) {
            try {
                $app.Quit()
            } catch {
                $cleanupErrors += "Excel application quit failed: $($_.Exception.Message)"
            }
        } else {
            $cleanupErrors += "Excel process identity was not verified before quit; launcher cleanup is required."
        }
        try { Release-ComObject $app } catch { $cleanupErrors += "Excel application release failed: $($_.Exception.Message)" }
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
