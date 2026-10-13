$ErrorActionPreference = "Stop"

if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    throw "Excel COM automation requires a single-threaded apartment (STA)."
}

$draftPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_DRAFT_PATH)
$outputPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_OUTPUT_PATH)
$pdfPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_PDF_PATH)
$resultPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_RESULT_PATH)
$logPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_LOG_PATH)
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
$artifact = $draft.artifact
if ($artifact.module_names.Count -lt 6 -or $artifact.module_names.Count -gt 7) {
    throw "The medium-depth financial workbook requires six or seven planned analysis modules."
}
$years = @($artifact.annual_financials | ForEach-Object { [int]$_.year })
$annualByYear = @{}
$detailByYear = @{}
$operationsByYear = @{}
$organizationByYear = @{}
$quarterByYear = @{}
foreach ($item in $artifact.annual_financials) { $annualByYear[[string]$item.year] = $item }
foreach ($item in $artifact.financial_details) { $detailByYear[[string]$item.year] = $item }
foreach ($item in $artifact.operating_metrics) { $operationsByYear[[string]$item.year] = $item }
foreach ($item in $artifact.organization_summaries) { $organizationByYear[[string]$item.year] = $item }
foreach ($year in $years) { $quarterByYear[[string]$year] = @($artifact.quarterly_financials | Where-Object { [int]$_.year -eq $year } | Sort-Object quarter) }

$app = $null
$workbook = $null
$workbookProperties = $null
$reopenedWorkbook = $null
$reopenedCollection = $null
$worksheetCollection = $null
$reopenedSheets = @{}
$usedRange = $null
$cell = $null
$failure = $null
$cleanupErrors = @()
$processId = $null
$evidence = $null
$expectedCells = @()
$sheetObjects = @{}
$formulaCount = 0
$formulaErrors = @()

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
        $actual = [string]$propertyType.InvokeMember("Value", $getFlags, $null, $property, $null)
        if ($actual -ne $Value) {
            throw "Office did not retain the $Name document property."
        }
    } finally {
        Release-ComObject $property
    }
}

function Set-Cell {
    param(
        [Parameter(Mandatory = $true)][object]$Sheet,
        [Parameter(Mandatory = $true)][int]$Row,
        [Parameter(Mandatory = $true)][int]$Column,
        [AllowNull()][object]$Value,
        [switch]$Formula
    )
    $targetCell = $Sheet.Cells.Item($Row, $Column)
    try {
        $cellType = [object].GetMethod("GetType").Invoke($targetCell, $null)
        $setFlags = [Reflection.BindingFlags]::Instance -bor
            [Reflection.BindingFlags]::Public -bor
            [Reflection.BindingFlags]::SetProperty
        if ($Formula) {
            [void]$cellType.InvokeMember("Formula", $setFlags, $null, $targetCell, @([string]$Value))
        } else {
            [void]$cellType.InvokeMember("Value2", $setFlags, $null, $targetCell, @($Value))
        }
    } finally {
        Release-ComObject $targetCell
    }
}

function Add-ExpectedCell {
    param(
        [string]$Sheet,
        [int]$Row,
        [int]$Column,
        [double]$Value,
        [string]$Label,
        [double]$Tolerance = 0.01
    )
    $script:expectedCells += [pscustomobject]@{
        Sheet = $Sheet; Row = $Row; Column = $Column
        Value = $Value; Label = $Label; Tolerance = $Tolerance
    }
}

function Format-Worksheet {
    param(
        [Parameter(Mandatory = $true)][object]$Sheet,
        [Parameter(Mandatory = $true)][int]$LastColumn,
        [Parameter(Mandatory = $true)][int]$LastRow,
        [switch]$NoHeaderRow
    )
    $lastColumnLetter = [string][char](64 + [Math]::Min($LastColumn, 26))
    $wideSheet = $LastColumn -gt 12
    $titleRange = $null
    $subtitleRange = $null
    $headerRange = $null
    $used = $null
    $contentRange = $null
    try {
        $used = $Sheet.UsedRange
        $used.Font.Name = "Aptos"
        $used.Font.Size = 10
        $used.Columns.AutoFit()
        for ($column = 1; $column -le $LastColumn; $column++) {
            $columnRange = $Sheet.Columns.Item($column)
            try {
                if ($columnRange.ColumnWidth -gt 28) {
                    $columnRange.ColumnWidth = 28
                }
            } finally {
                Release-ComObject $columnRange
            }
        }
        $contentRange = $Sheet.Range("A1:$lastColumnLetter$LastRow")
        $contentRange.WrapText = $true
        $contentRange.VerticalAlignment = -4160
        if ($wideSheet) {
            $titleRange = $Sheet.Range("A1:B1")
            $subtitleRange = $Sheet.Range("A2:B2")
            $subtitleRange.Merge()
            $subtitleRange.WrapText = $true
        } else {
            $titleRange = $Sheet.Range("A1:$($lastColumnLetter)1")
        }
        $titleRange.Merge()
        $titleRange.Font.Name = "Aptos Display"
        $titleRange.Font.Size = 16
        $titleRange.Font.Bold = $true
        $titleRange.Font.Color = 5197615
        if (-not $NoHeaderRow -and $LastRow -ge 3) {
            $headerRange = $Sheet.Range("A3:$($lastColumnLetter)3")
            $headerRange.Font.Bold = $true
            $headerRange.Font.Color = 16777215
            $headerRange.Interior.Color = 5197615
            $headerRange.WrapText = $true
        }
        $titleRange.WrapText = $wideSheet
        $used.Rows.AutoFit()
        if ($wideSheet) {
            $titleRange.RowHeight = 42
            $subtitleRange.RowHeight = 54
        } else {
            $titleRange.RowHeight = 28
        }
        $Sheet.PageSetup.Orientation = 2
        $Sheet.PageSetup.Zoom = $false
        $Sheet.PageSetup.FitToPagesWide = if ($wideSheet) { 3 } else { 1 }
        $Sheet.PageSetup.FitToPagesTall = $false
        $Sheet.PageSetup.PrintTitleRows = if (-not $NoHeaderRow -and $LastRow -ge 3) { '$1:$3' } else { '' }
        $Sheet.PageSetup.PrintTitleColumns = if ($wideSheet) { '$A:$B' } else { '' }
        $Sheet.PageSetup.CenterFooter = "Jimworks | Internal"
        $Sheet.PageSetup.RightFooter = "&P"
        $Sheet.PageSetup.LeftMargin = $Sheet.Application.InchesToPoints(0.35)
        $Sheet.PageSetup.RightMargin = $Sheet.Application.InchesToPoints(0.35)
        $Sheet.PageSetup.TopMargin = $Sheet.Application.InchesToPoints(0.5)
        $Sheet.PageSetup.BottomMargin = $Sheet.Application.InchesToPoints(0.5)
    } finally {
        Release-ComObject $contentRange
        Release-ComObject $used
        Release-ComObject $headerRange
        Release-ComObject $subtitleRange
        Release-ComObject $titleRange
    }
}

function Add-FormulaExpected {
    param(
        [object]$Sheet,
        [int]$Row,
        [int]$Column,
        [double]$Expected,
        [string]$Label,
        [double]$Tolerance = 0.01
    )
    Add-ExpectedCell -Sheet $Sheet.Name -Row $Row -Column $Column `
        -Value $Expected -Label $Label -Tolerance $Tolerance
}

try {
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Starting Excel authoring for $($draft.document_id)."
    $app = New-Object -ComObject Excel.Application
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $app.EnableEvents = $false
    $app.ScreenUpdating = $false
    $app.AskToUpdateLinks = $false
    $workbook = $app.Workbooks.Add()
    $worksheetCollection = $workbook.Worksheets
    $startSheet = $worksheetCollection.Item(1)
    $startSheet.Name = "Start Here"
    while ($worksheetCollection.Count -gt 1) {
        $extraSheet = $worksheetCollection.Item(2)
        try {
            $extraSheet.Delete()
        } finally {
            Release-ComObject $extraSheet
        }
    }
    $sheetObjects["Start Here"] = $startSheet
    foreach ($sheetName in @(
        "Quarterly Actuals", "Financial Analysis", "Product Mix",
        "Customer Segments", "Cost Drivers", "Operating KPIs",
        "Workforce", "Strategy and Risks"
    )) {
        $newSheet = $worksheetCollection.Add()
        $newSheet.Name = $sheetName
        $sheetObjects[$sheetName] = $newSheet
    }
    $workbookProperties = $workbook.BuiltinDocumentProperties
    Set-OfficeDocumentProperty -Properties $workbookProperties -Name "Title" -Value ([string]$draft.title)
    Set-OfficeDocumentProperty -Properties $workbookProperties -Name "Subject" -Value "Annual business and financial analysis"
    Set-OfficeDocumentProperty -Properties $workbookProperties -Name "Author" -Value "Jimworks"
    Release-ComObject $workbookProperties
    $workbookProperties = $null

    $start = $sheetObjects["Start Here"]
    Set-Cell $start 1 1 ([string]$draft.title)
    Set-Cell $start 2 1 "Fiscal year ended December 31, $($artifact.year) | Published $($draft.publication_date)"
    Set-Cell $start 4 1 "Purpose"
    Set-Cell $start 5 1 "Source-linked annual financial and operating analysis. All financial amounts are USD whole dollars; percentages are calculated from those source amounts."
    Set-Cell $start 7 1 "Analysis modules"
    for ($i = 0; $i -lt $artifact.module_names.Count; $i++) {
        Set-Cell $start ($i + 8) 1 ([string]$artifact.module_names[$i])
    }
    Set-Cell $start 17 1 "Source and control notes"
    Set-Cell $start 18 1 "Source: Jimworks canonical annual financial, quarterly financial, financial-detail, operating-metric, and organization-summary records available by $($draft.publication_date)."
    Set-Cell $start 19 1 "Quarterly actuals reconcile to the annual financial source in Financial Analysis."
    Set-Cell $start 20 1 "Product, customer-segment, COGS, and OpEx categories reconcile to their canonical annual totals."
    Set-Cell $start 21 1 "Revenue per year-end employee divides an annual flow by a point-in-time headcount; it is not an individual productivity measure."
    Set-Cell $start 22 1 "This workbook contains actuals and derived analysis, not forecasts, cash-flow projections, or customer-level profitability."
    Set-Cell $start 23 1 "Use formula cells for derived values; do not overwrite source actuals without an approved revision."
    $start.Range("A1:A23").WrapText = $true
    $start.Range("A4:A4").Font.Bold = $true
    $start.Range("A7:A7").Font.Bold = $true
    $start.Range("A17:A17").Font.Bold = $true

    $quarterSheet = $sheetObjects["Quarterly Actuals"]
    Set-Cell $quarterSheet 1 1 "Quarterly Actuals and Annual Tie-Out"
    Set-Cell $quarterSheet 2 1 "Source actuals by quarter; annual totals are formula-derived. Headcount is a point-in-time Q4 value, not a sum."
    $quarterHeaders = @("Year", "Period", "Period End", "Revenue (USD)", "COGS (USD)", "Operating Expenses (USD)", "Operating Income (USD)", "Year-End Headcount", "Gross Margin", "Operating Margin", "Sequential Revenue Growth")
    for ($column = 0; $column -lt $quarterHeaders.Count; $column++) {
        Set-Cell $quarterSheet 3 ($column + 1) $quarterHeaders[$column]
    }
    $annualRowByYear = @{}
    $row = 4
    foreach ($year in $years) {
        $firstQuarterRow = $row
        foreach ($quarter in $quarterByYear[[string]$year]) {
            Set-Cell $quarterSheet $row 1 ([string]$year)
            Set-Cell $quarterSheet $row 2 ("Q" + [int]$quarter.quarter)
            Set-Cell $quarterSheet $row 3 ([string]$quarter.period_end)
            Set-Cell $quarterSheet $row 4 ([int]$quarter.revenue_usd)
            Set-Cell $quarterSheet $row 5 ([int]$quarter.cogs_usd)
            Set-Cell $quarterSheet $row 6 ([int]$quarter.opex_usd)
            Set-Cell $quarterSheet $row 7 ([int]$quarter.operating_income_usd)
            Set-Cell $quarterSheet $row 8 ([int]$quarter.headcount)
            Set-Cell $quarterSheet $row 9 "=(D$row-E$row)/D$row" -Formula
            Set-Cell $quarterSheet $row 10 "=G$row/D$row" -Formula
            if ([int]$quarter.quarter -gt 1) {
                Set-Cell $quarterSheet $row 11 "=D$row/D$($row - 1)-1" -Formula
            }
            Add-FormulaExpected -Sheet $quarterSheet -Row $row -Column 9 -Expected ([double](($quarter.revenue_usd - $quarter.cogs_usd) / $quarter.revenue_usd)) -Label "$year Q$($quarter.quarter) gross margin" -Tolerance 0.000001
            Add-FormulaExpected -Sheet $quarterSheet -Row $row -Column 10 -Expected ([double]($quarter.operating_income_usd / $quarter.revenue_usd)) -Label "$year Q$($quarter.quarter) operating margin" -Tolerance 0.000001
            if ([int]$quarter.quarter -gt 1) {
                $previousQuarter = $quarterByYear[[string]$year][[int]$quarter.quarter - 2]
                Add-FormulaExpected -Sheet $quarterSheet -Row $row -Column 11 -Expected ([double]($quarter.revenue_usd / $previousQuarter.revenue_usd - 1)) -Label "$year Q$($quarter.quarter) sequential revenue growth" -Tolerance 0.000001
            }
            $row++
        }
        $lastQuarterRow = $row - 1
        Set-Cell $quarterSheet $row 1 "FY$year"
        Set-Cell $quarterSheet $row 2 "Annual total"
        Set-Cell $quarterSheet $row 3 ([string]$annualByYear[[string]$year].period_end)
        Set-Cell $quarterSheet $row 4 "=SUM(D$firstQuarterRow`:D$lastQuarterRow)" -Formula
        Set-Cell $quarterSheet $row 5 "=SUM(E$firstQuarterRow`:E$lastQuarterRow)" -Formula
        Set-Cell $quarterSheet $row 6 "=SUM(F$firstQuarterRow`:F$lastQuarterRow)" -Formula
        Set-Cell $quarterSheet $row 7 "=D$row-E$row-F$row" -Formula
        Set-Cell $quarterSheet $row 8 "=H$lastQuarterRow" -Formula
        Set-Cell $quarterSheet $row 9 "=(D$row-E$row)/D$row" -Formula
        Set-Cell $quarterSheet $row 10 "=G$row/D$row" -Formula
        $annualRowByYear[[string]$year] = $row
        Add-FormulaExpected -Sheet $quarterSheet -Row $row -Column 4 -Expected ([double]$annualByYear[[string]$year].revenue_usd) -Label "$year quarterly revenue total"
        Add-FormulaExpected -Sheet $quarterSheet -Row $row -Column 5 -Expected ([double]$annualByYear[[string]$year].cogs_usd) -Label "$year quarterly COGS total"
        Add-FormulaExpected -Sheet $quarterSheet -Row $row -Column 6 -Expected ([double]$annualByYear[[string]$year].opex_usd) -Label "$year quarterly OpEx total"
        Add-FormulaExpected -Sheet $quarterSheet -Row $row -Column 7 -Expected ([double]$annualByYear[[string]$year].operating_income_usd) -Label "$year quarterly operating income total"
        Add-FormulaExpected -Sheet $quarterSheet -Row $row -Column 8 -Expected ([double]$annualByYear[[string]$year].headcount) -Label "$year Q4 headcount"
        Add-FormulaExpected -Sheet $quarterSheet -Row $row -Column 9 -Expected ([double](($annualByYear[[string]$year].revenue_usd - $annualByYear[[string]$year].cogs_usd) / $annualByYear[[string]$year].revenue_usd)) -Label "$year annual gross margin" -Tolerance 0.000001
        Add-FormulaExpected -Sheet $quarterSheet -Row $row -Column 10 -Expected ([double]($annualByYear[[string]$year].operating_income_usd / $annualByYear[[string]$year].revenue_usd)) -Label "$year annual operating margin" -Tolerance 0.000001
        for ($column = 4; $column -le 11; $column++) {
            $inputCell = $quarterSheet.Cells.Item($firstQuarterRow, $column)
            try { $inputCell.Font.Color = 255 } finally { Release-ComObject $inputCell }
        }
        for ($quarterRow = $firstQuarterRow + 1; $quarterRow -le $lastQuarterRow; $quarterRow++) {
            for ($column = 4; $column -le 8; $column++) {
                $inputCell = $quarterSheet.Cells.Item($quarterRow, $column)
                try { $inputCell.Font.Color = 255 } finally { Release-ComObject $inputCell }
            }
        }
        for ($quarterRow = $firstQuarterRow; $quarterRow -le $lastQuarterRow; $quarterRow++) {
            for ($column = 9; $column -le 11; $column++) {
                $formulaCell = $quarterSheet.Cells.Item($quarterRow, $column)
                try { $formulaCell.Font.Color = 0 } finally { Release-ComObject $formulaCell }
            }
        }
        $row++
    }
    $quarterLastRow = $row - 1
    $quarterSheet.Range("D4:G$quarterLastRow").NumberFormat = '$#,##0;($#,##0);-'
    $quarterSheet.Range("I4:K$quarterLastRow").NumberFormat = "0.0%;(0.0%);-"

    $analysis = $sheetObjects["Financial Analysis"]
    Set-Cell $analysis 1 1 "Annual Financial Analysis and Controls"
    Set-Cell $analysis 2 1 "Calculated annual totals link to quarterly actuals; source annual values and variance checks are shown alongside."
    $analysisHeaders = @("Year", "Revenue Calc (USD)", "Revenue Source (USD)", "Revenue Difference (USD)", "COGS Calc (USD)", "COGS Source (USD)", "COGS Difference (USD)", "OpEx Calc (USD)", "OpEx Source (USD)", "OpEx Difference (USD)", "Operating Income Calc (USD)", "Operating Income Source (USD)", "Income Difference (USD)", "Gross Margin", "Operating Margin", "Revenue Growth")
    for ($column = 0; $column -lt $analysisHeaders.Count; $column++) {
        Set-Cell $analysis 3 ($column + 1) $analysisHeaders[$column]
    }
    $analysisRowByYear = @{}
    for ($i = 0; $i -lt $years.Count; $i++) {
        $year = $years[$i]
        $r = $i + 4
        $analysisRowByYear[[string]$year] = $r
        $annualRow = $annualRowByYear[[string]$year]
        $source = $annualByYear[[string]$year]
        Set-Cell $analysis $r 1 ([string]$year)
        Set-Cell $analysis $r 2 "='Quarterly Actuals'!D$annualRow" -Formula
        Set-Cell $analysis $r 3 ([int]$source.revenue_usd)
        Set-Cell $analysis $r 4 "=B$r-C$r" -Formula
        Set-Cell $analysis $r 5 "='Quarterly Actuals'!E$annualRow" -Formula
        Set-Cell $analysis $r 6 ([int]$source.cogs_usd)
        Set-Cell $analysis $r 7 "=E$r-F$r" -Formula
        Set-Cell $analysis $r 8 "='Quarterly Actuals'!F$annualRow" -Formula
        Set-Cell $analysis $r 9 ([int]$source.opex_usd)
        Set-Cell $analysis $r 10 "=H$r-I$r" -Formula
        Set-Cell $analysis $r 11 "=B$r-E$r-H$r" -Formula
        Set-Cell $analysis $r 12 ([int]$source.operating_income_usd)
        Set-Cell $analysis $r 13 "=K$r-L$r" -Formula
        Set-Cell $analysis $r 14 "=(B$r-E$r)/B$r" -Formula
        Set-Cell $analysis $r 15 "=K$r/B$r" -Formula
        if ($i -gt 0) {
            Set-Cell $analysis $r 16 "=B$r/B$($r - 1)-1" -Formula
        }
        foreach ($column in @(2, 4, 5, 7, 8, 10, 11, 13, 14, 15)) {
            $expected = switch ($column) {
                2 { [double]$source.revenue_usd }
                4 { 0.0 }
                5 { [double]$source.cogs_usd }
                7 { 0.0 }
                8 { [double]$source.opex_usd }
                10 { 0.0 }
                11 { [double]$source.operating_income_usd }
                13 { 0.0 }
                14 { [double](($source.revenue_usd - $source.cogs_usd) / $source.revenue_usd) }
                15 { [double]($source.operating_income_usd / $source.revenue_usd) }
            }
            Add-FormulaExpected -Sheet $analysis -Row $r -Column $column -Expected $expected -Label "$year financial-analysis column $column" -Tolerance 0.000001
        }
        if ($i -gt 0) {
            $previousYear = $years[$i - 1]
            $expectedGrowth = [double]($source.revenue_usd / $annualByYear[[string]$previousYear].revenue_usd - 1)
            Add-FormulaExpected -Sheet $analysis -Row $r -Column 16 -Expected $expectedGrowth -Label "$year revenue growth" -Tolerance 0.000001
        }
        foreach ($column in @(3, 6, 9, 12)) {
            $inputCell = $analysis.Cells.Item($r, $column)
            try { $inputCell.Font.Color = 255 } finally { Release-ComObject $inputCell }
        }
        foreach ($column in @(2, 5, 8, 11, 14, 15, 16)) {
            $formulaCell = $analysis.Cells.Item($r, $column)
            try {
                $formulaCell.Font.Color = if ($column -in @(2, 5, 8)) { 32768 } else { 0 }
            } finally { Release-ComObject $formulaCell }
        }
    }
    $analysis.Range("B4:M$($years.Count + 3)").NumberFormat = '$#,##0;($#,##0);-'
    $analysis.Range("N4:P$($years.Count + 3)").NumberFormat = "0.0%;(0.0%);-"

    $productSheet = $sheetObjects["Product Mix"]
    Set-Cell $productSheet 1 1 "Revenue Mix by Product and Service"
    Set-Cell $productSheet 2 1 "Annual source allocations reconcile to recognized revenue; shares are formula-derived."
    $productHeaders = @("Year", "Category", "Revenue (USD)", "Share of Revenue", "Annual Revenue Source (USD)", "Difference (USD)")
    for ($column = 0; $column -lt $productHeaders.Count; $column++) { Set-Cell $productSheet 3 ($column + 1) $productHeaders[$column] }
    $row = 4
    foreach ($year in $years) {
        $items = @($detailByYear[[string]$year].product_revenue)
        $first = $row
        foreach ($item in $items) {
            Set-Cell $productSheet $row 1 ([string]$year)
            Set-Cell $productSheet $row 2 ([string]$item.category)
            Set-Cell $productSheet $row 3 ([int]$item.revenue_usd)
            Set-Cell $productSheet $row 4 "=C$row/`$C`$$($first + $items.Count)" -Formula
            $row++
        }
        $totalRow = $row
        Set-Cell $productSheet $totalRow 1 ([string]$year)
        Set-Cell $productSheet $totalRow 2 "Total"
        Set-Cell $productSheet $totalRow 3 "=SUM(C$first`:C$($row - 1))" -Formula
        Set-Cell $productSheet $totalRow 4 "=SUM(D$first`:D$($row - 1))" -Formula
        Set-Cell $productSheet $totalRow 5 ([int]$annualByYear[[string]$year].revenue_usd)
        Set-Cell $productSheet $totalRow 6 "=C$totalRow-E$totalRow" -Formula
        Add-FormulaExpected -Sheet $productSheet -Row $totalRow -Column 3 -Expected ([double]$annualByYear[[string]$year].revenue_usd) -Label "$year product revenue total"
        Add-FormulaExpected -Sheet $productSheet -Row $totalRow -Column 4 -Expected 1.0 -Label "$year product share total" -Tolerance 0.000001
        Add-FormulaExpected -Sheet $productSheet -Row $totalRow -Column 6 -Expected 0.0 -Label "$year product revenue reconciliation"
        for ($r = $first; $r -lt $totalRow; $r++) {
            $itemIndex = $r - $first
            Add-FormulaExpected -Sheet $productSheet -Row $r -Column 4 -Expected ([double]($items[$itemIndex].revenue_usd / $annualByYear[[string]$year].revenue_usd)) -Label "$year product share $itemIndex" -Tolerance 0.000001
            $productSourceCell = $productSheet.Cells.Item($r, 3)
            try { $productSourceCell.Font.Color = 255 } finally { Release-ComObject $productSourceCell }
        }
        $productTotalSource = $productSheet.Cells.Item($totalRow, 5)
        try { $productTotalSource.Font.Color = 255 } finally { Release-ComObject $productTotalSource }
        $productTotalFormula = $productSheet.Cells.Item($totalRow, 3)
        try { $productTotalFormula.Font.Color = 0 } finally { Release-ComObject $productTotalFormula }
        $row++
    }
    $productLastRow = $row - 1
    $productSheet.Range("C4:C$productLastRow").NumberFormat = '$#,##0;($#,##0);-'
    $productSheet.Range("D4:D$productLastRow").NumberFormat = "0.0%;(0.0%);-"
    $productSheet.Range("E4:F$productLastRow").NumberFormat = '$#,##0;($#,##0);-'

    $segmentSheet = $sheetObjects["Customer Segments"]
    Set-Cell $segmentSheet 1 1 "Revenue Mix by Customer Segment"
    Set-Cell $segmentSheet 2 1 "Annual source allocations reconcile to recognized revenue; segment shares are formula-derived."
    $segmentHeaders = @("Year", "Customer Segment", "Revenue", "Share of Revenue", "Annual Revenue Source", "Difference")
    for ($column = 0; $column -lt $segmentHeaders.Count; $column++) { Set-Cell $segmentSheet 3 ($column + 1) $segmentHeaders[$column] }
    $row = 4
    foreach ($year in $years) {
        $items = @($detailByYear[[string]$year].customer_segment_revenue)
        $first = $row
        foreach ($item in $items) {
            Set-Cell $segmentSheet $row 1 ([string]$year)
            Set-Cell $segmentSheet $row 2 ([string]$item.segment)
            Set-Cell $segmentSheet $row 3 ([int]$item.revenue_usd)
            Set-Cell $segmentSheet $row 4 "=C$row/`$C`$$($first + $items.Count)" -Formula
            $row++
        }
        $totalRow = $row
        Set-Cell $segmentSheet $totalRow 1 ([string]$year)
        Set-Cell $segmentSheet $totalRow 2 "Total"
        Set-Cell $segmentSheet $totalRow 3 "=SUM(C$first`:C$($row - 1))" -Formula
        Set-Cell $segmentSheet $totalRow 4 "=SUM(D$first`:D$($row - 1))" -Formula
        Set-Cell $segmentSheet $totalRow 5 ([int]$annualByYear[[string]$year].revenue_usd)
        Set-Cell $segmentSheet $totalRow 6 "=C$totalRow-E$totalRow" -Formula
        Add-FormulaExpected -Sheet $segmentSheet -Row $totalRow -Column 3 -Expected ([double]$annualByYear[[string]$year].revenue_usd) -Label "$year segment revenue total"
        Add-FormulaExpected -Sheet $segmentSheet -Row $totalRow -Column 4 -Expected 1.0 -Label "$year segment share total" -Tolerance 0.000001
        Add-FormulaExpected -Sheet $segmentSheet -Row $totalRow -Column 6 -Expected 0.0 -Label "$year segment revenue reconciliation"
        for ($r = $first; $r -lt $totalRow; $r++) {
            $itemIndex = $r - $first
            Add-FormulaExpected -Sheet $segmentSheet -Row $r -Column 4 -Expected ([double]($items[$itemIndex].revenue_usd / $annualByYear[[string]$year].revenue_usd)) -Label "$year segment share $itemIndex" -Tolerance 0.000001
            $segmentSourceCell = $segmentSheet.Cells.Item($r, 3)
            try { $segmentSourceCell.Font.Color = 255 } finally { Release-ComObject $segmentSourceCell }
        }
        $segmentTotalSource = $segmentSheet.Cells.Item($totalRow, 5)
        try { $segmentTotalSource.Font.Color = 255 } finally { Release-ComObject $segmentTotalSource }
        $segmentTotalFormula = $segmentSheet.Cells.Item($totalRow, 3)
        try { $segmentTotalFormula.Font.Color = 0 } finally { Release-ComObject $segmentTotalFormula }
        $row++
    }
    $segmentLastRow = $row - 1
    $segmentSheet.Range("C4:C$segmentLastRow").NumberFormat = '$#,##0;($#,##0);-'
    $segmentSheet.Range("D4:D$segmentLastRow").NumberFormat = "0.0%;(0.0%);-"
    $segmentSheet.Range("E4:F$segmentLastRow").NumberFormat = '$#,##0;($#,##0);-'

    $costSheet = $sheetObjects["Cost Drivers"]
    Set-Cell $costSheet 1 1 "Cost Structure and Operating Leverage"
    Set-Cell $costSheet 2 1 "Source categories reconcile to reported COGS and OpEx; shares use formula-linked group totals and revenue."
    $costHeaders = @("Year", "Cost Group", "Category", "Amount", "Share of Group", "Share of Revenue", "Canonical Group Total", "Difference")
    for ($column = 0; $column -lt $costHeaders.Count; $column++) { Set-Cell $costSheet 3 ($column + 1) $costHeaders[$column] }
    $row = 4
    foreach ($year in $years) {
        $detail = $detailByYear[[string]$year]
        $annual = $annualByYear[[string]$year]
        foreach ($group in @(
            [pscustomobject]@{Name="COGS"; Items=@($detail.cost_of_goods_sold); Source=[int]$annual.cogs_usd},
            [pscustomobject]@{Name="Operating expenses"; Items=@($detail.operating_expenses); Source=[int]$annual.opex_usd}
        )) {
            $first = $row
            foreach ($item in $group.Items) {
                Set-Cell $costSheet $row 1 ([string]$year)
                Set-Cell $costSheet $row 2 $group.Name
                Set-Cell $costSheet $row 3 ([string]$item.category)
                Set-Cell $costSheet $row 4 ([int]$item.amount_usd)
                $row++
            }
            $totalRow = $row
            Set-Cell $costSheet $totalRow 1 ([string]$year)
            Set-Cell $costSheet $totalRow 2 $group.Name
            Set-Cell $costSheet $totalRow 3 "Total"
            Set-Cell $costSheet $totalRow 4 "=SUM(D$first`:D$($row - 1))" -Formula
            Set-Cell $costSheet $totalRow 5 "=SUM(E$first`:E$($row - 1))" -Formula
            $analysisRow = $analysisRowByYear[[string]$year]
            Set-Cell $costSheet $totalRow 6 "=D$totalRow/'Financial Analysis'!B$analysisRow" -Formula
            Set-Cell $costSheet $totalRow 7 ([int]$group.Source)
            Set-Cell $costSheet $totalRow 8 "=D$totalRow-G$totalRow" -Formula
            for ($r = $first; $r -lt $totalRow; $r++) {
                Set-Cell $costSheet $r 5 "=D$r/`$D`$$totalRow" -Formula
                Set-Cell $costSheet $r 6 "=D$r/'Financial Analysis'!B$analysisRow" -Formula
                $componentCell = $costSheet.Cells.Item($r, 4)
                try { $componentCell.Font.Color = 255 } finally { Release-ComObject $componentCell }
                $linkedCell = $costSheet.Cells.Item($r, 6)
                try { $linkedCell.Font.Color = 32768 } finally { Release-ComObject $linkedCell }
            }
            Add-FormulaExpected -Sheet $costSheet -Row $totalRow -Column 4 -Expected ([double]$group.Source) -Label "$year $($group.Name) total"
            Add-FormulaExpected -Sheet $costSheet -Row $totalRow -Column 5 -Expected 1.0 -Label "$year $($group.Name) share total" -Tolerance 0.000001
            Add-FormulaExpected -Sheet $costSheet -Row $totalRow -Column 8 -Expected 0.0 -Label "$year $($group.Name) reconciliation"
            $costSourceCell = $costSheet.Cells.Item($totalRow, 7)
            try { $costSourceCell.Font.Color = 255 } finally { Release-ComObject $costSourceCell }
            foreach ($column in @(4, 5, 6, 8)) {
                $costFormulaCell = $costSheet.Cells.Item($totalRow, $column)
                try { $costFormulaCell.Font.Color = if ($column -eq 6) { 32768 } else { 0 } } finally { Release-ComObject $costFormulaCell }
            }
            $row++
        }
    }
    $costLastRow = $row - 1
    $costSheet.Range("D4:D$costLastRow").NumberFormat = '$#,##0;($#,##0);-'
    $costSheet.Range("E4:F$costLastRow").NumberFormat = "0.0%;(0.0%);-"
    $costSheet.Range("G4:H$costLastRow").NumberFormat = '$#,##0;($#,##0);-'

    $kpiSheet = $sheetObjects["Operating KPIs"]
    Set-Cell $kpiSheet 1 1 "Operating Metrics and Customer Deployment"
    Set-Cell $kpiSheet 2 1 "Year-end counts and full-year flows remain separately labeled; comparisons are formula-derived where available."
    $kpiHeaders = @("Year", "Measure", "Actual", "Unit", "Prior-Year Actual", "Change (growth %, percentage points, or days)")
    for ($column = 0; $column -lt $kpiHeaders.Count; $column++) { Set-Cell $kpiSheet 3 ($column + 1) $kpiHeaders[$column] }
    $kpiSpecs = @(
        [pscustomobject]@{Key="customer_organizations_year_end"; Label="Customer organizations, year end"; Unit="organizations"; Mode="growth"},
        [pscustomobject]@{Key="connected_sites_year_end"; Label="Connected sites, year end"; Unit="sites"; Mode="growth"},
        [pscustomobject]@{Key="work_orders_completed"; Label="Work orders completed, full year"; Unit="orders"; Mode="growth"},
        [pscustomobject]@{Key="monthly_active_field_users_q4"; Label="Monthly active field users, Q4"; Unit="users"; Mode="growth"},
        [pscustomobject]@{Key="median_implementation_days"; Label="Median implementation time"; Unit="days"; Mode="days"},
        [pscustomobject]@{Key="platform_availability_pct"; Label="Platform availability"; Unit="percent"; Mode="points"},
        [pscustomobject]@{Key="logo_renewal_rate_pct"; Label="Logo renewal rate"; Unit="percent"; Mode="points"},
        [pscustomobject]@{Key="scheduled_work_completed_within_target_pct"; Label="Scheduled work completed within target"; Unit="percent"; Mode="points"}
    )
    $row = 4
    foreach ($year in $years) {
        $current = $operationsByYear[[string]$year]
        $prior = if ($year -gt $years[0]) { $operationsByYear[[string]($year - 1)] } else { $null }
        foreach ($spec in $kpiSpecs) {
            $actual = [double]$current.($spec.Key)
            if ($spec.Unit -eq "percent") { $actual /= 100 }
            Set-Cell $kpiSheet $row 1 ([string]$year)
            Set-Cell $kpiSheet $row 2 $spec.Label
            Set-Cell $kpiSheet $row 3 $actual
            Set-Cell $kpiSheet $row 4 $spec.Unit
            if ($null -ne $prior) {
                $oldValue = [double]$prior.($spec.Key)
                if ($spec.Unit -eq "percent") { $oldValue /= 100 }
                Set-Cell $kpiSheet $row 5 $oldValue
                if ($spec.Mode -eq "growth") {
                    Set-Cell $kpiSheet $row 6 "=C$row/E$row-1" -Formula
                    $growthInput = $kpiSheet.Cells.Item($row, 3)
                    try { $growthInput.Font.Color = 255 } finally { Release-ComObject $growthInput }
                    $priorInput = $kpiSheet.Cells.Item($row, 5)
                    try { $priorInput.Font.Color = 255 } finally { Release-ComObject $priorInput }
                    Add-FormulaExpected -Sheet $kpiSheet -Row $row -Column 6 -Expected ($actual / $oldValue - 1) -Label "$year $($spec.Label) growth" -Tolerance 0.000001
                } elseif ($spec.Mode -eq "days") {
                    Set-Cell $kpiSheet $row 6 "=E$row-C$row" -Formula
                    $dayChange = $kpiSheet.Cells.Item($row, 6)
                    try { $dayChange.NumberFormat = "0.0;(0.0);-" } finally { Release-ComObject $dayChange }
                    $growthInput = $kpiSheet.Cells.Item($row, 3)
                    try { $growthInput.Font.Color = 255 } finally { Release-ComObject $growthInput }
                    $priorInput = $kpiSheet.Cells.Item($row, 5)
                    try { $priorInput.Font.Color = 255 } finally { Release-ComObject $priorInput }
                    Add-FormulaExpected -Sheet $kpiSheet -Row $row -Column 6 -Expected ($oldValue - $actual) -Label "$year implementation days reduced"
                } else {
                    Set-Cell $kpiSheet $row 6 "=100*(C$row-E$row)" -Formula
                    $growthInput = $kpiSheet.Cells.Item($row, 3)
                    try { $growthInput.Font.Color = 255 } finally { Release-ComObject $growthInput }
                    $priorInput = $kpiSheet.Cells.Item($row, 5)
                    try { $priorInput.Font.Color = 255 } finally { Release-ComObject $priorInput }
                    Add-FormulaExpected -Sheet $kpiSheet -Row $row -Column 6 -Expected (100 * ($actual - $oldValue)) -Label "$year $($spec.Label) point change" -Tolerance 0.000001
                }
            } else {
                $currentInput = $kpiSheet.Cells.Item($row, 3)
                try { $currentInput.Font.Color = 255 } finally { Release-ComObject $currentInput }
            }
            if ($spec.Unit -eq "percent") {
                $actualCell = $kpiSheet.Cells.Item($row, 3)
                try { $actualCell.NumberFormat = "0.00%;(0.00%);-" } finally { Release-ComObject $actualCell }
                if ($null -ne $prior) {
                    $priorCell = $kpiSheet.Cells.Item($row, 5)
                    try { $priorCell.NumberFormat = "0.00%;(0.00%);-" } finally { Release-ComObject $priorCell }
                    $changeCell = $kpiSheet.Cells.Item($row, 6)
                    try {
                        if ($spec.Mode -eq "points") {
                            $changeCell.NumberFormat = '0.0;(0.0);- "pp"'
                        } elseif ($spec.Mode -eq "growth") {
                            $changeCell.NumberFormat = "0.0%;(0.0%);-"
                        }
                    } finally { Release-ComObject $changeCell }
                }
            }
            $row++
        }
    }
    $kpiLastRow = $row - 1
    $kpiSheet.Range("C4:C$kpiLastRow").NumberFormat = "#,##0.00;(#,##0.00);-"
    $kpiSheet.Range("E4:E$kpiLastRow").NumberFormat = "#,##0.00;(#,##0.00);-"

    $workforce = $sheetObjects["Workforce"]
    Set-Cell $workforce 1 1 "Year-End Functional Headcount and Revenue Scale"
    Set-Cell $workforce 2 1 "Headcount is point-in-time. Revenue per year-end employee is a scale measure, not an individual productivity measure."
    $workforceHeaders = @("Year", "Function", "Year-End Headcount", "Share of Headcount", "Annual Revenue", "Revenue / Year-End Employee")
    for ($column = 0; $column -lt $workforceHeaders.Count; $column++) { Set-Cell $workforce 3 ($column + 1) $workforceHeaders[$column] }
    $row = 4
    foreach ($year in $years) {
        $org = $organizationByYear[[string]$year]
        $annual = $annualByYear[[string]$year]
        $first = $row
        foreach ($functionName in $org.headcount_by_function.PSObject.Properties.Name) {
            $count = [int]$org.headcount_by_function.$functionName
            Set-Cell $workforce $row 1 ([string]$year)
            Set-Cell $workforce $row 2 $functionName
            Set-Cell $workforce $row 3 $count
            $row++
        }
        $totalRow = $row
        Set-Cell $workforce $totalRow 1 ([string]$year)
        Set-Cell $workforce $totalRow 2 "Total / scale"
        Set-Cell $workforce $totalRow 3 "=SUM(C$first`:C$($row - 1))" -Formula
        Set-Cell $workforce $totalRow 4 "=SUM(D$first`:D$($row - 1))" -Formula
        Set-Cell $workforce $totalRow 5 ([int]$annual.revenue_usd)
        Set-Cell $workforce $totalRow 6 "=E$totalRow/C$totalRow" -Formula
        for ($r = $first; $r -lt $totalRow; $r++) {
            Set-Cell $workforce $r 4 "=C$r/`$C`$$totalRow" -Formula
        }
        Add-FormulaExpected -Sheet $workforce -Row $totalRow -Column 3 -Expected ([double]$org.year_end_headcount) -Label "$year headcount total"
        Add-FormulaExpected -Sheet $workforce -Row $totalRow -Column 4 -Expected 1.0 -Label "$year headcount share total" -Tolerance 0.000001
        Add-FormulaExpected -Sheet $workforce -Row $totalRow -Column 6 -Expected ([double]($annual.revenue_usd / $org.year_end_headcount)) -Label "$year revenue per year-end employee"
        for ($r = $first; $r -lt $totalRow; $r++) {
            $functionInput = $workforce.Cells.Item($r, 3)
            try { $functionInput.Font.Color = 255 } finally { Release-ComObject $functionInput }
        }
        $headcountFormula = $workforce.Cells.Item($totalRow, 3)
        try { $headcountFormula.Font.Color = 0 } finally { Release-ComObject $headcountFormula }
        $annualRevenueInput = $workforce.Cells.Item($totalRow, 5)
        try { $annualRevenueInput.Font.Color = 255 } finally { Release-ComObject $annualRevenueInput }
        $row++
    }
    $workforceLastRow = $row - 1
    $workforce.Range("D4:D$workforceLastRow").NumberFormat = "0.0%;(0.0%);-"
    $workforce.Range("E4:F$workforceLastRow").NumberFormat = '$#,##0;($#,##0);-'

    $strategySheet = $sheetObjects["Strategy and Risks"]
    Set-Cell $strategySheet 1 1 "Strategy Execution and Risk Controls"
    Set-Cell $strategySheet 2 1 "Recorded year-end status and follow-through. This is an operating context sheet, not a forecast."
    $strategyHeaders = @("Topic", "Status / Evidence", "Control / Next Step", "Owner")
    for ($column = 0; $column -lt $strategyHeaders.Count; $column++) { Set-Cell $strategySheet 3 ($column + 1) $strategyHeaders[$column] }
    $row = 4
    foreach ($priority in $artifact.strategy.priorities) {
        Set-Cell $strategySheet $row 1 ([string]$priority.priority)
        Set-Cell $strategySheet $row 2 ([string]$priority.status + ": " + [string]$priority.evidence)
        Set-Cell $strategySheet $row 3 ([string]$priority.follow_through)
        Set-Cell $strategySheet $row 4 "Management"
        $row++
    }
    foreach ($risk in $artifact.risks.risk_factors) {
        $riskTitle = [Globalization.CultureInfo]::InvariantCulture.TextInfo.ToTitleCase(
            ([string]$risk.risk_id).Replace("-", " ")
        )
        Set-Cell $strategySheet $row 1 $riskTitle
        Set-Cell $strategySheet $row 2 ([string]$risk.exposure + " Evidence: " + [string]$risk.evidence)
        Set-Cell $strategySheet $row 3 ([string]$risk.mitigation + " Residual: " + [string]$risk.residual)
        $owner = [string]$risk.owner_id
        $ownerClaim = $draft.claims | Where-Object { $_.record_id -eq $owner } | Select-Object -First 1
        if ($null -eq $ownerClaim -or $null -eq $ownerClaim.value.name) {
            throw "The date-filtered draft has no canonical owner name for risk $($risk.risk_id)."
        }
        $ownerName = [string]$ownerClaim.value.name
        Set-Cell $strategySheet $row 4 $ownerName
        $row++
    }
    $nextPeriodPriorities = @($artifact.strategy.next_period_priorities)
    if ($nextPeriodPriorities.Count -gt 0) {
        Set-Cell $strategySheet $row 1 "Next-period priorities"
        Set-Cell $strategySheet $row 2 ([string]::Join("; ", $nextPeriodPriorities))
        Set-Cell $strategySheet $row 3 "Review through normal operating controls; not financial guidance."
        Set-Cell $strategySheet $row 4 "Management"
        $row++
    }

    Format-Worksheet -Sheet $quarterSheet -LastColumn 11 -LastRow $quarterLastRow
    $quarterSheet.Range("A1:K1").UnMerge()
    $quarterSheet.Range("A1:B1").Merge()
    $quarterSheet.Range("A1:B1").WrapText = $true
    $quarterSheet.Range("A2:B2").Merge()
    $quarterSheet.Range("A2:B2").WrapText = $true
    $quarterSheet.Rows.Item(1).RowHeight = 42
    $quarterSheet.Rows.Item(2).RowHeight = 54
    $quarterSheet.PageSetup.FitToPagesWide = 2
    $quarterSheet.PageSetup.PrintTitleColumns = '$A:$B'
    $quarterSheet.PageSetup.CenterHorizontally = $true
    Format-Worksheet -Sheet $analysis -LastColumn 16 -LastRow ($years.Count + 3)
    Format-Worksheet -Sheet $productSheet -LastColumn 6 -LastRow $productLastRow
    Format-Worksheet -Sheet $segmentSheet -LastColumn 6 -LastRow $segmentLastRow
    Format-Worksheet -Sheet $costSheet -LastColumn 8 -LastRow $costLastRow
    $costSheet.Columns.Item(4).ColumnWidth = 15
    $costSheet.Columns.Item(7).ColumnWidth = 19
    $costSheet.Columns.Item(8).ColumnWidth = 12
    $costSheet.Rows.Item(3).AutoFit()
    Format-Worksheet -Sheet $kpiSheet -LastColumn 6 -LastRow $kpiLastRow
    Format-Worksheet -Sheet $workforce -LastColumn 6 -LastRow $workforceLastRow
    Format-Worksheet -Sheet $strategySheet -LastColumn 4 -LastRow ($row - 1)
    $strategySheet.PageSetup.FitToPagesTall = 2
    Format-Worksheet -Sheet $start -LastColumn 1 -LastRow 23 -NoHeaderRow
    $start.Columns.Item(1).ColumnWidth = 95
    $start.Range("A1:A23").Rows.AutoFit()
    $start.Rows.Item(1).RowHeight = 28
    $start.PageSetup.FitToPagesTall = 1
    $start.PageSetup.CenterHorizontally = $true

    $app.CalculateFullRebuild()
    $workbook.SaveAs($outputPath, 51)
    $workbook.Save()
    $workbook.ExportAsFixedFormat(0, $pdfPath)
    $workbook.Close($false)
    foreach ($sheetObject in @($sheetObjects.Values)) {
        Release-ComObject $sheetObject
    }
    $sheetObjects.Clear()
    Release-ComObject $worksheetCollection
    $worksheetCollection = $null
    Release-ComObject $workbook
    $workbook = $null
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Excel saved and rendered the workbook."

    $reopenedWorkbook = $app.Workbooks.Open($outputPath, 0, $true)
    $reopenedCollection = $reopenedWorkbook.Worksheets
    $sheetNames = @()
    for ($sheetIndex = 1; $sheetIndex -le $reopenedCollection.Count; $sheetIndex++) {
        $sheet = $reopenedCollection.Item($sheetIndex)
        $sheetName = [string]$sheet.Name
        if ([string]::IsNullOrWhiteSpace($sheetName) -or $reopenedSheets.ContainsKey($sheetName)) {
            throw "Excel reopened a workbook with a missing or duplicate worksheet name."
        }
        $reopenedSheets[$sheetName] = $sheet
        $sheetNames += $sheetName
    }
    if (-not $reopenedSheets.ContainsKey("Financial Analysis")) {
        throw "Excel did not retain the Financial Analysis worksheet after reopen."
    }
    foreach ($expected in $expectedCells) {
        $checkSheet = $reopenedSheets[[string]$expected.Sheet]
        if ($null -eq $checkSheet) {
            throw "Excel did not retain expected worksheet $($expected.Sheet)."
        }
        $cell = $checkSheet.Cells.Item($expected.Row, $expected.Column)
        try {
            $actual = [double]$cell.Value2
            $formula = [string]$cell.Formula
            if (-not $formula.StartsWith("=")) {
                throw "$($expected.Label): expected a retained formula at $($expected.Sheet)!R$($expected.Row)C$($expected.Column)."
            }
            if ([Math]::Abs($actual - $expected.Value) -gt $expected.Tolerance) {
                throw "$($expected.Label): expected $($expected.Value), reopened cached value was $actual."
            }
        } finally {
            Release-ComObject $cell
            $cell = $null
        }
    }
    $formulaErrors = @()
    $formulaCount = 0
    foreach ($sheetName in $sheetNames) {
        $scanSheet = $reopenedSheets[$sheetName]
        try {
            $usedRange = $scanSheet.UsedRange
            for ($scanRow = 1; $scanRow -le $usedRange.Rows.Count; $scanRow++) {
                for ($scanColumn = 1; $scanColumn -le $usedRange.Columns.Count; $scanColumn++) {
                    $cell = $usedRange.Cells.Item($scanRow, $scanColumn)
                    try {
                        if ($cell.HasFormula) {
                            $formulaCount++
                        }
                        if ([string]$cell.Text -match "^#(REF!|DIV/0!|VALUE!|N/A|NAME\?|NUM!|NULL!)$") {
                            $formulaErrors += [string]$cell.Address($false, $false)
                        }
                    } finally {
                        Release-ComObject $cell
                        $cell = $null
                    }
                }
            }
        } finally {
            Release-ComObject $usedRange
            $usedRange = $null
        }
    }
    if ($formulaErrors.Count -gt 0) {
        throw "Excel formula errors found: $($formulaErrors -join ', ')"
    }
    if ($formulaCount -lt 20 -or $expectedCells.Count -lt 20) {
        throw "Excel did not retain the expected linked analysis formulas."
    }
    $newProcesses = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue |
        Where-Object { $baselinePids -notcontains $_.Id })
    if ($newProcesses.Count -ne 1) {
        throw "Could not verify a unique run-owned Excel process after rendering."
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
        formula_count = $formulaCount
        expected_formula_values_verified = $expectedCells.Count
        formula_errors = $formulaErrors.Count
        rendered_pdf = [ordered]@{ path = $pdfInfo.FullName; bytes = $pdfInfo.Length }
        sheets = $sheetNames
        analysis_modules = $artifact.module_names
    }
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Excel reopened cached formulas; verified $($expectedCells.Count) expected values and zero formula errors."
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
    if ($null -ne $cell) {
        try { Release-ComObject $cell } catch { $cleanupErrors += "Excel cell release failed: $($_.Exception.Message)" }
    }
    if ($null -ne $usedRange) {
        try { Release-ComObject $usedRange } catch { $cleanupErrors += "Excel used-range release failed: $($_.Exception.Message)" }
    }
    if ($null -ne $reopenedWorkbook) {
        try { $reopenedWorkbook.Close($false) } catch { $cleanupErrors += "Excel reopened workbook close failed: $($_.Exception.Message)" }
        foreach ($sheetObject in @($reopenedSheets.Values)) {
            try { Release-ComObject $sheetObject } catch { $cleanupErrors += "Excel reopened worksheet release failed: $($_.Exception.Message)" }
        }
        if ($null -ne $reopenedCollection) {
            try { Release-ComObject $reopenedCollection } catch { $cleanupErrors += "Excel reopened worksheet collection release failed: $($_.Exception.Message)" }
        }
        try { Release-ComObject $reopenedWorkbook } catch { $cleanupErrors += "Excel reopened workbook release failed: $($_.Exception.Message)" }
    }
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { $cleanupErrors += "Excel workbook close failed: $($_.Exception.Message)" }
        foreach ($sheetObject in @($sheetObjects.Values)) {
            try { Release-ComObject $sheetObject } catch { $cleanupErrors += "Excel worksheet release failed: $($_.Exception.Message)" }
        }
        if ($null -ne $workbookProperties) {
            try { Release-ComObject $workbookProperties } catch { $cleanupErrors += "Excel document properties release failed: $($_.Exception.Message)" }
        }
        try { Release-ComObject $workbook } catch { $cleanupErrors += "Excel workbook release failed: $($_.Exception.Message)" }
    }
    if ($null -ne $worksheetCollection) {
        try { Release-ComObject $worksheetCollection } catch { $cleanupErrors += "Excel worksheet collection release failed: $($_.Exception.Message)" }
    }
    if ($null -ne $app) {
        if ($null -ne $processId) {
            try {
                $app.Quit()
                Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Excel application quit."
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
