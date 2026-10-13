$ErrorActionPreference = "Stop"

if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    throw "PowerPoint COM automation requires a single-threaded apartment (STA)."
}

$contextPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_CONTEXT_PATH)
$draftPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_DRAFT_PATH)
$outputPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_OUTPUT_PATH)
$pdfPath = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_PDF_PATH)
$previewDirectory = [IO.Path]::GetFullPath($env:JIMWORKS_OFFICE_EXPORT_PREVIEW_DIRECTORY)
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
if (Test-Path -LiteralPath $previewDirectory) {
    throw "Refusing to overwrite an existing preview directory: $previewDirectory"
}
[void][IO.Directory]::CreateDirectory($previewDirectory)

$draft = Get-Content -LiteralPath $draftPath -Raw | ConvertFrom-Json
$context = Get-Content -LiteralPath $contextPath -Raw | ConvertFrom-Json
$companyName = [string]$context.company.data.name
if ([string]::IsNullOrWhiteSpace($companyName)) {
    throw "Date-filtered context does not contain a company name."
}
$sections = @($draft.sections)
$slideSpecs = if ($null -ne $draft.slides) {
    @($draft.slides)
} else {
    @($sections | ForEach-Object {
        [pscustomobject]@{
            heading = $_.heading
            body = @([string]$_.text)
            layout = "content"
        }
    })
}
if ($slideSpecs.Count -eq 0 -or $slideSpecs.Count -ne $sections.Count) {
    throw "PPTX draft must have one prepared slide for each grounded section."
}

$app = $null
$presentations = $null
$presentation = $null
$slides = $null
$slideObjects = @()
$failure = $null
$cleanupErrors = @()
$processId = $null
$evidence = $null
$expectedText = [System.Collections.Generic.List[string]]::new()

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

function Get-RgbColor {
    param([int]$Red, [int]$Green, [int]$Blue)
    $Red + ($Green -shl 8) + ($Blue -shl 16)
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
            throw "PowerPoint did not retain the $Name document property."
        }
    } finally {
        Release-ComObject $property
    }
}

function Get-TextLines {
    param([object]$Value, [string]$Label)
    if ($Value -is [string]) {
        $lines = @($Value -split "\r?\n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    } elseif ($Value -is [array] -and @($Value | Where-Object { $_ -isnot [string] }).Count -eq 0) {
        $lines = @($Value | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    } else {
        throw "$Label must be text or a list of text."
    }
    if ($lines.Count -eq 0) { throw "$Label must not be empty." }
    $lines
}

function Add-TextShape {
    param(
        [Parameter(Mandatory = $true)][object]$Slide,
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][double]$Left,
        [Parameter(Mandatory = $true)][double]$Top,
        [Parameter(Mandatory = $true)][double]$Width,
        [Parameter(Mandatory = $true)][double]$Height,
        [int]$FontSize = 16,
        [int]$Color = (Get-RgbColor 44 55 64),
        [bool]$Bold = $false,
        [int]$Alignment = 1
    )
    $shape = $Slide.Shapes.AddTextbox(1, $Left, $Top, $Width, $Height)
    $step = "hide text box outline"
    try {
        $shape.Line.Visible = 0
        $frame = $shape.TextFrame
        $step = "enable text wrapping"
        $frame.WordWrap = -1
        $step = "fit text in text box"
        $shape.TextFrame2.AutoSize = 2
        $range = $frame.TextRange
        $step = "set text"
        $range.Text = $Text
        $step = "set font family"
        $range.Font.Name = "Aptos"
        $step = "set font size"
        $range.Font.Size = $FontSize
        $step = "set font weight"
        $range.Font.Bold = if ($Bold) { -1 } else { 0 }
        $step = "set font color"
        $range.Font.Color.RGB = $Color
        $step = "set paragraph alignment"
        $range.ParagraphFormat.Alignment = $Alignment
    } catch {
        throw "PowerPoint could not $($step): $($_.Exception.Message)"
    }
    $shape
}

function Add-Card {
    param(
        [Parameter(Mandatory = $true)][object]$Slide,
        [Parameter(Mandatory = $true)][string]$Text,
        [Parameter(Mandatory = $true)][double]$Left,
        [Parameter(Mandatory = $true)][double]$Top,
        [Parameter(Mandatory = $true)][double]$Width,
        [Parameter(Mandatory = $true)][double]$Height,
        [int]$Fill = (Get-RgbColor 233 242 245),
        [int]$FontSize = 16
    )
    $shape = $Slide.Shapes.AddShape(5, $Left, $Top, $Width, $Height)
    $shape.Fill.Solid()
    $shape.Fill.ForeColor.RGB = $Fill
    $shape.Line.Visible = 0
    $frame = $shape.TextFrame
    $frame.WordWrap = -1
    $shape.TextFrame2.AutoSize = 2
    $frame.MarginLeft = 12
    $frame.MarginRight = 12
    $frame.MarginTop = 10
    $frame.MarginBottom = 8
    $range = $frame.TextRange
    $range.Text = $Text
    $range.Font.Name = "Aptos"
    $range.Font.Size = $FontSize
    $range.Font.Color.RGB = (Get-RgbColor 37 68 92)
    $range.ParagraphFormat.Alignment = 1
    $shape
}

function Add-Body {
    param(
        [Parameter(Mandatory = $true)][object]$Slide,
        [Parameter(Mandatory = $true)][string[]]$Lines,
        [Parameter(Mandatory = $true)][string]$Layout,
        [Parameter(Mandatory = $true)][double]$Top,
        [Parameter(Mandatory = $true)][double]$Height
    )
    $gap = 14.0
    if ($Layout -in @("cards", "grid", "stats", "workflow", "capabilities") -and $Lines.Count -gt 1) {
        $columns = if ($Layout -eq "cards") { [Math]::Min(2, $Lines.Count) } else { [Math]::Min(3, $Lines.Count) }
        $rows = [Math]::Ceiling($Lines.Count / $columns)
        $width = (900 - $gap * ($columns - 1)) / $columns
        $cardHeight = ($Height - $gap * ($rows - 1)) / $rows
        for ($index = 0; $index -lt $Lines.Count; $index++) {
            $row = [Math]::Floor($index / $columns)
            $column = $index % $columns
            $fill = switch ($index % 3) {
                0 { Get-RgbColor 233 242 245 }
                1 { Get-RgbColor 244 247 248 }
                2 { Get-RgbColor 232 242 239 }
            }
            [void](Add-Card -Slide $Slide -Text $Lines[$index] `
                -Left (30 + $column * ($width + $gap)) -Top ($Top + $row * ($cardHeight + $gap)) `
                -Width $width -Height $cardHeight -Fill $fill `
                -FontSize $(if ($Layout -eq "stats") { 20 } else { 15 }))
        }
        return
    }
    $text = [string]::Join("`r`n`r`n", $Lines)
    $fill = if ($Layout -eq "closing") { Get-RgbColor 37 68 92 } else { Get-RgbColor 233 242 245 }
    if ($Layout -in @("statement", "closing")) {
        [void](Add-Card -Slide $Slide -Text $text -Left 34 -Top $Top `
            -Width 892 -Height $Height -Fill $fill -FontSize 20)
    } else {
        $color = if ($Layout -eq "cover") { Get-RgbColor 37 68 92 } else { Get-RgbColor 44 55 64 }
        $fontSize = if ($Layout -eq "cover") { 22 } else { 15 }
        [void](Add-TextShape -Slide $Slide -Text $text -Left 38 -Top $Top `
            -Width 884 -Height $Height -FontSize $fontSize -Color $color)
    }
}

function Add-Chart {
    param(
        [Parameter(Mandatory = $true)][object]$Slide,
        [Parameter(Mandatory = $true)][object]$Chart,
        [Parameter(Mandatory = $true)][double]$Top
    )
    $labels = @($Chart.labels | ForEach-Object { [string]$_ })
    $values = @($Chart.values | ForEach-Object { [double]$_ })
    $maxMagnitude = [Math]::Max(1.0, [double](($values | ForEach-Object { [Math]::Abs($_) } | Measure-Object -Maximum).Maximum))
    if (-not [string]::IsNullOrWhiteSpace([string]$Chart.title)) {
        [void](Add-TextShape -Slide $Slide -Text ([string]$Chart.title) -Left 40 -Top $Top `
            -Width 880 -Height 26 -FontSize 13 -Color (Get-RgbColor 37 68 92) -Bold $true)
        $Top += 28
    }
    $rowHeight = [Math]::Min(27.0, [Math]::Max(16.0, (208.0 / [Math]::Max(1, $labels.Count))))
    $barLeft = 205.0
    $barMaxWidth = 540.0
    for ($index = 0; $index -lt $labels.Count; $index++) {
        $y = $Top + $index * $rowHeight
        $valueText = $values[$index].ToString("G", [Globalization.CultureInfo]::InvariantCulture)
        [void](Add-TextShape -Slide $Slide -Text $labels[$index] -Left 40 -Top $y `
            -Width 155 -Height ($rowHeight - 2) -FontSize 10)
        $barWidth = [Math]::Max(2.0, $barMaxWidth * [Math]::Abs($values[$index]) / $maxMagnitude)
        $bar = $Slide.Shapes.AddShape(1, $barLeft, $y + 3, $barWidth, ($rowHeight - 7))
        $bar.Fill.Solid()
        $bar.Fill.ForeColor.RGB = (Get-RgbColor 34 115 145)
        $bar.Line.Visible = 0
        [void](Add-TextShape -Slide $Slide -Text $valueText -Left ($barLeft + $barWidth + 8) `
            -Top $y -Width 165 -Height ($rowHeight - 2) -FontSize 10)
    }
}

function Add-PptTable {
    param(
        [Parameter(Mandatory = $true)][object]$Slide,
        [Parameter(Mandatory = $true)][object]$Specification,
        [Parameter(Mandatory = $true)][double]$Top
    )
    $headers = @($Specification.headers)
    $dataRows = @($Specification.rows)
    if ($Specification.headers -isnot [array] -or $headers.Count -eq 0 -or
        @($headers | Where-Object { $_ -isnot [string] -or [string]::IsNullOrWhiteSpace($_) }).Count -gt 0) {
        throw "PowerPoint tables require non-empty text headers."
    }
    if ($Specification.rows -isnot [array] -or $dataRows.Count -eq 0) {
        throw "PowerPoint tables require at least one data row."
    }
    foreach ($row in $dataRows) {
        if ($row -isnot [array] -or $row.Count -ne $headers.Count) {
            throw "Every PowerPoint table row must match the header count."
        }
    }
    $rows = $dataRows.Count + 1
    $columns = $headers.Count
    $tableHeight = [Math]::Min(210.0, [Math]::Max(56.0, 36.0 * $rows))
    $tableShape = $Slide.Shapes.AddTable($rows, $columns, 36, $Top, 888, $tableHeight)
    $table = $tableShape.Table
    for ($rowIndex = 0; $rowIndex -lt $rows; $rowIndex++) {
        for ($columnIndex = 0; $columnIndex -lt $columns; $columnIndex++) {
            $value = if ($rowIndex -eq 0) {
                $headers[$columnIndex]
            } else {
                $dataRows[$rowIndex - 1][$columnIndex]
            }
            if ($null -eq $value) { $text = "" }
            elseif ($value -is [bool]) { $text = if ($value) { "Yes" } else { "No" } }
            elseif ($value -is [string] -or $value -is [ValueType]) { $text = [string]$value }
            else { $text = $value | ConvertTo-Json -Depth 10 -Compress }
            $cell = $table.Cell($rowIndex + 1, $columnIndex + 1)
            try {
                $cell.Shape.TextFrame.TextRange.Text = $text
                $cell.Shape.Fill.Solid()
                $cell.Shape.Fill.ForeColor.RGB = if ($rowIndex -eq 0) {
                    Get-RgbColor 37 68 92
                } else {
                    Get-RgbColor 255 255 255
                }
                $cell.Shape.TextFrame.TextRange.Font.Name = "Aptos"
                $cell.Shape.TextFrame.TextRange.Font.Size = if ($rowIndex -eq 0) { 11 } else { 9 }
                $cell.Shape.TextFrame.TextRange.Font.Bold = if ($rowIndex -eq 0) { -1 } else { 0 }
                $cell.Shape.TextFrame.TextRange.Font.Color.RGB = if ($rowIndex -eq 0) {
                    Get-RgbColor 255 255 255
                } else {
                    Get-RgbColor 44 55 64
                }
            } finally {
                Release-ComObject $cell
            }
            $expectedText.Add($text)
        }
    }
    Release-ComObject $table
    Release-ComObject $tableShape
}

function Normalize-Text {
    param([string]$Value)
    [regex]::Replace($Value, "\s+", " ").Trim().ToLowerInvariant()
}

function Get-RunOwnedPowerPointProcessId {
    param([int[]]$BaselinePids)
    for ($attempt = 0; $attempt -lt 30; $attempt++) {
        $newProcesses = @(Get-Process -Name POWERPNT -ErrorAction SilentlyContinue |
            Where-Object { $BaselinePids -notcontains $_.Id })
        if ($newProcesses.Count -eq 1) {
            return [int]$newProcesses[0].Id
        }
        if ($newProcesses.Count -gt 1) {
            throw "Could not verify a unique run-owned PowerPoint process."
        }
        Start-Sleep -Milliseconds 100
    }
    throw "Could not verify a run-owned PowerPoint process."
}

try {
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) Starting generic PowerPoint authoring."
    $app = New-Object -ComObject PowerPoint.Application
    $processId = Get-RunOwnedPowerPointProcessId -BaselinePids $baselinePids
    $app.Visible = -1
    $app.WindowState = 2
    $presentations = $app.Presentations
    $presentation = $presentations.Add()
    $presentation.PageSetup.SlideWidth = 960
    $presentation.PageSetup.SlideHeight = 540
    $presentationProperties = $presentation.BuiltInDocumentProperties
    Set-OfficeDocumentProperty -Properties $presentationProperties -Name "Title" -Value ([string]$draft.title)
    Set-OfficeDocumentProperty -Properties $presentationProperties -Name "Author" -Value $companyName
    Release-ComObject $presentationProperties

    for ($index = 0; $index -lt $sections.Count; $index++) {
        $section = $sections[$index]
        $specification = $slideSpecs[$index]
        $heading = [string]$specification.heading
        $bodyLines = @(Get-TextLines -Value $specification.body -Label "Slide $($index + 1) body")
        $layout = if ([string]::IsNullOrWhiteSpace([string]$specification.layout)) { "content" } else { [string]$specification.layout }
        $slide = $presentation.Slides.Add($index + 1, 12)
        $slideObjects += $slide
        $slide.FollowMasterBackground = 0
        $slide.Background.Fill.Solid()
        $slide.Background.Fill.ForeColor.RGB = (Get-RgbColor 255 255 255)

        [void](Add-TextShape -Slide $slide -Text ([string]$draft.title) -Left 38 -Top 10 `
            -Width 884 -Height 20 -FontSize 9 -Color (Get-RgbColor 105 118 126) -Bold $true)
        $expectedText.Add([string]$draft.title)
        $titleColor = if ($layout -eq "closing") { Get-RgbColor 35 139 132 } else { Get-RgbColor 37 68 92 }
        [void](Add-TextShape -Slide $slide -Text $heading -Left 38 -Top 38 `
            -Width 884 -Height 45 -FontSize 24 -Color $titleColor -Bold $true)
        $expectedText.Add($heading)
        if (-not [string]::IsNullOrWhiteSpace([string]$specification.subtitle)) {
            [void](Add-TextShape -Slide $slide -Text ([string]$specification.subtitle) -Left 40 -Top 82 `
                -Width 880 -Height 26 -FontSize 13 -Color (Get-RgbColor 105 118 126))
            $expectedText.Add([string]$specification.subtitle)
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$specification.badge)) {
            [void](Add-TextShape -Slide $slide -Text ([string]$specification.badge) -Left 725 -Top 84 `
                -Width 195 -Height 22 -FontSize 9 -Color (Get-RgbColor 35 139 132) -Bold $true -Alignment 3)
            $expectedText.Add([string]$specification.badge)
        }
        $hasChart = $null -ne $specification.chart
        $sectionTables = @(Get-OptionalArray -Object $section -Name "tables")
        $slideTables = $sectionTables + @(Get-OptionalArray -Object $specification -Name "tables")
        if ($slideTables.Count -gt 1) {
            throw "Slide $($index + 1) has multiple tables; native PowerPoint exports one table per slide."
        }
        if ($hasChart -and $slideTables.Count -gt 0) {
            throw "Slide $($index + 1) cannot combine a chart and a table."
        }
        $hasData = $hasChart -or $slideTables.Count -gt 0
        $bodyHeight = if ($hasData) { 118.0 } else { 352.0 }
        $bodyTop = if ([string]::IsNullOrWhiteSpace([string]$specification.subtitle)) { 100.0 } else { 116.0 }
        Add-Body -Slide $slide -Lines $bodyLines -Layout $layout -Top $bodyTop -Height $bodyHeight
        foreach ($line in $bodyLines) { $expectedText.Add($line) }
        if ($hasChart) {
            if (-not [string]::IsNullOrWhiteSpace([string]$specification.chart.title)) {
                $expectedText.Add([string]$specification.chart.title)
            }
            foreach ($label in $specification.chart.labels) { $expectedText.Add([string]$label) }
            foreach ($value in $specification.chart.values) {
                $expectedText.Add(([double]$value).ToString("G", [Globalization.CultureInfo]::InvariantCulture))
            }
            Add-Chart -Slide $slide -Chart $specification.chart -Top 234
        } elseif ($slideTables.Count -gt 0) {
            $table = $slideTables[0]
            if (-not [string]::IsNullOrWhiteSpace([string]$table.title)) {
                [void](Add-TextShape -Slide $slide -Text ([string]$table.title) -Left 38 -Top 232 `
                    -Width 884 -Height 24 -FontSize 12 -Color (Get-RgbColor 37 68 92) -Bold $true)
                $expectedText.Add([string]$table.title)
            }
            Add-PptTable -Slide $slide -Specification $table -Top 258
        }
        $footerText = "$companyName  |  $($draft.publication_date)  |  $($index + 1)"
        [void](Add-TextShape -Slide $slide -Text $footerText -Left 38 -Top 514 `
            -Width 884 -Height 16 -FontSize 8 -Color (Get-RgbColor 105 118 126) -Alignment 3)
        $expectedText.Add($companyName)
        $expectedText.Add([string]$draft.publication_date)

        $notes = if ([string]::IsNullOrWhiteSpace([string]$specification.notes)) {
            [string]$section.text
        } else {
            [string]$specification.notes
        }
        $notesPage = $slide.NotesPage
        $notesPlaceholder = $notesPage.Shapes.Placeholders.Item(2)
        $notesPlaceholder.TextFrame.TextRange.Text = $notes
        $expectedText.Add($notes)
        Release-ComObject $notesPlaceholder
        Release-ComObject $notesPage
    }

    $presentation.SaveAs($outputPath, 24)
    $presentation.Close()
    foreach ($slideObject in $slideObjects) { Release-ComObject $slideObject }
    $slideObjects = @()
    Release-ComObject $presentation
    $presentation = $null
    $presentation = $presentations.Open($outputPath, -1, 0, 0)
    $slides = $presentation.Slides
    if ($slides.Count -ne $sections.Count) {
        throw "PowerPoint reopened $($slides.Count) slides; expected $($sections.Count)."
    }

    $actualText = [System.Collections.Generic.List[string]]::new()
    for ($slideIndex = 1; $slideIndex -le $slides.Count; $slideIndex++) {
        $slide = $slides.Item($slideIndex)
        $shapes = $slide.Shapes
        for ($shapeIndex = 1; $shapeIndex -le $shapes.Count; $shapeIndex++) {
            $shape = $shapes.Item($shapeIndex)
            try {
                if ([int]$shape.HasTextFrame -eq -1 -and [int]$shape.TextFrame.HasText -eq -1) {
                    $actualText.Add([string]$shape.TextFrame.TextRange.Text)
                }
                if ([int]$shape.HasTable -eq -1) {
                    $table = $shape.Table
                    for ($row = 1; $row -le $table.Rows.Count; $row++) {
                        for ($column = 1; $column -le $table.Columns.Count; $column++) {
                            $actualText.Add([string]$table.Cell($row, $column).Shape.TextFrame.TextRange.Text)
                        }
                    }
                    Release-ComObject $table
                }
            } finally {
                Release-ComObject $shape
            }
        }
        Release-ComObject $shapes
        $notesPage = $slide.NotesPage
        $notesPlaceholder = $notesPage.Shapes.Placeholders.Item(2)
        $actualText.Add([string]$notesPlaceholder.TextFrame.TextRange.Text)
        Release-ComObject $notesPlaceholder
        Release-ComObject $notesPage

        $previewPath = Join-Path $previewDirectory ("slide-{0:D3}.png" -f $slideIndex)
        $slide.Export($previewPath, "PNG", 1280, 720)
        Release-ComObject $slide
    }
    $actualNormalized = Normalize-Text ([string]::Join("`n", $actualText))
    foreach ($expected in $expectedText) {
        $normalizedExpected = Normalize-Text $expected
        if ($normalizedExpected.Length -gt 0 -and -not $actualNormalized.Contains($normalizedExpected)) {
            throw "PowerPoint reopen extraction did not contain expected text: $expected"
        }
    }
    $presentation.SaveAs($pdfPath, 32)
    if ($null -eq (Get-Process -Id $processId -ErrorAction SilentlyContinue)) {
        throw "The run-owned PowerPoint process exited before export verification."
    }
    $pptxInfo = Get-Item -LiteralPath $outputPath
    $pdfInfo = Get-Item -LiteralPath $pdfPath
    $previews = @(Get-ChildItem -LiteralPath $previewDirectory -File -Filter "*.png")
    if ($pptxInfo.Length -lt 1 -or $pdfInfo.Length -lt 1 -or $previews.Count -ne $sections.Count) {
        throw "PowerPoint did not create the PPTX, PDF, and one PNG preview per slide."
    }
    if (@($previews | Where-Object { $_.Length -lt 1 }).Count -gt 0) {
        throw "PowerPoint created an empty slide preview."
    }
    $evidence = [ordered]@{
        document_id = [string]$draft.document_id
        application = [string]$app.Name
        version = [string]$app.Version
        process_id = $processId
        pptx = [ordered]@{ path = $pptxInfo.FullName; bytes = $pptxInfo.Length }
        slide_count = [int]$slides.Count
        slides_rendered = $previews.Count
        preview_paths = @($previews | ForEach-Object { $_.FullName })
        extracted_text_match = $true
        rendered_pdf = [ordered]@{ path = $pdfInfo.FullName; bytes = $pdfInfo.Length }
    }
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) PowerPoint saved, reopened, extracted, and rendered every slide."
}
catch {
    $failure = @(
        $_.Exception.ToString()
        $_.InvocationInfo.PositionMessage
        $_.ScriptStackTrace
    ) -join "`n"
    Add-Content -LiteralPath $logPath -Value "$(Get-Date -Format o) PowerPoint export failed: $failure"
}
finally {
    if ($null -ne $presentation) {
        try { $presentation.Close() } catch { $cleanupErrors += "PowerPoint presentation close failed: $($_.Exception.Message)" }
    }
    foreach ($slideObject in $slideObjects) {
        try { Release-ComObject $slideObject } catch { $cleanupErrors += "PowerPoint slide release failed: $($_.Exception.Message)" }
    }
    foreach ($object in @($slides, $presentation, $presentations)) {
        if ($null -ne $object) {
            try { Release-ComObject $object } catch { $cleanupErrors += "PowerPoint COM object release failed: $($_.Exception.Message)" }
        }
    }
    if ($null -ne $app) {
        if ($null -ne $processId) {
            try { $app.Quit() } catch { $cleanupErrors += "PowerPoint application quit failed: $($_.Exception.Message)" }
        } else {
            $cleanupErrors += "PowerPoint process identity was not verified before quit; launcher cleanup is required."
        }
        try { Release-ComObject $app } catch { $cleanupErrors += "PowerPoint application release failed: $($_.Exception.Message)" }
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
