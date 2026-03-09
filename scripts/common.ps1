Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

function Get-ProjectRoot {
    return (Split-Path -Path $PSScriptRoot -Parent)
}

function Ensure-Directory {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Resolve-AbsolutePath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    return [System.IO.Path]::GetFullPath((Resolve-Path -LiteralPath $Path).Path)
}

function Write-JsonFile {
    param(
        [Parameter(Mandatory)]
        [object]$Data,
        [Parameter(Mandatory)]
        [string]$Path,
        [int]$Depth = 100
    )

    $json = $Data | ConvertTo-Json -Depth $Depth
    [System.IO.File]::WriteAllText($Path, $json, [System.Text.Encoding]::UTF8)
}

function Write-JsonLineFile {
    param(
        [Parameter(Mandatory)]
        [System.Collections.IEnumerable]$Items,
        [Parameter(Mandatory)]
        [string]$Path,
        [int]$Depth = 50
    )

    $writer = [System.IO.StreamWriter]::new($Path, $false, [System.Text.Encoding]::UTF8)
    try {
        foreach ($item in $Items) {
            $line = $item | ConvertTo-Json -Depth $Depth -Compress
            $writer.WriteLine($line)
        }
    }
    finally {
        $writer.Dispose()
    }
}

function Get-TimestampJst {
    return (Get-Date).ToString("yyyy-MM-dd HH:mm 'JST'")
}

function Convert-ExcelColor {
    param(
        $ColorValue
    )

    if ($null -eq $ColorValue) {
        return $null
    }

    try {
        $number = [int64]$ColorValue
    }
    catch {
        return $null
    }

    if ($number -lt 0) {
        return $null
    }

    $red = $number -band 0xFF
    $green = ($number -shr 8) -band 0xFF
    $blue = ($number -shr 16) -band 0xFF
    return ('#{0:X2}{1:X2}{2:X2}' -f $red, $green, $blue)
}

function Convert-VariantValue {
    param(
        $Value
    )

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return $null
    }

    if ($Value -is [DateTime]) {
        return $Value.ToString('o')
    }

    if ($Value -is [bool] -or
        $Value -is [byte] -or
        $Value -is [int16] -or
        $Value -is [int32] -or
        $Value -is [int64] -or
        $Value -is [single] -or
        $Value -is [double] -or
        $Value -is [decimal]) {
        return $Value
    }

    return [string]$Value
}

function Add-WarningMessage {
    param(
        [Parameter(Mandatory)]
        [System.Collections.Generic.List[string]]$Warnings,
        [Parameter(Mandatory)]
        [string]$Message
    )

    if (-not [string]::IsNullOrWhiteSpace($Message)) {
        $Warnings.Add($Message)
    }
}

function Release-ComReference {
    param(
        $Reference
    )

    if ($null -ne $Reference -and [System.Runtime.InteropServices.Marshal]::IsComObject($Reference)) {
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Reference)
    }
}

function New-ExcelApplication {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    return $excel
}

function Get-BorderNames {
    return [ordered]@{
        left = 7
        top = 8
        bottom = 9
        right = 10
        inside_vertical = 11
        inside_horizontal = 12
        diagonal_down = 5
        diagonal_up = 6
    }
}

function Get-CellHyperlink {
    param(
        $Cell
    )

    try {
        if ($Cell.Hyperlinks.Count -gt 0) {
            $link = $Cell.Hyperlinks.Item(1)
            return [ordered]@{
                address = if ([string]::IsNullOrWhiteSpace([string]$link.Address)) { $null } else { [string]$link.Address }
                sub_address = if ([string]::IsNullOrWhiteSpace([string]$link.SubAddress)) { $null } else { [string]$link.SubAddress }
                text_to_display = if ([string]::IsNullOrWhiteSpace([string]$link.TextToDisplay)) { $null } else { [string]$link.TextToDisplay }
            }
        }
    }
    catch {
        return $null
    }

    return $null
}

function Get-CellCommentText {
    param(
        $Cell
    )

    try {
        if ($null -ne $Cell.Comment) {
            return [string]$Cell.Comment.Text()
        }
    }
    catch {
        return $null
    }

    return $null
}

function Get-CellThreadedComment {
    param(
        $Cell
    )

    $commentThreaded = $null
    $repliesCollection = $null

    try {
        try {
            $commentThreaded = $Cell.CommentThreaded
        }
        catch {
            return $null
        }

        if ($null -eq $commentThreaded) {
            return $null
        }

        $text = $null
        try {
            $text = [string]$commentThreaded.Text()
        }
        catch {
            try {
                $text = [string]$commentThreaded.Text
            }
            catch {
                $text = $null
            }
        }

        $author = $null
        try {
            $author = [string]$commentThreaded.Author.Name
        }
        catch {
            try {
                $author = [string]$commentThreaded.Author
            }
            catch {
                $author = $null
            }
        }

        $createdAt = $null
        try {
            $createdAt = ([datetime]$commentThreaded.Date).ToString('o')
        }
        catch {
            $createdAt = $null
        }

        $replyList = [System.Collections.Generic.List[object]]::new()
        try {
            $repliesCollection = $commentThreaded.Replies
            if ($null -ne $repliesCollection) {
                foreach ($reply in $repliesCollection) {
                    $replyText = $null
                    $replyAuthor = $null
                    $replyDate = $null

                    try {
                        $replyText = [string]$reply.Text()
                    }
                    catch {
                        try {
                            $replyText = [string]$reply.Text
                        }
                        catch {
                            $replyText = $null
                        }
                    }

                    try {
                        $replyAuthor = [string]$reply.Author.Name
                    }
                    catch {
                        try {
                            $replyAuthor = [string]$reply.Author
                        }
                        catch {
                            $replyAuthor = $null
                        }
                    }

                    try {
                        $replyDate = ([datetime]$reply.Date).ToString('o')
                    }
                    catch {
                        $replyDate = $null
                    }

                    [void]$replyList.Add([ordered]@{
                        text = $replyText
                        author = $replyAuthor
                        created_at = $replyDate
                    })

                    Release-ComReference $reply
                }
            }
        }
        catch {
        }

        return [ordered]@{
            text = $text
            author = $author
            created_at = $createdAt
            replies = $replyList
        }
    }
    finally {
        if ($null -ne $repliesCollection) {
            Release-ComReference $repliesCollection
        }
        if ($null -ne $commentThreaded) {
            Release-ComReference $commentThreaded
        }
    }
}

function Get-CellFormula2 {
    param(
        $Cell
    )

    try {
        $formula2 = $Cell.Formula2
        if ($null -eq $formula2 -or [string]$formula2 -eq '') {
            $formula2 = $null
        }
    }
    catch {
        $formula2 = $null
    }

    if ($null -eq $formula2) {
        try {
            $fallbackFormula = $Cell.Formula
            if ($null -ne $fallbackFormula -and [string]$fallbackFormula -ne '') {
                return [string]$fallbackFormula
            }
        }
        catch {
        }

        return $null
    }

    return [string]$formula2
}

function Get-WorksheetFreezeState {
    param(
        [Parameter(Mandatory)]
        $Excel,
        [Parameter(Mandatory)]
        $Worksheet
    )

    $state = [ordered]@{
        enabled = $false
        split_row = 0
        split_column = 0
    }

    try {
        [void]$Worksheet.Activate()
        $window = $Excel.ActiveWindow
        if ($null -ne $window) {
            $state.enabled = [bool]$window.FreezePanes
            $state.split_row = [int]$window.SplitRow
            $state.split_column = [int]$window.SplitColumn
        }
    }
    catch {
        $state.enabled = $false
    }

    return $state
}

function Get-UsedRangeInfo {
    param(
        [Parameter(Mandatory)]
        $UsedRange
    )

    $firstRow = [int]$UsedRange.Row
    $firstColumn = [int]$UsedRange.Column
    $rowCount = [int]$UsedRange.Rows.Count
    $columnCount = [int]$UsedRange.Columns.Count

    return [ordered]@{
        address = [string]$UsedRange.Address($false, $false)
        first_row = $firstRow
        first_column = $firstColumn
        last_row = $firstRow + $rowCount - 1
        last_column = $firstColumn + $columnCount - 1
        row_count = $rowCount
        column_count = $columnCount
    }
}

function Get-StyleRecord {
    param(
        [Parameter(Mandatory)]
        $Cell
    )

    $record = [ordered]@{
        fill_color = $null
        font_color = $null
        horizontal_alignment = $null
        vertical_alignment = $null
        wrap_text = $null
        borders = [ordered]@{}
    }

    try {
        $record.fill_color = Convert-ExcelColor $Cell.Interior.Color
    }
    catch {
        $record.fill_color = $null
    }

    try {
        $record.font_color = Convert-ExcelColor $Cell.Font.Color
    }
    catch {
        $record.font_color = $null
    }

    try {
        $record.horizontal_alignment = [int]$Cell.HorizontalAlignment
    }
    catch {
        $record.horizontal_alignment = $null
    }

    try {
        $record.vertical_alignment = [int]$Cell.VerticalAlignment
    }
    catch {
        $record.vertical_alignment = $null
    }

    try {
        $record.wrap_text = [bool]$Cell.WrapText
    }
    catch {
        $record.wrap_text = $null
    }

    $borders = $null
    try {
        $borders = $Cell.Borders
        $overallLineStyle = $null
        try {
            $overallLineStyle = [int]$borders.LineStyle
        }
        catch {
            $overallLineStyle = $null
        }

        if ($null -eq $overallLineStyle -or $overallLineStyle -eq -4142) {
            foreach ($pair in (Get-BorderNames).GetEnumerator()) {
                $record.borders[$pair.Key] = $null
            }
        }
        else {
            foreach ($pair in (Get-BorderNames).GetEnumerator()) {
                $border = $null
                try {
                    $border = $borders.Item($pair.Value)
                    $lineStyle = $null
                    $weight = $null
                    $color = $null

                    try {
                        $lineStyle = [int]$border.LineStyle
                    }
                    catch {
                        $lineStyle = $null
                    }

                    try {
                        $weight = [int]$border.Weight
                    }
                    catch {
                        $weight = $null
                    }

                    try {
                        $color = Convert-ExcelColor $border.Color
                    }
                    catch {
                        $color = $null
                    }

                    $record.borders[$pair.Key] = [ordered]@{
                        line_style = $lineStyle
                        weight = $weight
                        color = $color
                    }
                }
                catch {
                    $record.borders[$pair.Key] = $null
                }
                finally {
                    if ($null -ne $border) {
                        Release-ComReference $border
                    }
                }
            }
        }
    }
    finally {
        if ($null -ne $borders) {
            Release-ComReference $borders
        }
    }

    return $record
}

function Convert-CoordinateToA1 {
    param(
        [Parameter(Mandatory)]
        [int]$Row,
        [Parameter(Mandatory)]
        [int]$Column
    )

    $col = $Column
    $letters = ''
    while ($col -gt 0) {
        $remainder = [int](($col - 1) % 26)
        $letters = [char][int](65 + $remainder) + $letters
        $col = [int][math]::Floor(($col - 1) / 26)
    }

    return '{0}{1}' -f $letters, $Row
}
