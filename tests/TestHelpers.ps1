. (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'scripts\common.ps1')

function New-TestWorkspace {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $root = Join-Path ([System.IO.Path]::GetTempPath()) ('Excel2LLM.Tests.{0}.{1}' -f $Name, [guid]::NewGuid().ToString('N'))
    Ensure-Directory -Path $root
    return $root
}

function New-MiniWorkbook {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [switch]$IncludeStyles,
        [switch]$IncludeMetadata
    )

    $excel = $null
    $workbook = $null

    try {
        $excel = New-ExcelApplication
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = 'Grid'
        $lookupSheet = $null

        for ($row = 1; $row -le 3; $row++) {
            for ($column = 1; $column -le 4; $column++) {
                $sheet.Cells.Item($row, $column).Value2 = "R$row-C$column"
            }
        }

        $sheet.Cells.Item(1, 3).Value2 = 10
        $sheet.Cells.Item(2, 3).Value2 = 20
        $sheet.Cells.Item(3, 3).Value2 = 30
        $sheet.Cells.Item(3, 4).Formula = '=A1&B1'
        $sheet.Range('A2:B2').Merge()
        $sheet.Range('A2').Value2 = 'Merged'

        if ($IncludeStyles) {
            $sheet.Range('A1').Interior.Color = 255
            $sheet.Range('A1').Font.Color = 16777215
            $sheet.Range('A1').WrapText = $true
            $sheet.Range('A1').Borders.Item(7).LineStyle = 1
        }

        if ($IncludeMetadata) {
            $lookupSheet = $workbook.Worksheets.Add()
            $lookupSheet.Name = 'Lookup'
            $lookupSheet.Cells.Item(1, 1).Value2 = 'LookupValue'

            $workbook.Names.Add('WorkbookInput', "='Grid'!`$B`$1") | Out-Null
            $workbook.Names.Add('ThresholdConstant', '=15') | Out-Null
            $workbook.Names.Add('LookupOnlyName', "='Lookup'!`$A`$1") | Out-Null

            $validationRange = $null
            $validation = $null
            $conditionalRange = $null
            $formatConditions = $null
            $condition = $null
            try {
                $validationRange = $sheet.Range('B2:B3')
                $validation = $validationRange.Validation
                $validation.Delete()
                $validation.Add(3, 1, 1, 'Yes,No') | Out-Null
                $validation.IgnoreBlank = $true
                $validation.InCellDropdown = $true
                $validation.InputTitle = '入力候補'
                $validation.InputMessage = 'Yes または No を入力してください'
                $validation.ErrorTitle = '入力エラー'
                $validation.ErrorMessage = '候補以外は入力できません'

                $conditionalRange = $sheet.Range('C1:C3')
                $formatConditions = $conditionalRange.FormatConditions
                $condition = $formatConditions.Add(1, 5, '15')
                $condition.StopIfTrue = $true
                $condition.Interior.Color = 65535
            }
            finally {
                if ($null -ne $condition) {
                    Release-ComReference $condition
                }
                if ($null -ne $formatConditions) {
                    Release-ComReference $formatConditions
                }
                if ($null -ne $conditionalRange) {
                    Release-ComReference $conditionalRange
                }
                if ($null -ne $validation) {
                    Release-ComReference $validation
                }
                if ($null -ne $validationRange) {
                    Release-ComReference $validationRange
                }
            }
        }

        $directory = Split-Path -Path $Path -Parent
        Ensure-Directory -Path $directory
        $workbook.SaveAs($Path, 51)
    }
    finally {
        if ($null -ne $workbook) {
            try {
                $workbook.Close($false)
            }
            catch {
            }
            Release-ComReference $workbook
        }
        if ($null -ne $excel) {
            try {
                $excel.Quit()
            }
            catch {
            }
            Release-ComReference $excel
        }
        if ($null -ne $lookupSheet) {
            Release-ComReference $lookupSheet
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Set-MetadataSqrefVariants {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [string]$ValidationSqref = 'B:B',
        [string]$ConditionalFormattingSqref = '3:3'
    )

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
        try {
            $worksheetEntries = @($archive.Entries | Where-Object { $_.FullName -like 'xl/worksheets/*.xml' } | Select-Object -ExpandProperty FullName)
        }
        finally {
            $archive.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }

    foreach ($sheetXmlPath in $worksheetEntries) {
        $sheetXml = Get-ZipEntryText -Path $Path -EntryPath $sheetXmlPath
        $updatedXml = $sheetXml.Replace('sqref="B2:B3"', ('sqref="' + $ValidationSqref + '"')).Replace('sqref="C1:C3"', ('sqref="' + $ConditionalFormattingSqref + '"'))
        if ($updatedXml -ne $sheetXml) {
            Set-ZipEntryText -Path $Path -EntryPath $sheetXmlPath -Content $updatedXml
        }
    }
}

function Invoke-ZipArchiveUpdate {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [scriptblock]$Action
    )

    $fileStream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try {
        $archive = [System.IO.Compression.ZipArchive]::new($fileStream, [System.IO.Compression.ZipArchiveMode]::Update, $false)
        try {
            & $Action $archive
        }
        finally {
            $archive.Dispose()
        }
    }
    finally {
        $fileStream.Dispose()
    }
}

function Get-ZipEntryText {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$EntryPath
    )

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $archive = [System.IO.Compression.ZipArchive]::new($stream, [System.IO.Compression.ZipArchiveMode]::Read, $false)
        try {
            $entry = $archive.GetEntry($EntryPath)
            if ($null -eq $entry) {
                throw "Zip entry was not found: $EntryPath"
            }

            $reader = [System.IO.StreamReader]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
            try {
                return $reader.ReadToEnd()
            }
            finally {
                $reader.Dispose()
            }
        }
        finally {
            $archive.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Set-ZipEntryText {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$EntryPath,
        [Parameter(Mandatory)]
        [string]$Content
    )

    Invoke-ZipArchiveUpdate -Path $Path -Action {
        param($Archive)

        $existing = $Archive.GetEntry($EntryPath)
        if ($null -ne $existing) {
            $existing.Delete()
        }

        $entry = $Archive.CreateEntry($EntryPath)
        $writer = [System.IO.StreamWriter]::new($entry.Open(), [System.Text.UTF8Encoding]::new($false))
        try {
            $writer.Write($Content)
        }
        finally {
            $writer.Dispose()
        }
    }
}

function Remove-ZipEntry {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$EntryPath
    )

    Invoke-ZipArchiveUpdate -Path $Path -Action {
        param($Archive)

        $entry = $Archive.GetEntry($EntryPath)
        if ($null -ne $entry) {
            $entry.Delete()
        }
    }
}

function Add-FilePaddingBytes {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [long]$TargetLengthBytes
    )

    $currentLength = (Get-Item -LiteralPath $Path).Length
    if ($currentLength -ge $TargetLengthBytes) {
        return
    }

    $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
        [void]$stream.Seek(0, [System.IO.SeekOrigin]::End)
        $buffer = New-Object byte[] (1024 * 1024)
        $remaining = $TargetLengthBytes - $currentLength
        while ($remaining -gt 0) {
            $writeLength = [int][Math]::Min($buffer.Length, $remaining)
            $stream.Write($buffer, 0, $writeLength)
            $remaining -= $writeLength
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Add-ZipPaddingEntry {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [long]$TargetLengthBytes
    )

    $currentLength = (Get-Item -LiteralPath $Path).Length
    if ($currentLength -ge $TargetLengthBytes) {
        return
    }

    $payloadLength = $TargetLengthBytes - $currentLength
    Invoke-ZipArchiveUpdate -Path $Path -Action {
        param($Archive)

        $existing = $Archive.GetEntry('xl/media/preflight-padding.bin')
        if ($null -ne $existing) {
            $existing.Delete()
        }

        $entry = $Archive.CreateEntry('xl/media/preflight-padding.bin', [System.IO.Compression.CompressionLevel]::NoCompression)
        $stream = $entry.Open()
        try {
            $buffer = New-Object byte[] (1024 * 1024)
            $remaining = $payloadLength
            while ($remaining -gt 0) {
                $writeLength = [int][Math]::Min($buffer.Length, $remaining)
                $stream.Write($buffer, 0, $writeLength)
                $remaining -= $writeLength
            }
        }
        finally {
            $stream.Dispose()
        }
    }
}

function New-PreflightWorkbookFixture {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [string]$Dimension,
        [switch]$RemoveWorkbookRelationships,
        [switch]$RemoveWorksheetDimension,
        [switch]$BreakWorksheetXml,
        [long]$PadToBytes
    )

    New-MiniWorkbook -Path $Path

    if ($RemoveWorkbookRelationships) {
        Remove-ZipEntry -Path $Path -EntryPath 'xl/_rels/workbook.xml.rels'
    }
    else {
        $sheetXmlPath = 'xl/worksheets/sheet1.xml'
        if ($BreakWorksheetXml) {
            Set-ZipEntryText -Path $Path -EntryPath $sheetXmlPath -Content '<worksheet>'
        }
        else {
            $sheetXml = Get-ZipEntryText -Path $Path -EntryPath $sheetXmlPath
            if ($RemoveWorksheetDimension) {
                $sheetXml = [System.Text.RegularExpressions.Regex]::Replace(
                    $sheetXml,
                    '<dimension\b[^>]*/>',
                    '',
                    [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
                )
            }
            elseif (-not [string]::IsNullOrWhiteSpace($Dimension)) {
                if ($sheetXml -match '<dimension\b[^>]*ref="[^"]+"[^>]*/>') {
                    $sheetXml = [System.Text.RegularExpressions.Regex]::Replace(
                        $sheetXml,
                        '(<dimension\b[^>]*ref=")[^"]+(")',
                        ('$1{0}$2' -f [System.Text.RegularExpressions.Regex]::Escape($Dimension)).Replace('\', ''),
                        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
                    )
                }
            }

            Set-ZipEntryText -Path $Path -EntryPath $sheetXmlPath -Content $sheetXml
        }
    }

    if ($PadToBytes -gt 0) {
        Add-ZipPaddingEntry -Path $Path -TargetLengthBytes $PadToBytes
    }
}

function New-CorruptWorkbookFile {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    $directory = Split-Path -Path $Path -Parent
    Ensure-Directory -Path $directory
    [System.IO.File]::WriteAllText($Path, 'this is not a valid xlsx archive', [System.Text.Encoding]::UTF8)
}
