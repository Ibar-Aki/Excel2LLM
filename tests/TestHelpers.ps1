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
        [switch]$IncludeStyles
    )

    $excel = $null
    $workbook = $null

    try {
        $excel = New-ExcelApplication
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = 'Grid'

        for ($row = 1; $row -le 3; $row++) {
            for ($column = 1; $column -le 4; $column++) {
                $sheet.Cells.Item($row, $column).Value2 = "R$row-C$column"
            }
        }

        $sheet.Cells.Item(3, 4).Formula = '=A1&B1'
        $sheet.Range('A2:B2').Merge()
        $sheet.Range('A2').Value2 = 'Merged'

        if ($IncludeStyles) {
            $sheet.Range('A1').Interior.Color = 255
            $sheet.Range('A1').Font.Color = 16777215
            $sheet.Range('A1').WrapText = $true
            $sheet.Range('A1').Borders.Item(7).LineStyle = 1
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
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}
