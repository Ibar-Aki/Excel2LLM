[CmdletBinding()]
param(
    [string]$OutputDir,
    [ValidateSet('all', 'mechanical', 'accounting')]
    [string]$Scenario = 'all',
    [ValidateSet('original', 'improved', 'both')]
    [string]$Variant = 'both'
)

. (Join-Path $PSScriptRoot 'common.ps1')

if (-not $OutputDir) {
    $OutputDir = Join-Path (Join-Path (Get-ProjectRoot) 'samples') 'domain_scenarios'
}

Ensure-Directory -Path $OutputDir

function Save-WorkbookAsXlsx {
    param(
        [Parameter(Mandatory)]
        $Workbook,
        [Parameter(Mandatory)]
        [string]$Path
    )

    Ensure-Directory -Path (Split-Path -Path $Path -Parent)
    $Workbook.SaveAs($Path, 51)
}

function Set-SheetStyleHeader {
    param(
        [Parameter(Mandatory)]
        $Range
    )

    $Range.Font.Bold = $true
    $Range.Interior.Color = 15773696
    $Range.Borders.Item(7).LineStyle = 1
    $Range.Borders.Item(8).LineStyle = 1
    $Range.Borders.Item(9).LineStyle = 1
    $Range.Borders.Item(10).LineStyle = 1
}

function Set-RowValues {
    param(
        [Parameter(Mandatory)]
        $Worksheet,
        [Parameter(Mandatory)]
        [int]$Row,
        [Parameter(Mandatory)]
        [object[]]$Values
    )

    for ($index = 0; $index -lt $Values.Count; $index++) {
        Set-CellValue -Cell $Worksheet.Cells.Item($Row, $index + 1) -Value $Values[$index]
    }
}

function Set-TableValues {
    param(
        [Parameter(Mandatory)]
        $Worksheet,
        [Parameter(Mandatory)]
        [int]$StartRow,
        [Parameter(Mandatory)]
        [object[][]]$Rows
    )

    $currentRow = $StartRow
    foreach ($rowValues in $Rows) {
        Set-RowValues -Worksheet $Worksheet -Row $currentRow -Values $rowValues
        $currentRow++
    }
}

function Set-CellValue {
    param(
        [Parameter(Mandatory)]
        $Cell,
        $Value
    )

    if ($null -eq $Value) {
        $Cell.ClearContents()
        return
    }

    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal] -or $Value -is [single]) {
        $Cell.Value2 = [double]$Value
        return
    }

    if ($Value -is [bool]) {
        $Cell.Value2 = [bool]$Value
        return
    }

    $Cell.Value2 = [string]$Value
}

function New-MechanicalWorkbook {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [bool]$Improved
    )

    $excel = $null
    $workbook = $null
    $inputSheet = $null
    $calcSheet = $null
    $reviewSheet = $null
    $window = $null

    try {
        $excel = New-ExcelApplication
        $workbook = $excel.Workbooks.Add()

        $inputSheet = $workbook.Worksheets.Item(1)
        $inputSheet.Name = 'Inputs'
        $calcSheet = $workbook.Worksheets.Add()
        $calcSheet.Name = if ($Improved) { 'ShaftSizing' } else { 'Calc' }
        $reviewSheet = $workbook.Worksheets.Add()
        $reviewSheet.Name = if ($Improved) { 'Checks' } else { 'Review' }

        Set-RowValues -Worksheet $inputSheet -Row 1 -Values @('Parameter', 'Value', 'Unit', 'Note')
        Set-SheetStyleHeader -Range $inputSheet.Range('A1:D1')
        $inputSheet.Range('A2').Value2 = 'Torque_Nm'
        $inputSheet.Range('B2').Value2 = 420
        $inputSheet.Range('C2').Value2 = 'N*m'
        $inputSheet.Range('D2').Value2 = 'Input torque'
        $inputSheet.Range('A3').Value2 = 'AllowableShear_MPa'
        $inputSheet.Range('B3').Value2 = 85
        $inputSheet.Range('C3').Value2 = 'MPa'
        $inputSheet.Range('D3').Value2 = 'Material allowable shear stress'
        $inputSheet.Range('A4').Value2 = 'ServiceFactor'
        $serviceFactor = if ($Improved) { 1.5 } else { 1.6 }
        Set-CellValue -Cell $inputSheet.Range('B4') -Value $serviceFactor
        $inputSheet.Range('C4').Value2 = '-'
        $serviceFactorNote = if ($Improved) { 'Reviewed service factor' } else { 'Temporary assumption' }
        Set-CellValue -Cell $inputSheet.Range('D4') -Value $serviceFactorNote
        $inputSheet.Range('A5').Value2 = 'ExistingDiameter_mm'
        $inputSheet.Range('B5').Value2 = 40
        $inputSheet.Range('C5').Value2 = 'mm'
        $inputSheet.Range('D5').Value2 = 'Current shaft diameter'
        $inputSheet.Range('A6').Value2 = 'TargetSafetyRatio'
        $inputSheet.Range('B6').Value2 = 1.1
        $inputSheet.Range('C6').Value2 = '-'
        $inputSheet.Range('D6').Value2 = 'Pass threshold'
        $inputSheet.Columns.Item(1).ColumnWidth = 24
        $inputSheet.Columns.Item(4).ColumnWidth = 28

        if (-not $Improved) {
            Set-RowValues -Worksheet $calcSheet -Row 1 -Values @('Item', 'Value', 'Unit')
            Set-SheetStyleHeader -Range $calcSheet.Range('A1:C1')
            $calcSheet.Range('A2').Value2 = 'Torque_Nmm'
            $calcSheet.Range('B2').Formula = '=Inputs!B2*1000'
            $calcSheet.Range('C2').Value2 = 'Nmm'
            $calcSheet.Range('A3').Value2 = 'DesignTorque'
            $calcSheet.Range('B3').Formula = '=B2*Inputs!B4'
            $calcSheet.Range('C3').Value2 = 'Nmm'
            $calcSheet.Range('A4').Value2 = 'ReqDia'
            $calcSheet.Range('B4').Formula = '=((16*B3)/(PI()*Inputs!B3))^(1/3)'
            $calcSheet.Range('C4').Value2 = 'mm'
            $calcSheet.Range('A5').Value2 = 'ExistingDia'
            $calcSheet.Range('B5').Formula = '=Inputs!B5'
            $calcSheet.Range('C5').Value2 = 'mm'
            $calcSheet.Range('A6').Value2 = 'StressRatio'
            $calcSheet.Range('B6').Formula = '=B5/B4'
            $calcSheet.Range('C6').Value2 = '-'
            $calcSheet.Range('A7').Value2 = 'Judge'
            $calcSheet.Range('B7').Formula = '=IF(B6>=Inputs!B6,"PASS","FAIL")'
            $calcSheet.Range('A9').Value2 = 'Tmp'
            $calcSheet.Range('B9').Formula = '=((16*B3)/(PI()*Inputs!B3))^(1/3)'
            $calcSheet.Range('A10').Value2 = 'Memo'
            $calcSheet.Range('B10').Value2 = 'Need to explain why 1000 is used'
            $calcSheet.Range('B10').AddComment('LLM should recommend clearer unit conversion and remove duplicate formula.') | Out-Null
            $calcSheet.Range('A12:C12').Merge()
            $calcSheet.Range('A12').Value2 = 'Review before release'
            $calcSheet.Range('A12').Interior.Color = 65535

            Set-RowValues -Worksheet $reviewSheet -Row 1 -Values @('CheckItem', 'Status')
            Set-SheetStyleHeader -Range $reviewSheet.Range('A1:B1')
            $reviewSheet.Range('A2').Value2 = 'Duplicate formulas reviewed'
            $reviewSheet.Range('B2').Value2 = 'No'
            $reviewSheet.Range('A3').Value2 = 'Units explicit'
            $reviewSheet.Range('B3').Value2 = 'No'
        }
        else {
            Set-RowValues -Worksheet $calcSheet -Row 1 -Values @('Step', 'Value', 'Unit', 'Explanation')
            Set-SheetStyleHeader -Range $calcSheet.Range('A1:D1')
            $calcSheet.Range('A2').Value2 = 'Torque_Nmm'
            $calcSheet.Range('B2').Formula = '=Inputs!B2*1000'
            $calcSheet.Range('C2').Value2 = 'Nmm'
            $calcSheet.Range('D2').Value2 = 'Convert N*m to Nmm'
            $calcSheet.Range('A3').Value2 = 'DesignTorque_Nmm'
            $calcSheet.Range('B3').Formula = '=B2*Inputs!B4'
            $calcSheet.Range('C3').Value2 = 'Nmm'
            $calcSheet.Range('D3').Value2 = 'Apply service factor'
            $calcSheet.Range('A4').Value2 = 'RequiredDiameter_mm'
            $calcSheet.Range('B4').Formula = '=ROUNDUP(((16*B3)/(PI()*Inputs!B3))^(1/3),1)'
            $calcSheet.Range('C4').Value2 = 'mm'
            $calcSheet.Range('D4').Value2 = 'Solid shaft formula'
            $calcSheet.Range('A5').Value2 = 'ExistingDiameter_mm'
            $calcSheet.Range('B5').Formula = '=Inputs!B5'
            $calcSheet.Range('C5').Value2 = 'mm'
            $calcSheet.Range('D5').Value2 = 'Current design value'
            $calcSheet.Range('A6').Value2 = 'SafetyRatio'
            $calcSheet.Range('B6').Formula = '=ROUND(B5/B4,3)'
            $calcSheet.Range('C6').Value2 = '-'
            $calcSheet.Range('D6').Value2 = 'Existing / required diameter'
            $calcSheet.Range('A7').Value2 = 'Margin_mm'
            $calcSheet.Range('B7').Formula = '=ROUND(B5-B4,2)'
            $calcSheet.Range('C7').Value2 = 'mm'
            $calcSheet.Range('D7').Value2 = 'Positive value means enough diameter'
            $calcSheet.Range('A8').Value2 = 'Recommendation'
            $calcSheet.Range('B8').Formula = '=IF(B6>=Inputs!B6,"Current diameter acceptable","Increase diameter or material strength")'
            $calcSheet.Range('B8').AddComment('Improved workbook exposes recommendation directly for the user.') | Out-Null

            Set-RowValues -Worksheet $reviewSheet -Row 1 -Values @('CheckItem', 'Result', 'Comment')
            Set-SheetStyleHeader -Range $reviewSheet.Range('A1:C1')
            $reviewSheet.Range('A2').Value2 = 'Safety ratio meets target'
            $reviewSheet.Range('B2').Formula = '=IF(ShaftSizing!B6>=Inputs!B6,"PASS","FAIL")'
            $reviewSheet.Range('C2').Value2 = 'Compares calculated safety ratio against target'
            $reviewSheet.Range('A3').Value2 = 'Unit conversion explicit'
            $reviewSheet.Range('B3').Value2 = 'PASS'
            $reviewSheet.Range('C3').Value2 = 'Torque conversion shown in ShaftSizing!B2'
            $reviewSheet.Range('A4').Value2 = 'Duplicate formulas removed'
            $reviewSheet.Range('B4').Value2 = 'PASS'
            $reviewSheet.Range('C4').Value2 = 'One source of truth for required diameter'
        }

        [void]$calcSheet.Activate()
        $window = $excel.ActiveWindow
        $window.SplitRow = 1
        $window.FreezePanes = $true

        Save-WorkbookAsXlsx -Workbook $workbook -Path $Path
        Write-Host "Created workbook -> $Path"
    }
    finally {
        if ($null -ne $window) { Release-ComReference $window }
        if ($null -ne $reviewSheet) { Release-ComReference $reviewSheet }
        if ($null -ne $calcSheet) { Release-ComReference $calcSheet }
        if ($null -ne $inputSheet) { Release-ComReference $inputSheet }
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch {}
            Release-ComReference $workbook
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch {}
            Release-ComReference $excel
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function New-AccountingWorkbook {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [bool]$Improved
    )

    $excel = $null
    $workbook = $null
    $transactionsSheet = $null
    $budgetSheet = $null
    $summarySheet = $null
    $checksSheet = $null
    $window = $null

    try {
        $excel = New-ExcelApplication
        $workbook = $excel.Workbooks.Add()

        $transactionsSheet = $workbook.Worksheets.Item(1)
        $transactionsSheet.Name = 'Transactions'
        $budgetSheet = $workbook.Worksheets.Add()
        $budgetSheet.Name = 'Budget'
        $summarySheet = $workbook.Worksheets.Add()
        $summarySheet.Name = 'Summary'
        $checksSheet = $workbook.Worksheets.Add()
        $checksSheet.Name = if ($Improved) { 'Checks' } else { 'Notes' }

        Set-RowValues -Worksheet $transactionsSheet -Row 1 -Values @('Date', 'Department', 'Type', 'Category', 'Amount', 'Note')
        Set-SheetStyleHeader -Range $transactionsSheet.Range('A1:F1')
        $transactions = @(
            @('2026-01-05', 'Design', 'Revenue', 'Project', 180000, 'Prototype order'),
            @('2026-01-08', 'Design', 'Cost', 'Material', 65000, 'Bearing and shaft stock'),
            @('2026-01-12', 'Assembly', 'Revenue', 'Project', 220000, 'Line support'),
            @('2026-01-14', 'Assembly', 'Cost', 'Labor', 90000, 'Overtime work'),
            @('2026-01-20', 'Design', 'Cost', 'Outsource', 25000, 'CAE review'),
            @('2026-01-21', 'Sales', 'Revenue', 'Service', 120000, 'Field engineering'),
            @('2026-01-24', 'Sales', 'Cost', 'Travel', 18000, 'Site visit'),
            @('2026-01-26', 'Assembly', 'Cost', 'Material', 30000, 'Fasteners'),
            @('2026-01-28', 'Sales', 'Cost', 'Commission', 25000, 'Channel fee')
        )

        $transactionRow = 2
        foreach ($row in $transactions) {
            for ($column = 0; $column -lt $row.Count; $column++) {
                Set-CellValue -Cell $transactionsSheet.Cells.Item($transactionRow, $column + 1) -Value $row[$column]
            }
            $transactionRow++
        }
        $transactionsSheet.Columns.Item(1).ColumnWidth = 14
        $transactionsSheet.Columns.Item(6).ColumnWidth = 24

        Set-RowValues -Worksheet $budgetSheet -Row 1 -Values @('Department', 'RevenueBudget', 'CostBudget')
        Set-SheetStyleHeader -Range $budgetSheet.Range('A1:C1')
        Set-TableValues -Worksheet $budgetSheet -StartRow 2 -Rows @(
            @('Design', 210000, 95000),
            @('Assembly', 240000, 110000),
            @('Sales', 130000, 55000)
        )

        if (-not $Improved) {
            Set-RowValues -Worksheet $summarySheet -Row 1 -Values @('Dept', 'Revenue', 'Cost', 'Profit', 'BudgetRev', 'BudgetCost', 'Memo')
            Set-SheetStyleHeader -Range $summarySheet.Range('A1:G1')
            Set-TableValues -Worksheet $summarySheet -StartRow 2 -Rows @(
                @('Design'),
                @('Assembly'),
                @('Sales')
            )
            for ($row = 2; $row -le 4; $row++) {
                $summarySheet.Cells.Item($row, 2).Formula = '=SUMIFS(Transactions!$E:$E,Transactions!$B:$B,$A' + $row + ',Transactions!$C:$C,"Revenue")'
                $summarySheet.Cells.Item($row, 3).Formula = '=SUMIFS(Transactions!$E:$E,Transactions!$B:$B,$A' + $row + ',Transactions!$C:$C,"Cost")'
                $summarySheet.Cells.Item($row, 4).Formula = '=B' + $row + '-C' + $row
                $summarySheet.Cells.Item($row, 5).Formula = '=VLOOKUP($A' + $row + ',Budget!$A$2:$C$4,2,FALSE)'
                $summarySheet.Cells.Item($row, 6).Formula = '=VLOOKUP($A' + $row + ',Budget!$A$2:$C$4,3,FALSE)'
                $summarySheet.Cells.Item($row, 7).Value2 = 'Need variance view'
            }
            $summarySheet.Range('A6').Value2 = 'Total'
            $summarySheet.Range('B6').Formula = '=SUM(B2:B4)'
            $summarySheet.Range('C6').Formula = '=SUM(C2:C4)'
            $summarySheet.Range('D6').Formula = '=B6-C6'
            $summarySheet.Range('F8:G8').Merge()
            $summarySheet.Range('F8').Value2 = 'Review summary formulas'
            $summarySheet.Range('F8').Interior.Color = 65535

            Set-RowValues -Worksheet $checksSheet -Row 1 -Values @('Item', 'Status')
            Set-SheetStyleHeader -Range $checksSheet.Range('A1:B1')
            $checksSheet.Range('A2').Value2 = 'Budget variance checked'
            $checksSheet.Range('B2').Value2 = 'No'
            $checksSheet.Range('A3').Value2 = 'Negative values checked'
            $checksSheet.Range('B3').Value2 = 'No'
        }
        else {
            Set-RowValues -Worksheet $summarySheet -Row 1 -Values @('Department', 'Revenue', 'Cost', 'Profit', 'RevenueBudget', 'CostBudget', 'RevenueVariance', 'CostVariance', 'ProfitMargin')
            Set-SheetStyleHeader -Range $summarySheet.Range('A1:I1')
            Set-TableValues -Worksheet $summarySheet -StartRow 2 -Rows @(
                @('Design'),
                @('Assembly'),
                @('Sales')
            )
            for ($row = 2; $row -le 4; $row++) {
                $summarySheet.Cells.Item($row, 2).Formula = '=SUMIFS(Transactions!$E:$E,Transactions!$B:$B,$A' + $row + ',Transactions!$C:$C,"Revenue")'
                $summarySheet.Cells.Item($row, 3).Formula = '=SUMIFS(Transactions!$E:$E,Transactions!$B:$B,$A' + $row + ',Transactions!$C:$C,"Cost")'
                $summarySheet.Cells.Item($row, 4).Formula = '=B' + $row + '-C' + $row
                $summarySheet.Cells.Item($row, 5).Formula = '=VLOOKUP($A' + $row + ',Budget!$A$2:$C$4,2,FALSE)'
                $summarySheet.Cells.Item($row, 6).Formula = '=VLOOKUP($A' + $row + ',Budget!$A$2:$C$4,3,FALSE)'
                $summarySheet.Cells.Item($row, 7).Formula = '=B' + $row + '-E' + $row
                $summarySheet.Cells.Item($row, 8).Formula = '=C' + $row + '-F' + $row
                $summarySheet.Cells.Item($row, 9).Formula = '=IF(B' + $row + '=0,0,D' + $row + '/B' + $row + ')'
            }
            $summarySheet.Range('A6').Value2 = 'Total'
            $summarySheet.Range('B6').Formula = '=SUM(B2:B4)'
            $summarySheet.Range('C6').Formula = '=SUM(C2:C4)'
            $summarySheet.Range('D6').Formula = '=SUM(D2:D4)'
            $summarySheet.Range('E6').Formula = '=SUM(E2:E4)'
            $summarySheet.Range('F6').Formula = '=SUM(F2:F4)'
            $summarySheet.Range('G6').Formula = '=SUM(G2:G4)'
            $summarySheet.Range('H6').Formula = '=SUM(H2:H4)'
            $summarySheet.Range('I6').Formula = '=IF(B6=0,0,D6/B6)'
            $summarySheet.Range('I2:I6').NumberFormat = '0.0%'

            Set-RowValues -Worksheet $checksSheet -Row 1 -Values @('CheckItem', 'Result', 'Comment')
            Set-SheetStyleHeader -Range $checksSheet.Range('A1:C1')
            $checksSheet.Range('A2').Value2 = 'Revenue variance visible'
            $checksSheet.Range('B2').Value2 = 'PASS'
            $checksSheet.Range('C2').Value2 = 'Summary sheet includes RevenueVariance'
            $checksSheet.Range('A3').Value2 = 'Cost variance visible'
            $checksSheet.Range('B3').Value2 = 'PASS'
            $checksSheet.Range('C3').Value2 = 'Summary sheet includes CostVariance'
            $checksSheet.Range('A4').Value2 = 'Profit margin visible'
            $checksSheet.Range('B4').Value2 = 'PASS'
            $checksSheet.Range('C4').Value2 = 'Summary sheet includes ProfitMargin'
        }

        [void]$summarySheet.Activate()
        $window = $excel.ActiveWindow
        $window.SplitRow = 1
        $window.FreezePanes = $true

        Save-WorkbookAsXlsx -Workbook $workbook -Path $Path
        Write-Host "Created workbook -> $Path"
    }
    finally {
        if ($null -ne $window) { Release-ComReference $window }
        if ($null -ne $checksSheet) { Release-ComReference $checksSheet }
        if ($null -ne $summarySheet) { Release-ComReference $summarySheet }
        if ($null -ne $budgetSheet) { Release-ComReference $budgetSheet }
        if ($null -ne $transactionsSheet) { Release-ComReference $transactionsSheet }
        if ($null -ne $workbook) {
            try { $workbook.Close($false) } catch {}
            Release-ComReference $workbook
        }
        if ($null -ne $excel) {
            try { $excel.Quit() } catch {}
            Release-ComReference $excel
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Get-VariantsToGenerate {
    param(
        [Parameter(Mandatory)]
        [string]$Variant
    )

    if ($Variant -eq 'both') {
        return @('original', 'improved')
    }

    return @($Variant)
}

$scenarios = if ($Scenario -eq 'all') { @('mechanical', 'accounting') } else { @($Scenario) }
$variants = Get-VariantsToGenerate -Variant $Variant

foreach ($scenarioName in $scenarios) {
    foreach ($variantName in $variants) {
        $isImproved = $variantName -eq 'improved'
        $fileName = '{0}_{1}.xlsx' -f $scenarioName, $variantName
        $targetPath = Join-Path $OutputDir $fileName

        if ($scenarioName -eq 'mechanical') {
            New-MechanicalWorkbook -Path $targetPath -Improved:$isImproved
            continue
        }

        if ($scenarioName -eq 'accounting') {
            New-AccountingWorkbook -Path $targetPath -Improved:$isImproved
        }
    }
}
