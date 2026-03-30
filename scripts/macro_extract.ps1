[CmdletBinding()]
param(
    [Parameter(Mandatory, Position = 0)]
    [Alias('ExcelPath')]
    [string]$WorkbookPath,
    [string]$OutputDir,
    [switch]$RedactPaths
)

. (Join-Path $PSScriptRoot 'common.ps1')

function Get-VbaComponentMetadata {
    param(
        [Parameter(Mandatory)]
        $Component
    )

    $componentType = [int]$Component.Type
    switch ($componentType) {
        1 {
            return [ordered]@{
                component_type = 'standard_module'
                component_scope = 'standard'
                extension = '.bas'
            }
        }
        2 {
            return [ordered]@{
                component_type = 'class_module'
                component_scope = 'class'
                extension = '.cls'
            }
        }
        3 {
            return [ordered]@{
                component_type = 'user_form'
                component_scope = 'form'
                extension = '.frm'
            }
        }
        100 {
            return [ordered]@{
                component_type = 'document_module'
                component_scope = 'document'
                extension = '.cls'
            }
        }
        default {
            return [ordered]@{
                component_type = 'unknown'
                component_scope = 'unknown'
                extension = '.txt'
            }
        }
    }
}

function Get-CodeModuleText {
    param(
        $Component
    )

    $codeModule = $null
    try {
        $codeModule = $Component.CodeModule
        if ($null -eq $codeModule) {
            return [ordered]@{
                text = ''
                line_count = 0
            }
        }

        $lineCount = [int]$codeModule.CountOfLines
        if ($lineCount -le 0) {
            return [ordered]@{
                text = ''
                line_count = 0
            }
        }

        return [ordered]@{
            text = [string]$codeModule.Lines(1, $lineCount)
            line_count = $lineCount
        }
    }
    finally {
        if ($null -ne $codeModule) {
            Release-ComReference $codeModule
        }
    }
}

function Write-VbaPromptFile {
    param(
        [Parameter(Mandatory)]
        [string]$WorkbookName,
        [Parameter(Mandatory)]
        [string]$PromptPath,
        [string]$JsonlRelativePath,
        [string]$ManifestRelativePath,
        [string]$AccessError,
        [bool]$HasReadableSource
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    [void]$lines.Add('あなたは VBA レビュー支援アシスタントです。')
    [void]$lines.Add('')
    [void]$lines.Add(('対象ブック: {0}' -f $WorkbookName))
    [void]$lines.Add('目的:')
    [void]$lines.Add('- VBA の役割を要約する')
    [void]$lines.Add('- 主要な処理フローと依存関係を整理する')
    [void]$lines.Add('- 不具合リスク、保守性、改善案を示す')
    [void]$lines.Add('')

    if ($HasReadableSource) {
        [void]$lines.Add('見るべき入力:')
        [void]$lines.Add(('- VBA ソース JSONL: {0}' -f $JsonlRelativePath))
        [void]$lines.Add(('- 抽出結果メタ情報: {0}' -f $ManifestRelativePath))
        [void]$lines.Add('')
        [void]$lines.Add('出力形式:')
        [void]$lines.Add('- 概要')
        [void]$lines.Add('- コンポーネント別の役割')
        [void]$lines.Add('- 問題点')
        [void]$lines.Add('- 改善案')
        [void]$lines.Add('')
        [void]$lines.Add('制約:')
        [void]$lines.Add('- 根拠のない推測はしない')
        [void]$lines.Add('- コードに書かれていない仕様は「不明」と明記する')
    }
    else {
        [void]$lines.Add('現状:')
        [void]$lines.Add('- readable な VBA ソースは抽出できませんでした')
        [void]$lines.Add(('- 抽出結果メタ情報: {0}' -f $ManifestRelativePath))
        if (-not [string]::IsNullOrWhiteSpace($AccessError)) {
            [void]$lines.Add(('- 原因メモ: {0}' -f $AccessError))
        }
        [void]$lines.Add('')
        [void]$lines.Add('対応候補:')
        [void]$lines.Add('- Excel の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を確認する')
        [void]$lines.Add('- 再度 VBA 抽出を実行する')
    }

    [System.IO.File]::WriteAllLines($PromptPath, $lines, [System.Text.Encoding]::UTF8)
}

try {
    $resolvedWorkbookPath = Resolve-AbsolutePath -Path $WorkbookPath
    $extension = [System.IO.Path]::GetExtension($resolvedWorkbookPath).ToLowerInvariant()
    if (@('.xlsm', '.xlam') -notcontains $extension) {
        throw 'VBA 抽出は .xlsm または .xlam にのみ対応しています。'
    }

    if (-not $OutputDir) {
        $OutputDir = Get-DefaultRunOutputDirectory -ExcelPath $resolvedWorkbookPath
    }

    $resolvedOutputDir = Get-NormalizedFullPath -Path $OutputDir
    Ensure-Directory -Path $resolvedOutputDir

    $vbaOutputDir = Join-Path $resolvedOutputDir 'vba'
    $modulesDir = Join-Path $vbaOutputDir 'modules'
    $rawDir = Join-Path $vbaOutputDir 'raw'
    Ensure-Directory -Path $vbaOutputDir
    Ensure-Directory -Path $modulesDir
    Ensure-Directory -Path $rawDir

    $workbookName = [System.IO.Path]::GetFileName($resolvedWorkbookPath)
    $displayWorkbookPath = if ($RedactPaths) { $workbookName } else { $resolvedWorkbookPath }
    $displayOutputDir = if ($RedactPaths) { [System.IO.Path]::GetFileName($resolvedOutputDir) } else { $resolvedOutputDir }
    $manifestPath = Join-Path $vbaOutputDir 'macro_manifest.json'
    $jsonlPath = Join-Path $vbaOutputDir 'vba_llm_package.jsonl'
    $promptPath = Join-Path $vbaOutputDir 'vba_prompt.txt'
    $warnings = [System.Collections.Generic.List[string]]::new()
    $componentRecords = [System.Collections.Generic.List[object]]::new()
    $llmRecords = [System.Collections.Generic.List[object]]::new()
    $rawRelativePath = $null
    $accessError = $null
    $rawExportStatus = 'missing'
    $sourceExportStatus = 'no_vba'
    $status = 'success'
    $hasVba = $false

    $archive = $null
    try {
        $archive = [System.IO.Compression.ZipFile]::OpenRead($resolvedWorkbookPath)
        $rawPath = Join-Path $rawDir 'vbaProject.bin'
        if (Export-ZipArchiveEntryToFile -Archive $archive -EntryPath 'xl/vbaProject.bin' -DestinationPath $rawPath) {
            $rawExportStatus = 'generated'
            $rawRelativePath = Get-RelativePathFromDirectory -Path $rawPath -BaseDirectory $vbaOutputDir
            $hasVba = $true
        }
    }
    catch {
        $rawExportStatus = 'failed'
        $status = 'warning'
        Add-WarningMessage -Warnings $warnings -Message ("raw VBA の保存に失敗しました: {0}" -f $_.Exception.Message)
    }
    finally {
        if ($null -ne $archive) {
            $archive.Dispose()
        }
    }

    $excel = $null
    $workbook = $null
    $vbProject = $null
    $vbComponents = $null
    try {
        $excel = New-ExcelApplication
        $workbook = $excel.Workbooks.Open($resolvedWorkbookPath, 0, $true)
        $vbProject = $workbook.VBProject
        if ($null -eq $vbProject) {
            throw 'VBA プロジェクトにアクセスできません。'
        }

        $vbComponents = $vbProject.VBComponents
        foreach ($component in $vbComponents) {
            $componentName = [string]$component.Name
            $componentTypeCode = [int]$component.Type
            $typeInfo = Get-VbaComponentMetadata -Component $component
            $componentText = Get-CodeModuleText -Component $component
            $hasMeaningfulCode = [int]$componentText.line_count -gt 0

            $componentPath = Join-Path $modulesDir ($componentName + $typeInfo.extension)
            $componentRelativePath = $null
            $additionalFiles = [System.Collections.Generic.List[string]]::new()
            $exportStatus = 'generated'
            $codeText = [string]$componentText.text

            try {
                if ([int]$component.Type -eq 3) {
                    try {
                        $component.Export($componentPath)
                        $componentRelativePath = Get-RelativePathFromDirectory -Path $componentPath -BaseDirectory $vbaOutputDir
                        if (Test-Path -LiteralPath $componentPath) {
                            $hasMeaningfulCode = $true
                        }
                        $frxPath = [System.IO.Path]::ChangeExtension($componentPath, '.frx')
                        if (Test-Path -LiteralPath $frxPath) {
                            [void]$additionalFiles.Add((Get-RelativePathFromDirectory -Path $frxPath -BaseDirectory $vbaOutputDir))
                            $hasMeaningfulCode = $true
                        }
                    }
                    catch {
                        $exportStatus = 'fallback'
                        if (-not [string]::IsNullOrWhiteSpace($codeText)) {
                            [System.IO.File]::WriteAllText($componentPath, $codeText, [System.Text.Encoding]::UTF8)
                            $componentRelativePath = Get-RelativePathFromDirectory -Path $componentPath -BaseDirectory $vbaOutputDir
                            $hasMeaningfulCode = $true
                            Add-WarningMessage -Warnings $warnings -Message ("UserForm {0} は export に失敗したためコードのみ保存しました。" -f $componentName)
                        }
                        else {
                            $exportStatus = 'skipped'
                            Add-WarningMessage -Warnings $warnings -Message ("UserForm {0} の export に失敗し、コードも取得できませんでした。" -f $componentName)
                        }
                    }
                }
                elseif (-not [string]::IsNullOrWhiteSpace($codeText)) {
                    [System.IO.File]::WriteAllText($componentPath, $codeText, [System.Text.Encoding]::UTF8)
                    $componentRelativePath = Get-RelativePathFromDirectory -Path $componentPath -BaseDirectory $vbaOutputDir
                    $hasMeaningfulCode = $true
                }
                else {
                    $exportStatus = 'skipped'
                }
            }
            finally {
                Release-ComReference $component
            }

            if ($hasMeaningfulCode) {
                $hasVba = $true
            }

            $componentRecord = [ordered]@{
                component_name = $componentName
                component_type = $typeInfo.component_type
                component_scope = $typeInfo.component_scope
                type_code = $componentTypeCode
                source_path = $componentRelativePath
                additional_files = @($additionalFiles.ToArray())
                line_count = [int]$componentText.line_count
                export_status = $exportStatus
            }
            [void]$componentRecords.Add($componentRecord)

            if ($hasMeaningfulCode -and -not [string]::IsNullOrWhiteSpace($componentRelativePath) -and -not [string]::IsNullOrWhiteSpace($codeText)) {
                [void]$llmRecords.Add([ordered]@{
                    component_name = $componentName
                    component_type = $typeInfo.component_type
                    component_scope = $typeInfo.component_scope
                    source_path = $componentRelativePath
                    code_text = $codeText
                    line_count = [int]$componentText.line_count
                    workbook_name = $workbookName
                })
            }
        }

        if ($hasVba) {
            $sourceExportStatus = 'generated'
        }
    }
    catch {
        $sourceExportStatus = 'failed'
        $status = 'warning'
        $accessError = $_.Exception.Message
        Add-WarningMessage -Warnings $warnings -Message ("可読な VBA ソースを抽出できませんでした: {0}" -f $accessError)
    }
    finally {
        if ($null -ne $vbComponents) {
            Release-ComReference $vbComponents
        }
        if ($null -ne $vbProject) {
            Release-ComReference $vbProject
        }
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

    if (-not $hasVba -and $sourceExportStatus -ne 'failed') {
        $sourceExportStatus = 'no_vba'
    }

    $llmPackageStatus = 'skipped'
    $promptStatus = 'generated'
    if ($llmRecords.Count -gt 0) {
        Write-JsonLineFile -Items $llmRecords -Path $jsonlPath -Depth 20
        $llmPackageStatus = 'generated'
    }

    $manifest = [ordered]@{
        generated_at = Get-TimestampJst
        generator = 'Excel2LLM Macro Extract'
        status = if ($status -eq 'success' -and $warnings.Count -gt 0) { 'warning' } else { $status }
        workbook_name = $workbookName
        workbook_path = $displayWorkbookPath
        output_directory = $displayOutputDir
        extension = $extension
        has_vba = $hasVba
        raw_export_status = $rawExportStatus
        raw_project_path = $rawRelativePath
        source_export_status = $sourceExportStatus
        access_error = $accessError
        component_count = $componentRecords.Count
        llm_package_status = $llmPackageStatus
        llm_package_path = if ($llmPackageStatus -eq 'generated') { 'vba_llm_package.jsonl' } else { $null }
        prompt_status = $promptStatus
        prompt_path = 'vba_prompt.txt'
        warnings = @($warnings.ToArray())
        components = @($componentRecords.ToArray())
    }

    Write-JsonFile -Data $manifest -Path $manifestPath -Depth 20
    Write-VbaPromptFile -WorkbookName $workbookName -PromptPath $promptPath -JsonlRelativePath 'vba_llm_package.jsonl' -ManifestRelativePath 'macro_manifest.json' -AccessError $accessError -HasReadableSource:($llmPackageStatus -eq 'generated')

    Write-Host '=== VBA 抽出結果 ==='
    Write-Host ('  対象ファイル: {0}' -f $workbookName)
    Write-Host ('  可読ソース抽出: {0}' -f $sourceExportStatus)
    Write-Host ('  raw 保存:      {0}' -f $rawExportStatus)
    Write-Host ('  モジュール数:   {0}' -f $componentRecords.Count)
    Write-Host ('  LLM 用 JSONL:  {0}' -f $(if ($llmPackageStatus -eq 'generated') { $jsonlPath } else { '未生成' }))
    Write-Host ('  manifest:      {0}' -f $manifestPath)
    if ($warnings.Count -gt 0) {
        Write-Host ('  警告数:        {0}' -f $warnings.Count)
    }

    $nextSteps = [System.Collections.Generic.List[string]]::new()
    if ($llmPackageStatus -eq 'generated') {
        [void]$nextSteps.Add(('LLM に渡す対象: {0}' -f $jsonlPath))
        [void]$nextSteps.Add(('レビュー用の完成文: {0}' -f $promptPath))
    }
    else {
        [void]$nextSteps.Add(('抽出結果メモ: {0}' -f $manifestPath))
        if (-not [string]::IsNullOrWhiteSpace($accessError)) {
            [void]$nextSteps.Add('Excel の VBA プロジェクトアクセス設定を確認してから再実行する')
        }
    }
    Write-NextStepBlock -Steps @($nextSteps.ToArray())
}
catch {
    Write-ErrorRecoverySteps -CommandName 'macro_extract'
    throw "macro_extract.ps1 の $($_.InvocationInfo.ScriptLineNumber) 行目: $($_.Exception.Message)"
}
