#Requires -Version 5.1
<#
.SYNOPSIS
  VBA ソースフォルダのコンポーネントを Excel ブックへインポートします。
  Excel は起動したまま実行してください。対象 .xlsm が未表示の場合は -WorkbookPath から自動で開きます。
  前提: [ファイル] → [オプション] → [トラスト センター] → [マクロの設定]
        「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」にチェック
#>
param(
    [string]$WorkbookName = "自動化_注文書_請求書.xlsm",
    [string]$WorkbookPath = "",
    [string]$SourceDir = "",
    [string]$ModuleName = "",
    [string]$CommitMessage = "",
    [switch]$All,
    [switch]$SkipGit
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$workspaceRoot = Split-Path -Parent $scriptDir

if ([string]::IsNullOrWhiteSpace($SourceDir)) {
    $sourceDir = Join-Path $scriptDir "運用中VBA"
} else {
    $sourceDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourceDir)
}

$vbextCtStdModule = 1
$vbextCtClassModule = 2
$vbextCtMsForm = 3
$vbextCtDocument = 100

function Read-TextFile([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)
    }
    return [System.Text.Encoding]::GetEncoding(932).GetString($bytes)
}

function Get-VbNameFromSource([string]$Path, [string]$DefaultName) {
    $text = Read-TextFile $Path
    if ($text -match 'Attribute VB_Name = "([^"]+)"') {
        return $Matches[1]
    }
    return $DefaultName
}

function Get-CodeLinesFromSource([string]$Path) {
    $lines = (Read-TextFile $Path) -split "`r?`n"
    $result = New-Object System.Collections.Generic.List[string]
    foreach ($line in $lines) {
        $trim = $line.Trim()
        if ($trim -match '^VERSION ') { continue }
        if ($trim -eq 'BEGIN') { continue }
        if ($trim -eq 'END') { continue }
        if ($trim -match '^Attribute ') { continue }
        $result.Add($line)
    }
    while ($result.Count -gt 0 -and [string]::IsNullOrWhiteSpace($result[$result.Count - 1])) {
        $result.RemoveAt($result.Count - 1)
    }
    return $result
}

function New-ImportableBasFile([string]$Path, [string]$ModuleName) {
    $text = Read-TextFile $Path
    if ($text -match 'Attribute VB_Name = ') {
        return $Path
    }
    $tempPath = Join-Path ([System.IO.Path]::GetTempPath()) ("vba_import_" + $ModuleName + ".bas")
    $header = "Attribute VB_Name = ""$ModuleName""" + [Environment]::NewLine
    [System.IO.File]::WriteAllText($tempPath, $header + $text, [System.Text.Encoding]::GetEncoding(932))
    return $tempPath
}

function Find-TargetWorkbook($Excel, [string]$TargetName, [switch]$AllowOrderInvoiceFallback) {
    foreach ($wb in $Excel.Workbooks) {
        if ($wb.Name -eq $TargetName) {
            return $wb
        }
    }
    if ($AllowOrderInvoiceFallback) {
        foreach ($wb in $Excel.Workbooks) {
            if ($wb.Name -like "*.xlsm" -and $wb.Name -like "*注文書*") {
                return $wb
            }
        }
    }
    return $null
}

function Resolve-WorkbookPath([string]$TargetName, [string]$ExplicitPath) {
    if (-not [string]::IsNullOrWhiteSpace($ExplicitPath)) {
        return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($ExplicitPath)
    }
    switch ($TargetName) {
        "自動化_注文書_請求書.xlsm" {
            return Join-Path $scriptDir "自動化_注文書_請求書.xlsm"
        }
        "単価一覧表作成ツール.xlsm" {
            return Join-Path $workspaceRoot "マスタデータ\単価マスタ\単価一覧表作成ツール.xlsm"
        }
    }
    return ""
}

function Get-OrOpen-TargetWorkbook($Excel, [string]$TargetName, [string]$TargetPath, [switch]$AllowOrderInvoiceFallback) {
    $workbook = Find-TargetWorkbookAcrossInstances $TargetName $AllowOrderInvoiceFallback
    if ($null -ne $workbook) {
        return $workbook
    }

    $workbook = Find-TargetWorkbook $Excel $TargetName $AllowOrderInvoiceFallback
    if ($null -ne $workbook) {
        return $workbook
    }
    if ([string]::IsNullOrWhiteSpace($TargetPath)) {
        return $null
    }
    if (-not (Test-Path $TargetPath)) {
        throw "Target workbook path not found: $TargetPath"
    }
    Write-Host "Opening workbook: $TargetPath"
    return $Excel.Workbooks.Open($TargetPath)
}

function Get-AllExcelApplications {
    $apps = New-Object System.Collections.Generic.List[object]
    $seen = New-Object System.Collections.Generic.HashSet[int]

    foreach ($process in @(Get-Process EXCEL -ErrorAction SilentlyContinue |
                            Where-Object { $_.MainWindowHandle -ne 0 } |
                            Sort-Object StartTime -Descending)) {
        try {
            $excel = [ExcelWindowAccess]::GetExcelApplication($process.MainWindowHandle)
            if ($null -ne $excel) {
                $hwnd = [int]$excel.Hwnd
                if (-not $seen.Contains($hwnd)) {
                    [void]$seen.Add($hwnd)
                    [void]$apps.Add($excel)
                }
            }
        }
        catch {
            # VBE 前面等で HWND から Application を取れない場合は次候補へ。
        }
    }

    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        if ($null -ne $excel) {
            $hwnd = [int]$excel.Hwnd
            if (-not $seen.Contains($hwnd)) {
                [void]$seen.Add($hwnd)
                [void]$apps.Add($excel)
            }
        }
    }
    catch {}

    return $apps
}

function Find-TargetWorkbookAcrossInstances([string]$TargetName, [switch]$AllowOrderInvoiceFallback) {
    foreach ($excel in @(Get-AllExcelApplications)) {
        $workbook = Find-TargetWorkbook $excel $TargetName $AllowOrderInvoiceFallback
        if ($null -ne $workbook) {
            return $workbook
        }
    }
    return $null
}

function Get-ExcelApplication {
    if ($null -eq ("ExcelWindowAccess" -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class ExcelWindowAccess
{
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(
        IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(
        IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(
        IntPtr hwnd, uint dwObjectID, ref Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out object ppvObject);

    public static object GetExcelApplication(IntPtr mainWindow)
    {
        var excelWindows = new List<IntPtr>();
        EnumChildWindows(mainWindow, delegate(IntPtr hWnd, IntPtr lParam)
        {
            var className = new StringBuilder(256);
            GetClassName(hWnd, className, className.Capacity);
            if (className.ToString() == "EXCEL7")
            {
                excelWindows.Add(hWnd);
            }
            return true;
        }, IntPtr.Zero);

        var dispatchId = new Guid("00020400-0000-0000-C000-000000000046");
        foreach (var excelWindow in excelWindows)
        {
            object nativeObject;
            int result = AccessibleObjectFromWindow(
                excelWindow, 0xFFFFFFF0, ref dispatchId, out nativeObject);
            if (result != 0 || nativeObject == null) continue;

            try
            {
                return nativeObject.GetType().InvokeMember(
                    "Application",
                    System.Reflection.BindingFlags.GetProperty,
                    null, nativeObject, null);
            }
            finally
            {
                Marshal.FinalReleaseComObject(nativeObject);
            }
        }
        return null;
    }
}
"@
    }

    # ウィンドウ表示中の Excel を優先する(GetActiveObject だと空の別インスタンスに繋がることがある)。
    foreach ($process in @(Get-Process EXCEL -ErrorAction SilentlyContinue |
                            Where-Object { $_.MainWindowHandle -ne 0 } |
                            Sort-Object StartTime -Descending)) {
        try {
            $excel = [ExcelWindowAccess]::GetExcelApplication($process.MainWindowHandle)
            if ($null -ne $excel) {
                return $excel
            }
        }
        catch {
            continue
        }
    }

    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
    }
    catch {
        throw "Running Excel application not found."
    }
}

function Remove-ComponentIfExists($VBProject, [string]$Name, [int[]]$AllowedTypes) {
    foreach ($comp in @($VBProject.VBComponents)) {
        if ($comp.Name -eq $Name -and ($AllowedTypes -contains [int]$comp.Type)) {
            $VBProject.VBComponents.Remove($comp)
            return $true
        }
    }
    return $false
}

function Find-BasicInfoWorksheet($Workbook) {
    if ($null -eq $Workbook) { return $null }
    foreach ($ws in @($Workbook.Worksheets)) {
        if ($ws.Name -eq "基本情報") { return $ws }
    }
    return $null
}

function Resolve-DocumentModule($VBProject, [string]$Name) {
    # Sheet1 / 基本情報は「表示名=基本情報」のシートへ必ず紐付ける。
    # CodeName だけを見ると、作り直しで Sheet110 等に変わったあと
    # 孤立した旧 Sheet1 へ書き込んでイベントが死ぬ事故が起きる。
    if ($Name -eq "Sheet1" -or $Name -eq "基本情報") {
        $workbook = $null
        try { $workbook = $VBProject.Parent } catch { $workbook = $null }
        $basicWs = Find-BasicInfoWorksheet $workbook
        if ($null -ne $basicWs) {
            $codeName = [string]$basicWs.CodeName
            $bound = $null
            foreach ($candidate in @($VBProject.VBComponents)) {
                if ($candidate.Name -eq $codeName -and [int]$candidate.Type -eq $vbextCtDocument) {
                    $bound = $candidate
                    break
                }
            }
            if ($null -ne $bound) {
                if ($bound.Name -ne "Sheet1") {
                    $conflict = $null
                    foreach ($candidate in @($VBProject.VBComponents)) {
                        if ($candidate.Name -eq "Sheet1" -and [int]$candidate.Type -eq $vbextCtDocument) {
                            $conflict = $candidate
                            break
                        }
                    }
                    if ($null -ne $conflict -and -not [object]::ReferenceEquals($conflict, $bound)) {
                        $orphanName = "Sheet1Orphan"
                        $n = 1
                        while ($true) {
                            $exists = $false
                            foreach ($candidate in @($VBProject.VBComponents)) {
                                if ($candidate.Name -eq $orphanName) { $exists = $true; break }
                            }
                            if (-not $exists) { break }
                            $orphanName = "Sheet1Orphan$n"
                            $n++
                        }
                        if ([int]$conflict.CodeModule.CountOfLines -gt 0) {
                            $conflict.CodeModule.DeleteLines(1, [int]$conflict.CodeModule.CountOfLines)
                        }
                        $conflict.Name = $orphanName
                        Write-Host "  Renamed conflicting document module to: $orphanName"
                    }
                    $bound.Name = "Sheet1"
                    Write-Host "  Restored basic-info CodeName to Sheet1 (was $codeName)"
                }
                return $bound
            }
        }
    }

    foreach ($candidate in @($VBProject.VBComponents)) {
        if ($candidate.Name -eq $Name -and [int]$candidate.Type -eq $vbextCtDocument) {
            return $candidate
        }
    }
    return $null
}

function Update-DocumentModule($VBProject, [string]$Name, [string]$FilePath) {
    $comp = Resolve-DocumentModule $VBProject $Name
    if ($null -eq $comp) {
        Write-Warning "Document module not found in workbook: $Name (skipped)"
        return
    }

    $lines = Get-CodeLinesFromSource $FilePath
    $codeModule = $comp.CodeModule
    $lineCount = [int]$codeModule.CountOfLines
    if ($lineCount -gt 0) {
        $codeModule.DeleteLines(1, $lineCount)
    }
    if ($lines.Count -gt 0) {
        $codeModule.AddFromString(($lines -join [Environment]::NewLine))
    }
    Write-Host "  Updated document module: $($comp.Name) (requested: $Name)"
}

function Remove-StandardModuleVariants($VBProject, [string]$ModuleName) {
    $escapedName = [regex]::Escape($ModuleName)
    $variantPattern = "^$escapedName(\d+)?$"
    foreach ($comp in @($VBProject.VBComponents)) {
        if ([int]$comp.Type -ne $vbextCtStdModule) { continue }
        if ($comp.Name -match $variantPattern) {
            $VBProject.VBComponents.Remove($comp)
        }
    }
}

function Remove-OrphanStandardModules($VBProject, [string[]]$ExpectedModuleNames) {
    $expected = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($name in $ExpectedModuleNames) {
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            [void]$expected.Add($name)
        }
    }

    foreach ($comp in @($VBProject.VBComponents)) {
        if ([int]$comp.Type -ne $vbextCtStdModule) { continue }
        if ($expected.Contains($comp.Name)) { continue }
        if ($comp.Name -match '^\d+$') {
            Write-Host "  Removed orphan standard module: $($comp.Name)"
            $VBProject.VBComponents.Remove($comp)
            continue
        }
        if ($comp.Name -match '^(mod_|ModuleExport)') {
            Write-Host "  Removed unexpected standard module: $($comp.Name)"
            $VBProject.VBComponents.Remove($comp)
        }
    }
}

function Save-WorkbookSafely($Workbook, [string]$LocalFallbackPath) {
    try {
        $Workbook.Save()
        return
    }
    catch {
        Write-Warning "Workbook.Save failed: $($_.Exception.Message)"
    }

    if ([string]::IsNullOrWhiteSpace($LocalFallbackPath)) {
        throw "Workbook save failed and no local fallback path is available."
    }

    $fallbackDir = Split-Path -Parent $LocalFallbackPath
    if (-not (Test-Path $fallbackDir)) {
        throw "Local fallback folder not found: $fallbackDir"
    }

    Write-Host "Saving repaired copy to local path: $LocalFallbackPath"
    $Workbook.SaveCopyAs($LocalFallbackPath)
}

function Import-StandardModule($VBProject, [string]$FilePath) {
    $moduleName = Get-VbNameFromSource $FilePath ([System.IO.Path]::GetFileNameWithoutExtension($FilePath))
    Remove-StandardModuleVariants $VBProject $moduleName
    $importPath = New-ImportableBasFile $FilePath $moduleName
    $imported = $VBProject.VBComponents.Import($importPath)
    if ($imported.Name -ne $moduleName) {
        $imported.Name = $moduleName
    }
    if ($importPath -ne $FilePath -and (Test-Path $importPath)) {
        Remove-Item $importPath -Force
    }
    Write-Host "  Imported standard module: $moduleName"
}

function Test-IsDocumentClassFile([string]$Path, [string]$BaseName) {
    if ($BaseName -eq "Sheet1" -or $BaseName -eq "ThisWorkbook" -or $BaseName -eq "基本情報") {
        return $true
    }
    if ($BaseName -match '^Sheet\d+$') {
        return $true
    }
    $text = Read-TextFile $Path
    # 通常のクラスモジュールは VERSION 1.0 CLASS ヘッダーを持つ。
    # ドキュメントモジュールの正本（Sheet1 / ThisWorkbook）は Option Explicit 始まり。
    if ($text -match '(?m)^VERSION\s+1\.0\s+CLASS') {
        return $false
    }
    return $false
}

function Import-ClassModule($VBProject, [string]$FilePath) {
    $moduleName = Get-VbNameFromSource $FilePath ([System.IO.Path]::GetFileNameWithoutExtension($FilePath))
    Remove-ComponentIfExists $VBProject $moduleName @($vbextCtClassModule) | Out-Null
    $imported = $VBProject.VBComponents.Import($FilePath)
    if ([int]$imported.Type -ne $vbextCtClassModule) {
        $VBProject.VBComponents.Remove($imported)
        throw "Class module import failed (unexpected type): $moduleName"
    }
    if ($imported.Name -ne $moduleName) {
        $imported.Name = $moduleName
    }
    Write-Host "  Imported class module: $moduleName"
}

function Update-FormCodeModule($VBProject, [string]$FormName, [string]$FilePath) {
    $comp = $null
    foreach ($candidate in @($VBProject.VBComponents)) {
        if ($candidate.Name -eq $FormName -and [int]$candidate.Type -eq $vbextCtMsForm) {
            $comp = $candidate
            break
        }
    }
    if ($null -eq $comp) {
        throw "Form module not found in workbook: $FormName"
    }

    $lines = Get-CodeLinesFromSource $FilePath
    $codeModule = $comp.CodeModule
    $lineCount = [int]$codeModule.CountOfLines
    if ($lineCount -gt 0) {
        $codeModule.DeleteLines(1, $lineCount)
    }
    if ($lines.Count -gt 0) {
        $codeModule.AddFromString(($lines -join [Environment]::NewLine))
    }
    Write-Host "  Updated form code: $FormName"
}

function Import-FormModule($VBProject, [string]$FilePath) {
    $formName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $sourceText = Read-TextFile $FilePath
    if ($sourceText -notmatch '(?m)^VERSION ') {
        Update-FormCodeModule $VBProject $formName $FilePath
        return
    }

    Remove-ComponentIfExists $VBProject $formName @($vbextCtMsForm) | Out-Null
    Remove-MisimportedFormModule $VBProject $formName
    $imported = $VBProject.VBComponents.Import($FilePath)
    if ([int]$imported.Type -ne $vbextCtMsForm) {
        $VBProject.VBComponents.Remove($imported)
        throw "Form import failed (not a UserForm): $formName"
    }
    Write-Host "  Imported form: $formName"
}

function Remove-MisimportedFormModule($VBProject, [string]$FormName) {
    foreach ($comp in @($VBProject.VBComponents)) {
        if ([int]$comp.Type -ne $vbextCtStdModule) { continue }
        if ($comp.CodeModule.CountOfLines -le 0) { continue }
        $preview = $comp.CodeModule.Lines(1, [Math]::Min(40, $comp.CodeModule.CountOfLines))
        if ($preview -match 'UserForm_Initialize' -and $preview -match 'SelectionConfirmed') {
            $VBProject.VBComponents.Remove($comp)
            Write-Host "  Removed misimported form module: $($comp.Name)"
        }
    }
}

function Get-ModuleNameBase([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
    return [System.IO.Path]::GetFileNameWithoutExtension($Name)
}

function Get-RelativePathFromBase {
    param(
        [string]$BasePath,
        [string]$TargetPath
    )

    $baseUri = New-Object System.Uri(($BasePath.TrimEnd('\') + '\'))
    $targetUri = New-Object System.Uri($TargetPath)
    $relativeUri = $baseUri.MakeRelativeUri($targetUri)
    return [System.Uri]::UnescapeDataString($relativeUri.ToString()).Replace('/', '\')
}

function Get-GitRepositoryRoot {
    $current = $workspaceRoot
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        if (Test-Path (Join-Path $current ".git")) {
            return $current
        }
        $parent = Split-Path -Parent $current
        if ($parent -eq $current) { break }
        $current = $parent
    }
    return $null
}

function Publish-GitChanges {
    param(
        [string[]]$ImportedFiles,
        [string]$TargetWorkbookName,
        [string]$Message
    )

    if ($SkipGit) {
        Write-Host "Git: skipped (-SkipGit)."
        return
    }

    $repoRoot = Get-GitRepositoryRoot
    if ($null -eq $repoRoot) {
        Write-Warning "Git: repository root not found. Skipping commit and push."
        return
    }

    if ($ImportedFiles.Count -eq 0) {
        Write-Host "Git: no imported source files. Skipping commit and push."
        return
    }

    Push-Location $repoRoot
    try {
        $currentBranch = (git rev-parse --abbrev-ref HEAD 2>$null).Trim()
        if ($currentBranch -ne "main") {
            Write-Host "Git: switching branch $currentBranch -> main"
            git checkout main | Out-Host
        }

        $pathsToAdd = New-Object System.Collections.Generic.List[string]
        foreach ($importedFile in $ImportedFiles) {
            $fullPath = [System.IO.Path]::GetFullPath($importedFile)
            $relativePath = Get-RelativePathFromBase $repoRoot $fullPath
            if ($relativePath -notmatch '\.\.') {
                [void]$pathsToAdd.Add($relativePath)
            }
        }

        if ($pathsToAdd.Count -eq 0) {
            Write-Host "Git: no repository-relative files to add."
            return
        }

        foreach ($pathToAdd in $pathsToAdd) {
            git add -- "$pathToAdd" | Out-Host
        }

        $pending = git status --porcelain -- $pathsToAdd
        if ([string]::IsNullOrWhiteSpace($pending)) {
            Write-Host "Git: no staged changes to commit."
            return
        }

        if ([string]::IsNullOrWhiteSpace($Message)) {
            $moduleNames = ($ImportedFiles | ForEach-Object {
                [System.IO.Path]::GetFileNameWithoutExtension($_)
            }) -join ", "
            $Message = "VBA反映: $moduleNames ($TargetWorkbookName)"
        }

        git commit -m $Message | Out-Host
        git push origin main | Out-Host
        Write-Host "Git: pushed to origin/main."
    }
    finally {
        Pop-Location
    }
}

function Test-ShouldImport([string]$FileName, [string]$SourceDirPath) {
    if ($FileName -eq "ModuleExport.bas" -and -not $All -and
        (Get-ModuleNameBase $ModuleName) -ne "ModuleExport") {
        return $false
    }
    # Sheet1.cls が正本。基本情報.cls は旧エクスポートで Sheet1 を上書きし VBA 破損の原因になるため除外。
    if ($FileName -eq "基本情報.cls" -and (Test-Path (Join-Path $SourceDirPath "Sheet1.cls"))) {
        return $false
    }
    if ($All) { return $true }
    if ([string]::IsNullOrWhiteSpace($ModuleName)) {
        return $FileName -eq "mod_VendorMaster.bas"
    }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    return ($base -eq (Get-ModuleNameBase $ModuleName))
}

try {
    if (-not (Test-Path $sourceDir)) {
        throw "Source folder not found: $sourceDir"
    }

    $encodingCheckScript = Join-Path $scriptDir "Check-VbaSourceEncoding.ps1"
    if (Test-Path $encodingCheckScript) {
        Write-Host "Checking VBA source encoding (CP932)..."
        & $encodingCheckScript -SourceDir $sourceDir -Quiet
        if ($LASTEXITCODE -ne 0) {
            throw "VBA source encoding check failed. Fix .bas/.frm/.cls to CP932 before import."
        }
    }

    $excel = Get-ExcelApplication
    $excel.DisplayAlerts = $false

    $resolvedWorkbookPath = Resolve-WorkbookPath $WorkbookName $WorkbookPath
    $allowOrderInvoiceFallback = ($WorkbookName -eq "自動化_注文書_請求書.xlsm")
    $workbook = Get-OrOpen-TargetWorkbook $excel $WorkbookName $resolvedWorkbookPath $allowOrderInvoiceFallback
    if ($null -eq $workbook) {
        throw "Target workbook not found. Open '$WorkbookName' in Excel, or pass -WorkbookPath."
    }

    Write-Host "Target workbook: $($workbook.Name)"
    Write-Host "Source folder: $sourceDir"
    if ($resolvedWorkbookPath -ne "") {
        Write-Host "Workbook path: $resolvedWorkbookPath"
    }

    $vbProject = $workbook.VBProject
    $files = Get-ChildItem -Path $sourceDir -File | Sort-Object Extension, Name
    $importedFiles = New-Object System.Collections.Generic.List[string]
    $expectedStandardModules = New-Object System.Collections.Generic.List[string]

    foreach ($file in $files) {
        if (-not (Test-ShouldImport $file.Name $sourceDir)) { continue }

        switch ($file.Extension.ToLower()) {
            ".bas" {
                $moduleName = Get-VbNameFromSource $file.FullName $file.BaseName
                Import-StandardModule $vbProject $file.FullName
                [void]$expectedStandardModules.Add($moduleName)
            }
            ".cls" {
                if (Test-IsDocumentClassFile $file.FullName $file.BaseName) {
                    $docName = Get-VbNameFromSource $file.FullName $file.BaseName
                    if ($file.BaseName -eq "基本情報") {
                        $docName = "Sheet1"
                    }
                    Update-DocumentModule $vbProject $docName $file.FullName
                }
                else {
                    Import-ClassModule $vbProject $file.FullName
                }
            }
            ".frm" {
                Import-FormModule $vbProject $file.FullName
            }
            default {
                continue
            }
        }
        [void]$importedFiles.Add($file.FullName)
    }

    if ($importedFiles.Count -eq 0) {
        throw "No matching VBA source files to import."
    }

    if ($All -and $expectedStandardModules.Count -gt 0) {
        Remove-OrphanStandardModules $vbProject $expectedStandardModules.ToArray()
    }

    Save-WorkbookSafely $workbook $resolvedWorkbookPath
    Write-Host "Done. Workbook saved."

    Publish-GitChanges -ImportedFiles $importedFiles.ToArray() `
        -TargetWorkbookName $WorkbookName `
        -Message $CommitMessage
}
catch {
    Write-Error $_
    if ($_.Exception.Message -match "2147467260|2146827284|拒否|denied|programmatic access") {
        Write-Host ""
        Write-Host "VBA project access denied. Enable:" -ForegroundColor Yellow
        Write-Host "  Excel Options -> Trust Center -> Macro Settings"
        Write-Host "  'Trust access to the VBA project object model'"
    }
    exit 1
}
