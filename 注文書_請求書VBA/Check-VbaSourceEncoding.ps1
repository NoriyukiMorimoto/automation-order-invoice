#Requires -Version 5.1
<#
.SYNOPSIS
  運用中VBA (.bas / .frm / .cls) が CP932 (Shift_JIS) 正本として妥当か検証する。
  UTF-8 のみデコード可能で CP932 として不正なファイルは VBE インポート時の文字化け原因になる。
#>
param(
    [string]$SourceDir = "",
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($SourceDir)) {
    $SourceDir = Join-Path $scriptDir "運用中VBA"
}
$SourceDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourceDir)

function Test-VbaSourceBytes([byte[]]$Bytes, [string]$PathLabel) {
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        return @{ Ok = $false; Reason = "UTF-8 BOM (CP932 正本に BOM は不要)" }
    }

    $enc932 = [System.Text.Encoding]::GetEncoding(932)
    $utf8Strict = New-Object System.Text.UTF8Encoding $false, $true

    $cp932Ok = $true
    try {
        $enc932.GetString($Bytes) | Out-Null
    }
    catch {
        $cp932Ok = $false
    }

    $utf8Ok = $true
    try {
        $utf8Strict.GetString($Bytes) | Out-Null
    }
    catch {
        $utf8Ok = $false
    }

    if ($cp932Ok -and (-not $utf8Ok)) {
        return @{ Ok = $true; Reason = "CP932" }
    }
    if ($cp932Ok -and $utf8Ok) {
        return @{ Ok = $true; Reason = "ASCII/共通" }
    }
    if ((-not $cp932Ok) -and $utf8Ok) {
        return @{ Ok = $false; Reason = "UTF-8 として読めるが CP932 として不正 (文字化け正本)" }
    }

    return @{ Ok = $false; Reason = "CP932/UTF-8 いずれでもデコード不可" }
}

$patterns = @("*.bas", "*.frm", "*.cls")
$failures = New-Object System.Collections.Generic.List[string]

foreach ($pattern in $patterns) {
    Get-ChildItem -Path $SourceDir -Filter $pattern -File | ForEach-Object {
        $bytes = [System.IO.File]::ReadAllBytes($_.FullName)
        $result = Test-VbaSourceBytes $bytes $_.FullName
        if (-not $result.Ok) {
            [void]$failures.Add("$($_.Name): $($result.Reason)")
        }
        elseif (-not $Quiet) {
            Write-Host "OK $($_.Name) ($($result.Reason))"
        }
    }
}

if ($failures.Count -gt 0) {
    Write-Error ("VBA ソースの文字コード検証に失敗しました:`n" + ($failures -join "`n"))
    exit 1
}

if (-not $Quiet) {
    Write-Host "All VBA sources passed CP932 validation."
}
exit 0
