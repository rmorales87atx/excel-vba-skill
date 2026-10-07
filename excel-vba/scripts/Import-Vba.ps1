<#
.SYNOPSIS
  Write edited VBA source files back into a workbook.
.DESCRIPTION
  Backs up the workbook, opens it in a hidden Excel instance with macros disabled,
  and for each selected file:
    - existing component      -> replace code body via CodeModule (header kept as-is)
    - existing .bas/.cls with member Attribute lines -> remove + re-import whole file
    - new .bas/.cls/.frm      -> VBComponents.Import
    - .doccls                 -> replace code body; the sheet must already exist
  Reads each component back and prints WARN on mismatch. Never deletes components.
.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File Import-Vba.ps1 -WorkbookPath .\Book.xlsm -SourceDir .\vba-src -Only Module1,Sheet1
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$WorkbookPath,
    [Parameter(Mandatory)][string]$SourceDir,
    [string[]]$Only,
    [switch]$NoBackup
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'VbaCommon.ps1')

$wbFull  = (Resolve-Path -LiteralPath $WorkbookPath).ProviderPath
$srcFull = (Resolve-Path -LiteralPath $SourceDir).ProviderPath
Assert-MacroCapable $wbFull

# Collect source files.
$files = Get-ChildItem -LiteralPath $srcFull -File |
    Where-Object { @('.bas', '.cls', '.frm', '.doccls') -contains $_.Extension.ToLowerInvariant() }
if ($Only) {
    $files = $files | Where-Object { $Only -icontains $_.BaseName }
    $missing = $Only | Where-Object { $n = $_; -not ($files | Where-Object { $_.BaseName -ieq $n }) }
    if ($missing) { throw "SOURCE_MISSING: no source file for: $($missing -join ', ')" }
}
if (-not $files) { throw "SOURCE_MISSING: no .bas/.cls/.frm/.doccls files selected in '$srcFull'." }

# Validate encoding before touching the workbook.
$sources = @{}
foreach ($f in $files) { $sources[$f.FullName] = Read-AnsiFile $f.FullName }

if (-not $NoBackup) {
    $stamp  = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backup = Join-Path ([IO.Path]::GetDirectoryName($wbFull)) (
        [IO.Path]::GetFileNameWithoutExtension($wbFull) + ".bak-$stamp" + [IO.Path]::GetExtension($wbFull))
    Copy-Item -LiteralPath $wbFull -Destination $backup
    Write-Output "BACKUP: $backup"
}

$session = $null
$changed = 0
try {
    $session = Open-VbaWorkbook -Path $wbFull -ReadOnly $false
    $proj = $session.Project

    if ($session.Workbook.VBASigned) {
        Write-Output "SIGNED_PROJECT: this import invalidates the VBA digital signature."
    }

    foreach ($f in $files) {
        $name = $f.BaseName
        $ext  = $f.Extension.ToLowerInvariant()
        $text = $sources[$f.FullName]
        $body = if ($ext -eq '.doccls') { ConvertTo-Crlf $text } else { Get-VbaBody $text }
        $comp = Find-VbaComponent $proj $name

        if ($ext -eq '.doccls') {
            if ($null -eq $comp -or [int]$comp.Type -ne 100) {
                throw "DOCUMENT_MODULE_MISSING: no document module named '$name' in the workbook."
            }
            Set-ModuleCode $comp $body
            Write-Output "UPDATED: $name (document module)"
        }
        elseif ($null -eq $comp) {
            $null = $proj.VBComponents.Import($f.FullName)
            $comp = Find-VbaComponent $proj $name
            if ($null -eq $comp) {
                throw "IMPORT_NAME_MISMATCH: imported '$($f.Name)' but no component named '$name' exists. Check the Attribute VB_Name header."
            }
            Write-Output "ADDED: $name"
        }
        elseif ($ext -ne '.frm' -and $body -match '(?m)^\s*Attribute\s+\w+\.VB_') {
            # Member attributes survive only through Import.
            $proj.VBComponents.Remove($comp)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($comp)
            $null = $proj.VBComponents.Import($f.FullName)
            $comp = Find-VbaComponent $proj $name
            if ($null -eq $comp) {
                throw "IMPORT_NAME_MISMATCH: re-import of '$name' produced a different component name. Restore from backup."
            }
            Write-Output "REPLACED: $name (re-imported to keep member attributes)"
        }
        else {
            Set-ModuleCode $comp $body
            Write-Output "UPDATED: $name"
        }

        # Read-back check (member Attribute lines are hidden by the VBE, so skip them).
        $expected = (($body -split "`r`n") | Where-Object { $_ -notmatch '^\s*Attribute\s+\w+\.VB_' }) -join "`r`n"
        $actual   = Get-ModuleCode $comp
        if (-not (Test-CodeEqual $expected $actual)) {
            Write-Output "WARN: read-back of '$name' differs from source. Re-export and diff."
        }
        $changed++
    }

    $session.Workbook.Save()
    Write-Output "DONE: $changed component(s) written. Macros were not run or compiled; run Debug > Compile VBAProject."
}
catch {
    Write-Error $_.Exception.Message
    Write-Output "NOT SAVED: workbook left unchanged on disk."
    Close-VbaWorkbook -Session $session -Save $false
    $session = $null
    exit 1
}
finally {
    if ($null -ne $session) { Close-VbaWorkbook -Session $session -Save $false }
}
