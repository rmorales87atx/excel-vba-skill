<#
.SYNOPSIS
  Export every VBA component of a workbook to text files for editing.
.DESCRIPTION
  Opens the workbook read-only in a hidden Excel instance with macros disabled.
  Writes .bas / .cls / .frm(+.frx) via VBComponent.Export, and document modules
  (ThisWorkbook, sheets) as code-only .doccls files. Writes vba-manifest.json.
.EXAMPLE
  powershell -NoProfile -ExecutionPolicy Bypass -File Export-Vba.ps1 -WorkbookPath .\Book.xlsm -OutDir .\vba-src
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$WorkbookPath,
    [Parameter(Mandatory)][string]$OutDir
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'VbaCommon.ps1')

$session = $null
try {
    $session = Open-VbaWorkbook -Path $WorkbookPath -ReadOnly $true

    if ($session.Workbook.VBASigned) {
        Write-Output "SIGNED_PROJECT: the VBA project is digitally signed. Any import will invalidate the signature."
    }

    New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
    $outFull = (Resolve-Path -LiteralPath $OutDir).ProviderPath

    $manifest = @()
    foreach ($comp in $session.Project.VBComponents) {
        $type = [int]$comp.Type
        if (-not $script:VbaExt.ContainsKey($type)) {
            Write-Output "WARN: skipped component '$($comp.Name)' with unsupported type $type."
            continue
        }
        $file = $comp.Name + $script:VbaExt[$type]
        $target = Join-Path $outFull $file

        if ($type -eq 100) {
            $code = Get-ModuleCode $comp
            [IO.File]::WriteAllText($target, (ConvertTo-Crlf $code), $script:Ansi)
        } else {
            if (Test-Path -LiteralPath $target) { Remove-Item -LiteralPath $target -Force }
            $comp.Export($target)   # VBE writes ANSI + CRLF; .frm also writes .frx
        }

        $manifest += [pscustomobject]@{ name = $comp.Name; type = $type; file = $file }
        Write-Output "EXPORTED: $file"
    }

    $manifest | ConvertTo-Json -Depth 3 |
        Set-Content -LiteralPath (Join-Path $outFull 'vba-manifest.json') -Encoding ASCII
    Write-Output "DONE: $($manifest.Count) component(s) exported to $outFull"
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
finally {
    Close-VbaWorkbook -Session $session -Save $false
}
