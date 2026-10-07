# Shared helpers for Export-Vba.ps1 and Import-Vba.ps1.
# Dot-source only. Requires Windows, desktop Excel, and VBA object model access.

Set-StrictMode -Version Latest

$script:Ansi = [System.Text.Encoding]::GetEncoding(1252)

# vbext_ComponentType -> file extension
$script:VbaExt = @{
    1   = '.bas'     # vbext_ct_StdModule
    2   = '.cls'     # vbext_ct_ClassModule
    3   = '.frm'     # vbext_ct_MSForm
    100 = '.doccls'  # vbext_ct_Document (code only, no header)
}

$script:MacroCapable = @('.xlsm', '.xlsb', '.xls', '.xlam', '.xla', '.xltm', '.xlt')

function Assert-MacroCapable([string]$Path) {
    $ext = [IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($script:MacroCapable -notcontains $ext) {
        throw "NO_VBA_CONTAINER: '$ext' files cannot contain VBA. Ask the user before converting to .xlsm."
    }
}

function Open-VbaWorkbook {
    param(
        [Parameter(Mandatory)][string]$Path,
        [bool]$ReadOnly = $true
    )
    $full = (Resolve-Path -LiteralPath $Path).ProviderPath
    Assert-MacroCapable $full

    $session = [pscustomobject]@{ Excel = $null; Workbook = $null; Project = $null; Path = $full }
    try {
        $session.Excel = New-Object -ComObject Excel.Application
        $session.Excel.Visible = $false
        $session.Excel.DisplayAlerts = $false
        $session.Excel.EnableEvents = $false
        $session.Excel.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: no macro runs

        # Open(FileName, UpdateLinks=0, ReadOnly)
        $session.Workbook = $session.Excel.Workbooks.Open($full, 0, $ReadOnly)
        if (-not $ReadOnly -and $session.Workbook.ReadOnly) {
            throw "WORKBOOK_IN_USE: '$full' opened read-only. Another process (likely the user's Excel) has it open."
        }

        try {
            $session.Project = $session.Workbook.VBProject
            $null = $session.Project.Protection
        } catch {
            throw "VBA_OBJECT_MODEL_BLOCKED: Excel denied programmatic access to the VBA project ('Trust access to the VBA project object model' is off or blocked by policy). Do not change this setting. Report to the user."
        }

        if ($session.Project.Protection -eq 1) {   # vbext_pp_locked
            throw "VBA_PROJECT_LOCKED: the VBA project is password-protected. Do not attempt to bypass it."
        }
        return $session
    } catch {
        Close-VbaWorkbook -Session $session -Save $false
        throw
    }
}

function Close-VbaWorkbook {
    param($Session, [bool]$Save = $false)
    if ($null -eq $Session) { return }
    try {
        if ($null -ne $Session.Workbook) {
            if ($Save) { $Session.Workbook.Save() }
            $Session.Workbook.Close($false)
        }
    } finally {
        if ($null -ne $Session.Excel) { $Session.Excel.Quit() }
        foreach ($o in @($Session.Project, $Session.Workbook, $Session.Excel)) {
            if ($null -ne $o) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) }
        }
        $Session.Project = $null; $Session.Workbook = $null; $Session.Excel = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Find-VbaComponent($Project, [string]$Name) {
    foreach ($c in $Project.VBComponents) {
        if ($c.Name -ieq $Name) { return $c }
    }
    return $null
}

function ConvertTo-Crlf([string]$Text) {
    return ($Text -replace "`r`n", "`n" -replace "`r", "`n" -replace "`n", "`r`n")
}

# Strip the export header (VERSION/BEGIN..END block and leading Attribute VB_* lines).
function Get-VbaBody([string]$Text) {
    $lines = (ConvertTo-Crlf $Text) -split "`r`n"
    $i = 0
    if ($lines.Count -gt 0 -and $lines[0] -match '^VERSION\s') {
        $i = 1
        # .cls closes with 'END', .frm closes with 'End' at column 0; nested designer blocks are indented.
        while ($i -lt $lines.Count -and $lines[$i] -notmatch '^(END|End)\s*$') { $i++ }
        $i++
    }
    while ($i -lt $lines.Count -and $lines[$i] -match '^Attribute VB_') { $i++ }
    if ($i -ge $lines.Count) { return '' }
    return ($lines[$i..($lines.Count - 1)] -join "`r`n")
}

function Get-ModuleCode($Component) {
    $cm = $Component.CodeModule
    $n = $cm.CountOfLines
    if ($n -gt 0) { return $cm.Lines(1, $n) }
    return ''
}

function Set-ModuleCode($Component, [string]$Code) {
    $cm = $Component.CodeModule
    $n = $cm.CountOfLines
    if ($n -gt 0) { $cm.DeleteLines(1, $n) }
    $Code = $Code.TrimEnd([char[]]"`r`n")
    if ($Code.Length -gt 0) { $cm.AddFromString($Code) }
}

# Case-insensitive, trailing-whitespace-insensitive comparison (the VBE normalizes both).
function Test-CodeEqual([string]$A, [string]$B) {
    $norm = {
        param($t)
        ((ConvertTo-Crlf $t) -split "`r`n" | ForEach-Object { $_.TrimEnd() }) -join "`n" |
            ForEach-Object { $_.TrimEnd("`n") }
    }
    return ((& $norm $A) -eq (& $norm $B))   # -eq on strings is case-insensitive
}

function Read-AnsiFile([string]$Path) {
    $bytes = [IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        throw "ENCODING_ERROR: '$Path' has a UTF-8 BOM. Save it as ANSI (Windows-1252) or pure ASCII."
    }
    if ($bytes.Length -ge 2 -and (($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) -or ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF))) {
        throw "ENCODING_ERROR: '$Path' is UTF-16. Save it as ANSI (Windows-1252) or pure ASCII."
    }
    return $script:Ansi.GetString($bytes)
}
