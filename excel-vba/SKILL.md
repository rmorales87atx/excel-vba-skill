---
name: excel-vba
description: Extract, read, and update VBA macro code inside Excel workbooks (.xlsm, .xlsb, .xls, .xlam). Use when a task needs to see macro source, change or add a VBA module, or diff VBA between workbook versions. Not for worksheet formulas, Office Scripts, or .xlsx files (they cannot contain VBA).
---

# Excel VBA: extract and update

## Facts that drive every decision

- VBA is stored in a binary OLE container (`vbaProject.bin`) as compressed source **plus** compiled p-code. **Never edit `vbaProject.bin` or the workbook zip directly.** Excel can execute the stale p-code instead of your new source, or flag the project as corrupt.
- The only supported write path is the VBA object model (VBIDE) through Excel COM automation. It requires: Windows, desktop Excel, and "Trust access to the VBA project object model" enabled.
- Reading does **not** require Excel. `oletools` (olevba) parses the file directly on any OS.
- `.xlsx` / `.xltx` cannot hold VBA. Saving VBA into them silently discards it.

## Choose the path

| Goal | Tool |
|---|---|
| Read, review, search, or summarize macros; no Excel available | `scripts/extract_vba.py` (read-only) |
| Export for an edit round trip | `scripts/Export-Vba.ps1` |
| Write changes back | `Export-Vba.ps1` → edit files → `Import-Vba.ps1` |

If the task is read-only, use `extract_vba.py` even on Windows. It cannot damage the workbook and does not depend on Trust Center settings.

## Workflow (write path)

1. **Check the file type.** If the extension is `.xlsx` or `.xltx`, stop. Ask the user whether they want a `.xlsm` copy. Never convert silently.
2. **Check that the workbook is closed.** The scripts start a separate hidden Excel instance. If the user has the file open, the script opens it read-only and the save fails. Ask the user to close it.
3. **Export:**
   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File <skill>/scripts/Export-Vba.ps1 -WorkbookPath "<workbook>" -OutDir "<dir>"
   ```
   Output: one file per component, plus `vba-manifest.json` (name, type, file).
4. **Edit** only the files the task needs. Follow the editing rules below.
5. **Import:**
   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File <skill>/scripts/Import-Vba.ps1 -WorkbookPath "<workbook>" -SourceDir "<dir>" -Only Module1,Sheet1
   ```
   Always pass `-Only` with the components you changed. The script copies the workbook to a timestamped backup before it opens it and prints the backup path.
6. **Verify:** export again to a new folder and diff it against your edited folder. The import script also reads back each component and prints `WARN` on any mismatch.
7. **Report to the user:** backup path, components changed, any warnings, and that they must run *Debug > Compile VBAProject* and test. The scripts disable macros, so nothing was executed or compiled.

## Exported file types

| Extension | Component | Notes |
|---|---|---|
| `.bas` | Standard module | Header: `Attribute VB_Name = "..."` |
| `.cls` | Class module | Header: `VERSION 1.0 CLASS` / `BEGIN`…`END` / `Attribute VB_*` lines |
| `.frm` + `.frx` | UserForm | `.frx` is binary designer data |
| `.doccls` | Document module (`ThisWorkbook`, sheet modules) | Code only, no header. This extension is a convention of these scripts, not an Excel format |

## Editing rules

- **Encoding: Windows-1252 (ANSI). No UTF-8, no BOM.** The import script rejects files with a BOM. If your edit tool can only write UTF-8, keep the file pure ASCII and build non-ASCII characters in code with `ChrW()`.
- **Line endings: CRLF.** The import script normalizes, but do not mix.
- **Do not edit header lines** (the `VERSION`/`BEGIN`…`END` block and the leading `Attribute VB_*` lines). For an existing component the import replaces only the code body, so header edits are ignored. Renaming a component is not supported by these scripts; tell the user.
- **UserForms: code only.** Do not edit the `.frx` file or the `Begin`…`End` designer block in the `.frm`. Layout changes are out of scope.
- **Member attributes** (for example `Attribute Item.VB_UserMemId = 0`) are invisible in the VBA editor but load-bearing. Keep them on the line directly after the procedure signature. If a `.bas`/`.cls` contains them, the import removes and re-imports the whole component to keep them.
- VBA limits: 1023 characters per physical line, 24 line continuations per statement. Keep `Option Explicit` where it exists.
- **New components:** create the file with a correct header (templates below). New document modules cannot be created; they exist only when a sheet exists.
- The scripts never delete components. If a task needs a component removed, tell the user.

### New standard module template (`NewModule.bas`)
```
Attribute VB_Name = "NewModule"
Option Explicit

Public Sub Example()
End Sub
```

### New class module template (`NewClass.cls`)
```
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NewClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
```

## Stop conditions: report, do not work around

Each script error starts with a code. On these codes, stop and tell the user:

- `VBA_OBJECT_MODEL_BLOCKED`: programmatic access is off. **Do not change the `AccessVBOM` registry value or Trust Center settings.** On managed machines this is a security control, often set by Group Policy. Offer the read-only path instead.
- `VBA_PROJECT_LOCKED`: the project has a password. Do not attempt to remove or bypass it.
- `NO_VBA_CONTAINER`: the file type cannot hold VBA.
- `ENCODING_ERROR`: re-save the file as ANSI/ASCII without a BOM, then retry.
- `DOCUMENT_MODULE_MISSING`: the target sheet does not exist in the workbook.
- `WORKBOOK_IN_USE`: another process has the file open. Ask the user to close it.
- `SOURCE_MISSING`: a name in `-Only` has no source file. Check the spelling.
- `IMPORT_NAME_MISMATCH`: the `Attribute VB_Name` header does not match the file name. On a re-import, tell the user to restore the backup.
- `MISSING_DEPENDENCY` (extract_vba.py): `oletools` is not installed. Ask before installing packages on a managed machine.
- PowerShell refuses `-ExecutionPolicy Bypass`: report it. Do not change machine execution policy.

Warn the user (do not stop) on:

- `SIGNED_PROJECT`: the VBA project has a digital signature. Any change invalidates it. If policy allows only signed macros, the macros will not run until the project is re-signed.

## Gotchas

- The VBA editor normalizes identifier case across the whole project. A re-export can show case-only changes in modules you did not touch. This is expected; mention it, do not "fix" it.
- If a script is killed mid-run, a hidden `EXCEL.EXE` may remain and hold a lock on the file. Tell the user. Do not kill Excel processes yourself; one of them may be the user's own session.
- `extract_vba.py` output keeps the `Attribute` headers and names document modules `.cls`. Treat it as read-only reference. Never feed it to `Import-Vba.ps1`.
- Files downloaded from the internet or email may carry Mark of the Web. Excel blocks their macros at run time; editing still works. Mention it if the user expects the macros to run.
