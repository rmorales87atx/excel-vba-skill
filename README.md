# excel-vba-skill

Kilo Code skill that lets an agent extract and update VBA code in Excel workbooks.

## Install

Copy the `excel-vba` folder into your skills directory:

- Project: `.kilo/skills/excel-vba/` (older Kilo versions: `.kilocode/skills/excel-vba/`)
- Global: `~/.kilo/skills/excel-vba/` (older Kilo versions: `~/.kilocode/skills/excel-vba/`)

## Requirements

- Read-only extraction (`scripts/extract_vba.py`): Python 3 and `oletools`. Any OS.
- Export/import (`scripts/*.ps1`): Windows, desktop Excel, and "Trust access to the VBA project object model" enabled.

See `excel-vba/SKILL.md` for the workflow and stop conditions.
