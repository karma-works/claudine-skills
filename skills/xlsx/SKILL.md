---
name: xlsx-manipulation
description: Read, edit, recalculate, and validate Excel files (.xlsx) without requiring Microsoft Excel. Use when working with spreadsheets for tasks like extracting data, updating cell values or formulas, recalculating formulas programmatically, checking for errors (#DIV/0!, #REF!, etc.), or batch processing Excel files.
---

# XLSX Manipulation

Headless Excel file operations using Python (openpyxl and formulas libraries).

## Prerequisites

Python 3.8+ with dependencies:
```bash
pip install openpyxl 'formulas[excel]' defusedxml
```

## Tools

All tools output JSON. The xlsx_engine.py script location is relative to this skill directory.

### xlsx_read

Extract values and formulas from a worksheet.

```bash
python3 scripts/xlsx_engine.py read --path FILE_PATH [--sheet SHEET_NAME]
```

- `FILE_PATH`: Path to .xlsx file
- `SHEET_NAME`: Optional, defaults to active sheet

Returns: JSON with sheet name and cells array (coordinates, values, formulas)

### xlsx_edit

Update cell values or formulas.

```bash
python3 scripts/xlsx_engine.py edit --path FILE_PATH --sheet SHEET_NAME --updates JSON_ARRAY
```

- `FILE_PATH`: Path to .xlsx file
- `SHEET_NAME`: Sheet to modify
- `JSON_ARRAY`: Array of `{"cell": "A1", "value": 100}` or `{"cell": "B1", "value": "=A1*2"}`

Returns: JSON with updated cells list

### xlsx_recalculate

Recalculate all formulas.

```bash
python3 scripts/xlsx_engine.py recalculate --path FILE_PATH
```

Use after editing to compute formula results. Some advanced Excel functions unsupported.

Returns: JSON with recalculated values

### xlsx_check_errors

Scan for Excel errors.

```bash
python3 scripts/xlsx_engine.py check_errors --path FILE_PATH
```

Detects: #REF!, #DIV/0!, #NAME?, #VALUE!, #N/A, #NULL!, #NUM!

Returns: JSON with errors list (sheet, cell, error type)

## Typical Workflow

For "Update cell B2 to 500 and show me C2's new value":

1. `xlsx_read` - Check current state
2. `xlsx_edit` - Update B2
3. `xlsx_recalculate` - Compute formulas
4. `xlsx_read` - Get C2's new value

## Limitations

- No VBA macros
- No .xls (binary format)
- No password-protected files
- Limited formula library support
