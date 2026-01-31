---
name: docx-manipulation
description: Read, edit, create, and manipulate Microsoft Word documents (.docx) without requiring Microsoft Word. Use when working with Word documents for tasks like extracting text, updating content, formatting text, working with tables, or batch processing documents.
---

# DOCX Manipulation

Headless Word document operations using Python (python-docx library).

## Prerequisites

Python 3.8+ with dependencies:
```bash
pip install python-docx
```

## Tools

All tools output JSON. The docx_engine.py script location is relative to this skill directory.

### docx_read

Extract text, paragraphs, and tables from a document.

```bash
python3 scripts/docx_engine.py read --path FILE_PATH [--include-tables]
```

- `FILE_PATH`: Path to .docx file
- `--include-tables`: Optional, include table data in output

Returns: JSON with document structure, paragraphs, and optional tables

### docx_edit_text

Update text in specific paragraphs by index.

```bash
python3 scripts/docx_engine.py edit_text --path FILE_PATH --paragraph-index INDEX --new-text TEXT
```

- `FILE_PATH`: Path to .docx file
- `INDEX`: Zero-based paragraph index to edit
- `TEXT`: New text content for the paragraph

Returns: JSON with update status and affected paragraph

### docx_insert_paragraph

Insert a new paragraph at a specific position.

```bash
python3 scripts/docx_engine.py insert_paragraph --path FILE_PATH --position POSITION --text TEXT [--style STYLE]
```

- `FILE_PATH`: Path to .docx file
- `POSITION`: Position to insert (0 = beginning, end = append)
- `TEXT`: Paragraph text content
- `STYLE`: Optional style name (e.g., 'Heading 1', 'Normal')

Returns: JSON with insert status and position

### docx_find_replace

Find and replace text across the document.

```bash
python3 scripts/docx_engine.py find_replace --path FILE_PATH --find TEXT --replace TEXT
```

- `FILE_PATH`: Path to .docx file
- `FIND`: Text to find
- `REPLACE`: Replacement text

Returns: JSON with number of replacements made

### docx_read_tables

Extract table data from the document.

```bash
python3 scripts/docx_engine.py read_tables --path FILE_PATH [--table-index INDEX]
```

- `FILE_PATH`: Path to .docx file
- `INDEX`: Optional, specific table index (0-based). If omitted, returns all tables.

Returns: JSON with table data as array of rows

### docx_edit_table_cell

Edit a specific cell in a table.

```bash
python3 scripts/docx_engine.py edit_table_cell --path FILE_PATH --table-index INDEX --row ROW --col COL --text TEXT
```

- `FILE_PATH`: Path to .docx file
- `INDEX`: Table index (0-based)
- `ROW`: Row index (0-based)
- `COL`: Column index (0-based)
- `TEXT`: New cell text content

Returns: JSON with update status

### docx_create

Create a new blank document.

```bash
python3 scripts/docx_engine.py create --path FILE_PATH
```

- `FILE_PATH`: Path for the new .docx file

Returns: JSON with creation status

### docx_add_table

Add a new table to the document.

```bash
python3 scripts/docx_engine.py add_table --path FILE_PATH --rows ROWS --cols COLS [--position POSITION] [--data JSON_ARRAY]
```

- `FILE_PATH`: Path to .docx file
- `ROWS`: Number of rows
- `COLS`: Number of columns
- `POSITION`: Optional position to insert (0 = beginning, end = append)
- `DATA`: Optional JSON array of row data

Returns: JSON with table insertion status

### docx_document_info

Get document metadata and statistics.

```bash
python3 scripts/docx_engine.py document_info --path FILE_PATH
```

- `FILE_PATH`: Path to .docx file

Returns: JSON with paragraph count, table count, word count, and core properties

## Typical Workflow

For "Update the title and add a new paragraph":

1. `docx_read` - Check current document structure
2. `docx_edit_text` - Update the title (paragraph 0)
3. `docx_insert_paragraph` - Add new content
4. `docx_read` - Verify changes

For "Find and replace a company name throughout the document":

1. `docx_find_replace` - Replace all occurrences
2. `docx_read` - Verify replacements

## Limitations

- No macro support (VBA)
- No .doc (binary format) support - only .docx
- No password-protected files
- Limited style support for complex formatting
- Track changes not supported
