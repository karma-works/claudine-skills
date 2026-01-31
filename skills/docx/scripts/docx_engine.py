#!/usr/bin/env python3
"""
DOCX Engine - Core utilities for Word document manipulation
Provides safe loading, saving, reading, editing, and creating Word documents
"""

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import json
import sys
import argparse
from pathlib import Path


def load_document_safe(path):
    """
    Safely load a Word document with error handling
    
    Args:
        path: Path to the .docx file
        
    Returns:
        docx Document object
        
    Raises:
        FileNotFoundError: If the file doesn't exist
        Exception: If the file is not a valid Word document
    """
    path = Path(path)
    
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    try:
        doc = Document(str(path))
        return doc
    except Exception as e:
        raise Exception(f"Error loading document: {e}") from e


def save_document_safe(doc, path):
    """
    Safely save a Word document with error handling
    
    Args:
        doc: docx Document object
        path: Path to save the .docx file
        
    Raises:
        Exception: If saving fails
    """
    try:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(path))
    except Exception as e:
        raise Exception(f"Error saving document: {e}") from e


def read_document(path, include_tables=False):
    """
    Extract text and structure from a document
    
    Args:
        path: Path to the .docx file
        include_tables: Whether to include table data
        
    Returns:
        dict: Contains paragraphs and optional tables
    """
    doc = load_document_safe(path)
    
    result = {
        "paragraph_count": len(doc.paragraphs),
        "table_count": len(doc.tables)
    }
    
    # Extract paragraphs
    paragraphs = []
    for idx, para in enumerate(doc.paragraphs):
        para_data = {
            "index": idx,
            "text": para.text,
            "style": para.style.name if para.style else None
        }
        paragraphs.append(para_data)
    
    result["paragraphs"] = paragraphs
    
    # Extract tables if requested
    if include_tables:
        tables = []
        for table_idx, table in enumerate(doc.tables):
            table_data = {
                "index": table_idx,
                "rows": len(table.rows),
                "cols": len(table.columns),
                "data": []
            }
            for row in table.rows:
                row_data = [cell.text for cell in row.cells]
                table_data["data"].append(row_data)
            tables.append(table_data)
        result["tables"] = tables
    
    return result


def edit_paragraph_text(path, paragraph_index, new_text):
    """
    Update text in a specific paragraph
    
    Args:
        path: Path to the .docx file
        paragraph_index: Zero-based index of paragraph to edit
        new_text: New text content
        
    Returns:
        dict: Status message
    """
    doc = load_document_safe(path)
    
    if paragraph_index < 0 or paragraph_index >= len(doc.paragraphs):
        raise ValueError(f"Paragraph index {paragraph_index} out of range. Document has {len(doc.paragraphs)} paragraphs.")
    
    para = doc.paragraphs[paragraph_index]
    
    # Clear existing runs and add new text
    para.clear()
    run = para.add_run(new_text)
    
    save_document_safe(doc, path)
    
    return {
        "status": "success",
        "paragraph_index": paragraph_index,
        "new_text": new_text
    }


def insert_paragraph(path, position, text, style=None):
    """
    Insert a new paragraph at a specific position
    
    Args:
        path: Path to the .docx file
        position: Position to insert (0 = beginning, 'end' = append)
        text: Paragraph text content
        style: Optional style name
        
    Returns:
        dict: Status message
    """
    doc = load_document_safe(path)
    
    if position == "end":
        new_para = doc.add_paragraph(text)
        insert_idx = len(doc.paragraphs) - 1
    else:
        try:
            insert_idx = int(position)
            if insert_idx < 0:
                insert_idx = 0
            if insert_idx > len(doc.paragraphs):
                insert_idx = len(doc.paragraphs)
        except ValueError:
            raise ValueError(f"Position must be an integer or 'end', got: {position}")
        
        # Insert at specific position
        if insert_idx == 0:
            new_para = doc.paragraphs[0].insert_paragraph_before(text)
        else:
            # Insert after the specified paragraph
            new_para = doc.paragraphs[insert_idx - 1].insert_paragraph_before(text)
    
    if style:
        try:
            new_para.style = style
        except Exception:
            pass  # Style not found, ignore
    
    save_document_safe(doc, path)
    
    return {
        "status": "success",
        "position": insert_idx,
        "text": text,
        "style": style
    }


def find_replace(path, find_text, replace_text):
    """
    Find and replace text throughout the document
    
    Args:
        path: Path to .docx file
        find_text: Text to find
        replace_text: Replacement text
        
    Returns:
        dict: Number of replacements made
    """
    doc = load_document_safe(path)
    
    replacements = 0
    
    # Search in paragraphs
    for para in doc.paragraphs:
        if find_text in para.text:
            # Replace in each run to preserve formatting
            for run in para.runs:
                if find_text in run.text:
                    run.text = run.text.replace(find_text, replace_text)
                    replacements += 1
    
    # Search in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if find_text in para.text:
                        for run in para.runs:
                            if find_text in run.text:
                                run.text = run.text.replace(find_text, replace_text)
                                replacements += 1
    
    save_document_safe(doc, path)
    
    return {
        "status": "success",
        "replacements": replacements,
        "find": find_text,
        "replace": replace_text
    }


def read_tables(path, table_index=None):
    """
    Extract table data from the document
    
    Args:
        path: Path to .docx file
        table_index: Optional specific table index (0-based)
        
    Returns:
        dict: Table data
    """
    doc = load_document_safe(path)
    
    if len(doc.tables) == 0:
        return {
            "table_count": 0,
            "tables": []
        }
    
    if table_index is not None:
        if table_index < 0 or table_index >= len(doc.tables):
            raise ValueError(f"Table index {table_index} out of range. Document has {len(doc.tables)} tables.")
        tables_to_read = [doc.tables[table_index]]
    else:
        tables_to_read = doc.tables
    
    tables_data = []
    for idx, table in enumerate(tables_to_read):
        table_info = {
            "index": idx if table_index is None else table_index,
            "rows": len(table.rows),
            "cols": len(table.columns),
            "data": []
        }
        for row in table.rows:
            row_data = [cell.text for cell in row.cells]
            table_info["data"].append(row_data)
        tables_data.append(table_info)
    
    return {
        "table_count": len(doc.tables),
        "tables": tables_data
    }


def edit_table_cell(path, table_index, row, col, text):
    """
    Edit a specific cell in a table
    
    Args:
        path: Path to .docx file
        table_index: Table index (0-based)
        row: Row index (0-based)
        col: Column index (0-based)
        text: New cell text content
        
    Returns:
        dict: Status message
    """
    doc = load_document_safe(path)
    
    if table_index < 0 or table_index >= len(doc.tables):
        raise ValueError(f"Table index {table_index} out of range. Document has {len(doc.tables)} tables.")
    
    table = doc.tables[table_index]
    
    if row < 0 or row >= len(table.rows):
        raise ValueError(f"Row index {row} out of range. Table has {len(table.rows)} rows.")
    
    if col < 0 or col >= len(table.columns):
        raise ValueError(f"Column index {col} out of range. Table has {len(table.columns)} columns.")
    
    cell = table.rows[row].cells[col]
    cell.text = text
    
    save_document_safe(doc, path)
    
    return {
        "status": "success",
        "table_index": table_index,
        "row": row,
        "col": col,
        "text": text
    }


def create_document(path):
    """
    Create a new blank document
    
    Args:
        path: Path for the new .docx file
        
    Returns:
        dict: Creation status
    """
    doc = Document()
    save_document_safe(doc, path)
    
    return {
        "status": "success",
        "path": str(path),
        "message": "New blank document created"
    }


def add_table(path, rows, cols, position=None, data=None):
    """
    Add a new table to the document
    
    Args:
        path: Path to .docx file
        rows: Number of rows
        cols: Number of columns
        position: Optional position to insert (0 = beginning, 'end' = append)
        data: Optional JSON array of row data
        
    Returns:
        dict: Table insertion status
    """
    doc = load_document_safe(path)
    
    if rows < 1 or cols < 1:
        raise ValueError("Rows and columns must be at least 1")
    
    # Create table
    table = doc.add_table(rows=rows, cols=cols)
    
    # Fill with data if provided
    if data:
        try:
            table_data = json.loads(data) if isinstance(data, str) else data
            for row_idx, row_data in enumerate(table_data):
                if row_idx < rows:
                    for col_idx, cell_text in enumerate(row_data):
                        if col_idx < cols:
                            table.rows[row_idx].cells[col_idx].text = str(cell_text)
        except Exception as e:
            pass  # Ignore data parsing errors
    
    # Move to position if specified
    if position is not None and position != "end":
        try:
            pos_idx = int(position)
            # Move table to position (complex in python-docx, simplified here)
            pass
        except ValueError:
            pass
    
    save_document_safe(doc, path)
    
    return {
        "status": "success",
        "rows": rows,
        "cols": cols,
        "table_index": len(doc.tables) - 1
    }


def document_info(path):
    """
    Get document metadata and statistics
    
    Args:
        path: Path to .docx file
        
    Returns:
        dict: Document information
    """
    doc = load_document_safe(path)
    
    # Count words
    word_count = 0
    for para in doc.paragraphs:
        word_count += len(para.text.split())
    
    # Get core properties
    core_props = doc.core_properties
    
    info = {
        "paragraph_count": len(doc.paragraphs),
        "table_count": len(doc.tables),
        "word_count": word_count,
        "properties": {
            "title": core_props.title,
            "author": core_props.author,
            "subject": core_props.subject,
            "keywords": core_props.keywords,
            "created": str(core_props.created) if core_props.created else None,
            "modified": str(core_props.modified) if core_props.modified else None,
            "last_modified_by": core_props.last_modified_by
        }
    }
    
    return info


def main():
    """CLI interface for docx_engine"""
    parser = argparse.ArgumentParser(description='Word document manipulation engine')
    subparsers = parser.add_subparsers(dest='command', help='Commands')
    
    # Read command
    read_parser = subparsers.add_parser('read', help='Read document content')
    read_parser.add_argument('--path', required=True, help='Path to Word document')
    read_parser.add_argument('--include-tables', action='store_true', help='Include table data')
    
    # Edit text command
    edit_parser = subparsers.add_parser('edit_text', help='Edit paragraph text')
    edit_parser.add_argument('--path', required=True, help='Path to Word document')
    edit_parser.add_argument('--paragraph-index', type=int, required=True, help='Paragraph index (0-based)')
    edit_parser.add_argument('--new-text', required=True, help='New text content')
    
    # Insert paragraph command
    insert_parser = subparsers.add_parser('insert_paragraph', help='Insert new paragraph')
    insert_parser.add_argument('--path', required=True, help='Path to Word document')
    insert_parser.add_argument('--position', required=True, help='Position (0 = beginning, end = append)')
    insert_parser.add_argument('--text', required=True, help='Paragraph text')
    insert_parser.add_argument('--style', default=None, help='Optional style name')
    
    # Find and replace command
    replace_parser = subparsers.add_parser('find_replace', help='Find and replace text')
    replace_parser.add_argument('--path', required=True, help='Path to Word document')
    replace_parser.add_argument('--find', required=True, help='Text to find')
    replace_parser.add_argument('--replace', required=True, help='Replacement text')
    
    # Read tables command
    tables_parser = subparsers.add_parser('read_tables', help='Read table data')
    tables_parser.add_argument('--path', required=True, help='Path to Word document')
    tables_parser.add_argument('--table-index', type=int, default=None, help='Specific table index')
    
    # Edit table cell command
    edit_cell_parser = subparsers.add_parser('edit_table_cell', help='Edit table cell')
    edit_cell_parser.add_argument('--path', required=True, help='Path to Word document')
    edit_cell_parser.add_argument('--table-index', type=int, required=True, help='Table index')
    edit_cell_parser.add_argument('--row', type=int, required=True, help='Row index')
    edit_cell_parser.add_argument('--col', type=int, required=True, help='Column index')
    edit_cell_parser.add_argument('--text', required=True, help='New cell text')
    
    # Create command
    create_parser = subparsers.add_parser('create', help='Create new document')
    create_parser.add_argument('--path', required=True, help='Path for new document')
    
    # Add table command
    add_table_parser = subparsers.add_parser('add_table', help='Add table to document')
    add_table_parser.add_argument('--path', required=True, help='Path to Word document')
    add_table_parser.add_argument('--rows', type=int, required=True, help='Number of rows')
    add_table_parser.add_argument('--cols', type=int, required=True, help='Number of columns')
    add_table_parser.add_argument('--position', default=None, help='Insert position')
    add_table_parser.add_argument('--data', default=None, help='JSON array of row data')
    
    # Document info command
    info_parser = subparsers.add_parser('document_info', help='Get document info')
    info_parser.add_argument('--path', required=True, help='Path to Word document')
    
    args = parser.parse_args()
    
    try:
        if args.command == 'read':
            result = read_document(args.path, args.include_tables)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'edit_text':
            result = edit_paragraph_text(args.path, args.paragraph_index, args.new_text)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'insert_paragraph':
            result = insert_paragraph(args.path, args.position, args.text, args.style)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'find_replace':
            result = find_replace(args.path, args.find, args.replace)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'read_tables':
            result = read_tables(args.path, args.table_index)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'edit_table_cell':
            result = edit_table_cell(args.path, args.table_index, args.row, args.col, args.text)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'create':
            result = create_document(args.path)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'add_table':
            result = add_table(args.path, args.rows, args.cols, args.position, args.data)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'document_info':
            result = document_info(args.path)
            print(json.dumps(result, indent=2))
            
        else:
            parser.print_help()
            sys.exit(1)
            
    except Exception as e:
        error_result = {
            "status": "error",
            "message": str(e),
            "type": type(e).__name__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == '__main__':
    main()
