---
name: pptx-manipulation
description: Read, edit, and manipulate PowerPoint files (.pptx) without requiring Microsoft PowerPoint. Use when working with presentations for tasks like extracting slide content, updating text, adding slides, reading slide notes, or batch processing PowerPoint files.
---

# PPTX Manipulation

Headless PowerPoint file operations using Python (python-pptx library).

## Prerequisites

Python 3.8+ with dependencies:
```bash
pip install python-pptx Pillow
```

## Tools

All tools output JSON. The pptx_engine.py script location is relative to this skill directory.

### pptx_read

Extract content from all slides or a specific slide.

```bash
python3 scripts/pptx_engine.py read --path FILE_PATH [--slide SLIDE_NUMBER]
```

- `FILE_PATH`: Path to .pptx file
- `SLIDE_NUMBER`: Optional, specific slide number (1-based index)

Returns: JSON with presentation info, slides, shapes, text content, and notes

### pptx_create

Create a new PowerPoint presentation.

```bash
python3 scripts/pptx_engine.py create --path FILE_PATH [--width WIDTH] [--height HEIGHT]
```

- `FILE_PATH`: Path for the new .pptx file
- `WIDTH`: Optional slide width in inches (default: 10)
- `HEIGHT`: Optional slide height in inches (default: 7.5)

Returns: JSON with creation status and slide dimensions

### pptx_edit

Update text content in specific shapes on slides.

```bash
python3 scripts/pptx_engine.py edit --path FILE_PATH --updates JSON_ARRAY
```

- `FILE_PATH`: Path to .pptx file
- `JSON_ARRAY`: Array of update objects with various identification methods:
  - `{"slide": 1, "title": "New Title"}` - Updates title shape (most reliable)
  - `{"slide": 1, "shape_name": "Title 1", "text": "New text"}` - Finds by shape name (case-insensitive substring match)
  - `{"slide": 1, "shape_text": "Old text", "text": "New text"}` - Finds by existing text content (case-insensitive substring match)
  - `{"slide": 1, "shape": 0, "text": "New text"}` - By index (less reliable, indices can shift)
  - `{"slide": 2, "notes": "Speaker notes"}` - Updates slide notes

**Recommendation**: Use `shape_name` or `shape_text` instead of `shape` (index) for more reliable editing.

Returns: JSON with updated shapes list

### pptx_add_slide

Add a new slide to the presentation.

```bash
python3 scripts/pptx_engine.py add_slide --path FILE_PATH --layout LAYOUT_INDEX [--title TITLE] [--content CONTENT]
```

- `FILE_PATH`: Path to .pptx file
- `LAYOUT_INDEX`: Slide layout index (0-based, typically 0=title, 1=title+content, 6=blank)
- `TITLE`: Optional title text
- `CONTENT`: Optional content text

Returns: JSON with new slide number and details

### pptx_get_layouts

List all available slide layouts in the presentation.

```bash
python3 scripts/pptx_engine.py get_layouts --path FILE_PATH
```

Returns: JSON with all available slide layouts and their indices

### pptx_extract_images

Extract all images from the presentation.

```bash
python3 scripts/pptx_engine.py extract_images --path FILE_PATH --output OUTPUT_DIR
```

- `FILE_PATH`: Path to .pptx file
- `OUTPUT_DIR`: Directory to save extracted images

Returns: JSON with extracted image paths

## Typical Workflow

For "Update slide 2 title to 'New Title' and add a new slide":

1. `pptx_read` - Check current presentation structure
2. `pptx_edit` - Update slide 2 title
3. `pptx_get_layouts` - Check available layouts
4. `pptx_add_slide` - Add a new slide with desired layout
5. `pptx_read` - Verify changes

## Limitations

- No VBA macros
- No .ppt (binary format)
- No password-protected files
- Limited support for complex animations
- Cannot modify embedded charts or tables directly
