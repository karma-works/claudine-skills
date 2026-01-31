---
name: pdf-manipulation
description: Read, edit, create, and manipulate PDF documents without requiring Adobe Acrobat. Use when working with PDFs for tasks like extracting text, merging or splitting files, adding watermarks, filling forms, extracting images, or converting pages to images.
---

# PDF Manipulation

Headless PDF operations using Python (pypdf, pdfplumber, reportlab libraries) and system tools.

## Prerequisites

Python 3.8+ with dependencies:
```bash
pip install pypdf pdfplumber reportlab pillow
```

Optional system tools for advanced features:
```bash
# macOS
brew install poppler qpdf

# Linux
sudo apt-get install poppler-utils qpdf
```

## Tools

All tools output JSON. The pdf_engine.py script location is relative to this skill directory.

### pdf_info

Get PDF metadata and page count.

```bash
python3 scripts/pdf_engine.py info --path FILE_PATH
```

- `FILE_PATH`: Path to .pdf file

Returns: JSON with page count, metadata (title, author, subject, creator), and file size

### pdf_read

Extract text from PDF pages.

```bash
python3 scripts/pdf_engine.py read --path FILE_PATH [--pages PAGES] [--method METHOD]
```

- `FILE_PATH`: Path to .pdf file
- `PAGES`: Optional page range (e.g., "1", "1-5", "1,3,5"). Default: all pages
- `METHOD`: Extraction method: "pypdf" (default, faster) or "pdfplumber" (better for complex layouts)

Returns: JSON with extracted text per page

### pdf_extract_tables

Extract tables from PDF pages.

```bash
python3 scripts/pdf_engine.py extract_tables --path FILE_PATH [--pages PAGES]
```

- `FILE_PATH`: Path to .pdf file
- `PAGES`: Optional page range. Default: all pages

Returns: JSON with tables as arrays of rows

### pdf_merge

Merge multiple PDFs into one.

```bash
python3 scripts/pdf_engine.py merge --output OUTPUT_PATH --inputs INPUT1 INPUT2 [INPUT3 ...]
```

- `OUTPUT_PATH`: Path for merged PDF
- `INPUTS`: Two or more paths to PDF files to merge

Returns: JSON with merge status and page count

### pdf_split

Split a PDF into individual pages or ranges.

```bash
python3 scripts/pdf_engine.py split --path FILE_PATH --output-dir OUTPUT_DIR [--pages PAGES]
```

- `FILE_PATH`: Path to .pdf file
- `OUTPUT_DIR`: Directory for output files
- `PAGES`: Optional page range to split. Default: all pages (one file per page)

Returns: JSON with list of created files

### pdf_extract_pages

Extract specific pages to a new PDF.

```bash
python3 scripts/pdf_engine.py extract_pages --path FILE_PATH --output OUTPUT_PATH --pages PAGES
```

- `FILE_PATH`: Source PDF path
- `OUTPUT_PATH`: Path for new PDF with extracted pages
- `PAGES`: Page range to extract (e.g., "1-5", "1,3,5,7")

Returns: JSON with extraction status

### pdf_rotate

Rotate pages in a PDF.

```bash
python3 scripts/pdf_engine.py rotate --path FILE_PATH --degrees DEGREES [--pages PAGES] [--output OUTPUT_PATH]
```

- `FILE_PATH`: Path to .pdf file
- `DEGREES`: Rotation angle (90, 180, 270)
- `PAGES`: Optional page range. Default: all pages
- `OUTPUT_PATH`: Optional output path. Default: overwrites input

Returns: JSON with rotation status

### pdf_add_watermark

Add text watermark to PDF pages.

```bash
python3 scripts/pdf_engine.py add_watermark --path FILE_PATH --text TEXT [--output OUTPUT_PATH] [--opacity OPACITY] [--position POSITION]
```

- `FILE_PATH`: Path to .pdf file
- `TEXT`: Watermark text
- `OUTPUT_PATH`: Optional output path. Default: overwrites input
- `OPACITY`: Opacity 0.0-1.0 (default: 0.3)
- `POSITION`: Position: "center", "diagonal", "bottom" (default: "diagonal")

Returns: JSON with watermark status

### pdf_extract_images

Extract images from PDF pages.

```bash
python3 scripts/pdf_engine.py extract_images --path FILE_PATH --output-dir OUTPUT_DIR [--pages PAGES]
```

- `FILE_PATH`: Path to .pdf file
- `OUTPUT_DIR`: Directory for extracted images
- `PAGES`: Optional page range. Default: all pages

Returns: JSON with list of extracted image files

### pdf_create

Create a new PDF from text or images.

```bash
python3 scripts/pdf_engine.py create --output OUTPUT_PATH [--text TEXT] [--images IMAGE1 IMAGE2 ...]
```

- `OUTPUT_PATH`: Path for new PDF
- `TEXT`: Optional text content for PDF
- `IMAGES`: Optional list of image paths to include

Returns: JSON with creation status

### pdf_encrypt

Add password protection to a PDF.

```bash
python3 scripts/pdf_engine.py encrypt --path FILE_PATH --password PASSWORD [--output OUTPUT_PATH]
```

- `FILE_PATH`: Path to .pdf file
- `PASSWORD`: Password for the PDF
- `OUTPUT_PATH`: Optional output path. Default: overwrites input

Returns: JSON with encryption status

### pdf_decrypt

Remove password protection from a PDF.

```bash
python3 scripts/pdf_engine.py decrypt --path FILE_PATH --password PASSWORD [--output OUTPUT_PATH]
```

- `FILE_PATH`: Path to encrypted .pdf file
- `PASSWORD`: Current password
- `OUTPUT_PATH`: Optional output path. Default: overwrites input

Returns: JSON with decryption status

### pdf_to_images

Convert PDF pages to images.

```bash
python3 scripts/pdf_engine.py to_images --path FILE_PATH --output-dir OUTPUT_DIR [--pages PAGES] [--format FORMAT] [--dpi DPI]
```

- `FILE_PATH`: Path to .pdf file
- `OUTPUT_DIR`: Directory for output images
- `PAGES`: Optional page range. Default: all pages
- `FORMAT`: Image format: "png", "jpg" (default: "png")
- `DPI`: Resolution (default: 150)

Returns: JSON with list of created image files

**Note**: Requires `pdftoppm` from poppler-utils for best quality.

## Typical Workflows

For "Extract text from a PDF":

1. `pdf_info` - Check page count and metadata
2. `pdf_read` - Extract text from desired pages

For "Merge multiple PDFs":

1. `pdf_merge` - Combine PDFs in order

For "Add watermark to all pages":

1. `pdf_add_watermark` - Apply watermark with desired text

For "Extract specific pages as a new document":

1. `pdf_info` - Check total page count
2. `pdf_extract_pages` - Extract desired pages

## Limitations

- No OCR support (scanned PDFs with images only won't extract text)
- No form filling (AcroForm support)
- No digital signature support
- Limited table extraction accuracy for complex layouts
- Watermarks are text-only (no image watermarks)
