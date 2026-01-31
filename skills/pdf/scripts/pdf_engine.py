#!/usr/bin/env python3
"""
PDF Engine - Core utilities for PDF file manipulation
Provides reading, editing, merging, splitting, watermarking, and more
"""

import json
import sys
import argparse
import os
from pathlib import Path
from io import BytesIO


def get_pdf_info(path):
    """Get PDF metadata and page count"""
    from pypdf import PdfReader
    
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    reader = PdfReader(str(path))
    metadata = reader.metadata or {}
    
    return {
        "path": str(path),
        "page_count": len(reader.pages),
        "file_size_bytes": path.stat().st_size,
        "metadata": {
            "title": metadata.get("/Title", None),
            "author": metadata.get("/Author", None),
            "subject": metadata.get("/Subject", None),
            "creator": metadata.get("/Creator", None),
            "producer": metadata.get("/Producer", None),
            "creation_date": str(metadata.get("/CreationDate", None)),
        },
        "is_encrypted": reader.is_encrypted
    }


def parse_page_range(page_range_str, total_pages):
    """Parse page range string to list of page indices (0-based)"""
    if not page_range_str:
        return list(range(total_pages))
    
    pages = set()
    parts = page_range_str.replace(" ", "").split(",")
    
    for part in parts:
        if "-" in part:
            start, end = part.split("-", 1)
            start = int(start) - 1  # Convert to 0-based
            end = int(end) - 1
            pages.update(range(start, end + 1))
        else:
            pages.add(int(part) - 1)
    
    # Filter valid pages and sort
    valid_pages = sorted([p for p in pages if 0 <= p < total_pages])
    return valid_pages


def read_pdf_text(path, pages=None, method="pypdf"):
    """Extract text from PDF pages"""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    if method == "pdfplumber":
        import pdfplumber
        
        result = {"pages": []}
        with pdfplumber.open(str(path)) as pdf:
            page_indices = parse_page_range(pages, len(pdf.pages))
            
            for idx in page_indices:
                page = pdf.pages[idx]
                text = page.extract_text() or ""
                result["pages"].append({
                    "page": idx + 1,
                    "text": text,
                    "char_count": len(text)
                })
        
        return result
    else:
        # Default: pypdf
        from pypdf import PdfReader
        
        reader = PdfReader(str(path))
        page_indices = parse_page_range(pages, len(reader.pages))
        
        result = {"pages": []}
        for idx in page_indices:
            page = reader.pages[idx]
            text = page.extract_text() or ""
            result["pages"].append({
                "page": idx + 1,
                "text": text,
                "char_count": len(text)
            })
        
        return result


def extract_tables(path, pages=None):
    """Extract tables from PDF pages using pdfplumber"""
    import pdfplumber
    
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    result = {"tables": []}
    
    with pdfplumber.open(str(path)) as pdf:
        page_indices = parse_page_range(pages, len(pdf.pages))
        
        for idx in page_indices:
            page = pdf.pages[idx]
            tables = page.extract_tables()
            
            for table_idx, table in enumerate(tables):
                result["tables"].append({
                    "page": idx + 1,
                    "table_index": table_idx,
                    "rows": table
                })
    
    return result


def merge_pdfs(output_path, input_paths):
    """Merge multiple PDFs into one"""
    from pypdf import PdfWriter
    
    if len(input_paths) < 2:
        raise ValueError("At least 2 input files required for merge")
    
    writer = PdfWriter()
    total_pages = 0
    
    for input_path in input_paths:
        path = Path(input_path)
        if not path.exists():
            raise FileNotFoundError(f"File not found: {path}")
        
        writer.append(str(path))
        # Count pages from this file
        from pypdf import PdfReader
        reader = PdfReader(str(path))
        total_pages += len(reader.pages)
    
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, "wb") as f:
        writer.write(f)
    
    return {
        "status": "success",
        "output_path": str(output_path),
        "total_pages": total_pages,
        "files_merged": len(input_paths)
    }


def split_pdf(path, output_dir, pages=None):
    """Split PDF into individual pages or ranges"""
    from pypdf import PdfReader, PdfWriter
    
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    reader = PdfReader(str(path))
    page_indices = parse_page_range(pages, len(reader.pages))
    
    created_files = []
    base_name = path.stem
    
    for idx in page_indices:
        writer = PdfWriter()
        writer.add_page(reader.pages[idx])
        
        output_file = output_dir / f"{base_name}_page_{idx + 1}.pdf"
        with open(output_file, "wb") as f:
            writer.write(f)
        
        created_files.append(str(output_file))
    
    return {
        "status": "success",
        "created_files": created_files,
        "file_count": len(created_files)
    }


def extract_pages(path, output_path, pages):
    """Extract specific pages to a new PDF"""
    from pypdf import PdfReader, PdfWriter
    
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    reader = PdfReader(str(path))
    page_indices = parse_page_range(pages, len(reader.pages))
    
    if not page_indices:
        raise ValueError("No valid pages specified")
    
    writer = PdfWriter()
    for idx in page_indices:
        writer.add_page(reader.pages[idx])
    
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output_path, "wb") as f:
        writer.write(f)
    
    return {
        "status": "success",
        "output_path": str(output_path),
        "pages_extracted": len(page_indices),
        "page_numbers": [i + 1 for i in page_indices]
    }


def rotate_pdf(path, degrees, pages=None, output_path=None):
    """Rotate pages in a PDF"""
    from pypdf import PdfReader, PdfWriter
    
    if degrees not in [90, 180, 270]:
        raise ValueError("Degrees must be 90, 180, or 270")
    
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    reader = PdfReader(str(path))
    page_indices = parse_page_range(pages, len(reader.pages))
    
    writer = PdfWriter()
    rotated_pages = []
    
    for idx, page in enumerate(reader.pages):
        if idx in page_indices:
            page.rotate(degrees)
            rotated_pages.append(idx + 1)
        writer.add_page(page)
    
    output = Path(output_path) if output_path else path
    output.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output, "wb") as f:
        writer.write(f)
    
    return {
        "status": "success",
        "output_path": str(output),
        "degrees": degrees,
        "rotated_pages": rotated_pages
    }


def add_watermark(path, text, output_path=None, opacity=0.3, position="diagonal"):
    """Add text watermark to PDF pages"""
    from pypdf import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.colors import Color
    
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    reader = PdfReader(str(path))
    writer = PdfWriter()
    
    for page in reader.pages:
        # Get page dimensions
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)
        
        # Create watermark
        watermark_buffer = BytesIO()
        c = canvas.Canvas(watermark_buffer, pagesize=(page_width, page_height))
        
        # Set watermark style
        c.setFillColor(Color(0.5, 0.5, 0.5, alpha=opacity))
        c.setFont("Helvetica", 40)
        
        if position == "diagonal":
            c.saveState()
            c.translate(page_width / 2, page_height / 2)
            c.rotate(45)
            c.drawCentredString(0, 0, text)
            c.restoreState()
        elif position == "center":
            c.drawCentredString(page_width / 2, page_height / 2, text)
        elif position == "bottom":
            c.setFont("Helvetica", 20)
            c.drawCentredString(page_width / 2, 30, text)
        
        c.save()
        watermark_buffer.seek(0)
        
        # Merge watermark with page
        from pypdf import PdfReader as WatermarkReader
        watermark_pdf = WatermarkReader(watermark_buffer)
        page.merge_page(watermark_pdf.pages[0])
        writer.add_page(page)
    
    output = Path(output_path) if output_path else path
    output.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output, "wb") as f:
        writer.write(f)
    
    return {
        "status": "success",
        "output_path": str(output),
        "watermark_text": text,
        "pages_watermarked": len(reader.pages)
    }


def extract_images(path, output_dir, pages=None):
    """Extract images from PDF pages"""
    from pypdf import PdfReader
    from PIL import Image
    
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    reader = PdfReader(str(path))
    page_indices = parse_page_range(pages, len(reader.pages))
    
    extracted_files = []
    image_count = 0
    
    for idx in page_indices:
        page = reader.pages[idx]
        
        if "/XObject" not in page["/Resources"]:
            continue
        
        x_objects = page["/Resources"]["/XObject"].get_object()
        
        for obj_name in x_objects:
            obj = x_objects[obj_name]
            
            if obj["/Subtype"] == "/Image":
                image_count += 1
                
                # Get image data
                width = obj["/Width"]
                height = obj["/Height"]
                
                try:
                    data = obj.get_data()
                    
                    # Determine format and save
                    if "/Filter" in obj:
                        filter_type = obj["/Filter"]
                        
                        if filter_type == "/DCTDecode":
                            # JPEG
                            ext = "jpg"
                            output_file = output_dir / f"page{idx + 1}_img{image_count}.{ext}"
                            with open(output_file, "wb") as f:
                                f.write(data)
                            extracted_files.append(str(output_file))
                        elif filter_type == "/FlateDecode":
                            # PNG/raw
                            ext = "png"
                            output_file = output_dir / f"page{idx + 1}_img{image_count}.{ext}"
                            
                            # Try to create image from raw data
                            color_space = obj.get("/ColorSpace", "/DeviceRGB")
                            if color_space == "/DeviceRGB":
                                mode = "RGB"
                            elif color_space == "/DeviceGray":
                                mode = "L"
                            else:
                                mode = "RGB"
                            
                            try:
                                img = Image.frombytes(mode, (width, height), data)
                                img.save(output_file)
                                extracted_files.append(str(output_file))
                            except Exception:
                                # Save raw data as fallback
                                with open(output_file.with_suffix(".bin"), "wb") as f:
                                    f.write(data)
                                extracted_files.append(str(output_file.with_suffix(".bin")))
                        else:
                            # Other format - save raw
                            output_file = output_dir / f"page{idx + 1}_img{image_count}.bin"
                            with open(output_file, "wb") as f:
                                f.write(data)
                            extracted_files.append(str(output_file))
                    else:
                        # No filter - save raw
                        output_file = output_dir / f"page{idx + 1}_img{image_count}.bin"
                        with open(output_file, "wb") as f:
                            f.write(data)
                        extracted_files.append(str(output_file))
                        
                except Exception as e:
                    # Skip problematic images
                    pass
    
    return {
        "status": "success",
        "output_dir": str(output_dir),
        "extracted_files": extracted_files,
        "image_count": len(extracted_files)
    }


def create_pdf(output_path, text=None, images=None):
    """Create a new PDF from text or images"""
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from PIL import Image
    
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    c = canvas.Canvas(str(output_path), pagesize=letter)
    page_width, page_height = letter
    
    content_added = False
    
    # Add text if provided
    if text:
        c.setFont("Helvetica", 12)
        
        # Split text into lines
        lines = text.split("\n")
        y = page_height - inch
        
        for line in lines:
            if y < inch:
                c.showPage()
                c.setFont("Helvetica", 12)
                y = page_height - inch
            
            c.drawString(inch, y, line)
            y -= 14  # Line spacing
        
        content_added = True
    
    # Add images if provided
    if images:
        for img_path in images:
            img_path = Path(img_path)
            if not img_path.exists():
                continue
            
            if content_added:
                c.showPage()
            
            # Get image dimensions
            with Image.open(img_path) as img:
                img_width, img_height = img.size
            
            # Scale to fit page
            max_width = page_width - 2 * inch
            max_height = page_height - 2 * inch
            
            scale = min(max_width / img_width, max_height / img_height, 1)
            
            draw_width = img_width * scale
            draw_height = img_height * scale
            
            x = (page_width - draw_width) / 2
            y = (page_height - draw_height) / 2
            
            c.drawImage(str(img_path), x, y, draw_width, draw_height)
            content_added = True
    
    if not content_added:
        # Create blank page if nothing provided
        c.drawString(inch, page_height - inch, "")
    
    c.save()
    
    return {
        "status": "success",
        "output_path": str(output_path),
        "has_text": bool(text),
        "image_count": len(images) if images else 0
    }


def encrypt_pdf(path, password, output_path=None):
    """Add password protection to a PDF"""
    from pypdf import PdfReader, PdfWriter
    
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    reader = PdfReader(str(path))
    writer = PdfWriter()
    
    for page in reader.pages:
        writer.add_page(page)
    
    writer.encrypt(password)
    
    output = Path(output_path) if output_path else path
    output.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output, "wb") as f:
        writer.write(f)
    
    return {
        "status": "success",
        "output_path": str(output),
        "encrypted": True
    }


def decrypt_pdf(path, password, output_path=None):
    """Remove password protection from a PDF"""
    from pypdf import PdfReader, PdfWriter
    
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    reader = PdfReader(str(path))
    
    if reader.is_encrypted:
        reader.decrypt(password)
    
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    
    output = Path(output_path) if output_path else path
    output.parent.mkdir(parents=True, exist_ok=True)
    
    with open(output, "wb") as f:
        writer.write(f)
    
    return {
        "status": "success",
        "output_path": str(output),
        "decrypted": True
    }


def pdf_to_images(path, output_dir, pages=None, format="png", dpi=150):
    """Convert PDF pages to images using pdftoppm"""
    import subprocess
    import shutil
    
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Check if pdftoppm is available
    if not shutil.which("pdftoppm"):
        # Fallback: use pdf2image if available
        try:
            from pdf2image import convert_from_path
            
            from pypdf import PdfReader
            reader = PdfReader(str(path))
            page_indices = parse_page_range(pages, len(reader.pages))
            
            # pdf2image uses 1-based indexing
            first_page = min(page_indices) + 1 if page_indices else 1
            last_page = max(page_indices) + 1 if page_indices else len(reader.pages)
            
            images = convert_from_path(str(path), dpi=dpi, first_page=first_page, last_page=last_page)
            
            created_files = []
            for i, img in enumerate(images):
                output_file = output_dir / f"page_{first_page + i}.{format}"
                img.save(output_file)
                created_files.append(str(output_file))
            
            return {
                "status": "success",
                "output_dir": str(output_dir),
                "created_files": created_files,
                "method": "pdf2image"
            }
        except ImportError:
            return {
                "status": "error",
                "message": "pdftoppm not found. Install poppler-utils or pip install pdf2image",
                "type": "DependencyError"
            }
    
    # Use pdftoppm
    base_name = output_dir / path.stem
    
    cmd = ["pdftoppm"]
    
    if format == "png":
        cmd.append("-png")
    else:
        cmd.append("-jpeg")
    
    cmd.extend(["-r", str(dpi)])
    
    # Handle page range
    if pages:
        from pypdf import PdfReader
        reader = PdfReader(str(path))
        page_indices = parse_page_range(pages, len(reader.pages))
        if page_indices:
            cmd.extend(["-f", str(min(page_indices) + 1)])
            cmd.extend(["-l", str(max(page_indices) + 1)])
    
    cmd.extend([str(path), str(base_name)])
    
    result = subprocess.run(cmd, capture_output=True, text=True)
    
    if result.returncode != 0:
        return {
            "status": "error",
            "message": result.stderr,
            "type": "ConversionError"
        }
    
    # List created files
    ext = "png" if format == "png" else "jpg"
    created_files = sorted([str(f) for f in output_dir.glob(f"{path.stem}*.{ext}")])
    
    return {
        "status": "success",
        "output_dir": str(output_dir),
        "created_files": created_files,
        "method": "pdftoppm"
    }


def main():
    """CLI interface for pdf_engine"""
    parser = argparse.ArgumentParser(description='PDF file manipulation engine')
    subparsers = parser.add_subparsers(dest='command', help='Commands')
    
    # Info command
    info_parser = subparsers.add_parser('info', help='Get PDF info')
    info_parser.add_argument('--path', required=True, help='Path to PDF file')
    
    # Read command
    read_parser = subparsers.add_parser('read', help='Extract text from PDF')
    read_parser.add_argument('--path', required=True, help='Path to PDF file')
    read_parser.add_argument('--pages', default=None, help='Page range (e.g., "1-5", "1,3,5")')
    read_parser.add_argument('--method', default='pypdf', choices=['pypdf', 'pdfplumber'], 
                            help='Extraction method')
    
    # Extract tables command
    tables_parser = subparsers.add_parser('extract_tables', help='Extract tables from PDF')
    tables_parser.add_argument('--path', required=True, help='Path to PDF file')
    tables_parser.add_argument('--pages', default=None, help='Page range')
    
    # Merge command
    merge_parser = subparsers.add_parser('merge', help='Merge PDFs')
    merge_parser.add_argument('--output', required=True, help='Output path')
    merge_parser.add_argument('--inputs', nargs='+', required=True, help='Input PDF files')
    
    # Split command
    split_parser = subparsers.add_parser('split', help='Split PDF into pages')
    split_parser.add_argument('--path', required=True, help='Path to PDF file')
    split_parser.add_argument('--output-dir', required=True, help='Output directory')
    split_parser.add_argument('--pages', default=None, help='Page range')
    
    # Extract pages command
    extract_parser = subparsers.add_parser('extract_pages', help='Extract specific pages')
    extract_parser.add_argument('--path', required=True, help='Path to PDF file')
    extract_parser.add_argument('--output', required=True, help='Output path')
    extract_parser.add_argument('--pages', required=True, help='Page range to extract')
    
    # Rotate command
    rotate_parser = subparsers.add_parser('rotate', help='Rotate PDF pages')
    rotate_parser.add_argument('--path', required=True, help='Path to PDF file')
    rotate_parser.add_argument('--degrees', required=True, type=int, choices=[90, 180, 270],
                              help='Rotation degrees')
    rotate_parser.add_argument('--pages', default=None, help='Page range')
    rotate_parser.add_argument('--output', default=None, help='Output path')
    
    # Watermark command
    watermark_parser = subparsers.add_parser('add_watermark', help='Add watermark')
    watermark_parser.add_argument('--path', required=True, help='Path to PDF file')
    watermark_parser.add_argument('--text', required=True, help='Watermark text')
    watermark_parser.add_argument('--output', default=None, help='Output path')
    watermark_parser.add_argument('--opacity', type=float, default=0.3, help='Opacity (0.0-1.0)')
    watermark_parser.add_argument('--position', default='diagonal', 
                                 choices=['center', 'diagonal', 'bottom'], help='Position')
    
    # Extract images command
    images_parser = subparsers.add_parser('extract_images', help='Extract images from PDF')
    images_parser.add_argument('--path', required=True, help='Path to PDF file')
    images_parser.add_argument('--output-dir', required=True, help='Output directory')
    images_parser.add_argument('--pages', default=None, help='Page range')
    
    # Create command
    create_parser = subparsers.add_parser('create', help='Create new PDF')
    create_parser.add_argument('--output', required=True, help='Output path')
    create_parser.add_argument('--text', default=None, help='Text content')
    create_parser.add_argument('--images', nargs='*', default=None, help='Image paths')
    
    # Encrypt command
    encrypt_parser = subparsers.add_parser('encrypt', help='Encrypt PDF')
    encrypt_parser.add_argument('--path', required=True, help='Path to PDF file')
    encrypt_parser.add_argument('--password', required=True, help='Password')
    encrypt_parser.add_argument('--output', default=None, help='Output path')
    
    # Decrypt command
    decrypt_parser = subparsers.add_parser('decrypt', help='Decrypt PDF')
    decrypt_parser.add_argument('--path', required=True, help='Path to PDF file')
    decrypt_parser.add_argument('--password', required=True, help='Current password')
    decrypt_parser.add_argument('--output', default=None, help='Output path')
    
    # To images command
    to_images_parser = subparsers.add_parser('to_images', help='Convert PDF to images')
    to_images_parser.add_argument('--path', required=True, help='Path to PDF file')
    to_images_parser.add_argument('--output-dir', required=True, help='Output directory')
    to_images_parser.add_argument('--pages', default=None, help='Page range')
    to_images_parser.add_argument('--format', default='png', choices=['png', 'jpg'], help='Image format')
    to_images_parser.add_argument('--dpi', type=int, default=150, help='Resolution')
    
    args = parser.parse_args()
    
    try:
        if args.command == 'info':
            result = get_pdf_info(args.path)
        elif args.command == 'read':
            result = read_pdf_text(args.path, args.pages, args.method)
        elif args.command == 'extract_tables':
            result = extract_tables(args.path, args.pages)
        elif args.command == 'merge':
            result = merge_pdfs(args.output, args.inputs)
        elif args.command == 'split':
            result = split_pdf(args.path, args.output_dir, args.pages)
        elif args.command == 'extract_pages':
            result = extract_pages(args.path, args.output, args.pages)
        elif args.command == 'rotate':
            result = rotate_pdf(args.path, args.degrees, args.pages, args.output)
        elif args.command == 'add_watermark':
            result = add_watermark(args.path, args.text, args.output, args.opacity, args.position)
        elif args.command == 'extract_images':
            result = extract_images(args.path, args.output_dir, args.pages)
        elif args.command == 'create':
            result = create_pdf(args.output, args.text, args.images)
        elif args.command == 'encrypt':
            result = encrypt_pdf(args.path, args.password, args.output)
        elif args.command == 'decrypt':
            result = decrypt_pdf(args.path, args.password, args.output)
        elif args.command == 'to_images':
            result = pdf_to_images(args.path, args.output_dir, args.pages, args.format, args.dpi)
        else:
            parser.print_help()
            sys.exit(1)
        
        print(json.dumps(result, indent=2, default=str))
        
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
