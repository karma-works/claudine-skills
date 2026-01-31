#!/usr/bin/env python3
"""
PPTX Engine - Core utilities for PowerPoint file manipulation
Provides safe loading, saving, reading, editing, and slide management
"""

from pptx import Presentation
from pptx.util import Inches, Pt
import json
import sys
import argparse
from pathlib import Path


def load_presentation_safe(path):
    """
    Safely load a PowerPoint presentation with error handling
    
    Args:
        path: Path to the .pptx file
        
    Returns:
        Presentation object
        
    Raises:
        FileNotFoundError: If the file doesn't exist
        Exception: If the file is not a valid PowerPoint file
    """
    path = Path(path)
    
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")
    
    try:
        prs = Presentation(str(path))
        return prs
    except Exception as e:
        raise Exception(f"Invalid PowerPoint file format: {path}. Error: {e}") from e


def save_presentation_safe(prs, path):
    """
    Safely save a PowerPoint presentation with error handling
    
    Args:
        prs: Presentation object
        path: Path to save the .pptx file
        
    Raises:
        Exception: If saving fails
    """
    try:
        path = Path(path)
        # Create parent directory if it doesn't exist
        path.parent.mkdir(parents=True, exist_ok=True)
        prs.save(str(path))
    except Exception as e:
        raise Exception(f"Error saving presentation: {e}") from e


def read_presentation(path, slide_number=None):
    """
    Extract content from presentation slides
    
    Args:
        path: Path to the .pptx file
        slide_number: Optional specific slide number (1-based index)
        
    Returns:
        dict: Contains presentation info and slide contents
    """
    prs = load_presentation_safe(path)
    
    result = {
        "total_slides": len(prs.slides),
        "slide_width": prs.slide_width,
        "slide_height": prs.slide_height,
        "slides": []
    }
    
    # Determine which slides to process
    if slide_number is not None:
        if slide_number < 1 or slide_number > len(prs.slides):
            raise ValueError(f"Slide number {slide_number} out of range. Presentation has {len(prs.slides)} slides.")
        slides_to_process = [(slide_number - 1, prs.slides[slide_number - 1])]
    else:
        slides_to_process = enumerate(prs.slides)
    
    # Process slides
    for idx, slide in slides_to_process:
        slide_data = {
            "slide_number": idx + 1,
            "shapes": []
        }
        
        # Extract shapes and text
        for shape_idx, shape in enumerate(slide.shapes):
            shape_info = {
                "shape_index": shape_idx,
                "shape_type": shape.shape_type.name if hasattr(shape.shape_type, 'name') else str(shape.shape_type),
                "name": shape.name
            }
            
            # Extract text if the shape has a text frame
            if hasattr(shape, "text_frame"):
                text_content = []
                for paragraph in shape.text_frame.paragraphs:
                    para_text = paragraph.text
                    if para_text:
                        text_content.append(para_text)
                
                if text_content:
                    shape_info["text"] = "\n".join(text_content)
            
            # Check if it's a title or has text
            if hasattr(shape, "text"):
                shape_info["full_text"] = shape.text
            
            # Add position and size
            shape_info["left"] = shape.left
            shape_info["top"] = shape.top
            shape_info["width"] = shape.width
            shape_info["height"] = shape.height
            
            slide_data["shapes"].append(shape_info)
        
        # Extract notes
        if slide.has_notes_slide:
            notes_text_frame = slide.notes_slide.notes_text_frame
            if notes_text_frame and notes_text_frame.text:
                slide_data["notes"] = notes_text_frame.text
        
        result["slides"].append(slide_data)
    
    return result


def edit_presentation(path, updates):
    """
    Update text content in presentation slides
    
    Args:
        path: Path to the .pptx file
        updates: List of dicts with slide, shape index, and new text
                 Examples:
                 - {"slide": 1, "shape": 0, "text": "New text"}
                 - {"slide": 1, "title": "New Title"} (finds first title shape)
                 - {"slide": 2, "notes": "New notes"}
    
    Returns:
        dict: Status message with updated items
    """
    prs = load_presentation_safe(path)
    updated_items = []
    
    for update in updates:
        slide_num = update.get("slide")
        
        if not slide_num or slide_num < 1 or slide_num > len(prs.slides):
            raise ValueError(f"Invalid slide number: {slide_num}. Presentation has {len(prs.slides)} slides.")
        
        slide = prs.slides[slide_num - 1]
        
        # Handle notes update
        if "notes" in update:
            if not slide.has_notes_slide:
                slide.notes_slide
            slide.notes_slide.notes_text_frame.text = update["notes"]
            updated_items.append(f"Slide {slide_num} notes")
            continue
        
        # Handle title update (find first shape with title)
        if "title" in update:
            title_updated = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame") and shape.name.lower().startswith("title"):
                    shape.text_frame.text = update["title"]
                    updated_items.append(f"Slide {slide_num} title")
                    title_updated = True
                    break
            
            if not title_updated:
                # Try to use the first text shape
                for shape in slide.shapes:
                    if hasattr(shape, "text_frame"):
                        shape.text_frame.text = update["title"]
                        updated_items.append(f"Slide {slide_num} first text shape (as title)")
                        title_updated = True
                        break
            
            if not title_updated:
                raise ValueError(f"No title or text shape found on slide {slide_num}")
            
            continue
        
        # Handle shape-specific update
        if "shape" in update:
            shape_idx = update["shape"]
            
            if shape_idx < 0 or shape_idx >= len(slide.shapes):
                raise ValueError(f"Shape index {shape_idx} out of range on slide {slide_num}")
            
            shape = slide.shapes[shape_idx]
            
            if "text" in update:
                if hasattr(shape, "text_frame"):
                    shape.text_frame.text = update["text"]
                    updated_items.append(f"Slide {slide_num}, Shape {shape_idx}")
                else:
                    raise ValueError(f"Shape {shape_idx} on slide {slide_num} does not support text")
            else:
                raise ValueError("Update must include 'text' when specifying a shape")
    
    # Save the presentation
    save_presentation_safe(prs, path)
    
    return {
        "status": "success",
        "updated_items": updated_items,
        "count": len(updated_items)
    }


def add_slide(path, layout_index=0, title=None, content=None):
    """
    Add a new slide to the presentation
    
    Args:
        path: Path to the .pptx file
        layout_index: Index of the slide layout to use (0-based)
        title: Optional title text
        content: Optional content text
        
    Returns:
        dict: Status and new slide information
    """
    prs = load_presentation_safe(path)
    
    # Check layout index
    if layout_index < 0 or layout_index >= len(prs.slide_layouts):
        raise ValueError(f"Layout index {layout_index} out of range. Available layouts: 0-{len(prs.slide_layouts)-1}")
    
    # Add slide
    slide_layout = prs.slide_layouts[layout_index]
    slide = prs.slides.add_slide(slide_layout)
    slide_number = len(prs.slides)
    
    # Set title if provided
    if title and hasattr(slide.shapes, 'title') and slide.shapes.title:
        slide.shapes.title.text = title
    
    # Set content if provided
    if content:
        # Try to find a content placeholder
        content_set = False
        for shape in slide.placeholders:
            # Skip title placeholder (usually index 0)
            if hasattr(shape, "text_frame") and shape.placeholder_format.idx != 0:
                shape.text = content
                content_set = True
                break
        
        if not content_set and hasattr(slide.shapes, 'title') and slide.shapes.title:
            # If no content placeholder, add a text box
            left = Inches(1)
            top = Inches(2)
            width = Inches(8)
            height = Inches(4)
            textbox = slide.shapes.add_textbox(left, top, width, height)
            textbox.text_frame.text = content
    
    # Save the presentation
    save_presentation_safe(prs, path)
    
    return {
        "status": "success",
        "slide_number": slide_number,
        "layout_index": layout_index,
        "layout_name": slide_layout.name,
        "title": title,
        "content": content if content else None
    }


def get_layouts(path):
    """
    Get all available slide layouts in the presentation
    
    Args:
        path: Path to the .pptx file
        
    Returns:
        dict: List of available layouts
    """
    prs = load_presentation_safe(path)
    
    layouts = []
    for idx, layout in enumerate(prs.slide_layouts):
        layout_info = {
            "index": idx,
            "name": layout.name,
            "placeholders": []
        }
        
        # Get placeholder information
        for shape in layout.placeholders:
            placeholder_info = {
                "idx": shape.placeholder_format.idx,
                "name": shape.name,
                "type": shape.placeholder_format.type.name if hasattr(shape.placeholder_format.type, 'name') else str(shape.placeholder_format.type)
            }
            layout_info["placeholders"].append(placeholder_info)
        
        layouts.append(layout_info)
    
    return {
        "total_layouts": len(layouts),
        "layouts": layouts
    }


def extract_images(path, output_dir):
    """
    Extract all images from the presentation
    
    Args:
        path: Path to the .pptx file
        output_dir: Directory to save extracted images
        
    Returns:
        dict: List of extracted image paths
    """
    prs = load_presentation_safe(path)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    
    extracted_images = []
    image_count = 0
    
    for slide_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if hasattr(shape, "image"):
                # Extract image
                image = shape.image
                image_bytes = image.blob
                
                # Determine extension
                ext = image.ext
                if not ext.startswith('.'):
                    ext = f'.{ext}'
                
                # Save image
                image_filename = f"slide_{slide_idx + 1}_image_{image_count}{ext}"
                image_path = output_path / image_filename
                
                with open(image_path, 'wb') as f:
                    f.write(image_bytes)
                
                extracted_images.append({
                    "slide": slide_idx + 1,
                    "filename": image_filename,
                    "path": str(image_path),
                    "content_type": image.content_type
                })
                
                image_count += 1
    
    return {
        "status": "success",
        "total_images": len(extracted_images),
        "output_directory": str(output_path),
        "images": extracted_images
    }


def main():
    """CLI interface for pptx_engine"""
    parser = argparse.ArgumentParser(description='PowerPoint file manipulation engine')
    subparsers = parser.add_subparsers(dest='command', help='Commands')
    
    # Read command
    read_parser = subparsers.add_parser('read', help='Read presentation data')
    read_parser.add_argument('--path', required=True, help='Path to PowerPoint file')
    read_parser.add_argument('--slide', type=int, default=None, help='Specific slide number (1-based)')
    
    # Edit command
    edit_parser = subparsers.add_parser('edit', help='Edit presentation data')
    edit_parser.add_argument('--path', required=True, help='Path to PowerPoint file')
    edit_parser.add_argument('--updates', required=True, help='JSON string of updates')
    
    # Add slide command
    add_parser = subparsers.add_parser('add_slide', help='Add a new slide')
    add_parser.add_argument('--path', required=True, help='Path to PowerPoint file')
    add_parser.add_argument('--layout', type=int, default=0, help='Layout index (default: 0)')
    add_parser.add_argument('--title', default=None, help='Slide title')
    add_parser.add_argument('--content', default=None, help='Slide content')
    
    # Get layouts command
    layouts_parser = subparsers.add_parser('get_layouts', help='Get available slide layouts')
    layouts_parser.add_argument('--path', required=True, help='Path to PowerPoint file')
    
    # Extract images command
    images_parser = subparsers.add_parser('extract_images', help='Extract images from presentation')
    images_parser.add_argument('--path', required=True, help='Path to PowerPoint file')
    images_parser.add_argument('--output', required=True, help='Output directory for images')
    
    args = parser.parse_args()
    
    try:
        if args.command == 'read':
            result = read_presentation(args.path, args.slide)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'edit':
            updates = json.loads(args.updates)
            result = edit_presentation(args.path, updates)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'add_slide':
            result = add_slide(args.path, args.layout, args.title, args.content)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'get_layouts':
            result = get_layouts(args.path)
            print(json.dumps(result, indent=2))
            
        elif args.command == 'extract_images':
            result = extract_images(args.path, args.output)
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
