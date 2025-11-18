#!/usr/bin/env python3
"""
Professional PowerPoint Generator
Generates clean, well-styled presentations matching corporate standards
"""

import argparse
import json
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import os

# Slide dimensions (16:9)
SLIDE_W = Inches(10)
SLIDE_H = Inches(5.625)

# Professional margins
MARGIN_LEFT = Inches(0.5)
MARGIN_RIGHT = Inches(0.5)
MARGIN_TOP = Inches(0.5)
MARGIN_BOTTOM = Inches(0.5)

def hex_to_rgb(hex_color):
    """Convert hex color to RGB tuple"""
    h = (hex_color or "#000000").lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def set_text_frame(text_frame, text, font_name="Calibri", font_size=11, 
                   bold=False, color="#000000", align=PP_ALIGN.LEFT):
    """Set text frame properties with proper formatting"""
    text_frame.clear()
    p = text_frame.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.bold = bold
    r, g, b = hex_to_rgb(color)
    p.font.color.rgb = RGBColor(r, g, b)
    p.alignment = align
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.05)
    text_frame.margin_right = Inches(0.05)

def add_background_shape(slide, color="#F5F5F5"):
    """Add a subtle background to slide"""
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        0, 0,
        SLIDE_W, SLIDE_H
    )
    bg.fill.solid()
    r, g, b = hex_to_rgb(color)
    bg.fill.fore_color.rgb = RGBColor(r, g, b)
    bg.line.fill.background()
    # Send to back
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)

def resize_image_if_needed(img_path, max_width_in=2.5):
    """Resize image if it's too large"""
    if not os.path.exists(img_path):
        return None
    
    try:
        img = Image.open(img_path)
        w, h = img.size
        dpi = img.info.get("dpi", (96, 96))[0] or 96
        w_in = w / dpi
        
        if w_in <= max_width_in:
            return img_path
        
        scale = max_width_in / w_in
        new_w = int(w * scale)
        new_h = int(h * scale)
        out_path = str(Path(img_path).with_suffix(".resized.png"))
        img.resize((new_w, new_h), Image.LANCZOS).save(out_path)
        return out_path
    except Exception as e:
        print(f"Error processing image {img_path}: {e}")
        return None

def create_title_slide(prs, slide_data, style):
    """Create professional title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background_shape(slide, "#FFFFFF")
    
    # Add logos
    left_logo = slide_data.get("left_logo")
    right_logo = slide_data.get("right_logo")
    
    if left_logo and os.path.exists(left_logo):
        img = resize_image_if_needed(left_logo, 2.0)
        if img:
            slide.shapes.add_picture(img, Inches(0.4), Inches(0.3), width=Inches(2.0))
    
    if right_logo and os.path.exists(right_logo):
        img = resize_image_if_needed(right_logo, 2.0)
        if img:
            slide.shapes.add_picture(img, SLIDE_W - Inches(2.4), Inches(0.3), width=Inches(2.0))
    
    # Title - centered and bold
    title_box = slide.shapes.add_textbox(
        Inches(1), Inches(2.2),
        SLIDE_W - Inches(2), Inches(1)
    )
    set_text_frame(title_box.text_frame, 
                   slide_data.get("title", "Project Report"),
                   font_name="Arial", font_size=36, bold=True,
                   color="#000000", align=PP_ALIGN.CENTER)
    
    # Subtitle
    if slide_data.get("subtitle"):
        subtitle_box = slide.shapes.add_textbox(
            Inches(1), Inches(3.4),
            SLIDE_W - Inches(2), Inches(0.8)
        )
        set_text_frame(subtitle_box.text_frame,
                       slide_data.get("subtitle", ""),
                       font_name="Arial", font_size=14, bold=False,
                       color="#666666", align=PP_ALIGN.CENTER)

def create_section_header(prs, slide_data, style):
    """Create section header slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background_shape(slide, "#FFFFFF")
    
    # Title
    title_box = slide.shapes.add_textbox(
        MARGIN_LEFT, MARGIN_TOP,
        SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT, Inches(0.7)
    )
    set_text_frame(title_box.text_frame,
                   slide_data.get("title", ""),
                   font_name="Arial", font_size=32, bold=True,
                   color="#000000", align=PP_ALIGN.LEFT)
    
    # Info bar with background
    if slide_data.get("info_bar"):
        bar_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            MARGIN_LEFT, Inches(1.1),
            SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT, Inches(0.5)
        )
        bar_shape.fill.solid()
        bar_shape.fill.fore_color.rgb = RGBColor(237, 238, 249)
        bar_shape.line.fill.background()
        
        text_frame = bar_shape.text_frame
        set_text_frame(text_frame, slide_data.get("info_bar", ""),
                       font_name="Calibri", font_size=11, bold=False,
                       color="#000000", align=PP_ALIGN.LEFT)

def create_three_column_slide(prs, slide_data, style):
    """Create three column feature slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background_shape(slide, "#FFFFFF")
    
    # Title
    title = slide_data.get("title", "")
    if title:
        title_box = slide.shapes.add_textbox(
            MARGIN_LEFT, MARGIN_TOP,
            SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT, Inches(0.6)
        )
        set_text_frame(title_box.text_frame, title,
                       font_name="Arial", font_size=28, bold=True,
                       color="#000000", align=PP_ALIGN.LEFT)
        content_top = Inches(1.3)
    else:
        content_top = MARGIN_TOP + Inches(0.2)
    
    # Description box if present
    desc = slide_data.get("description", "")
    if desc:
        desc_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            MARGIN_LEFT, content_top,
            SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT, Inches(0.5)
        )
        desc_shape.fill.solid()
        desc_shape.fill.fore_color.rgb = RGBColor(237, 238, 249)
        desc_shape.line.fill.background()
        set_text_frame(desc_shape.text_frame, desc,
                       font_name="Calibri", font_size=11, bold=False,
                       color="#000000", align=PP_ALIGN.LEFT)
        content_top += Inches(0.7)
    
    columns = slide_data.get("columns", [])
    if not columns:
        return
    
    col_width = (SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT - Inches(0.4)) / len(columns)
    
    for i, col in enumerate(columns):
        x_pos = MARGIN_LEFT + i * (col_width + Inches(0.2))
        
        # Icon
        icon_path = col.get("icon")
        icon_bottom = content_top
        if icon_path and os.path.exists(icon_path):
            img = resize_image_if_needed(icon_path, 0.6)
            if img:
                slide.shapes.add_picture(img, x_pos, content_top, width=Inches(0.5))
                icon_bottom = content_top + Inches(0.6)
        
        # Heading
        heading_box = slide.shapes.add_textbox(
            x_pos, icon_bottom + Inches(0.1),
            col_width, Inches(0.4)
        )
        set_text_frame(heading_box.text_frame,
                       col.get("heading", ""),
                       font_name="Arial", font_size=14, bold=True,
                       color="#000000", align=PP_ALIGN.LEFT)
        
        # Bullets
        bullets_top = icon_bottom + Inches(0.6)
        bullets_box = slide.shapes.add_textbox(
            x_pos, bullets_top,
            col_width, Inches(2.5)
        )
        text_frame = bullets_box.text_frame
        text_frame.word_wrap = True
        
        for idx, bullet in enumerate(col.get("bullets", [])):
            if idx == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            p.text = bullet
            p.level = 0
            p.font.name = "Calibri"
            p.font.size = Pt(10)
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.space_after = Pt(8)

def create_timeline_slide(prs, slide_data, style):
    """Create timeline table slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background_shape(slide, "#FFFFFF")
    
    # Title
    title = slide_data.get("title", "Timeline")
    title_box = slide.shapes.add_textbox(
        MARGIN_LEFT, MARGIN_TOP,
        SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT, Inches(0.6)
    )
    set_text_frame(title_box.text_frame, title,
                   font_name="Arial", font_size=28, bold=True,
                   color="#000000", align=PP_ALIGN.LEFT)
    
    # Description if present
    desc_top = Inches(1.1)
    desc = slide_data.get("description", "")
    if desc:
        desc_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            MARGIN_LEFT, desc_top,
            SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT, Inches(0.5)
        )
        desc_shape.fill.solid()
        desc_shape.fill.fore_color.rgb = RGBColor(237, 238, 249)
        desc_shape.line.fill.background()
        set_text_frame(desc_shape.text_frame, desc,
                       font_name="Calibri", font_size=10, bold=False,
                       color="#000000", align=PP_ALIGN.LEFT)
        table_top = desc_top + Inches(0.7)
    else:
        table_top = desc_top + Inches(0.1)
    
    rows = slide_data.get("rows", [])
    if not rows:
        return
    
    # Create table
    num_rows = len(rows) + 1  # +1 for header
    num_cols = 4  # Phase, Timeline, Status, Notes
    
    table_width = SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT
    table_height = Inches(0.35 * min(num_rows, 8))
    
    table = slide.shapes.add_table(
        num_rows, num_cols,
        MARGIN_LEFT, table_top,
        table_width, table_height
    ).table
    
    # Set column widths
    table.columns[0].width = Inches(2.0)  # Phase
    table.columns[1].width = Inches(2.2)  # Timeline
    table.columns[2].width = Inches(1.8)  # Status
    table.columns[3].width = Inches(3.0)  # Notes
    
    # Header row
    headers = ["Phase", "Timeline", "Status", "Notes"]
    for col_idx, header in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(243, 244, 246)
        text_frame = cell.text_frame
        text_frame.clear()
        p = text_frame.paragraphs[0]
        p.text = header
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
    
    # Data rows
    for row_idx, row_data in enumerate(rows, start=1):
        phase = row_data.get("phase", "")
        timeline = row_data.get("timeline", "")
        status = row_data.get("status", "")
        notes = row_data.get("notes", "")
        
        # Phase
        cell = table.cell(row_idx, 0)
        set_text_frame(cell.text_frame, phase, font_name="Calibri", 
                       font_size=10, bold=False, color="#000000")
        
        # Timeline
        cell = table.cell(row_idx, 1)
        set_text_frame(cell.text_frame, timeline, font_name="Calibri",
                       font_size=10, bold=False, color="#000000")
        
        # Status with icon
        cell = table.cell(row_idx, 2)
        status_text = status.title()
        if "complete" in status.lower():
            status_text = "✓ " + status_text
            color = "#16A34A"
        elif "progress" in status.lower():
            status_text = "◐ " + status_text
            color = "#3B82F6"
        elif "not" in status.lower():
            status_text = "✗ " + status_text
            color = "#EF4444"
        else:
            color = "#000000"
        set_text_frame(cell.text_frame, status_text, font_name="Calibri",
                       font_size=10, bold=False, color=color)
        
        # Notes
        cell = table.cell(row_idx, 3)
        set_text_frame(cell.text_frame, notes, font_name="Calibri",
                       font_size=9, bold=False, color="#666666")
        
        # Stripe rows
        if row_idx % 2 == 0:
            for col_idx in range(num_cols):
                table.cell(row_idx, col_idx).fill.solid()
                table.cell(row_idx, col_idx).fill.fore_color.rgb = RGBColor(250, 250, 251)

def create_table_slide(prs, slide_data, style):
    """Create generic table slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background_shape(slide, "#FFFFFF")
    
    # Title
    title = slide_data.get("title", "")
    if title:
        title_box = slide.shapes.add_textbox(
            MARGIN_LEFT, MARGIN_TOP,
            SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT, Inches(0.6)
        )
        set_text_frame(title_box.text_frame, title,
                       font_name="Arial", font_size=28, bold=True,
                       color="#000000", align=PP_ALIGN.LEFT)
        table_top = Inches(1.2)
    else:
        table_top = MARGIN_TOP + Inches(0.2)
    
    # Description
    desc = slide_data.get("description", "")
    if desc:
        desc_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            MARGIN_LEFT, table_top,
            SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT, Inches(0.5)
        )
        desc_shape.fill.solid()
        desc_shape.fill.fore_color.rgb = RGBColor(237, 238, 249)
        desc_shape.line.fill.background()
        set_text_frame(desc_shape.text_frame, desc,
                       font_name="Calibri", font_size=10, bold=False,
                       color="#000000", align=PP_ALIGN.LEFT)
        table_top += Inches(0.7)
    
    columns = slide_data.get("columns", [])
    rows = slide_data.get("rows", [])
    
    if not columns or not rows:
        return
    
    num_rows = len(rows) + 1
    num_cols = len(columns)
    
    table_width = SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT
    table_height = Inches(0.35 * min(num_rows, 10))
    
    table = slide.shapes.add_table(
        num_rows, num_cols,
        MARGIN_LEFT, table_top,
        table_width, table_height
    ).table
    
    # Set equal column widths
    col_width = table_width / num_cols
    for i in range(num_cols):
        table.columns[i].width = int(col_width)
    
    # Header
    for col_idx, header in enumerate(columns):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(243, 244, 246)
        text_frame = cell.text_frame
        text_frame.clear()
        p = text_frame.paragraphs[0]
        p.text = str(header)
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
    
    # Data rows
    for row_idx, row in enumerate(rows, start=1):
        for col_idx in range(num_cols):
            cell = table.cell(row_idx, col_idx)
            try:
                value = row[col_idx] if isinstance(row, (list, tuple)) else row.get(columns[col_idx], "")
            except:
                value = ""
            set_text_frame(cell.text_frame, str(value),
                           font_name="Calibri", font_size=10,
                           bold=False, color="#000000")
            
            if row_idx % 2 == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(250, 250, 251)

def create_bullets_slide(prs, slide_data, style):
    """Create bullet points slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_background_shape(slide, "#FFFFFF")
    
    # Title
    title = slide_data.get("title", "")
    if title:
        title_box = slide.shapes.add_textbox(
            MARGIN_LEFT, MARGIN_TOP,
            SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT, Inches(0.6)
        )
        set_text_frame(title_box.text_frame, title,
                       font_name="Arial", font_size=28, bold=True,
                       color="#000000", align=PP_ALIGN.LEFT)
        content_top = Inches(1.2)
    else:
        content_top = MARGIN_TOP + Inches(0.2)
    
    # Bullets
    bullets_box = slide.shapes.add_textbox(
        MARGIN_LEFT + Inches(0.2), content_top,
        SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT - Inches(0.4), Inches(3.5)
    )
    text_frame = bullets_box.text_frame
    text_frame.word_wrap = True
    
    for idx, bullet in enumerate(slide_data.get("bullets", [])):
        if idx == 0:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()
        p.text = bullet
        p.level = 0
        p.font.name = "Calibri"
        p.font.size = Pt(11)
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.space_after = Pt(12)

def generate_presentation(data_path, style_path, output_path):
    """Main generation function"""
    # Load data
    with open(data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Load style
    style = {}
    if os.path.exists(style_path):
        with open(style_path, 'r', encoding='utf-8') as f:
            style = json.load(f)
    
    # Create presentation
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    
    # Process slides
    for slide_data in data.get("slides", []):
        slide_type = slide_data.get("type", "bullets")
        
        if slide_type == "title":
            create_title_slide(prs, slide_data, style)
        elif slide_type == "section_header":
            create_section_header(prs, slide_data, style)
        elif slide_type == "three_column":
            create_three_column_slide(prs, slide_data, style)
        elif slide_type == "timeline":
            create_timeline_slide(prs, slide_data, style)
        elif slide_type == "table":
            create_table_slide(prs, slide_data, style)
        elif slide_type == "bullets":
            create_bullets_slide(prs, slide_data, style)
        elif slide_type in ("notes", "notes_slide"):
            create_bullets_slide(prs, slide_data, style)
        else:
            create_bullets_slide(prs, slide_data, style)
    
    # Save
    prs.save(output_path)
    print(f"✓ Presentation saved: {output_path}")
    print(f"  Total slides: {len(prs.slides)}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate professional PowerPoint presentations")
    parser.add_argument("--data", type=str, default="data.json", help="Path to data JSON file")
    parser.add_argument("--style", type=str, default="style.json", help="Path to style JSON file")
    parser.add_argument("--out", type=str, default="Professional_Report.pptx", help="Output file name")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.data):
        print(f"Error: Data file not found: {args.data}")
        exit(1)
    
    generate_presentation(args.data, args.style, args.out)