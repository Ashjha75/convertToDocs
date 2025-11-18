#!/usr/bin/env python3
"""
Professional PowerPoint Generator - Exact Style Match
Matches the reference screenshots precisely
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

# Exact margins from reference
MARGIN_LEFT = Inches(0.4)
MARGIN_RIGHT = Inches(0.4)
MARGIN_TOP = Inches(0.5)
CONTENT_WIDTH = SLIDE_W - MARGIN_LEFT - MARGIN_RIGHT

def hex_to_rgb(hex_color):
    """Convert hex color to RGB tuple"""
    h = (hex_color or "#000000").lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def add_text_with_style(slide, text, left, top, width, height, 
                        font_name="Arial", font_size=11, bold=False, 
                        color="#000000", align=PP_ALIGN.LEFT):
    """Add text box with exact styling"""
    textbox = slide.shapes.add_textbox(left, top, width, height)
    text_frame = textbox.text_frame
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.1)
    text_frame.margin_right = Inches(0.1)
    text_frame.margin_top = Inches(0.05)
    text_frame.margin_bottom = Inches(0.05)
    
    p = text_frame.paragraphs[0]
    p.text = text
    p.font.name = font_name
    p.font.size = Pt(font_size)
    p.font.bold = bold
    r, g, b = hex_to_rgb(color)
    p.font.color.rgb = RGBColor(r, g, b)
    p.alignment = align
    p.line_spacing = 1.15
    
    return textbox

def add_rounded_box(slide, left, top, width, height, fill_color="#E8E9F3", 
                    text="", font_size=11, text_color="#000000"):
    """Add rounded rectangle box with text"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left, top, width, height
    )
    shape.fill.solid()
    r, g, b = hex_to_rgb(fill_color)
    shape.fill.fore_color.rgb = RGBColor(r, g, b)
    shape.line.fill.background()
    
    if text:
        text_frame = shape.text_frame
        text_frame.word_wrap = True
        text_frame.margin_left = Inches(0.15)
        text_frame.margin_right = Inches(0.15)
        text_frame.margin_top = Inches(0.08)
        text_frame.margin_bottom = Inches(0.08)
        
        p = text_frame.paragraphs[0]
        p.text = text
        p.font.name = "Arial"
        p.font.size = Pt(font_size)
        p.font.bold = False
        r, g, b = hex_to_rgb(text_color)
        p.font.color.rgb = RGBColor(r, g, b)
        p.line_spacing = 1.15
    
    return shape

def resize_image(img_path, max_width_in=2.0):
    """Resize image if needed"""
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
    except:
        return None

def create_title_slide(prs, slide_data):
    """Create title slide matching reference"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # White background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(255, 255, 255)
    bg.line.fill.background()
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)
    
    # Left logo (Supreme Ventures)
    left_logo = slide_data.get("left_logo")
    if left_logo and os.path.exists(left_logo):
        img = resize_image(left_logo, 2.5)
        if img:
            slide.shapes.add_picture(img, Inches(0.4), Inches(0.4), width=Inches(2.5))
    
    # Right logo (New Fields)
    right_logo = slide_data.get("right_logo")
    if right_logo and os.path.exists(right_logo):
        img = resize_image(right_logo, 2.2)
        if img:
            slide.shapes.add_picture(img, SLIDE_W - Inches(2.6), Inches(0.3), width=Inches(2.2))
    
    # Main title - black, bold, large
    add_text_with_style(
        slide,
        slide_data.get("title", "Project Report"),
        Inches(0.4), Inches(2.1),
        SLIDE_W - Inches(0.8), Inches(0.8),
        font_name="Arial", font_size=40, bold=True,
        color="#000000", align=PP_ALIGN.LEFT
    )
    
    # Subtitle - smaller, gray
    if slide_data.get("subtitle"):
        add_text_with_style(
            slide,
            slide_data.get("subtitle", ""),
            Inches(0.4), Inches(3.0),
            SLIDE_W - Inches(0.8), Inches(0.6),
            font_name="Arial", font_size=12, bold=False,
            color="#4A4A4A", align=PP_ALIGN.LEFT
        )

def create_timeline_slide(prs, slide_data):
    """Create timeline slide EXACTLY matching reference images"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # White background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(255, 255, 255)
    bg.line.fill.background()
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)
    
    # Title - exact match
    add_text_with_style(
        slide,
        slide_data.get("title", "Project Timeline & Current Status"),
        MARGIN_LEFT, Inches(0.35),
        CONTENT_WIDTH, Inches(0.5),
        font_name="Arial", font_size=36, bold=True,
        color="#000000", align=PP_ALIGN.LEFT
    )
    
    # Description box - light purple rounded box
    desc = slide_data.get("description", "")
    if desc:
        add_rounded_box(
            slide,
            MARGIN_LEFT, Inches(0.95),
            CONTENT_WIDTH, Inches(0.45),
            fill_color="#E8E9F3",
            text=desc,
            font_size=10,
            text_color="#000000"
        )
    
    # Table
    rows = slide_data.get("rows", [])
    if not rows:
        return
    
    table_top = Inches(1.55)
    num_rows = len(rows) + 1
    num_cols = 4
    
    # Calculate table dimensions
    table_width = CONTENT_WIDTH
    row_height = Inches(0.38)
    table_height = row_height * num_rows
    
    table = slide.shapes.add_table(
        num_rows, num_cols,
        MARGIN_LEFT, table_top,
        table_width, table_height
    ).table
    
    # Set exact column widths from reference
    table.columns[0].width = Inches(2.0)   # Phase
    table.columns[1].width = Inches(2.2)   # Timeline
    table.columns[2].width = Inches(1.8)   # Status
    table.columns[3].width = Inches(3.2)   # Notes
    
    # Header row - light gray background
    headers = ["Phase", "Timeline", "Status", "Notes"]
    for col_idx, header in enumerate(headers):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(243, 244, 246)  # #F3F4F6
        
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.08)
        text_frame.margin_top = Inches(0.05)
        
        p = text_frame.paragraphs[0]
        p.text = header
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = PP_ALIGN.LEFT
    
    # Data rows
    for row_idx, row_data in enumerate(rows, start=1):
        phase = row_data.get("phase", "")
        timeline = row_data.get("timeline", "")
        status = row_data.get("status", "")
        notes = row_data.get("notes", "")
        
        # Determine status color and icon
        status_color = "#000000"
        status_icon = ""
        
        if "completed" in status.lower() and "partially" not in status.lower():
            status_color = "#16A34A"  # Green
            status_icon = "✓ "
        elif "partially" in status.lower() or "review" in status.lower():
            status_color = "#16A34A"  # Green
            status_icon = "✓ "
        elif "progress" in status.lower():
            status_color = "#3B82F6"  # Blue
            status_icon = "◐ "
        elif "not started" in status.lower():
            status_color = "#EF4444"  # Red
            status_icon = "✗ "
        
        # Apply row striping - alternating light blue
        for col_idx in range(num_cols):
            cell = table.cell(row_idx, col_idx)
            if row_idx % 2 == 1:  # Odd rows get light blue
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(209, 213, 227)  # #D1D5E3
            else:  # Even rows stay white
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        # Phase column
        cell = table.cell(row_idx, 0)
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.08)
        text_frame.margin_top = Inches(0.05)
        p = text_frame.paragraphs[0]
        p.text = phase
        p.font.name = "Arial"
        p.font.size = Pt(10)
        p.font.bold = False
        p.font.color.rgb = RGBColor(0, 0, 0)
        
        # Timeline column
        cell = table.cell(row_idx, 1)
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.08)
        text_frame.margin_top = Inches(0.05)
        p = text_frame.paragraphs[0]
        p.text = timeline
        p.font.name = "Arial"
        p.font.size = Pt(10)
        p.font.bold = False
        p.font.color.rgb = RGBColor(0, 0, 0)
        
        # Status column with color
        cell = table.cell(row_idx, 2)
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.08)
        text_frame.margin_top = Inches(0.05)
        p = text_frame.paragraphs[0]
        p.text = status_icon + status
        p.font.name = "Arial"
        p.font.size = Pt(10)
        p.font.bold = False
        r, g, b = hex_to_rgb(status_color)
        p.font.color.rgb = RGBColor(r, g, b)
        
        # Notes column
        cell = table.cell(row_idx, 3)
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.08)
        text_frame.margin_top = Inches(0.05)
        p = text_frame.paragraphs[0]
        p.text = notes
        p.font.name = "Arial"
        p.font.size = Pt(9)
        p.font.bold = False
        p.font.color.rgb = RGBColor(74, 74, 74)  # #4A4A4A

def create_three_column_slide(prs, slide_data):
    """Create three column slide matching reference"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # White background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(255, 255, 255)
    bg.line.fill.background()
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)
    
    # Title
    add_text_with_style(
        slide,
        slide_data.get("title", ""),
        MARGIN_LEFT, Inches(0.35),
        CONTENT_WIDTH, Inches(0.5),
        font_name="Arial", font_size=36, bold=True,
        color="#000000", align=PP_ALIGN.LEFT
    )
    
    # Description box
    desc = slide_data.get("description", "")
    if desc:
        add_rounded_box(
            slide,
            MARGIN_LEFT, Inches(0.95),
            CONTENT_WIDTH, Inches(0.45),
            fill_color="#E8E9F3",
            text=desc,
            font_size=10,
            text_color="#000000"
        )
    
    # Three columns
    columns = slide_data.get("columns", [])
    if not columns:
        return
    
    col_width = (CONTENT_WIDTH - Inches(0.4)) / 3
    content_top = Inches(1.55)
    
    for i, col in enumerate(columns):
        x_pos = MARGIN_LEFT + i * (col_width + Inches(0.2))
        
        # Column heading - bold
        add_text_with_style(
            slide,
            col.get("heading", ""),
            x_pos, content_top,
            col_width, Inches(0.4),
            font_name="Arial", font_size=13, bold=True,
            color="#000000", align=PP_ALIGN.LEFT
        )
        
        # Bullet points
        bullets_top = content_top + Inches(0.48)
        textbox = slide.shapes.add_textbox(
            x_pos, bullets_top,
            col_width, Inches(3.2)
        )
        text_frame = textbox.text_frame
        text_frame.word_wrap = True
        text_frame.margin_left = 0
        text_frame.margin_right = Inches(0.05)
        
        bullets = col.get("bullets", [])
        for idx, bullet in enumerate(bullets):
            if idx == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()
            
            p.text = bullet
            p.font.name = "Arial"
            p.font.size = Pt(9)
            p.font.bold = False
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.space_after = Pt(10)
            p.line_spacing = 1.2

def create_table_slide(prs, slide_data):
    """Create feature table slide matching reference"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # White background
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SLIDE_W, SLIDE_H)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(255, 255, 255)
    bg.line.fill.background()
    slide.shapes._spTree.remove(bg._element)
    slide.shapes._spTree.insert(2, bg._element)
    
    # Title
    add_text_with_style(
        slide,
        slide_data.get("title", ""),
        MARGIN_LEFT, Inches(0.35),
        CONTENT_WIDTH, Inches(0.5),
        font_name="Arial", font_size=36, bold=True,
        color="#000000", align=PP_ALIGN.LEFT
    )
    
    # Description box
    desc = slide_data.get("description", "")
    if desc:
        add_rounded_box(
            slide,
            MARGIN_LEFT, Inches(0.95),
            CONTENT_WIDTH, Inches(0.38),
            fill_color="#E8E9F3",
            text=desc,
            font_size=10,
            text_color="#000000"
        )
    
    # Table
    columns = slide_data.get("columns", [])
    rows = slide_data.get("rows", [])
    
    if not columns or not rows:
        return
    
    table_top = Inches(1.45)
    num_rows = len(rows) + 1
    num_cols = len(columns)
    
    table_width = CONTENT_WIDTH
    row_height = Inches(0.52)
    table_height = row_height * num_rows
    
    table = slide.shapes.add_table(
        num_rows, num_cols,
        MARGIN_LEFT, table_top,
        table_width, table_height
    ).table
    
    # Set column widths
    table.columns[0].width = Inches(1.8)  # Proposed Scope
    table.columns[1].width = Inches(5.5)  # Delivered Features
    table.columns[2].width = Inches(1.9)  # Status
    
    # Header row
    for col_idx, header in enumerate(columns):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(243, 244, 246)
        
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.08)
        text_frame.margin_top = Inches(0.05)
        
        p = text_frame.paragraphs[0]
        p.text = header
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
    
    # Data rows
    for row_idx, row in enumerate(rows, start=1):
        # Striping
        for col_idx in range(num_cols):
            cell = table.cell(row_idx, col_idx)
            if row_idx % 2 == 1:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(209, 213, 227)
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
        
        for col_idx in range(num_cols):
            cell = table.cell(row_idx, col_idx)
            text_frame = cell.text_frame
            text_frame.clear()
            text_frame.margin_left = Inches(0.08)
            text_frame.margin_top = Inches(0.08)
            text_frame.margin_right = Inches(0.08)
            
            value = row[col_idx] if isinstance(row, (list, tuple)) else row.get(columns[col_idx], "")
            
            p = text_frame.paragraphs[0]
            p.text = str(value)
            p.font.name = "Arial"
            p.font.size = Pt(9)
            p.font.bold = False
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.line_spacing = 1.15

def generate_presentation(data_path, style_path, output_path):
    """Main generation function"""
    with open(data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    
    for slide_data in data.get("slides", []):
        slide_type = slide_data.get("type", "bullets")
        
        if slide_type == "title":
            create_title_slide(prs, slide_data)
        elif slide_type == "timeline":
            create_timeline_slide(prs, slide_data)
        elif slide_type == "three_column":
            create_three_column_slide(prs, slide_data)
        elif slide_type == "table":
            create_table_slide(prs, slide_data)
    
    prs.save(output_path)
    print(f"✓ Presentation created: {output_path}")
    print(f"  Slides: {len(prs.slides)}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", default="data.json")
    parser.add_argument("--style", default="style.json")
    parser.add_argument("--out", default="Biller_Project_Report.pptx")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.data):
        print(f"Error: {args.data} not found")
        exit(1)
    
    generate_presentation(args.data, args.style, args.out)