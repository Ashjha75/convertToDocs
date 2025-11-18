#!/usr/bin/env python3
"""
PowerPoint Generator - Simple and Clean
Focus on getting the basics right
"""

import json
import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

def create_presentation(json_file, output_file):
    """Create presentation from JSON"""
    
    # Load JSON data
    with open(json_file, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Create new presentation
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    
    # Process each slide
    for slide_info in data.get("slides", []):
        slide_type = slide_info.get("type", "")
        
        if slide_type == "title":
            create_title_slide(prs, slide_info)
        elif slide_type == "timeline":
            create_timeline_slide(prs, slide_info)
        elif slide_type == "three_column":
            create_three_column_slide(prs, slide_info)
        elif slide_type == "table":
            create_features_table_slide(prs, slide_info)
    
    # Save
    prs.save(output_file)
    print(f"✓ Created: {output_file} ({len(prs.slides)} slides)")

def create_title_slide(prs, data):
    """Simple title slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    
    # Title
    left = Inches(0.5)
    top = Inches(2.0)
    width = Inches(9.0)
    height = Inches(1.5)
    
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    
    p = tf.paragraphs[0]
    p.text = data.get("title", "")
    p.font.name = "Arial"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    
    # Subtitle
    subtitle = data.get("subtitle", "")
    if subtitle:
        left = Inches(0.5)
        top = Inches(3.3)
        width = Inches(9.0)
        height = Inches(0.8)
        
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = subtitle
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.color.rgb = RGBColor(100, 100, 100)

def create_timeline_slide(prs, data):
    """Timeline slide with table"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title
    add_title(slide, data.get("title", "Project Timeline & Current Status"))
    
    # Description box
    desc = data.get("description", "")
    if desc:
        add_description_box(slide, desc, Inches(0.9))
    
    # Create table
    rows_data = data.get("rows", [])
    if not rows_data:
        return
    
    rows = len(rows_data) + 1  # +1 for header
    cols = 4
    
    left = Inches(0.5)
    top = Inches(1.5)
    width = Inches(9.0)
    height = Inches(3.6)
    
    # Add table
    graphic_frame = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = graphic_frame.table
    
    # Column widths
    table.columns[0].width = Inches(2.1)
    table.columns[1].width = Inches(2.3)
    table.columns[2].width = Inches(1.7)
    table.columns[3].width = Inches(2.9)
    
    # Header row
    headers = ["Phase", "Timeline", "Status", "Notes"]
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = header
        
        # Style header
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(243, 244, 246)
        
        paragraph = cell.text_frame.paragraphs[0]
        paragraph.font.name = "Arial"
        paragraph.font.size = Pt(11)
        paragraph.font.bold = True
        paragraph.font.color.rgb = RGBColor(0, 0, 0)
        
        # Add margins
        cell.text_frame.margin_left = Inches(0.1)
        cell.text_frame.margin_top = Inches(0.08)
    
    # Data rows
    for idx, row_data in enumerate(rows_data):
        row_idx = idx + 1
        
        # Alternating colors
        if row_idx % 2 == 1:
            row_color = RGBColor(209, 213, 227)  # Light blue
        else:
            row_color = RGBColor(255, 255, 255)  # White
        
        # Phase
        cell = table.cell(row_idx, 0)
        cell.text = row_data.get("phase", "")
        cell.fill.solid()
        cell.fill.fore_color.rgb = row_color
        format_cell(cell, 10)
        
        # Timeline
        cell = table.cell(row_idx, 1)
        cell.text = row_data.get("timeline", "")
        cell.fill.solid()
        cell.fill.fore_color.rgb = row_color
        format_cell(cell, 10)
        
        # Status (with color)
        cell = table.cell(row_idx, 2)
        status = row_data.get("status", "")
        
        # Add icon and determine color
        if "completed" in status.lower() and "partially" not in status.lower():
            cell.text = "✓ " + status
            color = RGBColor(22, 163, 74)
        elif "partially" in status.lower() or "review" in status.lower():
            cell.text = "✓ " + status
            color = RGBColor(22, 163, 74)
        elif "progress" in status.lower():
            cell.text = "◐ " + status
            color = RGBColor(59, 130, 246)
        elif "not started" in status.lower():
            cell.text = "✗ " + status
            color = RGBColor(239, 68, 68)
        else:
            cell.text = status
            color = RGBColor(0, 0, 0)
        
        cell.fill.solid()
        cell.fill.fore_color.rgb = row_color
        format_cell(cell, 10, color)
        
        # Notes
        cell = table.cell(row_idx, 3)
        cell.text = row_data.get("notes", "")
        cell.fill.solid()
        cell.fill.fore_color.rgb = row_color
        format_cell(cell, 9, RGBColor(80, 80, 80))

def create_three_column_slide(prs, data):
    """Three column layout"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title
    add_title(slide, data.get("title", ""))
    
    # Description
    desc = data.get("description", "")
    if desc:
        add_description_box(slide, desc, Inches(0.9))
    
    # Three columns
    columns_data = data.get("columns", [])
    if len(columns_data) != 3:
        return
    
    col_width = Inches(2.9)
    col_gap = Inches(0.25)
    start_x = Inches(0.5)
    start_y = Inches(1.6)
    
    for i, col in enumerate(columns_data):
        x = start_x + i * (col_width + col_gap)
        
        # Heading
        heading_box = slide.shapes.add_textbox(x, start_y, col_width, Inches(0.5))
        tf = heading_box.text_frame
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = col.get("heading", "")
        p.font.name = "Arial"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
        
        # Bullets
        bullets = col.get("bullets", [])
        bullet_box = slide.shapes.add_textbox(x, start_y + Inches(0.55), col_width, Inches(3.2))
        tf = bullet_box.text_frame
        tf.word_wrap = True
        tf.clear()
        
        for j, bullet in enumerate(bullets):
            if j == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()
            
            p.text = bullet
            p.font.name = "Arial"
            p.font.size = Pt(9)
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.space_after = Pt(10)
            p.line_spacing = 1.2

def create_features_table_slide(prs, data):
    """Features table"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title
    add_title(slide, data.get("title", ""))
    
    # Description
    desc = data.get("description", "")
    if desc:
        add_description_box(slide, desc, Inches(0.9))
    
    # Table
    columns = data.get("columns", [])
    rows_data = data.get("rows", [])
    
    if not columns or not rows_data:
        return
    
    rows = len(rows_data) + 1
    cols = len(columns)
    
    left = Inches(0.5)
    top = Inches(1.45)
    width = Inches(9.0)
    height = Inches(3.7)
    
    graphic_frame = slide.shapes.add_table(rows, cols, left, top, width, height)
    table = graphic_frame.table
    
    # Column widths
    table.columns[0].width = Inches(1.9)
    table.columns[1].width = Inches(5.3)
    table.columns[2].width = Inches(1.8)
    
    # Header
    for i, header in enumerate(columns):
        cell = table.cell(0, i)
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(243, 244, 246)
        
        p = cell.text_frame.paragraphs[0]
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
        
        cell.text_frame.margin_left = Inches(0.1)
        cell.text_frame.margin_top = Inches(0.08)
    
    # Data rows
    for idx, row in enumerate(rows_data):
        row_idx = idx + 1
        
        if row_idx % 2 == 1:
            row_color = RGBColor(209, 213, 227)
        else:
            row_color = RGBColor(255, 255, 255)
        
        for col_idx in range(cols):
            cell = table.cell(row_idx, col_idx)
            
            if isinstance(row, list):
                value = row[col_idx]
            else:
                value = row.get(columns[col_idx], "")
            
            cell.text = str(value)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_color
            format_cell(cell, 9)

def add_title(slide, title_text):
    """Add title to slide"""
    txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.5))
    tf = txBox.text_frame
    
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "Arial"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)

def add_description_box(slide, text, top_position):
    """Add rounded description box"""
    from pptx.enum.shapes import MSO_SHAPE
    
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.5), top_position,
        Inches(9.0), Inches(0.5)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(232, 233, 243)
    shape.line.fill.background()
    
    tf = shape.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = Inches(0.2)
    tf.margin_right = Inches(0.2)
    tf.margin_top = Inches(0.1)
    tf.margin_bottom = Inches(0.1)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    
    p = tf.paragraphs[0]
    p.text = text
    p.font.name = "Arial"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(0, 0, 0)
    p.alignment = PP_ALIGN.CENTER

def format_cell(cell, font_size, color=None):
    """Format table cell"""
    if color is None:
        color = RGBColor(0, 0, 0)
    
    p = cell.text_frame.paragraphs[0]
    p.font.name = "Arial"
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.line_spacing = 1.2
    
    cell.text_frame.margin_left = Inches(0.1)
    cell.text_frame.margin_top = Inches(0.08)
    cell.text_frame.margin_right = Inches(0.08)
    cell.text_frame.word_wrap = True

if __name__ == "__main__":
    import argparse
    
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", default="data.json")
    parser.add_argument("--out", default="Clean_Report.pptx")
    args = parser.parse_args()
    
    if not os.path.exists(args.data):
        print(f"ERROR: {args.data} not found")
        exit(1)
    
    create_presentation(args.data, args.out)