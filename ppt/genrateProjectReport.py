#!/usr/bin/env python3
"""
PowerPoint Generator - COMPLETE REWRITE
100% match to reference images
"""

import argparse
import json
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
import os

# Slide dimensions
SLIDE_W = Inches(10)
SLIDE_H = Inches(5.625)

FONT_FAMILY = "Arial"

PALETTE = {
    "ink": RGBColor(21, 24, 31),
    "muted": RGBColor(90, 98, 116),
    "card": RGBColor(242, 244, 252),
    "accent": RGBColor(122, 135, 255),
    "table_even": RGBColor(255, 255, 255),
    "table_odd": RGBColor(222, 228, 243),
    "chip": RGBColor(232, 233, 243),
}


def style_paragraph(paragraph, *, text="", size=12, bold=False, color=None,
                    align=None, line_spacing=1.2):
    paragraph.text = text
    paragraph.font.name = FONT_FAMILY
    paragraph.font.size = Pt(size)
    paragraph.font.bold = bold
    paragraph.font.color.rgb = color or PALETTE["ink"]
    paragraph.line_spacing = line_spacing
    if align:
        paragraph.alignment = align


def add_pill_text(slide, text, left, top, width, height):
    pill = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left,
        top,
        width,
        height,
    )
    pill.fill.solid()
    pill.fill.fore_color.rgb = PALETTE["chip"]
    pill.line.fill.background()

    frame = pill.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.2)
    frame.margin_right = Inches(0.2)
    frame.margin_top = Inches(0.1)
    frame.margin_bottom = Inches(0.1)
    frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    style_paragraph(frame.paragraphs[0], text=text, size=11, color=PALETTE["ink"], align=PP_ALIGN.CENTER)
    return pill


def hex_to_rgb(hex_str):
    """Convert hex to RGB"""
    h = hex_str.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def add_table_borders(table):
    """Add visible borders to all table cells"""
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            
            # Create border XML
            for border in ['lnL', 'lnR', 'lnT', 'lnB']:
                ln = parse_xml(f'<a:{border} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="12700"><a:solidFill><a:srgbClr val="D1D5E3"/></a:solidFill></a:{border}>')
                tcPr.append(ln)

def create_title_slide(prs, slide_data):
    """Title slide - clean and simple"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Add logos if they exist
    left_logo = slide_data.get("left_logo", "")
    right_logo = slide_data.get("right_logo", "")
    
    if left_logo and os.path.exists(left_logo):
        try:
            slide.shapes.add_picture(left_logo, Inches(0.5), Inches(0.4), width=Inches(2.0))
        except:
            print(f"Could not load logo: {left_logo}")
    
    if right_logo and os.path.exists(right_logo):
        try:
            slide.shapes.add_picture(right_logo, Inches(7.5), Inches(0.4), width=Inches(2.0))
        except:
            print(f"Could not load logo: {right_logo}")
    
    # Main title - bold with breathing room
    title_text = slide_data.get("title", "Project Report - Biller Management Portal")
    title_box = slide.shapes.add_textbox(Inches(0.65), Inches(1.4), Inches(8.7), Inches(1.7))
    frame = title_box.text_frame
    frame.word_wrap = True
    style_paragraph(
        frame.paragraphs[0],
        text=title_text,
        size=44,
        bold=True,
        color=PALETTE["ink"],
        line_spacing=1.15,
    )
    
    # Subtitle chip - keeps copy contained
    subtitle_text = slide_data.get("subtitle", "")
    if subtitle_text:
        add_pill_text(
            slide,
            subtitle_text,
            Inches(0.65),
            Inches(3.0),
            Inches(8.7),
            Inches(0.9),
        )

def create_timeline_slide(prs, slide_data):
    """Timeline slide with table"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.65), Inches(0.25), Inches(8.7), Inches(0.7))
    style_paragraph(
        title_box.text_frame.paragraphs[0],
        text=slide_data.get("title", "Project Timeline & Current Status"),
        size=34,
        bold=True,
    )
    
    # Description box - pill style for clarity
    desc_text = slide_data.get("description", "")
    if desc_text:
        add_pill_text(
            slide,
            desc_text,
            Inches(0.65),
            Inches(0.95),
            Inches(8.7),
            Inches(0.75),
        )
    
    # Create table
    rows_data = slide_data.get("rows", [])
    if not rows_data:
        return
    
    # Table dimensions
    num_rows = len(rows_data) + 1  # +1 for header
    num_cols = 4
    
    # Create table at position
    left = Inches(0.5)
    top = Inches(1.8)
    width = Inches(9.0)
    height = Inches(3.35)
    
    shapes = slide.shapes
    table_shape = shapes.add_table(num_rows, num_cols, left, top, width, height)
    table = table_shape.table
    
    # Set column widths
    table.columns[0].width = Inches(2.0)  # Phase
    table.columns[1].width = Inches(2.3)  # Timeline
    table.columns[2].width = Inches(1.7)  # Status
    table.columns[3].width = Inches(3.0)  # Notes
    
    # Style header row
    headers = ["Phase", "Timeline", "Status", "Notes"]
    for col_idx, header_text in enumerate(headers):
        cell = table.cell(0, col_idx)
        
        # Header background - light gray
        cell.fill.solid()
        cell.fill.fore_color.rgb = PALETTE["card"]
        
        # Header text
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_top = Inches(0.08)
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        style_paragraph(text_frame.paragraphs[0], text=header_text, size=11, bold=True)
    
    # Fill data rows
    for row_idx, row_data in enumerate(rows_data, start=1):
        phase = row_data.get("phase", "")
        timeline = row_data.get("timeline", "")
        status = row_data.get("status", "")
        notes = row_data.get("notes", "")
        
        # Alternating row colors - LIGHT BLUE for odd, WHITE for even
        if row_idx % 2 == 1:  # Odd rows
            row_fill = PALETTE["table_odd"]
        else:  # Even rows
            row_fill = PALETTE["table_even"]
        
        # Apply background to all cells in row
        for col_idx in range(num_cols):
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_fill
            cell.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        table.rows[row_idx].height = Inches(0.65)
        
        # Phase column
        cell = table.cell(row_idx, 0)
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_top = Inches(0.08)
        style_paragraph(text_frame.paragraphs[0], text=phase, size=11, color=PALETTE["ink"])
        
        # Timeline column
        cell = table.cell(row_idx, 1)
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_top = Inches(0.08)
        style_paragraph(text_frame.paragraphs[0], text=timeline, size=11)
        
        # Status column - with color coding
        cell = table.cell(row_idx, 2)
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_top = Inches(0.08)
        
        p = text_frame.paragraphs[0]
        
        # Determine status color and add icon
        status_lower = status.lower()
        if "completed" in status_lower and "partially" not in status_lower:
            p.text = "✓ " + status
            status_color = RGBColor(22, 163, 74)  # Green
        elif "partially" in status_lower or "review" in status_lower:
            p.text = "✓ " + status
            status_color = RGBColor(22, 163, 74)  # Green
        elif "progress" in status_lower:
            p.text = "◐ " + status
            status_color = RGBColor(59, 130, 246)  # Blue
        elif "not started" in status_lower:
            p.text = "✗ " + status
            status_color = RGBColor(239, 68, 68)  # Red
        else:
            p.text = status
            status_color = RGBColor(0, 0, 0)
        
        style_paragraph(p, text=p.text, size=11, color=status_color)
        
        # Notes column
        cell = table.cell(row_idx, 3)
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_top = Inches(0.08)
        text_frame.word_wrap = True
        style_paragraph(text_frame.paragraphs[0], text=notes, size=10, color=PALETTE["muted"], line_spacing=1.3)
    
    # Add borders to table
    add_table_borders(table)

def create_three_column_slide(prs, slide_data):
    """Three column layout slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.65), Inches(0.25), Inches(8.7), Inches(0.7))
    style_paragraph(
        title_box.text_frame.paragraphs[0],
        text=slide_data.get("title", "Our Solution: A Unified Three Pillar Platform"),
        size=34,
        bold=True,
    )
    
    # Description box
    desc_text = slide_data.get("description", "")
    if desc_text:
        add_pill_text(
            slide,
            desc_text,
            Inches(0.65),
            Inches(0.95),
            Inches(8.7),
            Inches(0.75),
        )
    
    # Three columns
    columns = slide_data.get("columns", [])
    if not columns or len(columns) != 3:
        return
    
    col_width = Inches(2.8)
    col_gap = Inches(0.3)
    start_x = Inches(0.5)
    start_y = Inches(1.8)
    card_height = Inches(3.4)
    
    for i, col_data in enumerate(columns):
        x_pos = start_x + i * (col_width + col_gap)
        card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            x_pos,
            start_y,
            col_width,
            card_height,
        )
        card.fill.solid()
        card.fill.fore_color.rgb = PALETTE["card"]
        card.line.color.rgb = PALETTE["card"]
        card.shadow.inherit = False
        
        heading_box = slide.shapes.add_textbox(x_pos + Inches(0.2), start_y + Inches(0.15), col_width - Inches(0.4), Inches(0.6))
        heading_box.text_frame.word_wrap = True
        style_paragraph(
            heading_box.text_frame.paragraphs[0],
            text=col_data.get("heading", ""),
            size=14,
            bold=True,
        )
        
        bullets = col_data.get("bullets", [])
        bullets_box = slide.shapes.add_textbox(
            x_pos + Inches(0.2),
            start_y + Inches(0.9),
            col_width - Inches(0.4),
            card_height - Inches(1.05),
        )
        text_frame = bullets_box.text_frame
        text_frame.word_wrap = True
        text_frame.clear()
        
        for idx, bullet_text in enumerate(bullets):
            paragraph = text_frame.paragraphs[0] if idx == 0 else text_frame.add_paragraph()
            style_paragraph(paragraph, text=bullet_text, size=10, color=PALETTE["muted"], line_spacing=1.35)
            paragraph.space_after = Pt(6)

def create_table_slide(prs, slide_data):
    """Features table slide"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    
    # Title
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(9.0), Inches(0.5))
    p = title_box.text_frame.paragraphs[0]
    p.text = slide_data.get("title", "Modules & Deep Dive On Features")
    p.font.name = "Arial"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 0, 0)
    
    # Description box
    desc_text = slide_data.get("description", "")
    if desc_text:
        from pptx.enum.shapes import MSO_SHAPE
        desc_shape = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.5), Inches(0.9),
            Inches(9.0), Inches(0.4)
        )
        desc_shape.fill.solid()
        desc_shape.fill.fore_color.rgb = RGBColor(232, 233, 243)
        desc_shape.line.fill.background()
        
        text_frame = desc_shape.text_frame
        text_frame.clear()
        text_frame.word_wrap = True
        text_frame.margin_left = Inches(0.2)
        text_frame.margin_right = Inches(0.2)
        text_frame.margin_top = Inches(0.08)
        text_frame.margin_bottom = Inches(0.08)
        text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        p = text_frame.paragraphs[0]
        p.text = desc_text
        p.font.name = "Arial"
        p.font.size = Pt(10)
        p.font.color.rgb = RGBColor(0, 0, 0)
        p.alignment = PP_ALIGN.CENTER
    
    # Create table
    columns = slide_data.get("columns", [])
    rows_data = slide_data.get("rows", [])
    
    if not columns or not rows_data:
        return
    
    num_rows = len(rows_data) + 1
    num_cols = len(columns)
    
    left = Inches(0.5)
    top = Inches(1.5)
    width = Inches(9.0)
    height = Inches(3.7)
    
    table_shape = slide.shapes.add_table(num_rows, num_cols, left, top, width, height)
    table = table_shape.table
    
    # Set column widths
    table.columns[0].width = Inches(1.8)  # Proposed Scope Item
    table.columns[1].width = Inches(5.4)  # Delivered Feature(s)
    table.columns[2].width = Inches(1.8)  # Status
    
    # Header row
    for col_idx, header_text in enumerate(columns):
        cell = table.cell(0, col_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(243, 244, 246)
        
        text_frame = cell.text_frame
        text_frame.clear()
        text_frame.margin_left = Inches(0.1)
        text_frame.margin_top = Inches(0.08)
        
        p = text_frame.paragraphs[0]
        p.text = header_text
        p.font.name = "Arial"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 0, 0)
    
    # Data rows
    for row_idx, row in enumerate(rows_data, start=1):
        # Alternating colors
        if row_idx % 2 == 1:
            row_fill = RGBColor(209, 213, 227)  # Light blue
        else:
            row_fill = RGBColor(255, 255, 255)  # White
        
        for col_idx in range(num_cols):
            cell = table.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_fill
            
            text_frame = cell.text_frame
            text_frame.clear()
            text_frame.margin_left = Inches(0.1)
            text_frame.margin_top = Inches(0.1)
            text_frame.margin_right = Inches(0.1)
            text_frame.word_wrap = True
            
            value = row[col_idx] if isinstance(row, list) else row.get(columns[col_idx], "")
            
            p = text_frame.paragraphs[0]
            p.text = str(value)
            p.font.name = "Arial"
            p.font.size = Pt(9)
            p.font.color.rgb = RGBColor(0, 0, 0)
            p.line_spacing = 1.2
    
    # Add borders
    add_table_borders(table)

def generate_presentation(data_path, output_path):
    """Main function to generate presentation"""
    # Load data
    with open(data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)
    
    # Create presentation
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    
    # Generate slides
    for slide_data in data.get("slides", []):
        slide_type = slide_data.get("type", "")
        
        if slide_type == "title":
            create_title_slide(prs, slide_data)
        elif slide_type == "timeline":
            create_timeline_slide(prs, slide_data)
        elif slide_type == "three_column":
            create_three_column_slide(prs, slide_data)
        elif slide_type == "table":
            create_table_slide(prs, slide_data)
    
    # Save
    prs.save(output_path)
    print(f"✓ Presentation created successfully!")
    print(f"  File: {output_path}")
    print(f"  Slides: {len(prs.slides)}")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate PowerPoint from JSON")
    parser.add_argument("--data", default="data.json", help="Input JSON file")
    parser.add_argument("--out", default="Final_Report.pptx", help="Output PPTX file")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.data):
        print(f"ERROR: File not found: {args.data}")
        exit(1)
    
    generate_presentation(args.data, args.out)