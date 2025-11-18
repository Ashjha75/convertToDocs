#!/usr/bin/env python3
"""
json_to_ppt_final_ui.py
Improved UI safety: prevents overflow, scales long text, keeps layout consistent.
Usage:
  python json_to_ppt_final_ui.py --data data.json --out Final_Report.pptx
"""
import argparse, json, os
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from PIL import Image

# Slide dimensions (user used 10 x 5.625 which is 16:9 scaled)
SLIDE_W = Inches(10)
SLIDE_H = Inches(5.625)

FONT_FAMILY = "Arial"   # keep consistent across slides

# Palette (keeps previous colors)
PALETTE = {
    "ink": RGBColor(21, 24, 31),
    "muted": RGBColor(90, 98, 116),
    "chip": RGBColor(232, 233, 243),
    "header_bg": RGBColor(243, 244, 246),
    "row_odd": RGBColor(209, 213, 227),
    "white": RGBColor(255,255,255),
    "status_green": RGBColor(22,163,74),
    "status_blue": RGBColor(59,130,246),
    "status_red": RGBColor(239,68,68),
    "notes_text": RGBColor(80,80,80),
}

# ---------- helpers ----------
def hex_to_rgb_tuple(hex_str):
    h = hex_str.lstrip("#")
    return tuple(int(h[i:i+2],16) for i in (0,2,4))

def apply_run_style(run, size_pt=None, bold=None, color=None):
    if size_pt is not None: run.font.size = Pt(size_pt)
    if bold is not None: run.font.bold = bool(bold)
    if color is not None: run.font.color.rgb = color
    run.font.name = FONT_FAMILY

def set_paragraph_text(par, text, size_pt=11, bold=False, color=PALETTE["ink"], align=None, line_spacing=1.15):
    # clear runs and set a single run
    par.clear()
    r = par.add_run()
    r.text = text
    apply_run_style(r, size_pt, bold, color)
    par.line_spacing = line_spacing
    if align: par.alignment = align

def ensure_image_fit(img_path, max_w_in=1.0):
    """Return path, resizing image if wider than max_w_in inches (save to .resized)."""
    try:
        img = Image.open(img_path)
    except Exception:
        return img_path
    w_px, h_px = img.size
    dpi = img.info.get("dpi", (96,96))[0] or 96
    w_in = w_px / dpi
    if w_in <= max_w_in:
        return img_path
    scale = max_w_in / w_in
    new_w = int(w_px * scale); new_h = int(h_px * scale)
    out = Path(img_path).with_suffix(".resized.png")
    img.resize((new_w, new_h), Image.LANCZOS).save(out)
    return str(out)

def add_table_borders(table):
    """Add light borders to each table cell (drawingml)."""
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            for border in ('lnL','lnR','lnT','lnB'):
                # light gray stroke (D1D5E3)
                ln = parse_xml(f'<a:{border} xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="12700">'
                               f'<a:solidFill><a:srgbClr val="D1D5E3"/></a:solidFill></a:{border}>')
                tcPr.append(ln)

def fit_text_to_box(text_frame, text, box_w_in, box_h_in, start_size=12, min_size=7, bold=False, color=PALETTE["ink"]):
    """
    Best-effort shrink-to-fit: lowers font size until text fits approx by char heuristic.
    python-pptx doesn't provide precise layout metrics; we use a heuristic:
      - max_chars_per_line ~ box_width_inches * chars_per_inch (approx)
      - estimated_lines = ceil(len(text)/max_chars_per_line)
      - required_height = lines * (font_pt * 1.2) converted to inches -> compare with box_h_in
    Loops decreasing font_size until it fits or until min_size reached.
    """
    text = str(text or "")
    if not text:
        text_frame.clear()
        p = text_frame.paragraphs[0]
        set_paragraph_text(p, "", size_pt=start_size, bold=bold, color=color)
        return

    chars_per_inch_at_12 = 10.5  # heuristic: ~10-11 monospace chars per inch at 12pt
    font_size = start_size

    while font_size >= min_size:
        max_cpl = max(6, int(box_w_in * (chars_per_inch_at_12 * (12.0 / font_size))))
        import math
        lines_needed = math.ceil(len(text) / max_cpl) or 1
        # line height inches = (font_pt * 1.2) / 72
        line_h_in = (font_size * 1.2) / 72.0
        required_h = lines_needed * line_h_in
        if required_h <= box_h_in:
            break
        font_size -= 1

    # apply result
    text_frame.clear()
    tfparagraph = text_frame.paragraphs[0]
    set_paragraph_text(tfparagraph, text, size_pt=font_size, bold=bold, color=color)
    # if multiline, add breaks (let PPT wrap)
    text_frame.word_wrap = True

# ---------- slide builders ----------
def create_title_slide(prs, slide_data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    left_logo = slide_data.get("left_logo", "")
    right_logo = slide_data.get("right_logo", "")

    if left_logo and os.path.exists(left_logo):
        try:
            slide.shapes.add_picture(ensure_image_fit(left_logo, max_w_in=2.0), Inches(0.5), Inches(0.35), width=Inches(2.0))
        except Exception:
            pass
    if right_logo and os.path.exists(right_logo):
        try:
            slide.shapes.add_picture(ensure_image_fit(right_logo, max_w_in=2.0), SLIDE_W - Inches(2.5), Inches(0.35), width=Inches(2.0))
        except Exception:
            pass

    # Title box (pulled up)
    title = slide_data.get("title", "")
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.6), SLIDE_W - Inches(1.2), Inches(1.1)).text_frame
    fit_text_to_box(tb, title, box_w_in=(SLIDE_W - Inches(1.2)).inches, box_h_in=1.1, start_size=44, min_size=20, bold=True, color=PALETTE["ink"])

    subtitle = slide_data.get("subtitle", "")
    if subtitle:
        # pill
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.85), SLIDE_W - Inches(1.2), Inches(0.7))
        pill.fill.solid(); pill.fill.fore_color.rgb = PALETTE["chip"]
        pill.line.fill.background()
        tf = pill.text_frame
        fit_text_to_box(tf, subtitle, box_w_in=(SLIDE_W - Inches(1.2)).inches, box_h_in=0.7, start_size=12, min_size=8, bold=False, color=PALETTE["muted"])

def create_timeline_slide(prs, slide_data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide_data.get("title", "Project Timeline & Current Status")
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(9.0), Inches(0.45)).text_frame
    set_paragraph_text(tb.paragraphs[0], title, size_pt=34, bold=True, color=PALETTE["ink"])

    desc = slide_data.get("description", "")
    if desc:
        desc_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.85), SLIDE_W - Inches(1.0), Inches(0.45))
        desc_shape.fill.solid(); desc_shape.fill.fore_color.rgb = PALETTE["chip"]
        desc_shape.line.fill.background()
        tf = desc_shape.text_frame
        fit_text_to_box(tf, desc, box_w_in=(SLIDE_W - Inches(1.0)).inches, box_h_in=0.45, start_size=11, min_size=8, bold=False, color=PALETTE["ink"])

    rows = slide_data.get("rows", [])
    if not rows:
        return

    # Table dims
    num_rows = len(rows) + 1
    num_cols = 4
    left = Inches(0.5); top = Inches(1.55); width = SLIDE_W - Inches(1.0); height = Inches(3.2)
    table_shape = slide.shapes.add_table(num_rows, num_cols, left, top, width, height)
    table = table_shape.table

    # column widths (tuned)
    table.columns[0].width = Inches(2.0)
    table.columns[1].width = Inches(2.2)
    table.columns[2].width = Inches(1.4)
    table.columns[3].width = width - (Inches(2.0)+Inches(2.2)+Inches(1.4))

    # header
    headers = ["Phase","Timeline","Status","Notes"]
    for i,h in enumerate(headers):
        cell = table.cell(0,i)
        cell.fill.solid(); cell.fill.fore_color.rgb = PALETTE["header_bg"]
        tf = cell.text_frame
        fit_text_to_box(tf, h, box_w_in=table.columns[i].width.inches, box_h_in=0.45, start_size=11, min_size=9, bold=True, color=PALETTE["ink"])

    # rows - alternating
    for r_idx, rdata in enumerate(rows, start=1):
        bg = PALETTE["row_odd"] if (r_idx % 2 == 1) else PALETTE["white"]
        for c in range(num_cols):
            cell = table.cell(r_idx, c)
            cell.fill.solid(); cell.fill.fore_color.rgb = bg

        # fill cells with wrapping + fit
        phase = rdata.get("phase","")
        timeline = rdata.get("timeline","")
        status = rdata.get("status","")
        notes = rdata.get("notes","")

        # Phase
        tf = table.cell(r_idx,0).text_frame
        fit_text_to_box(tf, phase, box_w_in=table.columns[0].width.inches, box_h_in=0.9, start_size=10, min_size=7, bold=False)

        # Timeline
        tf = table.cell(r_idx,1).text_frame
        fit_text_to_box(tf, timeline, box_w_in=table.columns[1].width.inches, box_h_in=0.9, start_size=10, min_size=7, bold=False)

        # Status with simple color coding
        tf = table.cell(r_idx,2).text_frame
        s_text = str(status or "")
        status_lower = s_text.lower()
        if "completed" in status_lower and "partially" not in status_lower:
            color = PALETTE["status_green"]; icon = "✓ "
        elif "partially" in status_lower or "review" in status_lower:
            color = PALETTE["status_green"]; icon = "✓ "
        elif "progress" in status_lower:
            color = PALETTE["status_blue"]; icon = "◐ "
        elif "not" in status_lower:
            color = PALETTE["status_red"]; icon = "✗ "
        else:
            color = PALETTE["ink"]; icon = ""
        fit_text_to_box(table.cell(r_idx,2).text_frame, f"{icon}{s_text}", box_w_in=table.columns[2].width.inches, box_h_in=0.9, start_size=10, min_size=7, bold=False, color=color)

        # Notes
        tf = table.cell(r_idx,3).text_frame
        fit_text_to_box(tf, notes, box_w_in=table.columns[3].width.inches, box_h_in=0.9, start_size=9, min_size=7, bold=False, color=PALETTE["notes_text"])

    add_table_borders(table)

def create_three_column_slide(prs, slide_data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide_data.get("title","")
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), SLIDE_W - Inches(1.0), Inches(0.45)).text_frame
    set_paragraph_text(tb.paragraphs[0], title, size_pt=34, bold=True, color=PALETTE["ink"])

    desc = slide_data.get("description","")
    if desc:
        desc_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.85), SLIDE_W - Inches(1.0), Inches(0.4))
        desc_shape.fill.solid(); desc_shape.fill.fore_color.rgb = PALETTE["chip"]
        desc_shape.line.fill.background()
        fit_text_to_box(desc_shape.text_frame, desc, box_w_in=(SLIDE_W - Inches(1.0)).inches, box_h_in=0.4, start_size=10, min_size=8, bold=False)

    columns = slide_data.get("columns", [])
    if not columns: return
    # geometry tuned for 3 columns
    col_w = Inches(2.8); gap = Inches(0.25); start_x = Inches(0.5); start_y = Inches(1.55)
    for i, col in enumerate(columns):
        x = start_x + i * (col_w + gap)
        # heading
        hb = slide.shapes.add_textbox(x, start_y, col_w, Inches(0.28)).text_frame
        set_paragraph_text(hb.paragraphs[0], col.get("heading",""), size_pt=13, bold=True, color=PALETTE["ink"])
        # bullets
        btf = slide.shapes.add_textbox(x, start_y + Inches(0.35), col_w, Inches(3.2)).text_frame
        bullets = col.get("bullets", [])
        if bullets:
            # ensure first paragraph exists
            for idx, b in enumerate(bullets):
                if idx == 0:
                    p = btf.paragraphs[0]
                    fit_text_to_box(btf, b, box_w_in=col_w.inches, box_h_in=3.2, start_size=10, min_size=8)
                else:
                    p = btf.add_paragraph()
                    p.text = b
                    p.font.name = FONT_FAMILY
                    p.font.size = Pt(9)
                    p.font.color.rgb = PALETTE["ink"]
                    p.line_spacing = 1.15

def create_table_slide(prs, slide_data):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide_data.get("title","")
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), SLIDE_W - Inches(1.0), Inches(0.45)).text_frame
    set_paragraph_text(tb.paragraphs[0], title, size_pt=34, bold=True, color=PALETTE["ink"])

    desc = slide_data.get("description","")
    if desc:
        desc_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(0.85), SLIDE_W - Inches(1.0), Inches(0.35))
        desc_shape.fill.solid(); desc_shape.fill.fore_color.rgb = PALETTE["chip"]
        desc_shape.line.fill.background()
        fit_text_to_box(desc_shape.text_frame, desc, box_w_in=(SLIDE_W - Inches(1.0)).inches, box_h_in=0.35, start_size=10, min_size=8)

    columns = slide_data.get("columns", [])
    rows = slide_data.get("rows", [])
    if not columns or not rows: return

    num_rows = len(rows) + 1
    num_cols = len(columns)
    left = Inches(0.5); top = Inches(1.45); width = SLIDE_W - Inches(1.0); height = Inches(3.6)
    table_shape = slide.shapes.add_table(num_rows, num_cols, left, top, width, height)
    table = table_shape.table

    # Example column widths tuned to screenshot proportions (modifiable)
    if num_cols == 3:
        table.columns[0].width = Inches(1.8)
        table.columns[1].width = Inches(5.4)
        table.columns[2].width = Inches(1.8)

    # header
    for i,h in enumerate(columns):
        cell = table.cell(0,i)
        cell.fill.solid(); cell.fill.fore_color.rgb = PALETTE["header_bg"]
        fit_text_to_box(cell.text_frame, h, box_w_in=table.columns[i].width.inches, box_h_in=0.45, start_size=11, min_size=9, bold=True)

    # data rows
    for r_idx, row in enumerate(rows, start=1):
        bg = PALETTE["row_odd"] if (r_idx % 2 == 1) else PALETTE["white"]
        for c_idx in range(num_cols):
            cell = table.cell(r_idx, c_idx)
            cell.fill.solid(); cell.fill.fore_color.rgb = bg
            val = row[c_idx] if isinstance(row, list) else row.get(columns[c_idx], "")
            fit_text_to_box(cell.text_frame, str(val), box_w_in=table.columns[c_idx].width.inches, box_h_in=0.85, start_size=10, min_size=7, bold=False, color=PALETTE["ink"])

    add_table_borders(table)

# ---------- main ----------
def generate_from_json(data_path: str, out_path: str):
    with open(data_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H

    slides = data.get("slides", [])
    for s in slides:
        t = s.get("type","")
        if t == "title": create_title_slide(prs, s)
        elif t == "timeline": create_timeline_slide(prs, s)
        elif t == "three_column": create_three_column_slide(prs, s)
        elif t == "table": create_table_slide(prs, s)
        else:
            # fallback to table or bullets if present
            if s.get("rows") and s.get("columns"):
                create_table_slide(prs,s)
            else:
                # produce a safe bullets slide
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                if s.get("title"):
                    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), SLIDE_W - Inches(1.2), Inches(0.5)).text_frame
                    set_paragraph_text(tb.paragraphs[0], s.get("title",""), size_pt=28, bold=True, color=PALETTE["ink"])
                bullets = s.get("bullets", [])
                if bullets:
                    btf = slide.shapes.add_textbox(Inches(0.8), Inches(1.05), SLIDE_W - Inches(1.6), Inches(4.5)).text_frame
                    first = True
                    for b in bullets:
                        if first:
                            fit_text_to_box(btf, b, box_w_in=(SLIDE_W - Inches(1.6)).inches, box_h_in=4.5, start_size=11, min_size=8)
                            first = False
                        else:
                            p = btf.add_paragraph(); p.text = b
                            p.font.name = FONT_FAMILY; p.font.size = Pt(9); p.font.color.rgb = PALETTE["ink"]

    prs.save(out_path)
    print("Saved:", out_path)

# ---------- CLI ----------
if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", default="data.json")
    parser.add_argument("--out", default="Final_Report.pptx")
    args = parser.parse_args()
    if not os.path.exists(args.data):
        print("ERROR: data.json not found:", args.data); raise SystemExit(1)
    generate_from_json(args.data, args.out)
