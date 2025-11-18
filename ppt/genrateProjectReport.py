#!/usr/bin/env python3
# json_to_ppt.py
"""
Generate a PPTX from data.json + style.json.
"""
import argparse
import json
import os
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_PARAGRAPH_ALIGNMENT
from PIL import Image

# ---------- Utilities ----------
def load_json(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2],16) for i in (0,2,4))

def safe_font_apply(run, style_token, style_cfg):
    if not style_token or style_token not in style_cfg.get("typography", {}):
        return
    t = style_cfg["typography"][style_token]
    if "family" in t:
        run.font.name = t["family"]
    if "size_pt" in t:
        run.font.size = Pt(t["size_pt"])
    if "weight_bold" in t:
        run.font.bold = bool(t["weight_bold"])
    if "color" in t:
        r,g,b = hex_to_rgb(t["color"])
        run.font.color.rgb = RGBColor(r,g,b)

def apply_text_style(paragraph, font_family=None, size_pt=None, bold=None, color=None):
    for run in paragraph.runs:
        if font_family:
            run.font.name = font_family
        if size_pt:
            run.font.size = Pt(size_pt)
        if bold is not None:
            run.font.bold = bold
        if color:
            r,g,b = hex_to_rgb(color)
            run.font.color.rgb = RGBColor(r,g,b)

def ensure_image_max_width(img_path, max_width_inches=6):
    # if image width > max_width, resize copy (temporary)
    try:
        img = Image.open(img_path)
    except Exception:
        return img_path
    w_px, h_px = img.size
    dpi = img.info.get('dpi', (96,96))[0]
    if dpi == 0: dpi = 96
    width_inches = w_px / dpi
    if width_inches <= max_width_inches:
        return img_path
    scale = max_width_inches / width_inches
    new_w = int(w_px * scale)
    new_h = int(h_px * scale)
    tmp_path = str(Path(img_path).with_suffix(".resized.png"))
    img.resize((new_w,new_h), Image.LANCZOS).save(tmp_path)
    return tmp_path

# ---------- Slide creation helpers ----------
def add_title_slide(prs, slide_data, style):
    slide_layout = prs.slide_layouts[6]  # blank
    slide = prs.slides.add_slide(slide_layout)
    # logos
    left_logo = slide_data.get("left_logo")
    right_logo = slide_data.get("right_logo")
    if left_logo:
        path = Path(left_logo)
        if path.exists():
            slide.shapes.add_picture(str(path), Inches(0.3), Inches(0.4), width=Inches(2.6))
    if right_logo:
        path = Path(right_logo)
        if path.exists():
            slide.shapes.add_picture(str(path), prs.slide_width - Inches(2.8), Inches(0.3), width=Inches(2.6))
    # title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), prs.slide_width - Inches(1.6), Inches(1.5))
    p = title_box.text_frame.paragraphs[0]
    p.text = slide_data.get("title","")
    # apply heading style if available
    heading_cfg = style.get("typography", {}).get("heading_font", {})
    apply_text_style(p, font_family=heading_cfg.get("family"), size_pt=heading_cfg.get("size_pt", 40),
                     bold=heading_cfg.get("weight_bold", True), color=style.get("colors",{}).get("primary_text"))
    # subtitle
    sub_text = slide_data.get("subtitle")
    if sub_text:
        sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(3.0), prs.slide_width - Inches(1.6), Inches(1.0))
        sp = sub_box.text_frame.paragraphs[0]
        sp.text = sub_text
        body_cfg = style.get("typography", {}).get("subheading_font", {})
        apply_text_style(sp, font_family=body_cfg.get("family"), size_pt=body_cfg.get("size_pt", 12),
                         bold=body_cfg.get("weight_bold", False), color=style.get("colors",{}).get("secondary_text"))
    return slide

def add_section_header_slide(prs, slide_data, style):
    slide_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(slide_layout)
    title = slide_data.get("title","")
    info_bar = slide_data.get("info_bar")
    # Title
    tbox = slide.shapes.add_textbox(Inches(0.6), Inches(0.6), prs.slide_width - Inches(1.2), Inches(1.2))
    p = tbox.text_frame.paragraphs[0]
    p.text = title
    hcfg = style["typography"].get("heading_font", {})
    apply_text_style(p, font_family=hcfg.get("family"), size_pt=hcfg.get("size_pt", 36), bold=True,
                     color=style["colors"].get("primary_text"))
    # info bar
    if info_bar:
        bar = slide.shapes.add_textbox(Inches(0.6), Inches(1.6), prs.slide_width - Inches(1.2), Inches(0.6))
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(*hex_to_rgb(style["colors"].get("muted_bar_bg","#EDEEF9")))
        bp = bar.text_frame.paragraphs[0]
        bp.text = info_bar
        scfg = style["typography"].get("subheading_font", {})
        apply_text_style(bp, font_family=scfg.get("family"), size_pt=scfg.get("size_pt", 12), bold=False,
                         color=style["colors"].get("primary_text"))
    return slide

def add_table_slide(prs, slide_data, style):
    cols = slide_data.get("columns", [])
    rows = slide_data.get("rows", [])
    title = slide_data.get("title")
    # compute rows count
    rcount = len(rows) + 1
    ccount = len(cols)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # title
    if title:
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), prs.slide_width - Inches(1.2), Inches(0.6))
        p = tb.text_frame.paragraphs[0]; p.text = title
        apply_text_style(p, font_family=style["typography"]["heading_font"]["family"],
                         size_pt=style["typography"]["heading_overrides"]["h2_size"],
                         bold=True, color=style["colors"]["primary_text"])
    # table position and size
    left = Inches(0.6); top = Inches(1.2)
    width = prs.slide_width - Inches(1.2)
    # create table
    table = slide.shapes.add_table(rcount, ccount, left, top, width, Inches(3.5)).table
    # header row
    for c, col_name in enumerate(cols):
        cell = table.cell(0,c)
        cell.text = col_name
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.name = style["typography"]["body_font"]["family"]
                run.font.size = Pt(style["table"]["header"]["size_pt"])
                run.font.bold = True
                run.font.color.rgb = RGBColor(*hex_to_rgb(style["colors"]["primary_text"]))
        # header background
        cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(*hex_to_rgb(style["colors"]["table_header_bg"]))
    # fill rows
    for r, row in enumerate(rows, start=1):
        for c in range(ccount):
            cell = table.cell(r,c)
            text = ""
            if isinstance(row, dict):
                # map by column name
                key = cols[c]
                text = str(row.get(key.lower(), "")) if row.get(key.lower()) is not None else ""
            else:
                try:
                    text = str(row[c])
                except Exception:
                    text = ""
            cell.text = text
            # style
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = style["typography"]["body_font"]["family"]
                    run.font.size = Pt(style["table"]["row"]["size_pt"])
                    run.font.bold = False
                    run.font.color.rgb = RGBColor(*hex_to_rgb(style["colors"]["primary_text"]))
        # row strip
        if style["table"].get("allow_row_strip", True):
            if r % 2 == 0:
                for c in range(ccount):
                    table.cell(r,c).fill.solid()
                    table.cell(r,c).fill.fore_color.rgb = RGBColor(*hex_to_rgb(style["colors"]["table_row_stripe"]))
    return slide

def add_three_column_slide(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    cols = slide_data.get("columns", [])
    # compute column widths
    total_w = prs.slide_width - Inches(1.2)
    col_w = total_w / 3
    top = Inches(1.0)
    left = Inches(0.6)
    for i, col in enumerate(cols):
        x = left + i * col_w
        # icon if present
        icon = col.get("icon")
        if icon and Path(icon).exists():
            img_path = ensure_image_max_width(icon, max_width_inches=1.2)
            slide.shapes.add_picture(img_path, x, top, width=Inches(0.9))
            text_left = x + Inches(1.0)
            text_width = col_w - Inches(1.0)
        else:
            text_left = x
            text_width = col_w
        # heading
        tb = slide.shapes.add_textbox(text_left, top, text_width, Inches(0.4))
        p = tb.text_frame.paragraphs[0]; p.text = col.get("heading","")
        apply_text_style(p, font_family=style["typography"]["body_font"]["family"],
                         size_pt=style["three_column_content"]["item_heading"]["size_pt"],
                         bold=True, color=style["colors"]["primary_text"])
        # bullets
        bullets = col.get("bullets",[])
        btop = top + Inches(0.5)
        tbox = slide.shapes.add_textbox(text_left, btop, text_width, Inches(3.5))
        tf = tbox.text_frame
        tf.word_wrap = True
        first = True
        for b in bullets:
            if first:
                p = tf.paragraphs[0]
                p.text = b
                p.level = 0
                first = False
            else:
                p = tf.add_paragraph()
                p.text = b
                p.level = 0
            apply_text_style(p, font_family=style["typography"]["body_font"]["family"],
                             size_pt=style["three_column_content"]["item_bullets"]["size_pt"],
                             bold=False, color=style["colors"]["primary_text"])
    return slide

def add_bullets_slide(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide_data.get("title")
    if title:
        tbox = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), prs.slide_width - Inches(1.2), Inches(0.6))
        p = tbox.text_frame.paragraphs[0]; p.text = title
        apply_text_style(p, font_family=style["typography"]["heading_font"]["family"],
                         size_pt=style["typography"]["heading_overrides"]["h2_size"], bold=True,
                         color=style["colors"]["primary_text"])
    bullets = slide_data.get("bullets",[])
    tbox = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), prs.slide_width - Inches(1.6), Inches(4.5))
    tf = tbox.text_frame
    tf.word_wrap = True
    first = True
    for b in bullets:
        if first:
            p = tf.paragraphs[0]; p.text = b; first = False
        else:
            p = tf.add_paragraph(); p.text = b
        p.level = 0
        apply_text_style(p, font_family=style["typography"]["body_font"]["family"],
                         size_pt=style["typography"]["body_font"]["size_pt"],
                         bold=False, color=style["colors"]["primary_text"])
    return slide

def add_notes_slide(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide_data.get("title","Notes")
    tbox = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), prs.slide_width - Inches(1.2), Inches(0.6))
    p = tbox.text_frame.paragraphs[0]; p.text = title
    apply_text_style(p, font_family=style["typography"]["heading_font"]["family"],
                     size_pt=style["typography"]["heading_overrides"]["h3_size"], bold=True,
                     color=style["colors"]["primary_text"])
    notes = slide_data.get("notes","")
    nbox = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), prs.slide_width - Inches(1.6), Inches(4.5))
    np = nbox.text_frame.paragraphs[0]; np.text = notes
    apply_text_style(np, font_family=style["typography"]["body_font"]["family"],
                     size_pt=style["typography"]["body_font"]["size_pt"], bold=False,
                     color=style["colors"]["secondary_text"])
    return slide

# ---------- Main generator ----------
def generate_ppt(data_path: Path, style_path: Path = None, output_name: str = None):
    data = load_json(data_path)
    style = {}
    if style_path and style_path.exists():
        style = load_json(style_path)
    # fallback very small defaults if style not provided
    if not style:
        style = {
            "colors":{"primary_text":"#0B0B0B","secondary_text":"#6B6B6B","muted_bar_bg":"#EDEEF9",
                      "table_header_bg":"#F3F4F6","table_row_stripe":"#FAFAFB","table_border":"#E1E4E8"},
            "typography":{"heading_font":{"family":"Georgia","size_pt":40,"weight_bold":True},
                          "body_font":{"family":"Calibri","size_pt":11}},
            "table":{"header":{"size_pt":11},"row":{"size_pt":11},"allow_row_strip":True}
        }
    prs = Presentation()
    # slide size 16:9
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # title slide if first slide is type title or presentation.title -> handled via slides array
    slides = data.get("slides", [])
    for s in slides:
        t = s.get("type","bullets")
        if t == "title":
            add_title_slide(prs, s, style)
        elif t == "section_header":
            add_section_header_slide(prs, s, style)
        elif t == "table":
            add_table_slide(prs, s, style)
        elif t == "three_column":
            add_three_column_slide(prs, s, style)
        elif t == "bullets":
            add_bullets_slide(prs, s, style)
        elif t == "notes_slide" or t == "notes":
            add_notes_slide(prs, s, style)
        else:
            # fallback to bullets for unknown types
            add_bullets_slide(prs, s, style)

    # output name
    if not output_name:
        title = data.get("presentation", {}).get("title", "presentation")
        safe = "".join(c for c in title if c.isalnum() or c in (" ", "-", "_")).rstrip()
        output_name = f"{safe}.pptx"
    prs.save(output_name)
    print("Saved:", output_name)

# ---------- CLI ----------
def main():
    ap = argparse.ArgumentParser(description="Generate PPTX from data.json + optional style.json")
    ap.add_argument("--data", required=True, type=Path, help="Path to data.json (content)")
    ap.add_argument("--style", required=False, type=Path, help="Path to style.json (theme)")
    ap.add_argument("--out", required=False, type=str, help="Output PPTX filename")
    args = ap.parse_args()
    if not args.data.exists():
        raise SystemExit("Data JSON not found")
    generate_ppt(args.data, args.style, args.out)

if __name__ == "__main__":
    main()
