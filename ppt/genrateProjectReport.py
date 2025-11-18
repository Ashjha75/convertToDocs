#!/usr/bin/env python3
# json_to_ppt.py
"""
Generate PPTX from data.json + style.json.

Usage (defaults editable in script):
    python json_to_ppt.py
    python json_to_ppt.py --data ./data.json --style ./style.json --out "My.pptx"
"""

import argparse, json, os
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image

# ----------------- CONFIG (edit here) -----------------
# Default paths (change these directly in the file if you prefer)
DATA_PATH_DEFAULT  = Path("data.json")
STYLE_PATH_DEFAULT = Path("style.json")   # optional - if missing script uses fallback style
OUTPUT_DEFAULT     = "Generated_Presentation.pptx"

# Slide size config (optional change)
SLIDE_WIDTH_IN  = 13.333   # 16:9 default width in inches
SLIDE_HEIGHT_IN = 7.5      # 16:9 default height in inches
# -----------------------------------------------------

# ---------- Helpers ----------
def load_json(path: Path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def hex_to_rgb(hex_color: str):
    hex_color = (hex_color or "#000000").lstrip('#')
    return tuple(int(hex_color[i:i+2],16) for i in (0,2,4))

def apply_text_style(paragraph, font_family=None, size_pt=None, bold=None, color=None):
    # ensure at least one run exists
    if not paragraph.runs:
        paragraph.add_run()
    for run in paragraph.runs:
        if font_family: run.font.name = font_family
        if size_pt: run.font.size = Pt(size_pt)
        if bold is not None: run.font.bold = bool(bold)
        if color:
            r,g,b = hex_to_rgb(color)
            run.font.color.rgb = RGBColor(r,g,b)

def ensure_image_max_width(img_path, max_width_inches=2.5):
    try:
        img = Image.open(img_path)
    except Exception:
        return img_path
    w_px, h_px = img.size
    dpi = img.info.get('dpi', (96,96))[0] or 96
    width_inches = w_px / dpi
    if width_inches <= max_width_inches:
        return img_path
    scale = max_width_inches / width_inches
    new_w = int(w_px * scale); new_h = int(h_px * scale)
    tmp_path = str(Path(img_path).with_suffix(".resized.png"))
    img.resize((new_w,new_h), Image.LANCZOS).save(tmp_path)
    return tmp_path

# ---------- Slide builders (modular) ----------
def add_title_slide(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # logos
    left_logo = slide_data.get("left_logo")
    right_logo = slide_data.get("right_logo")
    if left_logo and Path(left_logo).exists():
        slide.shapes.add_picture(str(left_logo), Inches(0.3), Inches(0.4), width=Inches(2.6))
    if right_logo and Path(right_logo).exists():
        slide.shapes.add_picture(str(right_logo), prs.slide_width - Inches(2.8), Inches(0.3), width=Inches(2.6))
    # title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), prs.slide_width - Inches(1.6), Inches(1.5))
    p = title_box.text_frame.paragraphs[0]; p.text = slide_data.get("title","")
    hcfg = style.get("typography", {}).get("heading_font", {})
    apply_text_style(p, font_family=hcfg.get("family"), size_pt=hcfg.get("size_pt", 40),
                     bold=hcfg.get("weight_bold", True), color=style.get("colors",{}).get("primary_text"))
    # subtitle
    sub_text = slide_data.get("subtitle")
    if sub_text:
        sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(3.0), prs.slide_width - Inches(1.6), Inches(1.0))
        sp = sub_box.text_frame.paragraphs[0]; sp.text = sub_text
        scfg = style.get("typography", {}).get("subheading_font", {})
        apply_text_style(sp, font_family=scfg.get("family"), size_pt=scfg.get("size_pt", 12),
                         bold=scfg.get("weight_bold", False), color=style.get("colors",{}).get("secondary_text"))
    return slide

def add_section_header_slide(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide_data.get("title","")
    info_bar = slide_data.get("info_bar")
    tbox = slide.shapes.add_textbox(Inches(0.6), Inches(0.6), prs.slide_width - Inches(1.2), Inches(1.2))
    p = tbox.text_frame.paragraphs[0]; p.text = title
    hcfg = style.get("typography", {}).get("heading_font", {})
    apply_text_style(p, font_family=hcfg.get("family"), size_pt=hcfg.get("size_pt", 36), bold=True,
                     color=style.get("colors",{}).get("primary_text"))
    if info_bar:
        bar = slide.shapes.add_textbox(Inches(0.6), Inches(1.6), prs.slide_width - Inches(1.2), Inches(0.6))
        bar.fill.solid(); bar.fill.fore_color.rgb = RGBColor(*hex_to_rgb(style["colors"].get("muted_bar_bg","#EDEEF9")))
        bp = bar.text_frame.paragraphs[0]; bp.text = info_bar
        scfg = style.get("typography", {}).get("subheading_font", {})
        apply_text_style(bp, font_family=scfg.get("family"), size_pt=scfg.get("size_pt", 12),
                         bold=False, color=style.get("colors",{}).get("primary_text"))
    return slide

def add_table_slide(prs, slide_data, style):
    cols = slide_data.get("columns", [])
    rows = slide_data.get("rows", [])
    title = slide_data.get("title")
    rcount = len(rows) + 1; ccount = len(cols)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    if title:
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), prs.slide_width - Inches(1.2), Inches(0.6))
        p = tb.text_frame.paragraphs[0]; p.text = title
        apply_text_style(p, font_family=style["typography"]["heading_font"]["family"],
                         size_pt=style["typography"].get("heading_overrides",{}).get("h2_size",28),
                         bold=True, color=style["colors"].get("primary_text"))
    left = Inches(0.6); top = Inches(1.2); width = prs.slide_width - Inches(1.2)
    table = slide.shapes.add_table(rcount, ccount, left, top, width, Inches(3.5)).table
    # header
    for c, col_name in enumerate(cols):
        cell = table.cell(0,c); cell.text = col_name
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.name = style["typography"]["body_font"]["family"]
                run.font.size = Pt(style["table"]["header"]["size_pt"])
                run.font.bold = True
                run.font.color.rgb = RGBColor(*hex_to_rgb(style["colors"]["primary_text"]))
        cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(*hex_to_rgb(style["colors"].get("table_header_bg","#F3F4F6")))
    # rows
    for r, row in enumerate(rows, start=1):
        for c in range(ccount):
            cell = table.cell(r,c)
            if isinstance(row, dict):
                key = cols[c]
                text = str(row.get(key, "")) if row.get(key) is not None else ""
            else:
                try: text = str(row[c])
                except: text = ""
            cell.text = text
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = style["typography"]["body_font"]["family"]
                    run.font.size = Pt(style["table"]["row"]["size_pt"])
                    run.font.bold = False
                    run.font.color.rgb = RGBColor(*hex_to_rgb(style["colors"]["primary_text"]))
        if style["table"].get("allow_row_strip", True) and (r % 2 == 0):
            for c in range(ccount):
                table.cell(r,c).fill.solid()
                table.cell(r,c).fill.fore_color.rgb = RGBColor(*hex_to_rgb(style["colors"]["table_row_stripe"]))
    return slide

def add_three_column_slide(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    cols = slide_data.get("columns", [])
    total_w = prs.slide_width - Inches(1.2); col_w = total_w / 3
    top = Inches(1.0); left = Inches(0.6)
    for i, col in enumerate(cols):
        x = left + i * col_w
        icon = col.get("icon")
        if icon and Path(icon).exists():
            img_path = ensure_image_max_width(icon, max_width_inches=1.2)
            slide.shapes.add_picture(img_path, x, top, width=Inches(0.9))
            text_left = x + Inches(1.0); text_width = col_w - Inches(1.0)
        else:
            text_left = x; text_width = col_w
        tb = slide.shapes.add_textbox(text_left, top, text_width, Inches(0.4))
        p = tb.text_frame.paragraphs[0]; p.text = col.get("heading","")
        apply_text_style(p, font_family=style["typography"]["body_font"]["family"],
                         size_pt=style["three_column_content"]["item_heading"]["size_pt"],
                         bold=True, color=style["colors"]["primary_text"])
        bullets = col.get("bullets",[])
        btop = top + Inches(0.5)
        tbox = slide.shapes.add_textbox(text_left, btop, text_width, Inches(3.5))
        tf = tbox.text_frame; tf.word_wrap = True
        first = True
        for b in bullets:
            if first:
                p = tf.paragraphs[0]; p.text = b; first = False
            else:
                p = tf.add_paragraph(); p.text = b
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
                         size_pt=style["typography"]["heading_overrides"].get("h2_size",28),
                         bold=True, color=style["colors"]["primary_text"])
    bullets = slide_data.get("bullets",[])
    tbox = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), prs.slide_width - Inches(1.6), Inches(4.5))
    tf = tbox.text_frame; tf.word_wrap = True
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
                     size_pt=style["typography"]["heading_overrides"].get("h3_size",20), bold=True,
                     color=style["colors"]["primary_text"])
    notes = slide_data.get("notes","")
    nbox = slide.shapes.add_textbox(Inches(0.8), Inches(1.2), prs.slide_width - Inches(1.6), Inches(4.5))
    np = nbox.text_frame.paragraphs[0]; np.text = notes
    apply_text_style(np, font_family=style["typography"]["body_font"]["family"],
                     size_pt=style["typography"]["body_font"]["size_pt"], bold=False,
                     color=style["colors"]["secondary_text"])
    return slide

# ---------- Main generator ----------
def generate_ppt(data_path: Path, style_path: Path, output_name: str):
    data = load_json(data_path)
    style = {}
    if style_path and style_path.exists():
        style = load_json(style_path)
    # fallback minimal style
    if not style:
        style = {
            "colors": {"primary_text":"#0B0B0B","secondary_text":"#6B6B6B","muted_bar_bg":"#EDEEF9",
                       "table_header_bg":"#F3F4F6","table_row_stripe":"#FAFAFB"},
            "typography": {"heading_font":{"family":"Georgia","size_pt":40,"weight_bold":True},
                           "body_font":{"family":"Calibri","size_pt":11},
                           "heading_overrides":{"h2_size":28,"h3_size":20}},
            "table":{"header":{"size_pt":11},"row":{"size_pt":11},"allow_row_strip":True},
            "three_column_content":{"item_heading":{"size_pt":14},"item_bullets":{"size_pt":11}}
        }
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_IN); prs.slide_height = Inches(SLIDE_HEIGHT_IN)

    slides = data.get("slides", [])
    for s in slides:
        t = s.get("type","bullets")
        if t == "title": add_title_slide(prs, s, style)
        elif t == "section_header": add_section_header_slide(prs, s, style)
        elif t == "table": add_table_slide(prs, s, style)
        elif t == "three_column": add_three_column_slide(prs, s, style)
        elif t == "bullets": add_bullets_slide(prs, s, style)
        elif t in ("notes_slide","notes"): add_notes_slide(prs, s, style)
        else: add_bullets_slide(prs, s, style)

    prs.save(output_name)
    print("Saved:", output_name)

# ---------- CLI ----------
def main():
    ap = argparse.ArgumentParser(description="Generate PPTX from data.json + optional style.json")
    ap.add_argument("--data", type=Path, default=DATA_PATH_DEFAULT, help="Path to data.json")
    ap.add_argument("--style", type=Path, default=STYLE_PATH_DEFAULT, help="Path to style.json")
    ap.add_argument("--out", type=str, default=OUTPUT_DEFAULT, help="Output PPTX filename")
    args = ap.parse_args()
    if not args.data.exists():
        raise SystemExit(f"Data JSON not found: {args.data}")
    generate_ppt(args.data, args.style, args.out)

if __name__ == "__main__":
    main()
