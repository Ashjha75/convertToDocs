#!/usr/bin/env python3
"""
json_to_ppt_fixed.py
Safer, layout-tight PPT generator tuned to Biller screenshots.
Edit DATA_PATH, STYLE_PATH, OUTPUT below or pass via CLI.
"""
import argparse, json
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image

# ---- CONFIG (edit here) ----
DATA_PATH  = Path("data.json")
STYLE_PATH = Path("style.json")
OUTPUT     = "Biller_Project_Report_fixed.pptx"
SLIDE_W_IN  = 13.333
SLIDE_H_IN  = 7.5
# ----------------------------

def load_json(path: Path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def hex_to_rgb(hex_color):
    h = (hex_color or "#000000").lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0,2,4))

def apply_run_style(run, family=None, size_pt=None, bold=None, color_hex=None):
    if family: run.font.name = family
    if size_pt: run.font.size = Pt(size_pt)
    if bold is not None: run.font.bold = bool(bold)
    if color_hex:
        r,g,b = hex_to_rgb(color_hex)
        run.font.color.rgb = RGBColor(r,g,b)

def set_paragraph_text(paragraph, text, family=None, size_pt=None, bold=None, color_hex=None):
    # ensure at least one run exists then set first run and remove others
    paragraph.clear()
    r = paragraph.add_run()
    r.text = text
    apply_run_style(r, family, size_pt, bold, color_hex)

def ensure_image(img_path, max_w_in=1.2):
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

# ------- Slide builders with tight geometry -------
def add_title_slide(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    # logos (left and right)
    left_logo = slide_data.get("left_logo")
    right_logo = slide_data.get("right_logo")
    if left_logo and Path(left_logo).exists():
        slide.shapes.add_picture(str(left_logo), Inches(0.3), Inches(0.25), width=Inches(2.6))
    if right_logo and Path(right_logo).exists():
        slide.shapes.add_picture(str(right_logo), prs.slide_width - Inches(2.9), Inches(0.18), width=Inches(2.6))

    # title — moved up (small top margin)
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), prs.slide_width - Inches(1.6), Inches(1.2))
    tf = title_box.text_frame
    set_paragraph_text(tf.paragraphs[0], slide_data.get("title",""),
                       family=style["typography"]["heading_font"]["family"],
                       size_pt=style["typography"]["heading_overrides"].get("h1_size",40),
                       bold=style["typography"]["heading_font"].get("weight_bold", True),
                       color_hex=style["colors"].get("primary_text"))
    # subtitle — small muted line under title
    subtitle = slide_data.get("subtitle")
    if subtitle:
        sub_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), prs.slide_width - Inches(1.6), Inches(0.6))
        st = sub_box.text_frame.paragraphs[0]
        set_paragraph_text(st, subtitle,
                           family=style["typography"]["subheading_font"]["family"],
                           size_pt=style["typography"]["subheading_font"].get("size_pt",14),
                           bold=False,
                           color_hex=style["colors"].get("secondary_text"))

def add_section_header(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # title near top but compact
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.5), prs.slide_width - Inches(1.2), Inches(1.0))
    set_paragraph_text(title_box.text_frame.paragraphs[0], slide_data.get("title",""),
                       family=style["typography"]["heading_font"]["family"],
                       size_pt=style["typography"]["heading_overrides"].get("h2_size",28),
                       bold=True, color_hex=style["colors"].get("primary_text"))
    # info bar: full width narrow bar directly under title
    info = slide_data.get("info_bar")
    if info:
        bar = slide.shapes.add_textbox(Inches(0.6), Inches(1.6), prs.slide_width - Inches(1.2), Inches(0.5))
        bar.fill.solid()
        r,g,b = hex_to_rgb(style["colors"].get("muted_bar_bg","#EDEEF9"))
        bar.fill.fore_color.rgb = RGBColor(r,g,b)
        p = bar.text_frame.paragraphs[0]
        set_paragraph_text(p, info,
                           family=style["typography"]["body_font"]["family"],
                           size_pt=style["typography"]["body_font"].get("size_pt",11),
                           bold=False, color_hex=style["colors"].get("primary_text"))

def add_table_slide(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide_data.get("title")
    if title:
        tbox = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), prs.slide_width - Inches(1.2), Inches(0.5))
        set_paragraph_text(tbox.text_frame.paragraphs[0], title,
                           family=style["typography"]["heading_font"]["family"],
                           size_pt=style["typography"]["heading_overrides"].get("h2_size",28),
                           bold=True, color_hex=style["colors"].get("primary_text"))
    cols = slide_data.get("columns", [])
    rows = slide_data.get("rows", [])
    if not cols: return
    rows_count = len(rows) + 1
    cols_count = len(cols)

    # Table full width, compact height per row
    left = Inches(0.6); top = Inches(1.0)
    table_width = prs.slide_width - Inches(1.2)
    # choose height: 0.6 header + 0.5 * numRows
    approx_height = Inches(0.6 + 0.5 * min(len(rows), 8))
    table = slide.shapes.add_table(rows_count, cols_count, left, top, table_width, approx_height).table

    # Header styling
    for c, name in enumerate(cols):
        cell = table.cell(0,c); cell.text = name
        for p in cell.text_frame.paragraphs:
            set_paragraph_text(p, name,
                               family=style["typography"]["body_font"]["family"],
                               size_pt=style["table"]["header"].get("size_pt",11),
                               bold=style["table"]["header"].get("weight_bold", True),
                               color_hex=style["colors"].get("primary_text"))
        cell.fill.solid()
        r,g,b = hex_to_rgb(style["colors"].get("table_header_bg"))
        cell.fill.fore_color.rgb = RGBColor(r,g,b)

    # Rows
    for r_idx, row in enumerate(rows, start=1):
        for c_idx in range(cols_count):
            cell = table.cell(r_idx, c_idx)
            try:
                text = row[c_idx] if not isinstance(row, dict) else row.get(cols[c_idx], "")
            except Exception:
                text = ""
            cell.text = str(text)
            for p in cell.text_frame.paragraphs:
                set_paragraph_text(p, str(text),
                                   family=style["typography"]["body_font"]["family"],
                                   size_pt=style["table"]["row"].get("size_pt",11),
                                   bold=False, color_hex=style["colors"].get("primary_text"))
        # stripe even rows
        if style["table"].get("allow_row_strip", True) and r_idx % 2 == 0:
            for c_idx in range(cols_count):
                table.cell(r_idx, c_idx).fill.solid()
                r,g,b = hex_to_rgb(style["colors"].get("table_row_stripe"))
                table.cell(r_idx, c_idx).fill.fore_color.rgb = RGBColor(r,g,b)

def add_three_column(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    cols = slide_data.get("columns", [])
    if not cols: return
    total_w = prs.slide_width - Inches(1.2)
    col_w = total_w / len(cols)
    left = Inches(0.6); top = Inches(0.9)

    for i, c in enumerate(cols):
        x = left + i * col_w
        icon = c.get("icon")
        text_left = x
        # icon on the left inside column
        if icon and Path(icon).exists():
            pict = ensure_image(icon, max_w_in=0.9)
            slide.shapes.add_picture(pict, x, top, width=Inches(0.8))
            text_left = x + Inches(0.9)
        # heading
        tb = slide.shapes.add_textbox(text_left, top, col_w - Inches(0.9), Inches(0.3))
        set_paragraph_text(tb.text_frame.paragraphs[0], c.get("heading",""),
                           family=style["typography"]["body_font"]["family"],
                           size_pt=style["three_column_content"]["item_heading"].get("size_pt",14),
                           bold=True, color_hex=style["colors"].get("primary_text"))
        # bullets
        btop = top + Inches(0.35)
        tbox = slide.shapes.add_textbox(text_left, btop, col_w - Inches(0.9), Inches(3.5))
        tf = tbox.text_frame
        first = True
        for b in c.get("bullets", []):
            if first:
                p = tf.paragraphs[0]; p.text = b; first = False
            else:
                p = tf.add_paragraph(); p.text = b
            p.level = 0
            set_paragraph_text(p, b,
                               family=style["typography"]["body_font"]["family"],
                               size_pt=style["three_column_content"]["item_bullets"].get("size_pt",11),
                               bold=False, color_hex=style["colors"].get("primary_text"))

def add_bullets(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide_data.get("title")
    if title:
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), prs.slide_width - Inches(1.2), Inches(0.5))
        set_paragraph_text(tb.text_frame.paragraphs[0], title,
                           family=style["typography"]["heading_font"]["family"],
                           size_pt=style["typography"]["heading_overrides"].get("h2_size",28),
                           bold=True, color_hex=style["colors"].get("primary_text"))
    # bullets block compact
    tbox = slide.shapes.add_textbox(Inches(0.8), Inches(1.05), prs.slide_width - Inches(1.6), Inches(5.5))
    tf = tbox.text_frame
    first = True
    for b in slide_data.get("bullets", []):
        if first:
            p = tf.paragraphs[0]; p.text = b; first=False
        else:
            p = tf.add_paragraph(); p.text = b
        p.level = 0
        set_paragraph_text(p, b,
                           family=style["typography"]["body_font"]["family"],
                           size_pt=style["typography"]["body_font"].get("size_pt",11),
                           bold=False, color_hex=style["colors"].get("primary_text"))

def add_notes(prs, slide_data, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide_data.get("title","Notes")
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.35), prs.slide_width - Inches(1.2), Inches(0.5))
    set_paragraph_text(tb.text_frame.paragraphs[0], title,
                       family=style["typography"]["heading_font"]["family"],
                       size_pt=style["typography"]["heading_overrides"].get("h3_size",20),
                       bold=True, color_hex=style["colors"].get("primary_text"))
    nb = slide.shapes.add_textbox(Inches(0.8), Inches(1.05), prs.slide_width - Inches(1.6), Inches(5.5))
    np = nb.text_frame.paragraphs[0]
    set_paragraph_text(np, slide_data.get("notes",""),
                      family=style["typography"]["body_font"]["family"],
                      size_pt=style["typography"]["body_font"].get("size_pt",11),
                      bold=False, color_hex=style["colors"].get("secondary_text"))

# ------- Main generator -------
def generate_ppt(data_path: Path, style_path: Path, output_name: str):
    data = load_json(data_path)
    style = {}
    if style_path.exists():
        style = load_json(style_path)
    # fallback minimal tokens
    style.setdefault("colors", {
        "primary_text":"#0B0B0B","secondary_text":"#6B6B6B","muted_bar_bg":"#EDEEF9",
        "table_header_bg":"#F3F4F6","table_row_stripe":"#FAFAFB"})
    style.setdefault("typography", {
        "heading_font":{"family":"Georgia","size_pt":40,"weight_bold":True},
        "subheading_font":{"family":"Georgia","size_pt":14},
        "body_font":{"family":"Calibri","size_pt":11},
        "heading_overrides":{"h1_size":40,"h2_size":28,"h3_size":20}})
    style.setdefault("three_column_content", {"item_heading":{"size_pt":14},"item_bullets":{"size_pt":11}})
    style.setdefault("table", {"header":{"size_pt":11},"row":{"size_pt":11},"allow_row_strip":True})

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W_IN)
    prs.slide_height = Inches(SLIDE_H_IN)

    slides = data.get("slides", [])
    # ensure one-to-one mapping (no blank pages)
    for s in slides:
        t = s.get("type","bullets")
        if t == "title":
            add_title_slide(prs, s, style)
        elif t == "section_header":
            add_section_header(prs, s, style)
        elif t == "table":
            add_table_slide(prs, s, style)
        elif t == "three_column":
            add_three_column(prs, s, style)
        elif t == "bullets":
            add_bullets(prs, s, style)
        elif t in ("notes","notes_slide"):
            add_notes(prs, s, style)
        else:
            # fallback
            add_bullets(prs, s, style)

    prs.save(output_name)
    print("Saved:", output_name)

# CLI
if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", type=Path, default=DATA_PATH)
    ap.add_argument("--style", type=Path, default=STYLE_PATH)
    ap.add_argument("--out", type=str, default=OUTPUT)
    args = ap.parse_args()
    if not args.data.exists():
        raise SystemExit("data.json not found at: " + str(args.data))
    generate_ppt(args.data, args.style, args.out)
