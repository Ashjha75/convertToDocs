#!/usr/bin/env python3
# json_to_ppt.py (FINAL SAFE VERSION)

import argparse, json
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from PIL import Image

# ----------------- DEFAULT CONFIG -----------------
DATA_PATH_DEFAULT  = Path("data.json")
STYLE_PATH_DEFAULT = Path("style.json")
OUTPUT_DEFAULT     = "Generated_Presentation.pptx"

SLIDE_WIDTH_IN  = 13.333
SLIDE_HEIGHT_IN = 7.5
# -------------------------------------------------

# ---------- HELPERS ----------
def load_json(path: Path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def hex_to_rgb(h):
    h = h.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0,2,4))

def safe_dict(d):
    return d if isinstance(d, dict) else {}

def safe_get(style, key, default):
    v = style.get(key)
    return v if isinstance(v, dict) else default

def apply_text(par, family=None, size=None, bold=None, color=None):
    if not par.runs:
        par.add_run()
    for r in par.runs:
        if family: r.font.name = family
        if size:   r.font.size = Pt(size)
        if bold is not None: r.font.bold = bool(bold)
        if color:
            r.font.color.rgb = RGBColor(*hex_to_rgb(color))

def ensure_img(img_path, max_w=2.5):
    try:
        img = Image.open(img_path)
    except:
        return img_path
    w, h = img.size
    dpi = img.info.get("dpi", (96,96))[0] or 96
    w_in = w/dpi
    if w_in <= max_w:
        return img_path
    scale = max_w/w_in
    new_w, new_h = int(w*scale), int(h*scale)
    out = Path(img_path).with_suffix(".resized.png")
    img.resize((new_w,new_h), Image.LANCZOS).save(out)
    return str(out)

# ---------- SLIDES ----------
def add_title(prs, s, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    hcfg = safe_get(style.get("typography", {}), "heading_font", {"family":"Calibri","size_pt":36,"weight_bold":True})
    scfg = safe_get(style.get("typography", {}), "subheading_font", {"family":"Calibri","size_pt":14})

    # Logos
    if s.get("left_logo") and Path(s["left_logo"]).exists():
        slide.shapes.add_picture(s["left_logo"], Inches(0.3), Inches(0.4), width=Inches(2.6))
    if s.get("right_logo") and Path(s["right_logo"]).exists():
        slide.shapes.add_picture(s["right_logo"], prs.slide_width - Inches(2.8), Inches(0.3), width=Inches(2.6))

    # Title
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.8), prs.slide_width - Inches(1.6), Inches(1.5))
    p = tb.text_frame.paragraphs[0]
    p.text = s.get("title","")
    apply_text(p, hcfg["family"], hcfg["size_pt"], hcfg["weight_bold"], 
               style.get("colors",{}).get("primary_text","#000"))

    # Subtitle
    if s.get("subtitle"):
        sb = slide.shapes.add_textbox(Inches(0.8), Inches(3.0), prs.slide_width - Inches(1.6), Inches(1))
        sp = sb.text_frame.paragraphs[0]
        sp.text = s["subtitle"]
        apply_text(sp, scfg["family"], scfg["size_pt"], False,
                   style.get("colors",{}).get("secondary_text","#666"))

def add_section(prs, s, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    hcfg = safe_get(style["typography"], "heading_font", {"family":"Calibri","size_pt":36})
    color_primary = style["colors"].get("primary_text","#000")

    # Title
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.6), prs.slide_width - Inches(1.2), Inches(1.2))
    p = tb.text_frame.paragraphs[0]
    p.text = s.get("title","")
    apply_text(p, hcfg["family"], hcfg["size_pt"], True, color_primary)

    # Info bar
    if s.get("info_bar"):
        bar = slide.shapes.add_textbox(Inches(0.6), Inches(1.6), prs.slide_width - Inches(1.2), Inches(0.6))
        bar.fill.solid()
        bar.fill.fore_color.rgb = RGBColor(*hex_to_rgb(style["colors"].get("muted_bar_bg","#EEE")))
        bp = bar.text_frame.paragraphs[0]
        bp.text = s["info_bar"]
        apply_text(bp, hcfg["family"], 14, False, color_primary)

def add_table(prs, s, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    cols = s.get("columns", [])
    rows = s.get("rows", [])
    h2 = style["typography"]["heading_overrides"].get("h2_size",28)

    # Title
    if s.get("title"):
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), prs.slide_width - Inches(1.2), Inches(0.6))
        p = tb.text_frame.paragraphs[0]
        p.text = s["title"]
        apply_text(p, style["typography"]["heading_font"]["family"], h2, True,
                   style["colors"]["primary_text"])

    # Table
    rcount = len(rows)+1
    ccount = len(cols)
    left = Inches(0.6); top = Inches(1.2)
    width = prs.slide_width - Inches(1.2)

    table = slide.shapes.add_table(rcount, ccount, left, top, width, Inches(4)).table

    # header
    for c,name in enumerate(cols):
        cell = table.cell(0,c); cell.text = name
        for p in cell.text_frame.paragraphs:
            apply_text(p, style["typography"]["body_font"]["family"],
                       style["table"]["header"]["size_pt"], True,
                       style["colors"]["primary_text"])
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(*hex_to_rgb(style["colors"]["table_header_bg"]))

    # rows
    for r,row in enumerate(rows, start=1):
        for c in range(ccount):
            cell = table.cell(r,c)
            try: cell.text = str(row[c])
            except: cell.text = ""
            for p in cell.text_frame.paragraphs:
                apply_text(p, style["typography"]["body_font"]["family"],
                           style["table"]["row"]["size_pt"], False,
                           style["colors"]["primary_text"])

        if style["table"].get("allow_row_strip",True) and r%2==0:
            for c in range(ccount):
                table.cell(r,c).fill.solid()
                table.cell(r,c).fill.fore_color.rgb = RGBColor(*hex_to_rgb(style["colors"]["table_row_stripe"]))

def add_three(prs, s, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # SAFE fallback
    three = safe_get(style, "three_column_content", {
        "columns": 3,
        "item_heading": {"size_pt": 14},
        "item_bullets": {"size_pt": 11}
    })

    cols = s.get("columns", [])
    if not isinstance(cols, list):
        cols = []

    total = prs.slide_width - Inches(1.2)
    col_w = total / max(1, len(cols))
    left = Inches(0.6)
    top = Inches(1.0)

    for i,col in enumerate(cols):
        x = left + i * col_w
        heading = col.get("heading","")
        bullets = col.get("bullets",[])

        # Icon
        icon = col.get("icon")
        if icon and Path(icon).exists():
            icon_p = ensure_img(icon, 1.3)
            slide.shapes.add_picture(icon_p, x, top, width=Inches(1.0))
            text_left = x + Inches(1.2)
        else:
            text_left = x

        # Heading
        tb = slide.shapes.add_textbox(text_left, top, col_w - Inches(0.3), Inches(0.5))
        p = tb.text_frame.paragraphs[0]
        p.text = heading
        apply_text(p, style["typography"]["body_font"]["family"],
                   three["item_heading"]["size_pt"], True,
                   style["colors"]["primary_text"])

        # Bullets
        tbox = slide.shapes.add_textbox(text_left, top+Inches(0.6), col_w-Inches(0.3), Inches(3.5))
        tf = tbox.text_frame
        tf.word_wrap = True

        first = True
        for b in bullets:
            if first:
                pp = tf.paragraphs[0]
                pp.text = b
                first = False
            else:
                pp = tf.add_paragraph()
                pp.text = b
            pp.level = 0
            apply_text(pp, style["typography"]["body_font"]["family"],
                       three["item_bullets"]["size_pt"], False,
                       style["colors"]["primary_text"])

def add_bullets(prs, s, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    h2 = style["typography"]["heading_overrides"]["h2_size"]
    # Title
    if s.get("title"):
        tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.4),
                                      prs.slide_width - Inches(1.2), Inches(0.6))
        p = tb.text_frame.paragraphs[0]
        p.text = s["title"]
        apply_text(p, style["typography"]["heading_font"]["family"], h2, True,
                   style["colors"]["primary_text"])

    # Bullets
    bl = s.get("bullets", [])
    box = slide.shapes.add_textbox(Inches(0.8), Inches(1.2),
                                   prs.slide_width - Inches(1.6), Inches(4.5))
    tf = box.text_frame
    tf.word_wrap = True

    first = True
    for b in bl:
        if first:
            p = tf.paragraphs[0]; p.text = b; first=False
        else:
            p = tf.add_paragraph(); p.text = b
        p.level = 0
        apply_text(p, style["typography"]["body_font"]["family"],
                   style["typography"]["body_font"]["size_pt"],
                   False, style["colors"]["primary_text"])

def add_notes(prs, s, style):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    h3 = style["typography"]["heading_overrides"]["h3_size"]

    # Title
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.4),
                                  prs.slide_width - Inches(1.2), Inches(0.6))
    p = tb.text_frame.paragraphs[0]
    p.text = s.get("title","Notes")
    apply_text(p, style["typography"]["heading_font"]["family"], h3, True,
               style["colors"]["primary_text"])

    # Notes
    nb = slide.shapes.add_textbox(Inches(0.8), Inches(1.2),
                                  prs.slide_width - Inches(1.6), Inches(4.5))
    np = nb.text_frame.paragraphs[0]
    np.text = s.get("notes","")
    apply_text(np, style["typography"]["body_font"]["family"],
               style["typography"]["body_font"]["size_pt"],
               False, style["colors"]["secondary_text"])

# ---------- MAIN ----------
def generate_ppt(data_path, style_path, output):
    data = load_json(data_path)

    # STYLE SAFE LOAD
    if style_path.exists():
        style = load_json(style_path)
    else:
        style = {}

    # fallback style enforced
    style.setdefault("colors", {})
    style.setdefault("typography", {})
    style.setdefault("three_column_content", {})
    style.setdefault("table", {"header":{"size_pt":11}, "row":{"size_pt":11}})

    prs = Presentation()
    prs.slide_width = Inches(SLIDE_WIDTH_IN)
    prs.slide_height = Inches(SLIDE_HEIGHT_IN)

    for s in data.get("slides", []):
        t = s.get("type", "bullets")

        if t == "title": add_title(prs, s, style)
        elif t == "section_header": add_section(prs, s, style)
        elif t == "table": add_table(prs, s, style)
        elif t == "three_column": add_three(prs, s, style)
        elif t == "bullets": add_bullets(prs, s, style)
        elif t in ("notes","notes_slide"): add_notes(prs, s, style)
        else:
            add_bullets(prs, s, style)

    prs.save(output)
    print("Saved:", output)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", type=Path, default=DATA_PATH_DEFAULT)
    ap.add_argument("--style", type=Path, default=STYLE_PATH_DEFAULT)
    ap.add_argument("--out", type=str, default=OUTPUT_DEFAULT)
    args = ap.parse_args()

    if not args.data.exists():
        raise SystemExit("data.json not found!")
    generate_ppt(args.data, args.style, args.out)

if __name__ == "__main__":
    main()
