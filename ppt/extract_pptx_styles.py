#!/usr/bin/env python3
# extract_pptx_styles.py
# Usage: python extract_pptx_styles.py "presentation.pptx"
# Produces ./pptx_styles.json

import sys, zipfile, json, xml.etree.ElementTree as ET, re, os

def srgb_from_elem(elem, ns):
    srgb = elem.find('.//a:srgbClr', ns)
    if srgb is not None and 'val' in srgb.attrib:
        return '#'+srgb.attrib['val']
    sysc = elem.find('.//a:sysClr', ns)
    if sysc is not None and 'lastClr' in sysc.attrib:
        return '#'+sysc.attrib['lastClr']
    return None

def parse_theme_colors(z, path):
    xml = z.read(path).decode('utf-8')
    root = ET.fromstring(xml)
    ns = {'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    cs = root.find('.//a:clrScheme', ns)
    out = {}
    if cs is None:
        return out
    for child in cs:
        tag = re.sub(r'\{.*\}', '', child.tag)
        name = tag.split('}')[-1]
        val = srgb_from_elem(child, ns)
        out[name] = val
    return out

def find_table_colors_in_slides(z):
    ns = {'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    found = set()
    for name in z.namelist():
        if name.startswith('ppt/slides/') or name.startswith('ppt/slideLayouts/') or name.startswith('ppt/slideMasters/'):
            data = z.read(name).decode('utf-8', errors='ignore')
            # quick regex for hex RGB in srgbClr val="RRGGBB"
            for m in re.finditer(r'srgbClr[^>]*val=[\'"]([0-9A-Fa-f]{6})', data):
                found.add('#'+m.group(1))
            # also look for sysClr lastClr
            for m in re.finditer(r'lastClr=[\'"]([0-9A-Fa-f]{6})', data):
                found.add('#'+m.group(1))
    return sorted(found)

def main(pptx_path):
    if not os.path.exists(pptx_path):
        print("File not found:", pptx_path); sys.exit(1)
    z = zipfile.ZipFile(pptx_path)
    themes = [n for n in z.namelist() if n.startswith('ppt/theme/') and n.endswith('.xml')]
    theme_colors = {}
    for t in themes:
        theme_colors[t] = parse_theme_colors(z, t)
    table_related_colors = find_table_colors_in_slides(z)

    out = {
        'theme_files': themes,
        'theme_colors': theme_colors,
        'table_related_colors_found_in_slides': table_related_colors
    }
    with open('pptx_styles.json','w') as f:
        json.dump(out,f,indent=2)
    print("Wrote pptx_styles.json with theme colors and table-related colors.")
    print("Theme files:", themes)

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python extract_pptx_styles.py <file.pptx>"); sys.exit(1)
    main(sys.argv[1])
