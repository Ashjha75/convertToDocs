"""
Configuration settings for PDF generation.
Optimized for Windows with professional styling.
"""

# PDF Generation Settings
PDF_CONFIG = {
    'presentational_hints': True,
    'optimize_images': True,
}

# Font Configuration (Windows-friendly fonts)
FONT_CONFIG = {
    'body_font': 'Calibri, Arial, sans-serif',
    'heading_font': 'Calibri, Arial, sans-serif',
    'code_font': 'Consolas, "Courier New", monospace',
    'base_size': '11pt',
}

# Professional CSS Template - High Quality Output
CSS_TEMPLATE = """
@page {
    size: A4;
    margin: 2.5cm 2cm;
    
    @top-right {
        content: "Page " counter(page) " of " counter(pages);
        font-size: 9pt;
        color: #666;
        font-family: Calibri, Arial, sans-serif;
    }
}

body {
    font-family: Calibri, Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.6;
    color: #1a1a1a;
    max-width: 100%;
}

/* Headings - Professional styling */
h1, h2, h3, h4, h5, h6 {
    font-family: Calibri, Arial, sans-serif;
    font-weight: bold;
    margin-top: 1.5em;
    margin-bottom: 0.5em;
    color: #1a1a1a;
    page-break-after: avoid;
}

h1 {
    font-size: 28pt;
    border-bottom: 3px solid #2563eb;
    padding-bottom: 0.3em;
    margin-top: 0;
    color: #1e40af;
}

h2 {
    font-size: 22pt;
    border-bottom: 2px solid #3b82f6;
    padding-bottom: 0.2em;
    color: #1e40af;
}

h3 {
    font-size: 18pt;
    color: #2563eb;
}

h4 {
    font-size: 14pt;
    color: #3b82f6;
}

h5 {
    font-size: 12pt;
    color: #60a5fa;
}

h6 {
    font-size: 11pt;
    color: #60a5fa;
}

/* Paragraphs */
p {
    margin: 0.8em 0;
    text-align: justify;
    orphans: 3;
    widows: 3;
}

/* Links */
a {
    color: #2563eb;
    text-decoration: none;
    font-weight: 500;
}

a:hover {
    text-decoration: underline;
}

/* Lists - Enhanced styling */
ul, ol {
    margin: 0.8em 0;
    padding-left: 2em;
}

li {
    margin: 0.4em 0;
    line-height: 1.5;
}

ul li {
    list-style-type: disc;
}

ul ul li {
    list-style-type: circle;
}

ol li {
    list-style-type: decimal;
}

/* Code blocks - Professional dark theme */
pre {
    background-color: #1e1e1e;
    border-radius: 8px;
    padding: 1.2em;
    overflow-x: auto;
    margin: 1.2em 0;
    page-break-inside: avoid;
    border-left: 4px solid #2563eb;
    box-shadow: 0 2px 4px rgba(0,0,0,0.1);
}

code {
    font-family: Consolas, "Courier New", monospace;
    font-size: 9.5pt;
    background-color: #f3f4f6;
    padding: 0.2em 0.4em;
    border-radius: 4px;
    color: #dc2626;
}

pre code {
    background-color: transparent;
    padding: 0;
    color: #d4d4d4;
    border-radius: 0;
}

/* Tables - Professional styling */
table {
    border-collapse: collapse;
    width: 100%;
    margin: 1.2em 0;
    page-break-inside: avoid;
    font-size: 10pt;
    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
}

th, td {
    border: 1px solid #d1d5db;
    padding: 0.7em 0.9em;
    text-align: left;
}

th {
    background-color: #2563eb;
    color: white;
    font-weight: bold;
    text-transform: uppercase;
    font-size: 9pt;
    letter-spacing: 0.5px;
}

tr:nth-child(even) {
    background-color: #f9fafb;
}

tr:hover {
    background-color: #eff6ff;
}

/* Blockquotes - Enhanced design */
blockquote {
    border-left: 5px solid #2563eb;
    padding-left: 1.2em;
    margin: 1.2em 0;
    color: #4b5563;
    font-style: italic;
    background-color: #f0f9ff;
    padding: 1em 1.2em;
    border-radius: 0 8px 8px 0;
    box-shadow: 0 2px 4px rgba(0,0,0,0.05);
}

blockquote p {
    margin: 0.5em 0;
}

/* Horizontal rule */
hr {
    border: none;
    border-top: 2px solid #e5e7eb;
    margin: 2em 0;
}

/* Images - Centered with shadow */
img {
    max-width: 100%;
    height: auto;
    display: block;
    margin: 1.5em auto;
    border-radius: 8px;
    box-shadow: 0 4px 6px rgba(0,0,0,0.1);
}

/* Strong and emphasis */
strong, b {
    font-weight: bold;
    color: #1a1a1a;
}

em, i {
    font-style: italic;
    color: #374151;
}

/* Syntax highlighting overrides */
.codehilite {
    background-color: #1e1e1e;
    border-radius: 8px;
    padding: 1.2em;
}

.codehilite pre {
    background-color: transparent;
    padding: 0;
    margin: 0;
    border: none;
    box-shadow: none;
}

/* Page breaks */
.page-break {
    page-break-after: always;
}

/* Print optimizations */
@media print {
    body {
        font-size: 10pt;
    }
    
    a {
        color: #1a1a1a;
        text-decoration: none;
    }
    
    pre, blockquote, table {
        page-break-inside: avoid;
    }
}
"""
