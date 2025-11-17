"""
Markdown to PDF converter using WeasyPrint for high-quality output.
Optimized for Windows with professional styling.
"""
import os
from pathlib import Path
from typing import Optional
import markdown
from weasyprint import HTML, CSS
from pygments.formatters import HtmlFormatter
from ..config.settings import PDF_CONFIG, CSS_TEMPLATE
from ..utils.markdown_parser import parse_markdown_with_code


class MdToPdfConverter:
    """Convert Markdown files to professional PDF documents."""
    
    def __init__(self, custom_css: Optional[str] = None):
        """
        Initialize the converter.
        
        Args:
            custom_css: Optional custom CSS string to override default styles
        """
        self.custom_css = custom_css or CSS_TEMPLATE
        
    def convert(self, input_path: str, output_path: str) -> None:
        """
        Convert a Markdown file to PDF.
        
        Args:
            input_path: Path to input .md file
            output_path: Path to output .pdf file
            
        Raises:
            FileNotFoundError: If input file doesn't exist
            IOError: If unable to write output file
        """
        # Validate input
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Input file not found: {input_path}")
        
        # Read markdown content
        with open(input_path, 'r', encoding='utf-8') as f:
            md_content = f.read()
        
        # Parse markdown to HTML with syntax highlighting
        html_content = parse_markdown_with_code(md_content)
        
        # Wrap in complete HTML document
        full_html = self._wrap_html(html_content, input_path)
        
        # Get Pygments CSS for code highlighting (VS Code Dark+ style)
        pygments_css = HtmlFormatter(style='monokai').get_style_defs('.codehilite')
        
        # Combine all CSS
        combined_css = f"{self.custom_css}\n{pygments_css}"
        
        # Convert to PDF with high quality settings
        try:
            HTML(string=full_html).write_pdf(
                output_path,
                stylesheets=[CSS(string=combined_css)],
                **PDF_CONFIG
            )
        except Exception as e:
            raise IOError(f"Failed to write PDF: {e}")
    
    def _wrap_html(self, body_content: str, source_file: str) -> str:
        """
        Wrap HTML content in a complete document structure.
        
        Args:
            body_content: HTML content from markdown
            source_file: Original markdown filename for title
            
        Returns:
            Complete HTML document string
        """
        title = Path(source_file).stem.replace('_', ' ').replace('-', ' ').title()
        
        return f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="generator" content="MD to PDF Converter">
    <title>{title}</title>
</head>
<body>
    {body_content}
</body>
</html>
"""
