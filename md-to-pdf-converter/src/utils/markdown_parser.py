"""
Markdown parsing utilities with code highlighting support.
Enhanced for professional document conversion.
"""
import markdown
from markdown.extensions.codehilite import CodeHiliteExtension
from markdown.extensions.fenced_code import FencedCodeExtension
from markdown.extensions.tables import TableExtension
from markdown.extensions.toc import TocExtension
from typing import List


def parse_markdown_with_code(md_content: str) -> str:
    """
    Parse Markdown content to HTML with syntax highlighting.
    
    Args:
        md_content: Raw markdown string
        
    Returns:
        HTML string with syntax-highlighted code blocks and enhanced formatting
    """
    md = markdown.Markdown(
        extensions=[
            FencedCodeExtension(),
            CodeHiliteExtension(
                linenums=False,
                css_class='codehilite',
                guess_lang=True,
                pygments_style='monokai'
            ),
            TableExtension(),
            TocExtension(
                toc_depth='2-6',
                title='Table of Contents'
            ),
            'extra',          # Includes: abbr, attr_list, def_list, fenced_code, footnotes, tables
            'nl2br',          # New line to break
            'sane_lists',     # Better list handling
            'smarty',         # Smart quotes and dashes
            'meta',           # Metadata support
        ],
        output_format='html5'
    )
    
    return md.convert(md_content)


def get_supported_languages() -> List[str]:
    """
    Get list of supported programming languages for syntax highlighting.
    
    Returns:
        List of language identifiers supported by Pygments
    """
    from pygments.lexers import get_all_lexers
    
    languages = []
    for lexer in get_all_lexers():
        languages.extend(lexer[1])
    
    return sorted(set(languages))
