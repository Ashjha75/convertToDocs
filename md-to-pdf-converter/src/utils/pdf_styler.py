"""
PDF styling utilities and custom style definitions.
"""
from typing import Dict, Optional


def apply_custom_styles(base_css: str, overrides: Optional[Dict[str, str]] = None) -> str:
    """
    Apply custom style overrides to base CSS.
    
    Args:
        base_css: Base CSS template
        overrides: Dictionary of CSS property overrides
        
    Returns:
        Modified CSS string
    """
    if not overrides:
        return base_css
    
    # Apply overrides
    modified_css = base_css
    
    for selector, properties in overrides.items():
        # Simple CSS override - can be expanded for more complex scenarios
        override_rule = f"\n{selector} {{\n"
        for prop, value in properties.items():
            override_rule += f"    {prop}: {value};\n"
        override_rule += "}\n"
        
        modified_css += override_rule
    
    return modified_css


def get_color_scheme(scheme: str = 'default') -> Dict[str, str]:
    """
    Get predefined color schemes for PDFs.
    
    Args:
        scheme: Color scheme name ('default', 'blue', 'green', 'corporate')
        
    Returns:
        Dictionary of color values
    """
    schemes = {
        'default': {
            'primary': '#2563eb',
            'secondary': '#3b82f6',
            'text': '#1a1a1a',
            'background': '#ffffff',
            'code_bg': '#1e1e1e',
        },
        'blue': {
            'primary': '#0066cc',
            'secondary': '#3399ff',
            'text': '#000000',
            'background': '#ffffff',
            'code_bg': '#002b4d',
        },
        'green': {
            'primary': '#059669',
            'secondary': '#10b981',
            'text': '#1a1a1a',
            'background': '#ffffff',
            'code_bg': '#064e3b',
        },
        'corporate': {
            'primary': '#1e3a8a',
            'secondary': '#3b82f6',
            'text': '#0f172a',
            'background': '#ffffff',
            'code_bg': '#1e293b',
        },
    }
    
    return schemes.get(scheme, schemes['default'])


def generate_custom_css(
    font_family: str = 'Calibri, Arial, sans-serif',
    font_size: str = '11pt',
    color_scheme: str = 'default'
) -> str:
    """
    Generate custom CSS based on parameters.
    
    Args:
        font_family: Base font family
        font_size: Base font size
        color_scheme: Color scheme name
        
    Returns:
        Custom CSS string
    """
    colors = get_color_scheme(color_scheme)
    
    return f"""
body {{
    font-family: {font_family};
    font-size: {font_size};
    color: {colors['text']};
}}

h1, h2 {{
    color: {colors['primary']};
}}

h3, h4 {{
    color: {colors['secondary']};
}}

a {{
    color: {colors['primary']};
}}

pre {{
    background-color: {colors['code_bg']};
}}
"""
