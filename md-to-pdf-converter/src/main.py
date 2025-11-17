#!/usr/bin/env python3
"""
Main entry point for MD to PDF converter.
Professional Markdown to PDF converter for Windows.
"""
import argparse
import os
import sys
from pathlib import Path
from typing import Optional
from converters.md_to_pdf import MdToPdfConverter


def validate_input_file(file_path: str) -> bool:
    """
    Validate input markdown file.
    
    Args:
        file_path: Path to markdown file
        
    Returns:
        True if valid, False otherwise
    """
    if not os.path.isfile(file_path):
        print(f"❌ Error: The file '{file_path}' does not exist.")
        return False
    
    if not file_path.lower().endswith(('.md', '.markdown', '.txt')):
        print(f"⚠️  Warning: '{file_path}' may not be a markdown file.")
    
    return True


def create_output_directory(output_path: str) -> None:
    """
    Create output directory if it doesn't exist.
    
    Args:
        output_path: Path to output PDF file
    """
    output_dir = os.path.dirname(output_path)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)


def format_file_size(size_bytes: int) -> str:
    """
    Format file size in human-readable format.
    
    Args:
        size_bytes: Size in bytes
        
    Returns:
        Formatted string (e.g., "1.5 MB")
    """
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} TB"


def main() -> None:
    """Main function to handle CLI arguments and conversion."""
    parser = argparse.ArgumentParser(
        description="Convert Markdown files to high-quality professional PDFs.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python main.py input.md output.pdf
  python main.py README.md docs/README.pdf -v
  python main.py notes.md report.pdf --verbose

Supported Markdown Features:
  • Headings (H1-H6)
  • Bold, italic, strikethrough
  • Lists (ordered and unordered)
  • Tables with styling
  • Code blocks with syntax highlighting
  • Blockquotes
  • Images
  • Links
  • Horizontal rules

For more information, visit: https://github.com/yourusername/md-to-pdf-converter
        """
    )
    
    parser.add_argument(
        "input",
        help="Path to the input Markdown file (.md, .markdown, .txt)"
    )
    parser.add_argument(
        "output",
        help="Path to save the output PDF file"
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose output with detailed information"
    )
    
    args = parser.parse_args()

    # Validate input file
    if not validate_input_file(args.input):
        sys.exit(1)
    
    # Ensure output has .pdf extension
    if not args.output.lower().endswith('.pdf'):
        args.output += '.pdf'
    
    # Create output directory if needed
    try:
        create_output_directory(args.output)
        if args.verbose and os.path.dirname(args.output):
            print(f"📁 Output directory: {os.path.dirname(args.output)}")
    except Exception as e:
        print(f"❌ Error creating output directory: {e}")
        sys.exit(1)

    # Convert MD to PDF
    converter = MdToPdfConverter()
    
    try:
        if args.verbose:
            print("=" * 60)
            print("📄 Markdown to PDF Converter")
            print("=" * 60)
            print(f"📥 Input:  {os.path.abspath(args.input)}")
            print(f"📤 Output: {os.path.abspath(args.output)}")
            print(f"⏳ Converting...")
        
        # Perform conversion
        converter.convert(args.input, args.output)
        
        # Success message
        print(f"✅ Successfully converted '{args.input}' to '{args.output}'")
        
        # Show file statistics
        if os.path.exists(args.output):
            file_size = os.path.getsize(args.output)
            print(f"📦 PDF size: {format_file_size(file_size)}")
            
            if args.verbose:
                print(f"📍 Full path: {os.path.abspath(args.output)}")
                print("=" * 60)
        
    except FileNotFoundError as e:
        print(f"❌ File not found: {e}")
        sys.exit(1)
    except IOError as e:
        print(f"❌ I/O error: {e}")
        sys.exit(1)
    except Exception as e:
        print(f"❌ An error occurred during conversion: {e}")
        if args.verbose:
            import traceback
            print("\n🔍 Detailed error information:")
            traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
