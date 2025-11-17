# 📄 Professional Markdown to PDF Converter

Convert your Markdown files into beautiful, high-quality PDFs with syntax highlighting and professional styling.

## ✨ Features

- 🎨 Professional PDF output with custom CSS
- 💻 Syntax highlighting for 100+ programming languages
- 📊 Beautiful tables, lists, and images
- 📝 Page numbers and headers
- 🔧 Fully customizable fonts and colors
- 🖥️ Cross-platform support (Windows, Linux, macOS)

## 🚀 Quick Start

```bash
# Clone the repository
git clone https://github.com/yourusername/md-to-pdf-converter.git
cd md-to-pdf-converter

# Install dependencies
pip install -r requirements.txt

# Convert your first markdown file
cd src
python main.py input.md output.pdf
```

## 📦 Installation

**Windows:**
```bash
pip install -r requirements.txt
```

**Linux/WSL:**
```bash
sudo apt-get install -y libpango-1.0-0 libpangoft2-1.0-0
pip install -r requirements.txt
```

## 💡 Usage

```bash
# Basic conversion
python main.py document.md output.pdf

# With verbose output
python main.py document.md output.pdf -v
```

## 📝 Supported Features

- Headings, bold, italic, strikethrough
- Code blocks with syntax highlighting
- Tables with professional styling
- Lists (ordered/unordered)
- Images, links, blockquotes
- And more!

## 🎨 Customization

Edit `src/config/settings.py` to customize fonts, colors, spacing, and CSS styles.

## 📄 License

MIT License

## 🤝 Contributing

Contributions, issues, and feature requests are welcome!
