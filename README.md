# 📄 OmniPDF: Layout-Preserving PDF to Word (DOCX) Converter

> **High-Fidelity PDF to Word converter** designed for **Computer Science papers**, STEM publications (arXiv, IEEE, ACM, Springer), brochures, forms, and reports — preserving exact visual layouts, 2-column reading order, native editable Word LaTeX math equations (`<m:oMath>`), algorithms, and tables without altering the format.

---

## 🌟 Key Features

- 🎓 **Academic & Computer Science Specialist**:
  - Accurately splits and reconstructs **2-column layouts** without interleaving left and right text streams.
  - Converts **LaTeX math formulas and symbols** into native Microsoft Word OMML equations (`<m:oMath>`), making them directly editable in Word.
  - Recognizes **pseudocode, algorithms, and code listings**, wrapping them in shaded callout boxes with preserved monospace fonts and indentation.
  - Preserves **tables with borders / LaTeX booktabs rules**, cell alignments, and headers.
  - Extracts and embeds high-resolution figures (300 DPI) linked to their captions.
  - Formats **References & Bibliographies** with proper hanging indents (`[1]`, `[2]`).
- 🎯 **Exact Layout Mode (Pixel-Perfect Lock)**:
  - Transparent, floating, absolutely positioned Word text boxes (`w:txbx`) over clean high-resolution page layers.
  - 100% visual fidelity for flyers, complex graphics, forms, and presentations.
- 📄 **Standard Flow Mode**:
  - Reconstructs flowing Word paragraphs and headings for standard text documents and books.
- ⚡ **Auto-Adaptive Intelligence**:
  - Pre-flight inspects PDF font density, column distribution, math presence, and automatically selects the optimal conversion engine.
- 🌐 **Modern Web User Interface**:
  - Interactive drag & drop web application with instant pre-flight analysis, mode selection, progress bar, and one-click download.
- 🤖 **AI Agent Skill Compatible**:
  - Compatible with Antigravity, Claude Code, Codex, and other AI agent workflows via `skills/pdf-to-word/SKILL.md`.

---

## 🚀 Quick Start

### 1. Installation

```bash
# Clone or navigate to the directory
cd /path/to/PDFEDITOR

# Create and activate virtual environment
python3 -m venv venv
source venv/bin/activate

# Install dependencies
pip install -e .
```

---

## 💻 Usage

### Command Line Interface (CLI)

```bash
# 1. Convert a Computer Science paper (Academic Mode)
python scripts/pdf2word.py convert paper.pdf -o paper.docx --mode academic

# 2. Convert with Exact Layout Lock (Pixel-Perfect)
python scripts/pdf2word.py convert brochure.pdf -o brochure.docx --mode exact

# 3. Auto-detect optimal mode
python scripts/pdf2word.py convert document.pdf -o document.docx --mode auto

# 4. Inspect PDF layout, columns, formulas, and compatibility before conversion
python scripts/pdf2word.py inspect paper.pdf

# 5. Batch convert an entire folder of PDFs
python scripts/pdf2word.py batch ./my_pdfs/ -o ./converted_docs/ --mode auto
```

### 🌐 Web User Interface

Launch the local web application:

```bash
python scripts/pdf2word.py serve --port 8000
```
Then open [http://127.0.0.1:8000](http://127.0.0.1:8000) in your browser.

---

## 🐍 Python API Usage

```python
from omnipdf import OmniConverter, ConversionMode, inspect_pdf

# 1. Inspect PDF structure
report = inspect_pdf("paper.pdf")
print("Detected Columns:", report["detected_columns"])
print("Math Density:", report["math_density"])
print("Recommended Mode:", report["recommended_mode"])

# 2. Convert PDF to Word
converter = OmniConverter(
    convert_math=True,      # Native Word OMML equations
    highlight_code=True,    # Monospace callout boxes
    extract_images=True,    # High-resolution figures
    image_dpi=300           # 300 DPI clarity
)

result = converter.convert(
    pdf_path="paper.pdf",
    docx_path="paper.docx",
    mode=ConversionMode.ACADEMIC
)

print(f"Converted in {result['duration_seconds']}s -> {result['output_file']}")
```

---

## 📁 Project Structure

```
PDFEDITOR/
├── src/
│   └── omnipdf/
│       ├── __init__.py
│       ├── core/
│       │   ├── converter.py          # Master conversion orchestrator
│       │   ├── inspector.py          # Pre-flight PDF layout & font analyzer
│       │   ├── layout_detector.py    # Geometric column & reading order parser
│       │   ├── math_engine.py        # LaTeX / Unicode to native Word OMML (<m:oMath>)
│       │   ├── code_detector.py      # Monospace algorithm & code block extractor
│       │   ├── table_extractor.py    # Bordered & borderless table builder
│       │   ├── academic_converter.py # CS & Academic Paper specialist
│       │   ├── exact_converter.py    # Pixel-perfect absolute layout converter
│       │   └── flow_converter.py     # Standard flowing text converter
│       ├── cli/
│       │   └── main.py               # Typer & Rich CLI
│       └── web/
│           ├── app.py                # FastAPI backend
│           └── templates/
│               └── index.html        # Modern drag & drop UI
├── skills/
│   └── pdf-to-word/
│       └── SKILL.md                  # Portable Agent Skill definition
├── scripts/
│   └── pdf2word.py                   # CLI wrapper script
├── tests/
│   └── test_converter.py             # Comprehensive test suite
├── pyproject.toml
└── README.md
```

---

## 📄 License
MIT License.
