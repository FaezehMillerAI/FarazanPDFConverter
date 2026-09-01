---
name: pdf-to-word
description: Convert any PDF file (including Computer Science, STEM, arXiv, IEEE, ACM papers, brochures, and reports) to an editable Word document (.docx) while preserving 100% layout fidelity, 2-column reading order, native LaTeX math OMML equations (<m:oMath>), code callouts, tables, and figures. Works as a portable Agent Skill and local CLI tool.
---

# PDF to Editable Word (DOCX) Skill

A layout-preserving converter that transforms any PDF into a fully editable Microsoft Word document (.docx).

## Features
- **Academic & Computer Science Paper Specialist (`academic`)**: Preserves IEEE, ACM, arXiv 2-column layouts, converts LaTeX math expressions to native Word OMML equations (`<m:oMath>`), formats algorithms and code listings into shaded callout boxes, and links figure/table captions.
- **Exact Layout Mode (`exact`)**: 100% visual layout lock with transparent, absolutely positioned floating text boxes over clean background layers for brochures, posters, and flyers.
- **Standard Flow Mode (`flow`)**: Flowing Word paragraphs and headings for standard text documents.
- **Auto-Adaptive (`auto`)**: Automatically inspects the document structure and chooses the optimal conversion strategy.

## Workflow for AI Agents

1. **Pre-flight Document Inspection**:
   Inspect the PDF layout, math density, code blocks, tables, and recommended mode:

   ```bash
   python scripts/pdf2word.py inspect INPUT.pdf --json
   ```

2. **Convert Document**:
   Convert the PDF using the recommended mode (or specify `academic`, `exact`, `flow`):

   ```bash
   python scripts/pdf2word.py convert INPUT.pdf -o OUTPUT.docx --mode academic
   ```

3. **Batch Conversion**:
   Convert all PDFs in a directory:

   ```bash
   python scripts/pdf2word.py batch /path/to/pdfs/ -o /path/to/output_docx/ --mode auto
   ```

4. **Launch Web Interface**:
   Launch the interactive GUI for end users:

   ```bash
   python scripts/pdf2word.py serve --port 8000
   ```
