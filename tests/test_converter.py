"""
Comprehensive test suite for OmniPDF: tests PDF inspector, layout detection,
math OMML conversion, code boxes, table extraction, and end-to-end PDF-to-DOCX conversion.
"""

import os
import unittest
import tempfile
import fitz
import docx

from omnipdf.core.inspector import PDFInspector
from omnipdf.core.math_engine import MathEngine
from omnipdf.core.code_detector import CodeDetector
from omnipdf.core.table_extractor import TableExtractor
from omnipdf.core.academic_converter import AcademicConverter
from omnipdf.core.exact_converter import ExactLayoutConverter
from omnipdf.core.converter import OmniConverter, ConversionMode


def create_sample_cs_paper_pdf(file_path: str):
    """Generate a realistic 2-column Computer Science paper PDF with math, code, and tables."""
    doc = fitz.open()
    page = doc.new_page(width=612, height=792)  # Standard US Letter (8.5 x 11 in)

    # 1. Title (Top Full-width, centered)
    page.insert_text((80, 50), "Deep Learning for High-Fidelity Document Processing", fontsize=16, fontname="helv")
    
    # 2. Authors
    page.insert_text((160, 75), "Alice Smith, Bob Jones — Dept. of Computer Science", fontsize=10, fontname="helv")

    # 3. Abstract Banner
    page.insert_text(
        (60, 110),
        "Abstract— In this paper, we propose a novel neural architecture for document conversion. We demonstrate\n"
        "state-of-the-art accuracy on multi-column layouts, mathematical equations, and complex tables.",
        fontsize=9,
        fontname="times-italic"
    )

    # 4. Left Column: Section I & Math Equations
    # Column 1 bounds: x in [54, 280]
    page.insert_text((54, 170), "I. INTRODUCTION", fontsize=11, fontname="helv")
    page.insert_text(
        (54, 190),
        "Document conversion from fixed-layout\n"
        "PDF to flowing formats has long been\n"
        "a challenging problem in computer science.\n"
        "Scientific papers contain complex\n"
        "structural elements including formulas.",
        fontsize=9.5,
        fontname="times-roman"
    )

    # Display Equation in Left Column
    page.insert_text((70, 280), "E = \\sum_{i=1}^n x_i^2 + \\int_0^\\infty f(t) dt", fontsize=10, fontname="courier")
    page.insert_text((250, 280), "(1)", fontsize=9.5, fontname="times-roman")

    # 5. Right Column: Section II & Code Block
    # Column 2 bounds: x in [330, 558]
    page.insert_text((330, 170), "II. ALGORITHM & METHOD", fontsize=11, fontname="helv")
    page.insert_text(
        (330, 190),
        "Algorithm 1: Layout-Aware Reflow\n"
        "Input: Bounding boxes B, Text spans S\n"
        "Output: Document hierarchy H\n"
        "1: def process_columns(spans):\n"
        "2:     columns = cluster_x(spans)\n"
        "3:     return sort_reading_order(columns)",
        fontsize=8.5,
        fontname="courier"
    )

    # 6. Table at bottom spanning width
    # Draw simple table grid
    rect = fitz.Rect(54, 380, 558, 460)
    page.draw_rect(rect, color=(0.2, 0.2, 0.2), width=1)
    page.draw_line((54, 405), (558, 405), color=(0.2, 0.2, 0.2), width=1)
    
    # Table Content
    page.insert_text((60, 398), "Model Name", fontsize=9, fontname="helv")
    page.insert_text((200, 398), "BLEU Score", fontsize=9, fontname="helv")
    page.insert_text((340, 398), "Layout Accuracy", fontsize=9, fontname="helv")
    page.insert_text((480, 398), "Time (s)", fontsize=9, fontname="helv")

    page.insert_text((60, 425), "Baseline Converter", fontsize=9, fontname="times-roman")
    page.insert_text((200, 425), "72.4", fontsize=9, fontname="times-roman")
    page.insert_text((340, 425), "81.2%", fontsize=9, fontname="times-roman")
    page.insert_text((480, 425), "1.84", fontsize=9, fontname="times-roman")

    page.insert_text((60, 445), "OmniPDF (Ours)", fontsize=9, fontname="times-bold")
    page.insert_text((200, 445), "98.7", fontsize=9, fontname="times-bold")
    page.insert_text((340, 445), "99.5%", fontsize=9, fontname="times-bold")
    page.insert_text((480, 445), "0.42", fontsize=9, fontname="times-bold")

    # Table Caption
    page.insert_text((180, 480), "Table 1: Benchmark evaluation on scientific paper dataset.", fontsize=9, fontname="helv")

    # 7. References
    page.insert_text((54, 520), "REFERENCES", fontsize=11, fontname="helv")
    page.insert_text((54, 540), "[1] J. Doe, 'Neural Document Parsing', IEEE Trans. PAMI, 2024.", fontsize=8.5, fontname="times-roman")
    page.insert_text((54, 560), "[2] A. Smith, 'Layout Preservation in Word Documents', ACM SIGGRAPH, 2025.", fontsize=8.5, fontname="times-roman")

    doc.save(file_path)
    doc.close()


class TestOmniPDF(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()
        self.sample_pdf = os.path.join(self.temp_dir, "sample_cs_paper.pdf")
        create_sample_cs_paper_pdf(self.sample_pdf)

    def test_pdf_inspector(self):
        """Verify pre-flight inspector correctly identifies CS paper, 2 columns, math, and code."""
        inspector = PDFInspector(self.sample_pdf)
        report = inspector.inspect()

        self.assertEqual(report.total_pages, 1)
        self.assertEqual(report.detected_columns, 2)
        self.assertEqual(report.document_type, "academic_or_cs_paper")
        self.assertEqual(report.recommended_mode, "academic")
        self.assertGreaterEqual(report.code_blocks_detected, 1)

    def test_math_engine_omml(self):
        """Test LaTeX to OMML XML conversion for fractions, superscripts, summations."""
        latex_expr = r"\frac{a+b}{2} = \sum_{i=1}^n x_i^2 + y^2"
        omml = MathEngine.convert_to_omml(latex_expr, is_display=True)

        self.assertIn("m:oMath", omml)
        self.assertIn("m:f", omml)  # Fraction
        self.assertIn("m:nary", omml)  # Summation
        self.assertIn("m:sSubSup", omml)  # Combined Subscript & Superscript (x_i^2)
        self.assertIn("m:sSup", omml)  # Superscript (y^2)

    def test_code_detector(self):
        """Test code detection on monospace text and algorithm patterns."""
        self.assertTrue(CodeDetector.is_monospace_font("CourierNewPSMT"))
        self.assertTrue(CodeDetector.is_code_line("def process_columns(spans):", "Courier"))
        self.assertTrue(CodeDetector.is_code_line("Algorithm 1: Layout-Aware Reflow"))

    def test_academic_converter(self):
        """Test end-to-end conversion in Academic mode."""
        out_docx = os.path.join(self.temp_dir, "academic_out.docx")
        converter = AcademicConverter(pdf_path=self.sample_pdf, docx_path=out_docx)
        stats = converter.convert()

        self.assertTrue(os.path.exists(out_docx))
        self.assertGreater(os.path.getsize(out_docx), 1000)
        self.assertEqual(stats["pages_processed"], 1)

        # Inspect generated DOCX
        doc = docx.Document(out_docx)
        doc_text = " ".join([p.text for p in doc.paragraphs])
        self.assertIn("Deep Learning", doc_text)
        self.assertIn("INTRODUCTION", doc_text)
        self.assertIn("REFERENCES", doc_text)

    def test_exact_layout_converter(self):
        """Test end-to-end conversion in Exact Layout mode."""
        out_docx = os.path.join(self.temp_dir, "exact_out.docx")
        converter = ExactLayoutConverter(pdf_path=self.sample_pdf, docx_path=out_docx, dpi=150)
        stats = converter.convert()

        self.assertTrue(os.path.exists(out_docx))
        self.assertGreater(os.path.getsize(out_docx), 1000)
        self.assertEqual(stats["pages_processed"], 1)
        self.assertGreaterEqual(stats["text_boxes_placed"], 5)

    def test_omni_converter_auto(self):
        """Test master OmniConverter in AUTO mode."""
        out_docx = os.path.join(self.temp_dir, "auto_out.docx")
        converter = OmniConverter()
        result = converter.convert(self.sample_pdf, out_docx, mode=ConversionMode.AUTO)

        self.assertEqual(result["status"], "success")
        self.assertTrue(os.path.exists(out_docx))
        self.assertEqual(result["mode_used"], "academic")


if __name__ == "__main__":
    unittest.main()
