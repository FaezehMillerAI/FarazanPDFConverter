"""
Flow Document Converter: Wraps and enhances pdf2docx for standard linear documents,
books, reports, and general text documents that require standard Word flowing paragraphs.
"""

import os
from typing import Optional, Callable, Dict, Any, Tuple
from pdf2docx import Converter


class FlowConverter:
    """Standard flowing text converter with automated table, paragraph, and heading detection."""

    def __init__(
        self,
        pdf_path: str,
        docx_path: str,
        on_progress: Optional[Callable[[int, int, str], None]] = None,
    ):
        self.pdf_path = pdf_path
        self.docx_path = docx_path
        self.on_progress = on_progress

    def convert(self, page_range: Optional[Tuple[int, int]] = None) -> Dict[str, Any]:
        """Convert PDF to flow Word document."""
        if not os.path.exists(self.pdf_path):
            raise FileNotFoundError(f"PDF file not found: {self.pdf_path}")

        os.makedirs(os.path.dirname(os.path.abspath(self.docx_path)), exist_ok=True)

        start_page = page_range[0] if page_range else 0
        end_page = page_range[1] if page_range else None

        if self.on_progress:
            self.on_progress(1, 1, "Parsing document flow and layout...")

        cv = Converter(self.pdf_path)
        try:
            cv.convert(self.docx_path, start=start_page, end=end_page)
        finally:
            cv.close()

        if self.on_progress:
            self.on_progress(1, 1, "Conversion complete.")

        return {
            "mode": "flow",
            "status": "success",
            "output_path": self.docx_path,
        }
