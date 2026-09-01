"""
OmniPDF: Advanced Layout-Preserving PDF to Word (DOCX) Converter
Specialized for Computer Science, Academic Papers, and Complex Layouts.
"""

from omnipdf.core.converter import OmniConverter, ConversionMode, convert_pdf
from omnipdf.core.inspector import PDFInspector, inspect_pdf

__version__ = "1.0.0"
__all__ = [
    "OmniConverter",
    "ConversionMode",
    "convert_pdf",
    "PDFInspector",
    "inspect_pdf",
    "__version__",
]
