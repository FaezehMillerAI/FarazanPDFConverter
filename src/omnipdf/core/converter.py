"""
OmniConverter: Master orchestrator for converting any PDF to editable Word (DOCX).
Supports Academic/CS, Exact Layout, Flow, and Auto-Adaptive modes.
"""

from enum import Enum
import time
import os
from typing import Optional, Callable, Dict, Any, Tuple

from omnipdf.core.inspector import PDFInspector
from omnipdf.core.academic_converter import AcademicConverter
from omnipdf.core.exact_converter import ExactLayoutConverter
from omnipdf.core.flow_converter import FlowConverter


class ConversionMode(str, Enum):
    AUTO = "auto"
    ACADEMIC = "academic"
    EXACT = "exact"
    FLOW = "flow"


class OmniConverter:
    """Unified PDF-to-Word conversion engine with auto-adaptive intelligence."""

    def __init__(
        self,
        convert_math: bool = True,
        highlight_code: bool = True,
        extract_images: bool = True,
        image_dpi: int = 300,
    ):
        self.convert_math = convert_math
        self.highlight_code = highlight_code
        self.extract_images = extract_images
        self.image_dpi = image_dpi

    def convert(
        self,
        pdf_path: str,
        docx_path: Optional[str] = None,
        mode: ConversionMode = ConversionMode.AUTO,
        page_range: Optional[Tuple[int, int]] = None,
        on_progress: Optional[Callable[[int, int, str], None]] = None,
    ) -> Dict[str, Any]:
        """Convert any PDF to Word document using the specified or auto-detected mode."""
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"Source PDF does not exist: {pdf_path}")

        if not docx_path:
            base_name, _ = os.path.splitext(pdf_path)
            docx_path = f"{base_name}.docx"

        start_time = time.time()

        # Pre-flight Inspection
        inspector = PDFInspector(pdf_path)
        report = inspector.inspect()

        # Resolve mode if AUTO
        selected_mode = mode
        if mode == ConversionMode.AUTO:
            if report.recommended_mode == "academic":
                selected_mode = ConversionMode.ACADEMIC
            elif report.recommended_mode == "exact":
                selected_mode = ConversionMode.EXACT
            else:
                selected_mode = ConversionMode.FLOW

        # Dispatch to appropriate engine
        stats = {}
        try:
            if selected_mode == ConversionMode.ACADEMIC:
                converter = AcademicConverter(
                    pdf_path=pdf_path,
                    docx_path=docx_path,
                    convert_math=self.convert_math,
                    highlight_code=self.highlight_code,
                    extract_images=self.extract_images,
                    image_dpi=self.image_dpi,
                    on_progress=on_progress,
                )
                stats = converter.convert(page_range=page_range)

            elif selected_mode == ConversionMode.EXACT:
                converter = ExactLayoutConverter(
                    pdf_path=pdf_path,
                    docx_path=docx_path,
                    dpi=self.image_dpi,
                    on_progress=on_progress,
                )
                stats = converter.convert(page_range=page_range)

            elif selected_mode == ConversionMode.FLOW:
                converter = FlowConverter(
                    pdf_path=pdf_path,
                    docx_path=docx_path,
                    on_progress=on_progress,
                )
                stats = converter.convert(page_range=page_range)

        except Exception as primary_error:
            # Automatic Fallback to Flow or Exact if primary engine fails
            fallback_mode = ConversionMode.EXACT if selected_mode != ConversionMode.EXACT else ConversionMode.FLOW
            if on_progress:
                on_progress(1, 1, f"Primary mode encountered an issue, falling back to {fallback_mode.value}...")

            if fallback_mode == ConversionMode.EXACT:
                fb_converter = ExactLayoutConverter(pdf_path=pdf_path, docx_path=docx_path, dpi=self.image_dpi)
                stats = fb_converter.convert(page_range=page_range)
            else:
                fb_converter = FlowConverter(pdf_path=pdf_path, docx_path=docx_path)
                stats = fb_converter.convert(page_range=page_range)

            stats["fallback_used"] = True
            stats["primary_error"] = str(primary_error)

        duration = round(time.time() - start_time, 2)

        output_size = os.path.getsize(docx_path) if os.path.exists(docx_path) else 0

        return {
            "status": "success",
            "input_file": pdf_path,
            "output_file": docx_path,
            "mode_used": selected_mode.value if isinstance(selected_mode, ConversionMode) else selected_mode,
            "duration_seconds": duration,
            "output_size_kb": round(output_size / 1024, 2),
            "document_type": report.document_type,
            "total_pages": report.total_pages,
            "stats": stats,
        }


def convert_pdf(
    pdf_path: str,
    docx_path: Optional[str] = None,
    mode: str = "auto",
    page_range: Optional[Tuple[int, int]] = None,
    convert_math: bool = True,
    highlight_code: bool = True,
    extract_images: bool = True,
    image_dpi: int = 300,
) -> Dict[str, Any]:
    """Top-level convenience function for converting a PDF to Word."""
    mode_enum = ConversionMode(mode.lower())
    converter = OmniConverter(
        convert_math=convert_math,
        highlight_code=highlight_code,
        extract_images=extract_images,
        image_dpi=image_dpi,
    )
    return converter.convert(pdf_path, docx_path, mode=mode_enum, page_range=page_range)
