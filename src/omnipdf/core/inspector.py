"""
PDF Pre-flight Inspector: Analyzes layout, fonts, math formulas, code blocks,
tables, images, and recommends the optimal conversion strategy.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional
import os
import re
import fitz  # PyMuPDF


@dataclass
class PageInspection:
    page_number: int
    width: float
    height: float
    orientation: str
    column_count: int
    text_blocks_count: int
    has_math: bool
    has_code: bool
    tables_count: int
    images_count: int
    drawings_count: int
    is_scanned: bool
    font_names: List[str] = field(default_factory=list)


@dataclass
class InspectionReport:
    file_path: str
    file_name: str
    file_size_bytes: int
    total_pages: int
    is_encrypted: bool
    is_scanned: bool
    detected_columns: int
    math_density: str
    code_blocks_detected: int
    tables_detected: int
    images_detected: int
    vector_drawings_detected: int
    document_type: str
    recommended_mode: str
    compatibility_score: int  # 0 to 100
    fonts: List[str] = field(default_factory=list)
    pages: List[PageInspection] = field(default_factory=list)
    notes: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "file_name": self.file_name,
            "file_size_kb": round(self.file_size_bytes / 1024, 2),
            "total_pages": self.total_pages,
            "is_encrypted": self.is_encrypted,
            "is_scanned": self.is_scanned,
            "detected_columns": self.detected_columns,
            "math_density": self.math_density,
            "code_blocks_detected": self.code_blocks_detected,
            "tables_detected": self.tables_detected,
            "images_detected": self.images_detected,
            "vector_drawings_detected": self.vector_drawings_detected,
            "document_type": self.document_type,
            "recommended_mode": self.recommended_mode,
            "compatibility_score": self.compatibility_score,
            "fonts": self.fonts[:15],
            "notes": self.notes,
            "pages": [
                {
                    "page": p.page_number,
                    "width": round(p.width, 1),
                    "height": round(p.height, 1),
                    "orientation": p.orientation,
                    "columns": p.column_count,
                    "tables": p.tables_count,
                    "images": p.images_count,
                    "has_math": p.has_math,
                    "has_code": p.has_code,
                    "is_scanned": p.is_scanned,
                }
                for p in self.pages
            ],
        }


class PDFInspector:
    """Deep PDF structure, layout, and content analyzer."""

    MATH_FONT_KEYWORDS = [
        "math", "cmr", "cmsy", "cmmi", "msbm", "msam", "stix", "symbol",
        "euler", "wasy", "bbm", "dsrom", "nimbusroman", "tex"
    ]
    
    MATH_SYMBOLS_REGEX = re.compile(
        r"[\u0370-\u03FF\u2100-\u214F\u2200-\u22FF\u2A00-\u2AFF\u27C0-\u27EF\u2980-\u29FF"
        r"±×÷·√∑∏∫∮∇∂∞≈≠≡≤≥⊂⊆⊃⊇∈∉∪∩∧∨¬⇒⇔→←↑↓]"
    )
    
    LATEX_KEYWORDS_REGEX = re.compile(
        r"(\\frac|\\sqrt|\\sum|\\int|\\prod|\\mathbf|\\mathcal|\\mathbb|\\text|\\alpha|\\beta|\\gamma|\\theta|\\lambda|\\sigma|\\omega|_\{|\^\{)"
    )

    CODE_FONT_KEYWORDS = [
        "courier", "consolas", "monaco", "menlo", "dejavu sans mono",
        "inconsolata", "source code pro", "fira", "monospace", "cmtt", "typewriter"
    ]

    CODE_SYNTAX_PATTERNS = [
        re.compile(r"^\s*(def|class|function|import|from|return|public|private|static|void|const|let|var)\s+\w+"),
        re.compile(r"^\s*(for|while|if|else|switch|case|try|catch|finally)\s*[\(\{]"),
        re.compile(r"^\s*(\{|\}|\[|\]|;|//|#|/\*)\s*$"),
        re.compile(r"^\s*Algorithm\s+\d+[:\.]"),
        re.compile(r"^\s*Input:\s+|^\s*Output:\s+"),
        re.compile(r"^\s*\d+:\s+[a-zA-Z_]"),  # Line numbered code
    ]

    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path

    def inspect(self) -> InspectionReport:
        if not os.path.exists(self.pdf_path):
            raise FileNotFoundError(f"PDF file not found: {self.pdf_path}")

        file_size = os.path.getsize(self.pdf_path)
        file_name = os.path.basename(self.pdf_path)

        doc = fitz.open(self.pdf_path)
        is_encrypted = doc.is_encrypted
        total_pages = len(doc)

        if total_pages == 0:
            doc.close()
            return InspectionReport(
                file_path=self.pdf_path,
                file_name=file_name,
                file_size_bytes=file_size,
                total_pages=0,
                is_encrypted=is_encrypted,
                is_scanned=False,
                detected_columns=1,
                math_density="none",
                code_blocks_detected=0,
                tables_detected=0,
                images_detected=0,
                vector_drawings_detected=0,
                document_type="empty",
                recommended_mode="flow",
                compatibility_score=0,
                notes=["The PDF document contains 0 pages."],
            )

        all_fonts = set()
        total_tables = 0
        total_images = 0
        total_drawings = 0
        total_math_occurrences = 0
        total_code_occurrences = 0
        scanned_page_count = 0
        page_columns_list = []
        page_inspections = []

        for page_idx in range(total_pages):
            page = doc[page_idx]
            rect = page.rect
            width, height = rect.width, rect.height
            orientation = "Landscape" if width > height else "Portrait"

            # 1. Text & Font Extraction
            text_dict = page.get_text("dict")
            blocks = text_dict.get("blocks", [])
            text_blocks = [b for b in blocks if b.get("type") == 0]
            
            # Extract fonts & math/code occurrences
            page_fonts = set()
            page_math = False
            page_code = False
            total_chars = 0

            for b in text_blocks:
                for line in b.get("lines", []):
                    line_text = "".join([span.get("text", "") for span in line.get("spans", [])]).strip()
                    total_chars += len(line_text)
                    
                    # Check code syntax
                    for pattern in self.CODE_SYNTAX_PATTERNS:
                        if pattern.search(line_text):
                            page_code = True
                            total_code_occurrences += 1
                            break

                    for span in line.get("spans", []):
                        font = span.get("font", "").lower()
                        text = span.get("text", "")
                        page_fonts.add(span.get("font", ""))
                        all_fonts.add(span.get("font", ""))

                        # Check math font
                        if any(mf in font for mf in self.MATH_FONT_KEYWORDS):
                            page_math = True
                            total_math_occurrences += 1
                        elif self.MATH_SYMBOLS_REGEX.search(text) or self.LATEX_KEYWORDS_REGEX.search(text):
                            page_math = True
                            total_math_occurrences += 1

                        # Check code font
                        if any(cf in font for cf in self.CODE_FONT_KEYWORDS):
                            page_code = True
                            total_code_occurrences += 1

            # 2. Table Detection
            try:
                table_finder = page.find_tables()
                tables_count = len(table_finder.tables) if table_finder else 0
            except Exception:
                tables_count = 0
            total_tables += tables_count

            # 3. Images & Drawings
            img_list = page.get_images()
            images_count = len(img_list)
            total_images += images_count

            drawings = page.get_drawings()
            drawings_count = len(drawings)
            total_drawings += drawings_count

            # 4. Scanned Page Check
            is_scanned = (total_chars < 50 and images_count > 0)
            if is_scanned:
                scanned_page_count += 1

            # 5. Multi-column Detection
            cols = self._detect_columns(text_blocks, width)
            page_columns_list.append(cols)

            page_inspections.append(
                PageInspection(
                    page_number=page_idx + 1,
                    width=width,
                    height=height,
                    orientation=orientation,
                    column_count=cols,
                    text_blocks_count=len(text_blocks),
                    has_math=page_math,
                    has_code=page_code,
                    tables_count=tables_count,
                    images_count=images_count,
                    drawings_count=drawings_count,
                    is_scanned=is_scanned,
                    font_names=list(page_fonts),
                )
            )

        doc.close()

        # Aggregate Statistics
        overall_scanned = (scanned_page_count / total_pages) >= 0.5
        avg_cols = round(sum(page_columns_list) / len(page_columns_list)) if page_columns_list else 1
        
        # Math density classification
        if total_math_occurrences > total_pages * 5:
            math_density = "heavy"
        elif total_math_occurrences > 0:
            math_density = "moderate"
        else:
            math_density = "none"

        # Document Type & Mode Recommendation
        notes = []
        if overall_scanned:
            document_type = "scanned_document"
            recommended_mode = "exact"
            notes.append("Document appears to be scanned or contains raster images without embedded text layer.")
        elif avg_cols >= 2 or math_density in ("heavy", "moderate") or total_code_occurrences > 0:
            document_type = "academic_or_cs_paper"
            recommended_mode = "academic"
            notes.append("Detected multi-column academic/computer science paper structure with formulas, code, or citations.")
        elif total_drawings > total_pages * 10 or total_images > total_pages * 3:
            document_type = "graphic_or_presentation"
            recommended_mode = "exact"
            notes.append("High density of vector drawings and graphics detected. Exact layout preservation recommended.")
        else:
            document_type = "standard_document"
            recommended_mode = "flow"
            notes.append("Standard linear layout detected. Flow mode will generate clean editable Word paragraphs and headings.")

        compatibility_score = 98 if not overall_scanned else 75

        return InspectionReport(
            file_path=self.pdf_path,
            file_name=file_name,
            file_size_bytes=file_size,
            total_pages=total_pages,
            is_encrypted=is_encrypted,
            is_scanned=overall_scanned,
            detected_columns=avg_cols,
            math_density=math_density,
            code_blocks_detected=total_code_occurrences,
            tables_detected=total_tables,
            images_detected=total_images,
            vector_drawings_detected=total_drawings,
            document_type=document_type,
            recommended_mode=recommended_mode,
            compatibility_score=compatibility_score,
            fonts=list(all_fonts),
            pages=page_inspections,
            notes=notes,
        )

    def _detect_columns(self, text_blocks: List[Dict], page_width: float) -> int:
        """Estimate number of text columns on a page using x-coordinate clustering."""
        if not text_blocks or len(text_blocks) < 4:
            return 1

        # Exclude headers (top 8%) and footers (bottom 8%)
        body_blocks = []
        for b in text_blocks:
            bbox = b.get("bbox", (0, 0, 0, 0))
            body_blocks.append(bbox)

        if not body_blocks:
            return 1

        # Check for 2-column layout: x0 distribution
        left_margin = page_width * 0.45
        right_margin = page_width * 0.55
        
        left_col_count = sum(1 for bbox in body_blocks if bbox[0] < left_margin and bbox[2] < right_margin)
        right_col_count = sum(1 for bbox in body_blocks if bbox[0] >= left_margin)

        if left_col_count >= 2 and right_col_count >= 2:
            return 2
        return 1


def inspect_pdf(pdf_path: str) -> Dict[str, Any]:
    """Convenience helper to inspect PDF and return dict."""
    inspector = PDFInspector(pdf_path)
    report = inspector.inspect()
    return report.to_dict()
