"""
Layout Detector: Analyzes page geometry, column boundaries, reading order,
section hierarchies, and classifies blocks into logical academic/document structures.
"""

from dataclasses import dataclass, field
from typing import List, Dict, Any, Optional, Tuple
import re
import fitz
from omnipdf.core.math_engine import MathEngine
from omnipdf.core.code_detector import CodeDetector


class BlockType:
    TITLE = "title"
    AUTHORS = "authors"
    ABSTRACT = "abstract"
    HEADING_1 = "heading_1"
    HEADING_2 = "heading_2"
    HEADING_3 = "heading_3"
    PARAGRAPH = "paragraph"
    DISPLAY_EQUATION = "display_equation"
    CODE_BLOCK = "code_block"
    TABLE = "table"
    FIGURE = "figure"
    CAPTION = "caption"
    FOOTNOTE = "footnote"
    REFERENCES = "references"
    HEADER = "header"
    FOOTER = "footer"


@dataclass
class TextSpan:
    text: str
    font: str
    size: float
    color: int
    flags: int  # 2 = italic, 16 = bold, 8 = monospace
    bbox: Tuple[float, float, float, float]
    is_bold: bool = False
    is_italic: bool = False
    is_mono: bool = False
    is_math: bool = False


@dataclass
class TextLine:
    spans: List[TextSpan] = field(default_factory=list)
    bbox: Tuple[float, float, float, float] = (0, 0, 0, 0)
    text: str = ""


@dataclass
class LayoutBlock:
    block_type: str
    bbox: Tuple[float, float, float, float]
    lines: List[TextLine] = field(default_factory=list)
    text: str = ""
    column: int = 0  # 0 = full width, 1 = left col, 2 = right col
    reading_index: int = 0
    metadata: Dict[str, Any] = field(default_factory=dict)


class LayoutDetector:
    """Performs geometric segmentation, column detection, and reading-order reconstruction."""

    HEADING_1_REGEX = re.compile(
        r"^([I|V|X]+\.\s+[A-Z\s]+|(\d+)\s+[A-Z][A-Za-z\s]+|ABSTRACT|INTRODUCTION|RELATED WORK|METHODOLOGY|EXPERIMENTS|RESULTS|CONCLUSION|REFERENCES)\b",
        re.IGNORECASE
    )
    HEADING_2_REGEX = re.compile(
        r"^([A-Z]\.\s+[A-Za-z\s]+|(\d+\.\d+)\s+[A-Za-z\s]+)\b"
    )
    HEADING_3_REGEX = re.compile(
        r"^(\d+\)\s+[A-Za-z\s]+|(\d+\.\d+\.\d+)\s+[A-Za-z\s]+)\b"
    )
    CAPTION_REGEX = re.compile(
        r"^(Fig(\.|ure)?\s*\d+[:\.]|Table\s*\d+[:\.]|Algorithm\s*\d+[:\.])",
        re.IGNORECASE
    )
    REFERENCE_ITEM_REGEX = re.compile(
        r"^(\[\d+\]|\d+\.)\s+"
    )

    def __init__(self, page: fitz.Page, page_num: int = 1):
        self.page = page
        self.page_num = page_num
        self.rect = page.rect
        self.width = self.rect.width
        self.height = self.rect.height

    def analyze(self, extracted_tables: Optional[List[Dict[str, Any]]] = None, enable_ocr: bool = True) -> List[LayoutBlock]:
        """Analyze page layout, detect column regions, classify blocks, and sort in reading order."""
        raw_dict = self.page.get_text("dict")
        raw_blocks = raw_dict.get("blocks", [])

        # Check if page is scanned and needs OCR
        text_char_count = sum([len(span.get("text", "")) for b in raw_blocks for l in b.get("lines", []) for span in l.get("spans", [])])
        if text_char_count < 40 and enable_ocr:
            from omnipdf.core.ocr_engine import OCRAgent
            ocr_agent = OCRAgent()
            if ocr_agent.is_page_scanned(self.page):
                ocr_dict = ocr_agent.extract_text_dict_with_ocr(self.page)
                if ocr_dict.get("blocks"):
                    raw_blocks = ocr_dict.get("blocks", [])

        # 1. Parse raw text blocks into structured LayoutBlocks
        parsed_blocks = []
        for b in raw_blocks:
            b_type = b.get("type", 0)
            bbox = b.get("bbox", (0, 0, 0, 0))

            if b_type == 0:  # Text block
                lines = []
                full_block_text = []

                for l in b.get("lines", []):
                    line_spans = []
                    line_text = ""
                    for s in l.get("spans", []):
                        stext = s.get("text", "")
                        sfont = s.get("font", "")
                        ssize = s.get("size", 10.0)
                        scolor = s.get("color", 0)
                        sflags = s.get("flags", 0)
                        sbbox = s.get("bbox", (0, 0, 0, 0))

                        is_bold = bool(sflags & 16) or ("bold" in sfont.lower()) or ("black" in sfont.lower())
                        is_italic = bool(sflags & 2) or ("italic" in sfont.lower()) or ("oblique" in sfont.lower())
                        is_mono = CodeDetector.is_monospace_font(sfont)
                        is_math = MathEngine.is_math_span(stext, sfont)

                        span_obj = TextSpan(
                            text=stext,
                            font=sfont,
                            size=ssize,
                            color=scolor,
                            flags=sflags,
                            bbox=sbbox,
                            is_bold=is_bold,
                            is_italic=is_italic,
                            is_mono=is_mono,
                            is_math=is_math,
                        )
                        line_spans.append(span_obj)
                        line_text += stext

                    if line_spans:
                        lines.append(TextLine(spans=line_spans, bbox=l.get("bbox", (0, 0, 0, 0)), text=line_text))
                        full_block_text.append(line_text)

                if lines:
                    block_obj = LayoutBlock(
                        block_type=BlockType.PARAGRAPH,
                        bbox=bbox,
                        lines=lines,
                        text="\n".join(full_block_text).strip(),
                    )
                    parsed_blocks.append(block_obj)

            elif b_type == 1:  # Image block
                img_bbox = b.get("bbox", (0, 0, 0, 0))
                parsed_blocks.append(
                    LayoutBlock(
                        block_type=BlockType.FIGURE,
                        bbox=img_bbox,
                        text="[FIGURE]",
                        metadata={"image_bytes": b.get("image", None), "ext": b.get("ext", "png")},
                    )
                )

        # 2. Add extracted tables as LayoutBlocks
        if extracted_tables:
            for tab in extracted_tables:
                tab_bbox = tab["bbox"]
                parsed_blocks.append(
                    LayoutBlock(
                        block_type=BlockType.TABLE,
                        bbox=tab_bbox,
                        text="[TABLE]",
                        metadata={"table_data": tab},
                    )
                )

        # 3. Filter out text that overlaps with tables
        filtered_blocks = []
        for b in parsed_blocks:
            if b.block_type == BlockType.TABLE or b.block_type == BlockType.FIGURE:
                filtered_blocks.append(b)
                continue

            # Check if block falls inside any table bbox
            inside_table = False
            if extracted_tables:
                for tab in extracted_tables:
                    t_bbox = tab["bbox"]
                    if (b.bbox[0] >= t_bbox[0] - 5 and b.bbox[2] <= t_bbox[2] + 5 and
                        b.bbox[1] >= t_bbox[1] - 5 and b.bbox[3] <= t_bbox[3] + 5):
                        inside_table = True
                        break
            if not inside_table:
                filtered_blocks.append(b)

        # 4. Classify Blocks (Title, Abstract, Headings, Math, Code, Captions, Header/Footer)
        classified_blocks = self._classify_blocks(filtered_blocks)

        # 5. Detect column regions and sort in correct reading order
        sorted_blocks = self._order_blocks_by_reading_flow(classified_blocks)

        return sorted_blocks

    def _classify_blocks(self, blocks: List[LayoutBlock]) -> List[LayoutBlock]:
        """Classify each block into its semantic role."""
        for b in blocks:
            if b.block_type in (BlockType.TABLE, BlockType.FIGURE):
                continue

            b_text = b.text.strip()
            if not b_text:
                continue

            first_line = b.lines[0].text.strip() if b.lines else ""
            max_size = max([s.size for l in b.lines for s in l.spans], default=10.0)
            has_mono = any([s.is_mono for l in b.lines for s in l.spans])

            # Header / Footer (Top / Bottom 6% of page)
            if b.bbox[3] < self.height * 0.06 and len(b_text) < 120:
                b.block_type = BlockType.HEADER
                continue
            if b.bbox[1] > self.height * 0.94 and len(b_text) < 120:
                b.block_type = BlockType.FOOTER
                continue

            # Title (Only on Page 1, large font > 13.5 pt near top)
            if self.page_num == 1 and max_size >= 13.5 and b.bbox[1] < self.height * 0.35:
                b.block_type = BlockType.TITLE
                continue

            # Abstract / Keywords
            if re.match(r"^(Abstract|ABSTRACT|Keywords|Index Terms)[:\s—\-]", first_line):
                b.block_type = BlockType.ABSTRACT
                continue

            # Captions (Fig. 1: ..., Table 2: ...)
            if self.CAPTION_REGEX.search(first_line):
                b.block_type = BlockType.CAPTION
                continue

            # Headings
            if max_size >= 11.0 or any(s.is_bold for l in b.lines for s in l.spans):
                if self.HEADING_1_REGEX.search(first_line) and len(first_line) < 80:
                    b.block_type = BlockType.HEADING_1
                    continue
                if self.HEADING_2_REGEX.search(first_line) and len(first_line) < 80:
                    b.block_type = BlockType.HEADING_2
                    continue
                if self.HEADING_3_REGEX.search(first_line) and len(first_line) < 80:
                    b.block_type = BlockType.HEADING_3
                    continue

            # Code / Algorithm block check
            line_dicts = [{"text": l.text, "font": l.spans[0].font if l.spans else ""} for l in b.lines]
            is_code, is_algo = CodeDetector.is_code_block(line_dicts)
            if is_code or is_algo:
                b.block_type = BlockType.CODE_BLOCK
                b.metadata["is_algorithm"] = is_algo
                continue

            # Display Equation check
            raw_lines = [l.text for l in b.lines]
            if MathEngine.is_display_equation_block(raw_lines):
                b.block_type = BlockType.DISPLAY_EQUATION
                continue

            # References check
            if self.REFERENCE_ITEM_REGEX.search(first_line):
                b.block_type = BlockType.REFERENCES
                continue

        return blocks

    def _order_blocks_by_reading_flow(self, blocks: List[LayoutBlock]) -> List[LayoutBlock]:
        """
        Sort blocks according to natural reading order:
        1. Running headers
        2. Top full-width blocks (Title, Authors, Abstract, full-span banners)
        3. Column 1 (Left column, top-to-bottom)
        4. Column 2 (Right column, top-to-bottom)
        5. Bottom full-width blocks (Footnotes, full-width Figures, Footers)
        """
        # Determine page column layout
        mid_x = self.width * 0.5
        col1_thresh = self.width * 0.48
        col2_thresh = self.width * 0.52

        headers = []
        footers = []
        top_full_width = []
        bottom_full_width = []
        col1_blocks = []
        col2_blocks = []

        # Find abstract or title bottom boundary on page 1
        top_split_y = 0.0
        for b in blocks:
            if b.block_type in (BlockType.TITLE, BlockType.AUTHORS, BlockType.ABSTRACT):
                top_split_y = max(top_split_y, b.bbox[3])

        # If abstract exists, everything above top_split_y + 10 is top_full_width
        for b in blocks:
            if b.block_type == BlockType.HEADER:
                headers.append(b)
                continue
            if b.block_type == BlockType.FOOTER:
                footers.append(b)
                continue

            # Check if block is in top full-width area
            if top_split_y > 0 and b.bbox[3] <= top_split_y + 15:
                top_full_width.append(b)
                b.column = 0
                continue

            # Check if block spans across both columns (e.g. width > 60% of page)
            block_w = b.bbox[2] - b.bbox[0]
            if block_w > self.width * 0.65:
                # Full width element (e.g. wide figure or wide table)
                if b.bbox[1] < self.height * 0.5:
                    top_full_width.append(b)
                else:
                    bottom_full_width.append(b)
                b.column = 0
                continue

            # Column 1 vs Column 2 assignment
            block_center_x = (b.bbox[0] + b.bbox[2]) / 2.0
            if block_center_x < mid_x:
                col1_blocks.append(b)
                b.column = 1
            else:
                col2_blocks.append(b)
                b.column = 2

        # Sort each section top-to-bottom
        headers.sort(key=lambda b: (b.bbox[1], b.bbox[0]))
        top_full_width.sort(key=lambda b: (b.bbox[1], b.bbox[0]))
        col1_blocks.sort(key=lambda b: (b.bbox[1], b.bbox[0]))
        col2_blocks.sort(key=lambda b: (b.bbox[1], b.bbox[0]))
        bottom_full_width.sort(key=lambda b: (b.bbox[1], b.bbox[0]))
        footers.sort(key=lambda b: (b.bbox[1], b.bbox[0]))

        # Final ordered sequence
        final_sequence = (
            headers +
            top_full_width +
            col1_blocks +
            col2_blocks +
            bottom_full_width +
            footers
        )

        for idx, b in enumerate(final_sequence):
            b.reading_index = idx

        return final_sequence
