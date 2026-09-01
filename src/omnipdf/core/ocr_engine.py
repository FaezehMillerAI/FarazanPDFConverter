"""
OCR Agent & Engine: Optical Character Recognition for scanned PDFs, image-based papers,
and non-searchable document pages using PyMuPDF OCR and pytesseract.
"""

import io
import os
import re
from typing import Optional, List, Dict, Any, Tuple
from PIL import Image
import fitz


class OCRAgent:
    """Intelligent OCR Engine for recognizing text, formulas, and layouts from scanned pages."""

    def __init__(self, language: str = "eng", dpi: int = 300):
        self.language = language
        self.dpi = dpi

    def is_page_scanned(self, page: fitz.Page, min_char_thresh: int = 40) -> bool:
        """Check if a page lacks embedded digital text and requires OCR."""
        text = page.get_text("text").strip()
        if len(text) < min_char_thresh:
            images = page.get_images()
            if len(images) > 0:
                return True
            # Check page drawing size
            rect = page.rect
            if rect.width > 0 and rect.height > 0:
                return True
        return False

    def extract_text_dict_with_ocr(self, page: fitz.Page) -> Dict[str, Any]:
        """
        Run OCR on a page and return a PyMuPDF-compatible text dictionary
        with blocks, lines, spans, and bounding boxes.
        """
        try:
            # 1. Try PyMuPDF built-in OCR textpage
            tp = page.get_textpage_ocr(
                dpi=self.dpi,
                language=self.language,
                full=True
            )
            raw_dict = page.get_text("dict", textpage=tp)
            blocks = raw_dict.get("blocks", [])
            if blocks and len(blocks) > 0:
                return raw_dict
        except Exception:
            pass

        # 2. Fallback: Render page to high-res image and run pytesseract
        return self._pytesseract_fallback(page)

    def _pytesseract_fallback(self, page: fitz.Page) -> Dict[str, Any]:
        """Fallback to pytesseract if PyMuPDF OCR is unavailable."""
        try:
            import pytesseract
            
            # Render page at high DPI
            zoom = self.dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img = Image.open(io.BytesIO(pix.tobytes("png")))

            # Get OCR bounding box data
            ocr_data = pytesseract.image_to_data(img, lang=self.language, output_type=pytesseract.Output.DICT)
            
            blocks_map = {}
            n_boxes = len(ocr_data["text"])
            
            scale = 72.0 / self.dpi

            for i in range(n_boxes):
                text = ocr_data["text"][i].strip()
                if not text:
                    continue

                conf = int(ocr_data["conf"][i])
                if conf < 30:  # Skip very low confidence noise
                    continue

                b_num = ocr_data["block_num"][i]
                l_num = ocr_data["line_num"][i]
                
                x = ocr_data["left"][i] * scale
                y = ocr_data["top"][i] * scale
                w = ocr_data["width"][i] * scale
                h = ocr_data["height"][i] * scale

                if b_num not in blocks_map:
                    blocks_map[b_num] = {"lines": {}, "bbox": [x, y, x + w, y + h]}
                
                b_entry = blocks_map[b_num]
                b_entry["bbox"][0] = min(b_entry["bbox"][0], x)
                b_entry["bbox"][1] = min(b_entry["bbox"][1], y)
                b_entry["bbox"][2] = max(b_entry["bbox"][2], x + w)
                b_entry["bbox"][3] = max(b_entry["bbox"][3], y + h)

                if l_num not in b_entry["lines"]:
                    b_entry["lines"][l_num] = []

                b_entry["lines"][l_num].append({
                    "text": text,
                    "font": "Calibri",
                    "size": max(8.0, h),
                    "color": 0,
                    "flags": 0,
                    "bbox": (x, y, x + w, y + h),
                })

            # Assemble into PyMuPDF dict format
            formatted_blocks = []
            for b_num, b_data in sorted(blocks_map.items()):
                formatted_lines = []
                for l_num, spans in sorted(b_data["lines"].items()):
                    line_x0 = min([s["bbox"][0] for s in spans])
                    line_y0 = min([s["bbox"][1] for s in spans])
                    line_x1 = max([s["bbox"][2] for s in spans])
                    line_y1 = max([s["bbox"][3] for s in spans])
                    formatted_lines.append({
                        "bbox": (line_x0, line_y0, line_x1, line_y1),
                        "spans": spans,
                    })

                if formatted_lines:
                    formatted_blocks.append({
                        "type": 0,
                        "bbox": tuple(b_data["bbox"]),
                        "lines": formatted_lines,
                    })

            return {"blocks": formatted_blocks}
        except Exception:
            return {"blocks": []}


def perform_ocr_on_pdf(pdf_path: str, output_path: Optional[str] = None) -> str:
    """Generate a searchable OCR-enhanced PDF copy."""
    if not output_path:
        base, _ = os.path.splitext(pdf_path)
        output_path = f"{base}_ocr.pdf"

    doc = fitz.open(pdf_path)
    ocr_doc = fitz.open()

    agent = OCRAgent()

    for page in doc:
        if agent.is_page_scanned(page):
            try:
                pdf_bytes = page.get_textpage_ocr(dpi=300).pdf
                ocr_page = fitz.open("pdf", pdf_bytes)
                ocr_doc.insert_pdf(ocr_page)
            except Exception:
                ocr_doc.insert_pdf(doc, from_page=page.number, to_page=page.number)
        else:
            ocr_doc.insert_pdf(doc, from_page=page.number, to_page=page.number)

    ocr_doc.save(output_path)
    ocr_doc.close()
    doc.close()
    return output_path
