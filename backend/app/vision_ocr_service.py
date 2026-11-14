"""PDF Auto-Detect & Google Cloud Vision OCR Service.

Workflow:
1. Auto-Detect PDF Type:
   - Text PDF (embedded text layer) -> Extract immediately with pdfplumber / pypdf (0 Vision quota used).
   - Scanned PDF / Image-only PDF   -> Render pages to PNG via pdf2image and invoke
     Google Cloud Vision API (DOCUMENT_TEXT_DETECTION) with language hints ["vi", "ko", "en"],
     rotating across GCP projects via ProjectPoolRotator (1,000 free requests/month/project).
2. Structure Preservation:
   - Normalizes extracted content into ordered paragraphs and 2D tables (ExtractedBlock).
   - Provides to_docx_bytes() to feed seamlessly into TranslationService, BilingualLayout,
     and AI Legal Review.
"""

from __future__ import annotations

import io
import logging
import os
import re
from typing import Any, Literal, Optional

import docx
from docx.shared import Pt
from google.api_core import exceptions as gcp_exceptions
from pydantic import BaseModel, Field
from pypdf import PdfReader

from app.key_rotator import get_project_rotator

logger = logging.getLogger(__name__)

MIN_TEXT_CHARS_PER_PAGE = int(os.getenv("PDF_TEXT_DETECT_MIN_CHARS_PER_PAGE", "40"))
VISION_OCR_DPI = int(os.getenv("VISION_OCR_DPI", "200"))
VISION_LANGUAGE_HINTS = ["vi", "ko", "en"]

_HEADING_RE = re.compile(
    r"^(ĐIỀU|Điều|CHƯƠNG|Chương|MỤC|Mục|PHẦN|Phần|ARTICLE|Article|SECTION|Section|제\s*\d+\s*조)\b"
)


class ExtractedTable(BaseModel):
    """Normalized 2D table extracted from a PDF page."""

    rows: list[list[str]] = Field(default_factory=list)


class ExtractedBlock(BaseModel):
    """Normalized content block (paragraph or table) in reading order."""

    kind: Literal["paragraph", "table"]
    page_number: int
    text: Optional[str] = None
    table: Optional[ExtractedTable] = None
    is_heading: bool = False


class PdfExtractionResult(BaseModel):
    """Structured extraction output passed to translation_service.py."""

    filename: str
    pdf_type: Literal["text_pdf", "scanned_pdf"]
    page_count: int
    char_count: int
    blocks: list[ExtractedBlock] = Field(default_factory=list)
    vision_requests_used: int = 0
    projects_used: list[str] = Field(default_factory=list)

    def to_docx_bytes(self) -> bytes:
        """Converts normalized paragraphs & tables into a clean .docx document."""
        doc = docx.Document()
        style = doc.styles["Normal"]
        style.font.name = "Calibri"
        style.font.size = Pt(11)

        for block in self.blocks:
            if block.kind == "paragraph" and block.text:
                text = block.text.strip()
                if not text:
                    continue
                if block.is_heading:
                    p = doc.add_paragraph()
                    run = p.add_run(text)
                    run.bold = True
                    run.font.size = Pt(11.5)
                else:
                    doc.add_paragraph(text)
            elif block.kind == "table" and block.table and block.table.rows:
                rows = block.table.rows
                num_cols = max((len(r) for r in rows), default=0)
                if num_cols == 0:
                    continue
                tbl = doc.add_table(rows=len(rows), cols=num_cols)
                tbl.style = "Table Grid"
                for r_idx, row_vals in enumerate(rows):
                    for c_idx in range(num_cols):
                        cell_text = row_vals[c_idx] if c_idx < len(row_vals) else ""
                        tbl.cell(r_idx, c_idx).text = cell_text.strip()

        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    def to_plain_text(self) -> str:
        """Converts normalized paragraphs & tables into plain text."""
        parts: list[str] = []
        for block in self.blocks:
            if block.kind == "paragraph" and block.text:
                text = block.text.strip()
                if text:
                    parts.append(text)
            elif block.kind == "table" and block.table and block.table.rows:
                for row in block.table.rows:
                    line = " | ".join(c.strip() for c in row if c and c.strip())
                    if line:
                        parts.append(line)
        return "\n\n".join(parts)


class VisionOCRService:
    """Auto-detects Text vs Scanned PDF and extracts normalized paragraphs & tables."""

    def detect_and_count_text(self, file_bytes: bytes) -> tuple[bool, int, int]:
        """Inspects PDF text layer using pypdf.

        Returns (is_text_pdf, page_count, extracted_non_whitespace_chars).
        """
        reader = PdfReader(io.BytesIO(file_bytes))
        page_count = len(reader.pages)
        if page_count == 0:
            raise ValueError("File PDF không có trang nào.")

        total_non_ws = 0
        pages_with_text = 0
        for page in reader.pages:
            raw_text = page.extract_text() or ""
            non_ws = len("".join(raw_text.split()))
            total_non_ws += non_ws
            if non_ws >= MIN_TEXT_CHARS_PER_PAGE:
                pages_with_text += 1

        avg_chars = total_non_ws / max(1, page_count)
        is_text_pdf = (pages_with_text >= max(1, page_count // 2)) and (
            avg_chars >= MIN_TEXT_CHARS_PER_PAGE
        )
        return is_text_pdf, page_count, total_non_ws

    def extract_pdf(self, file_bytes: bytes, filename: str = "contract.pdf") -> PdfExtractionResult:
        """Extracts normalized blocks from either a Text PDF or a Scanned PDF."""
        is_text_pdf, page_count, _ = self.detect_and_count_text(file_bytes)
        if is_text_pdf:
            logger.info("PDF '%s' detected as Text PDF (%d pages); extracting locally (0 Vision quota).", filename, page_count)
            return self._extract_text_pdf(file_bytes, filename, page_count)

        logger.info("PDF '%s' detected as Scanned PDF (%d pages); invoking Google Cloud Vision DOCUMENT_TEXT_DETECTION.", filename, page_count)
        return self._extract_scanned_pdf_vision(file_bytes, filename, page_count)

    def extract_image(self, file_bytes: bytes, filename: str = "image.png") -> PdfExtractionResult:
        """Extracts normalized blocks from an image (.png/.jpg/.jpeg) via Google Cloud Vision."""
        from google.cloud import vision
        from PIL import Image, ImageOps

        with Image.open(io.BytesIO(file_bytes)) as img:
            img = ImageOps.exif_transpose(img)
            if img.mode not in ("RGB", "L"):
                img = img.convert("RGB")
            img_buf = io.BytesIO()
            img.save(img_buf, format="PNG")
            png_bytes = img_buf.getvalue()

        rotator = get_project_rotator()
        pool_size = max(1, len(rotator.get_pool_summary()))
        vision_image = vision.Image(content=png_bytes)
        image_context = vision.ImageContext(language_hints=VISION_LANGUAGE_HINTS)

        response = None
        last_exc: Optional[Exception] = None
        projects_used: list[str] = []
        vision_requests_used = 0

        for attempt in range(pool_size):
            client, project = rotator.get_vision_client(estimated_requests=1)
            try:
                response = client.document_text_detection(
                    image=vision_image,
                    image_context=image_context,
                    timeout=60.0,
                )
                if response.error and response.error.message:
                    raise RuntimeError(f"Vision API error on image '{filename}': {response.error.message}")
                rotator.record_vision_usage(project.project_id, 1)
                vision_requests_used += 1
                projects_used.append(project.project_id)
                break
            except gcp_exceptions.ResourceExhausted as exc:
                logger.warning(
                    "Vision API quota exhausted on project %s (attempt %d/%d); rotating.",
                    project.project_id,
                    attempt + 1,
                    pool_size,
                )
                rotator.mark_vision_exhausted(project.project_id)
                last_exc = exc
                response = None
                continue

        if response is None:
            raise RuntimeError(
                f"All GCP projects exhausted Vision API quota for image '{filename}': {last_exc}"
            )

        blocks = self._parse_vision_page_response(response, 1, vision)
        char_count = self._count_block_chars(blocks)
        return PdfExtractionResult(
            filename=filename,
            pdf_type="scanned_pdf",
            page_count=1,
            char_count=char_count,
            blocks=blocks,
            vision_requests_used=vision_requests_used,
            projects_used=projects_used,
        )

    # ------------------------------------------------------------------
    # Text PDF Extraction (0 Vision API quota)
    # ------------------------------------------------------------------

    def _extract_text_pdf(
        self,
        file_bytes: bytes,
        filename: str,
        page_count: int,
    ) -> PdfExtractionResult:
        blocks: list[ExtractedBlock] = []
        try:
            import pdfplumber

            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                for page_idx, page in enumerate(pdf.pages, start=1):
                    tables = page.find_tables() or []
                    table_bboxes = [t.bbox for t in tables]

                    # Extract non-table text outside table bounding boxes
                    filtered_page = page
                    for bbox in table_bboxes:
                        filtered_page = filtered_page.outside_bbox(bbox)

                    page_text = filtered_page.extract_text(layout=False) or ""
                    for para in self._split_into_paragraphs(page_text):
                        blocks.append(
                            ExtractedBlock(
                                kind="paragraph",
                                page_number=page_idx,
                                text=para,
                                is_heading=bool(_HEADING_RE.match(para)),
                            )
                        )

                    for tbl in tables:
                        raw_rows = tbl.extract() or []
                        cleaned_rows = [
                            [(cell or "").replace("\n", " ").strip() for cell in row]
                            for row in raw_rows
                            if any((cell or "").strip() for cell in row)
                        ]
                        if cleaned_rows:
                            blocks.append(
                                ExtractedBlock(
                                    kind="table",
                                    page_number=page_idx,
                                    table=ExtractedTable(rows=cleaned_rows),
                                )
                            )
        except ImportError:
            logger.warning("pdfplumber not installed; falling back to pypdf text extraction.")
            reader = PdfReader(io.BytesIO(file_bytes))
            for page_idx, page in enumerate(reader.pages, start=1):
                page_text = page.extract_text() or ""
                for para in self._split_into_paragraphs(page_text):
                    blocks.append(
                        ExtractedBlock(
                            kind="paragraph",
                            page_number=page_idx,
                            text=para,
                            is_heading=bool(_HEADING_RE.match(para)),
                        )
                    )

        char_count = self._count_block_chars(blocks)
        return PdfExtractionResult(
            filename=filename,
            pdf_type="text_pdf",
            page_count=page_count,
            char_count=char_count,
            blocks=blocks,
            vision_requests_used=0,
            projects_used=[],
        )

    # ------------------------------------------------------------------
    # Scanned PDF OCR via Google Cloud Vision (DOCUMENT_TEXT_DETECTION)
    # ------------------------------------------------------------------

    def _extract_scanned_pdf_vision(
        self,
        file_bytes: bytes,
        filename: str,
        page_count: int,
    ) -> PdfExtractionResult:
        from google.cloud import vision
        from pdf2image import convert_from_bytes

        images = convert_from_bytes(file_bytes, dpi=VISION_OCR_DPI, fmt="png")
        rotator = get_project_rotator()
        pool_size = max(1, len(rotator.get_pool_summary()))

        blocks: list[ExtractedBlock] = []
        projects_used: list[str] = []
        vision_requests_used = 0

        for page_idx, pil_img in enumerate(images, start=1):
            img_buf = io.BytesIO()
            pil_img.save(img_buf, format="PNG")
            img_bytes = img_buf.getvalue()

            vision_image = vision.Image(content=img_bytes)
            image_context = vision.ImageContext(language_hints=VISION_LANGUAGE_HINTS)

            response = None
            last_exc: Optional[Exception] = None
            for attempt in range(pool_size):
                client, project = rotator.get_vision_client(estimated_requests=1)
                try:
                    response = client.document_text_detection(
                        image=vision_image,
                        image_context=image_context,
                        timeout=60.0,
                    )
                    if response.error and response.error.message:
                        raise RuntimeError(
                            f"Vision API error on page {page_idx}: {response.error.message}"
                        )
                    rotator.record_vision_usage(project.project_id, 1)
                    vision_requests_used += 1
                    if project.project_id not in projects_used:
                        projects_used.append(project.project_id)
                    break
                except gcp_exceptions.ResourceExhausted as exc:
                    logger.warning(
                        "Vision API quota exhausted on project %s (attempt %d/%d); rotating.",
                        project.project_id,
                        attempt + 1,
                        pool_size,
                    )
                    rotator.mark_vision_exhausted(project.project_id)
                    last_exc = exc
                    response = None
                    continue

            if response is None:
                raise RuntimeError(
                    f"All GCP projects exhausted Vision API quota at page {page_idx}: {last_exc}"
                )

            page_blocks = self._parse_vision_page_response(response, page_idx, vision)
            blocks.extend(page_blocks)

        char_count = self._count_block_chars(blocks)
        return PdfExtractionResult(
            filename=filename,
            pdf_type="scanned_pdf",
            page_count=len(images) or page_count,
            char_count=char_count,
            blocks=blocks,
            vision_requests_used=vision_requests_used,
            projects_used=projects_used,
        )

    def _parse_vision_page_response(
        self,
        response: Any,
        page_number: int,
        vision_mod: Any,
    ) -> list[ExtractedBlock]:
        """Reconstructs paragraphs and table grids from Vision full_text_annotation."""
        annotation = getattr(response, "full_text_annotation", None)
        if not annotation or not annotation.pages:
            raw_text = getattr(response, "text_annotations", None)
            if raw_text:
                fallback = raw_text[0].description or ""
                return [
                    ExtractedBlock(
                        kind="paragraph",
                        page_number=page_number,
                        text=p,
                        is_heading=bool(_HEADING_RE.match(p)),
                    )
                    for p in self._split_into_paragraphs(fallback)
                ]
            return []

        blocks_out: list[ExtractedBlock] = []
        table_block_type = vision_mod.Block.BlockType.TABLE

        for page in annotation.pages:
            for block in page.blocks:
                if block.block_type == table_block_type:
                    rows = self._cluster_block_paragraphs_into_table(block)
                    if rows:
                        blocks_out.append(
                            ExtractedBlock(
                                kind="table",
                                page_number=page_number,
                                table=ExtractedTable(rows=rows),
                            )
                        )
                    continue

                # For TEXT blocks, check if paragraphs form a multi-column table by geometry
                para_items = []
                for para in block.paragraphs:
                    text = self._extract_paragraph_text(para)
                    if not text:
                        continue
                    x_min, y_center, height = self._bbox_metrics(para.bounding_box)
                    para_items.append((x_min, y_center, height, text))

                if not para_items:
                    continue

                if self._looks_like_multicolumn_table(para_items):
                    rows = self._cluster_items_into_rows(para_items)
                    if rows:
                        blocks_out.append(
                            ExtractedBlock(
                                kind="table",
                                page_number=page_number,
                                table=ExtractedTable(rows=rows),
                            )
                        )
                else:
                    for _, _, _, text in para_items:
                        for sub in self._split_into_paragraphs(text):
                            blocks_out.append(
                                ExtractedBlock(
                                    kind="paragraph",
                                    page_number=page_number,
                                    text=sub,
                                    is_heading=bool(_HEADING_RE.match(sub)),
                                )
                            )

        return blocks_out

    @staticmethod
    def _extract_paragraph_text(para: Any) -> str:
        """Reconstructs paragraph text from Vision words & symbols (preserving VI diacritics & Hangul)."""
        parts: list[str] = []
        for word in para.words:
            for sym in word.symbols:
                parts.append(sym.text)
                brk = sym.property.detected_break.type_ if sym.property and sym.property.detected_break else 0
                # 1 = SPACE, 2 = SURE_SPACE, 3 = EOL_SURE_SPACE, 4 = HYPHEN, 5 = LINE_BREAK
                if brk in (1, 2):
                    parts.append(" ")
                elif brk in (3, 5):
                    parts.append("\n")
        return "".join(parts).strip()

    @staticmethod
    def _bbox_metrics(bbox: Any) -> tuple[float, float, float]:
        vertices = getattr(bbox, "vertices", None) or []
        if not vertices:
            return 0.0, 0.0, 12.0
        xs = [float(getattr(v, "x", 0) or 0) for v in vertices]
        ys = [float(getattr(v, "y", 0) or 0) for v in vertices]
        x_min = min(xs)
        y_min = min(ys)
        y_max = max(ys)
        return x_min, (y_min + y_max) / 2.0, max(8.0, y_max - y_min)

    def _cluster_block_paragraphs_into_table(self, block: Any) -> list[list[str]]:
        items = []
        for para in block.paragraphs:
            text = self._extract_paragraph_text(para)
            if not text:
                continue
            x_min, y_center, height = self._bbox_metrics(para.bounding_box)
            items.append((x_min, y_center, height, text.replace("\n", " ")))
        return self._cluster_items_into_rows(items)

    @staticmethod
    def _cluster_items_into_rows(
        items: list[tuple[float, float, float, str]],
    ) -> list[list[str]]:
        if not items:
            return []
        sorted_items = sorted(items, key=lambda it: (it[1], it[0]))
        rows_raw: list[list[tuple[float, float, float, str]]] = []

        for item in sorted_items:
            if not rows_raw:
                rows_raw.append([item])
                continue
            last_row = rows_raw[-1]
            avg_y = sum(it[1] for it in last_row) / len(last_row)
            avg_h = sum(it[2] for it in last_row) / len(last_row)
            if abs(item[1] - avg_y) <= max(10.0, avg_h * 0.65):
                last_row.append(item)
            else:
                rows_raw.append([item])

        table_rows: list[list[str]] = []
        for row in rows_raw:
            row_sorted = sorted(row, key=lambda it: it[0])
            table_rows.append([it[3].replace("\n", " ").strip() for it in row_sorted])
        return table_rows

    @staticmethod
    def _looks_like_multicolumn_table(
        items: list[tuple[float, float, float, str]],
    ) -> bool:
        """Detects if a TEXT block actually represents aligned multi-column table rows."""
        if len(items) < 4:
            return False
        sorted_items = sorted(items, key=lambda it: it[1])
        same_y_pairs = 0
        for i in range(len(sorted_items) - 1):
            _, y1, h1, _ = sorted_items[i]
            _, y2, _, _ = sorted_items[i + 1]
            if abs(y2 - y1) <= max(10.0, h1 * 0.6):
                same_y_pairs += 1
        return same_y_pairs >= 2

    @staticmethod
    def _split_into_paragraphs(text: str) -> list[str]:
        if not text:
            return []
        raw_paras = re.split(r"\n\s*\n+", text.strip())
        cleaned: list[str] = []
        for p in raw_paras:
            normalized = " ".join(line.strip() for line in p.splitlines() if line.strip())
            if normalized:
                cleaned.append(normalized)
        return cleaned

    @staticmethod
    def _count_block_chars(blocks: list[ExtractedBlock]) -> int:
        total = 0
        for b in blocks:
            if b.kind == "paragraph" and b.text:
                total += len(b.text)
            elif b.kind == "table" and b.table:
                for row in b.table.rows:
                    for cell in row:
                        total += len(cell)
        return total


_vision_ocr_singleton: Optional[VisionOCRService] = None


def get_vision_ocr_service() -> VisionOCRService:
    global _vision_ocr_singleton
    if _vision_ocr_singleton is None:
        _vision_ocr_singleton = VisionOCRService()
    return _vision_ocr_singleton
