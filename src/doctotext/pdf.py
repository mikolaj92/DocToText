from __future__ import annotations

from dataclasses import dataclass, field
from difflib import SequenceMatcher
from enum import StrEnum
from functools import lru_cache
from pathlib import Path

import fitz

from .common import PDF_MIME, DocumentBytes, DocumentError, output_filename
from .docx import TextSegment

A4_WIDTH = 595
A4_HEIGHT = 842
PAGE_MARGIN = 48
UNICODE_FONT_CANDIDATES = (
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/usr/local/share/fonts/DejaVuSans.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
)


class PdfExtractionMode(StrEnum):
    TEXT_LAYER = "text_layer"


@dataclass
class PdfDocument:
    filename: str
    pages: list[str]
    source_bytes: bytes = field(default=b"", repr=False)
    source_pages: list[str] | None = None
    extraction_mode: PdfExtractionMode = PdfExtractionMode.TEXT_LAYER

    def __post_init__(self) -> None:
        if self.source_pages is None:
            self.source_pages = list(self.pages)

    @classmethod
    def open_bytes(cls, data: bytes, *, filename: str = "document.pdf") -> PdfDocument:
        try:
            pdf = fitz.open(stream=data, filetype="pdf")
            pages = [page.get_text("text") or "" for page in pdf]
        except Exception as error:
            raise DocumentError("Nie udało się odczytać PDF.") from error

        if not any(page.strip() for page in pages):
            raise DocumentError("PDF nie ma warstwy tekstowej. Ten plik wymaga OCR.")
        return cls(
            filename=filename,
            pages=list(pages),
            source_bytes=data,
            source_pages=list(pages),
        )

    @property
    def segments(self) -> tuple[TextSegment, ...]:
        return tuple(
            TextSegment(id=f"p{index}", text=text, part=f"page:{index}", index=index)
            for index, text in enumerate(self.pages)
            if text.strip()
        )

    @property
    def texts(self) -> list[str]:
        return [segment.text for segment in self.segments]

    def apply_texts(self, texts) -> None:
        texts = list(texts)
        segment_indexes = [segment.index for segment in self.segments]
        if len(texts) != len(segment_indexes):
            raise ValueError(
                f"expected {len(segment_indexes)} text segments, got {len(texts)}"
            )
        for index, text in zip(segment_indexes, texts, strict=True):
            self.pages[index] = text

    def to_bytes(self) -> bytes:
        source_pages = self.source_pages or []
        if self.source_bytes and self.pages == source_pages:
            return self.source_bytes
        if not self.source_bytes:
            return _render_text_pdf(self.pages)
        if len(self.pages) != len(source_pages):
            raise DocumentError("Nie udało się zapisać PDF: zmieniła się liczba stron.")

        try:
            pdf = fitz.open(stream=self.source_bytes, filetype="pdf")
            for page_index, (source_text, anonymized_text) in enumerate(
                zip(source_pages, self.pages, strict=True)
            ):
                if source_text == anonymized_text:
                    continue

                page = pdf[page_index]
                if not _redact_page_changes(page, source_text, anonymized_text):
                    _replace_page_with_text(page, anonymized_text)

            return pdf.tobytes(garbage=4, deflate=True)
        except DocumentError:
            raise
        except Exception as error:
            raise DocumentError("Nie udało się zapisać PDF.") from error

    def to_document_bytes(self) -> DocumentBytes:
        return DocumentBytes(
            filename=output_filename(self.filename, "pdf"),
            content_type=PDF_MIME,
            data=self.to_bytes(),
        )


def _redact_page_changes(page, source_text: str, anonymized_text: str) -> bool:
    added_redaction = False
    for source_span, replacement_span in _changed_text_spans(source_text, anonymized_text):
        source = source_span.strip()
        if not source:
            continue
        if _unsafe_short_redaction_source(source):
            return False

        rects = _search_text_rects(page, source)
        if not rects:
            return False

        label = _redaction_label(replacement_span)
        for rect in rects:
            expanded = _expand_rect(page, rect)
            page.add_redact_annot(
                expanded,
                text=label,
                fill=(1, 1, 1),
                text_color=(0, 0, 0),
                fontsize=_redaction_font_size(expanded),
            )
            added_redaction = True

    if added_redaction:
        page.apply_redactions()
    return True


def _changed_text_spans(source_text: str, anonymized_text: str):
    matcher = SequenceMatcher(None, source_text, anonymized_text, autojunk=False)
    for tag, source_start, source_end, replacement_start, replacement_end in matcher.get_opcodes():
        if tag in {"equal", "insert"}:
            continue
        yield (
            source_text[source_start:source_end],
            anonymized_text[replacement_start:replacement_end],
        )


def _unsafe_short_redaction_source(text: str) -> bool:
    alnum_count = sum(character.isalnum() for character in text)
    return 0 < alnum_count < 2


def _search_text_rects(page, text: str):
    for candidate in _search_candidates(text):
        rects = page.search_for(candidate)
        if rects:
            return rects

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if len(lines) < 2:
        return []

    rects = []
    for line in lines:
        if _unsafe_short_redaction_source(line):
            return []
        line_rects = page.search_for(line)
        if not line_rects:
            return []
        rects.extend(line_rects)
    return rects


def _search_candidates(text: str) -> list[str]:
    candidates = [text.strip(), " ".join(text.split())]
    unique = []
    for candidate in candidates:
        if candidate and candidate not in unique:
            unique.append(candidate)
    return unique


def _redaction_label(text: str) -> str:
    label = " ".join(text.split())
    if not label:
        return ""
    if len(label) > 64 or any(ord(character) > 127 for character in label):
        return "****"
    return label


def _expand_rect(page, rect) -> fitz.Rect:
    expanded = fitz.Rect(rect)
    expanded.x0 = max(page.rect.x0, expanded.x0 - 0.75)
    expanded.y0 = max(page.rect.y0, expanded.y0 - 0.75)
    expanded.x1 = min(page.rect.x1, expanded.x1 + 0.75)
    expanded.y1 = min(page.rect.y1, expanded.y1 + 0.75)
    return expanded


def _redaction_font_size(rect: fitz.Rect) -> float:
    return max(5, min(10, rect.height * 0.7))


def _replace_page_with_text(page, text: str) -> None:
    page.add_redact_annot(page.rect, fill=(1, 1, 1))
    page.apply_redactions()
    _insert_page_text(page, text)


def _render_text_pdf(pages: list[str]) -> bytes:
    pdf = fitz.open()
    for page_text in pages:
        page = pdf.new_page(width=A4_WIDTH, height=A4_HEIGHT)
        _insert_page_text(page, page_text)
    return pdf.tobytes(garbage=4, deflate=True)


def _insert_page_text(page, text: str) -> None:
    target = fitz.Rect(
        page.rect.x0 + PAGE_MARGIN,
        page.rect.y0 + PAGE_MARGIN,
        page.rect.x1 - PAGE_MARGIN,
        page.rect.y1 - PAGE_MARGIN,
    )
    kwargs = {
        "fontsize": 10,
        "color": (0, 0, 0),
        "align": fitz.TEXT_ALIGN_LEFT,
    }
    font_path = _unicode_font_path()
    if font_path:
        kwargs["fontfile"] = font_path
        kwargs["fontname"] = "doctotextunicode"
    page.insert_textbox(target, text or "", **kwargs)


@lru_cache(maxsize=1)
def _unicode_font_path() -> str | None:
    for path in UNICODE_FONT_CANDIDATES:
        if Path(path).exists():
            return path
    return None
