from .common import (
    DOCX_MIME,
    MD_MIME,
    PDF_MIME,
    TXT_MIME,
    DocumentBytes,
    DocumentError,
    DocumentKind,
)
from .detection import DetectedDocumentType, detect_document_type
from .loader import document_to_bytes, load_document
from .pdf import PdfDocument, PdfExtractionMode
from .text import PlainTextDocument
from .docx import (
    DocxDocument,
    InlineSegment,
    InlineSegmentKind,
    SegmentReplacement,
    TextSegment,
    paragraph_to_inline_segments,
    rebuild_paragraph_from_inline,
    _advances_offset,
    _visible_text,
    _visible_len,
    _split_visible_offset,
    _insert_visible,
    _replace_visible_range,
    _rpr_at,
    _index_at_visible_offset,
    _copy_segment,
)

__all__ = [
    "DOCX_MIME",
    "MD_MIME",
    "PDF_MIME",
    "TXT_MIME",
    "DocumentBytes",
    "DocumentError",
    "DocumentKind",
    "DocxDocument",
    "InlineSegment",
    "InlineSegmentKind",
    "PdfExtractionMode",
    "PdfDocument",
    "PlainTextDocument",
    "SegmentReplacement",
    "TextSegment",
    "paragraph_to_inline_segments",
    "rebuild_paragraph_from_inline",
    "detect_document_type",
    "document_to_bytes",
    "load_document",
]
