from __future__ import annotations

from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

import fitz
import pytest
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from test_docx import write_docx

from doctotext import (
    DOCX_MIME,
    MD_MIME,
    PDF_MIME,
    TXT_MIME,
    DocumentError,
    DocumentKind,
    PdfExtractionMode,
    detect_document_type,
    document_to_bytes,
    load_document,
)

UNICODE_FONT_CANDIDATES = (
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/dejavu/DejaVuSans.ttf",
    "/usr/local/share/fonts/DejaVuSans.ttf",
    "/System/Library/Fonts/Supplemental/Arial.ttf",
    "/Library/Fonts/Arial Unicode.ttf",
    "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
)


def _pdf_bytes(*pages: str, font_name: str = "Helvetica") -> bytes:
    output = BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    pdf.setFont(font_name, 12)
    for index, text in enumerate(pages):
        if index:
            pdf.showPage()
            pdf.setFont(font_name, 12)
        pdf.drawString(48, 760, text)
    pdf.save()
    return output.getvalue()


def _blank_pdf_bytes() -> bytes:
    output = BytesIO()
    pdf = canvas.Canvas(output, pagesize=A4)
    pdf.showPage()
    pdf.save()
    return output.getvalue()


def _positioned_pdf_bytes(*lines: str) -> bytes:
    pdf = fitz.open()
    page = pdf.new_page(width=595, height=842)
    y = 100
    for line in lines:
        page.insert_text((48, y), line, fontsize=10)
        y += 13
    return pdf.tobytes()


def _pdf_text(data: bytes) -> str:
    pdf = fitz.open(stream=data, filetype="pdf")
    return _normalize_text("\n".join(page.get_text("text") or "" for page in pdf))


def _pdf_page_count(data: bytes) -> int:
    pdf = fitz.open(stream=data, filetype="pdf")
    return pdf.page_count


def _normalize_text(text: str) -> str:
    return text.replace("\xa0", " ")


def _unicode_font_name() -> str:
    font_path = next((path for path in UNICODE_FONT_CANDIDATES if Path(path).exists()), None)
    if font_path is None:
        pytest.skip("Unicode font unavailable for PDF fixture")

    font_name = "DocToTextTestUnicode"
    if font_name not in pdfmetrics.getRegisteredFontNames():
        pdfmetrics.registerFont(TTFont(font_name, font_path))
    return font_name


def test_detects_docx_from_bytes_before_metadata(tmp_path: Path) -> None:
    input_path = tmp_path / "input.docx"
    write_docx(input_path)
    data = input_path.read_bytes()

    detection = detect_document_type("upload.bin", "application/octet-stream", data)
    document = load_document("upload.bin", "application/octet-stream", data)

    assert detection.kind == DocumentKind.DOCX
    assert detection.source == "signature"
    assert document.texts == ["Hello world", "Second paragraph", "Header text"]


def test_detects_pdf_from_bytes_before_metadata() -> None:
    data = _pdf_bytes("Jan Kowalski")

    detection = detect_document_type("upload.bin", "application/octet-stream", data)
    document = load_document("upload.bin", "application/octet-stream", data)

    assert detection.kind == DocumentKind.PDF
    assert detection.source == "signature"
    assert document.texts == ["Jan Kowalski\n"]


def test_rejects_unknown_binary_document() -> None:
    with pytest.raises(DocumentError, match="Nieobsługiwany typ dokumentu"):
        load_document("upload.bin", "application/octet-stream", b"\x00\x01\x02\x03")


def test_load_docx_document_and_write_docx_bytes(tmp_path: Path) -> None:
    input_path = tmp_path / "input.docx"
    write_docx(input_path)

    document = load_document("input.docx", DOCX_MIME, input_path.read_bytes())
    document.apply_texts(["One", "Two", "Three"])
    output = document_to_bytes(document, "input.docx")

    assert output.filename == "input.anonimizowany.docx"
    assert output.content_type == DOCX_MIME
    with ZipFile(BytesIO(output.data)) as docx:
        assert "word/document.xml" in docx.namelist()


def test_load_pdf_document_and_write_pdf_bytes() -> None:
    document = load_document("input.pdf", PDF_MIME, _pdf_bytes("Jan Kowalski"))

    assert document.extraction_mode == PdfExtractionMode.TEXT_LAYER
    assert document.texts == ["Jan Kowalski\n"]
    document.apply_texts(["<PERSON>"])
    output = document_to_bytes(document, "input.pdf")

    assert output.filename == "input.anonimizowany.pdf"
    assert output.content_type == PDF_MIME
    assert output.data.startswith(b"%PDF")
    output_text = _pdf_text(output.data)
    assert "<PERSON>" in output_text
    assert "Jan Kowalski" not in output_text


def test_load_pdf_joins_soft_wrapped_lines_with_space() -> None:
    data = _positioned_pdf_bytes(
        "X" * 20
        + " Dane testowe obejmuja adres e-mail x, numer telefonu +48 514 222 333, rachun",
        "1140 2004 0000 3102 1234 5678, pojazd KR 7MZ18.",
    )

    raw_text = fitz.open(stream=data, filetype="pdf")[0].get_text("text")
    document = load_document("input.pdf", PDF_MIME, data)

    assert "rachun\n1140" in raw_text
    assert "rachun 1140" in document.texts[0]
    assert "rachun\n1140" not in document.texts[0]


def test_pdf_write_preserves_polish_text_and_page_count() -> None:
    data = _pdf_bytes(
        "Dane nie są fikcyjne. Zażółć gęślą jaźń. Jan Kowalski PESEL 44051401359",
        font_name=_unicode_font_name(),
    )
    document = load_document("input.pdf", PDF_MIME, data)

    assert "Dane nie są fikcyjne" in _normalize_text(document.texts[0])
    assert "Zażółć gęślą jaźń" in _normalize_text(document.texts[0])

    anonymized = (
        document.texts[0]
        .replace("Jan Kowalski", "****")
        .replace("44051401359", "****")
    )
    document.apply_texts([anonymized])
    output = document_to_bytes(document, "input.pdf")

    output_text = _pdf_text(output.data)
    assert _pdf_page_count(output.data) == _pdf_page_count(data)
    assert "Dane nie są fikcyjne" in output_text
    assert "Zażółć gęślą jaźń" in output_text
    assert "Jan Kowalski" not in output_text
    assert "44051401359" not in output_text


def test_pdf_write_keeps_original_page_count() -> None:
    data = _pdf_bytes("Jan Kowalski", "Anna Nowak")
    document = load_document("input.pdf", PDF_MIME, data)
    document.apply_texts(["<PERSON>\n", "<PERSON>\n"])
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)

    assert _pdf_page_count(output.data) == 2
    assert "<PERSON>" in output_text
    assert "Jan Kowalski" not in output_text
    assert "Anna Nowak" not in output_text


def test_pdf_write_redacts_changed_occurrence_by_offset() -> None:
    data = _pdf_bytes("Jan Kowalski oraz Jan Kowalski")
    document = load_document("input.pdf", PDF_MIME, data)
    document.apply_texts([document.texts[0].replace("Jan Kowalski", "<PERSON>", 1)])
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)

    assert "<PERSON>" in output_text
    assert output_text.count("Jan Kowalski") == 1


def test_pdf_write_rebuilds_page_when_text_is_inserted() -> None:
    data = _pdf_bytes("Pierwsze zdanie.")
    document = load_document("input.pdf", PDF_MIME, data)
    document.apply_texts(
        [document.texts[0].replace("Pierwsze zdanie.", "Pierwsze zdanie. Drugie zdanie.")]
    )
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)

    assert _pdf_page_count(output.data) == 1
    assert "Pierwsze zdanie. Drugie zdanie." in output_text


def test_pdf_write_rebuilds_page_when_replacement_is_longer() -> None:
    data = _pdf_bytes("Status: OK")
    document = load_document("input.pdf", PDF_MIME, data)
    document.apply_texts([document.texts[0].replace("OK", "bardzo dobrze")])
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)

    assert _pdf_page_count(output.data) == 1
    assert "Status: bardzo dobrze" in output_text
    assert "Status: OK" not in output_text


def test_pdf_write_removes_deleted_text() -> None:
    data = _pdf_bytes("Alpha Beta Gamma")
    document = load_document("input.pdf", PDF_MIME, data)
    document.apply_texts([document.texts[0].replace("Beta ", "")])
    output = document_to_bytes(document, "input.pdf")
    output_text = _pdf_text(output.data)

    assert _pdf_page_count(output.data) == 1
    assert "Alpha" in output_text
    assert "Gamma" in output_text
    assert "Beta" not in output_text


def test_pdf_without_text_layer_requires_ocr() -> None:
    with pytest.raises(DocumentError, match="wymaga OCR"):
        load_document("scan.pdf", PDF_MIME, _blank_pdf_bytes())


def test_load_text_document_and_write_txt_bytes() -> None:
    document = load_document("input.txt", "text/plain", "Zażółć".encode("cp1250"))

    assert document.texts == ["Zażółć"]
    document.apply_texts(["<TEXT>"])
    output = document_to_bytes(document, "input.txt")

    assert output.filename == "input.anonimizowany.txt"
    assert output.content_type == TXT_MIME
    assert output.data == b"<TEXT>"


def test_load_markdown_document_and_write_markdown_bytes() -> None:
    document = load_document("notes.md", "", b"# Title\n\nOld")

    assert document.texts == ["# Title\n\nOld"]
    document.apply_texts(["# Title\n\nNew"])
    output = document_to_bytes(document, "notes.md")

    assert output.filename == "notes.anonimizowany.md"
    assert output.content_type == MD_MIME
    assert output.data == b"# Title\n\nNew"
