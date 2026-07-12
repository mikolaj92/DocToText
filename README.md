# DocToText

Extract editable text segments from documents, modify them in memory, and write
the changed text back as document bytes.

DocToText is useful when you need one small interface for document text
round-trips across plain text files, DOCX files, and text-layer PDFs.

## Status

Early library. The public API is small, but not stable yet.

## Supported Formats

| Format | Read | Write | Notes |
| --- | --- | --- | --- |
| TXT/Markdown/text files | yes | same text format when known | Decodes UTF-8, UTF-16, CP1250, and Latin-1 fallback. |
| DOCX | yes | DOCX | Edits WordprocessingML text nodes in place inside the package. |
| PDF | text layer only | PDF | Patches original pages when edits fit; rebuilds the PDF as flowing text when edits require reflow. OCR is not bundled. |

`load_document` detects the document kind from bytes first, then falls back to
MIME type and file extension.

## Installation

From GitHub:

```bash
python -m pip install git+https://github.com/mikolaj92/DocToText.git
```

With `uv`:

```bash
uv add git+https://github.com/mikolaj92/DocToText.git
```

## Basic Usage

```python
from doctotext import DOCX_MIME, document_to_bytes, load_document

document = load_document("input.docx", DOCX_MIME, input_bytes)

updated_texts = [
    text.replace("old", "new")
    for text in document.texts
]

document.apply_texts(updated_texts)
output = document_to_bytes(document, "input.docx")

output.filename      # input.anonimizowany.docx
output.content_type  # application/vnd.openxmlformats-officedocument.wordprocessingml.document
output.data          # bytes
```

All loaded documents expose the same editing surface:

```python
document.texts
document.segments
document.apply_texts([...])
```

## Type Detection

```python
from doctotext import DocumentKind, detect_document_type

detection = detect_document_type(
    "upload.bin",
    "application/octet-stream",
    input_bytes,
)

if detection.kind == DocumentKind.DOCX:
    ...
```

## DOCX Markdown Bridge

DOCX documents can be exported to marker-based Markdown, edited, then applied
back to the original document structure.

```python
from doctotext import DocxDocument

document = DocxDocument.open("input.docx")
markdown = document.to_markdown()

# Keep doctotext markers intact.
edited_markdown = markdown.replace("old", "new")

document.apply_markdown(edited_markdown)
document.save_docx("output.docx")
```

## Development

```bash
uv sync --extra dev
uv run ruff check .
uv run pytest
```

## Limits

- DOCX output preserves archive entries and edits only selected Word XML story
  parts.
- Edited DOCX XML parts are rewritten, so output is not byte-identical.
- Replacement text inherits formatting through original text-node spans. Large
  length changes can move style boundaries.
- Document layout can change when replacement text length changes.
- PDF input must have a text layer. Scanned or image-only PDFs return an
  OCR-required error.
- PDF output is best-effort because PDF is fixed-layout, not an editable text
  document.
- Small PDF edits that fit the original text rectangles are applied on top of
  original pages.
- PDF edits that insert text, expand replacements, or need reflow rebuild the
  PDF as flowing text. Page count may shrink or grow.
- If changed PDF text cannot be located safely, DocToText fails closed for that
  page by rebuilding the page text instead of leaking the original text.

## License

MIT

## High-fidelity DOCX editing

For advanced use cases (reversible redaction, reinjection, structured plans):

```python
from doctotext import DocxDocument

doc = DocxDocument.open("input.docx")

# Stable container addressing
print([s.container_id for s in doc.segments])
# ['body:p:0', 'body:p:1', 'header:0', ...]

# Structured replacement using container_id (preferred)
doc.apply_replacements([
    {"container_id": "body:p:0", "text": "New first paragraph"},
], strict=True)

# Or by internal id
doc.apply_replacements([
    {"id": "s0", "text": "Replacement"},
], strict=True)

doc.save_docx("output.docx")
```

`apply_replacements(..., strict=True)` will raise on unknown targets or structural drift.

Existing `apply_texts` and `apply_markdown` remain supported.
