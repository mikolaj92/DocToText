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
| PDF | text layer only | new text PDF | Does not preserve original visual layout. OCR is not bundled. |

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
- PDF output is a newly rendered text PDF, not a patch of the original PDF page
  graphics.

## License

MIT
