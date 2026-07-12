from __future__ import annotations

import copy
from collections.abc import Iterable
from dataclasses import dataclass, replace
from io import BytesIO
from pathlib import Path
from typing import Any
from zipfile import ZipFile

from docx import Document as PyDocxDocument
from docx.document import Document as DocxDocumentType
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_P = qn("w:p")
W_R = qn("w:r")
W_T = qn("w:t")


@dataclass(frozen=True)
class TextSegment:
    """Stable textual segment inside a DOCX.

    container_id examples:
      body:p:0
      header:0
      table:0:r:0:c:0:p:0   (for table cells)
    paragraph_index is the global order index (body + tables + headers/footers).
    """

    id: str
    text: str
    part: str
    index: int
    container_id: str | None = None
    paragraph_index: int | None = None


@dataclass(frozen=True)
class SegmentReplacement:
    """Replacement targeting a segment or a sub-range inside it.

    Offsets are in characters of the segment's text.
    If start_offset and end_offset are None -> whole segment.
    """

    container_id: str | None = None
    id: str | None = None
    text: str = ""
    start_offset: int | None = None
    end_offset: int | None = None


@dataclass
class _ParaRef:
    """Internal mapping from our segment to python-docx paragraph + metadata."""

    id: str
    container_id: str
    paragraph_index: int
    paragraph: Paragraph
    part_name: str  # "body", "header:0", "table:0:r:0:c:0", etc.


class DocxDocument:
    """DOCX editing surface backed by python-docx (the proper library for the format).

    - Stable container_id + paragraph_index addressing.
    - Whole segment or offset-based partial replacements (run splitting).
    - Preserves formatting because we operate on runs.
    - Roundtrips via python-docx save.
    """

    def __init__(
        self,
        doc: DocxDocumentType,
        segments: list[TextSegment],
        refs: list[_ParaRef],
    ) -> None:
        self._doc = doc
        self._segments = segments
        self._refs = refs  # index-aligned with segments

    @classmethod
    def open(cls, path: str | Path) -> DocxDocument:
        path = Path(path)
        doc = PyDocxDocument(str(path))
        return cls._from_pydocx(doc)

    @classmethod
    def open_bytes(cls, data: bytes) -> DocxDocument:
        doc = PyDocxDocument(BytesIO(data))
        return cls._from_pydocx(doc)

    @classmethod
    def _from_pydocx(cls, doc: DocxDocumentType) -> DocxDocument:
        segments: list[TextSegment] = []
        refs: list[_ParaRef] = []

        paragraph_index = 0

        def add_paragraphs(paragraphs: list[Paragraph], prefix: str) -> None:
            nonlocal paragraph_index
            for local_idx, para in enumerate(paragraphs):
                text = "".join(run.text for run in para.runs)
                if not text:
                    # still count for paragraph_index stability (same as dike/posejdon)
                    paragraph_index += 1
                    continue

                container_id = f"{prefix}:p:{local_idx}" if prefix != "body" else f"body:p:{paragraph_index}"
                # For body we use global paragraph_index to match previous contract
                if prefix == "body":
                    cid = f"body:p:{paragraph_index}"
                else:
                    cid = f"{prefix}:p:{local_idx}"

                seg_id = f"s{len(segments)}"

                segments.append(
                    TextSegment(
                        id=seg_id,
                        text=text,
                        part="word/document.xml" if prefix.startswith("body") or prefix.startswith("table") else f"word/{prefix.split(':')[0]}.xml",
                        index=local_idx,
                        container_id=cid,
                        paragraph_index=paragraph_index,
                    )
                )
                refs.append(
                    _ParaRef(
                        id=seg_id,
                        container_id=cid,
                        paragraph_index=paragraph_index,
                        paragraph=para,
                        part_name=prefix,
                    )
                )
                paragraph_index += 1

        # Body paragraphs
        add_paragraphs(list(doc.paragraphs), "body")

        # Tables (cell paragraphs)
        for ti, table in enumerate(doc.tables):
            for ri, row in enumerate(table.rows):
                for ci, cell in enumerate(row.cells):
                    add_paragraphs(list(cell.paragraphs), f"table:{ti}:r:{ri}:c:{ci}")

        # Headers / Footers
        for si, section in enumerate(doc.sections):
            add_paragraphs(list(section.header.paragraphs), f"header:{si}")
            add_paragraphs(list(section.footer.paragraphs), f"footer:{si}")

        return cls(doc=doc, segments=segments, refs=refs)

    @property
    def segments(self) -> tuple[TextSegment, ...]:
        return tuple(self._segments)

    @property
    def texts(self) -> list[str]:
        return [s.text for s in self._segments]

    def to_markdown(self) -> str:
        blocks = [f"<!-- doctotext:{s.id} -->\n{s.text}" for s in self._segments]
        return "\n\n".join(blocks)

    # ------------------------------------------------------------------
    # Replacement API (supports full + offset ranges via python-docx)
    # ------------------------------------------------------------------

    def apply_texts(self, texts: Iterable[str], *, strict: bool = False) -> None:
        texts = list(texts)
        if len(texts) != len(self._segments):
            raise ValueError(f"expected {len(self._segments)} segments, got {len(texts)}")
        for i, txt in enumerate(texts):
            self._replace_full_segment(i, txt)

    def apply_replacements(
        self,
        replacements: list[dict[str, Any] | SegmentReplacement],
        *,
        strict: bool = False,
    ) -> None:
        by_container = {r.container_id: i for i, r in enumerate(self._segments)}
        by_id = {r.id: i for i, r in enumerate(self._segments)}

        for rep in replacements:
            if isinstance(rep, SegmentReplacement):
                idx = self._find_index(rep, by_container, by_id, strict)
                if idx is None:
                    continue
                self._apply_to_paragraph(idx, rep.text, rep.start_offset, rep.end_offset)
                continue

            # legacy / dict form
            idx = None
            if "container_id" in rep:
                idx = by_container.get(str(rep["container_id"]))
            elif "id" in rep:
                idx = by_id.get(str(rep["id"]))

            if idx is None:
                if strict:
                    raise ValueError(f"unknown target: {rep.get('container_id') or rep.get('id')}")
                continue

            text = str(rep.get("text", ""))
            start = rep.get("start_offset")
            end = rep.get("end_offset")
            self._apply_to_paragraph(idx, text, start, end)

    def apply_markdown(self, markdown: str, *, strict: bool = True) -> None:
        import re as _re
        by_id = {
            m.group("id"): m.group("text").rstrip("\n")
            for m in _re.finditer(r"<!-- doctotext:(?P<id>s\d+) -->\n(?P<text>.*?)(?=\n<!-- doctotext:s\d+ -->\n|\Z)", markdown, _re.DOTALL)
        }
        if strict:
            expected = {s.id for s in self._segments}
            actual = set(by_id.keys())
            missing = sorted(expected - actual)
            unknown = sorted(actual - expected)
            if missing or unknown:
                raise ValueError(f"markdown marker mismatch; missing={missing} unknown={unknown}")

        for i, seg in enumerate(self._segments):
            if seg.id in by_id:
                self._replace_full_segment(i, by_id[seg.id])

    # ------------------------------------------------------------------
    # Save / bytes
    # ------------------------------------------------------------------

    def save_docx(self, path: str | Path) -> None:
        self._doc.save(str(path))

    def to_bytes(self) -> bytes:
        buf = BytesIO()
        self._doc.save(buf)
        return buf.getvalue()

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def _find_index(
        self,
        rep: SegmentReplacement,
        by_container: dict[str, int],
        by_id: dict[str, int],
        strict: bool,
    ) -> int | None:
        if rep.container_id:
            idx = by_container.get(str(rep.container_id))
        elif rep.id:
            idx = by_id.get(str(rep.id))
        else:
            idx = None
        if idx is None and strict:
            raise ValueError(f"unknown replacement target: {rep.container_id or rep.id}")
        return idx

    def _replace_full_segment(self, index: int, text: str) -> None:
        ref = self._refs[index]
        para = ref.paragraph
        if not para.runs:
            # create one run
            run = para.add_run(text)
        else:
            # replace first run, clear the rest (preserves some formatting on first run)
            para.runs[0].text = text
            for run in para.runs[1:]:
                run.text = ""
        # update our view
        old = self._segments[index]
        self._segments[index] = replace(old, text=text)

    def _apply_to_paragraph(
        self,
        index: int,
        replacement: str,
        start: int | None,
        end: int | None,
    ) -> None:
        ref = self._refs[index]
        para = ref.paragraph
        full = "".join(r.text for r in para.runs)
        s = 0 if start is None else start
        e = len(full) if end is None else end

        if s == 0 and e == len(full):
            self._replace_full_segment(index, replacement)
            return

        # Use run-range logic (adapted from proven posejdon docx_runs)
        ranges = self._build_run_ranges(para)
        # split boundaries (right first)
        for r in reversed(ranges):
            if r[1] < e < r[2]:
                self._split_run(para, r[0], e - r[1])
                break
        ranges = self._build_run_ranges(para)
        for r in reversed(ranges):
            if r[1] < s < r[2]:
                self._split_run(para, r[0], s - r[1])
                break

        # now replace the affected runs
        ranges = self._build_run_ranges(para)
        affected = [r for r in ranges if r[1] >= s and r[2] <= e and r[2] > r[1]]
        if not affected:
            # fallback
            self._replace_full_segment(index, replacement)
            return

        first = True
        for r in affected:
            run = para.runs[r[0]]
            if first:
                run.text = replacement
                first = False
            else:
                run.text = ""

        # refresh our segment text
        new_text = "".join(r.text for r in para.runs)
        old = self._segments[index]
        self._segments[index] = replace(old, text=new_text)

    def _build_run_ranges(self, paragraph: Paragraph) -> list[tuple[int, int, int]]:
        out: list[tuple[int, int, int]] = []
        cur = 0
        for i, run in enumerate(paragraph.runs):
            ln = len(run.text)
            out.append((i, cur, cur + ln))
            cur += ln
        return out

    def _split_run(self, paragraph: Paragraph, run_index: int, offset: int) -> int:
        if offset <= 0:
            return run_index
        run = paragraph.runs[run_index]
        if offset >= len(run.text):
            return run_index + 1
        left = run.text[:offset]
        right = run.text[offset:]
        run.text = left

        # clone the underlying XML element
        cloned = copy.deepcopy(run._element)
        run._element.addnext(cloned)

        new_run = paragraph.runs[run_index + 1]
        new_run.text = right
        return run_index + 1


# ----------------------------------------------------------------------
# Helper to expose for advanced users if needed
# ----------------------------------------------------------------------

def _paragraph_text(p: Paragraph) -> str:
    return "".join(r.text for r in p.runs)
