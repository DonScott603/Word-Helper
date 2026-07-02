"""Apply character formatting to found text inside .docx files.

Finds occurrences of a search string and applies ONLY the formatting attributes
the caller explicitly enables (bold, italic, underline, strikethrough, font
name, size, colour). Any attribute left as ``None`` is not touched, so existing
formatting is preserved. Formatting is applied to exactly the matched
characters — runs are split at the match boundaries so nothing else changes.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field

from docx import Document
from docx.shared import Pt, RGBColor
from docx.text.run import Run

from docx_replace import _iter_document_regions


@dataclass
class FormatSpec:
    """Which attributes to apply. ``None`` means 'leave unchanged'."""
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    name: str | None = None          # font family
    size_pt: float | None = None     # font size in points
    color: str | None = None         # 6-hex-digit RGB, e.g. "FF0000"

    def is_empty(self) -> bool:
        return all(
            getattr(self, f) is None
            for f in ("bold", "italic", "underline", "strike", "name", "size_pt", "color")
        )


@dataclass
class FormatResult:
    path: str
    matches: int = 0
    ok: bool = True
    error: str = ""
    locations: dict = field(default_factory=dict)


def _apply_to_run(run, spec: FormatSpec):
    font = run.font
    if spec.bold is not None:
        font.bold = spec.bold
    if spec.italic is not None:
        font.italic = spec.italic
    if spec.underline is not None:
        font.underline = spec.underline
    if spec.strike is not None:
        font.strike = spec.strike
    if spec.name is not None:
        font.name = spec.name
    if spec.size_pt is not None:
        font.size = Pt(spec.size_pt)
    if spec.color is not None:
        font.color.rgb = RGBColor.from_string(spec.color)


def _runs_and_offsets(paragraph):
    runs = paragraph.runs
    offsets = []
    acc = 0
    for r in runs:
        offsets.append(acc)
        acc += len(r.text)
    return runs, offsets


def _split_run(run, local_offset):
    """Split ``run`` so the first ``local_offset`` chars stay and the rest move
    to a new run (same formatting) inserted immediately after."""
    text = run.text
    if local_offset <= 0 or local_offset >= len(text):
        return
    new_r = copy.deepcopy(run._r)
    run.text = text[:local_offset]
    new_run = Run(new_r, run._parent)
    new_run.text = text[local_offset:]
    run._r.addnext(new_r)


def _ensure_boundary(paragraph, pos):
    """Guarantee a run boundary exactly at character position ``pos``."""
    runs, offsets = _runs_and_offsets(paragraph)
    for i, run in enumerate(runs):
        start = offsets[i]
        end = start + len(run.text)
        if start < pos < end:
            _split_run(run, pos - start)
            return


def _apply_to_range(paragraph, start, end, spec: FormatSpec):
    _ensure_boundary(paragraph, start)
    _ensure_boundary(paragraph, end)
    runs, offsets = _runs_and_offsets(paragraph)
    for i, run in enumerate(runs):
        s = offsets[i]
        e = s + len(run.text)
        if e > s and s >= start and e <= end:
            _apply_to_run(run, spec)


def format_in_paragraph(paragraph, find, spec, match_case=True):
    """Apply ``spec`` to every occurrence of ``find`` in the paragraph.
    Returns the number of occurrences formatted."""
    if not paragraph.runs or not find:
        return 0
    full = "".join(r.text for r in paragraph.runs)
    haystack = full if match_case else full.lower()
    needle = find if match_case else find.lower()

    ranges = []
    i = 0
    while True:
        j = haystack.find(needle, i)
        if j < 0:
            break
        ranges.append((j, j + len(find)))
        i = j + len(find)  # non-overlapping

    # Applying formatting never changes character positions, so the ranges stay
    # valid even as runs get split. Apply left-to-right.
    for start, end in ranges:
        _apply_to_range(paragraph, start, end, spec)
    return len(ranges)


def apply_formatting_in_document(path, find, spec, match_case=True, scope="everywhere"):
    """Open ``path`` and apply ``spec`` to every occurrence of ``find`` in scope.
    Returns (document, FormatResult). Does NOT save."""
    result = FormatResult(path=path)
    if spec.is_empty():
        result.ok = False
        result.error = "no formatting options selected"
        return None, result
    try:
        document = Document(path)
    except Exception as exc:
        result.ok = False
        result.error = str(exc)
        return None, result

    for region, paragraph in _iter_document_regions(document):
        if scope == "body" and region != "body":
            continue
        made = format_in_paragraph(paragraph, find, spec, match_case=match_case)
        if made:
            result.matches += made
            result.locations[region] = result.locations.get(region, 0) + made

    return document, result
