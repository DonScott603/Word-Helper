"""Core image find-and-replace engine for .docx files.

Given an "old" image file and a "new" image file, this finds every embedded
image in a document that matches the old one and swaps in the new bytes. The
replacement is displayed at the original image's size/position (we only change
the image data, not the drawing that references it), so layout is preserved.

Matching strategies:
  * exact      — the embedded image bytes are identical to the old file (hash).
  * dimensions — (optional fallback) same pixel width/height and format. Useful
                 because Word sometimes re-compresses images on insert, so a
                 byte-exact match can miss.

Images inside headers and footers are covered too, because we iterate over every
part in the package. The document-preview thumbnail (/docProps/thumbnail.*) is
never touched.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass, field

from docx import Document
from docx.image.image import Image as DocxImage
from docx.opc.packuri import PackURI


@dataclass
class ImageInfo:
    content_type: str
    ext: str
    px_width: int
    px_height: int
    size: int


@dataclass
class ImageReplaceResult:
    path: str
    exact: int = 0          # images swapped via exact hash match
    dimensions: int = 0     # images swapped via dimension fallback
    ok: bool = True
    error: str = ""
    format_changed: bool = False
    notes: list = field(default_factory=list)

    @property
    def total(self) -> int:
        return self.exact + self.dimensions


def read_image_info(path_or_bytes) -> ImageInfo:
    """Return format/dimension info for an image given a path or raw bytes."""
    if isinstance(path_or_bytes, (bytes, bytearray)):
        blob = bytes(path_or_bytes)
    else:
        with open(path_or_bytes, "rb") as fh:
            blob = fh.read()
    im = DocxImage.from_blob(blob)
    return ImageInfo(
        content_type=im.content_type,
        ext=im.ext,
        px_width=im.px_width,
        px_height=im.px_height,
        size=len(blob),
    )


def _iter_image_parts(document):
    """Yield every real content-image part in the package, skipping the
    document-preview thumbnail."""
    for part in document.part.package.iter_parts():
        ct = getattr(part, "content_type", "") or ""
        if not ct.startswith("image/"):
            continue
        if str(part.partname).startswith("/docProps/"):
            continue
        yield part


def replace_images_in_document(path, old_bytes, new_bytes, match_by_dimensions=False):
    """Open ``path``, replace every image matching ``old_bytes`` with
    ``new_bytes``, and return (document, ImageReplaceResult). Does NOT save.
    """
    result = ImageReplaceResult(path=path)
    try:
        old_hash = hashlib.sha1(old_bytes).hexdigest()
        old_info = read_image_info(old_bytes)
        new_info = read_image_info(new_bytes)
    except Exception as exc:
        result.ok = False
        result.error = f"could not read source images: {exc}"
        return None, result

    try:
        document = Document(path)
    except Exception as exc:
        result.ok = False
        result.error = str(exc)
        return None, result

    for part in _iter_image_parts(document):
        try:
            blob = part.blob
        except Exception:
            continue

        method = None
        if hashlib.sha1(blob).hexdigest() == old_hash:
            method = "exact"
        elif match_by_dimensions:
            try:
                info = read_image_info(blob)
            except Exception:
                info = None
            if info and (info.px_width, info.px_height) == (
                old_info.px_width,
                old_info.px_height,
            ) and info.content_type == old_info.content_type:
                method = "dimensions"

        if method is None:
            continue

        # Swap the bytes.
        part._blob = new_bytes

        # If the new image is a different format, update the part's content type
        # and file extension so Word still recognises it.
        current_ext = str(part.partname).rsplit(".", 1)[-1].lower()
        if new_info.content_type != part.content_type or new_info.ext != current_ext:
            part._content_type = new_info.content_type
            base = str(part.partname).rsplit(".", 1)[0]
            new_partname = PackURI(f"{base}.{new_info.ext}")
            try:
                part.partname = new_partname
            except AttributeError:
                part._partname = new_partname
            result.format_changed = True

        if method == "exact":
            result.exact += 1
        else:
            result.dimensions += 1

    return document, result
