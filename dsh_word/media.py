"""Image sizing: the template decides the size, the source only supplies pixels."""

from __future__ import annotations

import re
from pathlib import Path

EMU_PER_INCH = 914_400
_UNITS = {
    "in": 914_400,
    "cm": 360_000,
    "mm": 36_000,
    "pt": 12_700,
    "emu": 1,
    "px": 9_525,
}
_LENGTH_RE = re.compile(r"^\s*([0-9]*\.?[0-9]+)\s*(in|cm|mm|pt|emu|px)?\s*$", re.I)


def parse_length(value) -> int | None:
    """Parse ``"2.5in"`` / ``"8cm"`` / ``12`` (EMU) into EMU."""
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError(f"not a length: {value!r}")
    if isinstance(value, (int, float)):
        return int(value)
    match = _LENGTH_RE.match(str(value))
    if not match:
        raise ValueError(f"not a length: {value!r}")
    number, unit = match.group(1), (match.group(2) or "in").lower()
    return int(round(float(number) * _UNITS[unit]))


def image_extent(
    path: str | Path,
    width: int | None = None,
    height: int | None = None,
    max_width: int | None = None,
) -> tuple[int, int]:
    """Return ``(cx, cy)`` in EMU, preserving aspect ratio.

    Precedence: explicit width, then explicit height, then the image's own DPI.
    ``max_width`` only ever shrinks the result, never grows it.
    """
    from PIL import Image

    with Image.open(path) as image:
        pixel_width, pixel_height = image.size
        dpi = image.info.get("dpi") or (96.0, 96.0)
    dpi_x = float(dpi[0]) or 96.0
    dpi_y = float(dpi[1]) or 96.0
    aspect = pixel_height / pixel_width if pixel_width else 1.0

    if width:
        resolved_width = int(width)
        resolved_height = int(height) if height else int(round(resolved_width * aspect))
    elif height:
        resolved_height = int(height)
        resolved_width = int(round(resolved_height / aspect)) if aspect else int(height)
    else:
        resolved_width = int(round(pixel_width / dpi_x * EMU_PER_INCH))
        resolved_height = int(round(pixel_height / dpi_y * EMU_PER_INCH))

    if max_width and resolved_width > max_width:
        scale = max_width / resolved_width
        resolved_width = int(max_width)
        resolved_height = int(round(resolved_height * scale))
    return resolved_width, resolved_height
