"""Layout verification through the real Word engine.

The deterministic checks read OOXML; they cannot see page breaks, overflowing
images or a formula that Word refuses to lay out.  This module closes that gap on
Windows: Word exports the document to PDF, then the pages are rasterised so a
human - or the Router - can look at the actual result.
"""

from __future__ import annotations

from pathlib import Path
from typing import Callable

WD_EXPORT_FORMAT_PDF = 17
WD_STATISTIC_PAGES = 2


def export_pdf(docx_path: str | Path, pdf_path: str | Path) -> tuple[str, int]:
    """Render with the installed Word via COM; returns (pdf path, page count)."""
    try:
        import win32com.client
    except ImportError as error:  # pragma: no cover - platform dependent
        raise RuntimeError("Word COM rendering needs pywin32 on Windows") from error

    source = str(Path(docx_path).resolve())
    target = str(Path(pdf_path).resolve())
    Path(pdf_path).parent.mkdir(parents=True, exist_ok=True)

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    document = None
    try:
        document = word.Documents.Open(source, False, True, False)
        document.ExportAsFixedFormat(target, WD_EXPORT_FORMAT_PDF)
        pages = int(document.ComputeStatistics(WD_STATISTIC_PAGES))
    finally:
        if document is not None:
            document.Close(0)
        word.Quit()
    return target, pages


def render_png(
    pdf_path: str | Path, png_path: str | Path, page: int = 0, dpi: int = 120
) -> tuple[str, int]:
    """Rasterise one page; returns (png path, page count)."""
    import pymupdf

    document = pymupdf.open(str(pdf_path))
    try:
        if page >= document.page_count:
            raise ValueError(f"page {page} out of range (0..{document.page_count - 1})")
        pixmap = document[page].get_pixmap(dpi=dpi)
        Path(png_path).parent.mkdir(parents=True, exist_ok=True)
        pixmap.save(str(png_path))
        return str(png_path), document.page_count
    finally:
        document.close()


def render_many(
    docx_paths: list[str | Path],
    out_dir: str | Path,
    dpi: int = 80,
    on_step: "Callable[[str, str], None] | None" = None,
) -> dict[str, str]:
    """Render several documents with ONE Word instance.

    Starting Word costs seconds; reusing it for a whole step sequence makes the
    live preview practical (roughly 0.5s per document after the first).
    """
    try:
        import win32com.client
    except ImportError as error:  # pragma: no cover - platform dependent
        raise RuntimeError("Word COM rendering needs pywin32 on Windows") from error

    destination = Path(out_dir)
    destination.mkdir(parents=True, exist_ok=True)
    manifest: dict[str, str] = {}

    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    try:
        for raw in docx_paths:
            source = Path(raw).resolve()
            pdf = destination / f"{source.stem}.pdf"
            png = destination / f"{source.stem}.png"
            document = None
            try:
                document = word.Documents.Open(str(source), False, True, False)
                document.ExportAsFixedFormat(str(pdf.resolve()), WD_EXPORT_FORMAT_PDF)
            finally:
                if document is not None:
                    document.Close(0)
            render_png(pdf, png, page=0, dpi=dpi)
            manifest[str(source)] = str(png.resolve())
            if on_step is not None:
                on_step(str(source), str(png.resolve()))
    finally:
        word.Quit()
    return manifest
