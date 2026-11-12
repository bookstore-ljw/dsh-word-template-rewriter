"""Live Word session helpers.

The panel can mirror the fill into a real Word window: after every atomic
operation the current step document is inserted into the open document, so the
user watches the document change instead of a static preview image.  Changed
paragraphs arrive already shaded by the fill step, which is how the "where did it
change" colour survives into Word.
"""

from __future__ import annotations

import os
from pathlib import Path


def _app():
    import win32com.client  # noqa: PLC0415 - platform dependent

    try:
        word = win32com.client.GetActiveObject("Word.Application")
    except Exception:  # noqa: BLE001 - no Word running yet
        word = win32com.client.DispatchEx("Word.Application")
    word.Visible = True
    return word


def _same(left: str, right: str) -> bool:
    return os.path.normcase(os.path.abspath(left)) == os.path.normcase(os.path.abspath(right))


def _find(word, target: str):
    resolved = os.path.abspath(target)
    for document in word.Documents:
        if _same(document.FullName, resolved):
            return document
    return None


def open_file(target: str | Path) -> str:
    path = os.path.abspath(str(target))
    word = _app()
    document = _find(word, path)
    if document is None:
        document = word.Documents.Open(path)
    document.Activate()
    return path


def sync(target: str | Path, source: str | Path, save: bool = True, focus: int | None = None) -> str:
    """Replace the open document's body with ``source`` (keeps the window)."""
    target_path = os.path.abspath(str(target))
    source_path = os.path.abspath(str(source))
    word = _app()
    document = _find(word, target_path)
    if document is None:
        document = word.Documents.Open(target_path)
    document.Content.Delete()
    document.Content.InsertFile(source_path)
    if focus is not None:
        _focus(document, focus)
    if save:
        document.Save()
    return target_path


def _focus(document, index: int) -> None:
    """Clear old colour, select the changed paragraph and scroll Word to it."""
    try:
        document.Content.HighlightColorIndex = 0
        paragraph = document.Paragraphs(index + 1)
        paragraph.Range.HighlightColorIndex = 7  # wdYellow
        paragraph.Range.Select()
        document.ActiveWindow.ScrollIntoView(paragraph.Range)
        document.Activate()
    except Exception:  # noqa: BLE001 - focus is cosmetic
        pass


def sync_block(target: str | Path, step: str | Path, index: int, save: bool = True) -> str:
    """Update one paragraph in place, highlight it and scroll the window to it."""
    target_path = os.path.abspath(str(target))
    word = _app()
    document = _find(word, target_path)
    if document is None:
        document = word.Documents.Open(target_path)
    from docx import Document  # noqa: PLC0415 - import cost on purpose

    source = Document(str(step))
    paragraphs = source.paragraphs
    if index >= len(paragraphs):
        return sync(target_path, step, save=save, focus=index)
    text = paragraphs[index].text
    try:
        document.Content.HighlightColorIndex = 0
        paragraph = document.Paragraphs(index + 1)
        paragraph.Range.Text = text
        paragraph.Range.HighlightColorIndex = 7
        paragraph.Range.Select()
        document.ActiveWindow.ScrollIntoView(paragraph.Range)
        document.Activate()
    except Exception:  # noqa: BLE001 - fall back to whole-document sync
        return sync(target_path, step, save=save, focus=index)
    if save:
        document.Save()
    return target_path
