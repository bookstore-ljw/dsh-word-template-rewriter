"""Typed slot values.  Every value is either literal content or a pointer into
source material; there is no "generate" variant on purpose."""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class ImageSpec:
    """An image copied into a slot at a size recorded on the template side."""

    path: str
    width: int | None = None  # EMU
    height: int | None = None  # EMU


@dataclass(frozen=True)
class FormulaSpec:
    """A deep-copied OMML node; stays editable inside Word."""

    element: object
    display: bool = True


@dataclass(frozen=True)
class TableSpec:
    rows: list[list[str]] = field(default_factory=list)

    def __post_init__(self) -> None:
        object.__setattr__(self, "rows", [list(row) for row in self.rows])
