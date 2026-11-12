"""Deterministic DOCX template filling and quality gates.

Two halves, deliberately separated:

* ``fill``   - copies content into a template without touching template formatting
* ``checks`` - read-only, rule driven verification of the produced document

Nothing in this package generates prose.  Slots are resolved from a plan that
points at source material; the engine only moves content into place.
"""

from .fill import FillError, FillPlan, FillResult, fill_docx, load_plan
from .checks import Finding, run_checks
from .router import Batch, BatchResult, Correction, RouterError, run_router

__all__ = [
    "FillError",
    "FillPlan",
    "FillResult",
    "fill_docx",
    "load_plan",
    "Finding",
    "run_checks",
    "Batch",
    "BatchResult",
    "Correction",
    "RouterError",
    "run_router",
]
__version__ = "0.1.0"
