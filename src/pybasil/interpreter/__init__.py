"""VBScript Tree-Walking Interpreter."""

from ._impl import Interpreter, run, VBSCRIPT_CONSTANTS  # noqa: F401

# Re-export runtime types that were historically importable from this module.
from ..runtime import VBScriptDictionary  # noqa: F401

__all__ = ["Interpreter", "run", "VBSCRIPT_CONSTANTS", "VBScriptDictionary"]