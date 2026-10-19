"""Universal Excel splitter engine."""
__version__ = "0.1.0"
__author__ = "Molyuu"
__email__ = "i@molyuu.com"

from .engine import (  # noqa: F401
    FILE_NAME_COL, SHEET_NAME_COL, INPUT_FILE, SOURCE_SHEET,
    Options, analyze, write_outputs, collect_inputs,
)
