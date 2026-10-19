"""Universal Excel splitter engine."""
from .engine import (  # noqa: F401
    FILE_NAME_COL, SHEET_NAME_COL, INPUT_FILE, SOURCE_SHEET,
    Options, analyze, write_outputs, collect_inputs,
)
