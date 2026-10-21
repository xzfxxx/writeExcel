"""writeExcel：按 txt 里的 JSON 数据填充 Excel 模板。"""

from .excel_writer import FillError, FillReport, fill, normalize_entries
from .parser import ParseError, loads

__all__ = ["FillError", "FillReport", "fill", "normalize_entries", "ParseError", "loads", "__version__"]

__version__ = "1.0.0"
