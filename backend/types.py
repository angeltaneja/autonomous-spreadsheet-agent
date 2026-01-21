# Sheet Schema Types
from dataclasses import dataclass
from typing import List

@dataclass
class SheetSummary:
    sheet_name: str
    row_count: int
    columns: List[str]
