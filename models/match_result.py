from dataclasses import dataclass
from typing import List, Optional

@dataclass
class MatchResult:
    source_index: int
    match_status: str
    species_id:Optional[int]
    scientific_name: str
    vernacular_name: str
    taxon_code: str
    note:str = ""
    