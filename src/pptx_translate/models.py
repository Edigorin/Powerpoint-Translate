from __future__ import annotations

from dataclasses import dataclass
from typing import Optional


@dataclass
class TranslatableUnit:
    """
    Represents a single text node to be translated.
    """

    id: str
    location: str
    source_text: str
    translated_text: Optional[str] = None
    context: Optional[str] = None
