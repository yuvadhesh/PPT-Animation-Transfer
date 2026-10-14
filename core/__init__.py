"""
PPT Animation Transfer Core Module
"""
from core.engine import transfer_presentation_animations
from core.models import (
    DeckTransferResult,
    EffectType,
    SlideStatus,
    SlideTransferResult,
    TriggerType,
)
from core.parser import PresentationParseError, validate_pptx_file

__all__ = [
    "transfer_presentation_animations",
    "PresentationParseError",
    "validate_pptx_file",
    "DeckTransferResult",
    "SlideTransferResult",
    "SlideStatus",
    "EffectType",
    "TriggerType",
]
