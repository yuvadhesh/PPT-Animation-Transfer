from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
import re
from typing import Any, Dict, List, Optional


def normalize_text(text: str) -> str:
    """Lowercase and strip all whitespace for invariant text matching."""
    if not text:
        return ""
    return re.sub(r"\s+", "", text).lower()


class EffectType(str, Enum):
    ENTRANCE = "Entrance"
    EMPHASIS = "Emphasis"
    EXIT = "Exit"
    MOTION_PATH = "Motion Path"
    MEDIA = "Media"
    OTHER = "Other"


class TriggerType(str, Enum):
    ON_CLICK = "On Click"
    WITH_PREVIOUS = "With Previous"
    AFTER_PREVIOUS = "After Previous"
    INTERACTIVE = "Interactive (Trigger)"
    OTHER = "Other"


class SlideStatus(str, Enum):
    PASS = "Pass"
    PARTIAL = "Partial"
    MANUAL_REVIEW = "Manual Review"
    FAILED = "Failed"
    SKIPPED = "Skipped"
    UNMATCHED = "Unmatched"


@dataclass
class BoundingBox:
    x: int
    y: int
    cx: int
    cy: int

    @property
    def left(self) -> int:
        return self.x

    @property
    def top(self) -> int:
        return self.y

    @property
    def width(self) -> int:
        return self.cx

    @property
    def height(self) -> int:
        return self.cy

    @property
    def right(self) -> int:
        return self.x + self.cx

    @property
    def bottom(self) -> int:
        return self.y + self.cy

    @property
    def area(self) -> float:
        return float(max(0, self.cx) * max(0, self.cy))

    def intersection_area(self, other: BoundingBox) -> float:
        ix1 = max(self.left, other.left)
        iy1 = max(self.top, other.top)
        ix2 = min(self.right, other.right)
        iy2 = min(self.bottom, other.bottom)
        if ix2 > ix1 and iy2 > iy1:
            return float((ix2 - ix1) * (iy2 - iy1))
        return 0.0

    def contains_ratio(self, other: BoundingBox) -> float:
        """Portion of `other` that lies inside `self` (0.0 to 1.0)."""
        if other.area <= 0:
            return 0.0
        return self.intersection_area(other) / other.area

    def overlap_ratio(self, other: BoundingBox) -> float:
        """Intersection over union / min area overlap."""
        denom = min(self.area, other.area)
        if denom <= 0:
            return 0.0
        return self.intersection_area(other) / denom

    def to_str(self) -> str:
        return f"x={self.x}, y={self.y}, w={self.cx}, h={self.cy}"


@dataclass
class ShapeInfo:
    shape_id: str
    shape_name: str
    shape_type: str
    text: str = ""
    normalized_text: str = ""
    left: int = 0
    top: int = 0
    width: int = 0
    height: int = 0
    bbox: Optional[BoundingBox] = None
    image_hash: Optional[str] = None
    is_group: bool = False
    parent_id: Optional[str] = None
    descendant_ids: List[str] = field(default_factory=list)
    combined_text: str = ""
    normalized_combined_text: str = ""
    paragraphs: List[str] = field(default_factory=list)
    normalized_paragraphs: List[str] = field(default_factory=list)
    xml_element: Optional[Any] = None

    def __post_init__(self):
        if not self.normalized_text and self.text:
            self.normalized_text = normalize_text(self.text)
        if not self.combined_text and self.text and self.is_group:
            self.combined_text = self.text
        if not self.normalized_combined_text and self.combined_text:
            self.normalized_combined_text = normalize_text(self.combined_text)
        if self.paragraphs and not self.normalized_paragraphs:
            self.normalized_paragraphs = [normalize_text(p) for p in self.paragraphs]
        if self.bbox is None and (self.left or self.top or self.width or self.height):
            self.bbox = BoundingBox(self.left, self.top, self.width, self.height)


@dataclass
class SlideData:
    slide_index: int  # 1-indexed
    title: str = ""
    all_text: str = ""
    shapes: Dict[str, ShapeInfo] = field(default_factory=dict)
    has_timing: bool = False
    has_transition: bool = False
    slide_part: Optional[Any] = None
    xml_root: Optional[Any] = None


@dataclass
class EffectData:
    order: int
    target_spid: str
    shape_name: str
    preset_class: str
    preset_id: str
    preset_subtype: str
    node_type: str
    delay: str
    dur: str
    paragraph_range: Optional[str] = None
    effect_name: str = ""
    effect_type: EffectType = EffectType.OTHER
    trigger: TriggerType = TriggerType.ON_CLICK
    is_interactive: bool = False
    trigger_spid: Optional[str] = None
    xml_element: Optional[Any] = None

    @property
    def shape_id(self) -> str:
        return self.target_spid

    @property
    def delay_ms(self) -> str:
        return self.delay

    @property
    def duration_ms(self) -> str:
        return self.dur


# Alias for backward compatibility if needed
AnimationItem = EffectData


@dataclass
class SkippedAnimation:
    slide_index: int
    shape_id: str
    shape_name: str
    effect_type: str
    effect_name: str
    reason: str


@dataclass
class ManualReviewItem:
    slide_index: int
    shape_name: str = ""
    orig_shape_id: str = ""
    shape_id: str = ""
    text_snippet: str = ""
    shape_type: str = ""
    bounding_box: str = ""
    bbox: str = ""
    effect_name: str = ""
    trigger: str = ""
    reason: str = ""

    def __post_init__(self):
        if not self.shape_id and self.orig_shape_id:
            self.shape_id = self.orig_shape_id
        elif not self.orig_shape_id and self.shape_id:
            self.orig_shape_id = self.shape_id
        if not self.bbox and self.bounding_box:
            self.bbox = self.bounding_box
        elif not self.bounding_box and self.bbox:
            self.bounding_box = self.bbox


@dataclass
class SlideVerificationResult:
    slide_index: int
    remediated_slide_index: Optional[int]
    passed: bool
    status: SlideStatus
    original_effects_count: int
    output_effects_count: int
    exact_matches_count: int
    skipped_count: int
    click_steps_orig: int
    click_steps_output: int
    covered_count: int = 0
    not_covered_count: int = 0
    mapped_direct_count: int = 0
    mapped_group_split_count: int = 0
    mapped_paragraph_level_count: int = 0
    manual_review_count: int = 0
    issues: List[str] = field(default_factory=list)


@dataclass
class SlideTransferResult:
    slide_index: int  # Original slide index
    remediated_slide_index: Optional[int] = None
    status: SlideStatus = SlideStatus.PASS
    original_animation_count: int = 0
    copied_animation_count: int = 0
    skipped_animation_count: int = 0
    mapped_direct_count: int = 0
    mapped_group_split_count: int = 0
    mapped_paragraph_level_count: int = 0
    manual_review_count: int = 0
    covered_count: int = 0
    not_covered_count: int = 0
    copied_transition: bool = False
    shape_map: Dict[str, Any] = field(default_factory=dict)  # orig_id -> rem_id or list of rem_ids
    animations: List[EffectData] = field(default_factory=list)
    skipped: List[SkippedAnimation] = field(default_factory=list)
    manual_review_items: List[ManualReviewItem] = field(default_factory=list)
    verification: Optional[SlideVerificationResult] = None
    error_message: Optional[str] = None


@dataclass
class DeckTransferResult:
    success: bool
    message: str = ""
    original_slide_count: int = 0
    remediated_slide_count: int = 0
    slides_with_animations: int = 0
    total_animations_copied: int = 0
    total_animations_skipped: int = 0
    total_exact_matches: int = 0
    total_failed_slides: int = 0
    total_mapped_direct: int = 0
    total_mapped_group_split: int = 0
    total_mapped_paragraph_level: int = 0
    total_manual_review: int = 0
    total_covered_effects: int = 0
    total_not_covered_effects: int = 0
    is_100_percent_faithful: bool = False
    unmatched_slides_count: int = 0
    slide_results: List[SlideTransferResult] = field(default_factory=list)
    manual_review_items: List[ManualReviewItem] = field(default_factory=list)
    unmatched_original_slides: List[int] = field(default_factory=list)
    unmatched_remediated_slides: List[int] = field(default_factory=list)
    output_pptx_bytes: Optional[bytes] = None
    excel_report_bytes: Optional[bytes] = None
