from __future__ import annotations

import math
import re
import unicodedata
from typing import Dict, List, Optional, Set, Tuple

from rapidfuzz import fuzz
from core.models import ManualReviewItem, ShapeInfo, SlideData


GREEK_TO_LATIN = {
    # Lowercase Greek
    "\u03b1": "a", "\u03b2": "b", "\u03b3": "g", "\u03b4": "d", "\u03b5": "e", "\u03f5": "e",
    "\u03b6": "z", "\u03b7": "h", "\u03b8": "q", "\u03d1": "q", "\u03b9": "i", "\u03ba": "k",
    "\u03bb": "l", "\u03bc": "m", "\u03bd": "v", "\u03be": "x", "\u03bf": "o", "\u03c0": "p", "\u03d6": "p",
    "\u03c1": "r", "\u03f1": "r", "\u03c2": "s", "\u03c3": "s", "\u03c4": "t", "\u03c5": "u",
    "\u03c6": "f", "\u03d5": "f", "\u03c7": "x", "\u03c8": "y", "\u03c9": "w",
    # Uppercase Greek
    "\u0391": "a", "\u0392": "b", "\u0393": "g", "\u0394": "d", "\u0395": "e",
    "\u0396": "z", "\u0397": "h", "\u0398": "q", "\u0399": "i", "\u039a": "k",
    "\u039b": "l", "\u039c": "m", "\u039d": "v", "\u039e": "x", "\u039f": "o", "\u03a0": "p",
    "\u03a1": "r", "\u03a3": "s", "\u03a4": "t", "\u03a5": "u",
    "\u03a6": "f", "\u03a7": "x", "\u03a8": "y", "\u03a9": "w",
}

MATH_MINUS_HYPHENS = {
    "\u2212": "-", "\u2010": "-", "\u2011": "-", "\u2012": "-", "\u2013": "-",
    "\u2014": "-", "\u2015": "-", "\ufe63": "-", "\uff0d": "-",
}


def normalize_equation_text(text: str) -> str:
    """
    Map symbol-font characters, Greek letters, minus signs, and combining marks to normalized Latin equivalents
    (e.g. w vs omega, D vs delta, hyphen vs minus).
    """
    if not text:
        return ""
    for k, v in MATH_MINUS_HYPHENS.items():
        text = text.replace(k, v)
    text = "".join(c for c in unicodedata.normalize("NFKD", text) if unicodedata.category(c) != "Mn")
    res = [GREEK_TO_LATIN.get(c, c) for c in text]
    return "".join(res)


def clean_text_alphanumeric(text: str) -> str:
    """Strip all non-alphanumeric characters for robust substring containment, equation-safe."""
    if not text:
        return ""
    eq_norm = normalize_equation_text(text)
    return re.sub(r"[\W_]+", "", eq_norm).lower()


def are_types_compatible(t1: str, t2: str) -> bool:
    """Check if two shape types are compatible."""
    if t1 == t2:
        return True
    u1, u2 = t1.upper(), t2.upper()
    is_g1 = "GROUP" in u1 or t1.lower() == "grpsp"
    is_g2 = "GROUP" in u2 or t2.lower() == "grpsp"
    if is_g1 != is_g2:
        return False

    # Tables
    t1_is_table = "TABLE" in u1 or "GRAPHICFRAME" in u1
    t2_is_table = "TABLE" in u2 or "GRAPHICFRAME" in u2
    if t1_is_table and t2_is_table:
        return True

    # OLE objects
    t1_is_ole = "OLE" in u1
    t2_is_ole = "OLE" in u2
    if t1_is_ole and t2_is_ole:
        return True

    # Pictures
    t1_is_pic = "PIC" in u1 or "PICTURE" in u1
    t2_is_pic = "PIC" in u2 or "PICTURE" in u2
    if t1_is_pic and t2_is_pic:
        return True

    # Connectors / lines / geometric shapes
    t1_is_line = any(k in u1 for k in ("LINE", "CONNECTOR"))
    t2_is_line = any(k in u2 for k in ("LINE", "CONNECTOR"))
    if t1_is_line and t2_is_line:
        return True

    # Geometric / auto shapes (arrows, braces, rectangles, callouts)
    t1_is_shape = any(k in u1 for k in ("AUTO_SHAPE", "SP", "GEOMETRY", "SHAPE"))
    t2_is_shape = any(k in u2 for k in ("AUTO_SHAPE", "SP", "GEOMETRY", "SHAPE"))
    if t1_is_shape and t2_is_shape:
        return True

    text_types = {"SP", "TEXT_BOX", "PLACEHOLDER", "AUTO_SHAPE", "1", "14", "17"}
    t1_is_text = any(k in u1 for k in text_types)
    t2_is_text = any(k in u2 for k in text_types)
    if t1_is_text and t2_is_text:
        return True

    return False


def find_matching_paragraph_index(
    orig_text: str,
    rem_paragraphs: List[str],
    used_indices: Optional[Set[int]] = None,
) -> Optional[int]:
    """
    Find matching paragraph index in rem_paragraphs for orig_text.
    Uses equation-safe normalization, exact match, substring containment, and fuzzy match.
    """
    if not orig_text or not rem_paragraphs:
        return None
    clean_orig = clean_text_alphanumeric(orig_text)
    if len(clean_orig) < 2:
        return None

    cleaned_paras = [clean_text_alphanumeric(p) for p in rem_paragraphs]

    # Pass 1: Exact match with an unused paragraph
    for idx, p_clean in enumerate(cleaned_paras):
        if used_indices and idx in used_indices:
            continue
        if p_clean == clean_orig:
            return idx

    # Pass 2: Substring containment among unused paragraphs
    for idx, p_clean in enumerate(cleaned_paras):
        if used_indices and idx in used_indices:
            continue
        if len(p_clean) >= 3 and (clean_orig in p_clean or p_clean in clean_orig):
            return idx

    # Pass 3: Fuzzy similarity >= 70% among unused paragraphs
    best_idx = None
    best_score = 0.0
    for idx, p_clean in enumerate(cleaned_paras):
        if used_indices and idx in used_indices:
            continue
        if len(p_clean) >= 2:
            sim = fuzz.ratio(clean_orig, p_clean)
            if sim >= 70.0 and sim > best_score:
                best_score = sim
                best_idx = idx
    if best_idx is not None:
        return best_idx

    # Pass 4: Fallback to exact match even if already used
    for idx, p_clean in enumerate(cleaned_paras):
        if p_clean == clean_orig:
            return idx

    return None


def compute_spatial_distance(s1: ShapeInfo, s2: ShapeInfo, slide_w: int = 9144000, slide_h: int = 5143500) -> float:
    """
    Compute normalized spatial distance between two shapes (0.0 = identical, 100.0+ = very distant).
    Default dimensions are standard 16:9 widescreen PPTX EMUs (9144000 x 5143500).
    """
    if slide_w <= 0:
        slide_w = 9144000
    if slide_h <= 0:
        slide_h = 5143500

    dx = abs(s1.left - s2.left) / slide_w
    dy = abs(s1.top - s2.top) / slide_h
    dw = abs(s1.width - s2.width) / slide_w
    dh = abs(s1.height - s2.height) / slide_h

    # Euclidean-style normalized penalty scaled to 0-100
    penalty = (math.sqrt(dx * dx + dy * dy) * 0.7 + (dw + dh) * 0.3) * 100.0
    return penalty


def match_slides(
    orig_slides: List[SlideData],
    rem_slides: List[SlideData],
) -> Tuple[Dict[int, int], List[int], List[int]]:
    """
    Match original slides to remediated slides.
    Returns:
    - matched: mapping orig_slide_index -> rem_slide_index
    - unmatched_original: list of orig_slide_index
    - unmatched_remediated: list of rem_slide_index
    """
    matched: Dict[int, int] = {}
    rem_indices = {s.slide_index: s for s in rem_slides}
    used_rem: Set[int] = set()

    # Rule 1: If counts match and presentations appear aligned, check 1-to-1 index matching
    if len(orig_slides) == len(rem_slides) and len(orig_slides) > 0:
        all_aligned = True
        for o_slide in orig_slides:
            r_slide = rem_indices.get(o_slide.slide_index)
            if not r_slide:
                all_aligned = False
                break
            if o_slide.title and r_slide.title:
                sim = fuzz.token_sort_ratio(
                    clean_text_alphanumeric(o_slide.title),
                    clean_text_alphanumeric(r_slide.title)
                )
                if sim < 40:
                    all_aligned = False
                    break
        if all_aligned:
            for o_slide in orig_slides:
                matched[o_slide.slide_index] = o_slide.slide_index
            return matched, [], []

    # Rule 2: Similarity matching based on title + text similarity
    for o_slide in orig_slides:
        best_rem_idx: Optional[int] = None
        best_score = -1.0

        o_norm_title = clean_text_alphanumeric(o_slide.title)
        o_norm_text = clean_text_alphanumeric(o_slide.all_text)

        # First check same index if available and unused
        same_idx_slide = rem_indices.get(o_slide.slide_index)
        if same_idx_slide and same_idx_slide.slide_index not in used_rem:
            r_norm_title = clean_text_alphanumeric(same_idx_slide.title)
            r_norm_text = clean_text_alphanumeric(same_idx_slide.all_text)
            title_sim = fuzz.token_sort_ratio(o_norm_title, r_norm_title) if (o_norm_title and r_norm_title) else 100
            text_sim = fuzz.token_sort_ratio(o_norm_text, r_norm_text) if (o_norm_text and r_norm_text) else 100
            if title_sim >= 70 and text_sim >= 60:
                best_rem_idx = same_idx_slide.slide_index
                best_score = title_sim * 0.6 + text_sim * 0.4

        if best_rem_idx is None:
            for r_slide in rem_slides:
                if r_slide.slide_index in used_rem:
                    continue

                r_norm_title = clean_text_alphanumeric(r_slide.title)
                r_norm_text = clean_text_alphanumeric(r_slide.all_text)

                score = 0.0
                if o_norm_title and r_norm_title:
                    score += 0.6 * fuzz.token_set_ratio(o_norm_title, r_norm_title)
                elif not o_norm_title and not r_norm_title:
                    score += 0.6 * 80.0

                if o_norm_text and r_norm_text:
                    score += 0.4 * fuzz.token_sort_ratio(o_norm_text, r_norm_text)
                elif not o_norm_text and not r_norm_text:
                    score += 0.4 * 80.0

                if score > best_score:
                    best_score = score
                    best_rem_idx = r_slide.slide_index

        if best_rem_idx is not None and best_score >= 40.0:
            matched[o_slide.slide_index] = best_rem_idx
            used_rem.add(best_rem_idx)

    unmatched_original = [s.slide_index for s in orig_slides if s.slide_index not in matched]
    unmatched_remediated = [s.slide_index for s in rem_slides if s.slide_index not in used_rem]

    return matched, unmatched_original, unmatched_remediated


class ShapeMatchResult(dict):
    """
    Mapping orig_shape_id -> rem_shape_id (or List[rem_shape_id] for group split).
    Inherits from dict for seamless backward compatibility.
    """
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.direct_matches: Dict[str, str] = {}
        self.group_splits: Dict[str, List[str]] = {}
        self.manual_reviews: List[ManualReviewItem] = []


def match_shapes(
    orig_shapes: Dict[str, ShapeInfo],
    rem_shapes: Dict[str, ShapeInfo],
    animated_orig_ids: Optional[Set[str]] = None,
    slide_index: int = 1,
    effect_info_map: Optional[Dict[str, Tuple[str, str]]] = None,
) -> ShapeMatchResult:
    """
    Match shapes according to the exact 5-step hierarchy:
    1. Same id and type, or text similarity >= 90%: direct match.
    2. If original is a GROUP and no remediated group matches:
       a. Find remediated shapes whose bbox lies >= 70% inside group bbox,
          or whose text is contained in group's combined text (or reverse).
       b. Map group to ALL of them (group split).
    3. If a remediated shape contains text of original shape (substring on normalized text), map to it.
    4. Pictures: image hash, then bounding box overlap.
    5. If still nothing: mark as "Manual review". Never drop silently.
    """
    result = ShapeMatchResult()
    used_rem_ids: Set[str] = set()
    used_rem_paragraphs: Dict[str, Set[int]] = {}

    # If animated_orig_ids is None, consider all original shapes
    targets_to_match = list(animated_orig_ids) if animated_orig_ids is not None else list(orig_shapes.keys())
    unmatched_orig_ids = list(targets_to_match)

    eff_map = effect_info_map or {}

    # -------------------------------------------------------------------------
    # STEP 0: Rebuilt slide check (Point 4)
    # If a slide was rebuilt (e.g. all content became one picture), do NOT put all
    # effects on that picture. List every original effect in Manual Review.
    # -------------------------------------------------------------------------
    rem_non_groups = [s for s in rem_shapes.values() if not s.is_group]
    rem_pics = [
        s for s in rem_non_groups
        if s.image_hash or "PIC" in s.shape_type.upper() or "picture" in s.shape_name.lower()
    ]
    orig_has_non_pics = any(
        not (orig_shapes[oid].image_hash or "PIC" in orig_shapes[oid].shape_type.upper() or "picture" in orig_shapes[oid].shape_name.lower())
        for oid in targets_to_match if oid in orig_shapes
    )
    is_rebuilt_as_picture = (
        len(rem_pics) == 1 and len(rem_non_groups) == 1 and orig_has_non_pics and len(targets_to_match) >= 1
    )

    if is_rebuilt_as_picture:
        for oid in targets_to_match:
            oshp = orig_shapes.get(oid)
            if not oshp:
                continue
            eff_name, trigger = eff_map.get(oid, ("Animation", "On Click"))
            snippet = oshp.text[:50] if oshp.text else (oshp.combined_text[:50] if oshp.combined_text else "No text")
            review_item = ManualReviewItem(
                slide_index=slide_index,
                shape_name=oshp.shape_name or f"Shape {oid}",
                shape_id=oid,
                text_snippet=snippet,
                shape_type=oshp.shape_type,
                bbox=oshp.bbox.to_str() if oshp.bbox else "Unknown",
                effect_name=eff_name,
                trigger=trigger,
                reason="Rebuilt slide (content converted to single picture)",
            )
            result.manual_reviews.append(review_item)
            result[oid] = []
        return result

    # -------------------------------------------------------------------------
    # STEP 1: Same id and compatible type, OR text similarity >= 90% (direct match)
    # -------------------------------------------------------------------------
    for oid in list(unmatched_orig_ids):
        oshp = orig_shapes.get(oid)
        if not oshp:
            continue

        # Sub-step 1a: Same id and compatible type
        if oid in rem_shapes and oid not in used_rem_ids:
            rshp = rem_shapes[oid]
            if are_types_compatible(oshp.shape_type, rshp.shape_type):
                result[oid] = oid
                result.direct_matches[oid] = oid
                used_rem_ids.add(oid)
                unmatched_orig_ids.remove(oid)
                continue

        # Sub-step 1b: Text similarity >= 90%
        clean_o = clean_text_alphanumeric(oshp.normalized_text or oshp.text)
        if len(clean_o) >= 3:
            best_rid: Optional[str] = None
            best_sim = -1.0
            for rid, rshp in rem_shapes.items():
                if rid in used_rem_ids:
                    continue
                if not are_types_compatible(oshp.shape_type, rshp.shape_type):
                    continue
                clean_r = clean_text_alphanumeric(rshp.normalized_text or rshp.text)
                if not clean_r:
                    continue
                sim = fuzz.ratio(clean_o, clean_r)
                if sim >= 90.0 and sim > best_sim:
                    best_sim = sim
                    best_rid = rid

            if best_rid is not None:
                result[oid] = best_rid
                result.direct_matches[oid] = best_rid
                used_rem_ids.add(best_rid)
                unmatched_orig_ids.remove(oid)
                continue

        # Sub-step 1c: Non-text shapes (groups without text, connectors, arrows, braces, tables, OLE)
        # Point 3: match by bounding box overlap and shape type, not by text
        clean_combined = clean_text_alphanumeric(oshp.combined_text or oshp.text)
        is_geom_or_table_or_ole = (
            len(clean_combined) < 3
            or "TABLE" in oshp.shape_type.upper()
            or "OLE" in oshp.shape_type.upper()
            or "LINE" in oshp.shape_type.upper()
            or any(k in oshp.shape_name.lower() for k in ("connector", "arrow", "brace", "bracket"))
        )
        if is_geom_or_table_or_ole and oshp.bbox:
            best_rid = None
            best_overlap = 0.0
            for rid, rshp in rem_shapes.items():
                if rid in used_rem_ids or (rshp.is_group and not oshp.is_group):
                    continue
                if not are_types_compatible(oshp.shape_type, rshp.shape_type):
                    continue
                if rshp.bbox:
                    overlap = oshp.bbox.overlap_ratio(rshp.bbox)
                    contains = oshp.bbox.contains_ratio(rshp.bbox)
                    eff_overlap = max(overlap, contains)
                    if eff_overlap >= 0.50 and eff_overlap > best_overlap:
                        best_overlap = eff_overlap
                        best_rid = rid

            if best_rid is not None:
                result[oid] = best_rid
                result.direct_matches[oid] = best_rid
                used_rem_ids.add(best_rid)
                unmatched_orig_ids.remove(oid)
                continue

        # Sub-step 1d: Multi-paragraph placeholder matching
        if len(clean_o) >= 2:
            best_rid: Optional[str] = None
            best_pidx: Optional[int] = None
            for rid, rshp in rem_shapes.items():
                if rshp.is_group:
                    continue
                if len(rshp.paragraphs) < 2:
                    continue
                used_p = used_rem_paragraphs.setdefault(rid, set())
                pidx = find_matching_paragraph_index(oshp.text, rshp.paragraphs, used_indices=used_p)
                if pidx is not None:
                    best_rid = rid
                    best_pidx = pidx
                    break
            if best_rid is not None and best_pidx is not None:
                result[oid] = best_rid
                result.direct_matches[oid] = best_rid
                used_rem_paragraphs[best_rid].add(best_pidx)
                if len(used_rem_paragraphs[best_rid]) >= len(rem_shapes[best_rid].paragraphs):
                    used_rem_ids.add(best_rid)
                unmatched_orig_ids.remove(oid)
                continue

    # -------------------------------------------------------------------------
    # STEP 2: Group Split Rule
    # If the original is a GROUP and no remediated group matches:
    # a. Find remediated top-level shapes (or group children) whose bounding box
    #    lies >= 70% inside the original group's bounding box, OR whose text
    #    is contained in the group's combined text (or the reverse).
    # b. Map the group to ALL of them.
    # -------------------------------------------------------------------------
    for oid in list(unmatched_orig_ids):
        oshp = orig_shapes.get(oid)
        if not oshp:
            continue

        if not oshp.is_group and "GROUP" not in oshp.shape_type.upper() and oshp.shape_type.lower() != "grpsp":
            continue

        # Check if already matched to a group
        if oid in result:
            continue

        matched_rids: List[str] = []
        clean_g = clean_text_alphanumeric(oshp.normalized_combined_text or oshp.normalized_text)

        # 2a: If oshp has descendants, match remediated shapes to descendants first
        if oshp.descendant_ids:
            for rid, rshp in rem_shapes.items():
                if rshp.is_group or rid in used_rem_ids:
                    continue
                clean_r = clean_text_alphanumeric(rshp.normalized_text or rshp.text)
                for d_id in oshp.descendant_ids:
                    d_shp = orig_shapes.get(d_id)
                    if not d_shp:
                        continue
                    # Match by id
                    if rid == d_id and are_types_compatible(d_shp.shape_type, rshp.shape_type):
                        matched_rids.append(rid)
                        break
                    # Match by text
                    clean_d = clean_text_alphanumeric(d_shp.normalized_text or d_shp.text)
                    if len(clean_d) >= 3 and len(clean_r) >= 3 and (clean_d in clean_r or clean_r in clean_d):
                        matched_rids.append(rid)
                        break
                    # Match by child bbox
                    if d_shp.bbox and rshp.bbox:
                        if max(d_shp.bbox.overlap_ratio(rshp.bbox), d_shp.bbox.contains_ratio(rshp.bbox)) >= 0.50:
                            if are_types_compatible(d_shp.shape_type, rshp.shape_type):
                                matched_rids.append(rid)
                                break

        # 2b: Fallback if no descendant matching was found
        if not matched_rids:
            for rid, rshp in rem_shapes.items():
                # Don't match whole groups as children
                if rshp.is_group or rid in used_rem_ids:
                    continue

                # Bounding box test: inside group bbox >= 70%
                bbox_ok = False
                if oshp.bbox and rshp.bbox:
                    ratio = oshp.bbox.contains_ratio(rshp.bbox)
                    if ratio >= 0.70:
                        bbox_ok = True

                # Text test: text contained in group's combined text (or reverse)
                text_ok = False
                clean_r = clean_text_alphanumeric(rshp.normalized_text)
                if len(clean_r) >= 3 and len(clean_g) >= 3:
                    if (clean_r in clean_g) or (clean_g in clean_r):
                        text_ok = True

                if bbox_ok or text_ok:
                    matched_rids.append(rid)

        if matched_rids:
            if len(matched_rids) == 1:
                result[oid] = matched_rids[0]
                result.direct_matches[oid] = matched_rids[0]
            else:
                result[oid] = matched_rids
                result.group_splits[oid] = matched_rids
            used_rem_ids.update(matched_rids)
            unmatched_orig_ids.remove(oid)

    # -------------------------------------------------------------------------
    # STEP 3: Substring text containment on normalized text
    # -------------------------------------------------------------------------
    for oid in list(unmatched_orig_ids):
        oshp = orig_shapes.get(oid)
        if not oshp:
            continue

        clean_o = clean_text_alphanumeric(oshp.normalized_text or oshp.text)
        if len(clean_o) >= 3:
            best_rid: Optional[str] = None
            best_score = -1.0
            for rid, rshp in rem_shapes.items():
                if rshp.is_group:
                    continue
                # If rid is already used, only allow if it has multiple paragraphs with unused paragraphs remaining
                if rid in used_rem_ids:
                    if len(rshp.paragraphs) < 2 or len(used_rem_paragraphs.get(rid, set())) >= len(rshp.paragraphs):
                        continue
                clean_r = clean_text_alphanumeric(rshp.normalized_text or rshp.text)
                if len(clean_r) >= 3 and (clean_o in clean_r or clean_r in clean_o):
                    # Prefer closest position
                    dist = compute_spatial_distance(oshp, rshp)
                    score = 100.0 - dist
                    if score > best_score:
                        best_score = score
                        best_rid = rid

            if best_rid is not None:
                result[oid] = best_rid
                result.direct_matches[oid] = best_rid
                rshp = rem_shapes[best_rid]
                if len(rshp.paragraphs) >= 2:
                    pidx = find_matching_paragraph_index(oshp.text, rshp.paragraphs, used_indices=used_rem_paragraphs.get(best_rid, set()))
                    if pidx is not None:
                        used_rem_paragraphs.setdefault(best_rid, set()).add(pidx)
                    if len(used_rem_paragraphs.get(best_rid, set())) >= len(rshp.paragraphs):
                        used_rem_ids.add(best_rid)
                else:
                    used_rem_ids.add(best_rid)
                unmatched_orig_ids.remove(oid)

    # -------------------------------------------------------------------------
    # STEP 4: Pictures: image hash, then bounding box overlap
    # -------------------------------------------------------------------------
    for oid in list(unmatched_orig_ids):
        oshp = orig_shapes.get(oid)
        if not oshp:
            continue

        is_pic = oshp.image_hash or "PIC" in oshp.shape_type.upper() or "picture" in oshp.shape_name.lower()
        if not is_pic:
            continue

        best_rid: Optional[str] = None

        # 4a: image hash
        if oshp.image_hash:
            for rid, rshp in rem_shapes.items():
                if rid in used_rem_ids:
                    continue
                if rshp.image_hash and rshp.image_hash == oshp.image_hash:
                    best_rid = rid
                    break

        # 4b: bounding box overlap >= 50%
        if best_rid is None and oshp.bbox:
            best_overlap = 0.0
            for rid, rshp in rem_shapes.items():
                if rid in used_rem_ids or rshp.is_group:
                    continue
                if rshp.bbox:
                    overlap = oshp.bbox.overlap_ratio(rshp.bbox)
                    if overlap >= 0.50 and overlap > best_overlap:
                        best_overlap = overlap
                        best_rid = rid

        if best_rid is not None:
            result[oid] = best_rid
            result.direct_matches[oid] = best_rid
            used_rem_ids.add(best_rid)
            unmatched_orig_ids.remove(oid)

    # -------------------------------------------------------------------------
    # STEP 5: If still nothing: mark as "Manual review". Never drop silently.
    # -------------------------------------------------------------------------
    for oid in unmatched_orig_ids:
        oshp = orig_shapes.get(oid)
        if not oshp:
            continue

        eff_name, trigger = eff_map.get(oid, ("Animation", "On Click"))
        snippet = oshp.text[:50] if oshp.text else (oshp.combined_text[:50] if oshp.combined_text else "No text")

        review_item = ManualReviewItem(
            slide_index=slide_index,
            shape_name=oshp.shape_name or f"Shape {oid}",
            shape_id=oid,
            text_snippet=snippet,
            shape_type=oshp.shape_type,
            bbox=oshp.bbox.to_str() if oshp.bbox else "Unknown",
            effect_name=eff_name,
            trigger=trigger,
            reason="No matching remediated shape found (shape deleted or structure changed)",
        )
        result.manual_reviews.append(review_item)
        result[oid] = []

    # Map remaining non-animated shapes for general structure / build lists
    for oid, oshp in orig_shapes.items():
        if oid in result:
            continue
        # Check by name and position
        best_rid = None
        for rid, rshp in rem_shapes.items():
            if rid in used_rem_ids:
                continue
            if oshp.shape_name and oshp.shape_name == rshp.shape_name and are_types_compatible(oshp.shape_type, rshp.shape_type):
                best_rid = rid
                break
        if best_rid:
            result[oid] = best_rid
            used_rem_ids.add(best_rid)

    return result
