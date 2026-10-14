from __future__ import annotations

import io
from typing import Any, List, Optional, Set, Tuple

from pptx import Presentation
from lxml import etree

from core.models import (
    SkippedAnimation,
    SlideData,
    SlideStatus,
    SlideVerificationResult,
)
from core.parser import NAMESPACES, extract_all_shapes_from_slide
from core.animation_transfer import extract_effects_from_timing, get_tag_name


class ValidationError(Exception):
    """Raised when presentation integrity validation fails."""
    pass


def validate_presentation_integrity(pptx_bytes: bytes) -> Tuple[bool, List[str], List[str]]:
    """
    Validate the output PPTX presentation:
    1. Reopen with python-pptx
    2. Check XML well-formedness
    3. Check unique cTn ids on each slide (hard-fail)
    4. Check every animation spid points to an existing shape on that slide (hard-fail)
    5. FAIL if any effect <p:cTn> has an empty <p:childTnLst/>
    6. Check build list entries:
       - If spid points to non-existent shape: hard-fail
       - If shape has no matching animation: auto-remove with a warning
    Returns (is_valid, list_of_hard_issues, list_of_warnings).
    """
    hard_issues: List[str] = []
    warnings: List[str] = []

    try:
        prs = Presentation(io.BytesIO(pptx_bytes))
    except Exception as e:
        return False, [f"Fatal: python-pptx cannot open generated presentation: {e}"], []

    for slide_idx, slide in enumerate(prs.slides, start=1):
        try:
            sld_xml = slide._element
            etree.tostring(sld_xml)
        except Exception as e:
            hard_issues.append(f"Slide {slide_idx}: XML is not well-formed: {e}")
            continue

        timing = sld_xml.find(f"{{{NAMESPACES['p']}}}timing")
        if timing is None:
            continue

        existing_shapes = extract_all_shapes_from_slide(slide)
        existing_ids = set(existing_shapes.keys())

        # 1. Unique cTn IDs
        ctn_ids: List[str] = []
        for elem in timing.iter():
            if isinstance(elem.tag, str) and elem.tag.endswith("cTn"):
                cid = elem.attrib.get("id")
                if cid:
                    ctn_ids.append(cid)

        if len(ctn_ids) != len(set(ctn_ids)):
            duplicates = [cid for cid in set(ctn_ids) if ctn_ids.count(cid) > 1]
            hard_issues.append(f"Slide {slide_idx}: Non-unique cTn id(s) detected: {duplicates}")

        # 2. FAIL if any effect <p:cTn> has an empty <p:childTnLst/>
        for elem in timing.iter():
            if get_tag_name(elem) == "cTn":
                preset_id = elem.attrib.get("presetID")
                preset_class = elem.attrib.get("presetClass")
                node_type = elem.attrib.get("nodeType", "")
                if preset_id or preset_class or node_type in ("clickEffect", "withEffect", "afterEffect"):
                    child_lst = elem.find(f"{{{NAMESPACES['p']}}}childTnLst")
                    if child_lst is not None and len(child_lst) == 0:
                        hard_issues.append(
                            f"Slide {slide_idx}: Effect <p:cTn id='{elem.attrib.get('id')}'> has empty <p:childTnLst/>."
                        )

        # 3. Every spTgt points to an existing shape
        animated_spids: Set[str] = set()
        for elem in timing.iter():
            if isinstance(elem.tag, str) and elem.tag.endswith("spTgt"):
                spid = elem.attrib.get("spid")
                if spid:
                    animated_spids.add(str(spid))
                    if str(spid) not in existing_ids:
                        hard_issues.append(
                            f"Slide {slide_idx}: Animation references non-existent shape ID '{spid}'."
                        )

        # 4. Check build list entries
        bldLst = timing.find(f".//{{{NAMESPACES['p']}}}bldLst")
        if bldLst is not None:
            for bld_child in list(bldLst):
                bld_spid = bld_child.attrib.get("spid")
                if bld_spid:
                    if str(bld_spid) not in existing_ids:
                        hard_issues.append(
                            f"Slide {slide_idx}: Build list references non-existent shape ID '{bld_spid}'."
                        )
                    elif str(bld_spid) not in animated_spids:
                        bldLst.remove(bld_child)
                        warnings.append(
                            f"Slide {slide_idx}: Warning: Auto-removed build list entry '{bld_spid}' with no matching animation."
                        )

            if len(bldLst) == 0:
                p = bldLst.getparent()
                if p is not None:
                    p.remove(bldLst)

    is_valid = len(hard_issues) == 0
    return is_valid, hard_issues, warnings


from core.matcher import clean_text_alphanumeric
from rapidfuzz import fuzz


def verify_slide_animations(
    orig_slide_data: SlideData,
    rem_slide_data: SlideData,
    shape_map: Any,
    skipped_animations: List[SkippedAnimation],
    manual_reviews: Optional[List[Any]] = None,
    mapped_paragraph_level: int = 0,
) -> SlideVerificationResult:
    """
    Verification by content (Point 5 & 6):
    - Verification must use text and geometry, not shape names or ids.
    - An original effect is 'covered' if an animated remediated shape contains
      the same normalized text, or (no text) overlaps the same bounding box.
    - Click step counts are compared.
    - Strict Pass rule: A slide is Pass ONLY IF click steps are equal and
      nothing is in Manual review.
    - Reports covered / not covered per slide.
    """
    orig_xml = orig_slide_data.xml_root
    rem_xml = rem_slide_data.xml_root

    orig_timing = orig_xml.find(f"{{{NAMESPACES['p']}}}timing") if orig_xml is not None else None
    rem_timing = rem_xml.find(f"{{{NAMESPACES['p']}}}timing") if rem_xml is not None else None

    orig_click_steps, orig_effects = extract_effects_from_timing(orig_timing, orig_slide_data.shapes)
    rem_click_steps, rem_effects = extract_effects_from_timing(rem_timing, rem_slide_data.shapes)

    issues: List[str] = []
    used_rem_indices: Set[int] = set()

    # 1. Check for hollow effects (empty childTnLst)
    if rem_timing is not None:
        for elem in rem_timing.iter():
            if get_tag_name(elem) == "cTn":
                preset_id = elem.attrib.get("presetID")
                preset_class = elem.attrib.get("presetClass")
                node_type = elem.attrib.get("nodeType", "")
                if preset_id or preset_class or node_type in ("clickEffect", "withEffect", "afterEffect"):
                    c_lst = elem.find(f"{{{NAMESPACES['p']}}}childTnLst")
                    if c_lst is not None and len(c_lst) == 0:
                        issues.append(f"Hollow effect found: <p:cTn id='{elem.attrib.get('id')}'> has empty childTnLst.")

    # 2. Content-based verification using text and geometry (NOT shape names or IDs)
    covered_count = 0
    not_covered_count = 0
    exact_matches_count = 0

    manual_items = list(manual_reviews) if manual_reviews is not None else []
    manual_review_count = len(manual_items)

    mapped_direct = 0
    mapped_split = 0
    if hasattr(shape_map, "direct_matches"):
        mapped_direct = len([k for k, v in shape_map.direct_matches.items() if v])
    if hasattr(shape_map, "group_splits"):
        mapped_split = len(shape_map.group_splits)

    for o_eff in orig_effects:
        o_sinfo = orig_slide_data.shapes.get(str(o_eff.target_spid))
        is_covered = False

        o_norm = ""
        if o_sinfo:
            o_norm = clean_text_alphanumeric(
                o_sinfo.normalized_text or o_sinfo.text or o_sinfo.combined_text
            )

        for r_idx, r_eff in enumerate(rem_effects):
            r_sinfo = rem_slide_data.shapes.get(str(r_eff.target_spid))
            if not r_sinfo:
                continue

            # A. Text matching (equation-safe, substring containment, fuzzy)
            if len(o_norm) >= 2:
                r_norm = clean_text_alphanumeric(r_sinfo.normalized_text or r_sinfo.text)
                # Check paragraphs if available
                para_match = False
                if r_sinfo.paragraphs:
                    for p in r_sinfo.paragraphs:
                        clean_p = clean_text_alphanumeric(p)
                        if len(clean_p) >= 2 and (clean_p == o_norm or o_norm in clean_p or clean_p in o_norm):
                            para_match = True
                            break

                if para_match or (len(r_norm) >= 2 and (o_norm == r_norm or o_norm in r_norm or r_norm in o_norm or fuzz.ratio(o_norm, r_norm) >= 80.0)):
                    is_covered = True
                    used_rem_indices.add(r_idx)
                    exact_matches_count += 1
                    break

            # B. Pictures: image hash match
            if o_sinfo and o_sinfo.image_hash and r_sinfo.image_hash:
                if o_sinfo.image_hash == r_sinfo.image_hash:
                    is_covered = True
                    used_rem_indices.add(r_idx)
                    exact_matches_count += 1
                    break

            # C. Geometry matching (bounding box overlap >= 50% or contains ratio >= 50%)
            if o_sinfo and o_sinfo.bbox and r_sinfo and r_sinfo.bbox:
                overlap = o_sinfo.bbox.overlap_ratio(r_sinfo.bbox)
                contains = o_sinfo.bbox.contains_ratio(r_sinfo.bbox)
                if max(overlap, contains) >= 0.50:
                    is_covered = True
                    used_rem_indices.add(r_idx)
                    exact_matches_count += 1
                    break

        if is_covered:
            covered_count += 1
        else:
            not_covered_count += 1

    # 3. Click steps comparison:
    if orig_click_steps != rem_click_steps and len(orig_effects) > 0:
        if len(skipped_animations) == 0 and manual_review_count == 0:
            issues.append(
                f"Click steps count mismatch: original has {orig_click_steps}, but output has {rem_click_steps}."
            )

    # 4. Strict status determination (Point 5):
    # A slide is Pass ONLY IF click steps are equal and nothing is in Manual review.
    has_hard_fail = any("Hollow" in iss or "non-existent" in iss for iss in issues)

    if has_hard_fail:
        status = SlideStatus.FAILED
        passed = False
    elif manual_review_count > 0:
        status = SlideStatus.MANUAL_REVIEW
        passed = False
    elif not_covered_count > 0:
        status = SlideStatus.MANUAL_REVIEW
        passed = False
    elif orig_click_steps != rem_click_steps and len(orig_effects) > 0:
        status = SlideStatus.MANUAL_REVIEW
        passed = False
    elif len(skipped_animations) > 0:
        status = SlideStatus.PARTIAL
        passed = False
    else:
        status = SlideStatus.PASS
        passed = True

    return SlideVerificationResult(
        slide_index=orig_slide_data.slide_index,
        remediated_slide_index=rem_slide_data.slide_index,
        passed=passed,
        status=status,
        original_effects_count=len(orig_effects),
        output_effects_count=len(rem_effects),
        exact_matches_count=exact_matches_count,
        skipped_count=len(skipped_animations),
        click_steps_orig=orig_click_steps,
        click_steps_output=rem_click_steps,
        covered_count=covered_count,
        not_covered_count=not_covered_count,
        mapped_direct_count=mapped_direct,
        mapped_group_split_count=mapped_split,
        mapped_paragraph_level_count=mapped_paragraph_level,
        manual_review_count=manual_review_count,
        issues=issues,
    )
