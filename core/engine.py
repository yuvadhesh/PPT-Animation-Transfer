from __future__ import annotations

import copy
import io
import re
from pathlib import Path
from typing import Callable, List, Optional, Union

from pptx import Presentation

from core.animation_transfer import (
    extract_effects_from_timing,
    transfer_slide_animations_and_transition,
)
from core.matcher import match_shapes, match_slides
from core.models import (
    DeckTransferResult,
    ManualReviewItem,
    SlideStatus,
    SlideTransferResult,
)
from core.parser import (
    NAMESPACES,
    PresentationParseError,
    extract_slide_data,
    validate_pptx_file,
)
from core.reporter import export_excel_report
from core.validator import validate_presentation_integrity, verify_slide_animations


def transfer_presentation_animations(
    original_source: Union[str, Path, bytes, io.BytesIO],
    remediated_source: Union[str, Path, bytes, io.BytesIO],
    copy_transitions: bool = True,
    progress_callback: Optional[Callable[[float, str], None]] = None,
    original_filename: Optional[str] = None,
    remediated_filename: Optional[str] = None,
) -> DeckTransferResult:
    """
    Main orchestration entry point:
    1. Validates both presentations with user-friendly error checking.
    2. Deep clones slides in-memory (never touches original files).
    3. Matches slides by index / text similarity.
    4. Cascades shape matching on every slide.
    5. Transfers and remaps all animation nodes and optional transitions.
    6. Prunes orphaned animations and cleans empty click containers.
    7. Copies media relationships where applicable.
    8. Validates output integrity with python-pptx and XML schema checks.
    9. Verifies animations shape-by-shape (Pass/Fail).
    10. Generates Excel report bytes.
    """
    if progress_callback:
        progress_callback(0.05, "Validating input presentations...")

    # Step 1: Validate input files
    validate_pptx_file(original_source, file_label="Original Presentation", filename=original_filename)
    validate_pptx_file(remediated_source, file_label="Remediated Presentation", filename=remediated_filename)

    # Read bytes into memory
    def get_bytes(src: Union[str, Path, bytes, io.BytesIO]) -> bytes:
        if isinstance(src, (str, Path)):
            with open(src, "rb") as f:
                return f.read()
        elif isinstance(src, io.BytesIO):
            return src.getvalue()
        elif isinstance(src, bytes):
            return src
        raise PresentationParseError("Invalid presentation source format.")

    orig_bytes = get_bytes(original_source)
    rem_bytes = get_bytes(remediated_source)

    if progress_callback:
        progress_callback(0.15, "Parsing slide contents and shapes...")

    # Step 2: Open with python-pptx
    orig_prs = Presentation(io.BytesIO(orig_bytes))
    rem_prs = Presentation(io.BytesIO(rem_bytes))

    orig_slide_datas = [
        extract_slide_data(slide, idx)
        for idx, slide in enumerate(orig_prs.slides, start=1)
    ]
    rem_slide_datas = [
        extract_slide_data(slide, idx)
        for idx, slide in enumerate(rem_prs.slides, start=1)
    ]

    rem_slides_by_idx = {s.slide_index: s for s in rem_slide_datas}

    if progress_callback:
        progress_callback(0.25, "Matching slides between presentations...")

    # Step 3: Match slides
    matched_slides, unmatched_orig, unmatched_rem = match_slides(
        orig_slide_datas, rem_slide_datas
    )

    total_matches = len(matched_slides)
    slide_results: List[SlideTransferResult] = []

    slides_with_animations = 0
    total_copied = 0
    total_skipped = 0

    # Step 4: Process each matched slide
    for i, (orig_idx, rem_idx) in enumerate(matched_slides.items(), start=1):
        if progress_callback:
            pct = 0.25 + 0.50 * (i / max(total_matches, 1))
            progress_callback(pct, f"Transferring animations for Slide {orig_idx} -> Slide {rem_idx}...")

        orig_s = orig_slide_datas[orig_idx - 1]
        rem_s = rem_slides_by_idx[rem_idx]

        # Save pre-transfer state of remediated slide timing so we can revert if failed
        orig_rem_timing_elem = rem_s.xml_root.find(f"{{{NAMESPACES['p']}}}timing")
        orig_rem_timing_backup = copy.deepcopy(orig_rem_timing_elem) if orig_rem_timing_elem is not None else None

        # Extract original effects to guide matching
        orig_timing = orig_s.xml_root.find(f"{{{NAMESPACES['p']}}}timing")
        _, orig_effects = extract_effects_from_timing(orig_timing, orig_s.shapes)
        animated_orig_ids = {eff.target_spid for eff in orig_effects if eff.target_spid}
        eff_info_map = {
            eff.target_spid: (
                eff.effect_name,
                eff.trigger.value if hasattr(eff.trigger, "value") else str(eff.trigger),
            )
            for eff in orig_effects
        }

        # Cascading shape mapping
        shape_map = match_shapes(
            orig_s.shapes,
            rem_s.shapes,
            animated_orig_ids=animated_orig_ids,
            slide_index=orig_idx,
            effect_info_map=eff_info_map,
        )

        try:
            # Transfer animations and transitions
            s_res = transfer_slide_animations_and_transition(
                orig_slide_data=orig_s,
                rem_slide_data=rem_s,
                shape_map=shape_map,
                copy_transitions=copy_transitions,
            )

            if s_res.original_animation_count > 0:
                slides_with_animations += 1

            total_copied += s_res.copied_animation_count
            total_skipped += s_res.skipped_animation_count

            # Run verification step
            verif = verify_slide_animations(
                orig_slide_data=orig_s,
                rem_slide_data=rem_s,
                shape_map=shape_map,
                skipped_animations=s_res.skipped,
                manual_reviews=s_res.manual_review_items,
                mapped_paragraph_level=s_res.mapped_paragraph_level_count,
            )
            s_res.verification = verif
            s_res.covered_count = verif.covered_count
            s_res.not_covered_count = verif.not_covered_count
            s_res.mapped_direct_count = verif.mapped_direct_count
            s_res.mapped_group_split_count = verif.mapped_group_split_count
            s_res.mapped_paragraph_level_count = verif.mapped_paragraph_level_count
            s_res.manual_review_count = verif.manual_review_count
            s_res.status = verif.status

            if not verif.passed and verif.status == SlideStatus.FAILED:
                s_res.error_message = "; ".join(verif.issues)
                # Revert timing on rem_slide so it is left unchanged
                cur_timing = rem_s.xml_root.find(f"{{{NAMESPACES['p']}}}timing")
                if cur_timing is not None:
                    rem_s.xml_root.remove(cur_timing)
                if orig_rem_timing_backup is not None:
                    rem_s.xml_root.append(orig_rem_timing_backup)
        except Exception as e:
            # Revert timing on failure
            cur_timing = rem_s.xml_root.find(f"{{{NAMESPACES['p']}}}timing")
            if cur_timing is not None:
                rem_s.xml_root.remove(cur_timing)
            if orig_rem_timing_backup is not None:
                rem_s.xml_root.append(orig_rem_timing_backup)

            s_res = SlideTransferResult(
                slide_index=orig_idx,
                remediated_slide_index=rem_idx,
                status=SlideStatus.FAILED,
                original_animation_count=0,
                copied_animation_count=0,
                skipped_animation_count=0,
                error_message=f"Slide transfer failed: {e}",
            )

        slide_results.append(s_res)

    # Process unmatched original slides
    for u_orig_idx in unmatched_orig:
        orig_s = orig_slide_datas[u_orig_idx - 1]
        slide_results.append(
            SlideTransferResult(
                slide_index=u_orig_idx,
                remediated_slide_index=None,
                status=SlideStatus.UNMATCHED,
                original_animation_count=len(orig_s.shapes),
                copied_animation_count=0,
                skipped_animation_count=0,
                error_message="Slide could not be matched to any remediated slide.",
            )
        )

    # Sort results by slide index
    slide_results.sort(key=lambda x: x.slide_index)

    if progress_callback:
        progress_callback(0.80, "Serializing output presentation...")

    # Step 5: Save remediated presentation
    out_bio = io.BytesIO()
    rem_prs.save(out_bio)
    output_pptx_bytes = out_bio.getvalue()

    if progress_callback:
        progress_callback(0.90, "Validating output presentation integrity...")

    # Step 6: Validate output presentation integrity
    is_valid, hard_issues, val_warnings = validate_presentation_integrity(output_pptx_bytes)

    # Distribute warnings to matching slide results
    for warn in val_warnings:
        m = re.search(r"Slide (\d+):", warn)
        if m:
            s_num = int(m.group(1))
            for s_res in slide_results:
                if s_res.remediated_slide_index == s_num or s_res.slide_index == s_num:
                    if s_res.verification:
                        s_res.verification.issues.append(warn)

    if not is_valid:
        # Check which slides caused the hard failure
        corrupted_rem_slides = set()
        for issue in hard_issues:
            m = re.search(r"Slide (\d+):", issue)
            if m:
                corrupted_rem_slides.add(int(m.group(1)))

        if corrupted_rem_slides:
            # Omit corrupted timing for those specific slides only, keep other slides
            for rem_idx in corrupted_rem_slides:
                if 1 <= rem_idx <= len(rem_prs.slides):
                    corrupted_slide = rem_prs.slides[rem_idx - 1]
                    bad_timing = corrupted_slide._element.find(f"{{{NAMESPACES['p']}}}timing")
                    if bad_timing is not None:
                        corrupted_slide._element.remove(bad_timing)
                # Mark slide as FAILED in report
                for s_res in slide_results:
                    if s_res.remediated_slide_index == rem_idx:
                        s_res.status = SlideStatus.FAILED
                        matching_issues = [iss for iss in hard_issues if f"Slide {rem_idx}:" in iss]
                        s_res.error_message = "; ".join(matching_issues)
                        if s_res.verification:
                            s_res.verification.passed = False
                            s_res.verification.issues.extend(matching_issues)

            # Re-serialize cleaned presentation
            out_bio = io.BytesIO()
            rem_prs.save(out_bio)
            output_pptx_bytes = out_bio.getvalue()
        else:
            # Presentation-level fatal error
            error_details = "\n".join(hard_issues)
            raise PresentationParseError(
                f"Generated presentation failed integrity validation:\n{error_details}"
            )

    # Step 7: Create final result
    all_manual_items: List[ManualReviewItem] = []
    for s in slide_results:
        all_manual_items.extend(s.manual_review_items)

    total_exact_matches = sum(
        s.verification.exact_matches_count
        for s in slide_results
        if s.verification is not None
    )
    total_failed_slides = sum(
        1 for s in slide_results
        if s.status == SlideStatus.FAILED or (s.verification is not None and s.verification.status == SlideStatus.FAILED)
    )
    total_manual_review = len(all_manual_items)
    total_mapped_direct = sum(s.mapped_direct_count for s in slide_results)
    total_mapped_group_split = sum(s.mapped_group_split_count for s in slide_results)
    total_mapped_paragraph_level = sum(s.mapped_paragraph_level_count for s in slide_results)
    total_covered_effects = sum(s.covered_count for s in slide_results)
    total_not_covered_effects = sum(s.not_covered_count for s in slide_results)

    is_100_percent_faithful = (
        total_failed_slides == 0
        and total_manual_review == 0
        and total_not_covered_effects == 0
        and len(unmatched_orig) == 0
        and len(unmatched_rem) == 0
    )

    deck_result = DeckTransferResult(
        success=True,
        message="Animation transfer completed successfully.",
        original_slide_count=len(orig_slide_datas),
        remediated_slide_count=len(rem_slide_datas),
        slides_with_animations=slides_with_animations,
        total_animations_copied=total_copied,
        total_animations_skipped=total_skipped,
        total_exact_matches=total_exact_matches,
        total_failed_slides=total_failed_slides,
        total_mapped_direct=total_mapped_direct,
        total_mapped_group_split=total_mapped_group_split,
        total_mapped_paragraph_level=total_mapped_paragraph_level,
        total_manual_review=total_manual_review,
        total_covered_effects=total_covered_effects,
        total_not_covered_effects=total_not_covered_effects,
        is_100_percent_faithful=is_100_percent_faithful,
        unmatched_slides_count=len(unmatched_orig) + len(unmatched_rem),
        slide_results=slide_results,
        manual_review_items=all_manual_items,
        unmatched_original_slides=unmatched_orig,
        unmatched_remediated_slides=unmatched_rem,
        output_pptx_bytes=output_pptx_bytes,
    )

    if progress_callback:
        progress_callback(0.95, "Generating Excel report...")

    # Step 8: Generate Excel report
    excel_bytes = export_excel_report(deck_result)
    deck_result.excel_report_bytes = excel_bytes

    if progress_callback:
        progress_callback(1.0, "Done!")

    return deck_result
