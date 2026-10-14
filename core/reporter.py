from __future__ import annotations

import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from core.models import DeckTransferResult


def export_excel_report(result: DeckTransferResult) -> bytes:
    """
    Generate an Excel report with comprehensive verification and transfer details:
    1. Summary & Verification: deck stats, faithfulness indicator, and per-slide verification table
    2. Verification: standalone verification table for automated audit
    3. Animations: detailed properties of all copied animations
    4. Skipped: detailed reasons for any skipped animations
    """
    wb = openpyxl.Workbook()

    # Style definitions
    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    card_header_fill = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    card_header_font = Font(name="Calibri", size=10, bold=True, color="E2E8F0")

    title_font = Font(name="Calibri", size=16, bold=True, color="0F172A")

    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    pass_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    pass_font = Font(name="Calibri", size=10, bold=True, color="166534")

    fail_fill = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    fail_font = Font(name="Calibri", size=10, bold=True, color="991B1B")

    warn_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    warn_font = Font(name="Calibri", size=10, bold=True, color="92400E")

    manual_fill = PatternFill(start_color="FFEDD5", end_color="FFEDD5", fill_type="solid")
    manual_font = Font(name="Calibri", size=10, bold=True, color="C2410C")

    # -------------------------------------------------------------------------
    # SHEET 1: Summary & Verification
    # -------------------------------------------------------------------------
    ws_summary = wb.active
    ws_summary.title = "Summary"
    ws_summary.views.sheetView[0].showGridLines = True

    ws_summary["A1"] = "PPT Animation Transfer Report"
    ws_summary["A1"].font = title_font
    ws_summary["A2"] = (
        "100% Faithful Transfer: " + ("YES (All effects matched or reported)" if result.is_100_percent_faithful else "NO (Discrepancies / Manual Review found)")
    )
    ws_summary["A2"].font = Font(
        name="Calibri",
        size=11,
        bold=True,
        color="166534" if result.is_100_percent_faithful else "991B1B",
    )

    # KPI summary table
    total_orig_effs = sum(
        s.verification.original_effects_count
        for s in result.slide_results
        if s.verification is not None
    )
    total_out_effs = sum(
        s.verification.output_effects_count
        for s in result.slide_results
        if s.verification is not None
    )

    kpis = [
        ("Original Slides", result.original_slide_count),
        ("Remediated Slides", result.remediated_slide_count),
        ("Original Effects", total_orig_effs),
        ("Output Effects", total_out_effs),
        ("Direct Mapped", result.total_mapped_direct),
        ("Group Split", result.total_mapped_group_split),
        ("Paragraph Mapped", result.total_mapped_paragraph_level),
        ("Manual Review", result.total_manual_review),
        ("Covered", result.total_covered_effects),
        ("Not Covered", result.total_not_covered_effects),
        ("100% Faithful", "YES" if result.is_100_percent_faithful else "NO"),
    ]

    for col_idx, (kpi_name, kpi_val) in enumerate(kpis, start=1):
        cell_h = ws_summary.cell(row=4, column=col_idx, value=kpi_name)
        cell_h.fill = card_header_fill
        cell_h.font = card_header_font
        cell_h.alignment = Alignment(horizontal="center", vertical="center")
        cell_h.border = thin_border

        cell_v = ws_summary.cell(row=5, column=col_idx, value=kpi_val)
        cell_v.font = Font(name="Calibri", size=14, bold=True, color="0F172A")
        cell_v.alignment = Alignment(horizontal="center", vertical="center")
        cell_v.border = thin_border
        if kpi_name == "100% Faithful":
            cell_v.fill = pass_fill if result.is_100_percent_faithful else fail_fill
            cell_v.font = pass_font if result.is_100_percent_faithful else fail_font

    # Per-slide Verification table (Point 5)
    verif_headers = [
        "Slide No",
        "Remediated Slide No",
        "Original Effects",
        "Output Effects",
        "Click Steps (Orig)",
        "Click Steps (Output)",
        "Mapped (Direct)",
        "Mapped (Group Split)",
        "Mapped (Paragraph)",
        "Manual Review",
        "Covered",
        "Not Covered",
        "Status",
        "Issues / Notes",
    ]

    start_row = 7
    for col_idx, header in enumerate(verif_headers, start=1):
        cell = ws_summary.cell(row=start_row, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    current_row = start_row + 1
    for s_res in result.slide_results:
        v = s_res.verification
        orig_effs = v.original_effects_count if v else s_res.original_animation_count
        out_effs = v.output_effects_count if v else s_res.copied_animation_count
        direct_cnt = v.mapped_direct_count if v else s_res.mapped_direct_count
        split_cnt = v.mapped_group_split_count if v else s_res.mapped_group_split_count
        para_cnt = v.mapped_paragraph_level_count if v else s_res.mapped_paragraph_level_count
        manual_cnt = v.manual_review_count if v else s_res.manual_review_count
        covered_cnt = v.covered_count if v else s_res.covered_count
        not_covered_cnt = v.not_covered_count if v else s_res.not_covered_count
        orig_clicks = v.click_steps_orig if v else 0
        out_clicks = v.click_steps_output if v else 0
        issues_str = "; ".join(v.issues) if (v and v.issues) else (s_res.error_message or "None")

        ws_summary.cell(row=current_row, column=1, value=s_res.slide_index).alignment = Alignment(horizontal="center")
        ws_summary.cell(
            row=current_row,
            column=2,
            value=s_res.remediated_slide_index if s_res.remediated_slide_index else "N/A",
        ).alignment = Alignment(horizontal="center")
        ws_summary.cell(row=current_row, column=3, value=orig_effs).alignment = Alignment(horizontal="center")
        ws_summary.cell(row=current_row, column=4, value=out_effs).alignment = Alignment(horizontal="center")
        ws_summary.cell(row=current_row, column=5, value=orig_clicks).alignment = Alignment(horizontal="center")
        ws_summary.cell(row=current_row, column=6, value=out_clicks).alignment = Alignment(horizontal="center")
        ws_summary.cell(row=current_row, column=7, value=direct_cnt).alignment = Alignment(horizontal="center")
        ws_summary.cell(row=current_row, column=8, value=split_cnt).alignment = Alignment(horizontal="center")
        ws_summary.cell(row=current_row, column=9, value=para_cnt).alignment = Alignment(horizontal="center")
        ws_summary.cell(row=current_row, column=10, value=manual_cnt).alignment = Alignment(horizontal="center")
        ws_summary.cell(row=current_row, column=11, value=covered_cnt).alignment = Alignment(horizontal="center")
        ws_summary.cell(row=current_row, column=12, value=not_covered_cnt).alignment = Alignment(horizontal="center")

        status_val = s_res.status.value
        status_cell = ws_summary.cell(row=current_row, column=13, value=status_val)
        status_cell.alignment = Alignment(horizontal="center")
        if status_val == "Pass":
            status_cell.fill = pass_fill
            status_cell.font = pass_font
        elif status_val == "Manual Review":
            status_cell.fill = manual_fill
            status_cell.font = manual_font
        elif status_val == "Partial":
            status_cell.fill = warn_fill
            status_cell.font = warn_font
        else:
            status_cell.fill = fail_fill
            status_cell.font = fail_font

        ws_summary.cell(row=current_row, column=14, value=issues_str).alignment = Alignment(horizontal="left")

        for c in range(1, 15):
            ws_summary.cell(row=current_row, column=c).border = thin_border
        current_row += 1

    # -------------------------------------------------------------------------
    # SHEET 2: Verification (Dedicated sheet for automated audits)
    # -------------------------------------------------------------------------
    ws_verif = wb.create_sheet(title="Verification")
    ws_verif.views.sheetView[0].showGridLines = True

    for col_idx, header in enumerate(verif_headers, start=1):
        cell = ws_verif.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    row_v = 2
    for s_res in result.slide_results:
        v = s_res.verification
        orig_effs = v.original_effects_count if v else s_res.original_animation_count
        out_effs = v.output_effects_count if v else s_res.copied_animation_count
        direct_cnt = v.mapped_direct_count if v else s_res.mapped_direct_count
        split_cnt = v.mapped_group_split_count if v else s_res.mapped_group_split_count
        para_cnt = v.mapped_paragraph_level_count if v else s_res.mapped_paragraph_level_count
        manual_cnt = v.manual_review_count if v else s_res.manual_review_count
        covered_cnt = v.covered_count if v else s_res.covered_count
        not_covered_cnt = v.not_covered_count if v else s_res.not_covered_count
        orig_clicks = v.click_steps_orig if v else 0
        out_clicks = v.click_steps_output if v else 0
        issues_str = "; ".join(v.issues) if (v and v.issues) else (s_res.error_message or "None")

        ws_verif.cell(row=row_v, column=1, value=s_res.slide_index).alignment = Alignment(horizontal="center")
        ws_verif.cell(
            row=row_v,
            column=2,
            value=s_res.remediated_slide_index if s_res.remediated_slide_index else "N/A",
        ).alignment = Alignment(horizontal="center")
        ws_verif.cell(row=row_v, column=3, value=orig_effs).alignment = Alignment(horizontal="center")
        ws_verif.cell(row=row_v, column=4, value=out_effs).alignment = Alignment(horizontal="center")
        ws_verif.cell(row=row_v, column=5, value=orig_clicks).alignment = Alignment(horizontal="center")
        ws_verif.cell(row=row_v, column=6, value=out_clicks).alignment = Alignment(horizontal="center")
        ws_verif.cell(row=row_v, column=7, value=direct_cnt).alignment = Alignment(horizontal="center")
        ws_verif.cell(row=row_v, column=8, value=split_cnt).alignment = Alignment(horizontal="center")
        ws_verif.cell(row=row_v, column=9, value=para_cnt).alignment = Alignment(horizontal="center")
        ws_verif.cell(row=row_v, column=10, value=manual_cnt).alignment = Alignment(horizontal="center")
        ws_verif.cell(row=row_v, column=11, value=covered_cnt).alignment = Alignment(horizontal="center")
        ws_verif.cell(row=row_v, column=12, value=not_covered_cnt).alignment = Alignment(horizontal="center")

        status_val = s_res.status.value
        status_cell = ws_verif.cell(row=row_v, column=13, value=status_val)
        status_cell.alignment = Alignment(horizontal="center")
        if status_val == "Pass":
            status_cell.fill = pass_fill
            status_cell.font = pass_font
        elif status_val == "Manual Review":
            status_cell.fill = manual_fill
            status_cell.font = manual_font
        elif status_val == "Partial":
            status_cell.fill = warn_fill
            status_cell.font = warn_font
        else:
            status_cell.fill = fail_fill
            status_cell.font = fail_font

        ws_verif.cell(row=row_v, column=14, value=issues_str).alignment = Alignment(horizontal="left")

        for c in range(1, 15):
            ws_verif.cell(row=row_v, column=c).border = thin_border
        row_v += 1

    # -------------------------------------------------------------------------
    # SHEET 3: Manual Review
    # -------------------------------------------------------------------------
    ws_manual = wb.create_sheet(title="Manual Review")
    ws_manual.views.sheetView[0].showGridLines = True

    manual_headers = [
        "Slide",
        "Original Shape Name",
        "Text Snippet",
        "Type",
        "Bounding Box",
        "Effect Name",
        "Trigger",
        "Reason",
    ]

    for col_idx, header in enumerate(manual_headers, start=1):
        cell = ws_manual.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    row_m = 2
    for s_res in result.slide_results:
        for m_item in s_res.manual_review_items:
            ws_manual.cell(row=row_m, column=1, value=m_item.slide_index).alignment = Alignment(horizontal="center")
            ws_manual.cell(row=row_m, column=2, value=m_item.shape_name)
            ws_manual.cell(row=row_m, column=3, value=m_item.text_snippet)
            ws_manual.cell(row=row_m, column=4, value=m_item.shape_type).alignment = Alignment(horizontal="center")
            ws_manual.cell(row=row_m, column=5, value=m_item.bbox)
            ws_manual.cell(row=row_m, column=6, value=m_item.effect_name)
            ws_manual.cell(row=row_m, column=7, value=m_item.trigger).alignment = Alignment(horizontal="center")

            r_cell = ws_manual.cell(row=row_m, column=8, value=m_item.reason)
            r_cell.fill = manual_fill
            r_cell.font = manual_font

            for c in range(1, 9):
                ws_manual.cell(row=row_m, column=c).border = thin_border
            row_m += 1

    # -------------------------------------------------------------------------
    # SHEET 4: Animations
    # -------------------------------------------------------------------------
    ws_anim = wb.create_sheet(title="Animations")
    ws_anim.views.sheetView[0].showGridLines = True

    anim_headers = [
        "Slide No",
        "Order",
        "Shape Name",
        "Shape ID",
        "Preset Class",
        "Preset ID",
        "Subtype",
        "Node Type",
        "Effect Name",
        "Trigger",
        "Delay (ms)",
        "Duration (ms)",
        "Paragraph Range",
    ]

    for col_idx, header in enumerate(anim_headers, start=1):
        cell = ws_anim.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    row_idx = 2
    for s_res in result.slide_results:
        for a in s_res.animations:
            ws_anim.cell(row=row_idx, column=1, value=s_res.slide_index).alignment = Alignment(horizontal="center")
            ws_anim.cell(row=row_idx, column=2, value=a.order).alignment = Alignment(horizontal="center")
            ws_anim.cell(row=row_idx, column=3, value=a.shape_name)
            ws_anim.cell(row=row_idx, column=4, value=a.target_spid).alignment = Alignment(horizontal="center")
            ws_anim.cell(row=row_idx, column=5, value=a.preset_class).alignment = Alignment(horizontal="center")
            ws_anim.cell(row=row_idx, column=6, value=a.preset_id).alignment = Alignment(horizontal="center")
            ws_anim.cell(row=row_idx, column=7, value=a.preset_subtype).alignment = Alignment(horizontal="center")
            ws_anim.cell(row=row_idx, column=8, value=a.node_type).alignment = Alignment(horizontal="center")
            ws_anim.cell(row=row_idx, column=9, value=a.effect_name)
            ws_anim.cell(row=row_idx, column=10, value=a.trigger.value if hasattr(a.trigger, "value") else str(a.trigger)).alignment = Alignment(horizontal="center")
            ws_anim.cell(row=row_idx, column=11, value=a.delay).alignment = Alignment(horizontal="right")
            ws_anim.cell(row=row_idx, column=12, value=a.dur).alignment = Alignment(horizontal="right")
            ws_anim.cell(row=row_idx, column=13, value=a.paragraph_range or "All").alignment = Alignment(horizontal="center")

            for c in range(1, 14):
                ws_anim.cell(row=row_idx, column=c).border = thin_border
            row_idx += 1

    # -------------------------------------------------------------------------
    # SHEET 5: Skipped
    # -------------------------------------------------------------------------
    ws_skip = wb.create_sheet(title="Skipped")
    ws_skip.views.sheetView[0].showGridLines = True

    skip_headers = [
        "Slide No",
        "Shape Name",
        "Shape ID",
        "Effect Type",
        "Effect Name",
        "Reason",
    ]

    for col_idx, header in enumerate(skip_headers, start=1):
        cell = ws_skip.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    row_idx = 2
    for s_res in result.slide_results:
        for sk in s_res.skipped:
            ws_skip.cell(row=row_idx, column=1, value=sk.slide_index).alignment = Alignment(horizontal="center")
            ws_skip.cell(row=row_idx, column=2, value=sk.shape_name)
            ws_skip.cell(row=row_idx, column=3, value=sk.shape_id).alignment = Alignment(horizontal="center")
            ws_skip.cell(row=row_idx, column=4, value=sk.effect_type).alignment = Alignment(horizontal="center")
            ws_skip.cell(row=row_idx, column=5, value=sk.effect_name)
            ws_skip.cell(row=row_idx, column=6, value=sk.reason)

            for c in range(1, 7):
                ws_skip.cell(row=row_idx, column=c).border = thin_border
            row_idx += 1

    # Adjust column widths for all sheets
    for ws in [ws_summary, ws_verif, ws_manual, ws_anim, ws_skip]:
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or "")
                if "\n" in val_str:
                    val_str = max(val_str.split("\n"), key=len)
                if len(val_str) > max_len:
                    max_len = len(val_str)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
