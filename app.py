from __future__ import annotations

from pathlib import Path
from typing import Optional

import pandas as pd
import streamlit as st

from core.engine import transfer_presentation_animations
from core.models import DeckTransferResult
from core.parser import PresentationParseError

# -----------------------------------------------------------------------------
# Streamlit Page Config & Theme
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="PPT Animation Transfer",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
      @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

      html, body, [class*="css"] {
          font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
      }

      .hero-container {
          background: linear-gradient(135deg, #0F172A 0%, #1E293B 50%, #0F172A 100%);
          border: 1px solid rgba(255, 255, 255, 0.1);
          border-radius: 14px;
          padding: 1.8rem 2.2rem;
          margin-bottom: 1.5rem;
          box-shadow: 0 10px 30px -10px rgba(0, 0, 0, 0.5);
      }
      .hero-title {
          font-size: 2.2rem;
          font-weight: 800;
          background: linear-gradient(90deg, #38BDF8, #818CF8);
          -webkit-background-clip: text;
          -webkit-text-fill-color: transparent;
          margin-bottom: 0.3rem;
      }
      .hero-subtitle {
          color: #94A3B8;
          font-size: 1.05rem;
          font-weight: 400;
      }

      .metric-card {
          background: rgba(30, 41, 59, 0.7);
          border: 1px solid rgba(255, 255, 255, 0.08);
          border-radius: 12px;
          padding: 1.2rem 1.0rem;
          text-align: center;
          transition: transform 0.2s ease, border-color 0.2s ease;
      }
      .metric-card:hover {
          transform: translateY(-2px);
          border-color: rgba(56, 189, 248, 0.4);
      }
      .metric-val {
          font-size: 2.2rem;
          font-weight: 800;
          line-height: 1.1;
      }
      .metric-lbl {
          font-size: 0.75rem;
          color: #94A3B8;
          text-transform: uppercase;
          letter-spacing: 0.06em;
          margin-top: 0.35rem;
          font-weight: 600;
      }

      .badge {
          display: inline-block;
          padding: 0.25rem 0.7rem;
          border-radius: 9999px;
          font-size: 0.75rem;
          font-weight: 700;
          text-transform: uppercase;
          letter-spacing: 0.04em;
      }
      .badge-done { background: rgba(34, 197, 94, 0.2); color: #4ADE80; border: 1px solid rgba(34, 197, 94, 0.4); }
      .badge-partial { background: rgba(245, 158, 11, 0.2); color: #FBBF24; border: 1px solid rgba(245, 158, 11, 0.4); }
      .badge-skipped { background: rgba(239, 68, 68, 0.2); color: #F87171; border: 1px solid rgba(239, 68, 68, 0.4); }
      .badge-pass { background: rgba(34, 197, 94, 0.2); color: #4ADE80; border: 1px solid rgba(34, 197, 94, 0.4); }
      .badge-fail { background: rgba(239, 68, 68, 0.2); color: #F87171; border: 1px solid rgba(239, 68, 68, 0.4); }

      .skip-card {
          background: #1E293B;
          border-left: 4px solid #EF4444;
          border-radius: 0 8px 8px 0;
          padding: 1.0rem 1.2rem;
          margin-bottom: 0.8rem;
      }
      .skip-title {
          font-weight: 700;
          font-size: 0.95rem;
          color: #F1F5F9;
      }
      .skip-desc {
          font-size: 0.88rem;
          color: #CBD5E1;
          margin: 0.3rem 0;
      }

      .info-banner {
          background: rgba(56, 189, 248, 0.1);
          border: 1px solid rgba(56, 189, 248, 0.3);
          border-radius: 8px;
          padding: 0.8rem 1.0rem;
          margin-bottom: 1rem;
          color: #E2E8F0;
          font-size: 0.9rem;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Hero Header
# -----------------------------------------------------------------------------
st.markdown(
    """
    <div class="hero-container">
      <div class="hero-title">PPT Animation Transfer</div>
      <div class="hero-subtitle">
        Preserve and transfer all complex animations, triggers, motion paths, delays, durations, build lists, and transitions with intelligent shape remapping.
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# -----------------------------------------------------------------------------
# Initialize Session State
# -----------------------------------------------------------------------------
if "transfer_result" not in st.session_state:
    st.session_state.transfer_result = None
if "output_pptx_bytes" not in st.session_state:
    st.session_state.output_pptx_bytes = None
if "excel_report_bytes" not in st.session_state:
    st.session_state.excel_report_bytes = None
if "output_filename" not in st.session_state:
    st.session_state.output_filename = "presentation_animated.pptx"
if "report_filename" not in st.session_state:
    st.session_state.report_filename = "animation_transfer_report.xlsx"

# -----------------------------------------------------------------------------
# Sidebar: File Upload & Configuration
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 📁 Presentation Files")

    orig_file = st.file_uploader(
        "Original Presentation (.pptx)",
        type=["pptx", "potx", "pptm"],
        help="The presentation containing the master animations and slide transitions.",
        key="upload_orig",
    )

    rem_file = st.file_uploader(
        "Remediated Presentation (.pptx)",
        type=["pptx", "potx", "pptm"],
        help="The remediated target presentation that needs animations transferred.",
        key="upload_rem",
    )

    st.markdown("---")
    st.markdown("### ⚙️ Options")
    copy_transitions = st.checkbox(
        "Also copy slide transitions",
        value=True,
        help="Copies <p:transition> elements and AlternateContent wrappers as well.",
    )

    st.markdown("---")
    start_transfer = st.button(
        "🚀 Transfer Animations",
        type="primary",
        disabled=not (orig_file and rem_file),
        use_container_width=True,
    )

# -----------------------------------------------------------------------------
# Processing Engine Trigger
# -----------------------------------------------------------------------------
if start_transfer and orig_file and rem_file:
    orig_bytes = orig_file.getvalue()
    rem_bytes = rem_file.getvalue()

    orig_name = orig_file.name
    rem_name = rem_file.name

    base_name = Path(rem_name).stem
    st.session_state.output_filename = f"{base_name}_animated.pptx"
    st.session_state.report_filename = f"{base_name}_animation_report.xlsx"

    progress_bar = st.progress(0.0)
    status_text = st.empty()

    def update_progress(pct: float, message: str):
        progress_bar.progress(pct)
        status_text.info(f"⏳ {message}")

    try:
        result = transfer_presentation_animations(
            original_source=orig_bytes,
            remediated_source=rem_bytes,
            copy_transitions=copy_transitions,
            progress_callback=update_progress,
            original_filename=orig_name,
            remediated_filename=rem_name,
        )

        st.session_state.transfer_result = result
        st.session_state.output_pptx_bytes = result.output_pptx_bytes
        st.session_state.excel_report_bytes = result.excel_report_bytes

        progress_bar.progress(1.0)
        status_text.success("✨ Animation transfer completed successfully!")

    except PresentationParseError as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ **Validation Error**: {e}")
    except Exception as e:
        progress_bar.empty()
        status_text.empty()
        st.error(f"❌ **Unexpected Error during animation transfer**: {e}")

# -----------------------------------------------------------------------------
# Result Display
# -----------------------------------------------------------------------------
res: Optional[DeckTransferResult] = st.session_state.transfer_result

if res is not None:
    # 0. Faithfulness Banner
    total_orig_effs = sum(
        s.verification.original_effects_count
        for s in res.slide_results
        if s.verification is not None
    )
    total_out_effs = sum(
        s.verification.output_effects_count
        for s in res.slide_results
        if s.verification is not None
    )

    if res.is_100_percent_faithful:
        st.markdown(
            """
            <div style="background: rgba(34, 197, 94, 0.15); border: 1px solid rgba(34, 197, 94, 0.5); border-radius: 10px; padding: 1.1rem 1.4rem; margin-bottom: 1.5rem; display: flex; align-items: center; gap: 14px;">
              <span style="font-size: 2rem;">🎯</span>
              <div>
                <strong style="font-size: 1.2rem; color: #4ADE80;">100% Faithful Transfer</strong><br>
                <span style="color: #CBD5E1; font-size: 0.92rem;">Every original animation effect is either an exact match or listed as deliberately pruned/skipped. Zero corruption, zero hollow effects, zero click-step discrepancies.</span>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.markdown(
            f"""
            <div style="background: rgba(239, 68, 68, 0.15); border: 1px solid rgba(239, 68, 68, 0.5); border-radius: 10px; padding: 1.1rem 1.4rem; margin-bottom: 1.5rem; display: flex; align-items: center; gap: 14px;">
              <span style="font-size: 2rem;">⚠️</span>
              <div>
                <strong style="font-size: 1.2rem; color: #F87171;">Discrepancies Detected ({res.total_failed_slides} Failed Slide(s))</strong><br>
                <span style="color: #CBD5E1; font-size: 0.92rem;">Some effects or click steps could not be verified faithfully. Please inspect the Verification tab for details.</span>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # 1. Summary Cards
    c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
    with c1:
        st.markdown(
            f"""
            <div class="metric-card">
              <div class="metric-val" style="color: #38BDF8;">{total_orig_effs}</div>
              <div class="metric-lbl">Original Effects</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c2:
        st.markdown(
            f"""
            <div class="metric-card">
              <div class="metric-val" style="color: #818CF8;">{total_out_effs}</div>
              <div class="metric-lbl">Output Effects</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c3:
        st.markdown(
            f"""
            <div class="metric-card">
              <div class="metric-val" style="color: #4ADE80;">{res.total_mapped_direct}</div>
              <div class="metric-lbl">Direct Mapped</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c4:
        st.markdown(
            f"""
            <div class="metric-card">
              <div class="metric-val" style="color: #34D399;">{res.total_mapped_group_split}</div>
              <div class="metric-lbl">Group Split</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c5:
        st.markdown(
            f"""
            <div class="metric-card">
              <div class="metric-val" style="color: {'#FB923C' if res.total_manual_review > 0 else '#94A3B8'};">
                {res.total_manual_review}
              </div>
              <div class="metric-lbl">Manual Review</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c6:
        st.markdown(
            f"""
            <div class="metric-card">
              <div class="metric-val" style="color: #22D3EE;">{res.total_covered_effects}</div>
              <div class="metric-lbl">Covered</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with c7:
        st.markdown(
            f"""
            <div class="metric-card">
              <div class="metric-val" style="color: {'#F87171' if res.total_not_covered_effects > 0 else '#4ADE80'};">
                {res.total_not_covered_effects}
              </div>
              <div class="metric-lbl">Not Covered</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # Highlight Manual Review Banner if items exist
    if res.total_manual_review > 0:
        st.markdown(
            f"""
            <div style="background: rgba(251, 146, 60, 0.15); border: 1px solid rgba(251, 146, 60, 0.5); border-radius: 10px; padding: 1.1rem 1.4rem; margin-bottom: 1.5rem; display: flex; align-items: center; gap: 14px;">
              <span style="font-size: 2rem;">🔍</span>
              <div>
                <strong style="font-size: 1.2rem; color: #FB923C;">Manual Review Required ({res.total_manual_review} effect(s))</strong><br>
                <span style="color: #CBD5E1; font-size: 0.92rem;">Certain effects targeted merged shapes or restructured components that require human review. These effects were never dropped silently. See the <strong>🔍 Manual Review</strong> tab.</span>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # 2. Download Center & Verification Instruction
    st.markdown(
        """
        <div class="info-banner">
          💡 <strong>Verification Recommended:</strong> Please download the new animated presentation, open it in 
          <strong>Microsoft PowerPoint</strong>, and play the slideshow (<strong>F5 / Shift+F5</strong>) for a final manual check.
        </div>
        """,
        unsafe_allow_html=True,
    )

    down_col1, down_col2 = st.columns(2)
    with down_col1:
        if st.session_state.output_pptx_bytes:
            st.download_button(
                label=f"📥 Download Animated PPTX ({st.session_state.output_filename})",
                data=st.session_state.output_pptx_bytes,
                file_name=st.session_state.output_filename,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                type="primary",
                use_container_width=True,
            )
    with down_col2:
        if st.session_state.excel_report_bytes:
            st.download_button(
                label="📊 Download Excel Report (.xlsx)",
                data=st.session_state.excel_report_bytes,
                file_name=st.session_state.report_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    st.markdown("<br>", unsafe_allow_html=True)

    # 3. Detailed Tabs
    tab_verify, tab_manual, tab_overview, tab_details, tab_skipped = st.tabs(
        [
            "✅ Verification Table",
            f"🔍 Manual Review ({res.total_manual_review})",
            "📋 Slide Summary",
            "🎬 Animation Details",
            f"⚠️ Skipped Animations ({res.total_animations_skipped})",
        ]
    )

    # --- TAB 1: Verification Table ---
    with tab_verify:
        st.markdown("#### Automated Content-Based Slide Verification")
        st.caption("Content-based verification confirms each original effect is covered by matching text or overlapping bounding box. Never relies solely on shape names or IDs.")

        verif_rows = []
        for s in res.slide_results:
            if s.verification:
                v = s.verification
                verif_rows.append(
                    {
                        "Slide No": s.slide_index,
                        "Remediated Slide": s.remediated_slide_index if s.remediated_slide_index else "Unmatched",
                        "Original Effects": v.original_effects_count,
                        "Output Effects": v.output_effects_count,
                        "Clicks (Orig -> Out)": f"{v.click_steps_orig} -> {v.click_steps_output}",
                        "Mapped (Direct)": v.mapped_direct_count,
                        "Mapped (Group Split)": v.mapped_group_split_count,
                        "Mapped (Paragraph)": v.mapped_paragraph_level_count,
                        "Manual Review": v.manual_review_count,
                        "Covered": v.covered_count,
                        "Not Covered": v.not_covered_count,
                        "Status": s.status.value,
                        "Issues / Notes": "; ".join(v.issues) if v.issues else "All effects verified",
                    }
                )

        df_verif = pd.DataFrame(verif_rows)
        st.dataframe(df_verif, use_container_width=True, hide_index=True)

    # --- TAB 2: Manual Review ---
    with tab_manual:
        st.markdown("#### Manual Review Items")
        st.caption("Animations that could not be mapped automatically or involve complex changes (e.g. text from two clicks merged into one placeholder, or slide rebuilt as one picture).")

        if not res.manual_review_items:
            st.success("🎉 Zero manual review items! All effects transferred cleanly.")
        else:
            m_rows = []
            for item in res.manual_review_items:
                m_rows.append(
                    {
                        "Slide": item.slide_index,
                        "Original Shape": item.shape_name,
                        "Shape ID": item.shape_id,
                        "Text Snippet": item.text_snippet,
                        "Type": item.shape_type,
                        "Bounding Box": item.bbox,
                        "Effect Name": item.effect_name,
                        "Trigger": item.trigger,
                        "Reason": item.reason,
                    }
                )
            df_manual = pd.DataFrame(m_rows)
            st.dataframe(df_manual, use_container_width=True, hide_index=True)

    # --- TAB 3: Slide Summary ---
    with tab_overview:
        st.markdown("#### Slide-by-Slide Transfer Status")

        slide_rows = []
        for s in res.slide_results:
            status_badge = s.status.value
            slide_rows.append(
                {
                    "Slide No": s.slide_index,
                    "Remediated Slide": s.remediated_slide_index if s.remediated_slide_index else "Unmatched",
                    "Original Animations": s.original_animation_count,
                    "Copied Count": s.copied_animation_count,
                    "Direct Mapped": s.mapped_direct_count,
                    "Group Split": s.mapped_group_split_count,
                    "Manual Review": s.manual_review_count,
                    "Covered": s.covered_count,
                    "Transition Copied": "Yes" if s.copied_transition else "No",
                    "Status": status_badge,
                }
            )

        df_slides = pd.DataFrame(slide_rows)
        st.dataframe(df_slides, use_container_width=True, hide_index=True)

    # --- TAB 3: Animation Details ---
    with tab_details:
        st.markdown("#### Per-Slide Animation Breakdown")
        active_slide_nums = [s.slide_index for s in res.slide_results if s.copied_animation_count > 0]

        if not active_slide_nums:
            st.info("No animations were transferred in this presentation.")
        else:
            selected_slide_num = st.selectbox("Select Slide to Inspect:", active_slide_nums)
            selected_slide = next(s for s in res.slide_results if s.slide_index == selected_slide_num)

            anim_rows = []
            for a in selected_slide.animations:
                anim_rows.append(
                    {
                        "Order": a.order,
                        "Shape Name": a.shape_name,
                        "Shape ID": a.target_spid,
                        "Preset Class": a.preset_class,
                        "Preset ID": a.preset_id,
                        "Node Type": a.node_type,
                        "Effect Name": a.effect_name,
                        "Trigger": a.trigger.value if hasattr(a.trigger, "value") else str(a.trigger),
                        "Delay": f"{a.delay} ms",
                        "Duration": f"{a.dur} ms",
                        "Paragraph Range": a.paragraph_range or "Full Shape",
                    }
                )

            df_anims = pd.DataFrame(anim_rows)
            st.dataframe(df_anims, use_container_width=True, hide_index=True)

    # --- TAB 4: Skipped Animations ---
    with tab_skipped:
        st.markdown("#### Orphaned & Skipped Animations")
        all_skipped = []
        for s in res.slide_results:
            all_skipped.extend(s.skipped)

        if not all_skipped:
            st.success("🎉 Zero animations were skipped! Every animation mapped cleanly.")
        else:
            for sk in all_skipped:
                st.markdown(
                    f"""
                    <div class="skip-card">
                      <div class="skip-title">Slide {sk.slide_index} — {sk.shape_name} (ID: {sk.shape_id})</div>
                      <div class="skip-desc">
                        <strong>Effect:</strong> {sk.effect_type} ({sk.effect_name})<br>
                        <strong>Reason:</strong> {sk.reason}
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

else:
    # Empty State Guidance
    st.info("👈 Upload your **Original** and **Remediated** PowerPoint presentations in the sidebar to begin.")
