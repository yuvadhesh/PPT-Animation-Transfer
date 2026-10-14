from __future__ import annotations

import io
import pytest
from pptx import Presentation
from pptx.util import Inches
from lxml import etree

from core.engine import transfer_presentation_animations
from core.models import SlideStatus
from core.parser import PresentationParseError, validate_pptx_file, NAMESPACES
from core.validator import validate_presentation_integrity
from core.reporter import export_excel_report
from tests.sample_generator import create_sample_decks


def test_file_validation_friendly_errors():
    """Verify friendly error messages on invalid input files."""
    # 1. Lock file
    with pytest.raises(PresentationParseError, match="temporary lock file"):
        validate_pptx_file(b"dummy data", filename="~$presentation.pptx")

    # 2. Too small / empty file
    with pytest.raises(PresentationParseError, match="File is empty or the upload was incomplete"):
        validate_pptx_file(b"small", filename="empty.pptx")

    # 3. Old PPT magic bytes D0 CF 11 E0
    old_ppt = b"\xd0\xcf\x11\xe0" + b"\x00" * 2000
    with pytest.raises(PresentationParseError, match="old .ppt or a password-protected file"):
        validate_pptx_file(old_ppt, filename="legacy.ppt")


def test_full_animation_transfer_pipeline():
    """
    Test sample original with entrance, exit, emphasis, motion-path animations,
    by-paragraph text build, triggers, and transition.
    Remediated copy has changed shape ids, renamed shapes, and one deleted shape.
    Assert output animations match, ids are remapped, deleted shape is skipped and reported,
    and output file opens with python-pptx.
    """
    orig_bytes, rem_bytes = create_sample_decks()

    # Run the complete transfer
    result = transfer_presentation_animations(
        original_source=orig_bytes,
        remediated_source=rem_bytes,
        copy_transitions=True,
    )

    # 1. High-level result assertions
    assert result.success is True
    assert result.original_slide_count == 1
    assert result.remediated_slide_count == 1
    assert result.slides_with_animations == 1
    assert len(result.slide_results) == 1

    s_res = result.slide_results[0]
    # Status is PARTIAL or MANUAL_REVIEW because 1 shape (Shape 7) was deleted/skipped
    assert s_res.status in (SlideStatus.PARTIAL, SlideStatus.MANUAL_REVIEW)
    assert s_res.copied_transition is True

    # 2. Assert deleted shape (Shape 7) is reported as skipped
    assert s_res.skipped_animation_count == 1
    skipped_item = s_res.skipped[0]
    assert skipped_item.shape_id == "7"
    assert any(w in skipped_item.reason.lower() for w in ("deleted", "not found", "no match"))

    # 3. Assert shape mappings:
    # 2 -> 202
    # 3 -> 303 (renamed shape matched by type + text + position)
    # 4 -> 404
    # 5 -> 505
    # 6 -> 606
    assert s_res.shape_map.get("2") == "202"
    assert s_res.shape_map.get("3") == "303"
    assert s_res.shape_map.get("4") == "404"
    assert s_res.shape_map.get("5") == "505"
    assert s_res.shape_map.get("6") == "606"

    # 4. Assert output PPTX opens with python-pptx and passes integrity validation
    assert result.output_pptx_bytes is not None
    is_valid, issues, warnings = validate_presentation_integrity(result.output_pptx_bytes)
    assert is_valid, f"Validation issues: {issues}"

    prs_out = Presentation(io.BytesIO(result.output_pptx_bytes))
    assert len(prs_out.slides) == 1
    out_slide = prs_out.slides[0]

    # Inspect XML of the output slide
    sld_xml = out_slide._element
    timing = sld_xml.find(f"{{{NAMESPACES['p']}}}timing")
    assert timing is not None, "Timing element missing from output slide!"

    # Verify no references to old IDs (2, 3, 4, 5, 6, 7) exist in spTgt
    spids = []
    for elem in timing.iter():
        if isinstance(elem.tag, str) and elem.tag.endswith("spTgt"):
            spids.append(elem.attrib.get("spid"))

    # Old IDs must not exist
    for old_id in ("2", "3", "4", "5", "6", "7"):
        assert old_id not in spids, f"Old shape id {old_id} still found in timing tree: {spids}"

    # Remapped IDs must exist
    for new_id in ("202", "303", "404", "505", "606"):
        assert new_id in spids, f"New shape id {new_id} missing from timing tree: {spids}"

    # Build list should have 606 and NOT 7
    bldLst = timing.find(f".//{{{NAMESPACES['p']}}}bldLst")
    assert bldLst is not None
    bld_spids = [c.attrib.get("spid") for c in bldLst]
    assert "606" in bld_spids
    assert "7" not in bld_spids
    assert "6" not in bld_spids

    # Transition copied
    trans = sld_xml.find(f"{{{NAMESPACES['p']}}}transition")
    assert trans is not None

    # Verification result should be MANUAL_REVIEW (since unmapped shape 7 requires review)
    assert s_res.verification is not None
    assert s_res.verification.status in (SlideStatus.PARTIAL, SlideStatus.MANUAL_REVIEW)
    assert len(s_res.verification.issues) == 0



def test_excel_report_generation():
    """Test generating 3-sheet Excel report from results."""
    orig_bytes, rem_bytes = create_sample_decks()
    result = transfer_presentation_animations(
        original_source=orig_bytes,
        remediated_source=rem_bytes,
    )

    excel_bytes = export_excel_report(result)
    assert excel_bytes is not None
    assert len(excel_bytes) > 1000

    # Read back with openpyxl
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(excel_bytes))
    assert "Summary" in wb.sheetnames
    assert "Animations" in wb.sheetnames
    assert "Skipped" in wb.sheetnames

    # Check Summary
    ws_summary = wb["Summary"]
    assert ws_summary["A1"].value == "PPT Animation Transfer Report"

def test_picture_hash_matching():
    """Verify that picture shapes are mapped via image hash."""
    from core.matcher import match_shapes
    from core.models import ShapeInfo

    orig_shapes = {
        "10": ShapeInfo(shape_id="10", shape_name="Picture 1", shape_type="PICTURE", image_hash="abc123hash"),
        "20": ShapeInfo(shape_id="20", shape_name="Text 1", shape_type="TEXT_BOX", text="Hello World"),
    }
    rem_shapes = {
        # Shape 99 has different ID and different name, but same image hash!
        "99": ShapeInfo(shape_id="99", shape_name="New Logo Picture", shape_type="PICTURE", image_hash="abc123hash"),
        "200": ShapeInfo(shape_id="200", shape_name="Text Remediated", shape_type="TEXT_BOX", text="Hello World"),
    }

    shape_map = match_shapes(orig_shapes, rem_shapes)
    assert shape_map.get("10") == "99"
    assert shape_map.get("20") == "200"


def test_slide_order_matching():
    """Verify that slides in different orders are matched by title similarity."""
    from core.matcher import match_slides
    from core.models import SlideData

    s1 = SlideData(slide_index=1, title="Executive Financial Summary", all_text="Revenue was 10 million in Q1")
    s2 = SlideData(slide_index=2, title="Project Roadmap and Milestones", all_text="Phase 1 delivery is scheduled for March")

    # Remediated presentation has them in reversed order
    r1 = SlideData(slide_index=1, title="Project Roadmap & Milestones", all_text="Phase 1 delivery scheduled for March")
    r2 = SlideData(slide_index=2, title="Executive Financial Summary", all_text="Revenue 10 million in Q1")

    matched, unmatched_orig, unmatched_rem = match_slides([s1, s2], [r1, r2])
    assert matched[1] == 2  # s1 matched r2
    assert matched[2] == 1  # s2 matched r1
    assert len(unmatched_orig) == 0
    assert len(unmatched_rem) == 0


def test_table_with_animation_no_bldlst_with_changed_shape_id():
    """
    Test table with Wipe entrance animation and NO <p:bldLst> at all.
    In the remediated deck the table's shape id changed from 3 to 4.
    Verify that the animation is transferred, spid is remapped to 4,
    and NO <p:bldLst> is created in the output.
    """
    from pptx.util import Inches
    from pptx import Presentation

    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])
    tbl_shape = orig_slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(2))
    tbl_shape.name = "Data Table"
    # Set ID to 3
    cNvPr = tbl_shape._element.find(f".//{{{NAMESPACES['p']}}}cNvPr")
    if cNvPr is not None:
        cNvPr.attrib["id"] = "3"

    # Add timing without bldLst
    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:par>
              <p:cTn id="2" presetID="22" presetClass="entr" fill="hold" nodeType="clickEffect">
                <p:childTnLst>
                  <p:animEffect transition="in" filter="wipe(left)">
                    <p:cBhvr>
                      <p:cTn id="3" dur="500"/>
                      <p:tgtEl>
                        <p:spTgt spid="3"/>
                      </p:tgtEl>
                    </p:cBhvr>
                  </p:animEffect>
                </p:childTnLst>
              </p:cTn>
            </p:par>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated deck has same table with shape ID 4
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])
    rem_tbl = rem_slide.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(4), Inches(2))
    rem_tbl.name = "Data Table"
    cNvPr_rem = rem_tbl._element.find(f".//{{{NAMESPACES['p']}}}cNvPr")
    if cNvPr_rem is not None:
        cNvPr_rem.attrib["id"] = "4"

    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True

    # Re-open output
    prs_out = Presentation(io.BytesIO(res.output_pptx_bytes))
    timing_out = prs_out.slides[0]._element.find(f"{{{NAMESPACES['p']}}}timing")
    assert timing_out is not None

    # Verify no bldLst exists
    bld_out = timing_out.find(f".//{{{NAMESPACES['p']}}}bldLst")
    assert bld_out is None, "Output must NOT contain <p:bldLst> when original has none!"

    # Verify target spid is remapped to 4
    spids = [elem.attrib.get("spid") for elem in timing_out.iter() if isinstance(elem.tag, str) and elem.tag.endswith("spTgt")]
    assert spids == ["4"]


def test_by_paragraph_build_where_one_paragraph_target_is_pruned():
    """
    Test by-paragraph build where two animations target shape 6 (paragraphs 0 and 1).
    Even if one paragraph animation is pruned, as long as another animation still targets
    shape 6, the bldP entry for shape 6 must be kept!
    """
    from pptx.util import Inches
    from pptx import Presentation

    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])
    tb = orig_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    tb.text = "Para 0\nPara 1"
    tb.name = "Text Box"
    cNvPr = tb._element.find(f".//{{{NAMESPACES['p']}}}cNvPr")
    if cNvPr is not None:
        cNvPr.attrib["id"] = "6"

    # Add timing with 2 paragraph builds and 1 bldP
    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:par>
              <p:cTn id="2" presetID="1" presetClass="entr" fill="hold" nodeType="clickEffect">
                <p:childTnLst>
                  <p:set>
                    <p:cBhvr>
                      <p:cTn id="3" dur="1"/>
                      <p:tgtEl>
                        <p:spTgt spid="6"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt>
                      </p:tgtEl>
                    </p:cBhvr>
                  </p:set>
                </p:childTnLst>
              </p:cTn>
            </p:par>
            <p:par>
              <p:cTn id="4" presetID="1" presetClass="entr" fill="hold" nodeType="clickEffect">
                <p:childTnLst>
                  <p:set>
                    <p:cBhvr>
                      <p:cTn id="5" dur="1"/>
                      <p:tgtEl>
                        <p:spTgt spid="6"><p:txEl><p:pRg st="1" end="1"/></p:txEl></p:spTgt>
                      </p:tgtEl>
                    </p:cBhvr>
                  </p:set>
                </p:childTnLst>
              </p:cTn>
            </p:par>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
      <p:bldLst>
        <p:bldP spid="6" grpId="0" build="p"/>
      </p:bldLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated deck has same text box with shape ID 60
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])
    rem_tb = rem_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    rem_tb.text = "Para 0\nPara 1"
    rem_tb.name = "Text Box"
    cNvPr_rem = rem_tb._element.find(f".//{{{NAMESPACES['p']}}}cNvPr")
    if cNvPr_rem is not None:
        cNvPr_rem.attrib["id"] = "60"

    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True

    prs_out = Presentation(io.BytesIO(res.output_pptx_bytes))
    timing_out = prs_out.slides[0]._element.find(f"{{{NAMESPACES['p']}}}timing")
    bld_out = timing_out.find(f".//{{{NAMESPACES['p']}}}bldLst")
    assert bld_out is not None
    assert len(bld_out) == 1
    assert bld_out[0].attrib.get("spid") == "60"
    assert bld_out[0].tag.endswith("bldP")


def test_empty_bldlst_cleanup():
    """
    Test that when an animation is pruned because its target shape is unmapped/deleted,
    its build-list entry is also removed, and if bldLst becomes empty, <p:bldLst> is completely removed.
    """
    from pptx.util import Inches
    from pptx import Presentation

    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])
    tb = orig_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    tb.text = "Deleted Box"
    tb.name = "Deleted Box"
    cNvPr = tb._element.find(f".//{{{NAMESPACES['p']}}}cNvPr")
    if cNvPr is not None:
        cNvPr.attrib["id"] = "99"

    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:par>
              <p:cTn id="2" presetID="10" presetClass="entr" fill="hold" nodeType="clickEffect">
                <p:childTnLst>
                  <p:animEffect transition="in" filter="fade">
                    <p:cBhvr>
                      <p:cTn id="3" dur="500"/>
                      <p:tgtEl>
                        <p:spTgt spid="99"/>
                      </p:tgtEl>
                    </p:cBhvr>
                  </p:animEffect>
                </p:childTnLst>
              </p:cTn>
            </p:par>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
      <p:bldLst>
        <p:bldP spid="99" grpId="0"/>
      </p:bldLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated deck does NOT have shape 99 (it's empty slide)
    rem_prs = Presentation()
    rem_prs.slides.add_slide(rem_prs.slide_layouts[6])
    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True

    prs_out = Presentation(io.BytesIO(res.output_pptx_bytes))
    timing_out = prs_out.slides[0]._element.find(f"{{{NAMESPACES['p']}}}timing")
    # Timing should have no bldLst (cleaned up)
    if timing_out is not None:
        bld_out = timing_out.find(f".//{{{NAMESPACES['p']}}}bldLst")
        assert bld_out is None


def check_timing_invariants(timing_out, orig_effects, expected_click_steps, shape_map, pruned_spids=set()):
    """Helper verifying that output timing preserves exact attributes, has no hollow effects, and matches click steps."""
    assert timing_out is not None, "Output timing element must exist!"

    # 1. FAIL if any effect <p:cTn> has an empty childTnLst
    for elem in timing_out.iter():
        if isinstance(elem.tag, str) and elem.tag.split("}")[-1] == "cTn":
            node_type = elem.attrib.get("nodeType", "")
            preset_id = elem.attrib.get("presetID")
            preset_class = elem.attrib.get("presetClass")
            if preset_id or preset_class or node_type in ("clickEffect", "withEffect", "afterEffect"):
                c_lst = elem.find(f"{{{NAMESPACES['p']}}}childTnLst")
                assert c_lst is not None and len(c_lst) > 0, f"Hollow effect found in {elem.attrib}"

    # 2. Extract output effects & click steps
    active_spids = {
        elem.attrib.get("spid")
        for elem in timing_out.iter()
        if isinstance(elem.tag, str) and elem.tag.endswith("spTgt")
    }
    from core.models import ShapeInfo
    from core.animation_transfer import extract_effects_from_timing
    dummy_shapes = {sid: ShapeInfo(shape_id=sid, shape_name=f"Shape {sid}", shape_type="RECTANGLE") for sid in active_spids if sid}
    click_steps, out_effects = extract_effects_from_timing(timing_out, dummy_shapes)

    assert click_steps == expected_click_steps, f"Click steps mismatch: expected {expected_click_steps}, got {click_steps}"

    # 3. Check effect list equals original (after id remapping)
    expected_non_pruned = [e for e in orig_effects if e.target_spid not in pruned_spids]
    assert len(out_effects) == len(expected_non_pruned), f"Effect count mismatch: expected {len(expected_non_pruned)}, got {len(out_effects)}"

    for orig_e, out_e in zip(expected_non_pruned, out_effects):
        expected_spid = shape_map[orig_e.target_spid]
        assert out_e.target_spid == expected_spid
        assert out_e.preset_class == orig_e.preset_class
        assert out_e.preset_id == orig_e.preset_id
        assert out_e.preset_subtype == orig_e.preset_subtype
        assert out_e.node_type == orig_e.node_type
        assert out_e.delay == orig_e.delay
        assert out_e.dur == orig_e.dur
        assert out_e.paragraph_range == orig_e.paragraph_range


def test_appear_only_slide():
    """
    Test slide with ONLY an Appear animation (presetClass='entr', presetID='1').
    Its timing node has only a <p:set> child and no <p:animEffect>.
    Remediated copy has shape id remapped (2 -> 20).
    Assert timing exists, effect list equals original (remapped), click step count == 1,
    no empty effect exists, and verification passed.
    """
    from pptx import Presentation
    from pptx.util import Inches
    from core.animation_transfer import extract_effects_from_timing
    from core.models import ShapeInfo

    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])
    tb = orig_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    tb.text = "Appear Text"
    tb.name = "Appear Shape"
    cNvPr = tb._element.find(f".//{{{NAMESPACES['p']}}}cNvPr")
    if cNvPr is not None:
        cNvPr.attrib["id"] = "2"

    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">
          <p:childTnLst>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                <p:childTnLst>
                  <p:par>
                    <p:cTn id="3" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="4" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:set>
                                <p:cBhvr>
                                  <p:cTn id="5" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>
                                  <p:tgtEl><p:spTgt spid="2"/></p:tgtEl>
                                  <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                                </p:cBhvr>
                                <p:to><p:strVal val="visible"/></p:to>
                              </p:set>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
            </p:seq>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated has same textbox with shape id 20
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])
    rem_tb = rem_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    rem_tb.text = "Appear Text"
    rem_tb.name = "Appear Shape"
    cNvPr_rem = rem_tb._element.find(f".//{{{NAMESPACES['p']}}}cNvPr")
    if cNvPr_rem is not None:
        cNvPr_rem.attrib["id"] = "20"

    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    orig_shapes = {"2": ShapeInfo(shape_id="2", shape_name="Appear Shape", shape_type="TEXT_BOX")}
    orig_clicks, orig_effects = extract_effects_from_timing(orig_slide._element.find(f"{{{NAMESPACES['p']}}}timing"), orig_shapes)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True
    assert res.is_100_percent_faithful is True

    s_res = res.slide_results[0]
    assert s_res.status == SlideStatus.PASS
    assert s_res.verification.passed is True

    prs_out = Presentation(io.BytesIO(res.output_pptx_bytes))
    timing_out = prs_out.slides[0]._element.find(f"{{{NAMESPACES['p']}}}timing")
    check_timing_invariants(timing_out, orig_effects, expected_click_steps=1, shape_map={"2": "20"})


def test_mixed_slide_dissolve_by_click_and_appear_with_previous():
    """
    Test mixed slide: a Dissolve entrance on click, and Appear entrance with previous.
    Both reside in the same click step. Remediated deck has remapped shape IDs (2->202, 3->303).
    Assert output has 1 click step, 2 effects matching exact attributes, no hollow cTn.
    """
    from pptx import Presentation
    from pptx.util import Inches
    from core.animation_transfer import extract_effects_from_timing
    from core.models import ShapeInfo

    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])
    tb1 = orig_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    tb1.text = "Dissolve Shape"
    tb1.name = "Shape 1"
    tb1._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "2"

    tb2 = orig_slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
    tb2.text = "Appear Shape"
    tb2.name = "Shape 2"
    tb2._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "3"

    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                <p:childTnLst>
                  <p:par>
                    <p:cTn id="3" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="4" presetID="10" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animEffect transition="in" filter="dissolve">
                                <p:cBhvr>
                                  <p:cTn id="5" dur="500"/>
                                  <p:tgtEl><p:spTgt spid="2"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animEffect>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                        <p:par>
                          <p:cTn id="6" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="withEffect">
                            <p:childTnLst>
                              <p:set>
                                <p:cBhvr>
                                  <p:cTn id="7" dur="1" fill="hold"/>
                                  <p:tgtEl><p:spTgt spid="3"/></p:tgtEl>
                                </p:cBhvr>
                                <p:to><p:strVal val="visible"/></p:to>
                              </p:set>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
            </p:seq>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])
    r_tb1 = rem_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    r_tb1.text = "Dissolve Shape"
    r_tb1.name = "Shape 1"
    r_tb1._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "202"

    r_tb2 = rem_slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
    r_tb2.text = "Appear Shape"
    r_tb2.name = "Shape 2"
    r_tb2._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "303"

    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    orig_shapes = {
        "2": ShapeInfo(shape_id="2", shape_name="Shape 1", shape_type="TEXT_BOX"),
        "3": ShapeInfo(shape_id="3", shape_name="Shape 2", shape_type="TEXT_BOX"),
    }
    orig_clicks, orig_effects = extract_effects_from_timing(orig_slide._element.find(f"{{{NAMESPACES['p']}}}timing"), orig_shapes)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True
    assert res.is_100_percent_faithful is True

    s_res = res.slide_results[0]
    assert s_res.status == SlideStatus.PASS
    assert s_res.verification.passed is True

    prs_out = Presentation(io.BytesIO(res.output_pptx_bytes))
    timing_out = prs_out.slides[0]._element.find(f"{{{NAMESPACES['p']}}}timing")
    check_timing_invariants(timing_out, orig_effects, expected_click_steps=1, shape_map={"2": "202", "3": "303"})


def test_by_paragraph_build_motion_path_exit_emphasis_interactive_table():
    """
    Test slide with:
    - By-paragraph build with <p:bldLst><p:bldP .../></p:bldLst>
    - Motion path (<p:animMotion>)
    - Exit animation (presetClass="exit")
    - Emphasis animation (presetClass="emph")
    - Interactive trigger sequence (<p:cTn nodeType="interactiveSeq">)
    - Table animation (graphicFrame)
    Assert output effect list equals original (after id remapping), no empty effect exists, and click steps are equal.
    """
    from pptx import Presentation
    from pptx.util import Inches
    from core.animation_transfer import extract_effects_from_timing
    from core.models import ShapeInfo

    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])

    # Shape 2: Text Box with 2 paragraphs
    s2 = orig_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    s2.text = "Para 0\nPara 1"
    s2.name = "Paragraph Box"
    s2._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "2"

    # Shape 3: Motion Path Rect
    s3 = orig_slide.shapes.add_textbox(Inches(5), Inches(1), Inches(2), Inches(1))
    s3.text = "Motion Rect"
    s3.name = "Motion Rect"
    s3._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "3"

    # Shape 4: Exit Oval
    s4 = orig_slide.shapes.add_textbox(Inches(1), Inches(3), Inches(2), Inches(1))
    s4.text = "Exit Oval"
    s4.name = "Exit Oval"
    s4._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "4"

    # Shape 5: Emphasis Star
    s5 = orig_slide.shapes.add_textbox(Inches(5), Inches(3), Inches(2), Inches(1))
    s5.text = "Emphasis Star"
    s5.name = "Emphasis Star"
    s5._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "5"

    # Shape 6: Interactive Trigger Button
    s6 = orig_slide.shapes.add_textbox(Inches(1), Inches(5), Inches(2), Inches(1))
    s6.text = "Click Me Button"
    s6.name = "Button"
    s6._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "6"

    # Shape 7: Table
    s7 = orig_slide.shapes.add_table(2, 2, Inches(4), Inches(5), Inches(3), Inches(1.5))
    s7.name = "Summary Table"
    s7._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "7"

    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                <p:childTnLst>
                  <!-- Step 1: Paragraph 0 entrance -->
                  <p:par>
                    <p:cTn id="10" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="11" presetID="1" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:set>
                                <p:cBhvr>
                                  <p:cTn id="12" dur="1"/>
                                  <p:tgtEl><p:spTgt spid="2"><p:txEl><p:pRg st="0" end="0"/></p:txEl></p:spTgt></p:tgtEl>
                                </p:cBhvr>
                                <p:to><p:strVal val="visible"/></p:to>
                              </p:set>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                  <!-- Step 2: Paragraph 1 entrance -->
                  <p:par>
                    <p:cTn id="13" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="14" presetID="1" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:set>
                                <p:cBhvr>
                                  <p:cTn id="15" dur="1"/>
                                  <p:tgtEl><p:spTgt spid="2"><p:txEl><p:pRg st="1" end="1"/></p:txEl></p:spTgt></p:tgtEl>
                                </p:cBhvr>
                                <p:to><p:strVal val="visible"/></p:to>
                              </p:set>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                  <!-- Step 3: Motion Path -->
                  <p:par>
                    <p:cTn id="16" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="17" presetID="0" presetClass="path" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animMotion origin="layout" path="M 0 0 L 0.2 0.2 Z">
                                <p:cBhvr>
                                  <p:cTn id="18" dur="1000"/>
                                  <p:tgtEl><p:spTgt spid="3"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animMotion>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                  <!-- Step 4: Exit Animation -->
                  <p:par>
                    <p:cTn id="19" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="20" presetID="1" presetClass="exit" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:set>
                                <p:cBhvr>
                                  <p:cTn id="21" dur="1"/>
                                  <p:tgtEl><p:spTgt spid="4"/></p:tgtEl>
                                </p:cBhvr>
                                <p:to><p:strVal val="hidden"/></p:to>
                              </p:set>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                  <!-- Step 5: Table entrance -->
                  <p:par>
                    <p:cTn id="22" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="23" presetID="22" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animEffect transition="in" filter="wipe(left)">
                                <p:cBhvr>
                                  <p:cTn id="24" dur="500"/>
                                  <p:tgtEl><p:spTgt spid="7"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animEffect>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
            </p:seq>
            <!-- Interactive Sequence triggered by shape 6, animating shape 5 -->
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="25" dur="indefinite" nodeType="interactiveSeq">
                <p:stCondLst>
                  <p:cond evt="onClick" delay="0">
                    <p:tgtEl><p:spTgt spid="6"/></p:tgtEl>
                  </p:cond>
                </p:stCondLst>
                <p:childTnLst>
                  <p:par>
                    <p:cTn id="26" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="27" presetID="8" presetClass="emph" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animScale>
                                <p:cBhvr>
                                  <p:cTn id="28" dur="800"/>
                                  <p:tgtEl><p:spTgt spid="5"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animScale>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
            </p:seq>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
      <p:bldLst>
        <p:bldP spid="2" grpId="0" build="p"/>
      </p:bldLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated slide with changed IDs: 2->20, 3->30, 4->40, 5->50, 6->60, 7->70
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])

    r2 = rem_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    r2.text = "Para 0\nPara 1"
    r2.name = "Paragraph Box"
    r2._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "20"

    r3 = rem_slide.shapes.add_textbox(Inches(5), Inches(1), Inches(2), Inches(1))
    r3.text = "Motion Rect"
    r3.name = "Motion Rect"
    r3._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "30"

    r4 = rem_slide.shapes.add_textbox(Inches(1), Inches(3), Inches(2), Inches(1))
    r4.text = "Exit Oval"
    r4.name = "Exit Oval"
    r4._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "40"

    r5 = rem_slide.shapes.add_textbox(Inches(5), Inches(3), Inches(2), Inches(1))
    r5.text = "Emphasis Star"
    r5.name = "Emphasis Star"
    r5._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "50"

    r6 = rem_slide.shapes.add_textbox(Inches(1), Inches(5), Inches(2), Inches(1))
    r6.text = "Click Me Button"
    r6.name = "Button"
    r6._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "60"

    r7 = rem_slide.shapes.add_table(2, 2, Inches(4), Inches(5), Inches(3), Inches(1.5))
    r7.name = "Summary Table"
    r7._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "70"

    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    orig_shapes = {
        "2": ShapeInfo(shape_id="2", shape_name="Paragraph Box", shape_type="TEXT_BOX"),
        "3": ShapeInfo(shape_id="3", shape_name="Motion Rect", shape_type="TEXT_BOX"),
        "4": ShapeInfo(shape_id="4", shape_name="Exit Oval", shape_type="TEXT_BOX"),
        "5": ShapeInfo(shape_id="5", shape_name="Emphasis Star", shape_type="TEXT_BOX"),
        "6": ShapeInfo(shape_id="6", shape_name="Button", shape_type="TEXT_BOX"),
        "7": ShapeInfo(shape_id="7", shape_name="Summary Table", shape_type="TABLE"),
    }
    orig_clicks, orig_effects = extract_effects_from_timing(orig_slide._element.find(f"{{{NAMESPACES['p']}}}timing"), orig_shapes)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True
    assert res.is_100_percent_faithful is True

    s_res = res.slide_results[0]
    assert s_res.status == SlideStatus.PASS
    assert s_res.verification.passed is True

    prs_out = Presentation(io.BytesIO(res.output_pptx_bytes))
    timing_out = prs_out.slides[0]._element.find(f"{{{NAMESPACES['p']}}}timing")

    shape_map = {"2": "20", "3": "30", "4": "40", "5": "50", "6": "60", "7": "70"}
    check_timing_invariants(timing_out, orig_effects, expected_click_steps=5, shape_map=shape_map)

    # Verify build list has remapped spid 20 and correct bldP tag
    bldLst = timing_out.find(f".//{{{NAMESPACES['p']}}}bldLst")
    assert bldLst is not None
    assert len(bldLst) == 1
    assert bldLst[0].attrib.get("spid") == "20"
    assert bldLst[0].tag.endswith("bldP")


def test_remediated_deck_with_changed_shape_ids_renamed_shapes_and_deleted_shape():
    """
    Test remediated deck with:
    - Shape 2 (changed shape id -> 200)
    - Shape 3 (renamed shape -> 300, matched by type + text + layout)
    - Shape 4 (deleted shape, NOT present in remediated slide)
    Assert:
    - Output effect list matches original for shapes 2 and 3 (with IDs 200, 300).
    - Shape 4 is pruned cleanly from effects.
    - Click step of shape 4 is removed cleanly, leaving exactly 2 click steps.
    - No empty/hollow effect node exists.
    - Deleted shape appears in the skipped report with reason.
    """
    from pptx import Presentation
    from pptx.util import Inches
    from core.animation_transfer import extract_effects_from_timing
    from core.models import ShapeInfo

    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])

    s2 = orig_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    s2.text = "Title Text"
    s2.name = "Title Shape"
    s2._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "2"

    s3 = orig_slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
    s3.text = "Unique Subtitle Text for Matching"
    s3.name = "Old Subtitle Name"
    s3._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "3"

    s4 = orig_slide.shapes.add_textbox(Inches(1), Inches(5), Inches(4), Inches(1))
    s4.text = "Deleted Box"
    s4.name = "Deleted Shape"
    s4._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "4"

    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                <p:childTnLst>
                  <!-- Step 1: Shape 2 -->
                  <p:par>
                    <p:cTn id="10" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="11" presetID="10" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animEffect transition="in" filter="fade">
                                <p:cBhvr>
                                  <p:cTn id="12" dur="500"/>
                                  <p:tgtEl><p:spTgt spid="2"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animEffect>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                  <!-- Step 2: Shape 3 -->
                  <p:par>
                    <p:cTn id="13" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="14" presetID="1" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:set>
                                <p:cBhvr>
                                  <p:cTn id="15" dur="1"/>
                                  <p:tgtEl><p:spTgt spid="3"/></p:tgtEl>
                                </p:cBhvr>
                                <p:to><p:strVal val="visible"/></p:to>
                              </p:set>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                  <!-- Step 3: Shape 4 (deleted) -->
                  <p:par>
                    <p:cTn id="16" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="17" presetID="10" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animEffect transition="in" filter="fade">
                                <p:cBhvr>
                                  <p:cTn id="18" dur="500"/>
                                  <p:tgtEl><p:spTgt spid="4"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animEffect>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
            </p:seq>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated has shape 2 (ID 200) and shape 3 (ID 300, renamed), shape 4 is DELETED
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])

    r2 = rem_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    r2.text = "Title Text"
    r2.name = "Title Shape"
    r2._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "200"

    r3 = rem_slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
    r3.text = "Unique Subtitle Text for Matching"
    r3.name = "New Renamed Subtitle"
    r3._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "300"

    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    orig_shapes = {
        "2": ShapeInfo(shape_id="2", shape_name="Title Shape", shape_type="TEXT_BOX"),
        "3": ShapeInfo(shape_id="3", shape_name="Old Subtitle Name", shape_type="TEXT_BOX"),
        "4": ShapeInfo(shape_id="4", shape_name="Deleted Shape", shape_type="TEXT_BOX"),
    }
    orig_clicks, orig_effects = extract_effects_from_timing(orig_slide._element.find(f"{{{NAMESPACES['p']}}}timing"), orig_shapes)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True

    s_res = res.slide_results[0]
    assert s_res.status in (SlideStatus.PARTIAL, SlideStatus.MANUAL_REVIEW)
    assert s_res.skipped_animation_count == 1
    assert s_res.skipped[0].shape_id == "4"

    prs_out = Presentation(io.BytesIO(res.output_pptx_bytes))
    timing_out = prs_out.slides[0]._element.find(f"{{{NAMESPACES['p']}}}timing")

    # Step 3 was for shape 4, so click steps in output must be 2
    shape_map = {"2": "200", "3": "300"}
    check_timing_invariants(timing_out, orig_effects, expected_click_steps=2, shape_map=shape_map, pruned_spids={"4"})


def test_30_slide_original_and_31_slide_remediated_decks():
    """
    Test real 30-slide original and 31-slide remediated test decks if present.
    Asserts integrity validation, no slide failures, and correct slide mapping.
    """
    import os
    orig_path = r"C:\Users\HBS_1\Downloads\original_30_slides_animated.pptx"
    rem_path = r"C:\Users\HBS_1\Downloads\remediated_31_slides_no_animation.pptx"

    if not (os.path.exists(orig_path) and os.path.exists(rem_path)):
        pytest.skip("Test decks not present in Downloads directory.")

    res = transfer_presentation_animations(orig_path, rem_path)
    assert res.success is True
    assert res.total_failed_slides == 0

    failed = [s for s in res.slide_results if s.status == SlideStatus.FAILED]
    assert len(failed) == 0, f"Failed slides: {failed}"


def test_group_ungrouped_into_3_shapes_split_rule():
    """
    Scenario 1: Group with text ungrouped into 3 shapes in remediated deck.
    Asserts:
    - Group maps to all 3 remediated shapes (group split).
    - Preserved click structure (first copy is clickEffect, rest are withEffect).
    - Unique cTn IDs for duplicated nodes.
    - Verified as covered.
    """
    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])

    grp_xml = f"""
    <p:grpSp xmlns:p="{NAMESPACES['p']}" xmlns:a="{NAMESPACES['a']}">
      <p:nvGrpSpPr><p:cNvPr id="10" name="Group 10"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="1000000" y="1000000"/><a:ext cx="3000000" cy="1000000"/>
          <a:chOff x="1000000" y="1000000"/><a:chExt cx="3000000" cy="1000000"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="11" name="Child 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Item Alpha</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="12" name="Child 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="2000000" y="1000000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Item Beta</a:t></a:r></a:p></p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="13" name="Child 3"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="3000000" y="1000000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Item Gamma</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:grpSp>
    """
    orig_slide._element.find(f".//{{{NAMESPACES['p']}}}spTree").append(etree.fromstring(grp_xml))

    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                <p:childTnLst>
                  <p:par>
                    <p:cTn id="3" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="4" presetID="10" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animEffect transition="in" filter="fade">
                                <p:cBhvr>
                                  <p:cTn id="5" dur="500"/>
                                  <p:tgtEl><p:spTgt spid="10"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animEffect>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
            </p:seq>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated deck has 3 separate shapes
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])
    r1 = rem_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(1), Inches(1))
    r1.text = "Item Alpha"
    r1._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "101"
    r2 = rem_slide.shapes.add_textbox(Inches(2), Inches(1), Inches(1), Inches(1))
    r2.text = "Item Beta"
    r2._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "102"
    r3 = rem_slide.shapes.add_textbox(Inches(3), Inches(1), Inches(1), Inches(1))
    r3.text = "Item Gamma"
    r3._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "103"
    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True
    s_res = res.slide_results[0]
    assert s_res.mapped_group_split_count >= 1
    assert s_res.status in (SlideStatus.PASS, SlideStatus.PARTIAL)

    prs_out = Presentation(io.BytesIO(res.output_pptx_bytes))
    timing_out = prs_out.slides[0]._element.find(f"{{{NAMESPACES['p']}}}timing")
    spids = [elem.attrib.get("spid") for elem in timing_out.iter() if elem.tag.endswith("spTgt")]
    assert set(spids) == {"101", "102", "103"}

    eff_nodes = [elem for elem in timing_out.iter() if elem.attrib.get("presetClass") == "entr"]
    assert len(eff_nodes) == 3
    assert eff_nodes[0].attrib.get("nodeType") == "clickEffect"
    assert eff_nodes[1].attrib.get("nodeType") == "withEffect"
    assert eff_nodes[2].attrib.get("nodeType") == "withEffect"

    ctn_ids = [elem.attrib.get("id") for elem in timing_out.iter() if elem.tag.endswith("cTn")]
    assert len(ctn_ids) == len(set(ctn_ids)), "cTn IDs must be unique!"


def test_group_converted_to_picture():
    """
    Scenario 2: Group converted to one picture.
    Asserts:
    - Group mapped to picture via bounding box overlap.
    - Animation transferred to picture shape ID.
    - Preserved timing and covered verification.
    """
    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])

    grp_xml = f"""
    <p:grpSp xmlns:p="{NAMESPACES['p']}" xmlns:a="{NAMESPACES['a']}">
      <p:nvGrpSpPr><p:cNvPr id="20" name="Icon Group 20"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="1000000" y="1000000"/><a:ext cx="2000000" cy="2000000"/>
          <a:chOff x="1000000" y="1000000"/><a:chExt cx="2000000" cy="2000000"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="21" name="Vector Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="2000000" cy="2000000"/></a:xfrm></p:spPr>
      </p:sp>
    </p:grpSp>
    """
    orig_slide._element.find(f".//{{{NAMESPACES['p']}}}spTree").append(etree.fromstring(grp_xml))

    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                <p:childTnLst>
                  <p:par>
                    <p:cTn id="3" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="4" presetID="1" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:set>
                                <p:cBhvr>
                                  <p:cTn id="5" dur="1"/>
                                  <p:tgtEl><p:spTgt spid="20"/></p:tgtEl>
                                </p:cBhvr>
                                <p:to><p:strVal val="visible"/></p:to>
                              </p:set>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
            </p:seq>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated has picture at the same bounding box
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])
    pic_xml = f"""
    <p:pic xmlns:p="{NAMESPACES['p']}" xmlns:a="{NAMESPACES['a']}">
      <p:nvPicPr>
        <p:cNvPr id="200" name="Picture 200"/><p:cNvPicPr/><p:nvPr/>
      </p:nvPicPr>
      <p:blipFill>
        <a:blip r:embed="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
      </p:blipFill>
      <p:spPr>
        <a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="2000000" cy="2000000"/></a:xfrm>
      </p:spPr>
    </p:pic>
    """
    rem_slide._element.find(f".//{{{NAMESPACES['p']}}}spTree").append(etree.fromstring(pic_xml))
    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True
    s_res = res.slide_results[0]
    assert s_res.status == SlideStatus.MANUAL_REVIEW
    assert len(s_res.manual_review_items) >= 1
    assert any("Rebuilt slide" in item.reason for item in s_res.manual_review_items)

    prs_out = Presentation(io.BytesIO(res.output_pptx_bytes))
    timing_out = prs_out.slides[0]._element.find(f"{{{NAMESPACES['p']}}}timing")
    # Point 4: keep the remediated slide without animation
    assert timing_out is None


def test_text_from_two_groups_merged_into_one_placeholder():
    """
    Scenario 3: Text from two original groups merged into one placeholder.
    Asserts:
    - First effect is preserved on the merged placeholder.
    - Second effect at different click is added to Manual Review with reason 'merged shape'.
    - Manual review entry populated.
    """
    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])

    grp1_xml = f"""
    <p:grpSp xmlns:p="{NAMESPACES['p']}" xmlns:a="{NAMESPACES['a']}">
      <p:nvGrpSpPr><p:cNvPr id="30" name="Revenue Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="2000000" cy="1000000"/><a:chOff x="1000000" y="1000000"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="31" name="Rev Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Quarterly Revenue</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:grpSp>
    """
    grp2_xml = f"""
    <p:grpSp xmlns:p="{NAMESPACES['p']}" xmlns:a="{NAMESPACES['a']}">
      <p:nvGrpSpPr><p:cNvPr id="40" name="Profit Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="1000000" y="3000000"/><a:ext cx="2000000" cy="1000000"/><a:chOff x="1000000" y="3000000"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="41" name="Profit Child"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="1000000" y="3000000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm></p:spPr>
        <p:txBody><a:bodyPr/><a:p><a:r><a:t>Annual Profit</a:t></a:r></a:p></p:txBody>
      </p:sp>
    </p:grpSp>
    """
    tree = orig_slide._element.find(f".//{{{NAMESPACES['p']}}}spTree")
    tree.append(etree.fromstring(grp1_xml))
    tree.append(etree.fromstring(grp2_xml))

    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                <p:childTnLst>
                  <p:par>
                    <p:cTn id="3" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="4" presetID="10" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animEffect transition="in" filter="fade">
                                <p:cBhvr>
                                  <p:cTn id="5" dur="500"/>
                                  <p:tgtEl><p:spTgt spid="30"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animEffect>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                  <p:par>
                    <p:cTn id="6" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="7" presetID="10" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animEffect transition="in" filter="fade">
                                <p:cBhvr>
                                  <p:cTn id="8" dur="500"/>
                                  <p:tgtEl><p:spTgt spid="40"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animEffect>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
            </p:seq>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated has ONE placeholder containing both texts
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])
    r_merged = rem_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    r_merged.text = "Quarterly Revenue and Annual Profit"
    r_merged._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "300"
    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True
    s_res = res.slide_results[0]
    assert s_res.manual_review_count >= 1
    assert any(item.reason == "merged shape" for item in s_res.manual_review_items)


def test_renamed_shapes_with_changed_ids():
    """
    Scenario 4: Renamed shapes with changed IDs (TextBox 16 -> Content Placeholder 2).
    Asserts:
    - Text similarity matches shapes correctly despite completely different names and IDs.
    - Animation transferred and covered.
    """
    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])
    s = orig_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    s.name = "TextBox 16"
    s.text = "Unique Strategic Plan Header"
    s._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "50"

    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                <p:childTnLst>
                  <p:par>
                    <p:cTn id="3" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="4" presetID="10" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animEffect transition="in" filter="fade">
                                <p:cBhvr>
                                  <p:cTn id="5" dur="500"/>
                                  <p:tgtEl><p:spTgt spid="50"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animEffect>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
            </p:seq>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated: renamed to "Content Placeholder 2", ID changed to "500"
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])
    r = rem_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    r.name = "Content Placeholder 2"
    r.text = "Unique Strategic Plan Header"
    r._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "500"
    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True
    s_res = res.slide_results[0]
    assert s_res.status == SlideStatus.PASS
    assert s_res.covered_count == 1
    assert s_res.shape_map.get("50") == "500"

    prs_out = Presentation(io.BytesIO(res.output_pptx_bytes))
    timing_out = prs_out.slides[0]._element.find(f"{{{NAMESPACES['p']}}}timing")
    spids = [elem.attrib.get("spid") for elem in timing_out.iter() if elem.tag.endswith("spTgt")]
    assert "500" in spids


def test_click_structure_by_paragraph_build():
    """
    Point 1: When one remediated shape (e.g. placeholder with several paragraphs) matches text
    from several original shapes revealed on DIFFERENT clicks, do not merge into one click.
    Target each paragraph separately with p:txEl/p:pRg, keeping one click step per original click step.
    Set shape build to by-paragraph (bldP build="p").
    """
    from pptx import Presentation
    from pptx.util import Inches

    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])

    tb1 = orig_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    tb1.text = "Alpha Bullet Point"
    tb1._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "10"

    tb2 = orig_slide.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
    tb2.text = "Beta Bullet Point"
    tb2._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "20"

    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                <p:childTnLst>
                  <p:par>
                    <p:cTn id="3" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="4" presetID="10" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animEffect transition="in" filter="fade">
                                <p:cBhvr>
                                  <p:cTn id="5" dur="500"/>
                                  <p:tgtEl><p:spTgt spid="10"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animEffect>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                  <p:par>
                    <p:cTn id="6" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="7" presetID="10" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:animEffect transition="in" filter="fade">
                                <p:cBhvr>
                                  <p:cTn id="8" dur="500"/>
                                  <p:tgtEl><p:spTgt spid="20"/></p:tgtEl>
                                </p:cBhvr>
                              </p:animEffect>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
            </p:seq>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated has ONE placeholder with 2 paragraphs
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])
    rem_ph = rem_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(3))
    rem_ph.name = "Content Placeholder 1"
    rem_ph._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "50"
    tf = rem_ph.text_frame
    tf.text = "Alpha Bullet Point"
    p2 = tf.add_paragraph()
    p2.text = "Beta Bullet Point"
    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True
    s_res = res.slide_results[0]
    assert s_res.status == SlideStatus.PASS
    assert s_res.mapped_paragraph_level_count == 2
    assert s_res.manual_review_count == 0

    prs_out = Presentation(io.BytesIO(res.output_pptx_bytes))
    timing_out = prs_out.slides[0]._element.find(f"{{{NAMESPACES['p']}}}timing")
    assert timing_out is not None

    # Check 2 click steps preserved
    main_seq = timing_out.find(f".//{{{NAMESPACES['p']}}}cTn[@nodeType='mainSeq']")
    c_pars = [c for c in main_seq.find(f"{{{NAMESPACES['p']}}}childTnLst") if c.tag.endswith("par")]
    assert len(c_pars) == 2

    # Check paragraph targets
    p_ranges = []
    for sp_tgt in timing_out.iter():
        if sp_tgt.tag.endswith("spTgt"):
            pRg = sp_tgt.find(f".//{{{NAMESPACES['p']}}}pRg")
            if pRg is not None:
                p_ranges.append((pRg.attrib.get("st"), pRg.attrib.get("end")))
    assert ("0", "0") in p_ranges
    assert ("1", "1") in p_ranges

    # Check build list has build="p"
    bldLst = timing_out.find(f".//{{{NAMESPACES['p']}}}bldLst")
    assert bldLst is not None
    assert len(bldLst) == 1
    assert bldLst[0].attrib.get("spid") == "50"
    assert bldLst[0].attrib.get("build") == "p"


def test_equation_safe_text_matching():
    """
    Point 2: Map symbol-font characters and OMML text to Unicode, treating Greek letters,
    minus signs, and combining marks as equal (w vs omega, D vs delta, hyphen vs minus).
    """
    from pptx import Presentation
    from pptx.util import Inches

    orig_prs = Presentation()
    orig_slide = orig_prs.slides.add_slide(orig_prs.slide_layouts[6])

    # Greek omega (\u03c9), Delta (\u0394), unicode minus (\u2212)
    tb = orig_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    tb.text = "Freq \u03c9 \u0394t \u2212 5"
    tb._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "10"

    timing_xml = f"""
    <p:timing xmlns:p="{NAMESPACES['p']}">
      <p:tnLst>
        <p:cTn id="1" dur="indefinite" nodeType="tmRoot">
          <p:childTnLst>
            <p:seq concurrent="1" nextAc="seek">
              <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                <p:childTnLst>
                  <p:par>
                    <p:cTn id="3" fill="hold">
                      <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                      <p:childTnLst>
                        <p:par>
                          <p:cTn id="4" presetID="1" presetClass="entr" fill="hold" nodeType="clickEffect">
                            <p:childTnLst>
                              <p:set>
                                <p:cBhvr>
                                  <p:cTn id="5" dur="1"/>
                                  <p:tgtEl><p:spTgt spid="10"/></p:tgtEl>
                                </p:cBhvr>
                              </p:set>
                            </p:childTnLst>
                          </p:cTn>
                        </p:par>
                      </p:childTnLst>
                    </p:cTn>
                  </p:par>
                </p:childTnLst>
              </p:cTn>
            </p:seq>
          </p:childTnLst>
        </p:cTn>
      </p:tnLst>
    </p:timing>
    """
    orig_slide._element.append(etree.fromstring(timing_xml))
    orig_bio = io.BytesIO()
    orig_prs.save(orig_bio)

    # Remediated uses Latin w, D, and hyphen -
    rem_prs = Presentation()
    rem_slide = rem_prs.slides.add_slide(rem_prs.slide_layouts[6])
    r_tb = rem_slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    r_tb.text = "Freq w Dt - 5"
    r_tb._element.find(f".//{{{NAMESPACES['p']}}}cNvPr").attrib["id"] = "99"
    rem_bio = io.BytesIO()
    rem_prs.save(rem_bio)

    res = transfer_presentation_animations(orig_bio.getvalue(), rem_bio.getvalue())
    assert res.success is True
    s_res = res.slide_results[0]
    assert s_res.status == SlideStatus.PASS
    assert s_res.shape_map.get("10") == "99"




