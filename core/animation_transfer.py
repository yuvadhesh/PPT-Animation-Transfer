from __future__ import annotations

import copy
from typing import Any, Dict, List, Optional, Set, Tuple

from lxml import etree

from core.models import (
    EffectData,
    EffectType,
    ManualReviewItem,
    ShapeInfo,
    SkippedAnimation,
    SlideData,
    SlideStatus,
    SlideTransferResult,
    TriggerType,
)
from core.parser import NAMESPACES
from core.matcher import find_matching_paragraph_index

# Known PowerPoint preset effect mappings
PRESET_EFFECT_NAMES = {
    "1": "Appear",
    "2": "Fly In",
    "3": "Blinds",
    "4": "Box",
    "5": "Checkerboard",
    "6": "Circle",
    "7": "Crawl",
    "8": "Diamond",
    "9": "Dissolve",
    "10": "Fade",
    "11": "Flash Once",
    "12": "Peek",
    "13": "Plus",
    "14": "Random Bars",
    "15": "Spiral",
    "16": "Split",
    "17": "Stretch",
    "18": "Strips",
    "19": "Swivel",
    "20": "Wipe",
    "21": "Zoom",
    "22": "Grow & Turn",
    "23": "Spin",
    "24": "Grow / Shrink",
    "25": "Desaturate",
    "26": "Darken",
    "27": "Lighten",
    "28": "Transparency",
    "29": "Object Color",
    "30": "Complementary Color",
    "31": "Line Color",
    "32": "Fill Color",
    "33": "Teeter",
    "34": "Font Color",
    "35": "Underline",
    "36": "Bold Flash",
    "37": "Bold Reveal",
    "38": "Wave",
    "39": "Brush Color",
    "40": "Color Wave",
    "41": "Blast",
    "42": "Blink",
    "43": "Float In",
    "44": "Bounce",
    "45": "Wheel",
    "46": "Swish",
    "47": "Compress",
    "48": "Ascend",
    "49": "Descend",
    "50": "Boomerang",
    "51": "Credits",
    "52": "Arc Up",
    "53": "Turn",
    "54": "Loops",
    "55": "Custom Path",
}

PRESET_CLASS_MAP = {
    "entr": EffectType.ENTRANCE,
    "exit": EffectType.EXIT,
    "emph": EffectType.EMPHASIS,
    "path": EffectType.MOTION_PATH,
    "mediacall": EffectType.MEDIA,
}


def get_tag_name(elem: Any) -> str:
    """Return local tag name if element is a valid tag, or empty string if Comment/PI."""
    if elem is None or not hasattr(elem, "tag") or not isinstance(elem.tag, str):
        return ""
    return elem.tag.split("}")[-1]


def tag_endswith(elem: Any, suffix: str) -> bool:
    """Check if element tag ends with suffix safely."""
    if elem is None or not hasattr(elem, "tag") or not isinstance(elem.tag, str):
        return False
    return elem.tag.endswith(suffix)


def is_effect_node(cTn_elem: etree._Element) -> bool:
    """
    Check if a <p:cTn> element represents an animation effect.
    An effect has presetClass or presetID, or nodeType in ('clickEffect', 'withEffect', 'afterEffect').
    """
    if get_tag_name(cTn_elem) != "cTn":
        return False
    node_type = cTn_elem.attrib.get("nodeType", "")
    preset_id = cTn_elem.attrib.get("presetID")
    preset_class = cTn_elem.attrib.get("presetClass")
    return bool(preset_class or preset_id or node_type in ("clickEffect", "withEffect", "afterEffect"))


def extract_effects_from_timing(
    timing_elem: Optional[etree._Element],
    shapes_dict: Dict[str, ShapeInfo],
) -> Tuple[int, List[EffectData]]:
    """
    Inspect timing XML and extract:
    1. click_steps_count
    2. list of EffectData describing every animation effect accurately
    """
    if timing_elem is None:
        return 0, []

    # 1. Count Click Steps in mainSeq
    click_steps_count = 0
    main_seq_ctn = None
    for elem in timing_elem.iter():
        if get_tag_name(elem) == "cTn" and elem.attrib.get("nodeType") == "mainSeq":
            main_seq_ctn = elem
            break

    if main_seq_ctn is not None:
        child_tn_lst = main_seq_ctn.find(f"{{{NAMESPACES['p']}}}childTnLst")
        if child_tn_lst is not None:
            # Each direct <p:par> child of mainSeq/childTnLst is a click step
            click_steps_count = sum(
                1 for c in child_tn_lst
                if isinstance(c.tag, str) and get_tag_name(c) == "par"
            )

    # 2. Extract Effect Nodes
    effects: List[EffectData] = []
    order = 1

    for elem in timing_elem.iter():
        if not is_effect_node(elem):
            continue

        preset_class = elem.attrib.get("presetClass", "")
        preset_id = elem.attrib.get("presetID", "")
        preset_subtype = elem.attrib.get("presetSubtype", "0")
        node_type = elem.attrib.get("nodeType", "")
        dur_val = elem.attrib.get("dur", "0")

        # Determine delay
        delay_val = "0"
        cond = elem.find(f".//{{{NAMESPACES['p']}}}cond")
        if cond is not None:
            delay_val = cond.attrib.get("delay", "0")

        # Find target spid and paragraph range
        spid = None
        p_range = None
        for sub in elem.iter():
            if tag_endswith(sub, "spTgt"):
                spid = sub.attrib.get("spid")
                pRg = sub.find(f".//{{{NAMESPACES['p']}}}pRg")
                if pRg is not None:
                    p_range = f"Para {pRg.attrib.get('st', '0')}-{pRg.attrib.get('end', '0')}"
                if spid:
                    break

        if not spid:
            continue

        # Determine effect name
        eff_name = "Animation"
        if preset_id and preset_id in PRESET_EFFECT_NAMES:
            eff_name = PRESET_EFFECT_NAMES[preset_id]
        else:
            for sub in elem.iter():
                flt = sub.attrib.get("filter")
                if flt:
                    eff_name = flt.capitalize()
                    break
                tag = get_tag_name(sub)
                if tag == "animMotion":
                    eff_name = "Motion Path"
                    break
                elif tag == "animRot":
                    eff_name = "Spin"
                    break
                elif tag == "animScale":
                    eff_name = "Grow/Shrink"
                    break
                elif tag == "cmd":
                    eff_name = "Media Action"
                    break

        # Determine effect type
        if preset_class in PRESET_CLASS_MAP:
            eff_type = PRESET_CLASS_MAP[preset_class]
        else:
            eff_type = EffectType.OTHER

        # Determine trigger
        trigger = TriggerType.ON_CLICK
        if node_type == "withEffect":
            trigger = TriggerType.WITH_PREVIOUS
        elif node_type == "afterEffect":
            trigger = TriggerType.AFTER_PREVIOUS
        elif node_type == "interactiveSeq":
            trigger = TriggerType.INTERACTIVE
        elif cond is not None:
            evt = cond.attrib.get("evt")
            if evt == "onClick":
                trigger = TriggerType.ON_CLICK
            elif evt == "onBegin":
                trigger = TriggerType.WITH_PREVIOUS
            elif evt == "onEnd":
                trigger = TriggerType.AFTER_PREVIOUS

        sinfo = shapes_dict.get(str(spid))
        sname = sinfo.shape_name if sinfo else f"Shape {spid}"

        effects.append(
            EffectData(
                order=order,
                target_spid=str(spid),
                shape_name=sname,
                preset_class=preset_class,
                preset_id=preset_id,
                preset_subtype=preset_subtype,
                node_type=node_type,
                delay=str(delay_val),
                dur=str(dur_val),
                paragraph_range=p_range,
                effect_name=eff_name,
                effect_type=eff_type,
                trigger=trigger,
                xml_element=elem,
            )
        )
        order += 1

    return click_steps_count, effects


def rewrite_shape_ids_in_timing(
    timing_root: etree._Element,
    shape_map: Dict[str, Any],
    slide_index: int,
    orig_shapes: Dict[str, ShapeInfo],
    rem_shapes: Optional[Dict[str, ShapeInfo]] = None,
    manual_reviews: Optional[List[ManualReviewItem]] = None,
) -> Tuple[etree._Element, List[SkippedAnimation], List[ManualReviewItem], int]:
    """
    Exact timing tree rewrite:
    1. Clones timing_root with copy.deepcopy.
    2. Identifies and removes ONLY effect nodes whose target shapes are unmapped.
       Never leaves hollow effects (empty childTnLst).
    3. Group split rule: when 1 original effect maps to N remediated shapes:
       - Duplicate effect node (deepcopy) into N copies.
       - 1st copy keeps original nodeType; rest are withEffect in same parent.
       - Assign unique cTn ids to all new nodes.
    4. By-paragraph build rule (Point 1):
       When one remediated shape (e.g. placeholder with several paragraphs) matches text
       from several original shapes revealed on DIFFERENT clicks:
       - Do not merge them into one click.
       - Target each paragraph separately with p:txEl/p:pRg (st="X" end="X").
       - Keep one click step per original click step with same nodeType, delay, dur, preset.
       - Set shape build to by-paragraph (bldP build="p") only when original steps differ.
    5. Conflict fallback: If a shape was matched by 2 different original effects at different clicks
       and cannot be split into paragraphs, keep first and add rest to Manual Review with reason 'merged shape'.
    6. Prunes empty click steps cleanly, keeping click steps if part of effects mapped.
    7. Copies <p:bldLst> from original with remapped spids, adding entries for each
       new split text shape or by-paragraph build shape.
    """
    cleaned_timing = copy.deepcopy(timing_root)
    skipped: List[SkippedAnimation] = []
    reviews: List[ManualReviewItem] = list(manual_reviews) if manual_reviews is not None else []
    rem_shapes_dict = rem_shapes or {}
    mapped_paragraph_level_count = 0
    shapes_needing_p_build: Set[str] = set()

    # Track max cTn id to guarantee uniqueness for new nodes
    existing_ctn_ids: Set[int] = set()
    for elem in cleaned_timing.iter():
        if tag_endswith(elem, "cTn"):
            cid = elem.attrib.get("id")
            if cid and cid.isdigit():
                existing_ctn_ids.add(int(cid))
    current_max_id = max(existing_ctn_ids or [100])

    def get_next_ctn_id() -> str:
        nonlocal current_max_id
        current_max_id += 1
        return str(current_max_id)

    # Locate mainSeq
    main_seq_ctn = None
    for elem in cleaned_timing.iter():
        if get_tag_name(elem) == "cTn" and elem.attrib.get("nodeType") == "mainSeq":
            main_seq_ctn = elem
            break

    # Collect click steps
    click_step_pars: List[etree._Element] = []
    if main_seq_ctn is not None:
        child_tn_lst = main_seq_ctn.find(f"{{{NAMESPACES['p']}}}childTnLst")
        if child_tn_lst is not None:
            click_step_pars = [
                c for c in child_tn_lst
                if isinstance(c.tag, str) and get_tag_name(c) == "par"
            ]

    # Pre-scan click steps to find remediated shapes targeted across multiple clicks by different original shapes
    # rid -> list of (click_idx, orig_spid)
    rid_to_click_origs: Dict[str, List[Tuple[int, str]]] = {}
    for c_idx, p_elem in enumerate(click_step_pars, start=1):
        for eff in p_elem.iter():
            if is_effect_node(eff):
                for sub in eff.iter():
                    if tag_endswith(sub, "spTgt"):
                        osp = sub.attrib.get("spid")
                        if osp and str(osp) in shape_map:
                            m_val = shape_map[str(osp)]
                            targets = [str(m_val)] if isinstance(m_val, (str, int)) else [str(t) for t in m_val]
                            for t_rid in targets:
                                rid_to_click_origs.setdefault(t_rid, []).append((c_idx, str(osp)))
                        break

    # For each rid, check if it matches text from several original shapes:
    # (rid, orig_spid) -> p_idx
    rid_orig_to_pidx: Dict[Tuple[str, str], int] = {}
    for rid, entries in rid_to_click_origs.items():
        distinct_origs = sorted(list({o for _, o in entries}))
        distinct_clicks = {c for c, _ in entries}
        if len(distinct_origs) > 1:
            rshp = rem_shapes_dict.get(rid)
            if rshp and len(rshp.paragraphs) >= 2:
                used_paras: Set[int] = set()
                for o_sp in distinct_origs:
                    o_info = orig_shapes.get(o_sp)
                    if o_info and o_info.text:
                        p_match = find_matching_paragraph_index(o_info.text, rshp.paragraphs, used_indices=used_paras)
                        if p_match is not None:
                            used_paras.add(p_match)
                            rid_orig_to_pidx[(rid, o_sp)] = p_match
                # Fallback for unassigned shapes: assign sequential unused paragraphs by vertical position
                unassigned = [o_sp for o_sp in distinct_origs if (rid, o_sp) not in rid_orig_to_pidx]
                if unassigned:
                    unassigned.sort(
                        key=lambda o: orig_shapes[o].bbox.top if (o in orig_shapes and orig_shapes[o].bbox) else 0
                    )
                    for o_sp in unassigned:
                        for p_i in range(len(rshp.paragraphs)):
                            if p_i not in used_paras:
                                used_paras.add(p_i)
                                rid_orig_to_pidx[(rid, o_sp)] = p_i
                                break
                # Set shape build to by-paragraph (bldP build="p") only when the original steps differ
                if len(distinct_clicks) > 1 and len(rid_orig_to_pidx) >= 2:
                    shapes_needing_p_build.add(rid)

    # Track remediated shape animations across click steps
    # maps rem_shape_id -> (click_step_number, orig_spid)
    seen_rem_shape_at_click: Dict[str, Tuple[int, str]] = {}

    effects_to_prune: List[etree._Element] = []
    pruned_orig_spids: Set[str] = set()
    processed_effects: Set[etree._Element] = set()

    def set_sp_tgt_nodes(eff: etree._Element, old_sp: str, new_sp: str, p_idx: Optional[int] = None):
        for sub in eff.iter():
            if tag_endswith(sub, "spTgt") and sub.attrib.get("spid") == old_sp:
                sub.attrib["spid"] = new_sp
                if p_idx is not None:
                    # Remove any existing txEl
                    for c in list(sub):
                        if tag_endswith(c, "txEl"):
                            sub.remove(c)
                    txEl = etree.SubElement(sub, f"{{{NAMESPACES['p']}}}txEl")
                    pRg = etree.SubElement(txEl, f"{{{NAMESPACES['p']}}}pRg")
                    pRg.attrib["st"] = str(p_idx)
                    pRg.attrib["end"] = str(p_idx)
            elif tag_endswith(sub, "subSp") and sub.attrib.get("id") == old_sp:
                sub.attrib["id"] = new_sp

    def _process_effect_elem(eff_elem: etree._Element, click_step_idx: int):
        nonlocal mapped_paragraph_level_count

        # Find target spid
        spTgt = None
        for sub in eff_elem.iter():
            if tag_endswith(sub, "spTgt"):
                spTgt = sub
                break

        if spTgt is None:
            return

        orig_spid = spTgt.attrib.get("spid")
        if not orig_spid:
            effects_to_prune.append(eff_elem)
            return

        orig_spid_str = str(orig_spid)
        mapped_val = shape_map.get(orig_spid_str)

        if mapped_val is None or (isinstance(mapped_val, list) and len(mapped_val) == 0):
            # Unmapped effect
            effects_to_prune.append(eff_elem)
            pruned_orig_spids.add(orig_spid_str)
            sinfo = orig_shapes.get(orig_spid_str)
            sname = sinfo.shape_name if sinfo else f"Shape {orig_spid_str}"
            p_class = eff_elem.attrib.get("presetClass", "")
            p_id = eff_elem.attrib.get("presetID", "")
            eff_name = PRESET_EFFECT_NAMES.get(p_id, "Animation")
            eff_type = PRESET_CLASS_MAP.get(p_class, EffectType.OTHER).value

            skipped.append(
                SkippedAnimation(
                    slide_index=slide_index,
                    shape_id=orig_spid_str,
                    shape_name=sname,
                    effect_type=eff_type,
                    effect_name=eff_name,
                    reason="Shape was deleted or not found in remediated slide",
                )
            )
            return

        # Normalize targets to list
        if isinstance(mapped_val, (str, int)):
            raw_targets = [str(mapped_val)]
        else:
            raw_targets = [str(t) for t in mapped_val]

        valid_targets: List[str] = []
        for rid in raw_targets:
            # Check if this effect targets a specific paragraph in a multi-paragraph shape
            if (rid, orig_spid_str) in rid_orig_to_pidx:
                seen_rem_shape_at_click[rid] = (click_step_idx, orig_spid_str)
                valid_targets.append(rid)
                continue

            # Check for merged text conflict:
            if rid in seen_rem_shape_at_click:
                prev_click, prev_orig = seen_rem_shape_at_click[rid]
                if prev_orig != orig_spid_str and prev_click != click_step_idx:
                    # Conflict: two different original shapes mapped to this remediated shape at different clicks
                    sinfo = orig_shapes.get(orig_spid_str)
                    sname = sinfo.shape_name if sinfo else f"Shape {orig_spid_str}"
                    p_id = eff_elem.attrib.get("presetID", "")
                    eff_name = PRESET_EFFECT_NAMES.get(p_id, "Animation")
                    rev_item = ManualReviewItem(
                        slide_index=slide_index,
                        shape_name=sname,
                        shape_id=orig_spid_str,
                        text_snippet=sinfo.text[:50] if (sinfo and sinfo.text) else "No text",
                        shape_type=sinfo.shape_type if sinfo else "UNKNOWN",
                        bbox=sinfo.bbox.to_str() if (sinfo and sinfo.bbox) else "Unknown",
                        effect_name=eff_name,
                        trigger="On Click",
                        reason="merged shape",
                    )
                    reviews.append(rev_item)
                    continue

            seen_rem_shape_at_click[rid] = (click_step_idx, orig_spid_str)
            valid_targets.append(rid)

        if not valid_targets:
            effects_to_prune.append(eff_elem)
            return

        # Remap target in the primary effect node
        first_rid = valid_targets[0]
        first_pidx = rid_orig_to_pidx.get((first_rid, orig_spid_str))
        if first_pidx is not None:
            mapped_paragraph_level_count += 1
        set_sp_tgt_nodes(eff_elem, orig_spid_str, first_rid, first_pidx)

        # Group split: if mapped to N shapes (N > 1), create copies for the rest
        if len(valid_targets) > 1:
            parent_container = eff_elem.getparent()
            if parent_container is not None:
                insert_ref = eff_elem
                for rid in valid_targets[1:]:
                    copy_node = copy.deepcopy(eff_elem)
                    copy_node.attrib["id"] = get_next_ctn_id()
                    copy_node.attrib["nodeType"] = "withEffect"

                    # Assign unique IDs to all nested cTn elements
                    for sub_c in copy_node.iter():
                        if tag_endswith(sub_c, "cTn") and sub_c is not copy_node:
                            sub_c.attrib["id"] = get_next_ctn_id()

                    split_pidx = rid_orig_to_pidx.get((rid, orig_spid_str))
                    if split_pidx is not None:
                        mapped_paragraph_level_count += 1
                    set_sp_tgt_nodes(copy_node, first_rid, rid, split_pidx)

                    ref_idx = parent_container.index(insert_ref)
                    parent_container.insert(ref_idx + 1, copy_node)
                    insert_ref = copy_node

                    # Prevent the new copy and its children from being treated as original effects
                    processed_effects.add(copy_node)
                    for sub_e in copy_node.iter():
                        processed_effects.add(sub_e)

    # Process all effects in mainSeq click steps
    for click_idx, par_elem in enumerate(click_step_pars, start=1):
        for elem in list(par_elem.iter()):
            if is_effect_node(elem) and elem not in effects_to_prune and elem not in processed_effects:
                processed_effects.add(elem)
                _process_effect_elem(elem, click_idx)

    # Process any effects outside mainSeq (e.g. interactive sequences)
    for elem in list(cleaned_timing.iter()):
        if is_effect_node(elem) and elem not in effects_to_prune and elem not in processed_effects:
            processed_effects.add(elem)
            _process_effect_elem(elem, 9999)

    # Prune unmapped effects
    for eff_cTn in effects_to_prune:
        parent = eff_cTn.getparent()
        if parent is None:
            continue
        parent_tag = get_tag_name(parent)
        if parent_tag == "par":
            cTn_children = [c for c in parent if get_tag_name(c) == "cTn"]
            if len(cTn_children) <= 1:
                grandparent = parent.getparent()
                if grandparent is not None:
                    grandparent.remove(parent)
                    continue
        parent.remove(eff_cTn)

    # Clean up empty click steps: keep step if at least 1 effect mapped
    if main_seq_ctn is not None:
        child_tn_lst = main_seq_ctn.find(f"{{{NAMESPACES['p']}}}childTnLst")
        if child_tn_lst is not None:
            for step_child in list(child_tn_lst):
                if not (isinstance(step_child.tag, str) and get_tag_name(step_child) == "par"):
                    continue
                remaining_effs = [e for e in step_child.iter() if is_effect_node(e)]
                if len(remaining_effs) == 0:
                    child_tn_lst.remove(step_child)

    # Clean up and remap interactive sequences
    for elem in list(cleaned_timing.iter()):
        if get_tag_name(elem) == "cTn" and elem.attrib.get("nodeType") == "interactiveSeq":
            trigger_spid = None
            st_cond_tgt = None
            stCondLst = elem.find(f"{{{NAMESPACES['p']}}}stCondLst")
            if stCondLst is not None:
                for sub in stCondLst.iter():
                    if tag_endswith(sub, "spTgt"):
                        trigger_spid = sub.attrib.get("spid")
                        st_cond_tgt = sub
                        break
            effs_inside = [e for e in elem.iter() if is_effect_node(e)]
            mapped_trigger = shape_map.get(str(trigger_spid)) if trigger_spid else None
            if not mapped_trigger or len(effs_inside) == 0:
                p_seq = elem.getparent()
                if p_seq is not None:
                    gp = p_seq.getparent()
                    if gp is not None:
                        gp.remove(p_seq)
            else:
                if isinstance(mapped_trigger, list):
                    mapped_trigger_id = str(mapped_trigger[0]) if mapped_trigger else str(trigger_spid)
                else:
                    mapped_trigger_id = str(mapped_trigger)
                if st_cond_tgt is not None:
                    st_cond_tgt.attrib["spid"] = mapped_trigger_id

    # Remove any hollow <p:cTn> that now has an empty childTnLst
    for elem in list(cleaned_timing.iter()):
        if get_tag_name(elem) == "cTn":
            if elem.attrib.get("nodeType") == "tmRoot":
                continue
            child_lst = elem.find(f"{{{NAMESPACES['p']}}}childTnLst")
            if child_lst is not None and len(child_lst) == 0:
                p = elem.getparent()
                if p is not None:
                    p.remove(elem)

    # Build List (<p:bldLst>) handling:
    active_animated_spids: Set[str] = set()
    for elem in cleaned_timing.iter():
        if tag_endswith(elem, "spTgt"):
            sp = elem.attrib.get("spid")
            if sp:
                active_animated_spids.add(str(sp))

    bldLst = cleaned_timing.find(f".//{{{NAMESPACES['p']}}}bldLst")
    if bldLst is not None:
        for bld_child in list(bldLst):
            old_bld_spid = bld_child.attrib.get("spid")
            if not old_bld_spid or str(old_bld_spid) not in shape_map or str(old_bld_spid) in pruned_orig_spids:
                bldLst.remove(bld_child)
                continue

            mapped_rids = shape_map[str(old_bld_spid)]
            if isinstance(mapped_rids, (str, int)):
                rids_list = [str(mapped_rids)]
            else:
                rids_list = [str(r) for r in mapped_rids]

            active_rids = [r for r in rids_list if r in active_animated_spids]
            if not active_rids:
                bldLst.remove(bld_child)
                continue

            # First active shape keeps the existing bld_child
            bld_child.attrib["spid"] = active_rids[0]

            # Additional split shapes get cloned bld_child entries
            for extra_rid in active_rids[1:]:
                copy_bld = copy.deepcopy(bld_child)
                copy_bld.attrib["spid"] = extra_rid
                bldLst.append(copy_bld)

    # For shapes needing by-paragraph build where original steps differed:
    for p_rid in shapes_needing_p_build:
        if p_rid not in active_animated_spids:
            continue
        if bldLst is None:
            bldLst = etree.Element(f"{{{NAMESPACES['p']}}}bldLst")
            tnLst = cleaned_timing.find(f"{{{NAMESPACES['p']}}}tnLst")
            if tnLst is not None:
                cleaned_timing.insert(cleaned_timing.index(tnLst) + 1, bldLst)
            else:
                cleaned_timing.append(bldLst)

        existing_bld = None
        for bld_c in bldLst:
            if bld_c.attrib.get("spid") == str(p_rid):
                existing_bld = bld_c
                break
        if existing_bld is not None:
            existing_bld.attrib["build"] = "p"
            existing_bld.attrib["grpId"] = "0"
        else:
            bldP = etree.SubElement(bldLst, f"{{{NAMESPACES['p']}}}bldP")
            bldP.attrib["spid"] = str(p_rid)
            bldP.attrib["grpId"] = "0"
            bldP.attrib["build"] = "p"

    if bldLst is not None and len(bldLst) == 0:
        p = bldLst.getparent()
        if p is not None:
            p.remove(bldLst)

    return cleaned_timing, skipped, reviews, mapped_paragraph_level_count


def copy_media_relationships(
    orig_slide_part: Any,
    rem_slide_part: Any,
) -> Tuple[int, List[str]]:
    """
    Copy audio / video / media relationship parts from original slide to remediated slide part.
    Returns (copied_count, errors_list).
    """
    copied_count = 0
    errors: List[str] = []

    if orig_slide_part is None or rem_slide_part is None:
        return 0, []

    try:
        for r_id, rel in orig_slide_part.rels.items():
            reltype = rel.reltype
            if any(k in reltype.lower() for k in ("audio", "video", "media")):
                try:
                    target_part = rel.target_part
                    already_exists = False
                    for existing_rel in rem_slide_part.rels.values():
                        if existing_rel.target_part == target_part:
                            already_exists = True
                            break

                    if not already_exists:
                        rem_slide_part.relate_to(target_part, reltype)
                        copied_count += 1
                except Exception as e:
                    errors.append(f"Could not copy media part {r_id}: {e}")
    except Exception as e:
        errors.append(f"Media extraction error: {e}")

    return copied_count, errors


def transfer_slide_animations_and_transition(
    orig_slide_data: SlideData,
    rem_slide_data: SlideData,
    shape_map: Any,
    copy_transitions: bool = True,
) -> SlideTransferResult:
    """
    Perform animation and transition transfer for a single slide pair.
    """
    res = SlideTransferResult(
        slide_index=orig_slide_data.slide_index,
        remediated_slide_index=rem_slide_data.slide_index,
        shape_map=shape_map,
    )

    orig_xml = orig_slide_data.xml_root
    rem_xml = rem_slide_data.xml_root

    if orig_xml is None or rem_xml is None:
        res.status = SlideStatus.FAILED
        res.error_message = "Slide XML not found."
        return res

    # 1. Extract original effects for reporting & verification
    orig_timing = orig_xml.find(f"{{{NAMESPACES['p']}}}timing")
    orig_click_steps, orig_effects = extract_effects_from_timing(orig_timing, orig_slide_data.shapes)
    res.original_animation_count = len(orig_effects)

    # Record direct vs group split stats from shape_map
    if hasattr(shape_map, "direct_matches"):
        res.mapped_direct_count = len([oid for oid in shape_map.direct_matches if shape_map.direct_matches[oid]])
    if hasattr(shape_map, "group_splits"):
        res.mapped_group_split_count = len(shape_map.group_splits)
    if hasattr(shape_map, "manual_reviews"):
        res.manual_review_items.extend(shape_map.manual_reviews)

    # 2. Copy transition if requested
    if copy_transitions:
        orig_transition = orig_xml.find(f"{{{NAMESPACES['p']}}}transition")
        orig_ac_transition: Optional[etree._Element] = None
        for ac in orig_xml.findall(f"{{{NAMESPACES['mc']}}}AlternateContent"):
            for sub in ac.iter():
                if tag_endswith(sub, "transition"):
                    orig_ac_transition = ac
                    break

        trans_to_copy = orig_ac_transition if orig_ac_transition is not None else orig_transition

        if trans_to_copy is not None:
            existing_trans = rem_xml.find(f"{{{NAMESPACES['p']}}}transition")
            if existing_trans is not None:
                rem_xml.remove(existing_trans)
            for ac in list(rem_xml.findall(f"{{{NAMESPACES['mc']}}}AlternateContent")):
                for sub in ac.iter():
                    if tag_endswith(sub, "transition"):
                        rem_xml.remove(ac)
                        break

            insert_idx = len(rem_xml)
            for idx, child in enumerate(rem_xml):
                tag = get_tag_name(child)
                if tag in ("timing", "extLst"):
                    insert_idx = idx
                    break

            new_trans = copy.deepcopy(trans_to_copy)
            rem_xml.insert(insert_idx, new_trans)
            res.copied_transition = True

    # 3. Copy and rewrite <p:timing>
    if orig_timing is not None:
        copy_media_relationships(orig_slide_data.slide_part, rem_slide_data.slide_part)

        cleaned_timing, skipped_list, updated_reviews, p_mapped_count = rewrite_shape_ids_in_timing(
            orig_timing,
            shape_map,
            orig_slide_data.slide_index,
            orig_slide_data.shapes,
            rem_shapes=rem_slide_data.shapes,
            manual_reviews=res.manual_review_items,
        )
        res.skipped = skipped_list
        res.skipped_animation_count = len(skipped_list)
        res.manual_review_items = updated_reviews
        res.manual_review_count = len(updated_reviews)
        res.mapped_paragraph_level_count = p_mapped_count

        # Extract remaining effects in cleaned_timing
        rem_click_steps, rem_effects = extract_effects_from_timing(cleaned_timing, rem_slide_data.shapes)

        if len(rem_effects) > 0:
            # Remove existing timing in rem_xml if any
            existing_timing = rem_xml.find(f"{{{NAMESPACES['p']}}}timing")
            if existing_timing is not None:
                rem_xml.remove(existing_timing)

            insert_idx = len(rem_xml)
            for idx, child in enumerate(rem_xml):
                tag = get_tag_name(child)
                if tag == "extLst":
                    insert_idx = idx
                    break

            rem_xml.insert(insert_idx, cleaned_timing)
            res.animations = rem_effects
            res.copied_animation_count = len(rem_effects)
        else:
            # Requirement 4: If a slide has animations but none can be mapped (e.g. rebuilt slide),
            # leave the remediated slide unchanged (without animation) and list slide as Manual review required
            existing_timing = rem_xml.find(f"{{{NAMESPACES['p']}}}timing")
            if existing_timing is not None:
                rem_xml.remove(existing_timing)
            res.copied_animation_count = 0
            res.animations = []
            if res.original_animation_count > 0:
                res.status = SlideStatus.MANUAL_REVIEW
                if not res.manual_review_items:
                    res.manual_review_items.append(
                        ManualReviewItem(
                            slide_index=orig_slide_data.slide_index,
                            shape_name="All shapes",
                            shape_id="",
                            text_snippet="Slide rebuilt or converted",
                            shape_type="SLIDE",
                            bbox="",
                            effect_name="All animations",
                            trigger="",
                            reason="Slide rebuilt or no animations could be mapped",
                        )
                    )
                res.manual_review_count = len(res.manual_review_items)
                return res

        # Determine slide status:
        # Do not mark a slide Pass if any effect is in Manual review.
        if res.manual_review_count > 0:
            res.status = SlideStatus.MANUAL_REVIEW
        elif res.original_animation_count == 0:
            res.status = SlideStatus.PASS
        elif res.copied_animation_count == 0:
            res.status = SlideStatus.MANUAL_REVIEW
        elif res.skipped_animation_count > 0:
            res.status = SlideStatus.PARTIAL
        else:
            res.status = SlideStatus.PASS
    else:
        # Original slide had no animations
        res.status = SlideStatus.PASS
        res.original_animation_count = 0
        res.copied_animation_count = 0
        res.skipped_animation_count = 0

    return res
