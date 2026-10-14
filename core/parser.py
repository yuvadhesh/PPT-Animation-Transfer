from __future__ import annotations

import hashlib
import io
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.shapes.base import BaseShape

from core.models import BoundingBox, ShapeInfo, SlideData, normalize_text


class PresentationParseError(Exception):
    """Friendly presentation validation/parsing exception."""
    pass


NAMESPACES = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "p14": "http://schemas.microsoft.com/office/powerpoint/2010/main",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
}


def tag_endswith(elem: Any, suffix: str) -> bool:
    """Check if element tag ends with suffix safely."""
    if elem is None or not hasattr(elem, "tag") or not isinstance(elem.tag, str):
        return False
    return elem.tag.endswith(suffix)


def validate_pptx_file(file_source: Union[str, Path, io.BytesIO, bytes], file_label: str = "File", filename: Optional[str] = None) -> None:
    """
    Validate that the file exists and is a valid PPTX file.
    Detects specific failure reasons and raises PresentationParseError with a friendly message:
    - File name starts with "~$": temporary lock file
    - File size 0 or under 1 KB: empty or incomplete upload
    - Magic bytes D0 CF 11 E0: legacy .ppt or password-protected file
    - Missing ppt/ structure or encrypted
    """
    display_name = filename or "presentation.pptx"
    raw_bytes: bytes

    if isinstance(file_source, (str, Path)):
        path = Path(file_source)
        display_name = path.name
        if not path.exists():
            raise PresentationParseError(f"{file_label} does not exist: {display_name}")

        if display_name.startswith("~$"):
            raise PresentationParseError(
                f"{file_label} ('{display_name}'): This is a temporary lock file, not the real presentation."
            )

        try:
            file_size = path.stat().st_size
        except Exception:
            file_size = 0

        if file_size <= 1024:
            raise PresentationParseError(
                f"{file_label} ('{display_name}'): File is empty or the upload was incomplete."
            )

        if path.suffix.lower() not in [".pptx", ".potx", ".pptm"]:
            raise PresentationParseError(
                f"{file_label} ('{display_name}'): File is not a valid PPTX. Please open it in PowerPoint and Save As .pptx."
            )

        try:
            with open(path, "rb") as f:
                raw_bytes = f.read()
        except Exception as e:
            raise PresentationParseError(f"Could not read {file_label} ('{display_name}'): {e}")

    elif isinstance(file_source, io.BytesIO):
        if display_name.startswith("~$"):
            raise PresentationParseError(
                f"{file_label} ('{display_name}'): This is a temporary lock file, not the real presentation."
            )
        raw_bytes = file_source.getvalue()
    elif isinstance(file_source, bytes):
        if display_name.startswith("~$"):
            raise PresentationParseError(
                f"{file_label} ('{display_name}'): This is a temporary lock file, not the real presentation."
            )
        raw_bytes = file_source
    else:
        raise PresentationParseError(f"Unsupported file source for {file_label}.")

    if len(raw_bytes) <= 1024:
        raise PresentationParseError(
            f"{file_label} ('{display_name}'): File is empty or the upload was incomplete."
        )

    # Magic bytes check for legacy PPT or compound file
    if len(raw_bytes) >= 4 and raw_bytes[:4] == b"\xd0\xcf\x11\xe0":
        raise PresentationParseError(
            f"{file_label} ('{display_name}'): This is an old .ppt or a password-protected file. Please re-save as .pptx."
        )

    # ZIP structure check
    try:
        bio = io.BytesIO(raw_bytes)
        with zipfile.ZipFile(bio, "r") as zf:
            names = zf.namelist()
            if not any(n.startswith("ppt/") for n in names):
                raise PresentationParseError(
                    f"{file_label} ('{display_name}'): File is not a valid PPTX. Please open it in PowerPoint and Save As .pptx."
                )
            for info in zf.infolist():
                if info.flag_bits & 0x1:
                    raise PresentationParseError(
                        f"{file_label} ('{display_name}'): File is password-protected or encrypted. Please remove encryption and save as standard .pptx."
                    )
    except zipfile.BadZipFile:
        raise PresentationParseError(
            f"{file_label} ('{display_name}'): File is corrupted or not a valid ZIP archive."
        )
    except PresentationParseError:
        raise
    except Exception as e:
        raise PresentationParseError(f"Could not inspect {file_label} archive ('{display_name}'): {e}")

    # python-pptx load check
    try:
        Presentation(io.BytesIO(raw_bytes))
    except Exception as e:
        raise PresentationParseError(
            f"{file_label} ('{display_name}'): python-pptx could not parse the presentation: {e}"
        )



def compute_image_hash(image_bytes: bytes) -> str:
    """Compute an MD5 hash of raw image bytes."""
    return hashlib.md5(image_bytes).hexdigest()


def extract_shape_text(shape: BaseShape) -> str:
    """Extract plain text from any shape, table, or group."""
    texts: List[str] = []
    if shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            p_text = "".join(r.text for r in p.runs).strip()
            if not p_text and p.text:
                p_text = p.text.strip()
            if p_text:
                texts.append(p_text)
    elif shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                c_text = cell.text.strip()
                if c_text:
                    texts.append(c_text)
    return " ".join(texts).strip()


def extract_all_shapes_from_slide(slide) -> Dict[str, ShapeInfo]:
    """
    Traverse slide shape tree recursively (including groups at any depth, placeholders,
    pictures, tables, and connectors).
    Applies group transforms (chOff/chExt to off/ext) to produce absolute slide-coordinate
    bounding boxes.
    For groups, aggregates descendants and combined text.
    Returns mapping shape_id (as string) -> ShapeInfo.
    """
    shapes_map: Dict[str, ShapeInfo] = {}

    # Map python-pptx shapes by id if available for extra metadata (e.g. MSO_SHAPE_TYPE)
    pptx_shapes_by_id: Dict[str, BaseShape] = {}
    try:
        def _index_pptx(s_list):
            for sh in s_list:
                pptx_shapes_by_id[str(sh.shape_id)] = sh
                if sh.shape_type == MSO_SHAPE_TYPE.GROUP and hasattr(sh, "shapes"):
                    _index_pptx(sh.shapes)
        _index_pptx(slide.shapes)
    except Exception:
        pass

    slide_part = getattr(slide, "part", None)
    sld_xml = getattr(slide, "_element", None)
    if sld_xml is None:
        return shapes_map

    spTree = sld_xml.find(f".//{{{NAMESPACES['p']}}}spTree")
    if spTree is None:
        spTree = sld_xml

    def _parse_xml_node(
        node,
        scale_x: float = 1.0,
        scale_y: float = 1.0,
        offset_x: float = 0.0,
        offset_y: float = 0.0,
        parent_id: Optional[str] = None,
    ) -> List[str]:
        """
        Recursively parse shapes in XML node.
        Returns list of shape IDs parsed directly inside this container.
        """
        container_shape_ids: List[str] = []

        for child in node:
            if not isinstance(child.tag, str):
                continue
            tag = child.tag.split("}")[-1]
            if tag not in ("sp", "pic", "graphicFrame", "grpSp", "cxnSp"):
                continue

            # Find cNvPr
            cNvPr = child.find(f".//{{{NAMESPACES['p']}}}cNvPr")
            if cNvPr is None:
                continue

            raw_id = cNvPr.attrib.get("id")
            if not raw_id:
                continue
            sid = str(raw_id)
            sname = cNvPr.attrib.get("name", "")
            container_shape_ids.append(sid)

            # Determine shape type
            pptx_s = pptx_shapes_by_id.get(sid)
            is_grp = (tag == "grpSp")
            if pptx_s is not None:
                try:
                    stype = str(pptx_s.shape_type)
                except Exception:
                    stype = "GROUP" if is_grp else tag.upper()
            elif is_grp:
                stype = "GROUP"
            else:
                stype = tag.upper()

            # Locate transform element
            xfrm = None
            if is_grp:
                grpPr = child.find(f"{{{NAMESPACES['p']}}}grpSpPr")
                if grpPr is not None:
                    xfrm = grpPr.find(f"{{{NAMESPACES['a']}}}xfrm")
            elif tag == "graphicFrame":
                xfrm = child.find(f"{{{NAMESPACES['p']}}}xfrm")
            else:
                spPr = child.find(f"{{{NAMESPACES['p']}}}spPr")
                if spPr is None:
                    spPr = child.find(f"{{{NAMESPACES['p']}}}picPr")
                if spPr is not None:
                    xfrm = spPr.find(f"{{{NAMESPACES['a']}}}xfrm")

            off = xfrm.find(f"{{{NAMESPACES['a']}}}off") if xfrm is not None else None
            ext = xfrm.find(f"{{{NAMESPACES['a']}}}ext") if xfrm is not None else None
            x = int(off.attrib.get("x", 0)) if off is not None else 0
            y = int(off.attrib.get("y", 0)) if off is not None else 0
            cx = int(ext.attrib.get("cx", 0)) if ext is not None else 0
            cy = int(ext.attrib.get("cy", 0)) if ext is not None else 0

            # Absolute coordinates in slide space
            abs_x = int(round(offset_x + scale_x * x))
            abs_y = int(round(offset_y + scale_y * y))
            abs_cx = int(round(scale_x * cx))
            abs_cy = int(round(scale_y * cy))
            bbox = BoundingBox(x=abs_x, y=abs_y, cx=abs_cx, cy=abs_cy)

            # Extract text and paragraphs
            paragraphs: List[str] = []
            norm_paragraphs: List[str] = []
            if not is_grp:
                # Find txBody in drawingml or presentationml namespace
                txBody = child.find(f".//{{{NAMESPACES['p']}}}txBody")
                if txBody is None:
                    txBody = child.find(f".//{{{NAMESPACES['a']}}}txBody")

                if txBody is not None:
                    p_elems = txBody.findall(f"{{{NAMESPACES['a']}}}p")
                    for p_elem in p_elems:
                        p_t = []
                        for sub_t in p_elem.iter():
                            if tag_endswith(sub_t, "t") and sub_t.text and sub_t.text.strip():
                                p_t.append(sub_t.text.strip())
                        if p_t:
                            p_str = " ".join(p_t).strip()
                            paragraphs.append(p_str)
                            norm_paragraphs.append(normalize_text(p_str))

                # Collect all text including OMML m:t and drawingml a:t
                all_t = []
                for sub_t in child.iter():
                    if tag_endswith(sub_t, "t") and sub_t.text and sub_t.text.strip():
                        all_t.append(sub_t.text.strip())
                text = " ".join(all_t).strip()
            else:
                text = ""
            norm_text = normalize_text(text)

            # Image hash for pictures
            img_hash = None
            if tag == "pic":
                blip = child.find(f".//{{{NAMESPACES['a']}}}blip")
                if blip is not None:
                    r_embed = blip.attrib.get(f"{{{NAMESPACES['r']}}}embed")
                    if r_embed and slide_part is not None:
                        try:
                            rel = slide_part.rels.get(r_embed)
                            if rel and hasattr(rel, "target_part") and hasattr(rel.target_part, "blob"):
                                img_hash = compute_image_hash(rel.target_part.blob)
                        except Exception:
                            pass
                if not img_hash and pptx_s and hasattr(pptx_s, "image") and pptx_s.image is not None:
                    try:
                        img_hash = compute_image_hash(pptx_s.image.blob)
                    except Exception:
                        pass

            sinfo = ShapeInfo(
                shape_id=sid,
                shape_name=sname,
                shape_type=stype,
                text=text,
                normalized_text=norm_text,
                paragraphs=paragraphs,
                normalized_paragraphs=norm_paragraphs,
                left=abs_x,
                top=abs_y,
                width=abs_cx,
                height=abs_cy,
                bbox=bbox,
                image_hash=img_hash,
                is_group=is_grp,
                parent_id=parent_id,
                xml_element=child,
            )
            shapes_map[sid] = sinfo

            if is_grp:
                chOff = xfrm.find(f"{{{NAMESPACES['a']}}}chOff") if xfrm is not None else None
                chExt = xfrm.find(f"{{{NAMESPACES['a']}}}chExt") if xfrm is not None else None
                ch_x = int(chOff.attrib.get("x", 0)) if chOff is not None else 0
                ch_y = int(chOff.attrib.get("y", 0)) if chOff is not None else 0
                ch_cx = int(chExt.attrib.get("cx", cx)) if chExt is not None else cx
                ch_cy = int(chExt.attrib.get("cy", cy)) if chExt is not None else cy

                local_scale_x = (cx / ch_cx) if ch_cx != 0 else 1.0
                local_scale_y = (cy / ch_cy) if ch_cy != 0 else 1.0

                next_scale_x = scale_x * local_scale_x
                next_scale_y = scale_y * local_scale_y
                next_offset_x = offset_x + scale_x * (x - ch_x * local_scale_x)
                next_offset_y = offset_y + scale_y * (y - ch_y * local_scale_y)

                child_ids = _parse_xml_node(
                    child,
                    scale_x=next_scale_x,
                    scale_y=next_scale_y,
                    offset_x=next_offset_x,
                    offset_y=next_offset_y,
                    parent_id=sid,
                )

                # Collect all descendants and combined text
                all_descendants: List[str] = []
                combined_texts: List[str] = []
                for cid in child_ids:
                    all_descendants.append(cid)
                    c_info = shapes_map.get(cid)
                    if c_info:
                        all_descendants.extend(c_info.descendant_ids)
                        if c_info.text:
                            combined_texts.append(c_info.text)
                        elif c_info.combined_text:
                            combined_texts.append(c_info.combined_text)

                sinfo.descendant_ids = all_descendants
                c_text = " ".join([t for t in combined_texts if t]).strip()
                sinfo.combined_text = c_text
                sinfo.normalized_combined_text = normalize_text(c_text)
                if not sinfo.text and c_text:
                    sinfo.text = c_text
                    sinfo.normalized_text = sinfo.normalized_combined_text

        return container_shape_ids

    _parse_xml_node(spTree, scale_x=1.0, scale_y=1.0, offset_x=0.0, offset_y=0.0, parent_id=None)
    return shapes_map


def extract_slide_data(slide, slide_index: int) -> SlideData:
    """Extract comprehensive metadata, shapes, title, and text for a slide."""
    shapes = extract_all_shapes_from_slide(slide)

    title = ""
    try:
        if slide.shapes.title and slide.shapes.title.has_text_frame:
            title = slide.shapes.title.text_frame.text.strip()
    except Exception:
        pass

    all_texts: List[str] = []
    if title:
        all_texts.append(title)
    for shp in shapes.values():
        if shp.text and shp.text != title:
            all_texts.append(shp.text)

    # Check for timing and transition in slide XML
    has_timing = False
    has_transition = False
    try:
        sld_xml = slide._element
        if sld_xml.find(f"{{{NAMESPACES['p']}}}timing") is not None:
            has_timing = True
        if sld_xml.find(f"{{{NAMESPACES['p']}}}transition") is not None:
            has_transition = True
        else:
            # Check mc:AlternateContent
            for ac in sld_xml.findall(f"{{{NAMESPACES['mc']}}}AlternateContent"):
                for sub in ac.iter():
                    if isinstance(sub.tag, str) and sub.tag.endswith("transition"):
                        has_transition = True
                        break
    except Exception:
        pass

    return SlideData(
        slide_index=slide_index,
        title=title,
        all_text=" ".join(all_texts).strip(),
        shapes=shapes,
        has_timing=has_timing,
        has_transition=has_transition,
        slide_part=slide.part,
        xml_root=slide._element,
    )
