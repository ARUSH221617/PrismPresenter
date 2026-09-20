import json
import logging
import copy
from pathlib import Path
from typing import Dict, Any, List, Optional, Callable, Tuple, Sequence
from pptx import Presentation
from pptx.util import Pt, Inches, Length
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE, MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from openai import OpenAI

from pptx_jahat.config import Config, DATA_DIR, OUTPUT_DIR
from pptx_jahat.tools.json_parser import safe_json_loads
from pptx_jahat.tools.image_gen import generate_image
from pptx_jahat.tools.preview import (
    render_pptx_file_previews,
    render_pptx_slide_to_image,
    image_to_base64_jpeg
)
from pptx_jahat.tools.human_touch import inspect_pptx_for_editing
from pptx_jahat.tools.pptx_builder import (
    _safe_update_text_frame,
    _replace_image_in_shape,
    _remove_shapes,
    _set_paragraph_rtl_and_fonts,
    _set_run_rtl_and_fonts,
    _is_rtl_text
)

logger = logging.getLogger("slide_verifier")


# ---------------------------------------------------------------------------
# 0. GEOMETRY, COLOR & SEMANTIC LAYOUT HELPERS (FULL ACTION EXECUTOR ACCESS)
# ---------------------------------------------------------------------------
def _parse_coordinate(val: Any, base_emu: Any = 12192000) -> Optional[int]:
    """Converts pt, in, px, %, float ratio or numeric value to EMU."""
    if val is None:
        return None
    effective_base = int(base_emu) if base_emu is not None else 12192000
    if isinstance(val, str):
        v = val.strip().lower()
        if v.endswith("%"):
            try:
                return int((float(v[:-1]) / 100.0) * effective_base)
            except ValueError:
                return None
        if v.endswith("pt"):
            try:
                return int(Pt(float(v[:-2])))
            except ValueError:
                return None
        if v.endswith("in") or v.endswith("inch") or v.endswith("inches"):
            num = v.rstrip("inches").rstrip("inch").rstrip("in")
            try:
                return int(Inches(float(num)))
            except ValueError:
                return None
        if v.endswith("px"):
            try:
                return int(Pt(float(v[:-2]) * 0.75))
            except ValueError:
                return None
        try:
            val = float(v)
        except ValueError:
            return None

    if isinstance(val, (int, float)):
        if 0.0 < float(val) <= 1.0:
            return int(float(val) * effective_base)
        if float(val) < 5000:
            return int(Pt(float(val)))
        return int(val)
    return None


def _parse_hex_color(val: Any) -> Optional[RGBColor]:
    """Parses hex color string (#RRGGBB or RRGGBB) to RGBColor."""
    if not val:
        return None
    s = str(val).strip().lstrip("#")
    if len(s) == 6:
        try:
            return RGBColor(int(s[0:2], 16), int(s[2:4], 16), int(s[4:6], 16))
        except ValueError:
            pass
    return None


def _resolve_shape_type(type_name: Optional[str]) -> Any:
    """Maps human-readable shape type names to MSO_SHAPE enum."""
    if not type_name:
        return MSO_SHAPE.RECTANGLE
    t = str(type_name).strip().upper().replace(" ", "_").replace("-", "_")
    mapping = {
        "RECTANGLE": MSO_SHAPE.RECTANGLE,
        "RECT": MSO_SHAPE.RECTANGLE,
        "ROUNDED_RECTANGLE": MSO_SHAPE.ROUNDED_RECTANGLE,
        "ROUNDED_RECT": MSO_SHAPE.ROUNDED_RECTANGLE,
        "CARD": MSO_SHAPE.ROUNDED_RECTANGLE,
        "BADGE": MSO_SHAPE.ROUNDED_RECTANGLE,
        "TAG": MSO_SHAPE.ROUNDED_RECTANGLE,
        "OVAL": MSO_SHAPE.OVAL,
        "CIRCLE": MSO_SHAPE.OVAL,
        "CALLOUT": MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
    }
    return mapping.get(t, MSO_SHAPE.ROUNDED_RECTANGLE)


def _clone_shape_to_slide(
    src_slide: Any,
    tgt_slide: Any,
    shape_idx: int,
    tgt_prs: Any,
    src_prs: Any
) -> Optional[Any]:
    """
    Clones a shape with full XML geometry, text, styling, and relationships
    from a source template slide into a target slide.
    """
    if shape_idx < 0 or shape_idx >= len(src_slide.shapes):
        return None
    src_shape: Any = src_slide.shapes[shape_idx]
    new_sp_elem = copy.deepcopy(src_shape._element)
    r_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    # Map and copy relationship targets (images, icons, etc.)
    for rId in new_sp_elem.xpath(".//@r:id", namespaces={"r": r_ns}):
        try:
            if rId in src_slide.part.rels:
                rel = src_slide.part.rels[rId]
                new_rId = tgt_slide.part.relate_to(rel.target_part, rel.reltype)
                for elem in new_sp_elem.xpath(f".//*[@r:id='{rId}']", namespaces={"r": r_ns}):
                    elem.set(f"{{{r_ns}}}id", new_rId)
        except Exception:
            pass

    tgt_slide.shapes._spTree.append(new_sp_elem)
    return tgt_slide.shapes[-1]


def extract_slide_semantic_layout(
    slide: Any,
    slide_width_emu: Any = 12192000,
    slide_height_emu: Any = 6858000
) -> str:
    """
    Produces a compact, token-efficient Semantic XML layout representation of a slide
    including shape indices, positions (pt & %), bounding dimensions, typography, fills,
    and text content for high-precision visual verification reasoning.
    """
    w_emu = int(slide_width_emu) if slide_width_emu is not None else 12192000
    h_emu = int(slide_height_emu) if slide_height_emu is not None else 6858000
    w_pt = round(w_emu / 12700, 1)
    h_pt = round(h_emu / 12700, 1)
    lines = [f'<slide width="{w_pt}pt" height="{h_pt}pt">']

    for sh_idx, shape_item in enumerate(slide.shapes):
        shape: Any = shape_item
        sh_type = "UNKNOWN"
        try:
            sh_type = str(shape.shape_type).split(".")[-1]
        except Exception:
            pass

        x_pt = round(getattr(shape, "left", 0) / 12700, 1)
        y_pt = round(getattr(shape, "top", 0) / 12700, 1)
        w_shape_pt = round(getattr(shape, "width", 0) / 12700, 1)
        h_shape_pt = round(getattr(shape, "height", 0) / 12700, 1)

        x_pct = round((getattr(shape, "left", 0) / w_emu) * 100, 1) if w_emu else 0
        y_pct = round((getattr(shape, "top", 0) / h_emu) * 100, 1) if h_emu else 0

        # Style attributes
        style_attrs = []
        try:
            if shape.fill and shape.fill.type == 1:
                col = shape.fill.fore_color.rgb
                style_attrs.append(f'fill="#{col[0]:02x}{col[1]:02x}{col[2]:02x}"')
        except Exception:
            pass

        try:
            if shape.line and shape.line.fill.type is not None:
                l_col = shape.line.color.rgb
                style_attrs.append(f'border="#{l_col[0]:02x}{l_col[1]:02x}{l_col[2]:02x}"')
        except Exception:
            pass

        style_str = " " + " ".join(style_attrs) if style_attrs else ""

        is_ph = getattr(shape, "is_placeholder", False)
        ph_str = ' is_placeholder="true"' if is_ph else ''

        text_content = ""
        text_attrs = []
        if getattr(shape, "has_text_frame", False):
            text_content = shape.text_frame.text.strip()
            paragraphs = shape.text_frame.paragraphs
            if paragraphs:
                first_p = paragraphs[0]
                if first_p.font and first_p.font.size:
                    text_attrs.append(f'size="{round(first_p.font.size.pt, 1)}pt"')
                if first_p.font and first_p.font.name:
                    text_attrs.append(f'font="{first_p.font.name}"')
                if first_p.runs and first_p.runs[0].font and first_p.runs[0].font.color:
                    try:
                        r_col = first_p.runs[0].font.color.rgb
                        text_attrs.append(f'color="#{r_col[0]:02x}{r_col[1]:02x}{r_col[2]:02x}"')
                    except Exception:
                        pass
                if _is_rtl_text(text_content):
                    text_attrs.append('rtl="true"')

        is_formula = False
        if hasattr(shape, "_element"):
            is_formula = any("math" in c.tag.lower() for c in shape._element.iter())
            if not text_content:
                math_texts = [t.text.strip() for t in shape._element.iter() if t.text and t.tag.endswith("}t")]
                if math_texts:
                    text_content = " ".join(math_texts)
                    is_formula = True

        formula_str = ' is_formula="true"' if is_formula else ''
        t_attr_str = " " + " ".join(text_attrs) if text_attrs else ""

        has_table = getattr(shape, "has_table", False)
        table_str = f' is_table="true" rows="{len(shape.table.rows)}" cols="{len(shape.table.columns)}"' if has_table else ''

        char_count = len(text_content)
        area_sq_pt = w_shape_pt * h_shape_pt
        overflow_flag = ' overflow_warning="true"' if (char_count > 250 and area_sq_pt < 12000) else ''

        clean_name = shape.name.replace('"', '&quot;')
        lines.append(
            f'  <shape index="{sh_idx}" name="{clean_name}" type="{sh_type}" '
            f'x="{x_pt}pt ({x_pct}%)" y="{y_pt}pt ({y_pct}%)" '
            f'w="{w_shape_pt}pt" h="{h_shape_pt}pt"{style_str}{ph_str}{formula_str}{table_str}{overflow_flag}>'
        )
        if text_content:
            sample_txt = text_content[:150].replace("<", "&lt;").replace(">", "&gt;").replace("\n", " ")
            lines.append(f'    <text{t_attr_str} chars="{char_count}">{sample_txt}</text>')
        lines.append('  </shape>')

    lines.append('</slide>')
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 1. ACTION APPLICATION: Apply Edit Structure to PowerPoint Presentation
# ---------------------------------------------------------------------------
DELETE_ACTION_NAMES = {
    "remove_shape", "remove_shapes", "delete_shape", "delete_shapes",
    "delete_element", "remove_element", "delete", "remove",
    "clear_shape", "prune_shape", "prune_shapes",
    "remove_formula", "delete_formula", "remove_math", "delete_math"
}


def delete_element_from_slide(
    slide: Any,
    action: Dict[str, Any],
    log_cb: Optional[Callable[[str], None]] = None
) -> int:
    """
    Deletes an element, shape, or formula equation from slide XML.
    Supports:
    - shape_index (int or str): removes shape by index (standard shape or AlternateContent)
    - shape_indices (list of int/str): removes multiple shapes in descending index order
    - formula_text / target_text / text: removes element containing matching text
    - action in ('remove_formula', 'delete_formula', 'remove_math', 'delete_math') or is_formula: removes math formula(s)
    - name / shape_name: removes shape matching name
    Returns number of elements deleted.
    """
    deleted_count = 0
    from pptx_jahat.tools.pptx_builder import _remove_shape

    # 1. Multiple shape indices
    sh_indices = action.get("shape_indices")
    if sh_indices and isinstance(sh_indices, (list, tuple)):
        clean_indices = sorted([int(i) for i in sh_indices if i is not None], reverse=True)
        for idx in clean_indices:
            if _remove_shape(slide, idx):
                deleted_count += 1
        if deleted_count > 0:
            return deleted_count

    # 2. Single shape index
    sh_idx = action.get("shape_index")
    if sh_idx is not None:
        try:
            idx = int(sh_idx)
            if _remove_shape(slide, idx):
                return 1
        except Exception:
            pass

    # 3. Text / formula_text substring matching
    raw_match = action.get("formula_text") or action.get("target_text") or action.get("text")
    if raw_match and str(raw_match).strip():
        clean_match = str(raw_match).strip().lower()

        # 3.1: Check AlternateContent elements (e.g. OpenXML math equations)
        alts = slide._element.xpath(".//*[local-name()='AlternateContent']")
        for alt in alts:
            alt_str = " ".join(t.text.strip() for t in alt.iter() if t.text and t.tag.endswith("}t")).lower()
            if clean_match in alt_str or alt_str in clean_match:
                p = alt.getparent()
                if p is not None:
                    p.remove(alt)
                    deleted_count += 1
        if deleted_count > 0:
            return deleted_count

        # 3.2: Check standard shapes
        for sh in list(slide.shapes):
            sh_str = " ".join(t.text.strip() for t in sh._element.iter() if t.text and t.tag.endswith("}t")).lower()
            if clean_match in sh_str or sh_str in clean_match:
                sp_elem = getattr(sh, "_element", None)
                if sp_elem is not None:
                    p = sp_elem.getparent()
                    if p is not None:
                        p.remove(sp_elem)
                        deleted_count += 1
        if deleted_count > 0:
            return deleted_count

    # 4. Formula actions without specific text (e.g. remove_formula, delete_formula, remove_math)
    atype = str(action.get("action", "")).lower()
    if atype in ("remove_formula", "delete_formula", "remove_math", "delete_math") or action.get("is_formula"):
        alts = slide._element.xpath(".//*[local-name()='AlternateContent']")
        for alt in alts:
            if any("math" in c.tag.lower() for c in alt.iter()):
                p = alt.getparent()
                if p is not None:
                    p.remove(alt)
                    deleted_count += 1
        if deleted_count > 0:
            return deleted_count

    # 5. Name match
    sname = action.get("shape_name") or action.get("name")
    if sname:
        clean_name = str(sname).strip().lower()
        for sh in list(slide.shapes):
            if sh.name.lower() == clean_name:
                sp_elem = getattr(sh, "_element", None)
                if sp_elem is not None:
                    p = sp_elem.getparent()
                    if p is not None:
                        p.remove(sp_elem)
                        return 1

    return deleted_count


def apply_verification_edits(
    pptx_path: Path | str,
    actions: List[Dict[str, Any]],
    log_cb: Optional[Callable[[str], None]] = None
) -> Dict[str, Any]:
    """
    Applies the structured edit plan generated by the Verification Agent (or user feedback)
    directly onto the PowerPoint presentation (.pptx) file.
    Supports: delete_shape, remove_formula, update_text, update_table, update_notes, generate_image, delete_slide.
    """
    def log(msg: str):
        if log_cb:
            log_cb(msg)
        logger.info(msg)

    p = Path(pptx_path)
    if not p.exists():
        raise FileNotFoundError(f"PPTX file not found at: {p}")

    if not actions:
        log(f"[Verification Edits] No actions to apply for {p.name}.")
        return {"success": True, "applied_count": 0, "file_path": str(p)}

    log(f"[Verification Edits] Applying {len(actions)} edit action(s) to '{p.name}'...")
    prs = Presentation(str(p))
    applied_count = 0

    # 1. Handle slide deletions first (in reverse index order)
    delete_indices = sorted(
        [
            int(act["slide_index"]) for act in actions
            if act.get("action") in ("delete_slide", "remove_slide") and act.get("slide_index") is not None
        ],
        reverse=True
    )
    for d_idx in delete_indices:
        if 0 <= d_idx < len(prs.slides):
            try:
                rId = prs.slides._sldIdLst[d_idx].rId
                prs.part.drop_rel(rId)
                prs.slides._sldIdLst.remove(prs.slides._sldIdLst[d_idx])
                applied_count += 1
                log(f"[✓] Deleted slide index {d_idx}")
            except Exception as ex:
                log(f"[!] Delete slide error at index {d_idx}: {ex}")

    # 1.5 Handle slide reordering if specified
    for act in actions:
        if str(act.get("action", "")).lower() == "reorder_slides":
            new_order = act.get("new_order", [])
            if isinstance(new_order, list) and len(new_order) == len(prs.slides):
                try:
                    old_sld_ids = list(prs.slides._sldIdLst)
                    prs.slides._sldIdLst.clear()
                    for idx in new_order:
                        prs.slides._sldIdLst.append(old_sld_ids[idx])
                    applied_count += 1
                    log(f"[✓] Reordered slides to: {new_order}")
                except Exception as ex:
                    log(f"[!] Reorder slides notice: {ex}")

    # 2. Handle shape / element / formula deletions
    for act in actions:
        atype = str(act.get("action", "")).lower()
        s_idx = act.get("slide_index")
        if s_idx is None or s_idx < 0 or s_idx >= len(prs.slides):
            continue
        slide = prs.slides[s_idx]

        if atype in DELETE_ACTION_NAMES:
            c = delete_element_from_slide(slide, act, log_cb=log)
            if c > 0:
                applied_count += c
                desc = act.get("formula_text") or act.get("target_text") or act.get("shape_index") or atype
                log(f"[✓] Slide {s_idx + 1}: Deleted unwanted element/formula ({desc}).")

    # 3. Handle shape updates, geometry, typography, styling, cloning & code execution (Full Action Executor Access)
    for act in actions:
        atype = str(act.get("action", "")).lower()
        s_idx = act.get("slide_index")

        if s_idx is None or s_idx < 0 or s_idx >= len(prs.slides):
            continue

        slide = prs.slides[s_idx]

        # -------------------------------------------------------------
        # A. Text and Content Updates
        # -------------------------------------------------------------
        if atype in ("update_text", "clear_text", "delete_text"):
            sh_idx = act.get("shape_index")
            new_text = act.get("new_text", "")

            # If empty text or explicit clear_text, delete the shape/formula entirely
            if atype in ("clear_text", "delete_text") or not new_text or not str(new_text).strip():
                c = delete_element_from_slide(slide, act, log_cb=log)
                if c > 0:
                    applied_count += c
                    log(f"[✓] Slide {s_idx + 1}: Cleared/deleted empty shape #{sh_idx}.")
            elif sh_idx is not None:
                try:
                    sh_idx_int = int(sh_idx)
                    if 0 <= sh_idx_int < len(slide.shapes):
                        shape: Any = slide.shapes[sh_idx_int]
                        if getattr(shape, "has_text_frame", False):
                            _safe_update_text_frame(
                                shape.text_frame,
                                str(new_text),
                                is_rtl=act.get("is_rtl"),
                                max_box_width_emu=getattr(shape, "width", None),
                                max_box_height_emu=getattr(shape, "height", None),
                                shape=shape
                            )
                            # Apply optional inline font size or color
                            f_val = act.get("font_size") or act.get("font_size_pt")
                            if f_val is not None:
                                f_sz = float(f_val)
                                for p_elem in shape.text_frame.paragraphs:
                                    for r_elem in p_elem.runs:
                                        r_elem.font.size = Pt(f_sz)
                            if act.get("font_color") or act.get("color"):
                                col = _parse_hex_color(act.get("font_color") or act.get("color"))
                                if col:
                                    for p_elem in shape.text_frame.paragraphs:
                                        for r_elem in p_elem.runs:
                                            r_elem.font.color.rgb = col
                            applied_count += 1
                            log(f"[✓] Slide {s_idx + 1} shape #{sh_idx_int}: Updated text ('{str(new_text)[:35]}...')")
                except Exception as ex:
                    log(f"[!] Update text error: {ex}")

        elif atype == "update_table":
            sh_idx = act.get("shape_index")
            tdata = act.get("table_data")
            if sh_idx is not None and tdata:
                try:
                    sh_idx_int = int(sh_idx)
                    if 0 <= sh_idx_int < len(slide.shapes):
                        shape: Any = slide.shapes[sh_idx_int]
                        if getattr(shape, "has_table", False):
                            for r_i, row in enumerate(tdata):
                                if r_i < len(shape.table.rows):
                                    for c_i, cell_val in enumerate(row):
                                        if c_i < len(shape.table.columns):
                                            cell = shape.table.cell(r_i, c_i)
                                            cell.text = str(cell_val)
                                            if cell.text_frame and cell.text_frame.paragraphs:
                                                p_elem = cell.text_frame.paragraphs[0]
                                                _set_paragraph_rtl_and_fonts(p_elem, is_rtl=_is_rtl_text(str(cell_val)))
                            applied_count += 1
                            log(f"[✓] Slide {s_idx + 1} shape #{sh_idx_int}: Updated table data.")
                except Exception as ex:
                    log(f"[!] Update table error: {ex}")

        elif atype == "update_notes":
            notes_str = act.get("notes")
            if notes_str is not None:
                try:
                    notes_slide: Any = getattr(slide, "notes_slide", None)
                    if notes_slide and hasattr(notes_slide, "notes_text_frame") and notes_slide.notes_text_frame:
                        notes_slide.notes_text_frame.text = str(notes_str)
                        applied_count += 1
                        log(f"[✓] Slide {s_idx + 1}: Updated speaker notes.")
                except Exception as ex:
                    log(f"[!] Speaker notes update notice: {ex}")

        elif atype in ("generate_image", "replace_image"):
            sh_idx = act.get("shape_index")
            img_prompt = act.get("prompt")
            if sh_idx is not None and img_prompt:
                try:
                    sh_idx_int = int(sh_idx)
                    if 0 <= sh_idx_int < len(slide.shapes):
                        shape: Any = slide.shapes[sh_idx_int]
                        log(f"[*] Generating AI image for slide {s_idx + 1} shape #{sh_idx_int}: '{img_prompt[:40]}...'")
                        img_file = generate_image(img_prompt)
                        if img_file and not img_file.startswith("Error"):
                            if _replace_image_in_shape(shape, img_file):
                                applied_count += 1
                                log(f"[✓] Slide {s_idx + 1} shape #{sh_idx_int}: Replaced image.")
                except Exception as ex:
                    log(f"[!] Image generation notice: {ex}")

        # -------------------------------------------------------------
        # B. Geometry & Layout (Move, Resize, Align)
        # -------------------------------------------------------------
        elif atype in ("move_shape", "set_position", "reposition_shape", "reposition"):
            sh_idx = act.get("shape_index")
            if sh_idx is not None:
                try:
                    sh_idx_int = int(sh_idx)
                    if 0 <= sh_idx_int < len(slide.shapes):
                        shape: Any = slide.shapes[sh_idx_int]
                        new_left = _parse_coordinate(act.get("left"), prs.slide_width)
                        new_top = _parse_coordinate(act.get("top"), prs.slide_height)
                        dx = _parse_coordinate(act.get("dx"), prs.slide_width)
                        dy = _parse_coordinate(act.get("dy"), prs.slide_height)
                        if new_left is not None:
                            shape.left = new_left
                        elif dx is not None:
                            shape.left += dx
                        if new_top is not None:
                            shape.top = new_top
                        elif dy is not None:
                            shape.top += dy
                        applied_count += 1
                        log(f"[✓] Slide {s_idx + 1} shape #{sh_idx_int}: Repositioned (left={shape.left // 12700}pt, top={shape.top // 12700}pt).")
                except Exception as ex:
                    log(f"[!] Move shape error: {ex}")

        elif atype in ("resize_shape", "set_dimensions", "set_size"):
            sh_idx = act.get("shape_index")
            if sh_idx is not None:
                try:
                    sh_idx_int = int(sh_idx)
                    if 0 <= sh_idx_int < len(slide.shapes):
                        shape: Any = slide.shapes[sh_idx_int]
                        new_w = _parse_coordinate(act.get("width"), prs.slide_width)
                        new_h = _parse_coordinate(act.get("height"), prs.slide_height)
                        dw = _parse_coordinate(act.get("dw"), prs.slide_width)
                        dh = _parse_coordinate(act.get("dh"), prs.slide_height)
                        if new_w is not None and new_w > 0:
                            shape.width = new_w
                        elif dw is not None and (shape.width + dw) > 0:
                            shape.width += dw
                        if new_h is not None and new_h > 0:
                            shape.height = new_h
                        elif dh is not None and (shape.height + dh) > 0:
                            shape.height += dh
                        applied_count += 1
                        log(f"[✓] Slide {s_idx + 1} shape #{sh_idx_int}: Resized (width={shape.width // 12700}pt, height={shape.height // 12700}pt).")
                except Exception as ex:
                    log(f"[!] Resize shape error: {ex}")

        elif atype in ("align_shapes", "distribute_shapes"):
            sh_indices = act.get("shape_indices", [])
            alignment = str(act.get("alignment", "left")).lower()
            try:
                shapes_to_align: List[Any] = [
                    slide.shapes[int(i)] for i in sh_indices
                    if 0 <= int(i) < len(slide.shapes)
                ]
                if len(shapes_to_align) >= 2:
                    if alignment == "left":
                        min_l = min(s.left for s in shapes_to_align)
                        for s in shapes_to_align:
                            s.left = min_l
                    elif alignment == "right":
                        max_r = max(s.left + s.width for s in shapes_to_align)
                        for s in shapes_to_align:
                            s.left = max_r - s.width
                    elif alignment == "center":
                        avg_cx = sum(s.left + s.width // 2 for s in shapes_to_align) // len(shapes_to_align)
                        for s in shapes_to_align:
                            s.left = avg_cx - s.width // 2
                    elif alignment == "top":
                        min_t = min(s.top for s in shapes_to_align)
                        for s in shapes_to_align:
                            s.top = min_t
                    elif alignment == "bottom":
                        max_b = max(s.top + s.height for s in shapes_to_align)
                        for s in shapes_to_align:
                            s.top = max_b - s.height
                    elif alignment == "middle":
                        avg_cy = sum(s.top + s.height // 2 for s in shapes_to_align) // len(shapes_to_align)
                        for s in shapes_to_align:
                            s.top = avg_cy - s.height // 2
                    applied_count += 1
                    log(f"[✓] Slide {s_idx + 1}: Aligned {len(shapes_to_align)} shapes ({alignment}).")
            except Exception as ex:
                log(f"[!] Align shapes error: {ex}")

        # -------------------------------------------------------------
        # C. Typography & Text Styling
        # -------------------------------------------------------------
        elif atype in ("format_text", "set_font", "set_typography", "adjust_font"):
            sh_idx = act.get("shape_index")
            if sh_idx is not None:
                try:
                    sh_idx_int = int(sh_idx)
                    if 0 <= sh_idx_int < len(slide.shapes):
                        shape = slide.shapes[sh_idx_int]
                        if getattr(shape, "has_text_frame", False):
                            tf = shape.text_frame
                            f_size = act.get("font_size") or act.get("font_size_pt")
                            f_name = act.get("font_name")
                            f_col = act.get("font_color") or act.get("color")
                            rgb_col = _parse_hex_color(f_col) if f_col else None
                            is_bold = act.get("bold")
                            is_italic = act.get("italic")
                            is_underline = act.get("underline")
                            align_str = str(act.get("alignment", "")).lower()
                            rtl_val = act.get("is_rtl") if act.get("is_rtl") is not None else act.get("rtl")
                            wrap_val = act.get("word_wrap")

                            align_dict = {
                                "left": PP_ALIGN.LEFT,
                                "center": PP_ALIGN.CENTER,
                                "right": PP_ALIGN.RIGHT,
                                "justify": PP_ALIGN.JUSTIFY
                            }

                            if wrap_val is not None:
                                tf.word_wrap = bool(wrap_val)

                            for p_elem in tf.paragraphs:
                                if align_str in align_dict:
                                    p_elem.alignment = align_dict[align_str]
                                if rtl_val is not None:
                                    _set_paragraph_rtl_and_fonts(p_elem, font_name=f_name, is_rtl=bool(rtl_val))
                                for r_elem in p_elem.runs:
                                    if f_size:
                                        r_elem.font.size = Pt(float(f_size))
                                    if f_name:
                                        r_elem.font.name = str(f_name)
                                    if rgb_col:
                                        r_elem.font.color.rgb = rgb_col
                                    if is_bold is not None:
                                        r_elem.font.bold = bool(is_bold)
                                    if is_italic is not None:
                                        r_elem.font.italic = bool(is_italic)
                                    if is_underline is not None:
                                        r_elem.font.underline = bool(is_underline)
                                    if rtl_val is not None:
                                        _set_run_rtl_and_fonts(r_elem, font_name=f_name, is_rtl=bool(rtl_val))
                            applied_count += 1
                            log(f"[✓] Slide {s_idx + 1} shape #{sh_idx_int}: Formatted typography.")
                except Exception as ex:
                    log(f"[!] Format text error: {ex}")

        # -------------------------------------------------------------
        # D. Shape Styling (Fill, Border, Rotation)
        # -------------------------------------------------------------
        elif atype in ("set_shape_style", "set_fill", "set_border", "set_style"):
            sh_idx = act.get("shape_index")
            if sh_idx is not None:
                try:
                    sh_idx_int = int(sh_idx)
                    if 0 <= sh_idx_int < len(slide.shapes):
                        shape = slide.shapes[sh_idx_int]
                        fill_col = act.get("fill_color") or act.get("fill")
                        if fill_col in ("none", "transparent", False) and ("fill_color" in act or "fill" in act):
                            shape.fill.background()
                        elif fill_col:
                            c = _parse_hex_color(fill_col)
                            if c:
                                shape.fill.solid()
                                shape.fill.fore_color.rgb = c

                        border_col = act.get("border_color") or act.get("border")
                        if border_col in ("none", "transparent", False) and ("border_color" in act or "border" in act):
                            shape.line.fill.background()
                        elif border_col:
                            c = _parse_hex_color(border_col)
                            if c:
                                shape.line.color.rgb = c
                        bw = act.get("border_width") or act.get("border_width_pt")
                        if bw is not None:
                            shape.line.width = Pt(float(bw))

                        rot = act.get("rotation")
                        if rot is not None:
                            shape.rotation = float(rot)
                        applied_count += 1
                        log(f"[✓] Slide {s_idx + 1} shape #{sh_idx_int}: Updated shape styling.")
                except Exception as ex:
                    log(f"[!] Set shape style error: {ex}")

        # -------------------------------------------------------------
        # E. Z-Order Arrangement
        # -------------------------------------------------------------
        elif atype in ("set_z_order", "reorder_shape", "z_order"):
            sh_idx = act.get("shape_index")
            pos = str(act.get("position", act.get("order", "bring_to_front"))).lower()
            if sh_idx is not None:
                try:
                    sh_idx_int = int(sh_idx)
                    if 0 <= sh_idx_int < len(slide.shapes):
                        shape = slide.shapes[sh_idx_int]
                        sp_elem = shape._element
                        sp_tree = slide.shapes._spTree
                        if pos in ("bring_to_front", "front", "top"):
                            sp_tree.append(sp_elem)
                        elif pos in ("send_to_back", "back", "bottom"):
                            sp_tree.insert(2, sp_elem)
                        applied_count += 1
                        log(f"[✓] Slide {s_idx + 1} shape #{sh_idx_int}: Adjusted z-order ({pos}).")
                except Exception as ex:
                    log(f"[!] Z-order error: {ex}")

        # -------------------------------------------------------------
        # F. Element Creation & Cloning from Template
        # -------------------------------------------------------------
        elif atype in ("add_shape", "create_shape"):
            try:
                stype_str = act.get("shape_type", "rounded_rectangle")
                sh_type = _resolve_shape_type(stype_str)
                l = _parse_coordinate(act.get("left", 50), prs.slide_width) or int(Pt(50))
                t = _parse_coordinate(act.get("top", 50), prs.slide_height) or int(Pt(50))
                w = _parse_coordinate(act.get("width", 200), prs.slide_width) or int(Pt(200))
                h = _parse_coordinate(act.get("height", 80), prs.slide_height) or int(Pt(80))

                new_sh: Any
                if str(stype_str).lower() in ("text_box", "textbox"):
                    new_sh = slide.shapes.add_textbox(Length(l), Length(t), Length(w), Length(h))
                else:
                    new_sh = slide.shapes.add_shape(sh_type, Length(l), Length(t), Length(w), Length(h))

                txt = act.get("text")
                if txt and getattr(new_sh, "has_text_frame", False):
                    _safe_update_text_frame(new_sh.text_frame, str(txt), shape=new_sh)

                f_col = act.get("fill_color") or act.get("fill")
                if f_col:
                    c = _parse_hex_color(f_col)
                    if c and hasattr(new_sh, "fill"):
                        new_sh.fill.solid()
                        new_sh.fill.fore_color.rgb = c
                b_col = act.get("border_color") or act.get("border")
                if b_col:
                    c = _parse_hex_color(b_col)
                    if c and hasattr(new_sh, "line"):
                        new_sh.line.color.rgb = c

                applied_count += 1
                log(f"[✓] Slide {s_idx + 1}: Created new shape '{stype_str}'.")
            except Exception as ex:
                log(f"[!] Add shape error: {ex}")

        elif atype in ("clone_shape_from_template", "copy_from_template"):
            tpl_file = act.get("template_file")
            tpl_sidx = int(act.get("template_slide_index", 0))
            tpl_sh_idx = act.get("template_shape_index")
            if tpl_sh_idx is not None:
                try:
                    t_path = None
                    if tpl_file:
                        t_path = DATA_DIR / tpl_file
                    if not t_path or not t_path.exists():
                        for f in DATA_DIR.glob("*.pptx"):
                            t_path = f
                            break
                    if t_path and t_path.exists():
                        tprs = Presentation(str(t_path))
                        if 0 <= tpl_sidx < len(tprs.slides):
                            tslide = tprs.slides[tpl_sidx]
                            cloned: Any = _clone_shape_to_slide(tslide, slide, int(tpl_sh_idx), prs, tprs)
                            if cloned:
                                if act.get("new_text") and getattr(cloned, "has_text_frame", False):
                                    _safe_update_text_frame(cloned.text_frame, str(act["new_text"]), shape=cloned)
                                nl = _parse_coordinate(act.get("left"), prs.slide_width)
                                nt = _parse_coordinate(act.get("top"), prs.slide_height)
                                if nl is not None:
                                    cloned.left = nl
                                if nt is not None:
                                    cloned.top = nt
                                applied_count += 1
                                log(f"[✓] Slide {s_idx + 1}: Cloned shape #{tpl_sh_idx} from template '{t_path.name}'.")
                except Exception as ex:
                    log(f"[!] Clone shape error: {ex}")

        elif atype == "duplicate_shape":
            sh_idx = act.get("shape_index")
            if sh_idx is not None:
                try:
                    sh_idx_int = int(sh_idx)
                    if 0 <= sh_idx_int < len(slide.shapes):
                        src_sh: Any = slide.shapes[sh_idx_int]
                        new_elem = copy.deepcopy(src_sh._element)
                        slide.shapes._spTree.append(new_elem)
                        new_sh: Any = slide.shapes[-1]
                        dx = _parse_coordinate(act.get("dx", 20), prs.slide_width) or int(Pt(20))
                        dy = _parse_coordinate(act.get("dy", 20), prs.slide_height) or int(Pt(20))
                        new_sh.left += dx
                        new_sh.top += dy
                        if act.get("new_text") and getattr(new_sh, "has_text_frame", False):
                            _safe_update_text_frame(new_sh.text_frame, str(act["new_text"]), shape=new_sh)
                        applied_count += 1
                        log(f"[✓] Slide {s_idx + 1}: Duplicated shape #{sh_idx_int}.")
                except Exception as ex:
                    log(f"[!] Duplicate shape error: {ex}")

        # -------------------------------------------------------------
        # G. Direct OpenXML Patch & Safe Python Script Execution
        # -------------------------------------------------------------
        elif atype in ("patch_oxml", "modify_oxml"):
            xp = act.get("xpath")
            sh_idx = act.get("shape_index")
            try:
                target_elem = (
                    slide.shapes[int(sh_idx)]._element
                    if (sh_idx is not None and 0 <= int(sh_idx) < len(slide.shapes))
                    else slide._element
                )
                if xp:
                    matches = target_elem.xpath(xp)
                    attrs = act.get("attributes", {})
                    for m in matches:
                        for k, v in attrs.items():
                            m.set(k, str(v))
                    applied_count += 1
                    log(f"[✓] Slide {s_idx + 1}: Patched OXML at xpath '{xp}'.")
            except Exception as ex:
                log(f"[!] Patch OXML error: {ex}")

        elif atype in ("execute_python", "run_code", "python_script"):
            code_str = act.get("code") or act.get("script")
            if code_str:
                try:
                    exec_scope = {
                        "prs": prs,
                        "slide": slide,
                        "shapes": list(slide.shapes),
                        "Pt": Pt,
                        "Inches": Inches,
                        "RGBColor": RGBColor,
                        "MSO_SHAPE": MSO_SHAPE,
                        "MSO_SHAPE_TYPE": MSO_SHAPE_TYPE,
                        "PP_ALIGN": PP_ALIGN,
                        "log": log,
                        "_safe_update_text_frame": _safe_update_text_frame,
                        "_set_paragraph_rtl_and_fonts": _set_paragraph_rtl_and_fonts,
                        "_set_run_rtl_and_fonts": _set_run_rtl_and_fonts,
                        "_parse_coordinate": _parse_coordinate,
                        "_parse_hex_color": _parse_hex_color,
                    }
                    sh_idx = act.get("shape_index")
                    if sh_idx is not None and 0 <= int(sh_idx) < len(slide.shapes):
                        exec_scope["shape"] = slide.shapes[int(sh_idx)]
                    exec(code_str, exec_scope)
                    applied_count += 1
                    log(f"[✓] Slide {s_idx + 1}: Executed custom Python script in Action Executor.")
                except Exception as py_ex:
                    log(f"[!] Custom script execution notice: {py_ex}")

    prs.save(str(p))
    log(f"[✓] Presentation saved with {applied_count} verification edit(s).")
    return {
        "success": True,
        "applied_count": applied_count,
        "file_path": str(p),
        "filename": p.name
    }


# ---------------------------------------------------------------------------
# 2. COMPARISON PREPARATION: Extract Template vs Generated Slides
# ---------------------------------------------------------------------------
def prepare_slide_comparison_data(
    pptx_path: Path | str,
    ai_plan: Optional[Dict[str, Any]],
    template_inventory: List[Dict[str, Any]],
    screenshot_width: int = 650,
    log_cb: Optional[Callable[[str], None]] = None
) -> List[Dict[str, Any]]:
    """
    Builds paired comparison entries between the chosen template slide and the generated slide.
    Each entry includes:
    - slide_index: 0-based index of generated slide
    - template_file: filename of template used
    - template_slide_index: 0-based slide index in template
    - template_screenshot: base64 JPEG of template slide
    - generated_screenshot: base64 JPEG of generated slide
    - template_slots: list of original template slots
    - generated_shapes: list of current generated slide shapes & text
    - target_section: section name / topic
    """
    def log(msg: str):
        if log_cb:
            log_cb(msg)
        logger.info(msg)

    p = Path(pptx_path)
    if not p.exists():
        return []

    # 1. Render all generated slides
    log(f"[Slide Verification] Rendering generated slide previews for comparison...")
    gen_previews: List[str] = []
    try:
        imgs = render_pptx_file_previews(p, target_width_px=screenshot_width)
        if isinstance(imgs, tuple):
            imgs = imgs[0]
        img_list = imgs if isinstance(imgs, list) else [imgs]
        for img in img_list:
            gen_previews.append(image_to_base64_jpeg(img, quality=82))
    except Exception as ex:
        log(f"[!] Could not render generated slide previews: {ex}")

    # 2. Inspect generated presentation structure
    gen_deck_info = inspect_pptx_for_editing(p)
    gen_slides = gen_deck_info.get("slides", [])
    try:
        gen_prs = Presentation(str(p))
    except Exception:
        gen_prs = None

    # Index template inventory by (template_file, slide_index)
    tpl_lookup: Dict[Tuple[str, int], Dict[str, Any]] = {}
    for entry in template_inventory:
        tfile = str(entry.get("template_file", ""))
        sidx = entry.get("slide_index", 0)
        tpl_lookup[(tfile, sidx)] = entry

    # Open source template presentations on-demand to extract screenshots and semantic layouts
    prs_cache: Dict[str, Any] = {}
    def get_template_prs(tname: str) -> Any:
        if tname not in prs_cache:
            cand = DATA_DIR / tname
            if cand.exists():
                try:
                    prs_cache[tname] = Presentation(str(cand))
                except Exception:
                    pass
        return prs_cache.get(tname)

    pairs: List[Dict[str, Any]] = []
    planned_slides = (ai_plan.get("slides", []) if ai_plan else [])

    for s_idx, gen_slide in enumerate(gen_slides):
        s_plan = planned_slides[s_idx] if s_idx < len(planned_slides) else {}
        tpl_name = s_plan.get("source_template") or (template_inventory[0]["template_file"] if template_inventory else "Template")
        tpl_sidx = s_plan.get("source_slide_index", 0)

        # Retrieve template screenshot
        tpl_entry = tpl_lookup.get((tpl_name, tpl_sidx)) or (template_inventory[0] if template_inventory else {})
        tpl_screenshot = tpl_entry.get("screenshot_base64")

        # If template screenshot or semantic layout not cached, render/extract directly
        tprs = get_template_prs(tpl_name)
        tpl_semantic = ""
        if tprs and 0 <= tpl_sidx < len(tprs.slides):
            try:
                tslide = tprs.slides[tpl_sidx]
                if not tpl_screenshot:
                    t_img = render_pptx_slide_to_image(tslide, tprs.slide_width, tprs.slide_height, target_width_px=screenshot_width)
                    tpl_screenshot = image_to_base64_jpeg(t_img, quality=82)
                tpl_semantic = extract_slide_semantic_layout(tslide, tprs.slide_width, tprs.slide_height)
            except Exception as ex:
                log(f"[!] Could not extract template screenshot/layout: {ex}")

        gen_screenshot = gen_previews[s_idx] if s_idx < len(gen_previews) else None
        gen_semantic = ""
        if gen_prs and s_idx < len(gen_prs.slides):
            try:
                gen_semantic = extract_slide_semantic_layout(gen_prs.slides[s_idx], gen_prs.slide_width, gen_prs.slide_height)
            except Exception:
                pass

        # Filter generated shapes to meaningful items (including formulas)
        meaningful_shapes = []
        for sh in gen_slide.get("shapes", []):
            if sh.get("text") or sh.get("has_table") or sh.get("is_picture") or sh.get("is_formula"):
                meaningful_shapes.append({
                    "shape_index": sh.get("shape_index"),
                    "name": sh.get("name"),
                    "type": sh.get("type"),
                    "text": sh.get("text"),
                    "has_table": sh.get("has_table"),
                    "table_data": sh.get("table_data") if sh.get("has_table") else None,
                    "is_picture": sh.get("is_picture"),
                    "is_formula": sh.get("is_formula", False)
                })

        pairs.append({
            "slide_index": s_idx,
            "slide_number": s_idx + 1,
            "title": gen_slide.get("title") or f"Slide {s_idx + 1}",
            "source_template": tpl_name,
            "source_slide_index": tpl_sidx,
            "target_section": s_plan.get("target_section", gen_slide.get("title", "")),
            "template_screenshot": tpl_screenshot,
            "generated_screenshot": gen_screenshot,
            "template_slots": tpl_entry.get("text_slots", []),
            "template_semantic_layout": tpl_semantic,
            "generated_semantic_layout": gen_semantic,
            "generated_shapes": meaningful_shapes,
            "speaker_notes": gen_slide.get("notes", "")
        })

    return pairs


# ---------------------------------------------------------------------------
# 3. VERIFICATION AGENT: Multi-Modal Slide Comparison & Edit Structure
# ---------------------------------------------------------------------------
def verify_slide_alignment_with_ai(
    slide_pair: Dict[str, Any],
    doc_context: Optional[Dict[str, Any]] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    timeout: Optional[float] = None
) -> Dict[str, Any]:
    """
    Evaluates one slide comparison with the 9Router Vision AI Agent.
    Receives:
      - Template slide screenshot + Semantic XML layout
      - Generated slide screenshot + Semantic XML layout
      - Full access to Action Executor for high-precision healing
    Returns:
      {
        "is_correct": bool,
        "score": int, # 0-100
        "detected_issues": ["Issue description..."],
        "edit_structure": {
          "summary_of_changes": "...",
          "actions": [
            {"action": "...", "slide_index": int, ...}
          ]
        }
      }
    """
    def log(msg: str):
        if log_cb:
            log_cb(msg)
        logger.info(msg)

    effective_timeout = float(timeout or Config.LLM_TIMEOUT)
    s_idx = slide_pair["slide_index"]
    tpl_name = slide_pair.get("source_template", "Template")
    tpl_sidx = slide_pair.get("source_slide_index", 0)

    client = OpenAI(
        api_key=Config.NINEROUTER_KEY or "dummy_key",
        base_url=Config.get_openai_base_url(),
        timeout=effective_timeout
    )

    system_prompt = (
        "You are PrismPresenter Autonomous Slide Verification & Quality Assurance Inspector. "
        "Your task is to compare a Chosen Template Slide against the Generated PowerPoint Slide to verify visual and content correctness.\n\n"
        "You have FULL UNRESTRICTED ACCESS to the Action Executor to heal, refine, and align the slide. "
        "You receive both visual screenshots AND high-precision Semantic XML layouts containing exact shape coordinates (pt & %), dimensions, and typography.\n\n"
        "Specifically check:\n"
        "1. Unreplaced formula text / equations or symbols from original template: If the template slide contained mathematical equations (e.g. 2n, n=6, math formulas, axis numbers) or specialized graphics that DO NOT match the generated slide's topic, YOU MUST ISSUE A DELETION ACTION TO REMOVE THEM:\n"
        "   {'action': 'delete_shape', 'slide_index': int, 'shape_index': int, 'formula_text': '2n'}\n"
        "   or {'action': 'remove_formula', 'slide_index': int, 'formula_text': '2n'}\n"
        "2. Missing extra elements or texts: Did the original template have subtitles, badges, tags, card numbers (01, 02), subheaders, or category labels that were left empty or missed in the generated slide? Use 'clone_shape_from_template' or 'add_shape' to restore them!\n"
        "3. Unreplaced placeholder text: Are there shapes still containing default dummy text (e.g. 'Lorem ipsum', 'Sample text', 'Header Here', 'Subtitle Goes Here', 'Click to edit')?\n"
        "4. Card / Multi-column completeness: If the template has 3 or 4 feature cards, were all cards populated with meaningful content from the topic, or was one left blank?\n"
        "5. Layout balance, geometry & clipping: Are titles overflowing or clipped? Use 'resize_shape', 'move_shape', or 'format_text' (reducing font_size) to heal them!\n\n"
        "FULL ACTION EXECUTOR CAPABILITIES AVAILABLE TO YOU:\n"
        "- 'delete_shape' / 'remove_formula': {'action': 'delete_shape', 'slide_index': int, 'shape_index': int, 'formula_text': str}\n"
        "- 'update_text': {'action': 'update_text', 'slide_index': int, 'shape_index': int, 'new_text': str, 'font_size': float, 'font_color': '#HEX'}\n"
        "- 'move_shape': {'action': 'move_shape', 'slide_index': int, 'shape_index': int, 'left': '120pt' or '15%', 'top': '60pt', 'dx': '10pt', 'dy': '-5pt'}\n"
        "- 'resize_shape': {'action': 'resize_shape', 'slide_index': int, 'shape_index': int, 'width': '350pt', 'height': '180pt', 'dw': '20pt', 'dh': '-10pt'}\n"
        "- 'align_shapes': {'action': 'align_shapes', 'slide_index': int, 'shape_indices': [int, ...], 'alignment': 'left'|'center'|'right'|'top'|'middle'|'bottom'}\n"
        "- 'format_text': {'action': 'format_text', 'slide_index': int, 'shape_index': int, 'font_size': float, 'font_name': str, 'font_color': '#HEX', 'bold': bool, 'alignment': 'right'|'left'|'center', 'is_rtl': bool}\n"
        "- 'set_shape_style': {'action': 'set_shape_style', 'slide_index': int, 'shape_index': int, 'fill_color': '#HEX'|'none', 'border_color': '#HEX', 'border_width': float}\n"
        "- 'set_z_order': {'action': 'set_z_order', 'slide_index': int, 'shape_index': int, 'position': 'bring_to_front'|'send_to_back'}\n"
        "- 'clone_shape_from_template': {'action': 'clone_shape_from_template', 'slide_index': int, 'template_shape_index': int, 'new_text': str, 'left': '...', 'top': '...'} (resurrects missing template cards, badges, icons directly from template!)\n"
        "- 'add_shape': {'action': 'add_shape', 'slide_index': int, 'shape_type': 'rounded_rectangle'|'text_box'|'badge', 'left': '...', 'top': '...', 'width': '...', 'height': '...', 'text': '...'}\n"
        "- 'duplicate_shape': {'action': 'duplicate_shape', 'slide_index': int, 'shape_index': int, 'dx': '20pt', 'dy': '0pt', 'new_text': '...'}\n"
        "- 'update_table': {'action': 'update_table', 'slide_index': int, 'shape_index': int, 'table_data': [[...]]}\n"
        "- 'update_notes': {'action': 'update_notes', 'slide_index': int, 'notes': str}\n"
        "- 'generate_image': {'action': 'generate_image', 'slide_index': int, 'shape_index': int, 'prompt': str}\n"
        "- 'execute_python': {'action': 'execute_python', 'slide_index': int, 'code': 'shape.left = Pt(100)'} (arbitrary programmatic python-pptx manipulation)\n\n"
        "Return ONLY valid JSON."
    )

    # Prepare slot descriptions for context
    tpl_slots_summary = []
    for slot in slide_pair.get("template_slots", []):
        text_sample = (slot.get("original_text") or "").strip()
        if text_sample or slot.get("is_title") or slot.get("is_table") or slot.get("is_formula"):
            tpl_slots_summary.append({
                "shape_index": slot.get("shape_index"),
                "sample_template_text": text_sample[:80],
                "is_title": slot.get("is_title", False),
                "is_table": slot.get("is_table", False),
                "is_formula": slot.get("is_formula", False)
            })

    gen_shapes_summary = []
    for sh in slide_pair.get("generated_shapes", []):
        entry = {
            "shape_index": sh.get("shape_index"),
            "name": sh.get("name"),
            "type": sh.get("type"),
            "text": (sh.get("text") or "")[:200],
            "has_table": sh.get("has_table", False)
        }
        if sh.get("is_formula"):
            entry["is_formula"] = True
        gen_shapes_summary.append(entry)

    tpl_xml = slide_pair.get("template_semantic_layout") or ""
    gen_xml = slide_pair.get("generated_semantic_layout") or ""
    semantic_xml_block = ""
    if tpl_xml or gen_xml:
        semantic_xml_block = f"""
Original Template Slide Layout (Semantic XML):
```xml
{tpl_xml[:2500]}
```

Generated Slide Current Layout (Semantic XML):
```xml
{gen_xml[:2500]}
```
"""

    prompt_text = f"""
Slide Under Inspection: Slide {s_idx + 1}
Target Section / Topic: {slide_pair.get('target_section', 'Slide Topic')}
Source Template: {tpl_name} (Slide {tpl_sidx + 1})
{semantic_xml_block}
Template Shapes & Sample Text:
{json.dumps(tpl_slots_summary, ensure_ascii=False, indent=2)}

Generated Slide Current Shapes & Text:
{json.dumps(gen_shapes_summary, ensure_ascii=False, indent=2)}

Speaker Notes: {slide_pair.get('speaker_notes', '')[:200]}

Instructions:
Evaluate if the generated slide accurately adapted the template without missing extra texts, subtitle slots, badges, or leaving unreplaced placeholder strings / unwanted formulas.
CRITICAL: If an unreplaced formula or math element from the template does not belong on this slide, you MUST formulate a 'delete_shape' or 'remove_formula' action to delete it!
You have FULL ACCESS to Action Executor: if shapes are clipped, misaligned, missing, or need styling/resizing, formulate the exact healing actions.

Output ONLY valid JSON adhering to:
{{
  "is_correct": false,
  "score": 75,
  "detected_issues": [
    "Unreplaced formula text from original template that does not match current topic.",
    "Template had a subtitle under the title that was left empty."
  ],
  "edit_structure": {{
    "summary_of_changes": "Deleted unreplaced template formula and added missing subtitle.",
    "actions": [
      {{
        "action": "delete_shape",
        "slide_index": {s_idx},
        "shape_index": 4,
        "formula_text": "2n"
      }},
      {{
        "action": "update_text",
        "slide_index": {s_idx},
        "shape_index": 2,
        "new_text": "Meaningful Subtitle Adapted From Topic"
      }}
    ]
  }}
}}
"""

    tpl_b64 = slide_pair.get("template_screenshot")
    gen_b64 = slide_pair.get("generated_screenshot")

    # Multimodal Vision reasoning (passing both template and generated screenshots)
    if tpl_b64 and gen_b64:
        user_content: List[Dict[str, Any]] = [
            {"type": "text", "text": prompt_text},
            {
                "type": "text",
                "text": "Image 1 below: ORIGINAL TEMPLATE SLIDE (inspect for badges, subtitles, cards, and extra elements):"
            },
            {"type": "image_url", "image_url": {"url": tpl_b64}},
            {
                "type": "text",
                "text": "Image 2 below: GENERATED SLIDE (inspect if elements or texts were missed or unpopulated):"
            },
            {"type": "image_url", "image_url": {"url": gen_b64}},
        ]
        try:
            v_model = Config.get_agent_model("verifier")
            v_think = Config.get_agent_think_level("verifier")
            log(f"[Verification Agent] Visual inspection for Slide {s_idx + 1} with 9Router Vision AI '{v_model}' (thinking: {v_think})...")
            messages_vision: Any = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content}
            ]
            res = Config.safe_chat_completion(
                client,
                "verifier",
                model=v_model,
                messages=messages_vision,
                temperature=0.2,
                timeout=min(effective_timeout, 45.0)
            )
            raw = res.choices[0].message.content or "{}"
            parsed = safe_json_loads(raw)
            if parsed and isinstance(parsed, dict) and "detected_issues" in parsed:
                return _normalize_verification_result(parsed, s_idx, slide_pair=slide_pair)
        except Exception as ex_vision:
            log(f"[Verification Agent Notice] Vision call notice for slide {s_idx + 1}: {ex_vision}. Using structural comparison...")

    # Fallback to structural text reasoning
    try:
        v_model = Config.get_agent_model("verifier")
        v_think = Config.get_agent_think_level("verifier")
        log(f"[Verification Agent] Structural inspection for Slide {s_idx + 1} with '{v_model}' (thinking: {v_think})...")
        res = Config.safe_chat_completion(
            client,
            "verifier",
            model=v_model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt_text}
            ],
            temperature=0.2,
            timeout=min(effective_timeout, 35.0)
        )
        raw = res.choices[0].message.content or "{}"
        parsed = safe_json_loads(raw)
        if parsed and isinstance(parsed, dict):
            return _normalize_verification_result(parsed, s_idx, slide_pair=slide_pair)
    except Exception as ex_text:
        log(f"[Verification Agent Warning] Structural comparison notice for slide {s_idx + 1}: {ex_text}")

    # Default heuristic check if AI calls fail
    return _heuristic_verification(slide_pair, s_idx)


def _normalize_verification_result(
    parsed: Dict[str, Any],
    s_idx: int,
    slide_pair: Optional[Dict[str, Any]] = None
) -> Dict[str, Any]:
    """Ensures consistent schema, enforces slide_index, and guarantees formula deletions."""
    is_correct = bool(parsed.get("is_correct", False))
    score = int(parsed.get("score", 90 if is_correct else 70))
    issues = parsed.get("detected_issues", [])
    if not isinstance(issues, list):
        issues = [str(issues)]

    edit_struct = parsed.get("edit_structure") or {}
    actions = edit_struct.get("actions", [])
    if not isinstance(actions, list):
        actions = []

    # Enforce slide_index
    for act in actions:
        if act.get("slide_index") is None:
            act["slide_index"] = s_idx

    # If any issue mentions unreplaced formula text / equations, ensure deletion action exists
    issue_text_blob = " ".join(str(i) for i in issues).lower()
    has_formula_issue = any(k in issue_text_blob for k in ["formula", "equation", "omath", "math"])
    has_delete_action = any(
        act.get("action") in DELETE_ACTION_NAMES or act.get("formula_text") or act.get("is_formula")
        for act in actions
    )

    if has_formula_issue and not has_delete_action:
        # Auto-inject deletion action for formula
        injected = False
        if slide_pair and slide_pair.get("generated_shapes"):
            for f_sh in slide_pair.get("generated_shapes", []):
                if f_sh.get("is_formula") or f_sh.get("type") == "FORMULA":
                    actions.append({
                        "action": "delete_shape",
                        "slide_index": s_idx,
                        "shape_index": f_sh.get("shape_index"),
                        "formula_text": f_sh.get("text", "")
                    })
                    injected = True
        if not injected:
            actions.append({
                "action": "remove_formula",
                "slide_index": s_idx
            })
        is_correct = False

    # If no issues and no actions, is_correct must be true
    if not issues and not actions:
        is_correct = True
        score = max(score, 95)
    elif actions:
        is_correct = False

    return {
        "is_correct": is_correct,
        "score": score,
        "detected_issues": issues,
        "edit_structure": {
            "summary_of_changes": edit_struct.get("summary_of_changes", "No modifications required." if is_correct else "Adjusted slide elements."),
            "actions": actions
        }
    }


def _heuristic_verification(slide_pair: Dict[str, Any], s_idx: int) -> Dict[str, Any]:
    """Lightweight rule-based verification when LLM is unreachable."""
    issues = []
    actions = []

    gen_shapes = slide_pair.get("generated_shapes", [])
    placeholder_keywords = ["lorem ipsum", "click to edit", "sample text", "subtitle goes here", "header here"]

    for sh in gen_shapes:
        txt = (sh.get("text") or "").lower()
        for kw in placeholder_keywords:
            if kw in txt:
                issues.append(f"Shape #{sh.get('shape_index')} appears to contain placeholder text: '{sh.get('text')}'")
                actions.append({
                    "action": "delete_shape",
                    "slide_index": s_idx,
                    "shape_index": sh.get("shape_index")
                })
                break

        # Check for unreplaced template formula
        if sh.get("is_formula") or sh.get("type") == "FORMULA":
            issues.append(f"Unreplaced formula text from original template that does not match current topic: '{sh.get('text')}'")
            actions.append({
                "action": "delete_shape",
                "slide_index": s_idx,
                "shape_index": sh.get("shape_index"),
                "formula_text": sh.get("text", "")
            })

    is_correct = len(issues) == 0
    return {
        "is_correct": is_correct,
        "score": 95 if is_correct else 75,
        "detected_issues": issues,
        "edit_structure": {
            "summary_of_changes": "Heuristic verification check completed.",
            "actions": actions
        }
    }


# ---------------------------------------------------------------------------
# 4. FULL PRESENTATION VERIFICATION PIPELINE
# ---------------------------------------------------------------------------
def run_presentation_visual_verification(
    pptx_path: Path | str,
    ai_plan: Optional[Dict[str, Any]],
    template_inventory: List[Dict[str, Any]],
    doc_context: Optional[Dict[str, Any]] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    timeout: Optional[float] = None,
    round_num: int = 1,
    max_rounds: int = 1
) -> Dict[str, Any]:
    """
    Step 4.5: Runs visual and structural verification for every slide in the presentation.
    Compares template slide screenshot vs generated slide screenshot.
    Supports iterative verification loops (round_num of max_rounds).
    """
    def log(msg: str):
        if log_cb:
            log_cb(msg)
        logger.info(msg)

    p = Path(pptx_path)
    log(f"[Step 4.5] Starting Visual Template Alignment & Slide Verification for '{p.name}' (Round {round_num}/{max_rounds})...")

    pairs = prepare_slide_comparison_data(
        pptx_path=p,
        ai_plan=ai_plan,
        template_inventory=template_inventory,
        screenshot_width=600,
        log_cb=log
    )

    if not pairs:
        log(f"[Step 4.5 Warning] No slide pairs could be constructed for comparison.")
        return {
            "all_correct": True,
            "total_issues_found": 0,
            "total_actions_planned": 0,
            "slides": [],
            "aggregated_actions": [],
            "round": round_num,
            "max_rounds": max_rounds
        }

    slide_results = []
    aggregated_actions = []
    total_issues = 0

    for pair in pairs:
        s_idx = pair["slide_index"]
        log(f"[Step 4.5 (Round {round_num}/{max_rounds})] Verifying Slide {s_idx + 1}/{len(pairs)} (Template: {pair.get('source_template')} #{pair.get('source_slide_index', 0) + 1})...")
        
        verif = verify_slide_alignment_with_ai(
            slide_pair=pair,
            doc_context=doc_context,
            log_cb=log,
            timeout=timeout
        )

        slide_entry = {
            "slide_index": s_idx,
            "slide_number": s_idx + 1,
            "title": pair.get("title", f"Slide {s_idx + 1}"),
            "source_template": pair.get("source_template"),
            "source_slide_index": pair.get("source_slide_index"),
            "target_section": pair.get("target_section"),
            "template_screenshot": pair.get("template_screenshot"),
            "generated_screenshot": pair.get("generated_screenshot"),
            "template_semantic_layout": pair.get("template_semantic_layout", ""),
            "generated_semantic_layout": pair.get("generated_semantic_layout", ""),
            "is_correct": verif.get("is_correct", True),
            "score": verif.get("score", 90),
            "detected_issues": verif.get("detected_issues", []),
            "edit_structure": verif.get("edit_structure", {"actions": []}),
            "generated_shapes": pair.get("generated_shapes", [])
        }

        slide_results.append(slide_entry)
        actions = verif.get("edit_structure", {}).get("actions", [])
        aggregated_actions.extend(actions)
        total_issues += len(verif.get("detected_issues", []))

    all_correct = (total_issues == 0 and len(aggregated_actions) == 0)
    log(f"[Step 4.5 (Round {round_num}/{max_rounds})] Slide verification complete: {len(slide_results)} slides checked, {total_issues} issue(s) detected, {len(aggregated_actions)} healing action(s) formulated.")

    return {
        "all_correct": all_correct,
        "total_issues_found": total_issues,
        "total_actions_planned": len(aggregated_actions),
        "slides": slide_results,
        "aggregated_actions": aggregated_actions,
        "round": round_num,
        "max_rounds": max_rounds
    }


# ---------------------------------------------------------------------------
# 5. HUMAN TOUCH CONVERSATION WITH VERIFICATION AGENT
# ---------------------------------------------------------------------------
def converse_with_verification_agent(
    current_verification_data: Dict[str, Any],
    user_message: str,
    active_slide_index: Optional[int] = None,
    conversation_history: Optional[List[Dict[str, str]]] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    timeout: Optional[float] = None
) -> Dict[str, Any]:
    """
    Interactive conversation between User and Verification Agent during Human Touch mode.
    Allows user to give instructions like:
    "The subheader on Slide 2 was deleted, put it back with 'Annual Review 2026'",
    "Slide 3 badge should say 'Phase 2'",
    "Don't remove the bottom card, just fix its text",
    "Move shape #2 30 points to the right and make font size 18pt".

    Returns updated verification report with refined issues and edit_structure actions.
    """
    def log(msg: str):
        if log_cb:
            log_cb(msg)
        logger.info(msg)

    effective_timeout = float(timeout or Config.LLM_TIMEOUT)
    log(f"[*] Human Touch: User conversing with Verification Agent (timeout={effective_timeout}s)...")

    client = OpenAI(
        api_key=Config.NINEROUTER_KEY or "dummy_key",
        base_url=Config.get_openai_base_url(),
        timeout=effective_timeout
    )

    system_prompt = (
        "You are PrismPresenter Autonomous Slide Verification & Healing Partner conversing directly with the human designer. "
        "You receive the current slide verification data (template screenshots, semantic XML layouts, current slide shapes, detected issues, proposed edit actions) "
        "and user instructions.\n\n"
        "You have FULL ACCESS to Action Executor to formulate any necessary healing actions:\n"
        "- 'update_text': {'action': 'update_text', 'slide_index': int, 'shape_index': int, 'new_text': str}\n"
        "- 'move_shape': {'action': 'move_shape', 'slide_index': int, 'shape_index': int, 'left': '...', 'top': '...', 'dx': '...', 'dy': '...'}\n"
        "- 'resize_shape': {'action': 'resize_shape', 'slide_index': int, 'shape_index': int, 'width': '...', 'height': '...'}\n"
        "- 'align_shapes': {'action': 'align_shapes', 'slide_index': int, 'shape_indices': [int, ...], 'alignment': 'left'|'center'|'right'|'top'|'middle'|'bottom'}\n"
        "- 'format_text': {'action': 'format_text', 'slide_index': int, 'shape_index': int, 'font_size': float, 'font_name': str, 'font_color': '#HEX', 'bold': bool, 'alignment': 'right'|'left'|'center', 'is_rtl': bool}\n"
        "- 'set_shape_style': {'action': 'set_shape_style', 'slide_index': int, 'shape_index': int, 'fill_color': '#HEX'|'none', 'border_color': '#HEX', 'border_width': float}\n"
        "- 'set_z_order': {'action': 'set_z_order', 'slide_index': int, 'shape_index': int, 'position': 'bring_to_front'|'send_to_back'}\n"
        "- 'clone_shape_from_template': {'action': 'clone_shape_from_template', 'slide_index': int, 'template_shape_index': int, 'new_text': str}\n"
        "- 'add_shape': {'action': 'add_shape', 'slide_index': int, 'shape_type': 'rounded_rectangle'|'text_box'|'badge', 'left': '...', 'top': '...', 'width': '...', 'height': '...', 'text': '...'}\n"
        "- 'duplicate_shape': {'action': 'duplicate_shape', 'slide_index': int, 'shape_index': int, 'dx': '20pt', 'dy': '0pt', 'new_text': '...'}\n"
        "- 'delete_shape' / 'remove_formula': {'action': 'delete_shape', 'slide_index': int, 'shape_index': int, 'formula_text': str}\n"
        "- 'update_table': {'action': 'update_table', 'slide_index': int, 'shape_index': int, 'table_data': [[...]]}\n"
        "- 'update_notes': {'action': 'update_notes', 'slide_index': int, 'notes': str}\n"
        "- 'delete_slide': {'action': 'delete_slide', 'slide_index': int}\n"
        "- 'execute_python': {'action': 'execute_python', 'slide_index': int, 'code': '...'}\n\n"
        "Your task is to:\n"
        "1. Understand the user's specific feedback or correction.\n"
        "2. Adjust or refine the detected issues and edit_structure actions to exactly match what the user wants.\n"
        "3. Provide a helpful, concise explanation of the adjustments made.\n\n"
        "Return ONLY valid JSON with 'agent_reply', 'updated_slides', and 'aggregated_actions'."
    )

    slides_info = current_verification_data.get("slides", [])
    simplified_slides = []
    for s in slides_info:
        simplified_slides.append({
            "slide_index": s.get("slide_index"),
            "slide_number": s.get("slide_number"),
            "title": s.get("title"),
            "source_template": s.get("source_template"),
            "is_correct": s.get("is_correct"),
            "detected_issues": s.get("detected_issues"),
            "edit_structure": s.get("edit_structure"),
            "semantic_layout": s.get("generated_semantic_layout", "")[:1200],
            "shapes": [
                {
                    "shape_index": sh.get("shape_index"),
                    "name": sh.get("name"),
                    "text": (sh.get("text") or "")[:120],
                    "is_formula": sh.get("is_formula", False)
                }
                for sh in s.get("generated_shapes", [])
            ]
        })

    user_payload = {
        "active_slide_focus": active_slide_index if active_slide_index is not None else "all",
        "user_message": user_message,
        "current_verification_state": simplified_slides,
        "history": conversation_history or []
    }

    user_text = f"""
User Conversation Instruction:
"{user_message}"

Active Slide Selected: {f'Slide {active_slide_index + 1}' if active_slide_index is not None else 'All Slides'}

Current Verification Status & Slide Shapes:
```json
{json.dumps(user_payload, ensure_ascii=False, indent=2)}
```

Output ONLY valid JSON adhering strictly to:
{{
  "agent_reply": "Clear explanation of how the user's instructions were incorporated into the edit structure.",
  "updated_slides": [
    {{
      "slide_index": 0,
      "is_correct": false,
      "detected_issues": ["Issue 1"],
      "edit_structure": {{
        "summary_of_changes": "...",
        "actions": [
          {{
            "action": "update_text",
            "slide_index": 0,
            "shape_index": 1,
            "new_text": "Updated text"
          }}
        ]
      }}
    }}
  ],
  "aggregated_actions": [
    {{
      "action": "update_text",
      "slide_index": 0,
      "shape_index": 1,
      "new_text": "Updated text"
    }}
  ]
}}
"""

    err_msg = "Unknown error"
    try:
        v_model = Config.get_agent_model("verifier")
        v_think = Config.get_agent_think_level("verifier")
        res = Config.safe_chat_completion(
            client,
            "verifier",
            model=v_model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_text}
            ],
            temperature=0.2,
            timeout=effective_timeout
        )
        content = res.choices[0].message.content or "{}"
        parsed = safe_json_loads(content)

        if parsed and isinstance(parsed, dict):
            agent_reply = parsed.get("agent_reply", "Understood. Refined the verification actions accordingly.")
            updated_slides_map = {us.get("slide_index"): us for us in parsed.get("updated_slides", [])}

            # Merge updates into existing slides_info
            merged_slides = []
            for s in slides_info:
                idx = s.get("slide_index")
                if idx in updated_slides_map:
                    us = updated_slides_map[idx]
                    merged = dict(s)
                    merged["is_correct"] = us.get("is_correct", s.get("is_correct"))
                    merged["detected_issues"] = us.get("detected_issues", s.get("detected_issues"))
                    merged["edit_structure"] = us.get("edit_structure", s.get("edit_structure"))
                    merged_slides.append(merged)
                else:
                    merged_slides.append(s)

            agg_actions = parsed.get("aggregated_actions", [])
            if not agg_actions:
                agg_actions = [
                    act for ms in merged_slides for act in ms.get("edit_structure", {}).get("actions", [])
                ]

            return {
                "success": True,
                "agent_reply": agent_reply,
                "slides": merged_slides,
                "aggregated_actions": agg_actions,
                "total_issues_found": sum(len(ms.get("detected_issues", [])) for ms in merged_slides),
                "total_actions_planned": len(agg_actions)
            }
    except Exception as ex:
        err_msg = str(ex)
        log(f"[!] Verification agent conversation notice: {ex}")

    return {
        "success": False,
        "agent_reply": f"Could not process conversation instruction: {err_msg}",
        "slides": slides_info,
        "aggregated_actions": current_verification_data.get("aggregated_actions", [])
    }
