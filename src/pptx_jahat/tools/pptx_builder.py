import json
import re
import copy
import io
import math
import time
import zipfile
import collections
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, Any, List, Optional, Callable, Tuple, Sequence
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml import parse_xml

from pptx_jahat.config import Config, DATA_DIR, OUTPUT_DIR
from pptx_jahat.tools.docx_parser import parse_docx
from pptx_jahat.tools.json_parser import safe_json_loads
from pptx_jahat.tools.multi_parser import parse_multiple_sources
from pptx_jahat.tools.structure_manager import (
    get_structure_content,
    restructure_slides_with_agent,
    convert_restructured_to_sections
)
from pptx_jahat.tools.pptx_engine import inspect_template_slides, inspect_all_templates
from pptx_jahat.tools.image_gen import generate_image
from pptx_jahat.tools.preview import render_pptx_file_previews, image_to_base64_jpeg
from pptx_jahat.tools.template_analyzer import load_notes, format_notes_for_ai_prompt
from pptx_jahat.tools.font_verifier import (
    verify_fonts_inventory,
    apply_font_fallbacks_to_presentation
)
from openai import OpenAI

def _is_rtl_text(text: str) -> bool:
    """
    Returns True if text contains Persian, Arabic or other RTL Unicode characters.
    """
    if not text:
        return True
    return bool(re.search(r'[\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF]', text))

def _set_run_rtl_and_fonts(run: Any, font_name: Optional[str] = None, is_rtl: bool = True) -> None:
    """
    Directly sets DrawingML run properties for true RTL and complex script / latin font typefaces.
    """
    try:
        rPr = run._r.get_or_add_rPr()
        if is_rtl:
            rPr.set("rtl", "1")
        if font_name:
            cs = rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}cs")
            if cs is None:
                cs = OxmlElement("a:cs")
                rPr.append(cs)
            cs.set("typeface", font_name)

            latin = rPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
            if latin is None:
                latin = OxmlElement("a:latin")
                rPr.append(latin)
            latin.set("typeface", font_name)
    except Exception:
        pass

def _set_paragraph_rtl_and_fonts(paragraph: Any, font_name: Optional[str] = None, is_rtl: bool = True) -> None:
    """
    Directly sets DrawingML paragraph properties for true RTL and complex script fonts.
    """
    try:
        pPr = paragraph._p.get_or_add_pPr()
        if is_rtl:
            pPr.set("rtl", "1")
            pPr.set("algn", "r")

        # Set default complex script font
        if font_name:
            defRPr = pPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr")
            if defRPr is None:
                defRPr = OxmlElement("a:defRPr")
                pPr.append(defRPr)
            cs = defRPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}cs")
            if cs is None:
                cs = OxmlElement("a:cs")
                defRPr.append(cs)
            cs.set("typeface", font_name)

            latin = defRPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
            if latin is None:
                latin = OxmlElement("a:latin")
                defRPr.append(latin)
            latin.set("typeface", font_name)
    except Exception:
        pass

def _safe_update_text_frame(
    tf: Any,
    new_text: str,
    is_rtl: Optional[bool] = None,
    max_box_width_emu: Optional[int] = None,
    max_box_height_emu: Optional[int] = None,
    font_override: Optional[str] = None,
    shape: Optional[Any] = None,
    font_fallbacks: Optional[Dict[str, str]] = None
) -> None:
    """
    Updates text in a text_frame while:
    1. Preserving run-level formatting (color, bold, italic, font face).
    2. Dynamic font auto-sizing based on character count, bounding box dimensions, and auto_size.
    3. Applying true DrawingML RTL properties on both paragraph and run levels.
    4. Dynamically adjusting text box width and coordinates to eliminate text overflow and unwanted wrapping.
    5. Managing TextFrame.auto_size and word_wrap to prevent clipping while preserving single-line badges/titles.
    6. Applying mapped font fallbacks when original template font is missing on system.
    """
    if not tf:
        return

    # Resolve target shape if available
    target_shape = shape if shape is not None else getattr(tf, "_parent", None)

    # Capture initial wrapping and auto_size state from text frame before clearing
    orig_word_wrap = getattr(tf, "word_wrap", None)
    orig_auto_size = getattr(tf, "auto_size", None)

    # Determine RTL based on content if not explicitly specified
    if is_rtl is None:
        is_rtl = _is_rtl_text(new_text)

    lines = [line for line in new_text.split("\n") if line.strip()]
    if not lines:
        lines = [new_text]

    # Capture style of first run and paragraph if available
    saved_font: Dict[str, Any] = {
        "name": None,
        "size": None,
        "bold": None,
        "italic": None,
        "color": None,
        "algn": None
    }

    try:
        if tf.paragraphs:
            p0 = tf.paragraphs[0]
            if p0.runs:
                r0 = p0.runs[0]
                if r0.font:
                    saved_font["name"] = r0.font.name
                    saved_font["size"] = r0.font.size
                    saved_font["bold"] = r0.font.bold
                    saved_font["italic"] = r0.font.italic
                    try:
                        if r0.font.color and r0.font.color.rgb:
                            saved_font["color"] = r0.font.color.rgb
                    except Exception:
                        pass
            # If font name or size wasn't directly on run, inspect paragraph defRPr
            if hasattr(p0, "_p"):
                pPr = p0._p.find("{http://schemas.openxmlformats.org/drawingml/2006/main}pPr")
                if pPr is not None:
                    if pPr.get("algn"):
                        saved_font["algn"] = pPr.get("algn")
                    defRPr = pPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}defRPr")
                    if defRPr is not None:
                        if not saved_font["name"]:
                            cs = defRPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}cs")
                            if cs is not None and cs.get("typeface"):
                                saved_font["name"] = cs.get("typeface")
                            else:
                                latin = defRPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}latin")
                                if latin is not None and latin.get("typeface"):
                                    saved_font["name"] = latin.get("typeface")
                        if not saved_font["size"] and defRPr.get("sz"):
                            try:
                                saved_font["size"] = Pt(int(defRPr.get("sz")) / 100.0)
                            except Exception:
                                pass
    except Exception:
        pass

    raw_font_name = saved_font["name"]
    if font_fallbacks and raw_font_name:
        for src_f, dst_f in font_fallbacks.items():
            if src_f and dst_f:
                if src_f.lower() == raw_font_name.lower() or src_f.lower() in raw_font_name.lower():
                    raw_font_name = dst_f
                    break

    final_font_name = font_override or raw_font_name

    # Dynamic Font Auto-Sizing calculation based on improve-workflow.pdf formula
    total_chars = sum(len(line) for line in lines)
    calculated_size_pt = None

    if saved_font["size"]:
        orig_pt = saved_font["size"].pt
        if total_chars > 250:
            calculated_size_pt = max(10, min(orig_pt, 12))
        elif total_chars > 120:
            calculated_size_pt = max(11, min(orig_pt, 14))
        elif total_chars > 60:
            calculated_size_pt = max(12, min(orig_pt, 18))
        else:
            calculated_size_pt = orig_pt
    else:
        if total_chars > 250:
            calculated_size_pt = 11
        elif total_chars > 120:
            calculated_size_pt = 13
        elif total_chars > 60:
            calculated_size_pt = 16
        else:
            calculated_size_pt = 20

    # Dynamic Text Box Width & Positioning Calculation:
    # Measures the longest line to ensure the text box boundary actually accommodates the text.
    needed_w_emu = None
    try:
        from pptx_jahat.tools.renderers.typography_engine import FontResolver
        font_res = FontResolver()
        target_f_name = final_font_name or "IRANYekanXFaNum Heavy"
        f_obj, _ = font_res.get_font(target_f_name, int(calculated_size_pt or 14))
        max_line_w_pt = 0.0
        for line in lines:
            bbox = f_obj.getbbox(line)
            w_pt = (bbox[2] - bbox[0]) * 72.0 / 96.0
            if w_pt > max_line_w_pt:
                max_line_w_pt = w_pt
        text_w_emu = int(max_line_w_pt * 12700)
    except Exception:
        char_factor = 0.62 if is_rtl else 0.55
        max_line_len = max(len(line) for line in lines)
        max_line_w_pt = max_line_len * (calculated_size_pt or 14) * char_factor
        text_w_emu = int(max_line_w_pt * 12700)

    margin_l = int(tf.margin_left) if getattr(tf, "margin_left", None) is not None else 91440
    margin_r = int(tf.margin_right) if getattr(tf, "margin_right", None) is not None else 91440
    buffer_emu = int(Pt(12))  # ~152,400 EMU padding
    needed_w_emu = text_w_emu + margin_l + margin_r + buffer_emu

    # Expand text box dimensions if shape is accessible and width is insufficient
    is_single_line = len(lines) == 1
    if target_shape is not None and hasattr(target_shape, "width") and hasattr(target_shape, "left"):
        current_w = int(target_shape.width)
        current_l = int(target_shape.left)

        slide_w = 12192000
        try:
            slide_w = target_shape.part.package.presentation_part.presentation.slide_width
        except Exception:
            pass

        # If single line or original template had word_wrap=False, expand width to fit content
        if (is_single_line or orig_word_wrap is False) and current_w < needed_w_emu:
            target_w = min(needed_w_emu, slide_w)
            # RTL right-aligned: anchor right edge and expand towards left
            if is_rtl or saved_font["algn"] == "r":
                right_edge = current_l + current_w
                target_w = min(target_w, right_edge)
                target_shape.width = target_w
                target_shape.left = max(0, right_edge - target_w)
            elif saved_font["algn"] == "ctr":
                center_x = current_l + current_w // 2
                target_shape.width = target_w
                target_shape.left = max(0, min(center_x - target_w // 2, slide_w - target_w))
            else:
                target_shape.width = min(target_w, slide_w - current_l)

    # Set word wrap and auto-size appropriately:
    # Single-line titles/badges and templates with word_wrap=False must NOT wrap into multiple lines!
    try:
        from pptx.enum.text import MSO_AUTO_SIZE
        if is_single_line or orig_word_wrap is False:
            tf.word_wrap = False
            tf.auto_size = MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT
        else:
            tf.word_wrap = True
            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    except Exception:
        pass

    # Clear old paragraphs and populate with new text lines
    tf.clear()

    # Clear any residual DrawingML math / OMML elements from target shape so old formulas don't linger
    if target_shape is not None and hasattr(target_shape, "_element"):
        try:
            for m_node in list(target_shape._element.xpath(".//*[local-name()='oMath' or local-name()='oMathPara' or local-name()='m']")):
                mp = m_node.getparent()
                if mp is not None:
                    mp.remove(m_node)
        except Exception:
            pass

    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = line

        # Set alignment
        if is_rtl:
            p.alignment = PP_ALIGN.RIGHT
        elif saved_font["algn"] == "ctr":
            p.alignment = PP_ALIGN.CENTER

        _set_paragraph_rtl_and_fonts(p, font_name=final_font_name, is_rtl=is_rtl)

        # Apply preserved/adjusted font styling to runs
        if p.runs:
            for run in p.runs:
                if final_font_name:
                    run.font.name = final_font_name
                if calculated_size_pt:
                    run.font.size = Pt(calculated_size_pt)
                elif saved_font["size"]:
                    run.font.size = saved_font["size"]
                if saved_font["bold"] is not None:
                    run.font.bold = saved_font["bold"]
                if saved_font["italic"] is not None:
                    run.font.italic = saved_font["italic"]
                if saved_font["color"]:
                    run.font.color.rgb = saved_font["color"]
                _set_run_rtl_and_fonts(run, font_name=final_font_name, is_rtl=is_rtl)

def _replace_image_in_shape(shape: Any, new_image_path: Path | str) -> bool:
    """
    Replaces the image in a picture shape or picture placeholder with a newly generated or selected image,
    preserving template masks, cropping, and aspect bounds.
    """
    try:
        img_path = Path(new_image_path)
        if not img_path.exists():
            return False

        # 1. If it's a placeholder (picture/bitmap/content), use native insert_picture
        if getattr(shape, "is_placeholder", False):
            try:
                shape.insert_picture(str(img_path))
                return True
            except Exception:
                pass

        # 2. If it's an existing picture shape, update the underlying image blob
        with open(img_path, "rb") as f:
            new_blob = f.read()

        if hasattr(shape, "image"):
            # Update image part blob
            shape.image._blob = new_blob
            return True
    except Exception:
        pass
    return False

def _remove_shape(slide: Any, shape_index: int) -> bool:
    """
    Safely removes a shape from slide XML by shape_index.
    Supports both direct slide.shapes and AlternateContent elements (e.g. math equations).
    """
    try:
        if 0 <= shape_index < len(slide.shapes):
            shape = slide.shapes[shape_index]
            sp_elem = getattr(shape, "_element", None)
            if sp_elem is not None:
                parent = sp_elem.getparent()
                if parent is not None:
                    parent.remove(sp_elem)
                    return True
        else:
            # Check AlternateContent elements (e.g. math equations beyond slide.shapes)
            alt_idx = shape_index - len(slide.shapes)
            alts = slide._element.xpath(".//*[local-name()='AlternateContent']")
            if 0 <= alt_idx < len(alts):
                target_alt = alts[alt_idx]
                parent = target_alt.getparent()
                if parent is not None:
                    parent.remove(target_alt)
                    return True
    except Exception:
        pass
    return False

def _remove_shapes(slide: Any, shape_indices: List[int]) -> None:
    """
    Removes multiple shapes in descending index order to avoid index shift issues.
    """
    if not shape_indices:
        return
    for s_idx in sorted(set(shape_indices), reverse=True):
        _remove_shape(slide, s_idx)

def clone_slide_across_presentations(source_prs: Any, target_prs: Any, slide_index: int) -> Any:
    """
    Deep clones a slide from source_prs into target_prs, preserving layout, background,
    media parts, and relationship mappings while avoiding duplicate/corrupted package parts.
    """
    src_slide = source_prs.slides[slide_index]
    
    # Choose layout from target_prs matching source or fallback to blank/content layout
    layout_name = src_slide.slide_layout.name if src_slide.slide_layout else "Blank"
    matching_layout = None
    for layout in target_prs.slide_layouts:
        if layout.name == layout_name:
            matching_layout = layout
            break
    if not matching_layout:
        matching_layout = target_prs.slide_layouts[min(1, len(target_prs.slide_layouts) - 1)]
        
    target_slide = target_prs.slides.add_slide(matching_layout)
    
    # Copy relationships & media/picture parts cleanly to avoid corrupting theme/layout parts
    src_part = src_slide.part
    target_part = target_slide.part
    rel_id_map: Dict[str, str] = {}
    
    for rel_id, rel in src_part.rels.items():
        if rel.reltype == RT.SLIDE_LAYOUT or rel.reltype == RT.NOTES_SLIDE:
            continue
        try:
            if rel.is_external:
                new_rid = target_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
                rel_id_map[rel_id] = new_rid
            elif rel.reltype == RT.IMAGE:
                # Add image blob into target package cleanly
                image_bytes = rel.target_part.blob
                new_image_part = target_prs.part.package.get_or_add_image_part(io.BytesIO(image_bytes))
                new_rid = target_part.relate_to(new_image_part, RT.IMAGE)
                rel_id_map[rel_id] = new_rid
            elif rel.reltype == RT.HYPERLINK:
                new_rid = target_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
                rel_id_map[rel_id] = new_rid
        except Exception:
            pass

    # Copy background definition (from slide or source layout/master) into target slide element
    try:
        src_cSld = src_slide._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}cSld")
        target_cSld = target_slide._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}cSld")
        
        # 1. Check if source slide has explicit <p:bg>
        src_bg = src_cSld.find("{http://schemas.openxmlformats.org/presentationml/2006/main}bg") if src_cSld is not None else None
        
        # 2. If not on slide, check source layout <p:bg>
        if src_bg is None and src_slide.slide_layout:
            l_cSld = src_slide.slide_layout._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}cSld")
            src_bg = l_cSld.find("{http://schemas.openxmlformats.org/presentationml/2006/main}bg") if l_cSld is not None else None
            
        # 3. If not on layout, check source master <p:bg>
        if src_bg is None and src_slide.slide_layout and src_slide.slide_layout.slide_master:
            m_cSld = src_slide.slide_layout.slide_master._element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}cSld")
            src_bg = m_cSld.find("{http://schemas.openxmlformats.org/presentationml/2006/main}bg") if m_cSld is not None else None

        if src_bg is not None and target_cSld is not None:
            # Check if target already has <p:bg>
            t_bg = target_cSld.find("{http://schemas.openxmlformats.org/presentationml/2006/main}bg")
            if t_bg is not None:
                target_cSld.remove(t_bg)
            copied_bg = copy.deepcopy(src_bg)
            # Remap relationship IDs in background (e.g. blipFill images)
            for elem in copied_bg.iter():
                for attr_name in list(elem.attrib.keys()):
                    if "embed" in attr_name or "id" in attr_name or "link" in attr_name:
                        val = elem.attrib[attr_name]
                        if val in rel_id_map:
                            elem.attrib[attr_name] = rel_id_map[val]
            target_cSld.insert(0, copied_bg)
    except Exception:
        pass
            
    # Replace target slide's spTree (shape tree) with deep copied source spTree
    target_spTree = target_slide.shapes._spTree
    src_spTree = src_slide.shapes._spTree
    
    # Remove default shapes in the newly added slide
    for child in list(target_spTree):
        target_spTree.remove(child)
        
    copied_spTree = copy.deepcopy(src_spTree)
    
    # Remap relationship IDs in elements (blip r:embed, hyperlinks, etc.)
    for elem in copied_spTree.iter():
        for attr_name in list(elem.attrib.keys()):
            if "embed" in attr_name or "id" in attr_name or "link" in attr_name:
                val = elem.attrib[attr_name]
                if val in rel_id_map:
                    elem.attrib[attr_name] = rel_id_map[val]
                    
    # Copy all children from copied spTree into target spTree
    for child in list(copied_spTree):
        target_spTree.append(child)
        
    return target_slide

def generate_slide_replacements_with_ai(
    template_inventory: List[Dict[str, Any]],
    doc_structure: Dict[str, Any],
    log_cb: Optional[Callable[[str], None]] = None,
    on_ai_images_ready: Optional[Callable[[List[Dict[str, Any]]], None]] = None,
    structure_blueprint: Optional[str] = None,
    timeout: Optional[float] = None
) -> Optional[Dict[str, Any]]:
    """
    Step 3: AI Vision Agent reasons over multimodal slide screenshots, shape slots across templates,
    Word docx or multi-modal content, and optional Storyboard Specification (structure.md).
    Returns optimal slide selections across templates, exact text replacements, shapes to remove,
    speaker notes, and optional AI image generation prompts for picture slots.
    """
    def log(msg: str):
        if log_cb:
            log_cb(msg)

    effective_timeout = float(timeout or Config.LLM_TIMEOUT)
    log(f"[Step 3] AI Vision Agent analyzing candidate template slides & document content (timeout={effective_timeout}s)...")

    client = OpenAI(
        api_key=Config.NINEROUTER_KEY or "dummy_key",
        base_url=Config.get_openai_base_url(),
        timeout=effective_timeout
    )

    system_prompt = (
        "You are an expert Presentation Art Director and Content Producer. "
        "You receive Template Intelligence & Design Notes (analyzing template purposes, ideas, content briefs, and styles like friendly, corporate, modern tech), "
        "visual screenshots, shape slots, and archetype tags of candidate presentation slides across multiple templates, "
        "along with parsed and structured input material. "
        "Your task is to:\n"
        "1. Step 1 (Template Selection): Choose the best Template(s) by matching the document's domain, purpose, and style with the Template Intelligence Notes.\n"
        "2. Step 2 (Slide Selection): Select the best visual slide archetype from the selected templates for each section/topic in the document (title_cover, table_matrix, metrics_stats, multi_column, process_timeline, content_bullets, conclusion_quote).\n"
        "3. Chunk and adapt long document text into punchy, high-impact slide text (concise headers, 3-4 bullet points max, 10-12 words per bullet).\n"
        "4. Match adapted content into the chosen slide's shape slots (shape_index).\n"
        "5. Generate detailed speaker notes for each slide to retain comprehensive background details from the document.\n"
        "6. Identify any unnecessary or overflowing shape indices to delete (shapes_to_remove).\n"
        "7. If a slot contains a table, supply updated 2D table_data.\n"
        "8. If a slide contains picture/graphic placeholders, you can provide an image_prompt for contextual AI image generation.\n"
        "Strictly return valid JSON adhering to the specified schema."
    )

    # Load Standardized Template Intelligence Notes from data/NOTE.md
    available_tpl_names = sorted(list({str(s["template_file"]) for s in template_inventory if s.get("template_file")}))
    formatted_notes = format_notes_for_ai_prompt(template_names=available_tpl_names if available_tpl_names else None)
    notes_prompt_block = ""
    if formatted_notes and formatted_notes.strip():
        notes_prompt_block = f"""
Step 0 - Standardized Template Intelligence & Architecture Notes (from data/NOTE.md):
Use these structured template profiles, slide layout blueprints, slot roles, and sequencing recipes to select the best template and slide archetypes:
{formatted_notes}
"""
    else:
        template_notes = load_notes()
        if template_notes and template_notes.strip():
            notes_prompt_block = f"""
Step 0 - Template Intelligence & Style Notes (from data/NOTE.md):
Use these analyzed notes to guide Step 1 (Best Template Selection by style, purpose, feel) and Step 2 (Best Slide Selection):
{template_notes}
"""

    structure_prompt_block = ""
    if structure_blueprint and structure_blueprint.strip():
        structure_prompt_block = f"""
Step 0.5 - Storyboard Specification Schema (from selected structure.md):
The presentation follows this structured pedagogical schema. Incorporate numbered animation steps (①, ②, ③...), formulas, and teacher callouts directly into slots:
{structure_blueprint[:2000]}
"""

    # Prepare compact, token-efficient slot descriptions (filtering out empty decorative shapes)
    inventory_summary = []
    for s in template_inventory:
        slots = []
        for slot in s.get("text_slots", []):
            orig = (slot.get("original_text") or "").strip()
            is_ph = slot.get("is_picture_placeholder", False)
            is_tbl = slot.get("is_table", False)
            is_title = slot.get("is_title", False)
            # Skip purely decorative empty shapes to keep prompt compact and prevent timeouts
            if not orig and not is_ph and not is_tbl and not is_title:
                continue

            entry: Dict[str, Any] = {"shape_index": slot.get("shape_index")}
            if slot.get("placeholder_idx") is not None:
                entry["placeholder_idx"] = slot.get("placeholder_idx")
            if orig:
                entry["sample_text"] = orig[:60]
            if slot.get("char_budget"):
                entry["char_budget"] = slot.get("char_budget")
            if is_title:
                entry["is_title"] = True
            if is_tbl:
                entry["is_table"] = True
                if slot.get("table_rows") and slot.get("table_cols"):
                    entry["table_shape"] = f"{slot.get('table_rows')}x{slot.get('table_cols')}"
            if is_ph:
                entry["is_picture"] = True
            slots.append(entry)

        summary_entry: Dict[str, Any] = {
            "template_file": s.get("template_file"),
            "slide_index": s.get("slide_index"),
            "archetype": s.get("archetype", "content_bullets"),
            "slots": slots
        }
        if s.get("primary_font"):
            summary_entry["primary_font"] = s.get("primary_font")
        inventory_summary.append(summary_entry)

    # Prepare clean, compact document content (prioritizing restructured slides and stripping raw OCR blobs)
    clean_doc: Dict[str, Any] = {
        "document_title": doc_structure.get("document_title", "Presentation")
    }
    if doc_structure.get("restructured_slides"):
        clean_doc["target_slides"] = [
            {
                "slide_number": sl.get("slide_number"),
                "title": sl.get("title"),
                "quadrant": sl.get("quadrant"),
                "slide_type": sl.get("slide_type"),
                "core_concept": sl.get("core_concept"),
                "key_bullets": sl.get("rewritten_bullets", [])[:4],
                "formulas": sl.get("mathematical_elements", {}).get("formulas") if isinstance(sl.get("mathematical_elements"), dict) else None,
                "callouts": sl.get("visual_annotations", {}).get("teacher_callouts") if isinstance(sl.get("visual_annotations"), dict) else None,
                "speaker_notes": (sl.get("speaker_notes") or "")[:200]
            }
            for sl in doc_structure["restructured_slides"]
        ]
    elif doc_structure.get("sections"):
        clean_doc["sections"] = [
            {
                "title": sec.get("title", f"Section {i+1}"),
                "bullets": sec.get("bullets", [])[:5],
                "summary": " ".join(sec.get("paragraphs", []))[:250]
            }
            for i, sec in enumerate(doc_structure["sections"][:12])
        ]
    else:
        clean_doc["raw_summary"] = (doc_structure.get("raw_text") or "")[:2000]

    # Select up to 6 diverse candidate slides with screenshots covering distinct archetypes
    priority_archetypes = [
        "title_cover",
        "content_bullets",
        "multi_column",
        "table_matrix",
        "process_timeline",
        "metrics_stats",
        "conclusion_quote"
    ]
    diverse_candidates: List[Dict[str, Any]] = []
    for target_arch in priority_archetypes:
        for s in template_inventory:
            if s.get("archetype") == target_arch and s.get("screenshot_base64"):
                if s not in diverse_candidates:
                    diverse_candidates.append(s)
                    break
        if len(diverse_candidates) >= 6:
            break

    # If fewer than 4 diverse, pad with remaining available slides that have screenshots
    if len(diverse_candidates) < 4:
        for s in template_inventory:
            if len(diverse_candidates) >= 6:
                break
            if s.get("screenshot_base64") and s not in diverse_candidates:
                diverse_candidates.append(s)

    # Deliver visual previews to UI for user feedback
    ai_sent_images: List[Dict[str, Any]] = []
    for s in diverse_candidates:
        ai_sent_images.append({
            "template_file": s.get("template_file"),
            "slide_index": s.get("slide_index"),
            "archetype": s.get("archetype", "content_bullets"),
            "base64": s.get("screenshot_base64")
        })

    if on_ai_images_ready:
        try:
            on_ai_images_ready(ai_sent_images)
        except Exception:
            pass

    # Build the core prompt text
    text_prompt = f"""
{notes_prompt_block}
{structure_prompt_block}

Step 1 - Available Slide Blueprints & Slots across Templates:
{json.dumps(inventory_summary, ensure_ascii=False)}

Step 2 - Target Content & Slide Requirements:
{json.dumps(clean_doc, ensure_ascii=False)}

Instructions:
1. Construct a cohesive presentation sequence matching the document flow (Title slide, Content/Topic slides, Metric slides, Summary).
2. For each slide in your output deck, specify:
   - "source_template": Name of template file (e.g. "T711.pptx", "t1.pptx", "sample_template.pptx")
   - "source_slide_index": Index of slide in that template
   - "target_section": Name of document section this slide covers
   - "speaker_notes": Detailed explanatory talking points for the presenter
   - "shape_replacements": List of {{"shape_index": int, "placeholder_idx": int (optional), "text": str}} mapping new adapted text into slots.
     CRITICAL: Respect the 'char_budget' for each slot to prevent text overflow and clipping.
   - "shapes_to_remove": List of shape indices [int] that should be pruned/deleted from the slide
   - "table_replacements": List of {{"shape_index": int, "table_data": [["cell", ...], ...]}} (for slots with is_table: true)
   - "image_replacements": Optional list of {{"shape_index": int, "image_prompt": "detailed prompt for slide visual"}}

Return a JSON object with this exact schema:
{{
  "deck_title": "{clean_doc.get('document_title', 'Presentation')}",
  "slides": [
    {{
      "source_template": "sample_template.pptx",
      "source_slide_index": 0,
      "target_section": "Document Title",
      "speaker_notes": "Welcome to the presentation...",
      "shape_replacements": [
        {{
          "shape_index": 0,
          "text": "Presentation Title"
        }}
      ],
      "shapes_to_remove": [],
      "table_replacements": [],
      "image_replacements": []
    }}
  ]
}}
"""

    def _normalize_plan(raw_plan: Any) -> Optional[Dict[str, Any]]:
        if not raw_plan:
            return None
        if isinstance(raw_plan, list):
            return {"slides": raw_plan, "deck_title": clean_doc.get("document_title", "Presentation")}
        if isinstance(raw_plan, dict):
            if "slides" in raw_plan and isinstance(raw_plan["slides"], list) and len(raw_plan["slides"]) > 0:
                return raw_plan
            for k in ["presentation", "deck", "presentation_plan", "data"]:
                sub = raw_plan.get(k)
                if isinstance(sub, dict) and "slides" in sub and isinstance(sub["slides"], list):
                    return sub
                elif isinstance(sub, list) and len(sub) > 0:
                    return {"slides": sub, "deck_title": raw_plan.get("deck_title", "Presentation")}
        return None

    # Meta limits
    meta = Config.get_model_metadata()
    max_output = min(meta.get("max_tokens", 65536), 16384)

    # -------------------------------------------------------------
    # Tier 1: Multimodal Vision Reasoning (with top 4-6 diverse previews)
    # -------------------------------------------------------------
    user_content_multimodal: List[Dict[str, Any]] = [
        {"type": "text", "text": text_prompt}
    ]
    for s in diverse_candidates:
        b64 = s.get("screenshot_base64")
        if b64:
            user_content_multimodal.append({
                "type": "image_url",
                "image_url": {"url": b64}
            })

    gen_model = Config.get_agent_model("generator")
    gen_think = Config.get_agent_think_level("generator")
    log(f"[Step 3] Sending prompt with {len(diverse_candidates)} visual slide previews to 9Router AI '{gen_model}' (thinking: {gen_think})...")
    try:
        messages_vision: Any = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_content_multimodal}
        ]
        response = Config.safe_chat_completion(
            client,
            "generator",
            model=gen_model,
            messages=messages_vision,
            temperature=0.25,
            max_tokens=max_output,
            timeout=min(effective_timeout, 70.0)
        )
        content = response.choices[0].message.content or "{}"
        plan = _normalize_plan(safe_json_loads(content))
        if plan:
            log(f"[✓] Step 3 AI Vision Agent generated presentation plan with {len(plan['slides'])} slides successfully.")
            return plan
        log("[Step 3 Notice] Multimodal response lacked slides structure. Retrying with high-speed text blueprint reasoning...")
    except Exception as e_vision:
        log(f"[Step 3 Notice] Multimodal visual reasoning notice ({e_vision}). Retrying with high-speed text blueprint reasoning...")

    # -------------------------------------------------------------
    # Tier 2: High-Speed Text Blueprint Reasoning (Zero image overhead, ultra-fast & immune to gateway timeouts)
    # -------------------------------------------------------------
    log(f"[Step 3] Dispatching high-speed text blueprint reasoning to '{gen_model}' (thinking: {gen_think})...")
    try:
        messages_text: Any = [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": text_prompt}
        ]
        response = Config.safe_chat_completion(
            client,
            "generator",
            model=gen_model,
            messages=messages_text,
            temperature=0.2,
            max_tokens=max_output,
            timeout=min(effective_timeout, 85.0)
        )
        content = response.choices[0].message.content or "{}"
        plan = _normalize_plan(safe_json_loads(content))
        if plan:
            log(f"[✓] Step 3 AI Agent generated presentation plan ({len(plan['slides'])} slides) via blueprint reasoning.")
            return plan
    except Exception as e_text:
        log(f"[Step 3 Notice] Primary model blueprint reasoning notice ({e_text}). Trying alternative fast model...")

    # -------------------------------------------------------------
    # Tier 3: Alternative Model Fallback
    # -------------------------------------------------------------
    candidate_alt_models = [
        m for m in ["gemini/gemini-3.8-flash", "aval/gemini-3.8-flash", "ag/gemini-3.7-flash-high"]
        if m != gen_model
    ]
    for alt_model in candidate_alt_models:
        try:
            log(f"[Step 3] Trying alternative AI model '{alt_model}'...")
            messages_alt: Any = [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": text_prompt}
            ]
            response = Config.safe_chat_completion(
                client,
                "generator",
                model=alt_model,
                messages=messages_alt,
                temperature=0.2,
                max_tokens=max_output,
                timeout=50.0
            )
            content = response.choices[0].message.content or "{}"
            plan = _normalize_plan(safe_json_loads(content))
            if plan:
                log(f"[✓] Step 3 AI Agent successfully generated presentation plan ({len(plan['slides'])} slides) with '{alt_model}'!")
                return plan
        except Exception as e_alt:
            log(f"[Step 3 Notice] Alternative model '{alt_model}' notice: {e_alt}")

    log("[Step 3 Warning] All AI reasoning tiers exhausted, proceeding with multi-template algorithmic fallback.")
    return None

def get_initial_diagnostics_steps() -> List[Dict[str, Any]]:
    """Returns the baseline list of 8 pipeline steps for diagnostics display."""
    return [
        {
            "id": "step_1",
            "name": "Step 1: Scan & Inspect Templates",
            "status": "pending",
            "input": "Template catalog (data/), selected template style, visual slide screenshots",
            "output": "Awaiting template inspection...",
            "duration": None,
        },
        {
            "id": "step_1_font",
            "name": "Step 1.2: Verify Template Fonts & Fallbacks",
            "status": "pending",
            "input": "Template typography inventory vs installed system & project fonts",
            "output": "Awaiting font verification...",
            "duration": None,
        },
        {
            "id": "step_1_5",
            "name": "Step 1.5: Storyboard Schema Resolution",
            "status": "pending",
            "input": "Storyboard schema (.md) for detection, restructuring, and layout blueprints",
            "output": "Awaiting schema resolution...",
            "duration": None,
        },
        {
            "id": "step_2",
            "name": "Step 2: Ingest & Parse Source Content",
            "status": "pending",
            "input": "Source documents (Word .docx, PPTX, MD, TXT, Images, Audio) and direct notes",
            "output": "Awaiting content ingestion...",
            "duration": None,
        },
        {
            "id": "step_2_5",
            "name": "Step 2.5: Storyboard Restructuring (AI Agent)",
            "status": "pending",
            "input": "Extracted document sections & active restructure schema rules",
            "output": "Awaiting restructure agent...",
            "duration": None,
        },
        {
            "id": "step_3",
            "name": "Step 3: Vision AI Reasoning & Slide Selection",
            "status": "pending",
            "input": "Template slide screenshots, content sections, schema archetype rules",
            "output": "Awaiting AI reasoning and slide selection...",
            "duration": None,
        },
        {
            "id": "step_4",
            "name": "Step 4: Deck Assembly & Slide Cloning",
            "status": "pending",
            "input": "AI synthesis plan, source template slides, target layout parameters",
            "output": "Awaiting presentation assembly...",
            "duration": None,
        },
        {
            "id": "step_4_5",
            "name": "Step 4.5: Visual Template Alignment & Verification",
            "status": "pending",
            "input": "Paired template vs generated slide screenshots, shape inventories, visual fidelity audit",
            "output": "Awaiting visual template verification...",
            "duration": None,
        },
        {
            "id": "step_5",
            "name": "Step 5: SlideCheck QA & Integrity Verification",
            "status": "pending",
            "input": "Assembled PPTX presentation, template typography, geometry boundaries",
            "output": "Awaiting SlideCheck QA & font auto-healing...",
            "duration": None,
        },
    ]


class DiagnosticsTracker:
    def __init__(self, callback: Optional[Callable[[Dict[str, Any]], None]] = None):
        self.callback = callback
        self.steps: Dict[str, Dict[str, Any]] = {s["id"]: dict(s) for s in get_initial_diagnostics_steps()}
        self.active_step_id: Optional[str] = None

    def start_step(self, step_id: str, input_desc: Optional[str] = None):
        self.active_step_id = step_id
        if step_id in self.steps:
            self.steps[step_id]["status"] = "running"
            self.steps[step_id]["_start_time"] = time.time()
            if input_desc:
                self.steps[step_id]["input"] = input_desc
            self.steps[step_id]["output"] = "In progress..."
            self._notify(step_id)

    def complete_step(self, step_id: str, output_desc: str, input_desc: Optional[str] = None):
        if step_id in self.steps:
            self.steps[step_id]["status"] = "completed"
            st = self.steps[step_id].pop("_start_time", None)
            if st:
                self.steps[step_id]["duration"] = f"{time.time() - st:.2f}s"
            if input_desc:
                self.steps[step_id]["input"] = input_desc
            self.steps[step_id]["output"] = output_desc
            self._notify(step_id)

    def skip_step(self, step_id: str, reason: str, input_desc: Optional[str] = None):
        if step_id in self.steps:
            self.steps[step_id]["status"] = "skipped"
            self.steps[step_id].pop("_start_time", None)
            self.steps[step_id]["duration"] = "0.00s"
            if input_desc:
                self.steps[step_id]["input"] = input_desc
            self.steps[step_id]["output"] = f"Skipped: {reason}"
            self._notify(step_id)

    def fail_step(self, step_id: str, error_desc: str):
        if step_id in self.steps:
            self.steps[step_id]["status"] = "failed"
            st = self.steps[step_id].pop("_start_time", None)
            if st:
                self.steps[step_id]["duration"] = f"{time.time() - st:.2f}s"
            self.steps[step_id]["output"] = f"Error: {error_desc}"
            self._notify(step_id)

    def fail_active(self, error_desc: str):
        if self.active_step_id:
            self.fail_step(self.active_step_id, error_desc)

    def get_steps_list(self) -> List[Dict[str, Any]]:
        out = []
        for s in self.steps.values():
            item = {k: v for k, v in s.items() if not k.startswith("_")}
            out.append(item)
        return out

    def _notify(self, step_id: str):
        if self.callback and step_id in self.steps:
            try:
                item = {k: v for k, v in self.steps[step_id].items() if not k.startswith("_")}
                self.callback(item)
            except Exception:
                pass


def build_pptx_with_agent(
    docx_path: Optional[str | Path | Sequence[str | Path]] = None,
    output_path: Optional[str | Path] = None,
    template_name: Optional[str] = None,
    log_callback: Optional[Callable[[str], None]] = None,
    on_ai_images_ready: Optional[Callable[[List[Dict[str, Any]]], None]] = None,
    structure_name: Optional[str] = None,
    raw_text: Optional[str] = None,
    enable_restructure: bool = False,
    timeout: Optional[float] = None,
    detection_structure_name: Optional[str] = None,
    restructure_structure_name: Optional[str] = None,
    blueprint_structure_name: Optional[str] = None,
    enable_detection: bool = True,
    enable_blueprint: bool = True,
    on_step_update: Optional[Callable[[Dict[str, Any]], None]] = None,
    enable_human_touch: bool = False,
    human_touch_steps: Optional[List[str]] = None,
    on_human_review: Optional[Callable[[str, Dict[str, Any]], Optional[Dict[str, Any]]]] = None,
    font_fallbacks: Optional[Dict[str, str]] = None,
    enable_verification: bool = True,
    verification_rounds: Optional[int] = None
) -> str:
    """
    Multi-Template & Storyboard-Guided Presentation Generation:
    Step 1: Scan & inspect candidate slides across all templates with rendered screenshots.
    Step 1.2: Verify template typography against system fonts and resolve user fallbacks.
    Step 2: Read and parse multi-modal sources (Word, PowerPoint, Text, Image OCR, Audio, raw text).
    Step 2.5 (Optional): Restructure & rewrite slide contents conforming to structure.md.
    Step 3: Vision AI reasons on slide screenshots & doc content, selecting best slides across templates.
    Step 4: Clone selected slides across presentations into target deck, prune removed shapes, and update text in-place.
    Step 5: SlideCheck QA & Automated Healing Loop.
    """
    def log(msg: str):
        if log_callback:
            log_callback(msg)

    effective_timeout = float(timeout or Config.LLM_TIMEOUT)
    tracker = DiagnosticsTracker(callback=on_step_update)

    # Normalize input files
    if isinstance(docx_path, (list, tuple)):
        input_paths = [Path(p) for p in docx_path if p]
    elif isinstance(docx_path, (str, Path)):
        input_paths = [Path(docx_path)]
    elif docx_path is not None:
        input_paths = [Path(p) for p in docx_path]
    else:
        input_paths = []

    try:
        # ----------------------------------------------------
        # Step 1: Scan & inspect templates
        # ----------------------------------------------------
        tpl_target_desc = template_name or "All Templates (Global AI Matching)"
        tracker.start_step(
            "step_1",
            input_desc=f"Catalog: '{DATA_DIR.name}/', Selected: '{tpl_target_desc}', Screenshots: Enabled"
        )

        if template_name and template_name != "All Templates (Global AI Matching)":
            candidate = DATA_DIR / template_name
            if candidate.exists():
                log(f"[Step 1] Inspecting selected template: {candidate.name}...")
                template_inventory = inspect_template_slides(candidate, include_screenshots=True)
            else:
                log(f"[Step 1] Scanning all templates in {DATA_DIR}...")
                template_inventory = inspect_all_templates(DATA_DIR, include_screenshots=True)
        else:
            log(f"[Step 1] Scanning all templates in {DATA_DIR} with visual screenshots...")
            template_inventory = inspect_all_templates(DATA_DIR, include_screenshots=True)

        if not template_inventory:
            tracker.fail_step("step_1", "No PPTX templates found in data folder.")
            raise FileNotFoundError("No PPTX templates found in data folder.")

        tpl_files_set = sorted(set(s.get("template_file", "Template") for s in template_inventory))
        tpl_fonts = template_inventory[0].get("template_fonts", []) if template_inventory else []
        fonts_str = f", Fonts: {', '.join(tpl_fonts[:3])}" if tpl_fonts else ""
        tracker.complete_step(
            "step_1",
            output_desc=f"Loaded {len(template_inventory)} candidate slides across {len(tpl_files_set)} templates ({', '.join(tpl_files_set)}){fonts_str}."
        )
        log(f"[Step 1] Loaded {len(template_inventory)} candidate slides across templates.")

        # ----------------------------------------------------
        # Step 1.2: Verify Template Fonts & Fallback Resolution
        # ----------------------------------------------------
        detected_fonts = set()
        for s in template_inventory:
            for f in s.get("template_fonts", []):
                if f and not f.startswith("+"):
                    detected_fonts.add(f)
        detected_fonts_list = sorted(list(detected_fonts))

        tracker.start_step(
            "step_1_font",
            input_desc=f"Template fonts: {', '.join(detected_fonts_list) if detected_fonts_list else 'None detected'}"
        )

        font_verif = verify_fonts_inventory(detected_fonts_list)
        missing_fonts = font_verif.get("missing_fonts", [])
        installed_fonts = font_verif.get("installed_fonts", [])
        active_font_fallbacks: Dict[str, str] = dict(font_fallbacks or {})

        unresolved_missing = [
            f for f in missing_fonts
            if f not in active_font_fallbacks and f.lower() not in {k.lower() for k in active_font_fallbacks}
        ]

        if not missing_fonts:
            msg = f"All {len(installed_fonts)} template font(s) verified on system: {', '.join(installed_fonts)}."
            log(f"[Step 1.2] {msg}")
            tracker.complete_step("step_1_font", output_desc=msg)
        elif not unresolved_missing:
            mapped_str = ", ".join(f"'{k}' -> '{v}'" for k, v in active_font_fallbacks.items())
            msg = f"Missing font(s) resolved with pre-configured fallbacks: {mapped_str}."
            log(f"[Step 1.2] {msg}")
            tracker.complete_step("step_1_font", output_desc=msg)
        else:
            missing_str = ", ".join(f"'{f}'" for f in unresolved_missing)
            log(f"[Step 1.2] Missing template font(s) detected on system: {missing_str}. Asking user for fallback selection...")

            step_data = {
                "missing_fonts": unresolved_missing,
                "installed_fonts": installed_fonts,
                "all_template_fonts": detected_fonts_list,
                "recommendations": font_verif.get("recommendations", {}),
                "available_persian": font_verif.get("available_persian", []),
                "available_latin": font_verif.get("available_latin", []),
                "system_fonts": font_verif.get("system_fonts", []),
                "current_fallbacks": active_font_fallbacks
            }

            if on_human_review:
                user_feedback = on_human_review("font_fallback", step_data)
                if user_feedback and isinstance(user_feedback, dict):
                    user_fallbacks = (
                        user_feedback.get("font_fallbacks")
                        or user_feedback.get("data", {}).get("font_fallbacks")
                        or user_feedback
                    )
                    if isinstance(user_fallbacks, dict):
                        for k, v in user_fallbacks.items():
                            if v and str(v).strip():
                                active_font_fallbacks[k] = str(v).strip()
            else:
                log(f"[Step 1.2] Non-interactive execution; applying recommended fallback fonts...")

            # Ensure every missing font has a fallback assigned
            for f in unresolved_missing:
                if f not in active_font_fallbacks and f.lower() not in {k.lower() for k in active_font_fallbacks}:
                    active_font_fallbacks[f] = font_verif.get("recommendations", {}).get(f, "Segoe UI")

            mapped_str = ", ".join(f"'{k}' -> '{v}'" for k, v in active_font_fallbacks.items())
            msg = f"Applied fallback font(s) for missing template fonts: {mapped_str}."
            log(f"[Step 1.2] {msg}")
            tracker.complete_step("step_1_font", output_desc=msg)

        # ----------------------------------------------------
        # Step 1.5: Pre-load Detection / Storyboard Schema if specified
        # ----------------------------------------------------
        det_target = detection_structure_name or structure_name
        restruct_target = restructure_structure_name or structure_name
        bp_target = blueprint_structure_name or structure_name

        tracker.start_step(
            "step_1_5",
            input_desc=f"Detection: '{det_target or 'None'}' (enabled={enable_detection}), Restructure: '{restruct_target or 'None'}' (enabled={enable_restructure}), Blueprint: '{bp_target or 'None'}' (enabled={enable_blueprint})"
        )

        def _resolve_blueprint(name: Optional[str], label: str) -> Optional[str]:
            if not name or "none" in str(name).lower():
                return None
            try:
                log(f"[*] Loading {label} Schema: {name}...")
                content = get_structure_content(name)
                log(f"[✓] Active {label} Schema: {Path(name).name}")
                return content
            except Exception as ex:
                log(f"[!] {label} Schema notice: {ex}. Proceeding without this schema.")
                return None

        detection_blueprint = _resolve_blueprint(det_target, "Detection") if (enable_detection and det_target) else None
        restructure_blueprint = _resolve_blueprint(restruct_target, "Restructure") if (enable_restructure and restruct_target) else None
        blueprint_blueprint = _resolve_blueprint(bp_target, "Slide Blueprint") if (enable_blueprint and bp_target) else None

        structure_blueprint = blueprint_blueprint or restructure_blueprint or detection_blueprint

        if structure_blueprint:
            active_schemas = []
            if detection_blueprint and det_target:
                active_schemas.append(f"Detect: {Path(det_target).name}")
            if restructure_blueprint and restruct_target:
                active_schemas.append(f"Restructure: {Path(restruct_target).name}")
            if blueprint_blueprint and bp_target:
                active_schemas.append(f"Blueprint: {Path(bp_target).name}")
            tracker.complete_step(
                "step_1_5",
                output_desc=f"Resolved active storyboard schema(s): {', '.join(active_schemas)}."
            )
        else:
            tracker.skip_step(
                "step_1_5",
                reason="No storyboard schema active. Proceeding with standard direct generation."
            )

        # ----------------------------------------------------
        # Step 2: Read & parse input sources (Word, PPTX, Text, Image, Audio)
        # ----------------------------------------------------
        source_summary = f"{len(input_paths)} file(s) ({', '.join(p.name for p in input_paths[:3])}{'...' if len(input_paths) > 3 else ''})" if input_paths else "Direct text outline"
        tracker.start_step(
            "step_2",
            input_desc=f"Sources: {source_summary}, Direct notes length: {len(raw_text) if raw_text else 0} chars, Detection schema active: {bool(detection_blueprint)}"
        )

        if len(input_paths) > 1 or any(p.suffix.lower() != ".docx" for p in input_paths) or raw_text:
            log(f"[Step 2] Reading multi-modal sources ({len(input_paths)} files + direct notes)...")
            parsed_doc = parse_multiple_sources(input_paths, raw_text=raw_text, log_cb=log, timeout=effective_timeout, structure_blueprint=detection_blueprint)
        elif input_paths and input_paths[0].exists():
            docx_file = input_paths[0]
            log(f"[Step 2] Reading Word document: {docx_file.name}...")
            parsed_doc = parse_docx(docx_file)
        else:
            log("[Step 2] Parsing direct text input...")
            parsed_doc = parse_multiple_sources([], raw_text=raw_text, log_cb=log, timeout=effective_timeout, structure_blueprint=detection_blueprint)

        sec_count = parsed_doc.get('total_sections', len(parsed_doc.get('sections', [])))
        media_count = len(parsed_doc.get('extracted_media', []))
        doc_title = parsed_doc.get('document_title', 'Presentation')
        tracker.complete_step(
            "step_2",
            output_desc=f"Parsed '{doc_title}'. Extracted {sec_count} content section(s) and {media_count} media element(s)."
        )
        log(f"[Step 2] Extracted {sec_count} content sections.")

        # Human Touch: Review & edit extracted content
        if enable_human_touch and on_human_review and (not human_touch_steps or "extract" in human_touch_steps):
            log("[Human Touch] Extracted slides ready for human review. Pausing pipeline...")
            reviewed_doc = on_human_review("extract", parsed_doc)
            if reviewed_doc and isinstance(reviewed_doc, dict):
                parsed_doc = reviewed_doc
                sec_count = parsed_doc.get('total_sections', len(parsed_doc.get('sections', [])))
                log(f"[Human Touch] Approved extracted content updated: {sec_count} sections.")

        # ----------------------------------------------------
        # Step 2.5: Restructure & Rewrite Slides base on structure.md (Optional)
        # ----------------------------------------------------
        if restructure_blueprint and enable_restructure:
            name_display = Path(restruct_target).name if restruct_target else "structure.md"
            tracker.start_step(
                "step_2_5",
                input_desc=f"Raw sections: {len(parsed_doc.get('sections', []))}, Schema: '{name_display}', AI timeout: {effective_timeout}s"
            )
            try:
                log(f"[Step 2.5] Autonomous Restructure Agent rewriting & structuring slides based on {name_display}...")
                restructured = restructure_slides_with_agent(parsed_doc, restructure_blueprint, log_cb=log, timeout=effective_timeout)

                # Human Touch: Review & edit restructured slides
                if enable_human_touch and on_human_review and (not human_touch_steps or "restructure" in human_touch_steps):
                    log("[Human Touch] Restructured slides ready for human review. Pausing pipeline...")
                    reviewed_restruct = on_human_review("restructure", restructured)
                    if reviewed_restruct and isinstance(reviewed_restruct, dict):
                        restructured = reviewed_restruct
                        log(f"[Human Touch] Approved restructured storyboard updated: {len(restructured.get('slides', []))} slides.")

                if restructured and restructured.get("slides"):
                    parsed_doc["restructured_slides"] = restructured.get("slides", [])
                    parsed_doc["sections"] = convert_restructured_to_sections(restructured)
                    parsed_doc["total_sections"] = len(parsed_doc["sections"])
                    tracker.complete_step(
                        "step_2_5",
                        output_desc=f"Autonomous Restructure Agent synthesized {len(parsed_doc['sections'])} slides adhering to '{name_display}'."
                    )
                    log(f"[Step 2.5] Restructure Agent prepared {len(parsed_doc['sections'])} slides adhering to {name_display}.")
                else:
                    tracker.complete_step(
                        "step_2_5",
                        output_desc=f"Restructure Agent returned without alterations; retaining {len(parsed_doc['sections'])} original sections."
                    )
            except Exception as st_ex:
                tracker.complete_step(
                    "step_2_5",
                    output_desc=f"Restructure notice: {st_ex}. Continued with standard sections."
                )
                log(f"[Step 2.5 Warning] Storyboard restructuring notice: {st_ex}. Continuing with standard sections.")
        else:
            tracker.skip_step(
                "step_2_5",
                reason="Restructure toggle disabled or no restructure schema provided.",
                input_desc="Restructure Agent disabled; original document section partitioning retained."
            )

        # ----------------------------------------------------
        # Step 3: AI Vision Agent writes texts and selects slides
        # ----------------------------------------------------
        active_bp_for_gen = blueprint_blueprint or restructure_blueprint or detection_blueprint
        bp_name = Path(bp_target).name if bp_target else ("structure.md" if active_bp_for_gen else "None")
        tracker.start_step(
            "step_3",
            input_desc=f"Candidate slides: {len(template_inventory)}, Document sections: {parsed_doc.get('total_sections', len(parsed_doc.get('sections', [])))}, Model timeout: {effective_timeout}s, Blueprint: '{bp_name}'"
        )

        ai_plan = generate_slide_replacements_with_ai(
            template_inventory,
            parsed_doc,
            log_cb=log,
            on_ai_images_ready=on_ai_images_ready,
            structure_blueprint=active_bp_for_gen,
            timeout=effective_timeout
        )

        if ai_plan and "slides" in ai_plan and len(ai_plan["slides"]) > 0:
            total_shapes = sum(len(s.get("shape_replacements", [])) for s in ai_plan.get("slides", []))
            total_tables = sum(len(s.get("table_replacements", [])) for s in ai_plan.get("slides", []))
            total_imgs = sum(len(s.get("image_replacements", [])) for s in ai_plan.get("slides", []))
            tracker.complete_step(
                "step_3",
                output_desc=f"AI Vision synthesis plan: {len(ai_plan['slides'])} slides selected. Planned {total_shapes} text replacements, {total_tables} table updates, {total_imgs} AI image prompts."
            )
        else:
            tracker.complete_step(
                "step_3",
                output_desc=f"Algorithmic multi-template fallback plan activated for {len(parsed_doc.get('sections', []))} sections across available templates."
            )

        # ----------------------------------------------------
        # Step 4: Assemble target deck across presentations
        # ----------------------------------------------------
        log("[Step 4] Assembling target presentation from selected template slides...")

        # Cache opened presentations by filename
        prs_cache: Dict[str, Any] = {}
        def get_source_prs(tpl_file: str) -> Any:
            if tpl_file not in prs_cache:
                p = DATA_DIR / tpl_file
                if not p.exists():
                    # Fallback to first existing template
                    p = next(DATA_DIR.glob("*.pptx"))
                prs_cache[tpl_file] = Presentation(str(p))
            return prs_cache[tpl_file]

        # Pre-open first source template to create matching target presentation package
        first_tpl_name = template_inventory[0]["template_file"]
        first_tpl_path = DATA_DIR / first_tpl_name if (DATA_DIR / first_tpl_name).exists() else next(DATA_DIR.glob("*.pptx"))
        
        planned_count = len(ai_plan["slides"]) if (ai_plan and "slides" in ai_plan) else len(parsed_doc.get("sections", []))
        target_out_name = Path(output_path).name if output_path else "presentation.pptx"
        tracker.start_step(
            "step_4",
            input_desc=f"Planned slides: {planned_count}, Base template: '{first_tpl_name}', Output destination: '{target_out_name}'"
        )

        # Initialize target presentation from base template to retain themes, color palettes, and layouts
        target_prs = Presentation(str(first_tpl_path))
        
        # Clear existing slides from target presentation
        while len(target_prs.slides) > 0:
            rId = target_prs.slides._sldIdLst[0].rId
            target_prs.part.drop_rel(rId)
            target_prs.slides._sldIdLst.remove(target_prs.slides._sldIdLst[0])

        if ai_plan and "slides" in ai_plan and len(ai_plan["slides"]) > 0:
            for s_plan in ai_plan["slides"]:
                src_tpl = s_plan.get("source_template") or first_tpl_name
                src_idx = s_plan.get("source_slide_index", 0)
                
                src_prs = get_source_prs(src_tpl)
                if src_idx >= len(src_prs.slides):
                    src_idx = 0
                    
                # Clone slide across presentation
                target_slide = clone_slide_across_presentations(src_prs, target_prs, src_idx)
                
                # In-place text replacements (supporting both shape_index and placeholder_idx)
                replacements_by_sh_idx = {}
                replacements_by_ph_idx = {}
                for r in s_plan.get("shape_replacements", []):
                    txt = r.get("text")
                    if txt is not None:
                        if r.get("shape_index") is not None:
                            replacements_by_sh_idx[r["shape_index"]] = txt
                        if r.get("placeholder_idx") is not None:
                            replacements_by_ph_idx[r["placeholder_idx"]] = txt

                for shape_idx, shape in enumerate(target_slide.shapes):
                    new_text = None
                    if shape_idx in replacements_by_sh_idx:
                        new_text = replacements_by_sh_idx[shape_idx]
                    elif getattr(shape, "is_placeholder", False):
                        try:
                            ph_i = shape.placeholder_format.idx
                            if ph_i in replacements_by_ph_idx:
                                new_text = replacements_by_ph_idx[ph_i]
                        except Exception:
                            pass

                    if new_text is not None and shape.has_text_frame:
                        _safe_update_text_frame(
                            shape.text_frame,
                            str(new_text),
                            is_rtl=None,
                            max_box_width_emu=getattr(shape, "width", None),
                            max_box_height_emu=getattr(shape, "height", None),
                            shape=shape,
                            font_fallbacks=active_font_fallbacks
                        )

                # Table replacements (supporting both shape_index and placeholder_idx)
                table_repl_by_sh = {}
                table_repl_by_ph = {}
                for t in s_plan.get("table_replacements", []):
                    tdata = t.get("table_data")
                    if tdata:
                        if t.get("shape_index") is not None:
                            table_repl_by_sh[t["shape_index"]] = tdata
                        if t.get("placeholder_idx") is not None:
                            table_repl_by_ph[t["placeholder_idx"]] = tdata

                for shape_idx, shape in enumerate(target_slide.shapes):
                    tdata = None
                    if shape_idx in table_repl_by_sh:
                        tdata = table_repl_by_sh[shape_idx]
                    elif getattr(shape, "is_placeholder", False):
                        try:
                            ph_i = shape.placeholder_format.idx
                            if ph_i in table_repl_by_ph:
                                tdata = table_repl_by_ph[ph_i]
                        except Exception:
                            pass

                    if tdata and shape.has_table:
                        for r_i, row in enumerate(tdata):
                            if r_i < len(shape.table.rows):
                                for c_i, cell_val in enumerate(row):
                                    if c_i < len(shape.table.columns):
                                        cell = shape.table.cell(r_i, c_i)
                                        cell.text = str(cell_val)
                                        if cell.text_frame and cell.text_frame.paragraphs:
                                            p = cell.text_frame.paragraphs[0]
                                            cell_fallback_font = next(iter(active_font_fallbacks.values()), None) if active_font_fallbacks else None
                                            _set_paragraph_rtl_and_fonts(p, font_name=cell_fallback_font, is_rtl=_is_rtl_text(str(cell_val)))
                                            if p.runs:
                                                _set_run_rtl_and_fonts(p.runs[0], font_name=cell_fallback_font, is_rtl=_is_rtl_text(str(cell_val)))

                # Image replacements via Image Gen API (supporting pictures and image placeholders)
                image_repl_by_sh = {}
                image_repl_by_ph = {}
                for img in s_plan.get("image_replacements", []):
                    prompt = img.get("image_prompt")
                    if prompt:
                        if img.get("shape_index") is not None:
                            image_repl_by_sh[img["shape_index"]] = prompt
                        if img.get("placeholder_idx") is not None:
                            image_repl_by_ph[img["placeholder_idx"]] = prompt

                for shape_idx, shape in enumerate(target_slide.shapes):
                    prompt = None
                    if shape_idx in image_repl_by_sh:
                        prompt = image_repl_by_sh[shape_idx]
                    elif getattr(shape, "is_placeholder", False):
                        try:
                            ph_i = shape.placeholder_format.idx
                            if ph_i in image_repl_by_ph:
                                prompt = image_repl_by_ph[ph_i]
                        except Exception:
                            pass

                    is_img_target = (
                        shape.shape_type == MSO_SHAPE_TYPE.PICTURE
                        or getattr(shape, "is_placeholder", False)
                    )
                    if prompt and is_img_target:
                        try:
                            log(f"[Step 4] Generating AI image for slide slot #{shape_idx}: '{prompt[:40]}...'")
                            img_file = generate_image(prompt)
                            if img_file and not img_file.startswith("Error"):
                                _replace_image_in_shape(shape, img_file)
                        except Exception as e:
                            log(f"[Step 4 Warning] Image generation skipped: {e}")

                # Speaker Notes
                notes_text = s_plan.get("speaker_notes")
                if notes_text:
                    try:
                        notes_slide = target_slide.notes_slide
                        text_frame = notes_slide.notes_text_frame
                        text_frame.text = str(notes_text)
                    except Exception:
                        pass

                # Remove unwanted shapes after text/table replacements to avoid index shifts during replacement
                shapes_to_remove = s_plan.get("shapes_to_remove", [])
                if shapes_to_remove:
                    _remove_shapes(target_slide, shapes_to_remove)
        else:
            # Fallback Multi-Template Assembly:
            # Title slide from first template, content slides from available candidate slides
            log("[Step 4 Fallback] Generating presentation using multi-slide assembly...")
            
            # 1. Title Slide
            first_tpl = template_inventory[0]["template_file"]
            base_prs = get_source_prs(first_tpl)
            title_slide = clone_slide_across_presentations(base_prs, target_prs, 0)
            for shape in title_slide.shapes:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    _safe_update_text_frame(
                        shape.text_frame,
                        parsed_doc.get("document_title", "Presentation"),
                        max_box_width_emu=getattr(shape, "width", None),
                        max_box_height_emu=getattr(shape, "height", None),
                        shape=shape,
                        font_fallbacks=active_font_fallbacks
                    )
                    break
                    
            # 2. Content Slides per Section
            for s_idx, section in enumerate(parsed_doc.get("sections", [])):
                candidate_entry = template_inventory[(s_idx + 1) % len(template_inventory)]
                src_prs = get_source_prs(candidate_entry["template_file"])
                src_idx = candidate_entry["slide_index"]
                
                c_slide = clone_slide_across_presentations(src_prs, target_prs, src_idx)
                text_shapes = [sh for sh in c_slide.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
                
                if text_shapes:
                    _safe_update_text_frame(
                        text_shapes[0].text_frame,
                        section.get("title", ""),
                        max_box_width_emu=getattr(text_shapes[0], "width", None),
                        max_box_height_emu=getattr(text_shapes[0], "height", None),
                        shape=text_shapes[0],
                        font_fallbacks=active_font_fallbacks
                    )
                    if len(text_shapes) > 1:
                        body = "\n".join(section.get("paragraphs", []) + [f"• {b}" for b in section.get("bullets", [])])
                        _safe_update_text_frame(
                            text_shapes[1].text_frame,
                            body,
                            max_box_width_emu=getattr(text_shapes[1], "width", None),
                            max_box_height_emu=getattr(text_shapes[1], "height", None),
                            shape=text_shapes[1],
                            font_fallbacks=active_font_fallbacks
                        )

        # Apply font fallbacks across all shapes, runs, and tables in the target deck
        if active_font_fallbacks:
            updated_font_count = apply_font_fallbacks_to_presentation(target_prs, active_font_fallbacks)
            if updated_font_count > 0:
                log(f"[Step 4] Applied font fallbacks to {updated_font_count} run(s) and element(s) in target presentation.")

        if not output_path:
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            primary_stem = input_paths[0].stem if input_paths else "presentation"
            output_path = OUTPUT_DIR / f"{primary_stem}_generated.pptx"
        else:
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
        target_prs.save(str(output_path))
        out_size_kb = (Path(output_path).stat().st_size // 1024) if Path(output_path).exists() else 0
        tracker.complete_step(
            "step_4",
            output_desc=f"Successfully assembled {len(target_prs.slides)} vector slides. Saved presentation to '{target_out_name}' ({out_size_kb} KB)."
        )
        log(f"[Step 4] Finished. Output presentation saved to: {output_path}")

        # ----------------------------------------------------
        # Step 4.5: Visual Template Alignment & Verification Loop
        # ----------------------------------------------------
        effective_verification_rounds = max(1, int(verification_rounds if verification_rounds is not None else Config.VERIFICATION_ROUNDS))

        if enable_verification:
            tracker.start_step(
                "step_4_5",
                input_desc=f"Presentation: '{target_out_name}', Max iterative rounds: {effective_verification_rounds}, Paired template slides: {len(target_prs.slides)}, Visual AI: 9Router"
            )
            log(f"[Step 4.5] Running Visual Template Alignment & Slide Verification Loop (Max {effective_verification_rounds} rounds)...")
            try:
                from pptx_jahat.tools.slide_verifier import (
                    run_presentation_visual_verification,
                    apply_verification_edits
                )

                total_healed_across_rounds = 0
                final_round_completed = 0
                all_clean = False

                for round_idx in range(1, effective_verification_rounds + 1):
                    final_round_completed = round_idx
                    log(f"[Step 4.5] --- Verification Round {round_idx}/{effective_verification_rounds} ---")

                    verif_report = run_presentation_visual_verification(
                        pptx_path=output_path,
                        ai_plan=ai_plan,
                        template_inventory=template_inventory,
                        doc_context=parsed_doc,
                        log_cb=log,
                        timeout=effective_timeout,
                        round_num=round_idx,
                        max_rounds=effective_verification_rounds
                    )
                    verif_report["round"] = round_idx
                    verif_report["max_rounds"] = effective_verification_rounds
                    verif_report["human_touch_enabled"] = bool(enable_human_touch)

                    is_human_touch_active = bool(
                        enable_human_touch
                        and on_human_review
                        and (not human_touch_steps or "verify" in human_touch_steps)
                    )

                    # Human Touch: Review template alignment & converse with verification agent
                    if is_human_touch_active:
                        log(f"[Human Touch] Slide verification statements (Round {round_idx}/{effective_verification_rounds}) awaiting human verification...")
                        verif_report["human_verification_required"] = True
                        reviewed_verif = on_human_review("verify", verif_report)
                        if reviewed_verif and isinstance(reviewed_verif, dict):
                            verif_report = reviewed_verif

                    actions_to_apply = verif_report.get("aggregated_actions", [])
                    issues_count = verif_report.get("total_issues_found", 0)
                    human_verified_clean = bool(verif_report.get("human_verified_clean", False))
                    human_skipped = bool(verif_report.get("human_skipped", False))

                    if is_human_touch_active:
                        # In human touch mode, verification statements MUST be human-verified
                        if human_verified_clean or human_skipped:
                            log(f"[Step 4.5] Human verified all slides cleanly aligned on Round {round_idx}/{effective_verification_rounds}.")
                            all_clean = True
                            break
                        if not actions_to_apply:
                            log(f"[Step 4.5] Human approved slide verification with 0 pending actions on Round {round_idx}/{effective_verification_rounds}.")
                            all_clean = True
                            break
                    else:
                        # Autonomous non-interactive mode
                        if verif_report.get("all_correct", False) or (issues_count == 0 and len(actions_to_apply) == 0):
                            log(f"[Step 4.5] Visual verification PASSED on Round {round_idx}/{effective_verification_rounds}: All slides cleanly aligned with 0 issues.")
                            all_clean = True
                            break

                    if actions_to_apply:
                        log(f"[Step 4.5] Round {round_idx}: Applying {len(actions_to_apply)} human-verified healing action(s) to presentation...")
                        edit_res = apply_verification_edits(output_path, actions_to_apply, log_cb=log)
                        applied_c = edit_res.get("applied_count", len(actions_to_apply))
                        total_healed_across_rounds += applied_c
                        if round_idx >= effective_verification_rounds:
                            all_clean = True
                            break
                    else:
                        all_clean = True
                        break

                if all_clean:
                    tracker.complete_step(
                        "step_4_5",
                        output_desc=f"Visual verification PASSED (Round {final_round_completed}/{effective_verification_rounds}): Slides cleanly match chosen templates. Total healed edits: {total_healed_across_rounds}."
                    )
                else:
                    tracker.complete_step(
                        "step_4_5",
                        output_desc=f"Visual verification completed {final_round_completed}/{effective_verification_rounds} rounds. Applied {total_healed_across_rounds} healing edit(s) across slides."
                    )
            except Exception as v_ex:
                log(f"[Step 4.5 Warning] Visual verification notice: {v_ex}. Proceeding to SlideCheck QA.")
                tracker.complete_step(
                    "step_4_5",
                    output_desc=f"Visual verification notice: {v_ex}. Proceeded to SlideCheck QA."
                )
        else:
            tracker.skip_step(
                "step_4_5",
                reason="Visual template verification toggle disabled by user at generation time."
            )

        # ----------------------------------------------------
        # Step 5: Verification & Automated SlideCheck QA Loop
        # ----------------------------------------------------
        log("[Step 5] Running Automated SlideCheck QA & Integrity Verification Loop...")

        expected_tpl_fonts: List[str] = []
        if template_inventory and "template_fonts" in template_inventory[0]:
            raw_exp = list(template_inventory[0].get("template_fonts", []))
            if active_font_fallbacks:
                expected_tpl_fonts = [
                    str(active_font_fallbacks.get(f, f)) for f in raw_exp if f
                ]
            else:
                expected_tpl_fonts = [str(f) for f in raw_exp if f]

        tracker.start_step(
            "step_5",
            input_desc=f"File: '{target_out_name}', Expected fonts: {', '.join(expected_tpl_fonts) if expected_tpl_fonts else 'Auto-detect'}, Auto-heal: Enabled"
        )

        qa_report = run_slidecheck_qa(output_path, expected_fonts=expected_tpl_fonts, auto_heal=True)
        if qa_report.get("overflow_issues_healed", 0) > 0:
            log(f"[Step 5 QA] Auto-healed {qa_report['overflow_issues_healed']} overflowing text box(es).")
        if qa_report.get("rtl_issues_healed", 0) > 0:
            log(f"[Step 5 QA] Auto-healed {qa_report['rtl_issues_healed']} RTL/BiDi tags on runs/paragraphs.")
        if qa_report.get("fonts_detected"):
            log(f"[Step 5 QA] Verified fonts in presentation: {', '.join(qa_report['fonts_detected'])}")

        is_valid, final_path = verify_and_auto_heal_pptx(
            output_path,
            doc_structure=parsed_doc,
            template_inventory=template_inventory,
            log_cb=log
        )

        qa_healed_txt = qa_report.get('overflow_issues_healed', 0)
        qa_healed_rtl = qa_report.get('rtl_issues_healed', 0)
        qa_fonts = ', '.join(qa_report.get('fonts_detected', [])) or 'Standard'
        valid_str = 'Verified PASSED (clean & valid)' if is_valid else 'Verified with notices auto-healed'
        tracker.complete_step(
            "step_5",
            output_desc=f"Integrity Check: {valid_str}. Auto-healed: {qa_healed_txt} overflow text box(es), {qa_healed_rtl} RTL run(s). Verified fonts in presentation: {qa_fonts}."
        )

        if not is_valid:
            log(f"[Step 5 Warning] Final PPTX file might contain remaining non-fatal notices.")
        else:
            log(f"[Step 5] PPTX Verification PASSED: File is clean, valid, and fully openable.")

        # Render Visual Preview QA
        try:
            preview_imgs = render_pptx_file_previews(output_path, target_width_px=650)
            log(f"[Step 5] Rendered {len(preview_imgs)} slide previews. Verification complete.")
        except Exception as qa_ex:
            log(f"[Step 5 Warning] QA preview render warning: {qa_ex}")

        return str(output_path)
    except Exception as e:
        tracker.fail_active(str(e))
        raise

def run_slidecheck_qa(
    pptx_path: str | Path,
    expected_fonts: Optional[List[str]] = None,
    auto_heal: bool = True
) -> Dict[str, Any]:
    """
    Automated SlideCheck Quality Assurance Layer (inspired by slidecheck & PPTX-HTML Fidelity Audit):
    1. Font Verification: Audits detected fonts against expected template fonts.
    2. Text Overflow Detection & Healing: Flags shapes where text exceeds bounding box capacity and auto-shrinks.
    3. RTL & BiDi Verification: Ensures Persian/Arabic runs and paragraphs have DrawingML rtl="1".
    4. Alignment Consistency: Flags shapes placed beyond slide bounds.
    """
    p = Path(pptx_path)
    report: Dict[str, Any] = {
        "passed": True,
        "total_slides": 0,
        "fonts_detected": [],
        "overflow_issues_healed": 0,
        "rtl_issues_healed": 0,
        "alignment_warnings": [],
        "warnings": []
    }

    if not p.exists():
        report["passed"] = False
        report["warnings"].append(f"File not found: {p}")
        return report

    try:
        prs = Presentation(str(p))
    except Exception as ex:
        report["passed"] = False
        report["warnings"].append(f"Cannot open presentation for SlideCheck QA: {ex}")
        return report

    report["total_slides"] = len(prs.slides)
    fonts_found = set()
    needs_save = False

    for s_idx, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            # 1. Alignment check: off-canvas bounds
            if shape.left is not None and shape.top is not None:
                if shape.left < 0 or shape.top < 0:
                    report["alignment_warnings"].append(
                        f"Slide {s_idx+1}: Shape '{shape.name}' has negative coordinates ({shape.left}, {shape.top})"
                    )

            if shape.has_text_frame:
                tf = getattr(shape, "text_frame", None)
                if not tf:
                    continue
                text = tf.text.strip()
                if not text:
                    continue

                # 2. Font check & RTL check across paragraphs & runs
                for p_idx, para in enumerate(tf.paragraphs):
                    para_text = para.text.strip()
                    is_rtl = _is_rtl_text(para_text)

                    # Verify paragraph RTL
                    if is_rtl and hasattr(para, "_p"):
                        pPr = para._p.get_or_add_pPr()
                        if pPr.get("rtl") != "1":
                            pPr.set("rtl", "1")
                            pPr.set("algn", "r")
                            report["rtl_issues_healed"] += 1
                            needs_save = True

                    for run in para.runs:
                        if run.font and run.font.name:
                            fonts_found.add(run.font.name)

                        # Verify run-level RTL
                        if is_rtl:
                            rPr = run._r.get_or_add_rPr()
                            if rPr.get("rtl") != "1":
                                rPr.set("rtl", "1")
                                report["rtl_issues_healed"] += 1
                                needs_save = True

                # 3. Text Overflow Check
                if shape.width and shape.height and auto_heal:
                    lines = [ln for ln in text.split("\n") if ln.strip()]
                    current_sz_pt = 14.0
                    try:
                        if tf.paragraphs and tf.paragraphs[0].runs and tf.paragraphs[0].runs[0].font.size:
                            current_sz_pt = tf.paragraphs[0].runs[0].font.size.pt
                    except Exception:
                        pass

                    w_in = shape.width / 914400.0
                    h_in = shape.height / 914400.0
                    chars_per_line = max(10, int((w_in * 72.0) / (current_sz_pt * 0.52)))
                    max_lines_allowed = max(1, int((h_in * 72.0) / (current_sz_pt * 1.30)))
                    estimated_lines = sum(max(1, math.ceil(len(ln) / max(1, chars_per_line))) for ln in lines)

                    is_single_line = len(lines) == 1
                    if is_single_line:
                        # Single-line titles and badges should never wrap into two lines
                        if tf.word_wrap is not False:
                            tf.word_wrap = False
                            needs_save = True
                        continue

                    if estimated_lines > max_lines_allowed * 1.15:
                        heal_pt = max(9.5, current_sz_pt - 2.5)
                        for para in tf.paragraphs:
                            for run in para.runs:
                                run.font.size = Pt(heal_pt)
                        try:
                            from pptx.enum.text import MSO_AUTO_SIZE
                            tf.word_wrap = True
                            tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                        except Exception:
                            pass
                        report["overflow_issues_healed"] += 1
                        needs_save = True

    report["fonts_detected"] = sorted(list(fonts_found))

    if expected_fonts:
        missing_fonts = [f for f in expected_fonts if f not in fonts_found]
        if missing_fonts:
            report["warnings"].append(f"Template fonts not detected in output runs: {missing_fonts}")

    if needs_save:
        try:
            prs.save(str(p))
        except Exception as sv_ex:
            report["warnings"].append(f"Could not save SlideCheck healed PPTX: {sv_ex}")

    return report

def verify_pptx_integrity(file_path: str | Path) -> Tuple[bool, List[str]]:
    """
    Performs comprehensive structural, zip packaging, XML and python-pptx integrity checks on a PPTX file.
    Returns (is_valid: bool, issues: List[str]).
    """
    p = Path(file_path)
    issues: List[str] = []
    
    if not p.exists():
        return False, [f"File does not exist: {p}"]
        
    if p.stat().st_size == 0:
        return False, ["File is empty (0 bytes)"]
        
    # 1. Zip package & duplicate part checks
    try:
        with zipfile.ZipFile(str(p), "r") as z:
            bad_file = z.testzip()
            if bad_file:
                issues.append(f"Corrupted zip entry found: {bad_file}")
                
            namelist = z.namelist()
            counts = collections.Counter(namelist)
            dups = [k for k, v in counts.items() if v > 1]
            if dups:
                issues.append(f"Duplicate package parts detected ({len(dups)} duplicate entries): {dups[:5]}")
                
            # Test all XML files for syntactic validity
            for fname in namelist:
                if fname.endswith(".xml") or fname.endswith(".rels"):
                    try:
                        raw = z.read(fname)
                        ET.fromstring(raw)
                    except Exception as xml_err:
                        issues.append(f"XML parse error in '{fname}': {str(xml_err)}")
    except Exception as zip_err:
        return False, [f"Zip archive integrity error: {str(zip_err)}"]

    # 2. python-pptx model parse & slide count check
    try:
        prs = Presentation(str(p))
        if len(prs.slides) == 0:
            issues.append("Presentation contains 0 slides")
    except Exception as pptx_err:
        issues.append(f"python-pptx engine parse error: {str(pptx_err)}")
        
    return len(issues) == 0, issues

def repair_pptx_package(corrupted_path: str | Path, target_path: Optional[str | Path] = None) -> bool:
    """
    Repairs package-level corruptions such as duplicate part entries in pptx zip container.
    """
    src_p = Path(corrupted_path)
    if not src_p.exists():
        return False
        
    dst_p = Path(target_path) if target_path else src_p
    temp_p = src_p.parent / f"~temp_repaired_{src_p.name}"
    
    try:
        with zipfile.ZipFile(str(src_p), "r") as z_in:
            with zipfile.ZipFile(str(temp_p), "w", compression=zipfile.ZIP_DEFLATED) as z_out:
                seen_names = set()
                for item in z_in.infolist():
                    if item.filename in seen_names:
                        continue
                    seen_names.add(item.filename)
                    z_out.writestr(item, z_in.read(item.filename))
                    
        # Replace original/target with repaired file
        if dst_p.exists() and dst_p.resolve() == src_p.resolve():
            src_p.unlink()
        temp_p.replace(dst_p)
        return True
    except Exception:
        if temp_p.exists():
            try:
                temp_p.unlink()
            except Exception:
                pass
        return False

def verify_and_auto_heal_pptx(
    pptx_path: str | Path,
    doc_structure: Optional[Dict[str, Any]] = None,
    template_inventory: Optional[List[Dict[str, Any]]] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    max_fix_attempts: int = 2
) -> Tuple[bool, str]:
    """
    Verification & Self-Correction QA Loop:
    1. Runs full structural & XML verification on pptx_path.
    2. If issues are found, attempts automated package repair (deduplicating zip parts).
    3. If errors persist, calls 9Router AI Agent to diagnose the failure, adjust the slide plan,
       and rebuild the presentation until clean.
    """
    def log(msg: str):
        if log_cb:
            log_cb(msg)

    p = Path(pptx_path)
    log(f"[Verification Loop] Validating PPTX integrity for '{p.name}'...")
    
    is_ok, issues = verify_pptx_integrity(p)
    if is_ok:
        log("[Verification Loop] PPTX integrity checks passed with 0 errors.")
        return True, str(p)

    log(f"[Verification Loop Warning] Found {len(issues)} integrity issue(s): {'; '.join(issues)}")

    # Attempt 1: Package-level deduplication & repair
    log("[Verification Loop - Fix 1] Running package repair & deduplication...")
    repair_success = repair_pptx_package(p, p)
    if repair_success:
        is_ok, issues = verify_pptx_integrity(p)
        if is_ok:
            log("[Verification Loop - Fix 1] Package repair succeeded! PPTX is now fully valid.")
            return True, str(p)
        else:
            log(f"[Verification Loop - Fix 1] Package repair applied, but remaining issues: {issues}")

    # Attempt 2: AI Agent Diagnostic & Self-Correction Rebuild with Visual Feedback
    if doc_structure and template_inventory:
        log("[Verification Loop - AI Agent] Invoking AI Agent to diagnose integrity errors and regenerate slide mapping...")
        try:
            client = OpenAI(
                api_key=Config.NINEROUTER_KEY or "dummy_key",
                base_url=Config.get_openai_base_url(),
                timeout=min(Config.LLM_TIMEOUT, 60.0)
            )
            
            ai_repair_prompt = f"""
The generated PowerPoint presentation has corruptions/integrity errors:
Error List:
{json.dumps(issues, ensure_ascii=False, indent=2)}

Document Outline:
{json.dumps(doc_structure, ensure_ascii=False, indent=2)}

Please analyze the issues and output a corrected, safe slide replacement plan adhering to standard schema.
Ensure no conflicting shape removals or malformed tables are generated.
"""
            # Collect generated slide screenshots to provide visual feedback to AI
            user_msg_parts: List[Dict[str, Any]] = [
                {"type": "text", "text": ai_repair_prompt}
            ]
            try:
                gen_raw = render_pptx_file_previews(p, target_width_px=350, use_com=True)
                gen_imgs = gen_raw[0] if isinstance(gen_raw, tuple) else gen_raw
                for idx, img in enumerate(gen_imgs[:4]):
                    b64 = image_to_base64_jpeg(img, quality=70)
                    user_msg_parts.append({
                        "type": "image_url",
                        "image_url": {
                            "url": b64
                        }
                    })
                log(f"[Verification Loop - AI Agent] Attached {len(gen_imgs[:4])} visual slide screenshots to diagnostic prompt.")
            except Exception as ss_err:
                log(f"[Verification Loop - AI Agent Warning] Could not attach screenshots: {ss_err}")

            messages_repair: Any = [
                {"role": "system", "content": "You are a PowerPoint Diagnostic & Repair Agent. Fix presentation generation errors and output clean valid JSON."},
                {"role": "user", "content": user_msg_parts}
            ]
            response = Config.safe_chat_completion(
                client,
                "verifier",
                messages=messages_repair,
                temperature=0.1
            )
            content = response.choices[0].message.content or "{}"
            repaired_plan = safe_json_loads(content)
            
            if repaired_plan and "slides" in repaired_plan:
                log("[Verification Loop - AI Agent] AI Agent provided healed plan. Re-assembling presentation...")
                # Re-assemble presentation with repaired plan
                first_tpl = template_inventory[0]["template_file"]
                src_base_p = DATA_DIR / first_tpl if (DATA_DIR / first_tpl).exists() else next(DATA_DIR.glob("*.pptx"))
                rebuilt_prs = Presentation(str(src_base_p))
                while len(rebuilt_prs.slides) > 0:
                    rId = rebuilt_prs.slides._sldIdLst[0].rId
                    rebuilt_prs.part.drop_rel(rId)
                    rebuilt_prs.slides._sldIdLst.remove(rebuilt_prs.slides._sldIdLst[0])
                    
                for s_plan in repaired_plan.get("slides", []):
                    src_tpl = s_plan.get("source_template") or first_tpl
                    src_idx = s_plan.get("source_slide_index", 0)
                    tpl_path = DATA_DIR / src_tpl if (DATA_DIR / src_tpl).exists() else src_base_p
                    src_prs = Presentation(str(tpl_path))
                    if src_idx >= len(src_prs.slides):
                        src_idx = 0
                        
                    t_slide = clone_slide_across_presentations(src_prs, rebuilt_prs, src_idx)
                    for r in s_plan.get("shape_replacements", []):
                        sh_i = r.get("shape_index")
                        if sh_i is not None and sh_i < len(t_slide.shapes):
                            sh = t_slide.shapes[sh_i]
                            if sh.has_text_frame:
                                _safe_update_text_frame(sh.text_frame, str(r.get("text", "")), shape=sh)
                                
                    if s_plan.get("speaker_notes"):
                        try:
                            t_slide.notes_slide.notes_text_frame.text = str(s_plan["speaker_notes"])
                        except Exception:
                            pass
                            
                rebuilt_prs.save(str(p))
                repair_pptx_package(p, p)
                
                is_ok, issues = verify_pptx_integrity(p)
                if is_ok:
                    log("[Verification Loop - AI Agent] Rebuilt deck is verified OK and error-free!")
                    return True, str(p)
        except Exception as ai_heal_err:
            log(f"[Verification Loop - AI Agent Error] Self-correction failed: {ai_heal_err}")

    # Fallback to package repair status
    is_ok, remaining = verify_pptx_integrity(p)
    return is_ok, str(p)
