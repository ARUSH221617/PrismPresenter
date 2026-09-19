import json
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional, Callable, Tuple, Sequence
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
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
    _is_rtl_text
)

logger = logging.getLogger("slide_verifier")


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

    # 3. Handle shape updates (text, tables, notes, images)
    for act in actions:
        atype = str(act.get("action", "")).lower()
        s_idx = act.get("slide_index")

        if s_idx is None or s_idx < 0 or s_idx >= len(prs.slides):
            continue

        slide = prs.slides[s_idx]

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
                                is_rtl=None,
                                max_box_width_emu=getattr(shape, "width", None),
                                max_box_height_emu=getattr(shape, "height", None),
                                shape=shape
                            )
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

    # Index template inventory by (template_file, slide_index)
    tpl_lookup: Dict[Tuple[str, int], Dict[str, Any]] = {}
    for entry in template_inventory:
        tfile = str(entry.get("template_file", ""))
        sidx = entry.get("slide_index", 0)
        tpl_lookup[(tfile, sidx)] = entry

    # Open source template presentations on-demand to extract screenshots if missing
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

        # If template screenshot not cached, render directly
        if not tpl_screenshot:
            tprs = get_template_prs(tpl_name)
            if tprs and 0 <= tpl_sidx < len(tprs.slides):
                try:
                    tslide = tprs.slides[tpl_sidx]
                    t_img = render_pptx_slide_to_image(tslide, tprs.slide_width, tprs.slide_height, target_width_px=screenshot_width)
                    tpl_screenshot = image_to_base64_jpeg(t_img, quality=82)
                except Exception as ex:
                    log(f"[!] Could not render template screenshot: {ex}")

        gen_screenshot = gen_previews[s_idx] if s_idx < len(gen_previews) else None

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
      - Template slide screenshot
      - Generated slide screenshot
      - Shapes / texts inventory of both
    Returns:
      {
        "is_correct": bool,
        "score": int, # 0-100
        "detected_issues": ["Issue description..."],
        "edit_structure": {
          "summary_of_changes": "...",
          "actions": [
            {"action": "update_text", "slide_index": int, "shape_index": int, "new_text": "..."}
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
        "Specifically check:\n"
        "1. Unreplaced formula text / equations or symbols from original template: If the template slide contained mathematical equations (e.g. 2n, n=6, math formulas, axis numbers) or specialized graphics that DO NOT match the generated slide's topic, YOU MUST ISSUE A DELETION ACTION TO REMOVE THEM:\n"
        "   {'action': 'delete_shape', 'slide_index': int, 'shape_index': int, 'formula_text': '2n'}\n"
        "   or {'action': 'remove_formula', 'slide_index': int, 'formula_text': '2n'}\n"
        "2. Missing extra elements or texts: Did the original template have subtitles, badges, tags, card numbers (01, 02), subheaders, or category labels that were left empty or missed in the generated slide?\n"
        "3. Unreplaced placeholder text: Are there shapes still containing default dummy text (e.g. 'Lorem ipsum', 'Sample text', 'Header Here', 'Subtitle Goes Here', 'Click to edit')?\n"
        "4. Card / Multi-column completeness: If the template has 3 or 4 feature cards, were all cards populated with meaningful content from the topic, or was one left blank?\n"
        "5. Layout balance & clipping: Are titles overflowing or clipped?\n\n"
        "If the slide is correct and well-populated, mark is_correct=true with empty actions.\n"
        "If issues are found, mark is_correct=false, list the detected_issues clearly, and formulate a precise 'actions' list in 'edit_structure' to heal the slide.\n\n"
        "Supported actions:\n"
        "- 'delete_shape': {'action': 'delete_shape', 'slide_index': int, 'shape_index': int, 'formula_text': str} (use to delete unwanted template formulas, equations, or mismatched elements)\n"
        "- 'remove_formula': {'action': 'remove_formula', 'slide_index': int, 'formula_text': str} (use to remove unreplaced math equations)\n"
        "- 'update_text': {'action': 'update_text', 'slide_index': int, 'shape_index': int, 'new_text': str}\n"
        "- 'remove_shapes': {'action': 'remove_shapes', 'slide_index': int, 'shape_indices': [int, ...]}\n"
        "- 'update_table': {'action': 'update_table', 'slide_index': int, 'shape_index': int, 'table_data': [[...]]}\n"
        "- 'update_notes': {'action': 'update_notes', 'slide_index': int, 'notes': str}\n\n"
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

    prompt_text = f"""
Slide Under Inspection: Slide {s_idx + 1}
Target Section / Topic: {slide_pair.get('target_section', 'Slide Topic')}
Source Template: {tpl_name} (Slide {tpl_sidx + 1})

Template Shapes & Sample Text:
{json.dumps(tpl_slots_summary, ensure_ascii=False, indent=2)}

Generated Slide Current Shapes & Text:
{json.dumps(gen_shapes_summary, ensure_ascii=False, indent=2)}

Speaker Notes: {slide_pair.get('speaker_notes', '')[:200]}

Instructions:
Evaluate if the generated slide accurately adapted the template without missing extra texts, subtitle slots, badges, or leaving unreplaced placeholder strings / unwanted formulas.
CRITICAL: If an unreplaced formula or math element from the template does not belong on this slide, you MUST formulate a 'delete_shape' or 'remove_formula' action to delete it!

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
    "Don't remove the bottom card, just fix its text".

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
        "You receive the current slide verification data (template screenshots, current slide shapes, detected issues, proposed edit actions) "
        "and user instructions.\n\n"
        "Your task is to:\n"
        "1. Understand the user's specific feedback or correction.\n"
        "2. Adjust or refine the detected issues and edit_structure actions to exactly match what the user wants.\n"
        "3. Provide a helpful, concise explanation of the adjustments made.\n\n"
        "Supported actions:\n"
        "- 'delete_shape': {'action': 'delete_shape', 'slide_index': int, 'shape_index': int, 'formula_text': str} (delete unwanted template formula, equation, or shape)\n"
        "- 'remove_formula': {'action': 'remove_formula', 'slide_index': int, 'formula_text': str} (remove unreplaced math formula)\n"
        "- 'update_text': {'action': 'update_text', 'slide_index': int, 'shape_index': int, 'new_text': str}\n"
        "- 'remove_shapes': {'action': 'remove_shapes', 'slide_index': int, 'shape_indices': [int, ...]}\n"
        "- 'update_table': {'action': 'update_table', 'slide_index': int, 'shape_index': int, 'table_data': [[...]]}\n"
        "- 'update_notes': {'action': 'update_notes', 'slide_index': int, 'notes': str}\n"
        "- 'delete_slide': {'action': 'delete_slide', 'slide_index': int}\n\n"
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
