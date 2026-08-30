import json
import re
import copy
from pathlib import Path
from typing import Dict, Any, List, Optional, Callable
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from pptx_jahat.config import Config, DATA_DIR, OUTPUT_DIR
from pptx_jahat.tools.docx_parser import parse_docx
from pptx_jahat.tools.pptx_engine import inspect_template_slides, inspect_all_templates
from openai import OpenAI

def _safe_update_text_frame(tf: Any, new_text: str, is_rtl: bool = True) -> None:
    """
    Updates text in a text_frame while preserving paragraph and run-level formatting
    such as font family, size, bold, italic, and colors.
    """
    if not tf:
        return
        
    lines = [line for line in new_text.split("\n") if line.strip()]
    if not lines:
        lines = [new_text]
        
    # Capture style of first run if available
    saved_font = {
        "name": None,
        "size": None,
        "bold": None,
        "italic": None,
        "color": None
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
    except Exception:
        pass

    # Clear old paragraphs and populate with new text lines
    tf.clear()
    
    for idx, line in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = line
        if is_rtl:
            p.alignment = PP_ALIGN.RIGHT
            
        # Apply preserved font styling to runs
        if p.runs and any(saved_font.values()):
            for run in p.runs:
                if saved_font["name"]: run.font.name = saved_font["name"]
                if saved_font["size"]: run.font.size = saved_font["size"]
                if saved_font["bold"] is not None: run.font.bold = saved_font["bold"]
                if saved_font["italic"] is not None: run.font.italic = saved_font["italic"]
                if saved_font["color"]: run.font.color.rgb = saved_font["color"]

def _remove_shape(slide: Any, shape_index: int) -> bool:
    """
    Safely removes a shape from slide XML by shape_index.
    """
    try:
        if 0 <= shape_index < len(slide.shapes):
            shape = slide.shapes[shape_index]
            sp_elem = shape._element
            parent = sp_elem.getparent()
            if parent is not None:
                parent.remove(sp_elem)
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

def clone_slide_across_presentations(source_prs: Presentation, target_prs: Presentation, slide_index: int) -> Any:
    """
    Deep clones a slide from source_prs into target_prs, preserving layout, media parts,
    and relationship mappings.
    """
    src_slide = source_prs.slides[slide_index]
    
    # Choose layout from target_prs matching source or fallback to blank layout
    layout_name = src_slide.slide_layout.name if src_slide.slide_layout else "Blank"
    matching_layout = None
    for layout in target_prs.slide_layouts:
        if layout.name == layout_name:
            matching_layout = layout
            break
    if not matching_layout:
        matching_layout = target_prs.slide_layouts[min(6, len(target_prs.slide_layouts) - 1)]
        
    target_slide = target_prs.slides.add_slide(matching_layout)
    
    # Copy relationships & media/picture parts
    src_part = src_slide.part
    target_part = target_slide.part
    
    for rel_id, rel in src_part.rels.items():
        if rel.reltype == RT.SLIDE_LAYOUT:
            continue
        try:
            if rel.is_external:
                target_part.rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
            else:
                target_part.relate_to(rel.target_part, rel.reltype)
        except Exception:
            pass
            
    # Replace target slide's spTree (shape tree) with deep copied source spTree
    target_spTree = target_slide.shapes._spTree
    src_spTree = src_slide.shapes._spTree
    
    # Remove default shapes in the newly added slide
    for child in list(target_spTree):
        target_spTree.remove(child)
        
    # Copy all children from source spTree into target spTree
    for child in list(src_spTree):
        target_spTree.append(copy.deepcopy(child))
        
    return target_slide

def generate_slide_replacements_with_ai(
    template_inventory: List[Dict[str, Any]],
    doc_structure: Dict[str, Any],
    log_cb: Optional[Callable[[str], None]] = None
) -> Dict[str, Any]:
    """
    Step 3: AI Vision Agent reasons over multimodal slide screenshots, shape slots across templates,
    and Word docx content.
    Returns optimal slide selections across templates, exact text replacements, and shapes to remove.
    """
    def log(msg: str):
        if log_cb:
            log_cb(msg)

    log("[Step 3] AI Vision Agent analyzing candidate template slides & document content...")

    client = OpenAI(
        api_key=Config.NINEROUTER_KEY or "dummy_key",
        base_url=f"{Config.NINEROUTER_URL.rstrip('/')}/v1"
    )

    system_prompt = (
        "You are an expert Presentation Art Director and Content Producer. "
        "You receive visual screenshots and shape slot metadata of candidate presentation slides across multiple templates, "
        "along with a parsed Word document. "
        "Your task is to:\n"
        "1. Select the best visual slide from the available templates for each section/topic in the document (Title slide, Section slides, Metric/Table slides, Conclusion).\n"
        "2. Match the document content into the chosen slide's shape slots (shape_index).\n"
        "3. Identify any unnecessary or overflowing shape indices to delete (shapes_to_remove).\n"
        "4. If a slot contains a table, supply updated 2D table_data.\n"
        "Strictly return valid JSON adhering to the specified schema."
    )

    # Prepare slot descriptions (without heavy base64 strings in the JSON text prompt)
    inventory_summary = []
    for s in template_inventory:
        summary_entry = {
            "template_file": s.get("template_file"),
            "slide_index": s.get("slide_index"),
            "layout_name": s.get("layout_name"),
            "text_slots": [
                {
                    "shape_index": slot.get("shape_index"),
                    "shape_name": slot.get("shape_name"),
                    "shape_type": slot.get("shape_type"),
                    "original_text": slot.get("original_text", "")[:120],
                    "is_title": slot.get("is_title", False),
                    "is_table": slot.get("is_table", False),
                    "table_shape": f"{slot.get('table_rows')}x{slot.get('table_cols')}" if slot.get("is_table") else None,
                    "is_decorative": slot.get("is_decorative", False)
                }
                for slot in s.get("text_slots", [])
            ]
        }
        inventory_summary.append(summary_entry)

    # Build multimodal user message content array
    user_content: List[Dict[str, Any]] = [
        {
            "type": "text",
            "text": f"""
Step 1 - Available Slide Blueprints & Slots:
{json.dumps(inventory_summary, ensure_ascii=False, indent=2)}

Step 2 - Input Word Document Outline & Content:
{json.dumps(doc_structure, ensure_ascii=False, indent=2)}

Visual Screenshots of Candidate Template Slides:
(See attached images corresponding to the candidate slides above)

Instructions:
1. Construct a cohesive presentation sequence matching the document flow (Title slide, Content/Topic slides, Metric slides, Summary).
2. For each slide in your output deck, specify:
   - "source_template": Name of template file (e.g. "T711.pptx", "t1.pptx", "sample_template.pptx")
   - "source_slide_index": Index of slide in that template
   - "target_section": Name of document section this slide covers
   - "shape_replacements": List of {{"shape_index": int, "text": str}} mapping new text into slots
   - "shapes_to_remove": List of shape indices [int] that should be pruned/deleted from the slide
   - "table_replacements": List of {{"shape_index": int, "table_data": [["cell", ...], ...]}}

Return a JSON object with this exact schema:
{{
  "deck_title": "Presentation Title",
  "slides": [
    {{
      "source_template": "sample_template.pptx",
      "source_slide_index": 0,
      "target_section": "Document Title",
      "shape_replacements": [
        {{
          "shape_index": 0,
          "text": "Presentation Title"
        }}
      ],
      "shapes_to_remove": [],
      "table_replacements": []
    }}
  ]
}}
"""
        }
    ]

    # Attach slide screenshots as multimodal image parts (limit to top 15 candidate slides to preserve tokens)
    attached_count = 0
    for s in template_inventory:
        b64 = s.get("screenshot_base64")
        if b64 and attached_count < 15:
            user_content.append({
                "type": "image_url",
                "image_url": {
                    "url": b64
                }
            })
            attached_count += 1

    log(f"[Step 3] Sending prompt with {attached_count} visual slide previews to 9Router AI...")

    try:
        response = client.chat.completions.create(
            model=Config.NINEROUTER_CHAT_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content}
            ],
            temperature=0.3
        )
        content = response.choices[0].message.content or "{}"
        match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", content, re.DOTALL)
        if match:
            json_str = match.group(1)
        else:
            json_str = content.strip()
            
        plan = json.loads(json_str)
        log("[Step 3] AI Vision Agent generated presentation plan successfully.")
        return plan
    except Exception as e:
        log(f"[Step 3 Warning] AI reasoning exception ({e}), using fallback multi-template mapper.")
        return None

def build_pptx_with_agent(
    docx_path: str | Path,
    output_path: Optional[str | Path] = None,
    template_name: Optional[str] = None,
    log_callback: Optional[Callable[[str], None]] = None
) -> str:
    """
    4-Step Vision-Guided Multi-Template Presentation Generation:
    Step 1: Scan & inspect candidate slides across all templates with rendered screenshots.
    Step 2: Read and parse Word DOCX structure.
    Step 3: Vision AI reasons on slide screenshots & doc content, selecting best slides across templates.
    Step 4: Clone selected slides across presentations into target deck, prune removed shapes, and update text in-place.
    """
    def log(msg: str):
        if log_callback:
            log_callback(msg)

    # ----------------------------------------------------
    # Step 1: Scan & inspect templates
    # ----------------------------------------------------
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
        raise FileNotFoundError("No PPTX templates found in data folder.")

    log(f"[Step 1] Loaded {len(template_inventory)} candidate slides across templates.")

    # ----------------------------------------------------
    # Step 2: Read and parse Word Document
    # ----------------------------------------------------
    docx_file = Path(docx_path)
    log(f"[Step 2] Reading Word document: {docx_file.name}...")
    parsed_doc = parse_docx(docx_file)
    log(f"[Step 2] Parsed {parsed_doc['total_sections']} sections from document.")

    # ----------------------------------------------------
    # Step 3: AI Vision Agent writes texts and selects slides
    # ----------------------------------------------------
    ai_plan = generate_slide_replacements_with_ai(template_inventory, parsed_doc, log_cb=log)

    # ----------------------------------------------------
    # Step 4: Assemble target deck across presentations
    # ----------------------------------------------------
    log("[Step 4] Assembling target presentation from selected template slides...")

    # Cache opened presentations by filename
    prs_cache: Dict[str, Presentation] = {}
    def get_source_prs(tpl_file: str) -> Presentation:
        if tpl_file not in prs_cache:
            p = DATA_DIR / tpl_file
            if not p.exists():
                # Fallback to first existing template
                p = next(DATA_DIR.glob("*.pptx"))
            prs_cache[tpl_file] = Presentation(str(p))
        return prs_cache[tpl_file]

    # Pre-open first source template to create matching target presentation
    first_tpl_name = template_inventory[0]["template_file"]
    base_src_prs = get_source_prs(first_tpl_name)
    
    # Initialize target presentation with matching dimensions
    target_prs = Presentation()
    target_prs.slide_width = base_src_prs.slide_width
    target_prs.slide_height = base_src_prs.slide_height
    
    # Clear any default blank slides in newly initialized presentation
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
            
            # In-place text replacements
            replacements = {r.get("shape_index"): r.get("text") for r in s_plan.get("shape_replacements", [])}
            for shape_idx, shape in enumerate(target_slide.shapes):
                if shape_idx in replacements and shape.has_text_frame:
                    new_text = replacements[shape_idx]
                    if new_text is not None:
                        _safe_update_text_frame(shape.text_frame, str(new_text))
                        
            # Table replacements
            table_replacements = {t.get("shape_index"): t.get("table_data") for t in s_plan.get("table_replacements", [])}
            for shape_idx, shape in enumerate(target_slide.shapes):
                if shape_idx in table_replacements and shape.has_table:
                    tdata = table_replacements[shape_idx]
                    if tdata:
                        for r_i, row in enumerate(tdata):
                            if r_i < len(shape.table.rows):
                                for c_i, cell_val in enumerate(row):
                                    if c_i < len(shape.table.columns):
                                        shape.table.cell(r_i, c_i).text = str(cell_val)
                                        
            # Remove unwanted shapes after text/table replacements to avoid index shifts during replacement
            shapes_to_remove = s_plan.get("shapes_to_remove", [])
            if shapes_to_remove:
                _remove_shapes(target_slide, shapes_to_remove)
    else:
        # Fallback Multi-Template Assembly:
        # Title slide from first template, content slides from available candidate slides
        log("[Step 4 Fallback] Generating presentation using multi-slide assembly...")
        
        # 1. Title Slide
        title_slide = clone_slide_across_presentations(base_src_prs, target_prs, 0)
        for shape in title_slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                _safe_update_text_frame(shape.text_frame, parsed_doc.get("document_title", "Presentation"))
                break
                
        # 2. Content Slides per Section
        for s_idx, section in enumerate(parsed_doc.get("sections", [])):
            candidate_entry = template_inventory[(s_idx + 1) % len(template_inventory)]
            src_prs = get_source_prs(candidate_entry["template_file"])
            src_idx = candidate_entry["slide_index"]
            
            c_slide = clone_slide_across_presentations(src_prs, target_prs, src_idx)
            text_shapes = [sh for sh in c_slide.shapes if sh.has_text_frame and sh.text_frame.text.strip()]
            
            if text_shapes:
                _safe_update_text_frame(text_shapes[0].text_frame, section.get("title", ""))
                if len(text_shapes) > 1:
                    body = "\n".join(section.get("paragraphs", []) + [f"• {b}" for b in section.get("bullets", [])])
                    _safe_update_text_frame(text_shapes[1].text_frame, body)

    if not output_path:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        output_path = OUTPUT_DIR / f"{docx_file.stem}_generated.pptx"
    else:
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        
    target_prs.save(str(output_path))
    log(f"[Step 4] Finished. Output presentation saved to: {output_path}")
    return str(output_path)
