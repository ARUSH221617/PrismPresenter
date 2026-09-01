import json
import re
from pathlib import Path
from typing import Dict, Any, List, Optional, Callable
from pptx import Presentation
from openai import OpenAI

from pptx_jahat.config import Config, DATA_DIR
from pptx_jahat.tools.pptx_engine import inspect_template_slides, classify_slide_archetype
from pptx_jahat.tools.preview import render_pptx_file_previews, image_to_base64_jpeg

NOTE_FILE = DATA_DIR / "NOTE.md"

NOTE_HEADER = """# PPTX Jahat — Template Intelligence & Design Notes
> Auto-generated AI analysis of reference PowerPoint templates for intelligent multi-template presentation synthesis.
> Used by the AI Presentation Agent to:
> - **Step 1:** Select optimal Template(s) matching document topic, narrative tone, and purpose.
> - **Step 2:** Select optimal slide archetypes and layouts for each document section.

---
"""

def load_notes(note_path: Optional[Path] = None) -> str:
    """Reads existing NOTE.md content or returns empty string if not found."""
    target = note_path or NOTE_FILE
    if target.exists():
        try:
            with open(target, "r", encoding="utf-8") as f:
                return f.read()
        except Exception:
            return ""
    return ""

def save_notes(content: str, note_path: Optional[Path] = None) -> None:
    """Saves full markdown content to NOTE.md."""
    target = note_path or NOTE_FILE
    target.parent.mkdir(parents=True, exist_ok=True)
    with open(target, "w", encoding="utf-8") as f:
        f.write(content.strip() + "\n")

def get_analyzed_templates(note_path: Optional[Path] = None) -> Dict[str, Dict[str, str]]:
    """
    Parses NOTE.md and returns a map of template_name -> {purpose, style, brief, raw_section}.
    """
    content = load_notes(note_path)
    if not content:
        return {}

    # Split by "# Template: " or "## Template: "
    sections = re.split(r"(?=(?:^|\n)##?\s+Template:\s+)", content)
    result: Dict[str, Dict[str, str]] = {}

    for sec in sections:
        sec = sec.strip()
        if not sec:
            continue
        m = re.match(r"##?\s+Template:\s+([^\n\r]+)", sec)
        if not m:
            continue
        tpl_name = m.group(1).strip().strip("`*")
        
        # Extract purpose
        purpose_match = re.search(r"(?:###?|[-*])\s*(?:🎯\s*)?Purpose[^\n:]*:\s*([^\n]+)", sec, re.IGNORECASE)
        if not purpose_match:
            purpose_match = re.search(r"###?\s+(?:🎯\s*)?Purpose[^\n]*\n+[-*]?\s*([^\n]+)", sec, re.IGNORECASE)
        purpose = purpose_match.group(1).strip() if purpose_match else "Analyzed Template"
        purpose = re.sub(r"^[-*]\s*", "", purpose).replace("*", "").strip()

        # Extract style
        style_match = re.search(r"(?:###?|[-*])\s*(?:🎨\s*)?Style[^\n:]*:\s*([^\n]+)", sec, re.IGNORECASE)
        if not style_match:
            style_match = re.search(r"###?\s+(?:🎨\s*)?Style[^\n]*\n+[-*]?\s*(?:\*\*(?:Tone|Vibe|Feel):\*\*\s*)?([^\n]+)", sec, re.IGNORECASE)
        style = style_match.group(1).strip() if style_match else "Standard"
        style = re.sub(r"^[-*]\s*", "", style).replace("*", "").strip()
        style = re.sub(r"^(?:Tone|Vibe|Feel):\s*", "", style).strip()

        # Extract brief
        brief_match = re.search(r"(?:###?|[-*])\s*(?:📝\s*)?Content Brief[^\n:]*:\s*([^\n]+)", sec, re.IGNORECASE)
        if not brief_match:
            brief_match = re.search(r"###?\s+(?:📝\s*)?Content Brief[^\n]*\n+[-*]?\s*([^\n]+)", sec, re.IGNORECASE)
        brief = brief_match.group(1).strip() if brief_match else "Presentation deck"
        brief = re.sub(r"^[-*]\s*", "", brief).replace("*", "").strip()

        result[tpl_name] = {
            "template_name": tpl_name,
            "purpose": purpose,
            "style": style,
            "brief": brief,
            "raw_section": sec
        }

    return result

def update_template_note_in_file(template_name: str, note_markdown: str, note_path: Optional[Path] = None) -> str:
    """
    Inserts or updates a specific template's note section in NOTE.md.
    Returns the updated full content.
    """
    target = note_path or NOTE_FILE
    current_content = load_notes(target)

    clean_note = note_markdown.strip()
    if not clean_note.startswith("# Template:") and not clean_note.startswith("## Template:"):
        clean_note = f"## Template: {template_name}\n\n" + clean_note

    if not current_content.strip():
        new_content = NOTE_HEADER + "\n\n" + clean_note + "\n"
        save_notes(new_content, target)
        return new_content

    # Regex search for existing template section
    pattern = rf"(^|\n)##?\s+Template:\s+{re.escape(template_name)}.*?(?=(?:\n##?\s+Template:|\Z))"
    if re.search(pattern, current_content, re.DOTALL):
        # Replace existing section
        new_content = re.sub(pattern, r"\1" + clean_note, current_content, flags=re.DOTALL)
    else:
        # Append new section
        new_content = current_content.rstrip() + "\n\n---\n\n" + clean_note + "\n"

    save_notes(new_content, target)
    return new_content

def _extract_template_summary_for_ai(pptx_path: Path) -> Dict[str, Any]:
    """
    Extracts deep architectural and visual metadata from PPTX for AI agent reasoning.
    """
    prs = Presentation(str(pptx_path))
    slide_w_in = round(prs.slide_width / 914400, 2)
    slide_h_in = round(prs.slide_height / 914400, 2)
    aspect_ratio = "16:9 (Widescreen)" if abs(slide_w_in / slide_h_in - 16/9) < 0.1 else ("4:3 (Standard)" if abs(slide_w_in / slide_h_in - 4/3) < 0.1 else f"{slide_w_in}x{slide_h_in}")

    slides_data = inspect_template_slides(pptx_path, include_screenshots=False)
    
    archetypes_count: Dict[str, int] = {}
    slides_breakdown: List[Dict[str, Any]] = []

    for s in slides_data:
        arch = s.get("archetype", "content_bullets")
        archetypes_count[arch] = archetypes_count.get(arch, 0) + 1
        
        # Collect text samples
        sample_texts = [
            slot.get("original_text", "").strip() 
            for slot in s.get("text_slots", []) 
            if slot.get("original_text") and not slot.get("is_decorative")
        ]
        
        has_table = any(slot.get("is_table") for slot in s.get("text_slots", []))
        total_slots = len(s.get("text_slots", []))

        slides_breakdown.append({
            "slide_index": s.get("slide_index"),
            "layout_name": s.get("layout_name"),
            "archetype": arch,
            "slot_count": total_slots,
            "has_table": has_table,
            "sample_texts": sample_texts[:4]
        })

    return {
        "file_name": pptx_path.name,
        "total_slides": len(prs.slides),
        "dimensions": f"{slide_w_in}in x {slide_h_in}in",
        "aspect_ratio": aspect_ratio,
        "archetypes_summary": archetypes_count,
        "slides_breakdown": slides_breakdown
    }

def analyze_template(
    pptx_path: Path | str,
    log_cb: Optional[Callable[[str], None]] = None,
    save_to_file: bool = True
) -> Dict[str, Any]:
    """
    Analyzes a single PPTX template using 9Router AI Agent.
    Generates a structured NOTE detailing:
    - Purpose & Ideal Use Cases
    - Content Brief & Structural Narrative
    - Core Idea & Visual Metaphors
    - Style, Tone & Feel (e.g. friendly, corporate, modern, dark tech)
    - Slide-by-slide Archetype Breakdown
    - AI Selection Guidelines
    Saves to data/NOTE.md if save_to_file is True.
    """
    path = Path(pptx_path)
    if not path.exists():
        raise FileNotFoundError(f"Template file not found: {pptx_path}")

    def log(msg: str):
        if log_cb:
            log_cb(msg)

    log(f"[*] Inspecting PPTX structure: {path.name}...")
    summary = _extract_template_summary_for_ai(path)
    log(f"[*] Extracted {summary['total_slides']} slides, aspect ratio: {summary['aspect_ratio']}.")

    # Render up to 6 key slide preview screenshots for multimodal vision reasoning
    log("[*] Rendering slide visual previews for Vision AI...")
    preview_images: List[Dict[str, Any]] = []
    try:
        rendered = render_pptx_file_previews(str(path), target_width_px=450)
        # Select first slide, middle slides, last slide up to 6 slides
        step = max(1, len(rendered) // 6)
        selected_indices = list(range(0, len(rendered), step))[:6]
        for idx in selected_indices:
            b64 = image_to_base64_jpeg(rendered[idx], quality=80)
            preview_images.append({
                "slide_index": idx,
                "base64": b64
            })
        log(f"[✓] Captured {len(preview_images)} visual slide snapshots for multimodal analysis.")
    except Exception as e:
        log(f"[!] Visual preview capture skipped ({e}). Proceeding with structural analysis.")

    # Build 9Router AI prompt
    client = OpenAI(
        api_key=Config.NINEROUTER_KEY or "dummy_key",
        base_url=f"{Config.NINEROUTER_URL.rstrip('/')}/v1",
        timeout=120.0
    )

    system_prompt = (
        "You are an expert Presentation Art Director, UX Design Architect, and Content Strategist. "
        "Your task is to analyze a PowerPoint presentation template (.pptx) and produce an in-depth, "
        "highly practical design and content intelligence NOTE.\n\n"
        "This NOTE will be saved into `data/NOTE.md` and directly consulted by an Autonomous AI Agent when:\n"
        "1. Step 1: Evaluating input Word documents and choosing the BEST matching template(s) by domain, tone, purpose, and visual style.\n"
        "2. Step 2: Selecting the best individual slides from this template for specific content sections.\n\n"
        "Structure your response in clear, standard Markdown with the following exact sections:\n"
        "## Template: <filename>\n"
        "### 🎯 Purpose & Best Use Cases\n"
        "### 📝 Content Brief & Narrative Flow\n"
        "### 💡 Core Concept & Visual Ideas\n"
        "### 🎨 Style, Mood & Aesthetic Feel (Explicitly describe tone like 'Friendly & Approachable', 'Executive Corporate', 'Vibrant Startup Pitch', 'Dark Minimal Tech', color palette, typography weight)\n"
        "### 📊 Slide Inventory & Archetype Highlights (Detail key slides, what layouts exist, e.g. 3-column stats, timeline, table comparison, quote, title)\n"
        "### 🤖 AI Selection Guidelines (When the AI Agent MUST pick this template over others, and exact slide mappings for common document chapters)\n"
    )

    user_content: List[Dict[str, Any]] = [
        {
            "type": "text",
            "text": f"""Analyze this PowerPoint presentation template:

Template Metadata & Slide Inventory:
{json.dumps(summary, ensure_ascii=False, indent=2)}

Please write the complete, thorough Template Intelligence Note in Markdown for `{path.name}`.
Ensure the Style and Feel (e.g. friendly, professional, bold, educational, playful) is clearly articulated so the AI Agent knows exactly when this template is the perfect match.
"""
        }
    ]

    for p in preview_images:
        user_content.append({
            "type": "image_url",
            "image_url": {
                "url": p["base64"]
            }
        })

    log(f"[*] Calling 9Router AI Model '{Config.NINEROUTER_CHAT_MODEL}' to analyze {path.name}...")

    try:
        response = client.chat.completions.create(
            model=Config.NINEROUTER_CHAT_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content}
            ],
            temperature=0.2
        )
        note_text = response.choices[0].message.content or ""
        log(f"[✓] AI analysis completed for {path.name}.")
    except Exception as e:
        log(f"[!] AI API call failed: {e}. Generating fallback structural note.")
        note_text = _generate_fallback_template_note(summary)

    if save_to_file:
        log(f"[*] Saving note to {NOTE_FILE}...")
        update_template_note_in_file(path.name, note_text, NOTE_FILE)
        log(f"[✓] Successfully updated {NOTE_FILE}.")

    # Parse key attributes for UI summary
    parsed_map = get_analyzed_templates()
    tpl_info = parsed_map.get(path.name, {
        "purpose": "General Presentation",
        "style": "Standard",
        "brief": "Multi-slide presentation"
    })

    return {
        "template_name": path.name,
        "total_slides": summary["total_slides"],
        "purpose": tpl_info["purpose"],
        "style": tpl_info["style"],
        "brief": tpl_info["brief"],
        "note_markdown": note_text
    }

def _generate_fallback_template_note(summary: Dict[str, Any]) -> str:
    """Fallback rule-based markdown note if LLM connection fails."""
    fname = summary["file_name"]
    archetypes = summary.get("archetypes_summary", {})
    arch_str = ", ".join([f"{k} ({v})" for k, v in archetypes.items()])
    
    return f"""## Template: {fname}

### 🎯 Purpose & Best Use Cases
- General business and educational slide presentation.
- Ideal for standard topic overviews, progress reports, and structured briefings.

### 📝 Content Brief & Narrative Flow
- **Total Slides:** {summary['total_slides']}
- **Slide Dimensions:** {summary['dimensions']} ({summary['aspect_ratio']})
- **Structure:** Title introduction followed by topic overview slides, content bullet points, and summary layouts.

### 💡 Core Concept & Visual Ideas
- Clean structured slide layouts with distinct content boxes and card containers.
- Archetypes present: {arch_str or 'Standard content bullets'}.

### 🎨 Style, Mood & Aesthetic Feel
- **Tone:** Professional, Clean & Balanced
- **Feel:** Structured, friendly readability, versatile color contrast.

### 📊 Slide Inventory & Archetype Highlights
{chr(10).join([f"- **Slide {s['slide_index']+1} ({s['layout_name']})**: {s['archetype']} archetype with {s['slot_count']} content slots." for s in summary.get('slides_breakdown', [])[:8]])}

### 🤖 AI Selection Guidelines
- **When to Choose:** Select when the document requires a dependable, cleanly-organized layout.
- **Slide Recommendations:** Use Slide 1 for Title/Intro, multi-slot slides for body points and metrics.
"""

def analyze_all_templates(
    data_dir: Optional[Path] = None,
    progress_cb: Optional[Callable[[int, int, str], None]] = None,
    log_cb: Optional[Callable[[str], None]] = None
) -> str:
    """
    Scans data directory for all PPTX templates and analyzes them sequentially.
    Updates data/NOTE.md after each template.
    Returns the complete compiled NOTE.md.
    """
    target_dir = data_dir or DATA_DIR
    pptx_files = sorted(list(target_dir.glob("*.pptx")))
    templates = [f for f in pptx_files if not f.name.endswith("_generated.pptx")]

    if not templates:
        if log_cb:
            log_cb("[!] No PPTX templates found in data directory.")
        return ""

    total = len(templates)
    if log_cb:
        log_cb(f"[*] Starting batch analysis of {total} templates into {NOTE_FILE}...")

    for idx, tpl in enumerate(templates):
        if progress_cb:
            progress_cb(idx + 1, total, tpl.name)
        if log_cb:
            log_cb(f"\n[{idx+1}/{total}] Processing {tpl.name}...")
        
        try:
            analyze_template(tpl, log_cb=log_cb, save_to_file=True)
        except Exception as e:
            if log_cb:
                log_cb(f"[!] Error analyzing {tpl.name}: {e}")

    if log_cb:
        log_cb(f"\n[✓] All {total} templates successfully analyzed and saved to {NOTE_FILE}.")

    return load_notes(NOTE_FILE)
