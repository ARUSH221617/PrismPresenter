import json
import re
from pathlib import Path
from typing import Dict, Any, List, Optional, Callable, Tuple
from pptx import Presentation
from openai import OpenAI

from pptx_jahat.config import Config, DATA_DIR
from pptx_jahat.tools.pptx_engine import (
    inspect_template_slides,
    classify_slide_archetype,
    extract_template_fonts,
)
from pptx_jahat.tools.preview import render_pptx_file_previews, image_to_base64_jpeg

NOTE_FILE = DATA_DIR / "NOTE.md"

STANDARD_NOTE_SCHEMA_SPEC = """# Standard Structured Template Schema for Template Intelligence Notes
Each PowerPoint template registered in `data/NOTE.md` MUST adhere to this 4-part structured specification:

## Template: <template_filename.pptx>

### 1. Template Profile & Visual Identity
- **Template File**: `<template_filename.pptx>`
- **Display Name**: <Human-readable template name>
- **Category / Domain**: <e.g. Business Strategy, Tech & Architecture, Academic Lecture, Marketing Pitch>
- **Aspect Ratio & Dimensions**: <e.g. 13.33in x 7.5in (16:9 (Widescreen))>
- **Total Slides**: <count> Slides
- **Tone & Mood**: <e.g. Professional, Clean & Balanced, High-Contrast Tech, Friendly Educational>
- **Color Theme**: <e.g. Light Theme (Clean White/Navy Accent), Dark Theme Minimalist>
- **Typography**: Primary: `<FontName>`, RTL/LTR: <RTL / LTR / Universal>
- **Content Density**: <Low-density / Medium-density / High-density>

### 2. Executive Purpose & Best Use Cases
- **Core Purpose**: <Clear 1-2 sentence statement of what this template achieves>
- **Ideal Use Cases**:
  - <Specific use case 1: e.g. Executive roadmaps and strategic initiative overviews>
  - <Specific use case 2: e.g. Step-by-step sequential workflows and onboarding pipelines>
  - <Specific use case 3: e.g. Quantitative metric scorecards and comparative benchmarks>
- **Avoid For**: <When not to use, e.g. Avoid for unstructured long-form text documents>

### 3. Slide Architecture & Slot Blueprint
| Slide Index | Archetype | Layout Pattern | Total Slots | Visual Elements | Best Content Fit |
| :---: | :--- | :--- | :---: | :--- | :--- |
| **0** | `title_cover` | Hero Center Banner | 3 | Title, Subtitle, Meta | Deck Cover, Section Title |
| **1** | `process_timeline` | 4-Step Pipeline | 8 | Step Badges, Body Cards | Sequential Steps, Chronology |
...

#### Slide Details & Slot Assignment Blueprint
##### 🔹 Slide 0 (Index 0) — `title_cover`
- **Visual Pattern**: Hero Title + Subtitle + Author/Date
- **Slot Breakdown**:
  - `Slot 0` (Title): Main Deck Headline [Max 60 chars]
  - `Slot 1` (Subtitle): Presentation Brief / Subhead [Max 120 chars]
  - `Slot 2` (Footer): Author, Organization, Date [Max 40 chars]
- **Removable Shapes**: []
- **AI Selection Intent**: Presentation title, major module transition.

##### 🔹 Slide 1 (Index 1) — `process_timeline`
- **Visual Pattern**: Horizontal 4-step progressive timeline
- **Slot Breakdown**:
  - `Slot 0` (Header): Slide Title [Max 50 chars]
  - `Slot 1-4` (Step Titles): Step 1 to 4 Headers [Max 25 chars each]
  - `Slot 5-8` (Step Bodies): Step Descriptions [Max 90 chars each]
- **Removable Shapes**: Shapes [4, 8] if only 3 steps needed
- **AI Selection Intent**: Workflow pipelines, milestone roadmaps, phased execution.

### 4. AI Generator Directives & Sequencing Recipes
- **Optimal Slide Flow Recipe**:
  - `Slide 1`: Slide 0 (`title_cover`) -> Deck Title & Topic Brief
  - `Slide 2`: Slide 1 (`process_timeline`) -> Agenda or Flow
  - `Slide 3`: Slide 2 (`content_bullets`) -> Core Definition & Context
  - `Slide 4`: Slide 7 (`metrics_stats`) -> Key KPIs & Data Evidence
  - `Slide 5`: Slide 9 (`conclusion_quote`) -> Action Items & Wrap-up
- **Slot Budget & Overflow Protection**: Respect character limits; condense paragraphs to crisp bullets.
- **Shape Pruning Strategy**: Prune optional badges/cards if document section has fewer points than template slots.
- **When AI Should Select This Template**: <Specific trigger conditions & keywords>
"""

NOTE_HEADER = """# PPTX Jahat — Template Intelligence & Design Notes
> Standardized AI Architectural Analysis of Reference PowerPoint Templates.
> Used by the AI Presentation Agent for intelligent, high-speed multi-template presentation synthesis:
> - **Step 1:** Select optimal Template(s) matching document topic, tone, color aesthetic, and purpose.
> - **Step 2:** Select optimal slide archetypes, layout patterns, and shape slots for each document section.
> - **Step 3:** Strictly conform content to slot character budgets and prune unused visual containers.

---
"""

def get_standard_note_schema() -> str:
    """Returns the official standardized template intelligence schema specification."""
    return STANDARD_NOTE_SCHEMA_SPEC.strip()

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

def get_analyzed_templates(note_path: Optional[Path] = None) -> Dict[str, Dict[str, Any]]:
    """
    Parses NOTE.md and returns a map of template_name -> structured attributes:
    - purpose, style, brief, raw_section (backward-compatible)
    - display_name, domain, tone, color_theme, typography, density, aspect_ratio, total_slides
    - slide_catalog (parsed table with index, archetype, pattern, slots, elements, best_fit)
    - flow_recipe (sequence of slide indices)
    - trigger_conditions (when AI should select)
    - is_structured (boolean flag)
    """
    content = load_notes(note_path)
    if not content:
        return {}

    # Split by "# Template: " or "## Template: "
    sections = re.split(r"(?=(?:^|\n)##?\s+Template:\s+)", content)
    result: Dict[str, Dict[str, Any]] = {}

    for sec in sections:
        sec = sec.strip()
        if not sec:
            continue
        m = re.match(r"##?\s+Template:\s+([^\n\r]+)", sec)
        if not m:
            continue
        tpl_name = m.group(1).strip().strip("`*")
        
        # Extract purpose
        purpose_match = re.search(r"(?:###?|[-*])\s*(?:🎯\s*)?Core Purpose[^\n:]*:\s*([^\n]+)", sec, re.IGNORECASE)
        if not purpose_match:
            purpose_match = re.search(r"(?:###?|[-*])\s*(?:🎯\s*)?Purpose[^\n:]*:\s*([^\n]+)", sec, re.IGNORECASE)
        if not purpose_match:
            purpose_match = re.search(r"###?\s+(?:🎯\s*)?(?:2\.\s*Executive\s+)?Purpose[^\n]*\n+[-*]?\s*([^\n]+)", sec, re.IGNORECASE)
        purpose = purpose_match.group(1).strip() if purpose_match else "Analyzed Template"
        purpose = re.sub(r"^[-*]\s*", "", purpose).replace("*", "").strip()

        # Extract style
        style_match = re.search(r"[-*]\s*\*\*Tone & Mood\*\*:\s*([^\n]+)", sec, re.IGNORECASE)
        if not style_match:
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
        brief = brief_match.group(1).strip() if brief_match else purpose
        brief = re.sub(r"^[-*]\s*", "", brief).replace("*", "").strip()

        # Structured metadata fields
        display_m = re.search(r"[-*]\s*\*\*Display Name\*\*:\s*([^\n]+)", sec, re.IGNORECASE)
        display_name = display_m.group(1).strip() if display_m else Path(tpl_name).stem

        domain_m = re.search(r"[-*]\s*\*\*Category\s*(?:/|&)?\s*Domain\*\*:\s*([^\n]+)", sec, re.IGNORECASE)
        domain = domain_m.group(1).strip() if domain_m else "General Business & Education"

        color_m = re.search(r"[-*]\s*\*\*Color Theme\*\*:\s*([^\n]+)", sec, re.IGNORECASE)
        color_theme = color_m.group(1).strip() if color_m else "Standard Presentation Theme"

        typo_m = re.search(r"[-*]\s*\*\*Typography\*\*:\s*([^\n]+)", sec, re.IGNORECASE)
        typography = typo_m.group(1).strip() if typo_m else "Standard"

        density_m = re.search(r"[-*]\s*\*\*Content Density\*\*:\s*([^\n]+)", sec, re.IGNORECASE)
        density = density_m.group(1).strip() if density_m else "Medium-density"

        dim_m = re.search(r"[-*]\s*\*\*Aspect Ratio & Dimensions\*\*:\s*([^\n]+)", sec, re.IGNORECASE)
        aspect_ratio = dim_m.group(1).strip() if dim_m else "16:9 (Widescreen)"

        slides_cnt_m = re.search(r"[-*]\s*\*\*Total Slides\*\*:\s*(\d+)", sec, re.IGNORECASE)
        total_slides_count = int(slides_cnt_m.group(1)) if slides_cnt_m else 0

        # Parse Slide Architecture Table
        slide_catalog: List[Dict[str, Any]] = []
        for line in sec.split("\n"):
            line_str = line.strip()
            if not line_str.startswith("|") or ":---" in line_str or "Layout Pattern" in line_str or "Archetype" in line_str:
                continue
            cols = [c.strip().strip("*`") for c in line_str.split("|")[1:-1]]
            if len(cols) >= 6:
                try:
                    s_idx = int(cols[0]) if cols[0].isdigit() else len(slide_catalog)
                except ValueError:
                    s_idx = len(slide_catalog)
                slot_cnt = int(cols[3]) if cols[3].isdigit() else 0
                slide_catalog.append({
                    "slide_index": s_idx,
                    "archetype": cols[1],
                    "layout_pattern": cols[2],
                    "slot_count": slot_cnt,
                    "visual_elements": cols[4],
                    "best_content_fit": cols[5]
                })

        # Flow Recipe parsing
        flow_recipe: List[int] = []
        flow_matches = re.findall(r"[-*]\s*`Slide\s*\d+`:\s*Slide\s*(\d+)", sec, re.IGNORECASE)
        if flow_matches:
            flow_recipe = [int(x) for x in flow_matches]

        # Trigger conditions
        trigger_m = re.search(r"[-*]\s*\*\*When AI Should Select This Template\*\*:\s*([^\n]+)", sec, re.IGNORECASE)
        if not trigger_m:
            trigger_m = re.search(r"[-*]\s*\*\*When to Choose:\*\*\s*([^\n]+)", sec, re.IGNORECASE)
        trigger_conditions = trigger_m.group(1).strip() if trigger_m else "Select when topic matches template style."

        is_structured = "### 1. Template Profile" in sec or "### 3. Slide Architecture" in sec

        result[tpl_name] = {
            "template_name": tpl_name,
            "display_name": display_name,
            "domain": domain,
            "purpose": purpose,
            "style": style,
            "brief": brief,
            "color_theme": color_theme,
            "typography": typography,
            "density": density,
            "aspect_ratio": aspect_ratio,
            "total_slides": total_slides_count or (len(slide_catalog) if slide_catalog else 0),
            "slide_catalog": slide_catalog,
            "flow_recipe": flow_recipe,
            "trigger_conditions": trigger_conditions,
            "is_structured": is_structured,
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

def _analyze_slide_layout_and_slots(
    slide_idx: int,
    archetype: str,
    slots: List[Dict[str, Any]],
    slide_w_in: float,
    slide_h_in: float
) -> Tuple[str, List[Dict[str, Any]], str, str, List[int]]:
    """
    Infers layout pattern, semantic slot roles, visual elements, best content fit,
    and prunable shapes for a template slide.
    """
    total_slots = len(slots)
    has_table = any(s.get("is_table") for s in slots)
    has_picture = any(s.get("is_picture_placeholder") for s in slots)

    # Detect numeric KPI callouts
    numeric_slots = []
    text_slots = []
    title_slots = []
    eyebrow_slots = []

    for s in slots:
        orig = s.get("original_text", "").strip()
        font_sz = s.get("font", {}).get("size_pt") or 14
        top_pos = s.get("bounds", {}).get("top") or 0
        left_pos = s.get("bounds", {}).get("left") or 0

        if s.get("is_title") or font_sz >= 20:
            title_slots.append(s)
        elif font_sz <= 13 and top_pos < 1800000 and len(orig) <= 25 and orig:
            eyebrow_slots.append(s)
        elif any(c.isdigit() for c in orig) and (font_sz >= 18 or "%" in orig or "$" in orig) and len(orig.split()) <= 4:
            numeric_slots.append(s)
        elif orig and not s.get("is_decorative"):
            text_slots.append(s)

    # Estimate horizontal columns from distinct left coordinates
    distinct_lefts = set()
    for s in slots:
        if not s.get("is_decorative") and s.get("bounds"):
            distinct_lefts.add(round(s["bounds"]["left"] / 914400, 1))
    col_clusters = len(distinct_lefts)

    # Determine layout pattern & best fit based on archetype and slot traits
    removable_shapes: List[int] = []

    if slide_idx == 0 or archetype == "title_cover":
        layout_pattern = "Hero Title + Subtitle Banner"
        visual_elements = "Big Title, Category Eyebrow, Presenter/Date Meta"
        best_fit = "Presentation Deck Cover, Module Title, Section Divider"
    elif archetype == "process_timeline":
        steps = max(len(numeric_slots), min(max(total_slots // 2, 3), 6))
        layout_pattern = f"Horizontal {steps}-Step Sequential Pipeline"
        visual_elements = f"{steps} Step Pills/Numbers, Arrow Flow, Milestone Cards"
        best_fit = "Sequential Workflows, Phased Roadmaps, Chronological Evolution"
        # Pruning candidate: last step slots
        if len(slots) >= 6:
            removable_shapes = [slots[-1]["shape_index"], slots[-2]["shape_index"]]
    elif archetype == "metrics_stats":
        cards = max(len(numeric_slots), min(max(total_slots // 4, 2), 6))
        layout_pattern = f"{cards}-Tile KPI Metric Dashboard Grid"
        visual_elements = f"{cards} Metric Stat Callouts, Trend Badges, KPI Labels"
        best_fit = "KPI Scorecards, Quantitative Performance, Key Data Highlights"
        if len(slots) >= 8:
            removable_shapes = [slots[-1]["shape_index"], slots[-2]["shape_index"]]
    elif archetype == "table_matrix" or has_table:
        layout_pattern = "Tabular Data Grid Matrix"
        visual_elements = "Structured Data Table, Column Headers, Row Cells"
        best_fit = "Tabular Comparisons, Multi-Attribute Specs, Financial Summaries"
    elif archetype == "multi_column":
        cols = min(max(col_clusters, 2), 4)
        layout_pattern = f"{cols}-Column Parallel Card Grid"
        visual_elements = f"{cols} Column Containers, Header Pills, Body Bullets"
        best_fit = "Multi-Pillar Comparisons, Feature Grids, Parallel Categories"
        if len(slots) >= 6:
            removable_shapes = [slots[-1]["shape_index"]]
    elif archetype == "conclusion_quote":
        layout_pattern = "Key Takeaway & Executive Quote Banner"
        visual_elements = "Large Callout Banner, Key Quote / Insight, Next Steps"
        best_fit = "Executive Summary, Closing Takeaways, Final Action Items"
    else:  # content_bullets
        if col_clusters >= 3:
            layout_pattern = f"3-Column Modular Content Cards"
            visual_elements = "3 Content Cards, Section Subheads, Detail Bullets"
            best_fit = "Multi-Card Architecture, Core Concepts, Pillar Breakdown"
            if len(slots) >= 6:
                removable_shapes = [slots[-1]["shape_index"]]
        elif col_clusters == 2:
            layout_pattern = "2-Column Split (Left/Right) Cards"
            visual_elements = "Split Dual Cards, Header, Detail Bullets"
            best_fit = "Problem vs Solution, Concept vs Application, Dual Compare"
        else:
            layout_pattern = "Structured Card & Bullet Points Container"
            visual_elements = "Main Headline, Structured Card Box, Bullet Items"
            best_fit = "Core Concepts, Explanatory Details, Problem-Solution Points"

    if has_picture:
        visual_elements += ", Image Slot"
    if has_table and "Data Table" not in visual_elements:
        visual_elements += ", Data Table"

    # Assign semantic slot roles
    slot_roles: List[Dict[str, Any]] = []
    for s in slots:
        sh_idx = s.get("shape_index", 0)
        orig = s.get("original_text", "").strip()
        budget = s.get("char_budget") or 80

        if s.get("is_title"):
            role = "Slide Title / Main Headline"
            budget = min(budget, 65)
        elif s in eyebrow_slots:
            role = "Category Eyebrow / Module Breadcrumb"
            budget = min(budget, 30)
        elif s in numeric_slots:
            role = "Metric Stat Value / KPI Number"
            budget = min(budget, 15)
        elif s.get("is_table"):
            r = s.get("table_rows", 3)
            c = s.get("table_cols", 3)
            role = f"Data Table ({r}x{c} Grid)"
        elif s.get("is_picture_placeholder"):
            role = "Visual / Picture Placeholder"
        elif "card" in s.get("shape_name", "").lower() or s.get("shape_type") == "AUTO_SHAPE":
            role = "Content Card Container"
            budget = min(budget, 120)
        else:
            role = "Body Text / Detail Bullet"
            budget = min(budget, 140)

        slot_roles.append({
            "shape_index": sh_idx,
            "role": role,
            "budget": budget,
            "sample": orig[:35] if orig else ""
        })

    return layout_pattern, slot_roles, visual_elements, best_fit, removable_shapes

def _extract_template_summary_for_ai(pptx_path: Path) -> Dict[str, Any]:
    """
    Extracts deep architectural and visual metadata from PPTX for AI agent reasoning.
    """
    prs = Presentation(str(pptx_path))
    slide_w_in = round(prs.slide_width / 914400, 2)
    slide_h_in = round(prs.slide_height / 914400, 2)
    aspect_ratio = "16:9 (Widescreen)" if abs(slide_w_in / slide_h_in - 16/9) < 0.1 else ("4:3 (Standard)" if abs(slide_w_in / slide_h_in - 4/3) < 0.1 else f"{slide_w_in}x{slide_h_in}")

    slides_data = inspect_template_slides(pptx_path, include_screenshots=False)
    template_fonts = extract_template_fonts(prs)
    primary_font = template_fonts[0] if template_fonts else "Aptos / Arial"

    archetypes_count: Dict[str, int] = {}
    slides_breakdown: List[Dict[str, Any]] = []

    for s in slides_data:
        arch = s.get("archetype", "content_bullets")
        archetypes_count[arch] = archetypes_count.get(arch, 0) + 1
        
        slots_data = s.get("text_slots", [])
        total_slots = len(slots_data)
        
        layout_pattern, slot_roles, visual_elements, best_fit, removable_shapes = _analyze_slide_layout_and_slots(
            slide_idx=s.get("slide_index", 0),
            archetype=arch,
            slots=slots_data,
            slide_w_in=slide_w_in,
            slide_h_in=slide_h_in
        )

        sample_texts = [
            slot.get("original_text", "").strip() 
            for slot in slots_data 
            if slot.get("original_text") and not slot.get("is_decorative")
        ]

        slides_breakdown.append({
            "slide_index": s.get("slide_index", 0),
            "layout_name": s.get("layout_name", f"Slide {s.get('slide_index', 0)+1}"),
            "archetype": arch,
            "layout_pattern": layout_pattern,
            "slot_count": total_slots,
            "visual_elements": visual_elements,
            "best_fit": best_fit,
            "slot_roles": slot_roles,
            "removable_shapes": removable_shapes,
            "sample_texts": sample_texts[:4]
        })

    # Infer domain & tone
    all_samples = " ".join([t for s in slides_breakdown for t in s.get("sample_texts", [])]).lower()
    is_educational = any(w in all_samples for w in ["فصل", "درس", "حل", "مسئله", "تمرین", "chapter", "lesson", "math", "theory"])
    is_technical = any(w in all_samples for w in ["system", "api", "architecture", "data", "code", "model", "server"])

    if is_educational:
        domain = "Education, Mathematics & Academic Problem Solving"
        tone = "Structured, Methodical & Highly Focused"
    elif is_technical:
        domain = "Technology Architecture, Engineering & Product Systems"
        tone = "Technical Minimalist, Crisp & Modern"
    else:
        domain = "Business Strategy, Executive Briefings & Operations"
        tone = "Professional, Clean & Balanced"

    avg_slots = sum(s["slot_count"] for s in slides_breakdown) / max(len(slides_breakdown), 1)
    if avg_slots > 14:
        density = "High-density (Comprehensive matrix & data slots)"
    elif avg_slots > 6:
        density = "Medium-density (Balanced cards & visual bullets)"
    else:
        density = "Low-density (High visual impact & focus)"

    return {
        "file_name": pptx_path.name,
        "total_slides": len(prs.slides),
        "dimensions": f"{slide_w_in}in x {slide_h_in}in",
        "aspect_ratio": aspect_ratio,
        "primary_font": primary_font,
        "template_fonts": template_fonts,
        "domain": domain,
        "tone": tone,
        "density": density,
        "archetypes_summary": archetypes_count,
        "slides_breakdown": slides_breakdown
    }

def _generate_fallback_template_note(summary: Dict[str, Any]) -> str:
    """
    Generates a production-ready, standardized Structured Template Note
    adhering strictly to the 4-part specification schema.
    """
    fname = summary["file_name"]
    stem = Path(fname).stem
    total_slides = summary["total_slides"]
    dim = summary["dimensions"]
    aspect = summary["aspect_ratio"]
    primary_font = summary.get("primary_font", "Aptos / Arial")
    domain = summary.get("domain", "General Business & Educational")
    tone = summary.get("tone", "Professional, Clean & Balanced")
    density = summary.get("density", "Medium-density")

    slides_bd = summary.get("slides_breakdown", [])

    # Build Slide Architecture Table
    table_rows = [
        "| Slide Index | Archetype | Layout Pattern | Total Slots | Visual Elements | Best Content Fit |",
        "| :---: | :--- | :--- | :---: | :--- | :--- |"
    ]
    for s in slides_bd:
        idx = s["slide_index"]
        arch = s["archetype"]
        pat = s["layout_pattern"]
        slots = s["slot_count"]
        elems = s["visual_elements"]
        fit = s["best_fit"]
        table_rows.append(f"| **{idx}** | `{arch}` | {pat} | {slots} | {elems} | {fit} |")
    table_md = "\n".join(table_rows)

    # Build Slide Details & Slot Assignment Blueprint
    slide_details_blocks = []
    for s in slides_bd[:10]:
        idx = s["slide_index"]
        arch = s["archetype"]
        pat = s["layout_pattern"]
        fit = s["best_fit"]
        removable = s.get("removable_shapes", [])
        rem_str = f"Shapes {removable} (prunable if fewer items)" if removable else "None"

        roles_list = []
        for r in s.get("slot_roles", [])[:6]:
            sample_part = f' (e.g. "{r["sample"]}")' if r.get("sample") else ""
            roles_list.append(f"  - `Slot {r['shape_index']}`: **{r['role']}** [Max {r['budget']} chars]{sample_part}")

        roles_md = "\n".join(roles_list) if roles_list else "  - Default content slots"

        slide_details_blocks.append(f"""##### 🔹 Slide {idx} (Index {idx}) — `{arch}`
- **Visual Pattern**: {pat}
- **Slot Breakdown**:
{roles_md}
- **Removable Shapes**: {rem_str}
- **AI Selection Intent**: {fit}.""")

    slide_details_md = "\n\n".join(slide_details_blocks)

    # Build Flow Recipe
    recipe_steps = []
    if slides_bd:
        recipe_steps.append(f"  - `Slide 1`: Slide 0 (`{slides_bd[0]['archetype']}`) -> Presentation Title & Topic Overview")
    if len(slides_bd) > 1:
        recipe_steps.append(f"  - `Slide 2`: Slide 1 (`{slides_bd[1]['archetype']}`) -> Agenda, Roadmap or Process Overview")
    if len(slides_bd) > 2:
        recipe_steps.append(f"  - `Slide 3`: Slide 2 (`{slides_bd[2]['archetype']}`) -> Core Concepts & Problem Framing")
    if len(slides_bd) > 4:
        recipe_steps.append(f"  - `Slide 4`: Slide 4 (`{slides_bd[4]['archetype']}`) -> In-depth Analysis, Metrics or Data")
    if len(slides_bd) > 1:
        last_s = slides_bd[-1]
        recipe_steps.append(f"  - `Slide 5`: Slide {last_s['slide_index']} (`{last_s['archetype']}`) -> Action Items & Wrap-up")
    recipe_md = "\n".join(recipe_steps)

    return f"""## Template: {fname}

### 1. Template Profile & Visual Identity
- **Template File**: `{fname}`
- **Display Name**: {stem.upper()} — {tone.split(',')[0]} Slide Deck
- **Category / Domain**: {domain}
- **Aspect Ratio & Dimensions**: {dim} ({aspect})
- **Total Slides**: {total_slides} Slides
- **Tone & Mood**: {tone}
- **Color Theme**: Light Theme (High-contrast clean presentation with accent containers)
- **Typography**: Primary: `{primary_font}`, RTL/LTR: RTL & Universal Multilingual
- **Content Density**: {density}

### 2. Executive Purpose & Best Use Cases
- **Core Purpose**: Reliable, highly structured slide deck engineered for clear instructional and strategic communications.
- **Ideal Use Cases**:
  - {domain} presentations and comprehensive briefings.
  - Step-by-step modular breakdowns and structured multi-column analysis.
  - Quantitative scorecards and milestone progress tracking.
- **Avoid For**: Avoid using for unstructured, paragraph-heavy prose with no visual hierarchy.

### 3. Slide Architecture & Slot Blueprint
{table_md}

#### Slide Details & Slot Assignment Blueprint
{slide_details_md}

### 4. AI Generator Directives & Sequencing Recipes
- **Optimal Slide Flow Recipe**:
{recipe_md}
- **Slot Budget & Overflow Protection**: Strictly enforce character budgets for each slot to prevent text clipping and container overflows.
- **Shape Pruning Strategy**: Prune optional card containers or trailing milestone badges when document sections have fewer items.
- **When AI Should Select This Template**: Prioritize `{fname}` when the document domain aligns with {domain} and requires a {tone.lower()} aesthetic.
"""

def analyze_template(
    pptx_path: Path | str,
    log_cb: Optional[Callable[[str], None]] = None,
    save_to_file: bool = True,
    timeout: Optional[float] = None
) -> Dict[str, Any]:
    """
    Analyzes a single PPTX template using 9Router AI Agent.
    Generates a structured NOTE following the standardized 4-part specification:
    1. Template Profile & Visual Identity
    2. Executive Purpose & Best Use Cases
    3. Slide Architecture & Slot Blueprint (with markdown catalog & slot roles)
    4. AI Generator Directives & Sequencing Recipes
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

    effective_timeout = float(timeout or Config.LLM_TIMEOUT)
    client = OpenAI(
        api_key=Config.NINEROUTER_KEY or "dummy_key",
        base_url=Config.get_openai_base_url(),
        timeout=effective_timeout
    )

    system_prompt = (
        "You are an expert Presentation Art Director, UX Design Architect, and AI Presentation Engine Strategist. "
        "Your task is to analyze a PowerPoint presentation template (.pptx) and produce a standardized, "
        "highly structured Template Intelligence Note.\n\n"
        "This NOTE will be saved into `data/NOTE.md` and directly consulted by an Autonomous AI Agent to:\n"
        "1. Step 1: Select the best template matching document domain, tone, and audience.\n"
        "2. Step 2: Select the optimal slide index and layout pattern for each document section.\n"
        "3. Step 3: Replace text in exact shape slots while respecting character budgets.\n\n"
        "You MUST strictly follow this exact 4-part Structured Template Schema:\n"
        "## Template: <filename>\n"
        "### 1. Template Profile & Visual Identity\n"
        "- **Template File**: `<filename>`\n"
        "- **Display Name**: <Name>\n"
        "- **Category / Domain**: <Domain>\n"
        "- **Aspect Ratio & Dimensions**: <Dimensions> (<Aspect>)\n"
        "- **Total Slides**: <Count> Slides\n"
        "- **Tone & Mood**: <Tone>\n"
        "- **Color Theme**: <Color Palette & Contrast>\n"
        "- **Typography**: Primary: `<Font>`, RTL/LTR: <Direction>\n"
        "- **Content Density**: <Density>\n\n"
        "### 2. Executive Purpose & Best Use Cases\n"
        "- **Core Purpose**: <Statement>\n"
        "- **Ideal Use Cases**:\n  - <Case 1>\n  - <Case 2>\n"
        "- **Avoid For**: <Guidance>\n\n"
        "### 3. Slide Architecture & Slot Blueprint\n"
        "| Slide Index | Archetype | Layout Pattern | Total Slots | Visual Elements | Best Content Fit |\n"
        "| :---: | :--- | :--- | :---: | :--- | :--- |\n"
        "<Populate rows for all slides>\n\n"
        "#### Slide Details & Slot Assignment Blueprint\n"
        "For key slides, detail the visual pattern, slot roles (title, eyebrow, column header, metric, body), character budgets, and removable shapes.\n\n"
        "### 4. AI Generator Directives & Sequencing Recipes\n"
        "- **Optimal Slide Flow Recipe**: Ordered list of slide indices for a standard 5-slide deck.\n"
        "- **Slot Budget & Overflow Protection**: Rules for char limits.\n"
        "- **Shape Pruning Strategy**: Which shapes to remove when content has fewer items.\n"
        "- **When AI Should Select This Template**: Trigger conditions.\n"
    )

    user_content: List[Dict[str, Any]] = [
        {
            "type": "text",
            "text": f"""Analyze this PowerPoint presentation template and generate its complete Structured Template Note:

Template Metadata & Slide Breakdown:
{json.dumps(summary, ensure_ascii=False, indent=2)}

Please write the complete, standardized Template Intelligence Note adhering to the 4-part schema for `{path.name}`.
Ensure the Slide Architecture table, slot character budgets, and AI selection directives are precise.
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

    a_model = Config.get_agent_model("analyzer")
    a_think = Config.get_agent_think_level("analyzer")
    log(f"[*] Calling 9Router AI Model '{a_model}' (thinking: {a_think}) to analyze {path.name}...")

    try:
        response = Config.safe_chat_completion(
            client,
            "analyzer",
            model=a_model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content}
            ],
            temperature=0.2
        )
        note_text = response.choices[0].message.content or ""
        log(f"[✓] AI analysis completed for {path.name}.")
    except Exception as e:
        log(f"[!] AI API call failed: {e}. Generating standardized structured fallback note.")
        note_text = _generate_fallback_template_note(summary)

    if save_to_file:
        log(f"[*] Saving structured note to {NOTE_FILE}...")
        update_template_note_in_file(path.name, note_text, NOTE_FILE)
        log(f"[✓] Successfully updated {NOTE_FILE}.")

    # Parse key attributes for UI summary
    parsed_map = get_analyzed_templates()
    tpl_info = parsed_map.get(path.name, {
        "purpose": "General Presentation",
        "style": summary.get("tone", "Standard"),
        "brief": "Multi-slide presentation",
        "display_name": path.stem,
        "domain": summary.get("domain", "General"),
        "color_theme": "Standard",
        "typography": summary.get("primary_font", "Standard"),
        "density": summary.get("density", "Medium-density"),
        "slide_catalog": []
    })

    return {
        "template_name": path.name,
        "display_name": tpl_info.get("display_name", path.stem),
        "domain": tpl_info.get("domain", "General"),
        "total_slides": summary["total_slides"],
        "purpose": tpl_info["purpose"],
        "style": tpl_info["style"],
        "brief": tpl_info["brief"],
        "color_theme": tpl_info.get("color_theme", "Standard"),
        "typography": tpl_info.get("typography", "Standard"),
        "density": tpl_info.get("density", "Medium-density"),
        "slide_catalog": tpl_info.get("slide_catalog", []),
        "is_structured": tpl_info.get("is_structured", True),
        "note_markdown": note_text
    }

def format_notes_for_ai_prompt(
    template_names: Optional[List[str]] = None,
    note_path: Optional[Path] = None
) -> str:
    """
    Compiles structured template intelligence into an ultra-high-signal, compact
    knowledge block for the AI PPTX presentation synthesis agent in `pptx_builder.py`.
    """
    analyzed_map = get_analyzed_templates(note_path)
    if not analyzed_map:
        raw = load_notes(note_path)
        return raw.strip() if raw else ""

    target_templates = template_names or list(analyzed_map.keys())
    entries = []

    for name in target_templates:
        info = analyzed_map.get(name)
        if not info:
            continue

        catalog_summary = []
        for sl in info.get("slide_catalog", []):
            catalog_summary.append(
                f"  - Slide {sl['slide_index']}: `{sl['archetype']}` | {sl['layout_pattern']} | {sl['slot_count']} slots | Best for: {sl['best_content_fit']}"
            )
        catalog_str = "\n".join(catalog_summary) if catalog_summary else "  - Slide inventory available in blueprint."

        recipe = info.get("flow_recipe", [])
        recipe_str = f"Recipe: {recipe}" if recipe else ""

        entry = f"""### Template: {info['template_name']} ({info.get('display_name', info['template_name'])})
- **Domain & Style**: {info.get('domain', 'General')} | Tone: {info.get('style', 'Standard')} | Density: {info.get('density', 'Medium')}
- **Purpose**: {info.get('purpose', '')}
- **Best Fit When**: {info.get('trigger_conditions', 'Topic matches domain')}
- **Slide Layout Catalog**:
{catalog_str}
{recipe_str}
"""
        entries.append(entry.strip())

    if not entries:
        return load_notes(note_path)

    return "\n\n".join(entries)

def generate_blank_structured_note(template_name: str, pptx_path: Optional[Path] = None) -> str:
    """
    Generates a pre-filled structured note boilerplate following the 4-part schema.
    If pptx_path is provided, inspects the file to populate real slide archetypes.
    """
    target = pptx_path or (DATA_DIR / template_name)
    if target.exists():
        summary = _extract_template_summary_for_ai(target)
        return _generate_fallback_template_note(summary)

    stem = Path(template_name).stem
    return f"""## Template: {template_name}

### 1. Template Profile & Visual Identity
- **Template File**: `{template_name}`
- **Display Name**: {stem.upper()} — Presentation Deck
- **Category / Domain**: Business Strategy & Executive Briefings
- **Aspect Ratio & Dimensions**: 13.33in x 7.5in (16:9 (Widescreen))
- **Total Slides**: 10 Slides
- **Tone & Mood**: Professional, Clean & Balanced
- **Color Theme**: Light Theme (Clean White/Navy Accent)
- **Typography**: Primary: `Aptos / Arial`, RTL/LTR: RTL & Universal Multilingual
- **Content Density**: Medium-density

### 2. Executive Purpose & Best Use Cases
- **Core Purpose**: Clear, high-impact corporate presentation deck with balanced typography and card containers.
- **Ideal Use Cases**:
  - Executive strategy presentations and quarterly briefings.
  - Phased project roadmaps and sequential workflows.
  - Key performance indicators and comparative benchmark analysis.
- **Avoid For**: Avoid using for unbroken walls of raw text or dense academic papers without visual structure.

### 3. Slide Architecture & Slot Blueprint
| Slide Index | Archetype | Layout Pattern | Total Slots | Visual Elements | Best Content Fit |
| :---: | :--- | :--- | :---: | :--- | :--- |
| **0** | `title_cover` | Hero Center Banner | 3 | Title, Subtitle, Meta | Deck Cover, Section Title |
| **1** | `process_timeline` | 4-Step Pipeline | 8 | Step Badges, Body Cards | Sequential Steps, Chronology |
| **2** | `content_bullets` | 2-Column Split Cards | 6 | Split Cards, Subheads | Core Theory, Problem vs Solution |
| **3** | `metrics_stats` | 4-Tile KPI Grid | 12 | Metric Values, Labels | KPI Performance, Key Numbers |
| **4** | `conclusion_quote` | Takeaway Banner | 4 | Large Callout, Next Steps | Summary, Action Items |

#### Slide Details & Slot Assignment Blueprint
##### 🔹 Slide 0 (Index 0) — `title_cover`
- **Visual Pattern**: Hero Title + Subtitle + Presenter Meta
- **Slot Breakdown**:
  - `Slot 0`: **Slide Title / Main Headline** [Max 60 chars]
  - `Slot 1`: **Category Eyebrow / Breadcrumb** [Max 30 chars]
  - `Slot 2`: **Presenter Meta / Date** [Max 40 chars]
- **Removable Shapes**: None
- **AI Selection Intent**: Presentation Deck Cover, Module Title, Section Divider.

### 4. AI Generator Directives & Sequencing Recipes
- **Optimal Slide Flow Recipe**:
  - `Slide 1`: Slide 0 (`title_cover`) -> Title & Presentation Overview
  - `Slide 2`: Slide 1 (`process_timeline`) -> Agenda or Workflow
  - `Slide 3`: Slide 2 (`content_bullets`) -> Core Problem & Framing
  - `Slide 4`: Slide 3 (`metrics_stats`) -> Key Results & KPIs
  - `Slide 5`: Slide 4 (`conclusion_quote`) -> Action Items & Wrap-up
- **Slot Budget & Overflow Protection**: Respect character limits; condense paragraphs to crisp bullets.
- **Shape Pruning Strategy**: Prune optional badges/cards if document section has fewer points than template slots.
- **When AI Should Select This Template**: Select when the presentation requires a clean, dependable corporate structure.
"""

def analyze_all_templates(
    data_dir: Optional[Path] = None,
    progress_cb: Optional[Callable[[int, int, str], None]] = None,
    log_cb: Optional[Callable[[str], None]] = None
) -> str:
    """
    Scans data directory for all PPTX templates and analyzes them sequentially.
    Updates data/NOTE.md after each template with structured notes.
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

