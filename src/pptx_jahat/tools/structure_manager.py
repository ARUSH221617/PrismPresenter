import os
import re
import json
import time
import base64
from pathlib import Path
from typing import Dict, Any, List, Optional, Callable, Sequence
from openai import OpenAI

from pptx_jahat.config import Config, DATA_DIR, STRUCTURES_DIR
from pptx_jahat.tools.json_parser import safe_json_loads
from pptx_jahat.tools.pptx_engine import inspect_template_slides, classify_slide_archetype
from pptx_jahat.tools.template_analyzer import load_notes, get_analyzed_templates

SAMPLE_STRUCTURE_FILE = STRUCTURES_DIR / "pptx-structure-yosefzadeh.md"

def list_structure_files() -> List[Dict[str, Any]]:
    """
    Returns a list of all structure.md files available in data/structure/.
    """
    STRUCTURES_DIR.mkdir(parents=True, exist_ok=True)
    files = sorted(list(STRUCTURES_DIR.glob("*.md")), key=lambda p: p.stat().st_mtime, reverse=True)
    
    result = []
    for f in files:
        try:
            stat = f.stat()
            size_kb = round(stat.st_size / 1024, 1)
            mtime_str = time.strftime("%Y-%m-%d %H:%M", time.localtime(stat.st_mtime))
            
            # Read first few lines for title and excerpt
            excerpt = ""
            title = f.stem
            try:
                with open(f, "r", encoding="utf-8") as rf:
                    lines = [rf.readline() for _ in range(6)]
                    for line in lines:
                        if line.startswith("# "):
                            title = line.replace("#", "").strip()
                            break
                    excerpt = "".join(lines[:3]).strip()
            except Exception:
                pass
                
            result.append({
                "filename": f.name,
                "name": f.stem,
                "title": title,
                "path": str(f.resolve()),
                "size": f"{size_kb} KB",
                "modified": mtime_str,
                "is_sample": f.name == "pptx-structure-yosefzadeh.md",
                "excerpt": excerpt
            })
        except Exception:
            pass
            
    return result

def get_structure_content(name_or_filename: str) -> str:
    """
    Reads the content of a structure.md file from data/structure/.
    """
    STRUCTURES_DIR.mkdir(parents=True, exist_ok=True)
    clean_name = Path(name_or_filename).name
    if not clean_name.endswith(".md"):
        clean_name += ".md"
        
    target = STRUCTURES_DIR / clean_name
    if not target.exists():
        # Fallback check for exact name
        target = STRUCTURES_DIR / name_or_filename
        if not target.exists():
            raise FileNotFoundError(f"Structure file not found: {name_or_filename}")
            
    with open(target, "r", encoding="utf-8") as f:
        return f.read()

def save_structure_content(name_or_filename: str, content: str) -> Dict[str, Any]:
    """
    Saves or updates a structure.md specification in data/structure/.
    """
    STRUCTURES_DIR.mkdir(parents=True, exist_ok=True)
    clean_name = Path(name_or_filename).name
    if not clean_name.endswith(".md"):
        clean_name += ".md"
        
    target = STRUCTURES_DIR / clean_name
    with open(target, "w", encoding="utf-8") as f:
        f.write(content.strip() + "\n")
        
    stat = target.stat()
    return {
        "success": True,
        "filename": target.name,
        "name": target.stem,
        "path": str(target.resolve()),
        "size": f"{round(stat.st_size / 1024, 1)} KB",
        "modified": time.strftime("%Y-%m-%d %H:%M", time.localtime(stat.st_mtime))
    }

def delete_structure_file(name_or_filename: str) -> bool:
    """
    Deletes a structure.md file from data/structure/.
    """
    clean_name = Path(name_or_filename).name
    if not clean_name.endswith(".md"):
        clean_name += ".md"
        
    target = STRUCTURES_DIR / clean_name
    if target.exists() and clean_name != "pptx-structure-yosefzadeh.md":
        target.unlink()
        return True
    return False

def build_structure_from_template(
    template_name_or_path: str | Path,
    structure_name: str,
    custom_instructions: Optional[str] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    timeout: Optional[float] = None
) -> Dict[str, Any]:
    """
    AI Agent analyzes a PPTX template's layouts, placeholders, archetypes, and design rules,
    and synthesizes a comprehensive Slide Storyboard Specification Schema (`structure.md`).
    Saves the file to data/structure/<structure_name>.md.
    """
    def log(msg: str):
        if log_cb:
            log_cb(msg)

    STRUCTURES_DIR.mkdir(parents=True, exist_ok=True)
    effective_timeout = float(timeout or Config.LLM_TIMEOUT)
    
    # Resolve template path
    tpl_path = Path(template_name_or_path)
    if not tpl_path.exists():
        tpl_path = DATA_DIR / tpl_path.name
    if not tpl_path.exists():
        # Fallback to first existing template
        existing = list(DATA_DIR.glob("*.pptx"))
        if not existing:
            raise FileNotFoundError(f"Template '{template_name_or_path}' not found.")
        tpl_path = existing[0]

    clean_structure_name = Path(structure_name).stem.strip()
    if not clean_structure_name:
        clean_structure_name = f"{tpl_path.stem}-structure"
    output_filename = f"{clean_structure_name}.md"
    output_path = STRUCTURES_DIR / output_filename

    log(f"[*] Analyzing template layouts and archetypes: {tpl_path.name}...")
    slides_info = inspect_template_slides(tpl_path, include_screenshots=False)
    
    # Sample structure context
    sample_spec = ""
    if SAMPLE_STRUCTURE_FILE.exists():
        try:
            with open(SAMPLE_STRUCTURE_FILE, "r", encoding="utf-8") as sf:
                sample_spec = sf.read()
        except Exception:
            pass

    notes_map = get_analyzed_templates()
    template_note = notes_map.get(tpl_path.name, {})

    log(f"[*] Found {len(slides_info)} slide layouts in {tpl_path.name}.")
    log(f"[*] Calling 9Router AI to generate storyboard structure specification '{output_filename}' (timeout={effective_timeout}s)...")

    client = OpenAI(
        api_key=Config.NINEROUTER_KEY or "dummy_key",
        base_url=f"{Config.NINEROUTER_URL.rstrip('/')}/v1",
        timeout=effective_timeout
    )

    system_prompt = (
        "You are an expert Presentation Instructional Designer and Slide Architect. "
        "Your task is to build a comprehensive, production-ready Slide Storyboard Specification Schema (`structure.md`) "
        "tailored to a PowerPoint template and the user's instructional/domain needs.\n\n"
        "You must follow the established schema design demonstrated in the reference sample:\n"
        "1. Document Metadata (Grade_Level, Chapter, Lesson_ID, Page_Number, Slide_Range)\n"
        "2. Layout / Quadrant Mapping (e.g. Quadrant_TR, Quadrant_TL, Quadrant_BR, Quadrant_BL or template layout positions)\n"
        "3. Slide Content Schema for each position:\n"
        "   - Slide Number & Title\n"
        "   - Quadrant / Position indicator\n"
        "   - Slide_Type (Agenda | Theory / Definition | Classification | Worked Example | Exercise | Summary)\n"
        "   - Core_Concept (Brief 1-line summary)\n"
        "   - Animation Sequence (Strictly numbered step indicators: ①, ②, ③, ④...)\n"
        "   - Mathematical / Core Structured Elements (Definitions / Rules, Formulas / Equations in LaTeX, Sets / Matrices / Key points)\n"
        "   - Visual / Non-Textual Annotations (Trees / Diagrams, Eliminations [crossed out], Confirmations [checkmarked], Key Invariants [boxed], Teacher Callouts)\n"
        "4. Concrete Transcription Example applying the schema to realistic domain slides.\n\n"
        "Return the complete Markdown document directly."
    )

    prompt = f"""
Reference Sample Specification (from data/structure/pptx-structure-yosefzadeh.md):
```markdown
{sample_spec[:2000]}
```

Target Template Information:
- Template Name: {tpl_path.name}
- Template Purpose: {template_note.get('purpose', 'General instructional / business presentation')}
- Template Style: {template_note.get('style', 'Modern structured presentation')}
- Slide Archetypes in Template:
{json.dumps([{"slide_index": s.get("slide_index"), "layout": s.get("layout_name"), "archetype": s.get("archetype"), "slots": len(s.get("text_slots", []))} for s in slides_info[:10]], indent=2)}

Requested Structure Name: {clean_structure_name}
User Custom Instructions: {custom_instructions or "Create an optimal slide storyboard specification matching this template's visual rhythm and pedagogical flow."}

Generate the full Markdown content for `{output_filename}`. Ensure it has all 4 sections clearly structured with actionable guidelines for the presentation rewriting agent.
"""

    try:
        response = client.chat.completions.create(
            model=Config.NINEROUTER_CHAT_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3
        )
        generated_md = response.choices[0].message.content or ""
        log("[✓] Storyboard structure schema generated by AI.")
    except Exception as ex:
        log(f"[!] AI generation failed ({ex}), generating robust archetype-based fallback structure.")
        generated_md = _generate_fallback_structure_spec(tpl_path, clean_structure_name, slides_info)

    # Clean code fences if AI wrapped the entire markdown
    cleaned = generated_md.strip()
    if cleaned.startswith("```markdown"):
        cleaned = cleaned[len("```markdown"):].strip()
    elif cleaned.startswith("```"):
        cleaned = cleaned[3:].strip()
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3].strip()

    with open(output_path, "w", encoding="utf-8") as wf:
        wf.write(cleaned + "\n")

    log(f"[✓] Structure specification saved to: {output_path}")

    return {
        "success": True,
        "name": clean_structure_name,
        "filename": output_filename,
        "path": str(output_path.resolve()),
        "content": cleaned
    }

def build_structure_from_sample_files(
    sample_file_paths: Sequence[str | Path],
    structure_name: str,
    custom_instructions: Optional[str] = None,
    log_cb: Optional[Callable[[str], None]] = None,
    timeout: Optional[float] = None
) -> Dict[str, Any]:
    """
    Analyzes multiple sample documents or photos (such as handwritten worksheets, paper notes,
    math problem sheets, or textbook pages) and generates a complete Slide Storyboard & Detection
    Specification Schema (`structure.md`) tailored to extract and transcribe slides from that material.
    Saves the specification into data/structure/<structure_name>.md.
    """
    def log(msg: str):
        if log_cb:
            log_cb(msg)

    STRUCTURES_DIR.mkdir(parents=True, exist_ok=True)
    effective_timeout = float(timeout or Config.LLM_TIMEOUT)

    clean_structure_name = Path(structure_name).stem.strip()
    if not clean_structure_name:
        clean_structure_name = f"structure-{int(time.time())}"
    output_filename = f"{clean_structure_name}.md"
    output_path = STRUCTURES_DIR / output_filename

    valid_paths = [Path(p) for p in sample_file_paths if Path(p).exists()]
    if not valid_paths:
        raise FileNotFoundError("No valid sample files provided for structure analysis.")

    log(f"[*] Ingesting {len(valid_paths)} sample file(s) for structure analysis...")

    # Load reference schema from sample if available
    sample_spec = ""
    if SAMPLE_STRUCTURE_FILE.exists():
        try:
            with open(SAMPLE_STRUCTURE_FILE, "r", encoding="utf-8") as sf:
                sample_spec = sf.read()
        except Exception:
            pass

    # Build multimodal user content
    image_exts = {".png", ".jpg", ".jpeg", ".webp"}
    attached_images_count = 0
    user_content: List[Dict[str, Any]] = []

    files_summary = []
    text_samples = []

    for p in valid_paths:
        ext = p.suffix.lower()
        files_summary.append({"filename": p.name, "size_kb": round(p.stat().st_size / 1024, 1)})
        if ext in image_exts and attached_images_count < 4:
            try:
                with open(p, "rb") as imf:
                    b64 = base64.b64encode(imf.read()).decode("utf-8")
                mime = "image/jpeg" if ext in (".jpg", ".jpeg") else "image/png"
                user_content.append({
                    "type": "image_url",
                    "image_url": {"url": f"data:{mime};base64,{b64}"}
                })
                attached_images_count += 1
                log(f"[✓] Attached sample visual page: {p.name}")
            except Exception as e:
                log(f"[!] Warning reading image {p.name}: {e}")
        elif ext in (".txt", ".md", ".json"):
            try:
                with open(p, "r", encoding="utf-8", errors="replace") as tf:
                    text_samples.append(f"--- Sample File: {p.name} ---\n" + tf.read()[:2000])
            except Exception:
                pass
        elif ext == ".docx":
            try:
                import docx
                d = docx.Document(str(p))
                d_paras = [pr.text.strip() for pr in d.paragraphs if pr.text.strip()][:15]
                text_samples.append(f"--- Sample Docx: {p.name} ---\n" + "\n".join(d_paras))
            except Exception:
                pass

    log(f"[*] Calling 9Router AI Model '{Config.NINEROUTER_CHAT_MODEL}' to analyze {len(valid_paths)} sample(s) (timeout={effective_timeout}s)...")

    client = OpenAI(
        api_key=Config.NINEROUTER_KEY or "dummy_key",
        base_url=f"{Config.NINEROUTER_URL.rstrip('/')}/v1",
        timeout=effective_timeout
    )

    system_prompt = (
        "You are an expert Presentation Instructional Designer, Visual OCR Architect, and Slide Storyboard Engineer. "
        "Your task is to analyze user-provided sample documents (e.g. photos of handwritten paper sheets, "
        "worksheet pages, 4-quadrant slide sheets, formulas, or lecture notes) and produce a complete, production-ready "
        "Slide Storyboard & Detection Specification Schema (`structure.md`).\n\n"
        "This specification will be used by an AI Vision & Parsing Agent to detect, extract, and structure slides from similar documents.\n\n"
        "You must follow the schema design demonstrated in the reference sample (`pptx-structure-yosefzadeh.md`):\n"
        "1. Document Metadata (Subject / Grade_Level, Chapter, Lesson_ID, Page_Number, Slide_Range)\n"
        "2. Quadrant / Layout Mapping (e.g. Quadrant_TR, Quadrant_TL, Quadrant_BR, Quadrant_BL or reading order breakdown)\n"
        "3. Slide Content Schema (for each position: Slide Number & Title, Quadrant, Slide_Type, Core_Concept, "
        "Animation Sequence strictly with circled step indicators ①, ②, ③, ④..., Mathematical Elements in LaTeX, "
        "and Visual Annotations: trees/diagrams, crossed-out eliminations, checkmarked confirmations, boxed invariants, teacher callouts like 'اگه دقت کنی!', 'نکته')\n"
        "4. Concrete Transcription Example directly transcribing and modeling one of the provided sample sheets/topics into structured slides!\n\n"
        "Return the complete Markdown document directly."
    )

    joined_text_samples = "\n".join(text_samples)
    json_files_summary = json.dumps(files_summary, indent=2)

    prompt_text = f"""
Reference Sample Specification (`pptx-structure-yosefzadeh.md`):
```markdown
{sample_spec[:2000]}
```

Sample Files Provided by User:
{json_files_summary}

Additional Text Samples:
{joined_text_samples}

Requested Structure Name: {clean_structure_name}
User Custom Guidance: {custom_instructions or "Analyze the visual layout, handwriting/text patterns, quadrant structure, step numbers, formulas, and pedagogical annotations to build the complete detection schema."}

Generate the full Markdown content for `{output_filename}`. Ensure it provides clear detection rules and a concrete transcription example based on the samples.
"""

    user_content.insert(0, {"type": "text", "text": prompt_text})

    try:
        response = client.chat.completions.create(
            model=Config.NINEROUTER_CHAT_MODEL,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_content}
            ],
            temperature=0.25,
            max_tokens=min(Config.get_model_metadata().get("max_tokens", 65536), 32768)
        )
        generated_md = response.choices[0].message.content or ""
        log("[✓] Storyboard detection structure successfully generated by AI from sample files.")
    except Exception as ex:
        log(f"[!] AI generation notice ({ex}). Generating archetype detection structure.")
        generated_md = _generate_fallback_sample_structure_spec(clean_structure_name, valid_paths)

    cleaned = generated_md.strip()
    if cleaned.startswith("```markdown"):
        cleaned = cleaned[len("```markdown"):].strip()
    elif cleaned.startswith("```"):
        cleaned = cleaned[3:].strip()
    if cleaned.endswith("```"):
        cleaned = cleaned[:-3].strip()

    with open(output_path, "w", encoding="utf-8") as wf:
        wf.write(cleaned + "\n")

    log(f"[✓] Structure specification saved to: {output_path}")

    return {
        "success": True,
        "name": clean_structure_name,
        "filename": output_filename,
        "path": str(output_path.resolve()),
        "content": cleaned
    }

def _generate_fallback_sample_structure_spec(structure_name: str, valid_paths: List[Path]) -> str:
    samples_str = ", ".join([p.name for p in valid_paths[:4]])
    return f"""# Slide Storyboard Specification Schema (`{structure_name}.md`)
Tailored for document detection from samples: {samples_str}.

## 1. Document Metadata
- **Subject / Topic**: [e.g. Mathematics, Science, Technical Paper]
- **Chapter / Module**: [Topic / Chapter Name]
- **Lesson_ID / Unit**: [e.g. Lesson 1]
- **Page_Number**: [Integer]
- **Slide_Range**: [e.g., Slides 1–4 per page]

---

## 2. Quadrant Layout Mapping
Each physical page or worksheet is partitioned into four slide quadrants in reading order:
- `Quadrant_TR` (Top-Right): Slide $N$ (Conceptual Introduction & Definitions)
- `Quadrant_TL` (Top-Left): Slide $N+1$ (Classification & Step-by-Step Procedure)
- `Quadrant_BR` (Bottom-Right): Slide $N+2$ (Worked Example & Application)
- `Quadrant_BL` (Bottom-Left): Slide $N+3$ (Special Cases, Parity Invariants & Exercises)

---

## 3. Slide Content Schema

For each quadrant, extract data using the following structure:

### Slide [Slide_Number]: [Slide Title]
- **Quadrant**: [TR | TL | BR | BL]
- **Slide_Type**: [Agenda | Theory / Definition | Classification | Worked Example | Exercise]
- **Core_Concept**: [Brief 1-line summary]

#### Animation Sequence (Numbered Steps)
Transcribe elements strictly following circled step indicators (`①`, `②`, `③`, `④`...):
1. **Step 1 (`①`)**: [Primary concept statement / formula]
2. **Step 2 (`②`)**: [Elaboration / rule application]
3. **Step 3 (`③`)**: [Intermediate calculation or deduction]
4. **Step 4 (`④`)**: [Key conclusion or final result]

#### Mathematical Elements
- **Definitions / Rules**: [Formal mathematical rule stated]
- **Formulas / Equations**: [Standard LaTeX formatting, e.g., $a \\times b = c$]
- **Factorizations / Sets**: [e.g., $24 = 2^3 \\times 3$]

#### Visual / Non-Textual Annotations
- **Trees / Diagrams**:
  - `Root`: [Number or Category]
  - `Branches`: [Branch 1, Branch 2]
- **Eliminations (Crossed Out)**: [List terms crossed out and reasoning]
- **Confirmations (Checkmarked)**: [List terms validated with `✓`]
- **Key Invariants (Boxed)**: [List boxed final answers or invariants]
- **Teacher Callouts**: [Notes prefixed with `اگه دقت کنی!`, `نکته`]

---

## 4. Transcription Example (Applied to Sample Sheet)

### Slide 1: Main Concept Definition
- **Quadrant**: TR
- **Slide_Type**: Theory / Definition
- **Core_Concept**: Definition and fundamental property.

#### Animation Sequence
1. **Step 1 (`①`)**: بیان تعریف دقیق مفهوم و ارائه دامنه مسئله.
2. **Step 2 (`②`)**: اگه دقت کنی! بررسی شرط‌های اولیه و ویژگی‌های پایدار.
3. **Step 3 (`③`)**: محاسبه گام‌به‌گام با استفاده از روابط ریاضی.

#### Visual / Non-Textual Annotations
- **Key Invariants**: Box around final solution formula.
- **Teacher Callouts**: `نکته کلیدی`: رعایت ترتیب عملیات الزامی است.
"""

def _generate_fallback_structure_spec(tpl_path: Path, structure_name: str, slides_info: List[Dict[str, Any]]) -> str:
    """Fallback generator for structure.md when LLM API is unavailable."""
    return f"""# Slide Storyboard Specification Schema (`{structure_name}.md`)
Based on template `{tpl_path.name}`.

## 1. Document Metadata
- **Subject / Domain**: [e.g. Mathematics, Science, Business, Technology]
- **Chapter / Topic**: [e.g. Topic Name]
- **Lesson_ID / Unit**: [e.g. Unit 1]
- **Page_Number**: [Integer]
- **Slide_Range**: [e.g. Slides 1–{min(len(slides_info), 6)}]

---

## 2. Quadrant & Layout Mapping
Each topic unit is structured into balanced slide segments matching template `{tpl_path.name}`:
- `Quadrant_TR` (Top-Right): Slide $N$ (Conceptual Overview & Core Definition)
- `Quadrant_TL` (Top-Left): Slide $N+1$ (Classification & Step-by-Step Procedure)
- `Quadrant_BR` (Bottom-Right): Slide $N+2$ (Worked Examples & Practical Application)
- `Quadrant_BL` (Bottom-Left): Slide $N+3$ (Summary, Invariants & Critical Exercise)

---

## 3. Slide Content Schema

For each slide, extract and format content according to this schema:

### Slide [Slide_Number]: [Slide Title]
- **Quadrant**: [TR | TL | BR | BL]
- **Slide_Type**: [Agenda | Theory / Definition | Classification | Worked Example | Exercise | Summary]
- **Core_Concept**: [Brief 1-line summary]

#### Animation Sequence (Numbered Steps)
Transcribe elements strictly following circled step indicators (`①`, `②`, `③`, `④`...):
1. **Step 1 (`①`)**: [Primary concept introduction / premise]
2. **Step 2 (`②`)**: [Elaboration / formula / rule application]
3. **Step 3 (`③`)**: [Detailed inference or intermediate result]
4. **Step 4 (`④`)**: [Key conclusion or invariant verification]

#### Mathematical & Structured Elements
- **Definitions / Rules**: [Formal rule or definition stated clearly]
- **Formulas / Equations**: [Standard LaTeX formatting, e.g., $E = mc^2$ or $a^2 + b^2 = c^2$]
- **Factorizations / Sets**: [e.g., Sets, matrices, or itemized factors]

#### Visual / Non-Textual Annotations
- **Trees / Diagrams**:
  - `Root`: [Main entity or category]
  - `Branches`: [Branch 1, Branch 2]
- **Eliminations (Crossed Out)**: [List terms excluded and reasoning]
- **Confirmations (Checkmarked)**: [List terms confirmed with `✓`]
- **Key Invariants (Boxed)**: [List boxed final answers or invariants]
- **Teacher Callouts**: [Notes prefixed with `نکته`, `اگه دقت کنی!`, `Important Note`]

---

## 4. Transcription Example

### Slide 1: Introduction & Foundation
- **Quadrant**: TR
- **Slide_Type**: Theory / Definition
- **Core_Concept**: Core principle definition and fundamental invariant.

#### Animation Sequence
1. **Step 1 (`①`)**: بیان مفهوم اولیه و تعریف دقیق مبانی موضوع.
2. **Step 2 (`②`)**: بررسی روابط حاکم و تحلیل گام‌به‌گام.
3. **Step 3 (`③`)**: اگه دقت کنی! نتیجه نهایی همواره مستقل از متغیرهای فرعی است.

#### Visual / Non-Textual Annotations
- **Key Invariants**: Boxed primary formula.
- **Teacher Callouts**: `نکته کلیدی`: رعایت توالی مراحل الزامی است.
"""

def restructure_slides_with_agent(
    content_data: Dict[str, Any],
    structure_blueprint: str,
    log_cb: Optional[Callable[[str], None]] = None,
    timeout: Optional[float] = None
) -> Dict[str, Any]:
    """
    Autonomous AI Agent that takes extracted slide texts (from docx, pptx, text, image, audio)
    and restructures/rewrites the text strictly conforming to the selected structure.md schema.
    Returns structured slides with animation sequences (①, ②...), formulas, teacher callouts,
    and punchy presentation bullet points.
    """
    def log(msg: str):
        if log_cb:
            log_cb(msg)

    effective_timeout = float(timeout or Config.LLM_TIMEOUT)
    log(f"[*] Restructure Agent activated: Analyzing extracted content against structure.md schema (timeout={effective_timeout}s)...")

    client = OpenAI(
        api_key=Config.NINEROUTER_KEY or "dummy_key",
        base_url=f"{Config.NINEROUTER_URL.rstrip('/')}/v1",
        timeout=effective_timeout
    )

    doc_title = content_data.get("document_title", "Presentation")
    raw_sections = content_data.get("sections", [])
    raw_text = content_data.get("raw_text", "")
    sources_summary = content_data.get("sources_summary", [])

    # Format extracted material into clean, token-efficient Markdown
    content_blocks = []
    for idx, sec in enumerate(raw_sections[:15]):
        stitle = sec.get("title", f"Topic {idx+1}")
        paras = sec.get("paragraphs", [])
        bullets = sec.get("bullets", [])

        block = [f"### {stitle}"]
        for p in paras:
            p_str = str(p).strip()
            if p_str:
                block.append(p_str)
        for b in bullets:
            b_str = str(b).strip()
            if b_str:
                block.append(f"- {b_str}")
        content_blocks.append("\n".join(block))

    formatted_content = "\n\n".join(content_blocks).strip()
    if not formatted_content and raw_text:
        formatted_content = raw_text[:4000]

    system_prompt = (
        "You are an expert Presentation Content Architect and Instructional Restructuring Agent. "
        "You receive raw extracted content from user input documents (Word, PowerPoint, Text, Image OCR, or Audio Transcripts) "
        "along with a strict Slide Storyboard Specification Schema (`structure.md`).\n\n"
        "Your task is to REWRITE and REORGANIZE the input content into discrete, highly structured slides "
        "strictly conforming to the provided structure.md specification.\n\n"
        "For every slide:\n"
        "1. Assign a Quadrant/Position ('TR', 'TL', 'BR', 'BL').\n"
        "2. Assign a Slide_Type ('Theory / Definition', 'Worked Example', 'Classification', 'Exercise', 'Summary').\n"
        "3. Provide a Core_Concept (1-line punchy summary).\n"
        "4. Transcribe an Animation Sequence with strictly numbered circled steps (①, ②, ③, ④...). Each step must be concise and impactful.\n"
        "5. Extract Mathematical Elements (Definitions, LaTeX formulas, factorizations/sets).\n"
        "6. Extract Visual/Non-Textual Annotations (Diagrams, eliminations, key boxed invariants, teacher callouts like 'اگه دقت کنی!' or 'نکته').\n"
        "7. Provide refined bullet points suitable for direct insertion into slide text boxes.\n"
        "8. Provide explanatory Speaker Notes.\n\n"
        "Output valid JSON adhering strictly to the schema."
    )

    sources_label = ", ".join([s.get("filename", "") for s in sources_summary if s.get("filename")]) or "Source Files"

    user_prompt = f"""
Strict Storyboard Specification Schema (`structure.md`):
```markdown
{structure_blueprint}
```

Input Material to Restructure & Rewrite:
Document Title: {doc_title}
Sources: {sources_label}

Extracted Content:
{formatted_content}

Instructions:
1. Reorganize and rewrite the input material into structured slides conforming to the storyboard specification above.
2. Ensure animation steps use circled numbers (①, ②, ③, ④...), math elements use LaTeX, and visual annotations include teacher callouts.
3. Return valid JSON only adhering to this structure:
{{
  "document_title": "{doc_title}",
  "total_slides": 4,
  "slides": [
    {{
      "slide_number": 1,
      "title": "Slide Title",
      "quadrant": "TR",
      "slide_type": "Theory / Definition",
      "core_concept": "Brief 1-line concept",
      "animation_steps": [
        {{"step": 1, "indicator": "①", "text": "Step 1 text"}},
        {{"step": 2, "indicator": "②", "text": "Step 2 text"}}
      ],
      "mathematical_elements": {{
        "definitions": "Key definition rule",
        "formulas": "$a \\times b = c$",
        "sets": "..."
      }},
      "visual_annotations": {{
        "diagrams": "Tree or diagram notes",
        "eliminations": "Crossed out terms",
        "confirmations": "Validated items",
        "key_invariants": "Boxed key results",
        "teacher_callouts": "اگه دقت کنی! ..."
      }},
      "rewritten_bullets": [
        "① First point with key detail",
        "② Second point with formula",
        "③ Concluding insight"
      ],
      "speaker_notes": "Detailed presenter speaking notes..."
    }}
  ]
}}
"""

    log(f"[*] Dispatching prompt to 9Router AI Model '{Config.NINEROUTER_CHAT_MODEL}' (timeout={Config.LLM_TIMEOUT}s)...")
    meta = Config.get_model_metadata()
    max_output = min(meta.get("max_tokens", 65536), 32768)

    for attempt in range(2):
        try:
            response = client.chat.completions.create(
                model=Config.NINEROUTER_CHAT_MODEL,
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt if attempt == 0 else (user_prompt[:3000] + "\n\nProvide 2-4 structured slides now.")}
                ],
                temperature=0.2,
                max_tokens=max_output
            )
            content = response.choices[0].message.content or "{}"
            restructured_data = safe_json_loads(content)
            log(f"[✓] Restructure Agent successfully rewritten {len(restructured_data.get('slides', []))} slides conforming to structure.md!")
            return restructured_data
        except Exception as e:
            if attempt == 0:
                log(f"[!] Restructure Agent attempt 1 notice: {e}. Retrying with concise prompt...")
            else:
                log(f"[!] Restructure Agent fallback: {e}. Programmatically adapting content to structure schema.")
                return _fallback_restructure(content_data, structure_blueprint)

    return _fallback_restructure(content_data, structure_blueprint)

def _fallback_restructure(content_data: Dict[str, Any], structure_blueprint: str) -> Dict[str, Any]:
    """
    Programmatic fallback that maps sections into quadrant and animation step structure.
    """
    doc_title = content_data.get("document_title", "Presentation")
    raw_sections = content_data.get("sections", [])
    if not raw_sections:
        raw_sections = [{"title": doc_title, "paragraphs": ["Content imported from sources."], "bullets": []}]

    quadrants = ["TR", "TL", "BR", "BL"]
    step_icons = ["①", "②", "③", "④", "⑤", "⑥"]
    slides = []

    for idx, sec in enumerate(raw_sections):
        quad = quadrants[idx % len(quadrants)]
        all_points = sec.get("paragraphs", []) + sec.get("bullets", [])
        if not all_points:
            all_points = ["Overview and conceptual details."]

        steps = []
        rewritten_bullets = []
        for p_idx, pt in enumerate(all_points[:4]):
            icon = step_icons[p_idx % len(step_icons)]
            steps.append({
                "step": p_idx + 1,
                "indicator": icon,
                "text": pt
            })
            rewritten_bullets.append(f"{icon} {pt}")

        slides.append({
            "slide_number": idx + 1,
            "title": sec.get("title", f"Slide {idx + 1}"),
            "quadrant": quad,
            "slide_type": "Theory / Definition" if idx == 0 else "Worked Example",
            "core_concept": f"Key concept of {sec.get('title', 'Topic')}",
            "animation_steps": steps,
            "mathematical_elements": {
                "definitions": "",
                "formulas": "",
                "sets": ""
            },
            "visual_annotations": {
                "diagrams": "",
                "eliminations": "",
                "confirmations": "",
                "key_invariants": "",
                "teacher_callouts": ""
            },
            "rewritten_bullets": rewritten_bullets,
            "speaker_notes": f"Explaining {sec.get('title', 'this slide')} according to structure specification."
        })

    return {
        "document_title": doc_title,
        "total_slides": len(slides),
        "slides": slides
    }

def convert_restructured_to_sections(restructured: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Converts restructured slides into section dicts compatible with pptx_builder.
    Preserves animation steps, formulas, and visual annotations.
    """
    sections = []
    for s in restructured.get("slides", []):
        title = s.get("title", "Slide")
        bullets = s.get("rewritten_bullets", [])
        if not bullets:
            bullets = [f"{st.get('indicator', '•')} {st.get('text', '')}" for st in s.get("animation_steps", [])]

        paragraphs = []
        core = s.get("core_concept")
        if core:
            paragraphs.append(f"مفهوم محوری: {core}")
            
        math_el = s.get("mathematical_elements", {})
        if isinstance(math_el, dict) and math_el.get("formulas"):
            paragraphs.append(f"فرمول / روابط: {math_el['formulas']}")

        vis = s.get("visual_annotations", {})
        if isinstance(vis, dict) and vis.get("teacher_callouts"):
            paragraphs.append(f"نکته: {vis['teacher_callouts']}")

        sections.append({
            "title": title,
            "level": 1,
            "quadrant": s.get("quadrant", "TR"),
            "slide_type": s.get("slide_type", "Content"),
            "paragraphs": paragraphs,
            "bullets": bullets,
            "tables": [],
            "speaker_notes": s.get("speaker_notes", "")
        })
        
    return sections
