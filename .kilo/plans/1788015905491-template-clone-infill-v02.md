# Implementation Plan: Exact Template Slide Cloning & AI Text Infill (v0.2)

## 1. Overview & Goal
Refactor the Docx-to-PPTX pipeline so that rather than constructing abstract geometric shapes from scratch, the system performs a **direct 4-step template cloning and in-place AI text infill**:
- **Step 1:** AI Agent / engine reads the selected base PPTX template (`data/*.pptx`), extracting its exact slide sequence, layout roles, shapes, placeholder IDs, and original sample texts.
- **Step 2:** AI Agent reads and parses the input Word document (`.docx`), extracting headings, problem statements, paragraphs, formulas, steps, bullet points, and tables.
- **Step 3:** AI Agent reasons over the template structure and document content to write contextual text replacements mapped specifically to each slide and shape index in the base presentation.
- **Step 4:** The generator opens a duplicate of the base PPTX presentation and updates the text frames directly, preserving 100% of the original visual design, geometry, positioning, typography, colors, animations, and decorative assets.
- **Step 5 (Documentation):** Update `PRD.md` to reflect version **v0.2**.

---

## 2. Architecture & Data Flow

```
[Selected Base PPTX]                  [Uploaded Word .docx]
        |                                       |
        v (Step 1)                              v (Step 2)
[Inspect PPTX Slides & Shapes]         [Parse Sections & Content]
        \                                       /
         \                                     /
          v                                   v
       +-----------------------------------------+
       |   Step 3: 9Router AI Agent Reasoning    |
       |  - Matches doc sections to slide slots  |
       |  - Writes fitted text per shape/card    |
       +-----------------------------------------+
                            |
                            v (Step 4)
       +-----------------------------------------+
       | Clone Base PPTX Presentation File       |
       | In-place update of text frames & runs   |
       | (Preserve all shapes, styles, geometry) |
       +-----------------------------------------+
                            |
                            v
               [Output Generated PPTX Deck]
```

---

## 3. Detailed Step-by-Step Execution Tasks

### Task 1: Enhance PPTX Structure Inspector (`src/pptx_jahat/tools/pptx_engine.py`)
- Provide a clean helper `inspect_template_slides(pptx_path: Path) -> List[Dict[str, Any]]`:
  - Enumerate each slide index.
  - Enumerate each text shape / table / placeholder with its `shape_id`, `name`, `text_sample`, `position`, `font_info`.
  - Provide a concise text representation of each slide's content slots for LLM prompt ingestion.

### Task 2: Implement Exact Template Clone & Infill Engine (`src/pptx_jahat/tools/pptx_builder.py`)
- **Step 1 (Read PPTX):** Load the base template presentation and build its structural slot inventory.
- **Step 2 (Read Docx):** Parse document sections, headings, bullets, and tables using `parse_docx`.
- **Step 3 (AI Content Mapping & Text Generation):**
  - Construct a structured prompt for 9Router LLM containing:
    - Base template slide list (with each slide's purpose and existing placeholder text slots).
    - Document content breakdown.
  - Prompt instructions: "Generate replacement text for each shape slot on each slide to teach the document's subject while respecting the layout intent of the template."
  - Return JSON mapped by slide index and shape identifier/index:
    ```json
    {
      "slides": [
        {
          "slide_index": 0,
          "shape_replacements": [
            {"shape_index": 0, "text": "عنوان فصل جدید"},
            {"shape_index": 1, "text": "توضیح زیرعنوان"}
          ]
        }
      ]
    }
    ```
- **Step 4 (In-place Slide Mutation & Saving):**
  - Open a fresh instance of the template PPTX via `Presentation(str(template_path))`.
  - For each slide and shape replacement, perform in-place text update on `shape.text_frame`.
  - Preserve run-level font attributes (family, size, bold, color) by updating text without destroying font metadata.
  - Save the modified clone to the output path.

### Task 3: Integrate with Autonomous AI Agent & GUI
- Update `src/pptx_jahat/agent.py` tool calling `build_pptx_with_agent`.
- Update `src/pptx_jahat/gui/app.py` progress logs and generator worker to reflect the 4-step template cloning pipeline.

### Task 4: Verification & Test Suite
- Run test suite with sample docx and template.
- Verify that slides in the generated PPTX are exact clones of the base template with customized text replacements.

### Task 5: Update PRD to v0.2 (`PRD.md`)
- Update PRD version to `v0.2`.
- Document the new 4-step Template Cloning & In-Place Text Infill architecture.
- Document Persian RTL text preservation and layout fidelity improvements.

---

## 4. Key Constraints & Edge Cases
1. **Mismatched Slide Counts:** If document has more sections than template slides, loop or duplicate layout patterns; if fewer, either keep remaining template slides or prune unused ones based on user preference.
2. **Font Formatting Preservation:** Modifying `text_frame.text` directly resets formatting; text updates must preserve existing paragraph/run font name, size, bold, and color properties.
3. **Persian / RTL Text Support:** Ensure right alignment and IRANYekan / B Nazanin font compatibility.
