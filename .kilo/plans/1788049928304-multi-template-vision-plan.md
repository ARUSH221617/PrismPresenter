# Implementation Plan: Multi-Template Global PPTX Engine & Vision-Guided Slide Infill

## 1. Goal
Upgrade the PPTX Jahat generation system so that:
1. **Multi-Template Global Scan:** The user does not select a single template. The system scans and understands **all** templates in `data/*.pptx`.
2. **Visual Screenshot Analysis for AI:** Before applying document content to presentation slides, the system captures screenshots/rendered previews of all candidate slides across all templates. These images (along with shape slot metadata) are sent to the 9Router Vision/LLM agent (`ag/gemini-3.7-flash-high` supports multimodal image input).
3. **AI Reasoning & Selection:** The AI agent analyzes the visual structure and textual slots of each slide image to match document sections, decide which exact template slide to pick for each section, and determine whether specific objects/shapes should be removed or replaced based on the incoming text content.
4. **Slide Assembly & Infill:** Assemble the selected exact slides from their source PPTX files into the final deck using XML slide cloning, delete any unwanted shapes identified by AI, and replace text/tables in-place while strictly preserving original formatting and layout styles.

---

## 2. Architectural Design & Pipeline

```
[All data/*.pptx Templates]                     [Uploaded Word .docx]
         │                                                │
         ▼ (Step 1)                                       ▼ (Step 2)
[Inspect & Render Candidate Slides]              [Parse Doc Structure]
 - Extract shape slots & metadata                 - Sections, headings,
 - Generate visual screenshots per slide             paragraphs, tables
         │                                                │
         └───────────────────────┬────────────────────────┘
                                 │
                                 ▼
         +────────────────────────────────────────────────+
         │ Step 3: Vision-Enabled 9Router AI Agent       │
         │ - Receives slide screenshot images + slot info │
         │ - Receives full parsed docx content            │
         │ - Selects best slide per section               │
         │ - Decides text replacements, tables, and       │
         │   shapes to remove (shape_index / shape_id)    │
         +────────────────────────────────────────────────+
                                 │
                                 ▼
         +────────────────────────────────────────────────+
         │ Step 4: Multi-Template Deck Assembly Engine    │
         │ - Creates target deck matching aspect ratio    │
         │ - Deep clones exact selected slides across     │
         │   different PPTX template files                │
         │ - Removes pruned shapes from slide XML         │
         │ - In-place text & table updates                │
         +────────────────────────────────────────────────+
                                 │
                                 ▼
                    [Output Generated PPTX Deck]
                                 │
                                 ▼
                    [Tkinter In-App Slide Viewer]
```

---

## 3. Detailed Component Changes

### 3.1 Template Screenshot & Catalog Enhancements (`src/pptx_jahat/tools/pptx_engine.py` & `preview.py`)
- In `preview.py`, add utility to render slide images directly to base64 or temporary image paths for AI vision analysis.
- In `pptx_engine.py`, build `inspect_all_templates(data_dir: Path)`:
  - Scans all `data/*.pptx` files.
  - Generates preview screenshots for every slide across all templates.
  - Produces a unified inventory of available slides: template name, slide index, visual screenshot (base64 or image uri), slide layout type, and list of text/table/graphic shape slots with coordinates and font info.

### 3.2 Vision-Guided AI Reasoning & Decision Engine (`src/pptx_jahat/tools/pptx_builder.py`)
- In `pptx_builder.py`, update `generate_slide_replacements_with_ai`:
  - Pass the slide screenshots as multimodal image parts (`data:image/png;base64,...`) alongside the structured slot metadata and docx outline to 9Router (`ag/gemini-3.7-flash-high`).
  - System prompt instructs AI to visually analyze each slide's aesthetic, layout flow, and shape density.
  - Return JSON schema:
    ```json
    {
      "deck_title": "Presentation Title",
      "slides": [
        {
          "source_template": "T711.pptx",
          "source_slide_index": 0,
          "target_section": "Introduction",
          "shape_replacements": [
            {
              "shape_index": 0,
              "text": "New Slide Title"
            },
            {
              "shape_index": 1,
              "text": "Fitted body bullet points..."
            }
          ],
          "shapes_to_remove": [2, 3],
          "table_replacements": [
            {
              "shape_index": 4,
              "table_data": [["Header 1", "Header 2"], ["Row 1", "Val 1"]]
            }
          ]
        }
      ]
    }
    ```

### 3.3 Multi-Template Slide Cloning & Shape Removal Engine (`src/pptx_jahat/tools/pptx_builder.py`)
- Implement `clone_slide_across_presentations(source_prs, target_prs, slide_index)`:
  - Performs slide-level XML cloning and relationship importing (images, media, layout references).
  - Maintains aspect ratio and geometry fidelity.
- Implement `_remove_shape(slide, shape_index)`:
  - Deletes shape element directly from slide `spTree` XML element.
- Implement text in-place update using `_safe_update_text_frame` preserving fonts, RTL direction, sizes, and colors.

### 3.4 GUI & Autonomous Agent Integration (`src/pptx_jahat/gui/app.py`, `agent.py`)
- Update GUI Tab 1:
  - Remove mandatory single-template selection or set default to "All Templates (AI Auto-Selection across data/*.pptx)".
  - Add visual status indicators showing: "Scanning all templates", "Rendering slide screenshots", "AI Vision analyzing slides & matching content", "Cloning & infilling slides".
  - Preview displays the newly assembled multi-template presentation.
- Update Autonomous Agent tool descriptions in `agent.py` so `build_pptx_from_docx` reflects multi-template and vision capabilities.

---

## 4. Verification & Validation Plan
1. **Unit Verification (`test_suite.py`):**
   - Verify `inspect_all_templates` discovers all `data/*.pptx` files and generates slide screenshots.
   - Verify `clone_slide_across_presentations` correctly copies slides from multiple distinct templates (e.g. Slide 0 from `T711.pptx` + Slide 2 from `T712.pptx` + Slide 1 from `T718.pptx`) into a unified presentation.
   - Verify `_remove_shape` successfully removes designated shape elements without corrupting the PPTX XML.
2. **End-to-End Test:**
   - Execute `build_pptx_with_agent` with `sample_document.docx` against the full `data/` directory.
   - Validate output presentation opens cleanly, contains slides chosen from across templates, has correctly replaced texts and pruned unused shapes.
3. **GUI Verification:**
   - Ensure in-app slide preview displays the generated presentation smoothly.
