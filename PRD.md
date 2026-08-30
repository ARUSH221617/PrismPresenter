# Product Requirements Document (PRD) — v0.2

**Project Name:** PPTX Jahat  
**Version:** 0.2  
**Status:** Implemented (Verified Architecture)  
**Author:** AI Agent / Product Engineering  
**Date:** 2026-08-30  

---

## 1. Executive Summary
**PPTX Jahat (v0.2)** is an AI-powered presentation generation system powered exclusively by **9Router**. In this version, the document-to-presentation pipeline adopts a **4-step exact template cloning and in-place AI text infill methodology**. Instead of recreating slide shapes from scratch, the AI Agent reads the user-selected base PowerPoint template, parses the input Word document (`.docx`), reasons over both to author fitted slide content, and clones the exact original template slides—mutating only the text frames and table data while preserving all master styles, layout geometry, colors, fonts, and animations.

---

## 2. Problem Statement & Motivation
- **Preserving Presentation Visual Quality:** Abstract programmatic generation often misses intricate design flourishes, custom borders, master layouts, and subtle brand accents present in designer-crafted PowerPoint templates.
- **Exact Layout Cloning with Intelligent Text Fitting:** Users want their new presentations to look identical to professional template decks (`data/*.pptx`), with only the text replaced by high-value content extracted from their Word documents.

---

## 3. Core 4-Step Pipeline (v0.2)

```
[Selected Base PPTX]                  [Uploaded Word .docx]
        │                                       │
        ▼ (Step 1)                              ▼ (Step 2)
[Inspect PPTX Slides & Shapes]         [Parse Sections & Content]
        │                                       │
        └───────────────────┬───────────────────┘
                            │
                            ▼
       +─────────────────────────────────────────+
       │   Step 3: 9Router AI Agent Reasoning    │
       │  - Maps doc sections to slide slots     │
       │  - Writes fitted text per shape/card    │
       +─────────────────────────────────────────+
                            │
                            ▼
       +─────────────────────────────────────────+
       │   Step 4: Exact Template Slide Cloning  │
       │  - Clones base PPTX presentation        │
       │  - In-place text update (safe font/RTL) │
       │  - Preserves 100% of shapes & design    │
       +─────────────────────────────────────────+
                            │
                            ▼
               [Output Generated PPTX Deck]
```

### Detailed Steps:
1. **Step 1 — Read & Inspect Base PPTX:** AI Agent inspects the user-selected PPTX template (`data/*.pptx`), mapping each slide, text box, title, card, and table into a structured slot inventory (`inspect_template_slides`).
2. **Step 2 — Read & Parse Word Document:** AI Agent reads the input `.docx` file, parsing titles, headings, problem statements, paragraphs, bullets, and tables (`parse_docx`).
3. **Step 3 — AI Content Generation & Slot Mapping:** The 9Router LLM (`Config.NINEROUTER_CHAT_MODEL`) reasons over the template slot inventory and document outline to draft customized, high-impact replacement texts for each slide and shape index.
4. **Step 4 — In-Place Slide Mutation:** The engine clones the base presentation and applies text replacements in-place (`_safe_update_text_frame`), preserving original font families, font sizes, colors, and formatting while supporting RTL/Persian alignment.

---

## 4. Key Features & Architecture Summary

### 4.1 9Router-Only AI Foundation
- All LLM reasoning, chat, and tool loops are routed exclusively through `9Router` (`NINEROUTER_URL`, `NINEROUTER_KEY`, `NINEROUTER_CHAT_MODEL`).
- Default model: `ag/gemini-3.7-flash-high`.
- Integrated web search (`/v1/search`) and fetch (`/v1/web/fetch`).
- Integrated image generation (`/v1/images/generations` binary endpoint).

### 4.2 PPTX Extraction & Component Library
- Automatically extracts shape metadata and saves `data/components/components.json`.
- Catalogs template layout types, shapes, dimensions, and visual assets.

### 4.3 Tkinter GUI & CLI
- **Tab 1: Docx to PPTX Generator & In-App Preview:**
  - File browser for `.docx` and template selector from `data/`.
  - **"Open in PowerPoint" Button:** Instantly launches the generated presentation in Microsoft PowerPoint.
  - **In-App Slide Preview Box:** Direct 2D interactive canvas slide viewer with Next/Previous slide navigation right within the desktop application.
  - Live 4-step progress and AI agent execution logs.
- **Tab 2: Templates & Components:** Live inspection and reload of extracted template assets.
- **Tab 3: Autonomous AI Agent:** Interactive terminal with direct tool invocation.
- **Tab 4: Settings (.env):** In-app configuration for 9Router endpoints and model parameters.
- **CLI Mode:** Available via `python -m pptx_jahat --cli`.

---

## 5. Technical Specifications

| Parameter | Specification |
|---|---|
| Runtime | Python 3.11+ / uv |
| Primary AI Gateway | 9Router (`/v1/chat/completions`, `/v1/search`, `/v1/web/fetch`, `/v1/images/generations`) |
| Presentation Engine | `python-pptx` (Exact Slide Cloning & In-Place Mutation) |
| Document Engine | `python-docx` |
| Image Engine | `Pillow (PIL)` |
| UI Framework | Tkinter (`ttk` clam theme) |

---

## 6. Verification & Test Results (v0.2)
- [x] Step 1 inspects template shape slots accurately.
- [x] Step 2 extracts Word document outline.
- [x] Step 3 9Router LLM generates complete slide text mappings.
- [x] Step 4 in-place text frame replacement clones exact template presentation without visual drift.
- [x] Full test execution verified via `uv run python test_suite.py`.

---

## 7. Future Roadmap (v0.3+)
- **v0.3:** Dynamic slide count expansion (duplicating middle content slide layouts when Word doc has more sections than template).
- **v0.4:** Native chart data replacement (mutating embedded Excel/PPTX charts directly from Word tables).
- **v0.5:** Slide thumbnail visual preview within Tkinter UI.
