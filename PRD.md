# Product Requirements Document (PRD) — v0.3

**Project Name:** PPTX Jahat  
**Version:** 0.3  
**Status:** Implemented (Verified Architecture)  
**Author:** AI Agent / Product Engineering  
**Date:** 2026-09-02  

---

## 1. Executive Summary
**PPTX Jahat (v0.3)** is a modern, web-powered AI presentation generation suite powered exclusively by **9Router**. In this version, the entire user interface has been migrated from legacy desktop Tkinter to a responsive **Flask Web Single-Page Application (SPA)** styled with **Tailwind CSS**, a dark crimson theme, and real-time **Server-Sent Events (SSE)**. 

Furthermore, presentation previewing has been upgraded to a **3-tier cascade rendering engine** (`Native PowerPoint COM -> Web Vector Engine -> Pure-Python PIL Fallback`), providing high-fidelity visual previewing, client-side vector DOM slide inspection, and reliable automated rendering for AI multimodal vision agent verification.

---

## 2. Problem Statement & Motivation
- **Modern Interactive Interface:** The previous desktop GUI was platform-constrained and limited in real-time responsiveness. A browser-based SPA enables richer layout previewing, drag-and-drop document uploads, instant slide navigation, and cross-device accessibility.
- **Robust Multi-Tier Slide Rendering:** Pure-Python raster rendering can encounter font or complex shape fidelity limitations, while COM automation requires Windows and PowerPoint. A 3-tier cascade (`Native PowerPoint -> Web Render Engine -> Pure PIL`) ensures maximum visual fidelity with automatic fallback.
- **Template Intelligence & Knowledge Base:** Users need deep design archetyping for template selection, documented in a structured markdown knowledge base (`data/NOTE.md`), with single and batch AI analysis capabilities.

---

## 3. Core Architecture & Pipeline (v0.3)

```
[Uploaded Word .docx]               [Selected Template / data/*.pptx]
        │                                           │
        ▼ (Step 1: Document Parsing)                ▼ (Step 2: Template & Archetype Analysis)
[Parse Sections & Content]                 [Inspect Slots, Archetypes & NOTE.md]
        │                                           │
        └─────────────────────┬─────────────────────┘
                              │
                              ▼
        +───────────────────────────────────────────+
        │    Step 3: 9Router AI Agent Reasoning     │
        │   - Multimodal Vision slide inspection    │
        │   - Maps doc sections to slide slots      │
        │   - Dynamic multi-template slide matching │
        +───────────────────────────────────────────+
                              │
                              ▼
        +───────────────────────────────────────────+
        │   Step 4: Slide Cloning & In-Place Infill │
        │   - Cross-presentation slide cloning      │
        │   - Safe text frame & font preservation   │
        │   - Persian/Arabic RTL alignment support  │
        +───────────────────────────────────────────+
                              │
                              ▼
        +───────────────────────────────────────────+
        │   Step 5: 3-Tier Cascade Preview Engine   │
        │   Tier 1: Native PowerPoint COM           │
        │   Tier 2: Web Vector Engine (HTML/SVG/DOM)│
        │   Tier 3: Pure-Python PIL SlideRenderer   │
        +───────────────────────────────────────────+
                              │
                              ▼
        [Output Generated PPTX & Web Live Preview]
```

---

## 4. Key Features & Components (v0.3)

### 4.1 Flask Web Single-Page Application (SPA)
- **Framework & Styling:** Flask backend, Tailwind CSS, shadcn-styled dark crimson aesthetic (`#0b0c10` canvas, `#e50914` crimson accents).
- **Auto-Browser Launch:** Automatically launches default browser at `http://127.0.0.1:5000` on startup (`pptx-jahat`).
- **Interactive CLI:** Maintained via `pptx-jahat --cli`.
- **Server-Sent Events (SSE):** Real-time streaming for generation execution logs, template analysis progress, and AI terminal reasoning steps.

### 4.2 Five Core Web Management Views
1. **Slide Synthesizer & Live Inspector:**
   - Word `.docx` file upload and template style selector.
   - Live SSE execution stream log terminal.
   - Interactive slide viewer with 3 sub-views:
     - *Live Deck Preview*: Interactive slide carousel with engine indicator badge.
     - *Visual Screenshots*: High-res slide visual snapshots.
     - *AI Test Snapshots*: Multimodal visual payloads sent to 9Router Vision Agent.
   - One-click "Open in PowerPoint" host launcher and direct file download.
2. **Template Intelligence & AI Analyzer:**
   - Repository grid displaying all `data/*.pptx` templates with slide counts and analysis status tags (`Analyzed` vs `Pending`).
   - Single-template and batch template AI analysis pipelines.
   - In-app Markdown editor for `data/NOTE.md` with Save and Reload actions.
3. **Deck & Template Manager:**
   - Dual-table management for Generated Decks (`data/output/`) and Reference Templates (`data/`).
   - Actions: Verify & Auto-Heal PPTX integrity, Duplicate, Rename, Delete, Download, and Launch in PowerPoint.
   - Slide preview carousel with active engine badge.
4. **Component Catalog:**
   - Extracted shapes, layout containers, and design primitives viewer (`data/components/components.json`).
   - Template re-scan and component extraction trigger.
5. **Autonomous AI Terminal & Engine Settings:**
   - Autonomous multi-step agent chat terminal with streaming tool-calling pipeline.
   - `.env` configuration manager for 9Router URL, API keys, and model parameters.

### 4.3 3-Tier Cascade Slide Rendering Engine
- **Tier 1 (Native PowerPoint COM):** Automated PowerPoint slide export on Windows hosts for exact native vector fidelity.
- **Tier 2 (Web Render Engine):**
  - Python vector engine (`PPTXWebRenderer`) converting slide elements, tables, RTL typography, and shapes into standards-compliant HTML5/SVG vector DOM structures.
  - Client-side JS parser/renderer (`pptx-web-renderer.js`) for direct interactive browser DOM slide visualization.
- **Tier 3 (Pure-Python PIL Fallback):** Robust `SlideRenderer` geometry and font engine for headless environments without COM or browser engine.
- **Configurable Mode:** `Config.RENDER_MODE = auto | native | web | pil`.

### 4.4 9Router AI Foundation
- Unified routing for all reasoning, chat, vision, search, and image generation via `9Router`.
- Default reasoning model: `ag/gemini-3.7-flash-high`.
- Integrated web search (`/v1/search`), web scraper/fetch (`/v1/web/fetch`), and image generation (`/v1/images/generations`).

---

## 5. Technical Specifications

| Parameter | Specification |
|---|---|
| Runtime | Python 3.11+ / uv |
| Web Server | Flask 3.1+, Flask-CORS |
| Frontend UI | Tailwind CSS, Lucide Icons, Vanilla JS SPA, SSE Streaming |
| AI Gateway | 9Router (`/v1/chat/completions`, `/v1/search`, `/v1/web/fetch`, `/v1/images/generations`) |
| Presentation Engine | `python-pptx` (Cross-Presentation Slide Cloning & In-Place Mutation) |
| Document Engine | `python-docx` |
| Rendering Engine | 3-Tier Cascade: Native PowerPoint COM -> Web Vector Engine -> Pure PIL |
| Configuration | Dynamic `.env` management via `Config.reload()` |

---

## 6. Verification & Test Suite (v0.3)
- [x] Web endpoints and REST APIs verified (`tests/test_web_app.py`).
- [x] Web vector slide rendering verified (`tests/test_web_renderer.py`).
- [x] 3-tier cascade render priority verified (`Native PowerPoint -> Web -> PIL`).
- [x] Real-time SSE streaming for generation, analyzer, and agent chat validated.
- [x] PPTX integrity verification and auto-healing validated.
- [x] CLI fallback mode (`--cli`) operational.

---

## 7. Future Roadmap (v0.4+)
- **v0.4:** Native chart data mutation (updating embedded Excel charts and series directly from Word tables).
- **v0.5:** Drag-and-drop slide reordering and custom component composition in Web UI.
