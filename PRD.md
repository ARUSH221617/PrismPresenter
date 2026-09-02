# Product Requirements Document (PRD) — v0.4

**Project Name:** PrismPresenter (formerly PPTX Jahat)  
**Version:** 0.4  
**Status:** Implemented & Verified Architecture  
**Author:** Amirreza Uneszadeh Shirazi (ARUSH) — `arush221617@gmail.com`  
**Date:** 2026-09-02  

---

## 1. Executive Summary
**PrismPresenter (v0.4)** is an autonomous, multi-modal presentation synthesis studio powered by **9Router**. In this version, the platform has been fully rebranded to **PrismPresenter** with a modern **shadcn/ui** and **Tailwind CSS** design system.

Key highlights of v0.4 include:
1. **Google Gemini-Style AI Terminal Chat**: Chat bubble thread with Markdown/code snippet rendering, collapsible execution reasoning drawer, suggestion chips, active tool capability toggles (Web Search/Scraper, PPTX Generator/Editor, NOTE.md context attachment), message copy/edit, and transcript export (Markdown & JSON).
2. **Interactive Component Primitives Catalog**: Card grid preview with image streaming endpoint (`/api/components/image/<filename>`), category filter chips (Headers, Metrics, Cards, Tables, Images), real-time search, and detailed JSON schema modal inspector.
3. **Full Dark / Light Mode System**: CSS variables and theme controller with localStorage persistence.
4. **Dedicated Help & Workflow Page**: Video player showing the brand reveal motion video on loop (`assets/videos/PrismPresenter_brand_reveal.mp4`), end-to-end architecture breakdown, and developer contact details.
5. **Full Multi-Device Responsiveness**: Adaptive layouts optimized for desktop, tablet, and mobile touchscreens.

---

## 2. Problem Statement & Motivation
- **High-End UI/UX Standards:** Traditional presentation generation tools present clunky settings or plain terminal text. Users need an intuitive, polished chat experience with Markdown formatting and copyable snippets.
- **Dynamic Tool Control:** Users need transparent control over which capabilities the AI Agent activates per prompt (Web Search vs Slide Synthesis vs Template Context).
- **Visual Primitives Exploration:** Extracted PPTX shapes and components must be browsable visually with live image previews and schema inspectors rather than static raw JSON.
- **Responsive & Accessible Theming:** Dark mode by default with instant light mode toggling and mobile-friendly touch responsiveness.

---

## 3. Core Architecture & Pipeline (v0.4)

```
[Uploaded Word .docx]               [Selected Template / data/*.pptx]
        │                                           │
        ▼ (Step 1: Document Ingestion)              ▼ (Step 2: Template Intelligence)
[Parse Sections & Content]                 [Classify Archetypes & NOTE.md]
        │                                           │
        └─────────────────────┬─────────────────────┘
                              │
                              ▼
        +───────────────────────────────────────────+
        │    Step 3: 9Router AI Agent Reasoning     │
        │   - Multimodal Vision slide inspection    │
        │   - Maps doc sections to slide slots      │
        │   - Dynamic multi-template slide matching │
        │   - Tool filtering (Search / Synthesis)   │
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
        │   Tier 1: Native PowerPoint COM (Windows) │
        │   Tier 2: Web Vector Engine (HTML5/DOM)   │
        │   Tier 3: Pure-Python PIL SlideRenderer   │
        +───────────────────────────────────────────+
                              │
                              ▼
        [Output Generated PPTX & Web Studio Preview]
```

---

## 4. Key Features & Components (v0.4)

### 4.1 Rebranding & Design System
- **Brand Identity:** Rebranded from PPTX Jahat to **PrismPresenter**.
- **UI Framework:** Tailwind CSS with atomic **shadcn/ui** tokens (Geist & Inter typography, Geist Mono for code/telemetry, HSL color variables).
- **Theme Modes:** Dark mode default with light mode switcher and persistent state.

### 4.2 Core Studio Views
1. **Slide Synthesizer & Studio (`/`):**
   - Drag-and-drop Word `.docx` file ingestion zone.
   - Template selection with intelligent auto-matching.
   - Live SSE synthesis log stream with one-click copy.
   - Live Deck Preview, High-Res Screenshots, and AI Payload inspectors.
   - Slide filmstrip thumbnail overview and fullscreen Lightbox modal (`F` key).
2. **Template Intelligence & Knowledge Base:**
   - Visual template repository table with slide counts and analysis status.
   - Single and batch AI template analysis pipelines.
   - Built-in editor for `data/NOTE.md` design rules.
3. **Deck & Template Manager:**
   - Dual file tables for Generated Presentations (`data/output/`) and Reference Templates (`data/`).
   - Integrity verification & XML auto-healing, Duplicate, Rename, Delete, Download, and PowerPoint launcher.
4. **Visual Component Primitives Catalog:**
   - Card grid previewing shapes, headers, metric callouts, tables, and extracted images (`/api/components/image/<filename>`).
   - Type filter chips and instant search bar.
   - Component schema detail modal with JSON copying and visual inspection.
5. **Autonomous AI Terminal (Google Gemini Style):**
   - Multi-turn conversational chat thread with user and AI avatars.
   - Rich Markdown formatting with syntax-styled code snippets and copy buttons.
   - Collapsible reasoning trace drawer for live tool execution streaming.
   - Floating input capsule with prompt suggestions, active model indicator, and tool toggles:
     - **Globe icon:** Enable/disable Web Search & Fetch tools.
     - **Inspect icon:** Attach `data/NOTE.md` template guidelines as context.
     - **Slide icon:** Enable/disable Slide Synthesis & editing tools.
   - Message-level actions (Copy response, Edit & reload user prompt).
   - Chat transcript export to Markdown (`.md`) or JSON (`.json`).
6. **Engine Settings:**
   - Interactive configuration editor for `.env` credentials, 9Router endpoints, and model routing.
7. **Help & Architecture:**
   - Top looping brand motion video (`assets/videos/PrismPresenter_brand_reveal.mp4`) that pauses when switching tabs.
   - 4-step end-to-end architecture pipeline summary.
   - Author and project contact details.

---

## 5. Technical Specifications

| Parameter | Specification |
|---|---|
| Product Name | PrismPresenter |
| Runtime | Python 3.11+ / uv |
| Web Backend | Flask 3.1+, Flask-CORS |
| Frontend UI | Tailwind CSS, shadcn/ui components, Marked.js, Lucide Icons |
| AI Gateway | 9Router (`/v1/chat/completions`, `/v1/search`, `/v1/web/fetch`, `/v1/images/generations`) |
| Primary Model | `ag/gemini-3.7-flash-high` |
| Presentation Engine | `python-pptx` (Cross-Presentation Slide Cloning & In-Place Mutation) |
| Document Parser | `python-docx` |
| Rendering Engine | 3-Tier Cascade: Native PowerPoint COM -> Web Vector Engine -> Pure PIL |
| Configuration | Dynamic `.env` hot-reload via `Config.reload()` |

---

## 6. Verification & Test Suite
- [x] Web endpoints and REST APIs verified (`tests/test_web_app.py`).
- [x] Web vector slide rendering verified (`tests/test_web_renderer.py`).
- [x] 3-tier cascade render priority verified (`Native PowerPoint -> Web -> PIL`).
- [x] Component image serving endpoint (`/api/components/image/<filename>`) verified.
- [x] Real-time SSE streaming for generation, analyzer, and agent chat validated.
- [x] Responsive layout tested for desktop, tablet, and mobile screens.
