# DESIGN.md — PrismPresenter Architecture & Design System

## System Architecture

```
                       +-----------------------------------+
                       |    PrismPresenter Web SPA / CLI   |
                       |    (Flask + Tailwind + shadcn/ui) |
                       +-----------------+-----------------+
                                         |
                                         v
                       +-----------------------------------+
                       |        Autonomous AI Agent        |
                       |      (LLM + Tool Orchestration)   |
                       +-----------------+-----------------+
                                         |
         +-------------------------------+-------------------------------+
         |                               |                               |
         v                               v                               v
+------------------+           +-------------------+           +-------------------+
| PPTX Extraction  |           | Word Docx Parser  |           | External Gateway  |
| & Component Lib  |           | & Slide Infiller  |           | (9Router Engine)  |
| (data/components)|           | (Docx -> PPTX)    |           | Search/Vision/IMG |
+------------------+           +-------------------+           +-------------------+
```

---

## Directory Structure
```
pptx-jahat/
├── assets/
│   ├── images/
│   │   ├── logo.jpeg
│   │   ├── logo-icon.jpeg
│   │   ├── logo-transparent.png
│   │   └── logo-icon-transparent.png
│   └── videos/
│       └── PrismPresenter_brand_reveal.mp4
├── data/
│   ├── components/
│   │   ├── components.json
│   │   ├── shapes/
│   │   └── images/
│   ├── output/
│   └── (reference pptx/docx templates)
├── src/
│   └── pptx_jahat/
│       ├── __init__.py         # App entrypoint & CLI banner
│       ├── __main__.py
│       ├── config.py           # .env config & reload
│       ├── agent.py            # 9Router AI autonomous agent & tool filtering
│       ├── web/
│       │   ├── app.py          # Flask REST API & SSE streaming server
│       │   ├── static/
│       │   │   ├── css/
│       │   │   │   └── custom.css # shadcn/ui tokens & Gemini chat styles
│       │   │   ├── js/
│       │   │   │   ├── app.js # SPA client logic, theme & chat state
│       │   │   │   └── pptx-web-renderer.js # Vector slide renderer
│       │   │   └── images/
│       │   └── templates/
│       │       └── index.html  # Main SPA template (7 tabbed views)
│       └── tools/
│           ├── filesystem.py   # Workspace filesystem tools
│           ├── exa_search.py   # Web search & page scraper
│           ├── pptx_engine.py  # Shape extraction & component catalog
│           ├── docx_parser.py  # Word document hierarchy parser
│           ├── pptx_builder.py # Slide cloning, in-place infill & auto-heal
│           ├── image_gen.py    # DALL-E / 9Router image generator
│           ├── preview.py      # Multi-tier preview coordinator
│           └── renderers/
│               ├── com_renderer.py       # Tier 1: Native COM
│               ├── web_renderer.py       # Tier 2: Vector HTML/SVG
│               └── color_resolver.py     # Tier 3: Pure PIL renderer
├── AGENT.md
├── DESIGN.md
├── PRD.md
├── pyproject.toml
└── .env.example
```

---

## Component Schema (`data/components/components.json`)
```json
{
  "templates": [
    {
      "source_file": "template1.pptx",
      "slide_dimensions": {"width_emu": 9144000, "height_emu": 5143500},
      "slides": [
        {
          "index": 0,
          "layout_name": "Title Slide",
          "components": [
            {
              "id": "comp_0_1",
              "type": "title_box",
              "label": "Main Header Box",
              "description": "Primary slide title with large font styling",
              "sample_text": "Sample Title",
              "position": {
                "left": 1000000,
                "top": 1200000,
                "width": 7144000,
                "height": 1000000
              },
              "font": {"name": "Arial", "size_pt": 40, "bold": true, "color": "#003366"},
              "fill": {"type": "none", "color": null},
              "line": {"color": null, "width_pt": null},
              "image_path": "data/components/images/comp_0_1.png"
            }
          ]
        }
      ]
    }
  ],
  "all_components": [],
  "component_counts_by_label": {}
}
```

---

## 9Router & Tool Integration
- **Chat & Tool Calling**: `9Router` `/v1/chat/completions` (OpenAI format with multi-provider routing).
- **Search & Scrape**: `9Router` `/v1/search` and `/v1/web/fetch` with toggle support in UI.
- **Image Generation**: `9Router` `/v1/images/generations` binary endpoint.
- **Dynamic Tool Permissions**: Tools can be enabled/disabled per chat query (`enable_search`, `enable_pptx_tools`).

---

## UI/UX Design System (v0.4)
- **Design Tokens**: Based on **shadcn/ui** with CSS variables (`--background`, `--card`, `--primary`, `--secondary`, `--border`, `--muted`, `--ring`).
- **Typography**: `Geist` & `Inter` for UI interfaces, `Geist Mono` for code blocks, logs, and telemetry.
- **Theme Modes**: Dark mode default with light mode switcher and localStorage persistence (`prism_theme`).
- **Google Gemini Chat Pill**: Floating input capsule with auto-expanding textarea, prompt suggestion chips, reasoning trace accordion, message copy/edit, and Markdown code snippet rendering with copy buttons.
- **Visual Component Explorer**: Card grid layout with image streaming endpoint (`/api/components/image/<filename>`), categorical filters, and JSON modal viewer.
