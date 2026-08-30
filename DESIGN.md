# DESIGN.md

## System Architecture

```
                       +-----------------------------+
                       |      Tkinter GUI / CLI      |
                       +--------------+--------------+
                                      |
                                      v
                       +-----------------------------+
                       |          AI Agent           |
                       |    (LLM + Tool Calling)     |
                       +--------------+--------------+
                                      |
         +----------------------------+----------------------------+
         |                            |                            |
         v                            v                            v
+------------------+         +------------------+         +------------------+
| PPTX Extraction  |         | Word Docx Parser |         | External APIs    |
| & Component Lib  |         | & Slide Mapper   |         | (OpenAI, Exa,    |
| (data/components)|         | (Docx -> PPTX)   |         | Image Gen)       |
+------------------+         +------------------+         +------------------+
```

## Directory Structure
```
pptx-jahat/
├── data/
│   ├── components/
│   │   ├── components.json
│   │   ├── shapes/
│   │   └── images/
│   └── (user pptx/docx files)
├── src/
│   └── pptx_jahat/
│       ├── __init__.py
│       ├── config.py           # .env config & validation
│       ├── agent.py            # LLM autonomous tool loop
│       ├── tools/
│       │   ├── filesystem.py   # Read/write/edit/delete/list
│       │   ├── exa_search.py   # Exa search & fetch
│       │   ├── pptx_engine.py  # Extraction & components.json
│       │   ├── docx_parser.py  # Word doc parsing & structure
│       │   ├── pptx_builder.py # Generates pptx from components + template
│       │   └── image_gen.py    # Image generation
│       └── gui/
│           ├── app.py          # Tkinter interface
│           └── components.py   # UI widgets & helpers
├── AGENT.md
├── DESIGN.md
├── pyproject.toml
└── .env.example
```

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
          "background_color": "#FFFFFF",
          "components": [
            {
              "id": "comp_0_1",
              "type": "title_box",
              "label": "Main Header Box",
              "description": "Primary slide title with large font styling",
              "left": 1000000,
              "top": 1200000,
              "width": 7144000,
              "height": 1000000,
              "font": {"name": "Arial", "size_pt": 40, "bold": true, "color": "#003366"},
              "fill": {"color": null, "type": "transparent"},
              "border": null
            }
          ]
        }
      ]
    }
  ]
}
```

## 9Router & External Integrations
- **Chat & Tool Calling**: Seamlessly routes through `9Router` `/v1/chat/completions` (OpenAI format with model routing and multi-provider combos) or standard OpenAI `/v1`.
- **Search**: `9Router` `/v1/search` with auto-fallback or direct Exa Search API.
- **Fetch**: `9Router` `/v1/web/fetch` (`jina-reader`, `firecrawl`, `tavily`, `exa`) with HTTP scrape fallback.
- **Image Generation**: `9Router` `/v1/images/generations` binary endpoint (supporting DALL-E, Imagen, FLUX, etc.) with OpenAI DALL-E fallback.
2. **Generation Flow**: Upload `.docx` -> Parse hierarchy (headings, bullets, paragraphs, tables) -> AI Agent selects best layout & components matching source design -> Generates slide deck + optional generated images.
3. **Interactive GUI**: Tabbed Tkinter UI for Component Manager, Docx-to-PPTX generator, Agent Chat, Settings (`.env` config editor).
