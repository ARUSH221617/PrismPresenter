# AGENT.md

## Overview
Autonomous AI Agent for PPTX extraction, templating, and Word-to-PPTX generation.

## Capabilities & Architecture
1. **Model Integration**: Powered entirely by **9Router** (`NINEROUTER_URL`, `NINEROUTER_KEY`, `NINEROUTER_CHAT_MODEL`). Supports all 9Router models and multi-provider auto-fallback combos.
2. **File & System Tools**:
   - `read_file(path)`
   - `write_file(path, content)`
   - `edit_file(path, old_text, new_text)`
   - `list_dir(path)`
   - `make_dir(path)`
   - `delete_file(path)`
3. **Web Search & Fetch**:
   - 9Router `/v1/search` (tavily, exa, brave-search, serper, perplexity, etc.).
   - 9Router `/v1/web/fetch` (jina-reader, firecrawl, exa, tavily).
4. **Slide & Component Engine**:
   - PPTX layout analyzer & shape extractor.
   - Component cataloger (`data/components/` + `data/components/components.json`).
   - Word (`.docx`) document structure parser.
   - Smart template assembler mapping document sections to slide layouts & extracted shapes.
5. **Image Generator**:
   - 9Router `/v1/images/generations` binary endpoint (`NINEROUTER_IMAGE_MODEL`, e.g. `gemini/gemini-3-pro-image-preview`, `openai/dall-e-3`, `flux`).
6. **Execution Modes**:
   - CLI / Direct Programmatic Agent Runner (`python -m pptx_jahat --cli`).
   - Tkinter GUI with asynchronous background task worker.

## Tool Registry
- `file_ops`: Full filesystem read/write/edit/delete inside workspace.
- `pptx_extractor`: Parse `data/*.pptx`, export shape specs & XML/images to `data/components/`.
- `pptx_generator`: Build new slides based on parsed `.docx` outlines and component blueprints.
- `image_generator`: Produce visual assets for slides and place them into presentations.
- `web_search`: Search external knowledge using Exa API.
