# Finalized Architecture Plan: Web-Enhanced PPTX Rendering Engine & Cascade Preview

## Executive Summary
Upgrade presentation rendering engine in PPTX Jahat to support full 3-tier cascade:
1. **Tier 1 (Native PowerPoint COM)**: Windows native PowerPoint automation for exact fidelity.
2. **Tier 2 (Web Render Engine)**:
   - **Frontend Canvas / DOM**: Client-side JS parser/renderer directly converts `.pptx` array buffer into interactive vector HTML/SVG/Canvas in web browser.
   - **Backend / Multimodal AI**: Python HTML/SVG vector DOM renderer for structured vector slide generation and AI Vision inspection.
3. **Tier 3 (Pure-Python PIL Fallback)**: Standalone `SlideRenderer` geometry and font engine for headless environments without COM or browser engine.

---

## 1. Engine Cascade Flow

```
                     PPTX Slide Rendering Request
                                │
                                ▼
                   Is PowerPoint COM Available?
                             /    \
                       YES  /      \  NO / Fails
                           ▼        ▼
               Native PowerPoint  Web Render Engine (Tier 2)
                   COM Export     ├── Frontend: Client JS PPTX Parser/DOM
                                  └── Backend: Python HTML/SVG Vector Engine
                                            │
                                            ▼ (If fails / unsupported)
                                  Pure PIL SlideRenderer (Tier 3)
```

Engine selection priorities in `render_pptx` & `render_pptx_file_previews`:
- Try `Native PowerPoint` export.
- If COM is unavailable or raises exception, invoke `Web Render Engine`.
- If Web rendering fails or cannot process slide elements, fallback to `Pure PIL`.
- Engine label returned in metadata: `"Native PowerPoint"` | `"Web Render Engine"` | `"Pure PIL"`.

---

## 2. Component Design & Changes

### A. Python Backend Web Renderer (`src/pptx_jahat/tools/renderers/web_renderer.py`)
- `PPTXWebRenderer`: Traverses slide shapes, text frames, tables, custom geometry XML (`a:custGeom`, `a:prstGeom`), and theme palettes.
- Outputs clean, responsive HTML5 and SVG vector elements with CSS transforms matching slide aspect ratio (16:9 or 4:3).
- Supports bidirectional Persian/Arabic and RTL text layout alignment.
- Generates base64 data URIs or self-contained HTML payloads for slide snapshots.

### B. Core Preview Cascade Update (`src/pptx_jahat/tools/preview.py`)
- Modify `render_pptx(source, width, use_com, return_engine_info)`:
  - Step 1: Execute `export_pptx_slides_com` if `use_com=True` and file exists on Windows.
  - Step 2: On COM failure or non-Windows, call `PPTXWebRenderer.render_pptx_web(...)`.
  - Step 3: If Web renderer encounters unhandled elements, fallback to `SlideRenderer` (PIL).
- Configurable via `Config.RENDER_MODE` (`auto`, `native`, `web`, `pil`).

### C. Client-Side JS Web PPTX Renderer (`src/pptx_jahat/web/static/js/pptx-web-renderer.js`)
- Client-side parser that accepts binary `.pptx` array buffer.
- Unzips PPTX package in browser, extracts slide XML, and constructs interactive DOM elements on viewer canvas.
- Provides fallback to server-rendered preview image when buffer parsing encounters complex embedded objects.

### D. Web API Route Enhancements (`src/pptx_jahat/web/app.py`)
- `POST /api/preview/render`: Returns slide images and `engine_name` ("Native PowerPoint", "Web Render Engine", "Pure PIL").
- `GET /api/preview/slide-html`: Returns raw HTML/SVG vector DOM string for interactive web previewing.
- `GET /api/generator/pptx-buffer`: Stream raw `.pptx` array buffer for client-side JS rendering.

### E. Frontend UI Integration (`src/pptx_jahat/web/templates/index.html` & `app.js`)
- Header and slide viewer badge dynamically reflects active engine:
  - 🟢 `⚡ Native PowerPoint`
  - 🔵 `🌐 Web Render Engine`
  - 🔴 `🎨 Pure PIL`
- Seamless switching between Live Vector DOM view and rendered raster screenshots.

---

## 3. Implementation Steps

1. **Step 1: Implement `web_renderer.py`**:
   - Construct XML-to-HTML/SVG vector DOM generator for python-pptx shapes, geometry, tables, and typography.
2. **Step 2: Update `preview.py` and `config.py`**:
   - Integrate 3-tier cascade (`COM -> Web -> PIL`) in `render_pptx` and `render_pptx_file_previews`.
   - Add `RENDER_MODE` property to `Config`.
3. **Step 3: Create `pptx-web-renderer.js`**:
   - Build client-side browser DOM slide visualizer.
4. **Step 4: Update Flask Web API Endpoints**:
   - Expose HTML/SVG rendering routes and arraybuffer endpoints in `app.py`.
5. **Step 5: Frontend SPA Wiring**:
   - Integrate engine badge and dual-render support in `app.js` and `index.html`.
6. **Step 6: Unit & Integration Testing**:
   - Write tests in `tests/test_web_renderer.py`.
   - Verify cascade failover behavior in test environment.
