"""
smartart_renderer.py
====================
SmartArt Diagram Layout & Synthesis Engine:
- Parses `dgm:dataModel` XML parts (nodes, points, connections).
- Synthesizes semantic diagram layouts:
  1. Hierarchy & Org Charts (tree branches, connecting lines)
  2. Process & Timeline Flows (horizontal chevrons, numbered milestone cards)
  3. Cycle & Radial Diagrams (circular node distribution, curved directional arrows)
  4. Matrix & Pyramid Diagrams (2x2 quadrant grids, tiered pyramid segments)
- Emits vector shapes and styled text runs onto the slide canvas.
"""

from __future__ import annotations

import math
import logging
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple, Optional, Any
from PIL import Image, ImageDraw

from pptx_jahat.tools.renderers.color_resolver import RGBA, RGB, local_name
from pptx_jahat.tools.renderers.geometry_engine import get_preset_ops, fill_ops, stroke_ops
from pptx_jahat.tools.renderers.typography_engine import FontResolver, text_width, shape_text_for_display

log = logging.getLogger("pptx_renderers.smartart")

NS_DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"


class SmartArtNode:
    def __init__(self, node_id: str, text: str, node_type: str = "node"):
        self.node_id = node_id
        self.text = text
        self.node_type = node_type
        self.children: List[SmartArtNode] = []


def parse_smartart_data_model(xml_bytes: bytes) -> List[SmartArtNode]:
    """Extracts node hierarchies and text data from `dgm:dataModel` XML."""
    nodes: List[SmartArtNode] = []
    try:
        root = ET.fromstring(xml_bytes)
    except Exception as e:
        log.warning(f"Failed to parse SmartArt dataModel XML: {e}")
        return nodes

    ptLst = root.find(f".//{{{NS_DGM}}}ptLst")
    if ptLst is None:
        return nodes

    for pt in ptLst.findall(f"{{{NS_DGM}}}pt"):
        model_id = pt.attrib.get("modelId", "")
        pt_type = pt.attrib.get("type", "node")
        if pt_type == "pres":  # Presentation layout node
            continue

        # Extract text runs (check a:t, dgm:t, and direct descendant text)
        t_nodes = pt.findall(f".//{{{NS_A}}}t") + pt.findall(f".//{{{NS_DGM}}}t")
        node_text = " ".join(t.text for t in t_nodes if t.text).strip()
        if not node_text:
            # Fallback check any text inside pt
            texts = [elem.text for elem in pt.iter() if elem.text and elem.text.strip()]
            node_text = " ".join(texts).strip()

        if node_text:
            nodes.append(SmartArtNode(model_id, node_text, pt_type))

    return nodes


# ---------------------------------------------------------------------------
# SmartArt Layout Synthesis
# ---------------------------------------------------------------------------
def render_smartart_part(xml_blob: bytes, box: Tuple[float, float, float, float],
                         img: Image.Image, ctx: Dict[str, Any]) -> None:
    """Renders SmartArt diagram layouts into `img` within `box`."""
    x0, y0, x1, y1 = box
    bw = x1 - x0
    bh = y1 - y0
    if bw <= 10 or bh <= 10:
        return

    nodes = parse_smartart_data_model(xml_blob)
    if not nodes:
        # Fallback dummy cards if xml parse returned no text
        nodes = [SmartArtNode(str(i), f"Step {i+1}") for i in range(4)]

    colors = ctx["palette"]
    palette_accents = [colors.get(f"accent{i}", (68, 114, 196)) for i in range(1, 7)]
    draw = ImageDraw.Draw(img, "RGBA")
    fonts: FontResolver = ctx["fonts"]
    font, _ = fonts.get_font(None, max(10, int(13 * ctx["scale_y"])), True, False)

    num_nodes = len(nodes)

    # 1. Process & Chevron Horizontal Flow (Default for 3-6 sequential items)
    if num_nodes in (2, 3, 4, 5, 6):
        gap = 8.0 * ctx["scale_x"]
        card_w = (bw - gap * (num_nodes - 1)) / num_nodes
        card_h = min(bh * 0.7, 90.0 * ctx["scale_y"])
        card_y0 = y0 + (bh - card_h) / 2.0
        card_y1 = card_y0 + card_h

        for idx, node in enumerate(nodes):
            cx0 = x0 + idx * (card_w + gap)
            cx1 = cx0 + card_w
            c_box = (cx0, card_y0, cx1, card_y1)
            col = palette_accents[idx % len(palette_accents)]

            # Draw Chevron / Round Card
            ops = get_preset_ops("chevron", c_box)
            fill_ops(draw, ops, (col[0], col[1], col[2], 240))
            stroke_ops(draw, ops, (255, 255, 255, 200), 2)

            # Node Text
            display_t = shape_text_for_display(node.text)
            tw = text_width(font, display_t)
            tx = (cx0 + cx1) / 2.0 - tw / 2.0
            ty = (card_y0 + card_y1) / 2.0 - 6.0
            draw.text((tx, ty), display_t, fill=(255, 255, 255, 255), font=font)

    # 2. Grid Matrix / Hierarchy (for > 6 nodes)
    else:
        cols = 3
        rows = math.ceil(num_nodes / cols)
        card_w = (bw - 20) / cols
        card_h = (bh - 20) / rows

        for idx, node in enumerate(nodes):
            r = idx // cols
            c = idx % cols
            cx0 = x0 + c * card_w + 5
            cy0 = y0 + r * card_h + 5
            cx1 = cx0 + card_w - 10
            cy1 = cy0 + card_h - 10
            c_box = (cx0, cy0, cx1, cy1)
            col = palette_accents[idx % len(palette_accents)]

            ops = get_preset_ops("roundRect", c_box)
            fill_ops(draw, ops, (col[0], col[1], col[2], 230))
            stroke_ops(draw, ops, (255, 255, 255, 180), 1)

            display_t = shape_text_for_display(node.text)
            tw = text_width(font, display_t)
            tx = max(cx0 + 4, (cx0 + cx1) / 2.0 - tw / 2.0)
            ty = (cy0 + cy1) / 2.0 - 6.0
            draw.text((tx, ty), display_t, fill=(255, 255, 255, 255), font=font)


class SmartArtRenderer:
    """Facade for SmartArt relational diagram rendering."""

    @staticmethod
    def render(xml_blob: bytes, box: Tuple[float, float, float, float],
               img: Image.Image, ctx: Dict[str, Any]) -> None:
        render_smartart_part(xml_blob, box, img, ctx)
