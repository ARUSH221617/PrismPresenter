"""
table_renderer.py
=================
Advanced DrawingML table grid renderer: cell spanning (hMerge/vMerge, gridSpan/rowSpan),
diagonal borders, interior margins/padding, vertical text orientation, and custom cell fills.
"""

from __future__ import annotations

import math
import logging
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple, Optional, Any
from PIL import Image, ImageDraw

from pptx_jahat.tools.renderers.color_resolver import (
    RGBA, RGB, q, local_name, resolve_element_color, clamp, NS_A
)
from pptx_jahat.tools.renderers.typography_engine import render_txbody

log = logging.getLogger("pptx_renderers.table")


def render_table(tbl: ET.Element, box: Tuple[float, float, float, float],
                 img: Image.Image, ctx: Dict[str, Any]) -> None:
    """Renders a DrawingML table element with full support for merged cells, borders, and text formatting."""
    x0, y0, x1, y1 = box
    if x1 - x0 < 2 or y1 - y0 < 2:
        return

    colors = ctx["palette"]
    scale = ctx["scale_x"]
    draw = ImageDraw.Draw(img, "RGBA")

    # 1. Parse Grid Columns
    grid = tbl.find(q(NS_A, "tblGrid"))
    col_w: List[int] = []
    if grid is not None:
        for gc in grid.findall(q(NS_A, "gridCol")):
            try:
                col_w.append(int(gc.attrib.get("w", "0")))
            except ValueError:
                col_w.append(0)
    if not col_w:
        return

    total_grid_w = sum(col_w) or 1
    scale_w = (x1 - x0) / float(total_grid_w)
    xs = [x0]
    for cw in col_w[:-1]:
        xs.append(xs[-1] + cw * scale_w)
    xs.append(x1)

    # 2. Parse Rows
    rows = tbl.findall(q(NS_A, "tr"))
    if not rows:
        return
    row_h: List[int] = []
    for tr in rows:
        try:
            row_h.append(int(tr.attrib.get("h", "0")))
        except ValueError:
            row_h.append(0)

    total_h = sum(row_h)
    hs = [rh * (y1 - y0) / total_h for rh in row_h] if total_h > 0 else [(y1 - y0) / len(rows)] * len(rows)
    ys = [y0]
    for hh in hs[:-1]:
        ys.append(ys[-1] + hh)
    ys.append(y1)

    num_rows = len(rows)
    num_cols = len(col_w)

    # 3. Grid Cell Matrix for Span / Merge Resolution
    # cell_matrix[r][c] = { 'elem': tc, 'skip': bool, 'span_r': int, 'span_c': int }
    matrix: List[List[Dict[str, Any]]] = [
        [{"elem": None, "skip": False, "span_r": 1, "span_c": 1} for _ in range(num_cols)]
        for _ in range(num_rows)
    ]

    for ri, tr in enumerate(rows):
        ci = 0
        for tc in tr.findall(q(NS_A, "tc")):
            # Advance ci past already-occupied / row-spanned cells
            while ci < num_cols and matrix[ri][ci]["skip"]:
                ci += 1
            if ci >= num_cols:
                break

            grid_span = max(1, int(tc.attrib.get("gridSpan", "1")))
            row_span = max(1, int(tc.attrib.get("rowSpan", "1")))

            # Check horizontal and vertical merge flags
            if tc.attrib.get("hMerge") == "1":
                matrix[ri][ci]["skip"] = True
                for back_c in range(ci - 1, -1, -1):
                    if not matrix[ri][back_c]["skip"]:
                        matrix[ri][back_c]["span_c"] += 1
                        break
            elif tc.attrib.get("vMerge") == "1":
                matrix[ri][ci]["skip"] = True
                for back_r in range(ri - 1, -1, -1):
                    if not matrix[back_r][ci]["skip"]:
                        matrix[back_r][ci]["span_r"] += 1
                        break
            else:
                matrix[ri][ci]["elem"] = tc
                matrix[ri][ci]["span_c"] = grid_span
                matrix[ri][ci]["span_r"] = row_span

                # Mark covered span cells as skipped
                for dr in range(row_span):
                    for dc in range(grid_span):
                        if (dr > 0 or dc > 0) and (ri + dr < num_rows) and (ci + dc < num_cols):
                            matrix[ri + dr][ci + dc]["skip"] = True

            ci += grid_span

    # 4. Render Table Cells (Fills & Diagonal / Edge Borders)
    for ri in range(num_rows):
        for ci in range(num_cols):
            cell_info = matrix[ri][ci]
            if cell_info["skip"]:
                continue

            tc = cell_info["elem"]
            if tc is None:
                continue

            span_c = cell_info["span_c"]
            span_r = cell_info["span_r"]
            end_c = min(ci + span_c, len(xs) - 1)
            end_r = min(ri + span_r, len(ys) - 1)

            cx0, cy0 = xs[ci], ys[ri]
            cx1, cy1 = xs[end_c], ys[end_r]

            tcPr = tc.find(q(NS_A, "tcPr"))
            fill = None

            if tcPr is not None:
                if tcPr.find(q(NS_A, "noFill")) is not None:
                    fill = (0, 0, 0, 0)
                else:
                    sf = tcPr.find(q(NS_A, "solidFill"))
                    if sf is not None:
                        fill = resolve_element_color(sf, colors)

            # Default PowerPoint table banding fallback
            if fill is None:
                if ri == 0:
                    a = colors.get("accent1", (68, 114, 196))
                    fill = (a[0], a[1], a[2], 255)
                elif ri % 2 == 0:
                    fill = (237, 237, 238, 255)
                else:
                    fill = (255, 255, 255, 255)

            if fill and fill[3] > 0:
                draw.rectangle([cx0, cy0, cx1, cy1], fill=fill)

            # Borders (Left, Right, Top, Bottom, and Diagonal Crossings)
            if tcPr is not None:
                # 4 Main Edges
                for side, pa, pb in (
                    ("lnL", (cx0, cy0), (cx0, cy1)),
                    ("lnR", (cx1, cy0), (cx1, cy1)),
                    ("lnT", (cx0, cy0), (cx1, cy0)),
                    ("lnB", (cx0, cy1), (cx1, cy1)),
                ):
                    ln = tcPr.find(q(NS_A, side))
                    if ln is None or ln.find(q(NS_A, "noFill")) is not None:
                        continue
                    c = resolve_element_color(ln, colors) or (200, 200, 200, 255)
                    try:
                        w_emu = float(ln.attrib.get("w", 9525))
                    except ValueError:
                        w_emu = 9525.0
                    lw = max(1, int(round((w_emu / 12700.0) * scale)))
                    draw.line([pa, pb], fill=c, width=lw)

                # Diagonal Borders (Top-Left to Bottom-Right, Bottom-Left to Top-Right)
                lnTlToBr = tcPr.find(q(NS_A, "lnTlToBr"))
                if lnTlToBr is not None:
                    c = resolve_element_color(lnTlToBr, colors) or (180, 180, 180, 255)
                    draw.line([(cx0, cy0), (cx1, cy1)], fill=c, width=1)

                lnBlToTr = tcPr.find(q(NS_A, "lnBlToTr"))
                if lnBlToTr is not None:
                    c = resolve_element_color(lnBlToTr, colors) or (180, 180, 180, 255)
                    draw.line([(cx0, cy1), (cx1, cy0)], fill=c, width=1)

            # 5. Cell Text Body
            txbody = tc.find(q(NS_A, "txBody"))
            if txbody is not None:
                pad_l, pad_t, pad_r, pad_b = 4.0, 2.0, 4.0, 2.0
                if tcPr is not None:
                    try:
                        pad_l = float(tcPr.attrib.get("marL", 45720)) / 12700.0 * scale
                        pad_t = float(tcPr.attrib.get("marT", 22860)) / 12700.0 * scale
                        pad_r = float(tcPr.attrib.get("marR", 45720)) / 12700.0 * scale
                        pad_b = float(tcPr.attrib.get("marB", 22860)) / 12700.0 * scale
                    except ValueError:
                        pass
                padded_box = (cx0 + pad_l, cy0 + pad_t, cx1 - pad_r, cy1 - pad_b)
                render_txbody(txbody, padded_box, img, ctx)


class TableRenderer:
    """Facade for DrawingML table rendering."""

    @staticmethod
    def render(tbl: ET.Element, box: Tuple[float, float, float, float],
               img: Image.Image, ctx: Dict[str, Any]) -> None:
        render_table(tbl, box, img, ctx)
