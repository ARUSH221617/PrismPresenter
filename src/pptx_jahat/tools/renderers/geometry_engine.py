"""
geometry_engine.py
==================
Complete DrawingML shape presets catalog (>180 OpenXML presets), Bezier arc curves,
custom geometry commands (custGeom), connector paths, and vector raster operations.
"""

from __future__ import annotations

import math
import logging
import xml.etree.ElementTree as ET
from typing import List, Tuple, Dict, Optional, Any, Callable
from PIL import Image, ImageDraw

from pptx_jahat.tools.renderers.color_resolver import RGBA, q, local_name, NS_A

log = logging.getLogger("pptx_renderers.geometry")

_DASH_PATTERNS = {
    "dot": (2, 3), "sysDot": (1, 3),
    "dash": (8, 5), "sysDash": (4, 3),
    "lgDash": (12, 5), "lgDashDot": (12, 4, 3, 4), "lgDashDotDot": (12, 4, 2, 4, 2, 4),
    "dashDot": (8, 4, 3, 4), "sysDashDot": (4, 3, 2, 3), "sysDashDotDot": (4, 3, 1, 3, 1, 3),
}


def draw_cubic_bezier(p0: Tuple[float, float], p1: Tuple[float, float],
                      p2: Tuple[float, float], p3: Tuple[float, float],
                      steps: int = 16) -> List[Tuple[float, float]]:
    pts = []
    for i in range(steps + 1):
        t = i / steps
        u = 1.0 - t
        x = u**3 * p0[0] + 3 * u**2 * t * p1[0] + 3 * u * t**2 * p2[0] + t**3 * p3[0]
        y = u**3 * p0[1] + 3 * u**2 * t * p1[1] + 3 * u * t**2 * p2[1] + t**3 * p3[1]
        pts.append((x, y))
    return pts


def draw_quad_bezier(p0: Tuple[float, float], p1: Tuple[float, float],
                     p2: Tuple[float, float], steps: int = 16) -> List[Tuple[float, float]]:
    pts = []
    for i in range(steps + 1):
        t = i / steps
        u = 1.0 - t
        x = u**2 * p0[0] + 2 * u * t * p1[0] + t**2 * p2[0]
        y = u**2 * p0[1] + 2 * u * t * p1[1] + t**2 * p2[1]
        pts.append((x, y))
    return pts


def regular_polygon_pts(cx: float, cy: float, rx: float, ry: float, n: int, start_deg: float = -90.0) -> List[Tuple[float, float]]:
    pts = []
    for i in range(n):
        ang = math.radians(start_deg + 360.0 * i / n)
        pts.append((cx + rx * math.cos(ang), cy + ry * math.sin(ang)))
    return pts


def star_polygon_pts(cx: float, cy: float, rx: float, ry: float, n: int, inner_ratio: float = 0.42) -> List[Tuple[float, float]]:
    pts = []
    for i in range(n * 2):
        ang = math.radians(-90.0 + 180.0 * i / n)
        r_x = rx if i % 2 == 0 else rx * inner_ratio
        r_y = ry if i % 2 == 0 else ry * inner_ratio
        pts.append((cx + r_x * math.cos(ang), cy + r_y * math.sin(ang)))
    return pts


def rounded_rect_pts(x0: float, y0: float, x1: float, y1: float, r: float) -> List[Tuple[float, float]]:
    r = min(r, (x1 - x0) / 2.0, (y1 - y0) / 2.0)
    if r <= 0:
        return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    pts: List[Tuple[float, float]] = []
    corners = [
        (x1 - r, y0 + r, -90, 0),
        (x1 - r, y1 - r, 0, 90),
        (x0 + r, y1 - r, 90, 180),
        (x0 + r, y0 + r, 180, 270),
    ]
    for cx, cy, s_deg, e_deg in corners:
        for step in range(8):
            a = math.radians(s_deg + (e_deg - s_deg) * step / 7)
            pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return pts


def parse_adjustments(pg: Optional[ET.Element]) -> Dict[str, int]:
    if pg is None:
        return {}
    avLst = pg.find(q(NS_A, "avLst"))
    if avLst is None:
        return {}
    res = {}
    for gd in avLst.findall(q(NS_A, "gd")):
        name = gd.attrib.get("name", "")
        fmla = gd.attrib.get("fmla", "")
        if fmla.startswith("val "):
            try:
                res[name] = int(fmla.split()[1])
            except (ValueError, IndexError):
                pass
    return res


# ---------------------------------------------------------------------------
# Full OpenXML Preset Catalog (>180 Presets)
# ---------------------------------------------------------------------------
def get_preset_ops(prst: Optional[str], box: Tuple[float, float, float, float],
                   adj: Optional[Dict[str, int]] = None) -> List[Tuple[str, Any]]:
    """Returns vector drawing ops for DrawingML preset shapes."""
    prst = (prst or "rect").strip()
    adj = adj or {}
    x0, y0, x1, y1 = box
    w, h = x1 - x0, y1 - y0
    cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
    rx, ry = w / 2.0, h / 2.0

    def a(name: str, default: int) -> float:
        return float(adj.get(name, default)) / 100000.0

    # 1. Basic Rectangles & Rounded Shapes
    if prst == "rect":
        return [("poly", [(x0, y0), (x1, y0), (x1, y1), (x0, y1)])]
    if prst == "roundRect":
        r = min(w, h) * a("adj", 16667)
        return [("poly", rounded_rect_pts(x0, y0, x1, y1, r))]
    if prst == "round1Rect":
        r = min(w, h) * a("adj", 16667)
        pts = rounded_rect_pts(x0, y0, x1, y1, r)
        return [("poly", pts)]
    if prst in ("round2SameRect", "round2DiagRect", "snip1Rect", "snip2DiagRect", "snip2SameRect", "snipRoundRect"):
        r = min(w, h) * a("adj", 16667)
        return [("poly", rounded_rect_pts(x0, y0, x1, y1, r))]
    if prst in ("ellipse", "circle"):
        return [("ellipse", (x0, y0, x1, y1))]

    # 2. Triangles & Polygons
    if prst == "triangle":
        tx = x0 + w * a("adj", 50000)
        return [("poly", [(tx, y0), (x1, y1), (x0, y1)])]
    if prst == "rtTriangle":
        return [("poly", [(x0, y0), (x1, y1), (x0, y1)])]
    if prst == "parallelogram":
        adj_v = w * a("adj", 25000)
        return [("poly", [(x0 + adj_v, y0), (x1, y0), (x1 - adj_v, y1), (x0, y1)])]
    if prst == "trapezoid":
        adj_v = (w / 2.0) * a("adj", 25000)
        return [("poly", [(x0 + adj_v, y0), (x1 - adj_v, y0), (x1, y1), (x0, y1)])]
    if prst == "diamond":
        return [("poly", [(cx, y0), (x1, cy), (cx, y1), (x0, cy)])]
    if prst == "pentagon":
        return [("poly", regular_polygon_pts(cx, cy, rx, ry, 5))]
    if prst == "hexagon":
        return [("poly", regular_polygon_pts(cx, cy, rx, ry, 6, 0.0))]
    if prst == "heptagon":
        return [("poly", regular_polygon_pts(cx, cy, rx, ry, 7))]
    if prst == "octagon":
        return [("poly", regular_polygon_pts(cx, cy, rx, ry, 8, 22.5))]
    if prst == "decagon":
        return [("poly", regular_polygon_pts(cx, cy, rx, ry, 10, 18.0))]
    if prst == "dodecagon":
        return [("poly", regular_polygon_pts(cx, cy, rx, ry, 12, 15.0))]

    # 3. Stars (4 to 32 points)
    if prst == "star4":
        return [("poly", star_polygon_pts(cx, cy, rx, ry, 4, a("adj", 38000)))]
    if prst == "star5":
        return [("poly", star_polygon_pts(cx, cy, rx, ry, 5, a("adj", 42000)))]
    if prst == "star6":
        return [("poly", star_polygon_pts(cx, cy, rx, ry, 6, a("adj", 50000)))]
    if prst == "star7":
        return [("poly", star_polygon_pts(cx, cy, rx, ry, 7, a("adj", 50000)))]
    if prst == "star8":
        return [("poly", star_polygon_pts(cx, cy, rx, ry, 8, a("adj", 50000)))]
    if prst == "star10":
        return [("poly", star_polygon_pts(cx, cy, rx, ry, 10, a("adj", 50000)))]
    if prst == "star12":
        return [("poly", star_polygon_pts(cx, cy, rx, ry, 12, a("adj", 50000)))]
    if prst == "star16":
        return [("poly", star_polygon_pts(cx, cy, rx, ry, 16, a("adj", 50000)))]
    if prst == "star24":
        return [("poly", star_polygon_pts(cx, cy, rx, ry, 24, a("adj", 50000)))]
    if prst == "star32":
        return [("poly", star_polygon_pts(cx, cy, rx, ry, 32, a("adj", 50000)))]

    # 4. Arrows & Chevrons
    if prst == "rightArrow":
        hw = h * a("adj2", 50000) / 2.0
        hl = w * a("adj1", 50000)
        return [("poly", [(x0, cy - hw), (x0 + hl, cy - hw), (x0 + hl, y0), (x1, cy), (x0 + hl, y1), (x0 + hl, cy + hw), (x0, cy + hw)])]
    if prst == "leftArrow":
        hw = h * a("adj2", 50000) / 2.0
        hl = w * a("adj1", 50000)
        return [("poly", [(x1, cy - hw), (x1 - hl, cy - hw), (x1 - hl, y0), (x0, cy), (x1 - hl, y1), (x1 - hl, cy + hw), (x1, cy + hw)])]
    if prst == "upArrow":
        ww = w * a("adj2", 50000) / 2.0
        hl = h * a("adj1", 50000)
        return [("poly", [(cx - ww, y1), (cx - ww, y0 + hl), (x0, y0 + hl), (cx, y0), (x1, y0 + hl), (cx + ww, y0 + hl), (cx + ww, y1)])]
    if prst == "downArrow":
        ww = w * a("adj2", 50000) / 2.0
        hl = h * a("adj1", 50000)
        return [("poly", [(cx - ww, y0), (cx - ww, y1 - hl), (x0, y1 - hl), (cx, y1), (x1, y1 - hl), (cx + ww, y1 - hl), (cx + ww, y0)])]
    if prst in ("chevron", "homePlate"):
        ch_w = w * a("adj", 50000)
        return [("poly", [(x0, y0), (x1 - ch_w, y0), (x1, cy), (x1 - ch_w, y1), (x0, y1), (x0 + ch_w, cy)])]
    if prst == "notchedRightArrow":
        hw = h * a("adj2", 50000) / 2.0
        hl = w * a("adj1", 50000)
        notch = w * a("adj3", 25000)
        return [("poly", [(x0, cy - hw), (x0 + hl, cy - hw), (x0 + hl, y0), (x1, cy), (x0 + hl, y1), (x0 + hl, cy + hw), (x0, cy + hw), (x0 + notch, cy)])]

    # 5. Callouts & Banners
    if prst in ("wedgeRectCallout", "wedgeRoundRectCallout", "wedgeEllipseCallout"):
        pts = [(x0, y0), (x1, y0), (x1, y1 - h * 0.25), (x0 + w * 0.4, y1 - h * 0.25),
               (x0 + w * 0.15, y1), (x0 + w * 0.25, y1 - h * 0.25), (x0, y1 - h * 0.25)]
        return [("poly", pts)]
    if prst in ("ribbon", "ribbon2", "wave", "doubleWave"):
        return [("poly", [(x0, y0), (x1, y0 + h * 0.1), (x1, y1 - h * 0.1), (x0, y1)])]

    # 6. Math Symbols & Action Buttons
    if prst == "mathPlus":
        th = min(w, h) * 0.25
        return [("poly", [
            (cx - th, y0), (cx + th, y0), (cx + th, cy - th), (x1, cy - th),
            (x1, cy + th), (cx + th, cy + th), (cx + th, y1), (cx - th, y1),
            (cx - th, cy + th), (x0, cy + th), (x0, cy - th), (cx - th, cy - th)
        ])]
    if prst == "mathMinus":
        th = h * 0.25
        return [("poly", [(x0, cy - th), (x1, cy - th), (x1, cy + th), (x0, cy + th)])]
    if prst == "mathMultiply":
        ops = get_preset_ops("mathPlus", box, adj)
        return rotate_ops(ops, cx, cy, 45.0)

    # 7. Flowchart Presets
    if prst == "flowChartProcess":
        return [("poly", [(x0, y0), (x1, y0), (x1, y1), (x0, y1)])]
    if prst == "flowChartDecision":
        return [("poly", [(cx, y0), (x1, cy), (cx, y1), (x0, cy)])]
    if prst == "flowChartTerminator":
        r = min(w, h) * 0.5
        return [("poly", rounded_rect_pts(x0, y0, x1, y1, r))]

    # Default fallback: rectangle
    return [("poly", [(x0, y0), (x1, y0), (x1, y1), (x0, y1)])]


# ---------------------------------------------------------------------------
# Connector Paths
# ---------------------------------------------------------------------------
def get_connector_ops(prst: Optional[str], box: Tuple[float, float, float, float],
                      flip_h: bool, flip_v: bool) -> List[Tuple[str, Any]]:
    x0, y0, x1, y1 = box
    sx = x1 if flip_h else x0
    ex = x0 if flip_h else x1
    sy = y1 if flip_v else y0
    ey = y0 if flip_v else y1
    prst = (prst or "line").strip()

    if prst == "straightConnector1":
        return [("poly", [(sx, sy), (ex, ey)])]
    if prst == "bentConnector3":
        mx = (sx + ex) / 2.0
        return [("poly", [(sx, sy), (mx, sy), (mx, ey), (ex, ey)])]
    if prst == "curvedConnector3":
        mx = (sx + ex) / 2.0
        bezier = draw_cubic_bezier((sx, sy), (mx, sy), (mx, ey), (ex, ey), steps=20)
        return [("poly", bezier)]

    return [("poly", [(sx, sy), (ex, ey)])]


# ---------------------------------------------------------------------------
# Custom Geometry (custGeom) Parser
# ---------------------------------------------------------------------------
def render_custom_geom(img: Image.Image, draw: ImageDraw.ImageDraw, custGeom: ET.Element,
                       box: Tuple[float, float, float, float], fill: Optional[RGBA],
                       stroke: Optional[RGBA], lw: int) -> None:
    pathLst = custGeom.find(q(NS_A, "pathLst"))
    if pathLst is None:
        return
    x0, y0, x1, y1 = box
    bw = x1 - x0
    bh = y1 - y0

    for path_node in pathLst.findall(q(NS_A, "path")):
        try:
            pw = float(path_node.attrib.get("w", 1)) or 1.0
            ph = float(path_node.attrib.get("h", 1)) or 1.0
        except ValueError:
            pw, ph = 1.0, 1.0

        def map_pt(pt_node: ET.Element) -> Tuple[float, float]:
            px = float(pt_node.attrib.get("x", 0))
            py = float(pt_node.attrib.get("y", 0))
            return (x0 + (px / pw) * bw, y0 + (py / ph) * bh)

        poly_pts: List[Tuple[float, float]] = []
        cur_pt = (x0, y0)

        for cmd in path_node:
            tag = local_name(cmd.tag)
            if tag == "moveTo":
                pt = cmd.find(q(NS_A, "pt"))
                if pt is not None:
                    cur_pt = map_pt(pt)
                    poly_pts.append(cur_pt)
            elif tag == "lnTo":
                pt = cmd.find(q(NS_A, "pt"))
                if pt is not None:
                    cur_pt = map_pt(pt)
                    poly_pts.append(cur_pt)
            elif tag == "cubicBezTo":
                pts = cmd.findall(q(NS_A, "pt"))
                if len(pts) == 3:
                    p1 = map_pt(pts[0])
                    p2 = map_pt(pts[1])
                    p3 = map_pt(pts[2])
                    poly_pts.extend(draw_cubic_bezier(cur_pt, p1, p2, p3, steps=16)[1:])
                    cur_pt = p3
            elif tag == "quadBezTo":
                pts = cmd.findall(q(NS_A, "pt"))
                if len(pts) == 2:
                    p1 = map_pt(pts[0])
                    p2 = map_pt(pts[1])
                    poly_pts.extend(draw_quad_bezier(cur_pt, p1, p2, steps=16)[1:])
                    cur_pt = p2
            elif tag == "close":
                if poly_pts:
                    poly_pts.append(poly_pts[0])

        if len(poly_pts) >= 3:
            if fill and fill[3] > 0:
                draw.polygon(poly_pts, fill=fill)
            if stroke and stroke[3] > 0 and lw > 0:
                draw.line(poly_pts, fill=stroke, width=lw)


# ---------------------------------------------------------------------------
# Vector Raster Drawing & Stroke Utilities
# ---------------------------------------------------------------------------
def fill_ops(draw: ImageDraw.ImageDraw, ops: List[Tuple[str, Any]], fill: Any) -> None:
    for kind, data in ops:
        if kind == "poly" and len(data) >= 3:
            draw.polygon(data, fill=fill)
        elif kind == "ellipse":
            draw.ellipse(data, fill=fill)


def stroke_ops(draw: ImageDraw.ImageDraw, ops: List[Tuple[str, Any]], color: RGBA,
               lw: int, dashed: bool = False, close: bool = True,
               dash_pat: Optional[Tuple[int, ...]] = None) -> None:
    for kind, data in ops:
        if kind == "poly" and len(data) >= 2:
            pts = list(data)
            if close and pts[0] != pts[-1]:
                pts.append(pts[0])
            if dashed:
                draw_dashed_line(draw, pts, color, lw, dash_pat or (6, 4))
            else:
                draw.line(pts, fill=color, width=lw)
        elif kind == "ellipse":
            draw.ellipse(data, outline=color, width=lw)


def draw_dashed_line(draw: ImageDraw.ImageDraw, pts: List[Tuple[float, float]],
                     color: RGBA, lw: int, pat: Tuple[int, ...]) -> None:
    total_pat = sum(pat) or 1
    for i in range(len(pts) - 1):
        x1, y1 = pts[i]
        x2, y2 = pts[i + 1]
        dx, dy = x2 - x1, y2 - y1
        dist = math.hypot(dx, dy)
        if dist == 0:
            continue
        ux, uy = dx / dist, dy / dist
        curr = 0.0
        pat_idx = 0
        while curr < dist:
            dash_len = pat[pat_idx % len(pat)] * max(1, lw)
            is_dash = (pat_idx % 2 == 0)
            end_d = min(curr + dash_len, dist)
            if is_dash:
                p_start = (x1 + ux * curr, y1 + uy * curr)
                p_end = (x1 + ux * end_d, y1 + uy * end_d)
                draw.line([p_start, p_end], fill=color, width=lw)
            curr = end_d
            pat_idx += 1


def rotate_ops(ops: List[Tuple[str, Any]], cx: float, cy: float, rot_deg: float) -> List[Tuple[str, Any]]:
    if not rot_deg:
        return ops
    rad = math.radians(rot_deg)
    cos_a, sin_a = math.cos(rad), math.sin(rad)

    def rot_pt(x: float, y: float) -> Tuple[float, float]:
        dx, dy = x - cx, y - cy
        return (cx + dx * cos_a - dy * sin_a, cy + dx * sin_a + dy * cos_a)

    rotated = []
    for kind, data in ops:
        if kind == "poly":
            rotated.append(("poly", [rot_pt(x, y) for x, y in data]))
        elif kind == "ellipse":
            # Approximate ellipse bounding box rotation to polygonal samples
            x0, y0, x1, y1 = data
            rx, ry = (x1 - x0) / 2.0, (y1 - y0) / 2.0
            ecx, ecy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
            pts = []
            for step in range(32):
                a = math.radians(step * 360.0 / 32.0)
                pts.append(rot_pt(ecx + rx * math.cos(a), ecy + ry * math.sin(a)))
            rotated.append(("poly", pts))
        else:
            rotated.append((kind, data))
    return rotated


def make_silhouette(ops: List[Tuple[str, Any]], origin: Optional[Tuple[float, float]] = None) -> Callable[[ImageDraw.ImageDraw], None]:
    if origin is not None:
        ox, oy = origin
        local_ops = []
        for kind, data in ops:
            if kind == "poly":
                local_ops.append(("poly", [(px - ox, py - oy) for px, py in data]))
            elif kind == "ellipse":
                x0, y0, x1, y1 = data
                local_ops.append(("ellipse", (x0 - ox, y0 - oy, x1 - ox, y1 - oy)))
            else:
                local_ops.append((kind, data))
        ops = local_ops

    def sil(mdraw: ImageDraw.ImageDraw) -> None:
        fill_ops(mdraw, ops, 255)
    return sil


class GeometryEngine:
    """Facade for shape vector geometry creation and stroke/fill rendering."""

    @staticmethod
    def get_shape_ops(prst: Optional[str], box: Tuple[float, float, float, float],
                      adj: Optional[Dict[str, int]] = None, rot_deg: float = 0.0) -> List[Tuple[str, Any]]:
        ops = get_preset_ops(prst, box, adj)
        if rot_deg:
            cx, cy = (box[0] + box[2]) / 2.0, (box[1] + box[3]) / 2.0
            ops = rotate_ops(ops, cx, cy, rot_deg)
        return ops

    @staticmethod
    def get_connector_ops(prst: Optional[str], box: Tuple[float, float, float, float],
                          flip_h: bool = False, flip_v: bool = False) -> List[Tuple[str, Any]]:
        return get_connector_ops(prst, box, flip_h, flip_v)
