"""
gradient_engine.py
==================
Advanced DrawingML gradient fills (linear, radial, rectangular, shape-contour paths),
specular 3D bevel lighting, and comprehensive shape filter effects (outer/inner shadow, glow, soft edge, reflection).
"""

from __future__ import annotations

import math
import logging
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple, Optional, Any, Callable
from PIL import Image, ImageDraw, ImageFilter, ImageChops

from pptx_jahat.tools.renderers.color_resolver import (
    RGBA, RGB, q, local_name, resolve_element_color, clamp, NS_A
)

log = logging.getLogger("pptx_renderers.gradient")


def parse_gradient(gradFill: ET.Element, colors: Dict[str, RGB],
                   ph_color: Optional[RGB] = None) -> Tuple[List[Tuple[float, RGBA]], float, Dict[str, Any]]:
    """
    Parses a `gradFill` element. Returns:
    - stops: List of (position 0.0..1.0, RGBA)
    - angle_deg: Linear angle in degrees (default 0.0)
    - info: Dict containing gradient path type ('linear', 'circle', 'rect', 'shape') and fillToRect specs.
    """
    stops: List[Tuple[float, RGBA]] = []
    gsLst = gradFill.find(q(NS_A, "gsLst"))
    if gsLst is not None:
        for gs in gsLst.findall(q(NS_A, "gs")):
            pos = float(gs.attrib.get("pos", 0)) / 100000.0
            c = resolve_element_color(gs, colors, ph_color)
            if c is not None:
                stops.append((clamp(pos, 0.0, 1.0), c))
    stops.sort(key=lambda s: s[0])
    if not stops:
        stops = [(0.0, (255, 255, 255, 255)), (1.0, (180, 180, 180, 255))]
    elif len(stops) == 1:
        stops = [(0.0, stops[0][1]), (1.0, stops[0][1])]
    else:
        if stops[0][0] > 0.0:
            stops.insert(0, (0.0, stops[0][1]))
        if stops[-1][0] < 1.0:
            stops.append((1.0, stops[-1][1]))

    angle_deg = 0.0
    info: Dict[str, Any] = {"type": "linear", "fill_to_rect": None}

    lin = gradFill.find(q(NS_A, "lin"))
    if lin is not None:
        try:
            angle_deg = float(lin.attrib.get("ang", 0)) / 60000.0
        except ValueError:
            angle_deg = 0.0
        info["type"] = "linear"

    path_elem = gradFill.find(q(NS_A, "path"))
    if path_elem is not None:
        p_type = path_elem.attrib.get("path", "circle")
        info["type"] = p_type
        fill_to = path_elem.find(q(NS_A, "fillToRect"))
        if fill_to is not None:
            info["fill_to_rect"] = {
                "l": float(fill_to.attrib.get("l", 50000)) / 100000.0,
                "t": float(fill_to.attrib.get("t", 50000)) / 100000.0,
                "r": float(fill_to.attrib.get("r", 50000)) / 100000.0,
                "b": float(fill_to.attrib.get("b", 50000)) / 100000.0,
            }

    return stops, angle_deg, info


def interp_stops(stops: List[Tuple[float, RGBA]], t: float) -> RGBA:
    t = clamp(t, 0.0, 1.0)
    for i in range(len(stops) - 1):
        p0, c0 = stops[i]
        p1, c1 = stops[i + 1]
        if p0 <= t <= p1:
            span = p1 - p0
            k = 0.0 if span == 0 else (t - p0) / span
            return (
                int(round(c0[0] + (c1[0] - c0[0]) * k)),
                int(round(c0[1] + (c1[1] - c0[1]) * k)),
                int(round(c0[2] + (c1[2] - c0[2]) * k)),
                int(round(c0[3] + (c1[3] - c0[3]) * k)),
            )
    return stops[-1][1]


def paint_gradient(img: Image.Image, box: Tuple[float, float, float, float],
                   stops: List[Tuple[float, RGBA]], angle_deg: float = 0.0,
                   silhouette_fn: Optional[Callable[[ImageDraw.ImageDraw], None]] = None,
                   grad_info: Optional[Dict[str, Any]] = None) -> None:
    """
    Renders high-quality linear, radial, rectangular, or path gradients into `img` within `box`,
    masked by `silhouette_fn`.
    """
    x0, y0, x1, y1 = box
    bw = int(math.ceil(x1 - x0))
    bh = int(math.ceil(y1 - y0))
    if bw <= 0 or bh <= 0:
        return

    g_type = grad_info.get("type", "linear") if grad_info else "linear"

    # Pre-render a downsampled gradient ramp for performance, then scale up
    sample_w = min(bw, 128)
    sample_h = min(bh, 128)
    grad = Image.new("RGBA", (sample_w, sample_h))
    pixels = grad.load()

    if g_type == "circle":
        # Radial gradient from center or focus point
        fill_to = grad_info.get("fill_to_rect") if grad_info else None
        cx = (fill_to["l"] + 1.0 - fill_to["r"]) / 2.0 * sample_w if fill_to else sample_w / 2.0
        cy = (fill_to["t"] + 1.0 - fill_to["b"]) / 2.0 * sample_h if fill_to else sample_h / 2.0
        max_r = math.hypot(max(cx, sample_w - cx), max(cy, sample_h - cy)) or 1.0
        for py in range(sample_h):
            for px in range(sample_w):
                dist = math.hypot(px - cx, py - cy)
                t = clamp(dist / max_r, 0.0, 1.0)
                pixels[px, py] = interp_stops(stops, t)

    elif g_type in ("rect", "shape"):
        # Rectangular / Box distance gradient
        fill_to = grad_info.get("fill_to_rect") if grad_info else None
        cx = sample_w * 0.5
        cy = sample_h * 0.5
        for py in range(sample_h):
            norm_y = abs(py - cy) / (sample_h * 0.5 or 1.0)
            for px in range(sample_w):
                norm_x = abs(px - cx) / (sample_w * 0.5 or 1.0)
                t = clamp(max(norm_x, norm_y), 0.0, 1.0)
                pixels[px, py] = interp_stops(stops, t)

    else:
        # Linear angled gradient
        rad = math.radians(angle_deg - 90.0)
        ux = math.cos(rad)
        uy = math.sin(rad)
        corners = [
            (0, 0),
            (sample_w, 0),
            (0, sample_h),
            (sample_w, sample_h)
        ]
        projs = [c[0] * ux + c[1] * uy for c in corners]
        min_p, max_p = min(projs), max(projs)
        span = max_p - min_p or 1.0

        for py in range(sample_h):
            row_dot = py * uy
            for px in range(sample_w):
                proj = px * ux + row_dot
                t = (proj - min_p) / span
                pixels[px, py] = interp_stops(stops, t)

    grad_full = grad.resize((bw, bh), Image.Resampling.BILINEAR)

    # Apply shape silhouette mask
    if silhouette_fn is not None:
        mask = Image.new("L", (bw, bh), 0)
        mdraw = ImageDraw.Draw(mask)
        # Shift coordinates to local box origin
        def local_draw(d):
            silhouette_fn(d)
        silhouette_fn(mdraw)
        
        # Check if silhouette mask was drawn at global coordinates or local
        # If mask is empty, try painting directly onto local bbox
        bbox = mask.getbbox()
        if not bbox:
            mdraw.rectangle([0, 0, bw, bh], fill=255)
            
        img.paste(grad_full, (int(x0), int(y0)), mask)
    else:
        img.paste(grad_full, (int(x0), int(y0)), grad_full)


def paint_3d_bevel(img: Image.Image, box: Tuple[float, float, float, float],
                   bevel_t: Optional[ET.Element], bevel_b: Optional[ET.Element],
                   scale: float) -> None:
    """
    Simulates DrawingML 3D Bevel and specular edge lighting highlights.
    """
    if bevel_t is None and bevel_b is None:
        return
    x0, y0, x1, y1 = box
    bw = int(round(x1 - x0))
    bh = int(round(y1 - y0))
    if bw <= 4 or bh <= 4:
        return

    # Parse top bevel width/height
    w_emu = float(bevel_t.attrib.get("w", 76200)) if bevel_t is not None else 76200
    h_emu = float(bevel_t.attrib.get("h", 76200)) if bevel_t is not None else 76200
    depth = max(1, int(round((w_emu / 12700.0) * scale * 0.5)))
    depth = min(depth, min(bw, bh) // 3)

    # Create specular highlight & shadow overlays
    overlay = Image.new("RGBA", (bw, bh), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    # Top & Left highlight (white with gradient alpha)
    for i in range(depth):
        alpha = int(round(120 * (1.0 - i / depth)))
        draw.line([(i, i), (bw - 1 - i, i)], fill=(255, 255, 255, alpha))
        draw.line([(i, i), (i, bh - 1 - i)], fill=(255, 255, 255, alpha))

    # Bottom & Right shadow (black with gradient alpha)
    for i in range(depth):
        alpha = int(round(100 * (1.0 - i / depth)))
        draw.line([(i, bh - 1 - i), (bw - 1 - i, bh - 1 - i)], fill=(0, 0, 0, alpha))
        draw.line([(bw - 1 - i, i), (bw - 1 - i, bh - 1 - i)], fill=(0, 0, 0, alpha))

    img.alpha_composite(overlay, (int(x0), int(y0)))


# ---------------------------------------------------------------------------
# Shape Effects (Shadows, Glow, Inner Shadow, Reflection, Soft Edge)
# ---------------------------------------------------------------------------
def parse_effect_lst(spPr: Optional[ET.Element], colors: Dict[str, RGB], scale: float) -> Dict[str, Any]:
    effects: Dict[str, Any] = {}
    if spPr is None:
        return effects
    el = spPr.find(q(NS_A, "effectLst"))
    if el is None:
        return effects

    # 1. outerShdw
    sh = el.find(q(NS_A, "outerShdw"))
    if sh is not None:
        col = resolve_element_color(sh, colors) or (0, 0, 0, 100)
        blur = float(sh.attrib.get("blurRad", 0)) / 12700.0 * scale
        dist = float(sh.attrib.get("dist", 0)) / 12700.0 * scale
        ang = math.radians(float(sh.attrib.get("dir", 0)) / 60000.0)
        effects["outerShdw"] = {
            "color": col, "blur": max(0.5, blur),
            "dx": dist * math.cos(ang), "dy": dist * math.sin(ang),
        }

    # 2. glow
    gl = el.find(q(NS_A, "glow"))
    if gl is not None:
        col = resolve_element_color(gl, colors) or (255, 255, 100, 150)
        rad = float(gl.attrib.get("rad", 0)) / 12700.0 * scale
        effects["glow"] = {"color": col, "rad": max(1.0, rad)}

    # 3. innerShdw
    ish = el.find(q(NS_A, "innerShdw"))
    if ish is not None:
        col = resolve_element_color(ish, colors) or (0, 0, 0, 120)
        blur = float(ish.attrib.get("blurRad", 0)) / 12700.0 * scale
        dist = float(ish.attrib.get("dist", 0)) / 12700.0 * scale
        ang = math.radians(float(ish.attrib.get("dir", 0)) / 60000.0)
        effects["innerShdw"] = {
            "color": col, "blur": max(0.5, blur),
            "dx": dist * math.cos(ang), "dy": dist * math.sin(ang),
        }

    # 4. softEdge
    se = el.find(q(NS_A, "softEdge"))
    if se is not None:
        rad = float(se.attrib.get("rad", 0)) / 12700.0 * scale
        effects["softEdge"] = {"rad": max(1.0, rad)}

    # 5. reflection
    refl = el.find(q(NS_A, "reflection"))
    if refl is not None:
        dist = float(refl.attrib.get("dist", 0)) / 12700.0 * scale
        stA = float(refl.attrib.get("stA", 50000)) / 100000.0
        endA = float(refl.attrib.get("endA", 300)) / 100000.0
        sy = float(refl.attrib.get("sy", 100000)) / 100000.0
        effects["reflection"] = {"dist": dist, "stA": stA, "endA": endA, "sy": sy}

    # 6. 3D bevels
    bevT = spPr.find(f".//{q(NS_A, 'bevelT')}")
    bevB = spPr.find(f".//{q(NS_A, 'bevelB')}")
    if bevT is not None or bevB is not None:
        effects["3dBevel"] = {"bevelT": bevT, "bevelB": bevB}

    return effects


def paint_shadow(img: Image.Image, ops: List[Any], sh: Dict[str, Any]) -> None:
    w, h = img.size
    shadow_layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(shadow_layer)
    from pptx_jahat.tools.renderers.geometry_engine import fill_ops
    fill_ops(draw, ops, sh["color"])
    blur_rad = max(1.0, sh.get("blur", 3.0))
    blurred = shadow_layer.filter(ImageFilter.GaussianBlur(blur_rad))
    dx = int(round(sh.get("dx", 2.0)))
    dy = int(round(sh.get("dy", 2.0)))
    img.alpha_composite(blurred, (dx, dy))


def paint_glow(img: Image.Image, ops: List[Any], gl: Dict[str, Any]) -> None:
    w, h = img.size
    glow_layer = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(glow_layer)
    from pptx_jahat.tools.renderers.geometry_engine import fill_ops
    fill_ops(draw, ops, gl["color"])
    blurred = glow_layer.filter(ImageFilter.GaussianBlur(gl["rad"]))
    img.alpha_composite(blurred)


def paint_inner_shadow(img: Image.Image, ops: List[Any], sh: Dict[str, Any]) -> None:
    w, h = img.size
    mask = Image.new("L", (w, h), 0)
    mdraw = ImageDraw.Draw(mask)
    from pptx_jahat.tools.renderers.geometry_engine import fill_ops
    fill_ops(mdraw, ops, 255)

    inv_mask = ImageChops.invert(mask)
    shifted_inv = Image.new("L", (w, h), 255)
    dx = int(round(sh.get("dx", 2.0)))
    dy = int(round(sh.get("dy", 2.0)))
    shifted_inv.paste(inv_mask, (dx, dy))

    blur_rad = max(1.0, sh.get("blur", 3.0))
    blurred_inv = shifted_inv.filter(ImageFilter.GaussianBlur(blur_rad))
    inner_shdw_mask = ImageChops.multiply(mask, blurred_inv)

    c = sh["color"]
    shdw_img = Image.new("RGBA", (w, h), (c[0], c[1], c[2], c[3]))
    img.paste(shdw_img, (0, 0), inner_shdw_mask)


def paint_soft_edge(img_layer: Image.Image, rad: float) -> None:
    if rad <= 0:
        return
    alpha = img_layer.split()[3]
    blurred_alpha = alpha.filter(ImageFilter.GaussianBlur(rad))
    img_layer.putalpha(blurred_alpha)


def paint_reflection(img: Image.Image, box: Tuple[float, float, float, float], refl: Dict[str, Any]) -> None:
    x0, y0, x1, y1 = box
    bw = int(round(x1 - x0))
    bh = int(round(y1 - y0))
    if bw <= 0 or bh <= 0:
        return
    refl_h = int(round(bh * refl["sy"]))
    if refl_h <= 0:
        return
    try:
        cropped = img.crop((int(x0), int(y0), int(x1), int(y1)))
        flipped = cropped.transpose(Image.Transpose.FLIP_TOP_BOTTOM).resize((bw, refl_h), Image.Resampling.BILINEAR)
        alpha_mask = Image.new("L", (bw, refl_h))
        for py in range(refl_h):
            t = py / refl_h
            a = int(round((refl["stA"] + (refl["endA"] - refl["stA"]) * t) * 255))
            for px in range(bw):
                alpha_mask.putpixel((px, py), a)
        flipped.putalpha(ImageChops.multiply(flipped.split()[3], alpha_mask))
        img.alpha_composite(flipped, (int(x0), int(y1 + refl["dist"])))
    except Exception as e:
        log.debug(f"Failed to paint reflection: {e}")


class GradientEngine:
    """Facade for rendering advanced gradient fills and effects."""

    @staticmethod
    def render_gradient(img: Image.Image, box: Tuple[float, float, float, float],
                        stops: List[Tuple[float, RGBA]], angle_deg: float = 0.0,
                        silhouette_fn: Optional[Callable[[ImageDraw.ImageDraw], None]] = None,
                        grad_info: Optional[Dict[str, Any]] = None) -> None:
        paint_gradient(img, box, stops, angle_deg, silhouette_fn, grad_info)

    @staticmethod
    def render_effects(img: Image.Image, ops: List[Any], box: Tuple[float, float, float, float],
                       effects: Dict[str, Any], scale: float) -> None:
        if effects.get("outerShdw"):
            paint_shadow(img, ops, effects["outerShdw"])
        if effects.get("glow"):
            paint_glow(img, ops, effects["glow"])
        if effects.get("innerShdw"):
            paint_inner_shadow(img, ops, effects["innerShdw"])
        if effects.get("3dBevel"):
            paint_3d_bevel(img, box, effects["3dBevel"]["bevelT"], effects["3dBevel"]["bevelB"], scale)
        if effects.get("reflection"):
            paint_reflection(img, box, effects["reflection"])
