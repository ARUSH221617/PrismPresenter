"""
media_resolver.py
=================
Vector Media & Multimedia Resolver:
- Direct SVG rasterization (cairosvg / svglib / Pillow)
- Windows Metafiles (EMF / WMF) conversion and rasterization
- Video & Audio poster frames with playback badge compositing.
"""

from __future__ import annotations

import io
import logging
from typing import Tuple, Optional
from PIL import Image, ImageDraw

log = logging.getLogger("pptx_renderers.media")


def rasterize_svg(svg_bytes: bytes, target_w: int, target_h: int) -> Optional[Image.Image]:
    """Rasterizes SVG vector data at high resolution using cairosvg / svglib / PIL."""
    if not svg_bytes:
        return None

    # Method 1: cairosvg
    try:
        import cairosvg
        png_data = cairosvg.svg2png(bytestring=svg_bytes, output_width=target_w, output_height=target_h)
        if png_data:
            return Image.open(io.BytesIO(png_data)).convert("RGBA")
    except Exception as e:
        log.debug(f"cairosvg unavailable or failed: {e}")

    # Method 2: svglib + reportlab
    try:
        from svglib.svglib import svg2rlg
        from reportlab.graphics import renderPM
        drawing = svg2rlg(io.BytesIO(svg_bytes))
        if drawing:
            buf = io.BytesIO()
            renderPM.drawToFile(drawing, buf, fmt="PNG", dpi=150)
            buf.seek(0)
            return Image.open(buf).convert("RGBA").resize((target_w, target_h), Image.Resampling.BILINEAR)
    except Exception as e:
        log.debug(f"svglib fallback failed: {e}")

    # Method 3: Direct Pillow SVG loader (if supported by plugin)
    try:
        return Image.open(io.BytesIO(svg_bytes)).convert("RGBA").resize((target_w, target_h), Image.Resampling.BILINEAR)
    except Exception:
        pass

    return None


def rasterize_emf_wmf(metafile_bytes: bytes, target_w: int, target_h: int) -> Optional[Image.Image]:
    """Decodes EMF/WMF vector metafiles using Pillow / pillow-wmf."""
    if not metafile_bytes:
        return None
    try:
        img = Image.open(io.BytesIO(metafile_bytes))
        return img.convert("RGBA").resize((target_w, target_h), Image.Resampling.BILINEAR)
    except Exception as e:
        log.debug(f"Pillow EMF/WMF decode failed: {e}")
        return None


def paint_media_poster(img: Image.Image, box: Tuple[float, float, float, float],
                       poster_blob: Optional[bytes] = None, is_video: bool = True) -> None:
    """Renders a multimedia poster frame with an elegant semi-transparent playback badge."""
    x0, y0, x1, y1 = box
    bw = int(round(x1 - x0))
    bh = int(round(y1 - y0))
    if bw <= 4 or bh <= 4:
        return

    # Draw poster image if present, else dark backdrop
    if poster_blob:
        try:
            p_img = Image.open(io.BytesIO(poster_blob)).convert("RGBA")
            p_img = p_img.resize((bw, bh), Image.Resampling.BILINEAR)
            img.paste(p_img, (int(x0), int(y0)), p_img)
        except Exception:
            ImageDraw.Draw(img).rectangle([x0, y0, x1, y1], fill=(40, 44, 52, 255))
    else:
        ImageDraw.Draw(img).rectangle([x0, y0, x1, y1], fill=(40, 44, 52, 255))

    # Play Icon Badge overlay
    draw = ImageDraw.Draw(img, "RGBA")
    cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
    r = min(bw, bh) * 0.15
    if r >= 8:
        draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=(0, 0, 0, 160), outline=(255, 255, 255, 220), width=2)
        # Play Triangle
        tri = [
            (cx - r * 0.3, cy - r * 0.45),
            (cx + r * 0.45, cy),
            (cx - r * 0.3, cy + r * 0.45),
        ]
        draw.polygon(tri, fill=(255, 255, 255, 240))


class MediaResolver:
    """Facade for vector and multimedia part decoding."""

    @staticmethod
    def load_svg(blob: bytes, w: int, h: int) -> Optional[Image.Image]:
        return rasterize_svg(blob, w, h)

    @staticmethod
    def load_metafile(blob: bytes, w: int, h: int) -> Optional[Image.Image]:
        return rasterize_emf_wmf(blob, w, h)

    @staticmethod
    def render_poster(img: Image.Image, box: Tuple[float, float, float, float],
                      poster_blob: Optional[bytes] = None, is_video: bool = True) -> None:
        paint_media_poster(img, box, poster_blob, is_video)
