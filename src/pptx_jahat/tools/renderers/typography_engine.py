"""
typography_engine.py
====================
Typography layout engine: FontResolver, multi-column text distribution,
WordArt warping, OMML math equation rendering, bullet formatting, RTL shaping.
"""

from __future__ import annotations

import io
import os
import math
import logging
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any, Callable
from PIL import Image, ImageDraw, ImageFont, ImageColor, ImageFilter

from pptx_jahat.tools.renderers.color_resolver import (
    RGBA, RGB, q, local_name, resolve_element_color, clamp, NS_A
)

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_BIDI = True
except Exception:
    HAS_BIDI = False

log = logging.getLogger("pptx_renderers.typography")


# ---------------------------------------------------------------------------
# RTL & Text Shaping
# ---------------------------------------------------------------------------
def is_rtl_text(text: str) -> bool:
    for ch in text:
        if ('\u0600' <= ch <= '\u06FF' or '\u0750' <= ch <= '\u077F'
                or '\uFB50' <= ch <= '\uFDFF' or '\uFE70' <= ch <= '\uFEFF'
                or '\u0590' <= ch <= '\u05FF'):
            return True
    return False


def shape_text_for_display(text: str) -> str:
    if not text or not HAS_BIDI or not is_rtl_text(text):
        return text
    try:
        return get_display(arabic_reshaper.reshape(text))
    except Exception:
        return text


# ---------------------------------------------------------------------------
# Font Resolver
# ---------------------------------------------------------------------------
class FontResolver:
    """Discovers system TTF/OTF fonts and resolves typeface matches with weight/style fallbacks."""

    _cached_system_fonts: Optional[Dict[str, Dict[str, str]]] = None

    def __init__(self):
        self.fonts: Dict[str, Dict[str, str]] = self._get_or_scan_fonts()
        self.major_font: Optional[str] = None
        self.minor_font: Optional[str] = None

    def set_theme_fonts(self, major: Optional[str], minor: Optional[str]) -> None:
        self.major_font = major
        self.minor_font = minor

    def theme_fonts(self) -> Dict[str, Optional[str]]:
        return {"major": self.major_font, "minor": self.minor_font}

    @classmethod
    def _get_or_scan_fonts(cls) -> Dict[str, Dict[str, str]]:
        if cls._cached_system_fonts is not None:
            return cls._cached_system_fonts

        scanned: Dict[str, Dict[str, str]] = {}
        dirs: List[Path] = []
        if os.name == "nt":
            windir = os.environ.get("WINDIR", "C:\\Windows")
            dirs.append(Path(windir) / "Fonts")
            local_app = os.environ.get("LOCALAPPDATA", "")
            if local_app:
                dirs.append(Path(local_app) / "Microsoft" / "Windows" / "Fonts")
        else:
            dirs.extend([
                Path("/usr/share/fonts"),
                Path("/usr/local/share/fonts"),
                Path("~/.fonts").expanduser(),
                Path("~/Library/Fonts").expanduser(),
                Path("/System/Library/Fonts"),
            ])
        # Project assets fonts
        dirs.append(Path(__file__).resolve().parent.parent.parent.parent / "assets" / "fonts")

        for d in dirs:
            if not d.exists():
                continue
            for ext in ("*.ttf", "*.otf", "*.TTF", "*.OTF"):
                for p in d.rglob(ext):
                    name_lower = p.stem.lower().replace("-", " ").replace("_", " ")
                    is_bold = "bold" in name_lower or "bd" in name_lower
                    is_italic = "italic" in name_lower or "oblique" in name_lower or "it" in name_lower
                    family = name_lower.replace("bold", "").replace("italic", "").replace("oblique", "").strip()
                    if family not in scanned:
                        scanned[family] = {}
                    style_key = f"{'bold_' if is_bold else ''}{'italic' if is_italic else 'regular'}"
                    scanned[family][style_key] = str(p)

        cls._cached_system_fonts = scanned
        return scanned

    def _match(self, family: str, bold: bool, italic: bool) -> Optional[str]:
        fam = family.lower().strip()
        fam_no_sp = fam.replace(" ", "")
        cand_keys = [k for k in self.fonts if k == fam or k.replace(" ", "") == fam_no_sp
                     or fam.startswith(k) or k.startswith(fam)]
        if not cand_keys:
            return None
        bucket = self.fonts[cand_keys[0]]
        desired = f"{'bold_' if bold else ''}{'italic' if italic else 'regular'}"
        if desired in bucket:
            return bucket[desired]
        if bold and "bold_regular" in bucket:
            return bucket["bold_regular"]
        if "regular" in bucket:
            return bucket["regular"]
        return next(iter(bucket.values()))

    def get_font(self, family: Optional[str], size_px: int,
                 bold: bool = False, italic: bool = False) -> Tuple[ImageFont.ImageFont, bool]:
        size_px = max(6, int(round(size_px)))
        path = None
        if family:
            path = self._match(family, bold, italic)
        if not path and self.minor_font:
            path = self._match(self.minor_font, bold, italic)
        if not path:
            for fallback in ("segoe ui", "calibri", "arial", "tahoma", "vazirmatn", "dejavu sans", "helvetica"):
                path = self._match(fallback, bold, italic)
                if path:
                    break
        if path:
            try:
                return ImageFont.truetype(path, size_px), False
            except Exception:
                pass
        try:
            return ImageFont.load_default(), bold
        except Exception:
            return ImageFont.load_default(), False


def text_width(font: ImageFont.ImageFont, s: str) -> float:
    try:
        bbox = font.getbbox(s)
        return float(bbox[2] - bbox[0])
    except Exception:
        try:
            return float(font.getlength(s))
        except Exception:
            return float(len(s) * 8)


# ---------------------------------------------------------------------------
# Run Style Model
# ---------------------------------------------------------------------------
class RunStyle:
    def __init__(self, font: ImageFont.ImageFont, size_px: int, color: RGBA,
                 underline: bool, faux_bold: bool, shadow: Optional[Dict[str, Any]] = None,
                 glow: Optional[Dict[str, Any]] = None, outline: Optional[Tuple[RGBA, int]] = None,
                 highlight: Optional[RGBA] = None, strike: bool = False,
                 warp: Optional[str] = None):
        self.font = font
        self.size_px = size_px
        self.color = color
        self.underline = underline
        self.faux_bold = faux_bold
        self.shadow = shadow
        self.glow = glow
        self.outline = outline
        self.highlight = highlight
        self.strike = strike
        self.warp = warp


def parse_run_style(rPr: Optional[ET.Element], colors: Dict[str, RGB],
                    fonts: FontResolver, scale_y: float) -> RunStyle:
    sz = 18.0
    bold = False
    italic = False
    underline = False
    strike = False
    color: RGBA = (0, 0, 0, 255)
    typeface = None
    shadow = None
    glow = None
    outline = None
    highlight = None

    if rPr is not None:
        sz_attr = rPr.attrib.get("sz")
        if sz_attr:
            try:
                sz = float(sz_attr) / 100.0
            except ValueError:
                pass
        bold = rPr.attrib.get("b") in ("1", "true")
        italic = rPr.attrib.get("i") in ("1", "true")
        underline = rPr.attrib.get("u") not in (None, "", "none")
        strike = rPr.attrib.get("strike") not in (None, "", "noStrike")

        # Color
        sf = rPr.find(q(NS_A, "solidFill"))
        if sf is not None:
            c = resolve_element_color(sf, colors)
            if c is not None:
                color = c
        elif rPr.find(q(NS_A, "noFill")) is not None:
            color = (0, 0, 0, 0)

        # Latin font
        latin = rPr.find(q(NS_A, "latin"))
        if latin is not None:
            typeface = latin.attrib.get("typeface")
        # Complex script font (Arabic/Farsi/Hebrew)
        cs = rPr.find(q(NS_A, "cs"))
        if cs is not None and not typeface:
            typeface = cs.attrib.get("typeface")

        # Effects
        sh = rPr.find(f".//{q(NS_A, 'outerShdw')}")
        if sh is not None:
            scol = resolve_element_color(sh, colors) or (0, 0, 0, 128)
            # dist is in EMU; convert directly: EMU * scale_y = pixels
            dist = float(sh.attrib.get("dist", 25400)) * scale_y
            ang = math.radians(float(sh.attrib.get("dir", 3240000)) / 60000.0)
            shadow = {"color": scol, "dx": dist * math.cos(ang), "dy": dist * math.sin(ang)}

        gl = rPr.find(f".//{q(NS_A, 'glow')}")
        if gl is not None:
            gcol = resolve_element_color(gl, colors) or (255, 255, 100, 150)
            glow = {"color": gcol, "rad": 2.0}

        ln = rPr.find(q(NS_A, "ln"))
        if ln is not None:
            lcol = resolve_element_color(ln, colors) or (0, 0, 0, 255)
            # ln w is in EMU; convert directly to pixels
            lw = max(1, int(round(float(ln.attrib.get("w", 12700)) * scale_y)))
            outline = (lcol, lw)

        hl = rPr.find(q(NS_A, "highlight"))
        if hl is not None:
            highlight = resolve_element_color(hl, colors)

    # Convert point size -> pixels. sz is in points; scale_y is EMU->px ratio;
    # 12700 EMU per point, so px = sz_pt * 12700 * scale_y.
    pt_to_px = 12700.0 * scale_y
    size_px = max(6, int(round(sz * pt_to_px)))
    font_obj, faux_bold = fonts.get_font(typeface, size_px, bold, italic)
    return RunStyle(font_obj, size_px, color, underline, faux_bold,
                    shadow, glow, outline, highlight, strike)


# ---------------------------------------------------------------------------
# Math Formula Rendering (OMML -> Matplotlib / PIL)
# ---------------------------------------------------------------------------
def omml_to_latex(elem: ET.Element) -> str:
    """Converts OpenXML Math elements into basic LaTeX math text representation."""
    texts = []
    for t in elem.iter():
        if local_name(t.tag) in ("t", "m:t"):
            if t.text:
                texts.append(t.text)
    return " ".join(texts)


def render_math_formula(elem: ET.Element, box: Tuple[float, float, float, float],
                        color: RGBA, scale_y: float) -> Optional[Image.Image]:
    """Renders an OMML / LaTeX math block into a crisp PIL image overlay."""
    latex_eq = omml_to_latex(elem)
    if not latex_eq.strip():
        return None
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        
        fig = plt.figure(figsize=(4, 1), dpi=150)
        fig.patch.set_alpha(0.0)
        c_norm = (color[0]/255.0, color[1]/255.0, color[2]/255.0)
        fig.text(0.5, 0.5, f"${latex_eq}$", fontsize=14, color=c_norm,
                 ha="center", va="center")
        buf = io.BytesIO()
        plt.savefig(buf, format="png", bbox_inches="tight", transparent=True, pad_inches=0.05)
        plt.close(fig)
        buf.seek(0)
        return Image.open(buf).convert("RGBA")
    except Exception as e:
        log.debug(f"Matplotlib math render failed, using text fallback: {e}")
        return None


# ---------------------------------------------------------------------------
# WordArt / Warped Text
# ---------------------------------------------------------------------------
def apply_wordart_warp(txt_img: Image.Image, warp_type: str) -> Image.Image:
    if not warp_type or warp_type in ("none", "textNoShape"):
        return txt_img
    w, h = txt_img.size
    if w <= 4 or h <= 4:
        return txt_img
    warped = Image.new("RGBA", (w, h + int(h * 0.4)), (0, 0, 0, 0))
    pixels_src = txt_img.load()
    pixels_dst = warped.load()
    
    # Arch Up / Wave deformation
    for x in range(w):
        t = x / (w or 1)
        if warp_type in ("textArchUp", "textCurveUp"):
            dy = -int(round(math.sin(t * math.pi) * (h * 0.35)))
        elif warp_type in ("textArchDown", "textCurveDown"):
            dy = int(round(math.sin(t * math.pi) * (h * 0.35)))
        elif warp_type in ("textWave1", "textWave2"):
            dy = int(round(math.sin(t * 2 * math.pi) * (h * 0.2)))
        else:
            dy = 0
        
        for y in range(h):
            target_y = y + int(h * 0.2) + dy
            if 0 <= target_y < warped.size[1]:
                pixels_dst[x, target_y] = pixels_src[x, y]
    return warped


# ---------------------------------------------------------------------------
# Multi-Column Text Flow & Text Body Rendering
# ---------------------------------------------------------------------------
def render_txbody(txbody: ET.Element, box: Tuple[float, float, float, float],
                  img: Image.Image, ctx: Dict[str, Any], rot_deg: float = 0.0) -> None:
    x0, y0, x1, y1 = box
    if x1 <= x0 or y1 <= y0:
        return

    bodyPr = txbody.find(q(NS_A, "bodyPr"))
    num_cols = 1
    col_gap = 10.0
    warp_type = None

    if bodyPr is not None:
        num_cols = max(1, min(4, int(bodyPr.attrib.get("numCol", "1"))))
        spc_col_emu = float(bodyPr.attrib.get("spcCol", 0))
        if spc_col_emu > 0:
            col_gap = (spc_col_emu / 12700.0) * ctx["scale_x"]
        prst_geom = bodyPr.find(q(NS_A, "prstTxWarp"))
        if prst_geom is not None:
            warp_type = prst_geom.attrib.get("prst")

    total_w = x1 - x0
    col_w = (total_w - col_gap * (num_cols - 1)) / num_cols if num_cols > 1 else total_w

    # Distribute paragraphs across columns
    paragraphs = txbody.findall(q(NS_A, "p"))
    if not paragraphs:
        return

    paras_per_col = math.ceil(len(paragraphs) / num_cols)
    draw = ImageDraw.Draw(img, "RGBA")

    for col_idx in range(num_cols):
        col_x0 = x0 + col_idx * (col_w + col_gap)
        col_x1 = col_x0 + col_w
        col_box = (col_x0, y0, col_x1, y1)
        col_paras = paragraphs[col_idx * paras_per_col: (col_idx + 1) * paras_per_col]
        _render_column_paragraphs(col_paras, col_box, img, draw, ctx, warp_type)


def _render_column_paragraphs(paragraphs: List[ET.Element], box: Tuple[float, float, float, float],
                             img: Image.Image, draw: ImageDraw.ImageDraw,
                             ctx: Dict[str, Any], warp_type: Optional[str]) -> None:
    x0, y0, x1, y1 = box
    curr_y = y0 + 4.0
    colors = ctx["palette"]
    fonts: FontResolver = ctx["fonts"]
    scale_y = ctx["scale_y"]

    for p in paragraphs:
        pPr = p.find(q(NS_A, "pPr"))
        algn = pPr.attrib.get("algn", "l") if pPr is not None else "l"
        is_rtl = pPr.attrib.get("rtl") == "1" if pPr is not None else False

        line_runs: List[Tuple[str, RunStyle]] = []
        for child in p:
            tag = local_name(child.tag)
            if tag == "r":
                t_elem = child.find(q(NS_A, "t"))
                text = t_elem.text if t_elem is not None and t_elem.text else ""
                if not text:
                    continue
                rPr = child.find(q(NS_A, "rPr"))
                style = parse_run_style(rPr, colors, fonts, scale_y)
                line_runs.append((text, style))
            elif tag in ("oMath", "m:oMath"):
                # Render embedded math formula
                m_img = render_math_formula(child, box, (0, 0, 0, 255), scale_y)
                if m_img:
                    img.alpha_composite(m_img, (int(x0), int(curr_y)))
                    curr_y += m_img.size[1] + 4.0

        if not line_runs:
            curr_y += 14.0 * scale_y
            continue

        # Render runs
        full_line_text = "".join(t for t, _ in line_runs)
        if HAS_BIDI and (is_rtl or is_rtl_text(full_line_text)):
            full_line_text = shape_text_for_display(full_line_text)
            algn = "r"

        r_x = x0 + 4.0
        line_h = 16.0 * scale_y

        for text_chunk, style in line_runs:
            if not text_chunk:
                continue
            txt_to_draw = shape_text_for_display(text_chunk) if (HAS_BIDI and is_rtl_text(text_chunk)) else text_chunk
            tw = text_width(style.font, txt_to_draw)

            if algn == "ctr":
                draw_x = (x0 + x1) / 2.0 - tw / 2.0
            elif algn == "r":
                draw_x = x1 - tw - 6.0
            else:
                draw_x = r_x

            # Draw highlight
            if style.highlight:
                draw.rectangle([draw_x, curr_y, draw_x + tw, curr_y + style.size_px], fill=style.highlight)

            # Draw text
            if style.color[3] > 0:
                draw.text((draw_x, curr_y), txt_to_draw, fill=style.color, font=style.font)
                if style.faux_bold:
                    draw.text((draw_x + 1, curr_y), txt_to_draw, fill=style.color, font=style.font)

            if style.underline:
                draw.line([(draw_x, curr_y + style.size_px + 1), (draw_x + tw, curr_y + style.size_px + 1)],
                          fill=style.color, width=1)
            if style.strike:
                draw.line([(draw_x, curr_y + style.size_px / 2.0), (draw_x + tw, curr_y + style.size_px / 2.0)],
                          fill=style.color, width=1)

            r_x += tw + 2.0
            line_h = max(line_h, style.size_px * 1.2)

        curr_y += line_h


class TypographyEngine:
    """Facade for high-level text, math, and WordArt typography rendering."""

    def __init__(self, fonts: FontResolver):
        self.fonts = fonts

    def render(self, txbody: ET.Element, box: Tuple[float, float, float, float],
               img: Image.Image, ctx: Dict[str, Any], rot_deg: float = 0.0) -> None:
        render_txbody(txbody, box, img, ctx, rot_deg)
