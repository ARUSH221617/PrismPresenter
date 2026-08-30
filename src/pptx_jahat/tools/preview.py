#!/usr/bin/env python3
"""
pptx_preview_engine.py  (v2.1 — bg & color fidelity fixes)
==========================================================
Pure-Python PPTX slide preview renderer (PIL only, no PowerPoint/LibreOffice).

New in this revision:
  * Slide backgrounds: p:bgPr solid/gradient/blip/noFill AND p:bgRef resolved
    against the theme's bgFillStyleLst (with phClr substitution). Background
    gradients and background images are actually painted now.
  * Colors: lumMod/lumOff computed in HSL (matches PowerPoint's
    Lighter/Darker theme variants), scheme aliases (bg1<->lt1, tx1<->dk1,
    bg2<->lt2, tx2<->dk2), sysClr defaults, prstClr/hslClr/scrgbClr support,
    robust hex parsing (3/6/8-digit, whitespace).
  * Shape style references (p:style): fillRef/lnRef/fontRef resolved from the
    theme's fillStyleLst/lnStyleLst — fixes shapes inheriting theme colors.
  * Theme is now resolved per slide master (multi-master decks), with
    fmtScheme kept for style/bgRef resolution.
"""

from __future__ import annotations

import io
import os
import math
import colorsys
import base64
import logging
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Optional, Tuple, Any, Dict

from pptx import Presentation
from PIL import Image, ImageDraw, ImageFont, ImageChops, ImageColor

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_BIDI = True
except Exception:
    HAS_BIDI = False

log = logging.getLogger("pptx_preview")

# ---------------------------------------------------------------------------
# Namespaces & small helpers
# ---------------------------------------------------------------------------
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

RGBA = Tuple[int, int, int, int]
RGB = Tuple[int, int, int]

try:
    RESAMPLE = Image.Resampling.LANCZOS
    BICUBIC = Image.Resampling.BICUBIC
except AttributeError:  # Pillow < 9.1
    RESAMPLE = Image.LANCZOS
    BICUBIC = Image.BICUBIC

_DEF_ACCENTS = [(79, 129, 189), (192, 80, 77), (155, 187, 89),
                (128, 100, 162), (75, 172, 198), (247, 150, 70)]

_DEFAULT_COLORS: Dict[str, RGB] = {
    "dk1": (0, 0, 0), "lt1": (255, 255, 255),
    "dk2": (80, 80, 70), "lt2": (238, 236, 225),
    "accent1": (79, 129, 189), "accent2": (192, 80, 77),
    "accent3": (155, 187, 89), "accent4": (128, 100, 162),
    "accent5": (75, 172, 198), "accent6": (247, 150, 70),
    "hlink": (0, 102, 204), "folHlink": (102, 102, 153),
    "bg1": (255, 255, 255), "bg2": (245, 245, 245),
    "tx1": (0, 0, 0), "tx2": (100, 100, 100),
}

# schemeClr value aliases used by PowerPoint
_SCHEME_ALIASES = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}

# sysClr fallbacks when lastClr is missing
_SYSCLR_DEFAULTS = {
    "window": (255, 255, 255), "windowtext": (0, 0, 0),
    "menu": (255, 255, 255), "menutext": (0, 0, 0),
    "buttonface": (240, 240, 240), "buttontext": (0, 0, 0),
    "buttonshadow": (160, 160, 160), "highlight": (0, 120, 215),
    "highlighttext": (255, 255, 255), "3dlight": (240, 240, 240),
    "infobackground": (255, 255, 225), "infotext": (0, 0, 0),
    "activetitle": (0, 90, 160), "graytext": (128, 128, 128),
}

_COLOR_TAGS = {"srgbClr", "schemeClr", "sysClr", "prstClr", "hslClr", "scrgbClr"}


def _q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def _local(tag: str) -> str:
    return tag.split("}")[-1]


def _int_attr(node: Optional[ET.Element], name: str, default: int) -> int:
    if node is None:
        return default
    v = node.attrib.get(name)
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def _str_attr(node: Optional[ET.Element], name: str, default: str) -> str:
    if node is None:
        return default
    return node.attrib.get(name, default)


def _rgba(c: Optional[Tuple[int, ...]]) -> Optional[RGBA]:
    if c is None:
        return None
    if len(c) == 3:
        return (c[0], c[1], c[2], 255)
    return (c[0], c[1], c[2], c[3])


def _clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def _hex_to_rgb(hv: Optional[str]) -> Optional[RGB]:
    """Robust hex color parsing: 3/6/8-digit, optional '#', any case."""
    if not hv:
        return None
    h = hv.strip().lstrip("#").strip()
    try:
        if len(h) == 3:
            return (int(h[0] * 2, 16), int(h[1] * 2, 16), int(h[2] * 2, 16))
        if len(h) >= 6:
            return (int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))
    except ValueError:
        return None
    return None


# ---------------------------------------------------------------------------
# Text shaping (Arabic / Hebrew / RTL)
# ---------------------------------------------------------------------------
def _is_rtl_text(text: str) -> bool:
    for ch in text:
        if ('\u0600' <= ch <= '\u06FF' or '\u0750' <= ch <= '\u077F'
                or '\uFB50' <= ch <= '\uFDFF' or '\uFE70' <= ch <= '\uFEFF'
                or '\u0590' <= ch <= '\u05FF'):
            return True
    return False


def _shape_text_for_display(text: str) -> str:
    if not text or not HAS_BIDI or not _is_rtl_text(text):
        return text
    try:
        return get_display(arabic_reshaper.reshape(text))
    except Exception:
        return text


# ---------------------------------------------------------------------------
# Theme model (colors + fonts + fmtScheme for style/bgRef resolution)
# ---------------------------------------------------------------------------
class Theme:
    """Holds a theme's color scheme, fonts and format scheme (fill/ln styles)."""

    def __init__(self, colors: Optional[Dict[str, RGB]] = None):
        self.colors: Dict[str, RGB] = dict(_DEFAULT_COLORS if colors is None else colors)
        self.major_font: Optional[str] = None
        self.minor_font: Optional[str] = None
        self.fmt_scheme: Optional[ET.Element] = None

    # -- parsing -------------------------------------------------------------
    def load_from_theme_part(self, blob: bytes) -> None:
        root = ET.fromstring(blob)
        clr = root.find(f".//{_q(NS_A, 'clrScheme')}")
        if clr is not None:
            for elem in clr:
                tag = _local(elem.tag)
                srgb = elem.find(_q(NS_A, "srgbClr"))
                sysc = elem.find(_q(NS_A, "sysClr"))
                rgb: Optional[RGB] = None
                if srgb is not None:
                    rgb = _hex_to_rgb(srgb.attrib.get("val"))
                if rgb is None and sysc is not None:
                    rgb = _hex_to_rgb(sysc.attrib.get("lastClr"))
                    if rgb is None:
                        rgb = _SYSCLR_DEFAULTS.get(
                            (sysc.attrib.get("val") or "").lower())
                if rgb:
                    self.colors[tag] = rgb
                    if tag == "lt1":
                        self.colors["bg1"] = rgb
                    elif tag == "lt2":
                        self.colors["bg2"] = rgb
                    elif tag == "dk1":
                        self.colors["tx1"] = rgb
                    elif tag == "dk2":
                        self.colors["tx2"] = rgb
        fs = root.find(f".//{_q(NS_A, 'fontScheme')}")
        if fs is not None:
            maj = fs.find(_q(NS_A, "majorFont"))
            mnr = fs.find(_q(NS_A, "minorFont"))
            if maj is not None:
                lat = maj.find(_q(NS_A, "latin"))
                self.major_font = lat.attrib.get("typeface") if lat is not None else None
            if mnr is not None:
                lat = mnr.find(_q(NS_A, "latin"))
                self.minor_font = lat.attrib.get("typeface") if lat is not None else None
        self.fmt_scheme = root.find(f".//{_q(NS_A, 'fmtScheme')}")

    @classmethod
    def from_presentation(cls, prs: Presentation) -> "Theme":
        theme = cls()
        try:
            for rel in prs.part.rels.values():
                if "theme" in rel.reltype:
                    theme.load_from_theme_part(rel.target_part.blob)
                    break
        except Exception:
            log.debug("presentation theme extraction failed", exc_info=True)
        return theme

    @property
    def theme_fonts(self) -> Dict[str, Optional[str]]:
        return {"major": self.major_font, "minor": self.minor_font}


def _extract_theme_color_palette(prs: Presentation) -> Dict[str, RGB]:
    """Backward-compatible helper."""
    return Theme.from_presentation(prs).colors


def _theme_for_slide(slide: Any, cache: Dict[int, Theme], fallback: Theme) -> Theme:
    """Resolves the theme belonging to a slide's master (multi-master decks)."""
    try:
        master = slide.slide_layout.slide_master
        key = id(master)
        t = cache.get(key)
        if t is None:
            t = Theme()
            loaded = False
            try:
                for rel in master.part.rels.values():
                    if "theme" in rel.reltype:
                        t.load_from_theme_part(rel.target_part.blob)
                        loaded = True
                        break
            except Exception:
                pass
            if not loaded:
                t.colors.update(fallback.colors)
                t.major_font = fallback.major_font
                t.minor_font = fallback.minor_font
                t.fmt_scheme = fallback.fmt_scheme
            cache[key] = t
        return t
    except Exception:
        return fallback


# ---------------------------------------------------------------------------
# Color resolution (DrawingML) — now HSL-correct
# ---------------------------------------------------------------------------
def _scheme_rgb(val: Optional[str], colors: Dict[str, RGB],
                ph_rgb: Optional[RGB]) -> Optional[RGB]:
    v = (val or "").strip()
    if not v:
        return None
    if v == "phClr":
        return ph_rgb
    probe = v[0].lower() + v[1:] if v else v
    for cand in (v, probe, _SCHEME_ALIASES.get(v), _SCHEME_ALIASES.get(probe)):
        if cand and cand in colors:
            return colors[cand]
    return None


def _parse_color_elem(elem: ET.Element, colors: Dict[str, RGB],
                      ph_rgb: Optional[RGB]) -> Tuple[Optional[RGB], ET.Element]:
    tag = _local(elem.tag)
    if tag == "srgbClr":
        return _hex_to_rgb(elem.attrib.get("val")), elem
    if tag == "sysClr":
        rgb = _hex_to_rgb(elem.attrib.get("lastClr"))
        if rgb is None:
            rgb = _SYSCLR_DEFAULTS.get((elem.attrib.get("val") or "").lower(), (0, 0, 0))
        return rgb, elem
    if tag == "schemeClr":
        return _scheme_rgb(elem.attrib.get("val"), colors, ph_rgb), elem
    if tag == "prstClr":
        try:
            rgb = ImageColor.getrgb(elem.attrib.get("val", "black"))
            return (rgb[0], rgb[1], rgb[2]), elem
        except Exception:
            return None, elem
    if tag == "hslClr":
        try:
            hue = _int_attr(elem, "hue", 0) / 60000.0
            sat_n = elem.find(_q(NS_A, "sat"))
            lum_n = elem.find(_q(NS_A, "lum"))
            sat = _int_attr(sat_n, "val", 0) / 100000.0 if sat_n is not None else 0.0
            lum = _int_attr(lum_n, "val", 0) / 100000.0 if lum_n is not None else 0.5
            r, g, b = colorsys.hls_to_rgb((hue % 360) / 360.0, _clamp(lum, 0, 1),
                                          _clamp(sat, 0, 1))
            return (int(r * 255), int(g * 255), int(b * 255)), elem
        except Exception:
            return None, elem
    if tag == "scrgbClr":
        try:
            parts = {}
            for ch in ("r", "g", "b"):
                n = elem.find(_q(NS_A, ch))
                parts[ch] = _int_attr(n, "val", 0) / 100000.0 if n is not None else 0.0
            return (int(_clamp(parts["r"], 0, 1) * 255),
                    int(_clamp(parts["g"], 0, 1) * 255),
                    int(_clamp(parts["b"], 0, 1) * 255)), elem
        except Exception:
            return None, elem
    return None, elem


def _resolve_element_color(elem: Optional[ET.Element],
                           colors: Dict[str, RGB],
                           ph_rgb: Optional[RGB] = None) -> Optional[RGBA]:
    """
    Parses a color spec (srgbClr/schemeClr/sysClr/prstClr/hslClr/scrgbClr,
    possibly wrapped in solidFill etc.) and applies transforms.
    lumMod/lumOff are applied in HSL space (PowerPoint semantics).
    Returns (R, G, B, Alpha 0-255).
    """
    if elem is None:
        return None

    base: Optional[RGB] = None
    color_node: Optional[ET.Element] = None

    tag = _local(elem.tag)
    if tag in _COLOR_TAGS:
        base, color_node = _parse_color_elem(elem, colors, ph_rgb)
    if base is None:  # direct child color
        for child in elem:
            if _local(child.tag) in _COLOR_TAGS:
                base, color_node = _parse_color_elem(child, colors, ph_rgb)
                break
    if base is None:  # deeper search
        for node in elem.iter():
            if _local(node.tag) in _COLOR_TAGS:
                base, color_node = _parse_color_elem(node, colors, ph_rgb)
                break
    if not base:
        return None

    r, g, b = base
    alpha = 255.0
    lum_mod = lum_off = shade = tint = None

    if color_node is not None:
        for mod in color_node:
            mt = _local(mod.tag)
            try:
                v = int(mod.attrib.get("val", "0")) / 100000.0
            except (TypeError, ValueError):
                continue
            if mt == "lumMod":
                lum_mod = v
            elif mt == "lumOff":
                lum_off = v
            elif mt == "shade":
                shade = v
            elif mt == "tint":
                tint = v
            elif mt == "alpha":
                alpha = _clamp(v, 0.0, 1.0)
            elif mt == "alphaMod":
                alpha = _clamp(alpha * v, 0.0, 1.0)
            elif mt == "alphaOff":
                alpha = _clamp(alpha + v, 0.0, 1.0)
            # satMod / hueOff / hueMod / gray / comp / inv ignored (rare)

    # HSL-luminance transforms — this is how PowerPoint builds
    # "Lighter 40%"/"Darker 25%" theme color variants.
    if lum_mod is not None or lum_off is not None:
        h, l, s = colorsys.rgb_to_hls(r / 255.0, g / 255.0, b / 255.0)
        if lum_mod is not None:
            l *= lum_mod
        if lum_off is not None:
            l += lum_off
        l = _clamp(l, 0.0, 1.0)
        r, g, b = colorsys.hls_to_rgb(h, l, s)

    if shade is not None:
        r, g, b = r * shade, g * shade, b * shade
    if tint is not None:
        r = r + (255 - r) * tint
        g = g + (255 - g) * tint
        b = b + (255 - b) * tint

    return (_clamp(int(round(r)), 0, 255), _clamp(int(round(g)), 0, 255),
            _clamp(int(round(b)), 0, 255), int(round(alpha * 255)))


def _parse_gradient(gradFill: ET.Element, colors: Dict[str, RGB],
                    ph_rgb: Optional[RGB] = None
                    ) -> Tuple[List[Tuple[float, RGBA]], float]:
    stops: List[Tuple[float, RGBA]] = []
    gsLst = gradFill.find(_q(NS_A, "gsLst"))
    if gsLst is not None:
        for gs in gsLst.findall(_q(NS_A, "gs")):
            try:
                pos = int(gs.attrib.get("pos", "0")) / 100000.0
            except ValueError:
                pos = 0.0
            rgba = _resolve_element_color(gs, colors, ph_rgb)
            if rgba:
                stops.append((_clamp(pos, 0.0, 1.0), rgba))
    stops.sort(key=lambda s: s[0])

    angle = 90.0
    lin = gradFill.find(_q(NS_A, "lin"))
    if lin is not None:
        angle = _int_attr(lin, "ang", 5400000) / 60000.0
    if not stops:
        stops = [(0.0, (230, 230, 230, 255)), (1.0, (255, 255, 255, 255))]
    return stops, angle


def _interp_stops(stops: List[Tuple[float, RGBA]], t: float) -> RGBA:
    if t <= stops[0][0]:
        return stops[0][1]
    if t >= stops[-1][0]:
        return stops[-1][1]
    for i in range(len(stops) - 1):
        p0, c0 = stops[i]
        p1, c1 = stops[i + 1]
        if p0 <= t <= p1:
            f = (t - p0) / (p1 - p0) if p1 > p0 else 0.0
            return (int(c0[0] + (c1[0] - c0[0]) * f),
                    int(c0[1] + (c1[1] - c0[1]) * f),
                    int(c0[2] + (c1[2] - c0[2]) * f),
                    int(c0[3] + (c1[3] - c0[3]) * f))
    return stops[-1][1]


def _resolve_theme_fill(fmt_scheme: Optional[ET.Element], lst_tag: str,
                        idx: int, ph_rgb: Optional[RGB],
                        colors: Dict[str, RGB]) -> Optional[tuple]:
    """Resolves a fillStyleLst/lnStyleLst/bgFillStyleLst entry (1-based idx)."""
    if fmt_scheme is None or idx < 1:
        return None
    lst = fmt_scheme.find(_q(NS_A, lst_tag))
    if lst is None:
        return None
    fills = [c for c in lst
             if _local(c.tag) in ("solidFill", "gradFill", "blipFill")]
    if idx - 1 >= len(fills):
        return None
    f = fills[idx - 1]
    tag = _local(f.tag)
    if tag == "solidFill":
        col = _resolve_element_color(f, colors, ph_rgb)
        return ("solid", col) if col else None
    if tag == "gradFill":
        stops, ang = _parse_gradient(f, colors, ph_rgb)
        return ("grad", stops, ang)
    return None


# ---------------------------------------------------------------------------
# Font resolution
# ---------------------------------------------------------------------------
class FontResolver:
    FALLBACKS = ["Calibri", "Segoe UI", "Arial", "Tahoma", "Helvetica",
                 "DejaVu Sans", "Liberation Sans", "Verdana"]

    def __init__(self):
        self._cache: Dict[tuple, Tuple[ImageFont.FreeTypeFont, bool]] = {}
        self._index: Dict[str, str] = {}
        self._theme_major: Optional[str] = None
        self._theme_minor: Optional[str] = None
        self._scan_dirs()

    def set_theme_fonts(self, major: Optional[str], minor: Optional[str]) -> None:
        self._theme_major = major
        self._theme_minor = minor

    @property
    def theme_fonts(self) -> Dict[str, Optional[str]]:
        return {"major": self._theme_major, "minor": self._theme_minor}

    def _scan_dirs(self) -> None:
        dirs: List[str] = []
        if os.name == "nt":
            windir = os.environ.get("WINDIR", r"C:\Windows")
            dirs.append(os.path.join(windir, "Fonts"))
            local = os.environ.get("LOCALAPPDATA")
            if local:
                dirs.append(os.path.join(local, "Microsoft", "Windows", "Fonts"))
        dirs += ["/usr/share/fonts", "/usr/local/share/fonts",
                 str(Path.home() / ".fonts"), "/Library/Fonts",
                 str(Path.home() / "Library" / "Fonts")]
        try:
            import matplotlib
            dirs.append(os.path.join(matplotlib.get_data_path(), "fonts", "ttf"))
        except Exception:
            pass
        for d in dirs:
            if not os.path.isdir(d):
                continue
            for root, _dirs, files in os.walk(d):
                for f in files:
                    if f.lower().endswith((".ttf", ".otf")):
                        self._index.setdefault(Path(f).stem.lower(),
                                               os.path.join(root, f))

    def _match(self, family: str, bold: bool, italic: bool) -> Optional[str]:
        fam = family.lower().replace(" ", "").replace("-", "")
        suffixes: List[str] = []
        if bold and italic:
            suffixes = ["bi", "bolditalic", "boldital", "bdit", "z"]
        elif bold:
            suffixes = ["bd", "bold", "b"]
        elif italic:
            suffixes = ["it", "italic", "i", "obi", "oblique"]
        suffixes.append("")
        for suf in suffixes:
            p = self._index.get(fam + suf)
            if p:
                return p
        for stem, path in self._index.items():
            if fam and fam in stem:
                if bold and not ("bold" in stem or stem.endswith("bd")):
                    continue
                if italic and not ("italic" in stem or stem.endswith("it")
                                   or "oblique" in stem):
                    continue
                return path
        return None

    def get_font(self, family: Optional[str], size_px: int,
                 bold: bool = False, italic: bool = False
                 ) -> Tuple[ImageFont.FreeTypeFont, bool]:
        family = family or self._theme_minor or "Calibri"
        size_px = max(5, int(size_px))
        key = (family, size_px, bold, italic)
        if key in self._cache:
            return self._cache[key]

        candidates = [family] + [f for f in self.FALLBACKS
                                 if f.lower() != family.lower()]
        for fam in candidates:
            p = self._match(fam, bold, italic)
            if p:
                try:
                    font = ImageFont.truetype(p, size_px)
                    self._cache[key] = (font, False)
                    return font, False
                except Exception:
                    continue
        for fam in candidates:
            try:
                font = ImageFont.truetype(f"{fam}.ttf", size_px)
                self._cache[key] = (font, bold)
                return font, bold
            except Exception:
                continue
        for fam in candidates:
            p = self._match(fam, False, False)
            if p:
                try:
                    font = ImageFont.truetype(p, size_px)
                    self._cache[key] = (font, bold)
                    return font, bold
                except Exception:
                    continue
        font = ImageFont.load_default()
        self._cache[key] = (font, False)
        return font, False


def _text_width(font: ImageFont.ImageFont, s: str) -> float:
    try:
        return font.getlength(s)
    except AttributeError:
        bbox = font.getbbox(s)
        return float(bbox[2] - bbox[0])


# ---------------------------------------------------------------------------
# Geometry: bezier, preset shapes, custom geometry  (unchanged logic)
# ---------------------------------------------------------------------------
def _draw_cubic_bezier(p0, p1, p2, p3, steps=16) -> List[Tuple[float, float]]:
    points = []
    for i in range(steps + 1):
        t = i / steps
        u = 1.0 - t
        x = (u ** 3) * p0[0] + 3 * (u ** 2) * t * p1[0] + 3 * u * (t ** 2) * p2[0] + (t ** 3) * p3[0]
        y = (u ** 3) * p0[1] + 3 * (u ** 2) * t * p1[1] + 3 * u * (t ** 2) * p2[1] + (t ** 3) * p3[1]
        points.append((x, y))
    return points


def _regular_pts(cx, cy, rx, ry, n, start_deg=-90.0) -> List[Tuple[float, float]]:
    pts = []
    for i in range(n):
        a = math.radians(start_deg + i * 360.0 / n)
        pts.append((cx + rx * math.cos(a), cy + ry * math.sin(a)))
    return pts


def _star_pts(cx, cy, rx, ry, n, inner=0.42) -> List[Tuple[float, float]]:
    pts = []
    for i in range(2 * n):
        r = 1.0 if i % 2 == 0 else inner
        a = math.radians(-90.0 + i * 180.0 / n)
        pts.append((cx + rx * r * math.cos(a), cy + ry * r * math.sin(a)))
    return pts


def _parse_adj(pg: Optional[ET.Element]) -> Dict[str, int]:
    adj: Dict[str, int] = {}
    if pg is None:
        return adj
    avLst = pg.find(_q(NS_A, "avLst"))
    if avLst is None:
        return adj
    for gd in avLst.findall(_q(NS_A, "gd")):
        fmla = gd.attrib.get("fmla", "val 0")
        try:
            adj[gd.attrib.get("name", "adj")] = int(fmla.split()[-1])
        except (ValueError, IndexError):
            pass
    return adj


def _arrow_ops(prst: str, box, adj) -> List[Tuple[str, list]]:
    x0, y0, x1, y1 = box
    w, h = x1 - x0, y1 - y0
    cy = (y0 + y1) / 2.0
    cx = (x0 + x1) / 2.0
    a1 = _clamp(adj.get("adj1", 50000) / 100000.0, 0.05, 0.95)
    a2 = _clamp(adj.get("adj2", 50000) / 100000.0, 0.05, 0.95)
    shaft = min(w, h) * a1
    head = min(w, h) * a2

    pts = [(x0, cy - shaft / 2), (x1 - head, cy - shaft / 2), (x1 - head, y0),
           (x1, cy), (x1 - head, y1), (x1 - head, cy + shaft / 2),
           (x0, cy + shaft / 2)]
    if prst == "rightArrow":
        mapped = pts
    elif prst == "leftArrow":
        mapped = [(x1 - (p[0] - x0), p[1]) for p in pts]
    elif prst == "upArrow":
        mapped = [(x0 + (p[1] - y0), y0 + (p[0] - x0)) for p in pts]
    elif prst == "downArrow":
        mapped = [(x1 - (p[1] - y0), y0 + (p[0] - x0)) for p in pts]
    else:
        mapped = pts

    if prst == "leftRightArrow":
        mapped = [(x0, cy), (x0 + head, y0), (x0 + head, cy - shaft / 2),
                  (x1 - head, cy - shaft / 2), (x1 - head, y0), (x1, cy),
                  (x1 - head, y1), (x1 - head, cy + shaft / 2),
                  (x0 + head, cy + shaft / 2), (x0 + head, y1)]
    elif prst == "upDownArrow":
        mapped = [(cx, y0), (cx - shaft / 2, y0 + head), (cx + shaft / 2, y0 + head),
                  (cx + shaft / 2, y1 - head), (x1, y1 - head), (cx, y1),
                  (x0, y1 - head), (cx - shaft / 2, y1 - head),
                  (cx - shaft / 2, y0 + head)]
        mapped = [(p[1], p[0]) for p in
                  [(cy, x0), (cy - shaft / 2, x0 + head), (cy + shaft / 2, x0 + head),
                   (cy + shaft / 2, x1 - head), (cy + head, x1), (cy, x1),
                   (cy - head, x1), (cy - head, x1 - head), (cy - shaft / 2, x0 + head)]]
    return [("poly", mapped)]


def _prst_ops(prst: Optional[str], box, adj: Optional[Dict[str, int]] = None
              ) -> List[Tuple[str, Any]]:
    adj = adj or {}
    x0, y0, x1, y1 = box
    w = max(1.0, x1 - x0)
    h = max(1.0, y1 - y0)
    cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
    prst = (prst or "rect").lower()

    def a(name, default):
        v = adj.get(name)
        return (v if v is not None else default) / 100000.0

    P = lambda pts: ("poly", pts)
    RECT = [P([(x0, y0), (x1, y0), (x1, y1), (x0, y1)])]

    if prst in ("rect", "flowChartProcess", "flowChartCollate", "flowChartOnlineStorage",
                "flowChartMagneticDisk", "flowChartMagneticDrum", "flowChartMagneticTape",
                "flowChartManualOperation", "flowChartPunchedTape", "textPlainText"):
        return RECT
    if prst in ("roundRect", "round1Rect", "round2DiagRect", "round2SameRect",
                "flowChartAlternateProcess", "flowChartDelay", "can"):
        r = min(w, h) * min(0.5, a("adj", 16667))
        return [("round", (x0, y0, x1, y1, r))]
    if prst in ("ellipse", "flowChartTerminator", "flowChartConnector",
                "flowChartOffpageConnector", "teardrop", "pie", "chord",
                "donut", "blockArc", "arc", "moon"):
        return [("ellipse", (x0, y0, x1, y1))]
    if prst == "triangle":
        return [P([(cx, y0), (x1, y1), (x0, y1)])]
    if prst == "rtTriangle":
        return [P([(x0, y0), (x0, y1), (x1, y1)])]
    if prst in ("diamond", "flowChartDecision"):
        return [P([(cx, y0), (x1, cy), (cx, y1), (x0, cy)])]
    if prst in ("parallelogram", "flowChartData"):
        off = min(w * a("adj", 25000), w * 0.5)
        return [P([(x0 + off, y0), (x1, y0), (x1 - off, y1), (x0, y1)])]
    if prst == "trapezoid":
        off = min(w * a("adj", 25000), w * 0.5)
        return [P([(x0 + off, y0), (x1 - off, y0), (x1, y1), (x0, y1)])]
    if prst in ("pentagon", "hexagon", "heptagon", "octagon", "decagon", "dodecagon"):
        n = {"pentagon": 5, "hexagon": 6, "heptagon": 7, "octagon": 8,
             "decagon": 10, "dodecagon": 12}.get(prst, 5)
        return [P(_regular_pts(cx, cy, w / 2, h / 2, n))]
    if prst.startswith("star"):
        try:
            n = int(prst[4:])
        except ValueError:
            n = 5
        n = max(3, min(24, n))
        inner = {5: 0.42, 4: 0.38, 6: 0.55, 7: 0.55, 8: 0.6, 10: 0.65, 12: 0.7}
        return [P(_star_pts(cx, cy, w / 2, h / 2, n, inner.get(n, 0.55)))]
    if prst in ("rightArrow", "leftArrow", "upArrow", "downArrow",
                "leftRightArrow", "upDownArrow"):
        return _arrow_ops(prst, box, adj)
    if prst in ("bentArrow", "curvedRightArrow", "curvedLeftArrow",
                "curvedUpArrow", "curvedDownArrow"):
        return [P([(x0, y1), (x0, cy), (x1 - min(w, h) * 0.3, cy),
                   (x1 - min(w, h) * 0.3, y0), (x1, cy),
                   (x1 - min(w, h) * 0.3, y1)])]
    if prst == "homePlate":
        n = min(w, h) * 0.25
        return [P([(x0, y0), (x1 - n, y0), (x1, cy), (x1 - n, y1), (x0, y1)])]
    if prst == "chevron":
        n = min(w, h) * 0.25
        return [P([(x0, y0), (x1 - n, y0), (x1, cy), (x1 - n, y1), (x0, y1), (x0 + n, cy)])]
    if prst in ("plus", "mathPlus", "cross"):
        t = min(w, h) * 0.33
        return [P([(cx - t, y0), (cx + t, y0), (cx + t, cy - t), (x1, cy - t),
                   (x1, cy + t), (cx + t, cy + t), (cx + t, y1), (cx - t, y1),
                   (cx - t, cy + t), (x0, cy + t), (x0, cy - t), (cx - t, cy - t)])]
    if prst == "cube":
        d = min(w, h) * 0.2
        return [P([(x0, y0 + d), (x1 - d, y0 + d), (x1 - d, y1), (x0, y1)]),
                P([(x0, y0 + d), (x0 + d, y0), (x1, y0), (x1 - d, y0 + d)])]
    if prst in ("line", "straightConnector1"):
        return [P([(x0, y0), (x1, y1)])]
    return RECT


def _connector_ops(prst: Optional[str], box, flip_h: bool, flip_v: bool
                   ) -> List[Tuple[str, list]]:
    x0, y0, x1, y1 = box
    p1, p2 = (x0, y0), (x1, y1)
    if flip_h and flip_v:
        p1, p2 = (x1, y1), (x0, y0)
    elif flip_h:
        p1, p2 = (x1, y0), (x0, y1)
    elif flip_v:
        p1, p2 = (x0, y1), (x1, y0)
    prst = (prst or "").lower()
    if prst.startswith("bentConnector3"):
        mid = (p1[0] + p2[0]) / 2.0
        return [("poly", [p1, (mid, p1[1]), (mid, p2[1]), p2])]
    if prst.startswith("bentConnector"):
        mid = (p1[1] + p2[1]) / 2.0
        return [("poly", [p1, (p1[0], mid), (p2[0], mid), p2])]
    if prst.startswith("curvedConnector3"):
        mid = (p1[0] + p2[0]) / 2.0
        return [("poly", _draw_cubic_bezier(p1, (mid, p1[1]), (mid, p2[1]), p2, 24))]
    if prst.startswith("curvedConnector"):
        mid = (p1[1] + p2[1]) / 2.0
        return [("poly", _draw_cubic_bezier(p1, (p1[0], mid), (p2[0], mid), p2, 24))]
    return [("poly", [p1, p2])]


def _render_custom_geom(img: Image.Image, draw: ImageDraw.ImageDraw,
                        custGeom: ET.Element, box,
                        fill_rgba: Optional[RGBA],
                        line_rgba: Optional[RGBA], line_w_px: int) -> None:
    x0, y0, x1, y1 = box
    w, h = x1 - x0, y1 - y0
    pathLst = custGeom.find(_q(NS_A, "pathLst"))
    if pathLst is None or w <= 0 or h <= 0:
        return

    for path in pathLst.findall(_q(NS_A, "path")):
        pw = float(path.attrib.get("w", w)) or w
        ph = float(path.attrib.get("h", h)) or h
        sx, sy = w / pw, h / ph
        fill_ok = path.attrib.get("fill", "norm") != "none"
        stroke_ok = path.attrib.get("stroke", "norm") != "none"

        def M(pt):
            return (x0 + float(pt.attrib.get("x", 0)) * sx,
                    y0 + float(pt.attrib.get("y", 0)) * sy)

        polys: List[list] = []
        cur: List[Tuple[float, float]] = []
        cur_pt = (x0, y0)

        def flush():
            if len(cur) >= 2:
                polys.append(list(cur))
            cur.clear()

        for cmd in path:
            tag = _local(cmd.tag)
            if tag == "moveTo":
                flush()
                pt = cmd.find(_q(NS_A, "pt"))
                if pt is not None:
                    cur_pt = M(pt)
                    cur.append(cur_pt)
            elif tag == "lnTo":
                pt = cmd.find(_q(NS_A, "pt"))
                if pt is not None:
                    cur_pt = M(pt)
                    cur.append(cur_pt)
            elif tag == "cubicBezTo":
                pts = cmd.findall(_q(NS_A, "pt"))
                if len(pts) == 3:
                    c1, c2, p3 = M(pts[0]), M(pts[1]), M(pts[2])
                    cur.extend(_draw_cubic_bezier(cur_pt, c1, c2, p3)[1:])
                    cur_pt = p3
            elif tag == "quadBezTo":
                pts = cmd.findall(_q(NS_A, "pt"))
                if len(pts) == 2:
                    c, p3 = M(pts[0]), M(pts[1])
                    c1 = (cur_pt[0] + 2 / 3 * (c[0] - cur_pt[0]),
                          cur_pt[1] + 2 / 3 * (c[1] - cur_pt[1]))
                    c2 = (p3[0] + 2 / 3 * (c[0] - p3[0]),
                          p3[1] + 2 / 3 * (c[1] - p3[1]))
                    cur.extend(_draw_cubic_bezier(cur_pt, c1, c2, p3)[1:])
                    cur_pt = p3
            elif tag == "close":
                if cur:
                    cur.append(cur[0])
                    flush()
        flush()

        for pts in polys:
            if fill_ok and fill_rgba and len(pts) >= 3:
                draw.polygon(pts, fill=fill_rgba)
            if stroke_ok and line_rgba and len(pts) >= 2:
                closed = fill_ok and len(pts) >= 3
                draw.line(pts + ([pts[0]] if closed else []),
                          fill=line_rgba, width=max(1, line_w_px), joint="curve")


# ---------------------------------------------------------------------------
# Paint helpers
# ---------------------------------------------------------------------------
def _fill_ops(draw: ImageDraw.ImageDraw, ops, fill: RGBA) -> None:
    for kind, data in ops:
        if kind == "ellipse":
            draw.ellipse(data, fill=fill)
        elif kind == "round":
            draw.rounded_rectangle(data[:4], radius=data[4], fill=fill)
        elif len(data) >= 3:
            draw.polygon(data, fill=fill)


def _stroke_ops(draw: ImageDraw.ImageDraw, ops, color: RGBA, lw: int,
                dashed: bool) -> None:
    for kind, data in ops:
        if kind == "ellipse":
            draw.ellipse(data, outline=color, width=max(1, lw))
            continue
        if kind == "round":
            draw.rounded_rectangle(data[:4], radius=data[4],
                                   outline=color, width=max(1, lw))
            continue
        pts = list(data)
        if len(pts) < 2:
            continue
        if len(pts) >= 3:
            pts = pts + [pts[0]]
        if dashed:
            _draw_dashed_line(draw, pts, color, max(1, lw))
        else:
            draw.line(pts, fill=color, width=max(1, lw), joint="curve")


def _draw_dashed_line(draw: ImageDraw.ImageDraw, pts, color: RGBA, lw: int,
                      dash: float = 8.0, gap: float = 5.0) -> None:
    for i in range(len(pts) - 1):
        ax, ay = pts[i]
        bx, by = pts[i + 1]
        seg = math.hypot(bx - ax, by - ay)
        if seg < 1:
            continue
        ux, uy = (bx - ax) / seg, (by - ay) / seg
        d = 0.0
        while d < seg:
            e = min(d + dash, seg)
            draw.line([(ax + ux * d, ay + uy * d), (ax + ux * e, ay + uy * e)],
                      fill=color, width=lw)
            d = e + gap


def _paint_gradient(img: Image.Image, box, stops, angle_deg: float,
                    silhouette) -> None:
    x0, y0, x1, y1 = [float(v) for v in box]
    w, h = x1 - x0, y1 - y0
    if w <= 0 or h <= 0 or not stops:
        return
    rad = math.radians(angle_deg)
    dx, dy = math.cos(rad), math.sin(rad)
    cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
    length = abs(w * dx) + abs(h * dy)
    if length <= 0:
        length = 1.0
    steps = max(2, min(512, int(length)))
    diag = math.hypot(w, h) * 0.75 + 2
    ex, ey = -dy, dx

    layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
    ldraw = ImageDraw.Draw(layer)
    for i in range(steps + 1):
        t = i / steps
        off = (t - 0.5) * length
        px, py = cx + dx * off, cy + dy * off
        ldraw.line([(px - ex * diag, py - ey * diag), (px + ex * diag, py + ey * diag)],
                   fill=_interp_stops(stops, t),
                   width=max(2, int(length / steps) + 2))

    mask = Image.new("L", img.size, 0)
    mdraw = ImageDraw.Draw(mask)
    silhouette(mdraw)
    combined = ImageChops.multiply(mask, layer.getchannel("A"))
    img.paste(layer, (0, 0), combined)


def _paint_blip(img: Image.Image, box, blob: bytes, silhouette) -> None:
    x0, y0, x1, y1 = [int(round(v)) for v in box]
    w, h = max(1, x1 - x0), max(1, y1 - y0)
    try:
        pic = Image.open(io.BytesIO(blob)).convert("RGBA").resize((w, h), RESAMPLE)
    except Exception:
        return
    mask = Image.new("L", img.size, 0)
    mdraw = ImageDraw.Draw(mask)
    silhouette(mdraw)
    combined = ImageChops.multiply(mask, pic.getchannel("A"))
    img.paste(pic, (0, 0), combined)


def _make_silhouette(ops):
    def sil(mdraw: ImageDraw.ImageDraw) -> None:
        for kind, data in ops:
            if kind == "ellipse":
                mdraw.ellipse(data, fill=255)
            elif kind == "round":
                mdraw.rounded_rectangle(data[:4], radius=data[4], fill=255)
            elif len(data) >= 3:
                mdraw.polygon(data, fill=255)
    return sil


# ---------------------------------------------------------------------------
# XML structure helpers
# ---------------------------------------------------------------------------
def _find_spPr(elem: ET.Element) -> Optional[ET.Element]:
    for tag in ("spPr", "grpSpPr"):
        for ns in (NS_P, NS_A):
            node = elem.find(_q(ns, tag))
            if node is not None:
                return node
    return None


def _find_xfrm(elem: ET.Element) -> Optional[ET.Element]:
    spPr = _find_spPr(elem)
    if spPr is not None:
        xf = spPr.find(_q(NS_A, "xfrm"))
        if xf is not None:
            return xf
    for ns in (NS_P, NS_A):
        xf = elem.find(_q(ns, "xfrm"))
        if xf is not None:
            return xf
    return None


def _is_hidden(elem: ET.Element) -> bool:
    for node in elem.iter():
        if _local(node.tag) == "cNvPr" and node.attrib.get("hidden") in ("1", "true"):
            return True
    return False


def _find_style(elem: ET.Element) -> Optional[ET.Element]:
    style = elem.find(_q(NS_P, "style"))
    if style is None:
        style = elem.find(_q(NS_A, "style"))
    return style


def _style_ref(elem: ET.Element, name: str):
    """Returns (idx, schemeClr element, ref element) for p:style/<name>Ref."""
    style = _find_style(elem)
    if style is None:
        return 0, None, None
    ref = style.find(_q(NS_A, name))
    if ref is None:
        return 0, None, None
    idx = _int_attr(ref, "idx", 0)
    scheme = ref.find(_q(NS_A, "schemeClr"))
    return idx, scheme, ref


def _blob_for_part(part: Any, rid: str) -> Optional[bytes]:
    if not part or not rid:
        return None
    try:
        return part.related_part(rid).blob
    except Exception:
        pass
    try:
        return part.rels[rid].target_part.blob
    except Exception:
        return None


def _blob_for_rid(shape: Any, rid: str) -> Optional[bytes]:
    try:
        return _blob_for_part(shape.part, rid)
    except Exception:
        return None


def _parse_fill(spPr: Optional[ET.Element], colors: Dict[str, RGB]):
    """Returns ("solid", rgba) | ("grad", stops, angle) | ("blip", rid)
    | ("none", None) | None (unspecified)."""
    if spPr is None:
        return None
    if spPr.find(_q(NS_A, "noFill")) is not None:
        return ("none", None)
    sf = spPr.find(_q(NS_A, "solidFill"))
    if sf is not None:
        col = _resolve_element_color(sf, colors)
        return ("solid", col) if col else ("none", None)
    gf = spPr.find(_q(NS_A, "gradFill"))
    if gf is not None:
        stops, ang = _parse_gradient(gf, colors)
        return ("grad", stops, ang)
    bf = spPr.find(_q(NS_A, "blipFill"))
    if bf is not None:
        blip = bf.find(_q(NS_A, "blip"))
        rid = blip.attrib.get(_q(NS_R, "embed")) if blip is not None else None
        if rid:
            return ("blip", rid)
    return None


def _parse_line(spPr: Optional[ET.Element], colors: Dict[str, RGB]
                ) -> Tuple[Optional[RGBA], float, bool]:
    if spPr is None:
        return None, 1.0, False
    ln = spPr.find(_q(NS_A, "ln"))
    if ln is None:
        return None, 1.0, False
    if ln.find(_q(NS_A, "noFill")) is not None:
        return "none", 0.0, False
    col = None
    sf = ln.find(_q(NS_A, "solidFill"))
    if sf is not None:
        col = _resolve_element_color(sf, colors)
    else:
        gf = ln.find(_q(NS_A, "gradFill"))
        if gf is not None:
            stops, _ = _parse_gradient(gf, colors)
            col = stops[0][1] if stops else None
    w_pt = _int_attr(ln, "w", 12700) / 12700.0
    dash_node = ln.find(_q(NS_A, "prstDash"))
    dashed = dash_node is not None and dash_node.attrib.get("val", "solid") != "solid"
    return col, w_pt, dashed


# ---------------------------------------------------------------------------
# Run styles & text body rendering
# ---------------------------------------------------------------------------
class _RunStyle:
    __slots__ = ("font", "size_px", "color", "underline", "faux_bold")

    def __init__(self, font, size_px, color, underline, faux_bold):
        self.font = font
        self.size_px = size_px
        self.color = color
        self.underline = underline
        self.faux_bold = faux_bold


def _parse_run_style(rPr: Optional[ET.Element], colors: Dict[str, RGB],
                     base_size_pt: float, default_color: RGBA,
                     default_family: str, fonts: FontResolver, scale_y: float
                     ) -> _RunStyle:
    size_pt = base_size_pt
    bold = italic = underline = False
    family = default_family
    color = default_color

    if rPr is not None:
        sz = rPr.attrib.get("sz")
        if sz:
            try:
                size_pt = int(sz) / 100.0
            except ValueError:
                pass
        bold = rPr.attrib.get("b") == "1"
        italic = rPr.attrib.get("i") == "1"
        u = rPr.attrib.get("u")
        underline = u is not None and u != "none"
        latin = rPr.find(_q(NS_A, "latin"))
        if latin is not None and latin.attrib.get("typeface"):
            family = latin.attrib["typeface"]
        sf = rPr.find(_q(NS_A, "solidFill"))
        if sf is not None:
            c = _resolve_element_color(sf, colors)
            if c:
                color = c
        elif rPr.find(_q(NS_A, "noFill")) is not None:
            color = (0, 0, 0, 0)

    size_px = max(5, int(round(size_pt * 12700 * scale_y)))
    font, faux = fonts.get_font(family, size_px, bold, italic)
    return _RunStyle(font, size_px, color, underline, faux)


def _wrap_tokens(words: List[Tuple[str, _RunStyle]], max_w: float,
                 wrap_on: bool) -> List[dict]:
    lines: List[dict] = []
    cur: List[Tuple[str, _RunStyle]] = []
    cur_w = 0.0

    def flush():
        nonlocal cur, cur_w
        if cur:
            lines.append({"words": cur, "w": cur_w, "shaped": None})
        cur, cur_w = [], 0.0

    for text, st in words:
        f = st.font
        w_w = _text_width(f, text)

        def hard_split(chunk_text, style):
            nonlocal cur, cur_w
            chunk = ""
            for ch in chunk_text:
                if _text_width(f, chunk + ch) > max_w and chunk:
                    lines.append({"words": [(chunk, style)],
                                  "w": _text_width(f, chunk), "shaped": None})
                    chunk = ch
                else:
                    chunk += ch
            cur, cur_w = [(chunk, style)], _text_width(f, chunk)

        if w_w > max_w and not cur:
            hard_split(text, st)
            continue
        sp_w = _text_width(f, " ") if cur else 0.0
        if wrap_on and cur and cur_w + sp_w + w_w > max_w:
            flush()
            if w_w > max_w:
                hard_split(text, st)
                continue
            cur, cur_w = [(text, st)], w_w
        else:
            cur.append((text, st))
            cur_w += sp_w + w_w
    flush()
    return lines


def _format_autonum(bu_type: str, n: int) -> str:
    typ = (bu_type or "").lower()
    if "alphalc" in typ:
        s = chr(ord('a') + (n - 1) % 26)
    elif "alphauc" in typ:
        s = chr(ord('A') + (n - 1) % 26)
    elif "romanlc" in typ:
        s = ["i", "ii", "iii", "iv", "v", "vi", "vii", "viii", "ix", "x"][(n - 1) % 10]
    elif "romanuc" in typ:
        s = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X"][(n - 1) % 10]
    else:
        s = str(n)
    if "parenboth" in typ:
        return f"({s})"
    if "paren" in typ:
        return f"{s})"
    return f"{s}."


def _para_spacing(pPr: Optional[ET.Element], scale_y: float):
    ln_pct, ln_pt = 1.0, None
    bef_px = aft_px = 0.0
    if pPr is None:
        return ln_pct, ln_pt, bef_px, aft_px
    ln = pPr.find(_q(NS_A, "lnSpc"))
    if ln is not None:
        pct = ln.find(_q(NS_A, "spcPct"))
        pts = ln.find(_q(NS_A, "spcPts"))
        if pct is not None:
            try:
                ln_pct = int(pct.attrib.get("val", "100000")) / 100000.0
            except ValueError:
                pass
        elif pts is not None:
            try:
                ln_pt = int(pts.attrib.get("val", "0")) / 100.0
            except ValueError:
                pass
    for key, out in (("spcBef", "bef"), ("spcAft", "aft")):
        node = pPr.find(_q(NS_A, key))
        if node is None:
            continue
        pts = node.find(_q(NS_A, "spcPts"))
        if pts is not None:
            try:
                px = int(pts.attrib.get("val", "0")) / 100.0 * 12700 * scale_y
                if out == "bef":
                    bef_px = px
                else:
                    aft_px = px
            except ValueError:
                pass
    return ln_pct, ln_pt, bef_px, aft_px


def _draw_text_segment(draw: ImageDraw.ImageDraw, x: float, y: float,
                       text: str, st: Optional[_RunStyle],
                       known_w: Optional[float] = None) -> None:
    if st is None or st.color is None or st.color[3] == 0 or not text:
        return
    kwargs: Dict[str, Any] = {}
    if st.faux_bold and st.size_px >= 12:
        kwargs = {"stroke_width": 1, "stroke_fill": st.color}
    draw.text((x, y), text, fill=st.color, font=st.font, **kwargs)
    if st.underline:
        ascent, _ = st.font.getmetrics()
        w = known_w if known_w is not None else _text_width(st.font, text)
        draw.line([(x, y + ascent + 1), (x + w, y + ascent + 1)],
                  fill=st.color, width=1)


def _render_txbody(txbody: ET.Element, box, img: Image.Image, ctx: Dict[str, Any],
                   hint_txbody: Optional[ET.Element] = None,
                   default_size_pt: Optional[float] = None,
                   default_color: Optional[RGBA] = None) -> None:
    x0, y0, x1, y1 = box
    if x1 - x0 < 4 or y1 - y0 < 4:
        return
    colors = ctx["palette"]
    fonts: FontResolver = ctx["fonts"]
    scale_x, scale_y = ctx["scale_x"], ctx["scale_y"]
    draw = ImageDraw.Draw(img, "RGBA")

    bodyPr = txbody.find(_q(NS_A, "bodyPr"))
    if bodyPr is None and hint_txbody is not None:
        bodyPr = hint_txbody.find(_q(NS_A, "bodyPr"))
    lIns = _int_attr(bodyPr, "lIns", 91440) * scale_x
    rIns = _int_attr(bodyPr, "rIns", 91440) * scale_x
    tIns = _int_attr(bodyPr, "tIns", 45720) * scale_y
    bIns = _int_attr(bodyPr, "bIns", 45720) * scale_y
    anchor = _str_attr(bodyPr, "anchor", "t")
    wrap_on = _str_attr(bodyPr, "wrap", "square") != "none"
    fit_scale = 1.0
    if bodyPr is not None:
        naf = bodyPr.find(_q(NS_A, "normAutofit"))
        if naf is not None:
            try:
                fit_scale = int(naf.attrib.get("fontScale", "100000")) / 100000.0
            except ValueError:
                pass

    paragraphs = txbody.findall(_q(NS_A, "p"))

    def _has_text(ps):
        for p in ps:
            for t in p.iter(_q(NS_A, "t")):
                if (t.text or "").strip():
                    return True
        return False

    if not _has_text(paragraphs):
        return

    hint_size, hint_color = None, None
    if hint_txbody is not None:
        for node in hint_txbody.iter():
            lt = _local(node.tag)
            if lt in ("rPr", "defRPr", "endParaRPr"):
                if hint_size is None and node.attrib.get("sz"):
                    try:
                        hint_size = int(node.attrib["sz"]) / 100.0
                    except ValueError:
                        pass
                if hint_color is None:
                    sf = node.find(_q(NS_A, "solidFill"))
                    if sf is not None:
                        hint_color = _resolve_element_color(sf, colors)
                if hint_size is not None and hint_color is not None:
                    break

    base_default = default_size_pt or hint_size or 18.0
    default_family = fonts.theme_fonts.get("minor") or "Calibri"
    eff_default_color = hint_color or default_color or \
        _rgba(colors.get("tx1", (0, 0, 0)))

    avail_w = (x1 - x0) - lIns - rIns
    inner_h = (y1 - y0) - tIns - bIns

    laid = []
    total_h = 0.0
    counters: Dict[int, int] = {}

    for p in paragraphs:
        pPr = p.find(_q(NS_A, "pPr"))
        algn = _str_attr(pPr, "algn", "l")
        lvl = _int_attr(pPr, "lvl", 0)
        marL = _int_attr(pPr, "marL", 0) * scale_x
        indent = _int_attr(pPr, "indent", 0) * scale_x
        ln_pct, ln_pt, bef_px, aft_px = _para_spacing(pPr, scale_y)

        bullet = None
        autonum = None
        if pPr is not None:
            if pPr.find(_q(NS_A, "buNone")) is not None:
                counters.pop(lvl, None)
            elif pPr.find(_q(NS_A, "buChar")) is not None:
                bullet = pPr.find(_q(NS_A, "buChar")).attrib.get("char", "•")
            elif pPr.find(_q(NS_A, "buAutoNum")) is not None:
                autonum = pPr.find(_q(NS_A, "buAutoNum")).attrib.get("type", "arabicPeriod")

        base_sz = base_default
        if pPr is not None:
            d = pPr.find(_q(NS_A, "defRPr"))
            if d is not None and d.attrib.get("sz"):
                try:
                    base_sz = int(d.attrib["sz"]) / 100.0
                except ValueError:
                    pass
        base_sz *= fit_scale

        words: List[Tuple[str, Optional[_RunStyle]]] = []
        first_style: Optional[_RunStyle] = None
        for child in p:
            lt = _local(child.tag)
            if lt in ("r", "fld"):
                rPr = child.find(_q(NS_A, "rPr"))
                tnode = child.find(_q(NS_A, "t"))
                text = (tnode.text or "") if tnode is not None else ""
                if lt == "fld" and child.attrib.get("type") == "slidenum":
                    text = str(ctx.get("slide_number") or text or "")
                st = _parse_run_style(rPr, colors, base_sz, eff_default_color,
                                      default_family, fonts, scale_y)
                first_style = first_style or st
                for w in text.split():
                    words.append((w, st))
            elif lt == "br":
                words.append(("\n", None))

        segments: List[List[Tuple[str, _RunStyle]]] = [[]]
        for wtext, st in words:
            if wtext == "\n" or st is None:
                segments.append([])
            else:
                segments[-1].append((wtext, st))

        prefix = ""
        if bullet:
            prefix = bullet
        elif autonum:
            counters[lvl] = counters.get(lvl, 0) + 1
            for k in [k2 for k2 in counters if k2 > lvl]:
                del counters[k2]
            prefix = _format_autonum(autonum, counters[lvl])

        line_h_default = max(6, int(base_sz * 12700 * scale_y * 1.2 * ln_pct))
        para_lines: List[dict] = []
        for seg in segments:
            if not seg:
                para_lines.append({"empty": True, "h": line_h_default})
                continue
            is_rtl = any(_is_rtl_text(w) for w, _ in seg)
            seg_lines = _wrap_tokens(seg, max(8.0, avail_w - marL), wrap_on)
            for ln in seg_lines:
                ln["rtl"] = is_rtl
                if is_rtl:
                    logical = " ".join(w for w, _ in ln["words"])
                    ln["shaped"] = _shape_text_for_display(logical)
                    ln["w"] = _text_width(ln["words"][0][1].font, ln["shaped"])
                ln["h"] = (max(8, int(ln_pt * 12700 * scale_y)) if ln_pt
                           else max(8, int(max(s.size_px for _, s in ln["words"])
                                           * 1.2 * ln_pct)))
            para_lines.extend(seg_lines)
        if not para_lines:
            para_lines = [{"empty": True, "h": line_h_default}]

        block_h = sum(l["h"] for l in para_lines) + bef_px + aft_px
        total_h += block_h
        laid.append({"lines": para_lines, "algn": algn, "marL": marL,
                     "indent": indent, "prefix": prefix, "bef": bef_px,
                     "aft": aft_px, "bullet_style": first_style})

    y = y0 + tIns
    if anchor == "ctr":
        y += max(0.0, (inner_h - total_h) / 2.0)
    elif anchor == "b":
        y += max(0.0, inner_h - total_h)

    for blk in laid:
        y += blk["bef"]
        first_line = True
        for ln in blk["lines"]:
            if ln.get("empty"):
                y += ln["h"]
                first_line = False
                continue
            algn = blk["algn"]
            if ln.get("rtl") and algn == "l":
                algn = "r"
            left = x0 + lIns + blk["marL"]
            inner = avail_w - blk["marL"]

            if blk["prefix"] and first_line and blk["bullet_style"]:
                bs = blk["bullet_style"]
                bw = _text_width(bs.font, blk["prefix"])
                bx = x0 + lIns + blk["marL"] + blk["indent"]
                if blk["indent"] >= 0:
                    bx = left - bw - 4
                _draw_text_segment(draw, bx, y + (ln["h"] - bs.size_px) / 2,
                                   blk["prefix"], bs, bw)

            if algn == "ctr":
                tx = left + max(0.0, (inner - ln["w"]) / 2.0)
            elif algn == "r":
                tx = left + max(0.0, inner - ln["w"])
            else:
                tx = left

            if ln.get("shaped") is not None:
                st = ln["words"][0][1]
                _draw_text_segment(draw, tx, y, ln["shaped"], st, ln["w"])
            else:
                cx = tx
                for word, st in ln["words"]:
                    _draw_text_segment(draw, cx, y, word, st)
                    cx += _text_width(st.font, word) + _text_width(st.font, " ")
            y += ln["h"]
            first_line = False
        y += blk["aft"]


# ---------------------------------------------------------------------------
# Pictures, tables, charts
# ---------------------------------------------------------------------------
def _render_picture(shape: Any, elem: ET.Element, box, img: Image.Image,
                    ctx: Dict[str, Any], flip_h: bool, flip_v: bool) -> None:
    x0, y0, x1, y1 = [int(round(v)) for v in box]
    x0, y0 = max(0, x0), max(0, y0)
    blipFill = elem.find(_q(NS_P, "blipFill"))
    if blipFill is None:
        blipFill = elem.find(_q(NS_A, "blipFill"))
    rid, src_rect = None, None
    if blipFill is not None:
        blip = blipFill.find(_q(NS_A, "blip"))
        if blip is not None:
            rid = blip.attrib.get(_q(NS_R, "embed")) or blip.attrib.get(_q(NS_R, "link"))
        sr = blipFill.find(_q(NS_A, "srcRect"))
        if sr is not None:
            src_rect = tuple(_int_attr(sr, k, 0) / 100000.0 for k in ("l", "t", "r", "b"))
    if not rid:
        return
    blob = _blob_for_rid(shape, rid)
    if not blob:
        return
    try:
        pic = Image.open(io.BytesIO(blob)).convert("RGBA")
    except Exception:
        return
    if src_rect:
        l, t, r, b = src_rect
        W, H = pic.size
        box_ = (int(l * W), int(t * H),
                max(int(l * W) + 1, int((1 - r) * W)),
                max(int(t * H) + 1, int((1 - b) * H)))
        pic = pic.crop(box_)
    if flip_h:
        pic = pic.transpose(Image.FLIP_LEFT_RIGHT)
    if flip_v:
        pic = pic.transpose(Image.FLIP_TOP_BOTTOM)
    tw, th = max(1, x1 - x0), max(1, y1 - y0)
    pic = pic.resize((tw, th), RESAMPLE)
    try:
        img.alpha_composite(pic, (x0, y0))
    except ValueError:
        img.paste(pic, (x0, y0), pic)


def _render_table(shape: Any, elem: ET.Element, box, img: Image.Image,
                  ctx: Dict[str, Any]) -> None:
    x0, y0, x1, y1 = box
    tdraw = ImageDraw.Draw(img, "RGBA")
    colors = ctx["palette"]

    graphic = elem.find(_q(NS_P, "graphic"))
    gd = graphic.find(_q(NS_A, "graphicData")) if graphic is not None else None
    tbl = gd.find(_q(NS_A, "tbl")) if gd is not None else None
    if tbl is None:
        return

    grid = tbl.find(_q(NS_A, "tblGrid"))
    col_ws = [_int_attr(g, "w", 0) for g in grid.findall(_q(NS_A, "gridCol"))] if grid else []
    rows = tbl.findall(_q(NS_A, "tr"))
    row_hs = [_int_attr(r, "h", 0) for r in rows]
    n_rows, n_cols = len(rows), len(col_ws)
    if not n_rows or not n_cols:
        return

    tblPr = tbl.find(_q(NS_A, "tblPr"))
    first_row_band = tblPr is not None and _str_attr(tblPr, "firstRow", "1") == "1"

    total_w = sum(col_ws) or 1
    total_h = sum(row_hs) or 1
    col_x = [x0]
    for w in col_ws:
        col_x.append(col_x[-1] + (x1 - x0) * w / total_w)
    row_y = [y0]
    for h in row_hs:
        row_y.append(row_y[-1] + (y1 - y0) * h / total_h)

    cellmap: Dict[Tuple[int, int], Any] = {}
    for r, tr in enumerate(rows):
        c = 0
        for tc in tr.findall(_q(NS_A, "tc")):
            while (r, c) in cellmap:
                c += 1
            gs = max(1, _int_attr(tc, "gridSpan", 1))
            rs = max(1, _int_attr(tc, "rowSpan", 1))
            cellmap[(r, c)] = (tc, gs, rs)
            for dr in range(rs):
                for dc in range(gs):
                    if dr or dc:
                        cellmap[(r + dr, c + dc)] = None
            c += gs

    border = (200, 200, 200, 255)
    for (r, c), entry in cellmap.items():
        if entry is None:
            continue
        tc, gs, rs = entry
        cx0, cy0 = col_x[c], row_y[r]
        cx1 = col_x[min(c + gs, n_cols)]
        cy1 = row_y[min(r + rs, n_rows)]

        tcPr = tc.find(_q(NS_A, "tcPr"))
        fill = None
        if tcPr is not None:
            sf = tcPr.find(_q(NS_A, "solidFill"))
            if sf is not None:
                fill = _resolve_element_color(sf, colors)
        if fill is None and first_row_band and r == 0:
            fill = _rgba(colors.get("accent1", _DEF_ACCENTS[0]))
        elif fill is None:
            fill = (255, 255, 255, 255) if r % 2 == 0 else (242, 242, 242, 255)

        tdraw.rectangle([cx0, cy0, cx1 - 1, cy1 - 1], fill=fill, outline=border)

        text_color = None
        if first_row_band and r == 0 and fill:
            lum = 0.3 * fill[0] + 0.59 * fill[1] + 0.11 * fill[2]
            if lum < 140:
                text_color = (255, 255, 255, 255)

        txbody = tc.find(_q(NS_A, "txBody"))
        if txbody is not None:
            _render_txbody(txbody, (cx0 + 5, cy0 + 2, cx1 - 5, cy1 - 2), img, ctx,
                           default_size_pt=14, default_color=text_color)


def _render_chart(shape: Any, box, tdraw: ImageDraw.ImageDraw,
                  ctx: Dict[str, Any]) -> None:
    x0, y0, x1, y1 = box
    colors_map, fonts = ctx["palette"], ctx["fonts"]
    scale_y = ctx["scale_y"]
    tdraw.rectangle(box, fill=(255, 255, 255, 255), outline=(190, 190, 190, 255))
    try:
        chart = shape.chart
        ctype = str(getattr(chart, "chart_type", "") or "").upper()
        plot = chart.plots[0]
        cats = ["" if c is None else str(c) for c in plot.categories]
        sers = []
        for i, s in enumerate(chart.series):
            try:
                name = str(s.name) if s.name is not None else f"Series {i + 1}"
            except Exception:
                name = f"Series {i + 1}"
            vals = []
            for v in (s.values or []):
                try:
                    vals.append(float(v))
                except (TypeError, ValueError):
                    vals.append(0.0)
            sers.append((name, vals))
        if not sers or not cats or not any(sers[0][1]):
            raise ValueError("no chart data")

        colors = [_rgba(colors_map.get(f"accent{i}", _DEF_ACCENTS[i - 1]))
                  for i in range(1, 7)]
        plot_box = (x0 + 44, y0 + 26, x1 - 10, y1 - 44)
        pw, ph = plot_box[2] - plot_box[0], plot_box[3] - plot_box[1]

        if "PIE" in ctype or "DOUGHNUT" in ctype:
            vals = sers[0][1]
            total = sum(vals) or 1.0
            ccx, ccy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
            rr = min(x1 - x0, y1 - y0) * 0.38
            start = -90.0
            for i, v in enumerate(vals):
                sweep = 360.0 * v / total
                tdraw.pieslice([ccx - rr, ccy - rr, ccx + rr, ccy + rr],
                               start, start + sweep, fill=colors[i % 6])
                start += sweep
        elif "LINE" in ctype or "AREA" in ctype or "XY_SCATTER" in ctype:
            maxv = max(max(v) for _, v in sers) or 1.0
            minv = min(0.0, min(min(v) for _, v in sers))
            rng = (maxv - minv) or 1.0
            n = max(1, len(cats) - 1)
            for si, (_, vals) in enumerate(sers[:6]):
                pts = [(plot_box[0] + pw * (i / n),
                        plot_box[3] - ph * (v - minv) / rng)
                       for i, v in enumerate(vals)]
                if pts:
                    tdraw.line(pts, fill=colors[si % 6], width=2, joint="curve")
            tdraw.line([plot_box[0], plot_box[3], plot_box[2], plot_box[3]],
                       fill=(120, 120, 120, 255))
            tdraw.line([plot_box[0], plot_box[1], plot_box[0], plot_box[3]],
                       fill=(120, 120, 120, 255))
        else:
            maxv = max(max(v) for _, v in sers) or 1.0
            minv = min(0.0, min(min(v) for _, v in sers))
            rng = (maxv - minv) or 1.0
            zero_y = plot_box[3] - ph * (0 - minv) / rng
            ncat = max(1, len(cats))
            gw = pw / ncat
            barw = gw * 0.7 / max(1, min(6, len(sers)))
            for si, (_, vals) in enumerate(sers[:6]):
                for i, v in enumerate(vals):
                    bx0 = plot_box[0] + gw * i + gw * 0.15 + barw * si
                    bh = ph * (v - minv) / rng
                    tdraw.rectangle([bx0, zero_y - bh, bx0 + barw, zero_y],
                                    fill=colors[si % 6])
            tdraw.line([plot_box[0], zero_y, plot_box[2], zero_y],
                       fill=(120, 120, 120, 255))
            tdraw.line([plot_box[0], plot_box[1], plot_box[0], plot_box[3]],
                       fill=(120, 120, 120, 255))

        f, _ = fonts.get_font(None, max(8, int(10 * 12700 * scale_y)), False, False)
        lx = x0 + 10
        for si, (name, _) in enumerate(sers[:6]):
            label = name[:16]
            tdraw.rectangle([lx, y1 - 28, lx + 10, y1 - 18], fill=colors[si % 6])
            tdraw.text((lx + 14, y1 - 30), label, fill=(60, 60, 60, 255), font=f)
            lx += 26 + _text_width(f, label)
            if lx > x1 - 60:
                break
    except Exception as exc:
        log.debug("chart sketch failed: %s", exc)
        f, _ = fonts.get_font(None, 12, False, False)
        tdraw.text(((x0 + x1) / 2 - 18, (y0 + y1) / 2), "Chart",
                   fill=(150, 150, 150, 255), font=f)


# ---------------------------------------------------------------------------
# Slide background resolution — full bgPr + bgRef support
# ---------------------------------------------------------------------------
def _paint_slide_background(img: Image.Image, slide: Any,
                            ctx: Dict[str, Any]) -> bool:
    """
    Walks slide -> layout -> master looking for <p:bg>, then paints it
    (solid / gradient / image) onto the canvas. bgRef entries are resolved
    against the theme's bgFillStyleLst with phClr substitution.
    """
    W, H = img.size
    theme: Theme = ctx["theme"]
    colors = theme.colors

    def full_rect(mdraw: ImageDraw.ImageDraw) -> None:
        mdraw.rectangle([0, 0, W - 1, H - 1], fill=255)

    containers = []
    for attr_chain in ((slide,), (slide, "slide_layout"),
                       (slide, "slide_layout", "slide_master")):
        try:
            obj = slide
            for attr in attr_chain[1:]:
                obj = getattr(obj, attr)
            containers.append(obj)
        except Exception:
            continue

    for cont in containers:
        bg = None
        try:
            cSld = cont.element.find(_q(NS_P, "cSld"))
            if cSld is not None:
                bg = cSld.find(_q(NS_P, "bg"))
        except Exception:
            bg = None
        if bg is None:
            continue

        fill = None
        bgRef = bg.find(_q(NS_P, "bgRef"))
        if bgRef is not None:
            idx = _int_attr(bgRef, "idx", 0)
            sc = bgRef.find(_q(NS_A, "schemeClr"))
            base = _resolve_element_color(sc, colors)
            if idx >= 1000:
                # bgRef indexes bgFillStyleLst as (idx - 1000), 1-based
                fill = _resolve_theme_fill(theme.fmt_scheme, "bgFillStyleLst",
                                           idx - 1000,
                                           base[:3] if base else None, colors)
            elif base is not None:
                fill = ("solid", base)
        else:
            bgPr = bg.find(_q(NS_P, "bgPr"))
            if bgPr is not None:
                f = _parse_fill(bgPr, colors)
                if f and f[0] == "blip":
                    blob = _blob_for_part(getattr(cont, "part", None), f[1])
                    if blob:
                        fill = ("blip", blob)
                else:
                    fill = f

        if fill is None:
            continue
        kind = fill[0]
        if kind == "solid" and fill[1]:
            d = ImageDraw.Draw(img, "RGBA")
            d.rectangle([0, 0, W - 1, H - 1], fill=_rgba(fill[1]))
            return True
        if kind == "grad":
            _paint_gradient(img, (0, 0, W, H), fill[1], fill[2], full_rect)
            return True
        if kind == "blip":
            try:
                pic = Image.open(io.BytesIO(fill[1])).convert("RGBA")
                pic = pic.resize((W, H), RESAMPLE)
                img.alpha_composite(pic)
            except Exception:
                log.debug("background image failed", exc_info=True)
            return True
        if kind == "none":
            return True  # explicitly transparent — keep base
    return False


# ---------------------------------------------------------------------------
# Shape leaf painting (now with p:style fill/line/font refs)
# ---------------------------------------------------------------------------
def _placeholder_fallback_size(shape: Any, ctx: Dict[str, Any]) -> Optional[Tuple[int, int]]:
    try:
        idx = shape.placeholder_format.idx
        lph = ctx.get("layout_ph", {}).get(idx)
        if lph is not None:
            w = getattr(lph, "width", None)
            h = getattr(lph, "height", None)
            if w and h:
                return int(w), int(h)
            for msh in lph.part.slide_layout.slide_master.shapes:
                if getattr(msh, "is_placeholder", False) and \
                        msh.placeholder_format.idx == idx:
                    if msh.width and msh.height:
                        return int(msh.width), int(msh.height)
    except Exception:
        pass
    return None


def _geom_of(shape: Any, elem: ET.Element, ctx: Dict[str, Any]):
    xf = _find_xfrm(elem)
    if xf is not None:
        off = xf.find(_q(NS_A, "off"))
        ext = xf.find(_q(NS_A, "ext"))
        left = _int_attr(off, "x", 0)
        top = _int_attr(off, "y", 0)
        w = _int_attr(ext, "cx", 0)
        h = _int_attr(ext, "cy", 0)
        rot = _int_attr(xf, "rot", 0)
        fh = xf.attrib.get("flipH") == "1"
        fv = xf.attrib.get("flipV") == "1"
    else:
        left = getattr(shape, "left", None) or 0
        top = getattr(shape, "top", None) or 0
        w = getattr(shape, "width", None) or 0
        h = getattr(shape, "height", None) or 0
        rot, fh, fv = 0, False, False

    if (w == 0 or h == 0) and getattr(shape, "is_placeholder", False):
        fb = _placeholder_fallback_size(shape, ctx)
        if fb:
            w = w or fb[0]
            h = h or fb[1]
    return left, top, w, h, rot, fh, fv


def _hint_txbody(shape: Any, ctx: Dict[str, Any]) -> Optional[ET.Element]:
    try:
        if getattr(shape, "is_placeholder", False):
            idx = shape.placeholder_format.idx
            lph = ctx.get("layout_ph", {}).get(idx)
            if lph is not None:
                return lph.element.find(_q(NS_P, "txBody"))
    except Exception:
        pass
    return None


def _paint_leaf(shape: Any, elem: ET.Element, box, timg: Image.Image,
                ctx: Dict[str, Any], rot: int, flip_h: bool, flip_v: bool) -> None:
    tdraw = ImageDraw.Draw(timg, "RGBA")
    tag = _local(elem.tag)
    theme: Theme = ctx["theme"]
    colors = theme.colors

    if tag == "pic":
        _render_picture(shape, elem, box, timg, ctx, flip_h, flip_v)
        return
    if tag == "graphicFrame":
        graphic = elem.find(_q(NS_P, "graphic"))
        gd = graphic.find(_q(NS_A, "graphicData")) if graphic is not None else None
        if gd is not None:
            uri = gd.attrib.get("uri", "")
            if gd.find(_q(NS_A, "tbl")) is not None or "table" in uri:
                _render_table(shape, elem, box, timg, ctx)
            elif "chart" in uri:
                _render_chart(shape, box, tdraw, ctx)
        return

    # ---- sp / cxnSp -------------------------------------------------------
    spPr = _find_spPr(elem)
    fill = _parse_fill(spPr, colors)
    line_col, line_pt, dashed = _parse_line(spPr, colors)

    # --- theme style refs (p:style) ----------------------------------------
    # fillRef: only when spPr says nothing about fill
    if fill is None:
        idx, sc, ref = _style_ref(elem, "fillRef")
        if idx and sc is not None:
            base = _resolve_element_color(sc, colors)
            if base:
                fill = _resolve_theme_fill(theme.fmt_scheme, "fillStyleLst",
                                           idx, base[:3], colors)

    # lnRef: only when spPr's <a:ln> says nothing about color
    if line_col is None:
        idx, sc, ref = _style_ref(elem, "lnRef")
        if idx and sc is not None:
            base = _resolve_element_color(sc, colors)
            if base:
                res = _resolve_theme_fill(theme.fmt_scheme, "lnStyleLst",
                                          idx, base[:3], colors)
                if res and res[0] == "solid" and res[1]:
                    line_col = res[1]
                    wattr = ref.attrib.get("w")
                    if wattr:
                        try:
                            line_pt = max(line_pt, int(wattr) / 12700.0)
                        except (TypeError, ValueError):
                            pass

    # fontRef: default text color from theme style
    font_default: Optional[RGBA] = None
    _fi, sc_f, _ref_f = _style_ref(elem, "fontRef")
    if sc_f is not None:
        c = _resolve_element_color(sc_f, colors)
        if c and c[3] > 0:
            font_default = c

    # python-pptx API fallback (style-referenced or theme colors)
    if fill is None:
        try:
            f = shape.fill
            if f.type == 1:
                rgb = f.fore_color.rgb
                fill = ("solid", (rgb[0], rgb[1], rgb[2], 255))
        except Exception:
            pass
    if line_col is None:
        try:
            lf = shape.line.fill
            if lf.type == 1:
                rgb = shape.line.color.rgb
                line_col = (rgb[0], rgb[1], rgb[2], 255)
                if shape.line.width:
                    line_pt = shape.line.width.pt
        except Exception:
            pass

    lw_px = 0
    if isinstance(line_col, tuple):
        lw_px = max(1, int(round(line_pt * 12700 * ctx["scale_x"])))

    pg = elem.find(f".//{_q(NS_A, 'prstGeom')}")
    custGeom = elem.find(f".//{_q(NS_A, 'custGeom')}")
    prst = pg.attrib.get("prst", "rect") if pg is not None else None
    adj = _parse_adj(pg)

    is_connector = (_local(elem.tag) == "cxnSp" or
                    (prst or "").startswith(("straightConnector", "bentConnector",
                                             "curvedConnector")))

    if custGeom is not None:
        _render_custom_geom(timg, tdraw, custGeom, box,
                            _rgba(fill[1]) if fill and fill[0] == "solid"
                            and fill[1] else None,
                            _rgba(line_col) if isinstance(line_col, tuple) else None,
                            lw_px)
    elif is_connector:
        ops = _connector_ops(prst, box, flip_h, flip_v)
        col = _rgba(line_col) if isinstance(line_col, tuple) else (100, 100, 100, 255)
        _stroke_ops(tdraw, ops, col, lw_px or 1, dashed)
    else:
        ops = _prst_ops(prst, box, adj)
        if flip_h or flip_v:
            bx0, by0, bx1, by1 = box
            mirrored = []
            for kind, data in ops:
                if kind == "poly":
                    data = [((bx1 - (px - bx0)) if flip_h else px,
                             (by1 - (py - by0)) if flip_v else py) for px, py in data]
                mirrored.append((kind, data))
            ops = mirrored

        if fill and fill[0] == "grad":
            _paint_gradient(timg, box, fill[1], fill[2], _make_silhouette(ops))
        elif fill and fill[0] == "blip":
            blob = _blob_for_rid(shape, fill[1])
            if blob:
                _paint_blip(timg, box, blob, _make_silhouette(ops))
        elif fill and fill[0] == "solid" and fill[1]:
            _fill_ops(tdraw, ops, _rgba(fill[1]))
        if isinstance(line_col, tuple):
            _stroke_ops(tdraw, ops, _rgba(line_col), lw_px, dashed)

    # ---- text --------------------------------------------------------------
    txbody = elem.find(_q(NS_P, "txBody")) or elem.find(_q(NS_A, "txBody"))
    if txbody is not None:
        _render_txbody(txbody, box, timg, ctx, hint_txbody=_hint_txbody(shape, ctx),
                       default_color=font_default)


# ---------------------------------------------------------------------------
# Recursive shape rendering (groups, rotation)
# ---------------------------------------------------------------------------
def _compose_parent(parent, ox, oy, fx, fy):
    if not parent:
        return (ox, oy, fx, fy)
    pox, poy, pfx, pfy = parent
    return (pox + pfx * ox, poy + pfy * oy, pfx * fx, pfy * fy)


def _render_shape_recursive(shape: Any, img: Image.Image, draw, scale_x: float,
                            scale_y: float, palette: Dict[str, RGB],
                            group_offset: Optional[Tuple[float, float, float, float]] = None,
                            ctx: Optional[Dict[str, Any]] = None,
                            fonts: Optional[FontResolver] = None) -> None:
    if ctx is None:
        theme = Theme(palette)
        ctx = {"palette": theme.colors, "theme": theme,
               "fonts": fonts or FontResolver(),
               "scale_x": scale_x, "scale_y": scale_y,
               "slide_number": None, "layout_ph": {}}
    try:
        elem = shape.element

        # ---- group shape ----
        if hasattr(shape, "shapes") and _local(elem.tag) in ("grpSp",):
            xf = _find_xfrm(elem)
            gox, goy, gfx, gfy = 0.0, 0.0, 1.0, 1.0
            if xf is not None:
                off = xf.find(_q(NS_A, "off"))
                ext = xf.find(_q(NS_A, "ext"))
                choff = xf.find(_q(NS_A, "chOff"))
                chext = xf.find(_q(NS_A, "chExt"))
                gx, gy = _int_attr(off, "x", 0), _int_attr(off, "y", 0)
                gw, gh = _int_attr(ext, "cx", 0) or 1, _int_attr(ext, "cy", 0) or 1
                cx0, cy0 = _int_attr(choff, "x", 0), _int_attr(choff, "y", 0)
                cw, ch = _int_attr(chext, "cx", 0) or 1, _int_attr(chext, "cy", 0) or 1
                gfx, gfy = gw / cw, gh / ch
                gox, goy = gx - cx0 * gfx, gy - cy0 * gfy
            comp = _compose_parent(group_offset, gox, goy, gfx, gfy)
            for child in shape.shapes:
                _render_shape_recursive(child, img, None, scale_x, scale_y,
                                        palette, group_offset=comp, ctx=ctx)
            return

        if _is_hidden(elem):
            return

        left, top, w, h, rot, flip_h, flip_v = _geom_of(shape, elem, ctx)
        if group_offset:
            gox, goy, gfx, gfy = group_offset
            left = gox + left * gfx
            top = goy + top * gfy
            w, h = w * gfx, h * gfy

        box = (left * scale_x, top * scale_y,
               (left + w) * scale_x, (top + h) * scale_y)
        if box[2] - box[0] < 1 or box[3] - box[1] < 1:
            if not (getattr(shape, "has_text_frame", False)
                    and shape.text_frame.text.strip()):
                return
            box = (box[0], box[1], box[0] + 2, box[1] + 2)

        def paint(timg: Image.Image):
            _paint_leaf(shape, elem, box, timg, ctx, rot, flip_h, flip_v)

        rot_deg = rot / 60000.0
        if rot_deg % 360 != 0:
            ccx = (box[0] + box[2]) / 2.0
            ccy = (box[1] + box[3]) / 2.0
            layer = Image.new("RGBA", img.size, (0, 0, 0, 0))
            paint(layer)
            layer = layer.rotate(-rot_deg, resample=BICUBIC, center=(ccx, ccy))
            img.alpha_composite(layer)
        else:
            paint(img)

    except Exception:
        log.exception("shape render failed: %s", getattr(shape, "name", "?"))


# ---------------------------------------------------------------------------
# Slide & file rendering (public API)
# ---------------------------------------------------------------------------
def render_pptx_slide_to_image(slide: Any,
                               slide_width_emu: int,
                               slide_height_emu: int,
                               target_width_px: int = 850,
                               palette: Optional[Dict[str, RGB]] = None,
                               fonts: Optional[FontResolver] = None,
                               slide_number: Optional[int] = None,
                               render_master: bool = True,
                               theme: Optional[Theme] = None) -> Image.Image:
    """
    Renders a high-fidelity 2D preview of a PPTX slide. Accepts either a
    Theme object (preferred) or a legacy palette dict.
    """
    if theme is None:
        theme = Theme(palette) if palette else Theme()

    aspect = slide_height_emu / slide_width_emu if slide_width_emu else 9 / 16
    target_h = max(1, int(round(target_width_px * aspect)))
    scale_x = target_width_px / slide_width_emu if slide_width_emu else 1.0
    scale_y = target_h / slide_height_emu if slide_height_emu else 1.0

    fonts = fonts or FontResolver()
    if theme.major_font or theme.minor_font:
        fonts.set_theme_fonts(theme.major_font, theme.minor_font)

    ctx: Dict[str, Any] = {
        "palette": theme.colors, "theme": theme, "fonts": fonts,
        "scale_x": scale_x, "scale_y": scale_y,
        "slide_number": slide_number, "layout_ph": {},
    }

    # layout placeholder index map (for geometry/style inheritance)
    slide_ph_idxs = set()
    try:
        for sh in slide.shapes:
            if getattr(sh, "is_placeholder", False):
                try:
                    slide_ph_idxs.add(sh.placeholder_format.idx)
                except Exception:
                    pass
        for lsh in slide.slide_layout.shapes:
            if getattr(lsh, "is_placeholder", False):
                try:
                    ctx["layout_ph"][lsh.placeholder_format.idx] = lsh
                except Exception:
                    pass
    except Exception:
        pass

    img = Image.new("RGBA", (target_width_px, target_h),
                    tuple(theme.colors.get("bg1", (255, 255, 255))) + (255,))

    # slide/layout/master background (solid, gradient, image, bgRef)
    try:
        _paint_slide_background(img, slide, ctx)
    except Exception:
        log.debug("background paint failed", exc_info=True)

    def _r(sh):
        _render_shape_recursive(sh, img, None, scale_x, scale_y,
                                theme.colors, ctx=ctx)

    if render_master:
        try:
            for msh in slide.slide_layout.slide_master.shapes:
                if not getattr(msh, "is_placeholder", False):
                    _r(msh)
        except Exception:
            log.debug("master render failed", exc_info=True)

        try:
            for lsh in slide.slide_layout.shapes:
                if not getattr(lsh, "is_placeholder", False):
                    _r(lsh)
                else:
                    try:
                        if lsh.placeholder_format.idx not in slide_ph_idxs:
                            _r(lsh)
                    except Exception:
                        _r(lsh)
        except Exception:
            log.debug("layout render failed", exc_info=True)

    for sh in slide.shapes:
        _r(sh)

    return img.convert("RGB")


def render_pptx_file_previews(pptx_path: Path | str,
                              target_width_px: int = 850,
                              render_master: bool = True,
                              progress=None) -> List[Image.Image]:
    """Renders all slides of a PPTX file into a list of PIL RGB images."""
    prs = Presentation(str(pptx_path))
    default_theme = Theme.from_presentation(prs)
    fonts = FontResolver()
    fonts.set_theme_fonts(default_theme.major_font, default_theme.minor_font)

    theme_cache: Dict[int, Theme] = {}
    images: List[Image.Image] = []
    slides = list(prs.slides)
    for i, slide in enumerate(slides, start=1):
        theme = _theme_for_slide(slide, theme_cache, default_theme)
        images.append(render_pptx_slide_to_image(
            slide, prs.slide_width, prs.slide_height,
            target_width_px=target_width_px,
            fonts=fonts, slide_number=i,
            render_master=render_master, theme=theme))
        if progress:
            try:
                progress(i, len(slides))
            except Exception:
                pass
    return images


# ---------------------------------------------------------------------------
# Encoding helpers
# ---------------------------------------------------------------------------
def image_to_base64_jpeg(img: Image.Image, quality: int = 85) -> str:
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="JPEG", quality=quality)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode("utf-8")


def image_to_base64_png(img: Image.Image) -> str:
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="PNG", optimize=True)
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode("utf-8")


__all__ = [
    "Theme",
    "FontResolver",
    "render_pptx_file_previews",
    "render_pptx_slide_to_image",
    "image_to_base64_jpeg",
    "image_to_base64_png",
    "_extract_theme_color_palette",
]


if __name__ == "__main__":
    import sys
    logging.basicConfig(level=logging.WARNING)
    if len(sys.argv) > 1:
        src = sys.argv[1]
        out_dir = sys.argv[2] if len(sys.argv) > 2 else "previews"
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        for i, im in enumerate(render_pptx_file_previews(src), 1):
            p = Path(out_dir) / f"slide_{i:03d}.png"
            im.save(p)
            print(f"saved {p}")
    else:
        print("usage: python pptx_preview_engine.py <file.pptx> [out_dir]")