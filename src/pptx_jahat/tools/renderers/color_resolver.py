"""
color_resolver.py
=================
DrawingML color space decoding, theme palette resolution, HSL/RGB transforms,
and theme format schemes (fillStyleLst, lnStyleLst, bgFillStyleLst, effectStyleLst).
"""

from __future__ import annotations

import colorsys
import logging
import xml.etree.ElementTree as ET
from typing import Dict, Optional, Tuple, Any, List

RGBA = Tuple[int, int, int, int]
RGB = Tuple[int, int, int]

log = logging.getLogger("pptx_renderers.color")

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

_DEFAULT_COLORS: Dict[str, RGB] = {
    "dk1": (0, 0, 0), "lt1": (255, 255, 255),
    "dk2": (68, 84, 106), "lt2": (237, 238, 239),
    "accent1": (68, 114, 196), "accent2": (237, 125, 49),
    "accent3": (165, 165, 165), "accent4": (91, 155, 213),
    "accent5": (112, 173, 71), "accent6": (255, 192, 0),
    "hlink": (5, 99, 193), "folHlink": (114, 114, 114),
    "bg1": (255, 255, 255), "bg2": (237, 238, 239),
    "tx1": (0, 0, 0), "tx2": (68, 84, 106),
    "phClr": (68, 114, 196),
}
_SCHEME_ALIASES = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}
_SYSCLR_DEFAULTS = {
    "window": (255, 255, 255), "windowtext": (0, 0, 0),
    "menu": (255, 255, 255), "menutext": (0, 0, 0),
    "buttonface": (240, 240, 240), "buttontext": (0, 0, 0),
    "buttonshadow": (160, 160, 160), "highlight": (0, 120, 215),
    "highlighttext": (255, 255, 255), "3dlight": (240, 240, 240),
    "infobackground": (255, 255, 225), "infotext": (0, 0, 0),
    "activetitle": (0, 90, 160), "graytext": (128, 128, 128),
    "activeborder": (255, 255, 255), "inactiveborder": (212, 208, 200),
    "inactivecaption": (128, 128, 128), "inactivecaptiontext": (0, 0, 0),
    "captiontext": (0, 0, 0), "applicationworkspace": (103, 103, 103),
    "scrollbar": (200, 200, 200), "threeddarkshadow": (105, 105, 105),
    "threedface": (240, 240, 240), "threedlight": (255, 255, 255),
    "threedshadow": (160, 160, 160), "windowframe": (0, 0, 0),
    "windowtext": (0, 0, 0),
}
_COLOR_TAGS = {"srgbClr", "schemeClr", "sysClr", "prstClr", "hslClr", "scrgbClr"}


def q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def local_name(tag: str) -> str:
    return tag.split("}")[-1]


def clamp(v: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, v))


def hex_to_rgb(hv: Optional[str]) -> Optional[RGB]:
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


def scheme_rgb(val: Optional[str], colors: Dict[str, RGB], ph_color: Optional[RGB] = None) -> RGB:
    if not val:
        return colors.get("accent1", (68, 114, 196))
    if val == "phClr":
        return ph_color if ph_color is not None else colors.get("accent1", (68, 114, 196))
    val = _SCHEME_ALIASES.get(val, val)
    return colors.get(val, colors.get("accent1", (68, 114, 196)))


def parse_color_elem(elem: ET.Element, colors: Dict[str, RGB], ph_color: Optional[RGB] = None) -> Optional[RGBA]:
    """Decodes a single DrawingML color element (srgbClr, schemeClr, sysClr, etc.) with full modifiers."""
    tag = local_name(elem.tag)
    rgb: Optional[RGB] = None

    if tag == "srgbClr":
        rgb = hex_to_rgb(elem.attrib.get("val"))
    elif tag == "schemeClr":
        rgb = scheme_rgb(elem.attrib.get("val"), colors, ph_color)
    elif tag == "sysClr":
        val = elem.attrib.get("val", "window")
        last_clr = elem.attrib.get("lastClr")
        rgb = hex_to_rgb(last_clr) or _SYSCLR_DEFAULTS.get(val, (128, 128, 128))
    elif tag == "prstClr":
        from PIL import ImageColor
        try:
            rgb = ImageColor.getrgb(elem.attrib.get("val", "black"))[:3]
        except Exception:
            rgb = (0, 0, 0)
    elif tag == "hslClr":
        try:
            h = float(elem.attrib.get("hue", 0)) / 60000.0 / 360.0
            s = float(elem.attrib.get("sat", 0)) / 100000.0
            l = float(elem.attrib.get("lum", 0)) / 100000.0
            r, g, b = colorsys.hls_to_rgb(h, l, s)
            rgb = (int(round(r * 255)), int(round(g * 255)), int(round(b * 255)))
        except Exception:
            rgb = (0, 0, 0)
    elif tag == "scrgbClr":
        try:
            r = float(elem.attrib.get("r", 0)) / 100000.0
            g = float(elem.attrib.get("g", 0)) / 100000.0
            b = float(elem.attrib.get("b", 0)) / 100000.0
            rgb = (int(round(clamp(r, 0, 1) * 255)),
                   int(round(clamp(g, 0, 1) * 255)),
                   int(round(clamp(b, 0, 1) * 255)))
        except Exception:
            rgb = (0, 0, 0)

    if rgb is None:
        return None

    # Apply DrawingML Color Modifiers (HSL transforms, tint, shade, alpha)
    r, g, b = rgb[0] / 255.0, rgb[1] / 255.0, rgb[2] / 255.0
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    alpha = 1.0

    lum_mod_node = elem.find(q(NS_A, "lumMod"))
    lum_off_node = elem.find(q(NS_A, "lumOff"))
    if lum_mod_node is not None or lum_off_node is not None:
        mod = float(lum_mod_node.attrib.get("val", 100000)) / 100000.0 if lum_mod_node is not None else 1.0
        off = float(lum_off_node.attrib.get("val", 0)) / 100000.0 if lum_off_node is not None else 0.0
        l = clamp(l * mod + off, 0.0, 1.0)

    sat_mod_node = elem.find(q(NS_A, "satMod"))
    sat_off_node = elem.find(q(NS_A, "satOff"))
    if sat_mod_node is not None or sat_off_node is not None:
        mod = float(sat_mod_node.attrib.get("val", 100000)) / 100000.0 if sat_mod_node is not None else 1.0
        off = float(sat_off_node.attrib.get("val", 0)) / 100000.0 if sat_off_node is not None else 0.0
        s = clamp(s * mod + off, 0.0, 1.0)

    hue_mod_node = elem.find(q(NS_A, "hueMod"))
    if hue_mod_node is not None:
        mod = float(hue_mod_node.attrib.get("val", 100000)) / 100000.0
        h = (h * mod) % 1.0

    # Shade & Tint
    shade_node = elem.find(q(NS_A, "shade"))
    if shade_node is not None:
        factor = float(shade_node.attrib.get("val", 100000)) / 100000.0
        r, g, b = r * factor, g * factor, b * factor
        h, l, s = colorsys.rgb_to_hls(r, g, b)

    tint_node = elem.find(q(NS_A, "tint"))
    if tint_node is not None:
        factor = float(tint_node.attrib.get("val", 100000)) / 100000.0
        r = r * factor + (1.0 - factor)
        g = g * factor + (1.0 - factor)
        b = b * factor + (1.0 - factor)
        h, l, s = colorsys.rgb_to_hls(r, g, b)

    inv_node = elem.find(q(NS_A, "inv"))
    if inv_node is not None:
        r, g, b = 1.0 - r, 1.0 - g, 1.0 - b
        h, l, s = colorsys.rgb_to_hls(r, g, b)

    # Alpha modifiers
    alpha_node = elem.find(q(NS_A, "alpha"))
    if alpha_node is not None:
        alpha = float(alpha_node.attrib.get("val", 100000)) / 100000.0
    alpha_mod_node = elem.find(q(NS_A, "alphaMod"))
    if alpha_mod_node is not None:
        alpha *= float(alpha_mod_node.attrib.get("val", 100000)) / 100000.0
    alpha_off_node = elem.find(q(NS_A, "alphaOff"))
    if alpha_off_node is not None:
        alpha += float(alpha_off_node.attrib.get("val", 0)) / 100000.0
    alpha = clamp(alpha, 0.0, 1.0)

    # Reconstruct RGB
    r_f, g_f, b_f = colorsys.hls_to_rgb(h, l, s)
    return (
        int(round(clamp(r_f, 0.0, 1.0) * 255)),
        int(round(clamp(g_f, 0.0, 1.0) * 255)),
        int(round(clamp(b_f, 0.0, 1.0) * 255)),
        int(round(alpha * 255)),
    )


def resolve_element_color(elem: Optional[ET.Element], colors: Dict[str, RGB], ph_color: Optional[RGB] = None) -> Optional[RGBA]:
    """Finds any color child tag inside `elem` and parses it into RGBA."""
    if elem is None:
        return None
    for child in elem:
        if local_name(child.tag) in _COLOR_TAGS:
            c = parse_color_elem(child, colors, ph_color)
            if c is not None:
                return c
    return None


def resolve_theme_fill(fmt_scheme: Optional[ET.Element], lst_tag: str, idx: int,
                       ph_color: Optional[RGB], colors: Dict[str, RGB]) -> Optional[Tuple[str, Any, Any]]:
    """Resolves format scheme style references (fillRef, lnRef, bgRef) index (1-based) against fmtScheme."""
    if fmt_scheme is None or idx <= 0:
        return None
    lst = fmt_scheme.find(q(NS_A, lst_tag))
    if lst is None:
        return None
    children = list(lst)
    if not children:
        return None
    c_idx = (idx - 1) % len(children)
    target = children[c_idx]
    tag = local_name(target.tag)

    if tag == "solidFill":
        c = resolve_element_color(target, colors, ph_color)
        return ("solid", c, None) if c else None
    elif tag == "gradFill":
        from pptx_jahat.tools.renderers.gradient_engine import parse_gradient
        stops, ang, info = parse_gradient(target, colors, ph_color)
        return ("grad", stops, (ang, info))
    elif tag == "blipFill":
        blip = target.find(q(NS_A, "blip"))
        rid = blip.attrib.get(q(NS_R, "embed")) if blip is not None else None
        return ("blip", rid, None)
    elif tag == "noFill":
        return ("solid", (0, 0, 0, 0), None)
    return None


class Theme:
    """A theme's color scheme, font mappings, and DrawingML format schemes."""

    def __init__(self, colors: Optional[Dict[str, RGB]] = None):
        self.colors: Dict[str, RGB] = dict(_DEFAULT_COLORS if colors is None else colors)
        self.major_font: Optional[str] = None
        self.minor_font: Optional[str] = None
        self.fmt_scheme: Optional[ET.Element] = None

    def load_from_theme_part(self, blob: bytes) -> None:
        try:
            root = ET.fromstring(blob)
        except Exception:
            return
        # Parse color scheme
        clr_scheme = root.find(f".//{q(NS_A, 'clrScheme')}")
        if clr_scheme is not None:
            for child in clr_scheme:
                tag = local_name(child.tag)
                for sc in child:
                    c_tag = local_name(sc.tag)
                    if c_tag == "srgbClr":
                        rgb = hex_to_rgb(sc.attrib.get("val"))
                        if rgb:
                            self.colors[tag] = rgb
                    elif c_tag == "sysClr":
                        lc = sc.attrib.get("lastClr")
                        rgb = hex_to_rgb(lc) or _SYSCLR_DEFAULTS.get(sc.attrib.get("val", ""), None)
                        if rgb:
                            self.colors[tag] = rgb

        # Parse font scheme
        font_scheme = root.find(f".//{q(NS_A, 'fontScheme')}")
        if font_scheme is not None:
            major = font_scheme.find(f".//{q(NS_A, 'majorFont')}/{q(NS_A, 'latin')}")
            minor = font_scheme.find(f".//{q(NS_A, 'minorFont')}/{q(NS_A, 'latin')}")
            if major is not None:
                self.major_font = major.attrib.get("typeface")
            if minor is not None:
                self.minor_font = minor.attrib.get("typeface")

        # Format scheme
        fmt_scheme = root.find(f".//{q(NS_A, 'fmtScheme')}")
        if fmt_scheme is not None:
            self.fmt_scheme = fmt_scheme

    @classmethod
    def from_presentation(cls, prs: Any) -> "Theme":
        theme = cls()
        try:
            for slide_master in prs.slide_masters:
                for rel in slide_master.part.rels.values():
                    if "theme" in rel.reltype:
                        theme.load_from_theme_part(rel.target_part.blob)
                        return theme
        except Exception:
            pass
        return theme


class ColorResolver:
    """Helper facade for high-level color resolution operations."""

    def __init__(self, theme: Theme):
        self.theme = theme

    def resolve_color(self, elem: Optional[ET.Element], ph_color: Optional[RGB] = None) -> Optional[RGBA]:
        return resolve_element_color(elem, self.theme.colors, ph_color)

    def get_scheme_rgb(self, val: str) -> RGB:
        return scheme_rgb(val, self.theme.colors)
