import io
import base64
import colorsys
import logging
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Optional, Tuple, Any, Dict
from pptx import Presentation
from PIL import Image, ImageDraw, ImageFont

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    HAS_BIDI = True
except Exception:
    HAS_BIDI = False

logger = logging.getLogger(__name__)

# DrawingML & OpenXML Namespaces
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def _is_rtl_text(text: str) -> bool:
    """Checks if text contains Arabic/Persian/Hebrew characters."""
    if not text: return False
    for ch in text:
        if '\u0600' <= ch <= '\u06FF' or '\u0750' <= ch <= '\u077F' or '\uFB50' <= ch <= '\uFDFF' or '\uFE70' <= ch <= '\uFEFF' or '\u0590' <= ch <= '\u05FF':
            return True
    return False

def _shape_text_for_display(text: str) -> str:
    """Correctly shapes Arabic/Persian letters with joining and RTL BiDi ordering."""
    if not text or not HAS_BIDI:
        return text
    if _is_rtl_text(text):
        try:
            reshaped = arabic_reshaper.reshape(text)
            return get_display(reshaped)
        except Exception:
            return text
    return text

def _extract_theme_color_palette(prs: Presentation) -> Dict[str, Tuple[int, int, int]]:
    """
    Extracts the color palette (dk1, lt1, accent1-6, etc.) from the theme1.xml part.
    """
    colors: Dict[str, Tuple[int, int, int]] = {
        "dk1": (0, 0, 0), "lt1": (255, 255, 255),
        "dk2": (80, 80, 70), "lt2": (238, 236, 225),
        "accent1": (232, 76, 34), "accent2": (255, 189, 71),
        "accent3": (182, 73, 38), "accent4": (255, 132, 39),
        "accent5": (204, 153, 0), "accent6": (178, 38, 0),
        "hlink": (204, 153, 0), "folHlink": (102, 102, 153),
        "bg1": (255, 255, 255), "bg2": (245, 245, 245),
        "tx1": (0, 0, 0), "tx2": (100, 100, 100)
    }

    try:
        for rel in prs.part.rels.values():
            if "theme" in rel.reltype:
                root = ET.fromstring(rel.target_part.blob)
                clrScheme = root.find(f".//{{{NS_A}}}clrScheme")
                if clrScheme is not None:
                    for elem in clrScheme:
                        tag = elem.tag.split("}")[-1]
                        srgb = elem.find(f"{{{NS_A}}}srgbClr")
                        sysClr = elem.find(f"{{{NS_A}}}sysClr")
                        hex_val = None
                        if srgb is not None and "val" in srgb.attrib:
                            hex_val = srgb.attrib["val"]
                        elif sysClr is not None:
                            hex_val = sysClr.attrib.get("lastClr", "000000")
                        
                        if hex_val:
                            try:
                                r = int(hex_val[0:2], 16)
                                g = int(hex_val[2:4], 16)
                                b = int(hex_val[4:6], 16)
                                colors[tag] = (r, g, b)
                                if tag == "lt1": colors["bg1"] = (r, g, b)
                                if tag == "lt2": colors["bg2"] = (r, g, b)
                                if tag == "dk1": colors["tx1"] = (r, g, b)
                                if tag == "dk2": colors["tx2"] = (r, g, b)
                            except Exception:
                                pass
    except Exception as e:
        logger.warning(f"Failed to extract theme: {e}")

    return colors

def _apply_lum_transforms(r: int, g: int, b: int, lum_mod: float = 1.0, lum_off: float = 0.0) -> Tuple[int, int, int]:
    """Applies DrawingML luminance modifications (lumMod and lumOff) using HSL conversion."""
    try:
        h, l, s = colorsys.rgb_to_hls(r / 255.0, g / 255.0, b / 255.0)
        l = l * lum_mod + lum_off
        l = max(0.0, min(1.0, l))
        r2, g2, b2 = colorsys.hls_to_rgb(h, l, s)
        return int(r2 * 255), int(g2 * 255), int(b2 * 255)
    except Exception:
        return r, g, b

def _resolve_element_color(
    elem: Optional[ET.Element],
    palette: Dict[str, Tuple[int, int, int]]
) -> Optional[Tuple[int, int, int, int]]:
    """
    Parses DrawingML color specs (srgbClr, schemeClr, lumMod, lumOff, alpha).
    Returns (R, G, B, Alpha 0-255).
    """
    if elem is None:
        return None

    srgb = elem.find(f".//{{{NS_A}}}srgbClr")
    scheme = elem.find(f".//{{{NS_A}}}schemeClr")

    base_rgb: Optional[Tuple[int, int, int]] = None
    color_node = None

    if srgb is not None and "val" in srgb.attrib:
        color_node = srgb
        hv = srgb.attrib["val"]
        try:
            base_rgb = (int(hv[0:2], 16), int(hv[2:4], 16), int(hv[4:6], 16))
        except Exception:
            pass
    elif scheme is not None and "val" in scheme.attrib:
        color_node = scheme
        scheme_val = scheme.attrib["val"]
        base_rgb = palette.get(scheme_val, (128, 128, 128))

    if not base_rgb:
        return None

    r, g, b = base_rgb
    alpha = 255

    # Check color transforms (lumMod, lumOff, alpha)
    if color_node is not None:
        lum_mod_node = color_node.find(f"{{{NS_A}}}lumMod")
        lum_off_node = color_node.find(f"{{{NS_A}}}lumOff")
        alpha_node = color_node.find(f"{{{NS_A}}}alpha")

        lum_mod_val = 1.0
        lum_off_val = 0.0

        if lum_mod_node is not None and "val" in lum_mod_node.attrib:
            lum_mod_val = int(lum_mod_node.attrib["val"]) / 100000.0

        if lum_off_node is not None and "val" in lum_off_node.attrib:
            lum_off_val = int(lum_off_node.attrib["val"]) / 100000.0

        if lum_mod_val != 1.0 or lum_off_val != 0.0:
            r, g, b = _apply_lum_transforms(r, g, b, lum_mod_val, lum_off_val)

        if alpha_node is not None and "val" in alpha_node.attrib:
            alpha = max(0, min(255, int(int(alpha_node.attrib["val"]) * 255 / 100000.0)))

    return (max(0, min(255, r)), max(0, min(255, g)), max(0, min(255, b)), alpha)

def _get_shape_colors_from_xml(
    sp_element: ET.Element,
    palette: Dict[str, Tuple[int, int, int]]
) -> Tuple[Optional[Tuple[int, int, int, int]], Optional[Tuple[int, int, int, int]], int]:
    """
    Extracts fill color (with alpha), line outline color, and outline width.
    """
    fill_rgba = None
    line_rgba = None
    line_w = 1

    spPr = sp_element.find(f"{{{NS_P}}}spPr")
    if spPr is None:
        spPr = sp_element.find(f"{{{NS_A}}}spPr")
    if spPr is None:
        spPr = sp_element

    # 1. Fill
    solid_fill = spPr.find(f"{{{NS_A}}}solidFill")
    no_fill = spPr.find(f"{{{NS_A}}}noFill")
    grad_fill = spPr.find(f"{{{NS_A}}}gradFill")

    if no_fill is not None:
        fill_rgba = None
    elif solid_fill is not None:
        fill_rgba = _resolve_element_color(solid_fill, palette)
    elif grad_fill is not None:
        gs_lst = grad_fill.find(f"{{{NS_A}}}gsLst")
        if gs_lst is not None:
            first_stop = gs_lst.find(f"{{{NS_A}}}gs")
            if first_stop is not None:
                fill_rgba = _resolve_element_color(first_stop, palette)

    # 2. Line Outline
    ln = spPr.find(f"{{{NS_A}}}ln")
    if ln is not None:
        if "w" in ln.attrib:
            try:
                line_w = max(1, int(int(ln.attrib["w"]) / 12700)) # roughly EMU to px
            except Exception:
                pass
        ln_solid = ln.find(f"{{{NS_A}}}solidFill")
        if ln_solid is not None:
            line_rgba = _resolve_element_color(ln_solid, palette)

    return fill_rgba, line_rgba, line_w

def _draw_cubic_bezier(p0, p1, p2, p3, steps=16) -> List[Tuple[float, float]]:
    """Calculates points along a cubic bezier segment."""
    points = []
    for i in range(steps + 1):
        t = i / steps
        u = 1.0 - t
        tt = t * t
        uu = u * u
        uuu = uu * u
        ttt = tt * t
        
        x = uuu * p0[0] + 3 * uu * t * p1[0] + 3 * u * tt * p2[0] + ttt * p3[0]
        y = uuu * p0[1] + 3 * uu * t * p1[1] + 3 * u * tt * p2[1] + ttt * p3[1]
        points.append((x, y))
    return points

def _render_custom_geom(
    draw: ImageDraw.Draw,
    custGeom_elem: ET.Element,
    shape_left_px: float,
    shape_top_px: float,
    shape_w_px: float,
    shape_h_px: float,
    fill_color: Optional[Tuple[int, int, int, int]],
    outline_color: Optional[Tuple[int, int, int, int]],
    line_width: int = 1
):
    """
    Renders custom geometry vector paths (moveTo, lnTo, cubicBezTo, close).
    """
    path_list = custGeom_elem.find(f"{{{NS_A}}}pathLst")
    if path_list is None:
        return

    for path_elem in path_list.findall(f"{{{NS_A}}}path"):
        pw = float(path_elem.attrib.get("w", shape_w_px or 1))
        ph = float(path_elem.attrib.get("h", shape_h_px or 1))

        if pw <= 0 or ph <= 0:
            continue

        sx = shape_w_px / pw
        sy = shape_h_px / ph

        polygons: List[List[Tuple[float, float]]] = []
        current_subpath: List[Tuple[float, float]] = []
        current_pt = (shape_left_px, shape_top_px)

        for cmd in path_elem:
            tag = cmd.tag.split("}")[-1]
            if tag == "moveTo":
                if current_subpath:
                    polygons.append(current_subpath)
                    current_subpath = []
                pt_node = cmd.find(f"{{{NS_A}}}pt")
                if pt_node is not None:
                    x = shape_left_px + float(pt_node.attrib["x"]) * sx
                    y = shape_top_px + float(pt_node.attrib["y"]) * sy
                    current_pt = (x, y)
                    current_subpath.append(current_pt)
            elif tag == "lnTo":
                pt_node = cmd.find(f"{{{NS_A}}}pt")
                if pt_node is not None:
                    x = shape_left_px + float(pt_node.attrib["x"]) * sx
                    y = shape_top_px + float(pt_node.attrib["y"]) * sy
                    current_pt = (x, y)
                    current_subpath.append(current_pt)
            elif tag == "cubicBezTo":
                pts = cmd.findall(f"{{{NS_A}}}pt")
                if len(pts) == 3:
                    p1 = (shape_left_px + float(pts[0].attrib["x"]) * sx, shape_top_px + float(pts[0].attrib["y"]) * sy)
                    p2 = (shape_left_px + float(pts[1].attrib["x"]) * sx, shape_top_px + float(pts[1].attrib["y"]) * sy)
                    p3 = (shape_left_px + float(pts[2].attrib["x"]) * sx, shape_top_px + float(pts[2].attrib["y"]) * sy)
                    curve_pts = _draw_cubic_bezier(current_pt, p1, p2, p3)
                    current_subpath.extend(curve_pts)
                    current_pt = p3
            elif tag == "close":
                if current_subpath:
                    current_subpath.append(current_subpath[0])
                    polygons.append(current_subpath)
                    current_subpath = []

        if current_subpath:
            polygons.append(current_subpath)
            
        for poly in polygons:
            if len(poly) >= 3 and fill_color:
                draw.polygon(poly, fill=fill_color[:3])
            if len(poly) >= 2 and outline_color:
                draw.line(poly, fill=outline_color[:3], width=max(1, line_width))

def _get_slide_bg_color(slide, palette: Dict[str, Tuple[int, int, int]]) -> Tuple[int, int, int]:
    """Determines the background color of the slide, falling back to layout and master."""
    bg_color = palette.get("bg1", (255, 255, 255))
    
    targets = [slide, getattr(slide, "slide_layout", None), getattr(slide, "slide_master", None)]
    
    for target in targets:
        if not target:
            continue
            
        # Check python-pptx API
        bg = getattr(target, "background", None)
        if bg and bg.fill and bg.fill.type is not None:
            try:
                if bg.fill.type == 1 and bg.fill.fore_color and getattr(bg.fill.fore_color, "rgb", None):
                    rgb = bg.fill.fore_color.rgb
                    return (rgb[0], rgb[1], rgb[2])
            except Exception:
                pass
                
        # Check XML structure as a more robust fallback
        bg_elem = target.element.find(f".//{{{NS_P}}}bgPr/{{{NS_A}}}solidFill")
        if bg_elem is not None:
            c = _resolve_element_color(bg_elem, palette)
            if c:
                return c[:3]
                
    return bg_color

def _render_shape_recursive(
    shape: Any,
    img: Image.Image,
    draw: ImageDraw.Draw,
    scale_x: float,
    scale_y: float,
    palette: Dict[str, Tuple[int, int, int]],
    group_offset: Optional[Tuple[float, float, float, float]] = None
):
    try:
        # Check if shape is a GroupShape (has .shapes)
        if hasattr(shape, "shapes"):
            grp_elem = shape.element
            xfrm = grp_elem.find(f"{{{NS_P}}}grpSpPr/{{{NS_A}}}xfrm")
            if xfrm is None:
                xfrm = grp_elem.find(f"{{{NS_A}}}xfrm")

            raw_gx = float(getattr(shape, "left", 0) or 0)
            raw_gy = float(getattr(shape, "top", 0) or 0)
            raw_gw = float(getattr(shape, "width", 1) or 1)
            raw_gh = float(getattr(shape, "height", 1) or 1)
            
            cx, cy, cw, ch = 0, 0, 1, 1
            if xfrm is not None:
                off = xfrm.find(f"{{{NS_A}}}off")
                ext = xfrm.find(f"{{{NS_A}}}ext")
                chOff = xfrm.find(f"{{{NS_A}}}chOff")
                chExt = xfrm.find(f"{{{NS_A}}}chExt")
                
                if off is not None and "x" in off.attrib: raw_gx = float(off.attrib["x"])
                if off is not None and "y" in off.attrib: raw_gy = float(off.attrib["y"])
                if ext is not None and "cx" in ext.attrib: raw_gw = float(ext.attrib["cx"])
                if ext is not None and "cy" in ext.attrib: raw_gh = float(ext.attrib["cy"])
                
                if chOff is not None and "x" in chOff.attrib: cx = float(chOff.attrib["x"])
                if chOff is not None and "y" in chOff.attrib: cy = float(chOff.attrib["y"])
                if chExt is not None and "cx" in chExt.attrib: cw = float(chExt.attrib["cx"])
                if chExt is not None and "cy" in chExt.attrib: ch = float(chExt.attrib["cy"])

            fx = raw_gw / cw if cw > 0 else 1.0
            fy = raw_gh / ch if ch > 0 else 1.0

            local_off_x = raw_gx - cx * fx
            local_off_y = raw_gy - cy * fy
            local_fx = fx
            local_fy = fy

            # Compose with parent group_offset
            if group_offset:
                p_go_x, p_go_y, p_gfx, p_gfy = group_offset
                new_go_x = p_go_x + local_off_x * p_gfx
                new_go_y = p_go_y + local_off_y * p_gfy
                new_fx = p_gfx * local_fx
                new_fy = p_gfy * local_fy
                next_offset = (new_go_x, new_go_y, new_fx, new_fy)
            else:
                next_offset = (local_off_x, local_off_y, local_fx, local_fy)

            for child in shape.shapes:
                _render_shape_recursive(child, img, draw, scale_x, scale_y, palette, next_offset)
            return

        # Calculate actual coordinates including group transforms
        raw_left = float(getattr(shape, "left", 0) or 0)
        raw_top = float(getattr(shape, "top", 0) or 0)
        raw_w = float(getattr(shape, "width", 0) or 0)
        raw_h = float(getattr(shape, "height", 0) or 0)

        if group_offset:
            go_x, go_y, gfx, gfy = group_offset
            calc_left = go_x + raw_left * gfx
            calc_top = go_y + raw_top * gfy
            calc_w = raw_w * gfx
            calc_h = raw_h * gfy
        else:
            calc_left = raw_left
            calc_top = raw_top
            calc_w = raw_w
            calc_h = raw_h

        left_px = int(calc_left * scale_x)
        top_px = int(calc_top * scale_y)
        width_px = max(1, int(calc_w * scale_x))
        height_px = max(1, int(calc_h * scale_y))
        right_px = left_px + width_px
        bottom_px = top_px + height_px

        # Extract Fill & Line
        fill_rgba, line_rgba, line_w = _get_shape_colors_from_xml(shape.element, palette)

        # Fallback to python-pptx standard properties if available
        if fill_rgba is None and hasattr(shape, "fill") and shape.fill:
            try:
                if shape.fill.type == 1 and getattr(shape.fill.fore_color, "rgb", None):
                    rgb = shape.fill.fore_color.rgb
                    fill_rgba = (rgb[0], rgb[1], rgb[2], 255)
            except Exception:
                pass

        if line_rgba is None and hasattr(shape, "line") and shape.line:
            try:
                if shape.line.fill and shape.line.fill.type is not None and getattr(shape.line.color, "rgb", None):
                    rgb = shape.line.color.rgb
                    line_rgba = (rgb[0], rgb[1], rgb[2], 255)
            except Exception:
                pass

        # 1. Custom Geometry (Vectors & Curves)
        custGeom = shape.element.find(f".//{{{NS_A}}}custGeom")
        if custGeom is not None:
            _render_custom_geom(draw, custGeom, left_px, top_px, width_px, height_px, fill_rgba, line_rgba, line_w)

        # 2. Standard Shapes (Rectangles, Ovals, Rounded Rectangles)
        else:
            shape_type_str = str(getattr(shape, "shape_type", ""))
            shape_name = shape.name.lower() if shape.name else ""

            if fill_rgba or line_rgba:
                fill_rgb = fill_rgba[:3] if fill_rgba else None
                outline_rgb = line_rgba[:3] if line_rgba else None

                if "oval" in shape_name or "circle" in shape_name:
                    draw.ellipse([(left_px, top_px), (right_px, bottom_px)], fill=fill_rgb, outline=outline_rgb, width=line_w)
                elif "rounded" in shape_name:
                    rad = max(4, int(min(width_px, height_px) * 0.15))
                    draw.rounded_rectangle([(left_px, top_px), (right_px, bottom_px)], radius=rad, fill=fill_rgb, outline=outline_rgb, width=line_w)
                elif "line" in shape_type_str or "connector" in shape_name:
                    draw.line([(left_px, top_px), (right_px, bottom_px)], fill=outline_rgb or fill_rgb or (100, 100, 100), width=max(1, line_w))
                else:
                    draw.rectangle([(left_px, top_px), (right_px, bottom_px)], fill=fill_rgb, outline=outline_rgb, width=line_w)

        # 3. Picture Rendering
        if hasattr(shape, "shape_type") and str(shape.shape_type).endswith("PICTURE"):
            try:
                img_bytes = io.BytesIO(shape.image.blob)
                pic = Image.open(img_bytes).convert("RGBA")
                pic = pic.resize((width_px, height_px), Image.Resampling.LANCZOS)
                img.paste(pic, (left_px, top_px), pic)
            except Exception:
                pass

        # 4. Chart Placeholder
        if getattr(shape, "has_chart", False):
            try:
                draw.rectangle([(left_px, top_px), (right_px, bottom_px)], fill=(235, 235, 235), outline=(180, 180, 180), width=2)
                try:
                    chart_font = ImageFont.truetype("tahoma.ttf", 16)
                except Exception:
                    chart_font = ImageFont.load_default()
                draw.text((left_px + 10, top_px + 10), "[Chart]", fill=(80, 80, 80), font=chart_font)
            except Exception:
                pass

        # 5. Table Rendering
        if getattr(shape, "has_table", False):
            table = shape.table
            rows = len(table.rows)
            cols = len(table.columns)
            if rows > 0 and cols > 0:
                row_h = height_px / rows
                col_w = width_px / cols
                for r_idx, row in enumerate(table.rows):
                    for c_idx, cell in enumerate(row.cells):
                        c_x1 = int(left_px + c_idx * col_w)
                        c_y1 = int(top_px + r_idx * row_h)
                        c_x2 = int(c_x1 + col_w)
                        c_y2 = int(c_y1 + row_h)

                        # Extract cell fill
                        c_fill = None
                        try:
                            c_rgba, _, _ = _get_shape_colors_from_xml(cell._tc, palette)
                            if c_rgba: c_fill = c_rgba[:3]
                        except Exception:
                            pass
                            
                        if c_fill is None:
                            c_fill = palette.get("accent1", (230, 80, 40)) if r_idx == 0 else (255, 255, 255)

                        draw.rectangle([(c_x1, c_y1), (c_x2, c_y2)], fill=c_fill, outline=(200, 200, 200))

                        cell_text = cell.text.strip()
                        if cell_text:
                            text_color = (255, 255, 255) if r_idx == 0 else (30, 40, 50)
                            try:
                                font = ImageFont.truetype("tahoma.ttf", max(8, int(11 * scale_y)))
                            except Exception:
                                font = ImageFont.load_default()
                            shaped_cell = _shape_text_for_display(cell_text[:30])
                            draw.text((c_x1 + 6, c_y1 + 4), shaped_cell, fill=text_color, font=font)

        # 6. Text Frame Rendering
        if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
            tf = shape.text_frame
            curr_y = top_px + 4

            for p in tf.paragraphs:
                txt = p.text.strip()
                if not txt:
                    continue

                pt_size = 14
                font_color = palette.get("tx1", (20, 20, 20))
                is_bold = False

                for r in p.runs:
                    if r.font:
                        if r.font.size: pt_size = r.font.size.pt
                        if r.font.bold: is_bold = True
                        
                        # Check text color via XML for theme support
                        rPr = r._r.find(f"{{{NS_A}}}rPr")
                        if rPr is not None:
                            c = _resolve_element_color(rPr, palette)
                            if c: font_color = c[:3]
                    break

                px_font_size = max(8, int(pt_size * scale_y * 1.05))
                font_name = "tahomabd.ttf" if is_bold else "tahoma.ttf"
                try:
                    font = ImageFont.truetype(font_name, px_font_size)
                except Exception:
                    try:
                        font = ImageFont.truetype("segoeui.ttf", px_font_size)
                    except Exception:
                        font = ImageFont.load_default()

                is_rtl = _is_rtl_text(txt)
                
                is_centered = False
                is_right = False
                align = p.alignment
                if align is not None:
                    align_name = align.name if hasattr(align, "name") else str(align)
                    if "CENTER" in align_name: is_centered = True
                    elif "RIGHT" in align_name: is_right = True

                lines_to_draw = txt.split("\n") if "\n" in txt else [txt]
                for raw_line in lines_to_draw:
                    words = raw_line.split()
                    curr_line = ""
                    for w in words:
                        test_line = f"{curr_line} {w}".strip()
                        bbox = draw.textbbox((0, 0), test_line, font=font)
                        w_w = bbox[2] - bbox[0]
                        if w_w > (width_px - 12) and curr_line:
                            shaped_display = _shape_text_for_display(curr_line)
                            line_bbox = draw.textbbox((0, 0), shaped_display, font=font)
                            line_w_px = line_bbox[2] - line_bbox[0]
                            
                            if is_centered:
                                draw_x = left_px + (width_px - line_w_px) // 2
                            elif is_right or is_rtl:
                                draw_x = right_px - line_w_px - 8
                            else:
                                draw_x = left_px + 8

                            draw.text((draw_x, curr_y), shaped_display, fill=font_color, font=font)
                            curr_y += int(px_font_size * 1.35)
                            curr_line = w
                        else:
                            curr_line = test_line

                    if curr_line:
                        shaped_display = _shape_text_for_display(curr_line)
                        line_bbox = draw.textbbox((0, 0), shaped_display, font=font)
                        line_w_px = line_bbox[2] - line_bbox[0]

                        if is_centered:
                            draw_x = left_px + (width_px - line_w_px) // 2
                        elif is_right or is_rtl:
                            draw_x = right_px - line_w_px - 8
                        else:
                            draw_x = left_px + 8

                        draw.text((draw_x, curr_y), shaped_display, fill=font_color, font=font)
                        curr_y += int(px_font_size * 1.35)

    except Exception as e:
        logger.debug(f"Error rendering shape: {e}")

def render_pptx_slide_to_image(
    slide: Any,
    slide_width_emu: int,
    slide_height_emu: int,
    target_width_px: int = 850,
    palette: Optional[Dict[str, Tuple[int, int, int]]] = None
) -> Image.Image:
    aspect_ratio = slide_height_emu / slide_width_emu if slide_width_emu else (9 / 16)
    target_height_px = int(target_width_px * aspect_ratio)

    scale_x = target_width_px / slide_width_emu if slide_width_emu else 1.0
    scale_y = target_height_px / slide_height_emu if slide_height_emu else 1.0

    theme_palette = palette or {}
    bg_color = _get_slide_bg_color(slide, theme_palette)
    
    # Initialize as RGBA to safely support transparent image pasting
    img = Image.new("RGBA", (target_width_px, target_height_px), color=bg_color + (255,))
    draw = ImageDraw.Draw(img)

    try:
        if hasattr(slide, "slide_layout") and slide.slide_layout:
            for layout_shape in slide.slide_layout.shapes:
                if not getattr(layout_shape, "is_placeholder", False):
                    _render_shape_recursive(layout_shape, img, draw, scale_x, scale_y, theme_palette)
    except Exception:
        pass

    for shape in slide.shapes:
        _render_shape_recursive(shape, img, draw, scale_x, scale_y, theme_palette)

    # Convert final image to RGB for standard output
    return img.convert("RGB")

def render_pptx_file_previews(pptx_path: Path | str, target_width_px: int = 850) -> List[Image.Image]:
    prs = Presentation(str(pptx_path))
    palette = _extract_theme_color_palette(prs)
    slide_w = prs.slide_width
    slide_h = prs.slide_height
    images = []

    for slide in prs.slides:
        img = render_pptx_slide_to_image(
            slide,
            slide_w,
            slide_h,
            target_width_px=target_width_px,
            palette=palette
        )
        images.append(img)

    return images

def image_to_base64_jpeg(img: Image.Image, quality: int = 85) -> str:
    buffered = io.BytesIO()
    rgb_img = img.convert("RGB")
    rgb_img.save(buffered, format="JPEG", quality=quality)
    img_b64 = base64.b64encode(buffered.getvalue()).decode("utf-8")
    return f"data:image/jpeg;base64,{img_b64}"