import os
import io
import math
import base64
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Optional, Tuple, Any, Dict
from pptx import Presentation
from PIL import Image, ImageDraw, ImageFont, ImageColor

# DrawingML & OpenXML Namespaces
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def _extract_theme_color_palette(prs: Presentation) -> Dict[str, Tuple[int, int, int]]:
    """
    Extracts the color palette (dk1, lt1, accent1-6, etc.) from the theme1.xml part.
    """
    colors: Dict[str, Tuple[int, int, int]] = {
        "dk1": (0, 0, 0),
        "lt1": (255, 255, 255),
        "dk2": (80, 80, 70),
        "lt2": (238, 236, 225),
        "accent1": (232, 76, 34),
        "accent2": (255, 189, 71),
        "accent3": (182, 73, 38),
        "accent4": (255, 132, 39),
        "accent5": (204, 153, 0),
        "accent6": (178, 38, 0),
        "hlink": (204, 153, 0),
        "folHlink": (102, 102, 153),
        "bg1": (255, 255, 255),
        "bg2": (245, 245, 245),
        "tx1": (0, 0, 0),
        "tx2": (100, 100, 100)
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
    except Exception:
        pass

    return colors


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
        lum_mod = color_node.find(f"{{{NS_A}}}lumMod")
        lum_off = color_node.find(f"{{{NS_A}}}lumOff")
        alpha_node = color_node.find(f"{{{NS_A}}}alpha")

        if lum_mod is not None and "val" in lum_mod.attrib:
            mod_val = int(lum_mod.attrib["val"]) / 100000.0
            r = int(r * mod_val)
            g = int(g * mod_val)
            b = int(b * mod_val)

        if lum_off is not None and "val" in lum_off.attrib:
            off_val = int(lum_off.attrib["val"]) / 100000.0
            r = min(255, int(r + 255 * off_val))
            g = min(255, int(g + 255 * off_val))
            b = min(255, int(b + 255 * off_val))

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

    # Check solidFill, gradFill, noFill in spPr
    spPr = sp_element.find(f"{{{NS_P}}}spPr")
    if spPr is None:
        spPr = sp_element.find(f"{{{NS_A}}}spPr")
    if spPr is None:
        # direct search
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
        # Resolve first stop color of gradient
        first_stop = grad_fill.find(f".//{{{NS_A}}}gs")
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

        polygon_points: List[Tuple[float, float]] = []
        current_pt = (shape_left_px, shape_top_px)

        for cmd in path_elem:
            tag = cmd.tag.split("}")[-1]
            if tag == "moveTo":
                pt_node = cmd.find(f"{{{NS_A}}}pt")
                if pt_node is not None:
                    x = shape_left_px + float(pt_node.attrib["x"]) * sx
                    y = shape_top_px + float(pt_node.attrib["y"]) * sy
                    current_pt = (x, y)
                    if not polygon_points:
                        polygon_points.append(current_pt)
            elif tag == "lnTo":
                pt_node = cmd.find(f"{{{NS_A}}}pt")
                if pt_node is not None:
                    x = shape_left_px + float(pt_node.attrib["x"]) * sx
                    y = shape_top_px + float(pt_node.attrib["y"]) * sy
                    current_pt = (x, y)
                    polygon_points.append(current_pt)
            elif tag == "cubicBezTo":
                pts = cmd.findall(f"{{{NS_A}}}pt")
                if len(pts) == 3:
                    p1 = (shape_left_px + float(pts[0].attrib["x"]) * sx, shape_top_px + float(pts[0].attrib["y"]) * sy)
                    p2 = (shape_left_px + float(pts[1].attrib["x"]) * sx, shape_top_px + float(pts[1].attrib["y"]) * sy)
                    p3 = (shape_left_px + float(pts[2].attrib["x"]) * sx, shape_top_px + float(pts[2].attrib["y"]) * sy)
                    curve_pts = _draw_cubic_bezier(current_pt, p1, p2, p3)
                    polygon_points.extend(curve_pts)
                    current_pt = p3
            elif tag == "close":
                if polygon_points:
                    polygon_points.append(polygon_points[0])

        if len(polygon_points) >= 3 and fill_color:
            draw.polygon(polygon_points, fill=fill_color[:3])
        if len(polygon_points) >= 2 and outline_color:
            draw.line(polygon_points, fill=outline_color[:3], width=max(1, line_width))


def _render_shape_recursive(
    shape: Any,
    img: Image.Image,
    draw: ImageDraw.Draw,
    scale_x: float,
    scale_y: float,
    palette: Dict[str, Tuple[int, int, int]]
):
    """
    Renders shapes, nested groups, tables, images, vector paths, and text with accurate coordinates.
    """
    try:
        # If shape is a GroupShape, iterate its children
        if hasattr(shape, "shapes"):
            for child in shape.shapes:
                _render_shape_recursive(child, img, draw, scale_x, scale_y, palette)
            return

        left_px = int(shape.left * scale_x)
        top_px = int(shape.top * scale_y)
        width_px = max(1, int(shape.width * scale_x))
        height_px = max(1, int(shape.height * scale_y))
        right_px = left_px + width_px
        bottom_px = top_px + height_px

        # Extract Fill & Line
        fill_rgba, line_rgba, line_w = _get_shape_colors_from_xml(shape.element, palette)

        # Fallback to python-pptx standard properties if available
        if fill_rgba is None and hasattr(shape, "fill") and shape.fill:
            try:
                if shape.fill.type == 1 and shape.fill.fore_color.rgb:
                    rgb = shape.fill.fore_color.rgb
                    fill_rgba = (rgb[0], rgb[1], rgb[2], 255)
            except Exception:
                pass

        if line_rgba is None and hasattr(shape, "line") and shape.line:
            try:
                if shape.line.fill.type is not None and shape.line.color.rgb:
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
            shape_name = shape.name.lower()

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

        # 4. Table Rendering
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

                        c_fill = palette.get("accent1", (230, 80, 40)) if r_idx == 0 else (255, 255, 255)
                        draw.rectangle([(c_x1, c_y1), (c_x2, c_y2)], fill=c_fill, outline=(200, 200, 200))

                        cell_text = cell.text.strip()
                        if cell_text:
                            text_color = (255, 255, 255) if r_idx == 0 else (30, 40, 50)
                            try:
                                font = ImageFont.truetype("arial.ttf", max(8, int(11 * scale_y)))
                            except Exception:
                                font = ImageFont.load_default()
                            draw.text((c_x1 + 6, c_y1 + 4), cell_text[:30], fill=text_color, font=font)

        # 5. Text Frame Rendering
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

                # Extract font specs from paragraph or first run
                for r in p.runs:
                    if r.font:
                        if r.font.size: pt_size = r.font.size.pt
                        if r.font.bold: is_bold = True
                        if r.font.color and hasattr(r.font.color, "rgb") and r.font.color.rgb:
                            font_color = (r.font.color.rgb[0], r.font.color.rgb[1], r.font.color.rgb[2])
                    break

                px_font_size = max(8, int(pt_size * scale_y * 1.05))
                font_name = "arialbd.ttf" if is_bold else "arial.ttf"
                try:
                    font = ImageFont.truetype(font_name, px_font_size)
                except Exception:
                    try:
                        font = ImageFont.truetype("segoeui.ttf", px_font_size)
                    except Exception:
                        font = ImageFont.load_default()

                # Alignments & wrapping
                words = txt.split()
                line = ""
                for w in words:
                    test_line = f"{line} {w}".strip()
                    bbox = draw.textbbox((0, 0), test_line, font=font)
                    w_w = bbox[2] - bbox[0]
                    if w_w > (width_px - 8) and line:
                        draw.text((left_px + 6, curr_y), line, fill=font_color, font=font)
                        curr_y += int(px_font_size * 1.35)
                        line = w
                    else:
                        line = test_line

                if line:
                    draw.text((left_px + 6, curr_y), line, fill=font_color, font=font)
                    curr_y += int(px_font_size * 1.35)

    except Exception:
        pass


def render_pptx_slide_to_image(
    slide: Any,
    slide_width_emu: int,
    slide_height_emu: int,
    target_width_px: int = 850,
    palette: Optional[Dict[str, Tuple[int, int, int]]] = None
) -> Image.Image:
    """
    Renders a comprehensive high-fidelity 2D visual preview of a PPTX slide with shapes,
    vectors, colors, containers, tables, pictures, and formatted text.
    """
    aspect_ratio = slide_height_emu / slide_width_emu if slide_width_emu else (9 / 16)
    target_height_px = int(target_width_px * aspect_ratio)

    scale_x = target_width_px / slide_width_emu if slide_width_emu else 1.0
    scale_y = target_height_px / slide_height_emu if slide_height_emu else 1.0

    theme_palette = palette or {}

    # 1. Slide Background Color
    bg_color = theme_palette.get("bg1", (255, 255, 255))
    try:
        if hasattr(slide, "background") and slide.background and slide.background.fill:
            if slide.background.fill.type == 1 and slide.background.fill.fore_color.rgb:
                rgb = slide.background.fill.fore_color.rgb
                bg_color = (rgb[0], rgb[1], rgb[2])
    except Exception:
        pass

    img = Image.new("RGB", (target_width_px, target_height_px), color=bg_color)
    draw = ImageDraw.Draw(img)

    # 2. Render Slide Layout Shapes (Base Template Layout)
    try:
        if hasattr(slide, "slide_layout") and slide.slide_layout:
            for layout_shape in slide.slide_layout.shapes:
                # Do not render title/body placeholders from layout if the slide itself overrides them
                if not getattr(layout_shape, "is_placeholder", False):
                    _render_shape_recursive(layout_shape, img, draw, scale_x, scale_y, theme_palette)
    except Exception:
        pass

    # 3. Render Slide Shapes
    for shape in slide.shapes:
        _render_shape_recursive(shape, img, draw, scale_x, scale_y, theme_palette)

    return img


def render_pptx_file_previews(pptx_path: Path | str, target_width_px: int = 850) -> List[Image.Image]:
    """
    Renders all slides in a PPTX file into a list of PIL Images using the extracted theme palette.
    """
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
    """
    Converts a PIL Image to a base64 encoded JPEG data URL string.
    """
    buffered = io.BytesIO()
    rgb_img = img.convert("RGB")
    rgb_img.save(buffered, format="JPEG", quality=quality)
    img_b64 = base64.b64encode(buffered.getvalue()).decode("utf-8")
    return f"data:image/jpeg;base64,{img_b64}"
