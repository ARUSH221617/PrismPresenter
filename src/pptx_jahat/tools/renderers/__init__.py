"""
pptx_jahat.tools.renderers
===========================
Specialized DrawingML / PresentationML rendering sub-engines:
- color_resolver: DrawingML color models, HSL transforms, theme palettes, format schemes.
- geometry_engine: 180+ OpenXML preset shapes, Bezier curves, connectors, 3D bevels.
- gradient_engine: Linear, radial, rectangular, shape-contour path gradients, specular lighting.
- typography_engine: FontResolver, multi-column text, WordArt, OMML Math formulas, RTL.
- chart_renderer: OpenXML 2D/3D & modern charts (c16) with matplotlib/PIL rasterization.
- smartart_renderer: SmartArt relational diagrams (hierarchies, chevrons, cycles, matrices).
- table_renderer: Advanced table grids with cell spanning (hMerge/vMerge), diagonal borders.
- media_resolver: SVG vector rasterization, EMF/WMF metafiles, video poster frames.
"""

from pptx_jahat.tools.renderers.color_resolver import (
    Theme,
    ColorResolver,
    resolve_element_color,
    resolve_theme_fill,
    parse_color_elem,
    scheme_rgb,
    hex_to_rgb,
    clamp,
    RGBA,
    RGB,
)
from pptx_jahat.tools.renderers.geometry_engine import (
    GeometryEngine,
    get_preset_ops,
    get_connector_ops,
    render_custom_geom,
    rotate_ops,
    stroke_ops,
    fill_ops,
    make_silhouette,
)
from pptx_jahat.tools.renderers.gradient_engine import (
    GradientEngine,
    paint_gradient,
    paint_3d_bevel,
    paint_shadow,
    paint_glow,
    paint_inner_shadow,
    paint_soft_edge,
    paint_reflection,
    parse_gradient,
    parse_effect_lst,
)
from pptx_jahat.tools.renderers.typography_engine import (
    FontResolver,
    TypographyEngine,
    render_txbody,
    render_math_formula,
    shape_text_for_display,
    is_rtl_text,
)
from pptx_jahat.tools.renderers.chart_renderer import (
    ChartRenderer,
    render_chart_part,
)
from pptx_jahat.tools.renderers.smartart_renderer import (
    SmartArtRenderer,
    render_smartart_part,
)
from pptx_jahat.tools.renderers.table_renderer import (
    TableRenderer,
    render_table,
)
from pptx_jahat.tools.renderers.media_resolver import (
    MediaResolver,
    rasterize_svg,
    rasterize_emf_wmf,
    paint_media_poster,
)

__all__ = [
    "Theme",
    "ColorResolver",
    "resolve_element_color",
    "resolve_theme_fill",
    "parse_color_elem",
    "scheme_rgb",
    "hex_to_rgb",
    "clamp",
    "RGBA",
    "RGB",
    "GeometryEngine",
    "get_preset_ops",
    "get_connector_ops",
    "render_custom_geom",
    "rotate_ops",
    "stroke_ops",
    "fill_ops",
    "make_silhouette",
    "GradientEngine",
    "paint_gradient",
    "paint_3d_bevel",
    "paint_shadow",
    "paint_glow",
    "paint_inner_shadow",
    "paint_soft_edge",
    "paint_reflection",
    "parse_gradient",
    "parse_effect_lst",
    "FontResolver",
    "TypographyEngine",
    "render_txbody",
    "render_math_formula",
    "shape_text_for_display",
    "is_rtl_text",
    "ChartRenderer",
    "render_chart_part",
    "SmartArtRenderer",
    "render_smartart_part",
    "TableRenderer",
    "render_table",
    "MediaResolver",
    "rasterize_svg",
    "rasterize_emf_wmf",
    "paint_media_poster",
]
