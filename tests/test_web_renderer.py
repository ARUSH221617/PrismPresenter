import pytest
from pathlib import Path
from pptx import Presentation
from pptx_jahat.config import DATA_DIR
from pptx_jahat.tools.renderers.web_renderer import (
    PPTXWebRenderer,
    render_slide_to_html,
    render_pptx_to_html_deck
)
from pptx_jahat.tools.preview import render_pptx

def test_web_renderer_single_slide():
    tpl_path = list(DATA_DIR.glob("*.pptx"))[0]
    prs = Presentation(str(tpl_path))
    slide = prs.slides[0]

    html_out = render_slide_to_html(slide, prs, width=800)
    assert isinstance(html_out, str)
    assert '<div class="pptx-web-slide"' in html_out
    assert "position:relative;" in html_out
    assert "width:800px;" in html_out

def test_web_renderer_full_deck():
    tpl_path = list(DATA_DIR.glob("*.pptx"))[0]
    deck_html = render_pptx_to_html_deck(str(tpl_path), width=800)
    assert isinstance(deck_html, list)
    assert len(deck_html) > 0
    for slide_html in deck_html:
        assert '<div class="pptx-web-slide"' in slide_html

def test_cascade_engine_info():
    tpl_path = list(DATA_DIR.glob("*.pptx"))[0]
    res = render_pptx(str(tpl_path), width=400, return_engine_info=True)
    assert isinstance(res, tuple)
    imgs, engine_name = res
    assert len(imgs) > 0
    assert engine_name in ("Native PowerPoint", "Web Render Engine", "Pure PIL")
