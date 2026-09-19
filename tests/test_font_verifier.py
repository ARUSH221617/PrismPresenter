import pytest
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx_jahat.tools.font_verifier import (
    get_available_system_fonts,
    is_font_available,
    recommend_font_fallback,
    verify_fonts_inventory,
    apply_font_fallbacks_to_presentation,
    is_persian_or_arabic_font,
    normalize_font_name
)


def test_get_available_system_fonts():
    catalog = get_available_system_fonts()
    assert "families" in catalog
    assert "persian_fonts" in catalog
    assert "latin_fonts" in catalog
    assert len(catalog["families"]) > 0
    # Segoe UI or Arial should be present on almost any Windows/mac/linux test env
    families_lower = [f.lower() for f in catalog["families"]]
    assert any("segoe" in f or "arial" in f or "calibri" in f or "dejavu" in f for f in families_lower)


def test_persian_arabic_classification():
    assert is_persian_or_arabic_font("IRANYekanXFaNum") is True
    assert is_persian_or_arabic_font("Amuzeh-New-Bold") is True
    assert is_persian_or_arabic_font("B Nazanin") is True
    assert is_persian_or_arabic_font("Vazirmatn") is True
    assert is_persian_or_arabic_font("Segoe UI") is False
    assert is_persian_or_arabic_font("Arial") is False
    assert is_persian_or_arabic_font("Montserrat") is False


def test_normalize_font_name():
    assert normalize_font_name("Amuzeh-New-Bold") == "amuzehnew"
    assert normalize_font_name("Segoe UI Bold") == "segoeui"
    assert normalize_font_name("IRANYekanX-Regular.ttf") == "iranyekanx"


def test_is_font_available():
    catalog = get_available_system_fonts()
    # Office theme placeholder must always be treated as available
    assert is_font_available("+mj-lt", catalog) is True
    assert is_font_available("+mn-lt", catalog) is True
    assert is_font_available("", catalog) is True

    # Check a completely fictitious font name
    assert is_font_available("CompletelyFictitiousFontName12345", catalog) is False


def test_recommend_font_fallback():
    catalog = get_available_system_fonts()
    persian_fb = recommend_font_fallback("Amuzeh-New-Bold", catalog)
    assert isinstance(persian_fb, str)
    assert len(persian_fb) > 0

    latin_fb = recommend_font_fallback("NonExistentFuturisticSans", catalog)
    assert isinstance(latin_fb, str)
    assert len(latin_fb) > 0


def test_verify_fonts_inventory():
    fonts = ["+mj-lt", "CompletelyFakeFontXYZ999"]
    res = verify_fonts_inventory(fonts)
    assert res["all_installed"] is False
    assert "CompletelyFakeFontXYZ999" in res["missing_fonts"]
    assert "CompletelyFakeFontXYZ999" in res["recommendations"]
    assert "system_fonts" in res


def test_apply_font_fallbacks_to_presentation(tmp_path):
    prs = Presentation()
    blank_slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(blank_slide_layout)

    # Add text box with a fake font
    txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "Testing font replacement"
    r.font.name = "FakeOldFont"

    # Add table with fake font
    table_shape = slide.shapes.add_table(1, 1, Inches(1), Inches(3), Inches(2), Inches(1))
    cell = table_shape.table.cell(0, 0)
    cell.text = "Table text"
    cell_p = cell.text_frame.paragraphs[0]
    cell_r = cell_p.runs[0]
    cell_r.font.name = "FakeOldFont"

    fallbacks = {"FakeOldFont": "IRANYekanXFaNum"}
    updated = apply_font_fallbacks_to_presentation(prs, fallbacks)
    assert updated >= 2

    # Check that font was replaced
    assert r.font.name == "IRANYekanXFaNum"
    assert cell_r.font.name == "IRANYekanXFaNum"

    # Save and reload
    out_file = tmp_path / "test_fonts_out.pptx"
    prs.save(str(out_file))

    prs2 = Presentation(str(out_file))
    r_reloaded = prs2.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
    assert r_reloaded.font.name == "IRANYekanXFaNum"


def test_windows_and_user_fonts_detection():
    catalog = get_available_system_fonts(force_rescan=True)
    # Windows environment should detect thousands of fonts from GDI, Registry, and user folders
    assert catalog["total_count"] > 100
    assert "font_paths" in catalog

    # Montserrat, Segoe UI, Calibri, and standard Persian fonts should be available
    families_lower = {f.lower() for f in catalog["families"]}
    assert any("segoe ui" in f for f in families_lower)
    assert any("calibri" in f for f in families_lower)
    assert is_font_available("Segoe UI", catalog) is True
    assert is_font_available("Calibri", catalog) is True


def test_distributor_tag_and_variant_normalization():
    catalog = get_available_system_fonts()
    # Normalized name strips pack tags and weights
    assert normalize_font_name("IRANYekan  [ @mimvid ] Bold") == "iranyekan"
    assert normalize_font_name("110_Besmellah_1(MRT).ttf") == "110besmellah1"
    assert normalize_font_name("Sahel Black FD") == "sahel"
    assert normalize_font_name("IRANSansXFaNum DemiBold") == "iransansx"
    assert normalize_font_name("IRANSansXV") == "iransansx"


def test_deep_template_font_extraction(tmp_path):
    from pptx_jahat.tools.pptx_engine import extract_template_fonts
    from pptx.oxml import parse_xml

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # 1. Textbox with DrawingML a:cs complex script font
    tx = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
    p = tx.text_frame.paragraphs[0]
    r = p.add_run()
    r.text = "Persian Complex Script"
    cs_elem = parse_xml('<a:cs xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" typeface="DeepPersianFont"/>')
    r._r.get_or_add_rPr().append(cs_elem)

    # 2. Table with unique font
    tbl = slide.shapes.add_table(1, 1, Inches(1), Inches(3), Inches(2), Inches(1))
    cell = tbl.table.cell(0, 0)
    cell.text = "Table text"
    cell_p = cell.text_frame.paragraphs[0]
    cell_r = cell_p.runs[0]
    cell_r.font.name = "TableUniqueFont"

    fonts = extract_template_fonts(prs)
    assert "DeepPersianFont" in fonts
    assert "TableUniqueFont" in fonts


def test_generation_font_verification_and_fallback(tmp_path):
    from pathlib import Path
    from pptx_jahat.tools.pptx_builder import build_pptx_with_agent
    from pptx_jahat.tools.pptx_engine import extract_template_fonts

    review_called = []
    def on_human_review(step, step_data):
        review_called.append((step, step_data))
        if step == "font_fallback":
            return {"font_fallbacks": {mf: "IRANYekanXFaNum" for mf in step_data.get("missing_fonts", [])}}
        return step_data

    out_file = tmp_path / "gen_with_fallbacks.pptx"
    res = build_pptx_with_agent(
        raw_text="Quantum Computing\nIntroduction to qubits and superposition.",
        output_path=out_file,
        template_name="T714.pptx",
        on_human_review=on_human_review
    )
    assert Path(res).exists()
    assert any(step == "font_fallback" for step, _ in review_called)

    prs = Presentation(res)
    detected = extract_template_fonts(prs)
    # The missing font 'Amuzeh-New-Bold' should have been replaced and not exist in output
    assert "Amuzeh-New-Bold" not in detected

