import json
import pytest
from pathlib import Path
from unittest.mock import MagicMock
from pptx import Presentation
from pptx.util import Inches, Pt

from pptx_jahat.config import DATA_DIR, Config
from pptx_jahat.tools.slide_verifier import (
    apply_verification_edits,
    prepare_slide_comparison_data,
    run_presentation_visual_verification,
    converse_with_verification_agent,
    _heuristic_verification,
    _normalize_verification_result
)
from pptx_jahat.tools.pptx_builder import get_initial_diagnostics_steps


@pytest.fixture(autouse=True)
def configure_test_env():
    prev_pil = Config.PURE_PIL_ACTIVE
    Config.PURE_PIL_ACTIVE = True
    yield
    Config.PURE_PIL_ACTIVE = prev_pil


@pytest.fixture
def sample_pptx(tmp_path):
    """Creates a sample 2-slide presentation."""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Slide 1
    s1 = prs.slides.add_slide(blank_layout)
    tb1 = s1.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
    tb1.text_frame.text = "Generated Title Slide"
    tb2 = s1.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(1))
    tb2.text_frame.text = "Lorem ipsum placeholder text"

    # Slide 2
    s2 = prs.slides.add_slide(blank_layout)
    tb3 = s2.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1))
    tb3.text_frame.text = "Slide 2 Content"
    table_sh = s2.shapes.add_table(2, 2, Inches(1), Inches(2.5), Inches(6), Inches(3))
    t = table_sh.table
    t.cell(0, 0).text = "A1"
    t.cell(0, 1).text = "B1"

    out_file = tmp_path / "test_deck.pptx"
    prs.save(str(out_file))
    return out_file


def test_diagnostics_includes_step_4_5():
    """Verify Step 4.5 is part of diagnostics steps list."""
    steps = get_initial_diagnostics_steps()
    step_ids = [s["id"] for s in steps]
    assert "step_4_5" in step_ids
    s45 = next(s for s in steps if s["id"] == "step_4_5")
    assert "Verification" in s45["name"]


def test_apply_verification_edits(sample_pptx):
    """Test applying text updates and shape removals."""
    actions = [
        {
            "action": "update_text",
            "slide_index": 0,
            "shape_index": 1,
            "new_text": "Healed Subtitle Text"
        },
        {
            "action": "update_table",
            "slide_index": 1,
            "shape_index": 1,
            "table_data": [["New A1", "New B1"]]
        }
    ]

    res = apply_verification_edits(sample_pptx, actions)
    assert res["success"] is True
    assert res["applied_count"] >= 2

    # Verify changes in PPTX
    prs = Presentation(str(sample_pptx))
    assert prs.slides[0].shapes[1].text_frame.text == "Healed Subtitle Text"
    assert prs.slides[1].shapes[1].table.cell(0, 0).text == "New A1"


def test_heuristic_verification():
    """Test heuristic detection of placeholder texts."""
    slide_pair = {
        "slide_index": 0,
        "generated_shapes": [
            {"shape_index": 0, "text": "Valid Title"},
            {"shape_index": 1, "text": "Click to edit text"}
        ]
    }
    res = _heuristic_verification(slide_pair, 0)
    assert res["is_correct"] is False
    assert len(res["detected_issues"]) == 1
    assert "Shape #1" in res["detected_issues"][0]


def test_prepare_slide_comparison_data(sample_pptx):
    """Test generating paired comparison entries between template and generated slide."""
    fake_inventory = [
        {
            "template_file": "T711.pptx",
            "slide_index": 0,
            "screenshot_base64": "data:image/jpeg;base64,fake",
            "text_slots": [{"shape_index": 0, "original_text": "Template Title"}]
        }
    ]
    ai_plan = {
        "slides": [
            {"source_template": "T711.pptx", "source_slide_index": 0, "target_section": "Intro"}
        ]
    }

    pairs = prepare_slide_comparison_data(
        sample_pptx,
        ai_plan=ai_plan,
        template_inventory=fake_inventory,
        screenshot_width=300
    )
    assert len(pairs) == 2
    assert pairs[0]["slide_index"] == 0
    assert pairs[0]["source_template"] == "T711.pptx"
    assert pairs[0]["template_screenshot"] == "data:image/jpeg;base64,fake"


def test_run_presentation_visual_verification_mocked(sample_pptx, monkeypatch):
    """Test full verification pipeline with mocked AI response."""
    fake_inventory = [
        {
            "template_file": "T711.pptx",
            "slide_index": 0,
            "screenshot_base64": "data:image/jpeg;base64,fake",
            "text_slots": [{"shape_index": 0, "original_text": "Template Title"}]
        }
    ]

    mock_response = MagicMock()
    mock_response.choices = [
        MagicMock(
            message=MagicMock(
                content=json.dumps({
                    "is_correct": False,
                    "score": 70,
                    "detected_issues": ["Missing template subtitle"],
                    "edit_structure": {
                        "summary_of_changes": "Added missing subtitle",
                        "actions": [
                            {
                                "action": "update_text",
                                "slide_index": 0,
                                "shape_index": 1,
                                "new_text": "Verified Subtitle"
                            }
                        ]
                    }
                })
            )
        )
    ]

    monkeypatch.setattr(
        "openai.resources.chat.completions.Completions.create",
        lambda *args, **kwargs: mock_response
    )

    verif = run_presentation_visual_verification(
        pptx_path=sample_pptx,
        ai_plan={"slides": [{"source_template": "T711.pptx", "source_slide_index": 0}]},
        template_inventory=fake_inventory
    )

    assert verif["all_correct"] is False
    assert verif["total_issues_found"] > 0
    assert len(verif["aggregated_actions"]) > 0
    assert verif["slides"][0]["detected_issues"] == ["Missing template subtitle"]


def test_converse_with_verification_agent_mocked(monkeypatch):
    """Test user conversation with Verification Agent."""
    mock_response = MagicMock()
    mock_response.choices = [
        MagicMock(
            message=MagicMock(
                content=json.dumps({
                    "agent_reply": "Understood. Restored the subtitle and preserved the badge.",
                    "updated_slides": [
                        {
                            "slide_index": 0,
                            "is_correct": False,
                            "detected_issues": ["User requested subtitle restoration"],
                            "edit_structure": {
                                "summary_of_changes": "Restored subtitle",
                                "actions": [
                                    {
                                        "action": "update_text",
                                        "slide_index": 0,
                                        "shape_index": 1,
                                        "new_text": "Annual Report 2026"
                                    }
                                ]
                            }
                        }
                    ],
                    "aggregated_actions": [
                        {
                            "action": "update_text",
                            "slide_index": 0,
                            "shape_index": 1,
                            "new_text": "Annual Report 2026"
                        }
                    ]
                })
            )
        )
    ]

    monkeypatch.setattr(
        "openai.resources.chat.completions.Completions.create",
        lambda *args, **kwargs: mock_response
    )

    current_data = {
        "slides": [
            {
                "slide_index": 0,
                "slide_number": 1,
                "title": "Title",
                "detected_issues": [],
                "edit_structure": {"actions": []},
                "generated_shapes": []
            }
        ]
    }

    res = converse_with_verification_agent(
        current_verification_data=current_data,
        user_message="Slide 1 is missing subtitle, set it to 'Annual Report 2026'",
        active_slide_index=0
    )

    assert res["success"] is True
    assert "Restored the subtitle" in res["agent_reply"]
    assert len(res["aggregated_actions"]) == 1
    assert res["aggregated_actions"][0]["new_text"] == "Annual Report 2026"


def test_web_verify_endpoints(tmp_path, monkeypatch):
    """Test /api/generator/verify/chat and /api/generator/verify/apply endpoints."""
    from pptx_jahat.web.app import create_app
    app = create_app()
    app.config["TESTING"] = True

    # Test verify chat endpoint
    monkeypatch.setattr(
        "pptx_jahat.tools.slide_verifier.converse_with_verification_agent",
        lambda **kwargs: {"success": True, "agent_reply": "Chat response OK", "slides": []}
    )

    with app.test_client() as client:
        chat_res = client.post("/api/generator/verify/chat", json={
            "message": "Check slide 1",
            "verification_data": {"slides": []}
        })
        assert chat_res.status_code == 200
        chat_json = chat_res.get_json()
        assert chat_json["success"] is True
        assert chat_json["agent_reply"] == "Chat response OK"

        # Test apply endpoint
        prs = Presentation()
        s = prs.slides.add_slide(prs.slide_layouts[6])
        tb = s.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(2))
        tb.text_frame.text = "Old"
        p = tmp_path / "deck.pptx"
        prs.save(str(p))

        apply_res = client.post("/api/generator/verify/apply", json={
            "file_path": str(p),
            "actions": [
                {"action": "update_text", "slide_index": 0, "shape_index": 0, "new_text": "New Text"}
            ]
        })
        assert apply_res.status_code == 200
        apply_json = apply_res.get_json()
        assert apply_json["success"] is True
        assert apply_json["applied_count"] == 1


def test_config_verification_rounds_default_and_env(monkeypatch):
    """Test VERIFICATION_ROUNDS default is 3 and configurable via env var."""
    monkeypatch.setattr("pptx_jahat.config.load_dotenv", lambda *args, **kwargs: None)
    monkeypatch.setenv("VERIFICATION_ROUNDS", "5")
    Config.reload()
    assert Config.VERIFICATION_ROUNDS == 5

    # Test fallback env variable
    monkeypatch.delenv("VERIFICATION_ROUNDS", raising=False)
    monkeypatch.setenv("VERIFICATION_MAX_ROUNDS", "4")
    Config.reload()
    assert Config.VERIFICATION_ROUNDS == 4

    # Reset
    monkeypatch.delenv("VERIFICATION_MAX_ROUNDS", raising=False)
    Config.reload()
    assert Config.VERIFICATION_ROUNDS == 3


def test_run_presentation_visual_verification_with_rounds(sample_pptx):
    """Test that verification returns round and max_rounds metadata."""
    fake_inventory = [
        {
            "template_file": "T711.pptx",
            "slide_index": 0,
            "screenshot_base64": "data:image/jpeg;base64,fake",
            "text_slots": [{"shape_index": 0, "original_text": "Template Title"}]
        }
    ]

    verif = run_presentation_visual_verification(
        pptx_path=sample_pptx,
        ai_plan={"slides": [{"source_template": "T711.pptx", "source_slide_index": 0}]},
        template_inventory=fake_inventory,
        round_num=2,
        max_rounds=3
    )
    assert verif["round"] == 2
    assert verif["max_rounds"] == 3


def test_config_api_verification_rounds():
    """Test GET /api/config returns VERIFICATION_ROUNDS."""
    from pptx_jahat.web.app import create_app
    app = create_app()
    app.config["TESTING"] = True

    with app.test_client() as client:
        res = client.get("/api/config")
        assert res.status_code == 200
        data = res.get_json()
        assert data["success"] is True
        assert "VERIFICATION_ROUNDS" in data["config"]
        assert data["config"]["VERIFICATION_ROUNDS"] == 3


def test_delete_shape_and_element_actions(sample_pptx):
    """Test various deletion action aliases (delete_shape, remove_shape, delete_element)."""
    prs = Presentation(str(sample_pptx))
    initial_shapes = len(prs.slides[0].shapes)
    assert initial_shapes == 2

    # Delete shape 1 using action 'delete_shape'
    actions = [
        {"action": "delete_shape", "slide_index": 0, "shape_index": 1}
    ]
    res = apply_verification_edits(sample_pptx, actions)
    assert res["success"] is True
    assert res["applied_count"] >= 1

    prs2 = Presentation(str(sample_pptx))
    assert len(prs2.slides[0].shapes) == initial_shapes - 1


def test_delete_unreplaced_template_formula_alternate_content(tmp_path):
    """Test deleting AlternateContent math formula elements (e.g. '2n') from slide."""
    import shutil
    tpl_path = DATA_DIR / "قالب جهت.pptx"
    if not tpl_path.exists():
        pytest.skip("Template 'قالب جهت.pptx' not available.")

    # Copy template to temp
    test_pptx = tmp_path / "test_formula_removal.pptx"
    shutil.copy(tpl_path, test_pptx)

    prs = Presentation(str(test_pptx))
    s3 = prs.slides[3]
    initial_alts = len(s3._element.xpath(".//*[local-name()='AlternateContent']"))
    assert initial_alts > 0

    # Delete formula containing '2n'
    actions = [
        {"action": "remove_formula", "slide_index": 3, "formula_text": "2n"}
    ]
    res = apply_verification_edits(test_pptx, actions)
    assert res["success"] is True
    assert res["applied_count"] >= 1

    # Verify formula was removed
    prs2 = Presentation(str(test_pptx))
    remaining_alts = len(prs2.slides[3]._element.xpath(".//*[local-name()='AlternateContent']"))
    assert remaining_alts < initial_alts


def test_action_executor_full_access(sample_pptx):
    """Test full access in Action Executor: move, resize, format, style, add, duplicate, execute_python."""
    actions = [
        {
            "action": "move_shape",
            "slide_index": 0,
            "shape_index": 0,
            "left": "150pt",
            "top": "80pt"
        },
        {
            "action": "resize_shape",
            "slide_index": 0,
            "shape_index": 0,
            "width": "500pt",
            "height": "120pt"
        },
        {
            "action": "format_text",
            "slide_index": 0,
            "shape_index": 0,
            "font_size": 24,
            "bold": True,
            "font_color": "#1E3A8A",
            "alignment": "center"
        },
        {
            "action": "set_shape_style",
            "slide_index": 0,
            "shape_index": 0,
            "fill_color": "#F0F4F8",
            "border_color": "#3B82F6",
            "border_width": 2
        },
        {
            "action": "add_shape",
            "slide_index": 0,
            "shape_type": "rounded_rectangle",
            "left": "100pt",
            "top": "300pt",
            "width": "200pt",
            "height": "60pt",
            "text": "New Action Badge"
        },
        {
            "action": "duplicate_shape",
            "slide_index": 0,
            "shape_index": 0,
            "dx": "20pt",
            "dy": "20pt",
            "new_text": "Duplicated Shape"
        },
        {
            "action": "execute_python",
            "slide_index": 0,
            "shape_index": 0,
            "code": "shape.name = 'ExecPyShape'"
        }
    ]

    res = apply_verification_edits(sample_pptx, actions)
    assert res["success"] is True
    assert res["applied_count"] >= 7

    prs = Presentation(str(sample_pptx))
    s0 = prs.slides[0]
    sh0 = s0.shapes[0]
    assert sh0.left == Pt(150)
    assert sh0.top == Pt(80)
    assert sh0.width == Pt(500)
    assert sh0.height == Pt(120)
    assert sh0.name == "ExecPyShape"

    # Verify added shape and duplicated shape exist
    assert len(s0.shapes) >= 4


def test_extract_slide_semantic_layout(sample_pptx):
    """Test generating semantic layout XML for a slide."""
    from pptx_jahat.tools.slide_verifier import extract_slide_semantic_layout
    prs = Presentation(str(sample_pptx))
    xml_str = extract_slide_semantic_layout(prs.slides[0], prs.slide_width, prs.slide_height)
    assert '<slide width=' in xml_str
    assert '<shape index="0"' in xml_str
    assert 'Generated Title Slide' in xml_str
    assert '</slide>' in xml_str



def test_auto_inject_deletion_when_formula_issue_detected():
    """Verify that _normalize_verification_result injects delete action when formula issue is detected."""
    parsed = {
        "is_correct": False,
        "score": 75,
        "detected_issues": [
            "Unreplaced formula text from original template that does not match current topic."
        ],
        "edit_structure": {
            "summary_of_changes": "Found unreplaced formula.",
            "actions": []
        }
    }

    slide_pair = {
        "slide_index": 2,
        "generated_shapes": [
            {"shape_index": 0, "name": "Title", "type": "TITLE", "text": "Topic Title"},
            {"shape_index": 5, "name": "Formula 2n", "type": "FORMULA", "text": "2n", "is_formula": True}
        ]
    }

    normalized = _normalize_verification_result(parsed, s_idx=2, slide_pair=slide_pair)
    assert normalized["is_correct"] is False
    actions = normalized["edit_structure"]["actions"]
    assert len(actions) == 1
    assert actions[0]["action"] in ("delete_shape", "remove_formula")
    assert actions[0]["slide_index"] == 2
    assert actions[0].get("shape_index") == 5 or actions[0].get("formula_text") == "2n"


