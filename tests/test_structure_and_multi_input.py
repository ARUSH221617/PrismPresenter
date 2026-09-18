import pytest
import tempfile
from pathlib import Path
from pptx_jahat.config import DATA_DIR, STRUCTURES_DIR
from pptx_jahat.tools.structure_manager import (
    list_structure_files,
    get_structure_content,
    save_structure_content,
    delete_structure_file,
    build_structure_from_template,
    restructure_slides_with_agent,
    convert_restructured_to_sections,
    _fallback_restructure
)
from pptx_jahat.tools.multi_parser import (
    is_supported_file,
    get_file_type_category,
    parse_text_content,
    parse_multiple_sources
)
from pptx_jahat.web.app import create_app

@pytest.fixture
def client():
    app = create_app()
    app.config["TESTING"] = True
    with app.test_client() as client:
        yield client

def test_list_structure_files():
    structures = list_structure_files()
    assert isinstance(structures, list)
    assert len(structures) >= 1
    sample = next((s for s in structures if s["filename"] == "pptx-structure-yosefzadeh.md"), None)
    assert sample is not None
    assert sample["is_sample"] is True

def test_get_structure_content():
    content = get_structure_content("pptx-structure-yosefzadeh.md")
    assert "Slide Storyboard Specification Schema" in content
    assert "Quadrant Layout Mapping" in content
    assert "Slide Content Schema" in content

def test_save_and_delete_structure_content():
    test_name = "test-auto-unit-struct.md"
    test_content = "# Unit Test Storyboard\n## 1. Metadata\n"
    res = save_structure_content(test_name, test_content)
    assert res["success"] is True

    # Read it back
    read_back = get_structure_content(test_name)
    assert "Unit Test Storyboard" in read_back

    # Sample cannot be deleted
    assert delete_structure_file("pptx-structure-yosefzadeh.md") is False

    # Test file can be deleted
    assert delete_structure_file(test_name) is True

def test_fallback_restructure_and_sections():
    sample_content = {
        "document_title": "Quantum Mechanics 101",
        "sections": [
            {
                "title": "Wave-Particle Duality",
                "paragraphs": ["Light exhibits properties of both waves and particles."],
                "bullets": ["Photon energy E = hf"]
            }
        ]
    }
    struct_schema = get_structure_content("pptx-structure-yosefzadeh.md")
    restructured = _fallback_restructure(sample_content, struct_schema)
    assert restructured["total_slides"] == 1
    slide = restructured["slides"][0]
    assert slide["quadrant"] == "TR"
    assert len(slide["animation_steps"]) > 0

    sections = convert_restructured_to_sections(restructured)
    assert len(sections) == 1
    assert sections[0]["title"] == "Wave-Particle Duality"
    assert len(sections[0]["bullets"]) > 0

def test_multi_parser_support_and_text():
    assert is_supported_file("notes.docx") is True
    assert is_supported_file("presentation.pptx") is True
    assert is_supported_file("outline.md") is True
    assert is_supported_file("diagram.png") is True
    assert is_supported_file("recording.mp3") is True
    assert is_supported_file("unsupported.xyz") is False

    assert get_file_type_category("paper.docx") == "Word Document"
    assert get_file_type_category("deck.pptx") == "PowerPoint Presentation"
    assert get_file_type_category("audio.mp3") == "Audio Recording / Speech"
    assert get_file_type_category("sketch.png") == "Image / Visual Diagram"

    # Test parsing multiple sources with text and direct input
    with tempfile.NamedTemporaryFile("w", suffix=".md", delete=False, encoding="utf-8") as tf:
        tf.write("# Topic A\nParagraph 1\n- Bullet 1\n## Subtopic B\nParagraph 2\n")
        tf_path = Path(tf.name)

    try:
        parsed = parse_multiple_sources([tf_path], raw_text="Important conclusion note.")
        assert parsed["total_sections"] >= 2
        assert len(parsed["sources_summary"]) == 2
    finally:
        if tf_path.exists():
            tf_path.unlink()

def test_structure_web_endpoints(client):
    # 1. List structures
    res = client.get("/api/structure/list")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert data["count"] >= 1

    # 2. Get specific structure
    res2 = client.get("/api/structure/get?name=pptx-structure-yosefzadeh.md")
    assert res2.status_code == 200
    d2 = res2.get_json()
    assert "Storyboard" in d2["content"]

    # 3. Save custom structure
    temp_name = "web-test-struct.md"
    res3 = client.post("/api/structure/save", json={
        "name": temp_name,
        "content": "# Web Test Storyboard\n## Schema..."
    })
    assert res3.status_code == 200

    # 4. Delete custom structure
    res4 = client.delete("/api/structure/delete", json={"name": temp_name})
    assert res4.status_code == 200

def test_multi_upload_endpoint(client):
    import io
    data = {
        "files": [
            (io.BytesIO(b"# Lecture Notes\nSection 1"), "test_lecture.txt"),
            (io.BytesIO(b"# Research Notes\nSection 2"), "test_paper.md")
        ]
    }
    res = client.post("/api/generator/upload-multi", data=data, content_type="multipart/form-data")
    assert res.status_code == 200
    d = res.get_json()
    assert d["success"] is True
    assert len(d["files"]) == 2

def test_timeout_config_and_metadata(client):
    res = client.get("/api/config")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "LLM_TIMEOUT" in data["config"]
    assert "model_metadata" in data
    assert "max_tokens" in data["model_metadata"]
    assert "context_length" in data["model_metadata"]

def test_safe_json_loads_latex_escapes():
    from pptx_jahat.tools.json_parser import safe_json_loads

    # Unescaped LaTeX backslashes that typically cause 'Invalid \escape'
    latex_json = r'''```json
    {
      "title": "Mathematics Test",
      "formula": "91 = 7 \times 13 and \cdot and \sqrt{4}",
      "rules": "x \in \mathbb{R}"
    }
    ```'''
    parsed = safe_json_loads(latex_json)
    assert parsed["title"] == "Mathematics Test"
    assert "91 = 7" in parsed["formula"]
    assert "x" in parsed["rules"]

def test_build_structure_from_sample_files():
    from pptx_jahat.tools.structure_manager import build_structure_from_sample_files, delete_structure_file

    with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False, encoding="utf-8") as tf:
        tf.write("# Sample Paper Sheet\nProblem: Is 37 prime?\nAnswer: Yes")
        tf_path = Path(tf.name)

    struct_name = "test-sample-detected-struct"
    try:
        res = build_structure_from_sample_files([tf_path], struct_name)
        assert res["success"] is True
        assert res["name"] == struct_name
        assert "Slide Storyboard" in res["content"]
        assert "Document Metadata" in res["content"]
        assert Path(res["path"]).exists()
    finally:
        delete_structure_file(f"{struct_name}.md")
        if tf_path.exists():
            tf_path.unlink()

def test_upload_samples_and_build_endpoints(client):
    import io
    # 1. Upload sample files for detection structure
    data = {
        "files": [
            (io.BytesIO(b"# Sample Paper Sheet 1\nQuadrant 1: Prime numbers"), "sample1.txt"),
            (io.BytesIO(b"# Sample Paper Sheet 2\nQuadrant 2: Worked examples"), "sample2.md")
        ]
    }
    res = client.post("/api/structure/upload-samples", data=data, content_type="multipart/form-data")
    assert res.status_code == 200
    d = res.get_json()
    assert d["success"] is True
    assert len(d["files"]) == 2

    # 2. Build structure from uploaded samples
    file_paths = [f["file_path"] for f in d["files"]]
    res2 = client.post("/api/structure/build-from-samples", json={
        "sample_files": file_paths,
        "structure_name": "test-endpoint-detection-struct"
    })
    assert res2.status_code == 200
    d2 = res2.get_json()
    assert d2["success"] is True
    assert "job_id" in d2



