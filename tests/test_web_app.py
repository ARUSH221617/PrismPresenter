import pytest
from pathlib import Path
from pptx_jahat.web.app import create_app
from pptx_jahat.config import DATA_DIR, OUTPUT_DIR, Config

@pytest.fixture
def client():
    app = create_app()
    app.config["TESTING"] = True
    with app.test_client() as client:
        yield client

def test_index_page(client):
    res = client.get("/")
    assert res.status_code == 200
    assert b"PRISMPRESENTER" in res.data or b"PPTX JAHAT" in res.data
    assert b"Generator" in res.data

def test_templates_api(client):
    res = client.get("/api/templates/list")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "templates" in data
    assert isinstance(data["templates"], list)
    if data["templates"]:
        first = data["templates"][0]
        assert "filename" in first
        assert "domain" in first
        assert "is_structured" in first

def test_template_schema_api(client):
    res = client.get("/api/templates/schema")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "schema" in data
    assert "Standard Structured Template Schema" in data["schema"]
    assert "1. Template Profile & Visual Identity" in data["schema"]
    assert "3. Slide Architecture & Slot Blueprint" in data["schema"]

def test_template_boilerplate_api(client):
    res = client.get("/api/templates/boilerplate?filename=T711.pptx")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "boilerplate" in data
    assert "T711.pptx" in data["boilerplate"]
    assert "Slide Architecture & Slot Blueprint" in data["boilerplate"]

def test_template_apply_schema_api(client):
    templates = list(DATA_DIR.glob("*.pptx"))
    target_name = templates[0].name if templates else "T711.pptx"
    res = client.post("/api/templates/apply-schema", json={"filename": target_name})
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "note" in data
    assert "Slide Architecture & Slot Blueprint" in data["note"]

def test_generator_templates_api(client):
    res = client.get("/api/generator/templates")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "templates" in data

def test_manager_decks_api(client):
    res = client.get("/api/manager/decks")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "generated" in data
    assert "reference" in data

def test_components_catalog_api(client):
    res = client.get("/api/components/catalog")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "catalog" in data

def test_config_api(client):
    res = client.get("/api/config")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "config" in data
    assert "NINEROUTER_URL" in data["config"]
    assert "RENDER_MODE" in data["config"]

def test_config_post_updates_model(client):
    prev_model = Config.NINEROUTER_CHAT_MODEL
    try:
        post_res = client.post("/api/config", json={
            "config": {
                "NINEROUTER_CHAT_MODEL": "ag/gemini-3.8-flash-high"
            }
        })
        assert post_res.status_code == 200
        data = post_res.get_json()
        assert data["success"] is True
        assert data["model"] == "ag/gemini-3.8-flash-high"
        assert Config.NINEROUTER_CHAT_MODEL == "ag/gemini-3.8-flash-high"
    finally:
        Config.NINEROUTER_CHAT_MODEL = prev_model

def test_models_api_default(client):
    res = client.get("/api/models")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "models" in data
    assert isinstance(data["models"], list)
    assert len(data["models"]) > 0
    assert "providers" in data
    assert "categories" in data
    assert "recommended" in data
    # Test that model structure contains expected fields
    sample = data["models"][0]
    assert "id" in sample
    assert "owned_by" in sample
    assert "provider" in sample
    assert "capabilities" in sample

def test_models_api_categories(client):
    for cat in ["chat", "image", "web", "search", "fetch", "all"]:
        res = client.get(f"/api/models?category={cat}")
        assert res.status_code == 200
        data = res.get_json()
        assert data["success"] is True
        assert "models" in data
        assert isinstance(data["models"], list)
        assert len(data["models"]) > 0

def test_models_api_with_custom_endpoint(client):
    # Test that offline / invalid URL returns fallback models safely without crashing
    res = client.get("/api/models?url=http://127.0.0.1:59999&refresh=1")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert data["connected"] is False
    assert len(data["models"]) > 0
    assert data["source"] == "fallback"

def test_index_page_contains_model_browser(client):
    res = client.get("/")
    assert res.status_code == 200
    # Check that the 9Router model suggestion elements are in the template
    assert b"btn-toggle-model-browser" in res.data
    assert b"cfg-chat-model-datalist" in res.data
    assert b"model-browser-drawer" in res.data
    assert b"quick-model-chips" in res.data

def test_generator_diagnostics_api(client):
    res = client.get("/api/generator/diagnostics")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "diagnostics" in data
    assert isinstance(data["diagnostics"], list)
    assert len(data["diagnostics"]) == 9
    # Check that each diagnostic item has status, input, output
    step_ids = [s["id"] for s in data["diagnostics"]]
    assert "step_1" in step_ids
    assert "step_1_font" in step_ids
    assert "step_1_5" in step_ids
    assert "step_2" in step_ids
    assert "step_4_5" in step_ids
    for step in data["diagnostics"]:
        assert "id" in step
        assert "name" in step
        assert "status" in step
        assert "input" in step
        assert "output" in step
        assert step["status"] in ("pending", "running", "completed", "skipped")

def test_system_fonts_api(client):
    res = client.get("/api/fonts/system")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "families" in data
    assert "persian_fonts" in data
    assert "latin_fonts" in data
    assert len(data["families"]) > 0
    assert "total_count" in data

def test_system_fonts_rescan_api(client):
    res = client.post("/api/fonts/system/rescan")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert data["total_count"] > 0

def test_verify_fonts_api(client):
    # Test GET with explicit font query
    res = client.get("/api/fonts/verify?fonts=FakeMissingFont123,IRANYekanXFaNum")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "missing_fonts" in data
    assert "FakeMissingFont123" in data["missing_fonts"]
    assert "recommendations" in data
    assert "FakeMissingFont123" in data["recommendations"]

    # Test POST with template name
    res_tpl = client.post("/api/fonts/verify", json={"template_name": "T711.pptx"})
    assert res_tpl.status_code == 200
    data_tpl = res_tpl.get_json()
    assert data_tpl["success"] is True
    assert "installed_fonts" in data_tpl
    assert "missing_fonts" in data_tpl

def test_dev_status_disabled_by_default(client):
    res = client.get("/api/dev/status")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert data["dev_mode"] is False
    assert data["auto_reload"] is False
    assert "boot_id" in data
    assert "watched_directories" in data

def test_dev_live_reload_endpoint_disabled_in_prod(client):
    res = client.get("/api/dev/live-reload")
    assert res.status_code == 404
    data = res.get_json()
    assert data["enabled"] is False

def test_dev_mode_active():
    dev_app = create_app(dev_mode=True)
    dev_app.config["TESTING"] = True
    with dev_app.test_client() as dev_client:
        # Check dev status
        res = dev_client.get("/api/dev/status")
        assert res.status_code == 200
        data = res.get_json()
        assert data["success"] is True
        assert data["dev_mode"] is True
        assert data["auto_reload"] is True
        assert len(data["watched_directories"]) >= 2

        # Check index page contains DEV AUTO-RELOAD badge
        idx_res = dev_client.get("/")
        assert idx_res.status_code == 200
        assert b"DEV AUTO-RELOAD" in idx_res.data
        assert b"/api/dev/live-reload" in idx_res.data

        # Check live reload SSE endpoint init event
        stream_res = dev_client.get("/api/dev/live-reload")
        assert stream_res.status_code == 200
        assert "text/event-stream" in stream_res.headers.get("Content-Type", "")
        # Read the first event from the generator
        first_chunk = next(stream_res.response)
        assert b"event: init" in first_chunk
        assert b"boot_id" in first_chunk
        stream_res.close()

def test_cli_dev_mode_invocation(monkeypatch):
    import sys
    from unittest.mock import MagicMock
    import pptx_jahat

    mock_app = MagicMock()
    mock_create_app = MagicMock(return_value=mock_app)
    monkeypatch.setattr(pptx_jahat, "create_app", mock_create_app)
    monkeypatch.setattr(sys, "argv", ["pptx-jahat", "--dev", "--no-browser", "--port", "5555"])

    pptx_jahat.main()

    mock_create_app.assert_called_once_with(dev_mode=True)
    mock_app.run.assert_called_once()
    kwargs = mock_app.run.call_args.kwargs
    assert kwargs["debug"] is True
    assert kwargs["use_reloader"] is True
    assert kwargs["port"] == 5555


def test_local_static_assets_served(client):
    """Ensure all critical frontend assets are hosted locally and served with HTTP 200."""
    local_assets = [
        "/static/js/tailwindcss.min.js",
        "/static/js/lucide.min.js",
        "/static/js/marked.min.js",
        "/static/fonts/Geist-Variable.woff2",
        "/static/fonts/GeistMono-Variable.woff2",
        "/static/fonts/Inter-Variable.woff2",
        "/static/css/custom.css",
        "/static/js/app.js",
        "/static/js/pptx-web-renderer.js",
    ]
    for asset_url in local_assets:
        res = client.get(asset_url)
        assert res.status_code == 200, f"Failed to load local asset: {asset_url}"
        assert len(res.data) > 0, f"Local asset is empty: {asset_url}"


def test_no_external_cdn_in_templates_or_css(client):
    """Ensure templates and css do not load scripts or styles from outside providers."""
    # Check index.html rendered page
    res = client.get("/")
    assert res.status_code == 200
    html_content = res.data.decode("utf-8")

    external_cdns = [
        "https://cdn.tailwindcss.com",
        "https://unpkg.com/lucide",
        "https://cdn.jsdelivr.net/npm/marked",
        "https://fonts.googleapis.com",
    ]
    for cdn in external_cdns:
        assert cdn not in html_content, f"Found external CDN in index.html: {cdn}"

    # Check that local assets are referenced
    assert "/static/js/tailwindcss.min.js" in html_content
    assert "/static/js/lucide.min.js" in html_content
    assert "/static/js/marked.min.js" in html_content

    # Check custom.css does not import external fonts
    css_res = client.get("/static/css/custom.css")
    assert css_res.status_code == 200
    css_content = css_res.data.decode("utf-8")
    assert "fonts.googleapis.com" not in css_content
    assert "/static/fonts/Geist-Variable.woff2" in css_content


def test_config_ping_api(client):
    res = client.post("/api/config/ping", json={"url": "http://127.0.0.1:99999", "key": ""})
    assert res.status_code == 200
    data = res.get_json()
    assert "latency_ms" in data

def test_config_diagnostics_api(client):
    res = client.get("/api/config/diagnostics")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "platform" in data
    assert "storage" in data
    assert "runtime" in data

def test_config_raw_api(client):
    res = client.get("/api/config/raw")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "raw" in data

def test_config_clean_cache_api(client):
    res = client.post("/api/config/clean-cache")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "cleaned_count" in data


def test_config_com_probe_api(client):
    res = client.post("/api/config/com-probe")
    assert res.status_code == 200
    data = res.get_json()
    assert "status" in data
    assert "latency_ms" in data


def test_config_clean_cache_dry_run_api(client):
    res = client.post("/api/config/clean-cache", json={"dry_run": True})
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "dry_run" in data
    assert data["dry_run"] is True
    assert "cleaned_count" in data


def test_config_backup_api(client):
    res = client.get("/api/config/backup")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "backup_exists" in data

