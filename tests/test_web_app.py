import pytest
from pathlib import Path
from pptx_jahat.web.app import create_app
from pptx_jahat.config import DATA_DIR, OUTPUT_DIR

@pytest.fixture
def client():
    app = create_app()
    app.config["TESTING"] = True
    with app.test_client() as client:
        yield client

def test_index_page(client):
    res = client.get("/")
    assert res.status_code == 200
    assert b"PPTX JAHAT" in res.data
    assert b"Slide Generator" in res.data

def test_templates_api(client):
    res = client.get("/api/templates/list")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "templates" in data
    assert isinstance(data["templates"], list)

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

def test_generator_diagnostics_api(client):
    res = client.get("/api/generator/diagnostics")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "diagnostics" in data
    assert isinstance(data["diagnostics"], list)
    assert len(data["diagnostics"]) == 7
    # Check that each diagnostic item has status, input, output
    for step in data["diagnostics"]:
        assert "id" in step
        assert "name" in step
        assert "status" in step
        assert "input" in step
        assert "output" in step
        assert step["status"] in ("pending", "running", "completed", "skipped")

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


