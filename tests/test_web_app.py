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
