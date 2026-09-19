import pytest
from unittest.mock import MagicMock
from pptx_jahat.config import Config, AGENTS_CONFIG_SCHEMA
from pptx_jahat.web.app import create_app


def test_agents_config_schema_contains_all_seven_agents():
    expected_agents = {"generator", "verifier", "autonomous", "analyzer", "structure", "editor", "parser"}
    assert set(AGENTS_CONFIG_SCHEMA.keys()) == expected_agents
    for aid, meta in AGENTS_CONFIG_SCHEMA.items():
        assert meta["id"] == aid
        assert "name" in meta
        assert "role" in meta
        assert "icon" in meta


def test_agent_model_fallback_to_primary():
    prev_models = dict(Config.AGENT_MODELS)
    try:
        Config.AGENT_MODELS["generator"] = ""
        assert Config.get_agent_model("generator") == Config.NINEROUTER_CHAT_MODEL

        Config.AGENT_MODELS["generator"] = "default"
        assert Config.get_agent_model("generator") == Config.NINEROUTER_CHAT_MODEL
    finally:
        Config.AGENT_MODELS = prev_models


def test_agent_model_custom_override():
    prev_models = dict(Config.AGENT_MODELS)
    try:
        Config.AGENT_MODELS["verifier"] = "openai/gpt-4o"
        assert Config.get_agent_model("verifier") == "openai/gpt-4o"
    finally:
        Config.AGENT_MODELS = prev_models


def test_agent_think_level_normalization():
    prev_levels = dict(Config.AGENT_THINK_LEVELS)
    try:
        Config.AGENT_THINK_LEVELS["generator"] = "high"
        assert Config.get_agent_think_level("generator") == "high"

        Config.AGENT_THINK_LEVELS["verifier"] = "low"
        assert Config.get_agent_think_level("verifier") == "low"

        Config.AGENT_THINK_LEVELS["autonomous"] = "medium"
        assert Config.get_agent_think_level("autonomous") == "medium"

        Config.AGENT_THINK_LEVELS["analyzer"] = "none"
        assert Config.get_agent_think_level("analyzer") == "none"

        Config.AGENT_THINK_LEVELS["structure"] = "default"
        assert Config.get_agent_think_level("structure") == "default"
    finally:
        Config.AGENT_THINK_LEVELS = prev_levels


def test_get_agent_llm_kwargs_reasoning_effort():
    prev_models = dict(Config.AGENT_MODELS)
    prev_levels = dict(Config.AGENT_THINK_LEVELS)
    try:
        Config.AGENT_MODELS["generator"] = "anthropic/claude-3-7-sonnet"
        Config.AGENT_THINK_LEVELS["generator"] = "high"

        kwargs = Config.get_agent_llm_kwargs("generator", temperature=0.2)
        assert kwargs["model"] == "anthropic/claude-3-7-sonnet"
        assert kwargs["reasoning_effort"] == "high"
        assert kwargs["temperature"] == 0.2

        # When think level is default or none, reasoning_effort is omitted
        Config.AGENT_THINK_LEVELS["generator"] = "default"
        kwargs_default = Config.get_agent_llm_kwargs("generator")
        assert "reasoning_effort" not in kwargs_default

        Config.AGENT_THINK_LEVELS["generator"] = "none"
        kwargs_none = Config.get_agent_llm_kwargs("generator")
        assert "reasoning_effort" not in kwargs_none
    finally:
        Config.AGENT_MODELS = prev_models
        Config.AGENT_THINK_LEVELS = prev_levels


def test_safe_chat_completion_fallback_when_reasoning_effort_unsupported():
    mock_client = MagicMock()
    # First call with reasoning_effort raises an error, second call succeeds
    error_call = Exception("Invalid parameter: reasoning_effort is extra forbidden for this model")
    mock_resp = MagicMock()
    mock_client.chat.completions.create.side_effect = [error_call, mock_resp]

    prev_levels = dict(Config.AGENT_THINK_LEVELS)
    try:
        Config.AGENT_THINK_LEVELS["autonomous"] = "high"
        res = Config.safe_chat_completion(mock_client, "autonomous", messages=[{"role": "user", "content": "Hi"}])
        assert res == mock_resp
        assert mock_client.chat.completions.create.call_count == 2
        # Verify second call stripped reasoning_effort
        second_call_kwargs = mock_client.chat.completions.create.call_args_list[1][1]
        assert "reasoning_effort" not in second_call_kwargs
    finally:
        Config.AGENT_THINK_LEVELS = prev_levels


def test_get_agents_config_payload():
    config_dict = Config.get_agents_config()
    assert len(config_dict) == 7
    for aid in AGENTS_CONFIG_SCHEMA:
        assert aid in config_dict
        item = config_dict[aid]
        assert "effective_model" in item
        assert "think_level" in item
        assert "is_custom_model" in item


def test_api_config_get_and_post_agents():
    app = create_app()
    client = app.test_client()

    # Test GET includes agents
    res = client.get("/api/config")
    assert res.status_code == 200
    data = res.get_json()
    assert data["success"] is True
    assert "agents" in data
    assert "generator" in data["agents"]
    assert "AGENT_MODEL_GENERATOR" in data["config"]
    assert "AGENT_THINK_LEVEL_GENERATOR" in data["config"]

    # Test POST updates agent model and think level
    post_res = client.post("/api/config", json={
        "config": {
            "AGENT_MODEL_GENERATOR": "claude-pro-agent",
            "AGENT_THINK_LEVEL_GENERATOR": "high",
            "AGENT_MODEL_VERIFIER": "openai/gpt-4o",
            "AGENT_THINK_LEVEL_VERIFIER": "low"
        }
    })
    assert post_res.status_code == 200
    post_data = post_res.get_json()
    assert post_data["success"] is True
    assert post_data["agents"]["generator"]["configured_model"] == "claude-pro-agent"
    assert post_data["agents"]["generator"]["think_level"] == "high"
    assert post_data["agents"]["verifier"]["configured_model"] == "openai/gpt-4o"
    assert post_data["agents"]["verifier"]["think_level"] == "low"
    assert Config.get_agent_model("generator") == "claude-pro-agent"
    assert Config.get_agent_think_level("generator") == "high"

    # Reset back to default
    client.post("/api/config", json={
        "config": {
            "AGENT_MODEL_GENERATOR": "",
            "AGENT_THINK_LEVEL_GENERATOR": "default",
            "AGENT_MODEL_VERIFIER": "",
            "AGENT_THINK_LEVEL_VERIFIER": "default"
        }
    })
