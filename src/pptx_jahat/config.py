import os
from typing import Optional
from pathlib import Path
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = DATA_DIR / "output"
COMPONENTS_DIR = DATA_DIR / "components"
SHAPES_DIR = COMPONENTS_DIR / "shapes"
IMAGES_DIR = COMPONENTS_DIR / "images"
STRUCTURES_DIR = DATA_DIR / "structure"

# Load .env file from project root with override=True so local settings always take precedence
env_file = BASE_DIR / ".env"
if env_file.exists():
    load_dotenv(dotenv_path=env_file, override=True)
else:
    load_dotenv(override=True)

# Ensure essential directories exist
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
COMPONENTS_DIR.mkdir(parents=True, exist_ok=True)
SHAPES_DIR.mkdir(parents=True, exist_ok=True)
IMAGES_DIR.mkdir(parents=True, exist_ok=True)
STRUCTURES_DIR.mkdir(parents=True, exist_ok=True)

# Dedicated Agents Registry Schema
AGENTS_CONFIG_SCHEMA = {
    "generator": {
        "id": "generator",
        "name": "Slide Synthesis & Blueprint Agent",
        "role": "Visual template matching, section-to-slide mapping & layout infill reasoning",
        "icon": "presentation",
        "category": "synthesis",
        "default_think": "default"
    },
    "verifier": {
        "id": "verifier",
        "name": "Slide Verification & QA Agent",
        "role": "Step 4.5 visual screenshot QA, typography auditing, and auto-healing",
        "icon": "check-check",
        "category": "qa",
        "default_think": "default"
    },
    "autonomous": {
        "id": "autonomous",
        "name": "Autonomous AI Assistant",
        "role": "Multi-turn Gemini-style terminal, web tool orchestration & autonomous tasks",
        "icon": "sparkles",
        "category": "chat",
        "default_think": "default"
    },
    "analyzer": {
        "id": "analyzer",
        "name": "Template Analyzer Agent",
        "role": "Extracts design intelligence, archetype classifications & NOTE.md documentation",
        "icon": "scan",
        "category": "analysis",
        "default_think": "default"
    },
    "structure": {
        "id": "structure",
        "name": "Storyboard & Structure Agent",
        "role": "Multimodal sample & handwriting analysis to build structure.md blueprints",
        "icon": "layout-template",
        "category": "structure",
        "default_think": "default"
    },
    "editor": {
        "id": "editor",
        "name": "Human Touch / Slide Editor Agent",
        "role": "Interactive user slide adjustments, conversational chat & content refinement",
        "icon": "sliders",
        "category": "editing",
        "default_think": "default"
    },
    "parser": {
        "id": "parser",
        "name": "Document Parser & Vision OCR Agent",
        "role": "Transcribes handwritten notes, worksheets, and complex multi-modal documents",
        "icon": "file-text",
        "category": "parser",
        "default_think": "default"
    }
}

class Config:
    # 9Router Configuration
    NINEROUTER_URL: str = os.getenv("NINEROUTER_URL", "http://localhost:20128")
    NINEROUTER_KEY: str = os.getenv("NINEROUTER_KEY", "")
    NINEROUTER_CHAT_MODEL: str = os.getenv("NINEROUTER_CHAT_MODEL", "ag/gemini-3.7-flash-high")
    NINEROUTER_SEARCH_MODEL: str = os.getenv("NINEROUTER_SEARCH_MODEL", "tavily")
    NINEROUTER_FETCH_MODEL: str = os.getenv("NINEROUTER_FETCH_MODEL", "jina-reader")
    NINEROUTER_IMAGE_MODEL: str = os.getenv("NINEROUTER_IMAGE_MODEL", "gemini/gemini-3-pro-image-preview")
    LLM_TIMEOUT: float = float(os.getenv("LLM_TIMEOUT", "300"))

    # Agent-Specific Model & Think Level Configuration
    # Each agent can have an independent model override and reasoning effort (think level)
    # Think level options: "default" (inherit/standard), "none" (reasoning disabled), "low", "medium", "high"
    AGENT_MODELS: dict = {
        aid: os.getenv(f"AGENT_MODEL_{aid.upper()}", "").strip()
        for aid in AGENTS_CONFIG_SCHEMA
    }
    AGENT_THINK_LEVELS: dict = {
        aid: os.getenv(f"AGENT_THINK_LEVEL_{aid.upper()}", "default").strip().lower()
        for aid in AGENTS_CONFIG_SCHEMA
    }
    AGENT_MODEL_GENERATOR: str = os.getenv("AGENT_MODEL_GENERATOR", "")
    AGENT_THINK_LEVEL_GENERATOR: str = os.getenv("AGENT_THINK_LEVEL_GENERATOR", "default")
    AGENT_MODEL_VERIFIER: str = os.getenv("AGENT_MODEL_VERIFIER", "")
    AGENT_THINK_LEVEL_VERIFIER: str = os.getenv("AGENT_THINK_LEVEL_VERIFIER", "default")
    AGENT_MODEL_AUTONOMOUS: str = os.getenv("AGENT_MODEL_AUTONOMOUS", "")
    AGENT_THINK_LEVEL_AUTONOMOUS: str = os.getenv("AGENT_THINK_LEVEL_AUTONOMOUS", "default")
    AGENT_MODEL_ANALYZER: str = os.getenv("AGENT_MODEL_ANALYZER", "")
    AGENT_THINK_LEVEL_ANALYZER: str = os.getenv("AGENT_THINK_LEVEL_ANALYZER", "default")
    AGENT_MODEL_STRUCTURE: str = os.getenv("AGENT_MODEL_STRUCTURE", "")
    AGENT_THINK_LEVEL_STRUCTURE: str = os.getenv("AGENT_THINK_LEVEL_STRUCTURE", "default")
    AGENT_MODEL_EDITOR: str = os.getenv("AGENT_MODEL_EDITOR", "")
    AGENT_THINK_LEVEL_EDITOR: str = os.getenv("AGENT_THINK_LEVEL_EDITOR", "default")
    AGENT_MODEL_PARSER: str = os.getenv("AGENT_MODEL_PARSER", "")
    AGENT_THINK_LEVEL_PARSER: str = os.getenv("AGENT_THINK_LEVEL_PARSER", "default")

    # Render Engine Cascade Configuration
    # Options: "auto" (Native COM -> Web Engine -> PIL), "native", "web", "pil"
    RENDER_MODE: str = os.getenv("RENDER_MODE", "auto").strip().lower()

    # If False (default is True), disables pure-Python PIL fallback and enforces native PowerPoint COM export.
    # If PowerPoint COM fails or is unavailable when PURE_PIL_ACTIVE is False, an error is raised.
    PURE_PIL_ACTIVE: bool = os.getenv("PURE_PIL_ACTIVE", "true").strip().lower() in ("1", "true", "yes", "on")

    # Step 4.5 Slide Verification Configuration (Number of iterative re-verification rounds)
    VERIFICATION_ROUNDS: int = int(os.getenv("VERIFICATION_ROUNDS", os.getenv("VERIFICATION_MAX_ROUNDS", "3")))

    _models_cache: dict = {}
    _models_cache_time: float = 0.0
    _category_cache: dict = {}

    @classmethod
    def get_clean_9router_url(cls, raw_url: Optional[str] = None) -> str:
        """Returns base 9Router gateway root URL without trailing slash or /v1 suffix."""
        url = (raw_url or cls.NINEROUTER_URL or "http://localhost:20128").strip().rstrip('/')
        if url.endswith('/v1'):
            url = url[:-3].rstrip('/')
        return url

    @classmethod
    def get_openai_base_url(cls) -> str:
        """Returns normalized base URL for OpenAI client, guaranteeing a clean single /v1 endpoint."""
        clean = cls.get_clean_9router_url()
        return f"{clean}/v1"

    @classmethod
    def get_9router_models(
        cls,
        category: str = "chat",
        base_url: Optional[str] = None,
        api_key: Optional[str] = None,
        force_refresh: bool = False
    ) -> dict:
        """
        Queries 9Router discovery endpoints (/v1/models, /v1/models/image, /v1/models/web)
        to discover models, providers, and capabilities.
        Caches responses for 60 seconds per endpoint.
        """
        import time
        import urllib.request
        import json

        clean_url = cls.get_clean_9router_url(base_url)
        key = cls.NINEROUTER_KEY if api_key is None else api_key
        now = time.time()

        fallback_chat = [
            {"id": "ag/gemini-3.8-flash-low", "owned_by": "ag", "context_length": 1048576, "max_completion_tokens": 65536, "capabilities": {"vision": True, "tools": True, "reasoning": True}},
            {"id": "ag/gemini-3.8-flash-high", "owned_by": "ag", "context_length": 1048576, "max_completion_tokens": 65536, "capabilities": {"vision": True, "tools": True, "reasoning": True}},
            {"id": "ag/gemini-3.7-flash-high", "owned_by": "ag", "context_length": 1048576, "max_completion_tokens": 65536, "capabilities": {"vision": True, "tools": True, "reasoning": True}},
            {"id": "light-code", "owned_by": "combo", "capabilities": {"tools": True}},
            {"id": "light-power", "owned_by": "combo", "capabilities": {"tools": True}},
            {"id": "claude-pro-agent", "owned_by": "combo", "capabilities": {"tools": True, "reasoning": True}},
            {"id": "openai/gpt-4o", "owned_by": "openai", "context_length": 128000, "max_completion_tokens": 16384, "capabilities": {"vision": True, "tools": True}},
            {"id": "anthropic/claude-3-7-sonnet", "owned_by": "anthropic", "context_length": 200000, "max_completion_tokens": 64000, "capabilities": {"vision": True, "tools": True, "reasoning": True}},
            {"id": "aval/deepseek-v4.1-flash", "owned_by": "aval", "capabilities": {"tools": True, "reasoning": True}},
        ]
        fallback_image = [
            {"id": "gemini/gemini-3.1-flash-image-preview", "owned_by": "gemini"},
            {"id": "gemini/gemini-3-pro-image-preview", "owned_by": "gemini"},
            {"id": "ag/gemini-3.1-flash-image", "owned_by": "ag"},
            {"id": "gemini/gemini-2.5-flash-image", "owned_by": "gemini"},
        ]
        fallback_web = [
            {"id": "exa/search", "owned_by": "exa", "kind": "webSearch"},
            {"id": "exa/fetch", "owned_by": "exa", "kind": "webFetch"},
            {"id": "tavily", "owned_by": "tavily", "kind": "webSearch"},
            {"id": "jina-reader", "owned_by": "jina", "kind": "webFetch"},
            {"id": "brave-search", "owned_by": "brave", "kind": "webSearch"},
            {"id": "firecrawl", "owned_by": "firecrawl", "kind": "webFetch"},
        ]

        def _fetch_endpoint(path: str, fallback: list) -> tuple[list, bool]:
            endpoint = f"{clean_url}{path}"
            cache_key = f"{clean_url}:{path}"
            if not force_refresh and hasattr(cls, "_category_cache"):
                cached = cls._category_cache.get(cache_key)
                if cached and (now - cached[1] < 60):
                    return cached[0], True

            headers = {}
            if key:
                headers["Authorization"] = f"Bearer {key}"

            try:
                req = urllib.request.Request(endpoint, headers=headers)
                with urllib.request.urlopen(req, timeout=4.0) as resp:
                    payload = json.loads(resp.read().decode("utf-8"))
                    items = payload.get("data", [])
                    if not hasattr(cls, "_category_cache"):
                        cls._category_cache = {}
                    cls._category_cache[cache_key] = (items, now)
                    return items, True
            except Exception:
                return fallback, False

        provider_names = {
            "combo": "9Router Combo",
            "ag": "Google (AG)",
            "gemini": "Google Gemini",
            "openrouter": "OpenRouter",
            "openai": "OpenAI",
            "anthropic": "Anthropic",
            "deepseek": "DeepSeek",
            "groq": "Groq",
            "tavily": "Tavily",
            "exa": "Exa",
            "aval": "Aval",
            "jina": "Jina AI",
            "brave": "Brave",
            "firecrawl": "Firecrawl"
        }

        def _format_item(m: dict, default_category: str = "chat") -> dict:
            mid = m.get("id", "")
            raw_owner = m.get("owned_by")
            if not raw_owner:
                raw_owner = mid.split("/")[0] if "/" in mid else "other"
            raw_owner_str = str(raw_owner).strip().lower()
            caps = m.get("capabilities") or {}
            ctx = m.get("context_length") or caps.get("contextWindow")
            max_t = m.get("max_completion_tokens") or caps.get("maxOutput")
            kind = m.get("kind")
            is_combo = (raw_owner_str == "combo")

            return {
                "id": mid,
                "owned_by": raw_owner_str,
                "provider": provider_names.get(raw_owner_str, raw_owner_str.upper() if len(raw_owner_str) <= 4 else raw_owner_str.capitalize()),
                "context_length": ctx,
                "max_tokens": max_t,
                "capabilities": {
                    "vision": bool(caps.get("vision")),
                    "tools": bool(caps.get("tools")),
                    "reasoning": bool(caps.get("reasoning")),
                    "search": bool(caps.get("search"))
                },
                "is_combo": is_combo,
                "kind": kind
            }

        # Query endpoints
        raw_chat, connected_chat = _fetch_endpoint("/v1/models", fallback_chat)
        raw_image, _ = _fetch_endpoint("/v1/models/image", fallback_image)
        raw_web, _ = _fetch_endpoint("/v1/models/web", fallback_web)

        # Update models cache for get_model_metadata
        if connected_chat:
            cls._models_cache = {m.get("id"): m for m in raw_chat if m.get("id")}
            cls._models_cache_time = now

        fmt_chat = [_format_item(m, "chat") for m in raw_chat if m.get("id")]
        fmt_image = [_format_item(m, "image") for m in raw_image if m.get("id")]
        fmt_web = [_format_item(m, "web") for m in raw_web if m.get("id")]

        fmt_search = [m for m in fmt_web if m.get("kind") == "webSearch" or "search" in m["id"].lower()]
        fmt_fetch = [m for m in fmt_web if m.get("kind") == "webFetch" or "fetch" in m["id"].lower() or "reader" in m["id"].lower()]

        # Collect unique provider keys for chat models
        providers_set = set(m["owned_by"] for m in fmt_chat)
        # Prioritize key providers
        priority = ["combo", "ag", "gemini", "openai", "anthropic", "deepseek", "openrouter", "aval"]
        ordered_providers = [p for p in priority if p in providers_set]
        for p in sorted(providers_set):
            if p not in ordered_providers:
                ordered_providers.append(p)

        recommended = [
            "ag/gemini-3.8-flash-low",
            "ag/gemini-3.8-flash-high",
            "light-code",
            "claude-pro-agent",
            "openai/gpt-4o",
            "anthropic/claude-3-7-sonnet",
            "aval/deepseek-v4.1-flash"
        ]
        # Filter recommended to only those available or known
        avail_ids = set(m["id"] for m in fmt_chat)
        active_rec = [r for r in recommended if r in avail_ids] or recommended[:4]

        # Determine returned models based on requested category
        cat_lower = category.strip().lower()
        if cat_lower == "image":
            selected_models = fmt_image
        elif cat_lower == "web":
            selected_models = fmt_web
        elif cat_lower == "search":
            selected_models = fmt_search
        elif cat_lower == "fetch":
            selected_models = fmt_fetch
        else:
            selected_models = fmt_chat

        return {
            "success": True,
            "connected": connected_chat,
            "category": cat_lower,
            "total": len(selected_models),
            "models": selected_models,
            "providers": ordered_providers,
            "recommended": active_rec,
            "categories": {
                "chat": fmt_chat,
                "image": fmt_image,
                "web": fmt_web,
                "search": fmt_search,
                "fetch": fmt_fetch
            },
            "source": "9router" if connected_chat else "fallback",
            "endpoint": f"{clean_url}/v1/models"
        }

    @classmethod
    def get_model_metadata(cls, model_id: Optional[str] = None) -> dict:
        """
        Queries 9Router /v1/models to detect context length, max tokens, and capabilities.
        Caches results in memory for 60 seconds.
        """
        import time
        import urllib.request
        import json
        target_id = model_id or cls.NINEROUTER_CHAT_MODEL
        now = time.time()

        if not cls._models_cache or (now - cls._models_cache_time > 60):
            try:
                clean_url = cls.get_clean_9router_url()
                url = f"{clean_url}/v1/models"
                headers = {}
                if cls.NINEROUTER_KEY:
                    headers["Authorization"] = f"Bearer {cls.NINEROUTER_KEY}"
                req = urllib.request.Request(url, headers=headers)
                with urllib.request.urlopen(req, timeout=5.0) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                    cls._models_cache = {m.get("id"): m for m in data.get("data", []) if m.get("id")}
                    cls._models_cache_time = now
            except Exception:
                pass

        model_obj = cls._models_cache.get(target_id, {})
        capabilities = model_obj.get("capabilities", {})

        max_output = (
            model_obj.get("max_completion_tokens")
            or capabilities.get("maxOutput")
            or 8192
        )
        context_window = (
            model_obj.get("context_length")
            or capabilities.get("contextWindow")
            or 128000
        )

        return {
            "model_id": target_id,
            "max_tokens": max_output,
            "context_length": context_window,
            "capabilities": capabilities,
            "owned_by": model_obj.get("owned_by", "unknown"),
            "detected": bool(model_obj)
        }

    @classmethod
    def get_agent_model(cls, agent_id: str) -> str:
        """
        Returns configured model for the specified agent.
        Falls back to NINEROUTER_CHAT_MODEL if unset, empty, or 'default'.
        """
        aid = str(agent_id or "").strip().lower()
        mod = cls.AGENT_MODELS.get(aid, "")
        if not mod:
            attr = getattr(cls, f"AGENT_MODEL_{aid.upper()}", "")
            if attr:
                mod = attr
        mod = str(mod).strip()
        if mod and mod.lower() not in ("default", "inherit", "none", ""):
            return mod
        return cls.NINEROUTER_CHAT_MODEL

    @classmethod
    def get_agent_think_level(cls, agent_id: str) -> str:
        """
        Returns configured think level for the specified agent:
        'default', 'none', 'low', 'medium', 'high'.
        """
        aid = str(agent_id or "").strip().lower()
        lvl = cls.AGENT_THINK_LEVELS.get(aid, "")
        if not lvl:
            attr = getattr(cls, f"AGENT_THINK_LEVEL_{aid.upper()}", "default")
            if attr:
                lvl = attr
        lvl = str(lvl).strip().lower()
        if lvl in ("none", "off", "0"):
            return "none"
        if lvl in ("low", "medium", "high"):
            return lvl
        return "default"

    @classmethod
    def get_agent_llm_kwargs(cls, agent_id: str, **kwargs) -> dict:
        """
        Prepares keyword arguments for client.chat.completions.create with per-agent model and think level.
        Preserves caller-provided model and reasoning_effort if explicitly passed.
        """
        call_kwargs = dict(kwargs)
        if "model" not in call_kwargs or not call_kwargs["model"]:
            call_kwargs["model"] = cls.get_agent_model(agent_id)

        think_level = cls.get_agent_think_level(agent_id)
        if think_level in ("low", "medium", "high") and "reasoning_effort" not in call_kwargs:
            call_kwargs["reasoning_effort"] = think_level
        return call_kwargs

    @classmethod
    def safe_chat_completion(cls, client, agent_id: str, **kwargs):
        """
        Executes client.chat.completions.create with agent-specific model and think level.
        If the target model/gateway rejects reasoning_effort, gracefully retries without it.
        """
        call_kwargs = cls.get_agent_llm_kwargs(agent_id, **kwargs)
        try:
            return client.chat.completions.create(**call_kwargs)
        except Exception as e:
            err_msg = str(e).lower()
            if "reasoning_effort" in call_kwargs and (
                "reasoning_effort" in err_msg
                or "extra forbidden" in err_msg
                or "unknown parameter" in err_msg
                or "unrecognized" in err_msg
                or "unexpected keyword" in err_msg
            ):
                call_kwargs.pop("reasoning_effort", None)
                return client.chat.completions.create(**call_kwargs)
            raise

    @classmethod
    def get_agents_config(cls) -> dict:
        """Returns structured dictionary of all agents and their current configuration state."""
        result = {}
        for aid, meta in AGENTS_CONFIG_SCHEMA.items():
            raw_model = cls.AGENT_MODELS.get(aid, "")
            if not raw_model:
                raw_model = getattr(cls, f"AGENT_MODEL_{aid.upper()}", "")
            raw_model = str(raw_model).strip()

            raw_think = cls.AGENT_THINK_LEVELS.get(aid, "")
            if not raw_think:
                raw_think = getattr(cls, f"AGENT_THINK_LEVEL_{aid.upper()}", "default")
            raw_think = str(raw_think).strip().lower()

            effective_model = cls.get_agent_model(aid)
            think_level = cls.get_agent_think_level(aid)
            is_custom = bool(raw_model and raw_model.lower() not in ("default", "inherit", "none", ""))

            result[aid] = {
                "id": aid,
                "name": meta["name"],
                "role": meta["role"],
                "icon": meta["icon"],
                "category": meta["category"],
                "configured_model": raw_model,
                "effective_model": effective_model,
                "is_custom_model": is_custom,
                "think_level": think_level
            }
        return result
    
    @classmethod
    def reload(cls, dotenv_path: Optional[Path] = None):
        target_env = dotenv_path or (BASE_DIR / ".env")
        if target_env.exists():
            load_dotenv(dotenv_path=target_env, override=True)
        else:
            load_dotenv(override=True)
        cls.NINEROUTER_URL = os.getenv("NINEROUTER_URL", "http://localhost:20128")
        cls.NINEROUTER_KEY = os.getenv("NINEROUTER_KEY", "")
        cls.NINEROUTER_CHAT_MODEL = os.getenv("NINEROUTER_CHAT_MODEL", "ag/gemini-3.7-flash-high")
        cls.NINEROUTER_SEARCH_MODEL = os.getenv("NINEROUTER_SEARCH_MODEL", "tavily")
        cls.NINEROUTER_FETCH_MODEL = os.getenv("NINEROUTER_FETCH_MODEL", "jina-reader")
        cls.NINEROUTER_IMAGE_MODEL = os.getenv("NINEROUTER_IMAGE_MODEL", "gemini/gemini-3-pro-image-preview")
        cls.LLM_TIMEOUT = float(os.getenv("LLM_TIMEOUT", "300"))
        cls.RENDER_MODE = os.getenv("RENDER_MODE", "auto").strip().lower()
        cls.PURE_PIL_ACTIVE = os.getenv("PURE_PIL_ACTIVE", "true").strip().lower() in ("1", "true", "yes", "on")
        cls.VERIFICATION_ROUNDS = int(os.getenv("VERIFICATION_ROUNDS", os.getenv("VERIFICATION_MAX_ROUNDS", "3")))
        cls.AGENT_MODELS = {}
        cls.AGENT_THINK_LEVELS = {}
        for aid in AGENTS_CONFIG_SCHEMA:
            mod_val = os.getenv(f"AGENT_MODEL_{aid.upper()}", "").strip()
            thk_val = os.getenv(f"AGENT_THINK_LEVEL_{aid.upper()}", "default").strip().lower()
            cls.AGENT_MODELS[aid] = mod_val
            cls.AGENT_THINK_LEVELS[aid] = thk_val
            setattr(cls, f"AGENT_MODEL_{aid.upper()}", mod_val)
            setattr(cls, f"AGENT_THINK_LEVEL_{aid.upper()}", thk_val)
        cls._models_cache = {}
        cls._models_cache_time = 0.0
        cls._category_cache = {}
