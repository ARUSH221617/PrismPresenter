import os
from typing import Optional
from pathlib import Path
from dotenv import load_dotenv

# Load .env file from project root or current working dir
load_dotenv()

BASE_DIR = Path(__file__).resolve().parent.parent.parent
DATA_DIR = BASE_DIR / "data"
OUTPUT_DIR = DATA_DIR / "output"
COMPONENTS_DIR = DATA_DIR / "components"
SHAPES_DIR = COMPONENTS_DIR / "shapes"
IMAGES_DIR = COMPONENTS_DIR / "images"
STRUCTURES_DIR = DATA_DIR / "structure"

# Ensure essential directories exist
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
COMPONENTS_DIR.mkdir(parents=True, exist_ok=True)
SHAPES_DIR.mkdir(parents=True, exist_ok=True)
IMAGES_DIR.mkdir(parents=True, exist_ok=True)
STRUCTURES_DIR.mkdir(parents=True, exist_ok=True)

class Config:
    # 9Router Configuration
    NINEROUTER_URL: str = os.getenv("NINEROUTER_URL", "http://localhost:20128")
    NINEROUTER_KEY: str = os.getenv("NINEROUTER_KEY", "")
    NINEROUTER_CHAT_MODEL: str = os.getenv("NINEROUTER_CHAT_MODEL", "ag/gemini-3.7-flash-high")
    NINEROUTER_SEARCH_MODEL: str = os.getenv("NINEROUTER_SEARCH_MODEL", "tavily")
    NINEROUTER_FETCH_MODEL: str = os.getenv("NINEROUTER_FETCH_MODEL", "jina-reader")
    NINEROUTER_IMAGE_MODEL: str = os.getenv("NINEROUTER_IMAGE_MODEL", "gemini/gemini-3-pro-image-preview")
    LLM_TIMEOUT: float = float(os.getenv("LLM_TIMEOUT", "300"))

    # Render Engine Cascade Configuration
    # Options: "auto" (Native COM -> Web Engine -> PIL), "native", "web", "pil"
    RENDER_MODE: str = os.getenv("RENDER_MODE", "auto").strip().lower()

    # If False (default is True), disables pure-Python PIL fallback and enforces native PowerPoint COM export.
    # If PowerPoint COM fails or is unavailable when PURE_PIL_ACTIVE is False, an error is raised.
    PURE_PIL_ACTIVE: bool = os.getenv("PURE_PIL_ACTIVE", "true").strip().lower() in ("1", "true", "yes", "on")

    _models_cache: dict = {}
    _models_cache_time: float = 0.0

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
                url = f"{cls.NINEROUTER_URL.rstrip('/')}/v1/models"
                req = urllib.request.Request(url, headers={"Authorization": f"Bearer {cls.NINEROUTER_KEY}"})
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
    def reload(cls):
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
