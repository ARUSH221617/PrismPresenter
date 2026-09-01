import os
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

# Ensure essential directories exist
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
COMPONENTS_DIR.mkdir(parents=True, exist_ok=True)
SHAPES_DIR.mkdir(parents=True, exist_ok=True)
IMAGES_DIR.mkdir(parents=True, exist_ok=True)

class Config:
    # 9Router Configuration
    NINEROUTER_URL: str = os.getenv("NINEROUTER_URL", "http://localhost:20128")
    NINEROUTER_KEY: str = os.getenv("NINEROUTER_KEY", "")
    NINEROUTER_CHAT_MODEL: str = os.getenv("NINEROUTER_CHAT_MODEL", "ag/gemini-3.7-flash-high")
    NINEROUTER_SEARCH_MODEL: str = os.getenv("NINEROUTER_SEARCH_MODEL", "tavily")
    NINEROUTER_FETCH_MODEL: str = os.getenv("NINEROUTER_FETCH_MODEL", "jina-reader")
    NINEROUTER_IMAGE_MODEL: str = os.getenv("NINEROUTER_IMAGE_MODEL", "gemini/gemini-3-pro-image-preview")

    # Render Engine Cascade Configuration
    # Options: "auto" (Native COM -> Web Engine -> PIL), "native", "web", "pil"
    RENDER_MODE: str = os.getenv("RENDER_MODE", "auto").strip().lower()

    # If False (default is True), disables pure-Python PIL fallback and enforces native PowerPoint COM export.
    # If PowerPoint COM fails or is unavailable when PURE_PIL_ACTIVE is False, an error is raised.
    PURE_PIL_ACTIVE: bool = os.getenv("PURE_PIL_ACTIVE", "true").strip().lower() in ("1", "true", "yes", "on")
    
    @classmethod
    def reload(cls):
        load_dotenv(override=True)
        cls.NINEROUTER_URL = os.getenv("NINEROUTER_URL", "http://localhost:20128")
        cls.NINEROUTER_KEY = os.getenv("NINEROUTER_KEY", "")
        cls.NINEROUTER_CHAT_MODEL = os.getenv("NINEROUTER_CHAT_MODEL", "ag/gemini-3.7-flash-high")
        cls.NINEROUTER_SEARCH_MODEL = os.getenv("NINEROUTER_SEARCH_MODEL", "tavily")
        cls.NINEROUTER_FETCH_MODEL = os.getenv("NINEROUTER_FETCH_MODEL", "jina-reader")
        cls.NINEROUTER_IMAGE_MODEL = os.getenv("NINEROUTER_IMAGE_MODEL", "gemini/gemini-3-pro-image-preview")
        cls.RENDER_MODE = os.getenv("RENDER_MODE", "auto").strip().lower()
        cls.PURE_PIL_ACTIVE = os.getenv("PURE_PIL_ACTIVE", "true").strip().lower() in ("1", "true", "yes", "on")
