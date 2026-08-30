import requests
import uuid
from pathlib import Path
from typing import Optional
from pptx_jahat.config import Config, IMAGES_DIR

def generate_image(prompt: str, output_name: Optional[str] = None) -> str:
    if not Config.NINEROUTER_URL:
        return "Error: NINEROUTER_URL is not configured in .env"
        
    filename = output_name or f"gen_{uuid.uuid4().hex[:8]}.png"
    if not filename.endswith((".png", ".jpg", ".jpeg")):
        filename += ".png"
    save_path = IMAGES_DIR / filename

    endpoint = f"{Config.NINEROUTER_URL.rstrip('/')}/v1/images/generations?response_format=binary"
    headers = {"Content-Type": "application/json"}
    if Config.NINEROUTER_KEY:
        headers["Authorization"] = f"Bearer {Config.NINEROUTER_KEY}"
    payload = {
        "model": Config.NINEROUTER_IMAGE_MODEL or "gemini/gemini-3-pro-image-preview",
        "prompt": prompt,
        "size": "1024x1024"
    }
    try:
        res = requests.post(endpoint, json=payload, headers=headers, timeout=60)
        if res.status_code == 200 and len(res.content) > 100:
            with open(save_path, "wb") as f:
                f.write(res.content)
            return str(save_path)
        else:
            return f"9Router image generation error {res.status_code}: {res.text[:200]}"
    except Exception as e:
        return f"9Router image request failed: {str(e)}"
