import re
import json
from typing import Any, Optional

def safe_json_loads(text: str, default: Optional[Any] = None) -> Any:
    """
    Robust JSON parser for LLM responses.
    Handles:
    - Markdown code fences (```json ... ```)
    - Unescaped LaTeX backslashes (e.g. \\cdot, \\times, \\sqrt, \\frac, \\pm, \\approx)
    - Trailing whitespace and non-standard JSON escapes from AI model outputs
    """
    if not text or not isinstance(text, str):
        return default if default is not None else {}

    # Extract JSON object or array from text
    match = re.search(r"```(?:json)?\s*(\{.*?\}|\[.*?\])\s*```", text, re.DOTALL)
    if match:
        raw = match.group(1).strip()
    else:
        obj_match = re.search(r"(\{.*\}|\[.*\])", text, re.DOTALL)
        raw = obj_match.group(1).strip() if obj_match else text.strip()

    # 1. Direct standard parse
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # 2. Fix unescaped backslashes that are invalid in JSON (e.g. \cdot, \sqrt, \alpha, \pm)
    fixed = re.sub(r'\\(?!["\\/bfnrt]|u[0-9a-fA-F]{4})', r'\\\\', raw)
    try:
        return json.loads(fixed)
    except json.JSONDecodeError:
        pass

    # 3. Fix LaTeX keywords that happen to start with valid escape chars (\times, \frac, \beta, \text)
    latex_fixed = re.sub(r'\\([a-zA-Z]{2,})', r'\\\\\1', raw)
    try:
        return json.loads(latex_fixed)
    except json.JSONDecodeError:
        pass

    # 4. Try lenient parsing with strict=False
    try:
        return json.loads(fixed, strict=False)
    except json.JSONDecodeError:
        pass

    try:
        return json.loads(raw, strict=False)
    except Exception:
        if default is not None:
            return default
        raise
