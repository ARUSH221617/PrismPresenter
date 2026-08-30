import requests
from typing import Dict, Any, List, Optional
from pptx_jahat.config import Config

def search_web(query: str, num_results: int = 5) -> str:
    if not Config.NINEROUTER_URL:
        return "Error: NINEROUTER_URL is not configured."
        
    url = f"{Config.NINEROUTER_URL.rstrip('/')}/v1/search"
    headers = {"Content-Type": "application/json"}
    if Config.NINEROUTER_KEY:
        headers["Authorization"] = f"Bearer {Config.NINEROUTER_KEY}"
    payload = {
        "model": Config.NINEROUTER_SEARCH_MODEL or "tavily",
        "query": query,
        "max_results": num_results
    }
    try:
        res = requests.post(url, json=payload, headers=headers, timeout=20)
        if res.status_code != 200:
            return f"9Router Search Error {res.status_code}: {res.text}"
        data = res.json()
        results = data.get("results", [])
        if not results:
            return f"No search results found for query: '{query}'"
        
        output = []
        for i, item in enumerate(results, 1):
            title = item.get("title", "No Title")
            url_str = item.get("url", "")
            snippet = (item.get("snippet") or item.get("content") or "")[:400].replace("\n", " ")
            output.append(f"{i}. [{title}]({url_str})\n   {snippet}\n")
        return "\n".join(output)
    except Exception as e:
        return f"9Router search request failed: {str(e)}"

def fetch_page_content(url: str) -> str:
    if not Config.NINEROUTER_URL:
        return "Error: NINEROUTER_URL is not configured."
        
    fetch_url = f"{Config.NINEROUTER_URL.rstrip('/')}/v1/web/fetch"
    headers = {"Content-Type": "application/json"}
    if Config.NINEROUTER_KEY:
        headers["Authorization"] = f"Bearer {Config.NINEROUTER_KEY}"
    payload = {
        "model": Config.NINEROUTER_FETCH_MODEL or "jina-reader",
        "url": url,
        "format": "markdown",
        "max_characters": 5000
    }
    try:
        res = requests.post(fetch_url, json=payload, headers=headers, timeout=20)
        if res.status_code == 200:
            data = res.json()
            content_obj = data.get("content", {})
            if isinstance(content_obj, dict):
                return content_obj.get("text", "")[:4000]
            elif isinstance(content_obj, str):
                return content_obj[:4000]
        return f"9Router fetch returned status {res.status_code}"
    except Exception as e:
        return f"9Router fetch failed: {str(e)}"
