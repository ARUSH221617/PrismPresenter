import os
import json
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional
from pptx_jahat.config import BASE_DIR, DATA_DIR

def _resolve_safe_path(path_str: str) -> Path:
    p = Path(path_str)
    if not p.is_absolute():
        p = (BASE_DIR / p).resolve()
    else:
        p = p.resolve()
    return p

def read_file(file_path: str) -> str:
    path = _resolve_safe_path(file_path)
    if not path.exists():
        return f"Error: File '{file_path}' does not exist."
    if path.is_dir():
        return f"Error: '{file_path}' is a directory, not a file."
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return f.read()
    except Exception as e:
        return f"Error reading file: {str(e)}"

def write_file(file_path: str, content: str) -> str:
    path = _resolve_safe_path(file_path)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(content)
        return f"Successfully wrote {len(content)} characters to {file_path}"
    except Exception as e:
        return f"Error writing file: {str(e)}"

def edit_file(file_path: str, old_text: str, new_text: str) -> str:
    path = _resolve_safe_path(file_path)
    if not path.exists():
        return f"Error: File '{file_path}' does not exist."
    try:
        with open(path, "r", encoding="utf-8") as f:
            content = f.read()
        if old_text not in content:
            return f"Error: old_text not found in {file_path}"
        new_content = content.replace(old_text, new_text, 1)
        with open(path, "w", encoding="utf-8") as f:
            f.write(new_content)
        return f"Successfully edited {file_path}"
    except Exception as e:
        return f"Error editing file: {str(e)}"

def list_dir(dir_path: str = "") -> str:
    path = _resolve_safe_path(dir_path) if dir_path else BASE_DIR
    if not path.exists():
        return f"Error: Directory '{dir_path}' does not exist."
    if not path.is_dir():
        return f"Error: '{dir_path}' is not a directory."
    try:
        entries = []
        for item in sorted(path.iterdir()):
            prefix = "[DIR] " if item.is_dir() else "[FILE]"
            entries.append(f"{prefix} {item.name}")
        return "\n".join(entries) if entries else "(Empty directory)"
    except Exception as e:
        return f"Error listing directory: {str(e)}"

def make_dir(dir_path: str) -> str:
    path = _resolve_safe_path(dir_path)
    try:
        path.mkdir(parents=True, exist_ok=True)
        return f"Directory created: {dir_path}"
    except Exception as e:
        return f"Error creating directory: {str(e)}"

def delete_file(file_path: str) -> str:
    path = _resolve_safe_path(file_path)
    if not path.exists():
        return f"Error: Path '{file_path}' does not exist."
    try:
        if path.is_dir():
            shutil.rmtree(path)
            return f"Directory deleted: {file_path}"
        else:
            path.unlink()
            return f"File deleted: {file_path}"
    except Exception as e:
        return f"Error deleting: {str(e)}"
