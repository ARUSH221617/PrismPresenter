import json
from typing import Dict, Any, List, Callable, Optional
from openai import OpenAI

from pptx_jahat.config import Config
from pptx_jahat.tools.filesystem import (
    read_file, write_file, edit_file, list_dir, make_dir, delete_file
)
from pptx_jahat.tools.exa_search import search_web, fetch_page_content
from pptx_jahat.tools.pptx_engine import extract_all_templates, get_components_catalog
from pptx_jahat.tools.pptx_builder import (
    build_pptx_with_agent,
    verify_pptx_integrity,
    repair_pptx_package,
    verify_and_auto_heal_pptx
)
from pptx_jahat.tools.image_gen import generate_image

TOOLS_DEFINITIONS = [
    {
        "type": "function",
        "function": {
            "name": "read_file",
            "description": "Read content from a file in the workspace or system",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Relative or absolute path of the file"}
                },
                "required": ["file_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "write_file",
            "description": "Write or overwrite content to a file",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Path to write to"},
                    "content": {"type": "string", "description": "Full file content"}
                },
                "required": ["file_path", "content"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "edit_file",
            "description": "Replace a specific substring in a file with new content",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Path to target file"},
                    "old_text": {"type": "string", "description": "Exact text snippet to replace"},
                    "new_text": {"type": "string", "description": "Replacement text snippet"}
                },
                "required": ["file_path", "old_text", "new_text"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "list_dir",
            "description": "List files and subdirectories inside a directory",
            "parameters": {
                "type": "object",
                "properties": {
                    "dir_path": {"type": "string", "description": "Directory path (leave empty for root)"}
                }
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "make_dir",
            "description": "Create a new directory recursively",
            "parameters": {
                "type": "object",
                "properties": {
                    "dir_path": {"type": "string", "description": "Directory path to create"}
                },
                "required": ["dir_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "delete_file",
            "description": "Delete a file or directory",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Path to delete"}
                },
                "required": ["file_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "search_web",
            "description": "Search the web using Exa search engine",
            "parameters": {
                "type": "object",
                "properties": {
                    "query": {"type": "string", "description": "Search query"},
                    "num_results": {"type": "integer", "description": "Number of results (default 5)"}
                },
                "required": ["query"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "fetch_page_content",
            "description": "Fetch text content from a given web URL",
            "parameters": {
                "type": "object",
                "properties": {
                    "url": {"type": "string", "description": "URL to scrape"}
                },
                "required": ["url"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "extract_all_templates",
            "description": "Scan data/*.pptx templates, extract shapes, styles & save catalog to data/components/components.json",
            "parameters": {
                "type": "object",
                "properties": {}
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "get_components_catalog",
            "description": "Get current JSON catalog of extracted shapes, styles, and template blueprints",
            "parameters": {
                "type": "object",
                "properties": {}
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "build_pptx_from_docx",
            "description": "Generate a new PPTX deck from a Word (.docx) document using multi-template scanning, visual slide screenshots, and 9Router Vision AI reasoning",
            "parameters": {
                "type": "object",
                "properties": {
                    "docx_path": {"type": "string", "description": "Path to input docx file"},
                    "output_path": {"type": "string", "description": "Optional output pptx path"},
                    "template_name": {"type": "string", "description": "Optional template filename from data folder (defaults to global multi-template scan)"}
                },
                "required": ["docx_path"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "generate_image",
            "description": "Generate an AI visual/illustration using DALL-E / image model for slides",
            "parameters": {
                "type": "object",
                "properties": {
                    "prompt": {"type": "string", "description": "Detailed image prompt"},
                    "output_name": {"type": "string", "description": "Optional filename to save in data/components/images"}
                },
                "required": ["prompt"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "verify_and_repair_pptx",
            "description": "Check a PowerPoint (.pptx) file for XML corruptions, duplicate package parts, or formatting errors and auto-repair it",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_path": {"type": "string", "description": "Path to the PPTX file to verify and fix"}
                },
                "required": ["file_path"]
            }
        }
    }
]

TOOL_HANDLERS: Dict[str, Callable] = {
    "read_file": read_file,
    "write_file": write_file,
    "edit_file": edit_file,
    "list_dir": list_dir,
    "make_dir": make_dir,
    "delete_file": delete_file,
    "search_web": search_web,
    "fetch_page_content": fetch_page_content,
    "extract_all_templates": lambda: json.dumps(extract_all_templates(), indent=2),
    "get_components_catalog": lambda: json.dumps(get_components_catalog(), indent=2),
    "build_pptx_from_docx": build_pptx_with_agent,
    "generate_image": generate_image,
    "verify_and_repair_pptx": lambda file_path: json.dumps({
        "is_valid": verify_and_auto_heal_pptx(file_path)[0],
        "file_path": str(file_path),
        "status": "Verified OK" if verify_and_auto_heal_pptx(file_path)[0] else "Repairs applied"
    }),
}

class AIAgent:
    def __init__(self, system_prompt: Optional[str] = None):
        self.system_prompt = system_prompt or (
            "You are PPTX Jahat Autonomous AI Agent. "
            "You have full access to workspace files, web search, PPTX extraction & component analysis, "
            "Word docx parsing, exact PPTX template slide cloning with AI text replacement, and Image generation. "
            "Execute tasks directly using your tool suite."
        )
        self.messages: List[Dict[str, Any]] = [
            {"role": "system", "content": self.system_prompt}
        ]

    def _get_client(self) -> OpenAI:
        base_url = f"{Config.NINEROUTER_URL.rstrip('/')}/v1"
        api_key = Config.NINEROUTER_KEY or "dummy_key"
        return OpenAI(api_key=api_key, base_url=base_url)

    def run(self, user_prompt: str, max_steps: int = 10, log_callback: Optional[Callable[[str], None]] = None) -> str:
        def log(msg: str):
            if log_callback:
                log_callback(msg)
                
        self.messages.append({"role": "user", "content": user_prompt})
        client = self._get_client()

        for step in range(max_steps):
            try:
                log(f"[Agent Step {step+1}] Calling 9Router model '{Config.NINEROUTER_CHAT_MODEL}' at {Config.NINEROUTER_URL}...")
                response = client.chat.completions.create(
                    model=Config.NINEROUTER_CHAT_MODEL,
                    messages=self.messages,
                    tools=TOOLS_DEFINITIONS,
                    tool_choice="auto",
                    temperature=0.2
                )
            except Exception as e:
                err_msg = f"9Router LLM API Call Error: {str(e)}"
                log(err_msg)
                return err_msg

            msg = response.choices[0].message
            self.messages.append(msg)

            if msg.tool_calls:
                for tool_call in msg.tool_calls:
                    fn_name = tool_call.function.name
                    args_str = tool_call.function.arguments
                    try:
                        args = json.loads(args_str) if args_str else {}
                    except Exception:
                        args = {}
                    
                    log(f"-> Executing Tool: {fn_name}({args})")
                    handler = TOOL_HANDLERS.get(fn_name)
                    if handler:
                        try:
                            result = handler(**args)
                            result_str = str(result)
                        except Exception as ex:
                            result_str = f"Execution Error: {str(ex)}"
                    else:
                        result_str = f"Error: Tool '{fn_name}' not found."

                    self.messages.append({
                        "role": "tool",
                        "tool_call_id": tool_call.id,
                        "name": fn_name,
                        "content": result_str
                    })
            else:
                final_content = msg.content or "(Done)"
                log(f"[Agent Completed]: {final_content[:150]}")
                return final_content

        return "Agent stopped: Reached maximum allowable tool steps."
