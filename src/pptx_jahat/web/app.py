import os
import io
import json
import time
import queue
import shutil
import base64
import subprocess
import threading
from pathlib import Path
from typing import Dict, Any, Optional, List
from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    Response,
    send_file,
    send_from_directory
)
from flask_cors import CORS
from PIL import Image

from pptx_jahat.config import Config, DATA_DIR, OUTPUT_DIR, COMPONENTS_DIR
from pptx_jahat.agent import AIAgent
from pptx_jahat.tools.pptx_builder import build_pptx_with_agent, verify_and_auto_heal_pptx
from pptx_jahat.tools.preview import render_pptx_file_previews, image_to_base64_jpeg, image_to_base64_png
from pptx_jahat.tools.template_analyzer import (
    analyze_template,
    analyze_all_templates,
    load_notes,
    save_notes,
    get_analyzed_templates,
    NOTE_FILE
)
from pptx_jahat.tools.pptx_engine import extract_all_templates, get_components_catalog

# Global Job Registry for SSE streams
JOBS: Dict[str, Dict[str, Any]] = {}
JOBS_LOCK = threading.Lock()

def create_app() -> Flask:
    template_dir = Path(__file__).parent / "templates"
    static_dir = Path(__file__).parent / "static"

    app = Flask(
        __name__,
        template_folder=str(template_dir),
        static_folder=str(static_dir)
    )
    CORS(app)
    app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB upload limit

    # Upload cache dir
    UPLOAD_CACHE = DATA_DIR / "uploads"
    UPLOAD_CACHE.mkdir(parents=True, exist_ok=True)

    @app.route("/")
    def index():
        return render_template("index.html")

    @app.route("/assets/videos/<filename>")
    def serve_asset_video(filename: str):
        assets_video_dir = Path(__file__).resolve().parent.parent.parent.parent / "assets" / "videos"
        target = assets_video_dir / filename
        if not target.exists():
            return jsonify({"error": "Video not found"}), 404
        return send_file(target, mimetype="video/mp4")

    # -------------------------------------------------------------
    # 1. GENERATOR ENDPOINTS
    # -------------------------------------------------------------
    @app.route("/api/generator/templates", methods=["GET"])
    def list_generator_templates():
        templates = [p.name for p in DATA_DIR.glob("*.pptx") if not p.name.endswith("_generated.pptx")]
        return jsonify({
            "success": True,
            "templates": templates
        })

    @app.route("/api/generator/upload", methods=["POST"])
    def upload_docx():
        if "file" not in request.files:
            return jsonify({"success": False, "error": "No file part in request"}), 400
        file = request.files["file"]
        if not file.filename or not file.filename.endswith(".docx"):
            return jsonify({"success": False, "error": "Invalid file. Must be a .docx document."}), 400

        filename = Path(file.filename).name
        target_path = UPLOAD_CACHE / filename
        file.save(str(target_path))

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        suggested_output = str(OUTPUT_DIR / f"{Path(filename).stem}_generated.pptx")

        return jsonify({
            "success": True,
            "filename": filename,
            "file_path": str(target_path.resolve()),
            "suggested_output": suggested_output
        })

    @app.route("/api/generator/generate", methods=["POST"])
    def start_generation():
        data = request.get_json() or {}
        docx_path = data.get("docx_path", "").strip()
        template_name = data.get("template_name", None)
        output_path = data.get("output_path", "").strip()

        if not docx_path or not Path(docx_path).exists():
            return jsonify({"success": False, "error": "Invalid or missing Word document path."}), 400

        if template_name and ("All Templates" in template_name or "No templates" in template_name):
            template_name = None

        if not output_path:
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            output_path = str(OUTPUT_DIR / f"{Path(docx_path).stem}_generated.pptx")

        job_id = f"gen_{int(time.time() * 1000)}"
        event_queue = queue.Queue()

        with JOBS_LOCK:
            JOBS[job_id] = {
                "id": job_id,
                "type": "generator",
                "queue": event_queue,
                "status": "running",
                "result": None,
                "error": None,
                "ai_images": []
            }

        def worker():
            def log_callback(msg: str):
                event_queue.put({"event": "log", "data": {"message": msg, "time": time.strftime("%H:%M:%S")}})

            def on_ai_images_ready(sent_images: List[Dict[str, Any]]):
                parsed = []
                for item in sent_images:
                    parsed.append({
                        "base64": item.get("base64"),
                        "template_file": item.get("template_file", "Template"),
                        "slide_index": item.get("slide_index", 0),
                        "archetype": item.get("archetype", "Archetype")
                    })
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["ai_images"] = parsed
                event_queue.put({"event": "ai_images", "data": {"images": parsed}})

            try:
                event_queue.put({"event": "status", "data": {"status": "Generating presentation..."}})
                res = build_pptx_with_agent(
                    docx_path,
                    output_path,
                    template_name,
                    log_callback=log_callback,
                    on_ai_images_ready=on_ai_images_ready
                )

                # Pre-render slides for instant UI loading
                previews = []
                try:
                    imgs, engine_name = render_pptx_file_previews(res, target_width_px=800, return_engine_info=True)
                    for idx, img in enumerate(imgs):
                        previews.append({
                            "slide_index": idx,
                            "data_url": image_to_base64_jpeg(img, quality=85)
                        })
                except Exception as ex:
                    engine_name = "None"
                    log_callback(f"Preview render warning: {str(ex)}")

                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "completed"
                        JOBS[job_id]["result"] = res

                event_queue.put({
                    "event": "completed",
                    "data": {
                        "pptx_path": res,
                        "filename": Path(res).name,
                        "engine_name": engine_name,
                        "previews": previews
                    }
                })
            except Exception as e:
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "error"
                        JOBS[job_id]["error"] = str(e)
                event_queue.put({"event": "error", "data": {"error": str(e)}})
            finally:
                event_queue.put({"event": "close", "data": {}})

        threading.Thread(target=worker, daemon=True).start()

        return jsonify({
            "success": True,
            "job_id": job_id,
            "output_path": output_path
        })

    @app.route("/api/generator/stream/<job_id>", methods=["GET"])
    def stream_generation_job(job_id: str):
        with JOBS_LOCK:
            job = JOBS.get(job_id)

        if not job:
            return jsonify({"error": "Job not found"}), 404

        q = job["queue"]

        def event_stream():
            while True:
                try:
                    item = q.get(timeout=30.0)
                    evt = item.get("event", "message")
                    data = json.dumps(item.get("data", {}))
                    yield f"event: {evt}\ndata: {data}\n\n"
                    if evt == "close":
                        break
                except queue.Empty:
                    yield f"event: ping\ndata: {{}}\n\n"

        return Response(event_stream(), mimetype="text/event-stream")

    @app.route("/api/preview/render", methods=["POST"])
    def render_presentation_previews():
        data = request.get_json() or {}
        file_path = data.get("file_path", "").strip()

        if not file_path or not Path(file_path).exists():
            return jsonify({"success": False, "error": "File does not exist."}), 400

        try:
            target_width = int(data.get("width", 800))
            res = render_pptx_file_previews(file_path, target_width_px=target_width, return_engine_info=True)
            if isinstance(res, tuple):
                imgs, engine_name = res
            else:
                imgs, engine_name = res, "Renderer"

            previews = []
            for idx, img in enumerate(imgs):
                previews.append({
                    "slide_index": idx,
                    "data_url": image_to_base64_jpeg(img, quality=85)
                })

            return jsonify({
                "success": True,
                "engine_name": engine_name,
                "slide_count": len(previews),
                "slides": previews
            })
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/preview/slide-html", methods=["POST"])
    def render_presentation_html():
        from pptx_jahat.tools.renderers.web_renderer import render_pptx_to_html_deck
        data = request.get_json() or {}
        file_path = data.get("file_path", "").strip()

        if not file_path or not Path(file_path).exists():
            return jsonify({"success": False, "error": "File does not exist."}), 400

        try:
            target_width = int(data.get("width", 800))
            slides_html = render_pptx_to_html_deck(file_path, width=target_width)
            return jsonify({
                "success": True,
                "engine_name": "Web Render Engine",
                "slide_count": len(slides_html),
                "slides": slides_html
            })
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    # -------------------------------------------------------------
    # 2. TEMPLATE INTELLIGENCE & NOTE.md ENDPOINTS
    # -------------------------------------------------------------
    @app.route("/api/templates/list", methods=["GET"])
    def list_templates():
        from pptx import Presentation
        pptx_files = sorted(list(DATA_DIR.glob("*.pptx")))
        templates = [f for f in pptx_files if not f.name.endswith("_generated.pptx")]
        analyzed_map = get_analyzed_templates()

        items = []
        for tpl in templates:
            try:
                prs = Presentation(str(tpl))
                slide_count = len(prs.slides)
                width_in = round(prs.slide_width / 914400, 2)
                height_in = round(prs.slide_height / 914400, 2)
                dim_str = f"{width_in}\" x {height_in}\""
            except Exception:
                slide_count = 0
                dim_str = "Unknown"

            is_analyzed = tpl.name in analyzed_map
            info = analyzed_map.get(tpl.name, {})

            items.append({
                "filename": tpl.name,
                "file_path": str(tpl.resolve()),
                "slide_count": slide_count,
                "dimensions": dim_str,
                "is_analyzed": is_analyzed,
                "purpose": info.get("purpose", "Not analyzed"),
                "style": info.get("style", "Not analyzed"),
                "brief": info.get("brief", "")
            })

        return jsonify({
            "success": True,
            "templates": items,
            "total_count": len(items),
            "analyzed_count": sum(1 for i in items if i["is_analyzed"])
        })

    @app.route("/api/templates/notes", methods=["GET", "POST"])
    def handle_template_notes():
        if request.method == "GET":
            content = load_notes(NOTE_FILE)
            return jsonify({
                "success": True,
                "file_path": str(NOTE_FILE.resolve()),
                "content": content
            })
        else:
            data = request.get_json() or {}
            content = data.get("content", "")
            save_notes(content, NOTE_FILE)
            return jsonify({
                "success": True,
                "message": f"Successfully updated {NOTE_FILE.name}"
            })

    @app.route("/api/templates/analyze", methods=["POST"])
    def analyze_single_template():
        data = request.get_json() or {}
        filename = data.get("filename", "").strip()
        tpl_path = DATA_DIR / filename

        if not tpl_path.exists():
            return jsonify({"success": False, "error": f"Template '{filename}' not found."}), 404

        job_id = f"tpl_{int(time.time() * 1000)}"
        event_queue = queue.Queue()

        with JOBS_LOCK:
            JOBS[job_id] = {
                "id": job_id,
                "type": "template_analysis",
                "queue": event_queue,
                "status": "running"
            }

        def worker():
            def log_cb(msg: str):
                event_queue.put({"event": "log", "data": {"message": msg, "time": time.strftime("%H:%M:%S")}})

            try:
                log_cb(f"[*] Starting AI Template Analysis: {filename}")
                res = analyze_template(tpl_path, log_cb=log_cb, save_to_file=True)
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "completed"
                        JOBS[job_id]["result"] = res
                event_queue.put({"event": "completed", "data": {"result": res, "filename": filename}})
            except Exception as e:
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "error"
                event_queue.put({"event": "error", "data": {"error": str(e)}})
            finally:
                event_queue.put({"event": "close", "data": {}})

        threading.Thread(target=worker, daemon=True).start()

        return jsonify({"success": True, "job_id": job_id})

    @app.route("/api/templates/analyze-all", methods=["POST"])
    def analyze_all_templates_batch():
        job_id = f"tpl_all_{int(time.time() * 1000)}"
        event_queue = queue.Queue()

        with JOBS_LOCK:
            JOBS[job_id] = {
                "id": job_id,
                "type": "template_batch_analysis",
                "queue": event_queue,
                "status": "running"
            }

        def worker():
            def log_cb(msg: str):
                event_queue.put({"event": "log", "data": {"message": msg, "time": time.strftime("%H:%M:%S")}})

            def progress_cb(current: int, total: int, current_name: str):
                event_queue.put({
                    "event": "progress",
                    "data": {
                        "current": current,
                        "total": total,
                        "percentage": round((current / total) * 100, 1),
                        "current_name": current_name
                    }
                })

            try:
                res = analyze_all_templates(DATA_DIR, progress_cb=progress_cb, log_cb=log_cb)
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "completed"
                event_queue.put({"event": "completed", "data": {"notes": res}})
            except Exception as e:
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "error"
                event_queue.put({"event": "error", "data": {"error": str(e)}})
            finally:
                event_queue.put({"event": "close", "data": {}})

        threading.Thread(target=worker, daemon=True).start()

        return jsonify({"success": True, "job_id": job_id})

    # -------------------------------------------------------------
    # 3. DECK & TEMPLATE MANAGER ENDPOINTS
    # -------------------------------------------------------------
    @app.route("/api/manager/decks", methods=["GET"])
    def get_all_manager_decks():
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        DATA_DIR.mkdir(parents=True, exist_ok=True)

        generated = []
        for p in sorted(OUTPUT_DIR.glob("*.pptx"), key=lambda x: x.stat().st_mtime if x.exists() else 0, reverse=True):
            try:
                st = p.stat()
                size_str = f"{st.st_size / 1024:.1f} KB" if st.st_size < 1024*1024 else f"{st.st_size / (1024*1024):.2f} MB"
                mtime_str = time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime))
                generated.append({
                    "filename": p.name,
                    "file_path": str(p.resolve()),
                    "size": size_str,
                    "modified": mtime_str,
                    "type": "generated"
                })
            except Exception:
                pass

        reference = []
        for p in sorted([f for f in DATA_DIR.glob("*.pptx") if not f.name.endswith("_generated.pptx")], key=lambda x: x.name):
            try:
                st = p.stat()
                size_str = f"{st.st_size / 1024:.1f} KB" if st.st_size < 1024*1024 else f"{st.st_size / (1024*1024):.2f} MB"
                mtime_str = time.strftime("%Y-%m-%d %H:%M", time.localtime(st.st_mtime))
                reference.append({
                    "filename": p.name,
                    "file_path": str(p.resolve()),
                    "size": size_str,
                    "modified": mtime_str,
                    "type": "reference"
                })
            except Exception:
                pass

        return jsonify({
            "success": True,
            "generated": generated,
            "reference": reference
        })

    @app.route("/api/manager/upload-template", methods=["POST"])
    def upload_reference_template():
        if "file" not in request.files:
            return jsonify({"success": False, "error": "No file uploaded"}), 400
        file = request.files["file"]
        if not file.filename or not file.filename.endswith(".pptx"):
            return jsonify({"success": False, "error": "File must be a .pptx presentation."}), 400

        filename = Path(file.filename).name
        target = DATA_DIR / filename
        file.save(str(target))

        return jsonify({"success": True, "filename": filename, "file_path": str(target.resolve())})

    @app.route("/api/manager/verify", methods=["POST"])
    def verify_deck():
        data = request.get_json() or {}
        file_path = data.get("file_path", "")
        if not file_path or not Path(file_path).exists():
            return jsonify({"success": False, "error": "File does not exist."}), 400

        is_ok, final_p = verify_and_auto_heal_pptx(file_path)
        return jsonify({
            "success": True,
            "is_valid": is_ok,
            "final_path": final_p,
            "filename": Path(final_p).name
        })

    @app.route("/api/manager/duplicate", methods=["POST"])
    def duplicate_deck():
        data = request.get_json() or {}
        file_path = data.get("file_path", "")
        if not file_path or not Path(file_path).exists():
            return jsonify({"success": False, "error": "File does not exist."}), 400

        orig_p = Path(file_path)
        copy_p = orig_p.parent / f"{orig_p.stem}_copy{orig_p.suffix}"
        idx = 1
        while copy_p.exists():
            copy_p = orig_p.parent / f"{orig_p.stem}_copy{idx}{orig_p.suffix}"
            idx += 1

        shutil.copy2(orig_p, copy_p)
        return jsonify({
            "success": True,
            "filename": copy_p.name,
            "file_path": str(copy_p.resolve())
        })

    @app.route("/api/manager/rename", methods=["POST"])
    def rename_deck():
        data = request.get_json() or {}
        file_path = data.get("file_path", "")
        new_name = data.get("new_name", "").strip()

        if not file_path or not Path(file_path).exists():
            return jsonify({"success": False, "error": "File does not exist."}), 400

        if not new_name:
            return jsonify({"success": False, "error": "New name cannot be empty."}), 400

        if not new_name.endswith(".pptx"):
            new_name += ".pptx"

        orig_p = Path(file_path)
        target_p = orig_p.parent / new_name

        if target_p.exists() and target_p != orig_p:
            return jsonify({"success": False, "error": "A file with this name already exists."}), 400

        orig_p.rename(target_p)
        return jsonify({
            "success": True,
            "filename": target_p.name,
            "file_path": str(target_p.resolve())
        })

    @app.route("/api/manager/delete", methods=["DELETE"])
    def delete_deck():
        data = request.get_json() or {}
        file_path = data.get("file_path", "")
        if not file_path or not Path(file_path).exists():
            return jsonify({"success": False, "error": "File does not exist."}), 400

        Path(file_path).unlink()
        return jsonify({"success": True, "message": "File deleted successfully."})

    @app.route("/api/manager/open", methods=["POST"])
    def open_deck_in_host():
        data = request.get_json() or {}
        file_path = data.get("file_path", "")
        if not file_path or not Path(file_path).exists():
            return jsonify({"success": False, "error": "File does not exist."}), 400

        try:
            os.startfile(file_path)
            return jsonify({"success": True, "message": f"Opened {file_path}"})
        except Exception:
            try:
                subprocess.Popen(["start", "", file_path], shell=True)
                return jsonify({"success": True, "message": f"Launched {file_path}"})
            except Exception as e:
                return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/manager/download", methods=["GET"])
    def download_deck():
        file_path = request.args.get("file", "")
        if not file_path or not Path(file_path).exists():
            return jsonify({"error": "File not found"}), 404
        return send_file(file_path, as_attachment=True, download_name=Path(file_path).name)

    # -------------------------------------------------------------
    # 4. COMPONENTS CATALOG ENDPOINTS
    # -------------------------------------------------------------
    @app.route("/api/components/image/<filename>", methods=["GET"])
    def get_component_image(filename: str):
        target = IMAGES_DIR / filename
        if not target.exists():
            return jsonify({"error": "Image not found"}), 404
        return send_file(target)

    @app.route("/api/components/catalog", methods=["GET"])
    def get_components_data():
        catalog = get_components_catalog()
        count = len(catalog.get("all_components", []))
        return jsonify({
            "success": True,
            "catalog": catalog,
            "count": count
        })

    @app.route("/api/components/extract", methods=["POST"])
    def run_components_extraction():
        try:
            catalog = extract_all_templates()
            count = len(catalog.get("all_components", []))
            return jsonify({
                "success": True,
                "catalog": catalog,
                "count": count,
                "message": f"Extraction completed. Found {count} components."
            })
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    # -------------------------------------------------------------
    # 5. AUTONOMOUS AI AGENT ENDPOINTS
    # -------------------------------------------------------------
    @app.route("/api/agent/chat", methods=["POST"])
    def run_agent_chat():
        data = request.get_json() or {}
        prompt = data.get("prompt", "").strip()
        enable_search = bool(data.get("enable_search", True))
        enable_pptx_tools = bool(data.get("enable_pptx_tools", True))

        if not prompt:
            return jsonify({"success": False, "error": "Prompt cannot be empty."}), 400

        job_id = f"agent_{int(time.time() * 1000)}"
        event_queue = queue.Queue()

        with JOBS_LOCK:
            JOBS[job_id] = {
                "id": job_id,
                "type": "agent_chat",
                "queue": event_queue,
                "status": "running"
            }

        def worker():
            agent = AIAgent(
                enable_search=enable_search,
                enable_pptx_tools=enable_pptx_tools
            )

            def log_callback(msg: str):
                event_queue.put({"event": "log", "data": {"message": msg, "time": time.strftime("%H:%M:%S")}})

            try:
                log_callback(f"[USER PROMPT]: {prompt}")
                reply = agent.run(prompt, log_callback=log_callback)
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "completed"
                        JOBS[job_id]["result"] = reply
                event_queue.put({"event": "completed", "data": {"response": reply}})
            except Exception as e:
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "error"
                event_queue.put({"event": "error", "data": {"error": str(e)}})
            finally:
                event_queue.put({"event": "close", "data": {}})

        threading.Thread(target=worker, daemon=True).start()

        return jsonify({"success": True, "job_id": job_id})

    # -------------------------------------------------------------
    # 6. CONFIGURATION & SETTINGS ENDPOINTS
    # -------------------------------------------------------------
    @app.route("/api/config", methods=["GET", "POST"])
    def handle_config():
        if request.method == "GET":
            return jsonify({
                "success": True,
                "config": {
                    "NINEROUTER_URL": Config.NINEROUTER_URL,
                    "NINEROUTER_KEY": Config.NINEROUTER_KEY,
                    "NINEROUTER_CHAT_MODEL": Config.NINEROUTER_CHAT_MODEL,
                    "NINEROUTER_SEARCH_MODEL": Config.NINEROUTER_SEARCH_MODEL,
                    "NINEROUTER_FETCH_MODEL": Config.NINEROUTER_FETCH_MODEL,
                    "NINEROUTER_IMAGE_MODEL": Config.NINEROUTER_IMAGE_MODEL,
                    "RENDER_MODE": Config.RENDER_MODE,
                    "PURE_PIL_ACTIVE": Config.PURE_PIL_ACTIVE
                }
            })
        else:
            data = request.get_json() or {}
            cfg = data.get("config", {})

            env_lines = []
            for k in [
                "NINEROUTER_URL",
                "NINEROUTER_KEY",
                "NINEROUTER_CHAT_MODEL",
                "NINEROUTER_SEARCH_MODEL",
                "NINEROUTER_FETCH_MODEL",
                "NINEROUTER_IMAGE_MODEL",
                "RENDER_MODE",
                "PURE_PIL_ACTIVE"
            ]:
                if k in cfg:
                    val = str(cfg[k]).strip()
                    env_lines.append(f"{k}={val}")

            env_path = Path(__file__).resolve().parent.parent.parent.parent / ".env"
            if env_path.exists():
                try:
                    backup_path = env_path.with_suffix(".env.backup")
                    shutil.copy2(env_path, backup_path)
                except Exception:
                    pass

            with open(env_path, "w", encoding="utf-8") as f:
                f.write("\n".join(env_lines) + "\n")

            Config.reload()
            return jsonify({
                "success": True,
                "message": "Configuration saved to .env and reloaded.",
                "model": Config.NINEROUTER_CHAT_MODEL,
                "render_mode": Config.RENDER_MODE
            })

    @app.route("/api/config/ping", methods=["POST"])
    def ping_gateway():
        import requests
        data = request.get_json() or {}
        target_url = (data.get("url") or Config.NINEROUTER_URL or "http://localhost:20128").strip().rstrip("/")
        api_key = (data.get("key") if "key" in data else Config.NINEROUTER_KEY).strip()

        if not target_url:
            return jsonify({"success": False, "error": "Gateway URL cannot be empty."})

        # Ensure protocol
        if not target_url.startswith("http://") and not target_url.startswith("https://"):
            target_url = f"http://{target_url}"

        headers = {}
        if api_key:
            headers["Authorization"] = f"Bearer {api_key}"

        start_time = time.time()
        try:
            test_url = f"{target_url}/v1/models"
            try:
                resp = requests.get(test_url, headers=headers, timeout=4)
            except Exception:
                test_url = f"{target_url}/health"
                try:
                    resp = requests.get(test_url, headers=headers, timeout=4)
                except Exception:
                    test_url = target_url
                    resp = requests.get(test_url, headers=headers, timeout=4)

            latency_ms = int((time.time() - start_time) * 1000)

            model_count = None
            if resp.status_code == 200:
                try:
                    body = resp.json()
                    if isinstance(body, dict) and "data" in body and isinstance(body["data"], list):
                        model_count = len(body["data"])
                except Exception:
                    pass

            is_ok = resp.status_code in (200, 204, 301, 302, 401)
            return jsonify({
                "success": is_ok,
                "status_code": resp.status_code,
                "latency_ms": latency_ms,
                "model_count": model_count,
                "url": target_url,
                "message": f"HTTP {resp.status_code} ({latency_ms}ms)"
            })
        except requests.exceptions.ConnectionError:
            return jsonify({
                "success": False,
                "error": f"Connection refused at {target_url}. Gateway is not running or port is closed.",
                "latency_ms": int((time.time() - start_time) * 1000)
            })
        except requests.exceptions.Timeout:
            return jsonify({
                "success": False,
                "error": f"Connection timed out (4s) to {target_url}.",
                "latency_ms": 4000
            })
        except Exception as e:
            return jsonify({
                "success": False,
                "error": f"Ping failed: {str(e)}",
                "latency_ms": int((time.time() - start_time) * 1000)
            })

    @app.route("/api/config/diagnostics", methods=["GET"])
    def get_diagnostics():
        import platform
        import sys

        com_status = "unavailable"
        com_detail = "Non-Windows OS"
        if sys.platform == "win32":
            try:
                import win32com.client
                com_status = "available"
                com_detail = "ActiveX / COM Automation library loaded"
            except Exception as e:
                com_status = "error"
                com_detail = str(e)

        output_count = 0
        output_size_bytes = 0
        if OUTPUT_DIR.exists():
            for p in OUTPUT_DIR.glob("**/*"):
                if p.is_file():
                    output_count += 1
                    try:
                        output_size_bytes += p.stat().st_size
                    except OSError:
                        pass
        output_size_mb = round(output_size_bytes / (1024 * 1024), 2)

        components_file = COMPONENTS_DIR / "components.json"
        component_count = 0
        if components_file.exists():
            try:
                with open(components_file, "r", encoding="utf-8") as f:
                    cdata = json.load(f)
                    component_count = len(cdata) if isinstance(cdata, list) else len(cdata.get("components", []))
            except Exception:
                pass

        templates_count = len(list(DATA_DIR.glob("T*.pptx")))

        return jsonify({
            "success": True,
            "platform": {
                "system": platform.system(),
                "release": platform.release(),
                "python": platform.python_version(),
                "os": sys.platform
            },
            "com_engine": {
                "status": com_status,
                "detail": com_detail
            },
            "storage": {
                "output_dir": str(OUTPUT_DIR),
                "output_files_count": output_count,
                "output_size_mb": output_size_mb,
                "components_count": component_count,
                "templates_count": templates_count
            },
            "runtime": {
                "render_mode": Config.RENDER_MODE,
                "pure_pil_active": Config.PURE_PIL_ACTIVE,
                "chat_model": Config.NINEROUTER_CHAT_MODEL
            }
        })

    @app.route("/api/config/clean-cache", methods=["POST"])
    def clean_render_cache():
        cleaned_count = 0
        freed_bytes = 0
        if OUTPUT_DIR.exists():
            for p in list(OUTPUT_DIR.glob("**/*.png")) + list(OUTPUT_DIR.glob("**/*.jpg")):
                try:
                    sz = p.stat().st_size
                    p.unlink()
                    cleaned_count += 1
                    freed_bytes += sz
                except Exception:
                    pass

        freed_mb = round(freed_bytes / (1024 * 1024), 2)
        return jsonify({
            "success": True,
            "cleaned_count": cleaned_count,
            "freed_mb": freed_mb,
            "message": f"Cleaned {cleaned_count} preview cache files, freeing {freed_mb} MB."
        })

    @app.route("/api/config/raw", methods=["GET", "POST"])
    def handle_raw_config():
        env_path = Path(__file__).resolve().parent.parent.parent.parent / ".env"
        if request.method == "GET":
            if env_path.exists():
                with open(env_path, "r", encoding="utf-8") as f:
                    content = f.read()
            else:
                example_path = Path(__file__).resolve().parent.parent.parent.parent / ".env.example"
                if example_path.exists():
                    with open(example_path, "r", encoding="utf-8") as f:
                        content = f.read()
                else:
                    content = ""
            return jsonify({"success": True, "raw": content})
        else:
            data = request.get_json() or {}
            raw_text = data.get("raw", "")
            if env_path.exists():
                try:
                    backup_path = env_path.with_suffix(".env.backup")
                    shutil.copy2(env_path, backup_path)
                except Exception:
                    pass
            with open(env_path, "w", encoding="utf-8") as f:
                f.write(raw_text)
            Config.reload()
            return jsonify({
                "success": True,
                "message": "Raw .env saved and reloaded successfully.",
                "model": Config.NINEROUTER_CHAT_MODEL
            })

    return app
