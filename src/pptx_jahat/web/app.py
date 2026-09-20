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

from pptx_jahat.config import Config, DATA_DIR, OUTPUT_DIR, COMPONENTS_DIR, STRUCTURES_DIR, IMAGES_DIR
from pptx_jahat.agent import AIAgent
from pptx_jahat.tools.pptx_builder import (
    build_pptx_with_agent,
    verify_and_auto_heal_pptx,
    get_initial_diagnostics_steps
)
from pptx_jahat.tools.preview import render_pptx_file_previews, image_to_base64_jpeg, image_to_base64_png
from pptx_jahat.tools.template_analyzer import (
    analyze_template,
    analyze_all_templates,
    load_notes,
    save_notes,
    get_analyzed_templates,
    get_standard_note_schema,
    generate_blank_structured_note,
    _extract_template_summary_for_ai,
    _generate_fallback_template_note,
    update_template_note_in_file,
    NOTE_FILE
)
from pptx_jahat.tools.pptx_engine import extract_all_templates, get_components_catalog, extract_template_fonts
from pptx_jahat.tools.font_verifier import get_available_system_fonts, verify_fonts_inventory
from pptx_jahat.tools.structure_manager import (
    list_structure_files,
    get_structure_content,
    save_structure_content,
    delete_structure_file,
    build_structure_from_template,
    build_structure_from_sample_files
)
from pptx_jahat.tools.multi_parser import (
    is_supported_file,
    get_file_type_category,
    parse_multiple_sources,
    ALL_SUPPORTED_EXTS
)
from pptx_jahat.tools.human_touch import (
    refine_extracted_content_with_ai,
    refine_restructured_slides_with_ai,
    edit_pptx_with_ai,
    inspect_pptx_for_editing
)

# Global Job Registry for SSE streams
JOBS: Dict[str, Dict[str, Any]] = {}
JOBS_LOCK = threading.Lock()

BOOT_ID = f"{int(time.time() * 1000)}_{os.getpid()}"

def create_app(dev_mode: Optional[bool] = None) -> Flask:
    template_dir = Path(__file__).parent / "templates"
    static_dir = Path(__file__).parent / "static"

    app = Flask(
        __name__,
        template_folder=str(template_dir),
        static_folder=str(static_dir)
    )
    CORS(app)
    app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200MB upload limit

    if dev_mode is None:
        dev_mode = (
            os.getenv("FLASK_DEBUG", "").lower() in ("1", "true", "yes")
            or os.getenv("DEV", "").lower() in ("1", "true", "yes")
            or os.getenv("DEBUG", "").lower() in ("1", "true", "yes")
            or os.getenv("RELOAD", "").lower() in ("1", "true", "yes")
            or os.getenv("FLASK_ENV", "").lower() == "development"
        )

    app.config["DEV_MODE"] = bool(dev_mode)
    app.config["BOOT_ID"] = BOOT_ID

    if dev_mode:
        app.config["DEBUG"] = True
        app.config["TEMPLATES_AUTO_RELOAD"] = True
        app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0
        app.jinja_env.auto_reload = True

    @app.context_processor
    def inject_dev_context():
        return {
            "dev_mode": app.config.get("DEV_MODE", False),
            "boot_id": app.config.get("BOOT_ID", BOOT_ID)
        }

    # Upload cache dir
    UPLOAD_CACHE = DATA_DIR / "uploads"
    UPLOAD_CACHE.mkdir(parents=True, exist_ok=True)

    def get_latest_frontend_mtime() -> tuple[float, str]:
        latest_mtime = 0.0
        latest_file = ""
        watch_dirs = [template_dir, static_dir]
        for d in watch_dirs:
            if not d.exists():
                continue
            for p in d.rglob("*"):
                if p.is_file() and p.suffix in (".html", ".css", ".js", ".json", ".svg", ".png", ".jpg"):
                    try:
                        m = p.stat().st_mtime
                        if m > latest_mtime:
                            latest_mtime = m
                            latest_file = p.name
                    except OSError:
                        pass
        return latest_mtime, latest_file

    @app.route("/api/dev/status", methods=["GET"])
    def dev_status():
        is_dev = bool(app.config.get("DEV_MODE", False))
        return jsonify({
            "success": True,
            "dev_mode": is_dev,
            "boot_id": app.config.get("BOOT_ID", BOOT_ID),
            "auto_reload": is_dev,
            "watched_directories": [
                str(template_dir.resolve()),
                str(static_dir.resolve())
            ]
        })

    @app.route("/api/dev/live-reload")
    def dev_live_reload():
        if not app.config.get("DEV_MODE", False):
            return jsonify({
                "success": False,
                "enabled": False,
                "error": "Live reload is only enabled when running in dev mode."
            }), 404

        def event_stream():
            try:
                yield f"event: init\ndata: {json.dumps({'boot_id': app.config.get('BOOT_ID', BOOT_ID), 'status': 'connected'})}\n\n"
                last_mtime, _ = get_latest_frontend_mtime()
                while True:
                    time.sleep(0.5)
                    curr_mtime, changed_file = get_latest_frontend_mtime()
                    if curr_mtime > last_mtime:
                        last_mtime = curr_mtime
                        yield f"event: reload\ndata: {json.dumps({'reason': 'file_changed', 'file': changed_file, 'time': time.time()})}\n\n"
                    else:
                        yield f": ping\n\n"
            except GeneratorExit:
                pass
            except Exception:
                pass

        return Response(
            event_stream(),
            mimetype="text/event-stream",
            headers={
                "Cache-Control": "no-cache",
                "Connection": "keep-alive",
                "X-Accel-Buffering": "no"
            }
        )

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

    @app.route("/api/fonts/system", methods=["GET"])
    def get_system_fonts_api():
        force_rescan = request.args.get("rescan", "").lower() in ("true", "1", "yes")
        catalog = get_available_system_fonts(force_rescan=force_rescan)
        return jsonify({
            "success": True,
            "families": catalog.get("families", []),
            "persian_fonts": catalog.get("persian_fonts", []),
            "latin_fonts": catalog.get("latin_fonts", []),
            "total_count": len(catalog.get("families", []))
        })

    @app.route("/api/fonts/system/rescan", methods=["POST"])
    def rescan_system_fonts_api():
        catalog = get_available_system_fonts(force_rescan=True)
        return jsonify({
            "success": True,
            "families": catalog.get("families", []),
            "persian_fonts": catalog.get("persian_fonts", []),
            "latin_fonts": catalog.get("latin_fonts", []),
            "total_count": len(catalog.get("families", []))
        })

    @app.route("/api/fonts/verify", methods=["GET", "POST"])
    def verify_template_fonts_api():
        if request.method == "POST":
            data = request.get_json() or {}
        else:
            data = request.args.to_dict()

        if data.get("rescan", "").lower() in ("true", "1", "yes"):
            get_available_system_fonts(force_rescan=True)

        template_name = data.get("template_name", "").strip()
        provided_fonts = data.get("fonts", [])
        if isinstance(provided_fonts, str):
            provided_fonts = [s.strip() for s in provided_fonts.split(",") if s.strip()]

        fonts_to_check = set()
        if isinstance(provided_fonts, list):
            for f in provided_fonts:
                if f and str(f).strip():
                    fonts_to_check.add(str(f).strip())

        from pptx import Presentation
        if template_name and template_name not in ("All Templates", "✨ All Templates (Global AI Intelligent Matching)", "✨ All Templates (Global AI Matching)"):
            candidate = DATA_DIR / template_name
            if candidate.exists() and candidate.suffix.lower() == ".pptx":
                try:
                    prs = Presentation(str(candidate))
                    tpl_f = extract_template_fonts(prs)
                    fonts_to_check.update(tpl_f)
                except Exception as ex:
                    return jsonify({"success": False, "error": f"Failed to inspect template '{template_name}': {ex}"}), 400
        elif not fonts_to_check:
            # Check all templates in DATA_DIR
            for p in DATA_DIR.glob("*.pptx"):
                if not p.name.endswith("_generated.pptx"):
                    try:
                        prs = Presentation(str(p))
                        fonts_to_check.update(extract_template_fonts(prs))
                    except Exception:
                        pass

        fonts_list = sorted(list(fonts_to_check))
        result = verify_fonts_inventory(fonts_list)
        result["success"] = True
        result["template_name"] = template_name or "All Templates"
        return jsonify(result)

    @app.route("/api/generator/upload", methods=["POST"])
    def upload_docx():
        if "file" not in request.files:
            return jsonify({"success": False, "error": "No file part in request"}), 400
        file = request.files["file"]
        if not file.filename:
            return jsonify({"success": False, "error": "Invalid file."}), 400

        filename = Path(file.filename).name
        target_path = UPLOAD_CACHE / filename
        file.save(str(target_path))

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        suggested_output = str(OUTPUT_DIR / f"{Path(filename).stem}_generated.pptx")

        return jsonify({
            "success": True,
            "filename": filename,
            "file_path": str(target_path.resolve()),
            "category": get_file_type_category(filename),
            "suggested_output": suggested_output
        })

    @app.route("/api/generator/upload-multi", methods=["POST"])
    def upload_multi_sources():
        uploaded_files = request.files.getlist("files")
        if not uploaded_files and "file" in request.files:
            uploaded_files = [request.files["file"]]

        if not uploaded_files:
            return jsonify({"success": False, "error": "No files uploaded."}), 400

        results = []
        for file in uploaded_files:
            if not file.filename:
                continue
            fname = Path(file.filename).name
            if not is_supported_file(fname):
                continue
            target = UPLOAD_CACHE / fname
            file.save(str(target))
            results.append({
                "filename": fname,
                "file_path": str(target.resolve()),
                "category": get_file_type_category(fname),
                "size_kb": round(target.stat().st_size / 1024, 1)
            })

        if not results:
            return jsonify({
                "success": False,
                "error": "No supported files provided. Supported: Word, PPTX, Text, Images, Audio."
            }), 400

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        suggested_output = str(OUTPUT_DIR / f"{Path(results[0]['filename']).stem}_generated.pptx")

        return jsonify({
            "success": True,
            "files": results,
            "primary_file": results[0]["file_path"],
            "suggested_output": suggested_output
        })

    @app.route("/api/generator/generate", methods=["POST"])
    def start_generation():
        data = request.get_json() or {}
        docx_path = data.get("docx_path", "").strip()
        source_files = data.get("source_files", [])
        raw_text = data.get("raw_text", "").strip()
        template_name = data.get("template_name", None)
        structure_name = data.get("structure_name", None)
        detection_structure_name = data.get("detection_structure_name", None)
        restructure_structure_name = data.get("restructure_structure_name", None)
        blueprint_structure_name = data.get("blueprint_structure_name", None)
        enable_detection = bool(data.get("enable_detection", True))
        enable_restructure = bool(data.get("enable_restructure", False))
        enable_blueprint = bool(data.get("enable_blueprint", True))
        enable_verification = bool(data.get("enable_verification", True))
        verification_rounds_val = data.get("verification_rounds")
        try:
            verification_rounds = int(verification_rounds_val) if verification_rounds_val is not None else Config.VERIFICATION_ROUNDS
        except (ValueError, TypeError):
            verification_rounds = Config.VERIFICATION_ROUNDS
        enable_human_touch = bool(data.get("enable_human_touch", False))
        human_touch_steps = data.get("human_touch_steps", ["extract", "restructure", "verify", "after_done"])
        output_path = data.get("output_path", "").strip()
        timeout_val = data.get("timeout", None)
        font_fallbacks = data.get("font_fallbacks", {})
        if not isinstance(font_fallbacks, dict):
            font_fallbacks = {}

        # Allow user to specify or override active 9Router chat model per generation
        chat_model_override = str(data.get("chat_model") or data.get("model") or "").strip()
        if chat_model_override:
            Config.NINEROUTER_CHAT_MODEL = chat_model_override
            os.environ["NINEROUTER_CHAT_MODEL"] = chat_model_override

        try:
            req_timeout = float(timeout_val) if timeout_val is not None else None
        except (ValueError, TypeError):
            req_timeout = None

        all_sources = []
        if source_files and isinstance(source_files, list):
            all_sources.extend([str(p).strip() for p in source_files if str(p).strip() and Path(str(p).strip()).exists()])
        if docx_path and docx_path not in all_sources and Path(docx_path).exists():
            all_sources.append(docx_path)

        if not all_sources and not raw_text:
            return jsonify({"success": False, "error": "Please provide at least one source file or text notes."}), 400

        if template_name and ("All Templates" in template_name or "No templates" in template_name):
            template_name = None

        def _clean_struct_name(val):
            if not val:
                return None
            s = str(val).strip()
            if not s or "none" in s.lower() or "select" in s.lower():
                return None
            return s

        structure_name = _clean_struct_name(structure_name)
        detection_structure_name = _clean_struct_name(detection_structure_name)
        restructure_structure_name = _clean_struct_name(restructure_structure_name)
        blueprint_structure_name = _clean_struct_name(blueprint_structure_name)

        if not output_path:
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            stem = Path(all_sources[0]).stem if all_sources else "presentation"
            output_path = str(OUTPUT_DIR / f"{stem}_generated.pptx")

        job_id = f"gen_{int(time.time() * 1000)}"
        event_queue = queue.Queue()
        human_event = threading.Event()

        with JOBS_LOCK:
            JOBS[job_id] = {
                "id": job_id,
                "type": "generator",
                "queue": event_queue,
                "status": "running",
                "result": None,
                "error": None,
                "ai_images": [],
                "diagnostics": get_initial_diagnostics_steps(),
                "enable_human_touch": enable_human_touch,
                "human_touch_steps": human_touch_steps,
                "waiting_for_human": False,
                "human_step": None,
                "human_event": human_event,
                "human_action": None,
                "step_data": None
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

            def on_step_update(step_info: Dict[str, Any]):
                with JOBS_LOCK:
                    if job_id in JOBS:
                        diag = JOBS[job_id].setdefault("diagnostics", get_initial_diagnostics_steps())
                        for i, s in enumerate(diag):
                            if s.get("id") == step_info.get("id"):
                                diag[i] = dict(step_info)
                                break
                        else:
                            diag.append(dict(step_info))
                event_queue.put({"event": "step_update", "data": step_info})

            def on_human_review(step: str, step_data: Dict[str, Any]) -> Optional[Dict[str, Any]]:
                with JOBS_LOCK:
                    if job_id not in JOBS:
                        return step_data
                    JOBS[job_id]["waiting_for_human"] = True
                    JOBS[job_id]["human_step"] = step
                    JOBS[job_id]["step_data"] = step_data
                    JOBS[job_id]["human_action"] = None
                    human_event.clear()

                step_titles = {
                    "font_fallback": "Step 1.2: Verify Template Fonts & Fallbacks",
                    "extract": "Step 2: Extract Slides from Inputs",
                    "restructure": "Step 2.5: Restructure Slides",
                    "verify": "Step 4.5: Visual Template Alignment & Verification"
                }
                title_str = step_titles.get(step, f"Step: {step}")
                log_callback(f"[Human Touch] {title_str} ready. Pausing for human review, edit, or rerun prompt...")
                event_queue.put({
                    "event": "human_review",
                    "data": {
                        "job_id": job_id,
                        "step": step,
                        "step_name": title_str,
                        "data": step_data
                    }
                })

                current_data = step_data
                while True:
                    signaled = human_event.wait(timeout=0.5)
                    with JOBS_LOCK:
                        job_state = JOBS.get(job_id, {})
                        if job_state.get("status") in ("error", "cancelled"):
                            raise RuntimeError("Job cancelled during human review.")

                    if signaled:
                        with JOBS_LOCK:
                            action_info = JOBS[job_id].get("human_action") or {}
                            human_event.clear()

                        action = action_info.get("action", "continue")
                        if action == "continue":
                            new_data = action_info.get("data") or current_data
                            with JOBS_LOCK:
                                if job_id in JOBS:
                                    JOBS[job_id]["waiting_for_human"] = False
                                    JOBS[job_id]["human_step"] = None
                                    JOBS[job_id]["step_data"] = new_data
                            log_callback(f"[Human Touch] User approved {title_str}. Resuming pipeline...")
                            event_queue.put({
                                "event": "human_action_accepted",
                                "data": {"job_id": job_id, "step": step, "action": "continue"}
                            })
                            return new_data

                        elif action == "rerun":
                            user_prompt = action_info.get("prompt", "").strip()
                            edited_data = action_info.get("data") or current_data
                            log_callback(f"[Human Touch] User requested rerun for {title_str} with prompt: '{user_prompt}'")
                            event_queue.put({
                                "event": "status",
                                "data": {"status": f"AI Rerunning {title_str}..."}
                            })
                            try:
                                if step == "extract":
                                    refined = refine_extracted_content_with_ai(
                                        edited_data,
                                        user_prompt,
                                        log_cb=log_callback,
                                        timeout=req_timeout
                                    )
                                elif step == "restructure":
                                    struct_bp = None
                                    if restructure_structure_name:
                                        try:
                                            struct_bp = get_structure_content(restructure_structure_name)
                                        except Exception:
                                            pass
                                    refined = refine_restructured_slides_with_ai(
                                        edited_data,
                                        user_prompt,
                                        structure_blueprint=struct_bp,
                                        log_cb=log_callback,
                                        timeout=req_timeout
                                    )
                                elif step == "verify":
                                    from pptx_jahat.tools.slide_verifier import converse_with_verification_agent
                                    chat_history = edited_data.get("conversation_history", [])
                                    chat_history.append({"role": "user", "content": user_prompt})
                                    active_sidx = None
                                    if isinstance(action_info.get("data"), dict):
                                        active_sidx = action_info.get("data", {}).get("active_slide_index")
                                    refined = converse_with_verification_agent(
                                        current_verification_data=edited_data,
                                        user_message=user_prompt,
                                        active_slide_index=active_sidx,
                                        conversation_history=chat_history,
                                        log_cb=log_callback,
                                        timeout=req_timeout
                                    )
                                    if refined.get("agent_reply"):
                                        chat_history.append({"role": "assistant", "content": refined.get("agent_reply", "")})
                                    refined["conversation_history"] = chat_history
                                else:
                                    refined = edited_data

                                current_data = refined
                                with JOBS_LOCK:
                                    if job_id in JOBS:
                                        JOBS[job_id]["step_data"] = current_data
                                        JOBS[job_id]["waiting_for_human"] = True

                                event_queue.put({
                                    "event": "human_review",
                                    "data": {
                                        "job_id": job_id,
                                        "step": step,
                                        "step_name": title_str,
                                        "data": current_data,
                                        "refined": True
                                    }
                                })
                            except Exception as ex:
                                log_callback(f"[!] Rerun notice: {ex}")
                                event_queue.put({
                                    "event": "human_review_error",
                                    "data": {"job_id": job_id, "step": step, "error": str(ex)}
                                })

                        elif action == "cancel":
                            raise RuntimeError("Generation cancelled by user during human review.")

            try:
                event_queue.put({"event": "status", "data": {"status": "Generating presentation..."}})
                res = build_pptx_with_agent(
                    docx_path=all_sources if len(all_sources) > 1 else (all_sources[0] if all_sources else None),
                    output_path=output_path,
                    template_name=template_name,
                    log_callback=log_callback,
                    on_ai_images_ready=on_ai_images_ready,
                    structure_name=structure_name,
                    raw_text=raw_text,
                    enable_restructure=enable_restructure,
                    timeout=req_timeout,
                    detection_structure_name=detection_structure_name,
                    restructure_structure_name=restructure_structure_name,
                    blueprint_structure_name=blueprint_structure_name,
                    enable_detection=enable_detection,
                    enable_blueprint=enable_blueprint,
                    on_step_update=on_step_update,
                    enable_human_touch=enable_human_touch,
                    human_touch_steps=human_touch_steps,
                    on_human_review=on_human_review,
                    font_fallbacks=font_fallbacks,
                    enable_verification=enable_verification,
                    verification_rounds=verification_rounds
                )

                # Pre-render slides for instant UI loading
                previews = []
                try:
                    imgs, engine_name = render_pptx_file_previews(res, target_width_px=800, return_engine_info=True)
                    img_list = imgs if isinstance(imgs, list) else [imgs]
                    for idx, img in enumerate(img_list):
                        previews.append({
                            "slide_index": idx,
                            "data_url": image_to_base64_jpeg(img, quality=85)
                        })
                except Exception as ex:
                    engine_name = "None"
                    log_callback(f"Preview render warning: {str(ex)}")

                with JOBS_LOCK:
                    current_diag = JOBS[job_id].get("diagnostics", []) if job_id in JOBS else []
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "completed"
                        JOBS[job_id]["result"] = res

                event_queue.put({
                    "event": "completed",
                    "data": {
                        "pptx_path": res,
                        "filename": Path(res).name,
                        "engine_name": engine_name,
                        "previews": previews,
                        "diagnostics": current_diag
                    }
                })
            except Exception as e:
                with JOBS_LOCK:
                    current_diag = JOBS[job_id].get("diagnostics", []) if job_id in JOBS else []
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "error"
                        JOBS[job_id]["error"] = str(e)
                event_queue.put({"event": "error", "data": {"error": str(e), "diagnostics": current_diag}})
            finally:
                event_queue.put({"event": "close", "data": {}})

        threading.Thread(target=worker, daemon=True).start()

        return jsonify({
            "success": True,
            "job_id": job_id,
            "output_path": output_path
        })

    @app.route("/api/generator/diagnostics", methods=["GET"])
    def get_generator_diagnostics():
        job_id = request.args.get("job_id")
        with JOBS_LOCK:
            if job_id and job_id in JOBS:
                return jsonify({
                    "success": True,
                    "job_id": job_id,
                    "diagnostics": JOBS[job_id].get("diagnostics", [])
                })
            # Return the latest generator job with diagnostics
            for j_id, j in reversed(list(JOBS.items())):
                if j.get("type") == "generator" and "diagnostics" in j:
                    return jsonify({
                        "success": True,
                        "job_id": j_id,
                        "diagnostics": j.get("diagnostics", [])
                    })
        return jsonify({
            "success": True,
            "job_id": None,
            "diagnostics": get_initial_diagnostics_steps()
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

    # -------------------------------------------------------------
    # 1.1. HUMAN TOUCH & INTERACTIVE WORKFLOW ENDPOINTS
    # -------------------------------------------------------------
    @app.route("/api/generator/human-action", methods=["POST"])
    def handle_human_action():
        data = request.get_json() or {}
        job_id = str(data.get("job_id") or "")
        action = data.get("action", "continue")
        prompt = data.get("prompt", "").strip()
        step = data.get("step")
        action_data = data.get("data")

        if not job_id:
            return jsonify({"success": False, "error": "job_id is required."}), 400

        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if not job:
                return jsonify({"success": False, "error": f"Job '{job_id}' not found."}), 404
            if not job.get("waiting_for_human"):
                return jsonify({"success": False, "error": f"Job '{job_id}' is not currently waiting for human input."}), 400

            job["human_action"] = {
                "action": action,
                "prompt": prompt,
                "step": step,
                "data": action_data
            }
            event = job.get("human_event")
            if event:
                event.set()

        return jsonify({
            "success": True,
            "job_id": job_id,
            "action": action,
            "message": f"Human touch action '{action}' signaled successfully."
        })

    @app.route("/api/generator/human-status/<job_id>", methods=["GET"])
    def get_human_status(job_id: str):
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if not job:
                return jsonify({"success": False, "error": "Job not found."}), 404
            return jsonify({
                "success": True,
                "job_id": job_id,
                "waiting_for_human": job.get("waiting_for_human", False),
                "human_step": job.get("human_step"),
                "step_data": job.get("step_data")
            })

    @app.route("/api/generator/extract", methods=["POST"])
    def api_extract_content():
        data = request.get_json() or {}
        source_files = data.get("source_files", [])
        docx_path = data.get("docx_path", "").strip()
        raw_text = data.get("raw_text", "").strip()
        structure_name = data.get("structure_name", None)
        timeout_val = data.get("timeout")

        try:
            req_timeout = float(timeout_val) if timeout_val is not None else None
        except (ValueError, TypeError):
            req_timeout = None

        all_sources = []
        if source_files and isinstance(source_files, list):
            all_sources.extend([str(p).strip() for p in source_files if str(p).strip() and Path(str(p).strip()).exists()])
        if docx_path and docx_path not in all_sources and Path(docx_path).exists():
            all_sources.append(docx_path)

        if not all_sources and not raw_text:
            return jsonify({"success": False, "error": "No source files or text provided."}), 400

        detection_bp = None
        if structure_name:
            try:
                detection_bp = get_structure_content(structure_name)
            except Exception:
                pass

        try:
            parsed = parse_multiple_sources(
                all_sources,
                raw_text=raw_text,
                timeout=req_timeout,
                structure_blueprint=detection_bp
            )
            return jsonify({"success": True, "data": parsed})
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/generator/extract/rerun", methods=["POST"])
    def api_rerun_extract_content():
        data = request.get_json() or {}
        current_data = data.get("current_data", {})
        prompt = data.get("prompt", "").strip()
        timeout_val = data.get("timeout")

        try:
            req_timeout = float(timeout_val) if timeout_val is not None else None
        except (ValueError, TypeError):
            req_timeout = None

        if not prompt:
            return jsonify({"success": False, "error": "Prompt is required to rerun extraction."}), 400

        try:
            refined = refine_extracted_content_with_ai(
                current_data,
                prompt,
                timeout=req_timeout
            )
            return jsonify({"success": True, "data": refined})
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/generator/restructure/rerun", methods=["POST"])
    def api_rerun_restructure_slides():
        data = request.get_json() or {}
        current_data = data.get("current_data", {})
        prompt = data.get("prompt", "").strip()
        structure_name = data.get("structure_name", None)
        timeout_val = data.get("timeout")

        try:
            req_timeout = float(timeout_val) if timeout_val is not None else None
        except (ValueError, TypeError):
            req_timeout = None

        if not prompt:
            return jsonify({"success": False, "error": "Prompt is required to rerun restructure."}), 400

        struct_bp = None
        if structure_name:
            try:
                struct_bp = get_structure_content(structure_name)
            except Exception:
                pass

        try:
            refined = refine_restructured_slides_with_ai(
                current_data,
                prompt,
                structure_blueprint=struct_bp,
                timeout=req_timeout
            )
            return jsonify({"success": True, "data": refined})
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/generator/verify/chat", methods=["POST"])
    def api_verify_chat():
        data = request.get_json() or {}
        verification_data = data.get("verification_data", {})
        user_message = data.get("message", "").strip()
        active_slide = data.get("active_slide_index", None)
        history = data.get("history", [])
        timeout_val = data.get("timeout", None)

        if not user_message:
            return jsonify({"success": False, "error": "Message cannot be empty."}), 400

        try:
            from pptx_jahat.tools.slide_verifier import converse_with_verification_agent
            result = converse_with_verification_agent(
                current_verification_data=verification_data,
                user_message=user_message,
                active_slide_index=active_slide,
                conversation_history=history,
                timeout=float(timeout_val) if timeout_val else None
            )
            return jsonify(result)
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/generator/verify/apply", methods=["POST"])
    def api_verify_apply():
        data = request.get_json() or {}
        file_path = data.get("file_path", "").strip()
        actions = data.get("actions", [])

        if not file_path or not Path(file_path).exists():
            return jsonify({"success": False, "error": "PPTX file path does not exist."}), 400

        try:
            from pptx_jahat.tools.slide_verifier import apply_verification_edits
            res = apply_verification_edits(file_path, actions)
            return jsonify(res)
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/generator/edit-pptx", methods=["POST"])
    def api_edit_presentation():
        data = request.get_json() or {}
        file_path = data.get("file_path", "").strip()
        prompt = data.get("prompt", "").strip()
        template_name = data.get("template_name")
        timeout_val = data.get("timeout")

        try:
            req_timeout = float(timeout_val) if timeout_val is not None else None
        except (ValueError, TypeError):
            req_timeout = None

        if not file_path or not Path(file_path).exists():
            return jsonify({"success": False, "error": "PPTX file path does not exist."}), 400
        if not prompt:
            return jsonify({"success": False, "error": "Custom prompt cannot be empty."}), 400

        try:
            res = edit_pptx_with_ai(
                pptx_path=file_path,
                user_prompt=prompt,
                template_name=template_name,
                timeout=req_timeout
            )
            return jsonify(res)
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/generator/inspect-deck", methods=["POST"])
    def api_inspect_deck():
        data = request.get_json() or {}
        file_path = data.get("file_path", "").strip()
        if not file_path or not Path(file_path).exists():
            return jsonify({"success": False, "error": "File does not exist."}), 400
        try:
            info = inspect_pptx_for_editing(file_path)
            return jsonify({"success": True, "deck": info})
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

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
                sw = prs.slide_width
                sh = prs.slide_height
                width_in = round(float(sw) / 914400, 2) if sw is not None else 10.0
                height_in = round(float(sh) / 914400, 2) if sh is not None else 7.5
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
                "brief": info.get("brief", ""),
                "display_name": info.get("display_name", tpl.stem),
                "domain": info.get("domain", "General Business & Educational"),
                "color_theme": info.get("color_theme", "Standard"),
                "typography": info.get("typography", "Standard"),
                "density": info.get("density", "Medium-density"),
                "slide_catalog": info.get("slide_catalog", []),
                "flow_recipe": info.get("flow_recipe", []),
                "trigger_conditions": info.get("trigger_conditions", ""),
                "is_structured": info.get("is_structured", False)
            })

        return jsonify({
            "success": True,
            "templates": items,
            "total_count": len(items),
            "analyzed_count": sum(1 for i in items if i["is_analyzed"])
        })

    @app.route("/api/templates/schema", methods=["GET"])
    def get_template_schema_spec():
        """Returns the official standardized template note schema specification."""
        return jsonify({
            "success": True,
            "schema": get_standard_note_schema()
        })

    @app.route("/api/templates/boilerplate", methods=["GET"])
    def get_template_boilerplate():
        """Returns structured note boilerplate pre-populated for a given template."""
        filename = request.args.get("filename", "").strip()
        tpl_path = (DATA_DIR / filename) if filename else None
        boilerplate = generate_blank_structured_note(filename, pptx_path=tpl_path)
        return jsonify({
            "success": True,
            "filename": filename,
            "boilerplate": boilerplate
        })

    @app.route("/api/templates/apply-schema", methods=["POST"])
    def apply_standard_schema_to_template():
        """
        Applies standardized structured template schema to a specific template or all templates.
        """
        data = request.get_json() or {}
        filename = data.get("filename", "").strip()
        apply_all = data.get("all", False)

        if apply_all:
            pptx_files = sorted(list(DATA_DIR.glob("*.pptx")))
            templates = [f for f in pptx_files if not f.name.endswith("_generated.pptx")]
            for tpl in templates:
                try:
                    summary = _extract_template_summary_for_ai(tpl)
                    note = _generate_fallback_template_note(summary)
                    update_template_note_in_file(tpl.name, note, NOTE_FILE)
                except Exception as e:
                    pass
            return jsonify({
                "success": True,
                "message": f"Successfully applied standardized schema to all {len(templates)} templates.",
                "content": load_notes(NOTE_FILE)
            })

        if not filename:
            return jsonify({"success": False, "error": "Filename is required"}), 400

        tpl_path = DATA_DIR / filename
        if not tpl_path.exists():
            return jsonify({"success": False, "error": f"Template '{filename}' not found."}), 404

        summary = _extract_template_summary_for_ai(tpl_path)
        note = _generate_fallback_template_note(summary)
        updated_content = update_template_note_in_file(filename, note, NOTE_FILE)

        return jsonify({
            "success": True,
            "message": f"Successfully formatted note for {filename} into standardized structured schema.",
            "note": note,
            "content": updated_content
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

    # -------------------------------------------------------------
    # 2.1. SLIDE STORYBOARD STRUCTURE ENDPOINTS (data/structure/*.md)
    # -------------------------------------------------------------
    @app.route("/api/structure/list", methods=["GET"])
    def api_list_structures():
        structures = list_structure_files()
        return jsonify({
            "success": True,
            "structures": structures,
            "count": len(structures)
        })

    @app.route("/api/structure/get", methods=["GET"])
    def api_get_structure():
        name = request.args.get("name", "").strip()
        if not name:
            return jsonify({"success": False, "error": "Structure name required."}), 400
        try:
            content = get_structure_content(name)
            return jsonify({
                "success": True,
                "name": name,
                "content": content
            })
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 404

    @app.route("/api/structure/save", methods=["POST"])
    def api_save_structure():
        data = request.get_json() or {}
        name = data.get("name", "").strip()
        content = data.get("content", "")
        if not name:
            return jsonify({"success": False, "error": "Structure name cannot be empty."}), 400
        try:
            res = save_structure_content(name, content)
            return jsonify(res)
        except Exception as e:
            return jsonify({"success": False, "error": str(e)}), 500

    @app.route("/api/structure/upload", methods=["POST"])
    def api_upload_structure():
        files = request.files.getlist("files")
        if not files and "file" in request.files:
            files = [request.files["file"]]

        if not files:
            return jsonify({"success": False, "error": "No structure files uploaded."}), 400

        STRUCTURES_DIR.mkdir(parents=True, exist_ok=True)
        results = []
        for file in files:
            if not file or not file.filename:
                continue
            fname = Path(file.filename).name
            if not fname.lower().endswith(".md"):
                fname += ".md"
            target = STRUCTURES_DIR / fname
            file.save(str(target))
            stat = target.stat()
            results.append({
                "filename": target.name,
                "name": target.stem,
                "path": str(target.resolve()),
                "size": f"{round(stat.st_size / 1024, 1)} KB",
                "modified": time.strftime("%Y-%m-%d %H:%M", time.localtime(stat.st_mtime))
            })

        if not results:
            return jsonify({"success": False, "error": "No valid Markdown (.md) structure files uploaded."}), 400

        return jsonify({
            "success": True,
            "structures": results,
            "uploaded": results[0]
        })

    @app.route("/api/structure/delete", methods=["DELETE"])
    def api_delete_structure():
        data = request.get_json() or {}
        name = data.get("name", "").strip()
        if not name:
            return jsonify({"success": False, "error": "Structure name required."}), 400
        deleted = delete_structure_file(name)
        if deleted:
            return jsonify({"success": True, "message": f"Deleted structure: {name}"})
        else:
            return jsonify({"success": False, "error": "Cannot delete protected sample or file not found."}), 400

    @app.route("/api/structure/build", methods=["POST"])
    def api_build_structure():
        data = request.get_json() or {}
        template_name = data.get("template_name", "").strip()
        structure_name = data.get("structure_name", "").strip()
        custom_instructions = data.get("custom_instructions", "").strip()
        timeout_val = data.get("timeout", None)

        try:
            struct_timeout = float(timeout_val) if timeout_val is not None else None
        except (ValueError, TypeError):
            struct_timeout = None

        if not template_name:
            return jsonify({"success": False, "error": "Template name is required."}), 400
        if not structure_name:
            structure_name = f"{Path(template_name).stem}-structure"

        job_id = f"struct_{int(time.time() * 1000)}"
        event_queue = queue.Queue()

        with JOBS_LOCK:
            JOBS[job_id] = {
                "id": job_id,
                "type": "structure_build",
                "queue": event_queue,
                "status": "running"
            }

        def worker():
            def log_cb(msg: str):
                event_queue.put({"event": "log", "data": {"message": msg, "time": time.strftime("%H:%M:%S")}})

            try:
                log_cb(f"[*] Starting Storyboard Schema Builder for template: {template_name}")
                res = build_structure_from_template(
                    template_name,
                    structure_name,
                    custom_instructions=custom_instructions,
                    log_cb=log_cb,
                    timeout=struct_timeout
                )
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "completed"
                        JOBS[job_id]["result"] = res
                event_queue.put({"event": "completed", "data": res})
            except Exception as e:
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "error"
                        JOBS[job_id]["error"] = str(e)
                event_queue.put({"event": "error", "data": {"error": str(e)}})
            finally:
                event_queue.put({"event": "close", "data": {}})

        threading.Thread(target=worker, daemon=True).start()

        return jsonify({"success": True, "job_id": job_id, "structure_name": structure_name})

    @app.route("/api/structure/upload-samples", methods=["POST"])
    def api_upload_structure_samples():
        files = request.files.getlist("files")
        if not files and "file" in request.files:
            files = [request.files["file"]]

        if not files:
            return jsonify({"success": False, "error": "No files uploaded."}), 400

        samples_dir = UPLOAD_CACHE / "samples"
        samples_dir.mkdir(parents=True, exist_ok=True)

        results = []
        for file in files:
            if not file.filename:
                continue
            fname = Path(file.filename).name
            if not is_supported_file(fname):
                continue
            target = samples_dir / fname
            file.save(str(target))
            results.append({
                "filename": fname,
                "file_path": str(target.resolve()),
                "category": get_file_type_category(fname),
                "size_kb": round(target.stat().st_size / 1024, 1)
            })

        if not results:
            return jsonify({
                "success": False,
                "error": "No supported files found. Supported: Images, Word docs, PPTX, Text."
            }), 400

        return jsonify({"success": True, "files": results})

    @app.route("/api/structure/build-from-samples", methods=["POST"])
    def api_build_structure_from_samples():
        data = request.get_json() or {}
        sample_files = data.get("sample_files", [])
        structure_name = data.get("structure_name", "").strip()
        custom_instructions = data.get("custom_instructions", "").strip()
        timeout_val = data.get("timeout", None)

        try:
            struct_timeout = float(timeout_val) if timeout_val is not None else None
        except (ValueError, TypeError):
            struct_timeout = None

        if not sample_files or not isinstance(sample_files, list):
            return jsonify({"success": False, "error": "At least one sample file is required."}), 400

        if not structure_name:
            first_stem = Path(sample_files[0]).stem
            structure_name = f"{first_stem}-detection-structure"

        job_id = f"struct_samples_{int(time.time() * 1000)}"
        event_queue = queue.Queue()

        with JOBS_LOCK:
            JOBS[job_id] = {
                "id": job_id,
                "type": "structure_samples_build",
                "queue": event_queue,
                "status": "running"
            }

        def worker():
            def log_cb(msg: str):
                event_queue.put({"event": "log", "data": {"message": msg, "time": time.strftime("%H:%M:%S")}})

            try:
                log_cb(f"[*] Starting AI Detection Structure Analysis on {len(sample_files)} sample file(s)...")
                res = build_structure_from_sample_files(
                    sample_file_paths=sample_files,
                    structure_name=structure_name,
                    custom_instructions=custom_instructions,
                    log_cb=log_cb,
                    timeout=struct_timeout
                )
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "completed"
                        JOBS[job_id]["result"] = res
                event_queue.put({"event": "completed", "data": res})
            except Exception as e:
                with JOBS_LOCK:
                    if job_id in JOBS:
                        JOBS[job_id]["status"] = "error"
                        JOBS[job_id]["error"] = str(e)
                event_queue.put({"event": "error", "data": {"error": str(e)}})
            finally:
                event_queue.put({"event": "close", "data": {}})

        threading.Thread(target=worker, daemon=True).start()

        return jsonify({"success": True, "job_id": job_id, "structure_name": structure_name})

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
    @app.route("/api/models", methods=["GET"])
    def list_models():
        """
        Discovers models available from 9Router gateway.
        Supports category query ('chat', 'image', 'web', 'search', 'fetch', 'all').
        Accepts optional overrides 'url' and 'key' for real-time testing before saving.
        """
        category = request.args.get("category", "chat").strip().lower()
        url = request.args.get("url", "").strip() or None
        key = request.args.get("key", "").strip() or None
        refresh = request.args.get("refresh", "").lower() in ("1", "true", "yes")

        result = Config.get_9router_models(
            category=category,
            base_url=url,
            api_key=key,
            force_refresh=refresh
        )
        return jsonify(result)

    @app.route("/api/config", methods=["GET", "POST"])
    def handle_config():
        from pptx_jahat.config import AGENTS_CONFIG_SCHEMA
        if request.method == "GET":
            cfg_dict = {
                "NINEROUTER_URL": Config.NINEROUTER_URL,
                "NINEROUTER_KEY": Config.NINEROUTER_KEY,
                "NINEROUTER_CHAT_MODEL": Config.NINEROUTER_CHAT_MODEL,
                "NINEROUTER_SEARCH_MODEL": Config.NINEROUTER_SEARCH_MODEL,
                "NINEROUTER_FETCH_MODEL": Config.NINEROUTER_FETCH_MODEL,
                "NINEROUTER_IMAGE_MODEL": Config.NINEROUTER_IMAGE_MODEL,
                "RENDER_MODE": Config.RENDER_MODE,
                "PURE_PIL_ACTIVE": Config.PURE_PIL_ACTIVE,
                "LLM_TIMEOUT": Config.LLM_TIMEOUT,
                "VERIFICATION_ROUNDS": Config.VERIFICATION_ROUNDS
            }
            for aid in AGENTS_CONFIG_SCHEMA:
                cfg_dict[f"AGENT_MODEL_{aid.upper()}"] = getattr(Config, f"AGENT_MODEL_{aid.upper()}", "")
                cfg_dict[f"AGENT_THINK_LEVEL_{aid.upper()}"] = getattr(Config, f"AGENT_THINK_LEVEL_{aid.upper()}", "default")

            return jsonify({
                "success": True,
                "config": cfg_dict,
                "agents": Config.get_agents_config(),
                "model_metadata": Config.get_model_metadata()
            })
        else:
            data = request.get_json() or {}
            cfg = data.get("config") if isinstance(data.get("config"), dict) else dict(data)

            agents_payload = data.get("agents")
            if isinstance(agents_payload, dict):
                for aid, ainfo in agents_payload.items():
                    if isinstance(ainfo, dict):
                        if "model" in ainfo:
                            cfg[f"AGENT_MODEL_{aid.upper()}"] = ainfo.get("model", "")
                        if "think_level" in ainfo:
                            cfg[f"AGENT_THINK_LEVEL_{aid.upper()}"] = ainfo.get("think_level", "default")

            from pptx_jahat.config import BASE_DIR
            env_path = BASE_DIR / ".env"

            # Parse existing .env to preserve existing values/comments
            existing_env = {}
            if env_path.exists():
                try:
                    with open(env_path, "r", encoding="utf-8") as f:
                        for line in f:
                            line_s = line.strip()
                            if line_s and not line_s.startswith("#") and "=" in line_s:
                                ek, ev = line_s.split("=", 1)
                                existing_env[ek.strip()] = ev.strip()
                except Exception:
                    pass

            managed_keys = [
                "NINEROUTER_URL",
                "NINEROUTER_KEY",
                "NINEROUTER_CHAT_MODEL",
                "NINEROUTER_SEARCH_MODEL",
                "NINEROUTER_FETCH_MODEL",
                "NINEROUTER_IMAGE_MODEL",
                "RENDER_MODE",
                "PURE_PIL_ACTIVE",
                "LLM_TIMEOUT",
                "VERIFICATION_ROUNDS"
            ]
            for aid in AGENTS_CONFIG_SCHEMA:
                managed_keys.append(f"AGENT_MODEL_{aid.upper()}")
                managed_keys.append(f"AGENT_THINK_LEVEL_{aid.upper()}")

            for k in managed_keys:
                if k in cfg and cfg[k] is not None:
                    val = str(cfg[k]).strip()
                    existing_env[k] = val
                    os.environ[k] = val
                    if hasattr(Config, k):
                        if k == "LLM_TIMEOUT":
                            try:
                                setattr(Config, k, float(val))
                            except ValueError:
                                pass
                        elif k == "VERIFICATION_ROUNDS":
                            try:
                                setattr(Config, k, int(val))
                            except ValueError:
                                pass
                        elif k == "PURE_PIL_ACTIVE":
                            setattr(Config, k, val.lower() in ("1", "true", "yes", "on"))
                        else:
                            setattr(Config, k, val)

            env_lines = [f"{k}={v}" for k, v in existing_env.items()]
            env_path = Path(__file__).resolve().parent.parent.parent.parent / ".env"
            if env_path.exists():
                try:
                    backup_path = env_path.with_suffix(".env.backup")
                    shutil.copy2(env_path, backup_path)
                except Exception:
                    pass
            with open(env_path, "w", encoding="utf-8") as f:
                f.write("\n".join(env_lines) + "\n")

            Config.reload(env_path)

            resp_cfg = {
                "NINEROUTER_URL": Config.NINEROUTER_URL,
                "NINEROUTER_KEY": Config.NINEROUTER_KEY,
                "NINEROUTER_CHAT_MODEL": Config.NINEROUTER_CHAT_MODEL,
                "NINEROUTER_SEARCH_MODEL": Config.NINEROUTER_SEARCH_MODEL,
                "NINEROUTER_FETCH_MODEL": Config.NINEROUTER_FETCH_MODEL,
                "NINEROUTER_IMAGE_MODEL": Config.NINEROUTER_IMAGE_MODEL,
                "RENDER_MODE": Config.RENDER_MODE,
                "PURE_PIL_ACTIVE": Config.PURE_PIL_ACTIVE,
                "LLM_TIMEOUT": Config.LLM_TIMEOUT,
                "VERIFICATION_ROUNDS": Config.VERIFICATION_ROUNDS
            }
            for aid in AGENTS_CONFIG_SCHEMA:
                resp_cfg[f"AGENT_MODEL_{aid.upper()}"] = getattr(Config, f"AGENT_MODEL_{aid.upper()}", "")
                resp_cfg[f"AGENT_THINK_LEVEL_{aid.upper()}"] = getattr(Config, f"AGENT_THINK_LEVEL_{aid.upper()}", "default")

            return jsonify({
                "success": True,
                "message": f"Configuration saved to .env and applied. Model: {Config.NINEROUTER_CHAT_MODEL}",
                "model": Config.NINEROUTER_CHAT_MODEL,
                "config": resp_cfg,
                "agents": Config.get_agents_config(),
                "model_metadata": Config.get_model_metadata(),
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
