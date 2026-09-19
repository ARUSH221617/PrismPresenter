import sys
import os
import argparse
import threading
import webbrowser
import time
from pathlib import Path
from pptx_jahat.web.app import create_app

def main() -> None:
    parser = argparse.ArgumentParser(
        prog="pptx-jahat",
        description="PrismPresenter — Autonomous Multi-Modal AI Presentation Studio"
    )
    parser.add_argument(
        "--cli",
        action="store_true",
        help="Launch interactive CLI AI agent session"
    )
    parser.add_argument(
        "--dev", "-d", "--reload",
        dest="dev",
        action="store_true",
        help="Run in development mode with auto-reload (Flask reloader & live-reload)"
    )
    parser.add_argument(
        "--no-reload",
        action="store_true",
        help="Disable auto-reload even if dev environment variables are set"
    )
    parser.add_argument(
        "--no-browser",
        action="store_true",
        help="Do not automatically open the web browser on startup"
    )
    parser.add_argument(
        "--host",
        default=os.getenv("HOST", "127.0.0.1"),
        help="Host interface to bind to (default: 127.0.0.1 or $HOST)"
    )
    parser.add_argument(
        "--port", "-p",
        type=int,
        default=int(os.getenv("PORT", 5000)),
        help="Port to listen on (default: 5000 or $PORT)"
    )

    args, unknown = parser.parse_known_args()

    if args.cli:
        from pptx_jahat.agent import AIAgent
        print("PrismPresenter AI Agent Interactive CLI. (Type 'exit' to quit)")
        agent = AIAgent()
        while True:
            try:
                user_input = input("\nYou > ").strip()
                if not user_input or user_input.lower() in ("exit", "quit"):
                    break
                print("\nAgent Thinking...\n")
                result = agent.run(user_input, log_callback=lambda m: print(f"  {m}"))
                print(f"\nResponse:\n{result}")
            except (KeyboardInterrupt, EOFError):
                break
    else:
        port = args.port
        host = args.host
        url = f"http://{host}:{port}"

        env_dev = (
            os.getenv("FLASK_DEBUG", "").lower() in ("1", "true", "yes")
            or os.getenv("DEV", "").lower() in ("1", "true", "yes")
            or os.getenv("DEBUG", "").lower() in ("1", "true", "yes")
            or os.getenv("RELOAD", "").lower() in ("1", "true", "yes")
            or os.getenv("FLASK_ENV", "").lower() == "development"
        )
        is_dev = bool((args.dev or env_dev) and not args.no_reload)
        no_browser = bool(args.no_browser or os.getenv("NO_BROWSER", "").lower() in ("1", "true", "yes"))

        # In dev mode with Werkzeug reloader, avoid opening the browser repeatedly on each reload
        is_reloader_child = os.environ.get("WERKZEUG_RUN_MAIN") == "true"

        if not no_browser and not is_reloader_child:
            def open_browser():
                time.sleep(1.2)
                try:
                    webbrowser.open(url)
                except Exception:
                    pass

            threading.Thread(target=open_browser, daemon=True).start()

        # Extra files to watch in dev mode
        extra_files = []
        env_path = Path(__file__).resolve().parent.parent.parent / ".env"
        if env_path.exists():
            extra_files.append(str(env_path))

        print(f"\n=======================================================")
        if is_dev:
            print(f"  ⚡ PRISMPRESENTER DEV SERVER (AUTO-RELOAD ACTIVE)")
            print(f"  Live Watch: Python sources, templates, static assets & .env")
        else:
            print(f"  ⚡ PRISMPRESENTER WEB GUI RUNNING AT: {url}")
        print(f"  URL: {url}")
        print(f"  Press Ctrl+C to stop server")
        print(f"=======================================================\n")

        app = create_app(dev_mode=is_dev)
        if is_dev:
            app.run(host=host, port=port, debug=True, use_reloader=True, extra_files=extra_files if extra_files else None)
        else:
            app.run(host=host, port=port, debug=False)

if __name__ == "__main__":
    main()
