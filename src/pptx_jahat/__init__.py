import sys

def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
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
        import os
        import threading
        import webbrowser
        import time
        from pptx_jahat.web.app import create_app

        port = int(os.getenv("PORT", 5000))
        host = os.getenv("HOST", "127.0.0.1")
        url = f"http://{host}:{port}"

        def open_browser():
            time.sleep(1.2)
            try:
                webbrowser.open(url)
            except Exception:
                pass

        threading.Thread(target=open_browser, daemon=True).start()

        print(f"\n=======================================================")
        print(f"  ⚡ PRISMPRESENTER WEB GUI RUNNING AT: {url}")
        print(f"  Press Ctrl+C to stop server")
        print(f"=======================================================\n")

        app = create_app()
        app.run(host=host, port=port, debug=False)

if __name__ == "__main__":
    main()
