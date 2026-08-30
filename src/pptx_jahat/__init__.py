import sys
from pptx_jahat.gui.app import run_gui
from pptx_jahat.agent import AIAgent

def main() -> None:
    if len(sys.argv) > 1 and sys.argv[1] == "--cli":
        print("PPTX Jahat AI Agent Interactive CLI. (Type 'exit' to quit)")
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
        run_gui()

if __name__ == "__main__":
    main()
