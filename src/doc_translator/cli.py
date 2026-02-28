import argparse
from pathlib import Path
from .main import translate_file


def find_project_root() -> Path:
    p = Path.cwd().resolve()
    for d in [p] + list(p.parents):
        if (d / "pyproject.toml").exists() or (d / "README.md").exists() or (d / ".git").exists():
            return d
    return p


def main():
    parser = argparse.ArgumentParser(description="Translate all supported files in origin/ to a target language or launch GUI")
    parser.add_argument("lang", nargs="?", help="Target language code (e.g., 'es')")
    parser.add_argument("--gui", action="store_true", help="Launch the GUI instead of running batch CLI")
    args = parser.parse_args()

    root = find_project_root()
    if args.gui:
        try:
            from .gui_customtk import main as gui_main
        except Exception:
            print("Failed to import GUI module. Ensure `customtkinter` is installed and the GUI files exist.")
            return
        gui_main()
        return
    origin = root / "origin"
    output = root / "output"
    origin.mkdir(parents=True, exist_ok=True)
    output.mkdir(parents=True, exist_ok=True)

    supported_exts = (".docx", ".pdf", ".xlsx", ".xls")
    files = [p for p in origin.iterdir() if p.is_file() and p.suffix.lower() in supported_exts]
    if not files:
        print(f"No supported files found in {origin}")
        return

    if not args.lang:
        print("Error: target language code is required for batch CLI. Example: python -m src.doc_translator.cli es")
        return

    for f in files:
        try:
            translate_file(str(f), args.lang, str(output))
        except Exception as e:
            print(f"Error translating {f.name}: {e}")


if __name__ == "__main__":
    main()
