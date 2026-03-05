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
    parser = argparse.ArgumentParser(
        description="Translate all supported files in origin/ to a target language or launch GUI",
    )
    parser.add_argument("lang", nargs="?", help="Target language code (e.g., 'es')")
    parser.add_argument("--source", "-s", default="auto",
                        help="Source language code (e.g. 'en'). 'auto' = translate all text (default)")
    parser.add_argument("--detector", "-d", default="fasttext",
                        choices=["fasttext", "lingua"],
                        help="Language detector engine (default: fasttext)")
    parser.add_argument("--gui", action="store_true", help="Launch the GUI instead of running batch CLI")
    args = parser.parse_args()

    root = find_project_root()
    if args.gui:
        try:
            from .gui.app import main as gui_main
        except Exception:
            import traceback
            traceback.print_exc()
            return
        gui_main()
        return
    origin = root / "origin"
    output = root / "output"
    output.mkdir(parents=True, exist_ok=True)

    supported_exts = (".docx", ".pdf", ".xlsx", ".xls")
    files = [p for p in origin.iterdir() if p.is_file() and p.suffix.lower() in supported_exts]
    if not files:
        print(f"No supported files found in {origin}")
        return

    if not args.lang:
        print("Error: target language code is required for batch CLI. Example: python -m src.verbilo.cli es")
        return

    print(f"Source language: {args.source}  |  Target language: {args.lang}")
    for f in files:
        try:
            result = translate_file(str(f), args.lang, str(output), source_lang=args.source, detector=args.detector)
            if result == "skipped-ocr":
                print(f"Skipped {f.name} (scanned/image PDF requiring OCR)")
            else:
                print(f"Translated {f.name}")
        except Exception as e:
            print(f"Error translating {f.name}: {e}")


if __name__ == "__main__":
    main()
