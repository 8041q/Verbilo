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
    parser.add_argument("--engine", "-e", default="google",
                        choices=["google", "google-cloud", "baidu"],
                        help="Translation engine (default: google)")
    parser.add_argument("--proxy", default=None,
                        help="HTTPS proxy URL (e.g. http://127.0.0.1:7890)")
    parser.add_argument("--google-api-key", default="",
                        help="Google Cloud Translation API key (for --engine google-cloud)")
    parser.add_argument("--baidu-appid", default="",
                        help="Baidu Translate App ID (for --engine baidu)")
    parser.add_argument("--baidu-appkey", default="",
                        help="Baidu Translate App Key (for --engine baidu)")
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

    proxies = {"https": args.proxy, "http": args.proxy} if args.proxy else None

    print(f"Source language: {args.source}  |  Target language: {args.lang}  |  Engine: {args.engine}")
    for f in files:
        try:
            result = translate_file(
                str(f), args.lang, str(output),
                source_lang=args.source,
                detector=args.detector,
                engine=args.engine,
                proxies=proxies,
                google_api_key=args.google_api_key,
                baidu_appid=args.baidu_appid,
                baidu_appkey=args.baidu_appkey,
            )
            if result == "skipped-ocr":
                print(f"Skipped {f.name} (scanned/image PDF requiring OCR)")
            else:
                print(f"Translated {f.name}")
        except Exception as e:
            print(f"Error translating {f.name}: {e}")


if __name__ == "__main__":
    main()
