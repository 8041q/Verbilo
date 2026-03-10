#!/usr/bin/env python3
# Usage: python scripts/build_nuitka.py --entry cli|gui --output dist/nuitka
import argparse
import os
import shlex
import subprocess
import sys

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--entry", choices=["cli", "gui"], default="cli")
    parser.add_argument("--output", default="dist")
    parser.add_argument("--python", default=sys.executable)
    args = parser.parse_args()

    repo_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    src_dir = os.path.join(repo_root, "src")

    # cli.py handles both modes: run as-is for CLI, pass --gui for the GUI.
    if args.entry == "cli":
        entry = os.path.join(src_dir, "verbilo", "cli.py")
    else:
        entry = os.path.join(src_dir, "verbilo_launcher.py")

    # Ensure output lives under the repository root unless an absolute
    if os.path.isabs(args.output):
        outdir = os.path.abspath(args.output)
    else:
        outdir = os.path.abspath(os.path.join(repo_root, args.output))
    os.makedirs(outdir, exist_ok=True)

    assets_abs_path = os.path.join(repo_root, "src", "verbilo", "assets")
    models_abs_path = os.path.join(repo_root, "models")
    favicon_ico = os.path.join(assets_abs_path, "favicon.ico")

    flags = [
        "--standalone",
        f"--output-dir={outdir}",

        # ── Tkinter / GUI ────────────────────────────────────────────────
        "--enable-plugin=tk-inter",
        "--include-package=PIL",
        "--include-package=pytablericons",
        "--include-package-data=customtkinter",

        # ── Language detectors (dynamically imported inside functions) ───
        "--include-package=fast_langdetect",
        "--include-package=fasttext",
        "--include-package=lingua",
        "--include-package-data=lingua",

        # ── Translation engine ───────────────────────────────────────────
        "--include-package=deep_translator",

        # ── Document converters ──────────────────────────────────────────
        "--include-package=docx",
        "--include-package-data=docx",
        "--include-package=openpyxl",
        "--include-package-data=openpyxl",

        # ── PDF (PyMuPDF) ────────────────────────────────────────────────
        # fitz.mupdf is a SWIG C extension that expands to ~1.75M lines of C.
        # MSVC runs out of heap on this file; Clang handles it far better.
        "--clang", # use Clang-cl (installed via Visual Studio) instead of MSVC
        "--lto=no", # disable link-time optimisation to reduce peak memory further
        "--low-memory", # serialise compilation and reduce Nuitka-side RAM usage
        "--include-package=fitz",

        # ── Verbilo assets ───────────────────────────────────────────────
        # icons.py resolves: Path(__file__).parent.parent / "assets" / "favicon.*" which maps to verbilo/assets/ inside the standalone dist folder.
        f"--include-data-files={favicon_ico}=verbilo/assets/favicon.ico",
        f"--include-data-files={os.path.join(assets_abs_path, 'favicon.jpg')}=verbilo/assets/favicon.jpg",

        # ── FastText language detection model ────────────────────────────
        # Bundled at <dist>/models/lid.176.bin; lang_detect.py points. FTLANG_CACHE to <dist>/models/ at runtime so fast_langdetect
        f"--include-data-files={os.path.join(models_abs_path, 'lid.176.bin')}=models/lid.176.bin",
    ]

    # Windows app icon embedded into the .exe
    if os.path.isfile(favicon_ico):
        flags.append(f"--windows-icon-from-ico={favicon_ico}")

    # GUI builds: hide the console window on Windows.
    # Comment this line out while debugging so you can see tracebacks.
    """if args.entry == "gui":
        flags.append("--windows-console-mode=disable")"""

    cmd = [args.python, "-m", "nuitka"] + flags + [entry]

    print("Running:", " ".join(shlex.quote(p) for p in cmd))
    subprocess.check_call(cmd)

    print("Built Nuitka output in:", outdir)


if __name__ == "__main__":
    main()