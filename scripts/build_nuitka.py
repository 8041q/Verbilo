#!/usr/bin/env python3
# Usage: python scripts/build_nuitka.py --entry cli|gui --output dist/nuitka
import argparse
import os
import shlex
import subprocess
import sys
from pathlib import Path

def write_version_from_pyproject():
    root = Path(__file__).resolve().parents[1]
    py = root / "pyproject.toml"
    if not py.exists():
        return
    import tomllib

    data = tomllib.loads(py.read_text())
    poetry = data.get("tool", {}).get("poetry", {})
    version = (
        data.get("project", {}).get("version")
        or poetry.get("version")
    )
    if not version:
        return
    build_date = (
        data.get("project", {}).get("build_date")
        or poetry.get("build_date")
        or ""
    )
    target = root / "src" / "verbilo" / "_version.py"
    target.write_text(
        "# Auto-generated at build time - do not edit, this file is overwritten on every build.\n"
        f'__version__ = "{version}"\n'
        f'__build_date__ = "{build_date}"\n'
    )

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--entry", choices=["cli", "gui"], default="cli")
    parser.add_argument("--output", default="dist")
    parser.add_argument("--python", default=sys.executable)
    args = parser.parse_args()
    write_version_from_pyproject()

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
    version_abs_path = os.path.join(repo_root, "src", "verbilo",)
    os.environ["PYTHONPATH"] = src_dir + os.pathsep + os.environ.get("PYTHONPATH", "")

    flags = [
        "--mode=standalone",
        f"--output-dir={outdir}",
        "--include-package=verbilo",
        #"--prefer-source-code", # will make build 2x/3x longer, but it's the most optimizing of the code
        "--assume-yes-for-downloads", # to make sure Nuitka is using cache to save time on next build

        # ── Tkinter / GUI ────────────────────────────────────────────────
        "--enable-plugin=tk-inter",
        "--include-package=PIL",
        "--include-package=pytablericons",
        "--include-package-data=pytablericons",
        "--include-package=customtkinter",
        "--include-package-data=customtkinter",

        # ── Language detectors (dynamically imported inside functions) ───
        "--include-package=fast_langdetect",
        "--include-package=fasttext",
        "--include-package=lingua",
        "--include-package-data=lingua",

        # ── Translation engine ───────────────────────────────────────────
        "--include-package=deep_translator",
        # ── System / runtime ────────────────────────────────────────────────
        "--include-package=platformdirs",
        # ── Document converters ──────────────────────────────────────────
        "--include-package=docx",
        "--include-package-data=docx",
        "--include-package=openpyxl",
        "--include-package-data=openpyxl",

        # ── PDF (PyMuPDF) ────────────────────────────────────────────────
        # fitz.mupdf is a SWIG C extension that expands to ~1.75M lines of C. MSVC runs out of heap on this file; Clang handles it far better.
        # Best to do is use --lto=no and --low-memory on first build, and then there is no need for it
        "--clang", # use Clang-cl (installed via Visual Studio) instead of MSVC
        "--lto=no", # disables link-time optimisation to reduce peak memory further
        "--low-memory", # serialise compilation and reduce Nuitka-side RAM usage
        "--enable-plugin=pylint-warnings",

        # ── Verbilo assets ───────────────────────────────────────────────
        # icons.py resolves: Path(__file__).parent.parent / "assets" / "favicon.*" which maps to verbilo/assets/ inside the standalone dist folder.
        f"--include-data-files={favicon_ico}=verbilo/assets/favicon.ico",
        f"--include-data-files={os.path.join(assets_abs_path, 'favicon.jpg')}=verbilo/assets/favicon.jpg",

        # ── FastText language detection model ────────────────────────────
        # Bundled at <dist>/models/lid.176.bin; lang_detect.py points. FTLANG_CACHE to <dist>/models/ at runtime so fast_langdetect
        f"--include-data-files={os.path.join(models_abs_path, 'lid.176.bin')}=models/lid.176.bin",
        "--nofollow-import-to=fasttext.tests",
        "--nofollow-import-to=fasttext.tests.test_script",

        # ── Model-download utilities: bytecode is fine ───────────────────
        # tqdm and robust_downloader are used only to download the FastText model at first run.  The model is pre-bundled (lid.176.bin)

        # ── Compilation report (for diagnostics) ─────────────────────────
        "--report=compilation-report.xml",

        # ── App Version Control ──────────────────────────────────────────
        "--include-module=verbilo._version",
        
        # ── Optimization ─────────────────────────────────────────────────
        "--python-flag=no_site",
        "--nofollow-import-to=81d243bd2c585b0f4821__mypyc",
        "--nofollow-import-to=__pycache__",
        "--nofollow-import-to=packaging",
        "--nofollow-import-to=pip",
        "--nofollow-import-to=pygame",
        "--nofollow-import-to=robust_downloader",
        "--nofollow-import-to=tqdm",
        "--collect-all=fitz",

        f"--output-filename=verbilo",
    ]

    # Windows app icon embedded into the .exe
    if os.path.isfile(favicon_ico):
        flags.append(f"--windows-icon-from-ico={favicon_ico}")

    # GUI builds: hide the console window on Windows.
    # Comment this line out while debugging so you can see tracebacks.
    if args.entry == "gui":
        flags.append("--windows-console-mode=disable")

    cmd = [args.python, "-m", "nuitka"] + flags + [entry]

    print("Running:", " ".join(shlex.quote(p) for p in cmd))
    subprocess.check_call(cmd)

    print("Built Nuitka output in:", outdir)


if __name__ == "__main__":
    main()