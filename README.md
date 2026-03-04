<p align="center">
  <img src=".\src\verbilo\assets/favicon.jpg" alt="Verbilo logo" width="220" />
</p>

<div align="center">

[![Releases](https://img.shields.io/github/v/release/8041q/Verbilo)](https://github.com/8041q/Verbilo/releases)
[![Release Notes](https://img.shields.io/badge/release-notes-blue)](https://github.com/8041q/Verbilo/releases)
[![Python](https://img.shields.io/badge/python-3.12%2B-brightgreen)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](https://opensource.org/licenses/MIT)
[![Stars](https://img.shields.io/github/stars/8041q/Verbilo?style=flat)](https://github.com/8041q/Verbilo/stargazers)
[![Issues](https://img.shields.io/github/issues/8041q/Verbilo)](https://github.com/8041q/Verbilo/issues)

</div>

### Verbilo — Portable
<p align="center"><em>Translate DOCX, XLSX and PDF into 130+ languages while preserving layout, styles, and images.</em></p>

---

## Why Verbilo?

 
Translating office documents should not mean losing formatting, tables, fonts, or images. Most translation tools export plain text or break layouts — Verbilo preserves the original file fidelity while translating only the text that needs translation (tries to :/).

**Problem → Solution**  
- Problem: Document translators commonly strip formatting, corrupt tables, or require manual rework.  
- Solution: Verbilo segments and selectively translates text while writing translations back into the original file structure (DOCX runs, XLSX cells, in-place PDF text).



## Feature Highlights

- **130+ Languages**: Translate to any language supported by Google Translate.  
- **Selective Translation**: Translate only text in a specified source language (or use auto).  
- **Formatting Preserved**: DOCX run-level styles and XLSX cell styles are preserved.  
- **In-place PDF Editing**: Uses PyMuPDF to replace text without breaking layout.  
- **Multi-Engine Detection**: Lingua, FastText, LangDetect — or Auto (majority vote).  
- **Batching for Efficiency**: Segments are batched to reduce API calls and avoid rate limits.  
- **Scanned-PDF Detection**: Image-only PDFs are detected and skipped with a log message.  
- **Segment-Aware**: Keeps technical strings intact (splits on `/` and newlines for safe translation).



## Quick Start — For Developers

1. Install core dependencies:

```bash
pip install -r requirements.txt
```

3. (GUI extras) Install UI helpers and icons (if any error appears):

```bash
pip install customtkinter
pip install pytablericons Pillow
```

4. Launch the GUI:

```bash
python -m src.verbilo.cli --gui
```

5. Or run CLI translations directly (examples below):

```bash
# Translate all supported files in `origin/` to Spanish
python -m src.verbilo.cli es

# Translate only English segments to Portuguese
python -m src.verbilo.cli pt --source en
```



## CLI Reference

Usage:

```text
python -m src.verbilo.cli [LANG] [--source CODE] [--detector ENGINE] [--gui]
```

| Argument | Description |
|---|---|
| `LANG` | Target language code (e.g., `es`, `pt`) |
| `--source CODE` | Source language code (default: `auto`) |
| `--detector ENGINE` | Detection engine: `auto` (default), `lingua`, `fasttext`, or `langdetect` |
| `--gui` | Launch the graphical interface |

Example commands:

```bash
python -m src.verbilo.cli es
python -m src.verbilo.cli pt --source en
python -m src.verbilo.cli --gui
```



## Programmatic API (quick example)

Use the core API when embedding Verbilo into scripts. `translate_file()` is the project’s core helper available in `src/verbilo/main.py`.

```python
from src.verbilo.main import translate_file

# This is an example; adapt params to match your needs.
translate_file("input.docx", "output.docx", target="es", source="auto")
```

(See `src/verbilo/main.py` for the exact function signature and options.)



## Project Structure

<details>
<summary>Click to expand the repository tree</summary>

```
output/          ← translated files (auto-created if needed)
src/
  verbilo/
    cli.py
    main.py         - `translate_file()` core API
    gui/
      app.py        - CustomTkinter GUI
      config.py
      helpers.py
      theme.py
      icons.py
    converters/
      docx_converter.py
      xlsx_converter.py
      pdf_converter.py
    translators/
      base.py
      dummy.py
      google.py
      lang_detect.py
    assets/
      __init__.py
    utils/
      io.py
pyproject.toml
requirements.txt
README.md
```

</details>



## Requirements & Notes

- **Python**: 3.12+  
- **Install (has all)**: `pip install -r requirements.txt`  
- GUI extras: `pip install customtkinter pytablericons Pillow`  
- Detection engines:
  - Lingua: high accuracy for short strings (heavier).  
  - FastText: very fast, good balance.  
  - LangDetect: lightweight.  
  - Auto: majority-vote across engines for robustness.

Notes:
- Scanned (image-only) PDFs are detected and skipped — they will be logged rather than producing broken output.  
- When a specific source language is set, local detection prevents unnecessary API translation calls.  
- Verbilo batches segments to reduce API usage and avoid rate limits.



## Contributing

Contributions are what makes open source great!

- Found a bug? Open an issue with steps to reproduce and a sample file if possible.  
- Want to add features? Fork, create a feature branch, and open a PR referencing the issue.  
- PR checklist:
  - Add tests for new behavior where practical. 
  - Keep changes focused and minimal. 
  - Follow existing code style. Run linters/formatters before submitting. 
  - Describe the change and rationale in the PR body.
- Development tips:
  - Use a virtual environment.  
  - Run unit tests (if present) locally before opening a PR.  
  - If changing translations or detectors, include small sample files demonstrating behavior.


## Acknowledgments

- Tabler Icons / `pytablericons` for GUI icons.  
- PyMuPDF for in-place PDF text editing.  
- Lingua, FastText, LangDetect for language detection options.  
- Google Translate (or the configured translation provider) for the translation backend.  
- Everyone who files issues and contributes patches.


## License

This project is released under the **MIT License** — see the `LICENSE` file for details.

