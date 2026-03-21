<p align="center">
  <img src=".\src\verbilo\assets/favicon.jpg" alt="Verbilo logo" width="220" />
</p>

<div align="center">

[![Releases](https://img.shields.io/github/v/release/8041q/Verbilo)](https://github.com/8041q/Verbilo/releases)
[![Release Notes](https://img.shields.io/badge/release-notes-blue)](https://github.com/8041q/Verbilo/releases)
[![Python](https://img.shields.io/badge/python-3.12%2B-brightgreen)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-AGPL)](https://www.gnu.org/licenses/agpl-3.0.en.html)
[![Stars](https://img.shields.io/github/stars/8041q/Verbilo?style=flat)](https://github.com/8041q/Verbilo/stargazers)
[![Issues](https://img.shields.io/github/issues/8041q/Verbilo)](https://github.com/8041q/Verbilo/issues)

</div>

### Verbilo — Portable
<p align="center"><em>Translate DOCX, XLSX and PDF into languages while preserving layout, styles, and images.</em></p>

---

## How it works (high level)

- Converts document content into translation units (runs, spans, rows, paragraphs)
- Sends grouped units in batches to translation backends with resilient HTTP retries and sub-batch fallbacks for large requests.
- Uses API-aware inline tagging where supported to preserve run/span formatting across the round-trip.
- Reconstructs translated text back into the original document structure, applying formatting where feasible.


## Feature Highlights

- **Multiple translation engines**: Google Translate (free), Google Cloud Translation API, Baidu, Azure, DeepL.
- **Proxy & resilience**: All engines use a resilient HTTP session with retries, backoff, timeouts, and optional HTTPS/HTTP proxy.
- **Selective Translation**: Translate only text in a specified source language (or use auto).
- **Formatting Preserved**: DOCX run-level, XLSX cell and in-place PDF Editing are preserved. 
- **Multi-Engine Detection**: Lingua, FastText — choose your preferred engine. (quality / speed)
- **Batching for Efficiency**: Segments are batched to reduce API calls and avoid rate limits.

## Known limitations

- Tag survival is API-dependent; inline tag preservation is not guaranteed on every backend.
- Z-order guard is conservative: it avoids translating text entirely covered by opaque graphics rather than rewriting PDF content streams to change stacking order.
- Very-short CJK tokens (1–2 characters) can behave inconsistently across translation APIs—prefer explicit source_lang to ensure correct source language.
- Extremely complex layouts (heavy overlays, rotated text, or nonstandard encodings) can still produce visual artifacts - manual verification recommended for critical documents.

## Quick Start — For Developers

1. Install core dependencies:

```bash
pip install -r requirements.txt
```

2. Launch the GUI:

```bash
cd src
python -m launch
```

3. If using GUI, and errors for UI helpers or icons appear:

```bash
pip install customtkinter
pip install pytablericons Pillow
```

***

### GUI translation engines & network settings

In the GUI sidebar you can choose the translation engine:

- Google Translate (free) ~ default, no API key required.
- Google Cloud Translation API ~ requires: API key for v2, Project ID & Account Credentials for v3.
- Baidu Translate ~ requires Baidu App ID and App Key.
- Microsoft Azure Translator ~ requires a Subscription Key and Region.
- DeepL ~ requires a DeepL API key (Free or Pro).
- Local (offline) ~ free and unlimited use, requires download of each language model source+target

**Settings → Network & API keys** to configure

If any API method is selected without credentials, the GUI will show a warning instead of starting the job.



</details>


## Project Structure

<details>
<summary>Click to expand</summary>

```
src/
  Origin/
  Output/
  verbilo/
    launch.py
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
      azure.py
      baidu.py
      base.py
      cache.py
      deepl.py
      factory.py
      google.py
      http_session.py
      lang_detect.py
      local.py
      usage.py
    assets/
      __init__.py
    utils/
      io.py
pyproject.toml
requirements.txt
README.md
```

</details>

## Nuitka build (Windows)

Prerequisites:

- A Python virtual environment (recommended) activated.
- `nuitka` installed in the virtualenv (`pip install nuitka`).

Build the GUI executable:

```bash
# From the repository root, with your virtualenv active
.venv\Scripts\python.exe scripts\build_nuitka.py --entry gui --output dist/nuitka
```

Notes:

- For a final GUI build without a console window, pass the flag `--windows-console-mode=disable` to the underlying Nuitka command (the helper script already exposes this behavior when appropriate).
- If build fails and you try with changes, clean the Nuitka cache at `%LOCALAPPDATA%\Nuitka\Nuitka\`

Run the built GUI directly by double-clicking the `verbilo.exe` in Explorer to launch without the console.

Troubleshooting:

- You will need the language model used by fasttext detector, run `download_models.py` to download `models/lid.176.bin`. For Local use, you will also need the OPUS-MT model, each language source+target is a diferent model.
- If paths or behavior differ, confirm you executed the commands from the repository root and that your virtualenv has `nuitka` installed.


## Requirements & Notes

- **Python**: 3.12+  
- **Install (has all)**: `pip install -r requirements.txt`  
- GUI extras: `pip install customtkinter pytablericons Pillow`  
- Detection engines:
  - Lingua: high accuracy for short strings (heavier).
  - FastText: very fast, good balance.

Notes:
- Scanned (image-only) PDFs are detected and skipped - they will be logged rather than producing broken output.  
- When a specific source language is set, local detection prevents unnecessary API translation calls.  
- Verbilo batches segments to reduce API usage and avoid rate limits.


## Contributing

Contributions are what makes open source great!

- Found a bug? Open an issue with steps to reproduce and a sample file if possible
- Want to add features? Fork, create a feature branch, and open a PR referencing the issue
- PR checklist:
  - Keep changes focused and minimal
  - Follow existing code style
- Development tips:
  - Run unit tests locally before
  - Any change needs to be documented, even if small


## Acknowledgments

- Tabler Icons / `pytablericons` for GUI icons.  
- PyMuPDF for in-place PDF text editing.  
- Lingua, FastText for language detection options.   
- Everyone who files issues and contributes patches.


## License

This project is released under the **GNU Affero General Public License v3 (AGPL-3.0-or-later)** — see the LICENSE file for details.
