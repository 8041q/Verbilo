#!/usr/bin/env python3
# Download helper for FastText and OPUS-MT models used by the app.
import argparse
import os
import sys
from pathlib import Path
from urllib.request import urlopen


FASTTEXT_URL = "https://dl.fbaipublicfiles.com/fasttext/supervised-models/lid.176.bin"

_REPO_ROOT = Path(__file__).resolve().parents[1]

# Default destination is anchored to the repo root, regardless of where the script is run from.
_DEFAULT_DEST = str(_REPO_ROOT / "models" / "lid.176.bin")

# OPUS-MT model base URL on HuggingFace (Helsinki-NLP) — used as fallback for .spm files.
_HF_BASE = "https://huggingface.co/Helsinki-NLP/opus-mt-{src}-{tgt}/resolve/main"

# Default directory for converted OPUS-MT CTranslate2 models.
_DEFAULT_OPUS_DIR = str(_REPO_ROOT / "models" / "opus-mt")

# Sentinel written after a successful CTranslate2 conversion.
_SENTINEL = "converted.ok"


def download(url, dest_path):
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    print(f"Downloading {url} -> {dest_path}")
    with urlopen(url) as r, open(dest_path, "wb") as f:
        f.write(r.read())


def download_opus_mt(src: str, tgt: str, dest_dir: str | None = None) -> Path:
    """Download a Helsinki-NLP OPUS-MT model and convert it to CTranslate2 format.

    Uses ``ctranslate2.converters.TransformersConverter`` which loads the model
    via the HuggingFace ``transformers`` library.  ``transformers`` is therefore
    required at **conversion time** only — it is **not** needed at translation
    runtime.

    The converted model is saved under ``<dest_dir>/<src>-<tgt>/`` with a
    ``converted.ok`` sentinel file that downstream code checks before loading.
    """
    dest_dir = dest_dir or _DEFAULT_OPUS_DIR
    out_path = Path(dest_dir) / f"{src}-{tgt}"

    # Already converted — skip.
    if (out_path / _SENTINEL).exists():
        print(f"Model {src}-{tgt} already converted at {out_path}")
        return out_path

    model_name = f"Helsinki-NLP/opus-mt-{src}-{tgt}"

    # Ensure required conversion-time dependencies are present.
    try:
        import ctranslate2   # noqa: F401
    except ImportError:
        print("ctranslate2 is not installed. Run: pip install ctranslate2", file=sys.stderr)
        sys.exit(1)
    try:
        import transformers  # noqa: F401
    except ImportError:
        print(
            "transformers is not installed (required for model conversion only). "
            "Run: pip install transformers",
            file=sys.stderr,
        )
        sys.exit(1)

    print(f"Converting {model_name} to CTranslate2 format ...")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    converter = ctranslate2.converters.TransformersConverter(
        model_name,
        copy_files=["source.spm", "target.spm", "tokenizer_config.json"],
    )
    converter.convert(str(out_path), force=True)

    # Also download SentencePiece models via direct URL if TransformersConverter
    # did not copy them (some older ctranslate2 versions skip copy_files).
    base_url = _HF_BASE.format(src=src, tgt=tgt)
    for sp_file in ("source.spm", "target.spm"):
        dest_sp = out_path / sp_file
        if not dest_sp.exists():
            url = f"{base_url}/{sp_file}"
            try:
                download(url, str(dest_sp))
            except Exception as exc:
                print(f"Warning: could not download {sp_file}: {exc}", file=sys.stderr)

    # Write sentinel so OpusMTTranslator knows the conversion is complete.
    (out_path / _SENTINEL).write_text("ok\n")
    print(f"Model {src}-{tgt} ready at {out_path}")
    return out_path


def main():
    parser = argparse.ArgumentParser(
        description="Download FastText and/or OPUS-MT models for Verbilo.",
    )
    sub = parser.add_subparsers(dest="command")

    # Sub-command: fasttext (original behaviour, also the default)
    ft = sub.add_parser("fasttext", help="Download the FastText language-detection model")
    ft.add_argument("--dest", default=_DEFAULT_DEST)

    # Sub-command: opus-mt
    opus = sub.add_parser("opus-mt", help="Download and convert an OPUS-MT translation model")
    opus.add_argument("src", help="Source language code, e.g. 'en'")
    opus.add_argument("tgt", help="Target language code, e.g. 'pt'")
    opus.add_argument("--dest-dir", default=_DEFAULT_OPUS_DIR,
                       help="Parent directory for converted models")

    args = parser.parse_args()

    if args.command == "opus-mt":
        try:
            download_opus_mt(args.src, args.tgt, args.dest_dir)
        except Exception as e:
            print(f"OPUS-MT download failed: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        # Default: download FastText model (preserves original behaviour).
        dest = getattr(args, "dest", _DEFAULT_DEST)
        try:
            download(FASTTEXT_URL, dest)
        except Exception as e:
            print("Download failed:", e, file=sys.stderr)
            sys.exit(1)
        print("Download complete")


if __name__ == "__main__":
    main()
