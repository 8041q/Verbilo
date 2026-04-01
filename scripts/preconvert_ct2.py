#!/usr/bin/env python3
"""Pre-convert OPUS-MT models to CTranslate2 format.

Run this script in your development environment (where `transformers` is
installed) to produce ready-to-use CT2 model directories.  The output can
then be uploaded to a HuggingFace repository so that standalone/Nuitka
builds can download them directly without needing `transformers`.

Models that already have third-party CT2 repos (gaudi/*-ctranslate2) do
NOT need this script — they are handled by _download_ct2_direct().

Usage:
    python scripts/preconvert_ct2.py                      # convert all missing
    python scripts/preconvert_ct2.py --slug en-pt          # convert one model
    python scripts/preconvert_ct2.py --output-dir ./ct2out # custom output dir

After conversion, upload each folder to HuggingFace and add a "ct2_repo"
field to models_catalogue.json.
"""

import argparse
import json
import shutil
import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
_CATALOGUE = _REPO_ROOT / "src" / "verbilo" / "assets" / "models_catalogue.json"
_DEFAULT_OUTPUT = _REPO_ROOT / "models" / "opus-mt"

# Models without a working pre-converted CT2 repo on HuggingFace.
# Change here the models you want to download and convert.
#
# Already converted (Guilherme-Torq): en-pt, en-ko, en-tr, pt-en, ro-en
#
# Broken gaudi repos (missing model.bin):
#   en-es, en-it, en-nl, en-jap, en-ar, en-pl, en-sv, en-da,
#   en-uk, en-cs, en-ro, en-hu, en-fi
#
# Tiny models (no ct2_repo exists at all):
#   fr-en, de-en, it-en, nl-en, ru-en, ar-en, ko-en, en-fr, en-de, en-ru
_MISSING_CT2 = [
    # Broken gaudi en->X repos (no model.bin in their HF repos)
    "en-es", "en-it", "en-nl", "en-jap", "en-ar", "en-sv",
    "en-da", "en-uk", "en-cs", "en-ro", "en-hu", "en-fi",
    # Tiny models (no ct2_repo at all)
    "fr-en", "de-en", "it-en", "nl-en", "ru-en", "ar-en", "ko-en",
    "en-fr", "en-de", "en-ru",
]

_COPY_FILES = ["source.spm", "target.spm", "tokenizer_config.json",
               "vocab.json", "shared_vocabulary.json"]


def _convert_one(slug: str, output_dir: Path) -> bool:
    """Download and convert a single model using TransformersConverter."""
    try:
        import ctranslate2
    except ImportError:
        print("ctranslate2 is required: pip install ctranslate2", file=sys.stderr)
        return False
    try:
        import transformers  # noqa: F401
    except ImportError:
        print("transformers is required: pip install transformers", file=sys.stderr)
        return False

    catalogue = json.loads(_CATALOGUE.read_text(encoding="utf-8"))
    entry = next((e for e in catalogue if e["slug"] == slug), None)
    if not entry:
        print(f"Slug '{slug}' not found in catalogue.", file=sys.stderr)
        return False

    hf_repo = entry["download_url"].rsplit("huggingface.co/", 1)[-1]
    pair = slug  # slug == canonical_name for these models
    out_path = output_dir / pair

    if (out_path / "converted.ok").exists():
        print(f"  [skip] {slug} already converted at {out_path}")
        return True

    print(f"Converting {slug} (repo: {hf_repo}) ...")

    try:
        converter = ctranslate2.converters.TransformersConverter(hf_repo)
        out_path.mkdir(parents=True, exist_ok=True)
        converter.convert(str(out_path), force=True)
    except Exception as e:
        print(f"  TransformersConverter failed for {slug}: {e}", file=sys.stderr)
        return False

    # Copy SentencePiece / tokenizer files from HF cache if not placed
    try:
        from huggingface_hub import hf_hub_download
        for fname in _COPY_FILES:
            dest = out_path / fname
            if not dest.exists():
                try:
                    cached = hf_hub_download(hf_repo, fname)
                    shutil.copy2(cached, str(dest))
                except Exception:
                    pass
    except ImportError:
        pass

    (out_path / "converted.ok").write_text("ok\n")
    print(f"  Done: {out_path}")
    return True


def main():
    parser = argparse.ArgumentParser(description=__doc__,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--slug", help="Convert a single model by slug")
    parser.add_argument("--output-dir", type=Path, default=_DEFAULT_OUTPUT)
    args = parser.parse_args()

    slugs = [args.slug] if args.slug else _MISSING_CT2
    ok = 0
    for slug in slugs:
        if _convert_one(slug, args.output_dir):
            ok += 1
        else:
            print(f"  FAILED: {slug}", file=sys.stderr)

    print(f"\n{ok}/{len(slugs)} models converted successfully.")
    if ok < len(slugs):
        sys.exit(1)


if __name__ == "__main__":
    main()
