#!/usr/bin/env python3
# Download helper for FastText and OPUS-MT models used by the app.
#
# Needed dependencies (not needed for the GUI, only this file):
#   ctranslate2, transformers, huggingface_hub, pyyaml

import argparse
import json
import os
import shutil
import ssl
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Optional
from urllib.request import urlopen, Request
from urllib.error import HTTPError


FASTTEXT_URL = "https://dl.fbaipublicfiles.com/fasttext/supervised-models/lid.176.bin"

_REPO_ROOT = Path(__file__).resolve().parents[1]
_DEFAULT_DEST = str(_REPO_ROOT / "models" / "lid.176.bin")
_HF_BASE = "https://huggingface.co/{repo}/resolve/main"
_OPUS_RAW_URLS = [
    "https://object.pouta.csc.fi/OPUS-MT-models/{slug}/opus+bt-2021-04-14.zip",
    "https://object.pouta.csc.fi/OPUS-MT-models/{slug}/opus-2020-02-26.zip",
]
_DEFAULT_OPUS_DIR = str(_REPO_ROOT / "models" / "opus-mt")
_SENTINEL = "converted.ok"
_UNDERSCORE_PREFIXES = ("tiny_",)
_COPY_FILES = ["source.spm", "target.spm", "tokenizer_config.json"]

_SSL_UNVERIFIED = ssl.create_default_context()
_SSL_UNVERIFIED.check_hostname = False
_SSL_UNVERIFIED.verify_mode = ssl.CERT_NONE


# ---------------------------------------------------------------------------
# HTTP helpers
# ---------------------------------------------------------------------------

def _open_url(url: str, method: str = "GET"):
    req = Request(url, headers={"User-Agent": "Mozilla/5.0"})
    req.method = method
    try:
        return urlopen(req)
    except Exception as e:
        is_ssl = isinstance(e, ssl.SSLCertVerificationError) or (
            isinstance(e, OSError)
            and any(k in str(e).upper() for k in ("SSL", "CERTIFICATE", "CERT_VERIFY"))
        )
        if is_ssl:
            return urlopen(req, context=_SSL_UNVERIFIED)
        raise


def _hf_head(repo: str) -> bool:
    try:
        with _open_url(f"https://huggingface.co/{repo}", method="HEAD") as r:
            return r.status == 200
    except HTTPError as e:
        return e.code != 404
    except Exception:
        return False


def _hf_file_exists(repo: str, filename: str) -> bool:
    try:
        with _open_url(f"{_HF_BASE.format(repo=repo)}/{filename}", method="HEAD") as r:
            return r.status == 200
    except Exception:
        return False


def download(url: str, dest_path: str) -> None:
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    print(f"Downloading {url} -> {dest_path}")
    with _open_url(url) as r:
        total = int(r.headers.get("Content-Length", 0))
        received = 0
        with open(dest_path, "wb") as f:
            while True:
                chunk = r.read(256 * 1024)  # 256 KB chunks
                if not chunk:
                    break
                f.write(chunk)
                received += len(chunk)
                print(f"PROGRESS {received} {total}", flush=True)


def _try_download(url: str, dest_path: str) -> bool:
    try:
        download(url, dest_path)
        return True
    except HTTPError as e:
        print(f"  HTTP {e.code} for {url}", file=sys.stderr)
        return False
    except Exception as e:
        print(f"  Download error: {e}", file=sys.stderr)
        return False


# ---------------------------------------------------------------------------
# Repo resolution
# ---------------------------------------------------------------------------

def _resolve_hf_repo(slug: str) -> str:
    org = "Helsinki-NLP"
    if any(slug.startswith(p) for p in _UNDERSCORE_PREFIXES):
        repo = f"{org}/opus-mt_{slug}"
        if _hf_head(repo):
            return repo
    repo = f"{org}/opus-mt-{slug}"
    if _hf_head(repo):
        return repo
    print(f"Warning: could not verify repo for '{slug}'. Trying '{repo}'.", file=sys.stderr)
    return repo


# ---------------------------------------------------------------------------
# Download all files from HF repo using hf_hub_download (no snapshot_download)
# ---------------------------------------------------------------------------

def _list_hf_repo_files(model_name: str) -> list:
    """Return list of filenames in a HuggingFace repo via the JSON API."""
    url = f"https://huggingface.co/api/models/{model_name}"
    try:
        with _open_url(url) as r:
            data = json.loads(r.read().decode())
        return [s["rfilename"] for s in data.get("siblings", [])]
    except Exception as e:
        print(f"  Could not list files for {model_name}: {e}", file=sys.stderr)
        return []


def _download_repo_files(model_name: str, local_dir: str) -> bool:
    """Download every file in a HuggingFace repo via plain HTTP.

    huggingface_hub (snapshot_download AND hf_hub_download) both have a
    WindowsPath .touch() bug that leaves all files as .incomplete.
    We bypass the library entirely: list files via the HF JSON API, then
    download each one with our own urllib-based download() function.
    """
    os.makedirs(local_dir, exist_ok=True)
    files = _list_hf_repo_files(model_name)
    if not files:
        return False

    print(f"Downloading {len(files)} files from {model_name} -> {local_dir}")
    base_url = _HF_BASE.format(repo=model_name)
    for filename in files:
        dest = Path(local_dir) / filename
        if dest.exists():
            print(f"  [skip] {filename}")
            continue
        dest.parent.mkdir(parents=True, exist_ok=True)
        if not _try_download(f"{base_url}/{filename}", str(dest)):
            print(f"  [warn] failed to download {filename}", file=sys.stderr)

    return True


# ---------------------------------------------------------------------------
# Conversion: Bergamot/tiny models (config.intgemm8bitalpha.yml)
# ---------------------------------------------------------------------------

def _convert_bergamot(local_dir: str, out_path: Path) -> bool:
    """Convert a Bergamot/tiny OPUS-MT model using MarianConverter.

    Tiny models ship:
      - config.intgemm8bitalpha.yml  lists model (.npz) and vocab (.spm) paths
      - model.npz                    Marian weights
      - *.spm                        SentencePiece vocab

    We read the yml to find the exact model_path and vocab_paths, then call
    ctranslate2.converters.MarianConverter directly.
    """
    try:
        import ctranslate2
        import yaml
    except ImportError as e:
        print(f"  Missing dependency: {e}", file=sys.stderr)
        return False

    local_path = Path(local_dir)

    # List what we actually have for diagnostics.
    all_files = list(local_path.rglob("*"))
    print(f"  Files in cache: {[f.name for f in all_files if f.is_file()]}")

    cfg_path = local_path / "config.intgemm8bitalpha.yml"
    if not cfg_path.exists():
        # Fall back: look for any decoder config
        cfg_path = next(local_path.rglob("decoder.yml"), None)
        if cfg_path is None:
            print(f"  No config.intgemm8bitalpha.yml or decoder.yml found in {local_dir}",
                  file=sys.stderr)
            return False

    print(f"  Reading config: {cfg_path}")
    with open(cfg_path, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)

    models_list = cfg.get("models") or cfg.get("model")
    vocabs_list = cfg.get("vocabs") or cfg.get("vocab")

    if not models_list or not vocabs_list:
        print(f"  'models'/'vocabs' keys not found in {cfg_path.name}. "
              f"Keys present: {list(cfg.keys())}", file=sys.stderr)
        return False

    if isinstance(models_list, str):
        models_list = [models_list]
    if isinstance(vocabs_list, str):
        vocabs_list = [vocabs_list]

    model_path = str(local_path / models_list[0])
    vocab_paths = [str(local_path / v) for v in vocabs_list]

    if not Path(model_path).exists():
        print(f"  Model file not found: {model_path}", file=sys.stderr)
        return False

    print(f"  model:  {model_path}")
    print(f"  vocabs: {vocab_paths}")

    try:
        out_path.mkdir(parents=True, exist_ok=True)
        converter = ctranslate2.converters.MarianConverter(model_path, vocab_paths)
        converter.convert(str(out_path), force=True)
        print("MarianConverter succeeded.")
        return True
    except Exception as e:
        print(f"  MarianConverter failed: {e}", file=sys.stderr)
        return False


# ---------------------------------------------------------------------------
# Conversion: standard MarianMT (model_type=marian in config.json)
# ---------------------------------------------------------------------------

def _convert_standard_marian(local_dir: str, out_path: Path) -> bool:
    """Convert a standard MarianMT model using ct2-transformers-converter."""
    # In a frozen/Nuitka build the converter CLI is not available (it requires
    # a full Python + transformers runtime), so skip the attempt entirely.
    if getattr(sys, "frozen", False) or globals().get("__compiled__"):
        print("  Skipping ct2-transformers-converter (frozen build).", file=sys.stderr)
        return False

    scripts_dir = Path(sys.executable).parent
    converter_bin = scripts_dir / "ct2-transformers-converter"
    if not converter_bin.exists():
        converter_bin = scripts_dir / "ct2-transformers-converter.exe"
    if not converter_bin.exists():
        print(f"  ct2-transformers-converter not found in {scripts_dir}.", file=sys.stderr)
        return False

    out_path.mkdir(parents=True, exist_ok=True)
    cmd = [str(converter_bin), "--model", local_dir,
           "--output_dir", str(out_path), "--force", "--copy_files", *_COPY_FILES]
    print(f"Running: {' '.join(cmd)}")
    result = subprocess.run(cmd)
    if result.returncode == 0:
        print("ct2-transformers-converter succeeded.")
        return True
    print("ct2-transformers-converter failed.", file=sys.stderr)
    return False


# ---------------------------------------------------------------------------
# Direct download of pre-converted CTranslate2 model from HuggingFace
# ---------------------------------------------------------------------------

_CT2_REQUIRED_FILES = ["model.bin", "source.spm", "target.spm"]
_CT2_OPTIONAL_FILES = [
    "shared_vocabulary.json", "config.json", "vocab.json",
    "tokenizer_config.json",
]


def _download_ct2_direct(ct2_repo: str, out_path: Path) -> bool:
    """Download a pre-converted CTranslate2 model directly from HuggingFace.

    These repos (e.g. gaudi/opus-mt-en-fr-ctranslate2) already contain the
    final model.bin + SentencePiece files — no conversion step needed.
    """
    files = _list_hf_repo_files(ct2_repo)
    if not files:
        print(f"  Could not list files for CT2 repo: {ct2_repo}", file=sys.stderr)
        return False

    # Verify required files exist in the repo
    for req in _CT2_REQUIRED_FILES:
        if req not in files:
            print(f"  CT2 repo {ct2_repo} missing required file: {req}", file=sys.stderr)
            return False

    out_path.mkdir(parents=True, exist_ok=True)
    base_url = _HF_BASE.format(repo=ct2_repo)
    wanted = [f for f in files if f in _CT2_REQUIRED_FILES or f in _CT2_OPTIONAL_FILES]

    print(f"Downloading pre-converted CT2 model from {ct2_repo} -> {out_path}")
    for filename in wanted:
        dest = out_path / filename
        if dest.exists():
            print(f"  [skip] {filename}")
            continue
        if not _try_download(f"{base_url}/{filename}", str(dest)):
            print(f"  Failed to download {filename} from {ct2_repo}", file=sys.stderr)
            return False

    print("CT2 direct download succeeded.")
    return True


# ---------------------------------------------------------------------------
# Conversion: raw zip from object.pouta.csc.fi (decoder.yml format)
# ---------------------------------------------------------------------------

def _convert_raw_zip(slug: str, out_path: Path) -> bool:
    """Download the raw Marian zip and convert with OpusMTConverter."""
    try:
        import ctranslate2
    except ImportError:
        return False

    with tempfile.TemporaryDirectory() as tmp:
        zip_path = os.path.join(tmp, "model.zip")
        downloaded = False
        for url_tpl in _OPUS_RAW_URLS:
            url = url_tpl.format(slug=slug)
            print(f"Trying raw zip: {url}")
            if _try_download(url, zip_path):
                downloaded = True
                break
        if not downloaded:
            return False

        with zipfile.ZipFile(zip_path) as zf:
            zf.extractall(tmp)

        decoder_yml = next(Path(tmp).rglob("decoder.yml"), None)
        if not decoder_yml:
            print("  decoder.yml not found in zip.", file=sys.stderr)
            return False

        try:
            out_path.mkdir(parents=True, exist_ok=True)
            converter = ctranslate2.converters.OpusMTConverter(str(decoder_yml.parent))
            converter.convert(str(out_path), force=True)
            print("OpusMTConverter succeeded.")
            return True
        except Exception as e:
            print(f"  OpusMTConverter failed: {e}", file=sys.stderr)
            return False


# ---------------------------------------------------------------------------
# Catalogue look-up for CT2 repo
# ---------------------------------------------------------------------------

def _lookup_ct2_repo(slug: str) -> Optional[str]:
    """Return the ct2_repo value from models_catalogue.json for this slug."""
    cat_paths = [
        Path(__file__).resolve().parent.parent / "src" / "verbilo" / "assets" / "models_catalogue.json",
        # Frozen build: catalogue is next to the exe
        Path(sys.executable).resolve().parent / "verbilo" / "assets" / "models_catalogue.json",
    ]
    for cat_path in cat_paths:
        if cat_path.is_file():
            try:
                catalogue = json.loads(cat_path.read_text(encoding="utf-8"))
                for entry in catalogue:
                    if entry.get("slug") == slug:
                        return entry.get("ct2_repo")
            except Exception:
                pass
    return None


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def _slug_to_pair(slug: str) -> str:
    """Derive the canonical src-tgt folder name from a slug.

    local.py expects the model directory to be named exactly "{src}-{tgt}",
    e.g. "en-fr", so it can build the path as model_dir / f"{src}-{tgt}".

    The slug may be longer (e.g. "tiny_eng-fra", "tc-big-en-fr") — we extract
    the last two dash-separated tokens that look like language codes, but the
    simplest and most reliable approach is to derive it from the HF repo
    language codes embedded in the slug:

      tiny_eng-fra  → eng-fra
      tc-big-en-fr  → en-fr
      en-fr         → en-fr  (already correct)

    We take everything after the last underscore (or the whole slug if no
    underscore), which gives us the "src-tgt" portion.
    """
    # Strip any leading variant prefix (tiny_, tc-big-, etc.) — keep only
    # the part after the final underscore.
    if "_" in slug:
        return slug.split("_", 1)[-1]   # "tiny_eng-fra" → "eng-fra"
    return slug                          # "en-fr" → "en-fr"


def download_opus_mt(slug: str, dest_dir: Optional[str] = None,
                     ct2_repo: Optional[str] = None,
                     hf_repo: Optional[str] = None) -> Path:
    dest_dir = dest_dir or _DEFAULT_OPUS_DIR
    # Output folder must be named "{src}-{tgt}" so local.py can find it.
    pair = _slug_to_pair(slug)
    out_path = Path(dest_dir) / pair

    if (out_path / _SENTINEL).exists():
        print(f"Model '{slug}' (folder: {pair}) already converted at {out_path}")
        return out_path

    # If no ct2_repo was explicitly passed, look it up from the catalogue.
    if ct2_repo is None:
        ct2_repo = _lookup_ct2_repo(slug)

    out_path.parent.mkdir(parents=True, exist_ok=True)

    # ── Strategy 1: pre-converted CTranslate2 model (no conversion needed) ──
    if ct2_repo:
        print(f"PHASE download", flush=True)
        print(f"Strategy: CT2 direct download from {ct2_repo}")
        if _download_ct2_direct(ct2_repo, out_path):
            (out_path / _SENTINEL).write_text("ok\n")
            print(f"\nModel '{slug}' ready at {out_path} (load as pair '{pair}')")
            return out_path
        print("CT2 direct download failed, falling back to conversion.", file=sys.stderr)

    # ── Remaining strategies require downloading the original HF repo ────────
    for pkg in ["ctranslate2"]:
        try:
            __import__(pkg)
        except ImportError:
            print(f"{pkg} is not installed. Run: pip install {pkg}", file=sys.stderr)
            sys.exit(1)

    print(f"PHASE download", flush=True)
    # Use explicit hf_repo from catalogue download_url when available;
    # fall back to slug-based resolution otherwise.
    if hf_repo:
        model_name = hf_repo
    else:
        model_name = _resolve_hf_repo(slug)
    print(f"Resolved HuggingFace repo: {model_name}")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    safe_name = model_name.replace("/", "--")
    local_dir = os.path.join(dest_dir, ".hf-cache", safe_name)

    # Download files first (always needed regardless of model type).
    if not _download_repo_files(model_name, local_dir):
        print("Failed to download model files.", file=sys.stderr)
        sys.exit(1)

    # Detect model type from what was actually downloaded.
    has_bergamot_cfg = (Path(local_dir) / "config.intgemm8bitalpha.yml").exists()
    has_decoder_yml  = (Path(local_dir) / "decoder.yml").exists()
    has_safetensors  = any(Path(local_dir).glob("*.safetensors"))
    has_pytorch_bin  = any(Path(local_dir).glob("pytorch_model*.bin"))

    print(f"  Model type detection:")
    print(f"    config.intgemm8bitalpha.yml : {has_bergamot_cfg}")
    print(f"    decoder.yml                : {has_decoder_yml}")
    print(f"    *.safetensors              : {has_safetensors}")
    print(f"    pytorch_model*.bin         : {has_pytorch_bin}")

    converted = False
    print("PHASE converting", flush=True)

    if has_bergamot_cfg or has_decoder_yml:
        # Bergamot/raw-Marian format — use MarianConverter via yml config
        print("Strategy: MarianConverter (Bergamot/decoder.yml)")
        converted = _convert_bergamot(local_dir, out_path)

    elif has_safetensors or has_pytorch_bin:
        # Standard HuggingFace Transformers format — use ct2-transformers-converter
        print("Strategy: ct2-transformers-converter (Transformers/safetensors)")
        converted = _convert_standard_marian(local_dir, out_path)
        if not converted:
            # CLI is unavailable in a frozen/standalone build (sys.executable is verbilo.exe);
            # fall back to the raw OPUS zip which only requires ctranslate2 (compiled in).
            print("Strategy: raw zip fallback (ct2-transformers-converter unavailable)")
            converted = _convert_raw_zip(slug, out_path)

    else:
        # Nothing useful downloaded — fall back to raw zip
        print("Strategy: raw zip fallback (OpusMTConverter)")
        converted = _convert_raw_zip(slug, out_path)

    if not converted:
        is_frozen = getattr(sys, "frozen", False) or globals().get("__compiled__")
        if is_frozen:
            print(
                f"\nERROR: Model '{slug}' cannot be downloaded in standalone mode.\n"
                "No pre-converted CTranslate2 version is available for this model.\n"
                "Run the conversion in a development environment first, or add a\n"
                "'ct2_repo' entry to models_catalogue.json pointing to a pre-converted repo.",
                file=sys.stderr,
            )
        else:
            print(
                f"\nConversion failed for '{slug}' (pair '{pair}').\n"
                "Check:\n"
                "  1. Slug at https://huggingface.co/Helsinki-NLP\n"
                "  2. pip install -U ctranslate2 transformers huggingface_hub",
                file=sys.stderr,
            )
        sys.exit(1)

    # Copy spm / tokenizer files if not already placed by the converter.
    for sp_file in _COPY_FILES:
        dest_sp = out_path / sp_file
        if not dest_sp.exists():
            local_copy = Path(local_dir) / sp_file
            if local_copy.exists():
                shutil.copy2(str(local_copy), str(dest_sp))
            else:
                try:
                    download(f"{_HF_BASE.format(repo=model_name)}/{sp_file}", str(dest_sp))
                except Exception as exc:
                    print(f"Warning: could not copy {sp_file}: {exc}", file=sys.stderr)

    (out_path / _SENTINEL).write_text("ok\n")
    print(f"\nModel '{slug}' ready at {out_path} (load as pair '{pair}')")
    return out_path


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Download FastText and/or OPUS-MT models.")
    sub = parser.add_subparsers(dest="command")

    ft = sub.add_parser("fasttext", help="Download the FastText language-detection model")
    ft.add_argument("--dest", default=_DEFAULT_DEST)

    opus = sub.add_parser("opus-mt", help="Download and convert an OPUS-MT translation model")
    opus.add_argument("slug", help=(
        "Model slug, e.g. 'en-fr', 'tiny_eng-fra', 'tc-big-en-fr'. "
        "Repo and conversion strategy resolved automatically. "
        "Output folder is always named by the language pair (e.g. eng-fra), "
        "matching what local.py expects under model_dir/src-tgt/."
    ))
    opus.add_argument("--dest-dir", default=_DEFAULT_OPUS_DIR)
    opus.add_argument("--ct2-repo", default=None,
                      help="HuggingFace repo with pre-converted CT2 model")
    opus.add_argument("--hf-repo", default=None,
                      help="HuggingFace repo name for the original model")

    args = parser.parse_args()

    if args.command == "opus-mt":
        try:
            download_opus_mt(args.slug, args.dest_dir, ct2_repo=args.ct2_repo,
                             hf_repo=args.hf_repo)
        except Exception as e:
            print(f"OPUS-MT download failed: {e}", file=sys.stderr)
            sys.exit(1)
    else:
        dest = getattr(args, "dest", _DEFAULT_DEST)
        try:
            download(FASTTEXT_URL, dest)
        except Exception as e:
            print("Download failed:", e, file=sys.stderr)
            sys.exit(1)
        print("Download complete")


if __name__ == "__main__":
    main()
