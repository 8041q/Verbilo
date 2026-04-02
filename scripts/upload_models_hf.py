#!/usr/bin/env python3
# Upload all 22 locally-converted opus-mt CTranslate2 models to HuggingFace.
#
# Usage:
#   set HF_TOKEN=hf_xxxxxxxxxxxxxxxxxxxx
#   python scripts/upload_models_hf.py
#
# Required dependency (not in requirements.txt):
#   pip install huggingface_hub

import os
import sys
from pathlib import Path

_REPO_ROOT = Path(__file__).resolve().parents[1]
_OPUS_DIR = _REPO_ROOT / "models" / "opus-mt"
_HF_NAMESPACE = "Guilherme-Torq"
_REPO_PREFIX = "c2t"


def main() -> int:
    token = os.environ.get("HF_TOKEN", "").strip()
    if not token:
        print(
            "Error: HF_TOKEN environment variable is not set.\n"
            "  Set it before running:  set HF_TOKEN=hf_...",
            file=sys.stderr,
        )
        return 1

    try:
        from huggingface_hub import HfApi
    except ImportError:
        print(
            "Error: huggingface_hub is not installed.\n"
            "  Run:  pip install huggingface_hub",
            file=sys.stderr,
        )
        return 1

    model_dirs = sorted(
        p for p in _OPUS_DIR.iterdir() if p.is_dir()
    )

    if not model_dirs:
        print(f"Error: No model directories found under {_OPUS_DIR}", file=sys.stderr)
        return 1

    total = len(model_dirs)
    print(f"Found {total} model directories under {_OPUS_DIR}")
    print()

    api = HfApi(token=token)
    failures: list[tuple[str, str]] = []

    for idx, model_dir in enumerate(model_dirs, start=1):
        lang_pair = model_dir.name
        repo_id = f"{_HF_NAMESPACE}/{_REPO_PREFIX}-{lang_pair}"
        print(f"[{idx}/{total}] {repo_id}")

        try:
            api.create_repo(
                repo_id=repo_id,
                repo_type="model",
                private=False,
                exist_ok=True,
            )
        except Exception as exc:
            msg = f"create_repo failed: {exc}"
            print(f"  ERROR: {msg}")
            failures.append((repo_id, msg))
            continue

        try:
            api.upload_folder(
                folder_path=str(model_dir),
                repo_id=repo_id,
                repo_type="model",
                commit_message=f"Upload CTranslate2 model for {lang_pair}",
            )
            print(f"  Done -> https://huggingface.co/{repo_id}")
        except Exception as exc:
            msg = f"upload_folder failed: {exc}"
            print(f"  ERROR: {msg}")
            failures.append((repo_id, msg))

    print()
    print(f"Finished: {total - len(failures)}/{total} succeeded.")

    if failures:
        print(f"\n{len(failures)} failure(s):")
        for repo_id, msg in failures:
            print(f"  {repo_id}: {msg}")
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
