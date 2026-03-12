#!/usr/bin/env python3
# Download helper for FastText and other models used by the app.
import argparse
import os
import sys
from pathlib import Path
from urllib.request import urlopen


FASTTEXT_URL = "https://dl.fbaipublicfiles.com/fasttext/supervised-models/lid.176.bin"

# Default destination is anchored to the repo root, regardless of where the script is run from.
_DEFAULT_DEST = str(Path(__file__).resolve().parents[1] / "models" / "lid.176.bin")


def download(url, dest_path):
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    print(f"Downloading {url} -> {dest_path}")
    with urlopen(url) as r, open(dest_path, "wb") as f:
        f.write(r.read())


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--dest", default=_DEFAULT_DEST)
    args = parser.parse_args()

    try:
        download(FASTTEXT_URL, args.dest)
    except Exception as e:
        print("Download failed:", e, file=sys.stderr)
        sys.exit(1)

    print("Download complete")


if __name__ == "__main__":
    main()
