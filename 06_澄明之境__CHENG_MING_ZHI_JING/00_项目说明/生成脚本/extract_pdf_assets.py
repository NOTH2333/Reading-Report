from __future__ import annotations

import argparse
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPO_ROOT = PROJECT_ROOT.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))

from book_extract_common import extract_book_assets


def main() -> None:
    parser = argparse.ArgumentParser(description="抽取《澄明之境》的章节文本与页图。")
    parser.add_argument("--pdf", type=str, default=None, help="PDF 路径。")
    args = parser.parse_args()
    manifest = extract_book_assets("chengming_zhijing", PROJECT_ROOT, provided_pdf=args.pdf)
    print(json_path(PROJECT_ROOT))
    print(manifest["source_pdf"])


def json_path(root: Path) -> str:
    return str(root / "05_中间素材" / "chapter_manifest.json")


if __name__ == "__main__":
    main()
