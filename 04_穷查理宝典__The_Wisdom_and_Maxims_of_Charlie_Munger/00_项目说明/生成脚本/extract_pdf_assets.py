from __future__ import annotations

import argparse
import json
from pathlib import Path

import fitz


ROOT = Path(__file__).resolve().parents[2]
INTERMEDIATE_DIR = ROOT / "05_中间素材"
IMAGE_DIR = ROOT / "04_配图" / "原书页图"

CHAPTERS = [
    {
        "id": "ch01",
        "number": "01",
        "title": "查理·芒格传略",
        "full_title": "第一章 查理·芒格传略",
        "start_page": 52,
        "end_page": 146,
        "opener_page": 52,
        "reference_pages": [52, 53, 97, 122],
    },
    {
        "id": "ch02",
        "number": "02",
        "title": "芒格的生活、学习和决策方法",
        "full_title": "第二章 芒格的生活、学习和决策方法",
        "start_page": 147,
        "end_page": 205,
        "opener_page": 147,
        "reference_pages": [147, 150, 162, 184],
    },
    {
        "id": "ch03",
        "number": "03",
        "title": "芒格主义：查理的即席谈话",
        "full_title": "第三章 芒格主义：查理的即席谈话",
        "start_page": 206,
        "end_page": 317,
        "opener_page": 206,
        "reference_pages": [206, 225, 253, 291],
    },
    {
        "id": "ch04",
        "number": "04",
        "title": "查理十一讲",
        "full_title": "第四章 查理十一讲",
        "start_page": 318,
        "end_page": 886,
        "opener_page": 318,
        "reference_pages": [318, 322, 349, 437, 538, 574, 658, 808],
    },
    {
        "id": "ch05",
        "number": "05",
        "title": "文章、报道与评论",
        "full_title": "第五章 文章、报道与评论",
        "start_page": 887,
        "end_page": 953,
        "opener_page": 887,
        "reference_pages": [887, 893, 901, 939, 945],
    },
]

STATIC_PAGE_RENDERS = [
    {"name": "book_cover_page", "page": 2},
    {"name": "book_toc_page", "page": 3},
    {"name": "book_year_chronology_page", "page": 972},
]


def find_pdf(provided_path: str | None) -> Path:
    if provided_path:
        pdf_path = Path(provided_path).expanduser().resolve()
        if not pdf_path.exists():
            raise FileNotFoundError(f"未找到 PDF: {pdf_path}")
        return pdf_path

    candidates = []
    search_roots = [
        Path.home() / "Downloads" / "Documents",
        Path.home() / "Downloads",
        ROOT,
    ]
    for base in search_roots:
        if not base.exists():
            continue
        candidates.extend(base.rglob("*穷查理宝典*.pdf"))

    if not candidates:
        raise FileNotFoundError("未在常见目录中找到《穷查理宝典》PDF，请用 --pdf 指定路径。")

    return sorted(candidates, key=lambda item: str(item))[0]


def render_page(doc: fitz.Document, page_no: int, out_path: Path, scale: float = 1.6) -> None:
    page = doc[page_no - 1]
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    pix.save(out_path)


def extract_text_for_range(doc: fitz.Document, start_page: int, end_page: int) -> str:
    chunks: list[str] = []
    for page_no in range(start_page, end_page + 1):
        text = doc[page_no - 1].get_text("text").replace("\x00", "").strip()
        chunks.append(f"===== PAGE {page_no} =====\n{text}\n")
    return "\n".join(chunks)


def main() -> None:
    parser = argparse.ArgumentParser(description="抽取《穷查理宝典》的章节文本与页图。")
    parser.add_argument("--pdf", type=str, default=None, help="PDF 路径。")
    args = parser.parse_args()

    pdf_path = find_pdf(args.pdf)
    doc = fitz.open(pdf_path)

    INTERMEDIATE_DIR.mkdir(parents=True, exist_ok=True)
    IMAGE_DIR.mkdir(parents=True, exist_ok=True)

    manifest = {
        "book_title": "穷查理宝典：查理·芒格智慧箴言录",
        "pdf_path": str(pdf_path),
        "page_count": doc.page_count,
        "chapters": CHAPTERS,
        "static_page_renders": STATIC_PAGE_RENDERS,
    }

    (INTERMEDIATE_DIR / "chapter_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    for chapter in CHAPTERS:
        chapter_text = extract_text_for_range(doc, chapter["start_page"], chapter["end_page"])
        text_path = INTERMEDIATE_DIR / f"{chapter['id']}_text.txt"
        text_path.write_text(chapter_text, encoding="utf-8")

        opener_out = IMAGE_DIR / f"{chapter['id']}_opener_page.png"
        render_page(doc, chapter["opener_page"], opener_out)
        for ref_page in chapter["reference_pages"]:
            render_page(doc, ref_page, IMAGE_DIR / f"{chapter['id']}_page_{ref_page}.png", scale=1.35)

    for page_info in STATIC_PAGE_RENDERS:
        render_page(doc, page_info["page"], IMAGE_DIR / f"{page_info['name']}.png", scale=1.5)

    print(f"PDF: {pdf_path}")
    print(f"页数: {doc.page_count}")
    print("已输出:")
    print(f"- {INTERMEDIATE_DIR / 'chapter_manifest.json'}")
    for chapter in CHAPTERS:
        print(f"- {INTERMEDIATE_DIR / f'{chapter['id']}_text.txt'}")
    print(f"- {IMAGE_DIR}")


if __name__ == "__main__":
    main()
