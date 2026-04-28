from __future__ import annotations

import json
from pathlib import Path

import fitz

from dual_book_specs import BOOK_PROJECTS


def find_pdf(search_terms: list[str], provided_path: str | None = None) -> Path:
    if provided_path:
        pdf_path = Path(provided_path).expanduser().resolve()
        if not pdf_path.exists():
            raise FileNotFoundError(f"未找到 PDF: {pdf_path}")
        return pdf_path

    search_roots = [
        Path.home() / "Downloads" / "Documents",
        Path.home() / "Downloads",
    ]
    candidates: list[Path] = []
    for base in search_roots:
        if not base.exists():
            continue
        for pdf in base.rglob("*.pdf"):
            name = pdf.name
            if any(term in name for term in search_terms):
                candidates.append(pdf)
    if not candidates:
        joined = " / ".join(search_terms)
        raise FileNotFoundError(f"未在常见目录中找到匹配 PDF: {joined}")
    return sorted(candidates, key=lambda item: str(item))[0]


def render_page(doc: fitz.Document, page_no: int, out_path: Path, scale: float = 1.55) -> None:
    page = doc[page_no - 1]
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    pix.save(out_path)


def extract_text_range(doc: fitz.Document, start_page: int, end_page: int) -> str:
    chunks: list[str] = []
    for page_no in range(start_page, end_page + 1):
        text = doc[page_no - 1].get_text("text").replace("\x00", "").strip()
        chunks.append(f"===== PAGE {page_no} =====\n{text}\n")
    return "\n".join(chunks)


def safe_name(raw: str) -> str:
    cleaned = raw.replace("：", "_").replace(" ", "")
    for token in ["（", "）", "(", ")", "——", "—", "？", "！", "、", "，", "。", "《", "》", "“", "”", "：", "；", "·", "…", "/", "\\"]:
        cleaned = cleaned.replace(token, "_")
    while "__" in cleaned:
        cleaned = cleaned.replace("__", "_")
    return cleaned.strip("_")


def write_section_text(text_dir: Path, prefix: str, title: str, text: str) -> Path:
    path = text_dir / f"{prefix}_{safe_name(title)}.txt"
    path.write_text(text, encoding="utf-8")
    return path


def extract_book_assets(book_key: str, project_root: Path, provided_pdf: str | None = None) -> dict:
    spec = BOOK_PROJECTS[book_key]
    pdf_path = find_pdf(spec["search_terms"], provided_path=provided_pdf)
    doc = fitz.open(pdf_path)

    text_dir = project_root / "05_中间素材" / "章节原文"
    image_dir = project_root / "04_配图" / "原书页图"
    text_dir.mkdir(parents=True, exist_ok=True)
    image_dir.mkdir(parents=True, exist_ok=True)

    manifest = {
        "book_id": spec["book_id"],
        "book_title": spec["title"],
        "source_pdf": str(pdf_path),
        "page_count": doc.page_count,
        "intro_sections": spec["intro_sections"],
        "chapters": spec["chapters"],
        "wrap_up_sections": spec["wrap_up_sections"],
        "static_page_renders": spec["static_page_renders"],
    }
    (project_root / "05_中间素材" / "chapter_manifest.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    for index, section in enumerate(spec["intro_sections"], start=1):
        text = extract_text_range(doc, section["start_page"], section["end_page"])
        write_section_text(text_dir, f"00_导读_{index:02d}", section["title"], text)
        render_page(doc, section["start_page"], image_dir / f"intro_{index:02d}_page_{section['start_page']:03d}.png", scale=1.35)

    for chapter in spec["chapters"]:
        text = extract_text_range(doc, chapter["start_page"], chapter["end_page"])
        write_section_text(text_dir, f"第{chapter['sequence_no']}章", chapter["title"], text)
        render_page(doc, chapter["opener_page"], image_dir / f"{chapter['id']}_opener_page.png")
        for ref_page in chapter["reference_pages"]:
            render_page(doc, ref_page, image_dir / f"{chapter['id']}_page_{ref_page}.png", scale=1.3)

    for index, section in enumerate(spec["wrap_up_sections"], start=1):
        text = extract_text_range(doc, section["start_page"], section["end_page"])
        write_section_text(text_dir, f"99_收束_{index:02d}", section["title"], text)
        render_page(doc, section["start_page"], image_dir / f"wrap_{index:02d}_page_{section['start_page']:03d}.png", scale=1.35)

    for page_info in spec["static_page_renders"]:
        render_page(doc, page_info["page"], image_dir / f"{page_info['name']}.png", scale=1.5)

    return manifest
