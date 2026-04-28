from __future__ import annotations

import json
import os
import re
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable

import fitz


BOOK_KEYWORD = "股票大作手回忆录"
ANNOTATED_KEYWORD = "丁圣元注疏版"
ORIGINAL_KEYWORD = "((美)埃德温·勒菲弗)"


@dataclass
class Entry:
    index: int
    kind: str
    title: str
    short_title: str
    start_page: int
    end_page: int
    page_count: int
    image_pages: list[int]
    extracted_images: list[str]
    quotes: list[str]
    text_path: str


def find_pdf(roots: Iterable[Path], needle: str) -> Path:
    for root in roots:
        if not root.exists():
            continue
        matches = sorted(
            p for p in root.glob("*.pdf") if needle in p.name and BOOK_KEYWORD in p.name
        )
        if matches:
            return matches[0]
    searched = ", ".join(str(root) for root in roots)
    raise FileNotFoundError(f"Could not find PDF containing {needle!r} in: {searched}")


def clean_page_text(text: str) -> str:
    text = text.replace("\u3000", " ")
    text = text.replace("\xa0", " ")
    lines = [line.strip() for line in text.splitlines()]
    kept: list[str] = []
    for line in lines:
        if not line:
            if kept and kept[-1] != "":
                kept.append("")
            continue
        if re.fullmatch(r"\d+", line):
            continue
        kept.append(line)
    return "\n".join(kept).strip()


def merge_hyphen_breaks(text: str) -> str:
    text = re.sub(r"(?<=\w)-\n(?=\w)", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text


def split_sentences(text: str) -> list[str]:
    normalized = re.sub(r"\s+", " ", text)
    pieces = re.split(r"(?<=[。！？；!?;])", normalized)
    return [p.strip() for p in pieces if len(p.strip()) >= 18]


def score_quote(sentence: str) -> int:
    score = 0
    keywords = [
        "市场",
        "投机",
        "趋势",
        "时机",
        "耐心",
        "错误",
        "赚钱",
        "亏损",
        "赔钱",
        "风险",
        "直觉",
        "大势",
        "纸带",
        "纪律",
        "规则",
        "最小阻力",
    ]
    for word in keywords:
        if word in sentence:
            score += 3
    if 20 <= len(sentence) <= 70:
        score += 4
    if "我" in sentence:
        score += 1
    if any(ch in sentence for ch in "“”\""):
        score += 1
    if re.search(r"\d", sentence):
        score -= 1
    return score


def choose_quotes(text: str, limit: int = 6) -> list[str]:
    seen: set[str] = set()
    ranked: list[tuple[int, str]] = []
    for sentence in split_sentences(text):
        if sentence in seen:
            continue
        seen.add(sentence)
        ranked.append((score_quote(sentence), sentence))
    ranked.sort(key=lambda item: (-item[0], item[1]))
    return [sentence for _, sentence in ranked[:limit]]


def shorten_title(title: str) -> str:
    short = re.sub(r"^\d+\s*", "", title).strip()
    short = short.replace("（", "(").replace("）", ")")
    return short


def entry_kind(title: str) -> str:
    if title.startswith("附录"):
        return "appendix"
    if re.match(r"^\d+\s", title):
        return "chapter"
    return "front_matter"


def extract_embedded_images(doc: fitz.Document, page_index: int, out_dir: Path) -> list[str]:
    page = doc.load_page(page_index)
    image_refs = page.get_images(full=True)
    saved: list[str] = []
    for i, image in enumerate(image_refs, start=1):
        xref = image[0]
        info = doc.extract_image(xref)
        ext = info.get("ext", "png")
        out_path = out_dir / f"page_{page_index + 1:03d}_img_{i}.{ext}"
        out_path.write_bytes(info["image"])
        saved.append(str(out_path))
    return saved


def main() -> None:
    workspace = Path(__file__).resolve().parents[2]
    source_pdf_dir = workspace / "source" / "pdfs"
    downloads = Path(os.environ["USERPROFILE"]) / "Downloads"
    pdf_roots = [source_pdf_dir, downloads]

    annotated_pdf = find_pdf(pdf_roots, ANNOTATED_KEYWORD)
    original_pdf = find_pdf(pdf_roots, ORIGINAL_KEYWORD)

    data_dir = workspace / "data"
    text_dir = data_dir / "chapter_texts"
    metadata_dir = data_dir / "metadata"
    image_dir = workspace / "output" / "assets" / "book_extracts"
    data_dir.mkdir(parents=True, exist_ok=True)
    text_dir.mkdir(parents=True, exist_ok=True)
    metadata_dir.mkdir(parents=True, exist_ok=True)
    image_dir.mkdir(parents=True, exist_ok=True)

    doc = fitz.open(annotated_pdf)
    toc = [(level, title.strip(), page) for level, title, page in doc.get_toc() if level == 1]

    entries: list[Entry] = []
    page_map: list[dict[str, object]] = []
    image_page_set = {
        idx + 1
        for idx in range(doc.page_count)
        if len(doc.load_page(idx).get_images(full=True)) > 0
    }

    for idx, (_, title, start_page) in enumerate(toc):
        end_page = toc[idx + 1][2] - 1 if idx + 1 < len(toc) else doc.page_count
        kind = entry_kind(title)
        short = shorten_title(title)
        page_texts: list[str] = []
        embedded_images: list[str] = []

        for page_number in range(start_page, end_page + 1):
            raw = doc.load_page(page_number - 1).get_text("text")
            cleaned = clean_page_text(raw)
            page_texts.append(cleaned)
            page_map.append(
                {
                    "page": page_number,
                    "entry_index": idx + 1,
                    "entry_title": title,
                    "kind": kind,
                    "text": cleaned,
                }
            )
            if page_number in image_page_set:
                embedded_images.extend(
                    extract_embedded_images(doc, page_number - 1, image_dir)
                )

        text = merge_hyphen_breaks("\n\n".join(page_texts))
        text_path = text_dir / f"{idx + 1:02d}_{kind}.txt"
        text_path.write_text(text, encoding="utf-8")

        entry = Entry(
            index=idx + 1,
            kind=kind,
            title=title,
            short_title=short,
            start_page=start_page,
            end_page=end_page,
            page_count=end_page - start_page + 1,
            image_pages=[page for page in range(start_page, end_page + 1) if page in image_page_set],
            extracted_images=embedded_images,
            quotes=choose_quotes(text),
            text_path=str(text_path),
        )
        entries.append(entry)

    payload = {
        "source": {
            "annotated_pdf": str(annotated_pdf),
            "original_pdf": str(original_pdf),
            "annotated_page_count": doc.page_count,
        },
        "entries": [asdict(entry) for entry in entries],
    }

    (metadata_dir / "book_structure.json").write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    (metadata_dir / "page_map.json").write_text(
        json.dumps(page_map, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    print(f"Wrote {metadata_dir / 'book_structure.json'}")
    print(f"Wrote {metadata_dir / 'page_map.json'}")
    print(f"Wrote chapter texts to {text_dir}")
    print(f"Extracted images to {image_dir}")


if __name__ == "__main__":
    main()
