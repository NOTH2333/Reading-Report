from __future__ import annotations

import argparse
import sys
from pathlib import Path

import fitz
from pypdf import PdfReader

from book_pdf_manifest import BOOK_JOBS
from markdown_pdf_renderer import MarkdownPdfRenderer


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="批量导出 5 本书的阅读报告与关键知识总结 PDF")
    parser.add_argument("--book-id", help="只处理指定 book_id")
    parser.add_argument("--validate-only", action="store_true", help="只校验源文件、图片与现有 PDF")
    parser.add_argument("--force", action="store_true", help="即使目标 PDF 已存在也重新生成")
    return parser.parse_args()


def select_jobs(book_id: str | None):
    if not book_id:
        return BOOK_JOBS
    selected = [job for job in BOOK_JOBS if job.book_id == book_id]
    if not selected:
        raise SystemExit(f"未找到 book_id={book_id!r}")
    return tuple(selected)


def validate_pdf(path: Path) -> tuple[int, str]:
    reader = PdfReader(str(path))
    page_count = len(reader.pages)
    if page_count <= 0:
        raise ValueError(f"PDF 没有页面: {path}")
    sample_texts: list[str] = []
    for page in reader.pages[: min(3, page_count)]:
        text = (page.extract_text() or "").strip()
        if text:
            sample_texts.append(text[:180].replace("\n", " "))
    fitz_doc = fitz.open(path)
    fitz_doc.load_page(0).get_pixmap(matrix=fitz.Matrix(1.2, 1.2), alpha=False)
    return page_count, " | ".join(sample_texts)


def render_doc(
    renderer: MarkdownPdfRenderer,
    *,
    sources: tuple[Path, ...],
    output: Path,
    title: str,
    subtitle: str,
    asset_base_dir: Path,
) -> None:
    renderer.render(
        sources=list(sources),
        output_path=output,
        title=title,
        subtitle=subtitle,
        asset_base_dir=asset_base_dir,
    )


def main() -> int:
    args = parse_args()
    jobs = select_jobs(args.book_id)
    renderer = MarkdownPdfRenderer()
    had_issues = False

    for job in jobs:
        print(f"== {job.book_id} ==")
        report_issues = renderer.validate_sources(list(job.report_sources), asset_base_dir=job.asset_base_dir)
        summary_issues = renderer.validate_sources(list(job.summary_sources), asset_base_dir=job.asset_base_dir)
        for issue in [*report_issues, *summary_issues]:
            had_issues = True
            print(f"[ERROR] {issue.source}: {issue.message}")
        if report_issues or summary_issues:
            continue

        for label, output in [("report", job.report_output), ("summary", job.summary_output)]:
            if output.exists():
                try:
                    pages, sample = validate_pdf(output)
                    print(f"[OK] existing {label}: {output} ({pages} pages)")
                    if sample:
                        print(f"      sample: {sample[:160]}")
                except Exception as exc:
                    had_issues = True
                    print(f"[ERROR] invalid existing {label}: {output} -> {exc}")

        if args.validate_only:
            continue

        skip_existing = job.skip_if_exists and not args.force

        if not (skip_existing and job.report_output.exists()):
            render_doc(
                renderer,
                sources=job.report_sources,
                output=job.report_output,
                title=f"{job.title}阅读报告",
                subtitle=f"{job.title} 阅读报告统一易读版",
                asset_base_dir=job.asset_base_dir,
            )
            print(f"[WRITE] {job.report_output}")
        else:
            print(f"[SKIP] report exists: {job.report_output}")

        if not (skip_existing and job.summary_output.exists()):
            render_doc(
                renderer,
                sources=job.summary_sources,
                output=job.summary_output,
                title=f"{job.title}关键知识总结",
                subtitle=f"{job.title} 关键知识总结统一易读版",
                asset_base_dir=job.asset_base_dir,
            )
            print(f"[WRITE] {job.summary_output}")
        else:
            print(f"[SKIP] summary exists: {job.summary_output}")

        for label, output in [("report", job.report_output), ("summary", job.summary_output)]:
            try:
                pages, sample = validate_pdf(output)
                print(f"[OK] generated {label}: {output} ({pages} pages)")
                if sample:
                    print(f"      sample: {sample[:160]}")
            except Exception as exc:
                had_issues = True
                print(f"[ERROR] generated {label}: {output} -> {exc}")

    return 1 if had_issues else 0


if __name__ == "__main__":
    sys.exit(main())
