from __future__ import annotations

from pathlib import Path

import fitz


ROOT = Path(__file__).resolve().parents[2]
PREVIEW_DIR = ROOT / "05_中间素材" / "导出预览"


if __name__ == "__main__":
    pdf_files = sorted(PREVIEW_DIR.glob("*.pdf"))
    if not pdf_files:
        raise FileNotFoundError(f"未在 {PREVIEW_DIR} 找到 PDF 预览文件。")
    for pdf_path in pdf_files:
        out_dir = PREVIEW_DIR / pdf_path.stem
        out_dir.mkdir(parents=True, exist_ok=True)
        doc = fitz.open(pdf_path)
        for page_index, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=fitz.Matrix(1.45, 1.45), alpha=False)
            pix.save(out_dir / f"slide_{page_index:02d}.png")
