from pathlib import Path
import fitz

ROOT = Path(__file__).resolve().parents[2]
PDF_DIR = ROOT / '_build' / '预览_PPT_PDF'
PNG_DIR = ROOT / '_build' / '预览_PPT_PNG'

if __name__ == '__main__':
    for pdf_path in sorted(PDF_DIR.glob('*.pdf')):
        out_dir = PNG_DIR / pdf_path.stem
        out_dir.mkdir(parents=True, exist_ok=True)
        doc = fitz.open(pdf_path)
        for page_index, page in enumerate(doc, start=1):
            pix = page.get_pixmap(matrix=fitz.Matrix(1.45, 1.45), alpha=False)
            pix.save(out_dir / f'slide_{page_index:02d}.png')
