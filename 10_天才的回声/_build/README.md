# 《天才的回声》 构建说明

1. `python scripts/build_pack.py`：生成章节原文、结构化 JSON、配图、Markdown 和 PDF。
2. `node scripts/build_ppts.js`：生成全部章节 PPTX。
3. `powershell -ExecutionPolicy Bypass -File scripts/export_ppts_to_pdf.ps1`：把 PPTX 导出为预览 PDF。
4. `python scripts/render_pdf_to_png.py`：把预览 PDF 转成逐页 PNG。
