# 目录说明

本项目已经按“源文件 / 中间数据 / 生成脚本 / 最终交付 / 临时检查”分层整理。

## 根目录结构

```text
项目根目录/
├─ source/
│  └─ pdfs/                  # 两本源 PDF，本地重建优先使用这里
├─ data/
│  ├─ chapter_texts/         # 从注疏版按目录切分出来的正文/附录文本
│  └─ metadata/              # 结构化 JSON：目录、页码映射、内容矩阵
├─ scripts/
│  ├─ build/                 # 生成链路脚本
│  ├─ helpers/
│  │  └─ pptxgenjs/          # PPT 生成辅助函数
│  └─ qa/                    # 渲染与检查脚本
├─ output/
│  ├─ ppt/                   # 最终 PPTX 交付物
│  ├─ md/                    # 最终 Markdown 交付物
│  └─ assets/                # PPT / Markdown 共用图片素材
├─ tmp/
│  ├─ logs/                  # 构建日志
│  └─ rendered/              # PPT 渲染后的图片检查结果
├─ package.json
└─ package-lock.json
```

## 各目录用途

- `source/pdfs/`
  - 保存源书 PDF。
  - `scripts/build/extract_book.py` 会优先从这里找书；如果这里缺失，才回退到 `Downloads`。

- `data/chapter_texts/`
  - 保存从 PDF 提取出来的章节文本。
  - 文件名按书签顺序编号，含前言、正文和附录。

- `data/metadata/`
  - `book_structure.json`：目录结构、章节范围、章节文本路径、提取图片路径。
  - `page_map.json`：每页文本和所属章节的映射。
  - `book_content.json`：生成 PPT 与 Markdown 使用的内容矩阵。

- `scripts/build/`
  - `extract_book.py`：提取目录、章节文本、页码映射、书内图像。
  - `build_content.py`：把书本内容整理成可直接生成交付物的结构化矩阵。
  - `generate_assets.py`：生成概念卡、路线图等图片素材。
  - `generate_decks.js`：批量生成 24 章 PPT 和附录 PPT。
  - `generate_markdown.py`：生成阅读报告与关键知识总结。

- `scripts/helpers/pptxgenjs/`
  - 存放 `PptxGenJS` 的布局、图像、文本辅助函数。

- `scripts/qa/`
  - `export_ppt_png.ps1`：把 `output/ppt/` 渲染到 `tmp/rendered/`，用于排版检查。

- `output/`
  - 这是最终交付目录，已保持不变，方便直接查看和分享。

- `tmp/`
  - 只放临时产物和日志，不放最终文件。

## 推荐重建顺序

在项目根目录执行：

```powershell
python scripts/build/extract_book.py
python scripts/build/build_content.py
python scripts/build/generate_assets.py
node scripts/build/generate_decks.js
python scripts/build/generate_markdown.py
powershell -ExecutionPolicy Bypass -File scripts/qa/export_ppt_png.ps1
```

## 当前交付物位置

- PPT：`output/ppt/`
- Markdown：`output/md/`
- 图片素材：`output/assets/`

## 维护原则

- 新的源材料放进 `source/`。
- 新的中间 JSON / 切分文本放进 `data/`。
- 生成类脚本放进 `scripts/build/`。
- 检查类脚本放进 `scripts/qa/`。
- 临时文件不要混入 `output/`。
