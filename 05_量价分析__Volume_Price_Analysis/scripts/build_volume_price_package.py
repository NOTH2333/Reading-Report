import json
import os
import re
import shutil
from pathlib import Path

import fitz
from PIL import Image, ImageDraw, ImageFont
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable,
    Image as RLImage,
    KeepTogether,
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from book_teaching_content import (
    BOOK_META,
    CHAPTERS,
    GLOBAL_GLOSSARY,
    READING_REPORT_SECTIONS,
    SUMMARY_PREFACE,
)


ROOT = Path(__file__).resolve().parent.parent
OUTPUT_ROOT = ROOT / "《量价分析》教学化交付包"
INDEX_DIR = OUTPUT_ROOT / "00_总索引"
PPT_DIR = OUTPUT_ROOT / "01_PPT"
REPORT_DIR = OUTPUT_ROOT / "02_阅读报告"
SUMMARY_DIR = OUTPUT_ROOT / "03_关键知识总结"
FIGURE_DIR = OUTPUT_ROOT / "04_原书配图素材"
DATA_DIR = OUTPUT_ROOT / "05_中间数据"
RAW_TEXT_DIR = DATA_DIR / "章节原文"
DIAGRAM_DIR = DATA_DIR / "教学示意图"

PDF_SOURCE = Path(
    r"c:\Users\霓\Downloads\Documents\量价分析：量价分析创始人威科夫的盘口解读方法 (安娜·库林) (z-library.sk, 1lib.sk, z-lib.sk).pdf"
)

INTRODUCTION_PAGES = {"序言": (9, 12), "致谢&免费交易者资源": (286, 289)}


def chapter_folder_name(chapter):
    return f"第{chapter['number']:02d}章_{chapter['title']}"


def ensure_dirs():
    for path in [
        OUTPUT_ROOT,
        INDEX_DIR,
        PPT_DIR,
        REPORT_DIR,
        SUMMARY_DIR,
        FIGURE_DIR,
        DATA_DIR,
        RAW_TEXT_DIR,
        DIAGRAM_DIR,
    ]:
        path.mkdir(parents=True, exist_ok=True)


def clean_text(text):
    text = text.replace("\u3000", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def pick_evenly_spaced(items, limit=3):
    if len(items) <= limit:
        return items
    positions = [0, len(items) // 2, len(items) - 1]
    picked = []
    for pos in positions:
        value = items[pos]
        if value not in picked:
            picked.append(value)
    while len(picked) < limit:
        for value in items:
            if value not in picked:
                picked.append(value)
            if len(picked) == limit:
                break
    return picked[:limit]


def read_outline_ranges(reader):
    top_items = []
    for item in reader.outline:
        if isinstance(item, list):
            continue
        title = getattr(item, "title", str(item))
        try:
            page = reader.get_destination_page_number(item) + 1
        except Exception:
            page = None
        top_items.append({"title": title, "page": page})
    ranges = []
    for index, item in enumerate(top_items):
        start = item["page"]
        next_page = (
            top_items[index + 1]["page"]
            if index + 1 < len(top_items) and top_items[index + 1]["page"]
            else len(reader.pages)
        )
        end = next_page - 1 if start else None
        ranges.append(
            {
                "title": item["title"],
                "start_page": start,
                "end_page": end,
                "page_count": end - start + 1 if start and end else None,
            }
        )
    return ranges


def export_pdf_assets():
    reader = PdfReader(str(PDF_SOURCE))
    doc = fitz.open(str(PDF_SOURCE))
    outline_ranges = read_outline_ranges(reader)
    manifest = []
    figure_manifest = []

    for chapter in CHAPTERS:
        folder = chapter_folder_name(chapter)
        raw_text_path = RAW_TEXT_DIR / f"{folder}.txt"
        figure_output_dir = FIGURE_DIR / folder
        figure_output_dir.mkdir(parents=True, exist_ok=True)

        raw_text_parts = []
        chapter_image_pages = []
        chapter_figures = []
        for page_number in range(chapter["start_page"], chapter["end_page"] + 1):
            page_text = reader.pages[page_number - 1].extract_text() or ""
            if page_text.strip():
                raw_text_parts.append(page_text.strip())
            page = doc.load_page(page_number - 1)
            images = page.get_images(full=True)
            if not images:
                continue
            chapter_image_pages.append(page_number)
            page_render_path = figure_output_dir / f"第{page_number:03d}页_原页.png"
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            pix.save(str(page_render_path))

            extracted_paths = []
            for image_index, image_info in enumerate(images, start=1):
                xref = image_info[0]
                img_pix = fitz.Pixmap(doc, xref)
                if img_pix.alpha:
                    img_pix = fitz.Pixmap(fitz.csRGB, img_pix)
                embedded_path = figure_output_dir / f"第{page_number:03d}页_图像{image_index}.png"
                img_pix.save(str(embedded_path))
                extracted_paths.append(embedded_path)
            chapter_figures.append(
                {
                    "page": page_number,
                    "page_render": page_render_path,
                    "embedded_images": extracted_paths,
                }
            )
            figure_manifest.append(
                {
                    "chapter": folder,
                    "page": page_number,
                    "page_render": str(page_render_path.relative_to(OUTPUT_ROOT)),
                    "embedded_images": [
                        str(path.relative_to(OUTPUT_ROOT)) for path in extracted_paths
                    ],
                }
            )

        raw_text_path.write_text(clean_text("\n\n".join(raw_text_parts)), encoding="utf-8")

        manifest.append(
            {
                "number": chapter["number"],
                "title": chapter["title"],
                "folder": folder,
                "start_page": chapter["start_page"],
                "end_page": chapter["end_page"],
                "page_count": chapter["end_page"] - chapter["start_page"] + 1,
                "raw_text_file": str(raw_text_path.relative_to(OUTPUT_ROOT)),
                "image_pages": chapter_image_pages,
                "figure_count": len(chapter_figures),
            }
        )
        chapter["folder"] = folder
        chapter["image_pages"] = chapter_image_pages
        chapter["figures"] = chapter_figures
        chapter["selected_figure_pages"] = pick_evenly_spaced(chapter_image_pages, limit=3)

    (DATA_DIR / "章节清单.json").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (DATA_DIR / "PDF目录页码索引.json").write_text(
        json.dumps(outline_ranges, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    (DATA_DIR / "原书配图索引.json").write_text(
        json.dumps(figure_manifest, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return manifest, outline_ranges


def get_font_path(candidates):
    for candidate in candidates:
        path = Path(candidate)
        if path.exists():
            return str(path)
    raise FileNotFoundError(f"未找到字体文件: {candidates}")


def register_fonts():
    normal = get_font_path(
        [
            r"C:\Windows\Fonts\msyh.ttc",
            r"C:\Windows\Fonts\simhei.ttf",
            r"C:\Windows\Fonts\simsun.ttc",
        ]
    )
    bold = get_font_path(
        [
            r"C:\Windows\Fonts\msyhbd.ttc",
            r"C:\Windows\Fonts\simhei.ttf",
            r"C:\Windows\Fonts\msyh.ttc",
        ]
    )
    pdfmetrics.registerFont(TTFont("VPARegular", normal))
    pdfmetrics.registerFont(TTFont("VPABold", bold))
    return normal


def pil_font(size, bold=False):
    candidates = (
        [
            r"C:\Windows\Fonts\msyhbd.ttc",
            r"C:\Windows\Fonts\simhei.ttf",
        ]
        if bold
        else [
            r"C:\Windows\Fonts\msyh.ttc",
            r"C:\Windows\Fonts\simsun.ttc",
        ]
    )
    for candidate in candidates:
        if Path(candidate).exists():
            return ImageFont.truetype(candidate, size=size)
    return ImageFont.load_default()


def wrap_text(draw, text, font, max_width):
    lines = []
    current = ""
    for char in text:
        test = current + char
        bbox = draw.textbbox((0, 0), test, font=font)
        if bbox[2] - bbox[0] <= max_width or not current:
            current = test
        else:
            lines.append(current)
            current = char
    if current:
        lines.append(current)
    return lines


def create_teaching_diagrams():
    diagrams = {}
    font_title = pil_font(44, bold=True)
    font_body = pil_font(28)
    font_small = pil_font(24)
    palette = BOOK_META["design_theme"]

    def base_canvas(title, subtitle):
        image = Image.new("RGB", (1600, 900), palette["bg"])
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((60, 50, 1540, 850), radius=36, outline=palette["primary"], width=4)
        draw.text((100, 90), title, fill=palette["primary"], font=font_title)
        lines = wrap_text(draw, subtitle, font_body, 1320)
        y = 170
        for line in lines:
            draw.text((100, y), line, fill=palette["text"], font=font_body)
            y += 42
        return image, draw

    image, draw = base_canvas("量像音量，价像脚步", "成交量告诉你市场用了多大力；价格告诉你这些力换来了什么结果。两者一起看，才不会被表面涨跌骗到。")
    for index, bar_height in enumerate([220, 340, 180, 420, 300, 520]):
        x0 = 180 + index * 80
        draw.rounded_rectangle((x0, 720 - bar_height, x0 + 42, 720), radius=8, fill=palette["secondary"])
    draw.line((760, 700, 840, 620, 950, 660, 1080, 500, 1230, 430, 1380, 360), fill=palette["accent"], width=12)
    draw.text((170, 740), "左边：量柱像音量", fill=palette["muted"], font=font_small)
    draw.text((980, 740), "右边：价格像脚步", fill=palette["muted"], font=font_small)
    draw.rounded_rectangle((850, 200, 1440, 300), radius=20, outline=palette["primary"], width=3)
    draw.text((890, 228), "大声但走不远，要警惕背离", fill=palette["primary"], font=font_body)
    path = DIAGRAM_DIR / "00_量与价_音量与脚步.png"
    image.save(path)
    diagrams["volume_voice"] = path

    image, draw = base_canvas("努力与结果", "量价分析最重要的问题不是‘涨了还是跌了’，而是‘市场付出了多少努力，又换来了多少结果’。")
    draw.rounded_rectangle((150, 260, 620, 640), radius=28, outline=palette["primary"], width=4, fill="#E6F0F5")
    draw.rounded_rectangle((970, 260, 1440, 640), radius=28, outline=palette["accent"], width=4, fill="#FFF2E0")
    draw.text((290, 300), "努力", fill=palette["primary"], font=font_title)
    draw.text((1090, 300), "结果", fill=palette["accent"], font=font_title)
    draw.text((220, 390), "成交量\n参与度\n推动力", fill=palette["text"], font=font_body, spacing=12)
    draw.text((1050, 390), "价差\n收盘位置\n延续性", fill=palette["text"], font=font_body, spacing=12)
    draw.line((650, 450, 930, 450), fill=palette["secondary"], width=10)
    draw.polygon([(930, 450), (880, 420), (880, 480)], fill=palette["secondary"])
    draw.rounded_rectangle((520, 690, 1080, 790), radius=20, outline=palette["secondary"], width=3)
    draw.text((575, 720), "努力和结果对不上，就是异常", fill=palette["secondary"], font=font_body)
    path = DIAGRAM_DIR / "01_努力与结果.png"
    image.save(path)
    diagrams["effort_result"] = path

    image, draw = base_canvas("吸筹与派筹", "大资金不是一口气完成买卖，而是常在区间里慢慢搬货。低位更像装仓库，高位更像清仓库。")
    draw.line((180, 620, 380, 760, 520, 580, 650, 690, 820, 500, 980, 560, 1180, 420, 1410, 300), fill=palette["primary"], width=10)
    draw.rounded_rectangle((330, 390, 710, 500), radius=20, outline=palette["secondary"], width=3)
    draw.text((405, 425), "吸筹：低位慢慢收货", fill=palette["secondary"], font=font_body)
    draw.rounded_rectangle((930, 180, 1310, 290), radius=20, outline=palette["accent"], width=3)
    draw.text((1005, 215), "派筹：高位慢慢出货", fill=palette["accent"], font=font_body)
    draw.rounded_rectangle((560, 720, 1070, 810), radius=20, outline=palette["primary"], width=3)
    draw.text((640, 750), "高潮常出现在旧阶段接近结束时", fill=palette["primary"], font=font_body)
    path = DIAGRAM_DIR / "02_吸筹与派筹.png"
    image.save(path)
    diagrams["accumulation_distribution"] = path

    image, draw = base_canvas("支撑、阻力与突破", "边界不是画出来就算数。重要的是价格碰到边界时，量有没有配合，后面能不能真的站稳。")
    draw.line((180, 570, 1430, 570), fill=palette["secondary"], width=8)
    draw.line((180, 340, 1430, 340), fill=palette["accent"], width=8)
    draw.text((210, 520), "支撑区", fill=palette["secondary"], font=font_body)
    draw.text((210, 290), "阻力区", fill=palette["accent"], font=font_body)
    draw.line((250, 500, 420, 450, 520, 480, 610, 430, 700, 510, 820, 470, 940, 240, 1120, 180, 1310, 220), fill=palette["primary"], width=10)
    draw.rounded_rectangle((980, 610, 1410, 760), radius=20, outline=palette["primary"], width=3)
    draw.text((1035, 640), "真突破：有量、站稳、\n回踩后还能继续", fill=palette["primary"], font=font_body, spacing=12)
    path = DIAGRAM_DIR / "03_支撑阻力与突破.png"
    image.save(path)
    diagrams["support_breakout"] = path

    (DATA_DIR / "教学示意图索引.json").write_text(
        json.dumps(
            {key: str(value.relative_to(OUTPUT_ROOT)) for key, value in diagrams.items()},
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    return diagrams


def build_dataset(diagrams):
    dataset = {
        "book_meta": BOOK_META,
        "chapters": [],
        "glossary": GLOBAL_GLOSSARY,
        "reading_report_sections": READING_REPORT_SECTIONS,
        "summary_preface": SUMMARY_PREFACE,
        "introduction_pages": INTRODUCTION_PAGES,
    }
    for chapter in CHAPTERS:
        selected_figures = []
        for selected_page in chapter["selected_figure_pages"]:
            figure = next(item for item in chapter["figures"] if item["page"] == selected_page)
            image_path = figure["embedded_images"][0] if figure["embedded_images"] else figure["page_render"]
            page_render = figure["page_render"]
            caption = f"原书图示：第 {selected_page} 页，对应《{chapter['title']}》中的关键图表。"
            explanation = (
                f"这张图适合用来回看本章的核心句子：{chapter['one_liner']} "
                f"观察时先看背景，再看量价是否彼此确认。"
            )
            selected_figures.append(
                {
                    "page": selected_page,
                    "caption": caption,
                    "explanation": explanation,
                    "image_path": str(image_path.relative_to(OUTPUT_ROOT)),
                    "page_render": str(page_render.relative_to(OUTPUT_ROOT)),
                }
            )
        if not selected_figures and chapter.get("fallback_figure_key"):
            fallback_path = diagrams[chapter["fallback_figure_key"]]
            selected_figures.append(
                {
                    "page": None,
                    "caption": "教学示意图：用自制图解释成交量为何像市场音量。",
                    "explanation": "本章原书没有独立插图，因此使用教学图帮助零基础读者先建立直觉。",
                    "image_path": str(fallback_path.relative_to(OUTPUT_ROOT)),
                    "page_render": str(fallback_path.relative_to(OUTPUT_ROOT)),
                }
            )

        chapter_data = {key: value for key, value in chapter.items() if key not in {"figures"}}
        chapter_data["selected_figures"] = selected_figures
        chapter_data["raw_text_file"] = str(
            (RAW_TEXT_DIR / f"{chapter['folder']}.txt").relative_to(OUTPUT_ROOT)
        )
        dataset["chapters"].append(chapter_data)

    (DATA_DIR / "术语表.json").write_text(
        json.dumps(GLOBAL_GLOSSARY, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    dataset_path = DATA_DIR / "章节教学内容.json"
    dataset_path.write_text(json.dumps(dataset, ensure_ascii=False, indent=2), encoding="utf-8")
    return dataset


def rel_path(from_dir, to_file):
    return os.path.relpath(to_file, start=from_dir).replace("\\", "/")


def write_markdown_documents(dataset):
    report_md = REPORT_DIR / "《量价分析》全书阅读报告.md"
    summary_md = SUMMARY_DIR / "《量价分析》按章节关键知识总结.md"

    report_lines = [
        f"# {BOOK_META['title_cn']}全书阅读报告",
        "",
        "## 文档说明",
        f"- 作者：{BOOK_META['author']}",
        f"- 英文原名：{BOOK_META['title_en']}",
        f"- 主源说明：{BOOK_META['source_note']}",
        f"- 读者定位：{BOOK_META['reader_position']}",
        "",
        "## 摘要",
        READING_REPORT_SECTIONS["abstract"],
        "",
        f"![量与价关系图]({rel_path(REPORT_DIR, DIAGRAM_DIR / '00_量与价_音量与脚步.png')})",
        "",
        "## 作者方法定位",
        READING_REPORT_SECTIONS["method_position"],
        "",
        "## 全书总主线",
    ]
    report_lines.extend([f"1. {line}" for line in READING_REPORT_SECTIONS["main_line"]])
    report_lines.extend(
        [
            "",
            "## 核心术语总表",
            "| 术语 | 儿童化解释 | 这项术语为什么重要 |",
            "| --- | --- | --- |",
        ]
    )
    for item in GLOBAL_GLOSSARY:
        report_lines.append(f"| {item['term']} | {item['simple']} | {item['importance']} |")

    report_lines.extend(["", "## 分章精读"])
    for chapter in dataset["chapters"]:
        report_lines.extend(
            [
                "",
                f"### 第{chapter['number']:02d}章 {chapter['title']}",
                "",
                f"- 原书页码：第 {chapter['start_page']} 页到第 {chapter['end_page']} 页",
                f"- 本章一句话：{chapter['one_liner']}",
                f"- 本章回答的问题：{chapter['big_question']}",
                f"- 给零基础读者的导语：{chapter['kid_hook']}",
                "",
                "#### 核心命题",
                chapter["chapter_report"]["core_thesis"],
                "",
                "#### 逻辑链",
            ]
        )
        report_lines.extend([f"1. {item}" for item in chapter["logic_chain"]])
        report_lines.extend(["", "#### 关键术语"])
        for item in chapter["concepts"]:
            report_lines.append(f"- **{item['term']}**：{item['simple']}")
        report_lines.extend(
            [
                "",
                "#### 对零基础读者的翻译",
                chapter["chapter_report"]["beginner"],
                "",
                "#### 这一章真正教会了什么",
                chapter["chapter_report"]["deep_gain"],
                "",
                "#### 关键图解",
            ]
        )
        for figure in chapter["selected_figures"]:
            report_lines.extend(
                [
                    "",
                    f"![{figure['caption']}]({rel_path(REPORT_DIR, OUTPUT_ROOT / figure['page_render'])})",
                    "",
                    f"图注：{figure['caption']}",
                    "",
                    figure["explanation"],
                ]
            )
        report_lines.extend(["", "#### 常见误解"])
        report_lines.extend([f"- {item}" for item in chapter["misreads"]])
        report_lines.extend(["", "#### 实战观察步骤"])
        report_lines.extend([f"1. {item}" for item in chapter["practice_steps"]])
        report_lines.extend(["", "#### 本章收获"])
        report_lines.extend([f"- {item}" for item in chapter["takeaways"]])
        report_lines.extend(["", "#### 与后续章节的衔接", chapter["chapter_report"]["bridge"]])

    report_lines.extend(["", "## 初学者常犯误区"])
    report_lines.extend([f"- {item}" for item in READING_REPORT_SECTIONS["beginner_mistakes"]])
    report_lines.extend(["", "## 推荐学习顺序"])
    report_lines.extend([f"1. {item}" for item in READING_REPORT_SECTIONS["learning_order"]])
    report_lines.extend(
        [
            "",
            f"![结构阅读框架图]({rel_path(REPORT_DIR, DIAGRAM_DIR / '01_努力与结果.png')})",
            "",
            "## 可执行观察框架",
        ]
    )
    report_lines.extend([f"1. {item}" for item in READING_REPORT_SECTIONS["execution_framework"]])
    report_lines.extend(["", "## 结论", READING_REPORT_SECTIONS["conclusion"], ""])
    report_md.write_text("\n".join(report_lines), encoding="utf-8")

    summary_lines = [
        f"# {BOOK_META['title_cn']}按章节关键知识总结",
        "",
        SUMMARY_PREFACE,
        "",
        f"![吸筹与派筹概览]({rel_path(SUMMARY_DIR, DIAGRAM_DIR / '02_吸筹与派筹.png')})",
        "",
    ]
    for chapter in dataset["chapters"]:
        summary_lines.extend(
            [
                f"## 第{chapter['number']:02d}章 {chapter['title']}",
                "",
                f"- 本章一句话：{chapter['one_liner']}",
                f"- 本章页码：第 {chapter['start_page']} 页到第 {chapter['end_page']} 页",
                "",
                "### 关键概念",
            ]
        )
        summary_lines.extend([f"- **{item['term']}**：{item['simple']}" for item in chapter["concepts"]])
        summary_lines.extend(["", "### 必看原图"])
        if chapter["selected_figures"]:
            figure = chapter["selected_figures"][0]
            summary_lines.extend(
                [
                    f"![{figure['caption']}]({rel_path(SUMMARY_DIR, OUTPUT_ROOT / figure['page_render'])})",
                    "",
                    f"- 图解重点：{figure['explanation']}",
                ]
            )
        summary_lines.extend(["", "### 识别信号"])
        summary_lines.extend([f"1. {item}" for item in chapter["logic_chain"][:3]])
        summary_lines.extend(["", "### 容易看错的地方"])
        summary_lines.extend([f"- {item}" for item in chapter["misreads"]])
        summary_lines.extend(["", "### 记忆口诀"])
        summary_lines.extend([f"- {item}" for item in chapter["memory_formula"]])
        summary_lines.extend(
            [
                "",
                "### 读完这一章后应该具备的判断能力",
                f"- {chapter['summary_capability']}",
                "",
            ]
        )
    summary_md.write_text("\n".join(summary_lines), encoding="utf-8")
    return report_md, summary_md


def build_styles():
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="VPATitle",
            parent=styles["Title"],
            fontName="VPABold",
            fontSize=24,
            leading=30,
            textColor=colors.HexColor(BOOK_META["design_theme"]["primary"]),
            spaceAfter=16,
            alignment=TA_CENTER,
        )
    )
    styles.add(
        ParagraphStyle(
            name="VPAH1",
            parent=styles["Heading1"],
            fontName="VPABold",
            fontSize=18,
            leading=24,
            textColor=colors.HexColor(BOOK_META["design_theme"]["primary"]),
            spaceAfter=10,
            spaceBefore=12,
        )
    )
    styles.add(
        ParagraphStyle(
            name="VPAH2",
            parent=styles["Heading2"],
            fontName="VPABold",
            fontSize=14,
            leading=20,
            textColor=colors.HexColor(BOOK_META["design_theme"]["secondary"]),
            spaceAfter=8,
            spaceBefore=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="VPABody",
            parent=styles["BodyText"],
            fontName="VPARegular",
            fontSize=10.6,
            leading=17,
            textColor=colors.HexColor(BOOK_META["design_theme"]["text"]),
            spaceAfter=6,
        )
    )
    styles.add(
        ParagraphStyle(
            name="VPACaption",
            parent=styles["BodyText"],
            fontName="VPARegular",
            fontSize=8.8,
            leading=12,
            textColor=colors.HexColor(BOOK_META["design_theme"]["muted"]),
            alignment=TA_CENTER,
            spaceAfter=10,
        )
    )
    return styles


def add_bullets(items, style):
    bullet_items = [
        ListItem(Paragraph(item, style), leftIndent=8, value="bullet") for item in items
    ]
    return ListFlowable(
        bullet_items,
        bulletType="bullet",
        start="circle",
        leftIndent=18,
        bulletFontName="VPARegular",
    )


def scaled_reportlab_image(image_path, max_width_cm=15.6):
    image = RLImage(str(image_path))
    max_width = max_width_cm * cm
    if image.drawWidth > max_width:
        ratio = max_width / image.drawWidth
        image.drawWidth *= ratio
        image.drawHeight *= ratio
    return image


def page_decor(canvas, doc):
    canvas.saveState()
    canvas.setFont("VPARegular", 9)
    canvas.setFillColor(colors.HexColor(BOOK_META["design_theme"]["muted"]))
    canvas.drawString(2 * cm, 1.3 * cm, BOOK_META["title_cn"])
    canvas.drawRightString(A4[0] - 2 * cm, 1.3 * cm, f"第 {doc.page} 页")
    canvas.restoreState()


def build_report_pdf(dataset, output_path):
    styles = build_styles()
    story = []
    story.append(Spacer(1, 2 * cm))
    story.append(Paragraph(f"{BOOK_META['title_cn']}全书阅读报告", styles["VPATitle"]))
    story.append(Paragraph(f"作者：{BOOK_META['author']}  |  英文原名：{BOOK_META['title_en']}", styles["VPABody"]))
    story.append(Spacer(1, 0.5 * cm))
    story.append(Paragraph("摘要", styles["VPAH1"]))
    story.append(Paragraph(READING_REPORT_SECTIONS["abstract"], styles["VPABody"]))
    story.append(scaled_reportlab_image(DIAGRAM_DIR / "00_量与价_音量与脚步.png"))
    story.append(Paragraph("图 1：量像音量，价像脚步。", styles["VPACaption"]))
    story.append(Paragraph("作者方法定位", styles["VPAH1"]))
    story.append(Paragraph(READING_REPORT_SECTIONS["method_position"], styles["VPABody"]))
    story.append(Paragraph("全书主线", styles["VPAH1"]))
    story.append(add_bullets(READING_REPORT_SECTIONS["main_line"], styles["VPABody"]))
    story.append(Paragraph("术语总表", styles["VPAH1"]))

    glossary_rows = [["术语", "儿童化解释", "重要性"]]
    for item in GLOBAL_GLOSSARY:
        glossary_rows.append([item["term"], item["simple"], item["importance"]])
    glossary_table = Table(glossary_rows, colWidths=[3.2 * cm, 7.0 * cm, 6.2 * cm], repeatRows=1)
    glossary_table.setStyle(
        TableStyle(
            [
                ("FONTNAME", (0, 0), (-1, -1), "VPARegular"),
                ("FONTNAME", (0, 0), (-1, 0), "VPABold"),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF2")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor(BOOK_META["design_theme"]["primary"])),
                ("GRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#D6D9DE")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("FONTSIZE", (0, 0), (-1, -1), 8.8),
                ("LEADING", (0, 0), (-1, -1), 11),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
            ]
        )
    )
    story.append(glossary_table)
    story.append(PageBreak())

    story.append(Paragraph("分章精读", styles["VPAH1"]))
    for chapter in dataset["chapters"]:
        story.append(Paragraph(f"第{chapter['number']:02d}章 {chapter['title']}", styles["VPAH1"]))
        story.append(Paragraph(f"原书页码：第 {chapter['start_page']} 页到第 {chapter['end_page']} 页。", styles["VPABody"]))
        story.append(Paragraph(f"本章一句话：{chapter['one_liner']}", styles["VPABody"]))
        story.append(Paragraph(f"本章回答的问题：{chapter['big_question']}", styles["VPABody"]))
        story.append(Paragraph("核心命题", styles["VPAH2"]))
        story.append(Paragraph(chapter["chapter_report"]["core_thesis"], styles["VPABody"]))
        story.append(Paragraph("逻辑链", styles["VPAH2"]))
        story.append(add_bullets(chapter["logic_chain"], styles["VPABody"]))
        story.append(Paragraph("对零基础读者的翻译", styles["VPAH2"]))
        story.append(Paragraph(chapter["chapter_report"]["beginner"], styles["VPABody"]))
        story.append(Paragraph("真正的收获", styles["VPAH2"]))
        story.append(Paragraph(chapter["chapter_report"]["deep_gain"], styles["VPABody"]))
        if chapter["selected_figures"]:
            story.append(Paragraph("关键图解", styles["VPAH2"]))
            figure = chapter["selected_figures"][0]
            story.append(scaled_reportlab_image(OUTPUT_ROOT / figure["page_render"], max_width_cm=13.8))
            story.append(Paragraph(figure["caption"], styles["VPACaption"]))
            story.append(Paragraph(figure["explanation"], styles["VPABody"]))
        story.append(Paragraph("常见误解", styles["VPAH2"]))
        story.append(add_bullets(chapter["misreads"], styles["VPABody"]))
        story.append(Paragraph("实战观察步骤", styles["VPAH2"]))
        story.append(add_bullets(chapter["practice_steps"], styles["VPABody"]))
        story.append(Paragraph("本章能力目标", styles["VPAH2"]))
        story.append(Paragraph(chapter["summary_capability"], styles["VPABody"]))
        story.append(HRFlowable(color=colors.HexColor("#D6D9DE"), width="100%"))
        story.append(Spacer(1, 0.25 * cm))

    story.append(Paragraph("初学者常犯误区", styles["VPAH1"]))
    story.append(add_bullets(READING_REPORT_SECTIONS["beginner_mistakes"], styles["VPABody"]))
    story.append(Paragraph("推荐学习顺序", styles["VPAH1"]))
    story.append(add_bullets(READING_REPORT_SECTIONS["learning_order"], styles["VPABody"]))
    story.append(Paragraph("可执行观察框架", styles["VPAH1"]))
    story.append(scaled_reportlab_image(DIAGRAM_DIR / "01_努力与结果.png"))
    story.append(Paragraph("图 2：努力与结果是整本书的底层判断框架。", styles["VPACaption"]))
    story.append(add_bullets(READING_REPORT_SECTIONS["execution_framework"], styles["VPABody"]))
    story.append(Paragraph("结论", styles["VPAH1"]))
    story.append(Paragraph(READING_REPORT_SECTIONS["conclusion"], styles["VPABody"]))

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=1.8 * cm,
        leftMargin=1.8 * cm,
        topMargin=1.8 * cm,
        bottomMargin=1.8 * cm,
        title=f"{BOOK_META['title_cn']}全书阅读报告",
        author="OpenAI Codex",
    )
    doc.build(story, onFirstPage=page_decor, onLaterPages=page_decor)


def build_summary_pdf(dataset, output_path):
    styles = build_styles()
    story = [
        Spacer(1, 1.6 * cm),
        Paragraph(f"{BOOK_META['title_cn']}按章节关键知识总结", styles["VPATitle"]),
        Paragraph(SUMMARY_PREFACE, styles["VPABody"]),
        scaled_reportlab_image(DIAGRAM_DIR / "02_吸筹与派筹.png"),
        Paragraph("图 1：用库存转换视角理解市场结构。", styles["VPACaption"]),
    ]
    for chapter in dataset["chapters"]:
        block = [
            Paragraph(f"第{chapter['number']:02d}章 {chapter['title']}", styles["VPAH1"]),
            Paragraph(f"本章一句话：{chapter['one_liner']}", styles["VPABody"]),
            Paragraph("关键概念", styles["VPAH2"]),
            add_bullets([f"{item['term']}：{item['simple']}" for item in chapter["concepts"]], styles["VPABody"]),
            Paragraph("识别信号", styles["VPAH2"]),
            add_bullets(chapter["logic_chain"][:3], styles["VPABody"]),
            Paragraph("容易看错的地方", styles["VPAH2"]),
            add_bullets(chapter["misreads"], styles["VPABody"]),
            Paragraph("记忆口诀", styles["VPAH2"]),
            add_bullets(chapter["memory_formula"], styles["VPABody"]),
            Paragraph("读完这一章后应该具备的判断能力", styles["VPAH2"]),
            Paragraph(chapter["summary_capability"], styles["VPABody"]),
        ]
        if chapter["selected_figures"]:
            figure = chapter["selected_figures"][0]
            block.extend(
                [
                    Paragraph("必看图示", styles["VPAH2"]),
                    scaled_reportlab_image(OUTPUT_ROOT / figure["page_render"], max_width_cm=11.8),
                    Paragraph(figure["caption"], styles["VPACaption"]),
                ]
            )
        story.append(KeepTogether(block))
        story.append(HRFlowable(color=colors.HexColor("#D6D9DE"), width="100%"))
        story.append(Spacer(1, 0.2 * cm))

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        rightMargin=1.8 * cm,
        leftMargin=1.8 * cm,
        topMargin=1.8 * cm,
        bottomMargin=1.8 * cm,
        title=f"{BOOK_META['title_cn']}按章节关键知识总结",
        author="OpenAI Codex",
    )
    doc.build(story, onFirstPage=page_decor, onLaterPages=page_decor)


def write_index(dataset):
    lines = [
        f"# {BOOK_META['title_cn']}教学化交付包",
        "",
        "## 交付说明",
        f"- 主源：`{PDF_SOURCE}`",
        "- 整体目标：把全书改写成零基础也能读懂的学习包，包含 12 个章节 PPT、全书阅读报告、按章节关键知识总结，以及中间数据与原书配图素材。",
        "- 使用顺序建议：先看 `02_阅读报告`，再按章节看 `01_PPT`，最后用 `03_关键知识总结` 反复复盘。",
        "",
        "## 文件夹总览",
        "- `00_总索引`：总说明与阅读顺序。",
        "- `01_PPT`：12 个章节的独立 PPT 文件夹，每章内含 `.pptx`、生成脚本与本章图片素材。",
        "- `02_阅读报告`：全书阅读报告的 Markdown 与 PDF 版本。",
        "- `03_关键知识总结`：按章节关键知识总结的 Markdown 与 PDF 版本。",
        "- `04_原书配图素材`：从原书 PDF 中提取并按章节整理的图像。",
        "- `05_中间数据`：章节清单、图像索引、术语表、章节原文、教学内容 JSON 与教学示意图。",
        "",
        "## 章节列表",
    ]
    for chapter in dataset["chapters"]:
        lines.append(
            f"- 第{chapter['number']:02d}章《{chapter['title']}》：原书第 {chapter['start_page']} 到 {chapter['end_page']} 页。"
        )
    lines.extend(
        [
            "",
            "## 重点文档",
            "- [阅读报告](../02_阅读报告/《量价分析》全书阅读报告.md)",
            "- [关键知识总结](../03_关键知识总结/《量价分析》按章节关键知识总结.md)",
            "",
            "## 备注",
            "- 序言和致谢没有单独制作 PPT，但内容已整合进阅读报告与中间索引。",
            "- 所有中文文件均统一命名，无随机缩写与未整理的临时文件。",
        ]
    )
    (INDEX_DIR / "README.md").write_text("\n".join(lines), encoding="utf-8")


def prepare_ppt_chapter_assets(dataset):
    for chapter in dataset["chapters"]:
        chapter_root = PPT_DIR / chapter["folder"]
        images_dir = chapter_root / "images"
        images_dir.mkdir(parents=True, exist_ok=True)
        for figure in chapter["selected_figures"]:
            source = OUTPUT_ROOT / figure["image_path"]
            target = images_dir / Path(figure["image_path"]).name
            if source.exists():
                shutil.copy2(source, target)


def main():
    if not PDF_SOURCE.exists():
        raise FileNotFoundError(f"未找到主源 PDF：{PDF_SOURCE}")
    ensure_dirs()
    register_fonts()
    export_pdf_assets()
    diagrams = create_teaching_diagrams()
    dataset = build_dataset(diagrams)
    prepare_ppt_chapter_assets(dataset)
    write_markdown_documents(dataset)
    build_report_pdf(dataset, REPORT_DIR / "《量价分析》全书阅读报告.pdf")
    build_summary_pdf(dataset, SUMMARY_DIR / "《量价分析》按章节关键知识总结.pdf")
    write_index(dataset)
    print(f"已完成文档和中间数据生成：{OUTPUT_ROOT}")


if __name__ == "__main__":
    main()
