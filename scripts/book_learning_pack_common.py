from __future__ import annotations

import json
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from dual_book_specs import BOOK_PROJECTS, palette_for


def choose_font(preferred: str) -> str:
    candidates = [
        Path("C:/Windows/Fonts") / preferred,
        Path("C:/Windows/Fonts/msyhbd.ttc"),
        Path("C:/Windows/Fonts/msyh.ttc"),
        Path("C:/Windows/Fonts/simhei.ttf"),
        Path("C:/Windows/Fonts/simsun.ttc"),
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    raise FileNotFoundError("未找到可用中文字体。")


FONT_BOLD = choose_font("msyhbd.ttc")
FONT_REGULAR = choose_font("msyh.ttc")


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REGULAR, size=size)


def hex_rgb(value: str) -> tuple[int, int, int]:
    value = value.strip().lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


def rgba(value: str, alpha: int = 255) -> tuple[int, int, int, int]:
    r, g, b = hex_rgb(value)
    return (r, g, b, alpha)


def card(draw: ImageDraw.ImageDraw, xy: tuple[int, int, int, int], fill_color, outline=None, radius: int = 24, width: int = 2) -> None:
    draw.rounded_rectangle(xy, radius=radius, fill=fill_color, outline=outline, width=width)


def wrap_text(draw: ImageDraw.ImageDraw, text: str, text_font: ImageFont.FreeTypeFont, max_width: int) -> list[str]:
    lines: list[str] = []
    for raw_line in text.splitlines():
        current = ""
        for char in raw_line:
            trial = current + char
            box = draw.textbbox((0, 0), trial, font=text_font)
            if box[2] - box[0] <= max_width or not current:
                current = trial
            else:
                lines.append(current)
                current = char
        lines.append(current or "")
    return lines or [""]


def draw_paragraph(draw: ImageDraw.ImageDraw, text: str, x: int, y: int, max_width: int, text_font: ImageFont.FreeTypeFont, fill, line_gap: int = 6) -> int:
    top = y
    for line in wrap_text(draw, text, text_font, max_width):
        draw.text((x, y), line, font=text_font, fill=fill)
        box = draw.textbbox((x, y), line or "字", font=text_font)
        y += box[3] - box[1] + line_gap
    return y - top


def load_manifest(project_root: Path) -> dict:
    path = project_root / "05_中间素材" / "chapter_manifest.json"
    return json.loads(path.read_text(encoding="utf-8"))


def build_book_map(spec: dict) -> list[dict]:
    intro = spec["content"]["intro_deck"]
    wrap = spec["content"]["wrap_up_deck"]
    chapter_count = len(spec["chapters"])
    return [
        {"title": "导读", "body": intro["core_message"]},
        *intro["section_cards"],
        {"title": "正文主章节", "body": f"全书按 {chapter_count} 个正文主章节拆分讲解，每章都转成儿童可读版。"},
        {"title": "全书收束", "body": wrap["core_message"]},
    ]


def unique_pages(chapter: dict) -> list[int]:
    ordered: list[int] = []
    for page in [chapter["start_page"], *chapter["reference_pages"], chapter["end_page"]]:
        if page not in ordered:
            ordered.append(page)
    return ordered


def merge_content(spec: dict, manifest: dict) -> dict:
    book = spec["content"]
    merged_chapters: list[dict] = []
    for chapter in manifest["chapters"]:
        extra = book["chapters"][chapter["id"]]
        item = {**chapter, **extra}
        item["number"] = chapter["sequence_no"]
        item["page_range"] = f"{chapter['start_page']}-{chapter['end_page']}"
        item["palette"] = palette_for(int(chapter["sequence_no"]))
        item["citation_pages"] = unique_pages(chapter)[:3]
        item["image_refs"] = [
            f"{chapter['id']}_opener_page.png",
            f"{chapter['id']}_core_map.png",
            f"{chapter['id']}_memory_card.png",
            f"{chapter['id']}_poster.png",
        ]
        item["chapter_gains"] = extra["must_remember"][:3]
        item["report_paragraphs"] = [
            f"这一章位于{chapter['source_part']}，它想回答的问题是：{extra['children_question']} 本章一句话可以概括为：{extra['core_message']}",
            f"如果把作者的表达翻成行动语言，就是：{extra['investment_meaning']} 也因此，本章不是要你背句子，而是要你知道以后该怎样做。",
        ]
        merged_chapters.append(item)

    special_decks = []
    for source in [book["intro_deck"], book["wrap_up_deck"]]:
        special_decks.append(
            {
                **source,
                "book_title": spec["title"],
                "image_refs": ["book_overview.png", "common_methods.png"],
            }
        )

    return {
        "audience": "完全零基础读者，孩子也能读懂",
        "tone": "白话、教学化、先类比后概念",
        "book": {
            "title": spec["title"],
            "reader_positioning": book["reader_positioning"],
            "abstract": book["abstract"],
            "front_matter_notes": book["front_matter_notes"],
            "main_thesis": book["main_thesis"],
            "cross_methods": book["cross_methods"],
            "daily_applications": book["daily_applications"],
            "limitations": book["limitations"],
            "conclusion": book["conclusion"],
            "appendix": book["appendix"],
        },
        "book_map": build_book_map(spec),
        "special_decks": special_decks,
        "chapters": merged_chapters,
        "cross_book_takeaways": book["cross_book_takeaways"],
    }


def report_image(filename: str) -> str:
    return f"../04_配图/阅读报告/{filename}"


def summary_image(filename: str) -> str:
    return f"../04_配图/关键知识总结/{filename}"


def ppt_image(filename: str) -> str:
    return f"../04_配图/PPT/{filename}"


def source_image(filename: str) -> str:
    return f"../04_配图/原书页图/{filename}"


def render_book_overview(content: dict, out_path: Path) -> None:
    image = Image.new("RGBA", (1700, 1000), rgba("F8FAFC"))
    draw = ImageDraw.Draw(image)
    draw.text((80, 60), f"{content['book']['title']} 全书结构总览", font=font(42, True), fill=hex_rgb("0F172A"))
    draw.text((80, 118), "先知道这本书的路线，再进入每章细读。", font=font(22), fill=hex_rgb("475569"))
    cards = content["book_map"]
    positions = [(80, 190), (580, 190), (1080, 190), (80, 510), (580, 510), (1080, 510)]
    for (x, y), item in zip(positions, cards[:6]):
        card(draw, (x, y, x + 420, y + 240), rgba("FFFFFF"), outline=hex_rgb("CBD5E1"))
        draw.text((x + 24, y + 22), item["title"], font=font(28, True), fill=hex_rgb("1D4ED8"))
        draw_paragraph(draw, item["body"], x + 24, y + 74, 372, font(20), hex_rgb("334155"), 5)
    image.save(out_path)


def render_common_methods(content: dict, out_path: Path) -> None:
    image = Image.new("RGBA", (1700, 1050), rgba("EFF6FF"))
    draw = ImageDraw.Draw(image)
    draw.text((80, 60), "全书反复出现的核心方法", font=font(42, True), fill=hex_rgb("1E3A8A"))
    draw.text((80, 118), "这几条方法会同时出现在阅读报告、总结和 PPT 里。", font=font(22), fill=hex_rgb("475569"))
    y = 190
    for index, item in enumerate(content["book"]["cross_methods"], start=1):
        card(draw, (80, y, 1620, y + 130), rgba("FFFFFF"), outline=hex_rgb("93C5FD"))
        draw.text((108, y + 22), f"{index}. {item['title']}", font=font(25, True), fill=hex_rgb("1D4ED8"))
        draw_paragraph(draw, item["body"], 110, y + 62, 1460, font(20), hex_rgb("334155"), 5)
        y += 150
    image.save(out_path)


def render_chapter_core_map(chapter: dict, out_path: Path) -> None:
    colors = chapter["palette"]
    image = Image.new("RGBA", (1700, 980), rgba(colors["bg"]))
    draw = ImageDraw.Draw(image)
    draw.text((70, 50), f"{chapter['full_title']} 核心地图", font=font(38, True), fill=hex_rgb(colors["primary"]))
    draw.text((70, 106), chapter["core_message"], font=font(22, True), fill=hex_rgb(colors["dark"]))
    positions = [(70, 200), (580, 200), (1090, 200)]
    for (x, y), item in zip(positions, chapter["key_concepts"][:3]):
        card(draw, (x, y, x + 440, y + 240), rgba("FFFFFF"), outline=hex_rgb(colors["secondary"]))
        draw.text((x + 24, y + 20), item["name"], font=font(26, True), fill=hex_rgb(colors["primary"]))
        draw_paragraph(draw, item["explain"], x + 24, y + 74, 390, font(20), hex_rgb("334155"), 5)
    card(draw, (70, 500, 1620, 760), rgba("FFFFFF"), outline=hex_rgb(colors["secondary"]))
    draw.text((96, 530), "给孩子的生活类比", font=font(24, True), fill=hex_rgb(colors["accent"]))
    draw_paragraph(draw, chapter["child_example"], 96, 580, 1460, font(22), hex_rgb("334155"), 6)
    draw.text((96, 835), "一句话复盘", font=font(22, True), fill=hex_rgb(colors["primary"]))
    draw_paragraph(draw, chapter["one_line_review"], 250, 830, 1280, font(20), hex_rgb("334155"), 5)
    image.save(out_path)


def render_summary_card(chapter: dict, out_path: Path) -> None:
    colors = chapter["palette"]
    image = Image.new("RGBA", (1600, 930), rgba("FFFFFF"))
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, 1600, 120), fill=rgba(colors["primary"]))
    draw.text((70, 34), f"{chapter['full_title']} 记忆卡", font=font(36, True), fill=hex_rgb("FFFFFF"))
    coords = [(70, 170), (810, 170), (70, 340), (810, 340), (70, 510)]
    for index, ((x, y), bullet) in enumerate(zip(coords, chapter["must_remember"][:5]), start=1):
        card(draw, (x, y, x + 700, y + 130), rgba(colors["bg"]), outline=hex_rgb(colors["secondary"]))
        draw.text((x + 22, y + 18), f"{index}.", font=font(22, True), fill=hex_rgb(colors["primary"]))
        draw_paragraph(draw, bullet, x + 72, y + 16, 600, font(21), hex_rgb("334155"), 5)
    card(draw, (70, 700, 1530, 840), rgba("F8FAFC"), outline=hex_rgb(colors["secondary"]))
    draw.text((94, 726), "一页带走", font=font(22, True), fill=hex_rgb(colors["primary"]))
    draw_paragraph(draw, chapter["one_line_review"], 220, 722, 1260, font(22), hex_rgb("334155"), 5)
    image.save(out_path)


def render_poster(chapter: dict, out_path: Path) -> None:
    colors = chapter["palette"]
    image = Image.new("RGBA", (1600, 900), rgba(colors["dark"]))
    draw = ImageDraw.Draw(image)
    card(draw, (60, 60, 1540, 840), rgba(colors["bg"], 245), outline=hex_rgb(colors["secondary"]))
    draw.text((110, 96), chapter["full_title"], font=font(40, True), fill=hex_rgb(colors["dark"]))
    draw.text((110, 160), chapter["core_message"], font=font(24, True), fill=hex_rgb(colors["primary"]))
    y = 260
    for item in chapter["key_concepts"][:3]:
        card(draw, (110, y, 1490, y + 120), rgba("FFFFFF"), outline=hex_rgb(colors["secondary"]))
        draw.text((138, y + 18), item["name"], font=font(24, True), fill=hex_rgb(colors["primary"]))
        draw_paragraph(draw, item["explain"], 360, y + 18, 1080, font(20), hex_rgb("334155"), 5)
        y += 140
    draw.text((110, 760), "立刻练习", font=font(22, True), fill=hex_rgb(colors["accent"]))
    draw_paragraph(draw, " / ".join(chapter["child_actions"][:3]), 220, 756, 1200, font(19), hex_rgb("334155"), 5)
    image.save(out_path)


def render_special_poster(deck: dict, title: str, out_path: Path) -> None:
    image = Image.new("RGBA", (1600, 900), rgba("E0F2FE"))
    draw = ImageDraw.Draw(image)
    card(draw, (60, 60, 1540, 840), rgba("FFFFFF"), outline=hex_rgb("93C5FD"))
    draw.text((100, 96), f"{title} {deck['deck_title']}", font=font(42, True), fill=hex_rgb("1E3A8A"))
    draw.text((100, 168), deck["core_message"], font=font(24, True), fill=hex_rgb("0F172A"))
    y = 270
    for item in deck["section_cards"][:4]:
        card(draw, (100, y, 1480, y + 110), rgba("EFF6FF"), outline=hex_rgb("BFDBFE"))
        draw.text((124, y + 16), item["title"], font=font(24, True), fill=hex_rgb("1D4ED8"))
        draw_paragraph(draw, item["body"], 340, y + 18, 1080, font(20), hex_rgb("334155"), 5)
        y += 130
    image.save(out_path)


def build_report(content: dict) -> str:
    book = content["book"]
    intro = content["special_decks"][0]
    wrap = content["special_decks"][1]
    lines = [
        f"# {book['title']} 阅读报告（儿童可读版）",
        "",
        f"> 适读定位：{book['reader_positioning']}",
        "",
        f"![全书结构图]({report_image('book_overview.png')})",
        "",
        "## 摘要",
        "",
    ]
    for paragraph in book["abstract"]:
        lines.extend([paragraph, ""])
    lines.extend(["## 导读：先怎样进入这本书", "", intro["why_it_matters"], "", intro["structure_summary"], "", "### 三个阅读钥匙", ""])
    for item in intro["reading_keys"]:
        lines.append(f"- {item}")
    lines.extend(["", f"![全书方法图]({report_image('common_methods.png')})", "", "## 全书主论点", ""])
    for index, item in enumerate(book["main_thesis"], start=1):
        lines.append(f"{index}. {item}")
    lines.extend(["", "## 章节精读", ""])
    for chapter in content["chapters"]:
        lines.extend([f"### {chapter['full_title']}", "", f"![章节页图]({source_image(f'{chapter['id']}_opener_page.png')})", "", f"![核心地图]({report_image(f'{chapter['id']}_core_map.png')})", ""])
        for paragraph in chapter["report_paragraphs"]:
            lines.extend([paragraph, ""])
        lines.extend(["本章关键概念：", ""])
        for item in chapter["key_concepts"]:
            lines.append(f"- {item['name']}：{item['explain']}")
        lines.extend(["", "代表性故事：", ""])
        for item in chapter["stories"]:
            lines.append(f"- {item['title']}：{item['summary']} 启示：{item['lesson']}")
        lines.extend(["", "孩子读完能拿走什么：", ""])
        for item in chapter["chapter_gains"]:
            lines.append(f"- {item}")
        lines.extend(["", "容易误读的地方：", ""])
        for item in chapter["misunderstandings"]:
            lines.append(f"- {item}")
        lines.extend(["", "立刻可练的动作：", ""])
        for item in chapter["child_actions"]:
            lines.append(f"- {item}")
        lines.extend(["", f"一句话复盘：{chapter['one_line_review']}", ""])
    lines.extend(["## 全书收束", "", wrap["why_it_matters"], "", wrap["structure_summary"], "", "### 如果只记 5 件事", ""])
    for item in wrap["key_takeaways"]:
        lines.append(f"- {item}")
    lines.extend(["", "## 跨章节共性方法", ""])
    for item in book["cross_methods"]:
        lines.extend([f"### {item['title']}", item["body"], ""])
    lines.extend(["## 现实应用", ""])
    for item in book["daily_applications"]:
        lines.append(f"- {item}")
    lines.extend(["", "## 局限与提醒", ""])
    for item in book["limitations"]:
        lines.append(f"- {item}")
    lines.extend(["", "## 结论", ""])
    for item in book["conclusion"]:
        lines.extend([item, ""])
    lines.extend(["## 附录说明", ""])
    for item in book["appendix"]:
        lines.append(f"- {item}")
    lines.append("")
    return "\n".join(lines)


def build_summary(content: dict) -> str:
    intro = content["special_decks"][0]
    wrap = content["special_decks"][1]
    lines = [
        f"# {content['book']['title']} 关键知识总结（按章节）",
        "",
        "这份总结只保留每一部分最该带走的骨架，方便第一次接触本书的人快速建立理解地图。",
        "",
        "## 00 导读",
        "",
        intro["core_message"],
        "",
    ]
    for item in intro["key_takeaways"]:
        lines.append(f"- {item}")
    lines.extend(["", f"![全书结构图]({report_image('book_overview.png')})", ""])
    for chapter in content["chapters"]:
        lines.extend([f"## {chapter['full_title']}", "", "### 这一章在讲什么", chapter["core_message"], "", f"![记忆卡]({summary_image(f'{chapter['id']}_memory_card.png')})", "", "### 必须记住的 5 个点"])
        for item in chapter["must_remember"]:
            lines.append(f"- {item}")
        lines.extend(["", "### 一个孩子也能懂的例子", chapter["child_example"], "", "### 一句话复盘", chapter["one_line_review"], "", f"![章节海报]({ppt_image(f'{chapter['id']}_poster.png')})", ""])
    lines.extend(["## 99 全书收束", "", wrap["core_message"], "", "### 如果只记五件事"])
    for item in wrap["key_takeaways"]:
        lines.append(f"- {item}")
    lines.extend(["", f"![全书方法图]({report_image('common_methods.png')})", ""])
    return "\n".join(lines)


def build_project_readme(spec: dict, content: dict) -> str:
    lines = [
        f"# {spec['title']} 教学化交付包",
        "",
        "## 主要入口",
        f"- `01_阅读报告/{spec['report_filename'].replace('.md', '.pdf')}`：阅读报告 PDF。",
        f"- `02_关键知识总结/{spec['summary_filename'].replace('.md', '.pdf')}`：按章节关键知识总结 PDF。",
        "- `03_PPT/`：章节与导读/收束 PPT。",
        "- `04_配图/`：阅读报告、知识总结、PPT 共用图片。",
        "- `05_中间素材/`：章节原文、清单、JSON 数据与导出预览。",
        "",
        "## 目录说明",
        "- `00_项目说明/`：脚本与项目说明。",
        "- `01_阅读报告/`：阅读报告 Markdown 与 PDF。",
        "- `02_关键知识总结/`：关键知识总结 Markdown 与 PDF。",
        "- `03_PPT/`：每个 deck 的 `.js` 与 `.pptx`。",
        "- `04_配图/`：原书页图、阅读报告图、总结图、PPT 海报图。",
        "- `05_中间素材/`：章节原文、`chapter_manifest.json`、`chapter_content.json`、导出预览。",
        "",
        "## 当前正文边界",
    ]
    for chapter in content["chapters"]:
        lines.append(f"- {chapter['full_title']}：{chapter['page_range']}（原书标签：{chapter['original_label']}）")
    lines.extend(["", "## 重建顺序", "1. 运行 `extract_pdf_assets.py` 抽取原书文本和页图。", "2. 运行 `build_learning_pack.py` 生成 JSON、配图、Markdown 和说明。", "3. 运行 `build_ppts.js` 生成全部 `.pptx`。", "4. 运行 `export_ppts_to_pdf.ps1` 与 `render_pdf_to_png.py` 输出校验预览。", ""])
    return "\n".join(lines)


def write_json(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def materialize_learning_pack(book_key: str, project_root: Path) -> dict:
    spec = BOOK_PROJECTS[book_key]
    manifest = load_manifest(project_root)
    content = merge_content(spec, manifest)

    report_dir = project_root / "01_阅读报告"
    summary_dir = project_root / "02_关键知识总结"
    report_img_dir = project_root / "04_配图" / "阅读报告"
    summary_img_dir = project_root / "04_配图" / "关键知识总结"
    ppt_img_dir = project_root / "04_配图" / "PPT"
    info_dir = project_root / "00_项目说明"
    for path in [report_dir, summary_dir, report_img_dir, summary_img_dir, ppt_img_dir, info_dir]:
        path.mkdir(parents=True, exist_ok=True)

    write_json(project_root / "05_中间素材" / "chapter_content.json", content)

    render_book_overview(content, report_img_dir / "book_overview.png")
    render_common_methods(content, report_img_dir / "common_methods.png")
    for chapter in content["chapters"]:
        render_chapter_core_map(chapter, report_img_dir / f"{chapter['id']}_core_map.png")
        render_summary_card(chapter, summary_img_dir / f"{chapter['id']}_memory_card.png")
        render_poster(chapter, ppt_img_dir / f"{chapter['id']}_poster.png")
    render_special_poster(content["special_decks"][0], spec["title"], ppt_img_dir / "deck00_poster.png")
    render_special_poster(content["special_decks"][1], spec["title"], ppt_img_dir / "deck99_poster.png")

    (report_dir / spec["report_filename"]).write_text(build_report(content), encoding="utf-8")
    (summary_dir / spec["summary_filename"]).write_text(build_summary(content), encoding="utf-8")
    readme_text = build_project_readme(spec, content)
    (project_root / "README.md").write_text(readme_text, encoding="utf-8")
    (info_dir / "README.md").write_text(readme_text, encoding="utf-8")
    return content
