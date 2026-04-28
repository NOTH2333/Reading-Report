from __future__ import annotations

import json
from pathlib import Path
from typing import Iterable

from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[2]
INTERMEDIATE_DIR = ROOT / "05_中间素材"
REPORT_DIR = ROOT / "01_阅读报告"
SUMMARY_DIR = ROOT / "02_关键知识总结"
REPORT_IMG_DIR = ROOT / "04_配图" / "阅读报告"
SUMMARY_IMG_DIR = ROOT / "04_配图" / "关键知识总结"
PPT_IMG_DIR = ROOT / "04_配图" / "PPT"
SOURCE_IMG_DIR = ROOT / "04_配图" / "原书页图"


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def hex_rgb(value: str) -> tuple[int, int, int]:
    value = value.strip().lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


def rgba(value: str, alpha: int = 255) -> tuple[int, int, int, int]:
    r, g, b = hex_rgb(value)
    return (r, g, b, alpha)


def choose_font_path(preferred: str) -> str:
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


FONT_BOLD = choose_font_path("msyhbd.ttc")
FONT_REGULAR = choose_font_path("msyh.ttc")


def font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REGULAR, size=size)


def wrap_text(draw: ImageDraw.ImageDraw, text: str, text_font: ImageFont.FreeTypeFont, max_width: int) -> list[str]:
    if not text:
        return [""]
    lines: list[str] = []
    for raw_line in text.splitlines():
        current = ""
        for char in raw_line:
            trial = current + char
            bbox = draw.textbbox((0, 0), trial, font=text_font)
            if bbox[2] - bbox[0] <= max_width or not current:
                current = trial
            else:
                lines.append(current)
                current = char
        lines.append(current or "")
    return lines


def draw_paragraph(
    draw: ImageDraw.ImageDraw,
    text: str,
    x: int,
    y: int,
    max_width: int,
    text_font: ImageFont.FreeTypeFont,
    fill: tuple[int, int, int],
    line_gap: int = 8,
) -> int:
    lines = wrap_text(draw, text, text_font, max_width)
    top = y
    for line in lines:
        draw.text((x, y), line, font=text_font, fill=fill)
        bbox = draw.textbbox((x, y), line or "字", font=text_font)
        y += bbox[3] - bbox[1] + line_gap
    return y - top


def card(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int, int, int],
    fill_color: tuple[int, int, int, int],
    outline: tuple[int, int, int] | None = None,
    radius: int = 26,
    width: int = 2,
) -> None:
    draw.rounded_rectangle(xy, radius=radius, fill=fill_color, outline=outline, width=width)


def render_overview(book: dict, chapters: list[dict]) -> None:
    image = Image.new("RGBA", (1600, 900), rgba("F7F4EE"))
    draw = ImageDraw.Draw(image)

    draw.text((90, 70), f"{book['title']} 全书结构总览", font=font(42, bold=True), fill=hex_rgb("2C2C2C"))
    draw.text((90, 130), "从“认识芒格”到“把思维方法用于现实世界”的五步路线图", font=font(24), fill=hex_rgb("55606E"))

    x = 70
    y = 230
    card_w = 280
    gap = 20
    for idx, chapter in enumerate(chapters):
        colors = chapter["palette"]
        card(draw, (x, y, x + card_w, y + 460), rgba(colors["bg"]), outline=hex_rgb(colors["primary"]), radius=30, width=4)
        draw.text((x + 24, y + 24), f"第{chapter['number']}章", font=font(22, bold=True), fill=hex_rgb(colors["primary"]))
        draw.text((x + 24, y + 60), chapter["title"], font=font(28, bold=True), fill=hex_rgb(colors["dark"]))
        h = draw_paragraph(draw, chapter["core_message"], x + 24, y + 115, card_w - 48, font(19), hex_rgb("39424E"), 6)
        draw.text((x + 24, y + 145 + h), "这章最像：", font=font(18, bold=True), fill=hex_rgb(colors["accent"]))
        draw_paragraph(draw, chapter["position_in_book"], x + 24, y + 180 + h, card_w - 48, font(17), hex_rgb("4A5561"), 6)
        draw.text((x + 24, y + 290 + h), "孩子要带走：", font=font(18, bold=True), fill=hex_rgb(colors["accent"]))
        base_y = y + 325 + h
        for bullet in chapter["must_remember"][:3]:
            draw.text((x + 30, base_y), "•", font=font(18, bold=True), fill=hex_rgb(colors["primary"]))
            used = draw_paragraph(draw, bullet, x + 52, base_y - 2, card_w - 82, font(16), hex_rgb("404A57"), 4)
            base_y += max(44, used + 4)
        if idx < len(chapters) - 1:
            arrow_x = x + card_w + 7
            draw.rounded_rectangle((arrow_x, y + 200, arrow_x + 6, y + 260), radius=3, fill=rgba(colors["accent"], 160))
            draw.polygon(
                [(arrow_x + 30, y + 230), (arrow_x + 5, y + 210), (arrow_x + 5, y + 250)],
                fill=rgba(colors["accent"], 200),
            )
        x += card_w + gap

    out_path = REPORT_IMG_DIR / "book_overview.png"
    out_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(out_path)


def render_common_methods(book: dict) -> None:
    image = Image.new("RGBA", (1600, 960), rgba("F4F8FB"))
    draw = ImageDraw.Draw(image)
    draw.text((90, 70), "全书反复出现的五个方法", font=font(42, bold=True), fill=hex_rgb("183446"))
    draw.text((90, 130), "这五个方法会贯穿阅读报告、关键总结和 PPT。", font=font(24), fill=hex_rgb("546A79"))
    y = 210
    for idx, method in enumerate(book["cross_methods"], start=1):
        card(draw, (90, y, 1510, y + 125), rgba("FFFFFF"), outline=hex_rgb("9AB6C5"), radius=26, width=2)
        draw.text((118, y + 26), f"{idx}. {method['title']}", font=font(25, bold=True), fill=hex_rgb("1B4D66"))
        draw_paragraph(draw, method["body"], 120, y + 66, 1360, font(20), hex_rgb("41515D"), 6)
        y += 145
    REPORT_IMG_DIR.mkdir(parents=True, exist_ok=True)
    image.save(REPORT_IMG_DIR / "common_methods.png")


def render_chapter_core_map(chapter: dict, target_dir: Path, suffix: str) -> None:
    colors = chapter["palette"]
    image = Image.new("RGBA", (1600, 900), rgba(colors["bg"]))
    draw = ImageDraw.Draw(image)
    draw.text((70, 55), f"第{chapter['number']}章核心地图", font=font(40, bold=True), fill=hex_rgb(colors["primary"]))
    draw.text((70, 112), chapter["title"], font=font(32, bold=True), fill=hex_rgb(colors["dark"]))
    card(draw, (70, 170, 1530, 285), rgba("FFFFFF", 220), outline=hex_rgb(colors["secondary"]), radius=28, width=2)
    draw.text((100, 192), "一句话抓住本章", font=font(22, bold=True), fill=hex_rgb(colors["accent"]))
    draw_paragraph(draw, chapter["core_message"], 100, 228, 1380, font(24), hex_rgb("3E4650"), 6)

    card_w = 460
    card_h = 170
    positions = [
        (70, 335),
        (570, 335),
        (1070, 335),
        (70, 540),
        (570, 540),
        (1070, 540),
    ]
    for (box_x, box_y), concept in zip(positions, chapter["key_concepts"][:6]):
        card(draw, (box_x, box_y, box_x + card_w, box_y + card_h), rgba("FFFFFF", 230), outline=hex_rgb(colors["secondary"]), radius=24, width=2)
        draw.text((box_x + 24, box_y + 22), concept["name"], font=font(24, bold=True), fill=hex_rgb(colors["primary"]))
        draw_paragraph(draw, concept["explain"], box_x + 24, box_y + 62, card_w - 48, font(18), hex_rgb("485360"), 5)

    draw.text((70, 765), "给孩子的例子：", font=font(22, bold=True), fill=hex_rgb(colors["accent"]))
    draw_paragraph(draw, chapter["child_example"], 220, 760, 1290, font(19), hex_rgb("39424E"), 5)

    target_dir.mkdir(parents=True, exist_ok=True)
    image.save(target_dir / f"{chapter['id']}_{suffix}.png")


def render_summary_card(chapter: dict) -> None:
    colors = chapter["palette"]
    image = Image.new("RGBA", (1500, 860), rgba("FFFFFF"))
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, 1500, 110), fill=rgba(colors["primary"]))
    draw.text((65, 28), f"第{chapter['number']}章知识记忆卡", font=font(36, bold=True), fill=hex_rgb("FFFFFF"))
    draw.text((65, 128), chapter["title"], font=font(30, bold=True), fill=hex_rgb(colors["dark"]))
    draw.text((65, 178), "必须记住的 5 个点", font=font(22, bold=True), fill=hex_rgb(colors["accent"]))
    x_positions = [65, 775]
    y_positions = [230, 395, 560]
    points = chapter["must_remember"][:5]
    coords = [(x_positions[0], y_positions[0]), (x_positions[1], y_positions[0]), (x_positions[0], y_positions[1]), (x_positions[1], y_positions[1]), (x_positions[0], y_positions[2])]
    for idx, (point, (x, y)) in enumerate(zip(points, coords), start=1):
        card(draw, (x, y, x + 650, y + 130), rgba(chapter["palette"]["bg"]), outline=hex_rgb(colors["secondary"]), radius=22, width=2)
        draw.text((x + 22, y + 18), f"{idx}.", font=font(22, bold=True), fill=hex_rgb(colors["primary"]))
        draw_paragraph(draw, point, x + 70, y + 16, 550, font(21), hex_rgb("404B57"), 5)
    card(draw, (65, 725, 1435, 815), rgba("F7F7F7"), outline=hex_rgb(colors["secondary"]), radius=18, width=2)
    draw.text((92, 748), "一句话复盘", font=font(20, bold=True), fill=hex_rgb(colors["primary"]))
    draw_paragraph(draw, chapter["one_line_review"], 240, 744, 1160, font(20), hex_rgb("404B57"), 5)
    SUMMARY_IMG_DIR.mkdir(parents=True, exist_ok=True)
    image.save(SUMMARY_IMG_DIR / f"{chapter['id']}_memory_card.png")


def render_ppt_poster(chapter: dict) -> None:
    colors = chapter["palette"]
    image = Image.new("RGBA", (1600, 900), rgba(colors["dark"]))
    draw = ImageDraw.Draw(image)
    card(draw, (55, 55, 1545, 845), rgba(colors["bg"], 245), outline=hex_rgb(colors["secondary"]), radius=36, width=3)
    draw.text((110, 95), f"第{chapter['number']}章", font=font(28, bold=True), fill=hex_rgb(colors["primary"]))
    draw.text((110, 140), chapter["title"], font=font(42, bold=True), fill=hex_rgb(colors["dark"]))
    draw.text((110, 205), "本章一句话", font=font(22, bold=True), fill=hex_rgb(colors["accent"]))
    draw_paragraph(draw, chapter["core_message"], 110, 242, 1360, font(25), hex_rgb("3C4652"), 6)
    draw.text((110, 350), "三个关键抓手", font=font(22, bold=True), fill=hex_rgb(colors["accent"]))
    y = 390
    for concept in chapter["key_concepts"][:3]:
        card(draw, (110, y, 1480, y + 95), rgba("FFFFFF"), outline=hex_rgb(colors["secondary"]), radius=18, width=2)
        draw.text((138, y + 18), concept["name"], font=font(22, bold=True), fill=hex_rgb(colors["primary"]))
        draw_paragraph(draw, concept["explain"], 400, y + 18, 1040, font(20), hex_rgb("414B57"), 5)
        y += 110
    draw.text((110, 735), "孩子能立刻练的动作", font=font(22, bold=True), fill=hex_rgb(colors["accent"]))
    action_text = " / ".join(chapter["child_actions"][:3])
    draw_paragraph(draw, action_text, 110, 772, 1370, font(20), hex_rgb("404A56"), 5)
    PPT_IMG_DIR.mkdir(parents=True, exist_ok=True)
    image.save(PPT_IMG_DIR / f"{chapter['id']}_poster.png")


def render_ch04_lecture_map(chapter: dict, target_dir: Path, name: str) -> None:
    colors = chapter["palette"]
    image = Image.new("RGBA", (1700, 1000), rgba(colors["bg"]))
    draw = ImageDraw.Draw(image)
    draw.text((80, 55), "第四章十一讲总地图", font=font(40, bold=True), fill=hex_rgb(colors["primary"]))
    draw.text((80, 112), "先看十一讲的入口，再去记每一讲的重点。", font=font(24), fill=hex_rgb("56616D"))
    x_positions = [80, 570, 1060]
    y_positions = [180, 340, 500, 660]
    entries = chapter["lecture_index"]
    for idx, item in enumerate(entries):
        col = idx % 3
        row = idx // 3
        x = x_positions[col]
        y = y_positions[row]
        card(draw, (x, y, x + 430, y + 128), rgba("FFFFFF"), outline=hex_rgb(colors["secondary"]), radius=24, width=2)
        draw.text((x + 22, y + 18), f"{idx + 1}.", font=font(22, bold=True), fill=hex_rgb(colors["primary"]))
        draw_paragraph(draw, item["title"], x + 64, y + 16, 340, font(18), hex_rgb(colors["dark"]), 4)
        draw_paragraph(draw, item["core"], x + 24, y + 72, 380, font(16), hex_rgb("44505D"), 4)
    target_dir.mkdir(parents=True, exist_ok=True)
    image.save(target_dir / name)


def write_teaching_table(chapter: dict) -> None:
    lines = [
        f"# {chapter['full_title']} 教学提炼表",
        "",
        f"- 章节范围：{chapter['page_range']}",
        f"- 本章一句话：{chapter['core_message']}",
        f"- 关键问题：{chapter['children_question']}",
        "",
        "## 章节主题",
        chapter["position_in_book"],
        "",
        "## 核心概念",
    ]
    for concept in chapter["key_concepts"]:
        lines.append(f"- {concept['name']}：{concept['explain']}")
    lines.extend(["", "## 典型故事或案例"])
    for story in chapter["stories"]:
        lines.append(f"- {story['title']}：{story['summary']} 启示：{story['lesson']}")
    lines.extend(["", "## 常见误解"])
    for item in chapter["misunderstandings"]:
        lines.append(f"- {item}")
    lines.extend(["", "## 孩子能做的行动建议"])
    for action in chapter["child_actions"]:
        lines.append(f"- {action}")
    (INTERMEDIATE_DIR / f"{chapter['id']}_教学提炼表.md").write_text("\n".join(lines), encoding="utf-8")


def report_image_md(filename: str) -> str:
    return f"../04_配图/阅读报告/{filename}"


def summary_image_md(filename: str) -> str:
    return f"../04_配图/关键知识总结/{filename}"


def source_image_md(filename: str) -> str:
    return f"../04_配图/原书页图/{filename}"


def build_report(book: dict, chapters: list[dict]) -> str:
    lines: list[str] = [
        f"# {book['title']} 阅读报告（儿童可读版）",
        "",
        f"> 适读定位：{book['reader_positioning']}",
        "",
        "![全书结构图](../04_配图/阅读报告/book_overview.png)",
        "",
        "## 摘要",
    ]
    for paragraph in book["abstract"]:
        lines.extend([paragraph, ""])
    lines.extend(
        [
            "## 前置理解：序言、导读与附录为什么重要",
            "![全书方法图](../04_配图/阅读报告/common_methods.png)",
            "",
        ]
    )
    for paragraph in book["front_matter_notes"]:
        lines.extend([paragraph, ""])
    lines.extend(["## 全书的三条主论点", ""])
    for idx, item in enumerate(book["main_thesis"], start=1):
        lines.append(f"{idx}. {item}")
    lines.extend(["", "## 章节精读", ""])

    for chapter in chapters:
        lines.extend(
            [
                f"### {chapter['full_title']}",
                "",
                f"![{chapter['title']}章首页]({source_image_md(f'{chapter['id']}_opener_page.png')})",
                "",
                f"![{chapter['title']}核心地图]({report_image_md(f'{chapter['id']}_core_map.png')})",
                "",
                f"本章一句话：{chapter['core_message']}",
                "",
                f"为什么重要：{chapter['position_in_book']}",
                "",
            ]
        )
        for paragraph in chapter["report_paragraphs"]:
            lines.extend([paragraph, ""])
        lines.append("本章最该抓住的关键概念：")
        lines.append("")
        for concept in chapter["key_concepts"]:
            lines.append(f"- {concept['name']}：{concept['explain']}")
        lines.extend(["", "本章里的代表性故事："])
        lines.append("")
        for story in chapter["stories"]:
            lines.append(f"- {story['title']}：{story['summary']} 启示：{story['lesson']}")
        lines.extend(["", "孩子读完这一章能收获什么：", ""])
        for gain in chapter["chapter_gains"]:
            lines.append(f"- {gain}")
        lines.extend(["", "容易误读的地方：", ""])
        for item in chapter["misunderstandings"]:
            lines.append(f"- {item}")
        lines.extend(["", "可以立刻实践的小动作：", ""])
        for action in chapter["child_actions"]:
            lines.append(f"- {action}")
        lines.extend(["", f"一句话复盘：{chapter['one_line_review']}", ""])
        if chapter["id"] == "ch04":
            lines.extend(
                [
                    f"![第四章十一讲地图]({report_image_md('ch04_lecture_map.png')})",
                    "",
                    "第四章十一讲极简索引：",
                    "",
                ]
            )
            for item in chapter["lecture_index"]:
                lines.append(f"- {item['title']}：{item['core']}")
            lines.append("")
        if chapter["id"] == "ch05":
            lines.extend(["本章关注的现实议题样本：", ""])
            for item in chapter["article_examples"]:
                lines.append(f"- {item}")
            lines.append("")

    lines.extend(["## 跨章节共性方法", ""])
    for item in book["cross_methods"]:
        lines.append(f"### {item['title']}")
        lines.extend([item["body"], ""])
    lines.extend(["## 芒格式思维的现实应用", ""])
    for item in book["daily_applications"]:
        lines.append(f"- {item}")
    lines.extend(["", "## 局限与容易误读的地方", ""])
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


def build_summary(book: dict, chapters: list[dict]) -> str:
    lines: list[str] = [
        f"# {book['title']} 关键知识总结（按章节）",
        "",
        "这份总结只保留每章最该记住的骨架，帮助第一次接触本书的读者快速建立理解地图。",
        "",
    ]
    for chapter in chapters:
        lines.extend(
            [
                f"## {chapter['full_title']}",
                "",
                "### 这一章在讲什么",
                chapter["core_message"],
                "",
                f"![{chapter['title']}记忆卡]({summary_image_md(f'{chapter['id']}_memory_card.png')})",
                "",
                "### 必须记住的 5 个点",
            ]
        )
        for item in chapter["must_remember"]:
            lines.append(f"- {item}")
        lines.extend(["", "### 一个孩子也能懂的例子", chapter["child_example"], "", "### 一句话复盘", chapter["one_line_review"], "", "### 本章配图", f"![{chapter['title']}概念海报](../04_配图/PPT/{chapter['id']}_poster.png)", ""])
        if chapter["id"] == "ch04":
            lines.extend(["### 第四章十一讲极简索引", ""])
            for item in chapter["lecture_index"]:
                lines.append(f"- {item['title']}：{item['core']}")
            lines.extend(["", f"![第四章十一讲地图]({summary_image_md('ch04_lecture_map.png')})", ""])
    return "\n".join(lines)


def build_readme(book: dict, chapters: list[dict]) -> str:
    lines = [
        f"# {book['title']} 儿童可读版学习资料包",
        "",
        "## 目录说明",
        "- `00_项目说明/`：项目说明与生成脚本。",
        "- `01_阅读报告/`：总阅读报告 Markdown。",
        "- `02_关键知识总结/`：按章节整理的关键知识总结 Markdown。",
        "- `03_PPT/`：5 份主章 PPT 及对应 `.js` 源文件。",
        "- `04_配图/`：阅读报告、关键知识总结、PPT 与原书页图所用图片。",
        "- `05_中间素材/`：章节纯文本、教学提炼表、数据源与导出预览。",
        "",
        "## 当前采用的章节边界",
    ]
    for chapter in chapters:
        lines.append(f"- {chapter['full_title']}：{chapter['page_range']}")
    lines.extend(
        [
            "",
            "## 生成流程",
            "1. 运行 `extract_pdf_assets.py`，抽取章节文本和原书页图。",
            "2. 运行 `build_learning_pack.py`，生成教学提炼表、配图、Markdown 成品。",
            "3. 运行 `build_ppts.js`，生成 5 份 PPT 与对应 JS 源文件。",
            "4. 运行 `export_ppts_to_pdf.ps1` 与 `render_pdf_to_png.py`，导出 PDF/PNG 做版面检查。",
            "",
            "## 重建提示",
            "- PPT 生成依赖 `pptxgenjs`。",
            "- 预览导出依赖本机 PowerPoint COM。",
            "- 所有 Markdown 都只引用本项目内的相对路径图片。"
        ]
    )
    return "\n".join(lines)


def main() -> None:
    manifest = load_json(INTERMEDIATE_DIR / "chapter_manifest.json")
    content = load_json(INTERMEDIATE_DIR / "chapter_content.json")
    chapters = content["chapters"]

    REPORT_IMG_DIR.mkdir(parents=True, exist_ok=True)
    SUMMARY_IMG_DIR.mkdir(parents=True, exist_ok=True)
    PPT_IMG_DIR.mkdir(parents=True, exist_ok=True)

    render_overview(content["book"], chapters)
    render_common_methods(content["book"])
    for chapter in chapters:
        render_chapter_core_map(chapter, REPORT_IMG_DIR, "core_map")
        render_summary_card(chapter)
        render_ppt_poster(chapter)
        write_teaching_table(chapter)

    ch04 = next(ch for ch in chapters if ch["id"] == "ch04")
    render_ch04_lecture_map(ch04, REPORT_IMG_DIR, "ch04_lecture_map.png")
    render_ch04_lecture_map(ch04, SUMMARY_IMG_DIR, "ch04_lecture_map.png")
    render_ch04_lecture_map(ch04, PPT_IMG_DIR, "ch04_lecture_map.png")

    report_markdown = build_report(content["book"], chapters)
    summary_markdown = build_summary(content["book"], chapters)
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    SUMMARY_DIR.mkdir(parents=True, exist_ok=True)
    (REPORT_DIR / "穷查理宝典_阅读报告_儿童可读版.md").write_text(report_markdown, encoding="utf-8")
    (SUMMARY_DIR / "穷查理宝典_关键知识总结_按章节.md").write_text(summary_markdown, encoding="utf-8")
    (ROOT / "00_项目说明" / "README.md").write_text(build_readme(content["book"], chapters), encoding="utf-8")

    print("已生成：")
    print(f"- {REPORT_DIR / '穷查理宝典_阅读报告_儿童可读版.md'}")
    print(f"- {SUMMARY_DIR / '穷查理宝典_关键知识总结_按章节.md'}")
    print(f"- {REPORT_IMG_DIR}")
    print(f"- {SUMMARY_IMG_DIR}")
    print(f"- {PPT_IMG_DIR}")
    print(f"- {ROOT / '00_项目说明' / 'README.md'}")
    print(f"- {manifest['pdf_path']}")


if __name__ == "__main__":
    main()
