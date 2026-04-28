from __future__ import annotations

import argparse
import json
import shutil
from pathlib import Path

import fitz
from PIL import Image, ImageDraw, ImageFont

from markdown_pdf_renderer import MarkdownPdfRenderer
from teaching_pack_config import BOOKS, OUTPUT_ROOT


HELPER_SOURCE = Path(
    r"G:\OBSIDIAN\06_澄明之境__CHENG_MING_ZHI_JING\00_项目说明\生成脚本\pptxgenjs_helpers"
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="生成三本书的章节化教学交付包")
    parser.add_argument("--book-id", choices=sorted(BOOKS.keys()), help="只生成指定书籍")
    return parser.parse_args()


def ensure_font(preferred: str) -> str:
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


FONT_BOLD = ensure_font("msyhbd.ttc")
FONT_REGULAR = ensure_font("msyh.ttc")


def font(size: int, *, bold: bool = False) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(FONT_BOLD if bold else FONT_REGULAR, size=size)


def hex_rgb(value: str) -> tuple[int, int, int]:
    value = value.strip().lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


def rgba(value: str, alpha: int = 255) -> tuple[int, int, int, int]:
    r, g, b = hex_rgb(value)
    return (r, g, b, alpha)


def card(
    draw: ImageDraw.ImageDraw,
    xy: tuple[int, int, int, int],
    *,
    fill_color,
    outline=None,
    radius: int = 26,
    width: int = 2,
) -> None:
    draw.rounded_rectangle(xy, radius=radius, fill=fill_color, outline=outline, width=width)


def wrap_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    text_font: ImageFont.FreeTypeFont,
    max_width: int,
) -> list[str]:
    lines: list[str] = []
    for raw_line in text.splitlines():
        raw_line = raw_line.strip()
        if not raw_line:
            lines.append("")
            continue
        current = ""
        for char in raw_line:
            trial = current + char
            box = draw.textbbox((0, 0), trial, font=text_font)
            if box[2] - box[0] <= max_width or not current:
                current = trial
            else:
                lines.append(current)
                current = char
        if current:
            lines.append(current)
    return lines or [""]


def draw_paragraph(
    draw: ImageDraw.ImageDraw,
    text: str,
    *,
    x: int,
    y: int,
    max_width: int,
    text_font: ImageFont.FreeTypeFont,
    fill,
    line_gap: int = 6,
) -> int:
    start_y = y
    for line in wrap_text(draw, text, text_font, max_width):
        draw.text((x, y), line, font=text_font, fill=fill)
        box = draw.textbbox((x, y), line or "字", font=text_font)
        y += box[3] - box[1] + line_gap
    return y - start_y


def safe_name(value: str) -> str:
    cleaned = value
    for token in [":", "：", "/", "\\", "?", "？", "!", "！", "\"", "'", "“", "”", "《", "》", "（", "）", "(", ")", "—", "——", "·", ",", "，"]:
        cleaned = cleaned.replace(token, "_")
    cleaned = cleaned.replace(" ", "")
    while "__" in cleaned:
        cleaned = cleaned.replace("__", "_")
    return cleaned.strip("_")


def dedupe_line(line: str) -> str:
    line = line.strip()
    if not line:
        return ""
    if len(line) % 2 == 0:
        half = len(line) // 2
        if line[:half] == line[half:]:
            return line[:half]
    return line


def clean_page_text(text: str) -> str:
    cleaned_lines: list[str] = []
    previous = ""
    for raw_line in text.replace("\x00", "").splitlines():
        line = dedupe_line(raw_line)
        if not line:
            continue
        if line == previous:
            continue
        cleaned_lines.append(line)
        previous = line
    return "\n".join(cleaned_lines).strip()


def extract_text_range(doc: fitz.Document, start_page: int, end_page: int) -> str:
    chunks: list[str] = []
    for page_no in range(start_page, end_page + 1):
        text = clean_page_text(doc[page_no - 1].get_text("text"))
        chunks.append(f"===== PAGE {page_no} =====\n{text}\n")
    return "\n".join(chunks)


def render_page(doc: fitz.Document, page_no: int, out_path: Path, *, scale: float = 1.45) -> None:
    page = doc[page_no - 1]
    pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    pix.save(out_path)


def book_defaults(book_id: str) -> dict:
    if book_id == "antifragile":
        return {
            "book_concept": "结构设计",
            "book_concept_explain": "这一章真正要训练的，不是预测剧情的能力，而是重新设计自己与风险和波动的关系。",
            "reader_question_prefix": "在充满波动和不确定的世界里，围绕这一章的主题，我们到底该重新看清什么结构问题？",
            "generic_example": "在投资与职业选择上，这意味着别把全部希望寄托在唯一预测上，而要先把结构安排得不容易被一击打穿。",
        }
    if book_id == "black_swan":
        return {
            "book_concept": "尾部意识",
            "book_concept_explain": "这一章不断提醒我们：真正重要的事情，常常不是平均值附近的小波动，而是少数极端、跳变、会改写结果的事件。",
            "reader_question_prefix": "如果世界经常被少数极端事件突然改写，这一章想纠正我们的哪种思维错觉？",
            "generic_example": "在职业、舆论和市场判断上，这意味着不要把昨天的平稳直接当成明天的保证。",
        }
    return {
        "book_concept": "经济学镜头",
        "book_concept_explain": "这一章的意义在于让读者多拿到一副看现实的镜头，从而更清楚地理解权衡、激励、制度和后果。",
        "reader_question_prefix": "围绕这一章的人物或争论，它为什么直到今天还值得我们反复回看？",
        "generic_example": "在看新闻和公共政策时，这意味着不要只急着站队，而要先问：代价是什么、激励是什么、制度会把事情推向哪里。",
    }


def split_title_tokens(title: str) -> list[str]:
    current = title
    for token in ["：", ":", "——", "—", "与", "和", "、", "（", "）", "(", ")", "“", "”"]:
        current = current.replace(token, "|")
    parts = [part.strip(" _-") for part in current.split("|")]
    unique: list[str] = []
    for part in parts:
        if not part or len(part) <= 1:
            continue
        if part not in unique:
            unique.append(part)
    return unique[:3]


def chapter_concepts(book_id: str, chapter: dict, focus: str, application: str) -> list[dict]:
    defaults = book_defaults(book_id)
    tokens = split_title_tokens(chapter["title"])
    concepts: list[dict] = []
    if tokens:
        concepts.append({"name": tokens[0], "explain": focus})
    if len(tokens) > 1:
        concepts.append({"name": tokens[1], "explain": application})
    else:
        concepts.append({"name": "现实含义", "explain": application})
    concepts.append({"name": defaults["book_concept"], "explain": defaults["book_concept_explain"]})
    return concepts[:3]


def chapter_question(book_id: str, chapter: dict) -> str:
    defaults = book_defaults(book_id)
    if book_id == "echo_of_genius":
        return f"{defaults['reader_question_prefix']} 换句话说，这一章到底想让零基础读者明白：{chapter['title']} 为什么不是历史旧闻，而是今天依然活着的问题。"
    return f"{defaults['reader_question_prefix']} 这章借“{chapter['title']}”这个入口，希望读者改变哪一种直觉？"


def chapter_takeaways(book_id: str, chapter: dict) -> list[str]:
    if book_id == "antifragile":
        return [
            "先看结构是否脆弱，再决定要不要上场。",
            "给自己保留后手和试错位，不要把所有资源押在单一判断上。",
            "允许小波动、小错误存在，用可承受的代价换长期韧性。",
        ]
    if book_id == "black_swan":
        return [
            "别把过去的平稳自动翻译成未来的安全。",
            "对讲得太圆满的解释保持警惕，极端事件经常发生在故事之外。",
            "与其追求精确预测，不如先减少致命暴露、争取正面意外。",
        ]
    return [
        "看任何现实问题时，先承认取舍和代价客观存在。",
        "别只看口号和立场，多问激励、制度和后果。",
        "把不同经济学家的镜头当成工具箱，而不是当成唯一真理。",
    ]


def chapter_misunderstandings(book_id: str, chapter: dict) -> list[str]:
    if book_id == "antifragile":
        return [
            "把反脆弱误解成单纯喜欢风险，仿佛波动越大越好。",
            "只记住概念姿态，却忘了真正重要的是结构安排和风险边界。",
            "把允许小试错误读成放弃纪律、放弃复盘。",
        ]
    if book_id == "black_swan":
        return [
            "把黑天鹅理解成“什么都能发生”，于是放弃任何分析。",
            "以为只要多收集一些信息，就能把极端事件完全装进模型。",
            "把事后解释的顺滑感，误当成事前就真的可以预测。",
        ]
    return [
        "把这位经济学家的观点当成放之四海而皆准的唯一答案。",
        "只记人物故事，不去理解这套观点到底在回答什么现实难题。",
        "把理论口号背下来，却不追问它在今天的适用边界。",
    ]


def chapter_real_examples(book_id: str, chapter: dict) -> list[str]:
    defaults = book_defaults(book_id)
    return [chapter["application"], defaults["generic_example"]]


def chapter_review_sentence(chapter: dict) -> str:
    focus = chapter["focus"].strip()
    return focus if focus.endswith("。") else focus + "。"


def chapter_report_paragraphs(book: dict, chapter: dict) -> list[str]:
    return [
        f"这一章抓住“{chapter['title']}”这个入口，想让读者看清的一件事是：{chapter['focus']}",
        f"如果把它放回全书，这一章承担的任务是把“{book['title']}”的主问题进一步落地：别只停留在概念表面，而要学会真正用结构和后果来判断事情。",
        f"对零基础读者来说，本章最有用的不是记住更多名词，而是把这条提醒变成动作：{chapter['application']}",
    ]


def chapter_content(book_id: str, book: dict, chapter: dict) -> dict:
    concepts = chapter_concepts(book_id, chapter, chapter["focus"], chapter["application"])
    must_remember = [
        chapter["focus"],
        f"通俗地说：{chapter['analogy']}",
        f"放回现实：{chapter['application']}",
        f"真正要防的是把“{chapter['title']}”只当成术语，而忘了它在提醒你重新安排决策结构。",
        f"一句话记住：{chapter_review_sentence(chapter)}",
    ]
    return {
        **chapter,
        "palette": book["theme"],
        "reader_question": chapter_question(book_id, chapter),
        "core_message": chapter["focus"],
        "position_in_book": f"它位于 {chapter['source_part']}，负责把全书的主问题推进到“{chapter['title']}”这个具体角度。",
        "key_concepts": concepts,
        "must_remember": must_remember,
        "plain_example": chapter["analogy"],
        "real_examples": chapter_real_examples(book_id, chapter),
        "misunderstandings": chapter_misunderstandings(book_id, chapter),
        "action_takeaways": chapter_takeaways(book_id, chapter),
        "one_line_review": chapter_review_sentence(chapter),
        "report_paragraphs": chapter_report_paragraphs(book, chapter),
        "study_questions": [
            f"如果只用一句白话解释“{chapter['title']}”，你会怎么说？",
            f"在你的生活、工作或投资判断里，哪里最能看到这一章讨论的问题？",
            "根据这一章的提醒，你最想马上改掉的一条决策习惯是什么？",
        ],
    }


def report_image(filename: str) -> str:
    return f"../40_配图/阅读报告/{filename}"


def summary_image(filename: str) -> str:
    return f"../40_配图/关键知识总结/{filename}"


def ppt_image(filename: str) -> str:
    return f"../40_配图/PPT/{filename}"


def source_image(filename: str) -> str:
    return f"../40_配图/原书页图/{filename}"


def render_book_overview(book: dict, chapters: list[dict], out_path: Path) -> None:
    image = Image.new("RGBA", (1700, 1000), rgba(book["theme"]["bg"]))
    draw = ImageDraw.Draw(image)
    draw.text((80, 60), f"{book['title']} 全书路线图", font=font(42, bold=True), fill=hex_rgb(book["theme"]["dark"]))
    draw.text((80, 118), "先知道这本书是怎样推进问题的，再进入每一章。", font=font(22), fill=hex_rgb("475569"))
    cards = [
        {"title": "阅读定位", "body": book["reader_positioning"]},
        *book["structure_cards"],
        {"title": "章节数量", "body": f"本次交付把全书拆成 {len(chapters)} 个编号章节，每章独立生成一份零基础教学版 PPT。"},
    ]
    positions = [(80, 190), (580, 190), (1080, 190), (80, 510), (580, 510), (1080, 510)]
    for (x, y), item in zip(positions, cards[:6]):
        card(draw, (x, y, x + 420, y + 240), fill_color=rgba("FFFFFF"), outline=hex_rgb(book["theme"]["secondary"]))
        draw.text((x + 24, y + 22), item["title"], font=font(28, bold=True), fill=hex_rgb(book["theme"]["primary"]))
        draw_paragraph(draw, item["body"], x=x + 24, y=y + 74, max_width=372, text_font=font(20), fill=hex_rgb("334155"))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(out_path)


def render_common_methods(book: dict, out_path: Path) -> None:
    image = Image.new("RGBA", (1700, 1050), rgba("FFFFFF"))
    draw = ImageDraw.Draw(image)
    draw.text((80, 60), "全书反复出现的关键理解方法", font=font(40, bold=True), fill=hex_rgb(book["theme"]["primary"]))
    draw.text((80, 118), "如果只记方法，不记枝节，这几条最值得带走。", font=font(22), fill=hex_rgb("475569"))
    y = 190
    for idx, item in enumerate(book["cross_methods"], start=1):
        card(draw, (80, y, 1620, y + 132), fill_color=rgba(book["theme"]["bg"]), outline=hex_rgb(book["theme"]["secondary"]))
        draw.text((108, y + 22), f"{idx}. {item['title']}", font=font(25, bold=True), fill=hex_rgb(book["theme"]["primary"]))
        draw_paragraph(draw, item["body"], x=110, y=y + 62, max_width=1460, text_font=font(20), fill=hex_rgb("334155"))
        y += 152
    out_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(out_path)


def render_chapter_core_map(chapter: dict, out_path: Path) -> None:
    colors = chapter["palette"]
    image = Image.new("RGBA", (1700, 1000), rgba(colors["bg"]))
    draw = ImageDraw.Draw(image)
    draw.text((70, 46), f"{chapter['full_title']} 核心地图", font=font(38, bold=True), fill=hex_rgb(colors["primary"]))
    draw.text((70, 102), chapter["core_message"], font=font(22, bold=True), fill=hex_rgb(colors["dark"]))
    positions = [(70, 190), (580, 190), (1090, 190)]
    for (x, y), item in zip(positions, chapter["key_concepts"]):
        card(draw, (x, y, x + 440, y + 240), fill_color=rgba("FFFFFF"), outline=hex_rgb(colors["secondary"]))
        draw.text((x + 24, y + 20), item["name"], font=font(26, bold=True), fill=hex_rgb(colors["primary"]))
        draw_paragraph(draw, item["explain"], x=x + 24, y=y + 74, max_width=390, text_font=font(20), fill=hex_rgb("334155"))
    card(draw, (70, 500, 1620, 780), fill_color=rgba("FFFFFF"), outline=hex_rgb(colors["secondary"]))
    draw.text((96, 530), "通俗比喻", font=font(24, bold=True), fill=hex_rgb(colors["accent"]))
    draw_paragraph(draw, chapter["plain_example"], x=96, y=578, max_width=1460, text_font=font(22), fill=hex_rgb("334155"))
    draw.text((96, 850), "一页带走", font=font(22, bold=True), fill=hex_rgb(colors["primary"]))
    draw_paragraph(draw, chapter["one_line_review"], x=240, y=846, max_width=1290, text_font=font(20), fill=hex_rgb("334155"))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(out_path)


def render_summary_card(chapter: dict, out_path: Path) -> None:
    colors = chapter["palette"]
    image = Image.new("RGBA", (1600, 930), rgba("FFFFFF"))
    draw = ImageDraw.Draw(image)
    draw.rectangle((0, 0, 1600, 120), fill=rgba(colors["primary"]))
    draw.text((70, 34), f"{chapter['full_title']} 记忆卡", font=font(36, bold=True), fill=hex_rgb("FFFFFF"))
    coords = [(70, 170), (810, 170), (70, 340), (810, 340), (70, 510)]
    for index, ((x, y), bullet) in enumerate(zip(coords, chapter["must_remember"][:5]), start=1):
        card(draw, (x, y, x + 700, y + 130), fill_color=rgba(colors["bg"]), outline=hex_rgb(colors["secondary"]))
        draw.text((x + 22, y + 18), f"{index}.", font=font(22, bold=True), fill=hex_rgb(colors["primary"]))
        draw_paragraph(draw, bullet, x=x + 72, y=y + 16, max_width=600, text_font=font(21), fill=hex_rgb("334155"))
    card(draw, (70, 700, 1530, 840), fill_color=rgba("F8FAFC"), outline=hex_rgb(colors["secondary"]))
    draw.text((94, 726), "一句话复盘", font=font(22, bold=True), fill=hex_rgb(colors["primary"]))
    draw_paragraph(draw, chapter["one_line_review"], x=250, y=722, max_width=1240, text_font=font(22), fill=hex_rgb("334155"))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(out_path)


def render_poster(chapter: dict, out_path: Path) -> None:
    colors = chapter["palette"]
    image = Image.new("RGBA", (1600, 900), rgba(colors["dark"]))
    draw = ImageDraw.Draw(image)
    card(draw, (60, 60, 1540, 840), fill_color=rgba(colors["bg"], 245), outline=hex_rgb(colors["secondary"]))
    draw.text((110, 96), chapter["full_title"], font=font(40, bold=True), fill=hex_rgb(colors["dark"]))
    draw.text((110, 158), chapter["core_message"], font=font(23, bold=True), fill=hex_rgb(colors["primary"]))
    y = 270
    for item in chapter["key_concepts"]:
        card(draw, (110, y, 1490, y + 118), fill_color=rgba("FFFFFF"), outline=hex_rgb(colors["secondary"]))
        draw.text((136, y + 18), item["name"], font=font(24, bold=True), fill=hex_rgb(colors["primary"]))
        draw_paragraph(draw, item["explain"], x=360, y=y + 18, max_width=1070, text_font=font(20), fill=hex_rgb("334155"))
        y += 138
    draw.text((110, 760), "马上可用", font=font(22, bold=True), fill=hex_rgb(colors["accent"]))
    draw_paragraph(draw, " / ".join(chapter["action_takeaways"]), x=220, y=756, max_width=1180, text_font=font(19), fill=hex_rgb("334155"))
    out_path.parent.mkdir(parents=True, exist_ok=True)
    image.save(out_path)


def build_report(book: dict, chapters: list[dict]) -> str:
    lines = [
        f"# {book['title']} 阅读报告",
        "",
        f"> 适读定位：{book['reader_positioning']}",
        "",
        f"![全书路线图]({report_image('book_overview.png')})",
        "",
        "## 摘要",
        "",
    ]
    for paragraph in book["abstract"]:
        lines.extend([paragraph, ""])
    lines.extend(["## 这本书到底在回答什么问题", ""])
    for index, item in enumerate(book["main_thesis"], start=1):
        lines.append(f"{index}. {item}")
    lines.extend(["", "## 全书推进逻辑", "", f"![全书方法图]({report_image('common_methods.png')})", ""])
    for card_item in book["structure_cards"]:
        lines.append(f"- {card_item['title']}：{card_item['body']}")
    lines.extend(["", "## 按章精读", ""])
    for chapter in chapters:
        lines.extend(
            [
                f"### {chapter['full_title']}",
                "",
                f"![原书起始页]({source_image(f'{chapter['id']}_opener_page.png')})",
                "",
                f"![核心地图]({report_image(f'{chapter['id']}_core_map.png')})",
                "",
                "#### 这一章在问什么",
                chapter["reader_question"],
                "",
            ]
        )
        for paragraph in chapter["report_paragraphs"]:
            lines.extend([paragraph, ""])
        lines.extend(["#### 三个关键概念", ""])
        for item in chapter["key_concepts"]:
            lines.append(f"- {item['name']}：{item['explain']}")
        lines.extend(["", "#### 两个现实例子", ""])
        for item in chapter["real_examples"]:
            lines.append(f"- {item}")
        lines.extend(["", "#### 容易误解的地方", ""])
        for item in chapter["misunderstandings"]:
            lines.append(f"- {item}")
        lines.extend(["", "#### 真正可以带走的三条收获", ""])
        for item in chapter["action_takeaways"]:
            lines.append(f"- {item}")
        lines.extend(["", "#### 一句话复盘", chapter["one_line_review"], ""])
    lines.extend(["## 现实意义", ""])
    for item in book["daily_applications"]:
        lines.append(f"- {item}")
    lines.extend(["", "## 可能争议与边界", ""])
    for item in book["limitations"]:
        lines.append(f"- {item}")
    lines.extend(["", "## 结论", ""])
    for item in book["conclusion"]:
        lines.extend([item, ""])
    return "\n".join(lines)


def build_summary(book: dict, chapters: list[dict]) -> str:
    lines = [
        f"# {book['title']} 关键知识总结（按章节）",
        "",
        "这份总结只保留每章最值得带走的骨架，适合第一次接触本书的人快速建立理解地图。",
        "",
        f"![全书路线图]({report_image('book_overview.png')})",
        "",
    ]
    for chapter in chapters:
        lines.extend(
            [
                f"## {chapter['full_title']}",
                "",
                "### 核心结论",
                chapter["core_message"],
                "",
                f"![记忆卡]({summary_image(f'{chapter['id']}_memory_card.png')})",
                "",
                "### 三个关键词",
            ]
        )
        for item in chapter["key_concepts"]:
            lines.append(f"- {item['name']}：{item['explain']}")
        lines.extend(
            [
                "",
                "### 一个通俗比喻",
                chapter["plain_example"],
                "",
                "### 一个现实启发",
                chapter["real_examples"][0],
                "",
                "### 一句记忆句",
                chapter["one_line_review"],
                "",
                f"![章节海报]({ppt_image(f'{chapter['id']}_poster.png')})",
                "",
            ]
        )
    return "\n".join(lines)


def manifest_for(book: dict, doc: fitz.Document) -> dict:
    return {
        "book_id": book["book_id"],
        "title": book["title"],
        "source_pdf": str(book["source_pdf"]),
        "page_count": doc.page_count,
        "chapters": [
            {
                "id": chapter["id"],
                "sequence_no": chapter["sequence_no"],
                "source_part": chapter["source_part"],
                "original_label": chapter["original_label"],
                "title": chapter["title"],
                "start_page": chapter["start_page"],
                "end_page": chapter["end_page"],
            }
            for chapter in book["chapters"]
        ],
    }


def write_json(path: Path, data: dict | list) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2, default=str), encoding="utf-8")


def ensure_structure(project_root: Path) -> dict[str, Path]:
    paths = {
        "original": project_root / "00_原书",
        "ppt": project_root / "10_PPT",
        "report": project_root / "20_阅读报告",
        "summary": project_root / "30_关键知识总结",
        "assets_root": project_root / "40_配图",
        "assets_report": project_root / "40_配图" / "阅读报告",
        "assets_summary": project_root / "40_配图" / "关键知识总结",
        "assets_ppt": project_root / "40_配图" / "PPT",
        "assets_source": project_root / "40_配图" / "原书页图",
        "build": project_root / "_build",
        "build_scripts": project_root / "_build" / "scripts",
        "build_text": project_root / "_build" / "章节原文",
        "build_preview_pdf": project_root / "_build" / "预览_PPT_PDF",
        "build_preview_png": project_root / "_build" / "预览_PPT_PNG",
    }
    for path in paths.values():
        path.mkdir(parents=True, exist_ok=True)
    return paths


def write_wrapper_scripts(book: dict, project_root: Path) -> None:
    scripts_dir = project_root / "_build" / "scripts"
    scripts_dir.mkdir(parents=True, exist_ok=True)
    helper_target = scripts_dir / "pptxgenjs_helpers"
    if HELPER_SOURCE.exists() and not helper_target.exists():
        shutil.copytree(HELPER_SOURCE, helper_target)

    (scripts_dir / "build_pack.py").write_text(
        "\n".join(
            [
                "from pathlib import Path",
                "import sys",
                "",
                "BOOK_ROOT = Path(__file__).resolve().parents[2]",
                "REPO = BOOK_ROOT.parent",
                "sys.path.insert(0, str(REPO / 'scripts'))",
                "",
                "from teaching_pack_builder import build_one",
                "",
                "if __name__ == '__main__':",
                f"    build_one('{book['book_id']}')",
                "",
            ]
        ),
        encoding="utf-8",
    )

    (scripts_dir / "build_ppts.js").write_text(
        "\n".join(
            [
                "\"use strict\";",
                "",
                "const path = require(\"path\");",
                "const { buildBook } = require(path.resolve(__dirname, \"../../../scripts/teaching_pack_ppt_builder.js\"));",
                "",
                "buildBook(path.resolve(__dirname, \"../..\")).catch((error) => {",
                "  console.error(error);",
                "  process.exit(1);",
                "});",
                "",
            ]
        ),
        encoding="utf-8",
    )

    (scripts_dir / "export_ppts_to_pdf.ps1").write_text(
        "\n".join(
            [
                "$ErrorActionPreference = 'Stop'",
                "$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)",
                "$pptDir = Join-Path $root '10_PPT'",
                "$outDir = Join-Path $root '_build\\预览_PPT_PDF'",
                "New-Item -ItemType Directory -Force -Path $outDir | Out-Null",
                "$pptFiles = Get-ChildItem -LiteralPath $pptDir -Filter *.pptx | Sort-Object Name",
                "if (-not $pptFiles) { throw \"未在 $pptDir 找到 PPTX 文件。\" }",
                "$powerPoint = New-Object -ComObject PowerPoint.Application",
                "try {",
                "  foreach ($file in $pptFiles) {",
                "    $pdfPath = Join-Path $outDir ($file.BaseName + '.pdf')",
                "    Write-Host \"Exporting $($file.Name) -> $pdfPath\"",
                "    $presentation = $powerPoint.Presentations.Open($file.FullName, $false, $false, $false)",
                "    try { $presentation.SaveAs($pdfPath, 32) }",
                "    finally { $presentation.Close() }",
                "  }",
                "}",
                "finally { $powerPoint.Quit() }",
                "",
            ]
        ),
        encoding="utf-8",
    )

    (scripts_dir / "render_pdf_to_png.py").write_text(
        "\n".join(
            [
                "from pathlib import Path",
                "import fitz",
                "",
                "ROOT = Path(__file__).resolve().parents[2]",
                "PDF_DIR = ROOT / '_build' / '预览_PPT_PDF'",
                "PNG_DIR = ROOT / '_build' / '预览_PPT_PNG'",
                "",
                "if __name__ == '__main__':",
                "    for pdf_path in sorted(PDF_DIR.glob('*.pdf')):",
                "        out_dir = PNG_DIR / pdf_path.stem",
                "        out_dir.mkdir(parents=True, exist_ok=True)",
                "        doc = fitz.open(pdf_path)",
                "        for page_index, page in enumerate(doc, start=1):",
                "            pix = page.get_pixmap(matrix=fitz.Matrix(1.45, 1.45), alpha=False)",
                "            pix.save(out_dir / f'slide_{page_index:02d}.png')",
                "",
            ]
        ),
        encoding="utf-8",
    )

    (project_root / "_build" / "README.md").write_text(
        "\n".join(
            [
                f"# {book['title']} 构建说明",
                "",
                "1. `python scripts/build_pack.py`：生成章节原文、结构化 JSON、配图、Markdown 和 PDF。",
                "2. `node scripts/build_ppts.js`：生成全部章节 PPTX。",
                "3. `powershell -ExecutionPolicy Bypass -File scripts/export_ppts_to_pdf.ps1`：把 PPTX 导出为预览 PDF。",
                "4. `python scripts/render_pdf_to_png.py`：把预览 PDF 转成逐页 PNG。",
                "",
            ]
        ),
        encoding="utf-8",
    )


def render_markdown_pdfs(book: dict, project_root: Path) -> None:
    renderer = MarkdownPdfRenderer()
    report_md = project_root / "20_阅读报告" / book["report_filename"]
    report_pdf = project_root / "20_阅读报告" / book["report_pdf_filename"]
    summary_md = project_root / "30_关键知识总结" / book["summary_filename"]
    summary_pdf = project_root / "30_关键知识总结" / book["summary_pdf_filename"]
    renderer.render(
        sources=[report_md],
        output_path=report_pdf,
        title=f"{book['title']} 阅读报告",
        subtitle=f"{book['title']} 零基础教学版阅读报告",
        asset_base_dir=project_root,
    )
    renderer.render(
        sources=[summary_md],
        output_path=summary_pdf,
        title=f"{book['title']} 关键知识总结",
        subtitle=f"{book['title']} 按章节关键知识总结",
        asset_base_dir=project_root,
    )


def build_one(book_id: str) -> Path:
    book = BOOKS[book_id]
    project_root = OUTPUT_ROOT / book["folder_name"]
    paths = ensure_structure(project_root)

    shutil.copy2(book["source_pdf"], paths["original"] / "00_原书.pdf")
    doc = fitz.open(book["source_pdf"])
    write_json(paths["build"] / "chapter_manifest.json", manifest_for(book, doc))

    chapters: list[dict] = []
    for chapter in book["chapters"]:
        raw_text = extract_text_range(doc, chapter["start_page"], chapter["end_page"])
        text_filename = f"{chapter['sequence_no']}_{safe_name(chapter['title'])}.txt"
        (paths["build_text"] / text_filename).write_text(raw_text, encoding="utf-8")
        render_page(doc, chapter["opener_page"], paths["assets_source"] / f"{chapter['id']}_opener_page.png")
        chapters.append(chapter_content(book_id, book, chapter))

    write_json(paths["build"] / "chapter_content.json", {"book": book, "chapters": chapters})

    render_book_overview(book, chapters, paths["assets_report"] / "book_overview.png")
    render_common_methods(book, paths["assets_report"] / "common_methods.png")
    for chapter in chapters:
        render_chapter_core_map(chapter, paths["assets_report"] / f"{chapter['id']}_core_map.png")
        render_summary_card(chapter, paths["assets_summary"] / f"{chapter['id']}_memory_card.png")
        render_poster(chapter, paths["assets_ppt"] / f"{chapter['id']}_poster.png")

    (paths["report"] / book["report_filename"]).write_text(build_report(book, chapters), encoding="utf-8")
    (paths["summary"] / book["summary_filename"]).write_text(build_summary(book, chapters), encoding="utf-8")

    (project_root / "README.md").write_text(
        "\n".join(
            [
                f"# {book['title']} 交付包",
                "",
                "## 主要入口",
                "- `00_原书/00_原书.pdf`：原书副本。",
                f"- `20_阅读报告/{book['report_pdf_filename']}`：阅读报告 PDF。",
                f"- `30_关键知识总结/{book['summary_pdf_filename']}`：按章节关键知识总结 PDF。",
                "- `10_PPT/`：按章节拆分的独立 PPTX。",
                "- `40_配图/`：PPT、阅读报告和总结共用图片。",
                "- `_build/`：章节原文、结构化 JSON、构建脚本与渲染预览。",
                "",
            ]
        ),
        encoding="utf-8",
    )

    write_wrapper_scripts(book, project_root)
    render_markdown_pdfs(book, project_root)
    return project_root


def main() -> int:
    args = parse_args()
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    selected = [BOOKS[args.book_id]] if args.book_id else [BOOKS[key] for key in ["antifragile", "black_swan", "echo_of_genius"]]
    for book in selected:
        build_one(book["book_id"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
