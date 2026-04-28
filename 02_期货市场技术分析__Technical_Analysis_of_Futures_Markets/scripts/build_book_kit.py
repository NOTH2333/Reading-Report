from __future__ import annotations

import html
import json
import math
import re
import shutil
from pathlib import Path

import fitz
from pypdf import PdfReader

from book_blueprints import BOOK_AUTHOR, BOOK_TITLE, CHAPTERS, GROUPED_REPORT_SECTIONS, PDF_PATH


ROOT = Path(__file__).resolve().parents[1]
REPORT_DIR = ROOT / "01_全书阅读报告"
PPT_DIR = ROOT / "02_章节PPT"
SUMMARY_DIR = ROOT / "03_章节关键知识总结"
COMMON_DIR = ROOT / "90_共用素材与底稿"
COMMON_ORIGINAL_DIR = COMMON_DIR / "原书裁图"
COMMON_REDRAWN_DIR = COMMON_DIR / "重绘图表"
COMMON_TEXT_DIR = COMMON_DIR / "提取文本"
COMMON_STYLE_DIR = COMMON_DIR / "统一样式与图标"
HELPER_SRC = Path(r"C:\Users\霓\.codex\skills\slides\assets\pptxgenjs_helpers")

FONT_STACK = "Microsoft YaHei, 'Noto Sans SC', sans-serif"
BG = "#F7F1E6"
INK = "#253238"
ACCENT = "#0F766E"
ACCENT_2 = "#C84C31"
ACCENT_3 = "#8FA49A"
CARD = "#FFFDF8"
LINE = "#D9CCB8"
GREEN = "#2F8F5B"
RED = "#C25B4A"


def ensure_dir(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def write_text(path: Path, content: str) -> None:
    ensure_dir(path.parent)
    path.write_text(content, encoding="utf-8")


def copy_file(src: Path, dest: Path) -> None:
    ensure_dir(dest.parent)
    shutil.copy2(src, dest)


def clean_page_text(text: str) -> str:
    lines = []
    for raw in text.splitlines():
        line = re.sub(r"\s+", " ", raw).strip()
        if not line:
            lines.append("")
            continue
        if "www.alltick.co" in line or "实时行情" in line or "高频行情数据接口" in line:
            continue
        lines.append(line)
    text = "\n".join(lines)
    text = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def chapter_page_count(chapter: dict) -> int:
    return chapter["pages"][1] - chapter["pages"][0] + 1


def target_slide_count(chapter: dict) -> int:
    page_count = chapter_page_count(chapter)
    if chapter["id"] == 1:
        return 20
    if page_count <= 19:
        return 16
    if page_count <= 29:
        return 20
    if page_count <= 39:
        return 24
    return 28


def extract_texts(reader: PdfReader) -> None:
    index = []
    for chapter in CHAPTERS:
        start, end = chapter["pages"]
        parts = []
        for page_num in range(start, end + 1):
            text = reader.pages[page_num - 1].extract_text() or ""
            parts.append(f"===== PDF页 {page_num} =====\n{clean_page_text(text)}\n")
        out_path = COMMON_TEXT_DIR / f"第{chapter['id']:02d}章_{chapter['title']}.txt"
        write_text(out_path, "\n".join(parts))
        index.append(
            {
                "id": chapter["id"],
                "title": chapter["title"],
                "pages": chapter["pages"],
                "page_count": chapter_page_count(chapter),
                "text_file": str(out_path.relative_to(ROOT)).replace("\\", "/"),
            }
        )
    write_text(COMMON_TEXT_DIR / "章节索引.json", json.dumps(index, ensure_ascii=False, indent=2))


def render_page_thumbnail(pdf: fitz.Document, page_num: int, out_path: Path) -> None:
    ensure_dir(out_path.parent)
    page = pdf.load_page(page_num - 1)
    rect = page.rect
    margin_x = rect.width * 0.04
    margin_y = rect.height * 0.04
    clip = fitz.Rect(rect.x0 + margin_x, rect.y0 + margin_y, rect.x1 - margin_x, rect.y1 - margin_y)
    pix = page.get_pixmap(matrix=fitz.Matrix(1.2, 1.2), clip=clip, alpha=False)
    pix.save(str(out_path))


def esc(value: str) -> str:
    return html.escape(value, quote=True)


def svg_wrap(body: str, width: int = 1200, height: int = 700, background: str = BG) -> str:
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}" viewBox="0 0 {width} {height}">'
        f'<rect width="{width}" height="{height}" fill="{background}"/>'
        f"{body}</svg>"
    )


def svg_text(x: float, y: float, text: str, size: int = 24, fill: str = INK, anchor: str = "start", weight: int = 500) -> str:
    return (
        f'<text x="{x}" y="{y}" fill="{fill}" font-family="{FONT_STACK}" font-size="{size}" '
        f'font-weight="{weight}" text-anchor="{anchor}">{esc(text)}</text>'
    )


def rounded_rect(x: float, y: float, w: float, h: float, fill: str = CARD, stroke: str = LINE, rx: float = 22, sw: float = 2) -> str:
    return f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{rx}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>'


def line(x1: float, y1: float, x2: float, y2: float, stroke: str = INK, sw: float = 4, dash: str | None = None) -> str:
    dash_attr = f' stroke-dasharray="{dash}"' if dash else ""
    return f'<line x1="{x1}" y1="{y1}" x2="{x2}" y2="{y2}" stroke="{stroke}" stroke-width="{sw}" stroke-linecap="round"{dash_attr}/>'


def circle(cx: float, cy: float, r: float, fill: str = CARD, stroke: str = ACCENT, sw: float = 3) -> str:
    return f'<circle cx="{cx}" cy="{cy}" r="{r}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}"/>'


def arrow(x1: float, y1: float, x2: float, y2: float, stroke: str = ACCENT, sw: float = 4) -> str:
    angle = math.atan2(y2 - y1, x2 - x1)
    head = 14
    hx1 = x2 - head * math.cos(angle - math.pi / 6)
    hy1 = y2 - head * math.sin(angle - math.pi / 6)
    hx2 = x2 - head * math.cos(angle + math.pi / 6)
    hy2 = y2 - head * math.sin(angle + math.pi / 6)
    return (
        line(x1, y1, x2, y2, stroke, sw)
        + f'<polygon points="{x2},{y2} {hx1},{hy1} {hx2},{hy2}" fill="{stroke}"/>'
    )


def pill(x: float, y: float, text: str, fill: str = ACCENT, fg: str = "#ffffff") -> str:
    width = 28 + len(text) * 16
    body = rounded_rect(x, y, width, 34, fill=fill, stroke=fill, rx=17, sw=0)
    body += svg_text(x + width / 2, y + 23, text, 18, fg, anchor="middle", weight=700)
    return body


def wrapped_svg_text(x: float, y: float, text: str, max_chars: int = 10, size: int = 22, fill: str = INK, anchor: str = "middle") -> str:
    parts = []
    for idx, line_text in enumerate(textwrap_list(text, max_chars)):
        parts.append(svg_text(x, y + idx * (size + 8), line_text, size=size, fill=fill, anchor=anchor))
    return "".join(parts)


def textwrap_list(text: str, max_chars: int) -> list[str]:
    clean = text.replace("\n", " ").strip()
    if len(clean) <= max_chars:
        return [clean]
    return [clean[i : i + max_chars] for i in range(0, len(clean), max_chars)]


def chapter_palette(chapter_id: int) -> tuple[str, str]:
    accents = [
        ("#0F766E", "#C84C31"),
        ("#1D4ED8", "#0F766E"),
        ("#8B5CF6", "#C84C31"),
        ("#C84C31", "#0F766E"),
        ("#0F766E", "#8B5CF6"),
        ("#1D4ED8", "#C84C31"),
    ]
    return accents[(chapter_id - 1) % len(accents)]


def build_concept_map_svg(chapter: dict) -> str:
    primary, secondary = chapter_palette(chapter["id"])
    center_x = 600
    center_y = 350
    points = [
        (600, 110),
        (920, 210),
        (920, 490),
        (600, 590),
        (280, 490),
        (280, 210),
    ]
    modules = chapter["modules"][:6]
    body = pill(46, 40, f"第{chapter['id']:02d}章")
    body += svg_text(150, 66, chapter["title"], 30, INK, weight=800)
    body += rounded_rect(390, 255, 420, 180, fill="#FFF8EE", stroke=primary, rx=30, sw=4)
    body += svg_text(center_x, 320, chapter["title"], 34, primary, anchor="middle", weight=800)
    body += wrapped_svg_text(center_x, 365, chapter["big_idea"], 18, 22, anchor="middle")
    for idx, module in enumerate(modules):
        px, py = points[idx]
        body += arrow(center_x, center_y, px, py, primary, 4)
        body += rounded_rect(px - 115, py - 48, 230, 96, fill=CARD, stroke=secondary, rx=24, sw=3)
        body += wrapped_svg_text(px, py - 4, module, 8, 22)
    return svg_wrap(body)


def build_process_svg(chapter: dict) -> str:
    primary, secondary = chapter_palette(chapter["id"])
    body = pill(46, 40, "应用流程", secondary)
    body += svg_text(170, 66, chapter["title"], 30, INK, weight=800)
    y_positions = [150, 280, 410, 540]
    for idx, (step, y) in enumerate(zip(chapter["application"], y_positions), start=1):
        body += circle(110, y, 34, fill="#FFF8EE", stroke=primary, sw=4)
        body += svg_text(110, y + 8, str(idx), 28, primary, anchor="middle", weight=800)
        body += rounded_rect(180, y - 42, 860, 84, fill=CARD, stroke=LINE, rx=24, sw=2)
        body += wrapped_svg_text(220, y - 4, step, 24, 24, anchor="start")
        if idx < len(y_positions):
            body += arrow(110, y + 34, 110, y + 86, primary, 4)
    return svg_wrap(body)


def polyline(points: list[tuple[float, float]], stroke: str = ACCENT, sw: float = 5, fill: str = "none") -> str:
    joined = " ".join(f"{x},{y}" for x, y in points)
    return f'<polyline points="{joined}" fill="{fill}" stroke="{stroke}" stroke-width="{sw}" stroke-linejoin="round" stroke-linecap="round"/>'


def build_specific_svg(chapter: dict) -> str:
    cid = chapter["id"]
    primary, secondary = chapter_palette(cid)
    body = pill(46, 40, chapter["core_visual_label"], primary)
    body += svg_text(46, 112, chapter["title"], 34, INK, weight=800)
    if cid == 1:
        body += circle(600, 260, 94, fill="#FFF8EE", stroke=primary, sw=4)
        body += wrapped_svg_text(600, 248, "市场行为\n包容一切", 8, 26)
        body += circle(360, 500, 90, fill=CARD, stroke=secondary, sw=4)
        body += wrapped_svg_text(360, 490, "趋势\n会延续", 8, 26)
        body += circle(840, 500, 90, fill=CARD, stroke=secondary, sw=4)
        body += wrapped_svg_text(840, 490, "历史\n会重演", 8, 26)
        body += arrow(560, 334, 405, 440, primary, 4)
        body += arrow(640, 334, 795, 440, primary, 4)
        body += arrow(440, 500, 760, 500, secondary, 4)
    elif cid == 2:
        body += polyline([(120, 500), (260, 420), (420, 470), (620, 260), (820, 320), (1080, 180)], primary, 6)
        body += polyline([(120, 540), (210, 510), (300, 530), (420, 480), (540, 520), (680, 460), (820, 490), (960, 430)], secondary, 4)
        body += polyline([(120, 590), (160, 570), (200, 592), (240, 572), (280, 590), (320, 570), (360, 592)], "#A9B6B4", 3)
        body += svg_text(1040, 164, "主要趋势", 26, primary, weight=800)
        body += svg_text(930, 418, "次级趋势", 24, secondary, weight=800)
        body += svg_text(380, 630, "日常波动", 22, "#64748B", weight=700)
    elif cid == 3:
        for idx, label in enumerate(["线图", "条形图", "点数图"]):
            x = 120 + idx * 350
            body += rounded_rect(x, 180, 260, 360, fill=CARD, stroke=LINE, rx=24, sw=2)
            body += svg_text(x + 130, 220, label, 28, primary, anchor="middle", weight=800)
        body += polyline([(150, 470), (210, 430), (260, 450), (320, 320), (360, 350)], primary, 5)
        body += line(520, 430, 520, 520, INK, 3)
        body += line(490, 470, 520, 470, INK, 3)
        body += line(520, 440, 550, 440, INK, 3)
        body += line(610, 340, 610, 500, INK, 3)
        body += line(580, 400, 610, 400, INK, 3)
        body += line(610, 360, 640, 360, INK, 3)
        for x, pattern in [(850, "XOX"), (920, "XXO"), (990, "OXO")]:
            for row, char in enumerate(pattern):
                color = primary if char == "X" else secondary
                body += svg_text(x, 380 + row * 58, char, 42, color, anchor="middle", weight=800)
    elif cid == 4:
        body += polyline([(140, 550), (260, 470), (350, 500), (470, 380), (600, 430), (760, 270), (1020, 320)], primary, 6)
        body += line(120, 575, 1060, 300, secondary, 4)
        body += line(120, 640, 1060, 365, secondary, 4, "14 8")
        body += svg_text(905, 284, "上升通道上沿", 22, secondary, weight=700)
        body += svg_text(910, 360, "趋势线", 22, primary, weight=700)
    elif cid == 5:
        body += polyline([(140, 500), (260, 360), (360, 440), (560, 250), (720, 440), (850, 360), (980, 520)], primary, 6)
        body += line(200, 470, 900, 470, secondary, 4, "10 10")
        body += svg_text(935, 463, "颈线", 22, secondary, weight=700)
        body += svg_text(260, 336, "左肩", 22, primary, anchor="middle")
        body += svg_text(560, 226, "头部", 22, primary, anchor="middle")
        body += svg_text(850, 336, "右肩", 22, primary, anchor="middle")
    elif cid == 6:
        body += rounded_rect(90, 180, 320, 360, fill=CARD, stroke=LINE, rx=22, sw=2)
        body += polyline([(120, 480), (180, 410), (250, 450), (320, 380), (380, 430)], primary, 5)
        body += line(120, 500, 390, 330, secondary, 3)
        body += line(120, 300, 390, 520, secondary, 3)
        body += svg_text(250, 560, "三角形", 24, primary, anchor="middle", weight=800)
        body += rounded_rect(450, 180, 300, 360, fill=CARD, stroke=LINE, rx=22, sw=2)
        body += polyline([(490, 300), (560, 230), (630, 300)], primary, 5)
        body += polyline([(630, 300), (690, 325), (740, 310)], secondary, 5)
        body += svg_text(600, 560, "旗形/三角旗", 24, primary, anchor="middle", weight=800)
        body += rounded_rect(820, 180, 280, 360, fill=CARD, stroke=LINE, rx=22, sw=2)
        body += line(860, 280, 1060, 280, secondary, 4)
        body += line(860, 460, 1060, 460, secondary, 4)
        body += polyline([(860, 350), (920, 310), (980, 380), (1040, 330)], primary, 5)
        body += svg_text(960, 560, "矩形", 24, primary, anchor="middle", weight=800)
    elif cid == 7:
        body += polyline([(120, 500), (260, 430), (400, 460), (560, 330), (760, 360), (1020, 210)], primary, 5)
        bars = [120, 180, 140, 260, 200, 340]
        for idx, bar in enumerate(bars):
            x = 180 + idx * 120
            body += f'<rect x="{x}" y="{600 - bar}" width="58" height="{bar}" fill="{secondary}" opacity="0.45"/>'
        body += polyline([(120, 620), (280, 590), (450, 560), (620, 500), (820, 470), (1020, 420)], GREEN, 5)
        body += svg_text(1035, 220, "价格", 22, primary, weight=800)
        body += svg_text(1035, 430, "持仓兴趣", 22, GREEN, weight=800)
        body += svg_text(1035, 320, "交易量柱", 22, secondary, weight=800)
    elif cid == 8:
        for idx, label in enumerate(["日线", "周线", "月线"]):
            y = 180 + idx * 140
            body += rounded_rect(180, y, 340, 90, fill=CARD, stroke=LINE, rx=18, sw=2)
            body += svg_text(250, y + 40, label, 24, primary, weight=800)
            body += polyline([(320, y + 62), (360, y + 48), (400, y + 54), (440, y + 30), (485, y + 34)], primary, 4)
        body += rounded_rect(700, 200, 320, 260, fill="#FFF8EE", stroke=secondary, rx=28, sw=3)
        body += svg_text(860, 250, "商品指数", 28, secondary, anchor="middle", weight=800)
        body += svg_text(860, 300, "看整体温度", 22, INK, anchor="middle")
        body += svg_text(860, 340, "不只看单一品种", 22, INK, anchor="middle")
        body += arrow(520, 320, 690, 320, primary, 5)
    elif cid == 9:
        body += polyline([(140, 540), (220, 500), (340, 520), (420, 420), (540, 450), (680, 300), (860, 360), (1020, 250)], "#94A3B8", 4)
        body += polyline([(120, 560), (220, 548), (320, 520), (420, 498), (520, 460), (620, 420), (720, 388), (820, 350), (920, 320), (1020, 290)], primary, 6)
        body += polyline([(120, 590), (220, 588), (320, 580), (420, 564), (520, 540), (620, 510), (720, 472), (820, 430), (920, 390), (1020, 350)], secondary, 6)
        body += circle(720, 472, 10, fill=secondary, stroke=secondary, sw=0)
        body += circle(720, 388, 10, fill=primary, stroke=primary, sw=0)
        body += svg_text(760, 396, "金叉", 22, primary, weight=800)
    elif cid == 10:
        body += polyline([(110, 320), (240, 280), (360, 290), (500, 220), (650, 235), (820, 180), (1000, 195)], primary, 5)
        body += polyline([(110, 560), (240, 520), (360, 525), (500, 530), (650, 500), (820, 510), (1000, 490)], secondary, 5)
        body += line(90, 540, 1040, 540, "#94A3B8", 3, "10 10")
        body += line(90, 470, 1040, 470, "#CBD5E1", 2, "10 10")
        body += line(90, 610, 1040, 610, "#CBD5E1", 2, "10 10")
        body += arrow(780, 190, 780, 500, ACCENT_2, 4)
        body += svg_text(820, 360, "价格创新高\n指标未创新高", 22, ACCENT_2, weight=700)
    elif cid == 11:
        cols = ["XXX", "OO", "XXXX", "OOO", "XXX"]
        x0 = 300
        for ci, col in enumerate(cols):
            for ri, char in enumerate(col):
                color = primary if char == "X" else secondary
                body += svg_text(x0 + ci * 90, 520 - ri * 62, char, 44, color, anchor="middle", weight=800)
        body += line(250, 560, 760, 560, "#CBD5E1", 2)
        body += svg_text(840, 330, "按价格变动记格子", 24, primary, weight=800)
        body += svg_text(840, 380, "不是按时间一格一格走", 22, INK)
    elif cid == 12:
        body += rounded_rect(140, 190, 360, 280, fill=CARD, stroke=LINE, rx=24, sw=2)
        body += svg_text(320, 240, "三点转向", 30, primary, anchor="middle", weight=800)
        body += svg_text(320, 300, "逆向移动 3 格", 24, INK, anchor="middle")
        body += svg_text(320, 340, "才允许换列", 24, INK, anchor="middle")
        body += rounded_rect(690, 190, 360, 280, fill="#FFF8EE", stroke=secondary, rx=24, sw=3)
        body += svg_text(870, 240, "参数优化", 30, secondary, anchor="middle", weight=800)
        body += svg_text(870, 300, "箱值 × 转向", 24, INK, anchor="middle")
        body += svg_text(870, 340, "比较稳定性", 24, INK, anchor="middle")
        body += arrow(500, 330, 680, 330, primary, 5)
    elif cid == 13:
        body += polyline([(120, 520), (240, 420), (360, 460), (520, 300), (680, 360), (860, 220), (1040, 300)], primary, 6)
        body += polyline([(1040, 300), (920, 380), (820, 340), (720, 420)], secondary, 6)
        for x, y, label in [(240, 420, "1"), (360, 460, "2"), (520, 300, "3"), (680, 360, "4"), (860, 220, "5"), (920, 380, "A"), (820, 340, "B"), (720, 420, "C")]:
            body += circle(x, y - 28, 18, fill="#FFF8EE", stroke=LINE, sw=2)
            body += svg_text(x, y - 20, label, 20, INK, anchor="middle", weight=800)
    elif cid == 14:
        wave1 = "M80 380 C180 260 280 500 380 380 S580 260 680 380 S880 500 980 380"
        wave2 = "M80 500 C180 430 280 560 380 500 S580 430 680 500 S880 560 980 500"
        body += f'<path d="{wave1}" fill="none" stroke="{primary}" stroke-width="6"/>'
        body += f'<path d="{wave2}" fill="none" stroke="{secondary}" stroke-width="4"/>'
        for x in [380, 680, 980]:
            body += line(x, 180, x, 580, "#CBD5E1", 3, "8 10")
        body += svg_text(1050, 384, "长周期", 22, primary, weight=800)
        body += svg_text(1050, 504, "短周期", 22, secondary, weight=800)
    elif cid == 15:
        xs = [130, 380, 630, 880]
        labels = ["规则想法", "回测检验", "执行系统", "复盘修正"]
        for x, label in zip(xs, labels):
            body += rounded_rect(x, 290, 180, 120, fill=CARD, stroke=LINE, rx=24, sw=2)
            body += wrapped_svg_text(x + 90, 350, label, 7, 26)
        for a, b in zip(xs[:-1], xs[1:]):
            body += arrow(a + 180, 350, b, 350, primary, 5)
    else:
        body += rounded_rect(140, 200, 260, 260, fill=CARD, stroke=LINE, rx=28, sw=2)
        body += rounded_rect(470, 200, 260, 260, fill=CARD, stroke=LINE, rx=28, sw=2)
        body += rounded_rect(800, 200, 260, 260, fill=CARD, stroke=LINE, rx=28, sw=2)
        body += svg_text(270, 330, "方向", 30, primary, anchor="middle", weight=800)
        body += svg_text(600, 330, "时机", 30, secondary, anchor="middle", weight=800)
        body += svg_text(930, 330, "资金", 30, GREEN, anchor="middle", weight=800)
        body += arrow(400, 330, 470, 330, primary, 5)
        body += arrow(730, 330, 800, 330, secondary, 5)
    return svg_wrap(body)


def build_global_roadmap_svg() -> str:
    body = pill(46, 40, "全书路线图")
    body += svg_text(180, 66, BOOK_TITLE, 30, INK, weight=800)
    cols = 4
    for idx, chapter in enumerate(CHAPTERS):
        row = idx // cols
        col = idx % cols
        x = 80 + col * 270
        y = 150 + row * 120
        primary, _ = chapter_palette(chapter["id"])
        body += rounded_rect(x, y, 220, 88, fill=CARD, stroke=primary, rx=18, sw=3)
        body += svg_text(x + 30, y + 34, f"第{chapter['id']:02d}章", 18, primary, weight=800)
        body += wrapped_svg_text(x + 110, y + 60, chapter["title"], 8, 20)
    return svg_wrap(body)


def build_global_premises_svg() -> str:
    fake = {
        "id": 1,
        "title": "技术分析三大前提",
        "big_idea": "市场行为包容消化一切，价格以趋势方式演变，历史会重演。",
        "modules": ["市场行为包容消化一切", "价格以趋势方式演变", "历史会重演"],
    }
    return build_concept_map_svg(fake)


def build_global_matrix_svg() -> str:
    body = pill(46, 40, "工具矩阵", ACCENT_2)
    body += svg_text(180, 66, "趋势、形态、指标、风控", 30, INK, weight=800)
    headers = ["层级", "核心问题", "代表工具"]
    xs = [120, 390, 690]
    for x, header in zip(xs, headers):
        body += rounded_rect(x, 160, 220, 64, fill="#FFF8EE", stroke=LINE, rx=18, sw=2)
        body += svg_text(x + 110, 200, header, 24, ACCENT_2, anchor="middle", weight=800)
    rows = [
        ("趋势", "市场往哪边走", "道氏理论、趋势线、均线"),
        ("结构", "在途中怎样整理或转身", "头肩形、三角形、矩形"),
        ("状态", "现在是过热还是衰竭", "RSI、动量、背离"),
        ("生存", "该押多少、错了怎么办", "止损、仓位、分散"),
    ]
    for idx, row in enumerate(rows):
        y = 260 + idx * 96
        for col, value in enumerate(row):
            x = xs[col]
            w = 220 if col < 2 else 360
            body += rounded_rect(x, y, w, 68, fill=CARD, stroke=LINE, rx=16, sw=2)
            body += wrapped_svg_text(x + w / 2, y + 42, value, 12 if col < 2 else 16, 22)
    return svg_wrap(body)


def build_global_indicator_svg() -> str:
    body = pill(46, 40, "指标家族", ACCENT)
    body += svg_text(180, 66, "技术分析常见工具分组", 30, INK, weight=800)
    groups = [
        ("趋势组", "趋势线、通道、均线", 180, 180, ACCENT),
        ("形态组", "反转形态、持续形态、点数图", 690, 180, ACCENT_2),
        ("动能组", "RSI、动量、背离", 180, 420, "#4C7A67"),
        ("风险组", "止损、仓位、分散", 690, 420, "#6B7280"),
    ]
    for title, desc, x, y, color in groups:
        body += rounded_rect(x, y, 320, 140, fill=CARD, stroke=color, rx=24, sw=3)
        body += svg_text(x + 160, y + 48, title, 30, color, anchor="middle", weight=800)
        body += wrapped_svg_text(x + 160, y + 92, desc, 12, 22)
    body += circle(600, 350, 70, fill="#FFF8EE", stroke=LINE, sw=3)
    body += wrapped_svg_text(600, 338, "一套完整的\n交易语言", 8, 22)
    for x, y in [(500, 250), (700, 250), (500, 450), (700, 450)]:
        body += arrow(600, 350, x, y, "#94A3B8", 4)
    return svg_wrap(body)


def build_global_tripod_svg() -> str:
    return build_specific_svg({"id": 16, "title": "交易三脚架", "core_visual_label": "交易三脚架与风险控制"})


def build_global_workflow_svg() -> str:
    body = pill(46, 40, "阅读与实战流程", "#1D4ED8")
    body += svg_text(230, 66, "看市场的五个固定动作", 30, INK, weight=800)
    steps = ["先看周期", "再判趋势", "接着找结构", "然后做验证", "最后管风险"]
    xs = [120, 330, 540, 750, 960]
    for idx, (x, step) in enumerate(zip(xs, steps), start=1):
        body += rounded_rect(x, 280, 170, 120, fill=CARD, stroke=LINE, rx=24, sw=2)
        body += circle(x + 36, 340, 22, fill="#FFF8EE", stroke="#1D4ED8", sw=3)
        body += svg_text(x + 36, 348, str(idx), 20, "#1D4ED8", anchor="middle", weight=800)
        body += wrapped_svg_text(x + 100, 334, step, 6, 24)
        if idx < len(xs):
            body += arrow(x + 170, 340, x + 210, 340, "#1D4ED8", 4)
    return svg_wrap(body)


def generate_assets(pdf: fitz.Document) -> dict[int, dict[str, Path]]:
    asset_index: dict[int, dict[str, Path]] = {}
    for chapter in CHAPTERS:
        common_prefix = f"第{chapter['id']:02d}章_{chapter['title']}"
        concept_master = COMMON_REDRAWN_DIR / f"{common_prefix}_概念地图.svg"
        visual_master = COMMON_REDRAWN_DIR / f"{common_prefix}_核心图示.svg"
        process_master = COMMON_REDRAWN_DIR / f"{common_prefix}_应用流程.svg"
        reference_master = COMMON_ORIGINAL_DIR / f"{common_prefix}_原书页缩略图.png"
        write_text(concept_master, build_concept_map_svg(chapter))
        write_text(visual_master, build_specific_svg(chapter))
        write_text(process_master, build_process_svg(chapter))
        render_page_thumbnail(pdf, chapter["representative_page"], reference_master)

        summary_assets = SUMMARY_DIR / chapter["dir_name"] / "assets"
        ppt_assets = PPT_DIR / chapter["dir_name"] / "assets"
        for folder in [summary_assets, ppt_assets]:
            ensure_dir(folder)
        summary_files = {
            "map": summary_assets / "01_概念地图.svg",
            "visual": summary_assets / "02_核心图示.svg",
            "process": summary_assets / "03_应用流程.svg",
            "reference": summary_assets / "04_原书页缩略图.png",
        }
        ppt_files = {
            "map": ppt_assets / "01_概念地图.svg",
            "visual": ppt_assets / "02_核心图示.svg",
            "process": ppt_assets / "03_应用流程.svg",
            "reference": ppt_assets / "04_原书页缩略图.png",
        }
        for dest_set in [summary_files, ppt_files]:
            copy_file(concept_master, dest_set["map"])
            copy_file(visual_master, dest_set["visual"])
            copy_file(process_master, dest_set["process"])
            copy_file(reference_master, dest_set["reference"])
        asset_index[chapter["id"]] = {"summary": summary_files, "ppt": ppt_files}
    return asset_index


def rel(path: Path, base: Path) -> str:
    return str(path.relative_to(base)).replace("\\", "/")


def render_summary_markdown(chapter: dict, assets: dict[str, Path]) -> str:
    start, end = chapter["pages"]
    lines = [
        f"# {chapter['full_title']}",
        "",
        f"> PDF页范围：{start}-{end}。核心图示：{chapter['core_visual_label']}。",
        "",
        f"**一句话总纲**：{chapter['big_idea']}",
        "",
        f"![概念地图](./{rel(assets['map'], assets['map'].parent.parent)})",
        "",
        f"![核心图示](./{rel(assets['visual'], assets['visual'].parent.parent)})",
        "",
        f"![应用流程](./{rel(assets['process'], assets['process'].parent.parent)})",
        "",
        f"![原书页缩略图](./{rel(assets['reference'], assets['reference'].parent.parent)})",
        "",
        "## 这章到底在讲什么",
        "",
        f"{chapter['why_it_matters']} 作者在这一章真正想训练的，不只是识别名词，而是把市场现象翻译成一套能重复使用的判断语言。",
        "",
        "## 本章核心术语",
        "",
    ]
    for term in chapter["terms"]:
        lines.append(f"- **{term['term']}**：{term['explain']}")
    lines += ["", "## 关键知识", ""]
    for idx, kp in enumerate(chapter["key_points"], start=1):
        lines += [
            f"### 关键知识 {idx}：{kp['title']}",
            "",
            f"{kp['idea']} 站在零基础读者角度，可以先把它理解成一句很朴素的话：市场在这里留下了一个可重复辨认的行为模式。",
            "",
            f"**怎么看**：{kp['how_to_read']}",
            "",
            f"**最容易错在哪里**：{kp['mistake']}",
            "",
            f"**真正能带走的收获**：{kp['harvest']}",
            "",
        ]
    lines += [
        "## 直观比喻",
        "",
        chapter["analogy"],
        "",
        "## 典型图示怎么读",
        "",
        f"上面的核心图示并不是为了让你死记图样，而是帮你抓住 `{chapter['core_visual_label']}` 背后的结构关系。真正该记住的是：先看背景，再看结构，再看确认，最后才谈动作。",
        "",
        "## 3 个最容易误解的问题",
        "",
    ]
    for item in chapter["misconceptions"]:
        lines += [f"- **{item['question']}**", f"  答：{item['answer']}"]
    lines += ["", "## 本章收获清单", ""]
    for takeaway in chapter["takeaways"]:
        lines.append(f"- {takeaway}")
    lines += [
        "",
        "## 如果讲给完全不懂的人听",
        "",
        f"你可以这样概括这一章：{chapter['big_idea']} 先把这件事讲成一个生活故事，再回到图表上找对应证据，理解会快很多。",
        "",
    ]
    return "\n".join(lines)


def render_report_markdown() -> str:
    chapter_lookup = {chapter["id"]: chapter for chapter in CHAPTERS}
    lines = [
        f"# 《{BOOK_TITLE}》全书阅读报告",
        "",
        f"> 作者：{BOOK_AUTHOR}",
        "",
        "> 目标读者：零基础到初学者，希望在读完全书后，不只知道名词，还知道这些工具为什么存在、怎样使用、哪里最容易犯错。",
        "",
        "![全书路线图](./assets/report_01_全书路线图.svg)",
        "",
        "![技术分析三大前提](./assets/report_02_三大前提.svg)",
        "",
        "## 阅读方法说明",
        "",
        "这份报告不是把原书改写成更短的摘抄，而是把它重组为一套更适合学习的知识结构。原书的核心价值，在于它不断提醒读者：技术分析既不是神秘学，也不是一套只适合短线投机的花招，而是一种围绕市场行为、趋势、结构、验证和风险控制展开的概率语言。",
        "",
        "对完全不懂的读者来说，最重要的不是一下子掌握所有图形和指标，而是先建立一条稳定的认知顺序：先确定自己在看哪个时间尺度，再判断大方向，然后识别图形和结构，接着用量能与指标做验证，最后把所有结论放进风险管理框架。这条顺序几乎贯穿了全书所有章节。",
        "",
        "作者写作本书时，既在讲工具，也在不断讲一种交易者的姿态：不要急着找神奇信号，不要把单一指标神化，不要把猜顶猜底当成聪明，不要把方向判断和资金管理拆开。正因为如此，这本书虽然写于更早时期，但方法论到今天仍然能成立。",
        "",
        "![工具矩阵](./assets/report_03_工具矩阵.svg)",
        "",
        "![指标家族](./assets/report_04_指标家族.svg)",
        "",
    ]
    for section in GROUPED_REPORT_SECTIONS:
        lines += [f"## {section['title']}", "", section["summary"], ""]
        if section["title"] == "一、全书导论":
            lines += [
                "这一部分不对应某一章，而是回答两个最基础的问题：这本书到底在训练什么能力，以及读者应该怎样安排自己的学习顺序。整本书可以看成一条从“会看图”到“会下计划”再到“会活下来”的学习路径。前半部建立市场语言，中部引入形态与指标，后半部开始转向规则、系统与资金管理。",
                "",
                "如果只学前面的趋势和形态，不学后面的系统与资金管理，读者往往会落入一个常见误区：会描述市场，却不会组织交易。反过来，如果只迷恋系统与参数，而没有理解前面的趋势、形态和量价关系，系统又会变成一堆缺乏市场语义的按钮。作者真正希望读者建立的是一整套前后连通的能力。",
                "",
                "![阅读流程](./assets/report_06_阅读流程.svg)",
                "",
            ]
            continue
        if section["title"] == "六、全书综合收获与适用边界":
            lines += [
                "从全书综合来看，作者最强的主张有三条。第一，市场价格往往比解释它的新闻更快，因此应优先尊重市场留下的证据。第二，单一工具永远不够，趋势、形态、量价、动能和风险管理必须连成链条。第三，真正成熟的交易不是追求次次正确，而是在不确定环境中持续保持优势。",
                "",
                "本书也有明显边界。许多图形和波浪判断都带有一定主观性；参数选择与系统优化存在过拟合风险；突发性事件和制度性变化并不会因为你画了趋势线就自动失效。因此，技术分析最适合作为“组织市场证据、帮助做概率判断”的框架，而不是“保证预测正确”的承诺。",
                "",
                "对今天的读者来说，本书最大的收获并不是某个具体指标，而是四种长期有用的能力：一是先用多周期和大背景建立语境；二是把图形看成群体心理的行为结果；三是用验证和风控抑制主观冲动；四是把任何交易想法转化为可以复盘的规则。只要这四种能力建立起来，这本书的价值就真正落地了。",
                "",
                "![交易三脚架](./assets/report_05_交易三脚架.svg)",
                "",
            ]
            continue
        for chapter_id in section["chapters"]:
            chapter = chapter_lookup[chapter_id]
            start, end = chapter["pages"]
            lines += [
                f"### {chapter['full_title']}（PDF页 {start}-{end}）",
                "",
                f"![{chapter['title']} 概念地图](./assets/chapter_{chapter_id:02d}_map.svg)",
                "",
                f"{chapter['why_it_matters']} 从全书结构看，这一章承担的任务是：把 `{chapter['title']}` 从一个术语，变成一套读图和决策的动作。它真正要教给读者的不是背答案，而是看到图表时先问什么、后问什么。",
                "",
                f"如果把这一章压缩成一句最好记的话，那就是：{chapter['big_idea']} 这句话看似简单，但背后实际上包含了作者反复强调的结构、确认和纪律三层意思。",
                "",
                "本章的主线可以概括为：",
            ]
            for kp in chapter["key_points"]:
                lines.append(f"- **{kp['title']}**：{kp['idea']} 读图时应当记住“{kp['how_to_read']}”，同时警惕“{kp['mistake']}”。最后真正沉淀下来的能力，是“{kp['harvest']}”。")
            lines += [
                "",
                f"如果把这一章讲给完全不懂市场的人听，最好的入口通常不是图表术语，而是生活类比：{chapter['analogy']} 当读者先抓住这个生活结构，再去看图上的线、形态或指标，就不容易陷入死背图样的状态。",
                "",
                "从实战角度看，这一章最重要的收获不只是“会不会认”，更是“会不会用”。作者不断提醒读者，任何工具都必须回到流程里去使用，而不是脱离背景独自下命令。换句话说，真正成熟的技术分析，不是看到一个信号就兴奋，而是先把它放到时间尺度、趋势环境、验证条件和风险控制里再判断。",
                "",
            ]
    return "\n".join(lines)


def build_slide_plan(chapter: dict, ppt_assets: dict[str, Path]) -> dict:
    fixed = [
        {"type": "cover", "title": chapter["full_title"], "subtitle": f"PDF页 {chapter['pages'][0]}-{chapter['pages'][1]} · 面向零基础的精讲版"},
        {"type": "objectives", "title": "学完这章你会什么", "bullets": chapter["takeaways"][:4]},
        {"type": "map", "title": "本章地图", "items": chapter["modules"]},
        {"type": "thesis", "title": "一句话主旨", "statement": chapter["big_idea"]},
        {"type": "why", "title": "为什么这一章重要", "statement": chapter["why_it_matters"], "analogy": chapter["analogy"]},
    ]
    for kp in chapter["key_points"]:
        fixed.append(
            {
                "type": "keypoint",
                "title": kp["title"],
                "idea": kp["idea"],
                "how": kp["how_to_read"],
                "mistake": kp["mistake"],
                "harvest": kp["harvest"],
            }
        )
    ending = [
        {"type": "misconceptions", "title": "常见误区", "items": chapter["misconceptions"]},
        {"type": "application", "title": "怎么用于交易理解", "steps": chapter["application"]},
        {"type": "takeaways", "title": "本章收获", "items": chapter["takeaways"]},
        {"type": "quiz", "title": "5题复盘页", "items": chapter["quiz"]},
    ]
    extras = [
        {"type": "terms", "title": "术语卡片 A", "items": chapter["terms"][:2]},
        {"type": "terms", "title": "术语卡片 B", "items": chapter["terms"][2:]},
        {"type": "diagram", "title": "核心图示", "image": Path("assets/02_核心图示.svg").as_posix(), "caption": chapter["core_visual_label"]},
        {"type": "reference", "title": "原书页缩略图", "image": Path("assets/04_原书页缩略图.png").as_posix(), "caption": f"来源：PDF页 {chapter['representative_page']}"},
        {"type": "diagram", "title": "应用流程", "image": Path("assets/03_应用流程.svg").as_posix(), "caption": "把本章方法放进流程"},
    ]
    for kp in chapter["key_points"]:
        extras.append(
            {
                "type": "deepdive",
                "title": f"{kp['title']}：怎么看",
                "bullets": [kp["idea"], f"观察重点：{kp['how_to_read']}", f"常见错误：{kp['mistake']}", f"真正收获：{kp['harvest']}"],
            }
        )
    extras += [
        {"type": "cards", "title": "关键收获卡片", "items": chapter["takeaways"][:4]},
        {"type": "cards", "title": "给零基础读者的讲法", "items": [chapter["analogy"], chapter["big_idea"], chapter["why_it_matters"], chapter["application"][0]]},
        {"type": "cards", "title": "本章判断流程", "items": chapter["application"]},
        {"type": "cards", "title": "记住这四件事", "items": [kp["harvest"] for kp in chapter["key_points"][:4]]},
    ]
    target = target_slide_count(chapter)
    slides = fixed[:]
    while len(slides) + len(ending) < target and extras:
        slides.append(extras.pop(0))
    slides.extend(ending)
    return {
        "id": chapter["id"],
        "title": chapter["title"],
        "fullTitle": chapter["full_title"],
        "dirName": chapter["dir_name"],
        "targetSlides": target,
        "pageRange": chapter["pages"],
        "pptFile": f"第{chapter['id']:02d}章_{chapter['title']}.pptx",
        "sourceFile": f"第{chapter['id']:02d}章_{chapter['title']}.js",
        "slides": slides,
    }


def render_readme() -> str:
    lines = [
        f"# {BOOK_TITLE} 学习资料包",
        "",
        "这套资料包围绕原书 16 章组织，目标是让零基础读者能按章节学习，并在每章结束后知道“这一章讲了什么、为什么重要、怎样使用、最容易错在哪里”。",
        "",
        "## 文件结构",
        "",
        "- `01_全书阅读报告/`：整本书的长篇阅读报告与总图示。",
        "- `02_章节PPT/`：16 份按章精讲 PPT，以及对应可重建的 `.js` 源文件。",
        "- `03_章节关键知识总结/`：16 份按章总结，适合快速复习或导入知识库。",
        "- `90_共用素材与底稿/`：原书页缩略图、重绘图表、提取文本和生成脚本中间产物。",
        "",
        "## 建议阅读顺序",
        "",
        "1. 先看 `01_全书阅读报告/`，建立全书地图。",
        "2. 再按章节阅读 `03_章节关键知识总结/`，熟悉核心概念。",
        "3. 需要授课、演示或系统复盘时，再打开 `02_章节PPT/`。",
        "",
        "## 章节对应表",
        "",
        "| 章节 | 标题 | PDF页 | PPT页数目标 |",
        "|---|---|---:|---:|",
    ]
    for chapter in CHAPTERS:
        start, end = chapter["pages"]
        lines.append(f"| 第{chapter['id']:02d}章 | {chapter['title']} | {start}-{end} | {target_slide_count(chapter)} |")
    lines += ["", "## 说明", "", "- 所有讲解文本均为原创整理与教学化重构，不是原书逐段摘抄。", "- 每章都同时提供图示、误区提醒、收获清单和复盘问题。", ""]
    return "\n".join(lines)


def main() -> None:
    for path in [REPORT_DIR, PPT_DIR, SUMMARY_DIR, COMMON_ORIGINAL_DIR, COMMON_REDRAWN_DIR, COMMON_TEXT_DIR, COMMON_STYLE_DIR]:
        ensure_dir(path)
    if HELPER_SRC.exists():
        helper_dest = COMMON_STYLE_DIR / "pptxgenjs_helpers"
        if helper_dest.exists():
            shutil.rmtree(helper_dest)
        shutil.copytree(HELPER_SRC, helper_dest)

    reader = PdfReader(PDF_PATH)
    extract_texts(reader)
    pdf = fitz.open(PDF_PATH)
    asset_index = generate_assets(pdf)

    report_assets = REPORT_DIR / "assets"
    ensure_dir(report_assets)
    report_svgs = {
        "report_01_全书路线图.svg": build_global_roadmap_svg(),
        "report_02_三大前提.svg": build_global_premises_svg(),
        "report_03_工具矩阵.svg": build_global_matrix_svg(),
        "report_04_指标家族.svg": build_global_indicator_svg(),
        "report_05_交易三脚架.svg": build_global_tripod_svg(),
        "report_06_阅读流程.svg": build_global_workflow_svg(),
    }
    for name, svg in report_svgs.items():
        write_text(report_assets / name, svg)
    for chapter in CHAPTERS:
        copy_file(asset_index[chapter["id"]]["summary"]["map"], report_assets / f"chapter_{chapter['id']:02d}_map.svg")

    for chapter in CHAPTERS:
        summary_assets = asset_index[chapter["id"]]["summary"]
        summary_path = SUMMARY_DIR / chapter["dir_name"] / f"第{chapter['id']:02d}章_关键知识总结.md"
        write_text(summary_path, render_summary_markdown(chapter, summary_assets))
        ensure_dir(PPT_DIR / chapter["dir_name"] / "review")

    report_path = REPORT_DIR / f"{BOOK_TITLE}_全书阅读报告.md"
    write_text(report_path, render_report_markdown())
    write_text(ROOT / "README.md", render_readme())

    decks = {"bookTitle": BOOK_TITLE, "chapters": [build_slide_plan(ch, asset_index[ch["id"]]["ppt"]) for ch in CHAPTERS]}
    write_text(COMMON_DIR / "chapter_decks.json", json.dumps(decks, ensure_ascii=False, indent=2))
    pdf.close()


if __name__ == "__main__":
    main()
