from __future__ import annotations

import json
import math
import textwrap
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont


WIDTH = 1600
HEIGHT = 900

PALETTE = {
    "forest": "#123524",
    "forest_2": "#1D4D38",
    "mint": "#D6E8D6",
    "cream": "#F5F0E1",
    "paper": "#EFE6CE",
    "ink": "#142013",
    "copper": "#B46A3C",
    "red": "#9A3D2E",
    "gold": "#C99D3A",
    "slate": "#5E6A64",
}


def font_candidates(*names: str) -> list[Path]:
    fonts_dir = Path("C:/Windows/Fonts")
    return [fonts_dir / name for name in names]


def load_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    candidates = (
        font_candidates("msyhbd.ttc", "msyh.ttc", "simhei.ttf", "simsun.ttc")
        if bold
        else font_candidates("msyh.ttc", "simsun.ttc", "arial.ttf")
    )
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


FONT_H1 = load_font(60, bold=True)
FONT_H2 = load_font(40, bold=True)
FONT_H3 = load_font(28, bold=True)
FONT_BODY = load_font(26)
FONT_SMALL = load_font(22)
FONT_TINY = load_font(18)


def make_canvas(color: str) -> Image.Image:
    return Image.new("RGB", (WIDTH, HEIGHT), color)


def add_texture(img: Image.Image, opacity: int = 24) -> None:
    draw = ImageDraw.Draw(img)
    for y in range(0, HEIGHT, 28):
        tone = 240 - (y % 56)
        draw.line([(0, y), (WIDTH, y)], fill=(tone, tone, tone), width=1)
    for x in range(0, WIDTH, 48):
        draw.line([(x, 0), (x, HEIGHT)], fill=(255, 255, 255), width=1)
    overlay = Image.new("RGBA", (WIDTH, HEIGHT), (255, 255, 255, 0))
    od = ImageDraw.Draw(overlay)
    for i in range(0, WIDTH, 120):
        od.line([(i, 0), (0, i)], fill=(255, 255, 255, opacity), width=1)
    img.alpha_composite(overlay.convert("RGBA")) if img.mode == "RGBA" else img.paste(
        Image.alpha_composite(img.convert("RGBA"), overlay).convert("RGB")
    )


def wrap_lines(text: str, width: int) -> str:
    lines = textwrap.wrap(text, width=width, break_long_words=False, break_on_hyphens=False)
    return "\n".join(lines)


def draw_paragraph(draw: ImageDraw.ImageDraw, xy: tuple[int, int], text: str, font, fill, width_chars: int, spacing: int = 10) -> int:
    wrapped = wrap_lines(text, width_chars)
    draw.multiline_text(xy, wrapped, font=font, fill=fill, spacing=spacing)
    bbox = draw.multiline_textbbox(xy, wrapped, font=font, spacing=spacing)
    return bbox[3] - bbox[1]


def rounded_panel(draw: ImageDraw.ImageDraw, box: tuple[int, int, int, int], fill: str, outline: str | None = None, radius: int = 28) -> None:
    draw.rounded_rectangle(box, radius=radius, fill=fill, outline=outline, width=2 if outline else 1)


def generate_theme_backgrounds(theme_dir: Path) -> dict[str, str]:
    theme_dir.mkdir(parents=True, exist_ok=True)

    cover = make_canvas(PALETTE["forest"])
    cover = cover.convert("RGBA")
    overlay = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    od = ImageDraw.Draw(overlay)
    for i in range(12):
        od.ellipse(
            (900 + i * 20, -180 + i * 24, 1600 + i * 80, 520 + i * 60),
            fill=(180, 106, 60, max(8, 28 - i)),
        )
    for x in range(120, WIDTH, 110):
        od.line([(x, 0), (x - 240, HEIGHT)], fill=(255, 255, 255, 18), width=2)
    cover = Image.alpha_composite(cover, overlay).convert("RGB")
    add_texture(cover)
    cover_path = theme_dir / "bg_cover.png"
    cover.save(cover_path)

    content = make_canvas(PALETTE["cream"])
    draw = ImageDraw.Draw(content)
    for y in range(90, HEIGHT, 90):
        draw.line([(80, y), (WIDTH - 80, y)], fill="#E7DAC0", width=1)
    for x in range(120, WIDTH, 160):
        draw.line([(x, 70), (x, HEIGHT - 70)], fill="#EFE7D4", width=1)
    draw.polygon([(0, 0), (420, 0), (0, 240)], fill="#F1E4C6")
    draw.polygon([(WIDTH, HEIGHT), (WIDTH - 420, HEIGHT), (WIDTH, HEIGHT - 240)], fill="#E9D9B4")
    content_path = theme_dir / "bg_content.png"
    content.save(content_path)

    closing = make_canvas(PALETTE["forest_2"])
    draw = ImageDraw.Draw(closing)
    for i in range(0, WIDTH, 90):
        draw.line([(i, 0), (i + 260, HEIGHT)], fill="#335945", width=2)
    for i in range(0, HEIGHT, 110):
        draw.line([(0, i), (WIDTH, i)], fill="#214A36", width=1)
    closing_path = theme_dir / "bg_closing.png"
    closing.save(closing_path)

    return {
        "cover": str(cover_path),
        "content": str(content_path),
        "closing": str(closing_path),
    }


def stage_color(stage: str) -> str:
    mapping = {
        "起步试错期": PALETTE["copper"],
        "方法成形期": PALETTE["gold"],
        "扩张与回撤期": PALETTE["red"],
        "专业化总结期": PALETTE["forest_2"],
        "收官警示期": "#5B4D8A",
    }
    return mapping.get(stage, PALETTE["copper"])


def create_card(title: str, badge: str, subtitle: str, points: list[str], rule: str, stage: str, out_path: Path) -> None:
    img = make_canvas(PALETTE["paper"])
    draw = ImageDraw.Draw(img)
    draw.rectangle((0, 0, WIDTH, 110), fill=stage_color(stage))
    draw.rectangle((0, HEIGHT - 90, WIDTH, HEIGHT), fill=PALETTE["forest"])

    rounded_panel(draw, (70, 150, 1530, 780), fill="#FBF7EC", outline="#D6C7A6", radius=36)
    rounded_panel(draw, (95, 175, 340, 410), fill=PALETTE["forest"], radius=28)

    draw.text((122, 210), badge, font=FONT_H1, fill=PALETTE["cream"])
    draw.text((380, 190), wrap_lines(title, 18), font=FONT_H1, fill=PALETTE["ink"], spacing=8)
    draw.text((382, 330), wrap_lines(subtitle, 28), font=FONT_BODY, fill=PALETTE["slate"], spacing=10)

    y = 440
    for idx, point in enumerate(points[:3], start=1):
        rounded_panel(draw, (120, y, 1480, y + 88), fill="#F0E7D0", radius=22)
        draw.ellipse((140, y + 22, 184, y + 66), fill=stage_color(stage))
        draw.text((152, y + 24), str(idx), font=FONT_SMALL, fill="#FFFFFF")
        draw.text((212, y + 20), wrap_lines(point, 42), font=FONT_BODY, fill=PALETTE["ink"])
        y += 106

    draw.text((110, 780), f"规则：{rule}", font=FONT_SMALL, fill=PALETTE["cream"])
    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path)


def create_roadmap(chapters: list[dict], out_path: Path) -> None:
    img = make_canvas(PALETTE["cream"])
    draw = ImageDraw.Draw(img)
    draw.rectangle((0, 0, WIDTH, 130), fill=PALETTE["forest"])
    draw.text((70, 34), "全书路线图：24章如何从‘会看数字’走到‘先活下来’", font=FONT_H2, fill=PALETTE["cream"])
    add_texture(img)

    card_w = 330
    card_h = 145
    x0 = 70
    y0 = 170
    gap_x = 35
    gap_y = 28
    for idx, chapter in enumerate(chapters):
        row = idx // 4
        col = idx % 4
        x = x0 + col * (card_w + gap_x)
        y = y0 + row * (card_h + gap_y)
        fill = "#FBF7EC"
        rounded_panel(draw, (x, y, x + card_w, y + card_h), fill=fill, outline="#D3C29A", radius=24)
        draw.rectangle((x, y, x + card_w, y + 16), fill=stage_color(chapter["stage"]))
        draw.text((x + 18, y + 24), f"第{chapter['chapter_no']}章", font=FONT_SMALL, fill=PALETTE["forest"])
        draw.text((x + 18, y + 58), wrap_lines(chapter["title"], 12), font=FONT_SMALL, fill=PALETTE["ink"], spacing=6)
        draw.text((x + 18, y + 112), chapter["stage"], font=FONT_TINY, fill=PALETTE["slate"])
    out_path.parent.mkdir(parents=True, exist_ok=True)
    img.save(out_path)


def main() -> None:
    workspace = Path(__file__).resolve().parents[2]
    data_path = workspace / "data" / "metadata" / "book_content.json"
    payload = json.loads(data_path.read_text(encoding="utf-8"))

    assets_dir = workspace / "output" / "assets"
    theme_dir = assets_dir / "theme"
    cards_dir = assets_dir / "cards"
    appendix_dir = assets_dir / "appendix"

    payload["theme"] = generate_theme_backgrounds(theme_dir)

    for chapter in payload["chapters"]:
        out_path = cards_dir / f"chapter_{chapter['chapter_no']:02d}.png"
        create_card(
            title=chapter["title"],
            badge=f"{chapter['chapter_no']:02d}",
            subtitle=chapter["one_line"],
            points=chapter["core_points"],
            rule=chapter["rule"],
            stage=chapter["stage"],
            out_path=out_path,
        )
        chapter["concept_card"] = str(out_path)
        chapter["visual_image"] = chapter["chart_image"] or str(out_path)

    appendix_card = appendix_dir / "appendix_overview.png"
    create_card(
        title=payload["appendix"]["title"],
        badge="附录",
        subtitle=payload["appendix"]["one_line"],
        points=payload["appendix"]["core_points"],
        rule=payload["appendix"]["rule"],
        stage="专业化总结期",
        out_path=appendix_card,
    )
    payload["appendix"]["concept_card"] = str(appendix_card)
    payload["appendix"]["visual_image"] = payload["appendix"]["chart_images"][0] if payload["appendix"]["chart_images"] else str(appendix_card)

    roadmap = assets_dir / "book_roadmap.png"
    create_roadmap(payload["chapters"], roadmap)
    payload["meta"]["roadmap_image"] = str(roadmap)

    data_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Updated {data_path}")


if __name__ == "__main__":
    main()
