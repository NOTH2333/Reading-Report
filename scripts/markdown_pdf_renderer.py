from __future__ import annotations

import html
import io
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import fitz
from PIL import Image as PILImage
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
    ListFlowable,
    ListItem,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


HEADING_RE = re.compile(r"^(#{1,6})\s+(.*)$")
IMAGE_RE = re.compile(r"^\s*!\[(.*?)\]\((.+?)\)\s*$")
UNORDERED_RE = re.compile(r"^\s*-\s+(.*)$")
ORDERED_RE = re.compile(r"^\s*\d+\.\s+(.*)$")
QUOTE_RE = re.compile(r"^\s*>\s?(.*)$")
TABLE_ALIGN_RE = re.compile(r"^\s*\|?(?:\s*:?-{3,}:?\s*\|)+\s*:?-{3,}:?\s*\|?\s*$")
HR_RE = re.compile(r"^\s*([-*_])(?:\s*\1){2,}\s*$")


@dataclass
class MarkdownSource:
    path: Path
    title: str
    blocks: list[dict[str, Any]]


@dataclass
class ValidationIssue:
    source: Path
    message: str


class MarkdownPdfRenderer:
    def __init__(self) -> None:
        self.page_width, self.page_height = A4
        self.left_margin = 1.8 * cm
        self.right_margin = 1.8 * cm
        self.top_margin = 1.8 * cm
        self.bottom_margin = 1.8 * cm
        self.content_width = self.page_width - self.left_margin - self.right_margin
        self._font_regular, self._font_bold = self._ensure_fonts()
        self.styles = self._build_styles()
        self._asset_search_cache: dict[tuple[Path, str], Path | None] = {}

    def validate_sources(self, sources: list[Path], asset_base_dir: Path | None = None) -> list[ValidationIssue]:
        issues: list[ValidationIssue] = []
        for source in sources:
            if not source.exists():
                issues.append(ValidationIssue(source=source, message="Markdown 源文件不存在"))
                continue
            parsed = self._parse_source(source)
            for block in parsed.blocks:
                if block["type"] != "image":
                    continue
                resolved = self._resolve_image(block["src"], source.parent, asset_base_dir)
                if resolved is None:
                    issues.append(
                        ValidationIssue(
                            source=source,
                            message=f"无法解析图片路径: {block['src']}",
                        )
                    )
        return issues

    def render(
        self,
        *,
        sources: list[Path],
        output_path: Path,
        title: str,
        subtitle: str | None = None,
        asset_base_dir: Path | None = None,
    ) -> None:
        parsed_sources = [self._parse_source(path) for path in sources]
        story: list[Any] = []
        story.extend(self._build_cover(title, subtitle, parsed_sources))
        if len(parsed_sources) > 1:
            story.extend(self._build_source_index(parsed_sources))

        for index, source in enumerate(parsed_sources):
            if story and not isinstance(story[-1], PageBreak):
                story.append(PageBreak())
            story.extend(self._render_blocks(source, asset_base_dir=asset_base_dir))
            if index != len(parsed_sources) - 1:
                story.append(PageBreak())

        output_path.parent.mkdir(parents=True, exist_ok=True)
        document = SimpleDocTemplate(
            str(output_path),
            pagesize=A4,
            leftMargin=self.left_margin,
            rightMargin=self.right_margin,
            topMargin=self.top_margin,
            bottomMargin=self.bottom_margin,
            title=title,
            author="OpenAI Codex",
        )
        document.build(
            story,
            onFirstPage=lambda canvas, doc: self._decorate_page(canvas, doc, title),
            onLaterPages=lambda canvas, doc: self._decorate_page(canvas, doc, title),
        )

    def _ensure_fonts(self) -> tuple[str, str]:
        regular_name = "VaultSans"
        bold_name = "VaultSansBold"
        registered = set(pdfmetrics.getRegisteredFontNames())
        if regular_name in registered and bold_name in registered:
            return regular_name, bold_name

        regular_path = self._pick_font_path(
            [
                Path(r"C:\Windows\Fonts\msyh.ttc"),
                Path(r"C:\Windows\Fonts\simsun.ttc"),
                Path(r"C:\Windows\Fonts\simhei.ttf"),
            ]
        )
        bold_path = self._pick_font_path(
            [
                Path(r"C:\Windows\Fonts\msyhbd.ttc"),
                Path(r"C:\Windows\Fonts\simhei.ttf"),
                Path(r"C:\Windows\Fonts\msyh.ttc"),
            ]
        )
        pdfmetrics.registerFont(TTFont(regular_name, str(regular_path)))
        pdfmetrics.registerFont(TTFont(bold_name, str(bold_path)))
        return regular_name, bold_name

    def _pick_font_path(self, candidates: list[Path]) -> Path:
        for candidate in candidates:
            if candidate.exists():
                return candidate
        raise FileNotFoundError("未找到可用的中文字体文件")

    def _build_styles(self) -> dict[str, ParagraphStyle]:
        base_styles = getSampleStyleSheet()
        return {
            "title": ParagraphStyle(
                name="VaultTitle",
                parent=base_styles["Title"],
                fontName=self._font_bold,
                fontSize=23,
                leading=29,
                textColor=colors.HexColor("#1F3A5F"),
                alignment=TA_CENTER,
                spaceAfter=0.3 * cm,
            ),
            "subtitle": ParagraphStyle(
                name="VaultSubtitle",
                parent=base_styles["BodyText"],
                fontName=self._font_regular,
                fontSize=11,
                leading=16,
                textColor=colors.HexColor("#5F6B7A"),
                alignment=TA_CENTER,
                spaceAfter=0.2 * cm,
            ),
            "h1": ParagraphStyle(
                name="VaultH1",
                parent=base_styles["Heading1"],
                fontName=self._font_bold,
                fontSize=18,
                leading=24,
                textColor=colors.HexColor("#1F3A5F"),
                spaceBefore=0.55 * cm,
                spaceAfter=0.2 * cm,
            ),
            "h2": ParagraphStyle(
                name="VaultH2",
                parent=base_styles["Heading2"],
                fontName=self._font_bold,
                fontSize=14,
                leading=20,
                textColor=colors.HexColor("#315B7D"),
                spaceBefore=0.45 * cm,
                spaceAfter=0.14 * cm,
            ),
            "h3": ParagraphStyle(
                name="VaultH3",
                parent=base_styles["Heading3"],
                fontName=self._font_bold,
                fontSize=12,
                leading=17,
                textColor=colors.HexColor("#426A88"),
                spaceBefore=0.32 * cm,
                spaceAfter=0.1 * cm,
            ),
            "h4": ParagraphStyle(
                name="VaultH4",
                parent=base_styles["Heading4"],
                fontName=self._font_bold,
                fontSize=11,
                leading=16,
                textColor=colors.HexColor("#426A88"),
                spaceBefore=0.22 * cm,
                spaceAfter=0.08 * cm,
            ),
            "body": ParagraphStyle(
                name="VaultBody",
                parent=base_styles["BodyText"],
                fontName=self._font_regular,
                fontSize=10.8,
                leading=17,
                textColor=colors.HexColor("#243447"),
                spaceAfter=0.12 * cm,
            ),
            "quote": ParagraphStyle(
                name="VaultQuote",
                parent=base_styles["BodyText"],
                fontName=self._font_regular,
                fontSize=10.2,
                leading=16,
                textColor=colors.HexColor("#4F5F73"),
                leftIndent=0.55 * cm,
                borderPadding=0.2 * cm,
                borderColor=colors.HexColor("#B3C4D7"),
                borderWidth=1,
                borderLeft=True,
                backColor=colors.HexColor("#F5F8FB"),
                spaceAfter=0.14 * cm,
            ),
            "caption": ParagraphStyle(
                name="VaultCaption",
                parent=base_styles["BodyText"],
                fontName=self._font_regular,
                fontSize=8.8,
                leading=12,
                textColor=colors.HexColor("#657489"),
                alignment=TA_CENTER,
                spaceAfter=0.16 * cm,
            ),
            "toc": ParagraphStyle(
                name="VaultToc",
                parent=base_styles["BodyText"],
                fontName=self._font_regular,
                fontSize=10.4,
                leading=16,
                textColor=colors.HexColor("#2E475E"),
                leftIndent=0.2 * cm,
                spaceAfter=0.08 * cm,
            ),
        }

    def _build_cover(self, title: str, subtitle: str | None, parsed_sources: list[MarkdownSource]) -> list[Any]:
        story: list[Any] = [
            Spacer(1, 3.4 * cm),
            Paragraph(self._format_inline(title), self.styles["title"]),
        ]
        if subtitle:
            story.append(Paragraph(self._format_inline(subtitle), self.styles["subtitle"]))
        story.extend(
            [
                Spacer(1, 0.35 * cm),
                Paragraph(
                    self._format_inline(
                        f"来源文档数：{len(parsed_sources)} 份 | 统一易读版 PDF 导出"
                    ),
                    self.styles["subtitle"],
                ),
                Spacer(1, 0.9 * cm),
                HRFlowable(color=colors.HexColor("#D3DEE8"), width="80%"),
                Spacer(1, 0.45 * cm),
                Paragraph(
                    self._format_inline("本文件由库级共享导出器生成，保留标题层级、表格、列表与图片。"),
                    self.styles["subtitle"],
                ),
                PageBreak(),
            ]
        )
        return story

    def _build_source_index(self, parsed_sources: list[MarkdownSource]) -> list[Any]:
        items = [Paragraph(self._format_inline(source.title), self.styles["toc"]) for source in parsed_sources]
        flowable = ListFlowable(items, bulletType="bullet", leftIndent=0.35 * cm)
        return [
            Paragraph("章节目录", self.styles["h1"]),
            flowable,
            PageBreak(),
        ]

    def _render_blocks(self, source: MarkdownSource, asset_base_dir: Path | None) -> list[Any]:
        story: list[Any] = []
        for block in source.blocks:
            block_type = block["type"]
            if block_type == "heading":
                style = {1: "h1", 2: "h2", 3: "h3"}.get(block["level"], "h4")
                story.append(Paragraph(self._format_inline(block["text"]), self.styles[style]))
                continue
            if block_type == "paragraph":
                story.append(Paragraph(self._format_inline(block["text"]), self.styles["body"]))
                continue
            if block_type == "quote":
                for item in block["items"]:
                    story.append(Paragraph(self._format_inline(item), self.styles["quote"]))
                continue
            if block_type == "list":
                story.append(self._build_list(block))
                story.append(Spacer(1, 0.08 * cm))
                continue
            if block_type == "table":
                story.append(self._build_table(block["rows"]))
                story.append(Spacer(1, 0.18 * cm))
                continue
            if block_type == "image":
                image_path = self._resolve_image(block["src"], source.path.parent, asset_base_dir)
                if image_path is None:
                    story.append(
                        Paragraph(
                            self._format_inline(f"[缺失图片] {block['src']}"),
                            self.styles["caption"],
                        )
                    )
                    continue
                story.append(self._build_image_flowable(image_path))
                caption = block["alt"].strip()
                if caption:
                    story.append(Paragraph(self._format_inline(caption), self.styles["caption"]))
                continue
            if block_type == "hr":
                story.append(HRFlowable(color=colors.HexColor("#D6DFE7"), width="100%"))
                story.append(Spacer(1, 0.16 * cm))
        return story

    def _build_list(self, block: dict[str, Any]) -> ListFlowable:
        items = [
            ListItem(Paragraph(self._format_inline(item), self.styles["body"]))
            for item in block["items"]
        ]
        bullet_type = "1" if block["ordered"] else "bullet"
        return ListFlowable(items, bulletType=bullet_type, leftIndent=0.45 * cm)

    def _build_table(self, rows: list[list[str]]) -> Table:
        column_count = max(len(row) for row in rows)
        normalized_rows: list[list[Any]] = []
        column_weights = [1] * column_count
        for row in rows:
            padded = row + [""] * (column_count - len(row))
            normalized_rows.append(
                [Paragraph(self._format_inline(cell), self.styles["body"]) for cell in padded]
            )
            for index, cell in enumerate(padded):
                column_weights[index] = max(column_weights[index], max(len(cell), 6))

        total_weight = sum(column_weights)
        col_widths = [self.content_width * weight / total_weight for weight in column_weights]
        table = Table(normalized_rows, colWidths=col_widths, repeatRows=1)
        table.setStyle(
            TableStyle(
                [
                    ("FONTNAME", (0, 0), (-1, 0), self._font_bold),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF0F6")),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
                    ("GRID", (0, 0), (-1, -1), 0.55, colors.HexColor("#D2DCE6")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 6),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                    ("TOPPADDING", (0, 0), (-1, -1), 5),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                ]
            )
        )
        return table

    def _build_image_flowable(self, image_path: Path) -> RLImage:
        suffix = image_path.suffix.lower()
        if suffix == ".svg":
            doc = fitz.open(image_path)
            page = doc.load_page(0)
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            buffer = io.BytesIO(pix.tobytes("png"))
            width_px, height_px = pix.width, pix.height
            image = RLImage(buffer)
            image._source_buffer = buffer
        else:
            width_px, height_px = self._read_image_size(image_path)
            image = RLImage(str(image_path))

        width_pt = width_px * 72 / 96
        height_pt = height_px * 72 / 96
        max_width = self.content_width
        max_height = 17.5 * cm
        scale = min(max_width / width_pt, max_height / height_pt, 1.0)
        image.drawWidth = width_pt * scale
        image.drawHeight = height_pt * scale
        image.hAlign = "CENTER"
        return image

    def _read_image_size(self, image_path: Path) -> tuple[int, int]:
        with PILImage.open(image_path) as image:
            return image.size

    def _resolve_image(
        self,
        raw_src: str,
        source_dir: Path,
        asset_base_dir: Path | None,
    ) -> Path | None:
        cleaned = raw_src.strip().strip("<>").split("#", 1)[0]
        direct = (source_dir / cleaned).resolve()
        if direct.exists():
            return direct

        if asset_base_dir is None:
            return None

        joined = (asset_base_dir / cleaned).resolve()
        if joined.exists():
            return joined

        cache_key = (asset_base_dir.resolve(), Path(cleaned).name)
        if cache_key in self._asset_search_cache:
            return self._asset_search_cache[cache_key]

        matches = list(asset_base_dir.rglob(Path(cleaned).name))
        if len(matches) == 1:
            self._asset_search_cache[cache_key] = matches[0]
            return matches[0]
        if matches:
            matches.sort(key=lambda path: (len(path.parts), len(str(path))))
            self._asset_search_cache[cache_key] = matches[0]
            return matches[0]

        self._asset_search_cache[cache_key] = None
        return None

    def _decorate_page(self, canvas, doc, title: str) -> None:
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#D3DEE8"))
        canvas.line(self.left_margin, self.page_height - 1.2 * cm, self.page_width - self.right_margin, self.page_height - 1.2 * cm)
        canvas.line(self.left_margin, 1.2 * cm, self.page_width - self.right_margin, 1.2 * cm)
        canvas.setFont(self._font_regular, 8.5)
        canvas.setFillColor(colors.HexColor("#6A7888"))
        canvas.drawString(self.left_margin, 0.75 * cm, title[:50])
        canvas.drawRightString(self.page_width - self.right_margin, 0.75 * cm, str(doc.page))
        canvas.restoreState()

    def _parse_source(self, path: Path) -> MarkdownSource:
        text = path.read_text(encoding="utf-8")
        blocks = self._parse_markdown(text)
        title = next(
            (block["text"] for block in blocks if block["type"] == "heading"),
            path.stem,
        )
        return MarkdownSource(path=path, title=title, blocks=blocks)

    def _parse_markdown(self, text: str) -> list[dict[str, Any]]:
        lines = text.splitlines()
        blocks: list[dict[str, Any]] = []
        index = 0
        while index < len(lines):
            raw_line = lines[index]
            line = raw_line.rstrip()
            stripped = line.strip()

            if not stripped:
                index += 1
                continue

            heading_match = HEADING_RE.match(stripped)
            if heading_match:
                blocks.append(
                    {
                        "type": "heading",
                        "level": len(heading_match.group(1)),
                        "text": heading_match.group(2).strip(),
                    }
                )
                index += 1
                continue

            image_match = IMAGE_RE.match(stripped)
            if image_match:
                blocks.append(
                    {
                        "type": "image",
                        "alt": image_match.group(1),
                        "src": image_match.group(2).strip(),
                    }
                )
                index += 1
                continue

            if self._is_table_start(lines, index):
                table_lines = [lines[index]]
                index += 2
                while index < len(lines):
                    candidate = lines[index].rstrip()
                    if not candidate.strip() or "|" not in candidate:
                        break
                    table_lines.append(candidate)
                    index += 1
                blocks.append({"type": "table", "rows": self._parse_table(table_lines)})
                continue

            if HR_RE.match(stripped):
                blocks.append({"type": "hr"})
                index += 1
                continue

            if QUOTE_RE.match(stripped):
                items: list[str] = []
                while index < len(lines):
                    quote_line = lines[index].rstrip()
                    match = QUOTE_RE.match(quote_line.strip())
                    if not match:
                        break
                    items.append(match.group(1).strip())
                    index += 1
                blocks.append({"type": "quote", "items": items})
                continue

            if UNORDERED_RE.match(stripped) or ORDERED_RE.match(stripped):
                ordered = bool(ORDERED_RE.match(stripped))
                item_regex = ORDERED_RE if ordered else UNORDERED_RE
                items: list[str] = []
                while index < len(lines):
                    candidate = lines[index].rstrip()
                    candidate_stripped = candidate.strip()
                    if not candidate_stripped:
                        break
                    match = item_regex.match(candidate_stripped)
                    if match:
                        items.append(match.group(1).strip())
                        index += 1
                        continue
                    if items and not self._is_structural_line(candidate_stripped):
                        items[-1] = f"{items[-1]} {candidate_stripped}"
                        index += 1
                        continue
                    break
                blocks.append(
                    {
                        "type": "list",
                        "ordered": ordered,
                        "items": [self._join_text_lines([item]) for item in items],
                    }
                )
                continue

            paragraph_lines = [stripped]
            index += 1
            while index < len(lines):
                candidate = lines[index].rstrip()
                candidate_stripped = candidate.strip()
                if not candidate_stripped or self._is_structural_line(candidate_stripped) or self._is_table_start(lines, index):
                    break
                paragraph_lines.append(candidate_stripped)
                index += 1
            blocks.append({"type": "paragraph", "text": self._join_text_lines(paragraph_lines)})

        return blocks

    def _is_table_start(self, lines: list[str], index: int) -> bool:
        if index + 1 >= len(lines):
            return False
        first = lines[index].strip()
        second = lines[index + 1].strip()
        return "|" in first and TABLE_ALIGN_RE.match(second) is not None

    def _parse_table(self, lines: list[str]) -> list[list[str]]:
        rows: list[list[str]] = []
        for line in lines:
            stripped = line.strip().strip("|")
            cells = [cell.strip() for cell in stripped.split("|")]
            rows.append(cells)
        return rows

    def _is_structural_line(self, stripped: str) -> bool:
        return any(
            [
                HEADING_RE.match(stripped),
                IMAGE_RE.match(stripped),
                QUOTE_RE.match(stripped),
                UNORDERED_RE.match(stripped),
                ORDERED_RE.match(stripped),
                HR_RE.match(stripped),
            ]
        )

    def _join_text_lines(self, lines: list[str]) -> str:
        text = " ".join(line.strip() for line in lines if line.strip())
        text = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[\u4e00-\u9fff])", "", text)
        text = re.sub(r"(?<=[\u4e00-\u9fff])\s+(?=[，。；：！？])", "", text)
        text = re.sub(r"(?<=[（])\s+", "", text)
        text = re.sub(r"\s+(?=[）])", "", text)
        return text.strip()

    def _format_inline(self, text: str) -> str:
        text = html.escape(text)
        placeholders: dict[str, str] = {}

        def stash(replacement: str) -> str:
            key = f"__MARK_{len(placeholders)}__"
            placeholders[key] = replacement
            return key

        text = re.sub(r"\[(.*?)\]\((.*?)\)", lambda m: stash(html.escape(m.group(1))), text)
        text = re.sub(
            r"`([^`]+)`",
            lambda m: stash(f'<font face="Courier">{m.group(1)}</font>'),
            text,
        )
        text = re.sub(r"\*\*(.+?)\*\*", lambda m: stash(f"<b>{m.group(1)}</b>"), text)
        text = re.sub(r"__(.+?)__", lambda m: stash(f"<b>{m.group(1)}</b>"), text)
        text = re.sub(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)", lambda m: stash(f"<i>{m.group(1)}</i>"), text)

        for key, value in placeholders.items():
            text = text.replace(key, value)
        return text
