from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


ROOT = Path(__file__).resolve().parent.parent


@dataclass(frozen=True)
class BookPdfJob:
    book_id: str
    title: str
    report_sources: tuple[Path, ...]
    report_output: Path
    summary_sources: tuple[Path, ...]
    summary_output: Path
    asset_base_dir: Path
    summary_mode: str
    skip_if_exists: bool = False


stock_root = ROOT / "01_股票大作手回忆录__GU_PIAO_DA_ZUO_SHOU"
technical_root = ROOT / "02_期货市场技术分析__Technical_Analysis_of_Futures_Markets"
psychology_root = ROOT / "03_交易心理分析__Trading_Psychology_Analysis"
munger_root = ROOT / "04_穷查理宝典__The_Wisdom_and_Maxims_of_Charlie_Munger"
volume_root = ROOT / "05_量价分析__Volume_Price_Analysis" / "《量价分析》教学化交付包"
chengming_root = ROOT / "06_澄明之境__CHENG_MING_ZHI_JING"
principles_root = ROOT / "07_原则__Principles"

technical_summary_sources = tuple(
    sorted((technical_root / "03_章节关键知识总结").glob("*/*.md"))
)

psychology_summary_dir = psychology_root / "deliverables" / "summaries"
psychology_summary_sources = (
    psychology_summary_dir / "00-章节索引.md",
    *tuple(
        sorted(
            path
            for path in psychology_summary_dir.glob("*.md")
            if path.name != "00-章节索引.md"
        )
    ),
)


BOOK_JOBS: tuple[BookPdfJob, ...] = (
    BookPdfJob(
        book_id="stock_operator",
        title="《股票大作手回忆录》",
        report_sources=(stock_root / "output" / "md" / "阅读报告.md",),
        report_output=stock_root / "output" / "md" / "阅读报告.pdf",
        summary_sources=(stock_root / "output" / "md" / "关键知识总结.md",),
        summary_output=stock_root / "output" / "md" / "关键知识总结.pdf",
        asset_base_dir=stock_root / "output" / "assets",
        summary_mode="single",
    ),
    BookPdfJob(
        book_id="technical_analysis",
        title="《期货市场技术分析》",
        report_sources=(technical_root / "01_全书阅读报告" / "期货市场技术分析_全书阅读报告.md",),
        report_output=technical_root / "01_全书阅读报告" / "期货市场技术分析_全书阅读报告.pdf",
        summary_sources=technical_summary_sources,
        summary_output=technical_root / "03_章节关键知识总结" / "期货市场技术分析_按章节关键知识总结.pdf",
        asset_base_dir=technical_root,
        summary_mode="concat",
    ),
    BookPdfJob(
        book_id="trading_psychology",
        title="《交易心理分析》",
        report_sources=(psychology_root / "deliverables" / "reports" / "阅读报告.md",),
        report_output=psychology_root / "deliverables" / "reports" / "阅读报告.pdf",
        summary_sources=psychology_summary_sources,
        summary_output=psychology_root / "deliverables" / "summaries" / "关键知识总结.pdf",
        asset_base_dir=psychology_root / "deliverables" / "assets",
        summary_mode="concat",
    ),
    BookPdfJob(
        book_id="charlie_munger",
        title="《穷查理宝典》",
        report_sources=(munger_root / "01_阅读报告" / "穷查理宝典_阅读报告_儿童可读版.md",),
        report_output=munger_root / "01_阅读报告" / "穷查理宝典_阅读报告_儿童可读版.pdf",
        summary_sources=(munger_root / "02_关键知识总结" / "穷查理宝典_关键知识总结_按章节.md",),
        summary_output=munger_root / "02_关键知识总结" / "穷查理宝典_关键知识总结_按章节.pdf",
        asset_base_dir=munger_root,
        summary_mode="single",
    ),
    BookPdfJob(
        book_id="volume_price_analysis",
        title="《量价分析》",
        report_sources=(volume_root / "02_阅读报告" / "《量价分析》全书阅读报告.md",),
        report_output=volume_root / "02_阅读报告" / "《量价分析》全书阅读报告.pdf",
        summary_sources=(volume_root / "03_关键知识总结" / "《量价分析》按章节关键知识总结.md",),
        summary_output=volume_root / "03_关键知识总结" / "《量价分析》按章节关键知识总结.pdf",
        asset_base_dir=volume_root,
        summary_mode="single",
        skip_if_exists=True,
    ),
    BookPdfJob(
        book_id="chengming_zhijing",
        title="《澄明之境》",
        report_sources=(chengming_root / "01_阅读报告" / "澄明之境_阅读报告_儿童可读版.md",),
        report_output=chengming_root / "01_阅读报告" / "澄明之境_阅读报告_儿童可读版.pdf",
        summary_sources=(chengming_root / "02_关键知识总结" / "澄明之境_关键知识总结_按章节.md",),
        summary_output=chengming_root / "02_关键知识总结" / "澄明之境_关键知识总结_按章节.pdf",
        asset_base_dir=chengming_root,
        summary_mode="single",
    ),
    BookPdfJob(
        book_id="principles",
        title="《原则》",
        report_sources=(principles_root / "01_阅读报告" / "原则_阅读报告_儿童可读版.md",),
        report_output=principles_root / "01_阅读报告" / "原则_阅读报告_儿童可读版.pdf",
        summary_sources=(principles_root / "02_关键知识总结" / "原则_关键知识总结_按章节.md",),
        summary_output=principles_root / "02_关键知识总结" / "原则_关键知识总结_按章节.pdf",
        asset_base_dir=principles_root,
        summary_mode="single",
    ),
)
