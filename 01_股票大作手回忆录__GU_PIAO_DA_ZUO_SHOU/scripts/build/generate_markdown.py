from __future__ import annotations

import json
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]


def load_json(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def rel(from_dir: Path, target: str | None) -> str | None:
    if not target:
        return None
    return Path(target).resolve().relative_to(ROOT.resolve()).as_posix()


def md_rel(from_dir: Path, target: str | None) -> str | None:
    if not target:
        return None
    return Path(os.path.relpath(Path(target).resolve(), start=from_dir.resolve())).as_posix()


def write_report(payload: dict, out_path: Path) -> None:
    md_dir = out_path.parent
    lines: list[str] = []
    meta = payload["meta"]

    lines.append("# 《股票大作手回忆录》详细阅读报告")
    lines.append("")
    lines.append(f"- 版本依据：{meta['source_edition']}")
    lines.append(f"- 面向读者：{meta['audience']}")
    lines.append(f"- 写作方式：{meta['tone']}")
    lines.append("")
    lines.append("## 摘要")
    lines.append("")
    for item in meta["summary"]:
        lines.append(f"- {item}")
    lines.append("")
    lines.append("## 为什么今天还值得读")
    lines.append("")
    lines.append("这本书真正讲的不是‘如何靠一条消息暴富’，而是交易者如何在反复犯错、反复赔钱、反复重建中，慢慢逼近一套更可靠的市场行为规则。它之所以到今天仍然重要，不是因为100多年前的品种和制度还能原样照搬，而是因为书中反复出现的人性问题几乎没有变：人会急、会贪、会怕、会犟、会想省掉等待和验证这一步。")
    lines.append("")
    lines.append("对零基础读者来说，这本书最大的价值有三层。第一层，它让你看到市场里的输赢并不只由方向决定，还由执行、时机、仓位和环境决定。第二层，它把‘消息’和‘价格’分开，让你理解为什么很多人不是输给市场，而是输给自己愿意听的故事。第三层，它把顺势、纪律、独立判断和风险控制放在同一条逻辑线上，而不是把它们当成互不相干的口号。")
    lines.append("")
    lines.append("## 人物与时代背景")
    lines.append("")
    lines.append("书中的主人公以拉里·利文斯顿为化名，对应的是杰西·劳雷斯顿·利弗莫尔。丁圣元注疏版的重要贡献，在于给原著补上了历史图表和文献背景，使读者不只是在听一段传奇故事，而是在看一个交易方法如何随着市场环境、个人错误和现实惩罚不断被修正。注疏版也帮助读者看清：利弗莫尔真正值得学的不是‘神准’，而是他怎样一次次从错误中提炼规则。")
    lines.append("")
    roadmap = md_rel(md_dir, payload["meta"]["roadmap_image"])
    if roadmap:
        lines.append(f"![全书路线图]({roadmap})")
        lines.append("")
    lines.append("## 全书反复出现的总规律")
    lines.append("")
    for item in meta["global_takeaways"]:
        lines.append(f"- {item}")
    lines.append("")
    lines.append("## 基础术语表")
    lines.append("")
    lines.append("| 术语 | 白话解释 |")
    lines.append("| --- | --- |")
    for term in meta["glossary"]:
        lines.append(f"| {term['term']} | {term['plain']} |")
    lines.append("")

    for chapter in payload["chapters"]:
        lines.append(f"## 第{chapter['chapter_no']}章：{chapter['title']}")
        lines.append("")
        lines.append(f"- 阶段：{chapter['stage']}")
        lines.append(f"- 页码：注疏版 p.{chapter['start_page']}-{chapter['end_page']}")
        lines.append(f"- 一句话：{chapter['one_line']}")
        lines.append("")
        concept = md_rel(md_dir, chapter["concept_card"])
        if concept:
            lines.append(f"![第{chapter['chapter_no']}章记忆图]({concept})")
            lines.append("")
        lines.append("### 发生了什么")
        lines.append("")
        lines.append(
            f"本章的故事线可以概括为：{chapter['storyline'][0]} 接着，{chapter['storyline'][1]} 然后，{chapter['storyline'][2]} 最后，{chapter['storyline'][3]} 这也是为什么本章虽然是人物故事，但本质上更像一堂关于交易流程升级的课程。"
        )
        lines.append("")
        lines.append("### 市场在发生什么")
        lines.append("")
        for item in chapter["market_context"]:
            lines.append(f"- {item}")
        lines.append("")
        lines.append("### 这章里最重要的三个动作")
        lines.append("")
        for trade in chapter["key_trades"]:
            lines.append(f"- **{trade['title']}**：{trade['detail']}")
        lines.append("")
        lines.append("### 他做对了什么")
        lines.append("")
        for item in chapter["strengths"]:
            lines.append(f"- {item}")
        lines.append("")
        lines.append("### 他做错了什么")
        lines.append("")
        for item in chapter["mistakes"]:
            lines.append(f"- {item}")
        lines.append("")
        lines.append("### 这一章真正要学会的原则")
        lines.append("")
        lines.append(f"- 交易原则：{chapter['principle']}")
        lines.append(f"- 心理提醒：{chapter['psychology']}")
        lines.append(f"- 风控提醒：{chapter['risk']}")
        lines.append("")
        lines.append("### 最容易误解什么")
        lines.append("")
        lines.append(f"- 典型误区：{chapter['misconception']}")
        lines.append(f"- 更好的记忆方式：{chapter['analogy']}")
        lines.append("")
        lines.append("### 如果把它翻成今天的新手场景")
        lines.append("")
        lines.append(chapter["modern"])
        lines.append("")
        for item in chapter["actions"]:
            lines.append(f"- {item}")
        lines.append("")
        lines.append("### 配图说明")
        lines.append("")
        lines.append(f"- 本章配图主旨：{chapter['chart_takeaway']}")
        if chapter["chart_image"]:
            chart = md_rel(md_dir, chapter["chart_image"])
            lines.append(f"- 书内图表/插图页段参考：p.{chapter['start_page']}-{chapter['end_page']}")
            if chart:
                lines.append("")
                lines.append(f"![第{chapter['chapter_no']}章书内图表]({chart})")
                lines.append("")
        else:
            lines.append("- 本章未直接使用书内图表，改用自制记忆图承接内容。")
            lines.append("")
        lines.append("### 本章结论")
        lines.append("")
        lines.append(
            f"这章最值得带走的一句话是“{chapter['rule']}”。如果把本章和全书其他章节放在一起看，它所贡献的并不是一条孤立技巧，而是对‘{chapter['core_points'][0]}、{chapter['core_points'][1]}、{chapter['core_points'][2]}’的进一步确认。"
        )
        lines.append("")

    appendix = payload["appendix"]
    appendix_card = md_rel(md_dir, appendix["concept_card"])
    lines.append("## 附录精华")
    lines.append("")
    lines.append(f"- 页码：注疏版 p.{appendix['start_page']}-{appendix['end_page']}")
    lines.append(f"- 一句话：{appendix['one_line']}")
    lines.append("")
    if appendix_card:
        lines.append(f"![附录总览]({appendix_card})")
        lines.append("")
    lines.append("### 附录一：为什么年表重要")
    lines.append("")
    lines.append("年表把利弗莫尔的多次大起大落按时间拉直，让读者看到：高手不是一路顺风，而是在不同环境中不断升级规则。用年表来看正文，会更容易理解为什么他早年的错误总围绕‘急、频繁、迷信局部波动’，而成熟后的关键词逐渐变成‘大势、时机、持有、独立判断’。")
    lines.append("")
    lines.append("### 附录二：交易规则的浓缩表达")
    lines.append("")
    lines.append("- 买上涨中的股票，卖下跌中的股票。")
    lines.append("- 市场没有明确趋势时，不要天天交易。")
    lines.append("- 等市场证明你的观点后再出手。")
    lines.append("- 交易有利润时继续持有，交易有亏损时从速了结。")
    lines.append("- 绝不摊平亏损，绝不追加保证金。")
    lines.append("- 市场永远比个人看法更有裁决权。")
    lines.append("")
    lines.append("### 附录三：最小阻力路线的现代解释")
    lines.append("")
    lines.append("最小阻力路线并不神秘。它的核心意思是：在当前的力量对比下，市场最容易继续往哪边走，你就优先尊重哪边。价格上升意味着买方之间的竞争更强，价格下降意味着卖方压力更大。趋势不是交易者强行制造出来的，而是竞争循环自然推出来的结果。")
    lines.append("")
    lines.append("## 给零基础读者的最终行动建议")
    lines.append("")
    lines.append("如果你读完整本书，只想带走最实用的做法，可以从下面五条开始。第一，任何时候先分清楚自己是在押小波动，还是在押大趋势。第二，没有确认就别急着重仓。第三，任何建议、贴士和权威解释都要让价格来盖章。第四，把不摊平、不嘴硬、不乱补钱写成硬规则。第五，真正把你带离新手状态的，不是某一次神操作，而是你有没有能力在大错之后把规则升级。")
    lines.append("")

    out_path.write_text("\n".join(lines), encoding="utf-8")


def write_summary(payload: dict, out_path: Path) -> None:
    md_dir = out_path.parent
    lines: list[str] = []
    lines.append("# 《股票大作手回忆录》关键知识总结")
    lines.append("")
    lines.append("这份文档按章节整理成复习卡，适合快速回看。")
    lines.append("")

    for chapter in payload["chapters"]:
        lines.append(f"## 第{chapter['chapter_no']}章：{chapter['title']}")
        lines.append("")
        lines.append(f"- 页码：注疏版 p.{chapter['start_page']}-{chapter['end_page']}")
        concept = md_rel(md_dir, chapter["concept_card"])
        if concept:
            lines.append(f"![第{chapter['chapter_no']}章卡片]({concept})")
            lines.append("")
        lines.append("### 3 个核心点")
        lines.append("")
        for item in chapter["core_points"]:
            lines.append(f"- {item}")
        lines.append("")
        lines.append("### 典型误区")
        lines.append("")
        lines.append(f"- {chapter['misconception']}")
        lines.append("")
        lines.append("### 1 条可执行规则")
        lines.append("")
        lines.append(f"- {chapter['rule']}")
        lines.append("")
        lines.append("### 1 个类比/例子")
        lines.append("")
        lines.append(f"- {chapter['analogy']}")
        lines.append("")
        lines.append("### 术语速记")
        lines.append("")
        for term in chapter["terms"]:
            lines.append(f"- **{term['name']}**：{term['plain']}")
        lines.append("")
        lines.append("### 行动清单")
        lines.append("")
        for item in chapter["actions"]:
            lines.append(f"- {item}")
        lines.append("")

    appendix = payload["appendix"]
    appendix_card = md_rel(md_dir, appendix["concept_card"])
    lines.append("## 附录：全书压缩包")
    lines.append("")
    if appendix_card:
        lines.append(f"![附录卡片]({appendix_card})")
        lines.append("")
    for item in appendix["core_points"]:
        lines.append(f"- {item}")
    lines.append("")
    lines.append("- 最值得立即落地的规则：等市场证明、让盈利跑、快认错、不摊平。")
    lines.append("- 最值得反复理解的概念：最小阻力路线，也就是趋势方向。")
    lines.append("")

    out_path.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    payload = load_json(ROOT / "data" / "metadata" / "book_content.json")
    out_dir = ROOT / "output" / "md"
    out_dir.mkdir(parents=True, exist_ok=True)
    write_report(payload, out_dir / "阅读报告.md")
    write_summary(payload, out_dir / "关键知识总结.md")
    print(f"Wrote markdown files to {out_dir}")


if __name__ == "__main__":
    main()
