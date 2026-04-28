from __future__ import annotations

from copy import deepcopy


PALETTES = [
    {"primary": "#0E7490", "secondary": "#67E8F9", "accent": "#155E75", "bg": "#ECFEFF", "dark": "#0F172A"},
    {"primary": "#1D4ED8", "secondary": "#93C5FD", "accent": "#1E40AF", "bg": "#EFF6FF", "dark": "#172554"},
    {"primary": "#047857", "secondary": "#6EE7B7", "accent": "#065F46", "bg": "#ECFDF5", "dark": "#022C22"},
    {"primary": "#B45309", "secondary": "#FCD34D", "accent": "#92400E", "bg": "#FFFBEB", "dark": "#451A03"},
    {"primary": "#7C3AED", "secondary": "#C4B5FD", "accent": "#6D28D9", "bg": "#F5F3FF", "dark": "#2E1065"},
    {"primary": "#BE123C", "secondary": "#FDA4AF", "accent": "#9F1239", "bg": "#FFF1F2", "dark": "#4C0519"},
]


def palette_for(index: int) -> dict:
    return deepcopy(PALETTES[(index - 1) % len(PALETTES)])


def chapter_rows(rows: list[tuple]) -> list[dict]:
    output: list[dict] = []
    for sequence_no, source_part, original_label, title, start_page, end_page, opener_page, reference_pages in rows:
        output.append(
            {
                "id": f"ch{sequence_no:02d}",
                "sequence_no": f"{sequence_no:02d}",
                "source_part": source_part,
                "original_label": original_label,
                "title": title,
                "full_title": f"第{sequence_no:02d}章 {title}",
                "start_page": start_page,
                "end_page": end_page,
                "opener_page": opener_page,
                "reference_pages": list(reference_pages),
            }
        )
    return output


def story(title: str, summary: str, lesson: str) -> dict:
    return {"title": title, "summary": summary, "lesson": lesson}


def concept(name: str, explain: str) -> dict:
    return {"name": name, "explain": explain}


def chapter_content(
    *,
    core_message: str,
    position_in_book: str,
    children_question: str,
    key_concepts: list[tuple[str, str]],
    must_remember: list[str],
    child_example: str,
    investment_meaning: str,
    misunderstandings: list[str],
    child_actions: list[str],
    one_line_review: str,
    stories: list[dict],
) -> dict:
    return {
        "core_message": core_message,
        "position_in_book": position_in_book,
        "children_question": children_question,
        "key_concepts": [concept(name, explain) for name, explain in key_concepts],
        "must_remember": must_remember,
        "child_example": child_example,
        "investment_meaning": investment_meaning,
        "misunderstandings": misunderstandings,
        "child_actions": child_actions,
        "one_line_review": one_line_review,
        "stories": stories,
    }
