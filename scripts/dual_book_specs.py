from __future__ import annotations

from book_specs_antifragile import ANTIFRAGILE_PROJECT
from book_specs_blackswan import BLACK_SWAN_PROJECT
from book_specs_chengming import CHENGMING_PROJECT
from book_specs_common import palette_for
from book_specs_principles import PRINCIPLES_PROJECT


BOOK_PROJECTS = {
    "chengming_zhijing": CHENGMING_PROJECT,
    "principles": PRINCIPLES_PROJECT,
    "antifragile": ANTIFRAGILE_PROJECT,
    "black_swan": BLACK_SWAN_PROJECT,
}


__all__ = ["BOOK_PROJECTS", "palette_for"]
