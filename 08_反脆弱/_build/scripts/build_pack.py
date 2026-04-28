from pathlib import Path
import sys

BOOK_ROOT = Path(__file__).resolve().parents[2]
REPO = BOOK_ROOT.parent
sys.path.insert(0, str(REPO / 'scripts'))

from teaching_pack_builder import build_one

if __name__ == '__main__':
    build_one('antifragile')
