from __future__ import annotations

import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]
REPO_ROOT = PROJECT_ROOT.parent
sys.path.insert(0, str(REPO_ROOT / "scripts"))

from book_learning_pack_common import materialize_learning_pack


if __name__ == "__main__":
    materialize_learning_pack("principles", PROJECT_ROOT)
