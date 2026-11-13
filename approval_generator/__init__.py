"""견적서에서 사용자 품의서 양식으로 변환하는 로컬 프로그램."""

import sys
from pathlib import Path

_local_packages = Path(__file__).resolve().parent.parent / ".runtime" / "site-packages"
if _local_packages.is_dir():
    sys.path.insert(0, str(_local_packages))

__version__ = "0.2.4"
