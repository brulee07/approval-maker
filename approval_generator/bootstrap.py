"""설치형 Python과 Windows 앱 번들 Python의 Tcl 경로를 함께 지원한다."""
import os
import shutil
import sys
from pathlib import Path


def prepare_tk():
    # WindowsApps 안의 Tcl 스크립트는 데스크톱 프로세스의 Tcl 파일 API에서
    # 직접 읽히지 않는 환경이 있다. 프로그램 폴더 안의 로컬 사본을 사용한다.
    if sys.platform != "win32" or "windowsapps" not in sys.base_prefix.lower():
        return
    source = Path(sys.base_prefix) / "tcl"
    destination = Path(__file__).resolve().parent.parent / ".runtime" / "tcl"
    destination.parent.mkdir(parents=True, exist_ok=True)
    (destination.parent / "python-path.txt").write_text(sys.executable, encoding="utf-8")
    marker = destination / "runtime-source.txt"
    if not marker.is_file() or marker.read_text(encoding="utf-8") != str(source):
        shutil.copytree(source, destination, dirs_exist_ok=True)
        marker.write_text(str(source), encoding="utf-8")
    for variable, pattern, script in [("TCL_LIBRARY", "tcl8.*", "init.tcl"), ("TK_LIBRARY", "tk8.*", "tk.tcl")]:
        folders = sorted(folder for folder in destination.glob(pattern) if (folder / script).is_file())
        if folders:
            os.environ[variable] = str(folders[-1])
