#!/usr/bin/env python3
"""Build the Windows executable with PyInstaller."""

from __future__ import annotations

import shutil
import subprocess
import sys
from pathlib import Path


APP_NAME = "艾松物流计费系统"
ROOT_DIR = Path(__file__).resolve().parent
ENTRY_FILE = ROOT_DIR / "windows_entry.py"
BUILD_DIR = ROOT_DIR / "build" / "windows"
DIST_DIR = ROOT_DIR / "dist" / "windows"
PACKAGE_DIR = DIST_DIR / APP_NAME


def run_command(command: list[str]) -> None:
    print("+", " ".join(command))
    subprocess.run(command, cwd=ROOT_DIR, check=True)


def clean_output_dirs() -> None:
    for path in [BUILD_DIR, DIST_DIR]:
        if path.exists():
            shutil.rmtree(path)
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    DIST_DIR.mkdir(parents=True, exist_ok=True)


def build_exe() -> Path:
    clean_output_dirs()
    run_command(
        [
            sys.executable,
            "-m",
            "PyInstaller",
            "--noconfirm",
            "--clean",
            "--windowed",
            "--name",
            APP_NAME,
            "--distpath",
            str(DIST_DIR),
            "--workpath",
            str(BUILD_DIR),
            "--specpath",
            str(BUILD_DIR),
            str(ENTRY_FILE),
        ]
    )
    exe_path = PACKAGE_DIR / f"{APP_NAME}.exe"
    if not exe_path.exists():
        raise FileNotFoundError(f"未找到打包结果：{exe_path}")
    return exe_path


def main() -> int:
    exe_path = build_exe()
    print(f"已生成 Windows 程序：{exe_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
