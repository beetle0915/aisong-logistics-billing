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


def configure_utf8_stdio() -> None:
    """Prefer UTF-8 for this process when the host allows it."""
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name)
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            try:
                reconfigure(encoding="utf-8")
            except Exception:
                pass


def ascii_join(parts: list[object]) -> str:
    """Render command lines with ASCII-only escaping for safe CI logs."""
    return " ".join(ascii(part) for part in parts)


def run_command(command: list[str]) -> None:
    print("+", ascii_join(command))
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
        raise FileNotFoundError(f"Build output not found: {ascii(str(exe_path))}")
    return exe_path


def main() -> int:
    configure_utf8_stdio()
    exe_path = build_exe()
    print("Windows build complete:", ascii(str(exe_path)))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
