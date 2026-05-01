#!/usr/bin/env python3
"""Build a lightweight macOS .app wrapper for the express fee GUI."""

from __future__ import annotations

import plistlib
import shutil
import stat
import sys
from pathlib import Path


APP_NAME = "艾松物流计费系统"
APP_VERSION = "7.0.1"
EXECUTABLE_NAME = "ExpressFeeCalculator"
BUNDLE_IDENTIFIER = "com.local.express-fee-calculator"
DEFAULT_PYTHON = Path(
    "/Users/beetle/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
)


def build_app_bundle(app_dir: Path) -> Path:
    app_dir = app_dir.resolve()
    express_root = app_dir.parent
    bundle_dir = express_root / f"{APP_NAME}.app"
    contents_dir = bundle_dir / "Contents"
    macos_dir = contents_dir / "MacOS"
    resources_dir = contents_dir / "Resources"

    macos_dir.mkdir(parents=True, exist_ok=True)
    resources_dir.mkdir(parents=True, exist_ok=True)
    copy_app_source(app_dir, resources_dir / "app")

    write_info_plist(contents_dir / "Info.plist")
    write_pkg_info(contents_dir / "PkgInfo")
    write_launcher(macos_dir / EXECUTABLE_NAME)
    return bundle_dir


def copy_app_source(app_dir: Path, bundle_app_dir: Path) -> None:
    if bundle_app_dir.exists():
        shutil.rmtree(bundle_app_dir)
    bundle_app_dir.mkdir(parents=True, exist_ok=True)
    shutil.copytree(
        app_dir / "express_app",
        bundle_app_dir / "express_app",
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
    )


def write_info_plist(path: Path) -> None:
    info = {
        "CFBundleName": APP_NAME,
        "CFBundleDisplayName": APP_NAME,
        "CFBundleExecutable": EXECUTABLE_NAME,
        "CFBundleIdentifier": BUNDLE_IDENTIFIER,
        "CFBundleInfoDictionaryVersion": "6.0",
        "CFBundlePackageType": "APPL",
        "CFBundleShortVersionString": APP_VERSION,
        "CFBundleVersion": APP_VERSION,
        "LSMinimumSystemVersion": "10.13",
        "NSHighResolutionCapable": True,
    }
    with path.open("wb") as file:
        plistlib.dump(info, file, sort_keys=False)


def write_pkg_info(path: Path) -> None:
    path.write_text("APPL????", encoding="ascii")


def write_launcher(path: Path) -> None:
    python_executable = DEFAULT_PYTHON if DEFAULT_PYTHON.exists() else Path(sys.executable)
    script = f"""#!/bin/zsh
set -e

CONTENTS_DIR="$(cd "$(dirname "$0")/.." && pwd)"
BUNDLE_DIR="$(cd "$CONTENTS_DIR/.." && pwd)"
APP_DIR="$CONTENTS_DIR/Resources/app"
CONFIG_DIR="$HOME/Library/Application Support/艾松物流计费系统"
LOG_DIR="$HOME/Library/Logs/艾松物流计费系统"
LOG_FILE="$LOG_DIR/启动日志.log"
PYTHON_EXECUTABLE="${{EXPRESS_APP_PYTHON:-{python_executable}}}"

mkdir -p "$CONFIG_DIR" "$LOG_DIR" 2>/dev/null || true
if touch "$LOG_FILE" >/dev/null 2>&1; then
  exec >> "$LOG_FILE" 2>&1
fi
echo "========== $(/bin/date '+%Y-%m-%d %H:%M:%S') =========="
echo "BUNDLE_DIR=$BUNDLE_DIR"
echo "APP_DIR=$APP_DIR"
echo "CONFIG_DIR=$CONFIG_DIR"
echo "PYTHON_EXECUTABLE=$PYTHON_EXECUTABLE"

show_error() {{
  /usr/bin/osascript -e "display dialog \\"$1\\" buttons {{\\"确定\\"}} default button \\"确定\\" with icon stop" >/dev/null 2>&1 || true
}}

if [ ! -d "$APP_DIR/express_app" ]; then
  show_error "找不到应用代码目录：$APP_DIR/express_app"
  exit 1
fi

if [ ! -x "$PYTHON_EXECUTABLE" ]; then
  show_error "找不到 Python 运行环境：$PYTHON_EXECUTABLE"
  exit 1
fi

cd "$APP_DIR"
export TK_SILENCE_DEPRECATION=1
export EXPRESS_APP_CONFIG_FILE="$CONFIG_DIR/config.json"
unset __PYVENV_LAUNCHER__

if [ "${{EXPRESS_APP_SELF_CHECK:-}}" = "1" ]; then
  "$PYTHON_EXECUTABLE" - "$APP_DIR" <<'PY'
import sys
sys.path.insert(0, sys.argv[1])
from express_app.gui.app import APP_TITLE
from express_app.core import ExpressFeeBatchJobConfig, run_express_fee_batch_job
print(APP_TITLE)
print(ExpressFeeBatchJobConfig.__name__, callable(run_express_fee_batch_job))
PY
  exit 0
fi

echo "starting GUI"
exec "$PYTHON_EXECUTABLE" - "$APP_DIR" <<'PY'
import sys
sys.path.insert(0, sys.argv[1])
from express_app.gui.app import main
main()
PY
"""
    path.write_text(script, encoding="utf-8")
    current_mode = path.stat().st_mode
    path.chmod(current_mode | stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)


def main() -> int:
    app_dir = Path(__file__).resolve().parent
    bundle_dir = build_app_bundle(app_dir)
    print(f"已生成：{bundle_dir}")
    print(f"启动文件：{bundle_dir / 'Contents' / 'MacOS' / EXECUTABLE_NAME}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
