#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import subprocess
from pathlib import Path

from prompt_history_lib import HISTORY_DIR


STATE_FILE = HISTORY_DIR / "trae_open_state.json"
TRAE_APP = "Trae CN.app"
TRAE_PROCESS = "TRAE CN"


def run_command(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(args, capture_output=True, text=True)


def load_state() -> dict[str, str]:
    if not STATE_FILE.exists():
        return {}
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def save_state(state: dict[str, str]) -> None:
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)
    STATE_FILE.write_text(
        json.dumps(state, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def get_window_titles() -> list[str]:
    script = (
        f'tell application "System Events" to tell process "{TRAE_PROCESS}" '
        'to get name of every window'
    )
    result = run_command("osascript", "-e", script)
    if result.returncode != 0:
        return []

    text = result.stdout.strip()
    if not text:
        return []
    return [item.strip() for item in text.split(", ") if item.strip()]


def is_project_open(project_path: Path, window_titles: list[str]) -> bool:
    project_name = project_path.name
    for title in window_titles:
        if project_name in title:
            return True
    return False


def open_project(project_path: Path) -> None:
    result = run_command("open", "-a", TRAE_APP, str(project_path))
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "打开 Trae 失败")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project-path", default=".", help="当前项目路径")
    args = parser.parse_args()

    project_path = Path(args.project_path).resolve()
    state = load_state()
    known_path = state.get(str(project_path))

    window_titles = get_window_titles()
    if is_project_open(project_path, window_titles):
        state[str(project_path)] = project_path.name
        save_state(state)
        print("already-open")
        return

    if known_path == project_path.name and not window_titles:
        print("unknown-window-state")
        return

    open_project(project_path)
    state[str(project_path)] = project_path.name
    save_state(state)
    print("opened")


if __name__ == "__main__":
    main()
