#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import subprocess
from pathlib import Path


def run_git(repo: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(repo), *args],
        check=True,
        capture_output=True,
        text=True,
    )
    return result.stdout


def parse_status_lines(text: str) -> list[dict[str, str]]:
    entries: list[dict[str, str]] = []
    for line in text.splitlines():
        if not line.strip():
            continue
        code = line[:2]
        path_text = line[3:]
        if " -> " in path_text:
            path_text = path_text.split(" -> ", 1)[1]
        entries.append({"code": code, "path": path_text, "raw": line})
    return entries


def latest_baseline_dir(repo: Path) -> Path:
    git_dir = run_git(repo, "rev-parse", "--git-dir").strip()
    git_dir_path = (repo / git_dir).resolve() if not Path(git_dir).is_absolute() else Path(git_dir)
    base_dir = git_dir_path / "solo-faster-baselines"
    candidates = sorted([path for path in base_dir.iterdir() if path.is_dir()])
    if not candidates:
        raise FileNotFoundError("未找到任何基线记录")
    return candidates[-1]


def same_file(a: Path, b: Path) -> bool:
    if not a.exists() and not b.exists():
        return True
    if not a.exists() or not b.exists():
        return False
    if a.is_file() and b.is_file():
        return a.read_bytes() == b.read_bytes()
    return a.resolve() == b.resolve()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo", default=".", help="目标仓库路径，默认当前目录")
    parser.add_argument(
        "--baseline",
        default="",
        help="基线路径，默认自动读取最近一次记录",
    )
    args = parser.parse_args()

    repo = Path(args.repo).resolve()
    baseline_dir = Path(args.baseline).resolve() if args.baseline else latest_baseline_dir(repo)
    metadata = json.loads((baseline_dir / "metadata.json").read_text(encoding="utf-8"))
    baseline_paths = set(metadata.get("dirty_paths", []))

    current_status = run_git(repo, "status", "--short")
    current_entries = parse_status_lines(current_status)
    current_paths = {entry["path"] for entry in current_entries}

    new_this_round: list[str] = []
    changed_again_this_round: list[str] = []
    unchanged_historical: list[str] = []

    for path in sorted(current_paths):
        if path not in baseline_paths:
            new_this_round.append(path)
            continue

        baseline_file = baseline_dir / "files" / path
        current_file = repo / path
        if same_file(baseline_file, current_file):
            unchanged_historical.append(path)
        else:
            changed_again_this_round.append(path)

    print(f"基线目录: {baseline_dir}")
    print("本轮新增改动文件:")
    if new_this_round:
        for path in new_this_round:
            print(f"- {path}")
    else:
        print("- 无")

    print("本轮继续修改的历史文件:")
    if changed_again_this_round:
        for path in changed_again_this_round:
            print(f"- {path}")
    else:
        print("- 无")

    print("仅延续历史、本次任务没有继续变化的文件:")
    if unchanged_historical:
        for path in unchanged_historical:
            print(f"- {path}")
    else:
        print("- 无")


if __name__ == "__main__":
    main()
