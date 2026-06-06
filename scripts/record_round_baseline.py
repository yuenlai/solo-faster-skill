#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
from datetime import datetime
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


def parse_index_entries(text: str) -> dict[str, dict[str, str]]:
    entries: dict[str, dict[str, str]] = {}
    for line in text.splitlines():
        if not line.strip():
            continue
        meta, path = line.split("\t", 1)
        mode, oid, stage = meta.split(" ", 2)
        entries[path] = {"mode": mode, "oid": oid, "stage": stage}
    return entries


def copy_dirty_files(repo: Path, snapshot_dir: Path, entries: list[dict[str, str]]) -> None:
    files_dir = snapshot_dir / "files"
    for entry in entries:
        rel_path = entry["path"]
        src = repo / rel_path
        if not src.exists() or not src.is_file():
            continue
        dest = files_dir / rel_path
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo", default=".", help="目标仓库路径，默认当前目录")
    parser.add_argument("--label", default="", help="可选标签，便于人工识别")
    args = parser.parse_args()

    repo = Path(args.repo).resolve()
    git_dir = run_git(repo, "rev-parse", "--git-dir").strip()
    git_dir_path = (repo / git_dir).resolve() if not Path(git_dir).is_absolute() else Path(git_dir)

    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    safe_label = args.label.strip().replace(" ", "-")
    folder_name = f"{timestamp}-{safe_label}" if safe_label else timestamp
    snapshot_dir = git_dir_path / "solo-faster-baselines" / folder_name
    snapshot_dir.mkdir(parents=True, exist_ok=True)

    status_text = run_git(repo, "status", "--short")
    diff_stat_text = run_git(repo, "diff", "--stat")
    unstaged_diff = run_git(repo, "diff", "--binary")
    staged_diff = run_git(repo, "diff", "--cached", "--binary")
    head = run_git(repo, "rev-parse", "HEAD").strip()
    index_entries = parse_index_entries(run_git(repo, "ls-files", "--stage"))
    entries = parse_status_lines(status_text)
    dirty_paths = sorted({entry["path"] for entry in entries})

    metadata = {
      "repo": str(repo),
      "git_dir": str(git_dir_path),
      "created_at": datetime.now().isoformat(timespec="seconds"),
      "label": args.label,
      "head": head,
      "dirty_entries": entries,
      "dirty_paths": dirty_paths,
      "index_entries": index_entries,
      "tracked_paths": sorted(index_entries),
    }

    (snapshot_dir / "metadata.json").write_text(
        json.dumps(metadata, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    (snapshot_dir / "status.txt").write_text(status_text, encoding="utf-8")
    (snapshot_dir / "diff_stat.txt").write_text(diff_stat_text, encoding="utf-8")
    (snapshot_dir / "diff.patch").write_text(unstaged_diff, encoding="utf-8")
    (snapshot_dir / "diff_cached.patch").write_text(staged_diff, encoding="utf-8")
    copy_dirty_files(repo, snapshot_dir, entries)

    print(snapshot_dir)


if __name__ == "__main__":
    main()
