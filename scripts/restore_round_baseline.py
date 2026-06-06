#!/usr/bin/env python3
"""Restore a repository to a solo-faster baseline without using reset --hard."""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
from datetime import datetime
from pathlib import Path


DEFAULT_EXCLUDES = {"solo-faster-prompts.xlsx"}


def run_git(repo: Path, *args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(["git", "-C", str(repo), *args], capture_output=True, text=True)
    if check and result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or f"git {' '.join(args)} failed")
    return result


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
    git_dir = run_git(repo, "rev-parse", "--git-dir").stdout.strip()
    git_dir_path = (repo / git_dir).resolve() if not Path(git_dir).is_absolute() else Path(git_dir)
    base_dir = git_dir_path / "solo-faster-baselines"
    candidates = sorted([path for path in base_dir.iterdir() if path.is_dir()])
    if not candidates:
        raise FileNotFoundError("未找到任何基线记录")
    return candidates[-1]


def backup_current_dirty(repo: Path, git_dir: Path, entries: list[dict[str, str]]) -> Path:
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup_dir = git_dir / "solo-faster-restore-backups" / timestamp
    files_dir = backup_dir / "files"
    backup_dir.mkdir(parents=True, exist_ok=True)
    (backup_dir / "status.txt").write_text(run_git(repo, "status", "--short").stdout, encoding="utf-8")
    (backup_dir / "diff.patch").write_text(run_git(repo, "diff", "--binary").stdout, encoding="utf-8")
    (backup_dir / "diff_cached.patch").write_text(run_git(repo, "diff", "--cached", "--binary").stdout, encoding="utf-8")
    for entry in entries:
        rel = entry["path"]
        src = repo / rel
        if not src.exists() or not src.is_file():
            continue
        dest = files_dir / rel
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)
    return backup_dir


def remove_path(path: Path) -> None:
    if path.is_dir() and not path.is_symlink():
        shutil.rmtree(path)
    elif path.exists() or path.is_symlink():
        path.unlink()


def checkout_head_file(repo: Path, rel: str, head: str) -> None:
    dest = repo / rel
    dest.parent.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(["git", "-C", str(repo), "show", f"{head}:{rel}"], capture_output=True)
    if result.returncode != 0:
        remove_path(dest)
        return
    dest.write_bytes(result.stdout)


def restore_index_entry(repo: Path, rel: str, baseline_index: dict[str, dict[str, str]], dry_run: bool, actions: list[dict[str, str]]) -> None:
    entry = baseline_index.get(rel)
    if entry:
        actions.append({"path": rel, "action": "restore-index-entry"})
        if not dry_run:
            run_git(repo, "update-index", "--add", "--cacheinfo", f"{entry['mode']},{entry['oid']},{rel}")
        return
    actions.append({"path": rel, "action": "remove-index-entry"})
    if not dry_run:
        run_git(repo, "update-index", "--force-remove", "--", rel, check=False)


def restore_baseline(repo: Path, baseline_dir: Path, excludes: set[str], dry_run: bool) -> dict[str, object]:
    metadata = json.loads((baseline_dir / "metadata.json").read_text(encoding="utf-8"))
    git_dir = Path(metadata["git_dir"]).resolve()
    baseline_dirty = set(metadata.get("dirty_paths", []))
    tracked_paths = set(metadata.get("tracked_paths", []))
    baseline_index = metadata.get("index_entries", {})
    head = metadata.get("head", "HEAD")
    current_entries = parse_status_lines(run_git(repo, "status", "--short").stdout)
    current_paths = {entry["path"] for entry in current_entries}
    excluded_paths = {path for path in current_paths.union(baseline_dirty) if Path(path).name in excludes}
    touched_paths = sorted((current_paths.union(baseline_dirty)) - excluded_paths)
    backup_dir = backup_current_dirty(repo, git_dir, current_entries) if current_entries and not dry_run else None

    actions: list[dict[str, str]] = []
    files_dir = baseline_dir / "files"
    for rel in touched_paths:
        dest = repo / rel
        if rel in baseline_dirty:
            snapshot = files_dir / rel
            actions.append({"path": rel, "action": "restore-dirty-snapshot" if snapshot.exists() else "restore-baseline-deletion"})
            if not dry_run:
                if snapshot.exists():
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(snapshot, dest)
                else:
                    remove_path(dest)
        elif rel in tracked_paths:
            actions.append({"path": rel, "action": "restore-head-file"})
            if not dry_run:
                checkout_head_file(repo, rel, head)
        else:
            actions.append({"path": rel, "action": "remove-new-path"})
            if not dry_run:
                remove_path(dest)
        restore_index_entry(repo, rel, baseline_index, dry_run, actions)

    return {
        "repo": str(repo),
        "baseline": str(baseline_dir),
        "backup": str(backup_dir) if backup_dir else "",
        "excluded": sorted(excluded_paths),
        "actions": actions,
        "dry_run": dry_run,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--repo", default=".")
    parser.add_argument("--baseline", default="")
    parser.add_argument("--exclude", action="append", default=[])
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    repo = Path(args.repo).expanduser().resolve()
    baseline_dir = Path(args.baseline).expanduser().resolve() if args.baseline else latest_baseline_dir(repo)
    excludes = set(DEFAULT_EXCLUDES).union(args.exclude)
    try:
        result = restore_baseline(repo, baseline_dir, excludes, args.dry_run)
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2))
        raise SystemExit(2) from exc
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
