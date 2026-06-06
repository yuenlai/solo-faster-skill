#!/usr/bin/env python3

from __future__ import annotations

import argparse
import json
from pathlib import Path

from prompt_history_lib import HISTORY_FILE, detect_prefix, load_history


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project-path", default=".", help="当前项目路径")
    parser.add_argument("--project-name", default="", help="可选项目名")
    parser.add_argument("--task-type", required=True, help="当前任务类型")
    parser.add_argument("--limit", type=int, default=5, help="返回最近几条历史")
    args = parser.parse_args()

    project_path = Path(args.project_path).resolve()
    project_name = args.project_name.strip()
    prefix = detect_prefix(project_name, project_path.name, str(project_path))
    task_type = args.task_type.strip()

    history = load_history()
    matched = [
        item for item in history
        if item.get("prefix", "") == prefix and item.get("task_type", "") == task_type
    ]
    matched = matched[-args.limit:]

    result = {
        "history_file": str(HISTORY_FILE),
        "project_path": str(project_path),
        "project_name": project_name or project_path.name,
        "prefix": prefix,
        "task_type": task_type,
        "matched_count": len(matched),
        "matched": matched,
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
