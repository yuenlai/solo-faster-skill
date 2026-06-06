#!/usr/bin/env python3

from __future__ import annotations

import argparse
from pathlib import Path

from prompt_history_lib import append_history, detect_prefix


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--project-path", default=".", help="当前项目路径")
    parser.add_argument("--project-name", default="", help="可选项目名")
    parser.add_argument("--task-type", required=True, help="当前任务类型")
    parser.add_argument("--angle", required=True, help="提示词切入点")
    parser.add_argument("--prompt", required=True, help="最终提示词")
    args = parser.parse_args()

    project_path = Path(args.project_path).resolve()
    project_name = args.project_name.strip() or project_path.name
    prefix = detect_prefix(project_name, project_path.name, str(project_path))

    append_history(
        {
            "project_path": str(project_path),
            "project_name": project_name,
            "prefix": prefix,
            "task_type": args.task_type.strip(),
            "angle": args.angle.strip(),
            "prompt": args.prompt.strip(),
        }
    )

    print("ok")


if __name__ == "__main__":
    main()
