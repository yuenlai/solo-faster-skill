#!/usr/bin/env python3
"""Self-check solo-faster v2 workbook and state-machine behavior."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
from pathlib import Path

from batch_prompt_workbook import (
    apply_outcome,
    bugfix_count,
    choose_next,
    initialize,
    insert_fix,
    read_workbook,
    update_row,
    validate_records,
    workbook_path,
)


SEND_SCRIPT = Path(__file__).resolve().with_name("send_batch_prompts_to_trae.py")


def run(*args: str, cwd: Path | None = None) -> None:
    result = subprocess.run(args, cwd=cwd, capture_output=True, text=True)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or f"{' '.join(args)} failed")


def run_result(*args: str, cwd: Path | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(args, cwd=cwd, capture_output=True, text=True)


def make_repo(root: Path) -> Path:
    repo = root / "demo-repo"
    repo.mkdir()
    (repo / "package.json").write_text('{"scripts":{"test":"echo ok"},"dependencies":{"vite":"latest"}}\n', encoding="utf-8")
    (repo / "src").mkdir()
    (repo / "src" / "main.js").write_text("console.log('demo');\n", encoding="utf-8")
    run("git", "init", cwd=repo)
    run("git", "add", ".", cwd=repo)
    run("git", "commit", "-m", "init", cwd=repo)
    return repo


def reason(text: str) -> str:
    return text


def main() -> None:
    checks: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(prefix="solo-faster-v2-") as temp:
        root = Path(temp)
        repo = make_repo(root)
        parent = root
        created = initialize(parent, repo, None, "中", "Web 前端", with_stubs=True, force=True)
        records = read_workbook(workbook_path(parent, None))
        validation = validate_records(records)
        checks.append({"name": "create default main rows", "ok": validation["ok"], "row_count": validation["row_count"]})

        apply_outcome(parent, None, 1, 1, "已完成", "", "", "", "单文件", "dry-run 完成")
        records = read_workbook(workbook_path(parent, None))
        checks.append({"name": "main 1 completed then next main 2", "ok": choose_next(records)["row"]["主提示词编号"] == "2"})

        sent_next_main = run_result(
            sys.executable,
            str(SEND_SCRIPT),
            "send",
            "--parent",
            str(parent),
            "--repo",
            str(repo),
            "--range",
            "2",
            "--dry-run",
        )
        next_payload = json.loads(sent_next_main.stdout) if sent_next_main.returncode == 0 else {}
        checks.append({
            "name": "send next main only after terminal main",
            "ok": sent_next_main.returncode == 0 and next_payload.get("sent") == [{"main": "2", "round": "1"}],
        })
        update_row(parent, None, 2, 1, {"执行状态": "已生成", "开始时间": "", "备注": "dry-run rollback next main send"})

        apply_outcome(
            parent,
            None,
            2,
            1,
            "未完成",
            reason("任务拆分漏掉主要交互路径，过早把局部显示变化当成完整结果。"),
            reason("用户完成操作后仍看不到关键结果反馈，页面状态和实际操作结果对不上。"),
            reason("任务拆分漏掉主要交互路径，用户完成操作后仍看不到关键结果反馈。"),
            "模块内多文件",
            "dry-run 未完成",
        )
        inserted = insert_fix(parent, None, 2, "修复操作完成后关键结果反馈没有同步显示的问题。")
        checks.append({"name": "failed main inserts round 2 fix", "ok": inserted["row"]["轮次"] == "2" and inserted["row"]["任务类型"] == "Bug修复"})

        skipped_fix = run_result(
            sys.executable,
            str(SEND_SCRIPT),
            "send",
            "--parent",
            str(parent),
            "--repo",
            str(repo),
            "--range",
            "3",
            "--dry-run",
        )
        checks.append({"name": "send refuses skipping pending fix", "ok": skipped_fix.returncode != 0 and "禁止绕过断点续跑顺序发送提示词" in skipped_fix.stderr})

        sent_fix = run_result(
            sys.executable,
            str(SEND_SCRIPT),
            "send",
            "--parent",
            str(parent),
            "--repo",
            str(repo),
            "--range",
            "2",
            "--dry-run",
        )
        sent_payload = json.loads(sent_fix.stdout) if sent_fix.returncode == 0 else {}
        checks.append({
            "name": "send keeps current main fix round",
            "ok": sent_fix.returncode == 0 and sent_payload.get("sent") == [{"main": "2", "round": "2"}],
        })

        apply_outcome(
            parent,
            None,
            2,
            2,
            "未完成",
            reason("主链路识别仍然偏在局部展示，没有把操作结果和后续反馈当成同一个业务对象生命周期。"),
            reason("用户再次操作后还是看不到可信的结果反馈，界面前后状态继续对不上。"),
            reason("主链路识别偏在局部展示，用户再次操作后还是看不到可信反馈。"),
            "模块内多文件",
            "dry-run 二轮未完成",
        )
        inserted3 = insert_fix(parent, None, 2, "修复结果反馈和后续状态仍然不一致的问题。")
        apply_outcome(
            parent,
            None,
            2,
            3,
            "未完成",
            reason("完成判断过早，核心结果反馈仍未按同一条业务链路验到末端。"),
            reason("用户操作完成后页面反馈仍然不可信，关键状态没有表现出一致结果。"),
            reason("完成判断过早，用户操作完成后页面反馈仍然不可信。"),
            "模块内多文件",
            "dry-run 三轮未完成",
        )
        records = read_workbook(workbook_path(parent, None))
        checks.append({"name": "round 3 failed stops fixing", "ok": choose_next(records)["row"]["主提示词编号"] == "3"})

        for main_number in range(3, 11):
            apply_outcome(
                parent,
                None,
                main_number,
                1,
                "未完成",
                reason("任务拆分漏掉关键状态归属，局部结果没有接到用户会继续查看的主路径。"),
                reason("用户完成核心操作后仍看不到一致的结果反馈，页面表现和预期不一致。"),
                reason("任务拆分漏掉关键状态归属，用户完成核心操作后仍看不到一致反馈。"),
                "单文件",
                "dry-run 未完成并生成修复",
            )
            fix = insert_fix(parent, None, main_number, "修复核心操作后的结果反馈仍不一致的问题。")
            apply_outcome(
                parent,
                None,
                main_number,
                int(fix["row"]["轮次"]),
                "已完成",
                "",
                "",
                "",
                "单文件",
                "dry-run 修复完成",
            )
        records = read_workbook(workbook_path(parent, None))
        checks.append({"name": "bugfix count reaches 10", "ok": bugfix_count(records) == 10})
        apply_outcome(
            parent,
            None,
            11,
            1,
            "未完成",
            reason("目标抓偏到局部展示，没有先把用户最关心的结果归属口径统一起来。"),
            reason("用户完成操作后关键结果还是不可见，界面没有给出可信反馈。"),
            reason("目标抓偏到局部展示，用户完成操作后关键结果还是不可见。"),
            "单文件",
            "dry-run 上限后不修复",
        )
        records = read_workbook(workbook_path(parent, None))
        checks.append({"name": "after 10 bugfixes next main only", "ok": choose_next(records)["row"]["主提示词编号"] == "12"})

        update_row(parent, None, 12, 1, {"执行状态": "已发送", "备注": "dry-run 已发送"})
        update_row(parent, None, 12, 1, {"执行状态": "Trae运行中", "备注": "dry-run 运行中"})
        update_row(parent, None, 12, 1, {"执行状态": "超时待人工", "备注": "dry-run 超时待人工"})
        records = read_workbook(workbook_path(parent, None))
        checks.append({"name": "timeout waits for human", "ok": choose_next(records)["action"] == "confirm"})
        update_row(parent, None, 12, 1, {"执行状态": "待验收", "备注": "dry-run 恢复后待验收"})
        records = read_workbook(workbook_path(parent, None))
        checks.append({"name": "resume pending acceptance", "ok": choose_next(records)["action"] == "accept"})

        final_validation = validate_records(records)
        checks.append({"name": "final workbook lint", "ok": final_validation["ok"], "errors": final_validation["errors"]})

        result = {
            "workbook": created["workbook"],
            "checks": checks,
            "ok": all(bool(item["ok"]) for item in checks),
        }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if not result["ok"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
