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
    audit_records,
    bugfix_count,
    classify_change_range,
    choose_next,
    detect_domain,
    difficulty_for,
    finalize_round,
    interleaved_main_task_order,
    initialize,
    insert_fix,
    read_workbook,
    update_row,
    write_workbook,
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
    run("git", "remote", "add", "origin", "git@github.com:yuenlai/solo-6600007.git", cwd=repo)
    run("git", "add", ".", cwd=repo)
    run("git", "commit", "-m", "init", cwd=repo)
    return repo


def make_fullstack_repo(root: Path) -> Path:
    repo = root / "fullstack-repo"
    repo.mkdir()
    (repo / "package.json").write_text('{"dependencies":{"vite":"latest","react":"latest"}}\n', encoding="utf-8")
    (repo / "src").mkdir()
    (repo / "src" / "App.tsx").write_text("export default function App() { return null; }\n", encoding="utf-8")
    (repo / "server").mkdir()
    (repo / "server" / "api.py").write_text("def handler():\n    return {}\n", encoding="utf-8")
    run("git", "init", cwd=repo)
    run("git", "remote", "add", "origin", "https://github.com/yuenlai/fullstack-demo.git", cwd=repo)
    run("git", "add", ".", cwd=repo)
    run("git", "commit", "-m", "init", cwd=repo)
    return repo


def reason(text: str) -> str:
    return text


def mark_session(parent: Path, main_number: int, round_number: int, session_id: str, status: str, note: str = "") -> None:
    update_row(
        parent,
        None,
        main_number,
        round_number,
        {
            "Trae Session ID": session_id,
            "执行状态": status,
            "备注": note or f"dry-run {status}",
        },
    )


def main_numbers_by_task_type(records: list[dict[str, str]], task_type: str) -> list[int]:
    return [
        int(record["主提示词编号"])
        for record in records
        if record.get("行类型") == "主提示词" and record.get("任务类型") == task_type
    ]


def mutate_repo(repo: Path, label: str) -> None:
    target = repo / "src" / "main.js"
    with target.open("a", encoding="utf-8") as handle:
        handle.write(f"console.log('{label}');\n")


def complete_and_commit(
    parent: Path,
    repo: Path,
    main_number: int,
    round_number: int,
    completed: str,
    change_range: str,
    process_reason: str = "",
    product_reason: str = "",
    merged_reason: str = "",
) -> str:
    session_id = f"dryrun-session-main-{main_number}-r{round_number}"
    update_row(parent, None, main_number, round_number, {"Trae Session ID": session_id})
    mutate_repo(repo, f"main-{main_number}-round-{round_number}")
    apply_outcome(parent, None, main_number, round_number, completed, process_reason, product_reason, merged_reason, change_range, f"dry-run {completed}")
    finalized = finalize_round(parent, None, repo, main_number, round_number, no_push=True)
    return str(finalized["commit_id"])


def bad_audit_case(root: Path, name: str, fields: dict[str, str]) -> dict[str, object]:
    case_root = root / name
    case_root.mkdir()
    repo = make_repo(case_root)
    parent = case_root
    initialize(parent, repo, None, "中", "Web 前端", with_stubs=True, force=True, task_counts="0-1代码生成=1")
    commit_id = complete_and_commit(parent, repo, 1, 1, "已完成", "单文件")
    update_row(parent, None, 1, 1, fields, strict_validate=False)
    if fields.get("不满意原因") and fields.get("任务是否完成", "已完成") == "已完成":
        path = workbook_path(parent, None)
        records = read_workbook(path)
        records[0]["不满意原因"] = fields["不满意原因"]
        write_workbook(path, records)
    audit = audit_records(parent, None, repo)
    return {"name": name, "ok": not audit["ok"], "errors": audit.get("errors", []), "git_errors": audit.get("git_errors", []), "valid_commit": commit_id}


def bad_duplicate_reason_case(root: Path) -> dict[str, object]:
    case_root = root / "bad-duplicate-reason"
    case_root.mkdir()
    repo = make_repo(case_root)
    parent = case_root
    initialize(parent, repo, None, "中", "Web 前端", with_stubs=True, force=True, task_counts="0-1代码生成=2")
    duplicate_tail = "用户完成操作后仍看不到新增结果，页面状态和预期反馈不一致。"
    complete_and_commit(
        parent,
        repo,
        1,
        1,
        "未完成",
        "单文件",
        "任务拆分只覆盖入口展示，没有处理结果回看链路。",
        duplicate_tail,
        "",
    )
    session_id = "dryrun-session-main-2-r1"
    update_row(parent, None, 2, 1, {"Trae Session ID": session_id})
    mutate_repo(repo, "bad-duplicate-reason-r2")
    run("git", "add", "src/main.js", cwd=repo)
    run("git", "commit", "-m", session_id, cwd=repo)
    commit_id = run_result("git", "rev-parse", "HEAD", cwd=repo).stdout.strip()
    update_row(
        parent,
        None,
        2,
        1,
        {
            "任务是否完成": "未完成",
            "执行状态": "已完成",
            "过程与产物是否满意": "不满意",
            "修改范围": "单文件",
            "不满意原因": "过程不满意：目标抓偏到局部按钮响应，没有验证保存后的对象归属。\n产物不满意：" + duplicate_tail,
            "Commit ID": commit_id,
        },
        strict_validate=False,
    )
    audit = audit_records(parent, None, repo)
    return {"name": "bad-duplicate-reason", "ok": not audit["ok"], "errors": audit.get("errors", []), "git_errors": audit.get("git_errors", [])}


def bad_duplicate_reason_write_case(root: Path) -> dict[str, object]:
    case_root = root / "bad-duplicate-reason-write"
    case_root.mkdir()
    repo = make_repo(case_root)
    parent = case_root
    initialize(parent, repo, None, "中", "Web 前端", with_stubs=True, force=True, task_counts="0-1代码生成=2")
    duplicate_tail = "用户完成操作后仍看不到新增结果，页面状态和预期反馈不一致。"
    complete_and_commit(
        parent,
        repo,
        1,
        1,
        "未完成",
        "单文件",
        "任务拆分只覆盖入口展示，没有处理结果回看链路。",
        duplicate_tail,
        "",
    )
    try:
        update_row(parent, None, 2, 1, {"Trae Session ID": "dryrun-session-main-2-r1"})
        mutate_repo(repo, "bad-duplicate-write")
        apply_outcome(
            parent,
            None,
            2,
            1,
            "未完成",
            "目标抓偏到局部按钮响应，没有验证保存后的对象归属。",
            duplicate_tail,
            "",
            "单文件",
            "dry-run duplicate write",
        )
    except SystemExit as exc:
        return {"name": "bad-duplicate-reason-write", "ok": True, "error": str(exc)}
    return {"name": "bad-duplicate-reason-write", "ok": False, "error": "duplicate reason write was accepted"}


def reason_check_regenerate_case(root: Path) -> dict[str, object]:
    case_root = root / "reason-check-regenerate"
    case_root.mkdir()
    repo = make_repo(case_root)
    parent = case_root
    initialize(parent, repo, None, "中", "Web 前端", with_stubs=True, force=True, task_counts="0-1代码生成=2")
    duplicate_tail = "用户完成操作后仍看不到新增结果，页面状态和预期反馈不一致。"
    complete_and_commit(
        parent,
        repo,
        1,
        1,
        "未完成",
        "单文件",
        "任务拆分只覆盖入口展示，没有处理结果回看链路。",
        duplicate_tail,
        "",
    )
    checked = run_result(
        sys.executable,
        str(Path(__file__).resolve().with_name("batch_prompt_workbook.py")),
        "reason-check",
        "--parent",
        str(parent),
        "--main",
        "2",
        "--round",
        "1",
        "--process-reason",
        "目标抓偏到局部按钮响应，没有验证保存后的对象归属。",
        "--product-reason",
        duplicate_tail,
    )
    payload = json.loads(checked.stdout) if checked.returncode == 0 else {}
    return {
        "name": "reason-check marks bad candidate for regeneration",
        "ok": checked.returncode == 0 and payload.get("ok") is False and payload.get("regenerate_required") is True,
        "payload": payload,
    }


def main() -> None:
    checks: list[dict[str, object]] = []
    with tempfile.TemporaryDirectory(prefix="solo-faster-v2-") as temp:
        root = Path(temp)
        repo = make_repo(root)
        parent = root
        created = initialize(parent, repo, None, "中", "Web 前端", with_stubs=True, force=True, task_counts=None)
        records = read_workbook(workbook_path(parent, None))
        validation = audit_records(parent, None, repo)
        checks.append({"name": "create default main rows", "ok": validation["ok"], "row_count": validation["row_count"]})
        checks.append({
            "name": "repo url uses markdown github format",
            "ok": records[0].get("Repo URL") == "[yuenlai/solo-6600007](https://github.com/yuenlai/solo-6600007)",
            "repo_url": records[0].get("Repo URL"),
        })
        first_twelve_types = [record["任务类型"] for record in records[:12]]
        checks.append({
            "name": "default distribution interleaves main task types",
            "ok": first_twelve_types[:6] == ["0-1代码生成", "Feature迭代", "0-1代码生成", "Feature迭代", "0-1代码生成", "Feature迭代"],
            "first_twelve_types": first_twelve_types,
        })
        custom_order = interleaved_main_task_order([("0-1代码生成", 4), ("Feature迭代", 3), ("代码理解", 1), ("工程化", 1)])
        checks.append({
            "name": "custom distribution scatters minority task types",
            "ok": custom_order == ["0-1代码生成", "Feature迭代", "0-1代码生成", "Feature迭代", "0-1代码生成", "代码理解", "0-1代码生成", "Feature迭代", "工程化"],
            "custom_order": custom_order,
        })
        high_difficulties = [difficulty_for("高", index) for index in range(1, 8)]
        checks.append({
            "name": "high difficulty mode avoids automatic hell difficulty",
            "ok": "地狱" not in high_difficulties and set(high_difficulties) == {"困难"},
            "difficulties": high_difficulties,
        })
        zero_to_one_mains = main_numbers_by_task_type(records, "0-1代码生成")

        complete_and_commit(parent, repo, 1, 1, "已完成", "单文件")
        records = read_workbook(workbook_path(parent, None))
        checks.append({"name": "main 1 completed then next main 2", "ok": choose_next(records)["row"]["主提示词编号"] == "2"})
        checks.append({
            "name": "finalize keeps markdown repo url",
            "ok": records[0].get("Repo URL") == "[yuenlai/solo-6600007](https://github.com/yuenlai/solo-6600007)",
            "repo_url": records[0].get("Repo URL"),
        })

        manual_next_main = run_result(
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
        manual_payload = json.loads(manual_next_main.stdout) if manual_next_main.returncode == 0 else {}
        manual_instruction = (manual_payload.get("instructions") or [{}])[0]
        checks.append({
            "name": "manual send instruction includes bracketed main and round",
            "ok": manual_next_main.returncode == 0
            and manual_instruction.get("turn_label") == "【第 2 个主提示词，第 1 轮】"
            and str(manual_instruction.get("instruction", "")).startswith("【第 2 个主提示词，第 1 轮】"),
            "instruction": manual_instruction,
        })

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
            "--auto-ui",
        )
        next_payload = json.loads(sent_next_main.stdout) if sent_next_main.returncode == 0 else {}
        checks.append({
            "name": "send next main only after terminal main",
            "ok": sent_next_main.returncode == 0 and next_payload.get("sent") == [{"main": "2", "round": "1"}],
        })
        update_row(parent, None, 2, 1, {"执行状态": "已生成", "开始时间": "", "备注": "dry-run rollback next main send"})

        complete_and_commit(
            parent,
            repo,
            2,
            1,
            "未完成",
            "模块内多文件",
            reason("任务拆分漏掉主要交互路径，过早把局部显示变化当成完整结果。"),
            reason("用户完成操作后仍看不到关键结果反馈，页面状态和实际操作结果对不上。"),
            reason("任务拆分漏掉主要交互路径，用户完成操作后仍看不到关键结果反馈。"),
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

        manual_fix = run_result(
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
        manual_fix_payload = json.loads(manual_fix.stdout) if manual_fix.returncode == 0 else {}
        manual_fix_instruction = (manual_fix_payload.get("instructions") or [{}])[0]
        checks.append({
            "name": "manual fix instruction includes bracketed main and round",
            "ok": manual_fix.returncode == 0
            and manual_fix_instruction.get("turn_label") == "【第 2 个主提示词，第 2 轮】"
            and str(manual_fix_instruction.get("instruction", "")).startswith("【第 2 个主提示词，第 2 轮】"),
            "instruction": manual_fix_instruction,
        })

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
            "--auto-ui",
        )
        sent_payload = json.loads(sent_fix.stdout) if sent_fix.returncode == 0 else {}
        checks.append({
            "name": "send keeps current main fix round",
            "ok": sent_fix.returncode == 0 and sent_payload.get("sent") == [{"main": "2", "round": "2"}],
        })

        complete_and_commit(
            parent,
            repo,
            2,
            2,
            "未完成",
            "模块内多文件",
            reason("主链路识别仍然偏在局部展示，没有把操作结果和后续反馈当成同一个业务对象生命周期。"),
            reason("用户再次操作后还是看不到可信的结果反馈，界面前后状态继续对不上。"),
            reason("主链路识别偏在局部展示，用户再次操作后还是看不到可信反馈。"),
        )
        inserted3 = insert_fix(parent, None, 2, "修复结果反馈和后续状态仍然不一致的问题。")
        complete_and_commit(
            parent,
            repo,
            2,
            3,
            "未完成",
            "模块内多文件",
            reason("完成判断过早，核心结果反馈仍未按同一条业务链路验到末端。"),
            reason("用户操作完成后页面反馈仍然不可信，关键状态没有表现出一致结果。"),
            reason("完成判断过早，用户操作完成后页面反馈仍然不可信。"),
        )
        records = read_workbook(workbook_path(parent, None))
        checks.append({"name": "round 3 failed stops fixing", "ok": choose_next(records)["row"]["主提示词编号"] == "3"})

        promoted_zero_to_one = zero_to_one_mains[1:6]
        promoted_reasons = [
            (
                "验收只看到了入口卡片出现，没有追到保存后的明细页回显。",
                "用户提交资料后仍找不到刚创建的记录，列表和详情之间没有形成可回看的结果。",
                "验收停在入口卡片，用户提交资料后仍找不到刚创建的记录。",
            ),
            (
                "目标拆分忽略了筛选条件和结果数量之间的联动关系。",
                "切换不同筛选项后列表内容没有给出对应变化，用户看到的结果仍像默认列表。",
                "目标拆分忽略筛选联动，切换条件后列表没有对应变化。",
            ),
            (
                "实现时把统计数字当成静态展示，没有接入用户操作后的重新计算链路。",
                "完成一次新增或删除后，页面上的统计值仍停在旧状态，用户看不到数据变化。",
                "统计展示没有接入操作后的重新计算，新增或删除后数字仍停在旧状态。",
            ),
            (
                "处理流程只覆盖了弹窗打开，没有验证提交关闭后主页面是否同步更新。",
                "用户在弹窗里保存内容后，主页面仍看不到对应条目，也缺少成功后的状态提示。",
                "弹窗提交后的主页面同步没有验证，保存内容后仍看不到对应条目。",
            ),
            (
                "任务判断过早停在默认数据展示，没有确认空态和真实数据态能互相切换。",
                "清空数据或重新添加内容后，页面仍沿用默认展示，用户看到的状态和操作结果脱节。",
                "默认数据展示被过早当成完成，空态和真实数据态没有可靠切换。",
            ),
        ]
        for main_number, (process_text, product_text, merged_text) in zip(promoted_zero_to_one, promoted_reasons):
            complete_and_commit(
                parent,
                repo,
                main_number,
                1,
                "未完成",
                "单文件",
                reason(process_text),
                reason(product_text),
                reason(merged_text),
            )
            fix = insert_fix(parent, None, main_number, "修复核心操作后的结果反馈仍不一致的问题。")
            complete_and_commit(
                parent,
                repo,
                main_number,
                int(fix["row"]["轮次"]),
                "已完成",
                "单文件",
                "",
                "",
                "",
            )
        records = read_workbook(workbook_path(parent, None))
        checks.append({"name": "bugfix count reaches 7", "ok": bugfix_count(records) == 7})
        checks.append({
            "name": "0-1 multi-round main count reaches 5",
            "ok": choose_next(records)["multi_round_counts"]["0-1代码生成"] == 5,
        })
        blocked_zero_to_one = zero_to_one_mains[5]
        complete_and_commit(
            parent,
            repo,
            blocked_zero_to_one,
            1,
            "未完成",
            "单文件",
            reason("目标抓偏到局部展示，没有先把用户最关心的结果归属口径统一起来。"),
            reason("用户完成操作后关键结果还是不可见，界面没有给出可信反馈。"),
            reason("目标抓偏到局部展示，用户完成操作后关键结果还是不可见。"),
        )
        records = read_workbook(workbook_path(parent, None))
        next_after_limit = choose_next(records)
        checks.append({
            "name": "after 5 multi-round mains next main only",
            "ok": next_after_limit["row"]["主提示词编号"] != str(blocked_zero_to_one),
            "next_main": next_after_limit["row"]["主提示词编号"],
        })

        timeout_main = int(next_after_limit["row"]["主提示词编号"])
        update_row(parent, None, timeout_main, 1, {"执行状态": "已发送", "备注": "dry-run 已发送"})
        mark_session(parent, timeout_main, 1, f"dryrun-session-main-{timeout_main}-r1", "Trae运行中", "dry-run 运行中")
        update_row(parent, None, timeout_main, 1, {"执行状态": "超时待人工", "备注": "dry-run 超时待人工"})
        records = read_workbook(workbook_path(parent, None))
        checks.append({"name": "timeout waits for human", "ok": choose_next(records)["action"] == "confirm"})
        update_row(parent, None, timeout_main, 1, {"执行状态": "待验收", "备注": "dry-run 恢复后待验收"})
        records = read_workbook(workbook_path(parent, None))
        checks.append({"name": "resume pending acceptance", "ok": choose_next(records)["action"] == "accept"})

        checks.append({
            "name": "web frontend multi-file changes stay inside frontend modules",
            "ok": classify_change_range(["src/main.js", "components/Panel.vue", "styles/app.css", "vite.config.ts"], "Web 前端") == "跨模块多文件",
        })
        checks.append({
            "name": "frontend and backend changes become cross-system",
            "ok": classify_change_range(["src/main.js", "server/api.py"], "全栈 Web 应用") == "跨系统多模块",
        })
        fullstack_repo = make_fullstack_repo(root)
        checks.append({
            "name": "domain detects frontend and backend as fullstack",
            "ok": detect_domain(fullstack_repo) == "全栈 Web 应用",
            "domain": detect_domain(fullstack_repo),
        })

        bad_cases = [
            bad_audit_case(root, "bad-short-commit-id", {"Commit ID": "abc123"}),
            bad_audit_case(root, "bad-random-commit-id", {"Commit ID": "f" * 40}),
            bad_audit_case(root, "bad-commit-message", {"Trae Session ID": "different-session-id"}),
            bad_audit_case(root, "bad-missing-commit-id", {"Commit ID": ""}),
            bad_audit_case(root, "bad-skipped-executed", {"执行状态": "已跳过"}),
            bad_audit_case(root, "bad-commit-no-change-range", {"修改范围": "无需修改"}),
            bad_audit_case(root, "bad-satisfied-reason", {"不满意原因": "过程不满意：目标抓偏到局部展示。\n产物不满意：页面反馈不一致。"}),
            bad_audit_case(
                root,
                "bad-browser-reason",
                {
                    "任务是否完成": "未完成",
                    "过程与产物是否满意": "不满意",
                    "不满意原因": "过程不满意：浏览器没验到关键链路。\n产物不满意：页面反馈不一致。",
                },
            ),
            bad_audit_case(
                root,
                "bad-vague-reason",
                {
                    "任务是否完成": "未完成",
                    "过程与产物是否满意": "不满意",
                    "不满意原因": "过程不满意：过程比较乱，没完成任务。\n产物不满意：代码有 bug，效果不好。",
                },
            ),
            bad_audit_case(
                root,
                "bad-external-reason",
                {
                    "任务是否完成": "未完成",
                    "过程与产物是否满意": "不满意",
                    "不满意原因": "过程不满意：模型请求失败导致没有继续处理。\n产物不满意：网络波动导致页面结果没有出来。",
                },
            ),
            bad_duplicate_reason_case(root),
            bad_duplicate_reason_write_case(root),
            reason_check_regenerate_case(root),
        ]
        checks.append({
            "name": "audit rejects known bad workbook patterns",
            "ok": all(item["ok"] for item in bad_cases),
            "cases": bad_cases,
        })

        final_validation = audit_records(parent, None, repo)
        checks.append({"name": "final workbook audit", "ok": final_validation["ok"], "errors": final_validation["errors"], "git_errors": final_validation["git_errors"]})

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
