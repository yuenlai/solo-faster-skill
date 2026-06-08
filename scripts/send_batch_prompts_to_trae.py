#!/usr/bin/env python3
"""Submit and monitor solo-faster v2 prompts in one Trae project window.

This script owns workbook state transitions, send ordering, and log-based
identity checks. Trae UI control may still be performed by computer-use in the
outer workflow when the native accessibility tree is insufficient to reliably
focus the current task input box.
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from tempfile import NamedTemporaryFile

from batch_prompt_workbook import pick, update_row, workbook_path, read_workbook, choose_next
from runtime_lock import (
    RuntimeLockError,
    acquire_runtime_lock,
    release_runtime_lock,
    runtime_lock_status,
    touch_owner,
)


SKILL_DIR = Path(__file__).resolve().parents[1]
ENSURE_TRAE = SKILL_DIR / "scripts" / "ensure_trae_project_open.py"
TRAE_APP = "Trae CN"
TRAE_PROCESS = "TRAE CN"
DEFAULT_LOCK_STAGE = "real-execution"


def run_command(*args: str, input_text: str | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(args, input=input_text, capture_output=True, text=True)


def lock_stage_label(stage: str, row: dict[str, str] | None = None) -> str:
    if not row:
        return stage
    main = row.get("主提示词编号", "").strip()
    round_number = row.get("轮次", "").strip()
    if main and round_number:
        return f"{stage}:main-{main}:round-{round_number}"
    return stage


def acquire_real_execution_lock(repo: Path, stage: str, dry_run: bool) -> dict[str, object]:
    if dry_run:
        return {"mode": "dry-run", "stage": stage, "repo": str(repo)}
    return acquire_runtime_lock(repo, stage, wait=True)


def heartbeat_real_execution_lock(repo: Path, stage: str, dry_run: bool) -> None:
    if dry_run:
        return
    touch_owner(repo, stage)


def copy_to_clipboard(text: str) -> None:
    result = run_command("pbcopy", input_text=text)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "复制提示词失败")


def open_project(project_path: Path) -> None:
    result = run_command(sys.executable, str(ENSURE_TRAE), "--project-path", str(project_path))
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "打开 Trae 项目失败")


def focused_window_title(project_name: str) -> str:
    script = f'''
on run argv
  set projectName to item 1 of argv
  tell application "{TRAE_APP}" to activate
  delay 0.5
  tell application "System Events"
    tell process "{TRAE_PROCESS}"
      set frontmost to true
      repeat with candidateWindow in windows
        if (name of candidateWindow contains projectName) then
          perform action "AXRaise" of candidateWindow
          delay 0.3
          return name of candidateWindow
        end if
      end repeat
      if (count of windows) > 0 then
        return "front-window:" & name of window 1
      end if
    end tell
  end tell
  return "no-window"
end run
'''
    result = run_command("osascript", "-e", script, project_name)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "切换 Trae 项目窗口失败")
    return result.stdout.strip()


def confirm_project_window(project_path: Path) -> str:
    title = focused_window_title(project_path.name)
    if project_path.name not in title:
        raise RuntimeError(f"未聚焦到目标 Trae 窗口：{title or '无窗口'}")
    return title


def submit_clipboard(project_name: str, before_enter_delay: float, after_enter_delay: float) -> None:
    script = f'''
on run argv
  set projectName to item 1 of argv
  set beforeEnterDelay to item 2 of argv as real
  tell application "{TRAE_APP}" to activate
  delay 0.3
  tell application "System Events"
    tell process "{TRAE_PROCESS}"
      set frontmost to true
      if (count of windows) = 0 then error "Trae 没有可用窗口"
      set targetWindow to window 1
      if (name of targetWindow does not contain projectName) then error "前台窗口不是目标项目：" & name of targetWindow
      set {{xPos, yPos}} to position of targetWindow
      set {{winWidth, winHeight}} to size of targetWindow
      set inputX to xPos + (winWidth div 2)
      set inputY to yPos + winHeight - 92
      click at {{inputX, inputY}}
      delay 0.2
      click at {{inputX, inputY}}
      delay 0.2
      keystroke "v" using command down
      delay beforeEnterDelay
      key code 36
    end tell
  end tell
  delay {after_enter_delay}
end run
'''
    result = run_command("osascript", "-e", script, project_name, str(before_enter_delay))
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "发送提示词失败")


def trigger_new_task_shortcut(project_name: str) -> str:
    before_signature = screenshot_signature()
    script = f'''
on run argv
  set projectName to item 1 of argv
  tell application "{TRAE_APP}" to activate
  delay 0.4
  tell application "System Events"
    tell process "{TRAE_PROCESS}"
      set frontmost to true
      if (count of windows) = 0 then error "Trae 没有可用窗口"
      if (name of window 1 does not contain projectName) then error "前台窗口不是目标项目：" & name of window 1
      key code 45 using {{control down, command down}}
      delay 1.2
      return "shortcut:new-task"
    end tell
  end tell
end run
'''
    result = run_command("osascript", "-e", script, project_name)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "触发 Trae 新任务快捷键失败")
    detection = await_new_task_ready(project_name, before_signature)
    if detection.get("state") != "ready":
        evidence = json.dumps(detection, ensure_ascii=False)
        raise RuntimeError(f"Trae 新任务快捷键执行后未检测到可输入的新任务状态：evidence={evidence}")
    return f"{result.stdout.strip() or 'shortcut:new-task'} evidence={json.dumps(detection, ensure_ascii=False)}"


def screenshot_signature() -> list[tuple[int, int, int]]:
    with NamedTemporaryFile(suffix=".png", delete=True) as image_file:
        result = run_command("screencapture", "-x", image_file.name)
        if result.returncode != 0:
            raise RuntimeError(result.stderr.strip() or "截图失败，无法验证 Trae 新任务")
        try:
            from PIL import Image
            image = Image.open(image_file.name).convert("RGB")
        except Exception as exc:  # noqa: BLE001
            raise RuntimeError(f"读取截图失败，无法验证 Trae 新任务：{exc}") from exc
        width, height = image.size
        crop = image.crop((0, 70, min(900, width), min(height, 760))).resize((90, 69))
        return list(crop.getdata())


def screenshot_diff_score(before: list[tuple[int, int, int]], after: list[tuple[int, int, int]]) -> float:
    if not before or len(before) != len(after):
        return 0.0
    total = 0
    for (r1, g1, b1), (r2, g2, b2) in zip(before, after):
        total += abs(r1 - r2) + abs(g1 - g2) + abs(b1 - b2)
    return total / (len(before) * 3)


def collect_window_affordance(project_name: str) -> dict[str, object]:
    script = f'''
on run argv
  set projectName to item 1 of argv
  tell application "System Events"
    tell process "{TRAE_PROCESS}"
      if (count of windows) = 0 then return "{{\\"state\\":\\"unknown\\",\\"reason\\":\\"no-window\\"}}"
      if (name of window 1 does not contain projectName) then return "{{\\"state\\":\\"unknown\\",\\"reason\\":\\"wrong-window\\"}}"
      set uiText to ""
      try
        set uiText to value of every static text of window 1 as text
      end try
      set buttonText to ""
      try
        set buttonText to name of every button of window 1 as text
      end try
      set descText to ""
      try
        set descText to description of every UI element of window 1 as text
      end try
      return "{{\\"state\\":\\"ok\\",\\"ui_text\\":" & quoted form of uiText & ",\\"button_text\\":" & quoted form of buttonText & ",\\"desc_text\\":" & quoted form of descText & "}}"
    end tell
  end tell
end run
'''
    result = run_command("osascript", "-e", script, project_name)
    if result.returncode != 0:
        return {"state": "unknown", "reason": result.stderr.strip() or "osascript failed"}
    text = result.stdout.strip()
    try:
        parsed = json.loads(text)
    except json.JSONDecodeError:
        normalized = re.sub(r":''", ':""', text)
        normalized = re.sub(r",'([^']*)'", lambda match: ',"' + match.group(1).replace('"', '\\"') + '"', normalized)
        normalized = re.sub(r":'([^']*)'", lambda match: ':"' + match.group(1).replace('"', '\\"') + '"', normalized)
        try:
            parsed = json.loads(normalized)
        except json.JSONDecodeError:
            return {"state": "unknown", "reason": "unparseable-affordance-output", "raw": text}
    if not isinstance(parsed, dict):
        return {"state": "unknown", "reason": "invalid-affordance-output", "raw": text}
    return parsed


def detect_new_task_ready(project_name: str, before_signature: list[tuple[int, int, int]] | None = None) -> dict[str, object]:
    focused_title = focused_window_title(project_name)
    if project_name not in focused_title:
        return {"state": "unknown", "reason": "wrong-window", "window": focused_title}

    affordance = collect_window_affordance(project_name)
    if affordance.get("state") != "ok":
        affordance["window"] = focused_title
        return affordance

    all_text = " ".join(
        [
            str(affordance.get("ui_text") or ""),
            str(affordance.get("button_text") or ""),
            str(affordance.get("desc_text") or ""),
        ]
    )

    after_signature = screenshot_signature()
    diff_score = screenshot_diff_score(before_signature or [], after_signature) if before_signature else 0.0
    has_old_task_markers = any(
        marker in all_text
        for marker in ["任务完成", "代码变更", "产物汇总", "重新变更", "验证结果", "最终技术栈"]
    )
    has_welcome_markers = any(
        marker in all_text
        for marker in ["擅长项目迭代", "智能任务规划", "自主编排智能体"]
    )
    has_input_ready = any(marker in all_text for marker in ["Send", "发送", "Ask", "提问", "新任务", "New Task"])
    evidence = {
        "diff_score": round(diff_score, 2),
        "has_old_task_markers": has_old_task_markers,
        "has_welcome_markers": has_welcome_markers,
        "has_input_ready": has_input_ready,
    }

    if has_old_task_markers:
        return {"state": "not-ready", "reason": "old-task-content-still-visible", "evidence": evidence, "window": focused_title}
    if has_welcome_markers and has_input_ready and diff_score >= 8.0:
        return {"state": "ready", "reason": "welcome-screen-visible", "evidence": evidence, "window": focused_title}
    if has_input_ready and diff_score >= 20.0:
        return {"state": "ready", "reason": "input-ready-with-significant-ui-change", "evidence": evidence, "window": focused_title}
    return {"state": "not-ready", "reason": "new-task-not-confirmed", "evidence": evidence, "window": focused_title}


def await_new_task_ready(
    project_name: str,
    before_signature: list[tuple[int, int, int]],
    timeout_seconds: float = 6.0,
    interval_seconds: float = 0.6,
) -> dict[str, object]:
    deadline = time.monotonic() + timeout_seconds
    last: dict[str, object] = {"state": "not-ready", "reason": "new-task-not-confirmed"}
    while time.monotonic() < deadline:
        last = detect_new_task_ready(project_name, before_signature)
        if last.get("state") == "ready":
            return last
        time.sleep(interval_seconds)
    return last


def detect_completion(
    project_name: str,
    repo_path: Path | None = None,
    since: str | None = None,
    session_ref: str | None = None,
    prompt: str | None = None,
) -> dict[str, object]:
    """Best-effort UI check. Returns a conservative state plus evidence.

    Important rule: running evidence wins. Completion is accepted only when the
    UI exposes a send/input-ready affordance and no running affordance is found.
    Generic words such as Done/完成/Apply are not enough because they also occur
    in completed checklist items while Trae is still generating.
    """
    focused_title = focused_window_title(project_name)
    if project_name not in focused_title:
        return {"state": "unknown", "reason": "wrong-window", "window": focused_title}

    visual = detect_completion_by_screenshot()
    if visual["state"] in {"running", "done"}:
        visual["window"] = focused_title
        return visual

    affordance = collect_window_affordance(project_name)
    if affordance.get("state") != "ok":
        affordance["window"] = focused_title
        return affordance
    all_text = " ".join(
        [
            str(affordance.get("ui_text") or ""),
            str(affordance.get("button_text") or ""),
            str(affordance.get("desc_text") or ""),
        ]
    )
    if any(marker in all_text for marker in ["Stop", "停止", "Cancel", "取消生成", "Generating", "生成中", "正在生成", "Thinking", "思考中", "Running", "运行中", "Applying", "正在应用", "Continue", "继续生成"]):
        parsed = {"state": "running", "reason": "running-affordance"}
    elif "任务完成" in all_text:
        parsed = {"state": "done", "reason": "task-complete-text"}
    elif any(marker in all_text for marker in ["Send", "发送", "Ask", "提问", "New Task", "新任务"]):
        parsed = {"state": "done", "reason": "input-ready-no-running-affordance"}
    else:
        parsed = {"state": "unknown", "reason": "no-decisive-affordance"}
    parsed["window"] = focused_title
    if parsed.get("state") == "unknown" and repo_path:
        log_detection = detect_completion_by_logs(repo_path, since, session_ref, prompt)
        if log_detection.get("state") in {"running", "done"}:
            log_detection["window"] = focused_title
            return log_detection
    return parsed


def detect_completion_by_logs(
    repo: Path,
    since: str | None = None,
    session_ref: str | None = None,
    prompt: str | None = None,
) -> dict[str, object]:
    log_root = Path.home() / "Library" / "Application Support" / "Trae CN" / "logs"
    if not log_root.exists():
        return {"state": "unknown", "reason": "trae-log-root-missing"}
    candidates = latest_ai_agent_logs()
    if not candidates:
        return {"state": "unknown", "reason": "trae-ai-agent-log-missing"}
    repo_text = str(repo)
    since_text = since.replace(" ", "T") if since else ""
    identity = parse_session_ref(session_ref)
    if not identity and prompt:
        identity = find_prompt_log_identity(candidates[:3], prompt, since_text)
    if not identity:
        return {"state": "unknown", "reason": "trae-log-no-prompt-session", "logs": [str(path) for path in candidates[:3]]}
    for log_path in candidates[:3]:
        evidence = latest_finish_log_evidence(log_path, repo_text, since_text, identity)
        if evidence:
            return {"state": "done", "reason": "trae-log-finish-tool", "evidence": evidence}
    return {"state": "unknown", "reason": "trae-log-no-matching-finish-tool", "identity": identity, "logs": [str(path) for path in candidates[:3]]}


def latest_ai_agent_logs() -> list[Path]:
    log_root = Path.home() / "Library" / "Application Support" / "Trae CN" / "logs"
    if not log_root.exists():
        return []
    return sorted(log_root.glob("*/Modular/ai-agent_*_stdout.log"), key=lambda path: path.stat().st_mtime, reverse=True)


def latest_renderer_logs() -> list[Path]:
    log_root = Path.home() / "Library" / "Application Support" / "Trae CN" / "logs"
    if not log_root.exists():
        return []
    return sorted(log_root.glob("*/window*/renderer.log"), key=lambda path: path.stat().st_mtime, reverse=True)


def trae_user_id() -> str:
    log_root = Path.home() / "Library" / "Application Support" / "Trae CN" / "logs"
    if not log_root.exists():
        return ""
    for log_path in sorted(log_root.glob("*/main.log"), key=lambda path: path.stat().st_mtime, reverse=True):
        text = read_log_tail(log_path, max_bytes=2_000_000)
        match = re.search(r'"userId":"(\d+)"', text)
        if match:
            return match.group(1)
    return ""


def read_log_tail(log_path: Path, max_bytes: int = 32_000_000) -> str:
    try:
        with log_path.open("rb") as handle:
            handle.seek(0, 2)
            size = handle.tell()
            handle.seek(max(0, size - max_bytes))
            return handle.read().decode("utf-8", errors="ignore")
    except OSError:
        return ""


def prompt_terms(prompt: str) -> list[str]:
    compact = " ".join(prompt.strip().split())
    terms: list[str] = []
    for sep in ["，", "。", "；", ",", ";"]:
        head = compact.split(sep, 1)[0].strip()
        if len(head) >= 8:
            terms.append(head)
            break
    if len(compact) >= 12:
        terms.append(compact[:18])
        terms.append(compact[:12])
    terms.append(compact)
    return list(dict.fromkeys(term for term in terms if term))


def parse_session_ref(session_ref: str | None) -> dict[str, str]:
    if not session_ref:
        return {}
    identity: dict[str, str] = {}
    for key in ["user_id", "trace_id", "session_id", "task_id", "message_id", "user_message_id", "run_id", "time", "app"]:
        match = re.search(rf"{key}=([A-Za-z0-9_-]+)", session_ref)
        if match:
            identity[key] = match.group(1)
    if identity:
        return identity
    full_match = re.search(
        r"\.?(?P<user_id>\d+):(?P<trace_id>[A-Za-z0-9]+)_(?P<session_id>[A-Za-z0-9]+)\.(?P<message_id>[A-Za-z0-9]+)\.(?P<user_message_id>[A-Za-z0-9]+):(?P<app>.+?)\.T\((?P<time>[^)]+)\)",
        session_ref,
    )
    if full_match:
        identity.update({key: value for key, value in full_match.groupdict().items() if value})
    return identity


def format_session_ref(identity: dict[str, str]) -> str:
    user_id = identity.get("user_id") or trae_user_id()
    trace_id = identity.get("trace_id", "")
    session_id = identity.get("session_id", "")
    message_id = identity.get("message_id", "")
    user_message_id = identity.get("user_message_id", "")
    app = identity.get("app") or "Trae CN"
    time_text = identity.get("time", "")
    if user_id and trace_id and session_id and message_id and user_message_id and time_text:
        return f".{user_id}:{trace_id}_{session_id}.{message_id}.{user_message_id}:{app}.T({time_text})"
    return ";".join(f"{key}={value}" for key, value in identity.items() if value)


def normalize_session_time(timestamp: str) -> str:
    try:
        dt = datetime.fromisoformat(timestamp)
    except ValueError:
        return ""
    return f"{dt.year}/{dt.month}/{dt.day} {dt.hour:02d}:{dt.minute:02d}:{dt.second:02d}"


def find_prompt_log_identity(log_paths: list[Path], prompt: str, since_text: str) -> dict[str, str]:
    terms = prompt_terms(prompt)
    best: dict[str, str] = {}
    best_time = ""
    for log_path in log_paths:
        text = read_log_tail(log_path)
        for line in text.splitlines():
            timestamp = log_timestamp(line)
            if since_text and timestamp and timestamp < since_text:
                continue
            if not any(term in line for term in terms):
                continue
            if "generate_session_title_and_icon_cloud" not in line and "title:" not in line:
                continue
            identity = extract_log_identity(line)
            session_match = re.search(r'session_id:\s*"([A-Za-z0-9_-]+)"', line)
            if session_match:
                identity["session_id"] = session_match.group(1)
            if identity.get("session_id"):
                if timestamp:
                    identity["time"] = normalize_session_time(timestamp)
                identity["app"] = "Trae CN"
                identity["user_id"] = trae_user_id()
                identity["prompt_term"] = next(term for term in terms if term in line)
                identity["log"] = str(log_path)
                if not best_time or (timestamp and timestamp >= best_time):
                    best = identity
                    best_time = timestamp
    if best:
        enrich_identity(best, since_text)
    return best


def await_prompt_log_identity(prompt: str, since_text: str, timeout_seconds: float = 8.0, interval_seconds: float = 1.0) -> dict[str, str]:
    deadline = time.monotonic() + timeout_seconds
    while time.monotonic() < deadline:
        identity = find_prompt_log_identity(latest_ai_agent_logs(), prompt, since_text)
        if identity:
            return identity
        time.sleep(interval_seconds)
    return {}


def extract_log_identity(line: str) -> dict[str, str]:
    identity: dict[str, str] = {}
    patterns = {
        "session_id": [r'session_id[:=]\s*"?([A-Za-z0-9_-]+)"?'],
        "trace_id": [r'trace_id[:=]\s*"?([A-Za-z0-9_-]+)"?'],
        "task_id": [r'"task_id":"([^"]+)"', r"task_id=([A-Za-z0-9_-]+)"],
        "message_id": [r"\bmessage_id=([A-Za-z0-9_-]+)", r'"message_id":"([^"]+)"'],
        "user_message_id": [r"user_message_id[:=]\s*\"?([A-Za-z0-9_-]+)\"?", r"turnId=([A-Za-z0-9_-]+)"],
        "run_id": [r'"agent_run_id":"([^"]+)"', r"run_id=([A-Za-z0-9_-]+)"],
    }
    for key, key_patterns in patterns.items():
        for pattern in key_patterns:
            match = re.search(pattern, line)
            if match:
                identity[key] = match.group(1)
                break
    return identity


def renderer_identity_for_session(session_id: str, since_text: str) -> dict[str, str]:
    identity: dict[str, str] = {}
    for log_path in latest_renderer_logs()[:3]:
        text = read_log_tail(log_path)
        for line in text.splitlines():
            if session_id not in line:
                continue
            timestamp = log_timestamp(line)
            if since_text and timestamp and timestamp < since_text:
                continue
            if "code_comp_trigger" in line:
                session_match = re.search(r'"session_id":"([A-Za-z0-9_-]+)"', line)
                user_message_match = re.search(r'"message_id":"([A-Za-z0-9_-]+)"', line)
                if session_match:
                    identity["session_id"] = session_match.group(1)
                if user_message_match:
                    identity["user_message_id"] = user_message_match.group(1)
                if timestamp:
                    identity["time"] = normalize_session_time(timestamp)
            elif "ChatStreamFrontResponseReporter" in line:
                trace_match = re.search(r'"traceId":"([A-Za-z0-9_-]+)"', line)
                if trace_match:
                    identity["trace_id"] = trace_match.group(1)
        if identity.get("time") and identity.get("user_message_id") and identity.get("trace_id"):
            break
    return identity


def ai_identity_for_session(session_id: str, since_text: str) -> dict[str, str]:
    identity: dict[str, str] = {}
    for log_path in latest_ai_agent_logs()[:3]:
        text = read_log_tail(log_path)
        lines = text.splitlines()
        anchor_index: int | None = None
        for idx, line in enumerate(lines):
            timestamp = log_timestamp(line)
            if since_text and timestamp and timestamp < since_text:
                continue
            if session_id not in line:
                continue
            if "new_user_message_id:" in line or "pre_user_message_id=" in line:
                user_message_match = re.search(r"(?:new_user_message_id:\s*|pre_user_message_id=)([A-Za-z0-9_-]+)", line)
                trace_match = re.search(r'trace_id="?([A-Za-z0-9_-]+)"?', line)
                if user_message_match:
                    identity["user_message_id"] = user_message_match.group(1)
                if trace_match:
                    identity["trace_id"] = trace_match.group(1)
                if timestamp and not identity.get("time"):
                    identity["time"] = normalize_session_time(timestamp)
                anchor_index = idx
                break
        if anchor_index is None:
            continue
        for line in lines[anchor_index: min(anchor_index + 800, len(lines))]:
            if session_id not in line:
                continue
            if identity.get("trace_id") and identity["trace_id"] not in line:
                continue
            task_match = re.search(r"task_id=([A-Za-z0-9_-]+)", line)
            middle_message_match = re.search(r"\bmessage_id=([A-Za-z0-9_-]+)", line)
            user_message_match = re.search(r"user_message_id:\s*([A-Za-z0-9_-]+)", line)
            if task_match and not identity.get("task_id"):
                identity["task_id"] = task_match.group(1)
            if middle_message_match and not identity.get("message_id"):
                identity["message_id"] = middle_message_match.group(1)
            if user_message_match and not identity.get("user_message_id"):
                identity["user_message_id"] = user_message_match.group(1)
            if identity.get("trace_id") and identity.get("message_id") and identity.get("user_message_id"):
                break
        if identity.get("trace_id") and identity.get("message_id") and identity.get("user_message_id"):
            break
    return identity


def enrich_identity(identity: dict[str, str], since_text: str) -> dict[str, str]:
    session_id = identity.get("session_id", "")
    if not session_id:
        return identity
    renderer_identity = renderer_identity_for_session(session_id, since_text)
    ai_identity = ai_identity_for_session(session_id, since_text)
    if renderer_identity.get("time"):
        identity["time"] = renderer_identity["time"]
    for source in [renderer_identity, ai_identity]:
        for key, value in source.items():
            if value and not identity.get(key):
                identity[key] = value

    if not identity.get("app"):
        identity["app"] = "Trae CN"
    if not identity.get("user_id"):
        identity["user_id"] = trae_user_id()
    return identity


def log_timestamp(line: str) -> str:
    match = re.match(r"(\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2})", line)
    return match.group(1) if match else ""


def line_matches_identity(line: str, identity: dict[str, str]) -> bool:
    if not identity:
        return False
    for key in ["trace_id", "session_id", "task_id", "message_id", "run_id"]:
        value = identity.get(key)
        if value and value in line:
            return True
    return False


def latest_finish_log_evidence(log_path: Path, repo_text: str, since_text: str, identity: dict[str, str]) -> dict[str, str] | None:
    text = read_log_tail(log_path)
    if not text:
        return None
    best: dict[str, str] | None = None
    for line in text.splitlines():
        if "tools::finish" not in line or repo_text not in line:
            continue
        if not line_matches_identity(line, identity):
            continue
        timestamp = log_timestamp(line)
        if since_text and timestamp and timestamp < since_text:
            continue
        best = {"log": str(log_path), "time": timestamp, "identity": format_session_ref(identity), "snippet": line[:500]}
    return best


def lookup_prompt_identity(prompt: str, since_text: str = "", log_limit: int = 3) -> dict[str, object]:
    logs = latest_ai_agent_logs()
    identity = find_prompt_log_identity(logs[:log_limit], prompt, since_text)
    return {"identity": identity, "logs": [str(path) for path in logs[:log_limit]]}


def detect_completion_by_screenshot() -> dict[str, object]:
    with NamedTemporaryFile(suffix=".png", delete=True) as image_file:
        result = run_command("screencapture", "-x", image_file.name)
        if result.returncode != 0:
            return {"state": "unknown", "reason": result.stderr.strip() or "screenshot-failed"}
        try:
            from PIL import Image
            image = Image.open(image_file.name).convert("RGB")
        except Exception as exc:  # noqa: BLE001
            return {"state": "unknown", "reason": f"screenshot-read-failed:{exc}"}
        width, height = image.size
        left = image.crop((0, 70, min(420, width), min(560, height)))
        bottom_right = image.crop((max(0, width - 320), max(0, height - 260), width, height))
        left_stats = color_stats(left)
        button_stats = color_stats(bottom_right)

    running_score = left_stats["blue"] + button_stats["green_stop"]
    done_score = left_stats["teal_done"] + button_stats["green_send"]
    evidence = {
        "left": left_stats,
        "bottom_right": button_stats,
        "running_score": running_score,
        "done_score": done_score,
    }
    if running_score >= 120:
        return {"state": "running", "reason": "visual-running-marker", "evidence": evidence}
    if done_score >= 180 and running_score < 80:
        return {"state": "done", "reason": "visual-done-marker", "evidence": evidence}
    return {"state": "unknown", "reason": "visual-no-decisive-marker", "evidence": evidence}


def color_stats(image: object) -> dict[str, int]:
    teal_done = 0
    blue = 0
    green_send = 0
    green_stop = 0
    for r, g, b in image.getdata():
        if 35 <= r <= 90 and 160 <= g <= 230 and 130 <= b <= 210:
            teal_done += 1
        if 20 <= r <= 90 and 80 <= g <= 160 and 170 <= b <= 255:
            blue += 1
        if 25 <= r <= 90 and 180 <= g <= 255 and 120 <= b <= 220:
            green_send += 1
        if 20 <= r <= 100 and 160 <= g <= 255 and 100 <= b <= 210:
            green_stop += 1
    return {
        "teal_done": teal_done,
        "blue": blue,
        "green_send": green_send,
        "green_stop": green_stop,
    }


def monitor_event(row: dict[str, str], event: str, detail: dict[str, object] | None = None) -> dict[str, object]:
    payload: dict[str, object] = {
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "main": row["主提示词编号"],
        "round": row["轮次"],
        "event": event,
    }
    if detail:
        payload["detail"] = detail
    return payload


def row_identity(row: dict[str, str]) -> tuple[int, int]:
    return int(row["主提示词编号"]), int(row["轮次"])


def row_send_mode(row: dict[str, str]) -> str:
    _, round_number = row_identity(row)
    return "new-task" if round_number == 1 else "same-task"


def manual_send_instruction(records: list[dict[str, str]], row: dict[str, str]) -> dict[str, object]:
    main, round_number = row_identity(row)
    new_task_required = requires_new_task_before_send(records, row)
    row_type = "主提示词" if round_number == 1 else "修复提示词"
    turn_label = f"【第 {main} 个主提示词，第 {round_number} 轮】"
    if new_task_required:
        instruction = f"{turn_label}【需要新建任务】请发送下面这条主提示词。"
    elif row_send_mode(row) == "same-task":
        instruction = f"{turn_label}【不要新建任务】请在当前任务下发送下面这条修复提示词。"
    else:
        instruction = f"{turn_label}请发送下面这条主提示词。"
    return {
        "main": str(main),
        "round": str(round_number),
        "row_type": row_type,
        "turn_label": turn_label,
        "new_task_required": new_task_required,
        "instruction": instruction,
        "prompt": row.get("提示词", "").strip(),
    }


def mark(parent: Path, workbook: str | None, row: dict[str, str], **fields: str) -> dict[str, object]:
    main, round_number = row_identity(row)
    return update_row(parent, workbook, main, round_number, fields)


def select_rows(parent: Path, workbook: str | None, number_range: str | None, limit: int | None, status_filter: str | None) -> list[dict[str, str]]:
    statuses = set(status_filter.split(",")) if status_filter else {"已生成"}
    selected = pick(parent, workbook, number_range, limit, statuses)
    return list(selected["items"])


def enforce_next_send(records: list[dict[str, str]], rows: list[dict[str, str]]) -> None:
    if not rows:
        return
    next_action = choose_next(records)
    if next_action.get("action") != "send":
        raise RuntimeError(f"当前下一步不是发送提示词：{json.dumps(next_action, ensure_ascii=False)}")
    expected = next_action.get("row")
    if not isinstance(expected, dict):
        raise RuntimeError(f"无法确认下一条待发送提示词：{json.dumps(next_action, ensure_ascii=False)}")
    expected_identity = row_identity(expected)
    selected_identity = row_identity(rows[0])
    if selected_identity != expected_identity:
        raise RuntimeError(
            "禁止绕过断点续跑顺序发送提示词："
            f"下一步应发送主提示词 {expected_identity[0]} 第 {expected_identity[1]} 轮，"
            f"但当前选择了主提示词 {selected_identity[0]} 第 {selected_identity[1]} 轮"
        )
    expected_main = expected_identity[0]
    expected_round = expected_identity[1]
    for row in rows[1:]:
        main, round_number = row_identity(row)
        if main != expected_main or round_number != expected_round + 1:
            raise RuntimeError("一次发送只能沿同一主提示词的连续轮次执行，不能跨到新的主提示词")
        expected_round = round_number


def previous_main_is_terminal(records: list[dict[str, str]], main_number: int) -> bool:
    if main_number <= 1:
        return False
    return terminal_for_send(records, main_number - 1)


def terminal_for_send(records: list[dict[str, str]], main_number: int) -> bool:
    rows = [row for row in records if int(row.get("主提示词编号") or 0) == main_number]
    if any(row.get("执行状态") == "已完成" or row.get("任务是否完成") == "已完成" for row in rows):
        return True
    if not rows:
        return False
    latest = max(rows, key=lambda row: int(row.get("轮次") or 0))
    return int(latest.get("轮次") or 0) >= 3 and (latest.get("执行状态") == "未完成" or latest.get("任务是否完成") == "未完成")


def requires_new_task_before_send(records: list[dict[str, str]], row: dict[str, str]) -> bool:
    main, round_number = row_identity(row)
    return round_number == 1 and previous_main_is_terminal(records, main)


def send_rows(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent).expanduser().resolve()
    repo = Path(args.repo).expanduser().resolve()
    records = read_workbook(workbook_path(parent, args.workbook))
    rows = select_rows(parent, args.workbook, args.range, args.limit, args.status_filter)
    if not args.allow_out_of_order:
        enforce_next_send(records, rows)
    if not args.auto_ui:
        instructions = [manual_send_instruction(records, row) for row in rows if row.get("提示词", "").strip()]
        return {
            "mode": "manual",
            "selected_count": len(rows),
            "instructions": instructions,
            "message": "低 CPU 模式默认不驱动 Trae UI，也不做窗口切换、截图检测或自动发送。",
            "next_step": "用户在 Trae 里手动发送后，先回填当前轮 Trae Session ID；只有回填成功后，才能把该轮推进到 Trae运行中。",
            "lock": {"mode": "disabled", "reason": "manual mode does not acquire runtime lock"},
        }

    lock_info = acquire_real_execution_lock(repo, f"{DEFAULT_LOCK_STAGE}:send", args.dry_run)
    try:
        sent: list[dict[str, str]] = []
        failed: list[dict[str, str]] = []
        if not args.dry_run:
            open_project(repo)
            time.sleep(args.open_delay)
            confirm_project_window(repo)
        for row in rows:
            prompt = row.get("提示词", "").strip()
            if not prompt:
                continue
            main, round_number = row_identity(row)
            try:
                if args.dry_run:
                    sent.append({"main": str(main), "round": str(round_number)})
                    continue
                heartbeat_real_execution_lock(repo, lock_stage_label("send", row), args.dry_run)
                confirm_project_window(repo)
                if requires_new_task_before_send(records, row) and not args.no_auto_new_task:
                    heartbeat_real_execution_lock(repo, lock_stage_label("new-task", row), args.dry_run)
                    trigger_new_task_shortcut(repo.name)
                copy_to_clipboard(prompt)
                submit_clipboard(repo.name, args.before_enter_delay, args.after_enter_delay)
                mark(
                    parent,
                    args.workbook,
                    row,
                    执行状态="已发送",
                    备注="已发送到 Trae，需立即从当前提示词对应的 SOLO Agent 回复复制 Session ID；回填前不得推进到 Trae运行中",
                )
                sent.append({"main": str(main), "round": str(round_number)})
            except Exception as exc:  # noqa: BLE001
                failed.append({"main": str(main), "round": str(round_number), "error": str(exc)})
        return {"selected_count": len(rows), "sent_count": len(sent), "sent": sent, "failed": failed, "lock": lock_info}
    finally:
        if not args.dry_run:
            release_runtime_lock(repo)


def monitor(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent).expanduser().resolve()
    repo = Path(args.repo).expanduser().resolve()
    records = read_workbook(workbook_path(parent, args.workbook))
    selected = [
        row
        for row in records
        if row.get("执行状态") in {"已发送", "Trae运行中", "超时待人工"}
        and (args.main is None or int(row.get("主提示词编号") or 0) == args.main)
        and (args.round is None or int(row.get("轮次") or 0) == args.round)
    ]
    events: list[dict[str, str]] = []
    if not selected:
        return {"monitored": 0, "events": events}
    if not args.active_monitor and not args.dry_run:
        return {
            "mode": "manual",
            "monitored": len(selected),
            "events": events,
            "message": "低 CPU 模式默认不持续轮询 Trae，不做截图分析、日志追踪或 UI 完成态检测。",
            "rows": [
                {
                    "main": row["主提示词编号"],
                    "round": row["轮次"],
                    "status": row["执行状态"],
                    "session_id": row.get("Trae Session ID", "").strip(),
                }
                for row in selected
            ],
            "next_step": "先确认所有已发送行都已回填 Trae Session ID；用户人工确认 Trae 已跑完后回复继续，再进入验收。",
            "lock": {"mode": "disabled", "reason": "manual monitor does not acquire runtime lock"},
        }

    lock_info = acquire_real_execution_lock(repo, f"{DEFAULT_LOCK_STAGE}:monitor", args.dry_run)
    try:
        if args.resume_timeout:
            for row in selected:
                if row.get("执行状态") == "超时待人工":
                    mark(parent, args.workbook, row, 执行状态="Trae运行中", 备注="用户确认后恢复监控")
                    row["执行状态"] = "Trae运行中"
        if args.dry_run:
            for row in selected:
                if args.simulate == "timeout":
                    mark(parent, args.workbook, row, 执行状态="超时待人工", 备注="dry-run 超时待人工")
                    events.append(monitor_event(row, "timeout", {"simulate": True}))
                elif args.simulate == "done":
                    mark(parent, args.workbook, row, 执行状态="待验收", 备注="dry-run Trae 完成，等待验收")
                    events.append(monitor_event(row, "done", {"simulate": True}))
                else:
                    events.append(monitor_event(row, "running", {"simulate": True}))
                print(json.dumps(events[-1], ensure_ascii=False), flush=True)
            return {"monitored": len(selected), "events": events, "lock": lock_info}

        confirm_project_window(repo)
        deadline = time.monotonic() + args.timeout_minutes * 60
        while True:
            first = selected[0]
            heartbeat_real_execution_lock(repo, lock_stage_label("monitor", first), args.dry_run)
            since = first.get("开始时间", "")
            detection = detect_completion(repo.name, repo, since, first.get("Trae Session ID", ""), first.get("提示词", ""))
            state = str(detection.get("state") or "unknown")
            if state == "done":
                for row in selected:
                    mark(parent, args.workbook, row, 执行状态="待验收", 备注="Trae 显示已完成，等待验收")
                    events.append(monitor_event(row, "done", detection))
                    print(json.dumps(events[-1], ensure_ascii=False), flush=True)
                break
            if time.monotonic() >= deadline:
                for row in selected:
                    mark(parent, args.workbook, row, 执行状态="超时待人工", 备注="超过 20 分钟仍未完成，请人工检查 Trae 后回复继续")
                    events.append(monitor_event(row, "timeout", detection))
                    print(json.dumps(events[-1], ensure_ascii=False), flush=True)
                break
            events.append(monitor_event(selected[0], state, detection))
            print(json.dumps(events[-1], ensure_ascii=False), flush=True)
            if args.once:
                break
            time.sleep(args.interval_minutes * 60)
        return {"monitored": len(selected), "events": events, "lock": lock_info}
    finally:
        if not args.dry_run:
            release_runtime_lock(repo)


def new_task(args: argparse.Namespace) -> dict[str, object]:
    repo = Path(args.repo).expanduser().resolve()
    lock_info = acquire_real_execution_lock(repo, f"{DEFAULT_LOCK_STAGE}:new-task", args.dry_run)
    try:
        if args.dry_run:
            return {"repo": str(repo), "new_task": "dry-run", "lock": lock_info}
        heartbeat_real_execution_lock(repo, f"{DEFAULT_LOCK_STAGE}:new-task", args.dry_run)
        confirm_project_window(repo)
        clicked = trigger_new_task_shortcut(repo.name)
        return {"repo": str(repo), "new_task": clicked, "lock": lock_info}
    finally:
        if not args.dry_run:
            release_runtime_lock(repo)


def lock_acquire(args: argparse.Namespace) -> dict[str, object]:
    repo = Path(args.repo).expanduser().resolve()
    return acquire_runtime_lock(repo, args.stage or DEFAULT_LOCK_STAGE, wait=not args.no_wait)


def lock_release(args: argparse.Namespace) -> dict[str, object]:
    repo = Path(args.repo).expanduser().resolve()
    return release_runtime_lock(repo, force=args.force)


def lock_status(args: argparse.Namespace) -> dict[str, object]:
    return runtime_lock_status()


def next_state(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent).expanduser().resolve()
    records = read_workbook(workbook_path(parent, args.workbook))
    return choose_next(records)


def lookup_identity(args: argparse.Namespace) -> dict[str, object]:
    return lookup_prompt_identity(args.prompt, args.since or "", args.log_limit)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="command", required=True)

    send = sub.add_parser("send")
    send.add_argument("--parent", required=True)
    send.add_argument("--repo", required=True)
    send.add_argument("--workbook")
    send.add_argument("--range")
    send.add_argument("--limit", type=int)
    send.add_argument("--status-filter")
    send.add_argument("--open-delay", type=float, default=2.5)
    send.add_argument("--before-enter-delay", type=float, default=0.8)
    send.add_argument("--after-enter-delay", type=float, default=2.5)
    send.add_argument("--dry-run", action="store_true")
    send.add_argument("--auto-ui", action="store_true", help="Explicitly enable Trae UI automation. Default is low-CPU manual mode.")
    send.add_argument("--allow-out-of-order", action="store_true", help="Manual recovery only: bypass next-state send protection")
    send.add_argument("--no-auto-new-task", action="store_true", help="Manual recovery only: do not create a new Trae task before a new main prompt")

    monitor_cmd = sub.add_parser("monitor")
    monitor_cmd.add_argument("--parent", required=True)
    monitor_cmd.add_argument("--repo", required=True)
    monitor_cmd.add_argument("--workbook")
    monitor_cmd.add_argument("--main", type=int)
    monitor_cmd.add_argument("--round", type=int)
    monitor_cmd.add_argument("--interval-minutes", type=float, default=2.0)
    monitor_cmd.add_argument("--timeout-minutes", type=float, default=20.0)
    monitor_cmd.add_argument("--once", action="store_true")
    monitor_cmd.add_argument("--resume-timeout", action="store_true")
    monitor_cmd.add_argument("--dry-run", action="store_true")
    monitor_cmd.add_argument("--active-monitor", action="store_true", help="Explicitly enable polling-based Trae monitoring. Default is low-CPU manual mode.")
    monitor_cmd.add_argument("--simulate", choices=["running", "done", "timeout"], default="running")

    task = sub.add_parser("new-task")
    task.add_argument("--repo", required=True)
    task.add_argument("--dry-run", action="store_true")

    lock_hold = sub.add_parser("lock-acquire")
    lock_hold.add_argument("--repo", required=True)
    lock_hold.add_argument("--stage")
    lock_hold.add_argument("--no-wait", action="store_true")

    lock_free = sub.add_parser("lock-release")
    lock_free.add_argument("--repo", required=True)
    lock_free.add_argument("--force", action="store_true")

    sub.add_parser("lock-status")

    nxt = sub.add_parser("next")
    nxt.add_argument("--parent", required=True)
    nxt.add_argument("--workbook")

    lookup = sub.add_parser("lookup-identity")
    lookup.add_argument("--prompt", required=True)
    lookup.add_argument("--since")
    lookup.add_argument("--log-limit", type=int, default=3)

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    commands = {
        "send": send_rows,
        "monitor": monitor,
        "new-task": new_task,
        "lock-acquire": lock_acquire,
        "lock-release": lock_release,
        "lock-status": lock_status,
        "next": next_state,
        "lookup-identity": lookup_identity,
    }
    try:
        result = commands[args.command](args)
    except Exception as exc:  # noqa: BLE001
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        raise SystemExit(2) from exc
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
