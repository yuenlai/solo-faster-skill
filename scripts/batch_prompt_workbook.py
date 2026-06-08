#!/usr/bin/env python3
"""Maintain the solo-faster v2 single-repository prompt workbook."""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
import zipfile
from datetime import datetime, timezone
from difflib import SequenceMatcher
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape


HEADERS = [
    "仓库路径",
    "主提示词编号",
    "轮次",
    "行类型",
    "执行状态",
    "Repo ID",
    "Trae Session ID",
    "提示词",
    "任务类型",
    "业务领域",
    "修改范围",
    "任务难度",
    "任务是否完成",
    "过程与产物是否满意",
    "不满意原因",
    "Repo URL",
    "Commit ID",
    "基线ID",
    "开始时间",
    "结束时间",
    "更新时间",
    "备注",
]

DEFAULT_WORKBOOK = "solo-faster-prompts.xlsx"
WORKBOOK_VERSION = "solo-faster-v2"
NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

DEFAULT_MAIN_TASK_DISTRIBUTION = [
    ("0-1代码生成", 20),
    ("Feature迭代", 20),
    ("代码理解", 1),
    ("代码重构", 1),
    ("工程化", 1),
]

TASK_TYPES = ["Bug修复", "0-1代码生成", "Feature迭代", "代码理解", "代码重构", "工程化", "代码测试"]
MULTI_ROUND_TASK_LIMITS = {
    "0-1代码生成": 5,
    "Feature迭代": 5,
}
DOMAINS = [
    "全栈 Web 应用",
    "Web 前端",
    "纯后端服务",
    "命令行工具",
    "科学计算",
    "3D / 交互可视化",
    "游戏开发",
    "桌面应用（含 GUI）",
    "AI/ML 应用",
    "数据分析与可视化",
    "自动化与工具脚本",
]
DIFFICULTIES = ["简单", "一般", "困难", "地狱"]
DIFFICULTY_MODES = {
    "低": ["简单", "一般"],
    "中": ["一般", "困难"],
    "高": ["困难"],
}
RANGES = ["单文件", "模块内多文件", "跨模块多文件", "跨系统多模块", "无需修改"]
STATUSES = [
    "待生成",
    "已生成",
    "已发送",
    "Trae运行中",
    "待验收",
    "验收中",
    "已完成",
    "已跳过",
    "超时待人工",
    "失败",
]
DONE_VALUES = ["已完成", "未完成"]
SATISFACTION_VALUES = ["满意", "不满意"]
ROW_TYPES = ["主提示词", "修复提示词"]
FORBIDDEN_TERMS = [
    "稳定",
    "收口",
    "收住",
    "落下来",
    "这一轮",
    "上一轮",
    "下一轮",
    "本轮",
    "当前轮",
    "当前轮次",
    "首轮",
    "次轮",
    "第几轮",
    "第一轮",
    "第二轮",
    "第三轮",
    "前一轮",
    "后一轮",
    "第2轮",
    "第3轮",
    "同上",
    "（继续）",
]
PUNCTUATION_RE = re.compile(r"[。！？.!?]$")
FULL_COMMIT_RE = re.compile(r"^[0-9a-f]{40}$")
GITHUB_REPO_MARKDOWN_RE = re.compile(r"^\[([A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+)\]\(https://github\.com/\1\)$")
BAD_REASON_TERMS = [
    "浏览器没验到",
    "浏览器没验证",
    "没验到",
    "没验证",
    "只通过代码判断",
    "我无法",
    "无法在 Codex",
    "无法在Codex",
    "不能按你的验收标准确认",
    "还不能按你的验收标准确认",
    "缺少应用内浏览器",
    "缺少浏览器",
    "页面结果佐证",
    "结果佐证",
    "暂时不能证明",
    "不能证明",
    "构建也顺利通过",
    "构建顺利通过",
    "已经把",
    "实现中已经",
    "暂时无法判断",
    "暂时无法判定",
    "无法判断",
    "无法判定",
    "实现不完整",
    "没有处理好",
    "没有核对结果",
    "缺少校验",
    "效果不好",
    "不行",
    "一般",
    "没完成任务",
    "代码有 bug",
    "代码有bug",
    "过程比较乱",
    "文件没改成功",
    "一直在报错",
    "写的代码根本不符合我的需求",
    "越改问题越多",
    "需求实现得不完整",
    "模型请求失败",
    "网络波动",
    "环境问题",
    "依赖源临时不可用",
]
REASON_DUPLICATE_LONG_FRAGMENT_LIMIT = 18
REASON_TRIGRAM_JACCARD_LIMIT = 0.30
REASON_PROMPT_JACCARD_LIMIT = 0.45


class WorkbookError(SystemExit):
    """Command-line workbook error."""


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def main_task_distribution(task_counts: str | None = None) -> list[tuple[str, int]]:
    if not task_counts:
        return list(DEFAULT_MAIN_TASK_DISTRIBUTION)
    counts_by_type: dict[str, int] = {}
    for raw_item in task_counts.split(","):
        item = raw_item.strip()
        if not item:
            continue
        if "=" not in item:
            raise WorkbookError(f"Invalid --task-counts item: {item}. Expected 任务类型=条数")
        task_type, count_text = item.split("=", 1)
        task_type = task_type.strip()
        count_text = count_text.strip()
        if task_type not in TASK_TYPES or task_type == "Bug修复":
            raise WorkbookError(f"Invalid main task type in --task-counts: {task_type}")
        if not count_text.isdigit():
            raise WorkbookError(f"Invalid count in --task-counts: {item}")
        counts_by_type[task_type] = int(count_text)
    distribution: list[tuple[str, int]] = []
    ordered_types = ["0-1代码生成", "Feature迭代", "代码理解", "代码重构", "工程化", "代码测试"]
    for task_type in ordered_types:
        count = counts_by_type.get(task_type, 0)
        if count > 0:
            distribution.append((task_type, count))
    if not distribution:
        raise WorkbookError("Invalid --task-counts: no positive main task counts")
    return distribution


def interleaved_main_task_order(distribution: list[tuple[str, int]]) -> list[str]:
    scheduled: list[tuple[float, int, int, str]] = []
    active_types = [(task_type, count) for task_type, count in distribution if count > 0]
    if not active_types:
        return []
    type_count = len(active_types)
    for order_index, (task_type, count) in enumerate(active_types):
        phase = order_index / type_count
        for occurrence_index in range(count):
            position = (occurrence_index + phase) / count
            scheduled.append((position, order_index, occurrence_index, task_type))
    scheduled.sort()
    return [task_type for _, _, _, task_type in scheduled]


def col_name(index: int) -> str:
    out = ""
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def cell_text(cell: ET.Element, shared_strings: list[str]) -> str:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        return "".join(node.text or "" for node in cell.findall(".//a:t", NS))
    value = cell.find("a:v", NS)
    if value is None or value.text is None:
        return ""
    if cell_type == "s":
        try:
            return shared_strings[int(value.text)]
        except (ValueError, IndexError):
            return ""
    return value.text


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for item in root.findall("a:si", NS):
        strings.append("".join(node.text or "" for node in item.findall(".//a:t", NS)))
    return strings


def read_workbook_raw(path: Path) -> tuple[list[str], list[dict[str, str]]]:
    if not path.exists():
        return HEADERS, []
    with zipfile.ZipFile(path) as zf:
        shared_strings = read_shared_strings(zf)
        sheet_name = "xl/worksheets/sheet1.xml"
        if sheet_name not in zf.namelist():
            return [], []
        root = ET.fromstring(zf.read(sheet_name))
    rows: list[list[str]] = []
    max_cols = 0
    for row in root.findall(".//a:sheetData/a:row", NS):
        values_by_index: dict[int, str] = {}
        for cell in row.findall("a:c", NS):
            ref = cell.attrib.get("r", "")
            match = re.match(r"([A-Z]+)", ref)
            if not match:
                continue
            index = 0
            for char in match.group(1):
                index = index * 26 + ord(char) - 64
            values_by_index[index] = cell_text(cell, shared_strings)
            max_cols = max(max_cols, index)
        values = [""] * max_cols
        for index, value in values_by_index.items():
            values[index - 1] = value
        rows.append(values)
    if not rows:
        return HEADERS, []
    headers = [item.strip() for item in rows[0] if item.strip()]
    records: list[dict[str, str]] = []
    for values in rows[1:]:
        record = {header: values[i] if i < len(values) else "" for i, header in enumerate(headers)}
        if any(str(value).strip() for value in record.values()):
            records.append(record)
    return headers, records


def ensure_v2_headers(headers: list[str], path: Path) -> None:
    if not headers:
        raise WorkbookError(f"Workbook has no header row: {path}")
    missing = [header for header in HEADERS if header not in headers]
    extra_required_old = {"子文件夹名称", "编号", "状态"}
    if missing or extra_required_old.intersection(headers):
        raise WorkbookError(
            "Existing workbook is not solo-faster v2. "
            f"Expected headers: {', '.join(HEADERS)}. "
            f"Found headers: {', '.join(headers)}. "
            "Create a v2 workbook or run a controlled migration before execution."
        )


def normalize_record(record: dict[str, str]) -> dict[str, str]:
    normalized = {header: str(record.get(header, "") or "") for header in HEADERS}
    repo_path = normalized.get("仓库路径", "").strip()
    main_number = normalized.get("主提示词编号", "").strip()
    if not normalized.get("Repo ID") and repo_path and main_number.isdigit():
        normalized["Repo ID"] = build_repo_id(Path(repo_path), int(main_number))
    legacy_process = str(record.get("过程不满意原因", "") or "").strip()
    legacy_product = str(record.get("产物不满意原因", "") or "").strip()
    if (legacy_process or legacy_product) and not normalized.get("不满意原因", "").strip():
        normalized["不满意原因"] = format_unsatisfied_reason(legacy_process, legacy_product)
    if normalized.get("执行状态") == "未完成":
        normalized["执行状态"] = "已完成"
        if not normalized.get("任务是否完成"):
            normalized["任务是否完成"] = "未完成"
    return normalized


def read_workbook(path: Path, *, strict: bool = True) -> list[dict[str, str]]:
    headers, records = read_workbook_raw(path)
    if path.exists() and strict:
        ensure_v2_headers(headers, path)
    return [normalize_record(record) for record in records]


def write_workbook(path: Path, records: list[dict[str, str]]) -> None:
    rows = [HEADERS] + [[record.get(header, "") for header in HEADERS] for record in records]
    sheet_rows = []
    for row_index, values in enumerate(rows, start=1):
        cells = []
        for col_index, value in enumerate(values, start=1):
            ref = f"{col_name(col_index)}{row_index}"
            text = escape(str(value), {'"': "&quot;"})
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{text}</t></is></c>')
        sheet_rows.append(f'<row r="{row_index}">{"".join(cells)}</row>')
    width_xml = "".join(
        f'<col min="{idx}" max="{idx}" width="{width}" customWidth="1"/>'
        for idx, width in enumerate([32, 14, 8, 12, 16, 20, 34, 92, 16, 22, 18, 12, 14, 18, 52, 32, 20, 36, 20, 20, 20, 36], start=1)
    )
    sheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<cols>{width_xml}</cols>
<sheetData>{''.join(sheet_rows)}</sheetData>
</worksheet>'''
    created = datetime.now(timezone.utc).isoformat()
    files = {
        "[Content_Types].xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>''',
        "_rels/.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>''',
        "xl/workbook.xml": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="prompts" sheetId="1" r:id="rId1"/></sheets>
</workbook>''',
        "xl/_rels/workbook.xml.rels": '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>''',
        "xl/worksheets/sheet1.xml": sheet_xml,
        "docProps/app.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>solo-faster</Application><DocSecurity>0</DocSecurity><Company>{WORKBOOK_VERSION}</Company>
</Properties>''',
        "docProps/core.xml": f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:creator>solo-faster</dc:creator><cp:lastModifiedBy>solo-faster</cp:lastModifiedBy>
<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>
</cp:coreProperties>''',
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            zf.writestr(name, content)


def default_workbook_name(repo: Path | None = None) -> str:
    if repo is None:
        return DEFAULT_WORKBOOK
    return f"{repo.name}-{DEFAULT_WORKBOOK}"


def workbook_candidates(parent: Path, workbook: str | None, repo: Path | None = None) -> list[Path]:
    if workbook:
        return [parent / workbook]
    candidates: list[Path] = []
    if repo is not None:
        candidates.append(parent / default_workbook_name(repo))
    named_workbooks = sorted(
        candidate
        for candidate in parent.glob(f"*-{DEFAULT_WORKBOOK}")
        if candidate.is_file()
    )
    if len(named_workbooks) == 1:
        candidates.append(named_workbooks[0])
    candidates.append(parent / DEFAULT_WORKBOOK)
    seen: set[Path] = set()
    ordered: list[Path] = []
    for candidate in candidates:
        if candidate in seen:
            continue
        seen.add(candidate)
        ordered.append(candidate)
    return ordered


def workbook_path(parent: Path, workbook: str | None, repo: Path | None = None) -> Path:
    candidates = workbook_candidates(parent, workbook, repo)
    for candidate in candidates:
        if candidate.exists():
            return candidate
    return candidates[0]


def run_git(repo: Path, *args: str) -> str:
    result = subprocess.run(["git", "-C", str(repo), *args], capture_output=True, text=True)
    if result.returncode != 0:
        raise WorkbookError(result.stderr.strip() or f"git {' '.join(args)} failed")
    return result.stdout


def run_git_result(repo: Path, *args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(["git", "-C", str(repo), *args], capture_output=True, text=True)


def repo_url(repo: Path) -> str:
    result = subprocess.run(["git", "-C", str(repo), "remote", "get-url", "origin"], capture_output=True, text=True)
    if result.returncode != 0:
        return ""
    return normalized_repo_url_from_remote(result.stdout.strip())


def ssh_remote_url(url: str) -> str:
    if url.startswith("https://github.com/"):
        path = url.removeprefix("https://github.com/")
        if path.endswith(".git"):
            path = path[:-4]
        return f"git@github.com:{path}.git"
    return url


def normalized_repo_url_from_remote(url: str) -> str:
    value = (url or "").strip()
    if not value:
        return ""
    markdown = re.fullmatch(r"\[([A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+)\]\(https://github\.com/([A-Za-z0-9_.-]+/[A-Za-z0-9_.-]+)(?:\.git)?\)", value)
    if markdown:
        path = markdown.group(1)
        if markdown.group(2) == path:
            return f"[{path}](https://github.com/{path})"
    if value.endswith(".git"):
        value = value[:-4]
    if value.startswith("git@github.com:"):
        path = value.removeprefix("git@github.com:")
        return f"[{path}](https://github.com/{path})"
    if value.startswith("https://github.com/"):
        path = value.removeprefix("https://github.com/")
        return f"[{path}](https://github.com/{path})"
    return value


def ensure_ssh_origin(repo: Path) -> tuple[str, str]:
    current = run_git(repo, "remote", "get-url", "origin").strip()
    target = ssh_remote_url(current)
    if target != current:
        run_git(repo, "remote", "set-url", "origin", target)
    return current, target


def is_git_repo(path: Path) -> bool:
    result = subprocess.run(["git", "-C", str(path), "rev-parse", "--is-inside-work-tree"], capture_output=True, text=True)
    return result.returncode == 0 and result.stdout.strip() == "true"


def has_code_signals(path: Path) -> bool:
    signals = [
        ".git",
        "package.json",
        "pyproject.toml",
        "requirements.txt",
        "go.mod",
        "Cargo.toml",
        "pom.xml",
        "build.gradle",
        "composer.json",
        "Gemfile",
        "Makefile",
    ]
    return any((path / signal).exists() for signal in signals)


def resolve_repo(start: Path) -> dict[str, object]:
    start = start.expanduser().resolve()
    if is_git_repo(start):
        root = Path(run_git(start, "rev-parse", "--show-toplevel").strip()).resolve()
        return {"status": "ok", "repo": str(root), "reason": "current directory is inside a git repository"}
    candidates = [child.resolve() for child in start.iterdir() if child.is_dir() and has_code_signals(child)]
    git_candidates = []
    for child in candidates:
        if is_git_repo(child):
            git_candidates.append(Path(run_git(child, "rev-parse", "--show-toplevel").strip()).resolve())
    unique = sorted(set(git_candidates or candidates))
    if len(unique) == 1:
        return {"status": "ok", "repo": str(unique[0]), "reason": "parent directory has exactly one code repository"}
    if len(unique) > 1:
        return {
            "status": "needs-user",
            "reason": "parent directory has multiple candidate repositories",
            "candidates": [str(item) for item in unique],
        }
    return {"status": "failed", "reason": "no git repository or single code child found", "candidates": []}


def detect_domain(repo: Path) -> str:
    files = {path.name.lower() for path in repo.iterdir() if path.is_file()}
    all_paths = []
    try:
        all_paths = run_git(repo, "ls-files").splitlines()
    except SystemExit:
        all_paths = []
    joined = "\n".join(all_paths).lower()
    frontend = has_frontend_signals(files, joined)
    backend = has_backend_signals(files, joined)
    if frontend and backend:
        return "全栈 Web 应用"
    if frontend:
        return "Web 前端"
    if any(name in files for name in ["pyproject.toml", "requirements.txt"]) and any(marker in joined for marker in ["sklearn", "tensorflow", "torch", "model", "notebook", ".ipynb"]):
        return "AI/ML 应用"
    if any(name in files for name in ["pyproject.toml", "requirements.txt"]) and any(marker in joined for marker in ["pandas", "plot", "chart", "visual", "dashboard"]):
        return "数据分析与可视化"
    if backend:
        return "纯后端服务"
    if any(marker in joined for marker in ["three", "webgl", "canvas", "d3", "visualization"]):
        return "3D / 交互可视化"
    if any(marker in joined for marker in ["game", "phaser", "unity", "godot"]):
        return "游戏开发"
    if any(marker in joined for marker in ["electron", "tauri", "qt", "tkinter", "gui"]):
        return "桌面应用（含 GUI）"
    if any(marker in joined for marker in ["cli", "argparse", "click", "commander"]):
        return "命令行工具"
    if any(name in files for name in ["pyproject.toml", "requirements.txt", "makefile"]) or "scripts/" in joined:
        return "自动化与工具脚本"
    return "全栈 Web 应用"


def has_frontend_signals(files: set[str], joined_paths: str) -> bool:
    frontend_files = {
        "package.json",
        "vite.config.js",
        "vite.config.ts",
        "next.config.js",
        "next.config.mjs",
        "nuxt.config.js",
        "nuxt.config.ts",
        "webpack.config.js",
        "tailwind.config.js",
        "tailwind.config.ts",
    }
    frontend_markers = [
        "src/app",
        "pages/",
        "components/",
        "views/",
        "router/",
        "routes/",
        "public/",
        "assets/",
        "vite",
        "next.config",
        "nuxt.config",
        ".tsx",
        ".jsx",
        ".vue",
        ".svelte",
    ]
    return bool(files.intersection(frontend_files) and any(marker in joined_paths for marker in frontend_markers))


def has_backend_signals(files: set[str], joined_paths: str) -> bool:
    backend_files = {
        "pyproject.toml",
        "requirements.txt",
        "go.mod",
        "pom.xml",
        "build.gradle",
        "composer.json",
        "gemfile",
        "manage.py",
    }
    backend_markers = [
        "api/",
        "server/",
        "backend/",
        "controllers/",
        "controller/",
        "services/",
        "routes/",
        "prisma/",
        "migrations/",
        "models/",
        "schema.sql",
        "app.py",
        "main.py",
        ".go",
        ".java",
        ".php",
    ]
    return bool(files.intersection(backend_files) or any(marker in joined_paths for marker in backend_markers))


def difficulty_for(mode: str, index: int) -> str:
    values = DIFFICULTY_MODES.get(mode or "中")
    if not values:
        raise WorkbookError(f"Invalid difficulty mode: {mode}. Expected one of: {', '.join(DIFFICULTY_MODES)}")
    return values[(index - 1) % len(values)]


def prompt_stub(repo: Path, task_type: str, number: int, domain: str, difficulty: str) -> str:
    repo_name = repo.name
    if task_type == "0-1代码生成":
        return f"新增贴合{repo_name}核心业务的独立能力，让用户能完成一个清晰的新流程，并能在主要入口看到结果反馈。"
    if task_type == "Feature迭代":
        return f"优化{repo_name}已有核心流程的联动体验，让用户操作后能立即看到一致的状态变化和结果反馈。"
    if task_type == "代码理解":
        return f"梳理{repo_name}的核心业务路径和关键模块职责，输出面向接手开发者的中文说明，并指出一个最值得优先验证的风险点。"
    if task_type == "代码重构":
        return f"重构{repo_name}中职责耦合较高的一段核心逻辑，保持用户可感知行为不变，并让后续扩展更容易理解。"
    if task_type == "工程化":
        return f"补强{repo_name}的工程化使用体验，让开发者能更快完成启动、检查和问题定位，并减少重复手工操作。"
    return f"完善{repo_name}的{domain}任务，难度按{difficulty}处理，并给出用户可感知的结果。"


def build_repo_id(repo: Path, main_number: int) -> str:
    return f"{repo.name}-{main_number}"


def empty_record(repo: Path, main_number: int, round_number: int, row_type: str, task_type: str, domain: str, difficulty: str) -> dict[str, str]:
    return {
        "仓库路径": str(repo),
        "主提示词编号": str(main_number),
        "轮次": str(round_number),
        "行类型": row_type,
        "执行状态": "待生成",
        "Repo ID": build_repo_id(repo, main_number),
        "Trae Session ID": "",
        "提示词": "",
        "任务类型": task_type,
        "业务领域": domain,
        "修改范围": "",
        "任务难度": difficulty,
        "任务是否完成": "",
        "过程与产物是否满意": "",
        "不满意原因": "",
        "Repo URL": repo_url(repo),
        "Commit ID": "",
        "基线ID": "",
        "开始时间": "",
        "结束时间": "",
        "更新时间": now_text(),
        "备注": "",
    }


def init_records(
    repo: Path,
    domain: str,
    difficulty_mode: str,
    *,
    with_stubs: bool,
    distribution: list[tuple[str, int]],
) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    ordered_task_types = interleaved_main_task_order(distribution)
    for number, task_type in enumerate(ordered_task_types, start=1):
        difficulty = difficulty_for(difficulty_mode, number)
        record = empty_record(repo, number, 1, "主提示词", task_type, domain, difficulty)
        if with_stubs:
            record["提示词"] = prompt_stub(repo, task_type, number, domain, difficulty)
            record["执行状态"] = "已生成"
        records.append(record)
    return records


def initialize(
    parent: Path,
    repo: Path,
    workbook: str | None,
    difficulty_mode: str,
    domain: str | None,
    with_stubs: bool,
    force: bool,
    task_counts: str | None,
) -> dict[str, object]:
    path = workbook_path(parent, workbook, repo)
    if path.exists() and not force:
        records = read_workbook(path)
        return {"workbook": str(path), "created": False, "count": len(records), "records": records}
    final_domain = domain or detect_domain(repo)
    if final_domain not in DOMAINS:
        raise WorkbookError(f"Invalid domain: {final_domain}. Expected one of: {', '.join(DOMAINS)}")
    distribution = main_task_distribution(task_counts)
    records = init_records(repo, final_domain, difficulty_mode, with_stubs=with_stubs, distribution=distribution)
    write_workbook(path, records)
    return {
        "workbook": str(path),
        "created": True,
        "count": len(records),
        "domain": final_domain,
        "records": records,
        "distribution": distribution,
    }


def main_rows(records: list[dict[str, str]]) -> list[dict[str, str]]:
    return [record for record in records if record.get("行类型") == "主提示词"]


def bugfix_count(records: list[dict[str, str]], repo: str | None = None) -> int:
    return sum(
        1
        for record in records
        if record.get("任务类型") == "Bug修复"
        and record.get("行类型") == "修复提示词"
        and (repo is None or record.get("仓库路径") == repo)
    )


def zero_to_one_main_count(records: list[dict[str, str]], repo: str | None = None) -> int:
    return sum(
        1
        for record in records
        if record.get("行类型") == "主提示词"
        and record.get("任务类型") == "0-1代码生成"
        and (repo is None or record.get("仓库路径") == repo)
    )


def rows_for_main(records: list[dict[str, str]], main_number: int) -> list[dict[str, str]]:
    return [record for record in records if int(record.get("主提示词编号") or 0) == main_number]


def latest_row_for_main(records: list[dict[str, str]], main_number: int) -> dict[str, str] | None:
    rows = rows_for_main(records, main_number)
    if not rows:
        return None
    return max(rows, key=lambda row: int(row.get("轮次") or 0))


def main_task_type(records: list[dict[str, str]], main_number: int) -> str:
    for record in rows_for_main(records, main_number):
        if record.get("行类型") == "主提示词":
            return record.get("任务类型", "")
    return ""


def multi_round_main_count(records: list[dict[str, str]], task_type: str, repo: str | None = None) -> int:
    mains: set[int] = set()
    for record in records:
        if repo is not None and record.get("仓库路径") != repo:
            continue
        if record.get("任务类型") != "Bug修复" or record.get("行类型") != "修复提示词":
            continue
        main_number = int(record.get("主提示词编号") or 0)
        if main_number > 0 and main_task_type(records, main_number) == task_type:
            mains.add(main_number)
    return len(mains)


def can_insert_fix_for_main(records: list[dict[str, str]], main_number: int, repo: str | None = None) -> bool:
    task_type = main_task_type(records, main_number)
    limit = MULTI_ROUND_TASK_LIMITS.get(task_type)
    if not limit:
        return True
    latest = latest_row_for_main(records, main_number)
    if latest and int(latest.get("轮次") or 0) >= 2:
        return True
    return multi_round_main_count(records, task_type, repo) < limit


def main_sort_key(record: dict[str, str]) -> tuple[int, int]:
    return (int(record.get("主提示词编号") or 0), int(record.get("轮次") or 0))


def send_priority_key(record: dict[str, str]) -> tuple[int, int, int]:
    row_type = record.get("行类型", "").strip()
    task_type = record.get("任务类型", "").strip()
    return (
        0 if row_type == "修复提示词" or task_type == "Bug修复" else 1,
        int(record.get("主提示词编号") or 0),
        int(record.get("轮次") or 0),
    )


def terminal_for_main(records: list[dict[str, str]], main_number: int) -> bool:
    rows = rows_for_main(records, main_number)
    if any(row.get("任务是否完成") == "已完成" for row in rows):
        return True
    latest = latest_row_for_main(records, main_number)
    if not latest:
        return False
    if int(latest.get("轮次") or 0) >= 3 and latest.get("任务是否完成") == "未完成":
        return True
    return False


def row_missing_required_fields(record: dict[str, str]) -> list[str]:
    missing: list[str] = []
    status = record.get("执行状态", "").strip()
    completion = record.get("任务是否完成", "").strip()
    session_id = record.get("Trae Session ID", "").strip()
    commit_id = record.get("Commit ID", "").strip()
    change_range = record.get("修改范围", "").strip()

    if status in {"Trae运行中", "待验收", "验收中", "超时待人工"} and not session_id:
        missing.append("Trae Session ID")
    if status == "已完成":
        if not session_id:
            missing.append("Trae Session ID")
        if not change_range:
            missing.append("修改范围")
        # Commit ID may be empty only when there were no code changes and the row note explains it.
        if not commit_id and "无需修改" != change_range:
            note = record.get("备注", "").strip()
            if "无代码改动" not in note and "没有代码改动" not in note:
                missing.append("Commit ID")
    return missing


def choose_next(records: list[dict[str, str]]) -> dict[str, object]:
    if not records:
        return {"action": "none", "reason": "empty workbook"}
    repo = records[0].get("仓库路径") or None
    fixes = bugfix_count(records, repo)
    fix_limit = zero_to_one_main_count(records, repo)
    multi_round_counts = {
        task_type: multi_round_main_count(records, task_type, repo)
        for task_type in MULTI_ROUND_TASK_LIMITS
    }
    conflict_statuses = {"已发送", "Trae运行中", "验收中", "超时待人工", "失败"}

    executable = sorted(records, key=main_sort_key)
    send_candidates: list[dict[str, str]] = []
    for record in executable:
        status = record.get("执行状态")
        main_number = int(record.get("主提示词编号") or 0)
        missing_fields = row_missing_required_fields(record)
        if missing_fields:
            return {
                "action": "confirm",
                "reason": f"row is missing required fields: {', '.join(missing_fields)}",
                "row": record,
                "missing_fields": missing_fields,
                "bugfix_count": fixes,
                "multi_round_counts": multi_round_counts,
            }
        if status in conflict_statuses:
            return {"action": "confirm", "reason": f"row is {status}", "row": record, "bugfix_count": fixes, "multi_round_counts": multi_round_counts}
        if status == "待验收":
            return {"action": "accept", "row": record, "bugfix_count": fixes, "multi_round_counts": multi_round_counts}
        if status == "已完成" and record.get("任务是否完成") == "未完成":
            latest = latest_row_for_main(records, main_number)
            if latest is record and int(record.get("轮次") or 0) < 3 and fixes < fix_limit and can_insert_fix_for_main(records, main_number, repo):
                return {"action": "insert-fix", "row": record, "bugfix_count": fixes, "multi_round_counts": multi_round_counts}
        if status in {"待生成", "已生成"} and not terminal_for_main(records, main_number):
            send_candidates.append(record)
    if send_candidates:
        return {
            "action": "send",
            "row": min(send_candidates, key=send_priority_key),
            "bugfix_count": fixes,
            "multi_round_counts": multi_round_counts,
        }
    return {"action": "none", "reason": "no executable rows", "bugfix_count": fixes, "multi_round_counts": multi_round_counts}


def parse_number_range(value: str | None) -> tuple[int, int] | None:
    if not value:
        return None
    match = re.fullmatch(r"\s*(\d+)\s*(?:-|~|至|到)\s*(\d+)\s*", value)
    if match:
        start = int(match.group(1))
        end = int(match.group(2))
        if start > end:
            start, end = end, start
        return start, end
    if re.fullmatch(r"\s*\d+\s*", value):
        number = int(value)
        return number, number
    raise WorkbookError(f"Invalid range: {value}")


def pick(parent: Path, workbook: str | None, number_range: str | None, limit: int | None, statuses: set[str] | None = None) -> dict[str, object]:
    path = workbook_path(parent, workbook)
    records = read_workbook(path)
    parsed_range = parse_number_range(number_range)
    picked: list[dict[str, str]] = []
    for record in records:
        prompt = record.get("提示词", "").strip()
        number_text = record.get("主提示词编号", "").strip()
        if not prompt or not number_text.isdigit():
            continue
        number = int(number_text)
        if parsed_range and not (parsed_range[0] <= number <= parsed_range[1]):
            continue
        if statuses and record.get("执行状态") not in statuses:
            continue
        picked.append(record)
    picked.sort(key=lambda item: (int(item.get("主提示词编号") or 0), int(item.get("轮次") or 0)))
    if limit is not None:
        picked = picked[:limit]
    return {"workbook": str(path), "count": len(picked), "items": picked}


def find_row(records: list[dict[str, str]], main_number: int, round_number: int) -> tuple[int, dict[str, str]]:
    for index, record in enumerate(records):
        if int(record.get("主提示词编号") or 0) == main_number and int(record.get("轮次") or 0) == round_number:
            return index, record
    raise WorkbookError(f"Row not found: main={main_number}, round={round_number}")


def validate_status(status: str | None) -> None:
    if status and status not in STATUSES:
        raise WorkbookError(f"Invalid status: {status}. Expected one of: {', '.join(STATUSES)}")


def validate_completion(value: str | None) -> None:
    if value and value not in DONE_VALUES:
        raise WorkbookError(f"Invalid completion value: {value}. Expected 已完成 or 未完成")


def validate_satisfaction(value: str | None) -> None:
    if value and value not in SATISFACTION_VALUES:
        raise WorkbookError(f"Invalid satisfaction value: {value}. Expected 满意 or 不满意")


def validate_output_text(label: str, value: str) -> list[str]:
    errors: list[str] = []
    if not value:
        return errors
    for term in FORBIDDEN_TERMS:
        if term in value:
            errors.append(f"{label} contains forbidden term: {term}")
    if label == "不满意原因":
        for term in BAD_REASON_TERMS:
            if term in value:
                errors.append(f"{label} contains invalid audit/process wording: {term}")
    if label in {"提示词", "不满意原因"} and not PUNCTUATION_RE.search(value.strip()):
        errors.append(f"{label} must end with punctuation")
    return errors


def valid_unsatisfied_reason(value: str) -> bool:
    stripped = value.strip()
    return bool(
        re.fullmatch(
            r"过程不满意：.+[。！？.!?]\n产物不满意：.+[。！？.!?]",
            stripped,
            flags=re.S,
        )
    )


def normalize_similarity_text(value: str) -> str:
    return re.sub(r"[\s，。！？、：；,.!?;:“”\"'（）()《》【】\[\]\-]", "", value or "")


def trigrams(value: str) -> set[str]:
    text = normalize_similarity_text(value)
    return {text[index : index + 3] for index in range(max(0, len(text) - 2))}


def trigram_jaccard(left: str, right: str) -> float:
    left_grams = trigrams(left)
    right_grams = trigrams(right)
    if not left_grams or not right_grams:
        return 0.0
    return len(left_grams & right_grams) / len(left_grams | right_grams)


def longest_common_text_length(left: str, right: str) -> int:
    left_text = normalize_similarity_text(left)
    right_text = normalize_similarity_text(right)
    blocks = SequenceMatcher(None, left_text, right_text, autojunk=False).get_matching_blocks()
    return max((block.size for block in blocks), default=0)


def split_unsatisfied_reason(value: str) -> tuple[str, str]:
    match = re.fullmatch(r"过程不满意：(.+)\n产物不满意：(.+)", value.strip(), flags=re.S)
    if not match:
        return "", ""
    return match.group(1).strip(), match.group(2).strip()


def validate_candidate_unsatisfied_reason(
    candidate: str,
    candidate_prompt: str,
    existing_records: list[dict[str, str]],
    *,
    current_main: int | None = None,
    current_round: int | None = None,
) -> list[str]:
    errors = validate_output_text("不满意原因", candidate)
    if not valid_unsatisfied_reason(candidate):
        errors.append("不满意原因 must merge as 过程不满意/产物不满意 lines")
    if candidate_prompt and trigram_jaccard(candidate, candidate_prompt) >= REASON_PROMPT_JACCARD_LIMIT:
        errors.append("不满意原因 is too similar to 提示词; rewrite the reason instead of restating the prompt")
    candidate_process, candidate_product = split_unsatisfied_reason(candidate)
    for index, record in enumerate(existing_records, start=2):
        try:
            same_row = (
                current_main is not None
                and current_round is not None
                and int(record.get("主提示词编号") or 0) == current_main
                and int(record.get("轮次") or 0) == current_round
            )
        except ValueError:
            same_row = False
        if same_row:
            continue
        existing_reason = record.get("不满意原因", "").strip()
        if not existing_reason:
            continue
        long_reason = longest_common_text_length(candidate, existing_reason)
        if long_reason >= REASON_DUPLICATE_LONG_FRAGMENT_LIMIT:
            errors.append(f"不满意原因 shares a long duplicate fragment with row {index} ({long_reason} chars)")
        score = trigram_jaccard(candidate, existing_reason)
        if score >= REASON_TRIGRAM_JACCARD_LIMIT:
            errors.append(f"不满意原因 trigram Jaccard is too high with row {index} ({score:.1%})")
        existing_process, existing_product = split_unsatisfied_reason(existing_reason)
        if candidate_process and existing_process:
            long_process = longest_common_text_length(candidate_process, existing_process)
            if long_process >= REASON_DUPLICATE_LONG_FRAGMENT_LIMIT:
                errors.append(f"过程不满意段 shares a long duplicate fragment with row {index} ({long_process} chars)")
        if candidate_product and existing_product:
            long_product = longest_common_text_length(candidate_product, existing_product)
            if long_product >= REASON_DUPLICATE_LONG_FRAGMENT_LIMIT:
                errors.append(f"产物不满意段 shares a long duplicate fragment with row {index} ({long_product} chars)")
    return errors


def validate_record(record: dict[str, str]) -> list[str]:
    errors: list[str] = []
    if not record.get("Repo ID", "").strip():
        errors.append("Repo ID is required")
    if record.get("任务类型") not in TASK_TYPES:
        errors.append(f"Invalid task type: {record.get('任务类型')}")
    if record.get("业务领域") and record.get("业务领域") not in DOMAINS:
        errors.append(f"Invalid domain: {record.get('业务领域')}")
    if record.get("修改范围") and record.get("修改范围") not in RANGES:
        errors.append(f"Invalid change range: {record.get('修改范围')}")
    if record.get("任务难度") and record.get("任务难度") not in DIFFICULTIES:
        errors.append(f"Invalid difficulty: {record.get('任务难度')}")
    repo_url_value = record.get("Repo URL", "").strip()
    if repo_url_value and not GITHUB_REPO_MARKDOWN_RE.fullmatch(repo_url_value):
        normalized_url = normalized_repo_url_from_remote(repo_url_value)
        if not GITHUB_REPO_MARKDOWN_RE.fullmatch(normalized_url):
            errors.append("Repo URL must use markdown GitHub format: [owner/repo](https://github.com/owner/repo)")
    if record.get("执行状态") not in STATUSES:
        errors.append(f"Invalid status: {record.get('执行状态')}")
    if record.get("任务是否完成") and record.get("任务是否完成") not in DONE_VALUES:
        errors.append(f"Invalid completion: {record.get('任务是否完成')}")
    if record.get("过程与产物是否满意") and record.get("过程与产物是否满意") not in SATISFACTION_VALUES:
        errors.append(f"Invalid satisfaction: {record.get('过程与产物是否满意')}")
    if record.get("行类型") not in ROW_TYPES:
        errors.append(f"Invalid row type: {record.get('行类型')}")
    if record.get("行类型") == "修复提示词" and record.get("任务类型") != "Bug修复":
        errors.append("Fix rows must use task type Bug修复")
    for label in ["提示词", "不满意原因"]:
        errors.extend(validate_output_text(label, record.get(label, "")))
    if record.get("任务是否完成") == "已完成" and record.get("过程与产物是否满意") and record.get("过程与产物是否满意") != "满意":
        errors.append("Completed rows must be 满意")
    if record.get("过程与产物是否满意") == "满意" and record.get("不满意原因", "").strip():
        errors.append("Satisfied rows must not include 不满意原因")
    if record.get("执行状态") == "已跳过":
        executed_fields = [
            field
            for field in ["Trae Session ID", "Commit ID", "任务是否完成", "过程与产物是否满意"]
            if record.get(field, "").strip()
        ]
        if executed_fields:
            errors.append(f"Skipped rows must not include execution evidence: {', '.join(executed_fields)}")
    if record.get("执行状态") == "已完成" and not record.get("Trae Session ID", "").strip():
        errors.append("Completed rows must include Trae Session ID")
    if record.get("执行状态") == "已完成" and record.get("修改范围") != "无需修改" and not record.get("Commit ID", "").strip():
        errors.append("Completed rows with code changes must include Commit ID")
    if record.get("Commit ID", "").strip() and record.get("修改范围") == "无需修改":
        errors.append("Rows with Commit ID must not use 修改范围=无需修改")
    if record.get("Commit ID", "").strip() and not FULL_COMMIT_RE.fullmatch(record.get("Commit ID", "").strip()):
        errors.append("Commit ID must be a 40-character git SHA")
    if record.get("任务是否完成") in DONE_VALUES or record.get("执行状态") == "已完成":
        if not record.get("修改范围", "").strip():
            errors.append("Terminal rows must include 修改范围")
    if record.get("任务是否完成") == "未完成":
        if record.get("过程与产物是否满意") != "不满意":
            errors.append("Unfinished rows must be 不满意")
        if not record.get("不满意原因", "").strip():
            errors.append("Unfinished rows must include 不满意原因")
        elif not valid_unsatisfied_reason(record.get("不满意原因", "")):
            errors.append("Unfinished rows must merge 不满意原因 as 过程不满意/产物不满意 lines")
    return errors


def validate_records(records: list[dict[str, str]]) -> dict[str, object]:
    errors: list[str] = []
    main = main_rows(records)
    if not main:
        errors.append("Expected at least one main prompt row")
    for index, record in enumerate(records, start=2):
        for error in validate_record(record):
            errors.append(f"Row {index}: {error}")
    errors.extend(validate_reason_similarity(records))
    by_main: dict[int, list[int]] = {}
    for record in records:
        try:
            by_main.setdefault(int(record.get("主提示词编号") or 0), []).append(int(record.get("轮次") or 0))
        except ValueError:
            errors.append(f"Invalid main/round number in record: {record}")
    for main_number, rounds in by_main.items():
        if main_number <= 0:
            errors.append(f"Main prompt number out of range: {main_number}")
        if max(rounds) > 3:
            errors.append(f"Main prompt {main_number} exceeds 3 rounds")
        if len(rounds) != len(set(rounds)):
            errors.append(f"Main prompt {main_number} has duplicate rounds")
        repo_ids = {record.get("Repo ID", "").strip() for record in records if int(record.get("主提示词编号") or 0) == main_number}
        if len(repo_ids) != 1:
            errors.append(f"Main prompt {main_number} must use exactly one Repo ID, found: {sorted(repo_ids)}")
    fixes = bugfix_count(records)
    fix_limit = zero_to_one_main_count(records)
    if fixes > fix_limit:
        errors.append(f"Bug修复 rows exceed 0-1代码生成 main row count: {fixes} > {fix_limit}")
    for task_type, limit in MULTI_ROUND_TASK_LIMITS.items():
        current = multi_round_main_count(records, task_type)
        if current > limit:
            errors.append(f"{task_type} multi-round main count exceeds limit: {current} > {limit}")
    return {"ok": not errors, "errors": errors, "row_count": len(records), "bugfix_count": fixes}


def validate_reason_similarity(records: list[dict[str, str]]) -> list[str]:
    errors: list[str] = []
    reason_rows: list[tuple[int, dict[str, str], str, str, str]] = []
    for index, record in enumerate(records, start=2):
        reason = record.get("不满意原因", "").strip()
        if not reason:
            continue
        process_reason, product_reason = split_unsatisfied_reason(reason)
        reason_rows.append((index, record, reason, process_reason, product_reason))
        prompt = record.get("提示词", "").strip()
        if prompt and trigram_jaccard(reason, prompt) >= REASON_PROMPT_JACCARD_LIMIT:
            errors.append(f"Row {index}: 不满意原因 is too similar to 提示词")
    for left_index, (row_index, _, reason, process_reason, product_reason) in enumerate(reason_rows):
        for other_index, _, other_reason, other_process, other_product in reason_rows[left_index + 1 :]:
            long_reason = longest_common_text_length(reason, other_reason)
            if long_reason >= REASON_DUPLICATE_LONG_FRAGMENT_LIMIT:
                errors.append(f"Rows {row_index}/{other_index}: 不满意原因 share a long duplicate fragment ({long_reason} chars)")
            score = trigram_jaccard(reason, other_reason)
            if score >= REASON_TRIGRAM_JACCARD_LIMIT:
                errors.append(f"Rows {row_index}/{other_index}: 不满意原因 trigram Jaccard too high ({score:.1%})")
            if process_reason and other_process:
                long_process = longest_common_text_length(process_reason, other_process)
                if long_process >= REASON_DUPLICATE_LONG_FRAGMENT_LIMIT:
                    errors.append(f"Rows {row_index}/{other_index}: 过程不满意段 share a long duplicate fragment ({long_process} chars)")
            if product_reason and other_product:
                long_product = longest_common_text_length(product_reason, other_product)
                if long_product >= REASON_DUPLICATE_LONG_FRAGMENT_LIMIT:
                    errors.append(f"Rows {row_index}/{other_index}: 产物不满意段 share a long duplicate fragment ({long_product} chars)")
    return errors


def update_row(
    parent: Path,
    workbook: str | None,
    main_number: int,
    round_number: int,
    fields: dict[str, str],
    *,
    strict_validate: bool = True,
) -> dict[str, object]:
    path = workbook_path(parent, workbook)
    records = read_workbook(path)
    index, row = find_row(records, main_number, round_number)
    for key, value in fields.items():
        if key not in HEADERS:
            raise WorkbookError(f"Unknown field: {key}")
        if key == "Repo URL":
            value = normalized_repo_url_from_remote(value)
        row[key] = value
    normalize_reason_fields(row)
    row["更新时间"] = now_text()
    if row.get("执行状态") in {"已发送", "Trae运行中"} and not row.get("开始时间"):
        row["开始时间"] = row["更新时间"]
    if row.get("执行状态") in {"已完成", "失败"} and not row.get("结束时间"):
        row["结束时间"] = row["更新时间"]
    validate_status(row.get("执行状态"))
    validate_completion(row.get("任务是否完成"))
    validate_satisfaction(row.get("过程与产物是否满意"))
    if strict_validate:
        errors = validate_record(row)
        if row.get("任务是否完成") == "未完成" and row.get("不满意原因", "").strip():
            errors.extend(
                validate_candidate_unsatisfied_reason(
                    row["不满意原因"],
                    row.get("提示词", ""),
                    records,
                    current_main=main_number,
                    current_round=round_number,
                )
            )
        if errors:
            raise WorkbookError("; ".join(errors))
    records[index] = row
    write_workbook(path, records)
    return {"workbook": str(path), "row": row}


def normalize_reason_fields(row: dict[str, str]) -> None:
    if row.get("任务是否完成") == "未完成":
        row["不满意原因"] = row.get("不满意原因", "").strip()
    if row.get("任务是否完成") == "已完成":
        row["不满意原因"] = ""


def insert_fix(parent: Path, workbook: str | None, main_number: int, prompt: str, note: str = "") -> dict[str, object]:
    path = workbook_path(parent, workbook)
    records = read_workbook(path)
    fixes = bugfix_count(records)
    fix_limit = zero_to_one_main_count(records)
    if fixes >= fix_limit:
        raise WorkbookError(f"Bug修复 row limit reached: current workbook 0-1代码生成 main row count is {fix_limit}")
    repo = records[0].get("仓库路径") if records else None
    if not can_insert_fix_for_main(records, main_number, repo):
        task_type = main_task_type(records, main_number)
        limit = MULTI_ROUND_TASK_LIMITS.get(task_type)
        if limit:
            raise WorkbookError(f"{task_type} multi-round main limit reached: limit is {limit}")
    rows = rows_for_main(records, main_number)
    if not rows:
        raise WorkbookError(f"Main prompt not found: {main_number}")
    max_round = max(int(row.get("轮次") or 0) for row in rows)
    if max_round >= 3:
        raise WorkbookError(f"Main prompt {main_number} already has 3 rounds")
    base = rows[0]
    new_round = max_round + 1
    record = empty_record(Path(base["仓库路径"]), main_number, new_round, "修复提示词", "Bug修复", base["业务领域"], base["任务难度"])
    record["提示词"] = prompt
    record["执行状态"] = "已生成"
    record["Repo URL"] = base.get("Repo URL", "")
    record["基线ID"] = base.get("基线ID", "")
    record["备注"] = note
    errors = validate_record(record)
    if errors:
        raise WorkbookError("; ".join(errors))
    insert_at = max(index for index, row in enumerate(records) if int(row.get("主提示词编号") or 0) == main_number) + 1
    records.insert(insert_at, record)
    write_workbook(path, records)
    return {"workbook": str(path), "row": record, "bugfix_count": fixes + 1}


def apply_outcome(
    parent: Path,
    workbook: str | None,
    main_number: int,
    round_number: int,
    completed: str,
    process_reason: str,
    product_reason: str,
    reason: str,
    change_range: str,
    note: str,
) -> dict[str, object]:
    if completed not in DONE_VALUES:
        raise WorkbookError("Outcome completion must be 已完成 or 未完成")
    path = workbook_path(parent, workbook)
    records = read_workbook(path)
    _, existing_row = find_row(records, main_number, round_number)
    fields: dict[str, str] = {
        "任务是否完成": completed,
        "执行状态": "已完成",
        "过程与产物是否满意": "满意" if completed == "已完成" else "不满意",
        "修改范围": change_range,
        "备注": note,
    }
    if completed == "未完成":
        fields["不满意原因"] = reason.strip() if valid_unsatisfied_reason(reason or "") else format_unsatisfied_reason(process_reason, product_reason)
        reason_errors = validate_candidate_unsatisfied_reason(
            fields["不满意原因"],
            existing_row.get("提示词", ""),
            records,
            current_main=main_number,
            current_round=round_number,
        )
        if reason_errors:
            raise WorkbookError("; ".join(reason_errors))
    else:
        fields["不满意原因"] = ""
    return update_row(parent, workbook, main_number, round_number, fields, strict_validate=False)


def format_unsatisfied_reason(process_reason: str, product_reason: str) -> str:
    process = process_reason.strip()
    product = product_reason.strip()
    return f"过程不满意：{process}\n产物不满意：{product}"


def changed_paths(repo: Path) -> list[str]:
    status = run_git(repo, "status", "--short")
    paths: list[str] = []
    workbook_names = {DEFAULT_WORKBOOK, default_workbook_name(repo)}
    for line in status.splitlines():
        if not line.strip():
            continue
        path_text = line[3:]
        if " -> " in path_text:
            path_text = path_text.split(" -> ", 1)[1]
        if Path(path_text).name in workbook_names:
            continue
        if Path(path_text).name in {".DS_Store"}:
            continue
        paths.append(path_text)
    return paths


def module_name(path: str) -> str:
    parts = Path(path).parts
    if len(parts) <= 1:
        return Path(path).stem
    if parts[0] in {"packages", "apps", "services"} and len(parts) > 1:
        return "/".join(parts[:2])
    if parts[0] in {"src", "app", "lib"}:
        if len(parts) == 2:
            return parts[0]
        return "/".join(parts[:2])
    return parts[0]


def detect_change_range(repo: Path, domain: str | None = None) -> str:
    paths = changed_paths(repo)
    return classify_change_range(paths, domain)


def classify_change_range(paths: list[str], domain: str | None = None) -> str:
    if not paths:
        return "无需修改"
    if len(paths) == 1:
        return "单文件"
    if spans_multiple_systems(paths, domain):
        return "跨系统多模块"
    modules = {module_name(path) for path in paths}
    if len(modules) == 1:
        return "模块内多文件"
    return "跨模块多文件"


def spans_multiple_systems(paths: list[str], domain: str | None = None) -> bool:
    systems = {system_name(path, domain) for path in paths}
    systems.discard("")
    if domain == "Web 前端" and systems and systems.issubset({"frontend", "frontend-config"}):
        return False
    return len(systems) >= 2


def system_name(path: str, domain: str | None = None) -> str:
    parts = Path(path).parts
    top = parts[0] if parts else path
    suffix = Path(path).suffix.lower()
    name = Path(path).name.lower()
    frontend_config_names = {
        "package.json",
        "package-lock.json",
        "pnpm-lock.yaml",
        "yarn.lock",
        "vite.config.js",
        "vite.config.ts",
        "webpack.config.js",
        "next.config.js",
        "next.config.mjs",
        "nuxt.config.js",
        "nuxt.config.ts",
        "tailwind.config.js",
        "tailwind.config.ts",
        "postcss.config.js",
        "tsconfig.json",
        "jsconfig.json",
    }
    frontend_tops = {
        "src",
        "app",
        "components",
        "pages",
        "views",
        "routes",
        "router",
        "store",
        "stores",
        "styles",
        "assets",
        "static",
        "public",
        "hooks",
        "composables",
        "utils",
        "lib",
    }
    if domain == "Web 前端":
        if top in frontend_tops or suffix in {".css", ".scss", ".sass", ".less", ".js", ".mjs", ".ts", ".tsx", ".jsx", ".vue", ".svelte", ".html"}:
            return "frontend"
        if name in frontend_config_names:
            return "frontend-config"
    if top in {"admin", "backend", "server", "api"}:
        return "backend-admin" if top == "admin" else "backend"
    if top in {"frontend", "client", "web", "pages", "public"}:
        return "frontend"
    if top in {"assets", "static"} and suffix in {".css", ".js", ".ts", ".tsx", ".jsx", ".vue", ".html"}:
        return "frontend"
    if top in {"includes", "config", "database", "migrations"} or name in {"api.php", "functions.php"}:
        return "backend"
    if suffix in {".css", ".js", ".ts", ".tsx", ".jsx", ".vue"}:
        return "frontend"
    if suffix in {".php", ".py", ".rb", ".go", ".java", ".kt", ".cs", ".sql"}:
        return "backend"
    return top


def lint_texts(parent: Path, workbook: str | None) -> dict[str, object]:
    path = workbook_path(parent, workbook)
    records = read_workbook(path)
    errors: list[str] = []
    for index, record in enumerate(records, start=2):
        for label in ["提示词", "不满意原因"]:
            for error in validate_output_text(label, record.get(label, "")):
                errors.append(f"Row {index}: {error}")
    return {"ok": not errors, "errors": errors}


def audit_git_integrity(records: list[dict[str, str]], repo: Path | None = None) -> list[str]:
    errors: list[str] = []
    for index, record in enumerate(records, start=2):
        commit_id = record.get("Commit ID", "").strip()
        if not commit_id:
            continue
        if not FULL_COMMIT_RE.fullmatch(commit_id):
            errors.append(f"Row {index}: Commit ID is not a full 40-character SHA: {commit_id}")
            continue
        row_repo = repo
        if row_repo is None and record.get("仓库路径", "").strip():
            row_repo = Path(record["仓库路径"]).expanduser().resolve()
        if row_repo is None or not row_repo.exists():
            errors.append(f"Row {index}: repository path is unavailable for Commit ID audit")
            continue
        object_type = run_git_result(row_repo, "cat-file", "-t", commit_id)
        if object_type.returncode != 0 or object_type.stdout.strip() != "commit":
            errors.append(f"Row {index}: Commit ID does not resolve to a local commit: {commit_id}")
            continue
        session_id = record.get("Trae Session ID", "").strip()
        message = run_git(row_repo, "log", "-1", "--format=%B", commit_id).strip()
        if session_id and message != session_id:
            errors.append(f"Row {index}: commit message must equal Trae Session ID")
    return errors


def audit_records(parent: Path, workbook: str | None, repo: Path | None = None) -> dict[str, object]:
    path = workbook_path(parent, workbook, repo)
    records = read_workbook(path)
    result = validate_records(records)
    git_errors = audit_git_integrity(records, repo)
    result["git_errors"] = git_errors
    result["ok"] = bool(result["ok"]) and not git_errors
    result["workbook"] = str(path)
    return result


def finalize_round(
    parent: Path,
    workbook: str | None,
    repo: Path,
    main_number: int,
    round_number: int,
    *,
    no_push: bool = False,
) -> dict[str, object]:
    path = workbook_path(parent, workbook, repo)
    records = read_workbook(path)
    _, row = find_row(records, main_number, round_number)
    session_id = row.get("Trae Session ID", "").strip()
    if not session_id:
        raise WorkbookError("Trae Session ID is required before git finalize")
    if row.get("执行状态") != "已完成" or row.get("任务是否完成") not in DONE_VALUES:
        raise WorkbookError("Run outcome before finalize-round; row must be terminal and include completion")

    domain = row.get("业务领域", "").strip() or None
    paths = changed_paths(repo)
    change_range = classify_change_range(paths, domain)
    repo_url_value = normalized_repo_url_from_remote(run_git(repo, "remote", "get-url", "origin").strip()) if run_git_result(repo, "remote", "get-url", "origin").returncode == 0 else ""
    if not paths:
        note = "无代码改动，需人工确认是否继续"
        result = update_row(
            parent,
            workbook,
            main_number,
            round_number,
            {
                "修改范围": "无需修改",
                "Commit ID": "",
                "Repo URL": repo_url_value,
                "备注": note,
            },
        )
        return {"workbook": str(path), "committed": False, "pushed": False, "change_range": "无需修改", "paths": [], "row": result["row"]}

    if repo_url_value:
        current_remote, push_remote = ensure_ssh_origin(repo)
        repo_url_value = normalized_repo_url_from_remote(push_remote)
    else:
        current_remote, push_remote = "", ""

    run_git(repo, "add", "--", *paths)
    run_git(repo, "commit", "-m", session_id)
    commit_id = run_git(repo, "rev-parse", "HEAD").strip()
    if not FULL_COMMIT_RE.fullmatch(commit_id):
        raise WorkbookError(f"git returned invalid Commit ID: {commit_id}")
    message = run_git(repo, "log", "-1", "--format=%B", commit_id).strip()
    if message != session_id:
        raise WorkbookError("Created commit message does not equal Trae Session ID")
    pushed = False
    if not no_push:
        run_git(repo, "push", "origin", "HEAD")
        pushed = True
    note_parts = [f"修改文件：{', '.join(paths)}"]
    if current_remote and current_remote != push_remote:
        note_parts.append("origin 已切换为 GitHub SSH")
    result = update_row(
        parent,
        workbook,
        main_number,
        round_number,
        {
            "修改范围": change_range,
            "Commit ID": commit_id,
            "Repo URL": repo_url_value,
            "备注": "；".join(note_parts),
        },
    )
    reread = read_workbook(path)
    _, persisted = find_row(reread, main_number, round_number)
    if persisted.get("Commit ID", "").strip() != commit_id:
        raise WorkbookError("Commit ID was not persisted to workbook")
    return {
        "workbook": str(path),
        "committed": True,
        "pushed": pushed,
        "commit_id": commit_id,
        "change_range": change_range,
        "paths": paths,
        "repo_url": repo_url_value,
        "row": result["row"],
    }


def command_scan(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    repo_result = resolve_repo(Path(args.repo).expanduser().resolve() if args.repo else parent)
    if repo_result["status"] != "ok":
        return repo_result
    repo = Path(str(repo_result["repo"]))
    result = initialize(
        parent,
        repo,
        args.workbook,
        args.difficulty_mode,
        args.domain,
        args.with_stubs,
        args.force,
        args.task_counts,
    )
    result["repo"] = str(repo)
    result["repo_resolution"] = repo_result
    return result


def command_status(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    path = workbook_path(parent, args.workbook)
    records = read_workbook(path)
    statuses: dict[str, int] = {}
    for record in records:
        statuses[record.get("执行状态", "")] = statuses.get(record.get("执行状态", ""), 0) + 1
    result = audit_records(parent, args.workbook)
    result.update({"workbook": str(path), "statuses": statuses, "next": choose_next(records)})
    return result


def command_update(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    fields: dict[str, str] = {}
    for assignment in args.set or []:
        if "=" not in assignment:
            raise WorkbookError(f"--set expects FIELD=VALUE, got: {assignment}")
        key, value = assignment.split("=", 1)
        fields[key] = value
    for key in HEADERS:
        value = getattr(args, field_to_arg(key), None)
        if value is not None:
            fields[key] = value
    if args.process_reason is not None or args.product_reason is not None:
        process_reason = args.process_reason or ""
        product_reason = args.product_reason or ""
        fields["不满意原因"] = format_unsatisfied_reason(process_reason, product_reason)
    return update_row(parent, args.workbook, args.main, args.round, fields, strict_validate=not args.no_strict_validate)


def field_to_arg(field: str) -> str:
    mapping = {
        "提示词": "prompt",
        "执行状态": "status",
        "Repo ID": "repo_id",
        "任务是否完成": "completed",
        "过程与产物是否满意": "satisfaction",
        "不满意原因": "reason",
        "Repo URL": "repo_url",
        "Commit ID": "commit_id",
        "基线ID": "baseline_id",
        "修改范围": "change_range",
        "备注": "note",
        "开始时间": "start_time",
        "结束时间": "end_time",
        "Trae Session ID": "trae_session_id",
    }
    return mapping.get(field, field)


def command_outcome(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    change_range = args.change_range
    repo = Path(args.repo).expanduser().resolve() if args.repo else None
    if not change_range and args.repo:
        path = workbook_path(parent, args.workbook, repo)
        records = read_workbook(path)
        _, row = find_row(records, args.main, args.round)
        change_range = detect_change_range(repo, row.get("业务领域", "").strip() or None)
    elif change_range and args.repo and not args.force_change_range:
        path = workbook_path(parent, args.workbook, repo)
        records = read_workbook(path)
        _, row = find_row(records, args.main, args.round)
        detected = detect_change_range(repo, row.get("业务领域", "").strip() or None)
        if detected != change_range:
            raise WorkbookError(f"Manual change range {change_range} does not match detected range {detected}; use --force-change-range to override")
    if not change_range:
        change_range = "无需修改"
    return apply_outcome(
        parent,
        args.workbook,
        args.main,
        args.round,
        args.completed,
        args.process_reason or "",
        args.product_reason or "",
        args.reason or "",
        change_range,
        args.note or "",
    )


def command_insert_fix(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    return insert_fix(parent, args.workbook, args.main, args.prompt, args.note or "")


def command_pick(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    statuses = set(args.status_filter.split(",")) if args.status_filter else None
    return pick(parent, args.workbook, args.range, args.limit, statuses)


def command_next(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    records = read_workbook(workbook_path(parent, args.workbook))
    validation = validate_records(records)
    git_errors = audit_git_integrity(records)
    if not validation["ok"] or git_errors:
        return {
            "action": "confirm",
            "reason": "workbook audit failed",
            "errors": validation["errors"],
            "git_errors": git_errors,
        }
    return choose_next(records)


def command_lint(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    result = audit_records(parent, args.workbook)
    text_result = lint_texts(parent, args.workbook)
    result["text_ok"] = text_result["ok"]
    result["text_errors"] = text_result["errors"]
    result["ok"] = result["ok"] and text_result["ok"]
    return result


def command_range(args: argparse.Namespace) -> dict[str, object]:
    repo = Path(args.repo or ".").expanduser().resolve()
    return {"repo": str(repo), "change_range": detect_change_range(repo, args.domain), "paths": changed_paths(repo)}


def command_audit(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    repo = Path(args.repo).expanduser().resolve() if args.repo else None
    return audit_records(parent, args.workbook, repo)


def command_reason_check(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    path = workbook_path(parent, args.workbook)
    records = read_workbook(path)
    _, row = find_row(records, args.main, args.round)
    candidate = args.reason or format_unsatisfied_reason(args.process_reason or "", args.product_reason or "")
    errors = validate_candidate_unsatisfied_reason(
        candidate,
        row.get("提示词", ""),
        records,
        current_main=args.main,
        current_round=args.round,
    )
    return {
        "ok": not errors,
        "errors": errors,
        "regenerate_required": bool(errors),
        "workbook": str(path),
        "main": args.main,
        "round": args.round,
    }


def command_finalize_round(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    repo = Path(args.repo).expanduser().resolve()
    return finalize_round(parent, args.workbook, repo, args.main, args.round, no_push=args.no_push)


def command_locate_project(args: argparse.Namespace) -> dict[str, object]:
    project = Path(args.project or ".").expanduser().resolve()
    repo_result = resolve_repo(project)
    if repo_result["status"] != "ok":
        return {"found": False, **repo_result}
    repo = Path(str(repo_result["repo"]))
    parent_candidates = [repo.parent, *repo.parents]
    for parent in parent_candidates:
        for path in workbook_candidates(parent, args.workbook, repo):
            if not path.exists():
                continue
            try:
                records = read_workbook(path)
            except SystemExit as exc:
                return {"found": False, "workbook": str(path), "error": str(exc)}
            if any(Path(record.get("仓库路径", "")).resolve() == repo for record in records if record.get("仓库路径")):
                return {
                    "found": True,
                    "project": str(project),
                    "repo": str(repo),
                    "parent": str(parent),
                    "workbook": str(path),
                    "next": choose_next(records),
                }
    return {"found": False, "project": str(project), "repo": str(repo)}


def command_migrate_headers(args: argparse.Namespace) -> dict[str, object]:
    parent = Path(args.parent or ".").expanduser().resolve()
    path = workbook_path(parent, args.workbook)
    records = read_workbook(path, strict=False)
    normalized = [normalize_record(record) for record in records]
    write_workbook(path, normalized)
    return {"workbook": str(path), "row_count": len(normalized), "headers": HEADERS}


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    sub = parser.add_subparsers(dest="command", required=True)

    scan = sub.add_parser("scan", help="Resolve one repository and create/read a v2 workbook")
    scan.add_argument("--parent", default=".")
    scan.add_argument("--repo")
    scan.add_argument("--workbook")
    scan.add_argument("--difficulty-mode", default="中", choices=sorted(DIFFICULTY_MODES))
    scan.add_argument("--domain", choices=DOMAINS)
    scan.add_argument("--task-counts", help="Main prompt distribution, e.g. 0-1代码生成=20,Feature迭代=20,代码理解=1,代码重构=1,工程化=1")
    scan.add_argument("--with-stubs", action="store_true", help="Fill generated rows with placeholder prompts for dry-run/self-test")
    scan.add_argument("--force", action="store_true", help="Overwrite an existing v2 workbook")

    status = sub.add_parser("status")
    status.add_argument("--parent", default=".")
    status.add_argument("--workbook")

    update = sub.add_parser("update")
    update.add_argument("--parent", default=".")
    update.add_argument("--workbook")
    update.add_argument("--main", type=int, required=True)
    update.add_argument("--round", type=int, required=True)
    update.add_argument("--prompt")
    update.add_argument("--status")
    update.add_argument("--completed")
    update.add_argument("--satisfaction")
    update.add_argument("--process-reason")
    update.add_argument("--product-reason")
    update.add_argument("--reason")
    update.add_argument("--baseline-id")
    update.add_argument("--change-range")
    update.add_argument("--note")
    update.add_argument("--start-time")
    update.add_argument("--end-time")
    update.add_argument("--trae-session-id")
    update.add_argument("--set", action="append", help="Set an arbitrary workbook field as FIELD=VALUE")
    update.add_argument("--no-strict-validate", action="store_true")

    outcome = sub.add_parser("outcome")
    outcome.add_argument("--parent", default=".")
    outcome.add_argument("--workbook")
    outcome.add_argument("--main", type=int, required=True)
    outcome.add_argument("--round", type=int, required=True)
    outcome.add_argument("--completed", required=True, choices=DONE_VALUES)
    outcome.add_argument("--process-reason", default="")
    outcome.add_argument("--product-reason", default="")
    outcome.add_argument("--reason", default="")
    outcome.add_argument("--change-range", choices=RANGES)
    outcome.add_argument("--force-change-range", action="store_true")
    outcome.add_argument("--repo")
    outcome.add_argument("--note", default="")

    insert_fix_cmd = sub.add_parser("insert-fix")
    insert_fix_cmd.add_argument("--parent", default=".")
    insert_fix_cmd.add_argument("--workbook")
    insert_fix_cmd.add_argument("--main", type=int, required=True)
    insert_fix_cmd.add_argument("--prompt", required=True)
    insert_fix_cmd.add_argument("--note", default="")

    pick_cmd = sub.add_parser("pick")
    pick_cmd.add_argument("--parent", default=".")
    pick_cmd.add_argument("--workbook")
    pick_cmd.add_argument("--range")
    pick_cmd.add_argument("--limit", type=int)
    pick_cmd.add_argument("--status-filter")

    next_cmd = sub.add_parser("next")
    next_cmd.add_argument("--parent", default=".")
    next_cmd.add_argument("--workbook")

    lint_cmd = sub.add_parser("lint")
    lint_cmd.add_argument("--parent", default=".")
    lint_cmd.add_argument("--workbook")

    range_cmd = sub.add_parser("change-range")
    range_cmd.add_argument("--repo", default=".")
    range_cmd.add_argument("--domain", choices=DOMAINS)

    audit_cmd = sub.add_parser("audit")
    audit_cmd.add_argument("--parent", default=".")
    audit_cmd.add_argument("--workbook")
    audit_cmd.add_argument("--repo")

    reason_check = sub.add_parser("reason-check")
    reason_check.add_argument("--parent", default=".")
    reason_check.add_argument("--workbook")
    reason_check.add_argument("--main", type=int, required=True)
    reason_check.add_argument("--round", type=int, required=True)
    reason_check.add_argument("--process-reason", default="")
    reason_check.add_argument("--product-reason", default="")
    reason_check.add_argument("--reason", default="")

    finalize_cmd = sub.add_parser("finalize-round")
    finalize_cmd.add_argument("--parent", default=".")
    finalize_cmd.add_argument("--workbook")
    finalize_cmd.add_argument("--repo", required=True)
    finalize_cmd.add_argument("--main", type=int, required=True)
    finalize_cmd.add_argument("--round", type=int, required=True)
    finalize_cmd.add_argument("--no-push", action="store_true", help="Create the commit and write Commit ID without pushing; intended for self-check only")

    locate = sub.add_parser("locate-project")
    locate.add_argument("--project", default=".")
    locate.add_argument("--workbook")

    migrate = sub.add_parser("migrate-headers")
    migrate.add_argument("--parent", default=".")
    migrate.add_argument("--workbook")

    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    commands = {
        "scan": command_scan,
        "status": command_status,
        "update": command_update,
        "outcome": command_outcome,
        "insert-fix": command_insert_fix,
        "pick": command_pick,
        "next": command_next,
        "lint": command_lint,
        "change-range": command_range,
        "audit": command_audit,
        "reason-check": command_reason_check,
        "finalize-round": command_finalize_round,
        "locate-project": command_locate_project,
        "migrate-headers": command_migrate_headers,
    }
    try:
        result = commands[args.command](args)
    except WorkbookError as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False, indent=2), file=sys.stderr)
        raise SystemExit(2) from exc
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
