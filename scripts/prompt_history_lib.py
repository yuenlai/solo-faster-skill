#!/usr/bin/env python3

from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parent.parent
HISTORY_DIR = SKILL_DIR / "data"
HISTORY_FILE = HISTORY_DIR / "prompt_history.jsonl"

PREFIX_PATTERNS = [
    re.compile(r"^([A-Za-z]+-\d+)"),
    re.compile(r"^([A-Za-z]+\d+)"),
    re.compile(r"^(\d{4,})"),
]


def ensure_history_dir() -> None:
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)


def detect_prefix(*names: str) -> str:
    for raw_name in names:
        if not raw_name:
            continue
        name = Path(raw_name).name.strip()
        for segment in re.split(r"[_\-\s]+", name):
            segment = segment.strip()
            if not segment:
                continue
            for pattern in PREFIX_PATTERNS:
                match = pattern.match(segment)
                if match:
                    return match.group(1).lower()
        for pattern in PREFIX_PATTERNS:
            match = pattern.match(name)
            if match:
                return match.group(1).lower()
    return ""


def load_history() -> list[dict]:
    if not HISTORY_FILE.exists():
        return []

    records: list[dict] = []
    with HISTORY_FILE.open("r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            records.append(json.loads(line))
    return records


def append_history(record: dict) -> None:
    ensure_history_dir()
    payload = dict(record)
    payload.setdefault("created_at", datetime.now().isoformat(timespec="seconds"))
    with HISTORY_FILE.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(payload, ensure_ascii=False) + "\n")

