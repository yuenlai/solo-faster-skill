#!/usr/bin/env python3
"""Global runtime lock for solo-faster real execution phase."""

from __future__ import annotations

import json
import os
import socket
import time
import uuid
from datetime import datetime
from pathlib import Path


LOCK_ROOT = Path.home() / ".codex" / "solo-faster-runtime-lock"
OWNER_FILE = LOCK_ROOT / "owner.json"
QUEUE_DIR = LOCK_ROOT / "queue"

DEFAULT_STALE_SECONDS = 60 * 60
DEFAULT_POLL_SECONDS = 2.0
DEFAULT_QUEUE_STALE_SECONDS = 8 * 60 * 60


class RuntimeLockError(RuntimeError):
    """Runtime lock failure."""


def now_text() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def ensure_lock_dirs() -> None:
    LOCK_ROOT.mkdir(parents=True, exist_ok=True)
    QUEUE_DIR.mkdir(parents=True, exist_ok=True)


def read_json(path: Path) -> dict[str, object]:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def write_json(path: Path, payload: dict[str, object]) -> None:
    ensure_lock_dirs()
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


def owner_payload(repo_path: Path, stage: str) -> dict[str, object]:
    return {
        "repo_path": str(repo_path),
        "stage": stage,
        "hostname": socket.gethostname(),
        "pid": os.getpid(),
        "acquired_at": now_text(),
        "last_seen_at": now_text(),
    }


def ticket_payload(repo_path: Path, stage: str, ticket_id: str) -> dict[str, object]:
    return {
        "ticket_id": ticket_id,
        "repo_path": str(repo_path),
        "stage": stage,
        "hostname": socket.gethostname(),
        "pid": os.getpid(),
        "created_at": now_text(),
        "created_at_epoch": time.time(),
    }


def create_ticket(repo_path: Path, stage: str) -> Path:
    ensure_lock_dirs()
    ticket_id = f"{time.time_ns():020d}-{os.getpid()}-{uuid.uuid4().hex}"
    ticket = QUEUE_DIR / f"{ticket_id}.json"
    write_json(ticket, ticket_payload(repo_path, stage, ticket_id))
    return ticket


def queue_entries() -> list[Path]:
    ensure_lock_dirs()
    return sorted(QUEUE_DIR.glob("*.json"), key=lambda path: path.name)


def cleanup_stale_queue_entries(queue_stale_seconds: int = DEFAULT_QUEUE_STALE_SECONDS) -> None:
    cutoff = time.time() - queue_stale_seconds
    for entry in queue_entries():
        payload = read_json(entry)
        created = float(payload.get("created_at_epoch") or 0)
        if created and created < cutoff:
            try:
                entry.unlink()
            except OSError:
                pass


def owner_is_stale(owner: dict[str, object], stale_seconds: int = DEFAULT_STALE_SECONDS) -> bool:
    timestamp = str(owner.get("last_seen_at") or owner.get("acquired_at") or "")
    if not timestamp:
        return True
    try:
        seen = datetime.strptime(timestamp, "%Y-%m-%d %H:%M:%S").timestamp()
    except ValueError:
        return True
    return seen < time.time() - stale_seconds


def current_owner(stale_seconds: int = DEFAULT_STALE_SECONDS) -> dict[str, object]:
    owner = read_json(OWNER_FILE)
    if owner and owner_is_stale(owner, stale_seconds):
        try:
            OWNER_FILE.unlink()
        except OSError:
            pass
        return {}
    return owner


def touch_owner(repo_path: Path, stage: str) -> dict[str, object]:
    owner = current_owner()
    if owner.get("repo_path") != str(repo_path):
        raise RuntimeLockError("cannot touch runtime lock owned by another repo")
    owner["stage"] = stage
    owner["last_seen_at"] = now_text()
    write_json(OWNER_FILE, owner)
    return owner


def acquire_runtime_lock(
    repo_path: Path,
    stage: str,
    *,
    wait: bool = True,
    poll_seconds: float = DEFAULT_POLL_SECONDS,
    stale_seconds: int = DEFAULT_STALE_SECONDS,
) -> dict[str, object]:
    ensure_lock_dirs()
    cleanup_stale_queue_entries()
    repo_path = repo_path.expanduser().resolve()

    owner = current_owner(stale_seconds)
    if owner.get("repo_path") == str(repo_path):
        return {"mode": "reentrant", "owner": touch_owner(repo_path, stage), "queue_position": 0}

    ticket = create_ticket(repo_path, stage)
    try:
        while True:
            owner = current_owner(stale_seconds)
            if owner.get("repo_path") == str(repo_path):
                return {"mode": "reentrant", "owner": touch_owner(repo_path, stage), "queue_position": 0}

            entries = queue_entries()
            position = next((index + 1 for index, entry in enumerate(entries) if entry == ticket), None)
            if position is None:
                ticket = create_ticket(repo_path, stage)
                continue

            if not owner and position == 1:
                payload = owner_payload(repo_path, stage)
                try:
                    fd = os.open(str(OWNER_FILE), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                    with os.fdopen(fd, "w", encoding="utf-8") as handle:
                        json.dump(payload, handle, ensure_ascii=False, indent=2)
                    try:
                        ticket.unlink()
                    except OSError:
                        pass
                    return {"mode": "acquired", "owner": payload, "queue_position": position}
                except FileExistsError:
                    pass

            if not wait:
                raise RuntimeLockError(
                    "runtime lock busy: "
                    f"owner={json.dumps(owner, ensure_ascii=False)} position={position}"
                )
            time.sleep(poll_seconds)
    finally:
        if ticket.exists():
            try:
                ticket.unlink()
            except OSError:
                pass


def release_runtime_lock(repo_path: Path, *, force: bool = False) -> dict[str, object]:
    ensure_lock_dirs()
    repo_path = repo_path.expanduser().resolve()
    owner = current_owner()
    if not owner:
        return {"released": False, "reason": "no-owner"}
    if not force and owner.get("repo_path") != str(repo_path):
        raise RuntimeLockError(
            "cannot release runtime lock owned by another repo: "
            f"{json.dumps(owner, ensure_ascii=False)}"
        )
    try:
        OWNER_FILE.unlink()
    except OSError as exc:  # noqa: BLE001
        raise RuntimeLockError(f"failed to release runtime lock: {exc}") from exc
    return {"released": True, "owner": owner}


def runtime_lock_status(stale_seconds: int = DEFAULT_STALE_SECONDS) -> dict[str, object]:
    ensure_lock_dirs()
    cleanup_stale_queue_entries()
    owner = current_owner(stale_seconds)
    queue = [read_json(entry) for entry in queue_entries()]
    return {"owner": owner, "queue": queue, "lock_root": str(LOCK_ROOT)}
