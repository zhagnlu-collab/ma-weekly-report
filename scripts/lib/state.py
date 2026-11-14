#!/usr/bin/env python3
"""运行状态与断点续跑。每个阶段成功后写入输入/输出指纹，重跑自动跳过相同阶段。"""
from __future__ import annotations

import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path


def digest_bytes(b: bytes | str) -> str:
    """兼容 str/bytes；编排器用字符串拼接阶段输入指纹。"""
    if isinstance(b, str):
        b = b.encode('utf-8')
    return hashlib.sha256(b).hexdigest()


def digest_path(p: Path) -> str:
    h = hashlib.sha256()
    with p.open('rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def digest_tree(root: Path, suffixes: tuple[str, ...] = ('.json', '.csv', '.xls', '.pdf', '.html')) -> str:
    h = hashlib.sha256()
    if not root.exists():
        return h.hexdigest()
    for p in sorted(root.rglob('*')):
        if p.is_file() and p.suffix.lower() in suffixes:
            h.update(p.relative_to(root).as_posix().encode())
            h.update(digest_path(p).encode())
    return h.hexdigest()


class RunState:
    def __init__(self, path: Path, as_of: str, extra: dict | None = None):
        self.path = path
        self.path.parent.mkdir(parents=True, exist_ok=True)
        if path.exists():
            self.data = json.loads(path.read_text(encoding='utf-8'))
            if self.data.get('as_of') != as_of:
                raise RuntimeError(f'运行目录已有其它 as-of 的状态: {self.data.get("as_of")}')
        else:
            self.data = {'as_of': as_of, 'created_at': _now(), 'stages': {},
                         'extra': extra or {}, 'status': 'RUNNING'}
            self.save()

    def save(self):
        self.data['updated_at'] = _now()
        self.path.write_text(json.dumps(self.data, ensure_ascii=False, indent=2), encoding='utf-8')

    def input_key(self, *parts: str) -> str:
        return digest_bytes('|'.join(parts).encode())

    def should_skip(self, stage: str, key: str) -> bool:
        s = self.data['stages'].get(stage)
        return bool(s and s.get('status') == 'OK' and s.get('input_key') == key)

    def start(self, stage: str, key: str):
        self.data['stages'][stage] = {'status': 'RUNNING', 'input_key': key, 'started_at': _now()}
        self.save()

    def ok(self, stage: str, key: str, **detail):
        # 保留字段必须最后写入：调用方 detail 里可能带 `status`
        # （例如 VALIDATE 返回门禁状态），若被展开覆盖会把 OK 改掉，
        # 导致 should_skip 永远不成立、每期重复执行该阶段。
        self.data['stages'][stage] = {**detail, 'status': 'OK', 'input_key': key,
                                      'finished_at': _now()}
        self.save()

    def fail(self, stage: str, key: str, error: str):
        self.data['stages'][stage] = {'status': 'FAIL', 'input_key': key,
                                      'finished_at': _now(), 'error': error}
        self.save()

    def stage(self, name: str) -> dict:
        return self.data['stages'].get(name, {})

    def set_status(self, status: str, **extra):
        self.data['status'] = status
        self.data['extra'].update(extra)
        self.save()


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')
