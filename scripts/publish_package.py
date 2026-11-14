#!/usr/bin/env python3
"""原子发布成果包 + 产出 artifact-manifest.json。

约定：
  - 全部文件先在与最终目录同一文件系统的临时 staging 目录拼齐；
  - 完整性复核（角色、哈希、无临时文件）通过后，用 os.replace 一次性改名；
  - 最终目录已存在时自动改用 _r2/_r3 修订目录，绝不覆盖已完成版本；
  - 任何一步失败都只留下 staging，不产生"半份成果包"。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
from datetime import datetime, timezone
from pathlib import Path

NATIVE = ('.xlsx', '.xlsm', '.docx', '.pptx')
REQUIRED = {f'xlsx_{i}' for i in range(14) if i != 5} | {'weekly_docx', 'dynamic_docx', 'market_ppt'}


def sha(p: Path) -> str:
    h = hashlib.sha256()
    with p.open('rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


def role(name: str) -> str | None:
    low = name.lower()
    if name.startswith('并购重组市场周报') and low.endswith('.pptx'):
        return 'market_ppt'
    if name.startswith('A股市场并购重组动态') and low.endswith('.docx'):
        return 'dynamic_docx'
    m = re.match(r'^(\d+)、', name)
    if m:
        if low.endswith(('.xlsx', '.xlsm')):
            return 'xlsx_' + m.group(1)
        if low.endswith('.docx') and m.group(1) == '5':
            return 'weekly_docx'
    return None


def choose_final(out_root: Path, dirname: str) -> tuple[Path, int]:
    candidate = out_root / dirname
    if not candidate.exists():
        return candidate, 1
    for n in range(2, 100):
        candidate = out_root / f'{dirname}_r{n}'
        if not candidate.exists():
            return candidate, n
    raise RuntimeError(f'修订目录已排到 _r99，请先清理 {out_root}')


def preflight(building: Path) -> dict:
    files = sorted(p for p in building.iterdir() if p.is_file()) if building.is_dir() else []
    native = [p for p in files if p.suffix.lower() in NATIVE and not p.name.startswith('~$')]
    pdfs = [p for p in files if p.suffix.lower() == '.pdf']
    roles = {r for p in native if (r := role(p.name))}
    temps = [p.name for p in files if p.name.startswith('~$') or p.suffix.lower() in ('.tmp', '.part')]
    missing = sorted(REQUIRED - roles)
    return {'ok': not missing and len(pdfs) == 3 and not temps,
            'native': len(native), 'pdfs': len(pdfs), 'missing_roles': missing,
            'temps': temps, 'files': [p.name for p in files]}


def atomic_publish(building: Path, out_root: Path, dirname: str,
                   extras: list[Path], meta: dict) -> dict:
    pf = preflight(building)
    if not pf['ok']:
        raise RuntimeError(f'发布前置检查失败: {pf}')
    out_root.mkdir(parents=True, exist_ok=True)
    final, revision = choose_final(out_root, dirname)
    stage = out_root / f'.publishing-{os.getpid()}-{datetime.now().strftime("%H%M%S")}'
    if stage.exists():
        shutil.rmtree(stage)
    try:
        shutil.copytree(building, stage)
        for p in extras or []:
            if p and p.is_file():
                shutil.copy2(p, stage / p.name)
        entries = []
        for p in sorted(stage.iterdir()):
            if p.is_file():
                entries.append({'name': p.name, 'bytes': p.stat().st_size, 'sha256': sha(p)})
        manifest = {
            'schema': 'ma-weekly-artifact-manifest/v1',
            'published_at': datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds'),
            'revision': revision,
            'final_dir': final.name,
            'status': meta.get('status', 'UNKNOWN'),
            'as_of': meta.get('as_of'),
            'period': meta.get('period'),
            'run_dir': meta.get('run_dir'),
            'publish_mode': 'atomic-os-replace',
            'counts': {'native': pf['native'], 'pdf': pf['pdfs'], 'total_files': len(entries)},
            'exceptions': meta.get('exceptions', []),
            'files': entries,
        }
        (stage / 'artifact-manifest.json').write_text(
            json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
        manifest['files'].append({
            'name': 'artifact-manifest.json',
            'bytes': (stage / 'artifact-manifest.json').stat().st_size,
            'sha256': sha(stage / 'artifact-manifest.json'),
        })
        # 二次复核：staging 内每个登记文件仍可读取
        for e in manifest['files'][:-1]:
            if sha(stage / e['name']) != e['sha256']:
                raise RuntimeError(f'staging 哈希漂移: {e["name"]}')
        os.replace(stage, final)
    except Exception:
        if stage.exists():
            shutil.rmtree(stage, ignore_errors=True)
        raise
    # 发布后复核：最终目录内容与 manifest 一致（manifest 自身除外）
    for e in manifest['files']:
        if e['name'] == 'artifact-manifest.json':
            continue
        q = final / e['name']
        if not q.is_file() or sha(q) != e['sha256']:
            raise RuntimeError(f'发布后复核失败: {e["name"]}')
    return {'manifest': manifest, 'final': str(final), 'revision': revision,
            'preflight': pf, 'verified': True}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--building', required=True, type=Path)
    ap.add_argument('--out-root', required=True, type=Path)
    ap.add_argument('--dirname', required=True)
    ap.add_argument('--status', default='UNKNOWN')
    ap.add_argument('--as-of', default='')
    ap.add_argument('--period', type=Path)
    ap.add_argument('--run-dir', type=Path)
    ap.add_argument('--exceptions', type=Path)
    ap.add_argument('--note', action='append', type=Path, default=[])
    ap.add_argument('--report', required=True, type=Path)
    args = ap.parse_args()
    period = json.loads(args.period.read_text(encoding='utf-8')) if args.period else {}
    exc = json.loads(args.exceptions.read_text(encoding='utf-8')) if args.exceptions else []
    if isinstance(exc, dict):
        exc = exc.get('soft_failures', []) + exc.get('hard_failures', [])
    meta = {'status': args.status, 'as_of': args.as_of,
            'period': {k: period.get(k) for k in ('start', 'end', 'year', 'issue_no')},
            'run_dir': str(args.run_dir) if args.run_dir else '',
            'exceptions': exc}
    result = atomic_publish(args.building, args.out_root, args.dirname, args.note, meta)
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'final': result['final'], 'revision': result['revision'],
                      'status': args.status, 'files': result['manifest']['counts']},
                     ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
