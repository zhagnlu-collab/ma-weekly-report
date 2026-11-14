#!/usr/bin/env python3
"""ma-weekly-report v4 无人值守执行器：日期进、成品出。

唯一必需输入：

    python scripts/auto_run.py --as-of 2026-11-07

程序自行完成：期次/区间/模板推导 → 官方源采集 → 公告与事件台账 → 指标快照
→ 16 个原生 Office 文件保真生成 → PDF 与逐页视觉核验 → 分级质量门禁
→ 《编制与待核说明》 → 原子发布。过程中不提问、不等人、不发消息。

阶段全部写入 run-state.json，输入指纹相同即跳过；失败只重跑失败阶段。
硬门禁失败或程序异常一律 BLOCKED，绝不发布半成品；软门禁不足出件但
强制标注 COMPLETE_WITH_EXCEPTIONS + 待复核稿说明，不冒称最终发布版。

PDF 默认用固定 LibreOffice 渲染并如实记录版式例外；若已有 Microsoft Office
「最适合打印」导出的 3 份 PDF，可用 --pdf-source-dir 直接装配，不中断流程。
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import traceback
import zipfile
from datetime import date, datetime, timezone
from pathlib import Path

SCRIPTS = Path(__file__).resolve().parent
SKILL = SCRIPTS.parent
sys.path.insert(0, str(SCRIPTS))
from lib import period as P  # noqa: E402
from lib.state import RunState, digest_bytes, digest_path, digest_tree  # noqa: E402

PY = sys.executable
STAGES = ['PLAN', 'ACQUIRE', 'EXTRACT', 'METRICS', 'BUILD',
          'RENDER', 'VALIDATE', 'NOTES', 'PUBLISH']
NATIVE = ('.xlsx', '.xlsm', '.docx', '.pptx')
ALL_SUFFIXES = NATIVE + ('.pdf', '.json', '.md')


class StageFailure(RuntimeError):
    pass


def now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')


def jdump(p: Path, data):
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


def jload(p: Path, default=None):
    if not p.is_file():
        return default
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except (ValueError, OSError):
        return default


# ---------------------------------------------------------------- 环境自举
def detect_renderer(root: Path, cfg: dict) -> tuple[Path | None, Path | None]:
    soffice_candidates = [
        os.environ.get('MA_WEEKLY_SOFFICE'), cfg.get('soffice'),
        str(Path.home() / '.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice'),
        shutil.which('soffice'),
        '/Applications/LibreOffice.app/Contents/MacOS/soffice',
    ]
    font_candidates = [
        os.environ.get('MA_WEEKLY_FONTCONFIG'), cfg.get('fontconfig'),
        str(root / 'e2e_test_ma_weekly_20261108/fonts-office.conf'),
    ]
    soffice = next((Path(c) for c in soffice_candidates if c and Path(c).is_file()), None)
    fontconfig = next((Path(c) for c in font_candidates if c and Path(c).is_file()), None)
    return soffice, fontconfig


def has_period_dirs(p: Path) -> bool:
    """判断目录是否像工作区根：含至少一个「并购周报基础数据-YYYY年第N期-…」目录。"""
    try:
        return any(x.is_dir() and x.name.startswith('并购周报基础数据-')
                   and '年第' in x.name and '至' in x.name
                   for x in p.iterdir())
    except OSError:
        return False


def resolve_root(args) -> Path:
    """工作区根自动探测。

    skill 装在工程目录里时，SKILL.parent 就是工作区；但装到用户级 skills/
    或放在测试子目录时不是。此时沿 cwd 与 skill 路径向上找含期次目录的根，
    保证「只给一个日期」在任意安装位置都能跑，无需额外参数。
    """
    if args.root:
        return Path(args.root).resolve()
    default = SKILL.parent.resolve()
    if has_period_dirs(default):
        return default
    seen: set[Path] = set()
    for seed in (Path.cwd(), SKILL):
        cur = seed.resolve()
        for cand in (cur, *list(cur.parents)[:6]):
            if cand in seen:
                continue
            seen.add(cand)
            if has_period_dirs(cand):
                return cand
    return default


def preflight_template(files: list[Path]) -> list[str]:
    """模板可读性预检，返回不可读文件说明清单。

    2026-11-08 五轮回放 Round 5 与离线端到端回归同时暴露：第43期模板被
    OneDrive 回收为云端占位后，PLAN 阶段直到 `shutil.copy2` 才抛裸
    `ETIMEDOUT`，既慢又无法指导处置。提前校验可给出可操作的
    `TEMPLATE_UNREADABLE` 判定，并继续守住「绝不发布半成品」。
    """
    bad: list[str] = []
    for p in files:
        try:
            with zipfile.ZipFile(p) as z:
                names = z.namelist()
            required = ('xl/workbook.xml' if p.suffix.lower() in ('.xlsx', '.xlsm')
                        else '[Content_Types].xml')
            if required not in names:
                bad.append(f'{p.name}: OOXML 缺少 {required}')
        except Exception as exc:
            bad.append(f'{p.name}: {type(exc).__name__}: {exc}')
    return bad


def prepare_template_snapshot(root: Path, prev: Path, end: str, out: Path) -> tuple[Path, list[dict]]:
    """Create a read-only-by-convention working snapshot of the previous package.

    BUILD mutates workbook/Word/PPT packages.  Keeping the source package under
    a run-local snapshot prevents a retry from reading files written by a
    previous attempt and makes the source hashes reproducible even when the
    OneDrive location changes during a run.
    """
    del root, end  # retained in the signature for call-site compatibility
    if not prev.is_dir():
        raise StageFailure(f'上期目录不存在: {prev}')
    out.mkdir(parents=True, exist_ok=True)
    source_files = sorted(
        p for p in prev.iterdir()
        if p.is_file() and p.suffix.lower() in NATIVE and not p.name.startswith('~$'))
    if not source_files:
        raise StageFailure(f'上期目录无原生 Office 文件: {prev}')
    unreadable = preflight_template(source_files)
    if unreadable:
        raise StageFailure(
            'TEMPLATE_UNREADABLE: 上期模板含不可读 Office 文件'
            '（常见为 OneDrive 云端占位）。请先在 Finder 中对模板目录执行'
            '「始终保留在此设备上」完成物化，或改用本地可读副本后重跑；'
            '不得把空模板/占位文件当作模板继续出件。\n'
            + '\n'.join('  - ' + x for x in unreadable[:20])
            + (f'\n  ... 另有 {len(unreadable) - 20} 个' if len(unreadable) > 20 else ''))
    expected = {p.name for p in source_files}
    for stale in out.iterdir():
        if stale.is_file() and stale.name not in expected:
            stale.unlink()
    records: list[dict] = []
    for src in source_files:
        dst = out / src.name
        if not dst.is_file() or digest_path(src) != digest_path(dst):
            shutil.copy2(src, dst)
        # First production operation: reset previous-week green on the copy.
        cleanup = {}
        if dst.suffix.lower() in ('.xlsx', '.xlsm'):
            from lib.clear_green import clear_green_fills
            cleanup = clear_green_fills(dst)
        records.append({
            'green_cleanup': cleanup,
            'source_sha256': digest_path(src),
            'source': str(src),
            'snapshot': str(dst),
            'bytes': dst.stat().st_size,
            'sha256': digest_path(dst),
        })
    return out, records


def cli(argv: list, log: Path, allow: tuple[int, ...] = (0,), timeout: int = 900) -> subprocess.CompletedProcess:
    cp = subprocess.run([PY, *[str(a) for a in argv]], cwd=str(SCRIPTS),
                        capture_output=True, text=True, timeout=timeout)
    log.parent.mkdir(parents=True, exist_ok=True)
    log.write_text(f'$ {PY} ' + ' '.join(str(a) for a in argv) +
                   f'\n\n[exit {cp.returncode}]\n[stdout]\n{cp.stdout}\n[stderr]\n{cp.stderr}\n',
                   encoding='utf-8')
    if cp.returncode not in allow:
        tail = (cp.stderr or cp.stdout or '').strip()[-800:]
        raise StageFailure(f'{Path(str(argv[0])).name} 退出码 {cp.returncode}: {tail}')
    return cp


def stage(state: RunState, name: str, key: str, fn, retries: int = 0,
          log_dir: Path | None = None):
    """幂等阶段执行器：命中指纹即跳过，失败自动重试，异常写入状态。"""
    if state.should_skip(name, key):
        state.data['stages'][name]['skipped_at'] = now()
        state.save()
        return {'skipped': True, 'detail': state.stage(name)}
    state.start(name, key)
    last = None
    for attempt in range(1, retries + 2):
        try:
            detail = fn(attempt) or {}
            state.ok(name, key, attempt=attempt, **{k: v for k, v in detail.items()
                                                     if isinstance(v, (str, int, float, bool, list, dict))})
            return detail
        except Exception as ex:  # noqa: BLE001 - 无人值守必须收口所有异常
            last = ex
            if log_dir:
                (log_dir / f'{name}.attempt{attempt}.error.txt').write_text(
                    ''.join(traceback.format_exception(type(ex), ex, ex.__traceback__)),
                    encoding='utf-8')
    state.fail(name, key, f'{type(last).__name__}: {last}')
    raise StageFailure(f'阶段 {name} 失败（重试 {retries} 次）：{last}')


# ---------------------------------------------------------------- 主流程
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--as-of', required=True, help='统计截止日 YYYY-MM-DD（唯一必需输入）')
    ap.add_argument('--root', help='工作区根目录；默认取 skill 所在目录的父目录')
    ap.add_argument('--work-root', type=Path, help='运行目录根；默认 <root>/.weekly-runs')
    ap.add_argument('--out-root', type=Path, help='成品发布根目录；默认 <root>')
    ap.add_argument('--prev', help='诊断用：显式指定上期目录；正常生产不需要')
    ap.add_argument('--config', type=Path, help='可选 JSON 配置（soffice/fontconfig/max_pdf 等）')
    ap.add_argument('--soffice', type=Path, help='可选：固定渲染器路径')
    ap.add_argument('--fontconfig', type=Path, help='可选：固定字体配置')
    ap.add_argument('--pdf-source-dir', type=Path,
                    help='可选：Office 原生导出的 3 份 PDF 目录；提供后不再调用 LibreOffice')
    ap.add_argument('--curation-decisions', type=Path,
                    help='可选：人工复核决定 JSON；默认读取运行目录/curation-decisions.json')
    ap.add_argument('--max-pdf', type=int, help='本期最多下载的公告原件数，默认 400（按公告角色优先）')
    ap.add_argument('--offline', action='store_true',
                    help='离线结构回放：复用运行目录已固化的 sources/，不联网')
    ap.add_argument('--fresh', action='store_true', help='强制新建运行目录（忽略已有状态）')
    ap.add_argument('--retries', type=int, default=1, help='单阶段失败重试次数，默认 1')
    ap.add_argument('--publish-policy', choices=['with-exceptions', 'complete-only'],
                    default=None,
                    help='软门禁未通过时是否仍出件（默认 with-exceptions，可由配置覆盖）')
    ap.add_argument('--timeout', type=int, default=900, help='单条子命令超时秒数')
    ap.add_argument('--plan-only', action='store_true', help='只推导计划，不执行后续阶段')
    args = ap.parse_args()

    root = resolve_root(args)
    cfg = jload(args.config, {}) if args.config else {}
    cfg = cfg.get('unattended', cfg)
    soffice = args.soffice or (Path(cfg['soffice']) if cfg.get('soffice') else None)
    fontconfig = args.fontconfig or (Path(cfg['fontconfig']) if cfg.get('fontconfig') else None)
    auto_soffice, auto_font = detect_renderer(root, cfg)
    soffice = soffice or auto_soffice
    fontconfig = fontconfig or auto_font
    max_pdf = args.max_pdf if args.max_pdf is not None else int(cfg.get('max_pdf', 400))
    work_root = (args.work_root or Path(cfg.get('work_root') or root / '.weekly-runs')).resolve()
    out_root = (args.out_root or Path(cfg.get('out_root') or root)).resolve()
    pdf_dir = (args.pdf_source_dir or (Path(cfg['pdf_source_dir'])
               if cfg.get('pdf_source_dir') else None))
    pdf_dir = pdf_dir.resolve() if pdf_dir else None
    publish_policy = args.publish_policy or cfg.get('publish_policy', 'with-exceptions')
    if publish_policy not in ('with-exceptions', 'complete-only'):
        raise SystemExit(f'publish_policy 非法: {publish_policy}')
    tz = cfg.get('timezone', 'Asia/Shanghai')
    cutoff = cfg.get('cutoff', '23:59:59')
    if args.retries == 1 and 'retries' in cfg:
        args.retries = int(cfg['retries'])
    if args.fresh:
        run_dir = work_root / f"{args.as_of}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
    else:
        run_dir = work_root / args.as_of
    run_dir.mkdir(parents=True, exist_ok=True)
    logs = run_dir / 'stage-logs'
    logs.mkdir(parents=True, exist_ok=True)

    state = RunState(run_dir / 'run-state.json', args.as_of,
                     extra={'root': str(root), 'offline': args.offline,
                            'publish_policy': publish_policy,
                            'pdf_source': str(pdf_dir) if pdf_dir else 'libreoffice'})
    summary: dict = {'as_of': args.as_of, 'version': '4.1.0', 'run_dir': str(run_dir),
                     'stages': {}, 'started_at': now()}
    print(json.dumps({'event': 'START', 'as_of': args.as_of, 'run_dir': str(run_dir)},
                     ensure_ascii=False), flush=True)

    def finish(status: str, **extra) -> int:
        summary.update({'finished_at': now(), 'status': status, **extra})
        summary['stages'] = state.data['stages']
        jdump(run_dir / 'auto-run-summary.json', summary)
        state.set_status(status, **{k: v for k, v in extra.items()
                                    if isinstance(v, (str, int, float, bool, list, dict))})
        print(json.dumps({'event': 'DONE', 'status': status, **extra}, ensure_ascii=False), flush=True)
        return 0 if status in ('COMPLETE', 'COMPLETE_WITH_EXCEPTIONS', 'PLAN_ONLY') else 2

    try:
        # ---------------- PLAN ----------------
        plan_key = digest_bytes('|'.join([args.as_of, args.prev or '', str(root),
                                          digest_path(SCRIPTS / 'lib/period.py'),
                                          digest_path(Path(__file__))]))
        def do_plan(_):
            per = P.resolve(root, args.as_of, args.prev)
            template_snapshot, template_sources = prepare_template_snapshot(
                root, Path(per.prev_dir), per.end, run_dir / 'template-snapshot')
            plan = {'as_of': per.as_of, 'period': P.as_dict(per),
                    'template_snapshot': str(template_snapshot),
                    'template_sources': template_sources,
                    'renderer': str(soffice) if soffice else None,
                    'fontconfig': str(fontconfig) if fontconfig else None,
                    'pdf_source': str(pdf_dir) if pdf_dir else 'libreoffice',
                    'offline': args.offline, 'max_pdf': max_pdf,
                    'out_root': str(out_root), 'work_root': str(work_root),
                    'timezone': tz, 'cutoff': cutoff,
                    'computer_use': '仅作网页改版/登录态受限及 Office 原生 PDF 导出的可选通道；正常生产不调用',
                    'notes': per.notes}
            jdump(run_dir / 'period.json', P.as_dict(per))
            jdump(run_dir / 'plan.json', plan)
            prev = Path(per.prev_dir)
            fingerprints = {p.name: digest_path(p) for p in sorted(prev.iterdir())
                            if p.is_file() and p.suffix.lower() in NATIVE}
            jdump(run_dir / 'prev-fingerprints.json', fingerprints)
            return {'out_dirname': per.out_dirname, 'prev': per.prev_dir,
                    'span': per.span_token, 'prev_native': len(fingerprints)}
        plan_detail = stage(state, 'PLAN', plan_key, do_plan, log_dir=logs)
        period_data = jload(run_dir / 'period.json')
        plan_data = jload(run_dir / 'plan.json')
        per = P.Period(**period_data)
        template_dir = Path(plan_data.get('template_snapshot') or per.prev_dir)
        summary['plan'] = plan_data
        print(json.dumps({'event': 'PLAN', 'out': per.out_dirname,
                          'span': per.span_token, 'prev': Path(per.prev_dir).name,
                          'notes': per.notes}, ensure_ascii=False), flush=True)
        if args.plan_only:
            # 不能用 kwarg `plan`：finish() 会覆盖 summary['plan'] 里的完整计划
            return finish('PLAN_ONLY', plan_summary=plan_detail)

        sources_dir = run_dir / 'sources'
        sources_dir.mkdir(parents=True, exist_ok=True)
        manifest_f = run_dir / 'source-manifest.json'
        ledger_f = run_dir / 'event-ledger.json'
        pending_f = run_dir / 'pending-review.json'
        metrics_f = run_dir / 'metrics.json'
        content_dir = run_dir / 'content'
        package_dir = run_dir / 'building'

        # ---------------- ACQUIRE ----------------
        # 幂等键只描述「请求」，不能包含 sources/ 的当前内容：
        # 首次采集会往 sources/ 写文件，若把 cached/empty 放进键，
        # 重跑时键必然变化，ACQUIRE 每次都会被重新执行。
        acq_key = digest_bytes('|'.join([
            per.start, per.end, str(max_pdf), 'offline' if args.offline else 'online',
            digest_path(SCRIPTS / 'acquire_sources.py')]))
        def do_acquire(_):
            if args.offline:
                if not manifest_f.is_file():
                    jdump(manifest_f, {
                        'generated_at': now(), 'period': {'start': per.start, 'end': per.end},
                        'offline_replay': True, 'sources': [],
                        'coverage': [{'group': g, 'status': 'NOT_RUN',
                                      'evidence': '离线回放未联网',
                                      'note': '结构回放，不得解释为无案例'}
                                     for g in ('SSE_MOMENTUM', 'SSE_COMMITTEE', 'SZSE_MOMENTUM',
                                               'CNINFO_FULLTEXT', 'BSE_MOMENTUM', 'CSRC_REGISTRATION')],
                        'limitations': ['离线回放：未采集官方源，覆盖矩阵全部 NOT_RUN']})
                if not (sources_dir / 'cninfo_announcements.json').is_file():
                    jdump(sources_dir / 'cninfo_announcements.json',
                          {'period': {'start': per.start, 'end': per.end},
                           'keywords': {}, 'total_unique': 0, 'offline_replay': True,
                           'announcements': []})
                if not (sources_dir / 'sse_momentum_window.json').is_file():
                    jdump(sources_dir / 'sse_momentum_window.json', [])
                return {'mode': 'offline'}
            cp = cli([SCRIPTS / 'acquire_sources.py', '--start', per.start, '--end', per.end,
                      '--out', sources_dir, '--manifest', manifest_f,
                      '--max-pdf', max_pdf,
                      *(['--soffice', soffice] if soffice else [])],
                     logs / 'ACQUIRE.log', timeout=args.timeout)
            return {'mode': 'online', 'stdout': cp.stdout.strip()[-400:]}
        stage(state, 'ACQUIRE', acq_key, do_acquire, args.retries, logs)
        print(json.dumps({'event': 'ACQUIRE', 'mode': 'offline' if args.offline else 'online'},
                         ensure_ascii=False), flush=True)

        # ---------------- EXTRACT ----------------
        curation_f = (args.curation_decisions or (run_dir / 'curation-decisions.json')).resolve()
        ext_key = digest_bytes('|'.join([digest_tree(sources_dir),
                                         digest_path(manifest_f) if manifest_f.is_file() else '',
                                         digest_path(SCRIPTS / 'extract_filings.py'),
                                         digest_path(SCRIPTS / 'normalize_ledger.py'),
                                         digest_path(curation_f) if curation_f.is_file() else '']))
        def do_extract(_):
            cp = cli([SCRIPTS / 'extract_filings.py',
                      '--sources', sources_dir, '--start', per.start, '--end', per.end,
                      '--out', ledger_f, '--db', run_dir / 'ledger.sqlite',
                      '--pending', pending_f],
                     logs / 'EXTRACT.log', timeout=args.timeout)
            detail = cp.stdout.strip()[-400:]
            if curation_f.is_file():
                raw_f = run_dir / 'event-ledger-raw.json'
                shutil.copy2(ledger_f, raw_f)
                cp2 = cli([SCRIPTS / 'normalize_ledger.py', raw_f, curation_f,
                           '--out', ledger_f,
                           '--report', run_dir / 'curation-report.json'],
                          logs / 'CURATE.log', timeout=args.timeout)
                detail += '\n' + cp2.stdout.strip()[-400:]
            return {'stdout': detail}
        stage(state, 'EXTRACT', ext_key, do_extract, args.retries, logs)

        # ---------------- METRICS ----------------
        met_key = digest_bytes(digest_path(ledger_f) + digest_path(SCRIPTS / 'weekly_ledger.py'))
        def do_metrics(_):
            cp = cli([SCRIPTS / 'weekly_ledger.py', ledger_f, '--out', metrics_f],
                     logs / 'METRICS.log', allow=(0, 1), timeout=args.timeout)
            if cp.returncode == 1 or not metrics_f.is_file():
                jdump(metrics_f, {'validation': 'FAIL', 'reason': '事件台账未通过口径校验',
                                  'cli_stdout': cp.stdout.strip()[-600:],
                                  'unconfirmed_events': [], 'coverage_declaration_complete': False})
            else:
                m = jload(metrics_f, {})
                if isinstance(m, dict) and 'validation' not in m:
                    m['validation'] = 'PASS'
                    jdump(metrics_f, m)
            return {'validation': jload(metrics_f, {}).get('validation')}
        stage(state, 'METRICS', met_key, do_metrics, log_dir=logs)

        # ---------------- BUILD ----------------
        bld_key = digest_bytes('|'.join([
            digest_path(ledger_f), digest_path(run_dir / 'period.json'),
            digest_path(run_dir / 'prev-fingerprints.json'),
            digest_path(SCRIPTS / 'build_package.py'),
            digest_path(SCRIPTS / 'build_reports.py'),
            digest_path(SCRIPTS / 'lib/ooxml.py'),
            digest_path(SCRIPTS / 'lib/period.py')]))
        def do_build(_):
            if content_dir.exists():
                shutil.rmtree(content_dir)
            content_dir.mkdir(parents=True)
            report = run_dir / 'build-report.json'
            build_stdout = cli([SCRIPTS / 'build_package.py',
                                '--prev', template_dir, '--out', content_dir,
                                '--period', run_dir / 'period.json', '--ledger', ledger_f,
                                '--report', report],
                               logs / 'BUILD.log', timeout=args.timeout).stdout.strip()[-400:]
            report_stdout = cli([SCRIPTS / 'build_reports.py', '--ledger', ledger_f,
                                 '--metrics', metrics_f, '--period', run_dir / 'period.json',
                                 '--outdir', content_dir, '--prev', template_dir],
                                logs / 'REPORTS.log', timeout=args.timeout).stdout.strip()[-400:]
            return {'stdout': build_stdout + '\n' + report_stdout}
        stage(state, 'BUILD', bld_key, do_build, args.retries, logs)

        # ---------------- RENDER ----------------

        if pdf_dir and not pdf_dir.is_dir():
            raise StageFailure(f'--pdf-source-dir 不存在或不是目录: {pdf_dir}')
        if not pdf_dir and (not soffice or not soffice.is_file()):
            raise StageFailure(f'PDF 渲染器不可用: {soffice or "未找到"}；'
                               '请设置 MA_WEEKLY_SOFFICE / --soffice，或提供 --pdf-source-dir')
        ren_key = digest_bytes('|'.join([
            digest_tree(content_dir, NATIVE),
            digest_tree(pdf_dir, ('.pdf', '.json')) if pdf_dir
            else (digest_path(soffice) if soffice.is_file() else str(soffice)),
            digest_path(fontconfig) if fontconfig and fontconfig.is_file() else '',
            digest_path(SCRIPTS / 'render_verify.py')]))
        def do_render(_):
            cmd = [SCRIPTS / 'render_verify.py', '--content', content_dir, '--out', package_dir,
                   '--prev', per.prev_dir,
                   '--visual', run_dir / 'visual', '--report', run_dir / 'render-report.json']
            if pdf_dir:
                cmd += ['--pdf-source-dir', pdf_dir]
            else:
                cmd += ['--soffice', soffice]
                if fontconfig:
                    cmd += ['--fontconfig', fontconfig]
            return {'stdout': cli(cmd, logs / 'RENDER.log', timeout=args.timeout).stdout.strip()[-400:]}
        stage(state, 'RENDER', ren_key, do_render, args.retries, logs)

        # ---------------- VALIDATE ----------------
        val_key = digest_bytes('|'.join([
            digest_tree(package_dir, ALL_SUFFIXES),
            digest_path(manifest_f) if manifest_f.is_file() else '',
            digest_path(pending_f) if pending_f.is_file() else '',
            digest_path(metrics_f) if metrics_f.is_file() else '',
            digest_path(run_dir / 'render-report.json'),
            digest_path(SCRIPTS / 'quality_gate.py')]))
        validation_f = run_dir / 'validation-report.json'
        def do_validate(_):
            cli([SCRIPTS / 'quality_gate.py', '--package', package_dir,
                 '--prev', per.prev_dir, '--period', run_dir / 'period.json',
                 '--run-dir', run_dir, '--report', validation_f],
                logs / 'VALIDATE.log', allow=(0, 2), timeout=args.timeout)
            return {'validation_status': jload(validation_f, {}).get('status')}
        stage(state, 'VALIDATE', val_key, do_validate, log_dir=logs)
        validation = jload(validation_f, {})
        status = validation.get('status', 'BLOCKED')
        exceptions = validation.get('soft_failures', []) + validation.get('hard_failures', [])
        print(json.dumps({'event': 'VALIDATE', 'status': status, 'exceptions': exceptions},
                         ensure_ascii=False), flush=True)

        # ---------------- NOTES ----------------
        notes_md = run_dir / '本期编制与待核说明.md'
        notes_docx = run_dir / '本期编制与待核说明.docx'
        note_key = digest_bytes('|'.join([status, digest_path(validation_f),
                                          digest_path(pending_f) if pending_f.is_file() else '',
                                          digest_path(SCRIPTS / 'build_notes.py')]))
        def do_notes(_):
            cli([SCRIPTS / 'build_notes.py', '--run-dir', run_dir,
                 '--period', run_dir / 'period.json', '--status', status,
                 '--md', notes_md, '--docx', notes_docx],
                logs / 'NOTES.log', timeout=args.timeout)
            return {'md': str(notes_md)}
        stage(state, 'NOTES', note_key, do_notes, log_dir=logs)

        # ---------------- PUBLISH ----------------
        if status == 'BLOCKED':
            state.data['stages']['PUBLISH'] = {'status': 'SKIPPED', 'reason': '硬门禁失败，禁止发布',
                                               'finished_at': now()}
            state.save()
            jdump(run_dir / 'blocked-report.json', {
                'as_of': args.as_of, 'status': 'BLOCKED',
                'hard_failures': validation.get('hard_failures', []),
                'diagnostics': {'building': str(package_dir), 'validation': str(validation_f),
                                'logs': str(logs)},
                'rule': '不把待核稿命名为最终成品；不发布部分文件'})
            return finish('BLOCKED', package=str(package_dir),
                          hard_failures=validation.get('hard_failures', []))
        if status == 'COMPLETE_WITH_EXCEPTIONS' and publish_policy == 'complete-only':
            state.data['stages']['PUBLISH'] = {'status': 'SKIPPED',
                                               'reason': 'publish-policy=complete-only 且存在软门禁例外',
                                               'finished_at': now()}
            state.save()
            return finish('COMPLETE_WITH_EXCEPTIONS', published=False,
                          package=str(package_dir), exceptions=exceptions)
        pub_key = digest_bytes('|'.join([status, digest_tree(package_dir, ALL_SUFFIXES),
                                         digest_path(notes_md), str(out_root),
                                         per.out_dirname, publish_policy]))
        def do_publish(_):
            report = run_dir / 'publish-report.json'
            cmd = [SCRIPTS / 'publish_package.py', '--building', package_dir,
                   '--out-root', out_root, '--dirname', per.out_dirname,
                   '--status', status, '--as-of', args.as_of,
                   '--period', run_dir / 'period.json', '--run-dir', run_dir,
                   '--exceptions', validation_f, '--note', notes_md,
                   '--report', report]
            cli(cmd, logs / 'PUBLISH.log', timeout=args.timeout)
            pub = jload(report, {})
            return {'final': pub.get('final'), 'revision': pub.get('revision')}
        pub = stage(state, 'PUBLISH', pub_key, do_publish, log_dir=logs)
        return finish(status, published=True, final=pub.get('final'),
                      exceptions=exceptions, notes=str(notes_md))

    except Exception as ex:  # noqa: BLE001 - 无人值守收口
        tb = ''.join(traceback.format_exception(type(ex), ex, ex.__traceback__))
        (logs / 'FATAL.txt').write_text(tb, encoding='utf-8')
        jdump(run_dir / 'blocked-report.json', {
            'as_of': args.as_of, 'status': 'BLOCKED',
            'error': f'{type(ex).__name__}: {ex}', 'traceback': tb[-4000:],
            'rule': '不把待核稿命名为最终成品；不发布部分文件'})
        return finish('BLOCKED', error=f'{type(ex).__name__}: {ex}', log=str(logs / 'FATAL.txt'))


# ---------------------------------------------------------------- 运行环境自举
# 主链子脚本分别依赖 openpyxl / python-docx / python-pptx / pypdf / lxml。
# v4.1.0 在 WorkBuddy Python 3.13 下会被 build_reports.ensure_libs() 注入
# Python 3.12 的 site-packages，PIL._imaging 等 C 扩展 ABI 不兼容导致 BUILD
# 阶段崩溃（2026-11-08 第42期盲测首次在线运行时实际触发）。由编排器在入口
# 一次性选定合格解释器并重新执行自身，保证所有子脚本继承同一解释器。
RUNTIME_MODULES = ('openpyxl', 'docx', 'pptx', 'pypdf', 'lxml')


def _runtime_ok(py: str) -> bool:
    try:
        cp = subprocess.run([py, '-c', 'import ' + ','.join(RUNTIME_MODULES)],
                            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                            timeout=30)
        return cp.returncode == 0
    except (OSError, subprocess.SubprocessError):
        return False


def ensure_runtime() -> None:
    """缺主链依赖时自动切到同机合格解释器；找不到则清晰阻断，不混用 ABI。"""
    import importlib.util
    missing = [m for m in RUNTIME_MODULES if importlib.util.find_spec(m) is None]
    if not missing:
        return
    if os.environ.get('MA_WEEKLY_RUNTIME_CHECKED') == '1':
        print(json.dumps({'event': 'BLOCKED', 'reason': 'RUNTIME_DEPENDENCY_MISSING',
                          'modules': missing,
                          'hint': 'pip install python-docx python-pptx openpyxl pypdf lxml'},
                         ensure_ascii=False), flush=True)
        raise SystemExit(2)
    home = Path.home()
    candidates = [os.environ.get('MA_WEEKLY_PYTHON'),
                  shutil.which('python3.12'), shutil.which('python3'),
                  str(home / '.cache/codex-runtimes/codex-primary-runtime'
                      '/dependencies/python/bin/python3'),
                  '/opt/homebrew/bin/python3.12', '/usr/local/bin/python3.12',
                  '/Library/Frameworks/Python.framework/Versions/3.12/bin/python3']
    seen = {str(Path(sys.executable).resolve())}
    for cand in candidates:
        if not cand:
            continue
        p = str(Path(cand).resolve())
        if p in seen or not Path(p).is_file():
            continue
        seen.add(p)
        if _runtime_ok(p):
            print(json.dumps({'event': 'RUNTIME_SWITCH', 'from': sys.executable,
                              'to': p}, ensure_ascii=False), flush=True)
            env = dict(os.environ, MA_WEEKLY_RUNTIME_CHECKED='1')
            os.execve(p, [p, str(Path(__file__).resolve()), *sys.argv[1:]], env)
    print(json.dumps({'event': 'BLOCKED', 'reason': 'NO_QUALIFIED_INTERPRETER',
                      'hint': '设置 MA_WEEKLY_PYTHON 指向已装依赖的 python，'
                              '或执行 pip install python-docx python-pptx openpyxl pypdf lxml'},
                     ensure_ascii=False), flush=True)
    raise SystemExit(2)


if __name__ == '__main__':
    ensure_runtime()
    sys.exit(main())
