#!/usr/bin/env python3
"""自动质量门禁：硬门禁失败即阻断发布；软门禁不足不阻断出件，但强制标注状态。

硬门禁保证"结构完整、文件可用、原件未改、无重复"；
软门禁反映"来源覆盖、结论性待核、PDF 版式"等无法由脚本单方面宣称通过的事项。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from zipfile import ZipFile

import openpyxl

REQUIRED = {f'xlsx_{i}' for i in range(14) if i != 5} | {'weekly_docx', 'dynamic_docx', 'market_ppt'}
NATIVE = ('.xlsx', '.xlsm', '.docx', '.pptx')


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


def sha(p: Path) -> str:
    return hashlib.sha256(p.read_bytes()).hexdigest()


def jload(p: Path, default=None):
    if not p.is_file():
        return default
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except (ValueError, OSError):
        return default


FORMULA_ERRORS = ('#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#N/A', '#NUM!', '#NULL!')


def formula_fingerprint(xlsx_files: list[Path]) -> tuple[set[str], set[str], list[str]]:
    """对照公式层与缓存层，返回（错误值键，缺缓存键，扫描异常）。

    键按「文件角色/工作表名/单元格坐标」归一化，跨期文件名不同也能比较。
    现场上期人工成果本身就含 #N/A 等 VLOOKUP 无匹配的缓存值，因此绝对
    「零错误」不是合理硬门禁；真正要阻断的是**本期新增**的错误值。

    两个工作簿都以 read_only 打开，避免把整本重新序列化。
    """
    errors: set[str] = set()
    missing: set[str] = set()
    failures: list[str] = []
    for p in xlsx_files:
        try:
            wf = openpyxl.load_workbook(p, read_only=True, data_only=False)
            wc = openpyxl.load_workbook(p, read_only=True, data_only=True)
            try:
                r = role(p.name) or p.stem
                for sf, sc in zip(wf.worksheets, wc.worksheets):
                    for rowf, rowc in zip(sf.iter_rows(), sc.iter_rows()):
                        for cf, cc in zip(rowf, rowc):
                            if cf.data_type != 'f':
                                continue
                            v = cc.value
                            tag = f'{r}/{sf.title}/{cf.coordinate}'
                            if isinstance(v, str) and v in FORMULA_ERRORS:
                                errors.add(f'{tag}:{v}')
                            elif v is None:
                                missing.add(tag)
            finally:
                wf.close()
                wc.close()
        except Exception as ex:
            failures.append(f'{p.name}: 公式缓存扫描失败 {type(ex).__name__}: {ex}')
    return errors, missing, failures


def office_manifest_status(run_dir: Path) -> dict:
    """可选增强：Office 原生 PDF 导出清单。不存在时返回 NOT_PROVIDED。"""
    f = run_dir / 'office-pdf' / 'office-export-manifest.json'
    if not f.is_file():
        return {'status': 'NOT_PROVIDED', 'exports': []}
    d = jload(f, {}) or {}
    return {'status': d.get('status', 'INCOMPLETE'), 'exports': d.get('exports', [])}


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--package', required=True, type=Path)
    ap.add_argument('--prev', required=True, type=Path)
    ap.add_argument('--period', required=True, type=Path)
    ap.add_argument('--run-dir', required=True, type=Path)
    ap.add_argument('--report', required=True, type=Path)
    args = ap.parse_args()
    period = json.loads(args.period.read_text(encoding='utf-8'))
    hard: list[dict] = []
    soft: list[dict] = []

    def check(cond: bool, item: str, detail: str, bucket=hard):
        bucket.append({'item': item, 'pass': bool(cond), 'detail': detail})

    pkg = args.package
    files = sorted(p for p in pkg.iterdir() if p.is_file()) if pkg.exists() else []
    native = [p for p in files if p.suffix.lower() in NATIVE and not p.name.startswith('~$')]
    pdfs = [p for p in files if p.suffix.lower() == '.pdf']
    roles: dict[str, str] = {}
    for p in native:
        r = role(p.name)
        if r:
            if r in roles:
                check(False, '角色唯一', f'重复角色 {r}: {roles[r]}, {p.name}')
            roles[r] = p.name
    missing = sorted(REQUIRED - set(roles))
    check(not missing, '16个原生文件角色齐全',
          f'共 {len(native)} 个；缺 {missing}' if missing else f'共 {len(native)} 个，角色完整')
    check(len(pdfs) == 3, '3份PDF', f'实际 {len(pdfs)} 份')

    bad_zip, bad_wb = [], []
    for p in native:
        try:
            with ZipFile(p) as z:
                if z.testzip():
                    bad_zip.append(p.name)
            if p.suffix.lower() in ('.xlsx', '.xlsm'):
                wb = openpyxl.load_workbook(p, read_only=True)
                if not wb.sheetnames:
                    bad_wb.append(p.name)
                wb.close()
        except Exception as ex:
            bad_zip.append(f'{p.name}: {type(ex).__name__}: {ex}')
    check(not bad_zip, 'Office包结构完整', f'损坏 {bad_zip}' if bad_zip else '全部 ZIP 校验通过')
    check(not bad_wb, 'Excel可打开且有工作表', f'异常 {bad_wb}' if bad_wb else '全部可读取')

    # v4.3.8 硬门禁：公告正文不得整段塞入 Excel 单元格。
    # 第 39→40 期盲测实测生成稿 14 处（最长 1200 字）、现场稿 0 处；
    # 现场人工稿中引用证券代码的长摘要只命中 1 个抬头特征，不会误伤。
    if str(Path(__file__).resolve().parent) not in sys.path:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
    from lib import celltext  # noqa: E402 - 与 build_package 共用同一判据
    cell_hits = celltext.scan_raw_filing_cells(
        [p for p in native if p.suffix.lower() in ('.xlsx', '.xlsm')])
    check(not cell_hits, 'Excel单元格无整段公告正文',
          f'{len(cell_hits)} 处，例：{[(h["sheet"], h["cell"], h["length"]) for h in cell_hits[:5]]}'
          if cell_hits else '全部单元格为有界展示文本')

    # 公式缓存：与上期对比，只阻断「本期新增」的错误值。
    # 上期人工成果自带 #N/A（VLOOKUP 无匹配）等缓存值，绝对零错误会误伤每期。
    xlsx_files = [p for p in native if p.suffix.lower() in ('.xlsx', '.xlsm')]
    prev_xlsx = ([p for p in sorted(args.prev.iterdir()) if p.is_file()
                  and p.suffix.lower() in ('.xlsx', '.xlsm')]
                 if args.prev.is_dir() else [])
    cur_err, cur_miss, scan_fail = formula_fingerprint(xlsx_files)
    prev_err, prev_miss, _ = formula_fingerprint(prev_xlsx)
    new_errors = sorted(cur_err - prev_err)
    new_missing = sorted(cur_miss - prev_miss)
    check(not scan_fail, 'Excel公式缓存可扫描', f'{scan_fail[:4]}' if scan_fail else '全部可扫描')
    check(not new_errors, '本期未新增Excel公式错误值',
          f'{new_errors[:8]}' if new_errors else
          f'无新增；沿用上期 {len(cur_err & prev_err)} 个既有缓存错误值')

    span = f'{period["start"]}至{period["end"]}'
    prev_span = f'{period["prev_start"]}至{period["prev_end"]}'
    stale = [p.name for p in files if prev_span in p.name]
    wrong = [p.name for p in native if span not in p.name and re.search(r'\d{8}至\d{8}', p.name)]
    check(not stale and not wrong, '文件名期次正确',
          f'残留上期 {stale}；日期不符 {wrong}' if (stale or wrong) else f'全部含 {span}')

    fp = args.run_dir / 'prev-fingerprints.json'
    before = json.loads(fp.read_text(encoding='utf-8')) if fp.exists() else {}
    changed = [n for n, h in before.items()
               if not (args.prev / n).exists() or sha(args.prev / n) != h]
    check(not changed, '上期原件未被改写', f'变化 {changed}' if changed else f'{len(before)} 个原件哈希一致')

    temps = [p.name for p in files if p.name.startswith('~$') or p.suffix.lower() == '.tmp']
    check(not temps, '无临时/半成品混入', f'发现 {temps}' if temps else '无')

    ledger_f = args.run_dir / 'event-ledger.json'
    dup = []
    if ledger_f.exists():
        led = json.loads(ledger_f.read_text(encoding='utf-8'))
        seen = set()
        for e in led.get('events', []):
            k = (e['deal_id'], e['type'], e['date'], e.get('sequence', 0))
            if k in seen:
                dup.append(k)
            seen.add(k)
    check(not dup, '事件台账无重复记录', f'重复 {len(dup)} 条' if dup else '无重复')

    render_f = args.run_dir / 'render-report.json'
    render = json.loads(render_f.read_text(encoding='utf-8')) if render_f.exists() else {}
    check(render.get('status') == 'PASS' and all(
        i.get('pages', 0) > 0 for i in render.get('inspections', [])),
        'PDF可打开且逐页渲染',
        f"status={render.get('status')}，页数={[i.get('pages') for i in render.get('inspections', [])]}")

    # ---- 软门禁 ----
    man_f = args.run_dir / 'source-manifest.json'
    cov = json.loads(man_f.read_text(encoding='utf-8')).get('coverage', []) if man_f.exists() else []
    incomplete = [c['group'] for c in cov if c.get('status') not in ('COMPLETE', 'NOT_APPLICABLE')]
    check(not incomplete, '来源覆盖完整',
          f'非完整组：{incomplete}' if incomplete else '全部 COMPLETE', soft)

    pending_f = args.run_dir / 'pending-review.json'
    pending_data = json.loads(pending_f.read_text(encoding='utf-8')) if pending_f.exists() else {}
    pend = pending_data.get('count', 0)
    blocking = pending_data.get('blocking_count', pend)
    check(blocking == 0, '无待核结论性字段', f'待核 {pend} 项，其中阻断结论 {blocking} 项' if pending_f.exists() else '缺失', soft)

    comparisons = render.get('page_comparisons', [])
    drift = [c for c in comparisons if c.get('prev_pages') and c['prev_pages'] != c['pages']]
    check(not drift, 'PDF分页与上期基准一致',
          f'{len(drift)} 份存在分页差异' if drift else '一致', soft)

    suspect = [(i.get('pdf'), i.get('blank_pages'), i.get('low_cjk_pages'))
               for i in render.get('inspections', [])
               if i.get('blank_pages') or i.get('low_cjk_pages')]
    check(not suspect, 'PDF无空白页/中文缺字页',
          f'{suspect}' if suspect else '逐页文本检查正常', soft)

    # 可选增强：固定 Office 发布环境。默认 LibreOffice 链路不因此阻断，
    # 但如实记为软例外，避免把预览 PDF 冒称为 Microsoft Office 发布版。
    om = office_manifest_status(args.run_dir)
    if render.get('pdf_source') == 'office-native':
        check(om['status'] == 'COMPLETE' and len(om['exports']) == 3,
              'Microsoft Office发布PDF清单',
              f"status={om['status']}，登记 {len(om['exports'])} 份")
    else:
        check(False, '发布级Office原生PDF（可选增强）',
              '默认固定 LibreOffice 预览渲染，分页不等价于 Microsoft Word/PowerPoint 发布版；'
              '如需发布级版式请导出三份 PDF 并用 --pdf-source-dir 装配', soft)

    check(not new_missing, 'Excel公式缓存已在Office重算',
          f'本期新增无缓存公式 {len(new_missing)} 个，例：{new_missing[:6]}；'
          f'沿用上期 {len(cur_miss & prev_miss)} 个' if new_missing else
          f'无新增；沿用上期 {len(cur_miss & prev_miss)} 个既有缺缓存公式（须在最终 Office 重算）',
          soft)

    metrics_f = args.run_dir / 'metrics.json'
    m = json.loads(metrics_f.read_text(encoding='utf-8')) if metrics_f.exists() else {}
    # 必须显式声明 PASS；缺字段或 FAIL 都不能算通过，避免把"没跑"当成"通过"
    check(m.get('validation') == 'PASS', '台账指标校验通过',
          f"validation={m.get('validation', 'MISSING')}", soft)

    hard_fail = [c for c in hard if not c['pass']]
    soft_fail = [c for c in soft if not c['pass']]
    status = 'BLOCKED' if hard_fail else ('COMPLETE_WITH_EXCEPTIONS' if soft_fail else 'COMPLETE')
    report = {'status': status, 'hard_gates': hard, 'soft_gates': soft,
              'hard_failures': [c['item'] for c in hard_fail],
              'soft_failures': [c['item'] for c in soft_fail],
              'package': str(pkg), 'files': len(files),
              'pdf_source': render.get('pdf_source'),
              'interpretation': {
                  'COMPLETE': '全部硬/软门禁通过，可作为完整成果发布',
                  'COMPLETE_WITH_EXCEPTIONS': '结构性完整，但存在来源覆盖/待核/版式例外；'
                                             '成果包已出件，必须随附《编制与待核说明》，不得冒称最终发布版',
                  'BLOCKED': '硬门禁失败，禁止发布',
              }[status]}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'status': status, 'hard_fail': len(hard_fail),
                      'soft_fail': len(soft_fail), 'files': len(files)}, ensure_ascii=False))
    return 0 if status != 'BLOCKED' else 2


if __name__ == '__main__':
    sys.exit(main())
