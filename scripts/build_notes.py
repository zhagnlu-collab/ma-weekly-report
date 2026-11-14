#!/usr/bin/env python3
"""生成《本期编制与待核说明》：来源覆盖、事件清单、勾稽、未决事项、版式状态。

状态不是 COMPLETE 时，说明首页明确标注"待复核稿 / 禁止作为最终发布版"，
避免把结构完整的自动产出误认为内容已全部核实。
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path


def load(p: Path, default):
    if not p or not p.is_file():
        return default
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except (ValueError, OSError):
        return default


def sections(run_dir: Path, period: dict, status: str) -> tuple[str, list[tuple[str, list[str]]]]:
    manifest = load(run_dir / 'source-manifest.json', {})
    ledger = load(run_dir / 'event-ledger.json', {})
    pending = load(run_dir / 'pending-review.json', {})
    validation = load(run_dir / 'validation-report.json', {})
    build = load(run_dir / 'build-report.json', {})
    render = load(run_dir / 'render-report.json', {})
    metrics = load(run_dir / 'metrics.json', {})
    office_manifest = load(run_dir / 'office-pdf' / 'office-export-manifest.json', {})

    span = f"{period.get('start', '')} 至 {period.get('end', '')}"
    title = f"并购周报 · 第{period.get('issue_no', '?')}期 编制与待核说明"
    if status == 'BLOCKED':
        banner = '【阻断】硬质量门禁未通过，未发布成果包；以下为诊断稿，禁止对外使用。'
    elif status == 'COMPLETE_WITH_EXCEPTIONS':
        banner = ('【待复核稿】结构已完整出件，但存在来源覆盖、结论性待核字段或版式例外；'
                  '未经人工复核不得命名为最终发布版、不得对外发送。')
    elif status == 'COMPLETE':
        banner = '【完整】硬/软质量门禁全部通过。'
    else:
        banner = f'【状态 {status}】请按下列门禁结果判断可用性。'

    out: list[tuple[str, list[str]]] = []
    out.append(('一、本期口径', [
        f"统计区间：{span}（Asia/Shanghai，闭区间，截止日 23:59:59）",
        f"期号：{period.get('year', '')}年第{period.get('issue_no', '?')}期；上期：{period.get('prev_dir', '')}",
        f"生成时间：{datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')}",
        f"自动运行状态：{status}",
        *[f"期次推导提示：{n}" for n in period.get('notes', [])],
    ]))

    cov = manifest.get('coverage', [])
    out.append(('二、检索覆盖', [
        f"来源记录 {len(manifest.get('sources', []))} 条；检索组 {len(cov)} 组。",
        *[f"{c.get('group')}：{c.get('status')} — {c.get('note', '')}" for c in cov],
        '说明：状态为程序声明的采集覆盖，不等于事实核验完成；列表页只产生候选。',
    ]))

    events = ledger.get('events', [])
    cases = {c['deal_id']: c for c in ledger.get('cases', [])}
    lines = [f"候选交易 {len(cases)} 笔，事件 {len(events)} 条（CONFIRMED "
             f"{sum(1 for e in events if e.get('status') == 'CONFIRMED')} 条）。"]
    for e in events[:60]:
        c = cases.get(e.get('deal_id'), {})
        lines.append(f"{e.get('date', '')}｜{c.get('code', '')} {c.get('name', '')}｜"
                     f"{e.get('type', '')}｜{e.get('status', '')}｜{e.get('note', '')[:60]}")
    if len(events) > 60:
        lines.append(f"……另有 {len(events) - 60} 条，详见 event-ledger.json。")
    out.append(('三、本期事件清单', lines))

    bridge = (build or {}).get('bridge', {})
    out.append(('四、统计口径与数字桥接', [
        f"指标校验：{metrics.get('validation', '未运行')}；"
        f"未确认事件 {len(metrics.get('unconfirmed_events', []))} 条。",
        f"年度累计按『上期正文基数 + 本期已核实增量』桥接，共应用 {len(bridge.get('applied', []))} 处、"
        f"跳过 {len(bridge.get('skipped', []))} 处。",
        *[f"{a.get('key')}：{a.get('old')} → {a.get('new')}（基数 {a.get('base')}，增量 {a.get('delta')}）"
          for a in bridge.get('applied', [])],
        *[f"未桥接：{s.get('key')} — {s.get('reason')}" for s in bridge.get('skipped', [])],
    ]))

    gates = validation.get('hard_gates', []) + validation.get('soft_gates', [])
    out.append(('五、已核对勾稽与门禁', [
        f"硬门禁失败 {len(validation.get('hard_failures', []))} 项；"
        f"软门禁失败 {len(validation.get('soft_failures', []))} 项。",
        *[f"[{'硬' if g in validation.get('hard_gates', []) else '软'}]"
          f"{'通过' if g.get('pass') else '未通过'}｜{g.get('item')}：{g.get('detail')}"
          for g in gates],
    ]))

    items = pending.get('items', [])
    out.append(('六、未决事项与待核字段', [
        f"待核线索 {pending.get('count', len(items))} 项，其中影响结论/入表的阻断项 {pending.get('blocking_count', len(items))} 项。"
        + ('结论性字段不得由程序猜测填列，须回公告全文或启用第二条解析链。' if items else ''),
        *[f"· {i.get('item', '')}｜{i.get('reason', '')}｜处理：{i.get('action', '')}"
          for i in items if i.get('blocking')][:40],
        *([f"……另有非阻断线索 {sum(1 for i in items if not i.get('blocking'))} 项，详见 pending-review.json。"]
          if any(not i.get('blocking') for i in items) else []),
    ]))

    out.append(('七、公式、图表与版式状态', [
        'Excel：采用 OOXML 局部补丁，保留公式、批注、合并区、图表与外链；'
        'openpyxl 不计算公式，缓存值沿用上期，涉及公式的汇总须在最终 Office 环境重算核对。',
        f"新增/更新行以 RGB(226,239,218) 标记，写入文件 "
        f"{len((build or {}).get('excel_plan', {}))} 个。",
        f"PDF 渲染器：{render.get('renderer', '未运行')}；"
        f"PDF 来源：{render.get('pdf_source', '未运行')}；"
        f"逐页 PNG {sum(len(i.get('pngs', [])) for i in render.get('inspections', []))} 张。",
        f"Office 原生导出清单：{office_manifest.get('status', '未提供')}，"
        f"登记 {len(office_manifest.get('exports', []))} 份。",
        f"版式限制：{render.get('limitation', '未运行渲染')}",
        '如需与 Microsoft Word/PowerPoint 发布版分页一致，须在固定 Office 环境以「最适合打印」'
        '另存为三份 PDF、登记导出清单，并用 --pdf-source-dir 重新装配后逐页比对。',
        'Excel 公式缓存若未重算，须在最终 Office 环境执行「公式→立即计算」并复核错误值；'
        '程序不会把未重算的缓存值声称为已刷新。',
    ]))

    out.append(('八、随附证据与运行文件', [
        'source-manifest.json（原始响应、URL、访问时间、哈希、检索页数）',
        'event-ledger.json / ledger.sqlite（交易、事件、来源、证据、覆盖）',
        'metrics.json（周度/年度/截止日指标快照）、build-report.json（文本替换与追加行审计）',
        'render-report.json + visual/（逐页 PNG）、validation-report.json、artifact-manifest.json',
        '引用证据请给出自然文件名＋页码/章节/原文摘录，或 Excel 表名＋单元格＋主体锚点。',
    ]))
    return f'# {title}\n\n> {banner}\n', out


def write_markdown(path: Path, header: str, blocks: list[tuple[str, list[str]]]):
    lines = [header]
    for h, items in blocks:
        lines.append(f'\n## {h}\n')
        lines.extend(f'- {x}' for x in items if x)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text('\n'.join(lines) + '\n', encoding='utf-8')


def write_docx(path: Path, header: str, blocks: list[tuple[str, list[str]]]) -> bool:
    try:
        from docx import Document
        from docx.shared import Pt
    except ImportError:
        return False
    doc = Document()
    doc.add_heading(header.lstrip('# ').strip(), level=1)
    banner = header.split('> ', 1)[-1].strip()
    p = doc.add_paragraph(banner)
    p.runs[0].bold = True
    for h, items in blocks:
        doc.add_heading(h, level=2)
        for x in items:
            if x:
                doc.add_paragraph(x, style='List Bullet')
    for s in doc.sections:
        pass
    path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(path)
    return True


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--run-dir', required=True, type=Path)
    ap.add_argument('--period', required=True, type=Path)
    ap.add_argument('--status', default='COMPLETE_WITH_EXCEPTIONS')
    ap.add_argument('--md', required=True, type=Path)
    ap.add_argument('--docx', type=Path)
    args = ap.parse_args()
    period = load(args.period, {})
    header, blocks = sections(args.run_dir, period, args.status)
    write_markdown(args.md, header, blocks)
    made = write_docx(args.docx, header, blocks) if args.docx else False
    print(json.dumps({'md': str(args.md), 'docx': str(args.docx) if made else None,
                      'sections': len(blocks), 'status': args.status}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
