#!/usr/bin/env python3
"""装配 3 份 PDF + 逐页 PNG 视觉核验。

两种 PDF 来源，均可无人值守：
  1. 默认 `--soffice`：固定 bundled LibreOffice + 固定字体配置导出。
     分页可能与 Microsoft Word/PowerPoint 发布版不同，脚本如实记录为软例外。
  2. `--pdf-source-dir`：使用 Microsoft Office（可由电脑操控）以「最适合打印」
     导出的 PDF，直接装配。仍需逐页渲染检查，不因来源可信就跳过。

绝不把「PDF 可打开」等同于「版式通过」。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

CJK_RE = re.compile(r'[\u4e00-\u9fff]')
NATIVE = ('.xlsx', '.xlsm', '.docx', '.pptx')


def convert(soffice: Path, src: Path, outdir: Path, fontconfig: Path | None) -> dict:
    if src.name.startswith('~$'):
        return {'source': src.name, 'pdf': None, 'ok': True, 'skipped': 'Office 临时锁文件',
                'bytes': 0, 'stdout': '', 'stderr': ''}
    env = os.environ.copy()
    if fontconfig:
        env['FONTCONFIG_FILE'] = str(fontconfig.resolve())
    pdf = outdir / (src.stem + '.pdf')
    if pdf.exists():
        pdf.unlink()
    last_stdout = last_stderr = ''
    for _attempt in range(2):
        with tempfile.TemporaryDirectory(prefix='ma-weekly-lo-') as profile:
            profile_uri = Path(profile).resolve().as_uri()
            cp = subprocess.run([str(soffice), f'-env:UserInstallation={profile_uri}',
                                 '--headless', '--nologo', '--nofirststartwizard',
                                 '--convert-to', 'pdf', '--outdir', str(outdir), str(src)],
                                capture_output=True, text=True, timeout=300, env=env)
        last_stdout, last_stderr = cp.stdout[-500:], cp.stderr[-500:]
        if cp.returncode == 0 and pdf.is_file() and pdf.stat().st_size > 0:
            return {'source': src.name, 'pdf': pdf.name, 'ok': True, 'bytes': pdf.stat().st_size,
                    'stdout': last_stdout, 'stderr': last_stderr}
    return {'source': src.name, 'pdf': None, 'ok': False, 'bytes': 0,
            'stdout': last_stdout, 'stderr': last_stderr,
            'error': f'转换失败，返回码 {cp.returncode}'}


def copy_office_pdfs(pdf_dir: Path, outdir: Path) -> list[dict]:
    """把电脑操控/Office 导出的 PDF 按角色装配，返回转换记录。"""
    wanted = {
        'weekly_docx': lambda n: n.startswith('5、') and n.lower().endswith('.pdf'),
        'dynamic_docx': lambda n: n.startswith('A股市场并购重组动态') and n.lower().endswith('.pdf'),
        'market_ppt': lambda n: n.startswith('并购重组市场周报') and n.lower().endswith('.pdf'),
    }
    out = []
    for role, match in wanted.items():
        cand = next((p for p in sorted(pdf_dir.iterdir()) if p.is_file() and match(p.name)), None)
        if cand is None:
            out.append({'role': role, 'ok': False, 'error': f'{pdf_dir} 未找到该角色 PDF'})
            continue
        dst = outdir / cand.name
        shutil.copy2(cand, dst)
        out.append({'role': role, 'source': cand.name, 'pdf': dst.name, 'ok': True,
                    'bytes': dst.stat().st_size})
    return out


def inspect_pdf(pdf: Path, visual_dir: Path, scale: float = 0.55) -> dict:
    result: dict = {'pdf': pdf.name, 'pages': 0, 'text_chars': 0, 'cjk_chars': 0,
                    'blank_pages': [], 'low_cjk_pages': [], 'pngs': []}
    try:
        from pypdf import PdfReader
        reader = PdfReader(str(pdf), strict=False)
        result['pages'] = len(reader.pages)
        sub = visual_dir / pdf.stem
        if sub.exists():
            shutil.rmtree(sub)
        sub.mkdir(parents=True, exist_ok=True)
        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ''
            result['text_chars'] += len(text)
            cjk = len(CJK_RE.findall(text))
            result['cjk_chars'] += cjk
            if len(text.strip()) < 20:
                result['blank_pages'].append(i + 1)
            elif cjk == 0 and len(text) > 40:
                result['low_cjk_pages'].append(i + 1)
        pdftoppm = shutil.which('pdftoppm')
        if not pdftoppm:
            raise RuntimeError('未找到 pdftoppm；请安装 poppler 并加入 PATH')
        prefix = sub / 'page'
        cp = subprocess.run([pdftoppm, '-png', '-r', str(max(48, int(scale * 144))),
                             str(pdf), str(prefix)], capture_output=True, text=True, timeout=300)
        if cp.returncode != 0:
            raise RuntimeError((cp.stderr or cp.stdout)[-500:])
        generated = sorted(sub.glob('page-*.png'))
        for i, png in enumerate(generated, start=1):
            target = sub / f'p{i:03d}.png'
            if png != target:
                png.replace(target)
            result['pngs'].append(str(target))
        if len(result['pngs']) != result['pages']:
            result['error'] = f'PNG页数 {len(result["pngs"])} 与PDF页数 {result["pages"]} 不一致'
    except Exception as ex:
        result['error'] = f'{type(ex).__name__}: {ex}'
    return result

def ref_pages(prev: Path) -> dict[str, int]:
    out = {}
    if not prev:
        return out
    for p in prev.glob('*.pdf'):
        try:
            from pypdf import PdfReader
            out[p.stem] = len(PdfReader(str(p), strict=False).pages)
        except Exception:
            pass
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--content', required=True, type=Path)
    ap.add_argument('--out', required=True, type=Path)
    ap.add_argument('--prev', type=Path)
    ap.add_argument('--soffice', type=Path)
    ap.add_argument('--fontconfig', type=Path)
    ap.add_argument('--pdf-source-dir', type=Path,
                    help='Office 原生导出的 3 份 PDF 所在目录；提供后不再调用 LibreOffice')
    ap.add_argument('--visual', type=Path)
    ap.add_argument('--report', required=True, type=Path)
    args = ap.parse_args()
    if args.out.exists():
        shutil.rmtree(args.out)
    args.out.mkdir(parents=True)
    visual = args.visual or args.out.parent / 'visual'
    visual.mkdir(parents=True, exist_ok=True)

    for p in sorted(args.content.iterdir()):
        if p.is_file() and p.suffix.lower() in NATIVE and not p.name.startswith('~$'):
            shutil.copy2(p, args.out / p.name)

    conversions: list[dict] = []
    if args.pdf_source_dir:
        if not args.pdf_source_dir.is_dir():
            print(json.dumps({'status': 'FAIL', 'errors': ['PDF来源目录不存在']},
                             ensure_ascii=False))
            return 1
        conversions = copy_office_pdfs(args.pdf_source_dir, args.out)
        engine = f'office-native:{args.pdf_source_dir}'
    else:
        if not args.soffice or not Path(args.soffice).is_file():
            print(json.dumps({'status': 'FAIL', 'errors': ['缺少 --soffice 渲染器']},
                             ensure_ascii=False))
            return 1
        for p in sorted(args.out.iterdir()):
            if p.suffix.lower() in ('.docx', '.pptx'):
                conversions.append(convert(args.soffice, p, args.out, args.fontconfig))
        engine = f'libreoffice:{args.soffice}'

    inspections = [inspect_pdf(p, visual) for p in sorted(args.out.glob('*.pdf'))]
    baseline = ref_pages(args.prev)
    comparisons = []
    for ins in inspections:
        stem = Path(ins['pdf']).stem
        key = next((k for k in baseline if k.split('_')[0][:6] == stem.split('_')[0][:6]), None)
        comparisons.append({'pdf': ins['pdf'], 'pages': ins['pages'],
                            'prev_pages': baseline.get(key) if key else None,
                            'baseline_matched': bool(key)})

    errors = [f"{c.get('source') or c.get('role')}:{c.get('error', 'conversion failed')}"
              for c in conversions if not c.get('ok')]
    for i in inspections:
        if i.get('error') or i['pages'] == 0:
            errors.append(i['pdf'] + ':' + str(i.get('error', 'no pages')))
    expected_prefixes = ('5、', 'A股市场并购重组动态', '并购重组市场周报')
    for prefix in expected_prefixes:
        if not any(ins['pdf'].startswith(prefix) for ins in inspections):
            errors.append(f'缺少 PDF 角色：{prefix}')
    if len(inspections) != 3:
        errors.append(f'PDF 数量应为 3，实际 {len(inspections)}')
    limitation = ('bundled LibreOffice 分页与 Microsoft Word 发布版可能不一致；'
                  '如需发布级版式，请以固定 Office 环境导出并提供 --pdf-source-dir'
                  if not args.pdf_source_dir else
                  'PDF 来自 Office 原生导出；仍须逐页检查文字、图表与母版漂移')
    report = {'status': 'PASS' if not errors else 'FAIL', 'errors': errors,
              'conversions': conversions, 'inspections': inspections,
              'page_comparisons': comparisons, 'renderer': engine,
              'fontconfig': str(args.fontconfig) if args.fontconfig else None,
              'pdf_source': 'office-native' if args.pdf_source_dir else 'libreoffice',
              'limitation': limitation}
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'status': report['status'], 'pdfs': len(inspections),
                      'pngs': sum(len(i['pngs']) for i in inspections),
                      'errors': errors}, ensure_ascii=False))
    return 0 if not errors else 1


if __name__ == '__main__':
    sys.exit(main())
