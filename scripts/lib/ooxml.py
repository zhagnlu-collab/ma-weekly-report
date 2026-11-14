#!/usr/bin/env python3
"""OOXML 局部补丁：只改动目标文本节点/单元格，原样复制其余 ZIP 部件。

现场 13 个 Excel 含批注、drawing、chart、externalLink；openpyxl 整本保存会丢失
这些对象。本模块在字节层面做最小替换，保证格式、公式、合并区、图表与外链保真。
"""
from __future__ import annotations

import html
import re
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET

TEXT_NODE_RE = re.compile(r'(<(?:\w+:)?t(?:\s[^>]*)?>)(.*?)(</(?:\w+:)?t>)', re.S)
CELL_RE = re.compile(r'<c\s+r="(?P<ref>[A-Z]+\d+)"(?P<attrs>[^>]*?)(?:/>|>(?P<body>.*?)</c>)', re.S)
ROW_RE = re.compile(r'<row(?P<attrs>\s[^>]*)>(?P<body>.*?)</row>', re.S)
ROW_TAG_RE = re.compile(r'<row(?P<attrs>\s[^>]*?)(?:/>|>.*?</row>)', re.S)
DIMENSION_RE = re.compile(r'(<dimension\s+ref="[A-Z]+\d+:[A-Z]+)(?P<last>\d+)(")', re.S)
XML_PARTS = ('.xml', '.rels')


def _patch_text_nodes(data: bytes, replacements: dict[str, str]) -> tuple[bytes, dict[str, int]]:
    """跨 run 的安全文本替换。

    把同一 XML 部件内所有 <t> 节点拼成 joined，求出不重叠的替换区间，
    再按每个节点的原始字符范围重建该节点内容：替换文本只写进「区间起点所在」
    的节点，区间覆盖到的其余字符（无论落在哪个节点）一律删除。
    这样既保留 run 级格式，也不会出现旧文本残留（例如
    "截至2026年11月7日2026年10月31日"）。

    键支持两种形式：
      "原文"        -> 精确子串
      "re:<正则>"   -> 正则匹配，替换值可用 \\1 等反向引用
    """
    if not replacements:
        return data, {}
    try:
        text = data.decode('utf-8')
    except UnicodeDecodeError:
        return data, {k: 0 for k in replacements}
    matches = list(TEXT_NODE_RE.finditer(text))
    if not matches:
        return data, {k: 0 for k in replacements}
    segments = [html.unescape(m.group(2)) for m in matches]
    joined = ''.join(segments)
    positions: list[tuple[int, int]] = []
    cursor = 0
    for seg in segments:
        positions.append((cursor, cursor + len(seg)))
        cursor += len(seg)

    raw_spans: list[tuple[int, int, str, str]] = []  # (start, end, new_text, key)
    for old, new in replacements.items():
        if old.startswith('re:'):
            try:
                for m in re.finditer(old[3:], joined):
                    raw_spans.append((m.start(), m.end(), m.expand(new), old))
            except re.error:
                continue
        else:
            start = 0
            while True:
                k = joined.find(old, start)
                if k < 0:
                    break
                raw_spans.append((k, k + len(old), new, old))
                start = k + len(old)
    # 消解重叠：同起点取更长者；之后按起点顺序只接受不重叠区间
    raw_spans.sort(key=lambda x: (x[0], -(x[1] - x[0])))
    spans: list[tuple[int, int, str, str]] = []
    last_end = -1
    for a, b, new, key in raw_spans:
        if a >= last_end and b > a:
            spans.append((a, b, new, key))
            last_end = b
    hits = {k: 0 for k in replacements}
    for _, _, _, key in spans:
        hits[key] += 1
    if not spans:
        return data, hits

    def rebuild_node(idx: int) -> str:
        s, e = positions[idx]
        out: list[str] = []
        pos = s
        for a, b, new, _ in spans:
            if b <= s or a >= e:
                continue
            if a >= pos:
                out.append(joined[pos:a])
                out.append(new)
                pos = max(pos, b)
            else:  # 区间起点在更早的节点，本节点只负责删掉被覆盖的残余
                pos = max(pos, min(b, e))
        out.append(joined[min(pos, e):e])
        return ''.join(out)

    out: list[str] = []
    last = 0
    for i, m in enumerate(matches):
        out.append(text[last:m.start(2)])
        out.append(html.escape(rebuild_node(i), quote=False))
        out.append(text[m.end(2):m.end()])
        last = m.end()
    out.append(text[last:])
    return ''.join(out).encode('utf-8'), hits


def patch_text(src: Path, dst: Path, replacements: dict[str, str],
               only_parts: tuple[str, ...] | None = None) -> dict[str, int]:
    """复制 Office 包并对 XML 文本节点做精确替换，返回每个替换式的命中数。"""
    total: dict[str, int] = {k: 0 for k in replacements}
    dst.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(src) as zin, ZipFile(dst, 'w', ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            raw = zin.read(info.filename)
            if (info.filename.endswith(XML_PARTS)
                    and (only_parts is None or any(k in info.filename for k in only_parts))):
                patched, hits = _patch_text_nodes(raw, replacements)
                raw = patched
                for k, v in hits.items():
                    total[k] += v
            zout.writestr(info, raw)
    return total


def sheet_parts(src: Path) -> dict[str, str]:
    """返回 {sheet 名: xl/worksheets/sheetN.xml}。"""
    with ZipFile(src) as z:
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid_to_target = {}
    for rel in rels:
        rid_to_target[rel.attrib.get('Id', '')] = rel.attrib.get('Target', '')
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    rid_attr = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    out = {}
    for sh in wb.iter(ns + 'sheet'):
        target = rid_to_target.get(sh.attrib.get(rid_attr, ''), '')
        if target.startswith('/'):
            part = target.lstrip('/')
        elif target.startswith('xl/'):
            part = target
        else:
            part = 'xl/' + target.lstrip('/')
        out[sh.attrib.get('name', '')] = part
    return out


def _cell_xml(ref: str, attrs: str, value) -> str:
    style = re.search(r'\ss="(\d+)"', attrs)
    s = f' s="{style.group(1)}"' if style else ''
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}"{s}><v>{value}</v></c>'
    v = html.escape(str(value), quote=False)
    return f'<c r="{ref}"{s} t="inlineStr"><is><t>{v}</t></is></c>'


def patch_cells(sheet_xml: bytes, updates: dict[str, str]) -> tuple[bytes, list[str]]:
    """把指定单元格改为 inlineStr，保留样式属性；不动 sharedStrings、公式与合并区。"""
    text = sheet_xml.decode('utf-8')
    done: list[str] = []

    def repl(m: re.Match) -> str:
        ref = m.group('ref')
        if ref not in updates:
            return m.group(0)
        done.append(ref)
        return _cell_xml(ref, m.group('attrs'), updates[ref])

    return CELL_RE.sub(repl, text).encode('utf-8'), done


def add_green_fill(styles_xml: bytes) -> tuple[bytes, int | None]:
    """在 styles.xml 追加 RGB(226,239,218) 实心填充及对应 cellXf，返回新 xf 索引。"""
    text = styles_xml.decode('utf-8')
    m = re.search(r'<fills\s+count="(\d+)"[^>]*>(.*?)</fills>', text, re.S)
    if not m:
        return styles_xml, None
    # Some production workbooks carry an incorrect count attribute.  The style
    # index is the actual ordinal position, so derive it from the XML children.
    fill_idx = len(re.findall(r'<fill\b', m.group(2)))
    new_fill = ('<fill><patternFill patternType="solid">'
                '<fgColor rgb="FFE2EFDA"/><bgColor indexed="64"/>'
                '</patternFill></fill>')
    text = (text[:m.start()] + f'<fills count="{fill_idx + 1}">' + m.group(2) + new_fill + '</fills>'
            + text[m.end():])
    x = re.search(r'<cellXfs\s+count="(\d+)"[^>]*>(.*?)</cellXfs>', text, re.S)
    if not x:
        return styles_xml, None
    xf_idx = len(re.findall(r'<xf\b', x.group(2)))
    base = re.search(r'<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>', x.group(2), re.S)
    tpl = base.group(0) if base else '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
    # Appended rows mix sequence numbers, text dates and amounts.  A single
    # cloned style must not inherit the template row's date format or values
    # such as “75” render as 1900-03-15 in Excel.
    tpl = re.sub(r'\snumFmtId="\d+"', ' numFmtId="0"', tpl)
    if 'numFmtId=' not in tpl:
        tpl = tpl.replace('<xf ', '<xf numFmtId="0" ', 1)
    tpl = re.sub(r'\sfillId="\d+"', f' fillId="{fill_idx}"', tpl)
    tpl = tpl if 'fillId=' in tpl else tpl.replace('<xf ', f'<xf fillId="{fill_idx}" ', 1)
    tpl = re.sub(r'\sapplyFill="[^"]*"', '', tpl).replace('<xf ', '<xf applyFill="1" ', 1)
    text = (text[:x.start()] + f'<cellXfs count="{xf_idx + 1}">' + x.group(2) + tpl + '</cellXfs>'
            + text[x.end():])
    return text.encode('utf-8'), xf_idx


def append_rows(sheet_xml: bytes, template_row: int, rows: list[dict[str, str]],
                style_index: int | None = None) -> tuple[bytes, int]:
    """以 template_row 为样式模板追加数据行；单元格引用随行号平移，值按列字母覆盖。"""
    text = sheet_xml.decode('utf-8')
    row_matches = list(ROW_RE.finditer(text))
    all_rows = list(ROW_TAG_RE.finditer(text))

    def row_number(match: re.Match) -> int:
        found = re.search(r'\br="(\d+)"', match.group('attrs'))
        return int(found.group(1)) if found else -1

    tpl = next((m for m in row_matches if row_number(m) == template_row), None)
    if tpl is None or not rows:
        return sheet_xml, 0
    tpl_xml = tpl.group(0)
    # Template sheets often contain styled but self-closing empty rows after the
    # last data row.  Count those too or appended records reuse an existing r=
    # and become invisible/overwritten in Excel.
    last = max(row_number(m) for m in all_rows)
    added = []
    for i, values in enumerate(rows, start=1):
        new_r = last + i

        xml = re.sub(r'\br="([A-Z]+)\d+"', lambda m: f'r="{m.group(1)}{new_r}"', tpl_xml)
        xml = re.sub(r'(<row\s+r=")\d+(")', lambda m: m.group(1) + str(new_r) + m.group(2), xml)
        if style_index is not None:
            xml = re.sub(r'(<c\s+r="[A-Z]+\d+"[^>]*?)\ss="\d+"',
                         lambda m: m.group(1) + f' s="{style_index}"', xml)

        # 模板行中未在 values 指定的单元格必须清空，禁止把上期数值/公式带入本期新增行
        def blank(m: re.Match) -> str:
            tag = m.group(0)
            ref_m = re.search(r'\br="([A-Z]+)(\d+)"', tag)
            if not ref_m or ref_m.group(1) in values:
                return tag
            s_m = re.search(r'\ss="(\d+)"', tag)
            s_attr = f' s="{style_index}"' if style_index is not None else (
                f' s="{s_m.group(1)}"' if s_m else '')
            return f'<c r="{ref_m.group(1)}{new_r}"{s_attr}/>'

        xml = CELL_RE.sub(blank, xml)
        for col, val in values.items():
            ref = f'{col}{new_r}'
            pattern = re.compile(r'<c\s+r="' + re.escape(ref) + r'"[^>]*?(?:/>|>.*?</c>)', re.S)
            mm = pattern.search(xml)
            if mm:
                attrs = mm.group(0)
                if style_index is not None:
                    attrs = re.sub(r'\ss="\d+"', f' s="{style_index}"', attrs)
                    if ' s=' not in attrs:
                        attrs = attrs.replace(f'r="{ref}"', f'r="{ref}" s="{style_index}"', 1)
                xml = pattern.sub(lambda _m: _cell_xml(ref, attrs, val), xml, count=1)
        added.append(xml)
    insert_at = text.rfind('</sheetData>')
    if insert_at < 0:
        return sheet_xml, 0
    text = text[:insert_at] + ''.join(added) + text[insert_at:]
    text = DIMENSION_RE.sub(lambda m: m.group(1) + str(last + len(rows)) + m.group(3), text, count=1)
    return text.encode('utf-8'), len(added)


def patch_workbook(src: Path, dst: Path, replacements: dict[str, str],
                   sheet_updates: dict[str, dict[str, str]] | None = None,
                   sheet_appends: dict[str, tuple[int, list[dict[str, str]]]] | None = None,
                   green_new_rows: bool = True
                   ) -> dict:
    """一次性完成文本替换 + 指定 sheet 的单元格改写 / 追加行，其余部件原样复制。"""
    parts = sheet_parts(src)
    sheet_updates = sheet_updates or {}
    sheet_appends = sheet_appends or {}
    target_parts = {parts[k]: v for k, v in sheet_updates.items() if k in parts}
    append_parts = {parts[k]: v for k, v in sheet_appends.items() if k in parts}
    report: dict = {'text_hits': {}, 'cells': {}, 'appended': {}, 'unmatched_sheets': [],
                    'green_xf': None}
    for k in list(sheet_updates) + list(sheet_appends):
        if k not in parts:
            report['unmatched_sheets'].append(k)
    green_xf = None
    dst.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(src) as zin, ZipFile(dst, 'w', ZIP_DEFLATED) as zout:
        infos = zin.infolist()
        if green_new_rows and append_parts:
            for info in infos:
                if info.filename == 'xl/styles.xml':
                    _, green_xf = add_green_fill(zin.read(info.filename))
                    report['green_xf'] = green_xf
                    break
        for info in infos:
            raw = zin.read(info.filename)
            if info.filename.endswith(XML_PARTS):
                raw, hits = _patch_text_nodes(raw, replacements)
                for key, n in hits.items():
                    report['text_hits'][key] = report['text_hits'].get(key, 0) + n
            if info.filename == 'xl/styles.xml' and green_xf is not None:
                raw, _ = add_green_fill(raw)
            if info.filename in target_parts:
                raw, cells = patch_cells(raw, target_parts[info.filename])
                report['cells'][info.filename] = cells
            if info.filename in append_parts:
                tpl, rows = append_parts[info.filename]
                raw, n = append_rows(raw, tpl, rows, green_xf)
                report['appended'][info.filename] = n
            zout.writestr(info, raw)
    return report


def collect_text(src: Path, only_parts: tuple[str, ...] | None = None) -> str:
    """只读：拼接包内所有 XML 文本节点，供统计句式正则发现替换锚点。"""
    chunks: list[str] = []
    with ZipFile(src) as z:
        for name in z.namelist():
            if not name.endswith(XML_PARTS):
                continue
            if only_parts and not any(k in name for k in only_parts):
                continue
            try:
                text = z.read(name).decode('utf-8')
            except UnicodeDecodeError:
                continue
            chunks.extend(html.unescape(m.group(2)) for m in TEXT_NODE_RE.finditer(text))
    return ''.join(chunks)


def read_cell_values(src: Path, sheet: str, refs: list[str]) -> dict[str, str]:
    """只读：取指定单元格的显示文本（兼容 sharedStrings）。"""
    parts = sheet_parts(src)
    part = parts.get(sheet)
    if not part:
        return {}
    with ZipFile(src) as z:
        root = ET.fromstring(z.read(part))
        shared: list[str] = []
        if 'xl/sharedStrings.xml' in z.namelist():
            sst = ET.fromstring(z.read('xl/sharedStrings.xml'))
            ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
            shared = [''.join(t.text or '' for t in si.iter(ns + 't')) for si in sst.iter(ns + 'si')]
    ns = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
    out = {}
    for c in root.iter(ns + 'c'):
        ref = c.attrib.get('r')
        if ref not in refs:
            continue
        t = c.attrib.get('t')
        if t == 'inlineStr':
            out[ref] = ''.join(x.text or '' for x in c.iter(ns + 't'))
        else:
            v = c.find(ns + 'v')
            val = v.text if v is not None and v.text else ''
            out[ref] = shared[int(val)] if t == 's' and val.isdigit() and int(val) < len(shared) else val
    return out
