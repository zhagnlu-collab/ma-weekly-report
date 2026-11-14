"""Remove previous-week green fills without reserializing workbook packages."""
import colorsys
import re
from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def clear_green_fills(path):
    path = Path(path)
    with ZipFile(path) as z:
        raw = z.read('xl/styles.xml')
        styles = ET.fromstring(raw)
        palette = styles.find(NS + 'colors/' + NS + 'indexedColors')
        from openpyxl.styles.colors import COLOR_INDEX
        indexed = [x.get('rgb') for x in palette] if palette is not None else COLOR_INDEX
        themes = []
        if 'xl/theme/theme1.xml' in z.namelist():
            theme = ET.fromstring(z.read('xl/theme/theme1.xml'))
            scheme = theme.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme')
            if scheme is not None:
                for item in scheme:
                    c = list(item)[0]
                    themes.append(c.get('lastClr') or c.get('val'))
        def green(c):
            if c is None: return False
            rgb = c.get('rgb')
            if rgb is None and c.get('indexed') is not None:
                i = int(c.get('indexed')); rgb = indexed[i] if i < len(indexed) else None
            if rgb is None and c.get('theme') is not None:
                i = int(c.get('theme')); rgb = themes[i] if i < len(themes) else None
            if not rgb or not re.fullmatch('[0-9a-fA-F]{6,8}', rgb): return False
            vals = [int(rgb[-6:][i:i+2], 16)/255 for i in (0,2,4)]
            tint = float(c.get('tint', '0'))
            vals = [v*(1+tint) if tint < 0 else v*(1-tint)+tint for v in vals]
            h,s,v = colorsys.rgb_to_hsv(*vals)
            return 65 <= h*360 <= 170 and s > .035 and v > 0
        fills = styles.find(NS+'fills')
        ids = set()
        for i, fill in enumerate(fills if fills is not None else []):
            pat = fill.find(NS+'patternFill')
            if pat is not None and pat.get('patternType') not in (None, 'none'):
                colors = [pat.find(NS+'fgColor')]
                if pat.get('patternType') != 'solid': colors.append(pat.find(NS+'bgColor'))
                if any(green(c) for c in colors): ids.add(i)
        if not ids: return {'cleared_fill_definitions': 0}
        text = raw.decode('utf-8')
        match = re.search(r'<(?:\w+:)?fills\b[^>]*>.*?</(?:\w+:)?fills>', text, re.S)
        if not match: raise ValueError('Cannot locate fills XML')
        n = [-1]
        def replace(m):
            n[0] += 1
            if n[0] not in ids: return m.group(0)
            prefix = m.group(1) or ''
            return f'<{prefix}fill><{prefix}patternFill patternType="none"/></{prefix}fill>'
        block = re.sub(r'<((?:\w+:)?)fill\b[^>]*(?:/>|>.*?</(?:\w+:)?fill>)', replace, match.group(0), flags=re.S)
        updated = (text[:match.start()] + block + text[match.end():]).encode('utf-8')
        entries = [(info, updated if info.filename == 'xl/styles.xml' else z.read(info.filename)) for info in z.infolist()]
    tmp = path.with_name(path.name+'.clear-green.tmp')
    try:
        with ZipFile(tmp, 'w') as out:
            for info, data in entries: out.writestr(info, data)
        tmp.replace(path)
    finally:
        if tmp.exists(): tmp.unlink()
    return {'cleared_fill_definitions': len(ids)}
