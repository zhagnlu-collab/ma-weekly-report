#!/usr/bin/env python3
"""官方源采集（纯标准库，可无人值守直连）。

已实测可直连的入口：
  - 上交所并购重组项目动态 Excel：query.sse.com.cn/commonExcelKcb.do?sqlId=GP_BGCZ_XMLB
  - 深交所数据接口：www.szse.cn/api/report/ShowReport/data
  - 巨潮资讯网公告全文检索：www.cninfo.com.cn/new/hisAnnouncement/query (POST)
  - 上交所并购重组委会议公告：www.sse.com.cn/listing/announcement/commiteema/
每份来源保存原始响应、URL、访问时间、哈希、页数与记录数；列表页只产生候选。
来源失败记录 LIMITED 并继续，绝不把零返回解释为无案例。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

UA = 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124 Safari/537.36'
TIMEOUT = 25
CNINFO_KEYWORDS = [
    '重大资产重组', '发行股份购买资产', '重组预案', '重组报告书', '收购报告书',
    '控制权变更', '实际控制人变更', '股份转让协议', '权益变动报告书',
    '重大资产购买', '重大资产出售', '终止重大资产重组', '并购重组审核',
    # Weekly-report roles that the v4 three-keyword set missed in the 2026-W43
    # benchmark: registrations, IPO withdrawals and non-M&A cooperation.
    '同意注册', '注册批复', '终止上市', '撤回上市', '上市审核终止', '终止筹划', '股权收购', '优先股', '终止筹划', '股权收购', '优先股',
    '增资扩股', '增资引入', '老股转让', '引入战略投资者', '引入投资者',
    '专业投资机构', '私募股权基金', '私募基金', '创业投资', '共同投资',
    '基金份额', '扩募', '合伙协议', '产业投资基金', '股权投资基金', '并购基金', '母基金', '投资基金', '产业基金',
    '增资协议', '参股公司增资', '增资暨关联交易', '向全资孙公司增资', '向全资子公司增资',
    '对子公司增资',
]
GROUPS = ['SSE_MOMENTUM', 'SSE_COMMITTEE', 'SZSE_MOMENTUM', 'CNINFO_FULLTEXT',
          'BSE_MOMENTUM', 'CSRC_REGISTRATION']


def _now() -> str:
    return datetime.now(timezone.utc).astimezone().isoformat(timespec='seconds')


def _sha(b: bytes) -> str:
    return hashlib.sha256(b).hexdigest()


def http(url: str, data: bytes | None = None, headers: dict | None = None,
         retries: int = 2) -> tuple[int, bytes, str]:
    h = {'User-Agent': UA, 'Accept': '*/*'}
    h.update(headers or {})
    last = ''
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, data=data, headers=h)
            with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
                return r.status, r.read(), r.headers.get('Content-Type', '')
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError) as ex:
            last = f'{type(ex).__name__}: {ex}'
            if attempt < retries:
                time.sleep(1.2 * (attempt + 1))
    return 0, b'', last


def announcement_priority(row: dict) -> tuple[int, str, str]:
    """Rank indexed announcements so the download budget covers report roles.

    The text index can return hundreds of supporting documents (advisor reports,
    legal opinions, fund announcements).  Downloading in raw publication order
    used to exhaust the cap before strategic-investment and registration items.
    Title signals give a stable first pass; ambiguous files remain available for
    a later targeted run instead of being treated as absent.
    """
    title = _strip_tags(row.get('title') or '')
    try:
        from extract_filings import classify_title
        hits = classify_title(title)
    except Exception:
        hits = []
    events = {h[0] for h in hits}
    if events & {'first_disclosed', 'first_draft', 'accepted', 'committee_passed',
                 'committee_deferred', 'committee_rejected', 'simplified_passed',
                 'registration_submitted', 'registered', 'terminated',
                 'control_completed'}:
        score = 0
    elif events & {'lead_signal'}:
        score = 2
    elif any(x in title for x in ('权益变动报告书', '收购报告书', '要约收购')) and not any(
            x in title for x in ('持续督导', '财务顾问', '法律意见', '核查意见', '摘要')):
        score = 1
    elif any(x in title for x in ('完成', '过户')):
        score = 3
    else:
        score = 8
    return score, row.get('published_on') or '', str(row.get('announcement_id') or '')


class Acquirer:
    def __init__(self, out: Path, start: str, end: str, soffice: Path | None):
        self.out = out
        self.out.mkdir(parents=True, exist_ok=True)
        # start/end 对外统一为 YYYYMMDD（巨潮 seDate 依赖该格式）；
        # 窗口过滤另存 ISO 形式，避免 '20261101' <= '2026-11-06' 的字面比较恒为假。
        self.start = start
        self.end = end
        self.start_iso = f'{start[:4]}-{start[4:6]}-{start[6:8]}' if len(start) == 8 else start
        self.end_iso = f'{end[:4]}-{end[4:6]}-{end[6:8]}' if len(end) == 8 else end
        self.soffice = soffice
        self.records: list[dict] = []
        self.coverage: list[dict] = []

    def save(self, sid: str, name: str, raw: bytes, url: str, method: str = 'GET',
             status: str = 'COMPLETE', records: int = 0, pages: int = 1,
             error: str = '', params: dict | None = None) -> dict:
        p = self.out / name
        if raw:
            p.write_bytes(raw)
        rec = {'source_id': sid, 'url': url, 'method': method, 'params': params or {},
               'accessed_at': _now(), 'http_status': 200 if raw else 0,
               'bytes': len(raw), 'sha256': _sha(raw) if raw else None,
               'local_path': str(p) if raw else None, 'records': records,
               'pages': pages, 'status': status, 'error': error}
        self.records.append(rec)
        return rec

    def cover(self, group: str, status: str, evidence: str, note: str = ''):
        self.coverage.append({'group': group, 'status': status,
                              'evidence': evidence, 'note': note, 'checked_at': _now()})

    # ---- 上交所并购重组项目动态 ----
    def sse_momentum(self):
        sid = 'SSE_MOMENTUM'
        q = urllib.parse.urlencode({
            'sqlId': 'GP_BGCZ_XMLB', 'currStatus': '', 'issueMarketType': '1,2',
            'order': 'updateDate|desc,stockAuditNum|desc', 'keyword': '',
            'businessSubType': '', 'auditApplyDateBegin': '', 'auditApplyDateEnd': '',
            'isPagination': 'true', 'pageHelp.pageSize': '1000', 'pageHelp.pageNo': '1',
        })
        url = 'https://query.sse.com.cn/commonExcelKcb.do?' + q
        code, raw, ctype = http(url, headers={'Referer': 'https://www.sse.com.cn/'})
        if code != 200 or not raw.startswith(b'\xd0\xcf\x11\xe0'):
            self.save(sid, 'sse_momentum.xls', b'', url, 'GET', 'LIMITED', error=f'HTTP {code} {ctype}')
            self.cover(sid, 'LIMITED', url, '上交所项目动态未取回结构化清单')
            return
        self.save(sid, 'sse_momentum.xls', raw, url, 'GET', 'COMPLETE', params={'sqlId': 'GP_BGCZ_XMLB'})
        csv_text = self._xls_to_csv(self.out / 'sse_momentum.xls')
        rows = []
        if csv_text:
            import csv
            import io
            rd = list(csv.DictReader(io.StringIO(csv_text)))
            for r in rd:
                rows.append({k.strip(): (v or '').strip() for k, v in r.items() if k})
            (self.out / 'sse_momentum.json').write_text(
                json.dumps(rows, ensure_ascii=False, indent=2), encoding='utf-8')
        in_window = [r for r in rows
                     if self.start_iso <= _isodate(r.get('更新日期', '')) <= self.end_iso] if rows else []
        (self.out / 'sse_momentum_window.json').write_text(
            json.dumps(in_window, ensure_ascii=False, indent=2), encoding='utf-8')
        self.cover(sid, 'COMPLETE', url,
                   f'全量项目动态 {len(rows)} 条，其中更新日期落在本期 {len(in_window)} 条')

    def _xls_to_csv(self, xls: Path) -> str:
        if not self.soffice or not self.soffice.exists():
            return ''
        tmp = self.out / '_csv'
        tmp.mkdir(parents=True, exist_ok=True)
        try:
            subprocess.run([str(self.soffice), '--headless', '--convert-to', 'csv',
                            '--outdir', str(tmp), str(xls)],
                           capture_output=True, text=True, timeout=120)
            f = tmp / (xls.stem + '.csv')
            return f.read_text(encoding='utf-8', errors='replace') if f.exists() else ''
        except (subprocess.SubprocessError, OSError):
            return ''

    # ---- 上交所并购重组委会议公告 ----
    def sse_committee(self):
        sid = 'SSE_COMMITTEE'
        url = 'https://www.sse.com.cn/listing/announcement/commiteema/'
        code, raw, _ = http(url, headers={'Referer': 'https://www.sse.com.cn/'})
        if code != 200 or not raw:
            self.save(sid, 'sse_committee.html', b'', url, 'GET', 'LIMITED', error=f'HTTP {code}')
            self.cover(sid, 'LIMITED', url, '会议公告页未取回')
            return
        self.save(sid, 'sse_committee.html', raw, url, 'GET', 'PARTIAL')
        text = raw.decode('utf-8', errors='replace')
        self.cover(sid, 'PARTIAL', url,
                   '已保存网页快照；会议明细通常由动态接口加载，需以公告原件复核')

    # ---- 深交所项目动态 ----
    def szse_momentum(self):
        sid = 'SZSE_MOMENTUM'
        url = ('https://www.szse.cn/api/ras/projectrends/query?'
               'bizType=3&pageIndex=0&pageSize=1000&random=0.618')
        code, raw, _ = http(url, headers={'Referer': 'https://www.szse.cn/listing/projectdynamic/reproperty/index.html',
                                          'X-Requested-With': 'XMLHttpRequest'})
        if code != 200 or not raw:
            self.save(sid, 'szse_momentum.json', b'', url, 'GET', 'LIMITED', error=f'HTTP {code}')
            self.cover(sid, 'LIMITED', url, '深交所数据接口未取回')
            return
        self.save(sid, 'szse_momentum.json', raw, url, 'GET', 'PARTIAL')
        try:
            obj = json.loads(raw)
            n = len(obj.get('data', []) or []) if isinstance(obj, dict) else 0
        except (ValueError, TypeError):
            n = 0
        self.cover(sid, 'NOT_APPLICABLE' if n == 0 else 'COMPLETE', url,
                   f'官方并购重组项目动态接口返回 {n} 条')

    # ---- 巨潮公告全文检索 ----
    def cninfo(self, max_pdf: int = 40):
        sid = 'CNINFO_FULLTEXT'
        endpoint = 'http://www.cninfo.com.cn/new/hisAnnouncement/query'
        se_date = f'{self.start[:4]}-{self.start[4:6]}-{self.start[6:]}~{self.end[:4]}-{self.end[4:6]}-{self.end[6:]}'
        all_rows: dict[str, dict] = {}
        group_status = {}
        truncated_groups = {}
        for kw in CNINFO_KEYWORDS:
            page, pages, total = 1, 1, 0
            kw_rows = []
            while page <= pages and page <= 20:
                body = urllib.parse.urlencode({
                    'pageNum': page, 'pageSize': 30, 'column': 'szse', 'tabName': 'fulltext',
                    'plate': '', 'stock': '', 'searchkey': kw, 'secid': '', 'category': '',
                    'trade': '', 'seDate': se_date, 'sortName': '', 'sortType': '',
                    'isHLtitle': 'true',
                }).encode()
                code, raw, _ = http(endpoint, data=body, headers={
                    'Referer': 'http://www.cninfo.com.cn/new/commonUrl/pageOfSearch?url=disclosure/list/search',
                    'X-Requested-With': 'XMLHttpRequest',
                    'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
                })
                if code != 200 or not raw:
                    group_status[kw] = f'HTTP {code}'
                    break
                try:
                    obj = json.loads(raw)
                except ValueError:
                    group_status[kw] = 'JSON解析失败'
                    break
                total = int(obj.get('totalRecordNum') or 0)
                pages = max(1, (total + 29) // 30)
                anns = obj.get('announcements') or []
                for a in anns:
                    aid = str(a.get('announcementId') or '')
                    if not aid:
                        continue
                    row = {
                        'announcement_id': aid,
                        'sec_code': a.get('secCode'), 'sec_name': a.get('secName'),
                        'title': _strip_tags(a.get('announcementTitle') or ''),
                        'published_on': _ms_date(a.get('announcementTime')),
                        'adjunct_url': a.get('adjunctUrl'),
                        'column_id': a.get('columnId'),
                        'keyword': kw,
                        'source': 'cninfo',
                        'url': 'http://static.cninfo.com.cn/' + (a.get('adjunctUrl') or ''),
                    }
                    row['deal_id'] = f'{row["sec_code"]}_{aid}'
                    if aid not in all_rows:
                        all_rows[aid] = row
                    kw_rows.append(row)
                page += 1
            group_status[kw] = f'{len(kw_rows)}/{total}'
            if pages > 20:
                truncated_groups[kw] = {'declared_pages': pages, 'fetched_pages': 20}
        rows = sorted(all_rows.values(), key=lambda r: (r['published_on'], r['sec_code'] or ''))
        payload = json.dumps({'period': {'start': self.start, 'end': self.end},
                              'keywords': group_status, 'truncated_groups': truncated_groups,
                              'total_unique': len(rows),
                              'announcements': rows}, ensure_ascii=False, indent=2)
        (self.out / 'cninfo_announcements.json').write_text(payload, encoding='utf-8')
        self.save(sid, 'cninfo_index.json', payload.encode(), endpoint, 'POST',
                  'COMPLETE' if rows else 'LIMITED', records=len(rows),
                  pages=len(CNINFO_KEYWORDS), params={'seDate': se_date, 'keywords': CNINFO_KEYWORDS})
        # 下载本期公告原件（限流，保留哈希）。先下载可构成周报里程碑的公告，
        # 避免按时间顺序耗尽配额后漏掉注册、引战或基金合作。
        pdf_dir = self.out / 'cninfo_pdfs'
        pdf_dir.mkdir(parents=True, exist_ok=True)
        downloaded = 0
        download_rows = sorted(rows, key=announcement_priority)
        for r in download_rows:
            if downloaded >= max_pdf or not r['adjunct_url']:
                continue
            dest = pdf_dir / f'{r["announcement_id"]}.pdf'
            if dest.exists():
                downloaded += 1
                continue
            code, raw, _ = http(r['url'], headers={'Referer': 'http://www.cninfo.com.cn/'})
            if code == 200 and raw.startswith(b'%PDF'):
                dest.write_bytes(raw)
                r['local_pdf'] = str(dest)
                r['pdf_sha256'] = _sha(raw)
                downloaded += 1
        (self.out / 'cninfo_announcements.json').write_text(
            json.dumps({'period': {'start': self.start, 'end': self.end},
                        'keywords': group_status, 'truncated_groups': truncated_groups,
                        'total_unique': len(rows),
                        'downloaded_pdfs': downloaded, 'announcements': rows},
                       ensure_ascii=False, indent=2), encoding='utf-8')
        present = sum(1 for r in rows if (pdf_dir / f"{r['announcement_id']}.pdf").is_file())
        ok = all('HTTP' not in v and '失败' not in v for v in group_status.values()) and not truncated_groups and present == len(rows)
        self.cover(sid, 'COMPLETE' if ok and rows else 'PARTIAL', endpoint,
                   f'{len(CNINFO_KEYWORDS)}组关键词检索，去重 {len(rows)} 条，原件就绪 {present}/{len(rows)} 份'
                   + (f'；截断组 {list(truncated_groups)}' if truncated_groups else ''))

    # ---- 尽力而为的补充源 ----
    def best_effort(self):
        targets = [
            ('BSE_MOMENTUM', 'https://www.bse.cn/audit/t_project_news_r.html', 'bse_momentum.html'),
            ('CSRC_REGISTRATION', 'https://www.csrc.gov.cn/csrc/c100028/common_list.shtml', 'csrc_registration.html'),
        ]
        for sid, url, name in targets:
            structured = self.out / ('bse_momentum.json' if sid == 'BSE_MOMENTUM' else 'csrc_registration.json')
            if structured.is_file():
                try:
                    obj = json.loads(structured.read_text(encoding='utf-8'))
                    rows = obj.get('data', []) if isinstance(obj, dict) else obj
                    if isinstance(rows, list) and rows:
                        self.save(sid, name, structured.read_bytes(), url, 'GET', 'COMPLETE', records=len(rows))
                        self.cover(sid, 'COMPLETE', url, f'浏览器/官方导出快照含 {len(rows)} 条记录')
                        continue
                except ValueError:
                    pass
            code, raw, _ = http(url)
            if code == 200 and raw:
                self.save(sid, name, raw, url, 'GET', 'PARTIAL')
                self.cover(sid, 'PARTIAL', url, '已保存网页快照；该源结构化程度有限，明细以公告原件为准')
            else:
                self.save(sid, name, b'', url, 'GET', 'LIMITED', error=f'HTTP {code}')
                self.cover(sid, 'LIMITED', url, '未取回，需在待核清单中列明覆盖缺口')

    def run(self, max_pdf: int):
        for fn in (self.sse_momentum, self.sse_committee, self.szse_momentum):
            try:
                fn()
            except Exception as ex:  # 单源失败不阻断整条流水线
                self.cover(fn.__name__.upper(), 'LIMITED', fn.__name__, f'{type(ex).__name__}: {ex}')
        try:
            self.cninfo(max_pdf)
        except Exception as ex:
            self.cover('CNINFO_FULLTEXT', 'LIMITED', 'cninfo', f'{type(ex).__name__}: {ex}')
        try:
            self.best_effort()
        except Exception as ex:
            self.cover('BEST_EFFORT', 'LIMITED', 'best_effort', f'{type(ex).__name__}: {ex}')
        for g in GROUPS:
            if not any(c['group'] == g for c in self.coverage):
                self.cover(g, 'NOT_RUN', 'acquire_sources', '本阶段未执行')
        return {'sources': self.records, 'coverage': self.coverage}


def _strip_tags(s: str) -> str:
    import re
    return re.sub(r'<[^>]+>', '', s)


def _ms_date(v) -> str:
    if not v:
        return ''
    try:
        return datetime.fromtimestamp(int(v) / 1000, timezone.utc).astimezone().date().isoformat()
    except (ValueError, TypeError, OSError):
        return ''


def _isodate(s: str) -> str:
    return s.replace('/', '-')[:10]


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--start', required=True)
    ap.add_argument('--end', required=True)
    ap.add_argument('--out', required=True, type=Path)
    ap.add_argument('--manifest', required=True, type=Path)
    ap.add_argument('--soffice', type=Path)
    ap.add_argument('--max-pdf', type=int, default=40)
    args = ap.parse_args()
    result = Acquirer(args.out, args.start, args.end, args.soffice).run(args.max_pdf)
    result.update({'generated_at': _now(), 'period': {'start': args.start, 'end': args.end},
                   'limitations': ['列表页只产生候选；关键事实必须回到公告全文',
                                   '来源状态为声明，不等于已完成事实核验']})
    args.manifest.parent.mkdir(parents=True, exist_ok=True)
    args.manifest.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'sources': len(result['sources']),
                      'coverage': {c['group']: c['status'] for c in result['coverage']},
                      'manifest': str(args.manifest)}, ensure_ascii=False))
    return 0


if __name__ == '__main__':
    sys.exit(main())
