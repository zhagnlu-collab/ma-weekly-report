#!/usr/bin/env python3
"""Generate clean-native Word/PowerPoint deliverables from a verified ledger.

This is the v4.1 fallback when a prior role-specific Word/PPT template cannot be
safely hydrated.  It deliberately favours truthful, reviewable content over
pretending that a newly generated layout is equivalent to the human template.
Excel workbooks remain template-derived so formulas, charts and historical sheets
are preserved.  The open WARN items in test_known_answer_43.py track the remaining
layout-fidelity gap.
"""
from __future__ import annotations
import argparse,json,re,sys
from datetime import date
from pathlib import Path

def ensure_libs():
    """只在**同版本**解释器的 site-packages 中补依赖，禁止跨 minor 版本注入。

    2026-11-08 第42期盲测发现的真实缺陷：本函数曾无条件把 Python 3.12 的
    site-packages 插进 Python 3.13 进程；python-pptx 导入 PIL.Image 时加载
    3.12 编译的 _imaging C 扩展，ABI 不兼容直接 ImportError，BUILD 阶段被
    硬门禁阻断。C 扩展绝不能跨解释器版本混用。
    """
    try:
        import docx, pptx  # noqa: F401
        return
    except ImportError:
        pass
    tag=f'python{sys.version_info.major}.{sys.version_info.minor}'
    runtime=Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies/python/lib'/tag/'site-packages'
    if (runtime/'pptx').is_dir() and (runtime/'docx').is_dir():
        sys.path.insert(0,str(runtime))
        return
    raise RuntimeError(
        f'当前解释器 Python {sys.version_info.major}.{sys.version_info.minor} 缺少 '
        'python-docx / python-pptx；拒绝注入其它 Python 版本的 site-packages'
        '（C 扩展 ABI 不兼容）。请通过 auto_run.py 启动（它会自动选择合格解释器），'
        '或执行 pip install python-docx python-pptx。')

ensure_libs()
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt, RGBColor
from pptx import Presentation
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches as PInches, Pt as PPt

NAVY='17365D'; BLUE='2F75B5'; LIGHT='DCE6F1'; PALE='EAF2F8'; ORANGE='C65911'; GREY='666666'; RED='C00000'; WHITE='FFFFFF'

def load(p): return json.loads(Path(p).read_text(encoding='utf-8'))
def cn_date(s):
    if not s:return ''
    d=date.fromisoformat(s);return f'{d.year}年{d.month}月{d.day}日'

def source_map(ledger):return {s['source_id']:s for s in ledger.get('sources',[])}
def case_map(ledger):return {c['deal_id']:c for c in ledger.get('cases',[])}

def event_rows(ledger,kind=None,types=None,journal=False):
    cases=case_map(ledger);srcs=source_map(ledger);out=[]
    for e in ledger.get('events',[]):
        c=cases.get(e['deal_id'],{})
        if e.get('status')!='CONFIRMED':continue
        if kind and c.get('kind') not in (kind if isinstance(kind,(set,list,tuple)) else {kind}):continue
        if types and e.get('type') not in types:continue
        s=srcs.get((e.get('source_ids') or [''])[0],{})
        out.append({'event':e,'case':c,'source':s})
    return sorted(out,key=lambda x:(x['event']['date'],x['case'].get('name','')))

def unique_event_rows(rows):
    best={}
    rank={'first_disclosed':10,'first_draft':20,'accepted':30,'inquiry':40,'replied':41,
          'suspended':45,'resumed':46,'committee_scheduled':50,'committee_passed':60,
          'simplified_passed':61,'registration_submitted':70,'registered':80,
          'terminated':90,'completed':95,'control_completed':96}
    for row in rows:
        cid=row['event']['deal_id']; score=(row['event'].get('event_date') or row['event'].get('date') or '',
             row['event'].get('announced_on') or '', rank.get(row['event'].get('type'),0))
        old=best.get(cid); old_score=(old['event'].get('event_date') or old['event'].get('date') or '',
             old['event'].get('announced_on') or '', rank.get(old['event'].get('type'),0)) if old else ('','',-1)
        if score>=old_score: best[cid]=row
    return sorted(best.values(),key=lambda x:(x['event'].get('date',''),x['case'].get('name','')))

def deal_union(*groups, dedupe_code=False):
    rows={}
    for group in groups:
        for x in group:
            key=x['case'].get('code') if dedupe_code else x['event']['deal_id']
            rows.setdefault(key,x)
    return sorted(rows.values(),key=lambda x:(x['event'].get('date',''),x['case'].get('name','')))

def safe_names(rows):
    return '、'.join(dict.fromkeys(x['case'].get('name','') for x in rows)) or '无'

def coverage_summary(ledger):
    gaps=[f"{x.get('group')}={x.get('status')}" for x in ledger.get('coverage',[]) if x.get('status') not in ('COMPLETE','NOT_APPLICABLE')]
    return '；'.join(gaps) if gaps else '核心来源均已完成'

def amounts(case,actual=False):
    vals=[]
    for a in case.get('amounts',[]):
        try:v=float(a.get('value_wan'))
        except (TypeError,ValueError):continue
        if a.get('currency','CNY')!='CNY':continue
        if actual and a.get('basis')!='actual_completion':continue
        if not actual and a.get('basis')=='actual_completion':continue
        vals.append((a.get('known_on',''),v))
    return sorted(vals)[-1][1] if vals else None

def amount_text(v):
    if v is None:return '未披露'
    if abs(v)>=10000:return f'{v/10000:.2f}亿元'
    return f'{v:,.2f}万元'

def wan_text(v):
    return '未披露' if v is None else f'{v:,.2f}'

def advisor_text(case):
    xs=[a.get('name','') for a in case.get('advisors',[]) if a.get('role')=='independent']
    return '、'.join(dict.fromkeys(xs)) or '未披露'

def locator(src,event):
    x=src.get('locator') or event.get('note','')
    return re.sub(r'^巨潮公告｜[^｜]+｜','',x)

def set_cell_shading(cell,color):
    tcPr=cell._tc.get_or_add_tcPr();shd=OxmlElement('w:shd')
    shd.set(qn('w:fill'),color);tcPr.append(shd)

def set_repeat_table_header(row):
    trPr=row._tr.get_or_add_trPr();el=OxmlElement('w:tblHeader');el.set(qn('w:val'),'true');trPr.append(el)

def set_run_font(run,size=10.5,bold=False,color=None,name='Arial Unicode MS'):
    run.font.name=name;run.font.size=Pt(size);run.font.bold=bold
    if color:run.font.color.rgb=RGBColor.from_string(color)
    run._element.rPr.rFonts.set(qn('w:eastAsia'),name)

def add_para(doc,text='',size=10.5,bold=False,color=None,align=None,space_after=4):
    p=doc.add_paragraph();r=p.add_run(text);set_run_font(r,size,bold,color)
    p.paragraph_format.space_after=Pt(space_after);p.paragraph_format.line_spacing=1.2
    if align is not None:p.alignment=align
    return p

def add_heading(doc,text,level=1):
    p=doc.add_paragraph();r=p.add_run(text);set_run_font(r,15 if level==1 else 12.5,True,NAVY)
    p.paragraph_format.space_before=Pt(8);p.paragraph_format.space_after=Pt(4)
    if level==1:
        pPr=p._p.get_or_add_pPr();pbdr=OxmlElement('w:pBdr');bottom=OxmlElement('w:bottom')
        bottom.set(qn('w:val'),'single');bottom.set(qn('w:sz'),'10');bottom.set(qn('w:space'),'2');bottom.set(qn('w:color'),BLUE);pbdr.append(bottom);pPr.append(pbdr)
    return p

def add_table(doc,headers,rows,widths=None):
    table=doc.add_table(rows=1,cols=len(headers));table.alignment=WD_TABLE_ALIGNMENT.CENTER;table.style='Table Grid'
    hdr=table.rows[0];
    for i,h in enumerate(headers):
        c=hdr.cells[i];c.text='';set_cell_shading(c,BLUE);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        p=c.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=p.add_run(h);set_run_font(r,8.5,True,WHITE)
    set_repeat_table_header(hdr)
    for ri,data in enumerate(rows):
        cells=table.add_row().cells
        for i,val in enumerate(data):
            c=cells[i];c.text='';c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if ri%2:set_cell_shading(c,PALE)
            p=c.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.CENTER if i not in (1,3) else WD_ALIGN_PARAGRAPH.LEFT
            r=p.add_run(str(val));set_run_font(r,8.2)
    if widths:
        for row in table.rows:
            for i,w in enumerate(widths):row.cells[i].width=Cm(w)
    doc.add_paragraph().paragraph_format.space_after=Pt(0)
    return table

def setup_doc():
    d=Document();sec=d.sections[0];sec.top_margin=Cm(1.55);sec.bottom_margin=Cm(1.55);sec.left_margin=Cm(1.7);sec.right_margin=Cm(1.7)
    st=d.styles['Normal'];st.font.name='Arial Unicode MS';st.font.size=Pt(10.5);st._element.rPr.rFonts.set(qn('w:eastAsia'),'Arial Unicode MS')
    d.core_properties.author='HTLH';d.core_properties.last_modified_by='HTLH';d.core_properties.comments='基于官方公告原件与已核验事件台账生成'
    return d

def add_footer(doc,text):
    for sec in doc.sections:
        p=sec.footer.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=p.add_run(text);set_run_font(r,8,False,GREY)

def role_file(folder: Path, number: str):
    return next((p for p in folder.glob(f'{number}、*.xlsx') if not p.name.startswith('~$')), None)

def sheet_code_rows(path,sheet):
    if not path or not path.is_file():return []
    import openpyxl,re
    wb=openpyxl.load_workbook(path,read_only=True,data_only=False)
    try:
        if sheet not in wb.sheetnames:return []
        ws=wb[sheet];out=[]
        for row in ws.iter_rows(min_row=1,max_row=min(ws.max_row,5000),min_col=1,max_col=min(ws.max_column,100),values_only=True):
            code=''
            for v in row:
                if isinstance(v,str):
                    m=re.fullmatch(r'\s*(\d{6})(?:\.(?:SH|SZ|BJ))?\s*',v,re.I)
                    if m:code=m.group(1);break
            if code:out.append((code,row))
        return out
    finally:wb.close()

def period_row_count(path, sheet, folder: Path):
    if not path or not path.is_file():return 0
    import openpyxl,re
    from datetime import date,datetime
    m=re.search(r'(\d{8})至(\d{8})',folder.name)
    if not m:return 0
    start=datetime.strptime(m.group(1),'%Y%m%d').date();end=datetime.strptime(m.group(2),'%Y%m%d').date()
    wb=openpyxl.load_workbook(path,read_only=True,data_only=False)
    try:
        if sheet not in wb.sheetnames:return 0
        ws=wb[sheet];date_col=None;total=0
        for r in range(1,min(8,ws.max_row)+1):
            for i,v in enumerate(next(ws.iter_rows(min_row=r,max_row=r,min_col=1,max_col=min(ws.max_column,30),values_only=True))):
                if isinstance(v,str) and ('首次披露日期' in v or '提示性公告披露日期' in v):date_col=i;break
            if date_col is not None:break
        if date_col is None:return 0
        for row in ws.iter_rows(min_row=2,max_row=ws.max_row,min_col=1,max_col=min(ws.max_column,50),values_only=True):
            v=row[date_col] if date_col<len(row) else None
            d=None
            if isinstance(v,datetime):d=v.date()
            elif isinstance(v,date):d=v
            elif isinstance(v,str):
                t=v.strip().replace('/','-')[:10]
                try:d=date.fromisoformat(t if len(t)==10 else f'{t[:4]}-{t[4:6]}-{t[6:8]}')
                except ValueError:pass
            if d and start<=d<=end:total+=1
        return total
    finally:wb.close()

def workbook_facts(folder: Path):
    folder=Path(folder);zero=role_file(folder,'0');month=role_file(folder,'4');control=role_file(folder,'7');coop=role_file(folder,'11');partial=role_file(folder,'12')
    def unique(sheet,p=zero):return len({x[0] for x in sheet_code_rows(p,sheet)})
    in_review=0
    import openpyxl
    if zero and zero.is_file():
        wb=openpyxl.load_workbook(zero,read_only=True,data_only=False)
        try:
            for sheet in ('上交所在审','深交所在审','北交所在审'):
                if sheet not in wb.sheetnames:continue
                ws=wb[sheet];header=None
                for r in range(1,min(8,ws.max_row)+1):
                    for i,v in enumerate(next(ws.iter_rows(min_row=r,max_row=r,min_col=1,max_col=min(ws.max_column,40),values_only=True))):
                        if isinstance(v,str) and '是否在审' in v:header=i;break
                    if header is not None:break
                rows=sheet_code_rows(zero,sheet)
                in_review+=len({code for code,row in rows if header is not None and header<len(row) and str(row[header]).strip()=='是'})
        finally:wb.close()
    return {'first_disclosed':unique('沪深北交易所首次披露'),'passed':unique('重组委通过'),
            'registered':unique('证监会注册'),'terminated':unique('撤回终止'),
            'cash_purchase':unique('现金重大购买'),'cash_sale':unique('现金重大出售'),
            'monthly_reviews':len({x[0] for x in sheet_code_rows(month,'重组委审核原始数据')}),
            'control_first':unique('2026年首次披露',control),'control_completed':unique('2026年已完成',control),
            'strategic':unique('上市公司子公司引战情况',coop),'professional':unique('上市公司与专业机构合作情况',coop),
            'current_strategic':period_row_count(coop,'上市公司子公司引战情况',folder),
            'current_professional':period_row_count(coop,'上市公司与专业机构合作情况',folder),
            'partial_transfer':unique('2026年首次披露（2026.9.13起统计）',partial),
            'current_partial_transfer':period_row_count(partial,'2026年首次披露（2026.9.13起统计）',folder),
            'in_review':in_review}

def build_weekly(ledger,metrics,period,out,prev_dir=None):
    d=setup_doc();start=period['start'];end=period['end'];issue=period['issue_no'];facts=workbook_facts(Path(out).parent);prev_facts=workbook_facts(Path(prev_dir)) if prev_dir else {}
    p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=p.add_run('并购重组周报');set_run_font(r,24,True,NAVY)
    add_para(d,f'{start[:4]}年第{issue}期｜{cn_date(start[:4]+"-"+start[4:6]+"-"+start[6:])}至{cn_date(end[:4]+"-"+end[4:6]+"-"+end[6:])}',11,False,GREY,WD_ALIGN_PARAGRAPH.CENTER,12)
    audit_first=unique_event_rows(event_rows(ledger,'ma_audit',{'first_disclosed','first_draft'}))
    cash=unique_event_rows(event_rows(ledger,{'ma_cash_purchase','ma_cash_sale'},{'first_disclosed','first_draft'}))
    # A combined cash purchase/sale is one transaction even when two role cases are retained for the two worksheets.
    first_all={x['case'].get('code') or x['event']['deal_id']:x for x in audit_first+cash}
    passed=unique_event_rows(event_rows(ledger,'ma_audit',{'committee_passed','simplified_passed'}))
    accepted=unique_event_rows(event_rows(ledger,'ma_audit',{'accepted'}))
    registered=unique_event_rows(event_rows(ledger,'ma_audit',{'registered'}))
    terms=unique_event_rows(event_rows(ledger,'ma_audit',{'terminated'}))
    control=unique_event_rows(event_rows(ledger,'control',{'first_disclosed'}))
    completed=unique_event_rows(event_rows(ledger,'control',{'control_completed'}))
    partial=unique_event_rows(event_rows(ledger,'partial_transfer',{'first_disclosed'}))
    strategic=unique_event_rows(event_rows(ledger,{'strategic_investment','professional_cooperation'},{'first_disclosed'}))
    leads=[]
    cases={c['deal_id']:c for c in ledger.get('cases',[])}
    for e in ledger.get('events',[]):
        if e.get('type')=='lead_signal':
            src=source_map(ledger).get((e.get('source_ids') or [''])[0],{})
            leads.append({'event':e,'case':cases.get(e['deal_id'],{}),'source':src})
    add_heading(d,'一、核心结论',1)
    add_para(d,f'本期确认预案/草案首次披露 {len(first_all)} 单，重组委审议通过 {len(passed)} 单，新增受理 {len(accepted)} 单，新增注册 {len(registered)} 单，终止 {len(terms)} 单。控制权首次披露 {len(control)} 单、控制权完成 {len(completed)} 单。',11,True,NAVY)
    add_para(d,f'统计口径区分“预案/报告书（草案）首次披露”和“筹划/停牌线索”。本期另有 {len(leads)} 条筹划或停牌线索，仅作后续跟踪，不计入上述首次披露项目。')
    review_text=f"工作簿状态勾稽显示在审 {facts.get('in_review','待核')} 单"
    if prev_facts:review_text+=f"（上期 {prev_facts.get('in_review','待核')} 单）"
    add_para(d,f'来源覆盖：{coverage_summary(ledger)}。{review_text}。')
    add_heading(d,'二、并购重组委审核',1)
    if passed:
        add_table(d,['会议日期','上市公司','审议结果','独立财务顾问','交易方式'],[[cn_date(x['event'].get('date')),x['case']['name'],'通过',advisor_text(x['case']),x['case'].get('facts',{}).get('payment_method','待核')] for x in passed],[2.5,3,1.8,4,4.5])
    else:add_para(d,'本期未识别到已完成公告核验的重组委通过项目。')
    add_para(d,f"本期会议数 {metrics.get('weekly',{}).get('committee_meetings',0)} 次；年度重组委已审核 {facts.get('monthly_reviews','待核')} 单；会议、公告日期分列。")
    add_heading(d,'三、预案或草案首次披露',1)
    rows=[]
    for x in deal_union(audit_first,cash,dedupe_code=True):
        c=x['case'];f=c.get('facts',{});stage='草案' if x['event'].get('type')=='first_draft' else '预案'
        rows.append([cn_date(x['event'].get('date')),c['name'],f.get('target','待核'),f.get('payment_method','待核'),wan_text(amounts(c)),stage,advisor_text(c),locator(x['source'],x['event'])])
    add_table(d,['日期','上市公司','交易标的','支付方式','交易金额（万元）','阶段','独立财务顾问','公告标题'],rows,[2.0,2.4,4.5,2.6,2.2,1.3,2.7,5.9])
    add_heading(d,'四、受理、注册与终止',1)
    add_para(d,f'新增受理：{safe_names(accepted)}。新增注册：{safe_names(registered)}。终止：{safe_names(terms)}。')
    add_para(d,'终止事实与“在审项目终止”分开判断；仅凭预案阶段终止公告，不从在审存量中扣减。')
    add_heading(d,'五、控制权、协议转让与引战合作',1)
    add_para(d,f'控制权首次披露：{safe_names(control)}；控制权完成：{safe_names(completed)}；协议转让：{safe_names(partial)}；子公司引战或专业合作：{safe_names(strategic)}。')
    add_heading(d,'六、筹划/停牌线索',1)
    if leads:
        add_table(d,['日期','公司','类型','公告标题'],[[cn_date(x['event'].get('date')),x['case'].get('name','待核'),x['case'].get('kind','待核'),locator(x['source'],x['event'])] for x in leads],[2.3,3.5,4,8])
    else:add_para(d,'本期无新增筹划/停牌线索。')
    add_heading(d,'七、编制口径与限制',1)
    for item in ['首次披露以预案或报告书（草案）为准；仅筹划、停牌或框架意向的公告不计入。','会议日、公告日、协议签署日、完成日分列，不以公告替代事件日。','金额、比例、控制权及顾问字段只在公告正文或结构化审核动态明确时填写。','来源覆盖与待核事项以随附《本期编制与待核说明》为准。']:
        p=d.add_paragraph(style='List Bullet');r=p.add_run(item);set_run_font(r,10)
    add_heading(d,'附录：主要证据索引',1)
    for x in deal_union(audit_first,cash,passed,terms,control,completed,partial,strategic,leads):
        src=x['source'];add_para(d,f"• {x['case'].get('name','')}｜{cn_date(x['event'].get('date'))}｜{locator(src,x['event'])}｜{src.get('file_or_url','')}",8.5,False,GREY)
    add_footer(d,f'HTLH｜并购重组周报｜{start[:4]}年第{issue}期');d.save(out)

def build_dynamic(ledger,metrics,period,out,prev_dir=None):
    d=setup_doc();start=period['start'];end=period['end'];issue=period['issue_no'];facts=workbook_facts(Path(out).parent);prev_facts=workbook_facts(Path(prev_dir)) if prev_dir else {}
    p=d.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=p.add_run('A股市场并购重组动态');set_run_font(r,24,True,NAVY)
    add_para(d,f'{cn_date(start[:4]+"-"+start[4:6]+"-"+start[6:])}至{cn_date(end[:4]+"-"+end[4:6]+"-"+end[6:])}',11,False,GREY,WD_ALIGN_PARAGRAPH.CENTER,12)
    audit=unique_event_rows(event_rows(ledger,'ma_audit',{'first_disclosed','first_draft'}));cash=unique_event_rows(event_rows(ledger,{'ma_cash_purchase','ma_cash_sale'},{'first_disclosed','first_draft'}));passed=unique_event_rows(event_rows(ledger,'ma_audit',{'committee_passed','simplified_passed'}));control=unique_event_rows(event_rows(ledger,'control',{'first_disclosed'}));completed=unique_event_rows(event_rows(ledger,'control',{'control_completed'}));partial=unique_event_rows(event_rows(ledger,'partial_transfer',{'first_disclosed'}));strategic=unique_event_rows(event_rows(ledger,{'strategic_investment','professional_cooperation'},{'first_disclosed'}));registered=unique_event_rows(event_rows(ledger,'ma_audit',{'registered'}));terms=unique_event_rows(event_rows(ledger,'ma_audit',{'terminated'}));accepted=unique_event_rows(event_rows(ledger,'ma_audit',{'accepted'}))
    cases={c['deal_id']:c for c in ledger.get('cases',[])};sources=source_map(ledger);leads=[{'event':e,'case':cases.get(e['deal_id'],{}),'source':sources.get((e.get('source_ids') or [''])[0],{})} for e in ledger.get('events',[]) if e.get('type')=='lead_signal']
    first_all={x['case'].get('code') or x['event']['deal_id']:x for x in audit+cash}
    add_heading(d,'市场概览',1)
    add_para(d,f'本期识别预案/草案首次披露 {len(first_all)} 单、重组委通过 {len(passed)} 单、新增受理 {len(accepted)} 单、新增注册 {len(registered)} 单、终止 {len(terms)} 单、控制权首次披露 {len(control)} 单、控制权完成 {len(completed)} 单、协议转让 {len(partial)} 单、子公司引战或专业合作 {len(strategic)} 单。另有 {len(leads)} 条筹划/停牌线索，不计入首次披露。')
    add_para(d,f"截至期末，基础工作簿累计口径为：审核类首次披露 {facts.get('first_disclosed','待核')} 单、重组委通过 {facts.get('passed','待核')} 单、注册 {facts.get('registered','待核')} 单、在审 {facts.get('in_review','待核')} 单。"+ (f"上期在审 {prev_facts.get('in_review','待核')} 单。" if prev_facts else ''))
    add_heading(d,'一、审核类交易与注册终止',1)
    if passed:
        add_table(d,['会议日期','公司','结果','独立财务顾问'],[[cn_date(x['event'].get('date')),x['case']['name'],'通过',advisor_text(x['case'])] for x in passed],[2.5,3,2,5])
    add_para(d,f'预案或草案首次披露：{safe_names(deal_union(audit,cash))}。新增受理：{safe_names(accepted)}。新增注册：{safe_names(registered)}。终止：{safe_names(terms)}。')
    add_heading(d,'二、控制权变更',1)
    if control:
        add_table(d,['日期','上市公司','原实际控制人','新实际控制人','控制股比','公告标题'],[[cn_date(x['event'].get('date')),x['case']['name'],x['case'].get('facts',{}).get('old_controller','待核'),x['case'].get('facts',{}).get('new_controller','待核'),f"{x['case'].get('facts',{}).get('control_ratio_pct','')}{'%' if x['case'].get('facts',{}).get('control_ratio_pct') else ''}",locator(x['source'],x['event'])] for x in control],[2.3,2.5,3.4,3.2,2.0,6.5])
    else:add_para(d,'本期未识别到已完成正文核验的实际控制人首次变更项目。')
    add_para(d,f'控制权完成：{safe_names(completed)}。')
    add_heading(d,'三、协议转让、现金重大交易与引战合作',1)
    for x in deal_union(partial,cash,strategic,dedupe_code=True):
        c=x['case'];add_para(d,f"• {c['name']}（{cn_date(x['event'].get('date'))}，{c.get('kind','')}）：{locator(x['source'],x['event'])}",9.5)
    if not deal_union(partial,cash,strategic,dedupe_code=True):add_para(d,'本期未识别到符合口径且已完成正文核验的新增项目。')
    add_heading(d,'四、筹划/停牌线索',1)
    for x in leads:add_para(d,f"• {x['case'].get('name','待核')}：{locator(x['source'],x['event'])}（仅作跟踪，不计入正式统计）",9.5)
    if not leads:add_para(d,'本期无新增筹划/停牌线索。')
    add_heading(d,'五、覆盖与待核',1)
    for item in [f'来源覆盖：{coverage_summary(ledger)}。','权益变动/收购报告书只有在正文明确实际控制人发生变化时进入控制权表；实控人未变但有量化协议转让的项目另列。','同一混合现金购买/出售交易在采购、出售两张专表可分别落表，但市场项目数按证券代码去重。','本稿为独立重建稿，正式发布前以《本期编制与待核说明》处理全部覆盖和字段缺口。']:
        p=d.add_paragraph(style='List Bullet');r=p.add_run(item);set_run_font(r,10)
    add_footer(d,f'HTLH｜A股市场并购重组动态｜{start[:4]}年第{issue}期');d.save(out)

def add_slide_title(prs,title,subtitle=''):
    s=prs.slides.add_slide(prs.slide_layouts[6]);bg=s.background.fill;bg.solid();bg.fore_color.rgb=PRGB(255,255,255)
    bar=s.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,PInches(0),PInches(0),prs.slide_width,PInches(0.15));bar.fill.solid();bar.fill.fore_color.rgb=PRGB.from_string(ORANGE);bar.line.fill.background()
    tb=s.shapes.add_textbox(PInches(0.55),PInches(0.35),prs.slide_width-PInches(1.1),PInches(0.55));p=tb.text_frame.paragraphs[0];p.text=title;p.font.name='Arial Unicode MS';p.font.size=PPt(24);p.font.bold=True;p.font.color.rgb=PRGB.from_string(NAVY)
    if subtitle:
        t=s.shapes.add_textbox(PInches(0.57),PInches(0.92),prs.slide_width-PInches(1.14),PInches(0.35));p=t.text_frame.paragraphs[0];p.text=subtitle;p.font.name='Arial Unicode MS';p.font.size=PPt(10);p.font.color.rgb=PRGB.from_string(GREY)
    return s

def add_kpi(s,x,y,w,h,label,value,color=ORANGE):
    box=s.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,PInches(x),PInches(y),PInches(w),PInches(h));box.fill.solid();box.fill.fore_color.rgb=PRGB.from_string(PALE);box.line.color.rgb=PRGB.from_string(LIGHT)
    tf=box.text_frame;tf.clear();tf.margin_left=PInches(.12);tf.margin_right=PInches(.12);tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    p=tf.paragraphs[0];p.text=str(value);p.alignment=PP_ALIGN.CENTER;p.font.name='Arial';p.font.size=PPt(28);p.font.bold=True;p.font.color.rgb=PRGB.from_string(color)
    p2=tf.add_paragraph();p2.text=label;p2.alignment=PP_ALIGN.CENTER;p2.font.name='Arial Unicode MS';p2.font.size=PPt(11);p2.font.color.rgb=PRGB.from_string(NAVY)

def add_bullets(slide,items,x=0.65,y=1.45,w=12.0,h=5.2,size=16):
    tb=slide.shapes.add_textbox(PInches(x),PInches(y),PInches(w),PInches(h));tf=tb.text_frame;tf.clear();tf.word_wrap=True
    for i,item in enumerate(items):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph();p.text='• '+item;p.level=0;p.font.name='Arial Unicode MS';p.font.size=PPt(size);p.font.color.rgb=PRGB.from_string('333333');p.space_after=PPt(8)

def add_ppt_table(slide,headers,rows,x=.65,y=1.55,w=12.0,h=4.5,font=11):
    shape=slide.shapes.add_table(len(rows)+1,len(headers),PInches(x),PInches(y),PInches(w),PInches(h));table=shape.table
    for j,h in enumerate(headers):
        c=table.cell(0,j);c.text=h;c.fill.solid();c.fill.fore_color.rgb=PRGB.from_string(BLUE);c.text_frame.paragraphs[0].font.bold=True;c.text_frame.paragraphs[0].font.color.rgb=PRGB.from_string(WHITE);c.text_frame.paragraphs[0].font.name='Arial Unicode MS';c.text_frame.paragraphs[0].font.size=PPt(font)
    for i,row in enumerate(rows,1):
        for j,v in enumerate(row):
            c=table.cell(i,j);c.text=str(v);c.fill.solid();c.fill.fore_color.rgb=PRGB.from_string('FFFFFF' if i%2 else 'F3F6FA');c.text_frame.paragraphs[0].font.name='Arial Unicode MS';c.text_frame.paragraphs[0].font.size=PPt(font-1);c.text_frame.paragraphs[0].font.color.rgb=PRGB.from_string('333333')

def build_ppt(ledger,metrics,period,out,prev_dir=None):
    prs=Presentation();prs.slide_width=PInches(13.333);prs.slide_height=PInches(7.5);facts=workbook_facts(Path(out).parent)
    audit=unique_event_rows(event_rows(ledger,'ma_audit',{'first_disclosed','first_draft'}));passed=unique_event_rows(event_rows(ledger,'ma_audit',{'committee_passed','simplified_passed'}));control=unique_event_rows(event_rows(ledger,'control',{'first_disclosed'}));completed=unique_event_rows(event_rows(ledger,'control',{'control_completed'}));partial=unique_event_rows(event_rows(ledger,'partial_transfer',{'first_disclosed'}));cash=unique_event_rows(event_rows(ledger,{'ma_cash_purchase','ma_cash_sale'},{'first_disclosed','first_draft'}));strategic=unique_event_rows(event_rows(ledger,{'strategic_investment','professional_cooperation'},{'first_disclosed'}));terms=unique_event_rows(event_rows(ledger,'ma_audit',{'terminated'}));regs=unique_event_rows(event_rows(ledger,'ma_audit',{'registered'}));accepted=unique_event_rows(event_rows(ledger,'ma_audit',{'accepted'}));first_all={x['case'].get('code') or x['event']['deal_id']:x for x in audit+cash}
    # Cover
    s=prs.slides.add_slide(prs.slide_layouts[6]);bg=s.background.fill;bg.solid();bg.fore_color.rgb=PRGB.from_string(NAVY)
    accent=s.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE,PInches(.7),PInches(1.25),PInches(.1),PInches(4.6));accent.fill.solid();accent.fill.fore_color.rgb=PRGB.from_string(ORANGE);accent.line.fill.background()
    tb=s.shapes.add_textbox(PInches(.95),PInches(1.45),PInches(10.8),PInches(1));p=tb.text_frame.paragraphs[0];p.text='A股并购重组市场周报';p.font.name='Arial Unicode MS';p.font.size=PPt(34);p.font.bold=True;p.font.color.rgb=PRGB.from_string(WHITE)
    tb=s.shapes.add_textbox(PInches(1),PInches(3),PInches(9),PInches(.8));p=tb.text_frame.paragraphs[0];p.text=f'{period["start"][:4]}年第{period["issue_no"]}期';p.font.name='Arial Unicode MS';p.font.size=PPt(20);p.font.color.rgb=PRGB.from_string('D9E2F3')
    tb=s.shapes.add_textbox(PInches(1),PInches(3.65),PInches(9),PInches(.6));p=tb.text_frame.paragraphs[0];p.text=f'{cn_date(period["start"][:4]+"-"+period["start"][4:6]+"-"+period["start"][6:])}至{cn_date(period["end"][:4]+"-"+period["end"][4:6]+"-"+period["end"][6:])}';p.font.name='Arial Unicode MS';p.font.size=PPt(16);p.font.color.rgb=PRGB.from_string('D9E2F3')
    # KPI
    s=add_slide_title(prs,'本周核心数据',f'官方公告与结构化审核动态；截至{cn_date(period["end"][:4]+"-"+period["end"][4:6]+"-"+period["end"][6:])}')
    for i,(lab,val) in enumerate([('首次披露',len(first_all)),('重组委通过',len(passed)),('新增受理',len(accepted)),('新增注册',len(regs)),('终止',len(terms))]):add_kpi(s,.65+i*2.5,1.55,2.2,1.35,lab,val,ORANGE if i<3 else NAVY)
    add_bullets(s,[f'预案/草案首次披露 {len(first_all)} 单；筹划/停牌公告另列，不混入首次披露。',f'重组委通过：{safe_names(passed)}。',f'新增受理：{safe_names(accepted)}；新增注册：{safe_names(regs)}。',f"期末工作簿在审 {facts.get('in_review','待核')} 单；来源覆盖：{coverage_summary(ledger)}。"],size=15)
    # First disclosures: plan/draft, including cash major transactions.
    s=add_slide_title(prs,'预案/草案首次披露','按证券代码去重；金额未核实时不填')
    first_rows=deal_union(audit,cash,dedupe_code=True)
    add_ppt_table(s,['日期','公司','交易标的','方式','金额','顾问'],[[cn_date(x['event']['date']),x['case']['name'],str(x['case'].get('facts',{}).get('target','待核'))[:26],x['case'].get('facts',{}).get('payment_method','待核'),amount_text(amounts(x['case'])),advisor_text(x['case'])] for x in first_rows],font=9)
    # passed
    s=add_slide_title(prs,'重组委审核结果',f"会议日与公告日分列；本期会议 {metrics.get('weekly',{}).get('committee_meetings',0)} 次")
    add_ppt_table(s,['会议日期','公司','结果','独立财务顾问'],[[cn_date(x['event']['date']),x['case']['name'],'通过',advisor_text(x['case'])] for x in passed],font=13)
    add_bullets(s,[f"通过项目：{safe_names(passed)}。",'会议日取重组委实际召开日；公告披露日只作为可追溯证据日期。','年度累计值仅在历史工作簿完成桥接后展示，不从本周台账外推。'],y=4.55,size=16)
    # control and partial
    s=add_slide_title(prs,'控制权与协议转让','实际控制人变化和“控制权不变”协议转让分别筛选')
    rows=[]
    for x in control:rows.append(['控制权',x['case']['name'],x['case'].get('facts',{}).get('old_controller','待核'),x['case'].get('facts',{}).get('new_controller','待核')])
    for x in partial:rows.append(['协议转让',x['case']['name'],x['case'].get('facts',{}).get('old_controller','实控人未变'),x['case'].get('facts',{}).get('new_controller','—')])
    add_ppt_table(s,['分类','公司','原实控人','新实控人'],rows,font=11,h=4.6)
    # cash/strategic
    s=add_slide_title(prs,'现金交易与子公司引战','只列正文明确构成重大资产重组或符合引战合作口径的项目')
    items=[]
    for x in cash:items.append(f"{x['case']['name']}：{x['case'].get('facts',{}).get('target','待核')}，现金交易。")
    for x in strategic:items.append(f"{x['case']['name']}：子公司引入战略投资者并拟签署股权转让协议。")
    add_bullets(s,items,size=17)
    # Interpretation
    s=add_slide_title(prs,'本期观察','以公告事实为边界，不用标题推断结论')
    payment_counts={}
    for x in audit:
        key=x['case'].get('facts',{}).get('payment_method','未核实');payment_counts[key]=payment_counts.get(key,0)+1
    observations=[f"审核类首次披露 {len(audit)} 单；支付方式分布：{('、'.join(f'{k}{v}单' for k,v in payment_counts.items()) or '待核')}。",f"控制权首次披露 {len(control)} 单、完成 {len(completed)} 单；协议转让项目 {len(partial)} 单。",'控股股东变化、实际控制人变化、协议转让分别按正文事实判断。',f'覆盖缺口：{coverage_summary(ledger)}。']
    add_bullets(s,observations,size=16)
    # Source
    s=add_slide_title(prs,'来源与待核','本页不列未核实数值')
    pdf_count=sum(1 for s in source_map(ledger).values() if str(s.get('local_path','')).lower().endswith('.pdf'))
    add_bullets(s,['基础来源：交易所项目动态、会议公告入口、巨潮公告全文及公司公告原件。',f'公告原件：{pdf_count} 份PDF已保存；引用事件保留标题、URL与正文页码（已获取时）。','仍待补事项见随附《本期编制与待核说明》，不在演示中补造数字。',f'来源覆盖：{coverage_summary(ledger)}。'],size=16)
    add_footer= f'HTLH｜{period["start"][:4]}年第{period["issue_no"]}期｜{cn_date(period["start"][:4]+"-"+period["start"][4:6]+"-"+period["start"][6:])}至{cn_date(period["end"][:4]+"-"+period["end"][4:6]+"-"+period["end"][6:])}'
    for slide in prs.slides:
        tb=slide.shapes.add_textbox(PInches(.5),PInches(7.12),PInches(12.3),PInches(.2));p=tb.text_frame.paragraphs[0];p.text=add_footer;p.alignment=PP_ALIGN.RIGHT;p.font.name='Arial Unicode MS';p.font.size=PPt(8);p.font.color.rgb=PRGB.from_string('7F7F7F')
    prs.core_properties.author='HTLH';prs.core_properties.last_modified_by='HTLH';prs.save(out)

def main():
    ap=argparse.ArgumentParser(description=__doc__);ap.add_argument('--ledger',required=True,type=Path);ap.add_argument('--metrics',required=True,type=Path);ap.add_argument('--period',required=True,type=Path);ap.add_argument('--outdir',required=True,type=Path);ap.add_argument('--prev',type=Path);args=ap.parse_args()
    argv=vars(args);ledger=load(argv['ledger']);metrics=load(argv['metrics']);period=load(argv['period']);out=argv['outdir'];out.mkdir(parents=True,exist_ok=True)
    tag=f'{period["start"]}至{period["end"]}'
    weekly=out/f'5、{period["year"]}年并购重组周报_{tag}.docx';dynamic=out/f'A股市场并购重组动态-{tag}.docx';ppt=out/f'并购重组市场周报_{tag}.pptx'
    build_weekly(ledger,metrics,period,weekly,argv['prev']);build_dynamic(ledger,metrics,period,dynamic,argv['prev']);build_ppt(ledger,metrics,period,ppt,argv['prev'])
    # Persist the workbook-derived as-of snapshot beside the ledger metrics so
    # downstream notes/tests do not mistake this week's event cohort for annual
    # or current-in-review totals.
    wf=workbook_facts(Path(out));metrics['workbook_snapshot']=wf
    metrics.setdefault('as_of_in_review',{})['projects']=wf.get('in_review')
    metrics['as_of_in_review']['source']='content 0 workbook: 是否在审=是'
    for key in ('first_disclosed','committee_passed','registered'):
        if key in metrics.get('annual',{}):
            metrics['annual'][key]['projects']=wf.get(key)
            metrics['annual'][key]['source']='content workbook'
    argv['metrics'].write_text(json.dumps(metrics,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'weekly':str(weekly),'dynamic':str(dynamic),'ppt':str(ppt)},ensure_ascii=False))
if __name__=='__main__':main()
