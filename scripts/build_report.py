#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""成绩分析与报告生成引擎（通用版）

输入：标准列契约 Excel + config.yaml
输出：Word 分析报告（.docx）+ 交互式 HTML（.html）

所有学校/班级/科目/考试/折算锚点/诊断归因均由配置注入，
本脚本不含任何真实个人数据。
"""
import argparse
import json
import math
import os
import sys
import warnings
from collections import defaultdict

warnings.filterwarnings('ignore')

try:
    import yaml
    import openpyxl
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.section import WD_ORIENT
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError as exc:
    sys.stderr.write('缺少依赖：%s\n请 pip install openpyxl python-docx PyYAML\n' % exc)
    sys.exit(2)


# ───────────────────────── 工具 ─────────────────────────

def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)


def num(v, nd=1, dash='-'):
    if not is_num(v):
        return dash
    return ('%.*f' % (nd, v)).rstrip('0').rstrip('.') if nd else str(int(v))


def signed(v, nd=1):
    if not is_num(v):
        return '-'
    n = round(v, nd)
    return ('+' if n > 0 else '') + num(n, nd)


def mean(vals):
    vals = [v for v in vals if is_num(v)]
    return sum(vals) / len(vals) if vals else None


def load_config(path):
    with open(path, encoding='utf-8') as f:
        cfg = yaml.safe_load(f)
    required = ['school', 'class_name', 'subjects', 'input', 'tier_rules']
    missing = [k for k in required if k not in cfg]
    if missing:
        raise ValueError('配置缺少必填字段：%s' % ', '.join(missing))
    if not cfg.get('exam_chain'):
        raise ValueError('配置缺少必填字段：exam_chain')
    if not cfg['input'].get('columns', {}).get('name'):
        raise ValueError('配置缺少 input.columns.name（姓名列映射）')
    return cfg


def tier_of(rank, cfg):
    if not is_num(rank):
        return cfg.get('missing_tier', '数据不全')
    for rule in cfg['tier_rules']:
        mx = rule.get('max')
        if mx is None or rank <= mx:
            return rule.get('tier', '未分层')
    return cfg['tier_rules'][-1].get('tier', '未分层')


def tier_order(cfg):
    return [r.get('tier', '未分层') for r in cfg['tier_rules']] + [
        cfg.get('missing_tier', '数据不全')]


def tier_colors(cfg):
    c = {r.get('tier', '未分层'): r.get('color', '#E8F6F2') for r in cfg['tier_rules']}
    c[cfg.get('missing_tier', '数据不全')] = '#ECECEC'
    return c


# ───────────────────────── 数据加载 ─────────────────────────

def load_students(cfg, path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError('Excel 为空')
    header = [str(h).strip() if h is not None else '' for h in rows[0]]
    idx = {h: i for i, h in enumerate(header) if h}
    cols = cfg['input']['columns']
    c_name = cols.get('name')
    c_sid = cols.get('sid')
    c_rank = cols.get('rank')
    if c_name not in idx:
        raise ValueError('Excel 中找不到姓名列：%s' % c_name)

    def cell(row, key):
        if key and key in idx and idx[key] < len(row):
            return row[idx[key]]
        return None

    students = []
    for r in rows[1:]:
        name = cell(r, c_name)
        if name is None or not str(name).strip():
            continue
        name = str(name).strip()
        rank = cell(r, c_rank)
        rank = float(rank) if is_num(rank) else None
        stu = {
            'name': name,
            'sid': str(cell(r, c_sid)).strip() if cell(r, c_sid) is not None else '',
            'rank': rank,
            'tier': tier_of(rank, cfg),
            'totals': {},
            'sections': {},
        }
        # 学科总分（按考试）
        totals_map = cols.get('totals', {}) or {}
        for sub, exams in totals_map.items():
            stu['totals'][sub] = {}
            for exam in cfg['exam_chain']:
                v = cell(r, (exams or {}).get(exam))
                stu['totals'][sub][exam] = float(v) if is_num(v) else None
        # 板块得分率（按学科 → 考试 → 板块）
        rates_map = cols.get('section_rates', {}) or {}
        for sub, exams in rates_map.items():
            stu['sections'][sub] = {}
            for exam in cfg['exam_chain']:
                blk_map = (exams or {}).get(exam, {}) or {}
                stu['sections'][sub][exam] = {}
                for blk, colname in blk_map.items():
                    v = cell(r, colname)
                    stu['sections'][sub][exam][blk] = float(v) if is_num(v) else None
        students.append(stu)
    if not students:
        raise ValueError('Excel 中没有有效学生行')
    return students


# ───────────────────────── 分析计算 ─────────────────────────

def subject_mean(students, sub, exam):
    return mean([s['totals'].get(sub, {}).get(exam) for s in students])


def weakest_section(stu, sub, exam):
    rates = stu.get('sections', {}).get(sub, {}).get(exam, {}) or {}
    valid = [(b, v) for b, v in rates.items() if is_num(v)]
    if not valid:
        return None, None
    return min(valid, key=lambda x: x[1])


def build_analysis(cfg, students):
    exams = cfg['exam_chain']
    latest = exams[-1]
    first = exams[0]
    out = {'students': students, 'subjects': {}, 'exams': exams,
           'latest': latest, 'first': first}

    for sub, meta in cfg['subjects'].items():
        full = meta.get('full', 150)
        sections = meta.get('sections', {}) or {}
        # 板块班均得分率
        sec_means = {}
        for exam in exams:
            sec_means[exam] = {}
            for blk in sections:
                vals = [s.get('sections', {}).get(sub, {}).get(exam, {}).get(blk)
                        for s in students]
                sec_means[exam][blk] = mean(vals)
        # 薄弱诊断
        diag = []
        for s in students:
            blk, rate = weakest_section(s, sub, latest)
            if blk is None:
                diag.append({'name': s['name'], 'tier': s['tier'], 'weak': '未配置',
                             'rate': None, 'cause': '未配置板块得分率数据',
                             'advice': '如需诊断，请在配置中填写 section_rates'})
                continue
            wr = (cfg.get('weak_reasons') or {}).get(blk) or cfg.get('default_reason') or {}
            diag.append({'name': s['name'], 'tier': s['tier'], 'weak': blk,
                         'rate': rate, 'cause': wr.get('cause', ''),
                         'advice': wr.get('advice', '')})
        diag.sort(key=lambda d: (d['rate'] if is_num(d['rate']) else 1e9))
        out['subjects'][sub] = {
            'full': full,
            'sections': sections,
            'sec_means': sec_means,
            'diag': diag,
            'means': {e: subject_mean(students, sub, e) for e in exams},
        }
    return out


# ───────────────────────── Word 生成 ─────────────────────────

def set_cell_shading(cell, hex_color):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color.lstrip('#'))
    tc_pr.append(shd)


def set_cell_text(cell, text, bold=False, color=None, size=10.5, align='center'):
    cell.text = ''
    p = cell.paragraphs[0]
    p.paragraph_format.line_spacing = 1.5
    p.alignment = {'center': WD_ALIGN_PARAGRAPH.CENTER,
                   'left': WD_ALIGN_PARAGRAPH.LEFT}.get(align, WD_ALIGN_PARAGRAPH.CENTER)
    r = p.add_run(str(text))
    r.font.name = 'Times New Roman'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), '华文楷体')
    r.font.size = Pt(size)
    r.bold = bold
    if color:
        r.font.color.rgb = RGBColor.from_string(color.lstrip('#'))


def repeat_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    el = OxmlElement('w:tblHeader')
    el.set(qn('w:val'), 'true')
    tr_pr.append(el)


def add_table(doc, headers, rows, widths=None, tier_col=None, color_map=None):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    for i, h in enumerate(headers):
        if widths and i < len(widths):
            for c in t.columns[i].cells:
                c.width = Cm(widths[i])
        set_cell_text(t.rows[0].cells[i], h, bold=True, color='FFFFFF', size=10.5)
        set_cell_shading(t.rows[0].cells[i], '3CB39B')
    repeat_header(t.rows[0])
    for row in rows:
        cells = t.add_row().cells
        tier_name = row[tier_col] if tier_col is not None and tier_col < len(row) else None
        fill = (color_map or {}).get(tier_name)
        for i, v in enumerate(row):
            if widths and i < len(widths):
                cells[i].width = Cm(widths[i])
            set_cell_text(cells[i], v, align='left' if i == 0 else 'center')
            if fill:
                set_cell_shading(cells[i], fill)
    return t


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), '华文楷体')
    r.bold = True
    r.font.size = Pt(14 if level == 1 else 12)
    r.font.color.rgb = RGBColor(0x2E, 0x8F, 0x79)


def add_text(doc, text, size=10.5):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r._element.rPr.rFonts.set(qn('w:eastAsia'), '华文楷体')
    r.font.size = Pt(size)


def setup_landscape(doc):
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = Cm(29.7), Cm(21.0)
    sec.left_margin = sec.right_margin = Cm(1.0)
    sec.top_margin = sec.bottom_margin = Cm(1.2)


def tier_width(ncols):
    base = 26.8 / ncols
    return [base] * ncols


def build_docx(cfg, analysis, out_path):
    doc = Document()
    setup_landscape(doc)
    title = '%s %s 学科能力与分层分析报告' % (cfg['school'], cfg['class_name'])
    add_heading(doc, title, 1)
    subtitle = '%s　|　数据区间：%s　|　生成日期：%s' % (
        cfg.get('term', ''), ' → '.join(analysis['exams']), cfg.get('update_date', ''))
    add_text(doc, subtitle, 10.5)
    if cfg.get('analyst'):
        add_text(doc, '编制：%s' % cfg['analyst'], 10.5)

    # 层次说明
    add_heading(doc, '层次折算说明', 2)
    tier_rows = [[r.get('tier', ''),
                  ('≤ %s' % num(r['max'], 0)) if is_num(r.get('max')) else '其余',
                  r.get('desc', '')] for r in cfg['tier_rules']]
    add_table(doc, ['层次', '校内排名', '说明'], tier_rows,
              widths=[4.0, 4.0, 18.8])

    rec = cfg.get('school_recommend')
    if rec:
        add_heading(doc, '各层次适配方向（示例）', 2)
        rr = []
        for t in tier_order(cfg):
            schools = rec.get(t, [])
            txt = '；'.join('%s %s' % (x.get('name', ''), x.get('note', '')) for x in schools)
            rr.append([t, txt or '—'])
        add_table(doc, ['层次', '适配方向'], rr, widths=[4.0, 22.8],
                  tier_col=0, color_map=tier_colors(cfg))

    exams = analysis['exams']
    latest = analysis['latest']
    first = analysis['first']
    for sub, meta in analysis['subjects'].items():
        add_heading(doc, '%s 学科分析' % sub, 1)
        full = meta['full']

        # 1) 能力摘要（板块得分率两考对比）
        blk_rows = []
        for blk, blk_full in meta['sections'].items():
            rv = meta['sec_means'][latest].get(blk)
            qv = meta['sec_means'][first].get(blk) if first != latest else None
            delta = (rv - qv) if (is_num(rv) and is_num(qv)) else None
            pos = ('薄弱' if rv is not None and rv < 60 else
                   '中等' if rv is not None and rv < 80 else
                   '优势' if rv is not None else '无数据')
            blk_rows.append([blk, num(blk_full, 0),
                             num(qv, 1) if is_num(qv) else '-',
                             num(rv, 1) if is_num(rv) else '-',
                             signed(delta, 1), pos])
        if blk_rows:
            heads = ['板块', '满分', '%s得分率' % first, '%s得分率' % latest, '变化', '定位']
            add_table(doc, heads, blk_rows, widths=tier_width(len(heads)))

        # 2) 学生分层
        tier_rows2 = []
        for s in sorted(analysis['students'],
                        key=lambda x: (tier_order(cfg).index(x['tier']),
                                       x['rank'] if is_num(x['rank']) else 1e9)):
            t = s['totals'].get(sub, {})
            q, y = t.get(first), t.get(latest)
            diff = (y - q) if (is_num(y) and is_num(q)) else None
            m = subject_mean(analysis['students'], sub, latest)
            dev = (y - m) if (is_num(y) and is_num(m)) else None
            tier_rows2.append([s['name'], s['tier'],
                               num(s['rank'], 0) if is_num(s['rank']) else '-',
                               num(q), num(y), signed(diff), signed(dev)])
        heads2 = ['姓名', '层次', '校内排名', first, latest, '较首次增减', '班均偏离']
        add_table(doc, heads2, tier_rows2, widths=tier_width(len(heads2)),
                  tier_col=1, color_map=tier_colors(cfg))

        # 3) 薄弱诊断
        diag_rows = [[d['name'], d['tier'], d['weak'],
                      num(d['rate'], 1), d['cause'], d['advice']]
                     for d in meta['diag']]
        heads3 = ['姓名', '层次', '最弱板块', '得分率', '归因', '训练建议']
        add_table(doc, heads3, diag_rows, widths=[2.8, 2.8, 2.8, 2.2, 7.0, 9.2],
                  tier_col=1, color_map=tier_colors(cfg))

        # 4) 学科小结
        add_heading(doc, '学科小结', 2)
        latest_means = [meta['means'].get(e) for e in exams]
        trend = ' → '.join(num(v) for v in latest_means)
        weak_blks = [b for b, v in meta['sec_means'][latest].items()
                     if is_num(v) and v < 60]
        add_text(doc, '%s班均变化：%s。%s' % (
            sub, trend,
            ('班级薄弱板块：%s，建议优先安排专项训练。' % '、'.join(weak_blks)) if weak_blks
            else '各板块整体达到中等以上，保持节奏并关注个体短板。'))

    add_heading(doc, '免责声明', 2)
    add_text(doc,
             '本文档的层次折算与适配方向为基于校内排名的经验估算，仅供学情预判参考，'
             '不构成对任何学生学业结果的保证；板块得分率来自所提供数据，'
             '请以原始成绩与教师研判为准。', 10.5)

    doc.save(out_path)


# ───────────────────────── HTML 生成 ─────────────────────────

HTML_TPL = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>__TITLE__</title>
<style>
  :root{--mint:#3CB39B;--mint-d:#2E8F79;--mint-l:#E8F6F2;--mint-ll:#F4FBF9;
    --ink:#1F3A34;--muted:#5C7269;--line:#DCEAE5}
  *{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
  body{margin:0;font-family:"PingFang SC","Microsoft YaHei",sans-serif;background:var(--mint-ll);
    color:var(--ink);line-height:1.5}
  .wrap{max-width:1280px;margin:0 auto;padding:20px 16px 60px}
  header{background:linear-gradient(135deg,var(--mint),var(--mint-d));color:#fff;
    border-radius:18px;padding:24px 26px;margin-bottom:18px}
  header h1{margin:0 0 8px;font-size:24px}
  header .note{font-size:13px;opacity:.92}
  .card{background:#fff;border-radius:16px;padding:18px;margin-bottom:18px;
    box-shadow:0 3px 12px rgba(31,58,52,.06);border:1px solid var(--line)}
  .card h2{margin:0 0 4px;font-size:18px;color:var(--mint-d)}
  .card .sub{font-size:12.5px;color:var(--muted);margin-bottom:12px}
  .toolbar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:10px}
  select,input[type=search]{border:1.5px solid var(--line);border-radius:10px;padding:8px 12px;
    font-size:14px;color:var(--ink);background:#fff;outline:none;min-width:130px}
  input[type=search]{min-width:200px}
  .count{font-size:12.5px;color:var(--muted);margin-left:auto}
  .table-scroll{overflow-x:auto;border-radius:12px;border:1px solid var(--line)}
  table{border-collapse:collapse;width:100%;font-size:13px;white-space:nowrap}
  thead th{background:var(--mint);color:#fff;font-weight:600;padding:10px;text-align:center;
    cursor:pointer;user-select:none;position:sticky;top:0}
  thead th.nosort{cursor:default}
  tbody td{padding:8px 10px;text-align:center;border-bottom:1px solid var(--line)}
  tbody tr:nth-child(even){background:var(--mint-ll)}
  tbody tr:hover{background:var(--mint-l)}
  td.t{text-align:left}
  .pill{padding:3px 10px;border-radius:20px;font-size:12px;font-weight:600}
  .up{color:#E06C4A;font-weight:600}.down{color:#3E8E5A;font-weight:600}
  .adv{text-align:left;font-size:12px;color:var(--muted);min-width:200px;white-space:normal}
  .rpt-btn{border:none;background:var(--mint);color:#fff;border-radius:8px;padding:5px 12px;
    font-size:12px;font-weight:600;cursor:pointer}
  .pdf-frame{height:70vh;min-height:480px;border:1px solid var(--line);border-radius:12px;overflow:hidden}
  .pdf-frame iframe{width:100%;height:100%;border:none;display:block}
  .pdf-empty{display:flex;align-items:center;justify-content:center;height:100%;color:var(--muted)}
  .viewer-bar{display:flex;flex-wrap:wrap;gap:10px;align-items:center;margin-bottom:10px}
  .vbtn{border:1.5px solid var(--mint);background:#fff;color:var(--mint-d);border-radius:10px;
    padding:7px 13px;font-size:13px;font-weight:600;cursor:pointer}
  footer{font-size:12px;color:var(--muted);margin-top:8px;line-height:1.7}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>__TITLE__</h1>
    <div class="note">__SUBTITLE__</div>
  </header>
  <div id="tables"></div>
  <div class="card" id="reportCard" __PDF_HIDDEN__>
    <h2>个人报告预览（可选）</h2>
    <div class="sub">需自行生成 PDF 并放入本 HTML 同目录；文件名按配置中的 pattern 命名。</div>
    <div class="viewer-bar">
      <select id="rptSel"></select>
      <button class="vbtn" id="rptPrev">← 上一人</button>
      <button class="vbtn" id="rptNext">下一人 →</button>
      <span class="count" id="rptInfo">尚未选择</span>
    </div>
    <div class="pdf-frame" id="pdfFrame"><div class="pdf-empty">选择学生后在此预览 PDF</div></div>
  </div>
  <footer>__FOOTER__</footer>
</div>
<script>
const DATA = __DATA__;
const TIER_ORDER = DATA.tierOrder;
const TIER_COLORS = DATA.tierColors;
const tierIdx = {}; TIER_ORDER.forEach((t,i)=>tierIdx[t]=i);
function fmt(v,type){
  if(v===null||v===undefined||v==='') return '-';
  if(type==='pct') return v+'%';
  if(type==='num') return (Math.round(v*10)/10).toString();
  if(type==='signed'){const n=Math.round(v*10)/10;return (n>0?'+':'')+n;}
  return String(v);
}
function tierPill(t){const c=TIER_COLORS[t]||'#E8F6F2';
  return '<span class="pill" style="background:'+c+'">'+t+'</span>';}
function renderTable(cfg){
  const cols=cfg.cols, rows=[...cfg.rows];
  let sortKey=null,sortDir=1,filterTier='all',search='';
  const card=document.createElement('div');card.className='card';
  card.innerHTML='<h2>'+cfg.title+'</h2><div class="sub">'+cfg.sub+'</div>'+
    '<div class="toolbar" id="tb_'+cfg.id+'"></div>'+
    '<div class="table-scroll"><table id="tb_'+cfg.id+'_t"><thead></thead><tbody></tbody></table></div>';
  document.getElementById('tables').appendChild(card);
  const toolbar=card.querySelector('.toolbar'),thead=card.querySelector('thead'),
    tbody=card.querySelector('tbody'),countEl=document.createElement('span');
  countEl.className='count';
  if(cfg.filter){
    const sel=document.createElement('select');
    sel.innerHTML='<option value="all">层次：全部</option>'+
      TIER_ORDER.map(t=>'<option value="'+t+'">'+t+'</option>').join('');
    sel.onchange=()=>{filterTier=sel.value;refresh();};
    const inp=document.createElement('input');inp.type='search';inp.placeholder='搜索姓名…';
    inp.oninput=()=>{search=inp.value.trim().toLowerCase();refresh();};
    toolbar.appendChild(sel);toolbar.appendChild(inp);
  }
  toolbar.appendChild(countEl);
  function refresh(){
    let vis=rows.filter(r=>{
      if(filterTier!=='all'&&r.tier!==filterTier) return false;
      if(search&&!String(r.name).toLowerCase().includes(search)) return false;
      return true;
    });
    if(sortKey){
      vis.sort((a,b)=>{
        const type=(cols.find(c=>c[0]===sortKey)||[])[2];
        const va=a[sortKey],vb=b[sortKey];
        const na=va===null||va===undefined,nb=vb===null||vb===undefined;
        if(na&&nb) return 0; if(na) return 1; if(nb) return -1;
        if(type==='tier') return (tierIdx[va]-tierIdx[vb])*sortDir;
        if(type==='num'||type==='signed') return (va-vb)*sortDir;
        return String(va).localeCompare(String(vb),'zh')*sortDir;
      });
    }
    countEl.textContent='共 '+vis.length+' 行';
    let h='';
    cols.forEach(([key,label,type])=>{
      const arrow=sortKey===key?(sortDir===1?' ▲':' ▼'):'';
      const ns=(type==='text'||type==='action')?'nosort':'';
      h+='<th class="'+ns+'" data-key="'+key+'">'+label+'<span class="arrow">'+arrow+'</span></th>';
    });
    thead.innerHTML='<tr>'+h+'</tr>';
    let b='';
    vis.forEach(r=>{
      b+='<tr>';
      cols.forEach(([key,label,type])=>{
        let cls=key==='name'?'t':(type==='adv'?'adv':'');
        let val=fmt(r[key],type);
        if(type==='tier') val=tierPill(r[key]);
        if(type==='signed'&&r[key]!==null&&r[key]!==undefined){
          const n=r[key];
          val='<span class="'+(n>0?'up':(n<0?'down':''))+'">'+val+'</span>';
        }
        if(type==='action'){
          val=(r.pdf)?'<button class="rpt-btn" data-name="'+r.name+'" data-pdf="'+r.pdf+'">查看</button>':'<span style="color:#aaa">—</span>';
        }
        b+='<td class="'+cls+'">'+val+'</td>';
      });
      b+='</tr>';
    });
    tbody.innerHTML=b;
    thead.querySelectorAll('th').forEach(th=>{
      if(th.classList.contains('nosort')) return;
      th.onclick=()=>{const k=th.dataset.key;
        if(sortKey===k) sortDir=-sortDir; else {sortKey=k;sortDir=1;}
        refresh();};
    });
    tbody.querySelectorAll('.rpt-btn').forEach(btn=>{
      btn.onclick=()=>openReport(btn.dataset.name,btn.dataset.pdf);
    });
  }
  refresh();
}
function openReport(name,pdf){
  document.getElementById('pdfFrame').innerHTML=
    '<iframe src="'+pdf+'" title="'+name+'"></iframe>';
  document.getElementById('rptInfo').textContent=name;
  document.getElementById('reportCard').scrollIntoView({behavior:'smooth'});
}
DATA.tables.forEach(renderTable);
(function(){
  if(!DATA.pdfEnabled) return;
  const sel=document.getElementById('rptSel');
  sel.innerHTML='<option value="-1">选择学生…</option>'+
    DATA.reports.map((r,i)=>'<option value="'+i+'">'+r.name+'</option>').join('');
  let cur=-1;
  function show(i){if(i<0||i>=DATA.reports.length) return;cur=i;
    sel.value=String(i);openReport(DATA.reports[i].name,DATA.reports[i].pdf);}
  sel.onchange=()=>show(Number(sel.value));
  document.getElementById('rptPrev').onclick=()=>show(cur-1);
  document.getElementById('rptNext').onclick=()=>show(cur+1);
})();
</script>
</body>
</html>
'''


def build_html(cfg, analysis, out_path, pdf_names):
    exams = analysis['exams']
    latest = analysis['latest']
    first = analysis['first']
    tables = []

    # 折算说明（静态表）
    rec_rows = []
    for rule in cfg['tier_rules']:
        t = rule.get('tier', '')
        rec_rows.append({'tier': t, 'range': ('≤ %s' % num(rule['max'], 0))
                         if is_num(rule.get('max')) else '其余',
                         'desc': rule.get('desc', ''),
                         'school': '；'.join(x.get('name', '') for x in
                                            (cfg.get('school_recommend') or {}).get(t, []))})
    tables.append({
        'id': 'tier_info', 'title': '层次折算说明与适配方向',
        'sub': '经验估算，仅供学情预判参考', 'filter': False,
        'cols': [['tier', '层次', 'tier'], ['range', '校内排名', 'text'],
                 ['desc', '说明', 'text'], ['school', '适配方向（示例）', 'text']],
        'rows': rec_rows,
    })

    for sub, meta in analysis['subjects'].items():
        # 能力摘要
        sr = []
        for blk, bf in meta['sections'].items():
            rv = meta['sec_means'][latest].get(blk)
            qv = meta['sec_means'][first].get(blk) if first != latest else None
            sr.append({'blk': blk, 'full': bf,
                       'q': round(qv, 1) if is_num(qv) else None,
                       'r': round(rv, 1) if is_num(rv) else None,
                       'd': round(rv - qv, 1) if (is_num(rv) and is_num(qv)) else None})
        tables.append({'id': '%s_summary' % sub, 'title': '%s 能力摘要：板块得分率对比' % sub,
                       'sub': '班级层面，点击列头排序', 'filter': False,
                       'cols': [['blk', '板块', 'text'], ['full', '满分', 'num'],
                                ['q', '%s得分率' % first, 'num'],
                                ['r', '%s得分率' % latest, 'num'],
                                ['d', '变化', 'signed']], 'rows': sr})
        # 学生分层
        tr = []
        for s in sorted(analysis['students'],
                        key=lambda x: (tier_order(cfg).index(x['tier']),
                                       x['rank'] if is_num(x['rank']) else 1e9)):
            t = s['totals'].get(sub, {})
            q, y = t.get(first), t.get(latest)
            m = subject_mean(analysis['students'], sub, latest)
            tr.append({'name': s['name'], 'tier': s['tier'],
                       'rank': round(s['rank']) if is_num(s['rank']) else None,
                       'q': round(q, 1) if is_num(q) else None,
                       'r': round(y, 1) if is_num(y) else None,
                       'd': round(y - q, 1) if (is_num(y) and is_num(q)) else None,
                       'dev': round(y - m, 1) if (is_num(y) and is_num(m)) else None,
                       'pdf': pdf_names.get(s['name'], '')})
        cols = [['name', '姓名', 'text'], ['tier', '层次', 'tier'],
                ['rank', '校内排名', 'num'], [first, first, 'num'],
                [latest, latest, 'num'], ['d', '较首次增减', 'signed'],
                ['dev', '班均偏离', 'signed']]
        if cfg.get('pdf_report', {}).get('enabled'):
            cols.append(['pdf', '报告', 'action'])
        tables.append({'id': '%s_tier' % sub, 'title': '%s 学生分层' % sub,
                       'sub': '可按层次筛选、姓名搜索，点击列头排序', 'filter': True,
                       'cols': cols, 'rows': tr})
        # 薄弱诊断
        dr = [{'name': d['name'], 'tier': d['tier'], 'weak': d['weak'],
               'rate': round(d['rate'], 1) if is_num(d['rate']) else None,
               'cause': d['cause'], 'advice': d['advice']} for d in meta['diag']]
        tables.append({'id': '%s_diag' % sub, 'title': '%s 薄弱诊断' % sub,
                       'sub': '默认按最弱板块得分率升序', 'filter': True,
                       'cols': [['name', '姓名', 'text'], ['tier', '层次', 'tier'],
                                ['weak', '最弱板块', 'text'], ['rate', '得分率', 'num'],
                                ['cause', '归因', 'adv'], ['advice', '训练建议', 'adv']],
                       'rows': dr})

    pdf_cfg = cfg.get('pdf_report') or {}
    pdf_enabled = bool(pdf_cfg.get('enabled'))
    reports = []
    if pdf_enabled:
        pattern = pdf_cfg.get('filename_pattern', 'report_{name}.pdf')
        for s in analysis['students']:
            reports.append({'name': s['name'],
                            'pdf': pattern.replace('{name}', s['name'])})
    data = {'tierOrder': tier_order(cfg), 'tierColors': tier_colors(cfg),
            'tables': tables, 'reports': reports, 'pdfEnabled': pdf_enabled}

    title = '%s %s 学科能力与分层分析（交互版）' % (cfg['school'], cfg['class_name'])
    subtitle = '%s　|　%s　|　折算为经验估算，仅供预判参考' % (
        cfg.get('term', ''), ' → '.join(exams))
    footer = ('免责声明：层次折算为基于校内排名的经验估算，仅供学情预判参考，'
              '不构成对任何学生学业结果的保证；请以原始成绩与教师研判为准。')
    html = (HTML_TPL.replace('__TITLE__', title)
            .replace('__SUBTITLE__', subtitle)
            .replace('__FOOTER__', footer)
            .replace('__DATA__', json.dumps(data, ensure_ascii=False))
            .replace('__PDF_HIDDEN__', '' if pdf_enabled else 'style="display:none"'))
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)


# ───────────────────────── 主入口 ─────────────────────────

def main():
    ap = argparse.ArgumentParser(description='成绩分析与报告生成引擎')
    ap.add_argument('--config', required=True, help='config.yaml 路径')
    ap.add_argument('--input', required=True, help='成绩 Excel 路径')
    ap.add_argument('--out-dir', required=True, help='输出目录')
    args = ap.parse_args()

    cfg = load_config(args.config)
    students = load_students(cfg, args.input)
    analysis = build_analysis(cfg, students)

    os.makedirs(args.out_dir, exist_ok=True)
    base = '%s_%s' % (cfg['school'], cfg['class_name'])
    docx_path = os.path.join(args.out_dir, base + '_学科能力分析.docx')
    html_path = os.path.join(args.out_dir, base + '_学科能力分析（交互版）.html')

    build_docx(cfg, analysis, docx_path)
    pdf_names = {}
    if cfg.get('pdf_report', {}).get('enabled'):
        pattern = cfg['pdf_report'].get('filename_pattern', 'report_{name}.pdf')
        pdf_names = {s['name']: pattern.replace('{name}', s['name']) for s in students}
    build_html(cfg, analysis, html_path, pdf_names)

    print('[OK] 学生数：%d，学科数：%d' % (len(students), len(cfg['subjects'])))
    print('[OK] %s' % docx_path)
    print('[OK] %s' % html_path)


if __name__ == '__main__':
    main()
