#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (C) E.D.STUDIO · 成绩分析报告工具
"""云端网页版：上传 Excel → 确认列 → 一键生成并下载 docx + 交互 HTML。

隐私：上传文件只写入临时目录，生成后立即读取为下载字节并随会话销毁；
服务器不保留任何学生数据，页面已明示。
"""
import io
import os
import sys
import tempfile

import streamlit as st
import openpyxl

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from wizard import guess
from engine import generate

st.set_page_config(page_title='成绩分析与报告生成', page_icon='📊', layout='wide')

st.markdown(
    '<style>'
    '.stButton>button{background-color:#3CBFAE;color:white;border-radius:10px;'
    'font-size:16px;padding:10px 24px;border:none}'
    '.stButton>button:hover{background-color:#2FA894;color:white}'
    'h1{color:#2FA894}'
    '</style>', unsafe_allow_html=True)

st.title('📊 成绩分析与报告生成')
st.caption('上传成绩 Excel，自动生成 Word 分析报告 + 交互式 HTML（筛选 / 排序 / 分档）')

with st.expander('🔒 数据隐私说明（请先阅读）', expanded=False):
    st.write(
        '• 上传的成绩表**仅用于本次生成**，写入临时文件后立即读取结果并随会话销毁；\n'
        '• 服务器**不保存、不记录**任何学生姓名与成绩，页面刷新或关闭后数据即消失；\n'
        '• 若涉及敏感成绩，建议使用本地一键包（数据完全不出本机）。')


def read_header(data):
    """从上传字节里读表头：前 12 行挑非空文本最多的一行。"""
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i >= 12:
            break
        rows.append(list(row))
    best = max(range(len(rows)),
               key=lambda i: sum(1 for c in rows[i] if isinstance(c, str) and c.strip()))
    return [str(c).strip() for c in rows[best] if c is not None and str(c).strip()], best + 1


uploaded = st.file_uploader('① 上传成绩 Excel（.xlsx）', type=['xlsx'])
if uploaded is None:
    st.info('请先上传 Excel 表格。表头建议包含：姓名、学号、校内排名，以及各科各次考试总分列。')
    st.stop()

data = uploaded.getvalue()
try:
    header, hrow = read_header(data)
except Exception as e:
    st.error('读取 Excel 失败：%s' % e)
    st.stop()
if not header:
    st.error('未识别到表头，请检查表格内容。')
    st.stop()

st.success('已识别表头（第 %d 行），共 %d 列。' % (hrow, len(header)))

# 基本信息
st.subheader('② 填写基本信息')
c1, c2, c3 = st.columns(3)
school = c1.text_input('学校名称', '某学校')
class_name = c2.text_input('班级名称', '某班级')
term = c3.text_input('学期', '本学期')

st.subheader('③ 考试与学科')
ex1, ex2 = st.columns(2)
exam1 = ex1.text_input('第一次考试', '期初')
exam2 = ex2.text_input('第二次考试', '期中')
exams = [exam1, exam2]
subj_raw = st.text_input('学科（逗号分隔）', '语文,数学,英语')
subjects = [s.strip() for s in subj_raw.replace('，', ',').split(',') if s.strip()] or ['语文', '数学', '英语']

# 列映射
st.subheader('④ 确认列对应关系（已自动识别，可修改）')
used = set()
name_col = st.selectbox('姓名列', header, index=max(0, guess(header, ['姓名', '学生姓名', '学生', 'name']) or 0))
sid_g = guess(header, ['学号', '考号', '编号', 'id'])
sid_col = st.selectbox('学号列', ['（无）'] + header, index=(sid_g + 1) if sid_g is not None else 0)
rank_col = st.selectbox('校内排名列（用于分档）', header,
                        index=max(0, guess(header, ['校内排名', '校排名', '排名', '位次']) or 0))

totals = {}
for si, sub in enumerate(subjects):
    cols = st.columns(len(exams))
    totals[sub] = {}
    for ei, exam in enumerate(exams):
        g = guess(header, [sub + exam, exam + sub])
        totals[sub][exam] = cols[ei].selectbox(
            '%s·%s' % (sub, exam), header, index=max(0, g or 0), key='tot_%d_%d' % (si, ei))

# 分档
st.subheader('⑤ 分档规则（按校内排名从高到低）')
t1, t2, t3 = st.columns(3)
m1 = t1.number_input('第1档 排名上限', 1, 100000, 50)
m2 = t2.number_input('第2档 排名上限', 1, 100000, 150)
tiers = [
    {'max': m1, 'tier': '拔尖层', 'color': '#FCE4D6', 'desc': '校内前 %d 名，经验折算仅供参考' % m1},
    {'max': m2, 'tier': '稳步层', 'color': '#E2EFDA', 'desc': '校内 %d–%d 名，经验折算仅供参考' % (m1, m2)},
    {'max': None, 'tier': '攻坚层', 'color': '#DDEBF7', 'desc': '其余学生，建议夯实基础'},
]

st.markdown('---')
if st.button('🚀 生成报告'):
    cfg = {
        'school': school, 'class_name': class_name, 'term': term, 'update_date': '',
        'analyst': '', 'exam_chain': exams,
        'subjects': {s: {'full': 150,
                         'sections': {'总分': 150}} for s in subjects},
        'input': {'columns': {
            'name': name_col,
            'sid': sid_col if sid_col != '（无）' else '学号',
            'rank': rank_col,
            'totals': totals,
        }},
        'tier_rules': tiers, 'missing_tier': '数据不全',
        'weak_reasons': {},
        'default_reason': {'cause': '该板块基础巩固不足', 'advice': '回归基础，按知识点清单逐条过关'},
        'school_recommend': {},
    }
    try:
        with st.spinner('正在生成，请稍候…'):
            with tempfile.TemporaryDirectory() as td:
                inp = os.path.join(td, 'input.xlsx')
                with open(inp, 'wb') as f:
                    f.write(data)
                docx_path, html_path, n = generate(cfg, inp, td)
                with open(docx_path, 'rb') as f:
                    docx_bytes = f.read()
                with open(html_path, 'rb') as f:
                    html_bytes = f.read()
                base = '%s_%s' % (school, class_name)
        st.success('✅ 生成完成！共处理 %d 名学生。' % n)
        dcol, hcol = st.columns(2)
        dcol.download_button('📄 下载 Word 分析报告', docx_bytes,
                             file_name=base + '_学科能力分析.docx',
                             mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        hcol.download_button('🌐 下载交互式 HTML', html_bytes,
                             file_name=base + '_学科能力分析（交互版）.html',
                             mime='text/html')
        st.caption('临时文件已随本次生成销毁，服务器未保留任何学生数据。')
    except Exception as e:
        st.error('生成失败：%s' % e)
        st.caption('请检查列映射是否正确（尤其是各科各次考试的总分列不要串列）。')
