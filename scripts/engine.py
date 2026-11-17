#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (C) E.D.STUDIO · 成绩分析报告工具
"""共享生成核心：本地向导与云端网页版都调用本模块的 generate()。

参数：
  cfg       —— 配置 dict（与 config.yaml 同结构）
  input_path—— 成绩 Excel 的本地路径
  out_dir   —— 输出目录
返回：(docx_path, html_path, student_count)
"""
import os
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
if HERE not in sys.path:
    sys.path.insert(0, HERE)

import build_report as E


def validate(cfg):
    required = ['school', 'class_name', 'subjects', 'input', 'tier_rules', 'exam_chain']
    miss = [k for k in required if k not in cfg]
    if miss:
        raise ValueError('配置缺少必填项：%s' % ', '.join(miss))
    if not cfg['input'].get('columns', {}).get('name'):
        raise ValueError('配置缺少 input.columns.name')
    return True


def generate(cfg, input_path, out_dir):
    validate(cfg)
    if not os.path.exists(input_path):
        raise FileNotFoundError(input_path)

    students = E.load_students(cfg, input_path)
    if not students:
        raise ValueError('未从 Excel 读取到任何学生，请检查表头与列映射是否正确。')
    analysis = E.build_analysis(cfg, students)

    os.makedirs(out_dir, exist_ok=True)
    base = '%s_%s' % (cfg['school'], cfg['class_name'])
    docx_path = os.path.join(out_dir, base + '_学科能力分析.docx')
    html_path = os.path.join(out_dir, base + '_学科能力分析（交互版）.html')

    E.build_docx(cfg, analysis, docx_path)
    pdf_names = {}
    if cfg.get('pdf_report', {}).get('enabled'):
        pattern = cfg['pdf_report'].get('filename_pattern', 'report_{name}.pdf')
        pdf_names = {s['name']: pattern.replace('{name}', s['name']) for s in students}
    E.build_html(cfg, analysis, html_path, pdf_names)
    return docx_path, html_path, len(students)
