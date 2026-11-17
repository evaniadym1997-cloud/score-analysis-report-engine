#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (C) E.D.STUDIO · 成绩分析报告工具
"""本地一键向导：选 Excel → 自动识别列 → 生成配置 → 调用引擎出报告。

设计目标：不懂命令行的老师双击即可用，全程只见 3 类提问。
隐私：所有数据只在本机处理，本脚本不联网、不上传任何内容。
"""
import os
import sys
import subprocess
import random
import string

try:
    import yaml
    import openpyxl
except ImportError as exc:
    sys.stderr.write('缺少依赖：%s\n请先运行安装脚本，或手动 pip install openpyxl python-docx PyYAML\n' % exc)
    sys.exit(2)

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
ENGINE = os.path.join(HERE, 'build_report.py')


# ───────────────────────── 小工具 ─────────────────────────

def ask(prompt, default=''):
    tip = (' [%s]' % default) if default else ''
    try:
        val = input('%s%s: ' % (prompt, tip)).strip()
    except EOFError:
        val = ''
    return val if val else default


def yes(prompt, default=True):
    d = 'Y/n' if default else 'y/N'
    v = ask(prompt + ' (%s)' % d, 'y' if default else 'n')
    return v.lower().startswith('y')


def pick_file():
    """弹出系统文件选择框；不可用时回退为手工输入路径。"""
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        path = filedialog.askopenfilename(
            title='请选择成绩 Excel 文件',
            filetypes=[('Excel 表格', '*.xlsx *.xls'), ('所有文件', '*.*')])
        root.destroy()
        if path:
            return path
    except Exception:
        pass
    print('\n（未能打开文件选择框，请手动粘贴文件路径）')
    while True:
        p = ask('Excel 文件完整路径', '').strip().strip('"')
        if p and os.path.exists(p):
            return p
        print('  找不到该文件，请检查路径后重试。')


def read_header(path):
    """读取 Excel 表头：在前 12 行里挑非空文本单元格最多的一行。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i >= 12:
            break
        rows.append(list(row))
    best_i, best_n = 0, -1
    for i, row in enumerate(rows):
        n = sum(1 for c in row if isinstance(c, str) and c.strip())
        if n > best_n:
            best_i, best_n = i, n
    header = [str(c).strip() if c is not None else '' for c in rows[best_i]]
    return header, best_i


def norm(s):
    return ''.join(ch for ch in str(s) if ch.strip() and ch not in '（）()【】[]:： 　')


def guess(headers, keys, exclude=()):
    """按关键字顺序在表头里精确匹配（ki 必须完整出现在表头里），返回列索引。

    keys 按顺序给；先命中的组优先级更高。这样「语文期中」不会被宽泛的「语文」
    抢先匹配到「语文期初」。
    """
    for k in keys:
        if not k:
            continue
        nk = norm(k)
        for i, h in enumerate(headers):
            if i in exclude:
                continue
            if nk == norm(h):
                return i
    # 精确不中再做包含匹配（仍按 keys 顺序），并要求关键词长度 >= 2
    for k in keys:
        if not k:
            continue
        nk = norm(k)
        if len(nk) < 2:
            continue
        for i, h in enumerate(headers):
            if i in exclude:
                continue
            if nk in norm(h):
                return i
    return None


def choose_column(headers, label, guessed, used, optional=False):
    """展示自动识别结果让用户确认；不对就列出来让用户选编号。"""
    if guessed is not None:
        if yes('  %s → 「%s」是否正确' % (label, headers[guessed]), True):
            used.add(guessed)
            return headers[guessed]
    print('  可选列：')
    for i, h in enumerate(headers):
        if h and i not in used:
            print('    %2d) %s' % (i + 1, h))
    if optional:
        if not yes('  是否提供「%s」列（跳过则留空）' % label, False):
            return ''
    while True:
        v = ask('  %s 请选择列号（直接回车留空）' % label, '')
        if not v:
            return ''
        if v.isdigit() and 1 <= int(v) <= len(headers) and (int(v) - 1) not in used:
            idx = int(v) - 1
            used.add(idx)
            return headers[idx]
        print('  无效选择，请重新输入。')


DEFAULT_SECTIONS = {
    '语文': {'阅读': 40, '古诗文': 35, '语言文字': 25, '作文': 50},
    '数学': {'选择': 40, '填空': 20, '解答': 90},
    '英语': {'听力': 30, '阅读': 40, '完形': 15, '语法': 15, '写作': 50},
}


def random_color():
    palette = ['#FCE4D6', '#E2EFDA', '#DDEBF7', '#FFF2CC', '#E4DFEC', '#F8CBAD']
    return random.choice(palette)


# ───────────────────────── 主流程 ─────────────────────────

def main():
    print('=' * 60)
    print('  成绩分析与报告生成工具 · 本地一键向导')
    print('=' * 60)
    print('  隐私说明：所有成绩数据只在本机处理，不会上传到任何服务器。')
    print('=' * 60)

    # 1. 基本信息
    print('\n【第 1 步】填写基本信息')
    school = ask('学校名称', '示例学校')
    class_name = ask('班级名称', '示例班级')
    term = ask('学期描述', '本学期')
    analyst = ask('署名/来源（可留空）', '教务处')

    # 2. 选文件 + 读表头
    print('\n【第 2 步】选择成绩 Excel')
    path = pick_file()
    print('  已选择：%s' % path)
    header, hrow = read_header(path)
    header = [h for h in header if h]
    print('  识别到表头（第 %d 行），共 %d 列：' % (hrow + 1, len(header)))
    for i, h in enumerate(header):
        print('    %2d) %s' % (i + 1, h))

    # 3. 考试名称
    print('\n【第 3 步】考试名称（两考：期初 + 期中）')
    exam1 = ask('第一次考试名称', '期初')
    exam2 = ask('第二次考试名称', '期中')
    exams = [exam1, exam2]

    # 4. 科目
    print('\n【第 4 步】学科（用逗号分隔，默认：语文,数学,英语）')
    subj_raw = ask('学科列表', '语文,数学,英语')
    subjects = [s.strip() for s in subj_raw.replace('，', ',').split(',') if s.strip()]
    if not subjects:
        subjects = ['语文', '数学', '英语']

    # 5. 列映射（自动识别 + 确认）
    print('\n【第 5 步】确认列对应关系（已自动识别，回车即采用）')
    used = set()

    name_col = choose_column(header, '姓名列', guess(header, ['姓名', '学生姓名', '学生', 'name']), used)
    sid_col = choose_column(header, '学号列', guess(header, ['学号', '考号', '编号', 'id']), used, optional=True)
    rank_col = choose_column(header, '校内排名列（用于分档）', guess(header, ['校内排名', '校排名', '排名', '位次']), used)

    totals = {}
    for sub in subjects:
        totals[sub] = {}
        for exam in exams:
            # 必须「科目名 + 考试名」同时出现；按 科目+考试 / 考试+科目 顺序精确优先
            g = guess(header, [sub + exam, exam + sub], exclude=set(used))
            totals[sub][exam] = choose_column(header, '%s·%s 总分列' % (sub, exam), g, used)

    # 6. 分档规则
    print('\n【第 6 步】分档规则（按校内排名从高到低）')
    tiers = []
    for i in range(3):
        t = ask('  第 %d 档名称' % (i + 1), ['拔尖层', '稳步层', '攻坚层'][i])
        mx = ask('    排名上限（数字，最后一档直接回车）', ['50', '150', ''][i])
        tiers.append({'max': int(mx) if mx.isdigit() else (None if i == 2 else 50 * (i + 1)),
                      'tier': t, 'color': random_color(),
                      'desc': '%s，按校内排名经验折算，仅供参考' % t})

    # 7. 生成配置
    subj_cfg = {}
    for sub in subjects:
        subj_cfg[sub] = {'full': 150,
                         'sections': DEFAULT_SECTIONS.get(sub, {'总分': 150})}
    cfg = {
        'school': school, 'class_name': class_name, 'term': term, 'update_date': '',
        'analyst': analyst, 'exam_chain': exams, 'subjects': subj_cfg,
        'input': {'columns': {
            'name': name_col, 'sid': sid_col or '学号', 'rank': rank_col,
            'totals': totals,
        }},
        'tier_rules': tiers, 'missing_tier': '数据不全',
        'weak_reasons': {},
        'default_reason': {'cause': '该板块基础巩固不足', 'advice': '回归基础，按知识点清单逐条过关'},
        'school_recommend': {},
    }
    cfg_path = os.path.join(ROOT, 'config.yaml')
    with open(cfg_path, 'w', encoding='utf-8') as f:
        yaml.safe_dump(cfg, f, allow_unicode=True, sort_keys=False)
    print('\n  配置已生成：%s' % cfg_path)

    # 8. 调用引擎
    print('\n【第 7 步】生成报告…')
    out_dir = os.path.join(ROOT, '输出结果')
    cmd = [sys.executable, ENGINE, '--config', cfg_path, '--input', path, '--out-dir', out_dir]
    proc = subprocess.run(cmd)
    if proc.returncode != 0:
        print('\n[失败] 生成过程出错，请把上面的提示发给技术支持。')
        input('\n按回车键退出…')
        sys.exit(1)

    print('\n' + '=' * 60)
    print('  完成！报告已保存到：')
    print('  %s' % out_dir)
    print('=' * 60)
    # 尽量自动打开输出目录
    try:
        if sys.platform == 'darwin':
            subprocess.run(['open', out_dir])
        elif sys.platform == 'win32':
            os.startfile(out_dir)
        elif sys.platform.startswith('linux'):
            subprocess.run(['xdg-open', out_dir])
    except Exception:
        pass
    input('\n按回车键退出…')


if __name__ == '__main__':
    main()
