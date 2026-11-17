#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# (C) E.D.STUDIO · 成绩分析报告工具
"""生成与 config.yaml 列契约完全对齐的【全虚构】示例 Excel。

所有姓名均为「示例学生01」式占位符，成绩由固定种子随机生成，
不含任何真实个人数据。用于先跑通流程。
"""
import argparse
import os
import random
import sys

try:
    import yaml
    import openpyxl
except ImportError as exc:
    sys.stderr.write('缺少依赖：%s\n请 pip install openpyxl PyYAML\n' % exc)
    sys.exit(2)


FAMILY = ['示例', '演示', '测试', '样例']
GIVEN = ['学生', '同学', '样本', '样例']


def synthetic_name(i):
    return '%s%s%02d' % (FAMILY[i % len(FAMILY)], GIVEN[i % len(GIVEN)], i + 1)


def main():
    ap = argparse.ArgumentParser(description='生成全虚构示例 Excel')
    ap.add_argument('--config', required=True)
    ap.add_argument('--out', default='sample.xlsx')
    ap.add_argument('--count', type=int, default=30, help='虚构学生数，默认 30')
    args = ap.parse_args()

    with open(args.config, encoding='utf-8') as f:
        cfg = yaml.safe_load(f)
    cols = cfg['input']['columns']
    exams = cfg['exam_chain']
    rng = random.Random(20260101)

    header = [cols.get('name', '姓名'), cols.get('sid', '学号'), cols.get('rank', '校内排名')]
    for sub, exams_map in (cols.get('totals') or {}).items():
        for exam in exams:
            header.append((exams_map or {}).get(exam, '%s%s' % (sub, exam)))
    for sub, exams_map in (cols.get('section_rates') or {}).items():
        for exam in exams:
            for blk in (exams_map.get(exam) or {}):
                header.append(exams_map[exam][blk])

    rows = []
    for i in range(args.count):
        row = {}
        row[cols.get('name', '姓名')] = synthetic_name(i)
        row[cols.get('sid', '学号')] = 'S%05d' % (i + 1)
        # 让虚构排名跨越配置的各档，便于演示筛选（最后一名落入兜底档）
        max_tier = None
        for rule in cfg['tier_rules']:
            if isinstance(rule.get('max'), (int, float)):
                max_tier = max(max_tier or 0, rule['max'])
        span = int(max_tier * 1.3) if max_tier else args.count
        row[cols.get('rank', '校内排名')] = 1 + round(i * span / max(args.count - 1, 1))
        for sub, exams_map in (cols.get('totals') or {}).items():
            full = (cfg['subjects'].get(sub) or {}).get('full', 150)
            for exam in exams:
                row[(exams_map or {}).get(exam, '%s%s' % (sub, exam))] = round(
                    rng.uniform(full * 0.55, full * 0.95), 1)
        for sub, exams_map in (cols.get('section_rates') or {}).items():
            for exam in exams:
                for blk, colname in (exams_map.get(exam) or {}).items():
                    row[colname] = round(rng.uniform(45, 95), 1)
        rows.append([row.get(h) for h in header])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '示例数据'
    ws.append(header)
    for r in rows:
        ws.append(r)
    wb.save(args.out)
    print('[OK] 已生成 %d 名虚构学生，列数 %d' % (args.count, len(header)))
    print('[OK] %s' % os.path.abspath(args.out))


if __name__ == '__main__':
    main()
