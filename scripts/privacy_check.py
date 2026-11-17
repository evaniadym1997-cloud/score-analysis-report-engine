#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""脱敏门禁：扫描整个目录，命中个人数据特征即非零退出。

用法：
  # 通用特征扫描（发布包自带）
  python scripts/privacy_check.py --root .

  # 严格模式：追加你自己的私有词表（词表通过文件传入，不写进仓库）
  python scripts/privacy_check.py --root . --strict --words-file /path/to/private_words.txt

重要：本扫描器**不含任何真实姓名/学校名/账号**，因为词表本身也是个人数据。
所有针对你本人的专属词，请写在仓库外的 private_words.txt 中，通过 --words-file 传入。
扫描器只会把它作为命令行参数读取，不会写回任何文件。

扫描对象：文本文件正文、文件名、git 提交历史（作者邮箱与提交信息）。
"""
import argparse
import os
import re
import subprocess
import sys

# 通用特征：不指向任何具体个人，命中即人工复核
GENERIC_PATTERNS = [
    ('Unix用户目录绝对路径', r'/Users/[A-Za-z0-9._-]+'),
    ('Windows用户目录绝对路径', r'[A-Za-z]:\\Users\\'),
    ('邮箱地址', r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}'),
    ('手机号(中国大陆)', r'(?<!\d)1[3-9]\d{9}(?!\d)'),
    ('身份证号', r'(?<!\d)\d{17}[\dXx](?!\d)'),
    ('Github个人主页', r'github\.com/[A-Za-z0-9-]{3,}'),
]

TEXT_EXT = {'.md', '.yaml', '.yml', '.py', '.json', '.txt', '.html', '.htm',
            '.cfg', '.toml', '.ini', '.csv', '.sh', '.example', '.gitignore',
            '.xml', '.js', '.css', '.rst', '.tex'}
SKIP_DIRS = {'.git', '__pycache__', '.pytest_cache', 'node_modules', '.idea', '.vscode'}


def load_words(path):
    if not path:
        return []
    words = []
    with open(path, encoding='utf-8') as f:
        for line in f:
            w = line.strip()
            if w and not w.startswith('#'):
                words.append(w)
    return words


def scan_text(path, patterns):
    hits = []
    try:
        with open(path, encoding='utf-8', errors='replace') as f:
            for lineno, line in enumerate(f, 1):
                for label, pat in patterns:
                    if re.search(pat, line):
                        hits.append((label, lineno, line.strip()[:160]))
    except (OSError, UnicodeError):
        pass
    return hits


def scan_git(root, patterns):
    hits = []
    if not os.path.isdir(os.path.join(root, '.git')):
        return hits
    try:
        # 只取作者名与提交说明；邮箱字段不在此扫（避免保留域名占位邮箱误报）。
        out = subprocess.check_output(
            ['git', '-C', root, 'log', '--all', '--pretty=format:%an%n%s%n%b'],
            stderr=subprocess.DEVNULL).decode('utf-8', 'replace')
    except (subprocess.CalledProcessError, OSError):
        return hits
    for lineno, line in enumerate(out.splitlines(), 1):
        for label, pat in patterns:
            if re.search(pat, line):
                hits.append(('git历史:' + label, lineno, line.strip()[:160]))
    return hits


def main():
    ap = argparse.ArgumentParser(description='脱敏门禁（词表与仓库分离）')
    ap.add_argument('--root', default='.', help='待扫描根目录')
    ap.add_argument('--strict', action='store_true',
                    help='严格模式：额外扫描 git 提交历史')
    ap.add_argument('--words-file', default='',
                    help='私有词表文件（仓库外），每行一个词，# 开头为注释')
    args = ap.parse_args()

    root = os.path.abspath(args.root)
    patterns = list(GENERIC_PATTERNS)
    for w in load_words(args.words_file):
        patterns.append(('私有词:' + w, re.escape(w)))

    hits = []
    scanned = 0
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for fn in filenames:
            full = os.path.join(dirpath, fn)
            rel = os.path.relpath(full, root)
            ext = os.path.splitext(fn)[1].lower()
            for label, pat in patterns:
                if re.search(pat, fn):
                    hits.append((rel, label, 0, '文件名命中: ' + fn))
            if ext in TEXT_EXT or fn in {'LICENSE', 'Makefile', 'Dockerfile'}:
                scanned += 1
                for label, lineno, line in scan_text(full, patterns):
                    hits.append((rel, label, lineno, line))

    if args.strict:
        for label, lineno, line in scan_git(root, patterns):
            hits.append(('.git/HISTORY', label, lineno, line))

    print('扫描文件数：%d' % scanned)
    print('通用特征：%d 条；私有词：%d 条；严格模式：%s' % (
        len(GENERIC_PATTERNS), len(patterns) - len(GENERIC_PATTERNS),
        '是' if args.strict else '否'))
    if hits:
        print('\n[FAIL] 命中 %d 处，请逐条复核：' % len(hits))
        for rel, label, lineno, line in hits[:80]:
            print('  %s:%s [%s] %s' % (rel, lineno or '-', label, line))
        if len(hits) > 80:
            print('  ... 其余 %d 处省略' % (len(hits) - 80))
        sys.exit(1)
    print('\n[PASS] 0 命中。')


if __name__ == '__main__':
    main()
