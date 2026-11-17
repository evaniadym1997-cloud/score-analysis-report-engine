#!/bin/bash
# (C) E.D.STUDIO · 成绩分析报告工具
# 成绩分析报告工具 · Mac 双击启动脚本
# 首次运行会自动创建虚拟环境并安装依赖（需联网一次），之后离线可用。
cd "$(dirname "$0")" || exit 1

echo "============================================================"
echo "  成绩分析与报告生成工具"
  echo "  by E.D.STUDIO"
echo "============================================================"

PY=""
for c in python3 python; do
  if command -v "$c" >/dev/null 2>&1; then PY="$c"; break; fi
done
if [ -z "$PY" ]; then
  echo "[错误] 未检测到 Python 3。请先安装 Python 3（https://www.python.org/downloads/）后重试。"
  read -r -p "按回车键退出…"; exit 1
fi

VENV=".venv"
if [ ! -x "$VENV/bin/python" ]; then
  echo "首次运行：正在创建运行环境…"
  "$PY" -m venv "$VENV" || { echo "[错误] 创建环境失败"; read -r -p "按回车退出…"; exit 1; }
fi

if ! "$VENV/bin/python" -c "import openpyxl, docx, yaml" >/dev/null 2>&1; then
  echo "首次运行：正在安装依赖（仅需联网一次，请稍候）…"
  "$VENV/bin/python" -m pip install --quiet --upgrade pip
  "$VENV/bin/python" -m pip install --quiet openpyxl python-docx PyYAML
fi

clear
"$VENV/bin/python" scripts/wizard.py
