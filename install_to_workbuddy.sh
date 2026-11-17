#!/bin/bash
# (C) E.D.STUDIO · 成绩分析报告工具
# 安装脚本：把本工具注册为 WorkBuddy 技能（用户级），并准备运行环境。
#
# 用法（三选一）：
#   bash install_to_workbuddy.sh          # 从仓库/解压目录直接安装
#   ./install_to_workbuddy.sh             # 同上（若已给可执行位）
#
# 做完这四件事：
#   1. 复制到 ~/.workbuddy/skills/score-analysis-report-engine/
#   2. 修复一键脚本的可执行位（git clone / 某些解压方式会丢权限）
#   3. 建虚拟环境并装依赖（openpyxl / python-docx / PyYAML）
#   4. 打印接下来怎么用

set -u

SRC="$(cd "$(dirname "$0")" && pwd)"
SKILL_NAME="score-analysis-report-engine"
DEST_DIR="$HOME/.workbuddy/skills/$SKILL_NAME"

echo "============================================================"
echo "  成绩分析报告工具 · 安装到 WorkBuddy"
echo "  by E.D.STUDIO"
echo "============================================================"

# ── 1. 复制到技能目录 ──
if [ ! -d "$HOME/.workbuddy/skills" ]; then
  echo "[1/4] 创建技能目录 $HOME/.workbuddy/skills"
  mkdir -p "$HOME/.workbuddy/skills" || { echo "[错误] 无权限创建目录"; exit 1; }
fi

if [ -e "$DEST_DIR" ]; then
  echo "[1/4] 目标已存在，先备份为 ${SKILL_NAME}.bak.$(date +%Y%m%d%H%M%S)"
  mv "$DEST_DIR" "${DEST_DIR}.bak.$(date +%Y%m%d%H%M%S)" || { echo "[错误] 备份失败"; exit 1; }
fi

echo "[1/4] 复制到 $DEST_DIR"
mkdir -p "$DEST_DIR" || { echo "[错误] 创建目标目录失败"; exit 1; }
# 复制全部内容（排除 .git 与临时文件）
(cd "$SRC" && tar cf - --exclude='.git' --exclude='.venv' --exclude='__pycache__' .) | (cd "$DEST_DIR" && tar xf -) \
  || { echo "[错误] 复制失败"; exit 1; }

# ── 2. 修复可执行位 ──
echo "[2/4] 修复一键脚本可执行位"
chmod +x "$DEST_DIR/一键运行_Mac.command" 2>/dev/null
chmod +x "$DEST_DIR/一键运行_Win.bat" 2>/dev/null
chmod +x "$DEST_DIR/scripts/"*.py 2>/dev/null

# ── 3. 准备运行环境 ──
PY=""
for c in python3 python; do
  if command -v "$c" >/dev/null 2>&1; then PY="$c"; break; fi
done
if [ -z "$PY" ]; then
  echo "[3/4] [跳过] 未检测到 Python 3，请先从 https://www.python.org/downloads/ 安装"
else
  echo "[3/4] 准备运行环境（首次需联网一次）"
  if [ ! -x "$DEST_DIR/.venv/bin/python" ]; then
    "$PY" -m venv "$DEST_DIR/.venv" 2>/dev/null || echo "      [提示] 创建虚拟环境失败，可改用系统 Python"
  fi
  VP="$DEST_DIR/.venv/bin/python"
  if [ -x "$VP" ]; then
    if ! "$VP" -c "import openpyxl, docx, yaml" >/dev/null 2>&1; then
      "$VP" -m pip install --quiet --upgrade pip >/dev/null 2>&1
      "$VP" -m pip install --quiet openpyxl python-docx PyYAML >/dev/null 2>&1
    fi
    if "$VP" -c "import openpyxl, docx, yaml" >/dev/null 2>&1; then
      echo "      ✅ 依赖就绪"
    else
      echo "      [提示] 依赖安装未完全成功，可在联网后重跑本脚本"
    fi
  fi
fi

# ── 4. 收尾提示 ──
cat <<EOF

[4/4] ✅ 安装完成

接下来怎么用：

  A. 在 WorkBuddy 对话框里直接说
       「用成绩分析报告引擎，把这份成绩 Excel 分析成学情报告」
     工具会自动识别并引导你完成列确认与报告生成。

  B. 想先看效果（不用配置任何东西），双击
       $DEST_DIR/一键运行_Mac.command
     会引导你选 Excel、自动识别列、并生成 Word + 交互网页。

  C. 网页版（浏览器里操作）：双击
       $DEST_DIR/一键运行_Mac.command 后选「网页版」，
     或手动执行：
       $DEST_DIR/.venv/bin/python $DEST_DIR/scripts/web_app.py

说明：所有数据只在本机处理，工具不上传任何学生信息。

============================================================
EOF
