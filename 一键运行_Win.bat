@echo off
rem (C) E.D.STUDIO · 成绩分析报告工具
chcp 65001 >nul
cd /d "%~dp0"
title 成绩分析与报告生成工具

echo ============================================================
echo   成绩分析与报告生成工具
 echo   by E.D.STUDIO
echo ============================================================

where python >nul 2>&1
if %errorlevel%==0 (set PY=python) else (
  where py >nul 2>&1
  if %errorlevel%==0 (set PY=py) else (
    echo [错误] 未检测到 Python 3。请先安装 Python 3 并勾选 "Add Python to PATH"。
    echo        下载地址：https://www.python.org/downloads/
    pause
    exit /b 1
  )
)

if not exist ".venv\Scripts\python.exe" (
  echo 首次运行：正在创建运行环境…
  %PY% -m venv .venv
  if errorlevel 1 ( echo [错误] 创建环境失败 & pause & exit /b 1 )
)

.venv\Scripts\python.exe -c "import openpyxl, docx, yaml" >nul 2>&1
if errorlevel 1 (
  echo 首次运行：正在安装依赖（仅需联网一次，请稍候）…
  .venv\Scripts\python.exe -m pip install --quiet --upgrade pip
  .venv\Scripts\python.exe -m pip install --quiet openpyxl python-docx PyYAML
)

cls
.venv\Scripts\python.exe scripts\wizard.py
