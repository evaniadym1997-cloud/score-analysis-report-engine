@echo off
REM (C) E.D.STUDIO · 成绩分析报告工具
REM 安装脚本（Windows）：把本工具注册为 WorkBuddy 技能（用户级），并准备运行环境。
REM 用法：双击本文件，或在命令行执行 install_to_workbuddy.bat

setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

set "SRC=%~dp0"
set "SKILL_NAME=score-analysis-report-engine"
set "DEST=%USERPROFILE%\.workbuddy\skills\%SKILL_NAME%"

echo ============================================================
echo   成绩分析报告工具 · 安装到 WorkBuddy
echo   by E.D.STUDIO
echo ============================================================

REM ── 1. 复制到技能目录 ──
if not exist "%USERPROFILE%\.workbuddy\skills" (
  echo [1/4] 创建技能目录
  mkdir "%USERPROFILE%\.workbuddy\skills" >nul 2>&1
)
if exist "%DEST%" (
  echo [1/4] 目标已存在，先备份
  if exist "%DEST%.bak" rd /s /q "%DEST%.bak"
  move "%DEST%" "%DEST%.bak" >nul 2>&1
)
echo [1/4] 复制到 %DEST%
mkdir "%DEST%" >nul 2>&1
xcopy "%SRC%" "%DEST%\" /E /H /Y /Q /EXCLUDE:.git >nul 2>&1
if exist "%DEST%\.venv" rd /s /q "%DEST%\.venv" >nul 2>&1

REM ── 2. 准备运行环境 ──
set "PY="
where py >nul 2>&1 && set "PY=py -3"
if not defined PY ( where python >nul 2>&1 && set "PY=python" )

if not defined PY (
  echo [3/4] [跳过] 未检测到 Python 3，请先从 https://www.python.org/downloads/ 安装
) else (
  echo [3/4] 准备运行环境（首次需联网一次）
  if not exist "%DEST%\.venv\Scripts\python.exe" (
    %PY% -m venv "%DEST%\.venv" >nul 2>&1
  )
  if exist "%DEST%\.venv\Scripts\python.exe" (
    "%DEST%\.venv\Scripts\python.exe" -c "import openpyxl, docx, yaml" >nul 2>&1
    if errorlevel 1 (
      "%DEST%\.venv\Scripts\python.exe" -m pip install --quiet --upgrade pip >nul 2>&1
      "%DEST%\.venv\Scripts\python.exe" -m pip install --quiet openpyxl python-docx PyYAML >nul 2>&1
    )
    "%DEST%\.venv\Scripts\python.exe" -c "import openpyxl, docx, yaml" >nul 2>&1
    if errorlevel 1 (
      echo       [提示] 依赖安装未完全成功，可在联网后重跑本脚本
    ) else (
      echo       [OK] 依赖就绪
    )
  )
)

echo.
echo [4/4] (OK) 安装完成
echo.
echo 接下来怎么用：
echo.
echo   A. 在 WorkBuddy 对话框里直接说：
echo      「用成绩分析报告引擎，把这份成绩 Excel 分析成学情报告」
echo.
echo   B. 想先看效果，双击：
echo      %DEST%\一键运行_Win.bat
echo.
echo   C. 网页版：双击 一键运行_Win.bat 后选「网页版」
echo.
echo 说明：所有数据只在本机处理，工具不上传任何学生信息。
echo ============================================================
pause
