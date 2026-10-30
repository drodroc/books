@echo off
chcp 65001 >nul
cd /d "%~dp0"
set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>nul && set "PY=python"
if not defined PY (
  echo 没有找到 Python。
  echo 请先安装 Python 3.10 或更新版本，安装时勾选 Add python.exe to PATH。
  echo https://www.python.org/downloads/windows/
  pause
  exit /b 1
)
echo 正在准备，第一次会稍慢。
%PY% -m pip install -r "%~dp0requirements.txt"
if errorlevel 1 (
  echo 安装依赖失败。
  pause
  exit /b 1
)
%PY% "%~dp0server.py"
echo.
echo 软件已退出。
pause
