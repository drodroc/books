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
%PY% -c "import openpyxl" >nul 2>nul && goto launch
echo 正在准备，第一次会稍慢。连不上官网时会改用国内镜像，再不行就用软件自带的安装包。
%PY% -m pip install --timeout 30 --retries 2 -i https://pypi.tuna.tsinghua.edu.cn/simple --trusted-host pypi.tuna.tsinghua.edu.cn -r "%~dp0requirements.txt"
if errorlevel 1 (
  echo 清华镜像没连上，换阿里云。
  %PY% -m pip install --timeout 30 --retries 2 -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com -r "%~dp0requirements.txt"
)
if errorlevel 1 (
  echo 镜像没连上，改用软件自带的安装包。
  %PY% -m pip install --no-index --find-links="%~dp0wheels" -r "%~dp0requirements.txt"
)
if errorlevel 1 (
  echo 安装依赖失败。请检查网络后再双击一次。
  pause
  exit /b 1
)
:launch
%PY% "%~dp0server.py"
echo.
echo 软件已退出。
pause
