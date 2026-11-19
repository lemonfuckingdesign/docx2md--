@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ============================================
echo   docx 转 Dify 知识库 Markdown
echo ============================================
echo.

REM ---------- 第 1 步：自动查找可用的 Python ----------
set "PY="

REM 1) 先试 PATH 里的命令
python --version >nul 2>&1 && set "PY=python"
if not defined PY (
    py -3 --version >nul 2>&1 && set "PY=py -3"
)
if not defined PY (
    python3 --version >nul 2>&1 && set "PY=python3"
)

REM 2) 再试常见安装路径（应对忘记勾选 Add to PATH 的情况）
if not defined PY (
    for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
        if exist "%%D\python.exe" set "PY="%%D\python.exe""
    )
)
if not defined PY (
    for /d %%D in ("C:\Program Files\Python3*") do (
        if exist "%%D\python.exe" set "PY="%%D\python.exe""
    )
)
if not defined PY (
    for /d %%D in ("C:\Program Files (x86)\Python3*") do (
        if exist "%%D\python.exe" set "PY="%%D\python.exe""
    )
)
if not defined PY (
    for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python31*") do (
        if exist "%%D\python.exe" set "PY="%%D\python.exe""
    )
)
REM 3) 最后试 Microsoft Store 版
if not defined PY (
    if exist "%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe" (
        "%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe" --version >nul 2>&1 && set "PY="%LOCALAPPDATA%\Microsoft\WindowsApps\python.exe""
    )
)

if not defined PY (
    echo [没有找到 Python]
    echo.
    echo 请先安装 Python，三种办法任选其一：
    echo   1. 到 https://www.python.org/downloads/ 下载安装包，
    echo      安装第一屏务必勾选底部的 "Add python.exe to PATH"
    echo   2. 如果已经装过但忘了勾选：重新运行安装包，选 Modify，
    echo      勾选 "Add Python to environment variables"
    echo   3. 装完关掉本窗口，重新双击本文件即可
    echo.
    echo 如果实在不想装：把文件发给我，我直接帮你转。
    echo.
    pause
    exit /b 1
)

echo [使用 Python] %PY%
%PY% --version
echo.

REM ---------- 第 2 步：检查依赖 ----------
%PY% -c "import docx" >nul 2>&1
if errorlevel 1 (
    echo [首次运行] 正在安装依赖 python-docx，请稍候 ...
    %PY% -m pip install -r "%~dp0requirements.txt" -i https://pypi.tuna.tsinghua.edu.cn/simple
    if errorlevel 1 (
        echo.
        echo [依赖安装失败] 尝试改用默认源重试 ...
        %PY% -m pip install -r "%~dp0requirements.txt"
    )
    echo.
    %PY% -c "import docx" >nul 2>&1
    if errorlevel 1 (
        echo [错误] 依赖安装仍未成功，请把上面的报错截图发给我。
        pause
        exit /b 1
    )
    echo [依赖就绪]
    echo.
)

REM ---------- 第 3 步：转换 ----------
echo 请输入存放 docx 的文件夹路径（直接回车 = 本文件夹）：
set /p SRCDIR=

if "%SRCDIR%"=="" (
    %PY% "%~dp0docx2md.py" "%~dp0."
) else (
    %PY% "%~dp0docx2md.py" "%SRCDIR%"
)

echo.
echo ============================================
echo  处理完成
echo  请查看上方日志，以及输出目录中的「处理清单.csv」
echo ============================================
pause
