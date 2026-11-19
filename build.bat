@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ============================================
echo   重新打包 docx2md 文档转换助手 exe
echo ============================================
echo.

REM 优先用带 tkinter + PyInstaller 的系统 Python（pythoncore-3.15）
set "BUILD_PY=%LOCALAPPDATA%\Python\pythoncore-3.15-64\python.exe"
if not exist "%BUILD_PY%" (
    set "BUILD_PY=py -3"
)

echo [使用 Python] %BUILD_PY%
"%BUILD_PY%" -m PyInstaller --version >nul 2>&1
if errorlevel 1 (
    echo [缺少 PyInstaller] 正在安装...
    "%BUILD_PY%" -m pip install pyinstaller
)

echo.
echo [开始打包] ...
"%BUILD_PY%" -m PyInstaller build.spec --noconfirm --clean
if errorlevel 1 (
    echo [打包失败] 请查看上方报错。
    pause
    exit /b 1
)

echo.
echo [复制发布文件] ...
if not exist "发布包" mkdir "发布包"
copy /Y "dist\docx2md-文档转换助手.exe" "发布包\" >nul
copy /Y "免安装版-使用说明.md" "发布包\使用说明.md" >nul
copy /Y "LICENSE" "发布包\" >nul

echo.
echo ============================================
echo   完成！exe 已生成在 dist\ 和 发布包\ 目录
echo ============================================
pause
