@echo off
chcp 65001 >nul
REM ============================================================
REM 超声波户表点检系统 - Windows 一键启动脚本
REM 用法：双击运行
REM ============================================================

cd /d %~dp0

echo ============================================
echo   超声波户表点检系统 - 启动中
echo ============================================
echo.

REM 1. 检查 Python
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo [错误] 未检测到 Python，请先安装 Python 3.8+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 2. 安装依赖（首次运行）
if not exist ".deps_ok" (
    echo [1/3] 正在安装依赖...
    pip install pandas openpyxl -q
    echo ok > .deps_ok
    echo   依赖安装完成
) else (
    echo [1/3] 依赖已就绪
)

REM 3. 启动协作服务器
echo.
echo [2/3] 正在启动协作服务器 (端口 8080)...
start "协作服务器 - 项目点检表" python 协作服务器_安全版.py 8080

REM 4. 设置定时任务（每天 8:20）
echo.
echo [3/3] 设置每日定时任务 (8:20)...
schtasks /create /tn "每日点检报表" /tr "python \"%cd%\更新点检表.py\"" /sc daily /st 08:20 /f >nul 2>&1
echo   定时任务已设置

REM 5. 显示访问地址
echo.
echo ============================================
echo   启动完成！
echo ============================================
echo.
echo   本机访问: http://localhost:8080
echo   用户管理: http://localhost:8080/admin/users
echo   默认账号: admin / admin123
echo.
echo   让朋友访问请运行 Cloudflare Tunnel
echo   或使用 ngrok: ngrok http 8080
echo.
echo   按任意键打开浏览器...
pause >nul
start http://localhost:8080
