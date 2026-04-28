@echo off
chcp 65001 >nul
echo ====================================
echo   爱敬舆论分析 - API服务启动
echo ====================================
echo.
echo 正在启动API服务...
echo 端口: 8000
echo 访问地址: http://localhost:8000
echo.
echo 按 Ctrl+C 停止服务
echo ====================================
cd /d "%~dp0"
python api_server.py 8000
pause
