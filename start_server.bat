@echo off
chcp 65001
title 局域网打印服务系统

echo 正在启动打印服务...
cd /d %~dp0

REM 检查Python环境
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误: 未找到Python环境，请先安装Python 3.13.2
    pause
    exit /b 1
)

REM 安装依赖
echo 检查并安装依赖...
pip install -r requirements.txt

REM 启动服务
echo 启动打印服务...
python print_server.py

pause
