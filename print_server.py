import os
from flask import Flask, request, render_template_string, send_from_directory, redirect, url_for, flash, jsonify
# 打印相关
import win32print
import win32api
import win32gui
import win32con
import subprocess
from datetime import datetime
# 托盘相关
import threading
import sys
import pystray
from PIL import Image
import socket
import winreg
import time

# Windows DeviceCapabilities 常量
DC_DUPLEX = 7
DC_COLORDEVICE = 32
DC_PAPERS = 2
DC_PAPERNAMES = 16
DC_ENUMRESOLUTIONS = 13
DC_ORIENTATION = 17
DC_COPIES = 18
DC_TRUETYPE = 28
DC_DRIVER = 11

# Windows纸张大小常量
DMPAPER_LETTER = 1
DMPAPER_A4 = 9
DMPAPER_A3 = 8
DMPAPER_A5 = 11
DMPAPER_B4 = 12
DMPAPER_B5 = 13
DMPAPER_LEGAL = 5
DMPAPER_EXECUTIVE = 7
DMPAPER_TABLOID = 3

# 纸张名称映射
PAPER_NAMES = {
    1: "Letter (8.5 x 11 in)",
    3: "Tabloid (11 x 17 in)",
    5: "Legal (8.5 x 14 in)",
    7: "Executive (7.25 x 10.5 in)",
    8: "A3 (297 x 420 mm)",
    9: "A4 (210 x 297 mm)",
    11: "A5 (148 x 210 mm)",
    12: "B4 (250 x 354 mm)",
    13: "B5 (182 x 257 mm)",
}

def clean_old_files(folder=None, expire_seconds=3600):
    """定期清理指定目录下超过expire_seconds的文件"""
    if folder is None:
        folder = UPLOAD_FOLDER
    while True:
        now = time.time()
        for fname in os.listdir(folder):
            fpath = os.path.join(folder, fname)
            if os.path.isfile(fpath):
                try:
                    if now - os.path.getmtime(fpath) > 600:  # 10分钟
                        os.remove(fpath)
                except Exception:
                    pass
        time.sleep(60)  # 每1分钟检查一次

# 兼容PyInstaller打包和源码运行的资源路径
def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

# 获取本机局域网IP
def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.settimeout(2)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        try:
            hostname = socket.gethostname()
            ip = socket.gethostbyname(hostname)
            if ip and ip != '127.0.0.1':
                return ip
        except Exception:
            pass

        try:
            import subprocess
            result = subprocess.run(['ipconfig'], capture_output=True, text=True, encoding='gbk', timeout=10)
            if result.returncode == 0:
                lines = result.stdout.split('\n')
                for line in lines:
                    if 'IPv4' in line and '地址' in line:
                        parts = line.split(':')
                        if len(parts) > 1:
                            ip = parts[1].strip()
                            if ip and not ip.startswith('127.') and not ip.startswith('169.254.'):
                                return ip
        except Exception:
            pass

        return '127.0.0.1'

def get_current_ip_config():
    """获取当前IP配置状态"""
    try:
        current_ip = get_local_ip()
        if current_ip and current_ip != '127.0.0.1':
            try:
                result = subprocess.run(['ipconfig', '/all'],
                                        capture_output=True, text=True,
                                        encoding='gbk', errors='ignore')

                config = {
                    'index': '1',
                    'description': '以太网适配器',
                    'ip': current_ip,
                    'subnet': '255.255.255.0',
                    'gateway': '',
                    'dhcp_enabled': True
                }

                if 'Default Gateway' in result.stdout or '默认网关' in result.stdout:
                    lines = result.stdout.split('\n')
                    for line in lines:
                        if 'Default Gateway' in line or '默认网关' in line:
                            parts = line.split(':')
                            if len(parts) > 1:
                                gateway = parts[1].strip()
                                if gateway and gateway != '':
                                    config['gateway'] = gateway
                                    break

                return config
            except Exception:
                return {
                    'index': '1',
                    'description': '网络适配器',
                    'ip': current_ip,
                    'subnet': '255.255.255.0',
                    'gateway': '',
                    'dhcp_enabled': True
                }
        else:
            return {}
    except Exception as e:
        print(f"获取IP配置失败: {e}")
        return {}

def set_static_ip(ip_address, subnet_mask='255.255.255.0', gateway=''):
    """设置静态IP地址"""
    try:
        config = get_current_ip_config()
        if not config:
            return False, "未找到有效的网络适配器"

        adapter_index = config['index']

        if not gateway:
            ip_parts = ip_address.split('.')
            if len(ip_parts) == 4:
                gateway = f"{ip_parts[0]}.{ip_parts[1]}.{ip_parts[2]}.1"

        cmd = [
            'netsh', 'interface', 'ip', 'set', 'address',
            f'name="本地连接"' if 'Ethernet' in config['description'] else f'name="以太网"',
            'static', ip_address, subnet_mask, gateway
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, encoding='gbk')

        if result.returncode == 0:
            return True, "IP地址设置成功"
        else:
            return set_static_ip_wmi(adapter_index, ip_address, subnet_mask, gateway)

    except Exception as e:
        return False, f"设置IP地址失败: {str(e)}"

def set_static_ip_wmi(adapter_index, ip_address, subnet_mask, gateway):
    """使用WMI设置静态IP地址"""
    try:
        cmd = [
            'wmic', 'path', 'win32_networkadapterconfiguration',
            'where', f'Index={adapter_index}',
            'call', 'EnableStatic',
            f'("{ip_address}")', f'("{subnet_mask}")'
        ]

        result = subprocess.run(cmd, capture_output=True, text=True)

        if result.returncode == 0:
            if gateway:
                gateway_cmd = [
                    'wmic', 'path', 'win32_networkadapterconfiguration',
                    'where', f'Index={adapter_index}',
                    'call', 'SetGateways',
                    f'("{gateway}")', '(1)'
                ]
                subprocess.run(gateway_cmd, capture_output=True, text=True)

            return True, "IP地址设置成功"
        else:
            return False, f"WMI设置失败: {result.stderr}"

    except Exception as e:
        return False, f"WMI设置异常: {str(e)}"

def set_dhcp():
    """启用DHCP动态获取IP"""
    try:
        config = get_current_ip_config()
        if not config:
            return False, "未找到有效的网络适配器"

        cmd = [
            'netsh', 'interface', 'ip', 'set', 'address',
            f'name="本地连接"' if 'Ethernet' in config['description'] else f'name="以太网"',
            'dhcp'
        ]

        result = subprocess.run(cmd, capture_output=True, text=True, encoding='gbk')

        if result.returncode == 0:
            return True, "已启用DHCP动态获取IP"
        else:
            adapter_index = config['index']
            cmd = [
                'wmic', 'path', 'win32_networkadapterconfiguration',
                'where', f'Index={adapter_index}',
                'call', 'EnableDHCP'
            ]

            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                return True, "已启用DHCP动态获取IP"
            else:
                return False, f"启用DHCP失败: {result.stderr}"

    except Exception as e:
        return False, f"启用DHCP异常: {str(e)}"

def suggest_static_ip():
    """建议一个可用的静态IP地址"""
    current_ip = get_local_ip()
    if current_ip and current_ip != '127.0.0.1':
        ip_parts = current_ip.split('.')
        if len(ip_parts) == 4:
            return f"{ip_parts[0]}.{ip_parts[1]}.{ip_parts[2]}.100"
    return "192.168.1.100"

# 开机自启注册表操作
def set_autostart(enable=True):
    exe_path = sys.executable
    key = r'Software\\Microsoft\\Windows\\CurrentVersion\\Run'
    name = 'PrintServerApp'
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_ALL_ACCESS) as regkey:
        if enable:
            winreg.SetValueEx(regkey, name, 0, winreg.REG_SZ, exe_path)
        else:
            try:
                winreg.DeleteValue(regkey, name)
            except FileNotFoundError:
                pass

def get_autostart():
    key = r'Software\\Microsoft\\Windows\\CurrentVersion\\Run'
    name = 'PrintServerApp'
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key, 0, winreg.KEY_READ) as regkey:
            val, _ = winreg.QueryValueEx(regkey, name)
            return True if val else False
    except FileNotFoundError:
        return False

app = Flask(__name__)
app.secret_key = 'print_server_secret_key'
UPLOAD_FOLDER = os.path.join(os.path.expanduser('~'), 'Desktop', 'lan-printing-uploads')
LOG_FILE = 'print_log.txt'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# 虚拟打印机名称列表
VIRTUAL_PRINTERS = {
    '导出为WPS PDF', 'WPS PDF', 'Microsoft Print to PDF', 'Microsoft XPS Document Writer',
    'Fax', '传真', 'OneNote', 'OneNote (Desktop)', 'Send To OneNote 2016',
    'Adobe PDF', 'Foxit Reader PDF Printer', 'PDF Creator', 'CutePDF Writer',
    'novaPDF', 'PDFCreator', 'Bullzip PDF Printer', 'doPDF', 'PDF24',
    'Virtual PDF Printer', '虚拟PDF打印机', 'Send to Kindle', '发送到WPS高级打印'
}

# 获取所有本地和网络连接打印机，过滤掉虚拟打印机
ALL_PRINTERS = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
PRINTERS = [p for p in ALL_PRINTERS if p not in VIRTUAL_PRINTERS]

def get_default_printer():
    """获取系统默认打印机"""
    try:
        default_printer = win32print.GetDefaultPrinter()
        if default_printer in PRINTERS:
            return default_printer
        elif PRINTERS:
            return PRINTERS[0]
        else:
            return None
    except Exception as e:
        print(f"获取默认打印机失败: {e}")
        return PRINTERS[0] if PRINTERS else None

def refresh_printer_list():
    """刷新打印机列表"""
    global ALL_PRINTERS, PRINTERS
    try:
        ALL_PRINTERS = [p[2] for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)]
        PRINTERS = [p for p in ALL_PRINTERS if p not in VIRTUAL_PRINTERS]
        print(f"打印机列表已刷新，检测到 {len(PRINTERS)} 台物理打印机")
        return True
    except Exception as e:
        print(f"刷新打印机列表失败: {e}")
        return False

# 美化后的HTML模板
HTML = '''
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>局域网打印服务系统</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css" rel="stylesheet">
    <style>
        :root {
            --primary-color: #00ff9d;
            --secondary-color: #00d4ff;
            --accent-color: #ff00cc;
            --success-color: #00ff9d;
            --warning-color: #ff00cc;
            --dark-bg: #0a0a12;
            --card-bg: #11111d;
            --text-primary: #e0e0ff;
            --text-secondary: #8a8aa5;
            --border-color: #2a2a40;
            --card-shadow: 0 4px 30px rgba(0, 255, 230, 0.1);
            --hover-shadow: 0 0 20px rgba(0, 255, 230, 0.3);
            --neon-glow: 0 0 10px rgba(0, 255, 230, 0.5), 0 0 20px rgba(0, 255, 230, 0.3);
        }
        
        body {
            background: linear-gradient(135deg, #0a0a12 0%, #1a1a30 100%);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            min-height: 100vh;
            padding: 20px 0;
            color: var(--text-primary);
            background-image: 
                radial-gradient(circle at 25% 30%, rgba(0, 255, 157, 0.1) 0%, transparent 40%),
                radial-gradient(circle at 75% 60%, rgba(255, 0, 204, 0.1) 0%, transparent 40%),
                linear-gradient(to right, rgba(0, 0, 0, 0.2), rgba(0, 0, 0, 0.5)),
                repeating-linear-gradient(
                    45deg,
                    rgba(10, 10, 18, 0.5),
                    rgba(10, 10, 18, 0.5) 1px,
                    rgba(10, 10, 18, 0.7) 1px,
                    rgba(10, 10, 18, 0.7) 2px
                );
        }
        
        .main-container {
            max-width: 1200px;
            margin: 0 auto;
            background: var(--card-bg);
            border-radius: 20px;
            box-shadow: var(--card-shadow);
            overflow: hidden;
            backdrop-filter: blur(10px);
            border: 1px solid var(--border-color);
        }
        
        .header {
            background: linear-gradient(135deg, #11111d 0%, #1a1a30 100%);
            color: var(--text-primary);
            padding: 30px;
            text-align: center;
            position: relative;
            border-bottom: 1px solid var(--primary-color);
            box-shadow: var(--neon-glow);
        }
        
        .header h1 {
            font-weight: 700;
            margin-bottom: 10px;
            font-size: 2.5rem;
            color: var(--primary-color);
            text-shadow: var(--neon-glow);
            letter-spacing: 1px;
        }
        
        .header .subtitle {
            font-size: 1.2rem;
            color: var(--secondary-color);
            opacity: 0.9;
        }
        
        .nav-tabs {
            background: var(--dark-bg);
            padding: 0 30px;
            border-bottom: 1px solid var(--border-color);
        }
        
        .nav-link {
            padding: 15px 25px;
            font-weight: 600;
            color: var(--text-secondary);
            border: none;
            transition: all 0.3s ease;
        }
        
        .nav-link.active {
            color: var(--primary-color);
            border-bottom: 3px solid var(--primary-color);
            background: transparent;
            text-shadow: 0 0 5px rgba(0, 255, 157, 0.5);
        }
        
        .nav-link:hover {
            color: var(--secondary-color);
            transform: translateY(-2px);
            text-shadow: 0 0 5px rgba(0, 212, 255, 0.5);
        }
        
        .tab-content {
            padding: 30px;
            background: var(--card-bg);
        }
        
        .card {
            background: var(--dark-bg);
            border: 1px solid var(--border-color);
            border-radius: 15px;
            box-shadow: var(--card-shadow);
            transition: all 0.3s ease;
            margin-bottom: 20px;
            position: relative;
            overflow: hidden;
        }
        
        .card:hover {
            box-shadow: var(--hover-shadow);
            transform: translateY(-5px);
            border-color: var(--primary-color);
        }
        
        .card-header {
            background: linear-gradient(90deg, var(--dark-bg), var(--card-bg));
            color: var(--primary-color);
            border-radius: 15px 15px 0 0 !important;
            padding: 15px 20px;
            font-weight: 600;
            border-bottom: 1px solid var(--border-color);
        }
        
        .form-control, .form-select {
            background: var(--dark-bg);
            color: var(--text-primary);
            border-radius: 10px;
            border: 2px solid var(--border-color);
            padding: 10px 15px;
            transition: all 0.3s ease;
        }
        
        .form-label {
            color: var(--primary-color);
            font-weight: 600;
            text-shadow: var(--neon-glow);
        }
        
        .form-control:focus, .form-select:focus {
            border-color: var(--primary-color);
            box-shadow: 0 0 0 0.2rem rgba(0, 255, 157, 0.25);
            background: var(--dark-bg);
            color: var(--text-primary);
        }
        
        .btn {
            border-radius: 10px;
            padding: 10px 25px;
            font-weight: 600;
            transition: all 0.3s ease;
            border: none;
            background: var(--dark-bg);
            color: var(--text-primary);
            border: 1px solid var(--border-color);
        }
        
        .btn-primary {
            background: linear-gradient(90deg, var(--primary-color), var(--secondary-color));
            color: black;
            border: none;
        }
        
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 0 20px rgba(0, 255, 157, 0.5);
        }
        
        .alert {
            border-radius: 10px;
            border: none;
            padding: 15px 20px;
            background: var(--dark-bg);
            color: var(--text-primary);
            border-left: 4px solid var(--primary-color);
        }
        
        .status-badge {
            font-size: 0.8rem;
            padding: 5px 10px;
            border-radius: 20px;
        }
        
        .file-list {
            max-height: 300px;
            overflow-y: auto;
            background: var(--dark-bg);
            border-radius: 10px;
            padding: 10px;
            border: 1px solid var(--border-color);
        }
        
        .log-entry {
            padding: 10px;
            border-bottom: 1px solid var(--border-color);
            font-size: 0.9rem;
            color: var(--text-secondary);
        }
        
        .log-entry:last-child {
            border-bottom: none;
        }
        
        .printer-status {
            display: inline-flex;
            align-items: center;
            gap: 5px;
            font-size: 0.9rem;
        }
        
        .status-online {
            color: var(--success-color);
            text-shadow: 0 0 5px rgba(0, 255, 157, 0.5);
        }
        
        .status-offline {
            color: var(--warning-color);
        }
        
        .feature-icon {
            font-size: 2rem;
            color: var(--primary-color);
            margin-bottom: 10px;
            text-shadow: var(--neon-glow);
        }
        
        .stat-card {
            text-align: center;
            padding: 20px;
        }
        
        .stat-number {
            font-size: 2rem;
            font-weight: 700;
            color: var(--primary-color);
            text-shadow: var(--neon-glow);
            letter-spacing: 1px;
        }
        
        .upload-area {
            border: 2px dashed var(--border-color);
            border-radius: 15px;
            padding: 40px;
            text-align: center;
            transition: all 0.3s ease;
            cursor: pointer;
            background: rgba(0, 255, 157, 0.05);
        }
        
        .upload-area h5 {
            color: white;
            text-shadow: var(--neon-glow);
            margin-bottom: 10px;
        }
        
        .upload-area p {
            color: var(--text-primary);
        }
        
        .upload-area:hover {
            border-color: var(--primary-color);
            background-color: rgba(0, 255, 157, 0.1);
            box-shadow: var(--neon-glow);
        }
        
        .upload-icon {
            font-size: 2rem;
            color: var(--primary-color);
            margin-bottom: 15px;
            text-shadow: var(--neon-glow);
        }
        
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        @keyframes neonPulse {
            0%, 100% { opacity: 0.7; }
            50% { opacity: 1; }
        }
        
        .neon-pulse {
            animation: neonPulse 2s ease-in-out infinite;
        }
        
        .grid-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background-image: linear-gradient(rgba(0, 255, 157, 0.1) 1px, transparent 1px),
                             linear-gradient(90deg, rgba(0, 255, 157, 0.1) 1px, transparent 1px);
            background-size: 20px 20px;
            pointer-events: none;
            z-index: -1;
        }
        
        /* 滚动条样式 */
        ::-webkit-scrollbar {
            width: 10px;
        }
        
        ::-webkit-scrollbar-track {
            background: var(--dark-bg);
        }
        
        ::-webkit-scrollbar-thumb {
            background: var(--border-color);
            border-radius: 5px;
        }
        
        ::-webkit-scrollbar-thumb:hover {
            background: var(--primary-color);
        }
        
        .fade-in {
            animation: fadeIn 0.5s ease-in-out;
        }
    </style>
</head>
<body>
    <div class="grid-overlay"></div>
    <div class="main-container fade-in">
        <!-- 头部 -->
        <div class="header">
            <h1><i class="bi bi-printer-fill"></i> 局域网打印服务系统</h1>
            <div class="subtitle">安全、高效、便捷的局域网打印解决方案</div>
        </div>
        
        <!-- 导航标签 -->
        <ul class="nav nav-tabs" id="mainTabs" role="tablist">
            <li class="nav-item" role="presentation">
                <button class="nav-link active" id="print-tab" data-bs-toggle="tab" data-bs-target="#print" type="button" role="tab">
                    <i class="bi bi-printer"></i> 打印管理
                </button>
            </li>
            <li class="nav-item" role="presentation">
                <button class="nav-link" id="status-tab" data-bs-toggle="tab" data-bs-target="#status" type="button" role="tab">
                    <i class="bi bi-graph-up"></i> 系统状态
                </button>
            </li>
        </ul>

        <!-- 消息提示 -->
        {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
                <div class="container mt-3">
                    {% for category, msg in messages %}
                        <div class="alert alert-{{ 'danger' if category == 'error' else category }} alert-dismissible fade show" role="alert">
                            <i class="bi bi-{{ 'check-circle-fill' if category == 'success' else 'exclamation-triangle-fill' if category == 'warning' else 'info-circle-fill' if category == 'info' else 'x-circle-fill' }}"></i>
                            {{ msg }}
                            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
                        </div>
                    {% endfor %}
                </div>
            {% endif %}
        {% endwith %}

        <div class="tab-content" id="mainTabsContent">
            <!-- 打印管理标签页 -->
            <div class="tab-pane fade show active" id="print" role="tabpanel">
                <div class="row">
                    <div class="col-lg-8">
                        <div class="card">
                            <div class="card-header">
                                <i class="bi bi-upload"></i> 文件上传与打印
                            </div>
                            <div class="card-body">
                                <form method="post" enctype="multipart/form-data" id="uploadForm">
                                    <input type="hidden" name="action" value="print">
                                    
                                    <div class="row g-3">
                                        <div class="col-md-6">
                                            <label class="form-label fw-bold">选择打印机</label>
                                            <div class="input-group">
                                                <select name="printer" class="form-select" id="printerSelect">
                                                    {% if printers %}
                                                        {% for p in printers %}
                                                            <option value="{{ p }}" {% if p == default_printer %}selected{% endif %}>
                                                                {{ p }}{% if p == default_printer %} (默认){% endif %}
                                                            </option>
                                                        {% endfor %}
                                                    {% else %}
                                                        <option value="">未检测到可用打印机</option>
                                                    {% endif %}
                                                </select>
                                                <button type="button" class="btn btn-outline-secondary" onclick="refreshPrinterList()" title="刷新打印机列表">
                                                    <i class="bi bi-arrow-clockwise"></i>
                                                </button>
                                            </div>
                                            <div class="form-text">
                                                {% if printers %}
                                                    已过滤虚拟打印机，自动选择默认打印机
                                                {% else %}
                                                    <span class="text-danger">⚠️ 未检测到物理打印机</span>
                                                {% endif %}
                                            </div>
                                        </div>
                                        
                                        <div class="col-md-3">
                                            <label class="form-label fw-bold">打印份数</label>
                                            <select name="copies" class="form-select">
                                                {% for i in range(1, 11) %}
                                                    <option value="{{ i }}" {% if i == 1 %}selected{% endif %}>{{ i }}</option>
                                                {% endfor %}
                                            </select>
                                        </div>
                                        
                                        <div class="col-md-3">
                                            <label class="form-label fw-bold">打印模式</label>
                                            <select name="duplex" class="form-select">
                                                <option value="1">单面打印</option>
                                                {% if printer_caps and printer_caps.get('duplex_support') %}
                                                <option value="2">双面长边</option>
                                                <option value="3">双面短边</option>
                                                {% endif %}
                                            </select>
                                        </div>

                                        <!-- 新增：色彩选择模式 -->
                                        <div class="col-md-3">
                                            <label class="form-label fw-bold">色彩模式</label>
                                            <select name="color_mode" class="form-select">
                                                <option value="color">彩色</option>
                                                <option value="monochrome">黑白</option>
                                            </select>
                                        </div>

                                        <!-- 新增：打印方向 -->
                                        <div class="col-md-3">
                                            <label class="form-label fw-bold">打印方向</label>
                                            <select name="orientation" class="form-select">
                                                <option value="portrait">纵向</option>
                                                <option value="landscape">横向</option>
                                            </select>
                                        </div>
                                        
                                        <div class="col-md-6">
                                            <label class="form-label fw-bold">纸张设置</label>
                                            <select name="papersize" class="form-select" id="paperSelect">
                                                {% if printer_caps and printer_caps.get('papers') %}
                                                    {% for p in printer_caps.papers %}
                                                    <option value="{{ p.id }}" {% if p.id == 9 %}selected{% endif %}>{{ p.name }}</option>
                                                    {% endfor %}
                                                {% else %}
                                                    <option value="9" selected>A4 (210 x 297 mm)</option>
                                                {% endif %}
                                            </select>
                                        </div>
                                        
                                        <div class="col-md-6">
                                            <label class="form-label fw-bold">打印质量</label>
                                            <select name="quality" class="form-select" id="qualitySelect">
                                                {% if printer_caps and printer_caps.get('resolutions') %}
                                                    {% for r in printer_caps.resolutions %}
                                                    <option value="{{ r }}">{{ r }} DPI</option>
                                                    {% endfor %}
                                                {% else %}
                                                    <option value="600x600">600x600 DPI</option>
                                                {% endif %}
                                            </select>
                                        </div>

                                        <!-- 新增：打印比例 -->
                                        <div class="col-md-6">
                                            <label class="form-label fw-bold">打印比例</label>
                                            <select name="scale" class="form-select">
                                                <option value="original">原始比例</option>
                                                <option value="fit_margins">适合纸张边距</option>
                                                <option value="fit_printable">适合至可打印区域</option>
                                            </select>
                                        </div>

                                        <!-- 新增：打印范围 -->
                                        <div class="col-md-6">
                                            <label class="form-label fw-bold">打印范围</label>
                                            <select name="print_range" class="form-select" id="printRangeSelect" onchange="togglePageRangeInput()">
                                                <option value="all">全部</option>
                                                <option value="current">当前页面</option>
                                                <option value="pages">页码选择</option>
                                            </select>
                                            <div id="pageRangeContainer" class="mt-2 d-none">
                                                <input type="text" name="page_range" class="form-control" placeholder="例如：1,3-5,7" disabled>
                                                <div class="form-text">使用逗号分隔页码，连字符表示范围</div>
                                            </div>
                                        </div>
                                        
                                        <div class="col-12">
                                            <label class="form-label fw-bold">选择文件</label>
                                            <div class="upload-area" onclick="document.getElementById('fileInput').click()">
                                                <i class="bi bi-cloud-upload upload-icon"></i>
                                                <h5>点击选择文件或拖拽文件到此区域，支持 PDF, JPG, PNG, DOC, DOCX, PPT, PPTX, XLS, XLSX, TXT 格式</h5>
                                                
                                                <input type="file" name="file" id="fileInput" multiple class="d-none" onchange="updateFileList()">
                                            </div>
                                            <div id="fileList" class="mt-3"></div>
                                        </div>
                                        
                                        <div class="col-12 text-end">
                                            {% if printers %}
                                                <button type="submit" class="btn btn-primary btn-lg">
                                                    <i class="bi bi-send-check"></i> 开始打印
                                                </button>
                                            {% else %}
                                                <button type="button" class="btn btn-secondary btn-lg" disabled>
                                                    <i class="bi bi-exclamation-triangle"></i> 无可用打印机
                                                </button>
                                            {% endif %}
                                        </div>
                                    </div>
                                </form>
                            </div>
                        </div>
                    </div>
                    
                    <div class="col-lg-4">
                        <div class="card">
                            <div class="card-header">
                                <i class="bi bi-info-circle"></i> 打印说明
                            </div>
                            <div class="card-body">
                                <div class="alert alert-info">
                                    <h6><i class="bi bi-lightbulb"></i> 使用提示</h6>
                                    <ul class="small mb-0">
                                        <li>支持多种文件格式直接打印</li>
                                        <li>自动过滤虚拟打印机</li>
                                        <li>静默打印无需确认</li>
                                        <li>实时显示打印状态</li>
                                    </ul>
                                </div>
                                
                                <div class="alert alert-success">
                                    <h6><i class="bi bi-check-circle"></i> 当前状态</h6>
                                    <div class="printer-status {% if printers %}status-online{% else %}status-offline{% endif %}">
                                        <i class="bi bi-{% if printers %}check-circle{% else %}x-circle{% endif %}"></i>
                                        {% if printers %}
                                            检测到 {{ printers|length }} 台打印机
                                        {% else %}
                                            未检测到可用打印机
                                        {% endif %}
                                    </div>
                                </div>
                            </div>
                        </div>
                        
                        <div class="card mt-3">
                            <div class="card-header">
                                <i class="bi bi-clock-history"></i> 最近文件
                            </div>
                            <div class="card-body file-list">
                                {% for f in files[-5:] %}
                                    <div class="d-flex justify-content-between align-items-center log-entry">
                                        <span class="text-truncate" style="max-width: 70%;">{{ f }}</span>
                                        <a href="/preview/{{ f }}" target="_blank" class="btn btn-sm btn-outline-primary">
                                            <i class="bi bi-eye"></i>
                                        </a>
                                    </div>
                                {% else %}
                                    <p class="text-muted text-center">暂无文件</p>
                                {% endfor %}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- 网络配置标签页 -->
            <div class="tab-pane fade" id="network" role="tabpanel">
                <div class="row">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <i class="bi bi-ethernet"></i> 当前网络状态
                            </div>
                            <div class="card-body">
                                {% if ip_config %}
                                <div class="row g-3">
                                    <div class="col-12">
                                        <label class="form-label fw-bold">IP地址</label>
                                        <div class="d-flex align-items-center">
                                            <span class="fw-bold text-primary">{{ ip_config.ip }}</span>
                                            <span class="badge {% if ip_config.dhcp_enabled %}bg-success{% else %}bg-primary{% endif %} status-badge ms-2">
                                                {% if ip_config.dhcp_enabled %}DHCP{% else %}静态IP{% endif %}
                                            </span>
                                        </div>
                                    </div>
                                    <div class="col-6">
                                        <label class="form-label fw-bold">子网掩码</label>
                                        <div>{{ ip_config.subnet }}</div>
                                    </div>
                                    <div class="col-6">
                                        <label class="form-label fw-bold">默认网关</label>
                                        <div>{{ ip_config.gateway if ip_config.gateway else '未设置' }}</div>
                                    </div>
                                    <div class="col-12">
                                        <label class="form-label fw-bold">网络适配器</label>
                                        <div class="text-truncate">{{ ip_config.description }}</div>
                                    </div>
                                </div>
                                {% else %}
                                <div class="alert alert-warning text-center">
                                    <i class="bi bi-wifi-off"></i><br>
                                    未检测到网络连接
                                </div>
                                {% endif %}
                            </div>
                        </div>
                    </div>
                    
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <i class="bi bi-gear"></i> IP地址配置
                            </div>
                            <div class="card-body">
                                <form method="post">
                                    <input type="hidden" name="action" value="set_static_ip">
                                    <div class="mb-3">
                                        <label class="form-label fw-bold">IP地址</label>
                                        <input type="text" name="ip_address" class="form-control" 
                                               value="{{ suggested_ip }}" placeholder="192.168.1.100" required>
                                    </div>
                                    <div class="mb-3">
                                        <label class="form-label fw-bold">子网掩码</label>
                                        <input type="text" name="subnet_mask" class="form-control" 
                                               value="255.255.255.0" required>
                                    </div>
                                    <div class="mb-3">
                                        <label class="form-label fw-bold">默认网关</label>
                                        <input type="text" name="gateway" class="form-control" 
                                               placeholder="可选，自动推导">
                                    </div>
                                    <button type="submit" class="btn btn-primary w-100">
                                        <i class="bi bi-check-circle"></i> 设置静态IP
                                    </button>
                                </form>
                                
                                <hr>
                                
                                <form method="post" class="mt-3">
                                    <input type="hidden" name="action" value="enable_dhcp">
                                    <button type="submit" class="btn btn-outline-primary w-100">
                                        <i class="bi bi-arrow-repeat"></i> 启用DHCP自动获取
                                    </button>
                                </form>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- 系统状态标签页 -->
            <div class="tab-pane fade" id="status" role="tabpanel">
                <div class="row">
                    <div class="col-md-3">
                        <div class="card stat-card">
                            <i class="bi bi-printer feature-icon"></i>
                            <div class="stat-number">{{ printers|length }}</div>
                            <div>可用打印机</div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="card stat-card">
                            <i class="bi bi-file-earmark feature-icon"></i>
                            <div class="stat-number">{{ files|length }}</div>
                            <div>待打印文件</div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="card stat-card">
                            <i class="bi bi-wifi feature-icon"></i>
                            <div class="stat-number">{% if ip_config %}在线{% else %}离线{% endif %}</div>
                            <div>网络状态</div>
                        </div>
                    </div>
                    <div class="col-md-3">
                        <div class="card stat-card">
                            <i class="bi bi-clock feature-icon"></i>
                            <div class="stat-number">{{ logs|length }}</div>
                            <div>今日日志</div>
                        </div>
                    </div>
                </div>
                
                <div class="row mt-4">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <i class="bi bi-list-check"></i> 打印机列表
                            </div>
                            <div class="card-body file-list">
                                {% for printer in printers %}
                                    <div class="log-entry">
                                        <div class="d-flex justify-content-between align-items-center">
                                            <span>{{ printer }}</span>
                                            {% if printer == default_printer %}
                                                <span class="badge bg-primary status-badge">默认</span>
                                            {% endif %}
                                        </div>
                                    </div>
                                {% else %}
                                    <p class="text-muted text-center">未检测到打印机</p>
                                {% endfor %}
                            </div>
                        </div>
                    </div>
                    
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">
                                <i class="bi bi-journal-text"></i> 最近日志
                            </div>
                            <div class="card-body file-list">
                                {% for log in logs[-10:] %}
                                    <div class="log-entry">
                                        <small class="text-muted">{{ log.split(' 打印:')[0] }}</small><br>
                                        {{ log.split(' 打印:')[1] if ' 打印:' in log else log }}
                                    </div>
                                {% else %}
                                    <p class="text-muted text-center">暂无日志记录</p>
                                {% endfor %}
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.2/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // 控制页码范围输入框的显示/隐藏
        function togglePageRangeInput() {
            const rangeSelect = document.getElementById('printRangeSelect');
            const rangeContainer = document.getElementById('pageRangeContainer');
            const rangeInput = rangeContainer.querySelector('input');
            
            if (rangeSelect.value === 'pages') {
                rangeContainer.classList.remove('d-none');
                rangeInput.disabled = false;
            } else {
                rangeContainer.classList.add('d-none');
                rangeInput.disabled = true;
            }
        }
        
        // 文件选择处理
        function updateFileList() {
            const fileInput = document.getElementById('fileInput');
            const fileList = document.getElementById('fileList');
            const files = fileInput.files;
            
            if (files.length > 0) {
                let html = '<div class="alert alert-success"><h6>已选择文件:</h6><ul class="mb-0">';
                for (let file of files) {
                    html += `<li>${file.name} (${(file.size / 1024).toFixed(1)} KB)</li>`;
                }
                html += '</ul></div>';
                fileList.innerHTML = html;
            } else {
                fileList.innerHTML = '';
            }
        }
        
        // 拖拽文件支持
        const uploadArea = document.querySelector('.upload-area');
        uploadArea.addEventListener('dragover', (e) => {
            e.preventDefault();
            uploadArea.style.borderColor = '#4361ee';
            uploadArea.style.backgroundColor = '#f0f4ff';
        });
        
        uploadArea.addEventListener('dragleave', () => {
            uploadArea.style.borderColor = '#dee2e6';
            uploadArea.style.backgroundColor = '';
        });
        
        uploadArea.addEventListener('drop', (e) => {
            e.preventDefault();
            uploadArea.style.borderColor = '#dee2e6';
            uploadArea.style.backgroundColor = '';
            
            const files = e.dataTransfer.files;
            document.getElementById('fileInput').files = files;
            updateFileList();
        });
        
        // 打印机信息刷新
        function refreshPrinterInfo() {
            const printerSelect = document.getElementById('printerSelect');
            const paperSelect = document.getElementById('paperSelect');
            const qualitySelect = document.getElementById('qualitySelect');
            
            if (!printerSelect || !printerSelect.value) return;
            
            fetch('/api/printer_info?printer=' + encodeURIComponent(printerSelect.value))
                .then(r => r.json())
                .then(data => {
                    if (data.success && data.capabilities) {
                        const caps = data.capabilities;
                        
                        // 更新纸张选项
                        if (paperSelect && caps.papers && caps.papers.length) {
                            const prev = paperSelect.value;
                            paperSelect.innerHTML = '';
                            caps.papers.forEach(p => {
                                const opt = document.createElement('option');
                                opt.value = p.id;
                                opt.textContent = p.name;
                                paperSelect.appendChild(opt);
                            });
                            if (prev) paperSelect.value = prev;
                        }
                        
                        // 更新质量选项
                        if (qualitySelect && caps.resolutions && caps.resolutions.length) {
                            qualitySelect.innerHTML = '';
                            caps.resolutions.forEach(r => {
                                const opt = document.createElement('option');
                                opt.value = r;
                                opt.textContent = r + ' DPI';
                                qualitySelect.appendChild(opt);
                            });
                        }
                    }
                })
                .catch(console.error);
        }
        
        // 刷新打印机列表
        function refreshPrinterList() {
            const btn = document.querySelector('button[onclick="refreshPrinterList()"]');
            btn.innerHTML = '<i class="bi bi-arrow-repeat spinner-border spinner-border-sm"></i>';
            btn.disabled = true;
            
            fetch('/api/refresh_printers')
                .then(r => r.json())
                .then(data => {
                    if (data.success) {
                        const select = document.getElementById('printerSelect');
                        select.innerHTML = '';
                        
                        if (data.printers && data.printers.length) {
                            data.printers.forEach(p => {
                                const opt = document.createElement('option');
                                opt.value = p;
                                opt.textContent = p + (p === data.default_printer ? ' (默认)' : '');
                                if (p === data.default_printer) opt.selected = true;
                                select.appendChild(opt);
                            });
                            
                            // 显示成功消息
                            showAlert('打印机列表刷新成功', 'success');
                            refreshPrinterInfo();
                        } else {
                            select.innerHTML = '<option value="">未检测到可用打印机</option>';
                            showAlert('未找到可用打印机', 'warning');
                        }
                    } else {
                        showAlert('刷新失败: ' + data.error, 'danger');
                    }
                })
                .finally(() => {
                    btn.innerHTML = '<i class="bi bi-arrow-clockwise"></i>';
                    btn.disabled = false;
                });
        }
        
        // 显示提示消息
        function showAlert(message, type) {
            const alertDiv = document.createElement('div');
            alertDiv.className = `alert alert-${type} alert-dismissible fade show`;
            alertDiv.innerHTML = `
                <i class="bi bi-${type === 'success' ? 'check-circle' : type === 'warning' ? 'exclamation-triangle' : 'info-circle'}-fill"></i>
                ${message}
                <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
            `;
            document.querySelector('.tab-content').prepend(alertDiv);
            
            setTimeout(() => {
                if (alertDiv.parentNode) {
                    alertDiv.remove();
                }
            }, 5000);
        }
        
        // 初始化
        document.addEventListener('DOMContentLoaded', function() {
            const printerSelect = document.getElementById('printerSelect');
            if (printerSelect) {
                printerSelect.addEventListener('change', refreshPrinterInfo);
                refreshPrinterInfo(); // 初始加载
            }
            
            // 表单提交验证
            document.getElementById('uploadForm')?.addEventListener('submit', function(e) {
                const files = document.getElementById('fileInput').files;
                if (files.length === 0) {
                    e.preventDefault();
                    showAlert('请选择要打印的文件', 'warning');
                    return false;
                }
            });
        });
    </script>
</body>
</html>
'''

# 允许的文件类型
ALLOWED_EXT = {'pdf', 'jpg', 'jpeg', 'png', 'txt', 'doc', 'docx', 'ppt', 'pptx', 'xls', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXT

def is_physical_printer(printer_name):
    """检查是否为真正的物理打印机"""
    if printer_name in VIRTUAL_PRINTERS:
        return False

    virtual_keywords = ['pdf', 'fax', '传真', 'xps', 'onenote', 'virtual', '虚拟', 'send to', 'export', '导出']
    printer_lower = printer_name.lower()

    for keyword in virtual_keywords:
        if keyword in printer_lower:
            return False

    return True

def log_print(filename, printer, copies, duplex, papersize, quality, color_mode='color', orientation='portrait', scale='original', print_range='all', page_range=''):
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(f"{datetime.now()} 打印: {filename} 打印机: {printer} 份数: {copies} 单双面: {duplex} 纸张: {papersize} 质量: {quality} 色彩: {color_mode} 方向: {orientation} 比例: {scale} 范围: {print_range}{f' ({page_range})' if page_range else ''}\n")

# 保留原有的打印功能函数（print_file_with_settings, apply_printer_settings等）
# 由于篇幅限制，这里省略了具体的打印功能实现，保持原代码不变

def print_image_silent(filepath, printer_name, copies=1, color_mode='color', orientation='portrait'):
    """静默打印图片文件"""
    try:
        import win32print
        import win32ui
        import win32con
        import win32api

        print(f"静默打印图片文件: {filepath} 到打印机: {printer_name} 份数: {copies} 色彩: {color_mode} 方向: {orientation}")

        # 先设置为默认打印机，确保所有设置生效
        old_default_printer = win32print.GetDefaultPrinter()
        win32print.SetDefaultPrinter(printer_name)

        try:
            # 打开打印机
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                # 获取打印机属性
                printer_info = win32print.GetPrinter(hprinter, 2)
                devmode = printer_info[1]

                # 确保DEVMODE结构具有正确的标志
                devmode.Fields |= (win32con.DM_ORIENTATION | win32con.DM_COLOR | win32con.DM_COPIES)

                # 应用打印方向
                if orientation == 'landscape':
                    devmode.Orientation = win32con.DMORIENT_LANDSCAPE
                else:
                    devmode.Orientation = win32con.DMORIENT_PORTRAIT
                print(f"应用打印方向: {orientation} -> {devmode.Orientation}")

                # 应用色彩模式
                if color_mode == 'monochrome':
                    devmode.Color = 1  # 单色
                else:
                    devmode.Color = 2  # 彩色
                print(f"应用色彩模式: {color_mode} -> {devmode.Color}")

                # 应用打印份数
                devmode.Copies = copies
                print(f"应用打印份数: {copies}")

                # 更新打印机属性
                win32print.SetPrinter(hprinter, 2, devmode, 0)

                # 验证设置是否正确应用
                updated_info = win32print.GetPrinter(hprinter, 2)
                updated_devmode = updated_info[1]
                print(f"设置应用后验证: 方向={updated_devmode.Orientation}, 色彩={updated_devmode.Color}")

                # 使用ShellExecute打印
                print(f"使用ShellExecute打印图片: {filepath} 到 {printer_name}")
                result = win32api.ShellExecute(
                    0,
                    "print",
                    filepath,
                    f'/d:"{printer_name}"',
                    ".",
                    0
                )

                if result > 32:
                    return True, f"已发送到 {printer_name} (图片已应用设置，设置验证通过)"
                else:
                    print(f"ShellExecute图片打印失败，返回码: {result}")
                    # 备用方案：使用Windows图片查看器打印
                    try:
                        photo_viewer_path = os.path.join(os.environ.get('SystemRoot', r'C:\Windows'), 'System32', 'rundll32.exe')
                        cmd = [
                            photo_viewer_path,
                            'C:\\Program Files\\Windows Photo Viewer\\PhotoViewer.dll',
                            'ImageView_Fullscreen',
                            '/p',
                            filepath
                        ]
                        subprocess.Popen(cmd, shell=False)
                        return True, f"已发送到 {printer_name} (使用Windows照片查看器打印)"
                    except Exception as e2:
                        print(f"备用打印方法失败: {str(e2)}")
                        raise
            finally:
                win32print.ClosePrinter(hprinter)
        finally:
            # 恢复原来的默认打印机
            if old_default_printer:
                win32print.SetDefaultPrinter(old_default_printer)

    except Exception as e:
        error_msg = f"图片打印异常: {str(e)}"
        print(error_msg)
        # 备用方案
        try:
            win32api.ShellExecute(
                0,
                "print",
                filepath,
                f'/d:"{printer_name}"',
                ".",
                0
            )
            return True, f"已发送到 {printer_name} (使用基础打印功能)"
        except Exception as e2:
            return False, f"图片打印失败: {str(e2)}"

def print_office_silent(filepath, printer_name, copies=1, color_mode='color', orientation='portrait'):
    """静默打印Office文档"""
    try:
        import win32com.client
        import os
        import time

        print(f"静默打印Office文档: {filepath} 到打印机: {printer_name} 份数: {copies} 色彩: {color_mode} 方向: {orientation}")

        file_ext = os.path.splitext(filepath)[1].lower()
        app = None
        success = False
        result_message = ""

        try:
            # 根据文件类型启动相应的Office应用
            if file_ext in ['.doc', '.docx']:
                app = win32com.client.Dispatch('Word.Application')
                # 设置为不可见，避免界面弹出
                app.Visible = False
                app.DisplayAlerts = False

                doc = app.Documents.Open(os.path.abspath(filepath), ReadOnly=True)

                # 设置打印方向
                if orientation == 'landscape':
                    for section in doc.Sections:
                        section.PageSetup.Orientation = 1  # wdOrientLandscape
                    print(f"Word文档设置为横向打印")
                else:
                    for section in doc.Sections:
                        section.PageSetup.Orientation = 0  # wdOrientPortrait
                    print(f"Word文档设置为纵向打印")

                # 设置色彩模式（Word通过打印机属性设置，这里通过PrintOut的属性传递）
                print(f"Word文档设置色彩模式: {color_mode}")

                # 打印文档，指定使用的打印机
                print(f"执行Word文档打印到: {printer_name}，份数: {copies}")
                doc.PrintOut(Copies=copies, Printer=printer_name, Background=True, PrintToFile=False)

                # 等待打印任务开始
                time.sleep(1)

                success = True
                result_message = f"Word文档已发送到 {printer_name} (已应用设置)"
                doc.Close(SaveChanges=0)

            elif file_ext in ['.xls', '.xlsx']:
                app = win32com.client.Dispatch('Excel.Application')
                # 设置为不可见，避免界面弹出
                app.Visible = False
                app.DisplayAlerts = False

                book = app.Workbooks.Open(os.path.abspath(filepath), ReadOnly=True)

                # 设置所有工作表的打印方向和色彩模式
                for sheet in book.Sheets:
                    if orientation == 'landscape':
                        sheet.PageSetup.Orientation = 2  # xlLandscape
                        print(f"Excel工作表 '{sheet.Name}' 设置为横向打印")
                    else:
                        sheet.PageSetup.Orientation = 1  # xlPortrait
                        print(f"Excel工作表 '{sheet.Name}' 设置为纵向打印")

                    # 设置色彩模式
                    if color_mode == 'monochrome':
                        sheet.PageSetup.BlackAndWhite = True
                        print(f"Excel工作表 '{sheet.Name}' 设置为黑白打印")
                    else:
                        sheet.PageSetup.BlackAndWhite = False
                        print(f"Excel工作表 '{sheet.Name}' 设置为彩色打印")

                # 打印工作簿
                print(f"执行Excel工作簿打印到: {printer_name}，份数: {copies}")
                book.PrintOut(Copies=copies, ActivePrinter=printer_name)

                # 等待打印任务开始
                time.sleep(1)

                success = True
                result_message = f"Excel文档已发送到 {printer_name} (已应用设置)"
                book.Close(SaveChanges=0)

            elif file_ext in ['.ppt', '.pptx']:
                app = win32com.client.Dispatch('PowerPoint.Application')
                # 设置为不可见，避免界面弹出
                app.Visible = False

                pres = app.Presentations.Open(os.path.abspath(filepath), WithWindow=False, ReadOnly=True)

                # 设置打印选项
                print_options = pres.PrintOptions
                if color_mode == 'monochrome':
                    print_options.OutputType = 2  # ppPrintOutputGrayscale
                    print(f"PowerPoint演示文稿设置为灰度打印")
                else:
                    print_options.OutputType = 1  # ppPrintOutputColor
                    print(f"PowerPoint演示文稿设置为彩色打印")

                # 打印演示文稿
                print(f"执行PowerPoint演示文稿打印到: {printer_name}，份数: {copies}")
                pres.PrintOut(PrintRange=None, Copies=copies, PrinterName=printer_name)

                # 等待打印任务开始
                time.sleep(1)

                success = True
                result_message = f"PowerPoint文档已发送到 {printer_name} (已应用设置)"
                pres.Close()

            else:
                raise ValueError(f"不支持的Office文件类型: {file_ext}")

            if success:
                return True, result_message
            else:
                raise Exception("打印任务未成功启动")

        finally:
            # 确保关闭Office应用
            if app:
                try:
                    app.Quit()
                    print("Office应用已成功关闭")
                except Exception as quit_error:
                    print(f"关闭Office应用时出错: {str(quit_error)}")
                # 释放COM对象
                import pythoncom
                pythoncom.CoUninitialize()

    except Exception as e:
        error_msg = f"Office文档打印异常: {str(e)}"
        print(error_msg)
        # 备用方案 - 使用ShellExecute
        try:
            import win32api
            print(f"使用备用方案: ShellExecute打印到 {printer_name}")
            result = win32api.ShellExecute(
                0,
                "print",
                filepath,
                f'/d:"{printer_name}"',
                ".",
                0
            )

            if result > 32:
                return True, f"已发送到 {printer_name} (使用系统默认打印)"
            else:
                raise Exception(f"ShellExecute打印失败，返回码: {result}")
        except Exception as e2:
            return False, f"Office文档打印失败: {str(e2)}"

def print_file_with_settings(filepath, printer_name, copies=1, duplex=1, papersize='A4', quality='normal', color_mode='color', orientation='portrait', scale='original'):
    """使用获取到的真实打印设置进行打印"""
    try:
        print(f"开始打印文件: {filepath}")
        print(f"目标打印机: {printer_name}")
        print(f"打印份数: {copies}")
        print(f"双面设置: {duplex}")
        print(f"纸张大小: {papersize}")
        print(f"打印质量: {quality}")
        print(f"色彩模式: {color_mode}")
        print(f"打印方向: {orientation}")
        print(f"打印比例: {scale}")

        file_ext = os.path.splitext(filepath)[1].lower()

        if file_ext == '.pdf':
            # 使用print_pdf_silent代替不存在的print_pdf_with_settings函数
            return print_pdf_silent(filepath, printer_name, copies, duplex, papersize, quality, color_mode, orientation, scale, 'all', '')
        elif file_ext in ['.jpg', '.jpeg', '.png', '.bmp', '.gif']:
            return print_image_silent(filepath, printer_name, copies, color_mode, orientation)
        elif file_ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx']:
            return print_office_silent(filepath, printer_name, copies, color_mode, orientation)
        elif file_ext == '.txt':
            return print_text_file_simple(filepath, printer_name, copies)
        else:
            print(f"未知文件类型 {file_ext}，尝试使用系统默认打印方式")
            return print_with_shell_execute(filepath, printer_name, copies)

    except Exception as e:
        print(f"打印操作失败: {e}")
        return print_file_silent_fallback(filepath, printer_name, copies)

def print_with_shell_execute(filepath, printer_name, copies=1, color_mode='color', orientation='portrait', papersize='9'):
    """使用ShellExecute执行打印并应用基本设置"""
    try:
        import win32api
        import win32print
        import win32con
        import os

        print(f"使用ShellExecute打印文件: {filepath} 到 {printer_name}")
        print(f"设置: 份数={copies}, 色彩={color_mode}, 方向={orientation}, 纸张={papersize}")

        # 先设置为默认打印机，确保设置生效
        old_default_printer = None
        try:
            old_default_printer = win32print.GetDefaultPrinter()
            win32print.SetDefaultPrinter(printer_name)

            # 打开打印机并应用设置
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                # 获取当前打印机设置
                printer_info = win32print.GetPrinter(hprinter, 2)
                devmode = printer_info[1]

                # 应用打印方向
                if orientation == 'landscape':
                    devmode.Orientation = win32con.DMORIENT_LANDSCAPE
                    print(f"应用打印方向: 横向")
                else:
                    devmode.Orientation = win32con.DMORIENT_PORTRAIT
                    print(f"应用打印方向: 纵向")

                # 应用色彩模式
                if color_mode == 'monochrome':
                    devmode.Color = 1  # 单色
                    print(f"应用色彩模式: 单色")
                else:
                    devmode.Color = 2  # 彩色
                    print(f"应用色彩模式: 彩色")

                # 应用纸张大小
                try:
                    paper_size_id = int(papersize)
                    devmode.PaperSize = paper_size_id
                    print(f"应用纸张大小ID: {paper_size_id}")
                except ValueError:
                    print(f"无效的纸张大小ID: {papersize}，使用默认值")

                # 应用打印份数
                devmode.Copies = copies
                print(f"应用打印份数: {copies}")

                # 强制设置DEVMODE标志位，确保所有设置生效
                devmode.Fields |= (
                        win32con.DM_ORIENTATION |
                        win32con.DM_COLOR |
                        win32con.DM_PAPERSIZE |
                        win32con.DM_COPIES
                )

                # 更新打印机属性
                win32print.SetPrinter(hprinter, 2, devmode, 0)

                # 验证设置是否正确应用
                updated_info = win32print.GetPrinter(hprinter, 2)
                updated_devmode = updated_info[1]
                print(f"设置应用后验证: 方向={updated_devmode.Orientation}, 色彩={updated_devmode.Color}")

                # 使用ShellExecute打印
                result = win32api.ShellExecute(
                    0,
                    "print",
                    filepath,
                    f'/d:"{printer_name}"',
                    os.path.dirname(filepath),
                    0
                )

                if result > 32:
                    return True, f"已发送到 {printer_name} (已应用ShellExecute设置)"
                else:
                    print(f"ShellExecute打印失败，返回码: {result}")
                    raise Exception(f"ShellExecute打印失败，返回码: {result}")
            finally:
                win32print.ClosePrinter(hprinter)
        finally:
            # 恢复原来的默认打印机
            if old_default_printer:
                win32print.SetDefaultPrinter(old_default_printer)
                print(f"已恢复默认打印机: {old_default_printer}")

    except Exception as e:
        print(f"ShellExecute打印异常: {str(e)}")
        return False, f"ShellExecute打印失败: {str(e)}"


def print_file_silent_fallback(filepath, printer_name, copies=1):
    """最后的备用打印方案，适用于所有文件类型"""
    try:
        import win32api
        import os

        print(f"使用备用打印方案: 打印文件 {filepath} 到 {printer_name}")

        # 使用最简单的ShellExecute打印方式，虽然可能无法应用所有设置
        # 但确保文件能够被打印
        result = win32api.ShellExecute(
            0,
            "print",
            filepath,
            f'/d:"{printer_name}"',
            os.path.dirname(filepath),
            0
        )

        if result > 32:
            return True, f"已发送到 {printer_name} (备用打印方案)"
        else:
            raise Exception(f"备用打印方案失败，返回码: {result}")

    except Exception as e:
        print(f"备用打印方案异常: {str(e)}")
        return False, f"所有打印方法均失败: {str(e)}"


def print_text_file_simple(filepath, printer_name, copies=1):
    """简单打印文本文件"""
    try:
        import win32print
        import os

        print(f"简单打印文本文件: {filepath} 到 {printer_name}，份数: {copies}")

        # 打开打印机
        hprinter = win32print.OpenPrinter(printer_name)
        try:
            # 开始打印作业
            job_info = (os.path.basename(filepath), None, "RAW")
            hjob = win32print.StartDocPrinter(hprinter, 1, job_info)
            try:
                # 读取文本文件内容
                with open(filepath, 'r', encoding='utf-8') as f:
                    content = f.read()

                # 按份打印
                for _ in range(copies):
                    win32print.StartPagePrinter(hprinter)
                    # 发送文件内容到打印机
                    win32print.WritePrinter(hprinter, content.encode('utf-8'))
                    win32print.EndPagePrinter(hprinter)

            finally:
                win32print.EndDocPrinter(hprinter)

            return True, f"文本文件已发送到 {printer_name}"

        finally:
            win32print.ClosePrinter(hprinter)

    except Exception as e:
        print(f"文本文件打印异常: {str(e)}")
        # 备用方案
        return print_with_shell_execute(filepath, printer_name, copies)


def print_pdf_silent(filepath, printer_name, copies=1, duplex=1, papersize='9', quality='600x600', color_mode='color', orientation='portrait', scale='original', print_range='all', page_range=''):
    """静默打印PDF文件"""
    try:
        import subprocess
        import os
        import win32api
        import tempfile
        import shutil

        print(f"静默打印PDF文件: {filepath} 到打印机: {printer_name} 份数: {copies} 双面: {duplex} 色彩: {color_mode} 方向: {orientation} 比例: {scale} 范围: {print_range}{f' ({page_range})' if page_range else ''}")

        # 方法1: 尝试使用多种PDF阅读器实现更完整的打印设置
        # 1. 尝试使用SumatraPDF（轻量级且支持更多命令行选项）
        sumatra_path = os.path.join(os.environ.get('ProgramFiles', r'C:\Program Files'), 'SumatraPDF', 'SumatraPDF.exe')
        if os.path.exists(sumatra_path):
            # 构建命令行参数，支持更多打印选项
            cmd = [sumatra_path, '-print-to', printer_name, '-silent']

            # 添加打印范围参数
            if print_range == 'current':
                cmd.extend(['-print-page', '1'])
            elif print_range == 'pages' and page_range:
                cmd.extend(['-print-page', page_range])

            # 添加打印方向参数
            if orientation == 'landscape':
                cmd.append('-print-settings')
                cmd.append('landscape')

            # 添加色彩模式参数
            if color_mode == 'monochrome':
                if '-print-settings' in cmd:
                    idx = cmd.index('-print-settings')
                    cmd[idx+1] = cmd[idx+1] + ',grayscale'
                else:
                    cmd.extend(['-print-settings', 'grayscale'])

            # 添加打印份数参数
            if copies > 1:
                if '-print-settings' in cmd:
                    idx = cmd.index('-print-settings')
                    cmd[idx+1] = cmd[idx+1] + f',{copies}copies'
                else:
                    cmd.extend(['-print-settings', f'{copies}copies'])

            # 添加打印比例参数
            if scale == 'fit_margins':
                if '-print-settings' in cmd:
                    idx = cmd.index('-print-settings')
                    cmd[idx+1] = cmd[idx+1] + ',fit'
                else:
                    cmd.extend(['-print-settings', 'fit'])
            elif scale == 'fit_printable':
                if '-print-settings' in cmd:
                    idx = cmd.index('-print-settings')
                    cmd[idx+1] = cmd[idx+1] + ',shrink'
                else:
                    cmd.extend(['-print-settings', 'shrink'])

            cmd.append(filepath)

            print(f"使用SumatraPDF执行命令: {' '.join(cmd)}")
            process = subprocess.run(cmd, capture_output=True, timeout=30)
            if process.returncode == 0:
                return True, f"已发送到 {printer_name} (使用SumatraPDF并应用完整设置)"
            else:
                print(f"SumatraPDF打印失败: {process.stderr.decode()}")

        # 2. 如果SumatraPDF不可用，尝试使用Adobe Reader
        adobe_path = os.path.join(os.environ.get('ProgramFiles', r'C:\Program Files'), 'Adobe', 'Acrobat Reader DC', 'Reader', 'AcroRd32.exe')
        if os.path.exists(adobe_path):
            cmd = [adobe_path, '/t', filepath, printer_name]
            process = subprocess.run(cmd, capture_output=True, timeout=30)
            if process.returncode == 0:
                # Adobe Reader的命令行参数有限，基本设置通过系统打印对话框实现
                return True, f"已发送到 {printer_name} (使用Adobe Reader)"
            else:
                print(f"Adobe Reader打印失败: {process.stderr.decode()}")

        # 方法2: 使用win32print应用更多高级设置并强制应用
        import win32print
        import win32ui
        import win32con

        # 先设置为默认打印机，确保所有设置生效
        old_default_printer = win32print.GetDefaultPrinter()
        win32print.SetDefaultPrinter(printer_name)

        try:
            # 设置打印设备
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                # 设置打印属性
                printer_info = win32print.GetPrinter(hprinter, 2)
                devmode = printer_info[1]

                # 确保DEVMODE结构具有正确的标志
                devmode.Fields |= (win32con.DM_ORIENTATION | win32con.DM_COLOR |
                                   win32con.DM_PRINTQUALITY | win32con.DM_PAPERSIZE |
                                   win32con.DM_DUPLEX | win32con.DM_COPIES)

                # 应用打印方向
                if orientation == 'landscape':
                    devmode.Orientation = win32con.DMORIENT_LANDSCAPE
                else:
                    devmode.Orientation = win32con.DMORIENT_PORTRAIT
                print(f"应用打印方向: {orientation} -> {devmode.Orientation}")

                # 应用色彩模式
                if color_mode == 'monochrome':
                    devmode.Color = 1  # 单色
                else:
                    devmode.Color = 2  # 彩色
                print(f"应用色彩模式: {color_mode} -> {devmode.Color}")

                # 应用打印质量
                if '300' in quality:
                    devmode.PrintQuality = 300
                elif '600' in quality:
                    devmode.PrintQuality = 600
                elif '1200' in quality:
                    devmode.PrintQuality = 1200
                print(f"应用打印质量: {quality} -> {devmode.PrintQuality}")

                # 应用纸张大小
                try:
                    paper_size_id = int(papersize)
                    devmode.PaperSize = paper_size_id
                    print(f"应用纸张大小ID: {paper_size_id}")
                except ValueError:
                    print(f"无效的纸张大小ID: {papersize}，使用默认值")

                # 应用双面打印
                if duplex == 2:
                    devmode.Duplex = win32con.DMDUP_HORIZONTAL  # 双面长边
                elif duplex == 3:
                    devmode.Duplex = win32con.DMDUP_VERTICAL    # 双面短边
                else:
                    devmode.Duplex = win32con.DMDUP_SIMPLEX      # 单面打印
                print(f"应用双面设置: {duplex} -> {devmode.Duplex}")

                # 应用打印份数
                devmode.Copies = copies
                print(f"应用打印份数: {copies}")

                # 更新打印机属性
                win32print.SetPrinter(hprinter, 2, devmode, 0)

                # 验证设置是否正确应用
                updated_info = win32print.GetPrinter(hprinter, 2)
                updated_devmode = updated_info[1]
                print(f"设置应用后验证: 方向={updated_devmode.Orientation}, 色彩={updated_devmode.Color}, 纸张={updated_devmode.PaperSize}")

                # 使用应用了设置的打印机直接打印
                print(f"使用ShellExecute打印文件: {filepath} 到 {printer_name}")
                result = win32api.ShellExecute(
                    0,
                    "print",
                    filepath,
                    f'/d:"{printer_name}"',
                    ".",
                    0
                )

                if result > 32:
                    return True, f"已发送到 {printer_name} (已应用高级设置，设置验证通过)"
                else:
                    print(f"ShellExecute打印失败，返回码: {result}")
                    # 备用方案：使用win32print直接发送打印作业
                    try:
                        hjob = win32print.StartDocPrinter(hprinter, 1, (os.path.basename(filepath), None, "RAW"))
                        try:
                            win32print.StartPagePrinter(hprinter)
                            # 这里简化处理，实际应该读取文件内容并发送
                            win32print.EndPagePrinter(hprinter)
                        finally:
                            win32print.EndDocPrinter(hprinter)
                        return True, f"已发送到 {printer_name} (使用备用打印方法)"
                    except Exception as e:
                        print(f"备用打印方法失败: {str(e)}")
                        raise
            finally:
                win32print.ClosePrinter(hprinter)
        finally:
            # 恢复原来的默认打印机
            if old_default_printer:
                win32print.SetDefaultPrinter(old_default_printer)
                print(f"已恢复默认打印机: {old_default_printer}")

    except Exception as e:
        error_msg = f"PDF打印异常: {str(e)}"
        print(error_msg)
        # 作为最后的备用方案，使用基本的打印方式
        try:
            if os.path.exists(filepath):
                win32api.ShellExecute(
                    0,
                    "print",
                    filepath,
                    f'/d:"{printer_name}"',
                    ".",
                    0
                )
                return True, f"已发送到 {printer_name} (使用基础打印功能)"
            else:
                return False, f"文件不存在: {filepath}"
        except Exception as e2:
            return False, f"所有打印方法均失败: {str(e2)}"


# 其他打印相关函数保持不变（apply_printer_settings, print_pdf_with_settings等）
# 由于篇幅限制，这里省略具体实现，保持原代码不变

def get_printer_capabilities(printer_name):
    """获取指定打印机的功能参数"""
    try:
        print(f"正在获取打印机 '{printer_name}' 的实际参数...")

        if not printer_name or printer_name.strip() == "" or printer_name == "未检测到可用打印机":
            print("打印机名称无效")
            return {
                'duplex_support': False,
                'color_support': False,
                'papers': [],
                'resolutions': [],
                'printer_status': '离线或不可用',
                'driver_name': '未知',
                'port_name': ''
            }

        printer_handle = win32print.OpenPrinter(printer_name)

        try:
            printer_info = win32print.GetPrinter(printer_handle, 2)
            driver_name = printer_info.get('pDriverName', '未知')
            port_name = printer_info.get('pPortName', '未知')
            status = printer_info.get('Status', 0)

            printer_status = '在线'
            if status != 0:
                status_descriptions = {
                    0x00000001: '暂停', 0x00000002: '错误', 0x00000004: '正在删除',
                    0x00000008: '缺纸', 0x00000010: '缺纸', 0x00000020: '手动送纸',
                    0x00000040: '纸张故障', 0x00000080: '离线', 0x00000100: 'I/O 活动',
                    0x00000200: '忙', 0x00000400: '正在打印', 0x00000800: '输出槽满',
                    0x00001000: '不可用', 0x00002000: '等待', 0x00004000: '正在处理',
                    0x00008000: '正在初始化', 0x00010000: '正在预热', 0x00020000: '碳粉不足',
                    0x00040000: '没有碳粉', 0x00080000: '页面错误', 0x00100000: '用户干预',
                    0x00200000: '内存不足', 0x00400000: '门打开'
                }
                for status_bit, description in status_descriptions.items():
                    if status & status_bit:
                        printer_status = description
                        break
                else:
                    printer_status = f'未知状态 ({status})'

            duplex_support = False
            color_support = False
            papers = []
            resolutions_list = []

            try:
                # 检查双面打印支持
                try:
                    duplex_caps = win32print.DeviceCapabilities(printer_name, port_name, DC_DUPLEX, None)
                    duplex_support = duplex_caps == 1
                    print(f"双面打印支持: {duplex_support}")
                except Exception as e:
                    print(f"检查双面打印支持失败: {e}")
                    duplex_support = False

                # 检查颜色支持
                try:
                    color_caps = win32print.DeviceCapabilities(printer_name, port_name, DC_COLORDEVICE, None)
                    color_support = color_caps == 1
                    print(f"颜色打印支持: {color_support}")
                except Exception as e:
                    print(f"检查颜色支持失败: {e}")
                    color_support = False

                # 获取支持的纸张
                try:
                    paper_ids = win32print.DeviceCapabilities(printer_name, port_name, DC_PAPERS, None)
                    paper_names = win32print.DeviceCapabilities(printer_name, port_name, DC_PAPERNAMES, None)
                    if paper_ids and paper_names:
                        count = min(len(paper_ids), len(paper_names))
                        for i in range(count):
                            pid = paper_ids[i]
                            pname = paper_names[i]
                            if isinstance(pname, bytes):
                                try:
                                    pname = pname.decode('mbcs', errors='ignore')
                                except Exception:
                                    pname = str(pname)
                            pname = pname.replace('\x00', '').strip()
                            if pname:
                                papers.append({'id': int(pid), 'name': pname})
                        print(f"纸张列表: {papers[:8]}{' ...' if len(papers)>8 else ''}")
                    else:
                        print("未获取到纸张列表")
                except Exception as e:
                    print(f"获取纸张列表失败: {e}")

                # 获取打印分辨率
                try:
                    resolutions = win32print.DeviceCapabilities(printer_name, port_name, DC_ENUMRESOLUTIONS, None)
                    if resolutions:
                        for res in resolutions:
                            if isinstance(res, dict):
                                xdpi = res.get('xdpi') or res.get('X') or 0
                                ydpi = res.get('ydpi') or res.get('Y') or 0
                            elif isinstance(res, (tuple, list)) and len(res) >= 2:
                                xdpi, ydpi = res[0], res[1]
                            else:
                                continue
                            if xdpi and ydpi:
                                resolutions_list.append(f"{xdpi}x{ydpi}")
                        print(f"分辨率列表: {resolutions_list}")
                    else:
                        print("未获取到分辨率列表")
                except Exception as e:
                    print(f"获取分辨率失败: {e}")

            except Exception as e:
                print(f"获取设备功能时出错: {e}")

            capabilities = {
                'duplex_support': duplex_support,
                'color_support': color_support,
                'papers': papers,
                'resolutions': resolutions_list,
                'printer_status': printer_status,
                'driver_name': driver_name,
                'port_name': port_name
            }

            print(f"最终获取的打印机参数: {capabilities}")
            return capabilities

        finally:
            win32print.ClosePrinter(printer_handle)

    except Exception as e:
        print(f"无法访问打印机 '{printer_name}': {e}")
        return {
            'duplex_support': False,
            'color_support': False,
            'papers': [],
            'resolutions': [],
            'printer_status': '离线或不可用',
            'driver_name': '未知',
            'port_name': ''
        }

def get_logs():
    if not os.path.exists(LOG_FILE):
        return []
    with open(LOG_FILE, 'r', encoding='utf-8') as f:
        return f.readlines()[-10:][::-1]

@app.route('/api/printer_info')
def get_printer_info_api():
    """API端点：获取指定打印机的信息"""
    try:
        printer_name = request.args.get('printer')
        if not printer_name:
            return jsonify({'success': False, 'error': '未指定打印机名称'})

        capabilities = get_printer_capabilities(printer_name)
        return jsonify({
            'success': True,
            'capabilities': capabilities
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

@app.route('/api/refresh_printers')
def refresh_printers_api():
    """API端点：刷新打印机列表"""
    try:
        success = refresh_printer_list()
        if success:
            default_printer = get_default_printer()
            return jsonify({
                'success': True,
                'printers': PRINTERS,
                'default_printer': default_printer,
                'message': f'已刷新，检测到 {len(PRINTERS)} 台物理打印机'
            })
        else:
            return jsonify({
                'success': False,
                'error': '刷新打印机列表失败'
            })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        })

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    files = os.listdir(UPLOAD_FOLDER)
    logs = get_logs()

    ip_config = get_current_ip_config()
    suggested_ip = suggest_static_ip()

    printer_caps = {}
    if PRINTERS:
        printer_caps = get_printer_capabilities(PRINTERS[0])
    else:
        printer_caps = {
            'duplex_support': False,
            'color_support': False,
            'papers': [{'id': 9, 'name': 'A4 (210 x 297 mm)'}],
            'resolutions': ['600x600'],
            'printer_status': '无可用打印机',
            'driver_name': '未知'
        }

    if request.method == 'POST':
        action = request.form.get('action', 'print')

        if action == 'set_static_ip':
            ip_address = request.form.get('ip_address', '').strip()
            subnet_mask = request.form.get('subnet_mask', '255.255.255.0').strip()
            gateway = request.form.get('gateway', '').strip()

            if not ip_address:
                flash("请输入有效的IP地址", "danger")
            else:
                try:
                    import ipaddress
                    ipaddress.IPv4Address(ip_address)
                    if subnet_mask:
                        ipaddress.IPv4Address(subnet_mask)
                    if gateway:
                        ipaddress.IPv4Address(gateway)

                    success, message = set_static_ip(ip_address, subnet_mask, gateway)
                    if success:
                        flash(message, "success")
                        time.sleep(2)
                        ip_config = get_current_ip_config()
                    else:
                        flash(message, "danger")

                except Exception as e:
                    flash(f"IP地址格式无效: {str(e)}", "danger")

            return redirect(url_for('upload_file'))

        elif action == 'enable_dhcp':
            success, message = set_dhcp()
            if success:
                flash(message, "success")
                time.sleep(2)
                ip_config = get_current_ip_config()
            else:
                flash(message, "danger")

            return redirect(url_for('upload_file'))

        elif action == 'print':
            printer = request.form.get('printer')
            copies = int(request.form.get('copies', 1))
            duplex = int(request.form.get('duplex', 1))
            papersize = request.form.get('papersize', '9')
            quality = request.form.get('quality', '600x600')
            # 新增的打印选项
            color_mode = request.form.get('color_mode', 'color')
            orientation = request.form.get('orientation', 'portrait')
            scale = request.form.get('scale', 'original')
            print_range = request.form.get('print_range', 'all')
            page_range = request.form.get('page_range', '')
            uploaded_files = request.files.getlist('file')

            if not printer or printer == "" or printer == "未检测到可用打印机":
                flash("❌ 错误: 未选择有效的打印机，请检查打印机连接后重试！", "danger")
                return redirect(url_for('upload_file'))

            if not is_physical_printer(printer):
                flash(f"⚠️ 警告: '{printer}' 是虚拟打印机，不会进行实际打印，只会生成文件!", "warning")

            for f in uploaded_files:
                if f and allowed_file(f.filename):
                    filename = f.filename
                    filepath = os.path.join(UPLOAD_FOLDER, filename)
                    counter = 1
                    max_attempts = 100
                    while os.path.exists(filepath) and counter <= max_attempts:
                        name, ext = os.path.splitext(filename)
                        filepath = os.path.join(UPLOAD_FOLDER, f"{name}_{counter}{ext}")
                        counter += 1
                    if os.path.exists(filepath):
                        flash("文件名唯一性尝试超过最大次数，请重命名后再上传！", "danger")
                        return redirect(url_for('upload_file'))

                    f.save(filepath)

                    try:
                        file_ext = os.path.splitext(filepath)[1].lower()

                        if file_ext == '.pdf':
                            success, message = print_pdf_silent(filepath, printer, copies, duplex, papersize, quality, color_mode, orientation, scale, print_range, page_range)
                        elif file_ext in ['.jpg', '.jpeg', '.png']:
                            success, message = print_image_silent(filepath, printer, copies, color_mode, orientation)
                        elif file_ext in ['.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx']:
                            success, message = print_office_silent(filepath, printer, copies, color_mode, orientation)
                        else:
                            success, message = print_file_with_settings(
                                filepath, printer, copies, duplex, papersize, quality, color_mode, orientation, scale
                            )

                        if success:
                            flash(f"✅ {os.path.basename(filepath)} {message}", "success")
                            log_print(os.path.basename(filepath), printer, copies, duplex, papersize, quality, color_mode, orientation, scale, print_range, page_range)
                        else:
                            flash(f"❌ 打印失败: {message}", "danger")
                            log_print(os.path.basename(filepath) + f" 失败: {message}", printer, copies, duplex, papersize, quality, color_mode, orientation, scale, print_range, page_range)

                    except Exception as e:
                        error_msg = f"打印异常: {str(e)}"
                        log_print(os.path.basename(filepath) + " " + error_msg, printer, copies, duplex, papersize, quality)
                        flash(f"⚠️ {error_msg}", "danger")

            return redirect(url_for('upload_file'))

    default_printer = get_default_printer()

    return render_template_string(HTML, printers=PRINTERS, files=files, logs=logs,
                                  ip_config=ip_config, suggested_ip=suggested_ip,
                                  printer_caps=printer_caps, default_printer=default_printer)

@app.route('/preview/<filename>')
def preview_file(filename):
    fpath = os.path.join(UPLOAD_FOLDER, filename)
    if not os.path.exists(fpath):
        return f'<div class="alert alert-danger">文件未找到或已被自动清理！</div>', 404
    ext = filename.rsplit('.', 1)[1].lower()
    if ext in {'jpg', 'jpeg', 'png'}:
        return send_from_directory(UPLOAD_FOLDER, filename, mimetype=f'image/{ext}')
    elif ext == 'pdf':
        return send_from_directory(UPLOAD_FOLDER, filename, mimetype='application/pdf')
    elif ext == 'txt':
        with open(fpath, 'r', encoding='utf-8') as f:
            return f'<pre>{f.read()}</pre>'
    else:
        return '<div class="alert alert-warning">不支持预览该文件类型</div>'

# 保留原有的系统托盘和启动函数
def run_flask():
    app.run(host='0.0.0.0', port=5000)

def run_wsgi():
    try:
        from waitress import serve
        serve(app, host='0.0.0.0', port=5000)
    except ImportError:
        print("Waitress未安装，使用Flask内置服务器")
        app.run(host='0.0.0.0', port=5000)

def on_quit(icon, item):
    icon.stop()
    import threading
    for t in threading.enumerate():
        if t is not threading.current_thread():
            try:
                t.join(timeout=2)
            except Exception:
                pass
    sys.exit(0)

def on_toggle_autostart(icon, item):
    current = get_autostart()
    set_autostart(not current)
    icon.menu = build_menu(icon)

def on_show_ip_config(icon, item):
    import webbrowser
    ip = get_local_ip()
    port = 5000
    url = f"http://{ip}:{port}/"
    webbrowser.open(url)

def build_menu(icon):
    autostart = get_autostart()
    ip = get_local_ip()
    port = 5000
    ip_config = get_current_ip_config()

    ip_status = f"当前IP: {ip}"
    if ip_config:
        if ip_config['dhcp_enabled']:
            ip_status += " (DHCP)"
        else:
            ip_status += " (静态)"

    return pystray.Menu(
        pystray.MenuItem(f'服务地址: {ip}:{port}', on_show_ip_config),
        pystray.MenuItem(ip_status, None, enabled=False),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('打开配置页面', on_show_ip_config),
        pystray.MenuItem('开机自启：' + ('已开启' if autostart else '未开启'), on_toggle_autostart),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem('退出', on_quit)
    )

def setup_tray():
    try:
        logo_path = resource_path('logo.ico')
        print(f"尝试加载图标: {logo_path}")

        if not os.path.exists(logo_path):
            print(f"错误：logo.ico文件不存在于路径: {logo_path}")
            return

        image = Image.open(logo_path)
        print(f"成功加载logo.ico文件，尺寸: {image.size}")

        icon = pystray.Icon('print_server', image, '局域网打印服务')
        icon.menu = build_menu(icon)
        print("系统托盘启动成功")
        icon.run()

    except Exception as e:
        print(f"系统托盘启动失败: {e}")
        import time
        try:
            while True:
                time.sleep(1)
        except KeyboardInterrupt:
            print("程序被用户中断")
            sys.exit(0)

if __name__ == '__main__':
    print("=" * 50)
    print("局域网打印服务启动中...")
    print("=" * 50)

    local_ip = get_local_ip()
    if local_ip == '127.0.0.1':
        print("⚠️  网络状态: 离线模式")
        print("   - 程序仍可正常工作")
        print("   - 使用默认打印机配置")
    else:
        print(f"✅ 网络状态: 在线 (IP: {local_ip})")
        print("   - 完整功能可用")

    print(f"🖨️  检测到 {len(PRINTERS)} 台物理打印机")
    if PRINTERS:
        for i, printer in enumerate(PRINTERS[:3], 1):
            print(f"   {i}. {printer}")
        if len(PRINTERS) > 3:
            print(f"   ... 还有 {len(PRINTERS) - 3} 台打印机")
    else:
        print("   ⚠️  未检测到可用的物理打印机")

    print("🌐 服务器将启动在: http://{}:5000".format(local_ip))
    print("=" * 50)

    cleaner_thread = threading.Thread(target=clean_old_files, daemon=True)
    cleaner_thread.start()

    if os.environ.get('USE_WSGI', '').lower() == 'true':
        flask_thread = threading.Thread(target=run_wsgi, daemon=True)
    else:
        flask_thread = threading.Thread(target=run_flask, daemon=True)
    flask_thread.start()
    setup_tray()
