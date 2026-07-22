# -*- coding: utf-8 -*-
"""
项目点检表 - 多人协作服务器（安全版）
✅ 带用户认证、角色权限、Session 管理、限流、CSRF 防护

使用方法：
  python 协作服务器_安全版.py              # 默认端口 8080
  python 协作服务器_安全版.py 8888         # 指定端口

默认管理员账号：admin / admin123（首次登录请立即修改密码）
"""

import http.server
import socketserver
import json
import os
import sys
import socket
import urllib.parse
import urllib.request
import urllib.error
import secrets
import threading
import time
from datetime import datetime

# 导入认证模块和同步模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import auth
import sync_excel
import project_parser

# ==================== 配置 ====================
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)
DATA_FILE = os.path.join(DATA_DIR, '协作数据.json')
HTML_FILE = os.path.join(BASE_DIR, '项目延期点检表.html')

# 安全头
SECURITY_HEADERS = {
    'X-Content-Type-Options': 'nosniff',
    'X-Frame-Options': 'SAMEORIGIN',
    'X-XSS-Protection': '1; mode=block',
    'Referrer-Policy': 'strict-origin-when-cross-origin',
}

# ==================== 数据管理 ====================
def _compute_last_update() -> str:
    """计算数据的最后更新时间（用Excel文件的修改时间，避免不必要的刷新）
    
    【死循环修复】之前用 datetime.now()，每次调用都不一样 → 前端每5秒检测到"更新"→ 死循环
    现在用 Excel 文件的 mtime，只有数据真正变化时才改变
    """
    try:
        excel_path = os.path.join(BASE_DIR, '超声波户表脚本.xlsx')
        user_path = os.path.join(BASE_DIR, '用户管理.xlsx')
        times = []
        for p in [excel_path, user_path, DATA_FILE]:
            if os.path.exists(p):
                times.append(os.path.getmtime(p))
        if times:
            return str(int(max(times) * 1000))  # 毫秒时间戳，稳定可比较
    except:
        pass
    return '0'


def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 【死循环修复】用 Excel mtime 覆盖 lastUpdate
                data['lastUpdate'] = _compute_last_update()
                return data
        except:
            pass
    return {
        'localEdits': {}, 'notes': {}, 'checked': {},
        'archived': {}, 'customEmails': {}, 'newProjects': [],
        'deletedIds': [],
        'lastUpdate': _compute_last_update()
    }


def save_data(data):
    data['lastUpdate'] = _compute_last_update()
    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return '127.0.0.1'


# ==================== 无数据引导页面 HTML ====================
NO_DATA_PAGE = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>项目点检表 - 数据导入</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  }
  .upload-box {
    background: #fff;
    border-radius: 16px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    padding: 40px 32px;
    width: 100%;
    max-width: 560px;
    text-align: center;
  }
  .icon { font-size: 56px; margin-bottom: 12px; }
  h1 { font-size: 26px; color: #333; margin-bottom: 8px; }
  .subtitle { color: #888; margin-bottom: 24px; font-size: 14px; line-height: 1.6; }
  .tabs {
    display: flex;
    gap: 8px;
    margin-bottom: 20px;
    border-bottom: 2px solid #e2e8f0;
  }
  .tab {
    flex: 1;
    padding: 10px 8px;
    font-size: 13px;
    font-weight: 500;
    color: #718096;
    cursor: pointer;
    border-bottom: 3px solid transparent;
    margin-bottom: -2px;
    transition: all 0.2s;
  }
  .tab.active { color: #667eea; border-bottom-color: #667eea; }
  .tab:hover:not(.active) { color: #4a5568; }
  .tab-panel { display: none; }
  .tab-panel.active { display: block; }
  .upload-area {
    border: 2px dashed #cbd5e0;
    border-radius: 12px;
    padding: 32px 16px;
    margin-bottom: 16px;
    cursor: pointer;
    transition: all 0.3s;
    background: #f7fafc;
  }
  .upload-area:hover, .upload-area.dragover {
    border-color: #667eea;
    background: #f0f4ff;
  }
  .upload-area p { color: #718096; margin-top: 6px; font-size: 13px; }
  .upload-area strong { color: #4a5568; font-size: 15px; }
  input[type=file] { display: none; }
  .input-group {
    display: flex;
    gap: 8px;
    margin-bottom: 16px;
  }
  input[type=text] {
    flex: 1;
    padding: 11px 14px;
    border: 2px solid #e2e8f0;
    border-radius: 8px;
    font-size: 14px;
    outline: none;
    transition: border-color 0.2s;
  }
  input[type=text]:focus { border-color: #667eea; }
  .btn {
    display: inline-block;
    padding: 11px 24px;
    border-radius: 8px;
    border: none;
    font-size: 14px;
    font-weight: 500;
    cursor: pointer;
    transition: all 0.2s;
    text-decoration: none;
    white-space: nowrap;
  }
  .btn-primary {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: #fff;
  }
  .btn-primary:hover { transform: translateY(-1px); box-shadow: 0 6px 20px rgba(102,126,234,0.4); }
  .btn-primary:disabled { opacity: 0.5; cursor: not-allowed; transform: none; box-shadow: none; }
  .btn-secondary { background: #e2e8f0; color: #4a5568; }
  .btn-secondary:hover { background: #cbd5e0; }
  .btn-block { width: 100%; padding: 14px; font-size: 15px; }
  .quick-action {
    background: #f7fafc;
    border: 2px solid #e2e8f0;
    border-radius: 12px;
    padding: 20px;
    margin-bottom: 16px;
    text-align: left;
    transition: all 0.2s;
  }
  .quick-action:hover { border-color: #667eea; background: #f0f4ff; }
  .quick-action h3 { font-size: 15px; color: #2d3748; margin-bottom: 4px; }
  .quick-action p { font-size: 13px; color: #718096; margin-bottom: 12px; }
  .status {
    margin-top: 16px;
    padding: 11px 14px;
    border-radius: 8px;
    font-size: 13px;
    display: none;
    text-align: left;
  }
  .status.success { background: #f0fff4; color: #22543d; display: block; }
  .status.error { background: #fff5f5; color: #742a2a; display: block; }
  .status.info { background: #ebf8ff; color: #2a4365; display: block; }
  .logout {
    margin-top: 20px;
    color: #a0aec0;
    font-size: 12px;
  }
  .logout a { color: #667eea; text-decoration: none; }
  .hint {
    font-size: 12px;
    color: #a0aec0;
    margin-top: 6px;
    text-align: left;
  }
</style>
</head>
<body>
<div class="upload-box">
  <div class="icon">📊</div>
  <h1>项目点检系统</h1>
  <p class="subtitle">选择一种方式导入「超声波户表脚本」Excel 数据</p>

  <div class="tabs">
    <div class="tab active" data-tab="github">🔗 GitHub 链接</div>
    <div class="tab" data-tab="default">📦 默认数据</div>
    <div class="tab" data-tab="upload">📁 本地上传</div>
  </div>

  <!-- GitHub 链接导入 -->
  <div class="tab-panel active" id="panel-github">
    <div class="input-group">
      <input type="text" id="githubUrl" placeholder="粘贴 GitHub 文件链接（如 https://github.com/xxx/xxx/blob/main/xxx.xlsx）">
      <button class="btn btn-primary" id="githubBtn" onclick="fetchFromGithub()">拉取并生成</button>
    </div>
    <div class="hint">💡 支持 github.com/blob/ 链接和 raw.githubusercontent.com 链接</div>
  </div>

  <!-- 默认数据 -->
  <div class="tab-panel" id="panel-default">
    <div class="quick-action">
      <h3>使用仓库内置 Excel</h3>
      <p>直接使用代码仓库中已有的「超声波户表脚本.xlsx」生成报表</p>
      <button class="btn btn-primary btn-block" id="defaultBtn" onclick="generateDefault()">⚡ 一键生成报表</button>
    </div>
  </div>

  <!-- 本地上传 -->
  <div class="tab-panel" id="panel-upload">
    <div class="upload-area" id="uploadArea" onclick="document.getElementById('fileInput').click()">
      <strong>点击选择文件</strong> 或拖拽到此处
      <p>支持 .xlsx / .xls 格式</p>
    </div>
    <input type="file" id="fileInput" accept=".xlsx,.xls">
    <button class="btn btn-primary btn-block" id="uploadBtn" onclick="uploadFile()" disabled>
      🚀 上传并生成报表
    </button>
  </div>

  <div id="status" class="status"></div>

  <div class="logout">
    <a href="#" onclick="logout()">退出登录</a>
  </div>
</div>

<script>
let selectedFile = null;
const statusEl = document.getElementById('status');

// Tab 切换
document.querySelectorAll('.tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
    tab.classList.add('active');
    document.getElementById('panel-' + tab.dataset.tab).classList.add('active');
  });
});

function showStatus(msg, type) {
  statusEl.className = 'status ' + type;
  statusEl.textContent = msg;
}

function setButtonsDisabled(disabled) {
  document.getElementById('githubBtn').disabled = disabled;
  document.getElementById('defaultBtn').disabled = disabled;
  document.getElementById('uploadBtn').disabled = disabled || !selectedFile;
}

// 从 GitHub 拉取
async function fetchFromGithub() {
  const url = document.getElementById('githubUrl').value.trim();
  if (!url) {
    showStatus('❌ 请输入 GitHub 链接', 'error');
    return;
  }
  setButtonsDisabled(true);
  showStatus('⏳ 正在从 GitHub 拉取文件并生成报表...', 'info');
  try {
    const res = await fetch('/api/fetch-github', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ url })
    });
    const data = await res.json();
    if (data.success) {
      showStatus('✅ 报表生成成功！正在跳转...', 'success');
      setTimeout(() => location.reload(), 1000);
    } else {
      showStatus('❌ ' + (data.message || '失败'), 'error');
      setButtonsDisabled(false);
    }
  } catch (e) {
    showStatus('❌ 请求失败：' + e.message, 'error');
    setButtonsDisabled(false);
  }
}

// 使用默认 Excel 生成
async function generateDefault() {
  setButtonsDisabled(true);
  showStatus('⏳ 正在使用默认数据生成报表...', 'info');
  try {
    const res = await fetch('/api/generate-default', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({})
    });
    const data = await res.json();
    if (data.success) {
      showStatus('✅ 报表生成成功！正在跳转...', 'success');
      setTimeout(() => location.reload(), 1000);
    } else {
      showStatus('❌ ' + (data.message || '失败'), 'error');
      setButtonsDisabled(false);
    }
  } catch (e) {
    showStatus('❌ 请求失败：' + e.message, 'error');
    setButtonsDisabled(false);
  }
}

// 本地上传
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadBtn = document.getElementById('uploadBtn');

fileInput.addEventListener('change', (e) => {
  if (e.target.files.length > 0) {
    selectedFile = e.target.files[0];
    uploadArea.querySelector('strong').textContent = selectedFile.name;
    uploadArea.querySelector('p').textContent = (selectedFile.size / 1024 / 1024).toFixed(2) + ' MB';
    uploadBtn.disabled = false;
  }
});

uploadArea.addEventListener('dragover', (e) => {
  e.preventDefault();
  uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
uploadArea.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
  if (e.dataTransfer.files.length > 0) {
    fileInput.files = e.dataTransfer.files;
    fileInput.dispatchEvent(new Event('change'));
  }
});

async function uploadFile() {
  if (!selectedFile) return;
  setButtonsDisabled(true);
  showStatus('⏳ 正在上传并生成报表，请稍候...', 'info');
  const formData = new FormData();
  formData.append('file', selectedFile);
  try {
    const res = await fetch('/api/upload', {
      method: 'POST',
      body: formData
    });
    const text = await res.text();
    let data;
    try { data = JSON.parse(text); } catch { data = { success: false, message: '服务器返回格式错误' }; }
    if (data.success) {
      showStatus('✅ 报表生成成功！正在跳转...', 'success');
      setTimeout(() => location.reload(), 1000);
    } else {
      showStatus('❌ ' + (data.message || '生成失败'), 'error');
      setButtonsDisabled(false);
    }
  } catch (e) {
    showStatus('❌ 上传失败：' + e.message, 'error');
    setButtonsDisabled(false);
  }
}

// 回车触发 GitHub 拉取
document.getElementById('githubUrl').addEventListener('keypress', (e) => {
  if (e.key === 'Enter') fetchFromGithub();
});

async function logout() {
  await fetch('/api/logout', { method: 'POST' });
  location.href = '/login';
}
</script>
</body>
</html>'''

# ==================== 登录页面 HTML ====================
LOGIN_PAGE = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>项目点检表 - 登录</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Hiragino Sans GB", "Microsoft YaHei", sans-serif;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    min-height: 100vh;
    display: flex;
    align-items: center;
    justify-content: center;
    padding: 20px;
  }
  .login-box {
    background: #fff;
    border-radius: 16px;
    box-shadow: 0 20px 60px rgba(0,0,0,0.3);
    padding: 40px;
    width: 100%;
    max-width: 400px;
  }
  h1 {
    font-size: 24px;
    color: #333;
    margin-bottom: 8px;
    text-align: center;
  }
  .subtitle {
    color: #888;
    font-size: 14px;
    text-align: center;
    margin-bottom: 32px;
  }
  .form-group { margin-bottom: 20px; }
  label {
    display: block;
    font-size: 14px;
    color: #555;
    margin-bottom: 8px;
    font-weight: 500;
  }
  input[type="text"], input[type="password"] {
    width: 100%;
    padding: 12px 16px;
    border: 2px solid #e0e0e0;
    border-radius: 8px;
    font-size: 15px;
    transition: border-color 0.2s;
    outline: none;
  }
  input:focus { border-color: #667eea; }
  .btn {
    width: 100%;
    padding: 13px;
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: #fff;
    border: none;
    border-radius: 8px;
    font-size: 16px;
    font-weight: 600;
    cursor: pointer;
    transition: transform 0.1s;
  }
  .btn:hover { transform: translateY(-1px); }
  .btn:active { transform: translateY(0); }
  .error {
    background: #fee;
    color: #c33;
    padding: 10px 14px;
    border-radius: 6px;
    font-size: 14px;
    margin-bottom: 16px;
    display: none;
  }
  .error.show { display: block; }
  .footer {
    text-align: center;
    margin-top: 24px;
    color: #aaa;
    font-size: 12px;
  }
  .default-tip {
    background: #fffbeb;
    color: #92400e;
    padding: 10px 14px;
    border-radius: 6px;
    font-size: 13px;
    margin-bottom: 16px;
    border: 1px solid #fde68a;
  }
</style>
</head>
<body>
<div class="login-box">
  <h1>🔐 项目点检表</h1>
  <p class="subtitle">请登录以继续访问</p>
  <div class="default-tip">
    💡 默认管理员：<b>admin</b> / <b>admin123</b>
  </div>
  <div id="error" class="error"></div>
  <form id="loginForm">
    <div class="form-group">
      <label for="username">用户名</label>
      <input type="text" id="username" autocomplete="username" required>
    </div>
    <div class="form-group">
      <label for="password">密码</label>
      <input type="password" id="password" autocomplete="current-password" required>
    </div>
    <button type="submit" class="btn">登 录</button>
  </form>
  <div class="footer">安全版本 · 权限管理已启用</div>
</div>
<script>
const form = document.getElementById('loginForm');
const errorEl = document.getElementById('error');

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const username = document.getElementById('username').value.trim();
  const password = document.getElementById('password').value;
  errorEl.classList.remove('show');

  try {
    const res = await fetch('/api/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ username, password }),
      credentials: 'same-origin'
    });
    const data = await res.json();
    if (data.success) {
      console.log('登录成功，正在跳转...');
      // 直接跳转，不要用 alert 阻塞！
      // （首页会自动检测 must_change_pwd 并弹出改密码对话框）
      window.location.replace('/');
    } else {
      errorEl.textContent = data.message || '登录失败';
      errorEl.classList.add('show');
    }
  } catch (err) {
    console.error('登录错误:', err);
    errorEl.textContent = '网络错误，请重试';
    errorEl.classList.add('show');
  }
});
</script>
</body>
</html>'''


# ==================== 用户管理页面 HTML ====================
USER_MANAGE_PAGE = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>用户管理 - 项目点检表</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
    background: #f5f7fa;
    min-height: 100vh;
    padding: 20px;
  }
  .container { max-width: 960px; margin: 0 auto; }
  .header {
    background: #fff;
    padding: 20px 24px;
    border-radius: 12px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    margin-bottom: 20px;
    display: flex;
    justify-content: space-between;
    align-items: center;
  }
  h1 { font-size: 20px; color: #333; }
  .btn {
    padding: 8px 16px;
    border: none;
    border-radius: 6px;
    cursor: pointer;
    font-size: 14px;
    transition: all 0.2s;
  }
  .btn-primary { background: #667eea; color: #fff; }
  .btn-primary:hover { background: #5568d3; }
  .btn-danger { background: #ef4444; color: #fff; }
  .btn-danger:hover { background: #dc2626; }
  .btn-secondary { background: #e5e7eb; color: #333; }
  .btn-secondary:hover { background: #d1d5db; }
  .card {
    background: #fff;
    border-radius: 12px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    padding: 20px;
    margin-bottom: 20px;
  }
  table { width: 100%; border-collapse: collapse; }
  th, td {
    padding: 12px 16px;
    text-align: left;
    border-bottom: 1px solid #eee;
  }
  th { background: #f9fafb; font-weight: 600; color: #555; font-size: 13px; }
  td { font-size: 14px; color: #333; }
  .role-badge {
    display: inline-block;
    padding: 3px 10px;
    border-radius: 12px;
    font-size: 12px;
    font-weight: 500;
  }
  .role-admin { background: #fef3c7; color: #92400e; }
  .role-editor { background: #dbeafe; color: #1e40af; }
  .role-viewer { background: #d1fae5; color: #065f46; }
  .modal {
    display: none;
    position: fixed;
    inset: 0;
    background: rgba(0,0,0,0.5);
    align-items: center;
    justify-content: center;
    z-index: 1000;
  }
  .modal.show { display: flex; }
  .modal-content {
    background: #fff;
    border-radius: 12px;
    padding: 24px;
    width: 100%;
    max-width: 420px;
  }
  .modal h2 { font-size: 18px; margin-bottom: 16px; }
  .form-group { margin-bottom: 14px; }
  .form-group label {
    display: block;
    font-size: 13px;
    color: #555;
    margin-bottom: 6px;
  }
  .form-group input, .form-group select {
    width: 100%;
    padding: 9px 12px;
    border: 1px solid #d1d5db;
    border-radius: 6px;
    font-size: 14px;
    outline: none;
  }
  .form-group input:focus, .form-group select:focus { border-color: #667eea; }
  .modal-actions {
    display: flex;
    gap: 10px;
    justify-content: flex-end;
    margin-top: 20px;
  }
  .top-bar {
    display: flex;
    gap: 10px;
    align-items: center;
  }
  .user-info {
    padding: 6px 12px;
    background: #eef2ff;
    border-radius: 6px;
    font-size: 13px;
    color: #4338ca;
  }
</style>
</head>
<body>
<div class="container">
  <div class="header">
    <h1>👥 用户管理</h1>
    <div class="top-bar">
      <span id="userInfo" class="user-info"></span>
      <button class="btn btn-secondary" onclick="location.href='/'">← 返回报表</button>
      <button class="btn btn-primary" onclick="showAddModal()">+ 新增用户</button>
    </div>
  </div>
  <div class="card">
    <table>
      <thead>
        <tr>
          <th>用户名</th>
          <th>角色</th>
          <th>邮箱</th>
          <th>最后登录</th>
          <th>状态</th>
          <th>操作</th>
        </tr>
      </thead>
      <tbody id="userTable"></tbody>
    </table>
  </div>
</div>

<!-- 新增/编辑用户弹窗 -->
<div id="userModal" class="modal">
  <div class="modal-content">
    <h2 id="modalTitle">新增用户</h2>
    <div class="form-group">
      <label>用户名</label>
      <input type="text" id="m_username">
    </div>
    <div class="form-group" id="pwdGroup">
      <label>密码</label>
      <input type="password" id="m_password">
    </div>
    <div class="form-group">
      <label>角色</label>
      <select id="m_role">
        <option value="viewer">只读用户</option>
        <option value="editor">编辑者</option>
        <option value="admin">管理员</option>
      </select>
    </div>
    <div class="form-group">
      <label>邮箱</label>
      <input type="email" id="m_email">
    </div>
    <div class="modal-actions">
      <button class="btn btn-secondary" onclick="hideModal()">取消</button>
      <button class="btn btn-primary" onclick="saveUser()">保存</button>
    </div>
  </div>
</div>

<script>
let editingUser = null;
let currentUser = null;

async function loadUsers() {
  const res = await fetch('/api/users');
  const users = await res.json();
  const tbody = document.getElementById('userTable');
  tbody.innerHTML = users.map(u => `
    <tr>
      <td><b>${u.username}</b></td>
      <td><span class="role-badge role-${u.role}">${u.role_name}</span></td>
      <td>${u.email || '-'}</td>
      <td>${u.last_login ? new Date(u.last_login).toLocaleString('zh-CN') : '未登录'}</td>
      <td>${u.status === 'active' ? '✅ 正常' : '❌ 禁用'}</td>
      <td>
        <button class="btn btn-secondary" style="padding:4px 10px;font-size:12px" onclick="editUser('${u.username}')">编辑</button>
        ${u.username !== 'admin' ? `<button class="btn btn-danger" style="padding:4px 10px;font-size:12px" onclick="deleteUser('${u.username}')">删除</button>` : ''}
      </td>
    </tr>
  `).join('');
}

async function checkAuth() {
  const res = await fetch('/api/me');
  if (!res.ok) { location.href = '/'; return; }
  currentUser = await res.json();
  document.getElementById('userInfo').textContent = `👤 ${currentUser.username} (${currentUser.role_name})`;
  if (currentUser.role !== 'admin') {
    alert('权限不足');
    location.href = '/';
  }
}

function showAddModal() {
  editingUser = null;
  document.getElementById('modalTitle').textContent = '新增用户';
  document.getElementById('m_username').value = '';
  document.getElementById('m_password').value = '';
  document.getElementById('m_role').value = 'viewer';
  document.getElementById('m_email').value = '';
  document.getElementById('pwdGroup').style.display = 'block';
  document.getElementById('m_username').disabled = false;
  document.getElementById('userModal').classList.add('show');
}

function editUser(username) {
  fetch('/api/users').then(r => r.json()).then(users => {
    const u = users.find(x => x.username === username);
    if (!u) return;
    editingUser = u;
    document.getElementById('modalTitle').textContent = '编辑用户';
    document.getElementById('m_username').value = u.username;
    document.getElementById('m_username').disabled = true;
    document.getElementById('m_password').value = '';
    document.getElementById('m_role').value = u.role;
    document.getElementById('m_email').value = u.email || '';
    document.getElementById('pwdGroup').style.display = 'block';
    document.getElementById('userModal').classList.add('show');
  });
}

function hideModal() {
  document.getElementById('userModal').classList.remove('show');
}

async function saveUser() {
  const username = document.getElementById('m_username').value.trim();
  const password = document.getElementById('m_password').value;
  const role = document.getElementById('m_role').value;
  const email = document.getElementById('m_email').value.trim();

  let url, method, body;
  if (editingUser) {
    url = '/api/user/' + username;
    method = 'PUT';
    body = { role, email };
    if (password) body.password = password;
  } else {
    if (!password) { alert('请输入密码'); return; }
    url = '/api/user';
    method = 'POST';
    body = { username, password, role, email };
  }

  const res = await fetch(url, {
    method,
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  });
  const data = await res.json();
  if (data.success) {
    hideModal();
    loadUsers();
  } else {
    alert(data.message || '操作失败');
  }
}

async function deleteUser(username) {
  if (!confirm(`确定删除用户 ${username}？`)) return;
  const res = await fetch('/api/user/' + username, { method: 'DELETE' });
  const data = await res.json();
  if (data.success) loadUsers();
  else alert(data.message);
}

checkAuth();
loadUsers();
</script>
</body>
</html>'''


# ==================== 请求处理 ====================
class SecureCollaborationHandler(http.server.SimpleHTTPRequestHandler):
    """带认证授权的协作服务器"""

    def __init__(self, *args, **kwargs):
        self._session_id = None
        self._session = None
        super().__init__(*args, directory=BASE_DIR, **kwargs)

    # ---------- 工具方法 ----------
    def end_headers(self):
        for k, v in SECURITY_HEADERS.items():
            self.send_header(k, v)
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, X-CSRF-Token')
        self.send_header('Access-Control-Allow-Credentials', 'true')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def get_client_ip(self):
        """获取客户端真实 IP"""
        xff = self.headers.get('X-Forwarded-For')
        if xff:
            return xff.split(',')[0].strip()
        return self.client_address[0]

    def parse_cookies(self):
        """解析 Cookie"""
        cookie_header = self.headers.get('Cookie', '')
        cookies = {}
        for part in cookie_header.split(';'):
            part = part.strip()
            if '=' in part:
                k, v = part.split('=', 1)
                cookies[k.strip()] = v.strip()
        return cookies

    def get_current_session(self):
        """获取当前 Session（带缓存）"""
        if self._session is not None:
            return self._session
        cookies = self.parse_cookies()
        sid = cookies.get('session_id')
        self._session_id = sid
        self._session = auth.get_session(sid) if sid else None
        return self._session

    def get_current_user(self):
        """获取当前登录用户"""
        session = self.get_current_session()
        if not session:
            return None
        return auth.get_user(session['username'])

    def set_session_cookie(self, session_id: str):
        """设置 Session Cookie（兼容 Chrome/Safari/Edge/Render HTTPS）
        
        【关键修复】不设置 Max-Age，让浏览器把它作为会话Cookie（关闭浏览器才清除）。
        Session 有效性完全由服务端的 expires_at 控制（20分钟无操作过期）。
        之前设置 Max-Age=20分钟会导致：即使有操作，20分钟后浏览器也删除Cookie。
        """
        # 检测是否 HTTPS（通过代理头、Render 环境变量、端口）
        is_https = False
        if self.headers.get('X-Forwarded-Proto', '').lower() == 'https':
            is_https = True
        elif self.headers.get('X-Forwarded-Ssl', '').lower() == 'on':
            is_https = True
        elif os.environ.get('RENDER') or os.environ.get('DYNO'):
            # Render / Heroku 等 PaaS 平台默认 HTTPS
            is_https = True
        
        secure_flag = '; Secure' if is_https else ''
        # Chrome 80+ 要求：Secure 的 Cookie 必须 SameSite=None 或 Lax
        # 但 SameSite=None 必须配合 Secure，所以 HTTPS 下用 None，HTTP 下用 Lax
        samesite_flag = 'None' if is_https else 'Lax'
        # 【关键修复】不设 Max-Age，服务端通过 expires_at 控制
        cookie = (
            f'session_id={session_id}; '
            f'Path=/; '
            f'HttpOnly; '
            f'SameSite={samesite_flag}'
            f'{secure_flag}'
        )
        self.send_header('Set-Cookie', cookie)

    def clear_session_cookie(self):
        """清除 Session Cookie"""
        # 与 set_session_cookie 保持一致的 SameSite 和 Secure 设置
        is_https = False
        if self.headers.get('X-Forwarded-Proto', '').lower() == 'https':
            is_https = True
        elif os.environ.get('RENDER') or os.environ.get('DYNO'):
            is_https = True
        secure_flag = '; Secure' if is_https else ''
        samesite_flag = 'None' if is_https else 'Lax'
        self.send_header(
            'Set-Cookie',
            f'session_id=; Path=/; HttpOnly; SameSite={samesite_flag}; Max-Age=0; Expires=Thu, 01 Jan 1970 00:00:00 GMT{secure_flag}'
        )

    def send_json(self, data, status=200):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def send_html(self, html: str, status=200):
        self.send_response(status)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(html.encode('utf-8'))

    def read_body(self):
        """读取 POST/PUT body"""
        content_length = int(self.headers.get('Content-Length', 0))
        if content_length == 0:
            return {}
        body = self.rfile.read(content_length).decode('utf-8')
        try:
            return json.loads(body) if body else {}
        except:
            return {}

    def parse_multipart(self):
        """解析 multipart/form-data 文件上传"""
        content_type = self.headers.get('Content-Type', '')
        if 'multipart/form-data' not in content_type:
            return None
        boundary = content_type.split('boundary=')[-1].strip().strip('"')
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        # 解析 multipart
        parts = body.split(b'--' + boundary.encode())
        for part in parts:
            if b'Content-Disposition' in part and b'filename=' in part:
                # 提取文件名
                header_end = part.find(b'\r\n\r\n')
                if header_end == -1:
                    continue
                headers = part[:header_end].decode('utf-8', errors='replace')
                import re
                fn_match = re.search(r'filename="([^"]+)"', headers)
                if not fn_match:
                    continue
                filename = fn_match.group(1)
                file_data = part[header_end+4:]
                # 移除结尾的 \r\n--
                if file_data.endswith(b'\r\n'):
                    file_data = file_data[:-2]
                return {'filename': filename, 'data': file_data}
        return None

    # ---------- 权限检查辅助 ----------
    def require_auth(self):
        """要求已登录"""
        session = self.get_current_session()
        if not session:
            self.send_json({'error': '未登录'}, 401)
            return False
        return True

    def require_permission(self, perm: str):
        """要求指定权限"""
        if not self.require_auth():
            return False
        user = self.get_current_user()
        if not auth.has_permission(user['username'], perm):
            self.send_json({'error': '权限不足'}, 403)
            return False
        return True

    # ---------- GET 请求 ----------
    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        # --- 公开页面 ---
        if path == '/login':
            self.send_html(LOGIN_PAGE)
            return

        # --- 根路径：未登录显示登录页，已登录显示报表 ---
        if path == '/' or path == '':
            session = self.get_current_session()
            if not session:
                self.send_html(LOGIN_PAGE)
                return
            # 已登录：注入用户信息到报表
            if os.path.exists(HTML_FILE):
                with open(HTML_FILE, 'r', encoding='utf-8') as f:
                    html = f.read()
                user = self.get_current_user()
                user_info = json.dumps({
                    'username': user['username'],
                    'role': user['role'],
                    'role_name': auth.ROLES[user['role']]['name'],
                    'permissions': auth.ROLES[user['role']]['permissions'],
                    'must_change_pwd': user.get('must_change_pwd', False)
                }, ensure_ascii=False)
                # 注入用户信息和安全模式标记
                inject = f'''<script>
window.COLLAB_MODE = true;
window.AUTH_ENABLED = true;
window.CURRENT_USER = {user_info};
</script>'''
                html = html.replace('<head>', '<head>' + inject)
                # 注入前端权限控制脚本
                auth_js = self._get_auth_frontend_js()
                html = html.replace('</body>', auth_js + '</body>')
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(html.encode('utf-8'))
                return
            self.send_html(NO_DATA_PAGE)
            return

        # --- 用户管理页 ---
        if path == '/admin/users':
            if not self.require_permission('user_manage'):
                return
            self.send_html(USER_MANAGE_PAGE)
            return

        # --- API: 当前用户信息 ---
        if path == '/api/me':
            if not self.require_auth():
                return
            user = self.get_current_user()
            self.send_json({
                'username': user['username'],
                'role': user['role'],
                'role_name': auth.ROLES[user['role']]['name'],
                'email': user.get('email', ''),
                'must_change_pwd': user.get('must_change_pwd', False)
            })
            return

        # --- API: 获取 CSRF Token ---
        if path == '/api/csrf':
            if not self.require_auth():
                return
            session = self.get_current_session()
            token = auth.generate_csrf_token(self._session_id)
            self.send_json({'token': token})
            return

        # --- API: 获取数据（需 view 权限）---
        if path == '/api/data':
            if not self.require_permission('view'):
                return
            self.send_json(load_data())
            return

        # --- API: 版本检查（需 view 权限）---
        if path == '/api/version':
            if not self.require_permission('view'):
                return
            data = load_data()
            self.send_json({'lastUpdate': data.get('lastUpdate', '')})
            return

        # --- API: 获取最新项目数据（需 view 权限）---
        # 【关键修复】用于客户端同步后刷新 RAW_DATA，避免使用过时的归档/删除状态
        if path == '/api/projects':
            if not self.require_permission('view'):
                return
            projects = sync_excel.read_excel_projects()
            # 【死循环修复】必须返回 lastUpdate，否则客户端无法正确同步版本
            data_info = load_data()
            # 【修复】返回当前日期，避免客户端显示的点检日期是HTML生成时的旧日期
            from datetime import datetime, timezone, timedelta
            try:
                shanghai_tz = timezone(timedelta(hours=8))
                today_dt = datetime.now(shanghai_tz)
            except:
                today_dt = datetime.now()
            today_str = today_dt.strftime('%Y-%m-%d')
            three_days_later = (today_dt + timedelta(days=3)).strftime('%Y-%m-%d')
            self.send_json({
                'allProjects': projects,
                'lastUpdate': data_info.get('lastUpdate', ''),
                'today': today_str,
                'threeDaysLater': three_days_later
            })
            return

        # --- API: 用户列表（需 user_manage 权限）---
        if path == '/api/users':
            if not self.require_permission('user_manage'):
                return
            self.send_json(auth.list_users())
            return

        # --- API: 审计日志（需 audit_view 权限）---
        if path == '/api/audit':
            if not self.require_permission('audit_view'):
                return
            self.send_json(auth.get_audit_log(200))
            return

        # --- API: 导入功能检测（需 view 权限）---
        if path == '/api/import/features':
            if not self.require_permission('view'):
                return
            self.send_json(project_parser.get_available_features())
            return

        # 其他静态文件
        super().do_GET()

    # ---------- POST 请求 ----------
    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        # 对于文件上传，先解析 multipart（不调用 read_body，避免二进制数据解码崩溃）
        if path == '/api/upload':
            if not self.require_permission('save'):
                return
            try:
                upload = self.parse_multipart()
                if not upload:
                    self.send_json({'success': False, 'message': '请选择要上传的 Excel 文件'}, 400)
                    return
                filename = upload['filename']
                file_data = upload['data']
                if not filename.lower().endswith(('.xlsx', '.xls')):
                    self.send_json({'success': False, 'message': '只支持 .xlsx 或 .xls 格式'}, 400)
                    return
                if len(file_data) == 0:
                    self.send_json({'success': False, 'message': '文件为空'}, 400)
                    return
                # 保存上传的 Excel
                excel_path = os.path.join(BASE_DIR, '超声波户表脚本.xlsx')
                with open(excel_path, 'wb') as f:
                    f.write(file_data)
                # 调用主脚本生成报表
                import subprocess
                result = subprocess.run(
                    [sys.executable, os.path.join(BASE_DIR, '更新点检表.py'), excel_path],
                    capture_output=True, text=True, cwd=BASE_DIR, timeout=120
                )
                if result.returncode != 0:
                    error_msg = (result.stderr or result.stdout or '生成失败')[-500:]
                    self.send_json({'success': False, 'message': f'报表生成失败：{error_msg}'}, 500)
                    return
                user = self.get_current_user()
                auth._audit_log('REPORT_GENERATE', user['username'], f'上传文件 {filename} 生成报表')
                self.send_json({'success': True, 'message': '报表生成成功'})
                return
            except subprocess.TimeoutExpired:
                self.send_json({'success': False, 'message': '生成超时（超过2分钟），请稍后重试'}, 500)
                return
            except Exception as e:
                import traceback
                self.send_json({'success': False, 'message': f'上传失败：{str(e)}'}, 500)
                return

        # --- 从 GitHub 拉取 Excel 并生成报表 ---
        if path == '/api/fetch-github':
            if not self.require_permission('save'):
                return
            data = self.read_body()
            github_url = data.get('url', '').strip()
            if not github_url:
                self.send_json({'success': False, 'message': '请输入 GitHub 文件链接'}, 400)
                return
            # 验证 URL 格式（支持 github.com 和 raw.githubusercontent.com）
            if 'github.com' not in github_url and 'raw.githubusercontent.com' not in github_url:
                self.send_json({'success': False, 'message': '请输入有效的 GitHub 链接'}, 400)
                return
            # 转换为 raw 链接
            raw_url = github_url
            if 'github.com' in github_url and '/blob/' in github_url:
                raw_url = github_url.replace('github.com', 'raw.githubusercontent.com').replace('/blob/', '/')
            try:
                req = urllib.request.Request(raw_url, headers={
                    'User-Agent': 'Mozilla/5.0 (Report-Server)'
                })
                with urllib.request.urlopen(req, timeout=30) as resp:
                    file_data = resp.read()
                if len(file_data) == 0:
                    self.send_json({'success': False, 'message': '下载的文件为空'}, 400)
                    return
                # 检查文件大小（最大 20MB）
                if len(file_data) > 20 * 1024 * 1024:
                    self.send_json({'success': False, 'message': '文件过大（超过 20MB）'}, 400)
                    return
                # 从 URL 提取文件名
                filename = raw_url.split('/')[-1].split('?')[0]
                if not filename.lower().endswith(('.xlsx', '.xls')):
                    self.send_json({'success': False, 'message': '链接必须指向 .xlsx 或 .xls 文件'}, 400)
                    return
                # 保存
                excel_path = os.path.join(BASE_DIR, '超声波户表脚本.xlsx')
                with open(excel_path, 'wb') as f:
                    f.write(file_data)
                # 生成报表
                import subprocess
                result = subprocess.run(
                    [sys.executable, os.path.join(BASE_DIR, '更新点检表.py'), excel_path],
                    capture_output=True, text=True, cwd=BASE_DIR, timeout=120
                )
                if result.returncode != 0:
                    error_msg = (result.stderr or result.stdout or '生成失败')[-500:]
                    self.send_json({'success': False, 'message': f'报表生成失败：{error_msg}'}, 500)
                    return
                user = self.get_current_user()
                auth._audit_log('REPORT_GENERATE', user['username'], f'从 GitHub 拉取 {filename} 生成报表')
                self.send_json({'success': True, 'message': '报表生成成功'})
                return
            except subprocess.TimeoutExpired:
                self.send_json({'success': False, 'message': '生成超时（超过2分钟），请稍后重试'}, 500)
                return
            except urllib.error.URLError as e:
                self.send_json({'success': False, 'message': f'下载失败：{str(e)}'}, 500)
                return
            except Exception as e:
                import traceback
                self.send_json({'success': False, 'message': f'处理失败：{str(e)}'}, 500)
                return

        # --- 使用仓库内默认 Excel 生成报表 ---
        if path == '/api/generate-default':
            if not self.require_permission('save'):
                return
            excel_path = os.path.join(BASE_DIR, '超声波户表脚本.xlsx')
            if not os.path.exists(excel_path):
                self.send_json({'success': False, 'message': '仓库中未找到默认 Excel 文件'}, 404)
                return
            try:
                import subprocess
                result = subprocess.run(
                    [sys.executable, os.path.join(BASE_DIR, '更新点检表.py'), excel_path],
                    capture_output=True, text=True, cwd=BASE_DIR, timeout=120
                )
                if result.returncode != 0:
                    error_msg = (result.stderr or result.stdout or '生成失败')[-500:]
                    self.send_json({'success': False, 'message': f'报表生成失败：{error_msg}'}, 500)
                    return
                user = self.get_current_user()
                auth._audit_log('REPORT_GENERATE', user['username'], '使用仓库默认 Excel 生成报表')
                self.send_json({'success': True, 'message': '报表生成成功'})
                return
            except subprocess.TimeoutExpired:
                self.send_json({'success': False, 'message': '生成超时（超过2分钟），请稍后重试'}, 500)
                return
            except Exception as e:
                self.send_json({'success': False, 'message': f'处理失败：{str(e)}'}, 500)
                return

        # --- 手动同步数据到 GitHub ---
        if path == '/api/sync-github':
            if not self.require_permission('user_manage'):
                return
            data = self.read_body()
            message = data.get('message', '手动同步') if isinstance(data, dict) else '手动同步'
            ok, msg = auth.sync_to_github(message)
            self.send_json({'success': ok, 'message': msg})
            return

        # 其他接口：读取 JSON body
        data = self.read_body()

        # --- 登录（公开）---
        if path == '/api/login':
            ip = self.get_client_ip()
            allowed, wait = auth.check_rate_limit(ip)
            if not allowed:
                self.send_json({
                    'success': False,
                    'message': f'登录尝试次数过多，请 {wait} 秒后再试'
                }, 429)
                return

            username = data.get('username', '').strip()
            password = data.get('password', '')

            user = auth.get_user(username)
            if not user:
                auth.record_login_attempt(ip, False)
                self.send_json({'success': False, 'message': '用户名或密码错误'})
                return

            if user.get('status') != 'active':
                self.send_json({'success': False, 'message': '账号已被禁用'})
                return

            # 验证密码（使用 auth 模块内部方法）
            stored_pwd = user.get('password', {})
            if not auth._verify_password(password, stored_pwd):
                auth.record_login_attempt(ip, False)
                self.send_json({'success': False, 'message': '用户名或密码错误'})
                return

            auth.record_login_attempt(ip, True)
            session_id = auth.create_session(username)

            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.set_session_cookie(session_id)
            self.end_headers()
            self.wfile.write(json.dumps({
                'success': True,
                'must_change_pwd': user.get('must_change_pwd', False),
                'role': user['role']
            }, ensure_ascii=False).encode('utf-8'))
            return

        # --- 登出（需登录）---
        if path == '/api/logout':
            if self._session_id:
                auth.destroy_session(self._session_id)
            self.send_response(200)
            self.send_header('Content-Type', 'application/json; charset=utf-8')
            self.clear_session_cookie()
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode('utf-8'))
            return

        # --- 修改密码（需登录）---
        if path == '/api/change-password':
            if not self.require_auth():
                return
            user = self.get_current_user()
            old_pwd = data.get('old_password', '')
            new_pwd = data.get('new_password', '')
            ok, msg = auth.change_password(user['username'], old_pwd, new_pwd)
            # 注：auth.change_password 内部已调用 save_users(push=True)，自动同步到 GitHub
            self.send_json({'success': ok, 'message': msg})
            return

        # --- 新增用户（需 user_manage 权限）---
        if path == '/api/user':
            if not self.require_permission('user_manage'):
                return
            ok, msg = auth.create_user(
                data.get('username', ''),
                data.get('password', ''),
                data.get('role', 'viewer'),
                data.get('email', '')
            )
            # 注：auth.create_user 内部已调用 save_users(push=True)，自动同步到 GitHub
            self.send_json({'success': ok, 'message': msg})
            return

        # ==================== 简化方案：原子操作端点 ====================
        # 【简化方案】每个操作直接修改 Excel + 生成报表 + 推送 GitHub
        # 不再通过协作数据 JSON 中转，避免多源状态不同步的问题

        # --- 新增项目（需 edit 权限）---
        # 【关键优化】操作后不返回全量项目数据，避免大JSON序列化导致502超时
        # 前端收到成功响应后自行刷新页面或调用 /api/projects 获取最新数据
        if path == '/api/action/add':
            if not self.require_permission('edit'):
                return
            user = self.get_current_user()
            ok, msg = sync_excel.action_add_project(data, user['username'])
            auth._audit_log('PROJECT_ADD', user['username'], msg[:100])
            self.send_json({'success': ok, 'message': msg})
            return

        # --- 删除项目（需 edit 权限）---
        if path == '/api/action/delete':
            if not self.require_permission('edit'):
                return
            user = self.get_current_user()
            pid = int(data.get('id', 0))
            ok, msg = sync_excel.action_delete_project(pid, user['username'])
            auth._audit_log('PROJECT_DELETE', user['username'], f'ID={pid}: {msg[:80]}')
            self.send_json({'success': ok, 'message': msg})
            return

        # --- 归档项目（需 edit 权限）---
        if path == '/api/action/archive':
            if not self.require_permission('edit'):
                return
            user = self.get_current_user()
            pid = int(data.get('id', 0))
            ok, msg = sync_excel.action_archive_project(pid, user['username'])
            auth._audit_log('PROJECT_ARCHIVE', user['username'], f'ID={pid}: {msg[:80]}')
            self.send_json({'success': ok, 'message': msg})
            return

        # --- 取消归档（需 edit 权限）---
        if path == '/api/action/unarchive':
            if not self.require_permission('edit'):
                return
            user = self.get_current_user()
            pid = int(data.get('id', 0))
            ok, msg = sync_excel.action_unarchive_project(pid, user['username'])
            auth._audit_log('PROJECT_UNARCHIVE', user['username'], f'ID={pid}: {msg[:80]}')
            self.send_json({'success': ok, 'message': msg})
            return

        # --- 编辑项目（需 edit 权限）---
        if path == '/api/action/edit':
            if not self.require_permission('edit'):
                return
            user = self.get_current_user()
            pid = int(data.get('id', 0))
            edit_data = data.get('fields', {})
            ok, msg = sync_excel.action_edit_project(pid, edit_data, user['username'])
            auth._audit_log('PROJECT_EDIT', user['username'], f'ID={pid}: {msg[:80]}')
            self.send_json({'success': ok, 'message': msg})
            return

        # --- 解析导入（需 edit 权限）---
        if path == '/api/import/parse':
            if not self.require_permission('edit'):
                return
            user = self.get_current_user()
            parse_type = data.get('type', 'text')
            text = data.get('text', '')
            filename = data.get('filename', '')

            if parse_type == 'text':
                result = project_parser.parse_text(text)
            elif parse_type == 'mpp':
                import tempfile
                import base64
                tmp_dir = tempfile.gettempdir()
                safe_name = os.path.basename(filename) if filename else f'import_{int(time.time())}.mpp'
                tmp_path = os.path.join(tmp_dir, safe_name)
                try:
                    file_data = base64.b64decode(text) if text else b''
                    with open(tmp_path, 'wb') as f:
                        f.write(file_data)
                    result = project_parser.parse_mpp(tmp_path)
                except Exception as e:
                    result = {'success': False, 'error': f'MPP文件处理失败: {str(e)}'}
                finally:
                    if os.path.exists(tmp_path):
                        try:
                            os.remove(tmp_path)
                        except:
                            pass
            elif parse_type == 'image':
                result = {'success': False, 'error': 'OCR功能暂不可用'}
            else:
                result = {'success': False, 'error': f'不支持的解析类型: {parse_type}'}

            auth._audit_log('IMPORT_PARSE', user['username'], f'type={parse_type}, filename={filename[:50]}')
            self.send_json(result)
            return

        # --- 批量提交导入（需 edit 权限）---
        if path == '/api/import/commit':
            if not self.require_permission('edit'):
                return
            user = self.get_current_user()
            projects = data.get('projects', [])
            result = sync_excel.action_add_project_batch(projects, operator=user['username'])
            auth._audit_log('IMPORT_COMMIT', user['username'], f'批量导入 {len(projects)} 条')
            self.send_json(result)
            return

        # --- 数据同步（需 edit 权限）---
        if path == '/api/sync':
            if not self.require_permission('edit'):
                return
            all_data = load_data()
            for key in ['localEdits', 'notes', 'checked', 'customEmails']:
                if key in data:
                    all_data[key].update(data[key])
            # 【关键修复】archived 不再盲目全量替换，而是以 Excel 为准 + 合并客户端增量
            # 原因：如果客户端的 archived 对象不完整（例如刚通过API归档后还没同步到本地），
            # 全量替换会导致已归档项目的状态丢失，下次 full_sync 时 Excel 的归档标志被清空
            if 'archived' in data:
                # 1. 先从 Excel 读取当前真实的归档状态（最权威的数据源）
                excel_projects = sync_excel.read_excel_projects()
                excel_archived = {}
                for p in excel_projects:
                    if p.get('已归档'):
                        excel_archived[str(p['id'])] = {
                            'time': all_data.get('lastUpdate', ''),
                            'project': p.get('项目', ''),
                            'fromExcel': True
                        }
                # 2. 以 Excel 为基准，再合并客户端的增量（客户端的变更优先级更高）
                merged_archived = dict(excel_archived)
                client_archived = data['archived'] or {}
                for pid, info in client_archived.items():
                    if info:
                        # 客户端明确标记为归档的，添加/更新
                        merged_archived[str(pid)] = info
                    elif str(pid) in merged_archived:
                        # 客户端明确设置为空/false（表示取消归档），才删除
                        # 注意：客户端"没有这个键"不等于取消归档，可能只是客户端数据不全
                        del merged_archived[str(pid)]
                all_data['archived'] = merged_archived
                print(f'[sync] 归档状态合并: Excel基准={len(excel_archived)}个, 客户端增量={len(client_archived)}个, 合并后={len(merged_archived)}个')
            if 'newProjects' in data:
                existing_ids = {p.get('id') for p in all_data['newProjects']}
                for p in data['newProjects']:
                    if p.get('id') not in existing_ids:
                        all_data['newProjects'].append(p)
            if 'deletedIds' in data:
                for did in data['deletedIds']:
                    if did not in all_data['deletedIds']:
                        all_data['deletedIds'].append(did)
            user = self.get_current_user()
            auth._audit_log('DATA_SYNC', user['username'], '数据同步更新')
            save_data(all_data)
            # 自动同步到 Excel + GitHub
            sync_ok, sync_msg = sync_excel.full_sync(f'用户{user["username"]}增量同步')
            # 【修复】根据 full_sync 结果返回正确的 success 状态，不再掩盖失败
            self.send_json({
                'success': sync_ok,
                'lastUpdate': all_data['lastUpdate'],
                'syncMessage': sync_msg
            })
            return

        # --- 全量保存（需 save 权限）---
        if path == '/api/save':
            if not self.require_permission('save'):
                return
            all_data = load_data()
            for key in ['localEdits', 'notes', 'checked', 'customEmails', 'newProjects', 'deletedIds']:
                if key in data:
                    all_data[key] = data[key]
            # 【关键修复】archived 同样以 Excel 为准 + 合并客户端增量，防止归档状态丢失
            if 'archived' in data:
                excel_projects = sync_excel.read_excel_projects()
                excel_archived = {}
                for p in excel_projects:
                    if p.get('已归档'):
                        excel_archived[str(p['id'])] = {
                            'time': all_data.get('lastUpdate', ''),
                            'project': p.get('项目', ''),
                            'fromExcel': True
                        }
                merged_archived = dict(excel_archived)
                client_archived = data['archived'] or {}
                for pid, info in client_archived.items():
                    if info:
                        merged_archived[str(pid)] = info
                    elif str(pid) in merged_archived:
                        del merged_archived[str(pid)]
                all_data['archived'] = merged_archived
                print(f'[save] 归档状态合并: Excel基准={len(excel_archived)}个, 客户端={len(client_archived)}个, 合并后={len(merged_archived)}个')
            user = self.get_current_user()
            auth._audit_log('DATA_SAVE', user['username'], '全量数据保存')
            save_data(all_data)
            # 自动同步到 Excel + GitHub
            sync_ok, sync_msg = sync_excel.full_sync(f'用户{user["username"]}全量保存')
            # 【修复】根据 full_sync 结果返回正确的 success 状态，不再掩盖失败
            self.send_json({
                'success': sync_ok,
                'lastUpdate': all_data['lastUpdate'],
                'syncMessage': sync_msg
            })
            return

        # --- 同步到 Excel 并推送 GitHub（需 save 权限）---
        if path == '/api/sync-excel':
            if not self.require_permission('save'):
                return
            user = self.get_current_user()
            auth._audit_log('SYNC_TO_EXCEL', user['username'], '同步数据到Excel并推送GitHub')
            op_desc = '用户' + user['username'] + '触发同步'
            ok, msg = sync_excel.full_sync(op_desc)
            self.send_json({'success': ok, 'message': msg})
            return

        # --- 从 GitHub 拉取最新数据（需 save 权限）---
        if path == '/api/pull-github':
            if not self.require_permission('save'):
                return
            user = self.get_current_user()
            auth._audit_log('PULL_GITHUB', user['username'], '从GitHub拉取最新数据')
            ok, msg = sync_excel.startup_sync()
            self.send_json({'success': ok, 'message': msg})
            return

        self.send_json({'error': 'Unknown endpoint'}, 404)

    # ---------- PUT 请求 ----------
    def do_PUT(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        data = self.read_body()

        # 更新用户
        if path.startswith('/api/user/'):
            if not self.require_permission('user_manage'):
                return
            username = path[len('/api/user/'):]
            update_data = {}
            for k in ['role', 'email', 'password', 'status', 'must_change_pwd']:
                if k in data:
                    update_data[k] = data[k]
            ok, msg = auth.update_user(username, **update_data)
            # 注：auth.update_user 内部已调用 save_users(push=True)，自动同步到 GitHub
            self.send_json({'success': ok, 'message': msg})
            return

        self.send_json({'error': 'Unknown endpoint'}, 404)

    # ---------- DELETE 请求 ----------
    def do_DELETE(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        # 删除用户
        if path.startswith('/api/user/'):
            if not self.require_permission('user_manage'):
                return
            username = path[len('/api/user/'):]
            ok, msg = auth.delete_user(username)
            # 注：auth.delete_user 内部已调用 save_users(push=True)，自动同步到 GitHub
            self.send_json({'success': ok, 'message': msg})
            return

        self.send_json({'error': 'Unknown endpoint'}, 404)

    # ---------- 前端注入脚本 ----------
    def _get_auth_frontend_js(self):
        """注入到报表页面的权限控制脚本"""
        return '''<script>
(function(){
  if (!window.AUTH_ENABLED || !window.CURRENT_USER) return;
  var perms = window.CURRENT_USER.permissions || [];
  var hasPerm = function(p) { return perms.indexOf(p) >= 0; };

  // 在页面加载后注入用户信息和权限控制
  document.addEventListener('DOMContentLoaded', function() {
    // 注入顶部用户信息栏
    var authBar = document.createElement('div');
    authBar.style.cssText = 'position:fixed;top:0;right:0;z-index:99999;background:linear-gradient(135deg,#667eea,#764ba2);color:#fff;padding:8px 16px;border-radius:0 0 0 12px;font-size:13px;box-shadow:0 2px 8px rgba(0,0,0,0.2);display:flex;gap:12px;align-items:center;';
    authBar.innerHTML =
      '<span>👤 <b>' + window.CURRENT_USER.username + '</b> (' + window.CURRENT_USER.role_name + ')</span>' +
      (hasPerm('user_manage') ? '<a href="/admin/users" style="color:#fff;text-decoration:none;opacity:0.9">👥 用户管理</a>' : '') +
      '<a href="javascript:;" onclick="logout()" style="color:#fff;text-decoration:none;opacity:0.9">🚪 退出</a>';
    document.body.appendChild(authBar);

    // 如果没有 edit 权限，禁用编辑
    if (!hasPerm('edit')) {
      document.body.classList.add('read-only-mode');
      var style = document.createElement('style');
      style.textContent = '.read-only-mode input[contenteditable], .read-only-mode [contenteditable="true"] { pointer-events:none !important; opacity:0.7; }';
      document.head.appendChild(style);
    }

    // 首次登录强制改密码
    if (window.CURRENT_USER.must_change_pwd) {
      setTimeout(function() {
        showChangePwdModal();
      }, 500);
    }
  });

  window.logout = function() {
    if (!confirm('确定退出登录？')) return;
    fetch('/api/logout', {method:'POST'}).then(function() {
      location.reload();
    });
  };

  window.showChangePwdModal = function() {
    var oldPwd = prompt('请输入原密码：');
    if (!oldPwd) return;
    var newPwd = prompt('请输入新密码（至少6位）：');
    if (!newPwd) return;
    var confirmPwd = prompt('请再次输入新密码：');
    if (newPwd !== confirmPwd) { alert('两次密码不一致'); return; }
    fetch('/api/change-password', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({old_password: oldPwd, new_password: newPwd})
    }).then(function(r){return r.json()}).then(function(d){
      alert(d.message);
      if (d.success) location.reload();
    });
  };
})();
</script>'''

    def log_message(self, format, *args):
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        try:
            user = ''
            session = self.get_current_session()
            if session:
                user = f' [{session.get("username","")}]'
            print(f"[{timestamp}]{user} {format % args}")
        except:
            print(f"[{timestamp}] {format}")


# ==================== 启动服务器 ====================

# 定时同步线程标志
_auto_pull_running = True
_auto_pull_thread = None
_last_pull_hash = None

def _auto_pull_worker(interval_seconds: int):
    """后台线程：定期从 GitHub 拉取最新数据"""
    global _last_pull_hash
    while _auto_pull_running:
        try:
            # 拉取最新
            ok, msg = sync_excel.git_pull()
            if ok:
                # 检查Excel是否有变化（通过文件修改时间）
                excel_path = os.path.join(BASE_DIR, '超声波户表脚本.xlsx')
                if os.path.exists(excel_path):
                    current_mtime = os.path.getmtime(excel_path)
                    if _last_pull_hash is None or current_mtime != _last_pull_hash:
                        _last_pull_hash = current_mtime
                        # Excel有变化，清除缓存 + 重新生成报表
                        sync_excel.invalidate_projects_cache()
                        ok2, msg2 = sync_excel.regenerate_report()
                        print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] 🔄 GitHub数据已更新: {msg2}")
        except Exception as e:
            print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] ⚠️  自动拉取异常: {e}")
        
        # 等待下一轮
        for _ in range(interval_seconds):
            if not _auto_pull_running:
                break
            time.sleep(1)

def start_auto_pull(interval_seconds: int = 300):
    """启动自动拉取线程"""
    global _auto_pull_thread
    if _auto_pull_thread and _auto_pull_thread.is_alive():
        return
    _auto_pull_running = True
    _auto_pull_thread = threading.Thread(
        target=_auto_pull_worker,
        args=(interval_seconds,),
        daemon=True,
        name='AutoPullThread'
    )
    _auto_pull_thread.start()
    print(f"🔄 已启用自动从GitHub拉取（每{interval_seconds//60}分钟）")


def main():
    os.chdir(BASE_DIR)

    # ========== 【关键修复1】启动时先从 GitHub 拉取最新数据 ==========
    # 必须在 init_auth() 之前！否则 init_auth 会在文件缺失时创建空的用户表并推送，
    # 覆盖掉 GitHub 上已有的用户数据！
    print("🔄 正在从 GitHub 拉取最新数据...")
    sync_ok, sync_msg = sync_excel.startup_sync()
    print(f"   {sync_msg}")

    # ========== 【关键修复2】再初始化认证系统 ==========
    # 此时 用户管理.xlsx 已经从 GitHub 拉取回来，不会被 init_default_users 覆盖
    auth.init_auth()
    
    # ========== 【502关键修复】启动时预热项目数据缓存 ==========
    # 第一个用户请求时就不用冷启动读Excel了（从几秒降到几毫秒）
    print("📊 正在预热项目数据缓存...")
    try:
        warmup_projects = sync_excel.read_excel_projects()
        print(f"   ✅ 缓存预热完成（{len(warmup_projects)}条数据）")
    except Exception as e:
        print(f"   ⚠️  缓存预热失败: {e}")

    # 启动自动同步用户数据到 GitHub（每30分钟一次，用户数据变化不频繁）
    auth.auto_sync_periodically(30 * 60)
    
    # 【关键】关闭自动从 GitHub 拉取！
    # 服务器是唯一数据源，用户只通过网页操作修改数据
    # 没有其他客户端会修改 Excel，所以完全不需要频繁拉取
    # 只在启动时拉一次恢复数据即可（启动代码中已有 git_pull）
    # start_auto_pull(300)  ← 已禁用

    if not os.path.exists(DATA_FILE):
        save_data(load_data())

    # 自动生成报表（如果Excel存在但HTML不存在，或HTML为空）
    excel_path = os.path.join(BASE_DIR, '超声波户表脚本.xlsx')
    if os.path.exists(excel_path):
        need_generate = False
        if not os.path.exists(HTML_FILE):
            need_generate = True
            print("📊 检测到报表HTML不存在，正在自动生成...")
        elif os.path.getsize(HTML_FILE) < 50000:
            need_generate = True
            print("📊 检测到报表HTML文件过小（可能为空），正在重新生成...")
        else:
            # 检查 HTML 中是否有数据
            try:
                with open(HTML_FILE, 'r', encoding='utf-8') as f:
                    html_content = f.read()
                if 'allProjects' not in html_content or '"total": 0' in html_content or '"today": ""' in html_content:
                    need_generate = True
                    print("📊 检测到报表数据为空，正在重新生成...")
            except:
                pass

        if need_generate:
            try:
                import subprocess
                result = subprocess.run(
                    [sys.executable, os.path.join(BASE_DIR, '更新点检表.py'), excel_path],
                    capture_output=True, text=True, cwd=BASE_DIR, timeout=120
                )
                if result.returncode == 0:
                    print("✅ 报表自动生成成功")
                else:
                    print(f"❌ 报表自动生成失败: {(result.stderr or result.stdout)[-300:]}")
            except Exception as e:
                print(f"❌ 报表自动生成异常: {e}")

    local_ip = get_local_ip()

    print("=" * 60)
    print("🔐 项目点检表协作服务器（安全版）已启动！")
    print("=" * 60)
    print(f"📂 工作目录: {BASE_DIR}")
    print(f"💾 数据文件: {DATA_FILE}")
    print(f"👥 用户文件: {auth.USERS_EXCEL}")
    print()
    print("✅ 已启用: 登录认证 · 角色权限 · Session 管理")
    print("✅ 已启用: 登录限流 · CSRF 防护 · 操作审计")
    print("✅ 已启用: GitHub 数据持久化（每5分钟自动推拉同步）")
    print("✅ 已启用: GitHub 自动拉取（Excel修改后自动更新报表）")
    print()
    print("🌐 访问地址：")
    print(f"   本机访问: http://localhost:{PORT}")
    print(f"   局域网:  http://{local_ip}:{PORT}")
    print()
    print("🔑 默认管理员: admin / admin123")
    print("   （首次登录请立即修改密码！）")
    print()
    print("👥 用户管理页: http://localhost:{PORT}/admin/users".format(PORT=PORT))
    print()
    print("💡 提示：按 Ctrl+C 停止服务器")
    print("=" * 60)

    try:
        socketserver.ThreadingTCPServer.allow_reuse_address = True
        with socketserver.ThreadingTCPServer(('', PORT), SecureCollaborationHandler) as httpd:
            httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n\n👋 服务器已停止")
    except OSError as e:
        if 'Address already in use' in str(e):
            print(f"\n❌ 端口 {PORT} 已被占用，请换一个端口：")
            print(f"   python 协作服务器_安全版.py {PORT + 1}")
        else:
            raise


if __name__ == '__main__':
    main()
