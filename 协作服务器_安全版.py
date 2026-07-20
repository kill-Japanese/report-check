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
import secrets
from datetime import datetime

# 导入认证模块
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import auth

# ==================== 配置 ====================
PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8080
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_FILE = os.path.join(BASE_DIR, '协作数据.json')
HTML_FILE = os.path.join(BASE_DIR, '项目延期点检表.html')

# 安全头
SECURITY_HEADERS = {
    'X-Content-Type-Options': 'nosniff',
    'X-Frame-Options': 'SAMEORIGIN',
    'X-XSS-Protection': '1; mode=block',
    'Referrer-Policy': 'strict-origin-when-cross-origin',
}

# ==================== 数据管理 ====================
def load_data():
    if os.path.exists(DATA_FILE):
        try:
            with open(DATA_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {
        'localEdits': {}, 'notes': {}, 'checked': {},
        'archived': {}, 'customEmails': {}, 'newProjects': [],
        'lastUpdate': datetime.now().isoformat()
    }

def save_data(data):
    data['lastUpdate'] = datetime.now().isoformat()
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
      body: JSON.stringify({ username, password })
    });
    const data = await res.json();
    if (data.success) {
      if (data.must_change_pwd) {
        alert('首次登录请修改密码！');
      }
      window.location.reload();
    } else {
      errorEl.textContent = data.message || '登录失败';
      errorEl.classList.add('show');
    }
  } catch (err) {
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
        """设置 Session Cookie"""
        self.send_header(
            'Set-Cookie',
            f'session_id={session_id}; Path=/; HttpOnly; SameSite=Lax; Max-Age={auth.SESSION_TIMEOUT}'
        )

    def clear_session_cookie(self):
        """清除 Session Cookie"""
        self.send_header('Set-Cookie', 'session_id=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0')

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
            self.send_html('<h1>报表文件不存在</h1>', 404)
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

        # 其他静态文件
        super().do_GET()

    # ---------- POST 请求 ----------
    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
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
            self.send_json({'success': ok, 'message': msg})
            return

        # --- 数据同步（需 edit 权限）---
        if path == '/api/sync':
            if not self.require_permission('edit'):
                return
            all_data = load_data()
            for key in ['localEdits', 'notes', 'checked', 'archived', 'customEmails']:
                if key in data:
                    all_data[key].update(data[key])
            if 'newProjects' in data:
                existing_ids = {p.get('id') for p in all_data['newProjects']}
                for p in data['newProjects']:
                    if p.get('id') not in existing_ids:
                        all_data['newProjects'].append(p)
            user = self.get_current_user()
            auth._audit_log('DATA_SYNC', user['username'], '数据同步更新')
            save_data(all_data)
            self.send_json({'success': True, 'lastUpdate': all_data['lastUpdate']})
            return

        # --- 全量保存（需 save 权限）---
        if path == '/api/save':
            if not self.require_permission('save'):
                return
            all_data = load_data()
            for key in ['localEdits', 'notes', 'checked', 'archived', 'customEmails', 'newProjects']:
                if key in data:
                    all_data[key] = data[key]
            user = self.get_current_user()
            auth._audit_log('DATA_SAVE', user['username'], '全量数据保存')
            save_data(all_data)
            self.send_json({'success': True, 'lastUpdate': all_data['lastUpdate']})
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
def main():
    os.chdir(BASE_DIR)
    auth.init_auth()

    if not os.path.exists(DATA_FILE):
        save_data(load_data())

    local_ip = get_local_ip()

    print("=" * 60)
    print("🔐 项目点检表协作服务器（安全版）已启动！")
    print("=" * 60)
    print(f"📂 工作目录: {BASE_DIR}")
    print(f"💾 数据文件: {DATA_FILE}")
    print(f"👥 用户文件: {auth.USERS_FILE}")
    print()
    print("✅ 已启用: 登录认证 · 角色权限 · Session 管理")
    print("✅ 已启用: 登录限流 · CSRF 防护 · 操作审计")
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
