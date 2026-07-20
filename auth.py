# -*- coding: utf-8 -*-
"""
项目点检表 - 权限管理与安全模块

功能：
- 用户管理（增删改查）
- 密码加密存储（PBKDF2-HMAC-SHA256）
- 角色权限（admin / editor / viewer）
- Session 管理（带过期时间）
- 登录限流（防止暴力破解）
- CSRF Token 防护
- 操作日志审计
"""

import hashlib
import hmac
import json
import os
import uuid
import secrets
import time
import threading
from datetime import datetime, timedelta
from functools import wraps

# ==================== 配置 ====================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 数据文件
USERS_FILE = os.path.join(BASE_DIR, 'users.json')
SESSIONS_FILE = os.path.join(BASE_DIR, 'sessions.json')
RATE_LIMIT_FILE = os.path.join(BASE_DIR, 'rate_limit.json')
AUDIT_LOG_FILE = os.path.join(BASE_DIR, 'audit.log')

# 安全配置
PASSWORD_HASH_ITERATIONS = 200_000  # PBKDF2 迭代次数
SESSION_TIMEOUT = 8 * 60 * 60        # Session 过期时间（8小时）
MAX_LOGIN_ATTEMPTS = 5                # 最大登录失败次数
LOGIN_LOCKOUT_TIME = 15 * 60          # 锁定时间（15分钟）
CSRF_TOKEN_TTL = 3600                  # CSRF Token 有效期（1小时）

# 角色定义
ROLES = {
    'admin': {
        'name': '管理员',
        'permissions': [
            'view',           # 查看报表
            'edit',           # 编辑数据
            'save',           # 全量保存
            'user_manage',    # 用户管理
            'role_manage',    # 角色管理
            'audit_view',     # 查看审计日志
            'system_config',  # 系统配置
        ]
    },
    'editor': {
        'name': '编辑者',
        'permissions': [
            'view',
            'edit',
            'save',
        ]
    },
    'viewer': {
        'name': '只读用户',
        'permissions': [
            'view',
        ]
    }
}

# 线程锁
_lock = threading.Lock()


# ==================== 工具函数 ====================

def _safe_read_json(filepath, default):
    """安全读取 JSON 文件"""
    with _lock:
        if not os.path.exists(filepath):
            return default
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return default


def _safe_write_json(filepath, data):
    """安全写入 JSON 文件"""
    with _lock:
        tmp = filepath + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, filepath)


def _hash_password(password: str, salt: bytes = None) -> dict:
    """使用 PBKDF2-HMAC-SHA256 加密密码"""
    if salt is None:
        salt = secrets.token_bytes(32)
    key = hashlib.pbkdf2_hmac(
        'sha256',
        password.encode('utf-8'),
        salt,
        PASSWORD_HASH_ITERATIONS
    )
    return {
        'hash': key.hex(),
        'salt': salt.hex(),
        'iterations': PASSWORD_HASH_ITERATIONS,
        'algorithm': 'PBKDF2-HMAC-SHA256'
    }


def _verify_password(password: str, stored: dict) -> bool:
    """验证密码"""
    try:
        result = _hash_password(password, bytes.fromhex(stored['salt']))
        return hmac.compare_digest(result['hash'], stored['hash'])
    except Exception:
        return False


def _audit_log(action: str, username: str, detail: str = ''):
    """记录审计日志"""
    try:
        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        entry = f"[{timestamp}] {action} | 用户: {username} | {detail}\n"
        with open(AUDIT_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(entry)
    except Exception:
        pass


# ==================== 用户管理 ====================

def load_users() -> dict:
    """加载所有用户"""
    return _safe_read_json(USERS_FILE, {})


def save_users(users: dict):
    """保存用户"""
    _safe_write_json(USERS_FILE, users)


def init_default_users():
    """初始化默认用户（首次运行）"""
    users = load_users()
    if not users:
        # 创建默认管理员
        default_pwd = 'admin123'
        users['admin'] = {
            'username': 'admin',
            'password': _hash_password(default_pwd),
            'role': 'admin',
            'email': '',
            'created_at': datetime.now().isoformat(),
            'last_login': None,
            'must_change_pwd': True,  # 首次登录必须改密码
            'status': 'active'
        }
        save_users(users)
        _audit_log('USER_CREATE', 'system', '创建默认管理员 admin / admin123')
        print(f"\n⚠️  已创建默认管理员账号: admin / {default_pwd}")
        print("    请登录后立即修改密码！\n")
    return users


def get_user(username: str) -> dict:
    """获取用户信息"""
    users = load_users()
    return users.get(username)


def create_user(username: str, password: str, role: str = 'viewer',
                email: str = '') -> tuple[bool, str]:
    """创建用户"""
    if role not in ROLES:
        return False, f'无效的角色: {role}'
    if len(username) < 3 or len(username) > 32:
        return False, '用户名长度需在 3-32 之间'
    if len(password) < 6:
        return False, '密码长度至少 6 位'

    users = load_users()
    if username in users:
        return False, '用户名已存在'

    users[username] = {
        'username': username,
        'password': _hash_password(password),
        'role': role,
        'email': email,
        'created_at': datetime.now().isoformat(),
        'last_login': None,
        'must_change_pwd': False,
        'status': 'active'
    }
    save_users(users)
    _audit_log('USER_CREATE', 'admin', f'创建用户: {username}, 角色: {role}')
    return True, '用户创建成功'


def update_user(username: str, **kwargs) -> tuple[bool, str]:
    """更新用户信息"""
    users = load_users()
    if username not in users:
        return False, '用户不存在'

    if 'role' in kwargs and kwargs['role'] not in ROLES:
        return False, f'无效的角色: {kwargs["role"]}'
    if 'password' in kwargs:
        if len(kwargs['password']) < 6:
            return False, '密码长度至少 6 位'
        users[username]['password'] = _hash_password(kwargs['password'])
        del kwargs['password']

    for k, v in kwargs.items():
        if k in ('role', 'email', 'status', 'must_change_pwd'):
            users[username][k] = v

    save_users(users)
    _audit_log('USER_UPDATE', 'admin', f'更新用户: {username}, 字段: {list(kwargs.keys())}')
    return True, '用户更新成功'


def delete_user(username: str) -> tuple[bool, str]:
    """删除用户"""
    if username == 'admin':
        return False, '不能删除默认管理员'

    users = load_users()
    if username not in users:
        return False, '用户不存在'

    del users[username]
    save_users(users)
    _audit_log('USER_DELETE', 'admin', f'删除用户: {username}')
    return True, '用户删除成功'


def change_password(username: str, old_password: str, new_password: str) -> tuple[bool, str]:
    """用户修改自己的密码"""
    users = load_users()
    if username not in users:
        return False, '用户不存在'

    if not _verify_password(old_password, users[username]['password']):
        return False, '原密码错误'
    if len(new_password) < 6:
        return False, '新密码长度至少 6 位'
    if old_password == new_password:
        return False, '新密码不能与原密码相同'

    users[username]['password'] = _hash_password(new_password)
    users[username]['must_change_pwd'] = False
    save_users(users)
    _audit_log('PASSWORD_CHANGE', username, '修改密码成功')
    return True, '密码修改成功'


def list_users() -> list:
    """列出所有用户（去除密码）"""
    users = load_users()
    result = []
    for u in users.values():
        safe = {k: v for k, v in u.items() if k != 'password'}
        safe['role_name'] = ROLES.get(u['role'], {}).get('name', u['role'])
        result.append(safe)
    return result


# ==================== 权限检查 ====================

def has_permission(username: str, permission: str) -> bool:
    """检查用户是否有指定权限"""
    user = get_user(username)
    if not user or user.get('status') != 'active':
        return False
    role = user.get('role', 'viewer')
    return permission in ROLES.get(role, {}).get('permissions', [])


def require_permission(permission: str):
    """装饰器：要求权限"""
    def decorator(func):
        @wraps(func)
        def wrapper(self, *args, **kwargs):
            session = self.get_current_session()
            if not session:
                self.send_json({'error': '未登录'}, 401)
                return
            username = session.get('username')
            if not has_permission(username, permission):
                self.send_json({'error': '权限不足'}, 403)
                return
            return func(self, *args, **kwargs)
        return wrapper
    return decorator


# ==================== Session 管理 ====================

def load_sessions() -> dict:
    return _safe_read_json(SESSIONS_FILE, {})


def save_sessions(sessions: dict):
    _safe_write_json(SESSIONS_FILE, sessions)


def cleanup_sessions():
    """清理过期 Session"""
    sessions = load_sessions()
    now = time.time()
    expired = [sid for sid, s in sessions.items()
               if s.get('expires_at', 0) < now]
    for sid in expired:
        del sessions[sid]
    if expired:
        save_sessions(sessions)
        _audit_log('SESSION_CLEANUP', 'system', f'清理了 {len(expired)} 个过期 Session')
    return len(expired)


def create_session(username: str) -> str:
    """创建 Session，返回 Session ID"""
    cleanup_sessions()
    sessions = load_sessions()

    # 先清除该用户的旧 Session（单点登录）
    for sid, s in list(sessions.items()):
        if s.get('username') == username:
            del sessions[sid]

    session_id = secrets.token_urlsafe(48)
    now = time.time()
    sessions[session_id] = {
        'username': username,
        'created_at': now,
        'expires_at': now + SESSION_TIMEOUT,
        'last_activity': now,
    }
    save_sessions(sessions)

    # 更新用户最后登录时间
    users = load_users()
    if username in users:
        users[username]['last_login'] = datetime.now().isoformat()
        save_users(users)

    _audit_log('LOGIN', username, '登录成功')
    return session_id


def get_session(session_id: str) -> dict:
    """获取 Session（自动续期）"""
    if not session_id:
        return None
    cleanup_sessions()
    sessions = load_sessions()
    session = sessions.get(session_id)
    if not session:
        return None
    if session.get('expires_at', 0) < time.time():
        del sessions[session_id]
        save_sessions(sessions)
        return None

    # 自动续期（每次访问重置过期时间）
    sessions[session_id]['last_activity'] = time.time()
    sessions[session_id]['expires_at'] = time.time() + SESSION_TIMEOUT
    save_sessions(sessions)

    return session


def destroy_session(session_id: str):
    """销毁 Session（登出）"""
    sessions = load_sessions()
    if session_id in sessions:
        username = sessions[session_id].get('username', '')
        del sessions[session_id]
        save_sessions(sessions)
        _audit_log('LOGOUT', username, '登出成功')


# ==================== 登录限流 ====================

def load_rate_limit() -> dict:
    return _safe_read_json(RATE_LIMIT_FILE, {})


def save_rate_limit(data: dict):
    _safe_write_json(RATE_LIMIT_FILE, data)


def check_rate_limit(ip: str) -> tuple[bool, int]:
    """检查是否被限流，返回 (是否允许, 剩余秒数)"""
    data = load_rate_limit()
    entry = data.get(ip, {})
    now = time.time()

    # 检查是否被锁定
    if entry.get('locked_until', 0) > now:
        return False, int(entry['locked_until'] - now)

    return True, 0


def record_login_attempt(ip: str, success: bool):
    """记录登录尝试"""
    data = load_rate_limit()
    now = time.time()
    entry = data.get(ip, {'attempts': 0, 'first_attempt': 0})

    if success:
        # 登录成功，重置计数
        data[ip] = {'attempts': 0, 'first_attempt': 0}
    else:
        if now - entry.get('first_attempt', 0) > LOGIN_LOCKOUT_TIME:
            # 超过锁定窗口，重置
            entry = {'attempts': 1, 'first_attempt': now}
        else:
            entry['attempts'] += 1
            if entry['attempts'] >= MAX_LOGIN_ATTEMPTS:
                entry['locked_until'] = now + LOGIN_LOCKOUT_TIME
                _audit_log('RATE_LIMIT', 'system',
                           f'IP {ip} 因登录失败次数过多被锁定 {LOGIN_LOCKOUT_TIME//60} 分钟')
        data[ip] = entry

    save_rate_limit(data)


# ==================== CSRF 防护 ====================

_csrf_tokens = {}  # session_id -> {token, expires_at}


def generate_csrf_token(session_id: str) -> str:
    """生成 CSRF Token"""
    token = secrets.token_urlsafe(32)
    _csrf_tokens[session_id] = {
        'token': token,
        'expires_at': time.time() + CSRF_TOKEN_TTL
    }
    return token


def verify_csrf_token(session_id: str, token: str) -> bool:
    """验证 CSRF Token"""
    entry = _csrf_tokens.get(session_id)
    if not entry:
        return False
    if entry['expires_at'] < time.time():
        del _csrf_tokens[session_id]
        return False
    return hmac.compare_digest(entry['token'], token)


# ==================== 审计日志 ====================

def get_audit_log(limit: int = 100) -> list:
    """获取审计日志"""
    if not os.path.exists(AUDIT_LOG_FILE):
        return []
    try:
        with open(AUDIT_LOG_FILE, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        return [l.strip() for l in lines[-limit:]]
    except Exception:
        return []


# ==================== 初始化 ====================

def init_auth():
    """初始化认证系统"""
    init_default_users()
    cleanup_sessions()
    # 定期清理 Session（每小时一次）
    def _periodic_cleanup():
        while True:
            time.sleep(3600)
            try:
                cleanup_sessions()
            except Exception:
                pass

    t = threading.Thread(target=_periodic_cleanup, daemon=True)
    t.start()


if __name__ == '__main__':
    # 测试
    init_auth()
    print("用户列表:", list_users())
    print("角色:", list(ROLES.keys()))
