# -*- coding: utf-8 -*-
"""
GitHub API 同步模块 - 当 git 命令不可用时的兜底方案

【零依赖设计】使用 Python 内置 urllib 实现 HTTP 请求，
不依赖 requests 或任何第三方库，确保在任何 Python 环境中都能工作。

功能：
1. 从 GitHub 拉取文件（替代 git pull/checkout）
2. 推送文件变更到 GitHub（替代 git add/commit/push）
3. 自动检测 git 命令是否可用，不可用时自动切换到 API 模式
"""

import os
import sys
import json
import base64
import subprocess
from datetime import datetime

# 【关键】优先用内置 urllib，不依赖 requests
try:
    from urllib.request import Request, urlopen
    from urllib.error import HTTPError, URLError
    from urllib.parse import quote as url_quote
    HAS_URLLIB = True
except ImportError:
    HAS_URLLIB = False
    def url_quote(s, safe=''):
        return s

# requests 作为可选（如果有的话），但不依赖
try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
GITHUB_API_BASE = 'https://api.github.com'

# ==================== HTTP 请求封装（零依赖） ====================

def _http_request(url: str, method: str = 'GET', headers: dict = None,
                  data: bytes = None, timeout: int = 30) -> tuple[int, bytes, dict]:
    """
    通用 HTTP 请求（优先用 urllib，requests 兜底）
    返回: (status_code, response_body_bytes, headers_dict)
    """
    if headers is None:
        headers = {}
    
    # 优先用 urllib（零依赖）
    if HAS_URLLIB:
        try:
            req = Request(url, data=data, headers=headers, method=method)
            resp = urlopen(req, timeout=timeout)
            resp_body = resp.read()
            resp_headers = dict(resp.headers.items()) if hasattr(resp, 'headers') else {}
            return resp.getcode(), resp_body, resp_headers
        except HTTPError as e:
            resp_body = e.read() if hasattr(e, 'read') else b''
            resp_headers = dict(e.headers.items()) if hasattr(e, 'headers') else {}
            return e.code, resp_body, resp_headers
        except URLError as e:
            return 0, str(e.reason).encode('utf-8'), {}
        except Exception as e:
            return 0, str(e).encode('utf-8'), {}
    
    # urllib 不可用时才用 requests（极端情况）
    elif HAS_REQUESTS:
        try:
            resp = requests.request(method, url, headers=headers, data=data, timeout=timeout)
            return resp.status_code, resp.content, dict(resp.headers)
        except Exception as e:
            return 0, str(e).encode('utf-8'), {}
    
    return 0, b'No HTTP library available', {}


# ==================== 环境检测 ====================

def has_git_command() -> bool:
    """检测 git 命令是否可用"""
    try:
        result = subprocess.run(
            ['git', '--version'],
            capture_output=True, timeout=5
        )
        return result.returncode == 0
    except (FileNotFoundError, OSError):
        return False


def has_git_repo() -> bool:
    """检测 .git 目录是否存在"""
    return os.path.exists(os.path.join(BASE_DIR, '.git'))


def should_use_github_api() -> bool:
    """判断是否应该使用 GitHub API 模式（git 不可用时）"""
    return not has_git_command() or not has_git_repo()


# ==================== GitHub 配置解析 ====================

def _get_github_config() -> tuple[str, str, str]:
    """
    从环境变量或 git 配置中获取 GitHub 认证信息
    返回: (token, owner, repo)
    """
    token = os.environ.get('GITHUB_TOKEN', '')
    owner = os.environ.get('GITHUB_OWNER', '')
    repo = os.environ.get('GITHUB_REPO', '')
    
    # 尝试从 git remote URL 解析
    if not token or not owner or not repo:
        try:
            result = subprocess.run(
                ['git', 'remote', '-v'],
                capture_output=True, text=True, cwd=BASE_DIR, timeout=5
            )
            for line in result.stdout.split('\n'):
                if 'origin' in line and 'github.com' in line:
                    url = line.split()[1]
                    if 'github.com' in url:
                        # 解析各种格式:
                        # https://token@github.com/owner/repo.git
                        # https://github.com/owner/repo.git
                        # git@github.com:owner/repo.git
                        url_clean = url.rstrip('.git').rstrip('/')
                        
                        # 提取 token
                        if '@github.com' in url_clean:
                            token_part = url_clean.split('@github.com')[0]
                            if 'https://' in token_part or 'http://' in token_part:
                                auth_part = token_part.split('://')[1]
                                if ':' in auth_part:
                                    token = auth_part.split(':')[1]
                                else:
                                    token = auth_part
                            path_part = url_clean.split('@github.com/')[1]
                        elif 'github.com/' in url_clean:
                            path_part = url_clean.split('github.com/')[1]
                        elif 'github.com:' in url_clean:
                            path_part = url_clean.split('github.com:')[1]
                        else:
                            continue
                        
                        parts = path_part.rstrip('/').split('/')
                        if len(parts) >= 2:
                            owner = parts[0]
                            repo = '/'.join(parts[1:])
                            break
        except:
            pass
    
    # 如果还是没有，用默认值（本项目）
    if not owner:
        owner = 'kill-Japanese'
    if not repo:
        repo = 'report-check'
    
    return token, owner, repo


# ==================== GitHub API 核心操作 ====================

def github_api_get_file(path: str, branch: str = 'main') -> tuple[bool, bytes, str]:
    """
    从 GitHub 获取文件内容
    返回: (成功, 文件内容(bytes), 文件SHA)
    """
    token, owner, repo = _get_github_config()
    if not token:
        return False, b'', ''
    
    # 【关键修复】URL 路径中的中文字符需要百分号编码
    encoded_path = url_quote(path, safe='/')
    url = f'{GITHUB_API_BASE}/repos/{owner}/{repo}/contents/{encoded_path}?ref={branch}'
    headers = {
        'Authorization': f'token {token}',
        'Accept': 'application/vnd.github.v3+json',
        'User-Agent': 'report-check-server'
    }
    
    status_code, resp_body, _ = _http_request(url, 'GET', headers=headers)
    
    if status_code == 200:
        try:
            data = json.loads(resp_body.decode('utf-8'))
            content = base64.b64decode(data['content'])
            sha = data.get('sha', '')
            return True, content, sha
        except Exception as e:
            print(f'[GitHub API] 解析文件内容失败: {e}')
            return False, b'', ''
    elif status_code == 404:
        return True, b'', ''  # 文件不存在不算错误
    else:
        err_msg = resp_body.decode('utf-8', errors='replace')[:300]
        print(f'[GitHub API] 获取文件失败: HTTP {status_code} {err_msg}')
        return False, b'', ''


def github_api_push_file(path: str, content: bytes, message: str,
                         sha: str = '', branch: str = 'main') -> tuple[bool, str]:
    """
    推送文件到 GitHub（创建或更新）
    返回: (成功, 消息)
    """
    token, owner, repo = _get_github_config()
    if not token:
        return False, '缺少 GitHub Token'
    
    # 【关键修复】URL 路径中的中文字符需要百分号编码
    encoded_path = url_quote(path, safe='/')
    url = f'{GITHUB_API_BASE}/repos/{owner}/{repo}/contents/{encoded_path}'
    headers = {
        'Authorization': f'token {token}',
        'Accept': 'application/vnd.github.v3+json',
        'Content-Type': 'application/json',
        'User-Agent': 'report-check-server'
    }
    
    # 如果没有提供 SHA，尝试获取
    if not sha:
        ok, _, existing_sha = github_api_get_file(path, branch)
        if ok:
            sha = existing_sha
    
    payload = {
        'message': f'[API同步] {message} - {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}',
        'content': base64.b64encode(content).decode('utf-8'),
        'branch': branch
    }
    if sha:
        payload['sha'] = sha  # 更新文件需要 SHA
    
    data_bytes = json.dumps(payload, ensure_ascii=False).encode('utf-8')
    status_code, resp_body, _ = _http_request(url, 'PUT', headers=headers, data=data_bytes)
    
    if status_code in (200, 201):
        return True, '已同步到 GitHub'
    else:
        err_msg = resp_body.decode('utf-8', errors='replace')[:300]
        print(f'[GitHub API] 推送文件失败: HTTP {status_code} {err_msg}')
        return False, f'API推送失败: HTTP {status_code}'


# ==================== 高级操作（替代 git pull/push） ====================

# 需要同步的关键文件
CRITICAL_FILES = [
    '用户管理.xlsx',
    '超声波户表脚本.xlsx',
]


def github_api_pull() -> tuple[bool, str]:
    """
    从 GitHub 拉取所有关键文件（替代 git pull）
    返回: (成功, 消息)
    """
    messages = []
    restored = []
    
    for filename in CRITICAL_FILES:
        filepath = os.path.join(BASE_DIR, filename)
        existed_before = os.path.exists(filepath)
        
        ok, content, _ = github_api_get_file(filename)
        if ok and content:
            with open(filepath, 'wb') as f:
                f.write(content)
            if not existed_before:
                restored.append(f'{filename}(新建)')
            else:
                restored.append(f'{filename}(已同步)')
        elif not ok:
            messages.append(f'{filename}拉取失败')
    
    if restored:
        messages.append(f'已同步 {len(restored)} 个文件: {", ".join(restored)}')
    
    return True, '（拉取成功；' + '；'.join(messages) + '）' if messages else '拉取成功'


def github_api_push(message: str = '同步数据') -> tuple[bool, str]:
    """
    将所有关键文件的变更推送到 GitHub（替代 git push）
    返回: (成功, 消息)
    """
    pushed = []
    errors = []
    
    for filename in CRITICAL_FILES:
        filepath = os.path.join(BASE_DIR, filename)
        if not os.path.exists(filepath):
            continue
        
        try:
            with open(filepath, 'rb') as f:
                content = f.read()
        except Exception as e:
            errors.append(f'{filename}读取失败: {e}')
            continue
        
        ok, msg = github_api_push_file(filename, content, message)
        if ok:
            pushed.append(filename)
        else:
            errors.append(f'{filename}: {msg}')
    
    # 同时推送 HTML 报表
    html_file = '项目延期点检表.html'
    html_path = os.path.join(BASE_DIR, html_file)
    if os.path.exists(html_path):
        try:
            with open(html_path, 'rb') as f:
                content = f.read()
            ok, msg = github_api_push_file(html_file, content, message)
            if ok:
                pushed.append(html_file)
        except:
            pass
    
    # 推送 data 目录中的 JSON 文件
    data_dir = os.path.join(BASE_DIR, 'data')
    if os.path.exists(data_dir):
        for fname in os.listdir(data_dir):
            if fname.endswith('.json') or fname.endswith('.log'):
                fpath = os.path.join(data_dir, fname)
                try:
                    with open(fpath, 'rb') as f:
                        content = f.read()
                    ok, _ = github_api_push_file(f'data/{fname}', content, message)
                    if ok:
                        pushed.append(f'data/{fname}')
                except:
                    pass
    
    if pushed:
        return True, f'已推送 {len(pushed)} 个文件到 GitHub'
    elif errors:
        return False, f'推送失败: {"; ".join(errors[:3])}'
    else:
        return True, '无变更需要推送'


# ==================== 统一接口（自动切换模式） ====================

def unified_pull() -> tuple[bool, str]:
    """统一拉取接口：自动选择 git 命令或 GitHub API"""
    if should_use_github_api():
        print('[同步] 使用 GitHub API 模式拉取数据')
        return github_api_pull()
    else:
        try:
            from sync_excel import git_pull
            return git_pull()
        except:
            return github_api_pull()


def unified_push(message: str = '同步数据') -> tuple[bool, str]:
    """统一推送接口：自动选择 git 命令或 GitHub API"""
    if should_use_github_api():
        print('[同步] 使用 GitHub API 模式推送数据')
        return github_api_push(message)
    else:
        try:
            from sync_excel import git_push
            return git_push(message)
        except:
            return github_api_push(message)


if __name__ == '__main__':
    print('=== GitHub API 同步模块测试（零依赖） ===')
    print(f'git 命令可用: {has_git_command()}')
    print(f'.git 目录存在: {has_git_repo()}')
    print(f'使用 API 模式: {should_use_github_api()}')
    print(f'urllib 可用: {HAS_URLLIB}')
    print(f'requests 可用: {HAS_REQUESTS}')
    
    token, owner, repo = _get_github_config()
    print(f'仓库: {owner}/{repo}')
    print(f'Token: {"***" + token[-4:] if token else "未找到"}')
    
    # 测试 HTTP 请求
    print('\n--- 测试 HTTP 请求 ---')
    status, body, _ = _http_request('https://api.github.com', 'GET')
    print(f'  GitHub API 连通性: HTTP {status}')
