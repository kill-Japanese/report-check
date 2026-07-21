# -*- coding: utf-8 -*-
"""
GitHub API 同步模块 - 当 git 命令不可用时的兜底方案

使用 GitHub REST API 直接操作仓库文件，不依赖 git 命令行工具。
适用于 Render 等容器环境中 git 命令不可用的场景。

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
import requests
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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
    
    # 尝试从 git remote URL 解析
    if not token:
        try:
            result = subprocess.run(
                ['git', 'remote', '-v'],
                capture_output=True, text=True, cwd=BASE_DIR, timeout=5
            )
            for line in result.stdout.split('\n'):
                if 'origin' in line and 'github.com' in line:
                    # 解析: https://token@github.com/owner/repo.git
                    url = line.split()[1]
                    if '@github.com' in url:
                        parts = url.split('@github.com/')[1].rstrip('.git').split('/')
                        if len(parts) >= 2:
                            owner = parts[0]
                            repo = '/'.join(parts[1:])
                            # 从 URL 中提取 token
                            if 'https://' in url:
                                token_part = url.split('https://')[1].split('@')[0]
                                if ':' in token_part:
                                    token = token_part.split(':')[1]
                                else:
                                    token = token_part
                            return token, owner, repo
        except:
            pass
    
    # 从环境变量获取 owner/repo
    owner = os.environ.get('GITHUB_OWNER', '')
    repo = os.environ.get('GITHUB_REPO', '')
    
    # 如果没有从 git 配置中解析到，尝试已知的配置
    if not owner or not repo:
        # 从已知的 remote URL 推断（本项目）
        owner = 'kill-Japanese'
        repo = 'report-check'
    
    return token, owner, repo


# ==================== GitHub API 核心操作 ====================

GITHUB_API_BASE = 'https://api.github.com'


def github_api_get_file(path: str, branch: str = 'main') -> tuple[bool, bytes, str]:
    """
    从 GitHub 获取文件内容
    返回: (成功, 文件内容(bytes), 文件SHA)
    """
    token, owner, repo = _get_github_config()
    if not token:
        return False, b'', ''
    
    url = f'{GITHUB_API_BASE}/repos/{owner}/{repo}/contents/{path}'
    headers = {
        'Authorization': f'token {token}',
        'Accept': 'application/vnd.github.v3+json'
    }
    params = {'ref': branch}
    
    try:
        resp = requests.get(url, headers=headers, params=params, timeout=30)
        if resp.status_code == 200:
            data = resp.json()
            content = base64.b64decode(data['content'])
            sha = data.get('sha', '')
            return True, content, sha
        elif resp.status_code == 404:
            return True, b'', ''  # 文件不存在不算错误
        else:
            print(f'[GitHub API] 获取文件失败: {resp.status_code} {resp.text[:200]}')
            return False, b'', ''
    except Exception as e:
        print(f'[GitHub API] 获取文件异常: {e}')
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
    
    url = f'{GITHUB_API_BASE}/repos/{owner}/{repo}/contents/{path}'
    headers = {
        'Authorization': f'token {token}',
        'Accept': 'application/vnd.github.v3+json'
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
    
    try:
        resp = requests.put(url, headers=headers, json=payload, timeout=30)
        if resp.status_code in (200, 201):
            return True, '已同步到 GitHub'
        else:
            err_msg = resp.text[:300]
            print(f'[GitHub API] 推送文件失败: {resp.status_code} {err_msg}')
            return False, f'API推送失败: {resp.status_code}'
    except Exception as e:
        print(f'[GitHub API] 推送文件异常: {e}')
        return False, f'API推送异常: {str(e)}'


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
        
        # 读取本地文件内容
        try:
            with open(filepath, 'rb') as f:
                content = f.read()
        except Exception as e:
            errors.append(f'{filename}读取失败: {e}')
            continue
        
        # 推送到 GitHub
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
        # 使用原有的 git pull（从 sync_excel 导入）
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
        # 使用原有的 git push（从 sync_excel 导入）
        try:
            from sync_excel import git_push
            return git_push(message)
        except:
            return github_api_push(message)


if __name__ == '__main__':
    print('=== GitHub API 同步模块测试 ===')
    print(f'git 命令可用: {has_git_command()}')
    print(f'.git 目录存在: {has_git_repo()}')
    print(f'使用 API 模式: {should_use_github_api()}')
    
    token, owner, repo = _get_github_config()
    print(f'仓库: {owner}/{repo}')
    print(f'Token: {"***" + token[-4:] if token else "未找到"}')
