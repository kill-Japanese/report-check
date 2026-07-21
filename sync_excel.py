# -*- coding: utf-8 -*-
"""
Excel 同步模块 - 以 GitHub Excel 为唯一数据源
功能：
  1. 将 Web 操作（新增/删除/归档/编辑）写回原始 Excel
  2. 提交并推送到 GitHub
  3. 重新生成 HTML 报表
  4. 启动时从 GitHub 拉取最新数据
"""

import os
import sys
import json
import subprocess
from datetime import datetime
import pandas as pd
from openpyxl import load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, 'data')
COLLAB_FILE = os.path.join(DATA_DIR, '协作数据.json')
EXCEL_FILE = os.path.join(BASE_DIR, '超声波户表脚本.xlsx')
HTML_FILE = os.path.join(BASE_DIR, '项目延期点检表.html')

# Excel 列配置（0-based 索引）
# 【严重修复】原使用U列(20)/V列(21)会覆盖延期计算公式，
# 导致Excel延期工时计算错误。改为使用完全空闲的A列(0)和B列(1)。
# 注意：openpyxl写入时使用 1-based 列号，所以 COL_ARCHIVED+1 = 第2列(B列)
COL_ARCHIVED = 0   # 第1列(A列)用于存放归档标志（原U列/20，改为A列/0，避免覆盖公式）
COL_DELETED = 1    # 第2列(B列)用于存放删除标志（原V列/21，改为B列/1，软删除避免合并单元格破坏）

# ==================== Git 仓库保障 ====================

def ensure_git_repo() -> tuple[bool, str]:
    """确保 Git 仓库存在并且配置正确（Render 部署环境保障）
    
    【关键修复】Render 部署时可能不会保留 .git 目录，
    导致所有 Git 操作失败，数据无法持久化。
    
    修复策略：
    1. 检查 .git 目录是否存在
    2. 如果不存在，尝试从环境变量或已知配置初始化 Git 仓库
    3. 检查 remote origin 是否配置正确
    4. 确保用户信息配置正确（用于 commit）
    """
    git_dir = os.path.join(BASE_DIR, '.git')
    
    try:
        # 1. 检查 .git 目录是否存在
        if not os.path.exists(git_dir):
            print(f"[Git] .git 目录不存在，正在初始化...")
            
            # 初始化 Git 仓库
            init = subprocess.run(
                ['git', 'init'],
                capture_output=True, text=True, cwd=BASE_DIR, timeout=10
            )
            if init.returncode != 0:
                return False, f'Git 初始化失败: {init.stderr[:200]}'
            
            # 尝试从环境变量获取 remote URL
            remote_url = os.environ.get('GIT_REMOTE_URL', '')
            if not remote_url:
                # 尝试从常见的 Render 环境变量中推断
                # Render 会设置一些环境变量，但不包含完整的 repo URL
                # 这里使用一个默认的占位符，用户需要在环境变量中配置
                print(f"[Git] 警告: 未设置 GIT_REMOTE_URL 环境变量")
            
            if remote_url:
                subprocess.run(
                    ['git', 'remote', 'add', 'origin', remote_url],
                    capture_output=True, cwd=BASE_DIR, timeout=10
                )
                print(f"[Git] 已设置 remote origin: {remote_url[:50]}...")
        else:
            # .git 存在，检查 remote 配置
            remote_check = subprocess.run(
                ['git', 'remote', '-v'],
                capture_output=True, text=True, cwd=BASE_DIR, timeout=5
            )
            if 'origin' not in remote_check.stdout:
                print(f"[Git] 警告: 未配置 remote origin")
        
        # 2. 确保 Git 用户信息配置正确（用于 commit）
        subprocess.run(
            ['git', 'config', 'user.email', 'server@report-check.local'],
            capture_output=True, cwd=BASE_DIR, timeout=5
        )
        subprocess.run(
            ['git', 'config', 'user.name', 'Report Check Server'],
            capture_output=True, cwd=BASE_DIR, timeout=5
        )
        
        # 3. 确保分支名为 main
        branch = subprocess.run(
            ['git', 'branch', '--show-current'],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=5
        )
        current_branch = branch.stdout.strip()
        if current_branch and current_branch != 'main':
            subprocess.run(
                ['git', 'branch', '-M', 'main'],
                capture_output=True, cwd=BASE_DIR, timeout=5
            )
            print(f"[Git] 已将分支 {current_branch} 重命名为 main")
        
        return True, 'Git 仓库就绪'
        
    except Exception as e:
        return False, f'Git 仓库初始化失败: {str(e)}'


# ==================== Git 操作 ====================

def git_pull() -> tuple[bool, str]:
    """从 GitHub 拉取最新数据（双路径：git命令优先，GitHub API兜底）
    
    【关键修复】Render 环境中 git 命令可能不存在，
    自动检测并回退到 GitHub REST API 模式。
    """
    # 检查 git 命令是否可用
    git_available = False
    try:
        result = subprocess.run(['git', '--version'], capture_output=True, timeout=5)
        git_available = (result.returncode == 0) and os.path.exists(os.path.join(BASE_DIR, '.git'))
    except (FileNotFoundError, OSError):
        git_available = False
    
    if not git_available:
        # 回退到 GitHub API 模式
        try:
            from github_sync import github_api_pull
            print('[sync] git 不可用，使用 GitHub API 模式拉取')
            return github_api_pull()
        except Exception as api_e:
            return False, f'git不可用且API拉取失败: {api_e}'
    
    # git 可用，使用原有逻辑
    try:
        ensure_ok, ensure_msg = ensure_git_repo()
        if not ensure_ok:
            return False, f'Git仓库不可用: {ensure_msg}'
        if not os.path.exists(os.path.join(BASE_DIR, '.git')):
            return False, '未检测到 Git 仓库'

        fetch = subprocess.run(
            ['git', 'fetch', 'origin', 'main'],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=30
        )
        if fetch.returncode != 0:
            return False, f'fetch失败: {fetch.stderr[:200]}'

        critical_files = ['用户管理.xlsx', '超声波户表脚本.xlsx']
        restored = []
        for f in critical_files:
            fpath = os.path.join(BASE_DIR, f)
            existed_before = os.path.exists(fpath)
            checkout = subprocess.run(
                ['git', 'checkout', 'origin/main', '--', f],
                capture_output=True, text=True, cwd=BASE_DIR, timeout=10
            )
            if checkout.returncode == 0 and os.path.exists(fpath):
                restored.append(f'{f}(新建)' if not existed_before else f'{f}(已同步)')

        pull = subprocess.run(
            ['git', 'pull', 'origin', 'main'],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=30
        )

        msg_parts = ['拉取成功']
        if restored:
            msg_parts.append(f'已同步 {len(restored)} 个文件: {", ".join(restored)}')
        return True, '（' + '；'.join(msg_parts) + '）'
    except subprocess.TimeoutExpired:
        return False, '拉取超时'
    except Exception as e:
        return False, f'拉取失败: {str(e)}'

def git_push(message: str = '同步数据') -> tuple[bool, str]:
    """将变更提交并推送到 GitHub（双路径：git命令优先，GitHub API兜底）
    
    【关键修复】Render 环境中 git 命令可能不存在，
    自动检测并回退到 GitHub REST API 模式。
    """
    # 检查 git 命令是否可用
    git_available = False
    try:
        result = subprocess.run(['git', '--version'], capture_output=True, timeout=5)
        git_available = (result.returncode == 0) and os.path.exists(os.path.join(BASE_DIR, '.git'))
    except (FileNotFoundError, OSError):
        git_available = False
    
    if not git_available:
        # 回退到 GitHub API 模式
        try:
            from github_sync import github_api_push
            print('[sync] git 不可用，使用 GitHub API 模式推送')
            return github_api_push(message)
        except Exception as api_e:
            return False, f'git不可用且API推送失败: {api_e}'
    
    # git 可用，使用原有逻辑
    try:
        ensure_ok, ensure_msg = ensure_git_repo()
        if not ensure_ok:
            return False, f'Git仓库不可用: {ensure_msg}'
        if not os.path.exists(os.path.join(BASE_DIR, '.git')):
            return False, '未检测到 Git 仓库'

        ahead = subprocess.run(
            ['git', 'rev-list', '--count', 'origin/main..HEAD'],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=10
        )
        try:
            ahead_count = int(ahead.stdout.strip())
        except ValueError:
            ahead_count = 0

        if ahead_count > 0:
            push = subprocess.run(
                ['git', 'push', 'origin', 'main'],
                capture_output=True, text=True, cwd=BASE_DIR, timeout=30
            )
            if push.returncode != 0:
                return False, f'推送失败（重试 {ahead_count} 个待推送提交）: {push.stderr[:200]}'
            return True, f'已推送 {ahead_count} 个待提交到 GitHub'

        result = subprocess.run(
            ['git', 'status', '--porcelain'],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=10
        )
        if not result.stdout.strip():
            return True, '无变更，无需推送'

        subprocess.run(['git', 'add', '-A'], capture_output=True, cwd=BASE_DIR, timeout=10)
        commit_msg = f'[数据同步] {message} - {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
        commit_result = subprocess.run(
            ['git', 'commit', '-m', commit_msg],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=10
        )
        if commit_result.returncode != 0:
            if 'nothing to commit' in (commit_result.stdout + commit_result.stderr):
                return True, '无变更'
            return False, f'提交失败: {commit_result.stderr[:200]}'

        push_result = subprocess.run(
            ['git', 'push', 'origin', 'main'],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=30
        )
        if push_result.returncode != 0:
            subprocess.run(
                ['git', 'reset', '--soft', 'HEAD~1'],
                capture_output=True, cwd=BASE_DIR, timeout=10
            )
            return False, f'推送失败（已撤销本地提交，可重试）: {push_result.stderr[:200]}'
        return True, '已同步到 GitHub'
    except subprocess.TimeoutExpired:
        return False, '同步超时'
    except Exception as e:
        return False, f'同步失败: {str(e)}'

# ==================== 协作数据读写 ====================

def load_collab_data() -> dict:
    """加载协作数据（新增项目、归档、编辑等）"""
    default = {
        'localEdits': {}, 'notes': {}, 'checked': {},
        'archived': {}, 'customEmails': {}, 'newProjects': [],
        'deletedIds': [],
        'lastUpdate': datetime.now().isoformat()
    }
    if os.path.exists(COLLAB_FILE):
        try:
            with open(COLLAB_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            # 确保所有默认键都存在
            for k, v in default.items():
                if k not in data:
                    data[k] = v
            return data
        except:
            pass
    return default

def save_collab_data(data: dict):
    """保存协作数据"""
    data['lastUpdate'] = datetime.now().isoformat()
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(COLLAB_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

# ==================== Excel 读写（核心） ====================

def read_excel_projects() -> list:
    """从原始 Excel 读取所有项目资源（与更新点检表.py 逻辑一致）"""
    if not os.path.exists(EXCEL_FILE):
        return []
    
    df = pd.read_excel(EXCEL_FILE, sheet_name='任务计划表', header=None)
    projects = []
    current_dept = None
    current_project = None
    current_start = None
    current_end = None
    current_desc = None
    
    for idx in range(3, len(df)):
        row = df.iloc[idx]
        
        if pd.notna(row[4]):
            val = row[4]
            current_dept = str(val) if not isinstance(val, float) else val
        
        if pd.notna(row[5]):
            current_project = str(row[5])
            current_start = row[6] if pd.notna(row[6]) else None
            current_end = row[7] if pd.notna(row[7]) else None
            current_desc = str(row[8]) if pd.notna(row[8]) else ''
        else:
            # 即使项目名（F列）没值（合并单元格的后续行），
            # 也要检查项目描述（I列）是否有独立值（合并被取消后可能有编辑值）
            if pd.notna(row[8]):
                current_desc = str(row[8])
            # 同样检查开始/结束时间
            if pd.notna(row[6]):
                current_start = row[6]
            if pd.notna(row[7]):
                current_end = row[7]
        
        resource_type = str(row[9]) if pd.notna(row[9]) else ''
        resource_name = str(row[10]) if pd.notna(row[10]) else ''
        
        # 清理资源名称
        if resource_name:
            resource_name = resource_name.strip()
            if resource_name.startswith('@'):
                resource_name = resource_name[1:]
            elif resource_name.startswith(' @'):
                resource_name = resource_name[2:]
            resource_name = resource_name.strip()
            if '(' in resource_name:
                resource_name = resource_name[:resource_name.index('(')].strip()
        
        has_resource = resource_type.strip() or resource_name.strip()
        
        # 读取归档标志（第21列，U列）
        archived_flag = ''
        if COL_ARCHIVED < len(row) and pd.notna(row[COL_ARCHIVED]):
            archived_flag = str(row[COL_ARCHIVED]).strip()
        is_archived = archived_flag in ('已归档', '1', 'true', 'True', 'YES', 'yes', 'Y', 'y')
        
        # 读取删除标志（第22列，V列）- 软删除，已删除的项目不返回
        deleted_flag = ''
        if COL_DELETED < len(row) and pd.notna(row[COL_DELETED]):
            deleted_flag = str(row[COL_DELETED]).strip()
        is_deleted = deleted_flag in ('已删除', '1', 'true', 'True', 'YES', 'yes', 'Y', 'y')
        
        if is_deleted:
            continue  # 跳过已删除的项目
        
        if current_project and has_resource:
            projects.append({
                'id': idx,
                '部门': current_dept if current_dept and not (isinstance(current_dept, float) and pd.isna(current_dept)) else '',
                '项目': current_project,
                '项目开始时间': current_start,
                '项目结束时间': current_end,
                '项目描述': current_desc,
                '资源类型': resource_type,
                '资源名称': resource_name,
                '资源开始时间': row[11] if pd.notna(row[11]) else None,
                '资源结束时间': row[12] if pd.notna(row[12]) else None,
                '日平均工时': row[13] if pd.notna(row[13]) else 0,
                '已归档': is_archived,
            })
    
    return projects

def _find_merged_range(ws, row: int, col: int):
    """查找单元格所属的合并区域，返回 (min_row, max_row, min_col, max_col) 或 None"""
    from openpyxl.cell.cell import MergedCell
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        return None
    for merged_range in ws.merged_cells.ranges:
        if (merged_range.min_row <= row <= merged_range.max_row and
            merged_range.min_col <= col <= merged_range.max_col):
            return merged_range
    return None


def _safe_write_cell(ws, row: int, col: int, value) -> bool:
    """
    安全地写入Excel单元格，自动处理合并单元格：
    - 如果单元格不是合并单元格，直接写入
    - 如果单元格是合并区域的左上角，直接写入
    - 如果单元格在合并区域内但不是左上角：
      1. 先取消合并
      2. 将原值填充到原合并区域的所有单元格
      3. 修改目标单元格的值
    返回: 是否成功写入
    """
    from openpyxl.cell.cell import MergedCell
    
    merged_range = _find_merged_range(ws, row, col)
    
    if merged_range is None:
        # 不是合并单元格，直接写入
        ws.cell(row=row, column=col, value=value)
        return True
    
    # 是合并单元格
    min_r, max_r = merged_range.min_row, merged_range.max_row
    min_c, max_c = merged_range.min_col, merged_range.max_col
    
    if row == min_r and col == min_c:
        # 恰好是合并区域的左上角，直接写入
        ws.cell(row=row, column=col, value=value)
        return True
    
    # 在合并区域内但不是左上角：需要先取消合并
    # 1. 先获取合并区域的原值
    original_value = ws.cell(row=min_r, column=min_c).value
    
    # 2. 取消合并
    ws.unmerge_cells(str(merged_range))
    
    # 3. 将原值填充到原合并区域的所有单元格
    for r in range(min_r, max_r + 1):
        for c in range(min_c, max_c + 1):
            ws.cell(row=r, column=c, value=original_value)
    
    # 4. 写入目标单元格的新值
    ws.cell(row=row, column=col, value=value)
    return True


def apply_collab_to_excel() -> tuple[bool, str, int]:
    """
    将协作数据（新增/删除/归档/编辑）应用到原始 Excel
    返回: (成功, 消息, 变更数量)
    
    执行顺序（关键！避免行号漂移）：
    1. 先处理归档（不改变行号，使用初始映射）
    2. 再处理编辑（不改变行号，使用初始映射）
    3. 再处理删除（从大到小删除，使用初始映射）
    4. 最后处理新增（追加到末尾，不影响已有行号）
    """
    collab = load_collab_data()
    changes = 0
    errors = []
    
    if not os.path.exists(EXCEL_FILE):
        return False, '原始 Excel 文件不存在', 0
    
    # 读取所有现有项目（仅一次，作为行号映射的基准）
    existing_projects = read_excel_projects()
    # id是df的idx，Excel行号=idx+1（因为df从0开始，Excel从1开始，且前3行是表头）
    # 实际上：df.idx=3 对应 Excel第4行，所以 Excel行号 = idx + 1
    id_to_excel_row = {p['id']: p['id'] + 1 for p in existing_projects}
    
    try:
        wb = load_workbook(EXCEL_FILE)
        ws = wb['任务计划表']
        
        # ============== 1. 先处理归档（不改变行号） ==============
        archived = collab.get('archived', {})
        
        # 构建需要归档的行号集合
        archived_rows = set()
        if archived:
            for pid, arch_info in archived.items():
                try:
                    pid_int = int(pid)
                except:
                    pid_int = pid
                row_num = id_to_excel_row.get(pid_int) or id_to_excel_row.get(str(pid_int))
                if row_num and arch_info:
                    archived_rows.add(row_num)
        
        # 【修复】只处理需要变更的行，不盲目清空所有行的U列（避免覆盖原有公式）
        # - 需要归档的行：写入"已归档"
        # - 不需要归档但当前U列是"已归档"的行：清空（恢复）
        # - 其他行（原有公式）：不做任何修改
        for p in existing_projects:
            row_num = id_to_excel_row.get(p['id'])
            if not row_num:
                continue
            
            current_val = ws.cell(row=row_num, column=COL_ARCHIVED + 1).value
            is_current_archived = (str(current_val).strip() == '已归档') if current_val else False
            
            if row_num in archived_rows:
                # 需要归档
                if not is_current_archived:
                    ws.cell(row=row_num, column=COL_ARCHIVED + 1, value='已归档')
                    changes += 1
                    print(f"   📦 归档: Excel第{row_num}行")
            else:
                # 不需要归档：只有之前被标记为"已归档"的才清空，保留原有公式
                if is_current_archived:
                    ws.cell(row=row_num, column=COL_ARCHIVED + 1, value='')
                    changes += 1
                    print(f"   📤 取消归档: Excel第{row_num}行")
        
        # ============== 1.5. 处理编辑（localEdits）- 不改变行号，在删除前应用 ==============
        # 字段名到Excel列号的映射（1-based）
        FIELD_TO_COLUMN = {
            '部门': 5,           # E列
            '项目': 6,           # F列
            '项目开始时间': 7,   # G列
            '项目结束时间': 8,   # H列
            '项目描述': 9,       # I列
            '资源类型': 10,      # J列
            '资源名称': 11,      # K列
            '资源开始时间': 12,  # L列
            '资源结束时间': 13,  # M列
            '日平均工时': 14,    # N列
        }
        
        local_edits = collab.get('localEdits', {})
        if local_edits:
            # 先构建要删除的ID集合，跳过已删除项目的编辑
            deleted_ids_for_edit = set()
            for did in collab.get('deletedIds', []):
                try:
                    deleted_ids_for_edit.add(int(did))
                except:
                    deleted_ids_for_edit.add(did)
                    deleted_ids_for_edit.add(str(did))
            
            for pid, edits in local_edits.items():
                # pid 可能是整数或字符串
                try:
                    pid_int = int(pid)
                except:
                    pid_int = pid
                
                # 跳过已删除项目的编辑
                if pid_int in deleted_ids_for_edit or str(pid_int) in deleted_ids_for_edit:
                    continue
                
                row_num = id_to_excel_row.get(pid_int) or id_to_excel_row.get(str(pid_int))
                if not row_num:
                    continue
                
                if not isinstance(edits, dict):
                    continue
                
                for field, value in edits.items():
                    col = FIELD_TO_COLUMN.get(field)
                    if col:
                        # 跳过空日期值
                        if field in ('项目开始时间', '项目结束时间', '资源开始时间', '资源结束时间'):
                            if not value or value in ['', '1900-01-01', '2100-01-01']:
                                continue
                        try:
                            # 使用安全写入，自动处理合并单元格
                            write_ok = _safe_write_cell(ws, row_num, col, value)
                            if write_ok:
                                changes += 1
                                print(f"   ✏️  编辑: Excel第{row_num}行, {field}={value}")
                            else:
                                errors.append(f"编辑失败: 第{row_num}行 {field}")
                        except Exception as e:
                            # 单个字段编辑失败不影响其他操作
                            err_msg = f"编辑第{row_num}行{field}失败: {str(e)}"
                            errors.append(err_msg)
                            print(f"   ⚠️  {err_msg}")
        
        # ============== 2. 再处理删除（软删除：标记V列为"已删除"，不物理删除行） ==============
        # 注意：不能用 ws.delete_rows() 物理删除，因为Excel有大量合并单元格
        # 物理删除会导致合并区域错乱，丢失多行数据
        deleted_ids = set()
        for did in collab.get('deletedIds', []):
            try:
                deleted_ids.add(int(did))
            except:
                deleted_ids.add(did)
        
        if deleted_ids:
            for p in existing_projects:
                pid = p['id']
                if pid in deleted_ids or str(pid) in deleted_ids:
                    row_num = id_to_excel_row.get(pid) or id_to_excel_row.get(str(pid))
                    if row_num:
                        # 软删除：在V列(COL_DELETED+1)标记"已删除"
                        try:
                            _safe_write_cell(ws, row_num, COL_DELETED + 1, '已删除')
                            changes += 1
                            print(f"   🗑️  软删除: Excel第{row_num}行 ({p.get('项目', '')})")
                        except Exception as e:
                            err_msg = f"软删除第{row_num}行失败: {str(e)}"
                            errors.append(err_msg)
                            print(f"   ⚠️  {err_msg}")
        
        # ============== 3. 最后处理新增项目（追加到末尾） ==============
        new_projects = collab.get('newProjects', [])
        
        if new_projects:
            # 【修复】构建 Excel 中已有的项目集合，防止重复添加
            # （当 git_push 失败后重试时，协作数据未被清空，需要去重）
            existing_set = set()
            for p in existing_projects:
                key = (str(p.get('项目', '')).strip(), str(p.get('资源名称', '')).strip())
                existing_set.add(key)
            
            last_row = ws.max_row
            
            for np in new_projects:
                # 去重检查：项目名+资源名称完全一致则跳过
                np_key = (str(np.get('项目', '')).strip(), str(np.get('资源名称', '')).strip())
                if np_key in existing_set:
                    print(f"   ⏭️  跳过重复新增: {np_key[0]} / {np_key[1]}")
                    continue
                last_row += 1
                
                # 列5: 部门（列E=5）
                ws.cell(row=last_row, column=5, value=np.get('部门', ''))
                # 列6: 项目名（列F=6）
                ws.cell(row=last_row, column=6, value=np.get('项目', ''))
                # 列7: 项目开始（列G=7）
                start_val = np.get('项目开始时间', '')
                if start_val and start_val not in ['', '1900-01-01']:
                    ws.cell(row=last_row, column=7, value=start_val)
                # 列8: 项目结束（列H=8）
                end_val = np.get('项目结束时间', '')
                if end_val and end_val not in ['', '2100-01-01']:
                    ws.cell(row=last_row, column=8, value=end_val)
                # 列9: 项目描述（列I=9）
                ws.cell(row=last_row, column=9, value=np.get('项目描述', ''))
                # 列10: 资源类型（列J=10）
                ws.cell(row=last_row, column=10, value=np.get('资源类型', ''))
                # 列11: 资源名称（列K=11）
                ws.cell(row=last_row, column=11, value=np.get('资源名称', ''))
                # 列12: 资源开始（列L=12）
                res_start = np.get('资源开始时间', '')
                if res_start and res_start not in ['', '1900-01-01']:
                    ws.cell(row=last_row, column=12, value=res_start)
                # 列13: 资源结束（列M=13）
                res_end = np.get('资源结束时间', '')
                if res_end and res_end not in ['', '2100-01-01']:
                    ws.cell(row=last_row, column=13, value=res_end)
                # 列14: 日平均工时（列N=14）
                ws.cell(row=last_row, column=14, value=np.get('日平均工时', 0) or 0)
                # 列21: 归档标志（列U=21）
                if np.get('已归档'):
                    ws.cell(row=last_row, column=COL_ARCHIVED + 1, value='已归档')
                
                changes += 1
                print(f"   ➕ 新增: {np.get('项目', '')} / {np.get('资源名称', '')}")
        
        wb.save(EXCEL_FILE)
        
    except Exception as e:
        return False, f'写入 Excel 失败: {str(e)}', changes
    
    if errors:
        error_detail = '; '.join(errors[:5])  # 最多显示5个错误
        if len(errors) > 5:
            error_detail += f' 等{len(errors)}个错误'
        return True, f'已应用 {changes} 项变更到 Excel (部分警告: {error_detail})', changes
    
    return True, f'已应用 {changes} 项变更到 Excel', changes

def regenerate_report() -> tuple[bool, str]:
    """重新生成 HTML 报表"""
    try:
        script_path = os.path.join(BASE_DIR, '更新点检表.py')
        result = subprocess.run(
            [sys.executable, script_path],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=60
        )
        if result.returncode != 0:
            return False, f'生成报表失败: {result.stderr[:300]}'
        return True, '报表已重新生成'
    except subprocess.TimeoutExpired:
        return False, '生成报表超时'
    except Exception as e:
        return False, f'生成报表异常: {str(e)}'

# ==================== 全量同步入口 ====================

def full_sync(operation: str = '未知操作') -> tuple[bool, str]:
    """
    执行完整同步流程：
    1. 将协作数据应用到 Excel
    2. 重新生成报表
    3. 推送到 GitHub
    4. 【关键修复】只有推送成功后才清空协作数据，防止推送失败时数据丢失
    
    返回: (整体是否成功, 消息)
    - 只有 Excel写入+报表生成+Git推送 全部成功才返回 True
    - 任何一步失败都返回 False，协作数据保留以便下次重试
    """
    messages = []
    
    # 步骤1: 应用协作数据到 Excel（最关键的一步）
    ok, msg, changes = apply_collab_to_excel()
    messages.append(msg)
    if not ok:
        return False, '; '.join(messages)
    
    if changes > 0:
        # 步骤2: 重新生成报表
        ok, msg = regenerate_report()
        messages.append(msg)
        if not ok:
            # 报表生成失败但数据已在Excel中，保留协作数据以便下次重试
            messages.append('警告：报表生成失败，协作数据保留待重试')
            return False, '；'.join(messages)
        
        # 步骤3: 推送到 GitHub
        push_ok, push_msg = git_push(f'{operation}，{changes}项变更')
        messages.append(push_msg)
        if not push_ok:
            # 【关键修复】推送失败时不清空协作数据，保留以便下次重试
            messages.append('错误：GitHub推送失败，协作数据已保留待重试（请勿重启服务器！）')
            return False, '；'.join(messages)
        
        # 【关键修复】只有推送成功后才清空协作数据
        collab = load_collab_data()
        collab['newProjects'] = []
        collab['deletedIds'] = []
        collab['archived'] = {}
        collab['localEdits'] = {}
        collab['notes'] = {}
        collab['checked'] = {}
        save_collab_data(collab)
    else:
        # 【修复】changes=0 但协作数据不为空时，说明是重试场景
        # （Excel已被修改但上次推送失败，协作数据还保留着）
        # 直接检查是否有未提交的变更并尝试推送
        collab = load_collab_data()
        has_pending = (
            collab.get('newProjects') or 
            collab.get('deletedIds') or 
            collab.get('archived') or 
            collab.get('localEdits')
        )
        if has_pending:
            messages.append('检测到待推送的协作数据（重试场景）')
            # 重新生成报表（确保报表是最新的）
            ok, msg = regenerate_report()
            messages.append(msg)
            if not ok:
                messages.append('警告：报表生成失败')
                return False, '；'.join(messages)
            # 尝试推送
            push_ok, push_msg = git_push(f'{operation}，重试推送')
            messages.append(push_msg)
            if not push_ok:
                messages.append('错误：GitHub推送失败，协作数据已保留待重试')
                return False, '；'.join(messages)
            # 推送成功，清空协作数据
            collab['newProjects'] = []
            collab['deletedIds'] = []
            collab['archived'] = {}
            collab['localEdits'] = {}
            collab['notes'] = {}
            collab['checked'] = {}
            save_collab_data(collab)
        else:
            # 【关键修复】即使没有待处理的协作数据，也要检查是否有其他文件变更
            # （比如超声波户表脚本.xlsx 被外部直接修改、报表重新生成等）
            # 确保所有变更都能被推送到 GitHub，避免重新部署后丢失
            has_file_changes = False
            try:
                status = subprocess.run(
                    ['git', 'status', '--porcelain',
                     '超声波户表脚本.xlsx', '用户管理.xlsx',
                     '项目延期点检表.html', 'data/'],
                    capture_output=True, text=True, cwd=BASE_DIR, timeout=10
                )
                if status.stdout.strip():
                    has_file_changes = True
                    print(f"[同步] 检测到非协作数据的文件变更:\n{status.stdout.strip()}")
            except:
                pass
            
            if has_file_changes:
                messages.append('检测到文件变更，正在同步...')
                # 重新生成报表（确保报表是最新的）
                ok, msg = regenerate_report()
                messages.append(msg)
                if not ok:
                    messages.append('警告：报表生成失败')
                # 尝试推送所有变更
                push_ok, push_msg = git_push(f'{operation}，文件变更同步')
                messages.append(push_msg)
                if not push_ok:
                    messages.append('错误：GitHub推送失败')
                    return False, '；'.join(messages)
            else:
                messages.append('无需要同步的变更')
    
    return True, '；'.join(messages)

def startup_sync() -> tuple[bool, str]:
    """服务器启动时同步：确保Git仓库 + 拉取最新 + 生成报表
    
    【关键修复】Render 部署时 .git 目录可能不存在，
    必须先确保 Git 仓库可用，再进行拉取操作。
    """
    messages = []
    
    # 先确保 Git 仓库存在（Render 部署环境保障）
    ensure_ok, ensure_msg = ensure_git_repo()
    messages.append(ensure_msg)
    if not ensure_ok:
        messages.append('警告：Git仓库不可用，将无法同步数据到GitHub')
    
    ok, msg = git_pull()
    messages.append(msg)
    ok2, msg2 = regenerate_report()
    messages.append(msg2)
    return (ok and ok2), '；'.join(messages)

# ==================== 测试 ====================

if __name__ == '__main__':
    print('=== 测试 Excel 同步模块 ===')
    print(f'Excel 文件: {EXCEL_FILE}')
    print(f'协作数据文件: {COLLAB_FILE}')
    
    projects = read_excel_projects()
    print(f'\\n读取到 {len(projects)} 条资源记录')
    
    collab = load_collab_data()
    print(f'协作数据: newProjects={len(collab.get("newProjects", []))}, '
          f'archived={len(collab.get("archived", {}))}, '
          f'localEdits={len(collab.get("localEdits", {}))}')
    
    if len(sys.argv) > 1 and sys.argv[1] == '--sync':
        print('\\n=== 执行全量同步 ===')
        ok, msg = full_sync('手动触发')
        print(f'结果: {"成功" if ok else "失败"} - {msg}')
