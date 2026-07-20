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
COL_ARCHIVED = 20  # 第21列(U列)用于存放归档标志

# ==================== Git 操作 ====================

def git_pull() -> tuple[bool, str]:
    """从 GitHub 拉取最新数据"""
    try:
        if not os.path.exists(os.path.join(BASE_DIR, '.git')):
            return False, '未检测到 Git 仓库'
        result = subprocess.run(
            ['git', 'pull', 'origin', 'main'],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=30
        )
        if result.returncode != 0:
            return False, f'拉取失败: {result.stderr[:200]}'
        return True, '拉取成功'
    except subprocess.TimeoutExpired:
        return False, '拉取超时'
    except Exception as e:
        return False, f'拉取失败: {str(e)}'

def git_push(message: str = '同步数据') -> tuple[bool, str]:
    """将变更提交并推送到 GitHub"""
    try:
        if not os.path.exists(os.path.join(BASE_DIR, '.git')):
            return False, '未检测到 Git 仓库'
        
        # 检查是否有变更
        result = subprocess.run(
            ['git', 'status', '--porcelain'],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=10
        )
        if not result.stdout.strip():
            return True, '无变更，无需推送'
        
        subprocess.run(['git', 'add', '-A'], capture_output=True, cwd=BASE_DIR, timeout=10)
        commit_msg = f'[数据同步] {message} - {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}'
        subprocess.run(['git', 'commit', '-m', commit_msg], capture_output=True, cwd=BASE_DIR, timeout=10)
        push_result = subprocess.run(
            ['git', 'push', 'origin', 'main'],
            capture_output=True, text=True, cwd=BASE_DIR, timeout=30
        )
        if push_result.returncode != 0:
            return False, f'推送失败: {push_result.stderr[:200]}'
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

def apply_collab_to_excel() -> tuple[bool, str, int]:
    """
    将协作数据（新增/删除/归档/编辑）应用到原始 Excel
    返回: (成功, 消息, 变更数量)
    
    执行顺序（关键！避免行号漂移）：
    1. 先处理归档（不改变行号，使用初始映射）
    2. 再处理删除（从大到小删除，使用初始映射）
    3. 最后处理新增（追加到末尾，不影响已有行号）
    """
    collab = load_collab_data()
    changes = 0
    
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
        
        # 先清空所有已有的归档标志
        for p in existing_projects:
            row_num = id_to_excel_row.get(p['id'])
            if row_num:
                ws.cell(row=row_num, column=COL_ARCHIVED + 1, value='')
        
        # 再写入新的归档标志
        if archived:
            for pid, arch_info in archived.items():
                # pid 可能是整数或字符串
                try:
                    pid_int = int(pid)
                except:
                    pid_int = pid
                
                row_num = id_to_excel_row.get(pid_int) or id_to_excel_row.get(str(pid_int))
                if row_num and arch_info:
                    ws.cell(row=row_num, column=COL_ARCHIVED + 1, value='已归档')
                    changes += 1
                    print(f"   📦 归档: Excel第{row_num}行")
        
        # ============== 2. 再处理删除（从大到小删除，避免行号漂移） ==============
        deleted_ids = set()
        for did in collab.get('deletedIds', []):
            try:
                deleted_ids.add(int(did))
            except:
                deleted_ids.add(did)
        
        if deleted_ids:
            # 获取要删除的Excel行号，从大到小排序
            rows_to_delete = []
            for p in existing_projects:
                if p['id'] in deleted_ids or str(p['id']) in deleted_ids:
                    rows_to_delete.append(id_to_excel_row[p['id']])
            
            rows_to_delete.sort(reverse=True)
            for row_num in rows_to_delete:
                ws.delete_rows(row_num)
                changes += 1
                print(f"   🗑️  删除Excel第{row_num}行")
        
        # ============== 3. 最后处理新增项目（追加到末尾） ==============
        new_projects = collab.get('newProjects', [])
        
        if new_projects:
            last_row = ws.max_row
            
            for np in new_projects:
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
    
    # 处理编辑（localEdits）- 计入变更数
    local_edits = collab.get('localEdits', {})
    if local_edits:
        changes += len(local_edits)
    
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
    4. 清空已处理的协作数据（已写入 Excel 的部分）
    """
    messages = []
    
    # 步骤1: 应用协作数据到 Excel
    ok, msg, changes = apply_collab_to_excel()
    messages.append(msg)
    if not ok:
        return False, '; '.join(messages)
    
    if changes > 0:
        # 步骤2: 重新生成报表
        ok, msg = regenerate_report()
        messages.append(msg)
        if not ok:
            return False, '; '.join(messages)
        
        # 步骤3: 推送到 GitHub
        ok, msg = git_push(f'{operation}，{changes}项变更')
        messages.append(msg)
        if not ok:
            return False, '; '.join(messages)
        
        # 步骤4: 清空已写入 Excel 的协作数据（已持久化到Excel，避免下次重复应用）
        collab = load_collab_data()
        # 已写入Excel的：新项目、删除ID、归档标志
        if collab.get('newProjects'):
            collab['newProjects'] = []
        if collab.get('deletedIds'):
            collab['deletedIds'] = []
        if collab.get('archived'):
            collab['archived'] = {}
        save_collab_data(collab)
    else:
        messages.append('无需要同步的变更')
    
    return True, '；'.join(messages)

def startup_sync() -> tuple[bool, str]:
    """服务器启动时同步：拉取最新 + 生成报表"""
    messages = []
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
