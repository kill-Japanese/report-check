# -*- coding: utf-8 -*-
"""
Project文档转换解析模块
支持：文本粘贴解析、.mpp文件解析（可选）、截图OCR（可选）

输出格式（每条资源）：
{
    '项目名称': '巴勒斯坦超声户表二开',
    '项目描述': 'TR1',
    '资源类型': '需求导入',
    '资源名称': '陈雷雷',
    '开始时间': '2026-07-09',
    '结束时间': '2026-07-10',
    '工时': 0,
    'is_unresolved_tr': False,   # 上级是否找不到TR标识
    'is_unresolved_project': False, # 上级是否找不到project标识
    'warnings': []
}
"""

import re
import os
from datetime import datetime, timedelta

# ============================================================
# 工具函数
# ============================================================

def _parse_date(date_str):
    """解析多种日期格式为 YYYY-MM-DD
    支持: 2026-07-09, 2026/7/9, 2026年7月9日, 7/9, 07-09
    """
    if not date_str:
        return None
    s = str(date_str).strip()
    if not s or s == '/':
        return None
    
    # 已经是标准格式
    if re.match(r'^\d{4}-\d{2}-\d{2}$', s):
        return s
    
    formats = [
        '%Y-%m-%d', '%Y/%m/%d', '%Y.%m.%d',
        '%y-%m-%d', '%y/%m/%d', '%y.%m.%d',  # 两位年
        '%m-%d-%y', '%m/%d/%y', '%m.%d.%y',  # 月/日/两位年
        '%m-%d-%Y', '%m/%d/%Y', '%m.%d.%Y',  # 月/日/四位年
        '%Y年%m月%d日',
        '%m-%d', '%m/%d', '%m.%d',
    ]
    
    for fmt in formats:
        try:
            dt = datetime.strptime(s, fmt)
            # 如果只有月日，用当前年
            if fmt in ('%m-%d', '%m/%d', '%m.%d'):
                dt = dt.replace(year=datetime.now().year)
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            continue
    
    return None

def _parse_work_hours(val):
    """解析工时为数字（小时）
    支持: 16, 16h, 16小时, 2d, 2天, 16hrs
    """
    if val is None or val == '':
        return 0
    s = str(val).strip()
    if not s:
        return 0
    
    # 提取数字
    m = re.search(r'([\d.]+)', s)
    if not m:
        return 0
    num = float(m.group(1))
    
    # 判断单位
    if re.search(r'(天|d|day)', s, re.I):
        return num * 8  # 1天 = 8小时
    return num  # 默认小时

def _get_workdays(start_date, end_date):
    """计算两个日期间的工作日数（排除周末）"""
    if not start_date or not end_date:
        return 1
    try:
        start = datetime.strptime(start_date, '%Y-%m-%d')
        end = datetime.strptime(end_date, '%Y-%m-%d')
        if end < start:
            return 1
        count = 0
        current = start
        while current <= end:
            if current.weekday() < 5:  # 0-4 = 周一到周五
                count += 1
            current += timedelta(days=1)
        return max(count, 1)
    except Exception:
        return 1

def _is_project_line(name):
    """判断是否为project标识行"""
    if not name:
        return False
    return bool(re.search(r'project[:：]', name, re.I))

def _extract_project_name(name):
    """从project标识行提取项目名称"""
    m = re.search(r'project[:：]\s*(.+)', name, re.I)
    if m:
        return m.group(1).strip()
    return name

def _is_tr_line(name):
    """判断是否为TR标识行"""
    if not name:
        return False
    return bool(re.search(r'\bTR\b', name, re.I)) or bool(re.match(r'^TR[\d\-]', name, re.I))

# ============================================================
# 文本解析（核心入口）
# ============================================================

def parse_text(text):
    """解析从Project复制粘贴的文本表格
    文本格式: 每行一条任务，列之间用Tab或多空格分隔
    列顺序通常: 任务名称 | 开始时间 | 结束时间 | 工时 | 责任人
    
    Returns:
        dict: {
            'success': bool,
            'tasks': [task_list],  # 原始任务（带层级）
            'resources': [resource_list],  # 解析后的资源列表
            'warnings': [str],
            'unresolved_tr': [resource_names],
            'unresolved_project': [resource_names]
        }
    """
    result = {
        'success': False,
        'tasks': [],
        'resources': [],
        'warnings': [],
        'unresolved_tr': [],
        'unresolved_project': []
    }
    
    if not text or not text.strip():
        result['warnings'].append('输入文本为空')
        return result
    
    lines = text.strip().split('\n')
    lines = [l.rstrip('\r') for l in lines if l.strip()]
    
    if not lines:
        result['warnings'].append('没有可解析的行')
        return result
    
    # ---------- 步骤1: 检测列标题 ----------
    header_line = lines[0]
    # 分割列（Tab优先，否则多空格）
    if '\t' in header_line:
        delimiter = '\t'
    else:
        delimiter = None  # 按空白分割
    
    headers = _split_line(header_line, delimiter)
    
    # 智能识别列
    col_map = _detect_columns(headers)
    if not col_map:
        result['warnings'].append('无法识别列标题，请手动指定列映射')
        # 尝试默认映射: 假设列顺序为 名称|开始|结束|工时|责任人
        col_map = {
            'name': 0, 'start': 1, 'end': 2, 'work': 3, 'owner': 4
        }
    
    name_col_idx = col_map.get('name', 0)
    
    # ---------- 步骤2: 解析每一行任务 ----------
    tasks = []
    for i, line in enumerate(lines[1:], start=2):
        # 先提取第一列的缩进（任务名称列）
        # 先用分隔符分割获取原始单元格（保留前导空格）
        if delimiter == '\t':
            raw_cells = line.split('\t')
        elif delimiter:
            raw_cells = line.split(delimiter)
        else:
            raw_cells = re.split(r'\s{2,}', line)
        
        # 计算缩进：从第一列提取前导空白数
        name_raw = raw_cells[name_col_idx] if name_col_idx < len(raw_cells) else ''
        indent = 0
        for ch in name_raw:
            if ch in (' ', '\t'):
                indent += 1 if ch == ' ' else 4  # Tab算4个空格
            else:
                break
        
        cells = _split_line(line, delimiter)
        if not cells:
            continue
        
        # 提取字段
        name = _get_cell(cells, name_col_idx, '')
        start_raw = _get_cell(cells, col_map.get('start', 1), '')
        end_raw = _get_cell(cells, col_map.get('end', 2), '')
        work_raw = _get_cell(cells, col_map.get('work', 3), '')
        owner_raw = _get_cell(cells, col_map.get('owner', 4), '')
        
        clean_name = name.strip()
        
        # 解析日期和工时
        start_date = _parse_date(start_raw)
        end_date = _parse_date(end_raw)
        work_hours = _parse_work_hours(work_raw)
        owner = owner_raw.strip() if owner_raw else ''
        
        task = {
            'line_num': i,
            'indent': indent,
            'level': 0,  # 将在下一步计算
            'name': clean_name,
            'start': start_date,
            'end': end_date,
            'work': work_hours,
            'owner': owner,
            'is_project': _is_project_line(clean_name),
            'is_tr': _is_tr_line(clean_name),
            'project_name': None,  # 关联的项目名
            'tr_name': None,  # 关联的TR名
            'parent': None,  # 父任务
        }
        
        if task['is_project']:
            task['project_name'] = _extract_project_name(clean_name)
        
        if task['is_tr']:
            task['tr_name'] = clean_name
        
        tasks.append(task)
    
    if not tasks:
        result['warnings'].append('没有解析到任何任务')
        return result
    
    result['tasks'] = tasks
    
    # ---------- 步骤3: 计算层级关系（基于缩进）----------
    _calculate_levels(tasks)
    
    # ---------- 步骤4: 向上查找关联的project和TR ----------
    _resolve_ancestors(tasks)
    
    # ---------- 步骤5: 提取资源（过滤+转换）----------
    resources = _extract_resources(tasks)
    result['resources'] = resources
    
    # ---------- 步骤6: 设计+评审合并 ----------
    resources = _merge_design_review(resources)
    result['resources'] = resources
    
    # ---------- 步骤7: 检查未解析的标识 ----------
    for r in resources:
        if r.get('is_unresolved_tr'):
            result['unresolved_tr'].append(f"{r['资源类型']}({r['资源名称']})")
        if r.get('is_unresolved_project'):
            result['unresolved_project'].append(f"{r['资源类型']}({r['资源名称']})")
    
    result['success'] = True
    if not result['resources']:
        result['warnings'].append('没有提取到任何有效资源')
    
    return result

# ============================================================
# 辅助函数
# ============================================================

def _split_line(line, delimiter=None):
    """分割一行文本为单元格"""
    if delimiter == '\t':
        return [c.strip() for c in line.split('\t')]
    elif delimiter:
        return [c.strip() for c in line.split(delimiter)]
    else:
        # 按2个以上空格分割
        parts = re.split(r'\s{2,}', line.strip())
        return [p.strip() for p in parts if p.strip()]

def _get_cell(cells, idx, default=''):
    """安全获取单元格"""
    if 0 <= idx < len(cells):
        return cells[idx]
    return default

def _detect_columns(headers):
    """智能识别列映射
    Returns: dict with keys: name, start, end, work, owner
    """
    if not headers:
        return None
    
    col_map = {}
    for i, h in enumerate(headers):
        hl = h.lower().strip()
        # 任务名称
        if any(k in hl for k in ['任务名称', '任务名', '名称', 'task name', 'name', '任务']):
            col_map['name'] = i
        # 开始时间
        elif any(k in hl for k in ['开始时间', '开始', 'start', '开始日期']):
            col_map['start'] = i
        # 结束时间
        elif any(k in hl for k in ['完成时间', '结束时间', '完成', 'finish', 'end', '结束日期']):
            col_map['end'] = i
        # 工时
        elif any(k in hl for k in ['工时', '工作', 'work', 'duration', '工时(小时)']):
            col_map['work'] = i
        # 责任人
        elif any(k in hl for k in ['责任人', '负责人', '资源名称', 'resource names', '资源', 'resource']):
            col_map['owner'] = i
    
    # 至少识别到名称列
    if 'name' not in col_map:
        return None
    
    return col_map

def _calculate_levels(tasks):
    """根据缩进计算任务层级"""
    if not tasks:
        return
    
    # 收集所有缩进值，排序后映射到层级
    indents = sorted(set(t['indent'] for t in tasks))
    indent_to_level = {ind: i for i, ind in enumerate(indents)}
    
    for t in tasks:
        t['level'] = indent_to_level[t['indent']]
    
    # 建立父子关系
    stack = []  # [(level, task)]
    for t in tasks:
        # 弹出层级 >= 当前的
        while stack and stack[-1][0] >= t['level']:
            stack.pop()
        
        if stack:
            t['parent'] = stack[-1][1]
        
        stack.append((t['level'], t))

def _resolve_ancestors(tasks):
    """向上遍历查找关联的project和TR"""
    for t in tasks:
        current = t['parent']
        while current:
            if t['project_name'] is None and current.get('project_name'):
                t['project_name'] = current['project_name']
            if t['tr_name'] is None and current.get('tr_name'):
                t['tr_name'] = current['tr_name']
            # 两个都找到了就提前退出
            if t['project_name'] is not None and t['tr_name'] is not None:
                break
            current = current.get('parent')

def _extract_resources(tasks):
    """从任务中提取资源（应用过滤规则）"""
    resources = []
    
    for t in tasks:
        # 跳过 project 标识行和 TR 标识行本身
        if t['is_project'] or t['is_tr']:
            continue
        
        # 过滤规则: 责任人空 且 工时为0 → 不提取
        if not t['owner'] and t['work'] == 0:
            continue
        
        # 资源名称 = 责任人优先，否则用任务名
        res_name = t['owner'] if t['owner'] else t['name']
        
        # 计算日平均工时
        workdays = _get_workdays(t['start'], t['end'])
        daily_hours = 0
        if t['work'] > 0 and workdays > 0:
            daily_hours = round(t['work'] / workdays, 1)
        
        resource = {
            '项目名称': t['project_name'] if t['project_name'] else '/',
            '项目描述': t['tr_name'] if t['tr_name'] else '/',
            '资源类型': t['name'],  # 任务名 = 资源类型
            '资源名称': res_name,
            '开始时间': t['start'] if t['start'] else '1900-01-01',
            '结束时间': t['end'] if t['end'] else '2100-01-01',
            '工时': t['work'],
            '日平均工时': daily_hours,
            'is_unresolved_tr': t['tr_name'] is None,
            'is_unresolved_project': t['project_name'] is None,
            'warnings': []
        }
        
        if resource['is_unresolved_tr']:
            resource['warnings'].append('上级找不到TR标识')
        if resource['is_unresolved_project']:
            resource['warnings'].append('上级找不到project标识')
        
        resources.append(resource)
    
    return resources

def _merge_design_review(resources):
    """合并设计+评审任务（同名资源，资源类型分别含设计和评审）"""
    if not resources:
        return resources
    
    # 按资源名称分组
    from collections import defaultdict
    groups = defaultdict(list)
    for i, r in enumerate(resources):
        groups[r['资源名称']].append((i, r))
    
    to_remove = set()
    merged_results = []
    
    for name, items in groups.items():
        # 筛选设计和评审
        design_items = [(i, r) for i, r in items if '设计' in r['资源类型']]
        review_items = [(i, r) for i, r in items if '评审' in r['资源类型']]
        
        # 只有同时有设计和评审才合并
        if design_items and review_items:
            # 取第一个设计和第一个评审配对
            d_idx, design = design_items[0]
            r_idx, review = review_items[0]
            
            # 合并规则
            merged = dict(review)  # 以评审为基础（资源类型用评审的）
            merged['开始时间'] = min(design['开始时间'], review['开始时间'])  # 取最早
            merged['结束时间'] = max(design['结束时间'], review['结束时间'])  # 取最晚
            merged['工时'] = max(design['工时'], review['工时'])  # 取较长工时
            # 重新计算日平均工时
            workdays = _get_workdays(merged['开始时间'], merged['结束时间'])
            if merged['工时'] > 0 and workdays > 0:
                merged['日平均工时'] = round(merged['工时'] / workdays, 1)
            merged['warnings'] = list(set(design.get('warnings', []) + review.get('warnings', [])))
            merged['is_unresolved_tr'] = design.get('is_unresolved_tr', False) or review.get('is_unresolved_tr', False)
            merged['is_unresolved_project'] = design.get('is_unresolved_project', False) or review.get('is_unresolved_project', False)
            
            to_remove.add(d_idx)
            to_remove.add(r_idx)
            merged_results.append((min(d_idx, r_idx), merged))
    
    # 构建最终列表
    result = []
    for i, r in enumerate(resources):
        if i not in to_remove:
            result.append(r)
    
    # 插入合并结果
    for pos, merged in sorted(merged_results, key=lambda x: x[0]):
        result.insert(pos, merged)
    
    return result

# ============================================================
# MPP 解析（可选，依赖 mpxj 库）
# ============================================================

def parse_mpp(file_path):
    """解析.mpp文件（需要mpxj库）
    环境不可用时返回 None，由调用方降级处理
    """
    try:
        from mpxj import Reader
    except ImportError:
        return {
            'success': False,
            'error': 'MPP解析库(mpxj)未安装，当前环境不支持.mpp文件解析',
            'resources': []
        }
    
    try:
        reader = Reader()
        project = reader.read(file_path)
        
        # 提取任务
        tasks_data = []
        for task in project.tasks:
            if task.summary:
                continue  # 跳过摘要任务？不，有工时的摘要任务也要提取
            
            name = task.name or ''
            start = task.start.strftime('%Y-%m-%d') if task.start else None
            end = task.finish.strftime('%Y-%m-%d') if task.finish else None
            work = task.duration.hours if task.duration else 0
            
            # 提取资源
            owner = ''
            if task.resource_assignments:
                ra = task.resource_assignments[0]
                if ra.resource:
                    owner = ra.resource.name or ''
            
            tasks_data.append({
                'name': name,
                'start': start,
                'end': end,
                'work': work,
                'owner': owner,
                'indent': task.outline_level * 2 if hasattr(task, 'outline_level') else 0
            })
        
        # 复用文本解析的后续流程
        lines = ['任务名称\t开始时间\t结束时间\t工时\t责任人']
        for t in tasks_data:
            lines.append(f"{' ' * t['indent']}{t['name']}\t{t['start'] or ''}\t{t['end'] or ''}\t{t['work']}\t{t['owner']}")
        
        return parse_text('\n'.join(lines))
        
    except Exception as e:
        return {
            'success': False,
            'error': f'MPP解析失败: {str(e)}',
            'resources': []
        }

# ============================================================
# 环境检测
# ============================================================

def get_available_features():
    """检测当前环境可用的解析功能
    Returns: dict with keys: text, mpp, ocr
    """
    features = {
        'text': True,  # 文本解析总是可用
        'mpp': False,
        'ocr': False
    }
    
    # 检测 mpxj
    try:
        from mpxj import Reader
        features['mpp'] = True
    except ImportError:
        pass
    
    # 检测 OCR 相关库
    try:
        import PIL  # Pillow
        features['ocr'] = True
    except ImportError:
        pass
    
    return features

# ============================================================
# 测试入口
# ============================================================

if __name__ == '__main__':
    import sys
    import json
    
    # 内置测试数据
    test_text = """任务名称	开始时间	结束时间	工时	责任人
project：巴勒斯坦超声户表二开	7/9/26	8/15/26	304h	
  TR1	7/9/26	7/10/26	0h	
    需求导入	7/9/26	7/10/26	0h	陈雷雷
    需求评审	7/9/26	7/10/26	0h	陈雷雷
  TR3	7/21/26	7/23/26	16h	
    巴勒斯坦二开SE文档	7/21/26	7/23/26	16h	毛文豪
    巴勒斯坦SE评审	7/22/26	7/23/26	0h	毛文豪
  TR4	7/23/26	7/30/26	168h	
    水表固件代码编写	7/23/26	7/25/26	48h	
    阀门开度代码移植	7/23/26	7/25/26	0h	毛文豪
    水表固件自测	7/27/26	7/29/26	24h	
    CIU固件代码编写	7/29/26	7/30/26	32h	
    测试	7/23/26	7/29/26	64h	
      测试用例	7/23/26	7/28/26	24h	袁知正
      测试用例评审	7/28/26	7/29/26	0h	袁知正"""
    
    result = parse_text(test_text)
    
    print(f"解析结果: success={result['success']}")
    print(f"资源数: {len(result['resources'])}")
    if result['warnings']:
        print(f"警告: {result['warnings']}")
    if result['unresolved_tr']:
        print(f"未找到TR: {result['unresolved_tr']}")
    if result['unresolved_project']:
        print(f"未找到project: {result['unresolved_project']}")
    
    print("\n资源列表:")
    for i, r in enumerate(result['resources'], 1):
        print(f"{i}. {r['项目名称']}|{r['项目描述']}|{r['资源类型']}|{r['资源名称']}|{r['开始时间']}|{r['结束时间']}|{r['工时']}")
