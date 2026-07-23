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
    """从任务中提取资源（应用过滤规则）
    注意：按SKILL.MD规范，工时列直接存储日平均工时（h/天），不需要再做转换
    """
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
        
        # 按SKILL.MD规范：工时列直接就是日平均工时（h/天），不再计算
        # 日平均工时 = 工时列的值（四舍五入到1位小数）
        daily_hours = round(float(t['work']), 1) if t['work'] else 0
        
        resource = {
            '项目名称': t['project_name'] if t['project_name'] else '/',
            '项目描述': t['tr_name'] if t['tr_name'] else '/',
            '资源类型': t['name'],  # 任务名 = 资源类型
            '资源名称': res_name,
            '开始时间': t['start'] if t['start'] else '1900-01-01',
            '结束时间': t['end'] if t['end'] else '2100-01-01',
            '工时': t['work'],  # 日平均工时（与日平均工时字段值一致）
            '日平均工时': daily_hours,  # 直接等于工时列
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
    """合并设计+评审任务（同名资源，资源类型分别含设计和评审）
    约束：除了资源名称一致外，资源类型的基础名称（去掉设计/评审）也必须匹配
    例如：固件编码设计 + 固件编码评审 → 合并（基础名都是"固件编码"）
         需求评审 + SE文档设计 → 不合并（基础名不同）
    """
    if not resources:
        return resources
    
    def _get_base_name(type_name):
        """去掉资源类型中的"设计"或"评审"，得到基础名称"""
        base = type_name
        base = base.replace('设计', '').replace('评审', '')
        return base.strip()
    
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
        
        # 只有同时有设计和评审才尝试合并
        if design_items and review_items:
            # 按基础名称匹配（只有基础名称相同才合并）
            matched_reviews = set()
            for d_idx, design in design_items:
                d_base = _get_base_name(design['资源类型'])
                for r_idx, review in review_items:
                    if r_idx in matched_reviews:
                        continue
                    r_base = _get_base_name(review['资源类型'])
                    # 基础名称必须匹配（或一方为空另一方也为空的情况不合并）
                    if d_base and r_base and d_base == r_base:
                        # 合并规则
                        merged = dict(review)  # 以评审为基础（资源类型用评审的）
                        merged['开始时间'] = min(design['开始时间'], review['开始时间'])  # 取最早
                        merged['结束时间'] = max(design['结束时间'], review['结束时间'])  # 取最晚
                        merged['工时'] = max(design['工时'], review['工时'])  # 取较大的日平均工时
                        # 按SKILL.MD规范：日平均工时直接等于工时列（不再除以工作日数）
                        merged['日平均工时'] = merged['工时']
                        merged['warnings'] = list(set(design.get('warnings', []) + review.get('warnings', [])))
                        merged['is_unresolved_tr'] = design.get('is_unresolved_tr', False) or review.get('is_unresolved_tr', False)
                        merged['is_unresolved_project'] = design.get('is_unresolved_project', False) or review.get('is_unresolved_project', False)
                        
                        to_remove.add(d_idx)
                        to_remove.add(r_idx)
                        matched_reviews.add(r_idx)
                        merged_results.append((min(d_idx, r_idx), merged))
                        break
    
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
# PDF 解析（依赖 pdfplumber 库）
# ============================================================

def parse_pdf(pdf_path):
    """解析PDF格式的Project计划表
    
    支持 Microsoft Project 导出的PDF甘特图，从左侧任务表格提取数据
    
    Returns:
        dict: 同 parse_text 的返回格式
    """
    # 检测 pdfplumber
    try:
        import pdfplumber
    except ImportError:
        return {
            'success': False,
            'error': 'PDF解析库(pdfplumber)未安装，当前环境不支持PDF文件解析',
            'resources': []
        }
    
    def _is_gray_color(color):
        """判断颜色是否为灰色（RGB各分量接近相等，且亮度低于阈值）
        返回 True 表示是灰色/浅色文字，应被过滤
        """
        if not color:
            return False  # 无颜色信息，不过滤（默认黑色）
        try:
            # color 通常是元组，如 (0.7, 0.7, 0.7) 或 CMYK (0, 0, 0, 0.3)
            if isinstance(color, (list, tuple)):
                if len(color) == 3:
                    # RGB: 灰色 = R ≈ G ≈ B，且值不在黑色范围(0~0.3)内
                    r, g, b = color
                    if r < 0 and g < 0 and b < 0:
                        return False  # 负值通常是特殊标记
                    # 判断灰度：RGB差值都很小
                    diff = max(r, g, b) - min(r, g, b)
                    avg = (r + g + b) / 3
                    if diff < 0.15 and avg > 0.3:
                        return True  # 明显的灰色
                    # 很浅的颜色（接近白色）
                    if avg > 0.85:
                        return True
                    return False
                elif len(color) == 4:
                    # CMYK: 灰色 = C≈M≈Y≈0, K较高
                    c, m, y, k = color
                    if c < 0.1 and m < 0.1 and y < 0.1 and k > 0.15:
                        return True
                    return False
                elif len(color) == 1:
                    # 灰度值: 0=黑, 1=白
                    v = color[0]
                    if 0.3 < v < 1.0:
                        return True
                    return False
        except:
            pass
        return False
    
    def _build_gray_char_regions(page):
        """收集页面上所有灰色字体字符的坐标区域，返回用于后续过滤的矩形列表
        每个矩形用 (x0, y0, x1, y1) 表示，带少许扩展容差
        """
        gray_regions = []
        try:
            chars = page.chars
            if not chars:
                return gray_regions
            for ch in chars:
                color = ch.get('non_stroking_color')
                if _is_gray_color(color):
                    gray_regions.append((ch['x0'], ch['top'], ch['x1'], ch['bottom']))
        except Exception as e:
            print(f'[PDF] 收集灰色字符区域失败: {e}')
        return gray_regions
    
    try:
        all_tasks = []
        debug_info = []
        
        with pdfplumber.open(pdf_path) as pdf:
            debug_info.append(f'PDF共{len(pdf.pages)}页')
            for page_idx, page in enumerate(pdf.pages):
                # 【新增】收集灰色字体字符区域，用于过滤干扰文字
                gray_char_cache = _build_gray_char_regions(page)
                gray_filtered = False
                if gray_char_cache:
                    debug_info.append(f'第{page_idx+1}页: 检测到{len(gray_char_cache)}个灰色字符，尝试过滤')
                    # 过滤 page.chars：排除灰色字符
                    try:
                        original_chars = page.chars
                        filtered_chars = [
                            ch for ch in original_chars
                            if not _is_gray_color(ch.get('non_stroking_color'))
                        ]
                        if len(filtered_chars) < len(original_chars):
                            removed_count = len(original_chars) - len(filtered_chars)
                            debug_info.append(f'第{page_idx+1}页: 已过滤{removed_count}个灰色字符，剩余{len(filtered_chars)}个字符')
                            # 替换 page 的 chars 缓存，后续 extract_tables 将使用过滤后的字符
                            page._chars = filtered_chars
                            page.flush_cache()
                            gray_filtered = True
                        else:
                            debug_info.append(f'第{page_idx+1}页: 灰色字符过滤未生效（可能颜色格式不匹配）')
                    except Exception as e:
                        debug_info.append(f'第{page_idx+1}页: 灰色字符过滤失败: {e}')
                
                # 尝试多种表格提取策略
                all_tables_attempts = []
                
                # 策略1: 默认（竖线检测）- 优先
                try:
                    t1 = page.extract_tables()
                    all_tables_attempts.append(('默认', t1))
                except:
                    pass
                
                # 策略2: 文本位置检测（text strategy）
                try:
                    t2 = page.extract_tables(table_settings={
                        'vertical_strategy': 'text',
                        'horizontal_strategy': 'text',
                        'snap_tolerance': 5,
                    })
                    all_tables_attempts.append(('文本位置', t2))
                except:
                    pass
                
                # 策略3: 混合策略
                try:
                    t3 = page.extract_tables(table_settings={
                        'vertical_strategy': 'lines',
                        'horizontal_strategy': 'text',
                        'intersection_tolerance': 5,
                    })
                    all_tables_attempts.append(('混合', t3))
                except:
                    pass
                
                # 选最优策略：优先选能识别出任务表且列数在合理范围(8-20)的
                best_tables = []
                best_strategy = ''
                best_score = -1
                for strategy, tables in all_tables_attempts:
                    if not tables:
                        continue
                    for table in tables:
                        if len(table) < 2:
                            continue
                        first_row = [str(c).replace('\n', ' ').strip() if c else '' for c in table[0]]
                        is_task = any('标识号' in c or '任务名称' in c for c in first_row)
                        n_cols = len(first_row)
                        # 评分: 是任务表+100分, 列数在8-15之间额外加分
                        score = (100 if is_task else 0) + (10 if 8 <= n_cols <= 18 else 0) - abs(n_cols - 11)
                        if score > best_score:
                            best_score = score
                            best_tables = tables
                            best_strategy = strategy
                
                tables = best_tables if best_tables else page.extract_tables()
                debug_info.append(f'第{page_idx+1}页: 策略={best_strategy or "默认"}, {len(tables)}个表格')
                
                if not tables:
                    continue
                for t_idx, table in enumerate(tables):
                    if len(table) < 2:
                        continue
                    
                    # 【关键修复】检查是否是任务表格：表头可能在第1行或第2行（标题行占用第1行）
                    header_row_idx = -1
                    header_row = None
                    for check_idx in range(min(3, len(table))):
                        check_row = [str(c).replace('\n', ' ').strip() if c else '' for c in table[check_idx]]
                        is_task = any('标识号' in c or '任务名称' in c for c in check_row)
                        if is_task:
                            header_row_idx = check_idx
                            header_row = check_row
                            break
                    
                    if header_row_idx == -1:
                        continue
                    
                    debug_info.append(f'  表格{t_idx+1}: 识别为任务表(表头在第{header_row_idx+1}行)，共{len(table)}行, {len(header_row)}列, 表头={header_row}')
                    
                    # 找到列索引（用检测到的表头行）
                    col_idx = {}
                    for i, h in enumerate(header_row):
                        h_clean = h.replace('\n', '').strip()
                        if '标识号' in h_clean:
                            col_idx['id'] = i
                        elif '任务名称' in h_clean or 'Task Name' in h_clean:
                            col_idx['name'] = i
                        elif '工期' in h_clean and '工时' not in h_clean:
                            col_idx['duration'] = i
                        elif '开始时间' in h_clean or '开始' in h_clean:
                            col_idx['start'] = i
                        elif '完成时间' in h_clean or '结束时间' in h_clean or '完成' in h_clean:
                            col_idx['end'] = i
                        elif '交付物' in h_clean:
                            col_idx['deliverable'] = i
                        elif '责任人' in h_clean or '负责人' in h_clean or '资源名称' in h_clean:
                            col_idx['owner'] = i
                        elif '工时' in h_clean or '工作时间' in h_clean:
                            col_idx['work'] = i
                    
                    debug_info.append(f'  列映射: {col_idx}')
                    
                    # 必须有 id 和 name 列
                    if 'id' not in col_idx or 'name' not in col_idx:
                        debug_info.append(f'  ⚠ 缺少id或name列，跳过')
                        continue
                    
                    # 解析数据行（从表头行的下一行开始）
                    for row in table[header_row_idx + 1:]:
                        def _is_gantt_noise(s):
                            """判断字符串是否为纯甘特图噪音（无有效数据）"""
                            if not s or not s.strip():
                                return True
                            t = s.strip()
                            # 纯符号/数字/大写字母的组合，不含中文
                            if not re.search(r'[\u4e00-\u9fa5]', t):
                                # 检查是否是典型的甘特图标签模式
                                # 如: CCC ::, 888 66, ，，，, :::, 000 00, 222 6 44
                                if re.match(r'^[A-Z0-9，,.:：；;、\-_/\\|\s]+$', t):
                                    return True
                                # 重复字符3次以上: CCC, 888, ，，，
                                if re.match(r'^(.)\1{2,}', t):
                                    return True
                            return False

                        def _strip_gantt_inline(text):
                            """移除插入在文字中间的甘特图噪音
                            PDF甘特图标签可能与任务名称文字重叠，导致标签字符插入到名称中间。
                            例如：
                              固件概要设计000评审 → 固件概要设计评审
                              内测/单元888测试 → 内测/单元测试
                              可生产性评000估 → 可生产性评估
                              固件测试问题修CCC改 → 固件测试问题修改
                              生产工艺及规22 111程制定 → 生产工艺及规程制定
                              试流问题888整改 → 试流问题整改
                              T样机可生产00性 → T样机可生产性
                              硬件概要设计 ，，→ 硬件概要设计（末尾逗号）
                              硬件测试问题修改::: → 硬件测试问题修改（末尾冒号）
                              测试方案(000大纲）设计 → 测试方案(大纲）设计
                            """
                            if not text:
                                return text
                            result = text
                            
                            # 反复清理，直到没有更多变化
                            for _ in range(10):
                                new_result = result
                                
                                # 模式1: 纯重复数字(2+位)夹在中文之间 → 评000估 → 评估
                                # 如：评000估、修CCC改、888测试
                                new_result = re.sub(r'([\u4e00-\u9fa5])\d{2,}([\u4e00-\u9fa5])', r'\1\2', new_result)
                                # 也处理数字+空格+数字的变体：规22 111程 → 规程
                                new_result = re.sub(r'([\u4e00-\u9fa5])\s*\d{2,}\s*\d*\s*([\u4e00-\u9fa5])', r'\1\2', new_result)
                                # 更通用的：中文+数字/空格混合串+中文 → 移除中间的数字空格
                                new_result = re.sub(r'([\u4e00-\u9fa5])[\d\s]{2,}([\u4e00-\u9fa5])', r'\1\2', new_result)
                                
                                # 模式2: 纯重复大写字母(2+位)夹在中文之间 → 修CCC改
                                new_result = re.sub(r'([\u4e00-\u9fa5])[A-Z]{2,}([\u4e00-\u9fa5])', r'\1\2', new_result)
                                
                                # 模式3: 数字夹在中文和括号之间 → 测试方案(000大纲）→ 测试方案(大纲）
                                new_result = re.sub(r'([\(\（])\d{2,}([\u4e00-\u9fa5])', r'\1\2', new_result)
                                new_result = re.sub(r'([\u4e00-\u9fa5])\d{2,}([\)\）])', r'\1\2', new_result)
                                
                                # 模式4: 中文后面紧跟纯数字(2+位)再紧跟中文/括号 → 可生产00性 → 可生产性
                                # 注意：这会与模式1部分重叠，但模式1要求两端都是中文字符
                                # 这里扩展到：中文+数字+中文/左括号
                                new_result = re.sub(r'([\u4e00-\u9fa5])\d{2,}([\u4e00-\u9fa5（\(])', r'\1\2', new_result)
                                
                                if new_result == result:
                                    break
                                result = new_result
                            
                            return result.strip()
                        
                        def _strip_gantt_suffix(text):
                            """移除任务名称开头和末尾的甘特图噪音"""
                            if not text:
                                return text
                            result = text
                            # 反复移除末尾的甘特图标签（可能有多层）
                            for _ in range(5):
                                # 空格 + 纯大写字母(2+) + 可选符号: " CCC ::"
                                new_result = re.sub(r'\s+[A-Z]{2,}[\s:：；;、\-_/\\|]*$', '', result)
                                # 空格 + 纯数字(2+) + 可选符号: " 888 66"
                                new_result = re.sub(r'\s+\d{2,}[\s:：；;、\-_/\\|]*$', '', new_result)
                                # 末尾纯符号串: " ::"  " :::"
                                new_result = re.sub(r'\s+[:：；;、\-_/\\|]{2,}$', '', new_result)
                                # 末尾直接跟的纯数字(2+)：测试用例设计888
                                new_result = re.sub(r'\d{2,}$', '', new_result)
                                # 末尾直接跟的纯大写字母(2+)：文档CCC
                                new_result = re.sub(r'[A-Z]{2,}$', '', new_result)
                                # 末尾直接跟的中文标点符号(2+重复): ，，， :::
                                new_result = re.sub(r'[，,。.：:；;、]{2,}$', '', new_result)
                                if new_result == result:
                                    break
                                result = new_result
                            return result.strip()

                        def _clean_lines(raw_val, col_type='name'):
                            """清理单元格（处理甘特图标签混入的垃圾数据）
                            col_type: name(任务名称), date(日期), work(工时), owner(责任人), other
                            """
                            if raw_val is None:
                                return ''
                            s = str(raw_val)
                            if not s.strip():
                                return ''
                            
                            # 把换行符当作空格处理（统一处理所有行的混合内容）
                            s_normalized = re.sub(r'\n+', ' ', s).strip()
                            
                            if col_type == 'date':
                                # 日期：提取并修复（处理污染的年份和年月之间的垃圾字符）
                                def _extract_date(s):
                                    if not s:
                                        return None
                                    s = re.sub(r'\n+', ' ', s)
                                    
                                    # 方法1: 找 "月D日" 模式，然后往前找年份
                                    # 同时处理年月之间有垃圾字符的情况（如2026年CCC7月23日、2026888年7月31日）
                                    def _fix_year(y_str):
                                        """从污染的年份字符串中提取正确的4位年份"""
                                        if not y_str:
                                            return '2026'
                                        y = y_str.strip()
                                        # 已经是正确的4位年份
                                        if len(y) == 4 and y.isdigit() and 1900 <= int(y) <= 2100:
                                            return y
                                        # 从数字串中找合理的4位年份（1900-2100）
                                        digits = re.findall(r'\d+', y)
                                        for d in digits:
                                            # 找4位合理年份
                                            for i in range(len(d) - 3):
                                                four = d[i:i+4]
                                                if 1900 <= int(four) <= 2100:
                                                    return four
                                        # 找不到合理年份，取最后2位加20
                                        all_nums = ''.join(digits)
                                        if len(all_nums) >= 2:
                                            return '20' + all_nums[-2:]
                                        return '2026'
                                    
                                    m = re.search(r'(\d{1,2})\s*月\s*(\d{1,2})\s*日', s)
                                    if m:
                                        month, day = m.group(1), m.group(2)
                                        # 从月份位置往前，找"年"字，再往前找数字
                                        pre = s[:m.start(1)]
                                        year_m = re.search(r'(\d+)\s*年\s*$', pre)
                                        if year_m:
                                            y = _fix_year(year_m.group(1))
                                            return f'{y}年{month}月{day}日'
                                        # 没找到年字，往前找最近的数字串
                                        digits = re.findall(r'(\d+)', pre)
                                        if digits:
                                            # 从后往前找合理的年份
                                            for d in reversed(digits):
                                                fixed = _fix_year(d)
                                                if fixed != '2026' or len(d) >= 2:
                                                    return f'{fixed}年{month}月{day}日'
                                            y = _fix_year(digits[-1])
                                            return f'{y}年{month}月{day}日'
                                    
                                    # 方法2: 标准格式
                                    m = re.search(r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s)
                                    if m:
                                        return f'{m.group(1)}-{m.group(2)}-{m.group(3)}'
                                    return None
                                
                                extracted = _extract_date(s_normalized)
                                if extracted:
                                    return extracted
                                return s_normalized
                            
                            elif col_type == 'work':
                                # 工时：优先取中间的合理数字（甘特图标签通常在两端）
                                # 如 222 6 44 → 6是工时，00 2 888 → 2是工时
                                def _is_gantt_num(num_str):
                                    """判断数字是否是甘特图标签（重复数字如00, 11, 44, 888）"""
                                    if re.match(r'^(\d)\1{1,}$', num_str):
                                        return True
                                    return False
                                
                                if '\n' in s:
                                    lines = [l.strip() for l in s.split('\n') if l.strip()]
                                    # 从最后一行（通常是真实数据行）往前找
                                    for l in reversed(lines):
                                        nums = re.findall(r'([\d.]+)', l)
                                        valid_nums = []
                                        for n in nums:
                                            try:
                                                val = float(n)
                                                if 0 <= val <= 200:
                                                    valid_nums.append(n)
                                            except:
                                                pass
                                        if valid_nums:
                                            # 过滤掉甘特图标签（重复数字）
                                            real_nums = [n for n in valid_nums if not _is_gantt_num(n)]
                                            if not real_nums:
                                                real_nums = valid_nums  # 全是标签就用原值
                                            # 优先取中间值（两端是甘特图标签）
                                            if len(real_nums) >= 3:
                                                return real_nums[len(real_nums)//2]
                                            elif len(real_nums) == 2:
                                                # 取非标签的那个，如果都不是标签取第一个非零或最后一个
                                                g0 = _is_gantt_num(real_nums[0])
                                                g1 = _is_gantt_num(real_nums[1])
                                                if g0 and not g1:
                                                    return real_nums[1]
                                                elif g1 and not g0:
                                                    return real_nums[0]
                                                else:
                                                    # 0小时也是合理工时，不再因为值为0就回退
                                                    return real_nums[0]
                                            else:
                                                return real_nums[0]
                                # 没换行的情况
                                nums = re.findall(r'([\d.]+)', s_normalized)
                                valid_nums = []
                                for n in nums:
                                    try:
                                        val = float(n)
                                        if 0 <= val <= 200:
                                            valid_nums.append((n, val))
                                    except:
                                        pass
                                if valid_nums:
                                    # 优先取中间值，排除明显的甘特图标签（重复数字如00, 888, 666）
                                    def _is_gantt_label(num_str, val):
                                        # 重复数字2次以上：00, 888, 66, 222
                                        if re.match(r'^(\d)\1{1,}$', num_str):
                                            return True
                                        return False
                                    
                                    real_values = [(n, v) for n, v in valid_nums if not _is_gantt_label(n, v)]
                                    if real_values:
                                        if len(real_values) >= 3:
                                            return real_values[len(real_values)//2][0]
                                        elif len(real_values) == 2:
                                            return real_values[1][0]
                                        else:
                                            return real_values[0][0]
                                    # 全是标签，取第一个有效值（0也是合理工时）
                                    if valid_nums:
                                        return valid_nums[0][0]
                                    return valid_nums[-1][0]
                                if nums:
                                    return nums[-1]
                                return ''
                            
                            elif col_type == 'owner':
                                # 先判断是否纯甘特图噪音
                                if _is_gantt_noise(s_normalized):
                                    return ''
                                # 责任人：优先从已知人名列表提取（支持模糊匹配）
                                known_names = ['陈雷雷', '毛文豪', '袁知正', '文春英', '肖庆杨', '文善英']
                                # 精确子串匹配
                                for n in known_names:
                                    if n in s_normalized:
                                        return n
                                # 按字序模糊匹配
                                for n in known_names:
                                    idx = 0
                                    match = True
                                    for ch in n:
                                        pos = s_normalized.find(ch, idx)
                                        if pos == -1:
                                            match = False
                                            break
                                        idx = pos + 1
                                    if match:
                                        return n
                                # 没有已知人名，找2-4个中文字
                                m = re.search(r'[\u4e00-\u9fa5]{2,4}', s_normalized)
                                if m:
                                    return m.group(0)
                                # 如果看起来像噪音（无中文，纯符号数字），返回空
                                if _is_gantt_noise(s_normalized):
                                    return ''
                                return s_normalized.strip()
                            
                            elif col_type == 'name':
                                # 任务名称：
                                # 1. 如果有换行，取包含中文/TR/project的最长行
                                result = s_normalized
                                if '\n' in s:
                                    lines = [l.strip() for l in s.split('\n') if l.strip()]
                                    candidates = []
                                    for l in lines:
                                        if re.match(r'^[\d，,.:：；;、\-_/\\|]+$', l):
                                            continue
                                        if len(re.findall(r'[\u4e00-\u9fa5]', l)) > 0 or re.match(r'^TR[\dA-Za-z]*', l, re.I) or re.match(r'^project[:：]', l, re.I):
                                            candidates.append(l)
                                    if candidates:
                                        result = max(candidates, key=len)
                                # 2. 去掉开头的纯数字/符号垃圾前缀（，，，、888、00等）
                                result = re.sub(r'^[\d，,.:：；;、\-_/\\|\s]+', '', result)
                                # 3. 用专用函数移除末尾甘特图噪音（CCC ::、888 66等）
                                result = _strip_gantt_suffix(result)
                                # 4. 用专用函数移除文字中间的甘特图噪音（000、888、CCC等）
                                result = _strip_gantt_inline(result)
                                # 5. TR行特殊处理：只保留 TR+数字/字母
                                if re.match(r'^TR[\dA-Za-z]*', result, re.I):
                                    m = re.match(r'^(TR[\dA-Za-z]*)', result, re.I)
                                    if m:
                                        result = m.group(1)
                                return result.strip()
                            
                            else:
                                return s_normalized
                        
                        def get_cell(idx, col_type='other'):
                            if idx is None or idx >= len(row):
                                return ''
                            return _clean_lines(row[idx], col_type)
                        
                        task_id = get_cell(col_idx.get('id'))
                        if not task_id or not re.match(r'^\d+$', task_id):
                            continue
                        
                        name = get_cell(col_idx.get('name'), 'name')
                        if not name:
                            continue
                        
                        duration = get_cell(col_idx.get('duration'))
                        start_raw = get_cell(col_idx.get('start'), 'date')
                        end_raw = get_cell(col_idx.get('end'), 'date')
                        deliverable = get_cell(col_idx.get('deliverable'))
                        owner = get_cell(col_idx.get('owner'), 'owner')
                        work_raw = get_cell(col_idx.get('work'), 'work')
                        
                        # 清理责任人：从混合文本中提取真实人名
                        known_names = ['陈雷雷', '毛文豪', '袁知正', '文春英', '肖庆杨']
                        
                        def _extract_owner(s):
                            """从字符串中提取已知人名（支持模糊匹配）"""
                            if not s:
                                return ''
                            # 精确子串匹配
                            for n in known_names:
                                if n in s:
                                    return n
                            # 模糊匹配: 按字序匹配（处理"毛文文档豪"→"毛文豪"的情况）
                            for n in known_names:
                                idx = 0
                                match = True
                                for ch in n:
                                    pos = s.find(ch, idx)
                                    if pos == -1:
                                        match = False
                                        break
                                    idx = pos + 1
                                if match:
                                    return n
                            return ''
                        
                        owner_clean = _extract_owner(owner)
                        # 如果交付物列包含人名，也提取
                        if not owner_clean:
                            owner_clean = _extract_owner(deliverable)
                            if owner_clean:
                                deliverable = deliverable.replace(owner_clean, '').strip('，, ')
                        
                        # 如果从责任人列中提取到了人名，但原始内容比人名长
                        # 说明有交付物内容被错误地划到了责任人列，需要还回去
                        if owner_clean and owner and owner != owner_clean:
                            # 从owner中移除匹配到的人名（按字序移除）
                            leftover = owner
                            for ch in owner_clean:
                                pos = leftover.find(ch)
                                if pos != -1:
                                    leftover = leftover[:pos] + leftover[pos+1:]
                            leftover = leftover.strip('，, ')
                            if leftover:
                                deliverable = (deliverable + leftover) if deliverable else leftover
                        
                        # 如果责任人包含"报告"、"文档"等交付物关键词，清理掉
                        if not owner_clean and owner:
                            if any(kw in owner for kw in ['报告', '文档', '配置', '设计', '评审', '产物', '申请']):
                                deliverable = (deliverable + '，' + owner).strip('，') if deliverable else owner
                                owner_clean = ''
                            else:
                                owner_clean = owner
                        else:
                            owner_clean = owner_clean or owner
                        
                        # 解析日期和工时
                        start = _parse_date(start_raw)
                        end = _parse_date(end_raw)
                        work = _parse_work_hours(work_raw)
                        
                        all_tasks.append({
                            'id': int(task_id),
                            'name': name,
                            'duration': duration,
                            'start': start,
                            'end': end,
                            'deliverable': deliverable,
                            'owner': owner_clean,
                            'work': work,
                        })
        
        debug_info.append(f'✅ PDF提取任务数: {len(all_tasks)}')
        
        if not all_tasks:
            return {
                'success': False,
                'error': 'PDF中未找到有效的任务表格，请确认是Microsoft Project导出的PDF',
                'resources': [],
                'warnings': debug_info
            }
        
        # 注意：不按ID排序，保持PDF原文档的行顺序
        
        # 检查是否有显式project行，没有的话把首个顶级任务标记为project
        has_project = any(re.match(r'^project[:：]', t['name'], re.I) for t in all_tasks)
        if not has_project and all_tasks:
            all_tasks[0]['name'] = 'project: ' + all_tasks[0]['name']
            debug_info.append(f'⚠ 未检测到project标识，将首个任务标记为project: {all_tasks[0]["name"]}')
        
        # 推断缩进层级
        tasks_with_indent = []
        for i, t in enumerate(all_tasks):
            name = t['name']
            if re.match(r'^project[:：]', name, re.I):
                indent = 0
            elif re.match(r'^TR[\dA-Z]*', name, re.I):
                indent = 2
            else:
                # 子任务：找最近的TR或project祖先
                indent = 4
                for j in range(i-1, -1, -1):
                    prev = all_tasks[j]
                    if re.match(r'^TR[\dA-Z]*', prev['name'], re.I):
                        indent = 4
                        break
                    elif re.match(r'^project[:：]', prev['name'], re.I):
                        indent = 2
                        break
            tasks_with_indent.append((t, indent))
        
        # 构造文本格式，复用 parse_text 的后续流程
        lines = ['任务名称\t开始时间\t结束时间\t工时\t责任人']
        for t, indent in tasks_with_indent:
            padded_name = ' ' * indent + t['name']
            lines.append(f"{padded_name}\t{t['start'] or ''}\t{t['end'] or ''}\t{t['work']}h\t{t['owner']}")
        
        text_content = '\n'.join(lines)
        debug_info.append(f'构造文本共{len(lines)}行, 前5行预览:')
        for l in lines[:5]:
            debug_info.append(f'  {l[:80]}')
        
        result = parse_text(text_content)
        # 把PDF调试信息合并到结果中
        if 'warnings' not in result:
            result['warnings'] = []
        result['warnings'] = debug_info + result['warnings']
        return result
        
    except Exception as e:
        return {
            'success': False,
            'error': f'PDF解析失败: {str(e)}',
            'resources': [],
            'warnings': debug_info if 'debug_info' in dir() else [str(e)]
        }

# ============================================================
# 环境检测
# ============================================================

def get_available_features():
    """检测当前环境可用的解析功能
    Returns: dict with keys: text, mpp, pdf, ocr
    """
    features = {
        'text': True,  # 文本解析总是可用
        'mpp': False,
        'pdf': False,
        'ocr': False
    }
    
    # 检测 mpxj
    try:
        from mpxj import Reader
        features['mpp'] = True
    except ImportError:
        pass
    
    # 检测 pdfplumber
    try:
        import pdfplumber
        features['pdf'] = True
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
