# feature/project-import 分支回退说明

## 主干备份信息
- **主干分支**: main
- **备份 Hash**: `e8f0ca7a14b8e563f85adb50c91453cbc67c8889`
- **备份日期**: 2026-07-22

## 当前分支
- **分支名**: feature/project-import
- **最新 Commit**: `3fcba0d` - feat: 新增Project文档批量导入功能

## 改动文件清单

### 新增文件
| 文件 | 说明 |
|------|------|
| `project_parser.py` | Project文档解析模块（文本/MPP/OCR 三种方式） |
| `feature/project-import/main_backup_hash.txt` | 主干备份hash |

### 修改文件
| 文件 | 改动行数 | 说明 |
|------|---------|------|
| `sync_excel.py` | +220行 | 新增3个辅助函数(_normalize_date, _copy_formulas, _format_resource_name) + action_add_project_batch() + 修复action_add_project() |
| `协作服务器_安全版.py` | +85行 | 新增3个API: /api/import/features, /api/import/parse, /api/import/commit |
| `更新点检表.py` | +600行 | 新增导入按钮 + 导入模态框 + 8个JS函数 |

## 功能说明

本分支新增了从 Microsoft Project 文档批量导入项目的功能：

1. **三种输入方式**：
   - 文本粘贴（Tab/空格分隔的Project表格）
   - .mpp 文件上传（需 mpxj 库，Render环境可能不可用）
   - 截图OCR（暂不可用）

2. **智能解析**：
   - 自动识别 project: 和 TR 标识
   - 跨层级向上查找关联
   - 设计+评审任务自动合并（同名资源）
   - 日平均工时自动计算

3. **预览编辑**：
   - 解析后展示可编辑的预览表格
   - 支持修改字段、删除行
   - TR/project 标识缺失警告

4. **批量写入**：
   - 一次打开Excel写入所有行
   - O-W列公式自动复制
   - 资源名称自动转换为Excel格式
   - G/H列（项目开始/结束）按项目分组自动计算
   - 7道写入前校验 + 50条数量限制

## 回退步骤

### 方式一：切回主干（推荐）
```bash
git checkout main
```

### 方式二：从备份hash恢复
```bash
git checkout e8f0ca7a14b8e563f85adb50c91453cbc67c8889
```

### 方式三：删除本分支（彻底清除）
```bash
git checkout main
git branch -D feature/project-import
git push origin --delete feature/project-import
```

## 合并到主干（用户确认后执行）

```bash
git checkout main
git merge feature/project-import
git push origin main
```
