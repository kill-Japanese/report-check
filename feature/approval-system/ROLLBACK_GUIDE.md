# feature/approval-system 分支回退说明

## 上一分支备份信息
- **来源分支**: feature/project-import
- **起点 Commit Hash**: `bc3d16c`
- **起点日期**: 2026-07-23
- **主干分支**: main
- **主干备份 Hash**: `e8f0ca7a14b8e563f85adb50c91453cbc67c8889`

## 当前分支
- **分支名**: feature/approval-system
- **最新 Commit**: (开发中)

## 功能说明

本分支新增**归档、编辑审批系统**：

### 核心变更

1. **权限体系调整**：
   - 新增 `submit_approval` 权限（所有登录用户）
   - 新增 `approve` 审批权限（admin + editor）
   - 新增 `delete` 删除权限（仅 admin）

2. **审批流程**：
   - viewer：归档/编辑 → 提交审批 → editor/admin 审批 → 生效
   - editor：归档/编辑 → 直接生效；可审批 viewer 的申请
   - admin：归档/编辑/删除 → 直接生效；可审批；可管理用户

3. **数据存储**：
   - X列（第24列）：审批状态标识（PENDING_ARCHIVE / PENDING_UNARCHIVE / PENDING_EDIT）
   - 「操作记录」Sheet（超声波户表脚本.xlsx新增）：所有操作日志（归档、编辑、删除、审批）

4. **前端变更**：
   - 角色感知的操作按钮（viewer走审批、editor/admin直接生效）
   - 审批面板（待我审批 + 我发起的）
   - 红点提示待审批数量
   - 删除按钮仅 admin 可见

### 改动文件清单

| 文件 | 改动类型 | 说明 |
|------|---------|------|
| `auth.py` | 修改 | 新增3个权限项 + 3个审计事件 |
| `sync_excel.py` | 修改 | X列审批状态 + 操作记录Sheet + 14个新函数 |
| `协作服务器_安全版.py` | 修改 | 7个新API（审批提交/列表/通过/拒绝/计数 + 操作记录） |
| `更新点检表.py` | 修改 | 角色感知按钮 + 审批面板 + 编辑字段限制 |

## 回退步骤

### 方式一：切回上一分支（推荐，保留审批功能历史）
```bash
git checkout feature/project-import
```

### 方式二：从起点hash恢复
```bash
git checkout bc3d16c
```

### 方式三：删除本分支（彻底清除审批功能）
```bash
git checkout feature/project-import
git branch -D feature/approval-system
git push origin --delete feature/approval-system
```

### 方式四：回退到主干（无任何功能分支）
```bash
git checkout main
```

## 合并到上一分支（用户确认后执行）

```bash
git checkout feature/project-import
git merge feature/approval-system
git push origin feature/project-import
```

## 合并到主干（最终确认后执行）

```bash
git checkout main
git merge feature/approval-system
git push origin main
```
