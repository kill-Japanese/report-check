# feature/approval-frontend 分支回退说明

## 上一分支备份信息
- **来源分支**: feature/approval-system
- **起点 Commit Hash**: `482ddd8`
- **起点日期**: 2026-07-23
- **主干分支**: main
- **主干备份 Hash**: `e8f0ca7a14b8e563f85adb50c91453cbc67c8889`

## 当前分支
- **分支名**: feature/approval-frontend
- **最新 Commit**: (开发中)

## 功能说明

本分支在 `feature/approval-system`（后端审批系统基础上，完成**前端审批面板**开发：

### 核心变更

1. **角色感知按钮**：
   - viewer：归档/编辑按钮点击后走审批流程（提交申请）
   - editor/admin：归档/编辑直接生效
   - 删除按钮仅 admin 可见

2. **审批面板**：
   - 待我审批列表（editor/admin 可见）
   - 我发起的审批列表
   - 审批通过/拒绝操作
   - 红点提示待审批数量

3. **编辑字段限制**：
   - 允许编辑：负责人、开始时间、结束时间、资源类型、工时
   - 其他字段只读

4. **审批状态展示**：
   - 项目列表中显示审批状态标签（待归档/已通过/已拒绝）

### 改动文件清单

| 文件 | 改动类型 | 说明 |
|------|---------|------|
| `更新点检表.py` | 修改 | 角色感知按钮 + 审批面板 + 编辑字段限制 + 审批状态展示 |

## 回退步骤

### 方式一：切回上一分支（推荐，保留前端审批功能历史）
```bash
git checkout feature/approval-system
```

### 方式二：从起点hash恢复
```bash
git checkout 482ddd8
```

### 方式三：删除本分支（彻底清除前端审批功能）
```bash
git checkout feature/approval-system
git branch -D feature/approval-frontend
git push origin --delete feature/approval-frontend
```

### 方式四：回退到主干（无任何功能分支）
```bash
git checkout main
```

## 合并到上一分支（用户确认后执行）

```bash
git checkout feature/approval-system
git merge feature/approval-frontend
git push origin feature/approval-system
```

## 合并到主干（最终确认后执行）

```bash
git checkout main
git merge feature/approval-frontend
git push origin main
```
