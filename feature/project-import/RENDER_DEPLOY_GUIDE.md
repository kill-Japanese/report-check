# Render 部署分支切换操作指南

## 重要说明

**Render 不会自动切换分支。**

您的服务配置了 `autoDeployTrigger: 'off'`（已关闭自动部署），所以：
- 即使 GitHub 上有新的分支或提交，Render 也不会自动部署
- 每次部署都需要您手动操作
- 默认部署的是 `main` 分支

---

## 方式一：在 Render 控制台切换分支（推荐用于测试）

如果您想在 Render 上直接测试 `feature/project-import` 分支，按以下步骤操作：

### 步骤 1：进入 Render 控制台

1. 打开 [https://dashboard.render.com](https://dashboard.render.com)
2. 登录您的账号
3. 找到您的服务 `report-check` 并点击进入

### 步骤 2：切换部署分支

1. 在服务详情页，点击顶部的 **Settings**（设置）选项卡
2. 找到 **Build & Deploy**（构建和部署）区域
3. 找到 **Branch**（分支）选项
4. 点击下拉框，选择 `feature/project-import`
5. 点击 **Save Changes**（保存更改）

### 步骤 3：手动触发部署

1. 回到服务详情页的 **Events**（事件）或 **Logs**（日志）选项卡
2. 点击右上角的 **Manual Deploy**（手动部署）按钮
3. 选择 **Deploy latest commit**（部署最新提交）
4. 等待部署完成（通常 2-5 分钟）

### 步骤 4：验证功能

部署完成后，访问您的 Render 服务地址，测试：
- 是否出现「📥 从Project导入」按钮
- 文本粘贴解析是否正常
- 批量导入是否正常写入

### 切回 main 分支

测试完成后，如果要恢复主干版本：
1. 重复步骤 2，将分支切回 `main`
2. 重复步骤 3，点击 Manual Deploy

---

## 方式二：合并到 main 后再部署（推荐用于正式上线）

如果您确认功能没问题，想正式上线，按以下步骤操作：

### 步骤 1：在本地合并分支

在 TRAE 中或您的本地终端执行：

```bash
# 切换到 main 分支
git checkout main

# 合并 feature 分支
git merge feature/project-import

# 推送到 GitHub
git push origin main
```

### 步骤 2：在 Render 上部署

1. 打开 [Render 控制台](https://dashboard.render.com)
2. 进入 `report-check` 服务
3. 确保 Settings → Build & Deploy → Branch 是 `main`
4. 点击 **Manual Deploy** → **Deploy latest commit**

---

## 方式三：创建 Preview Environment（高级，可选）

如果您想同时保留 main 的稳定版本和 feature 的测试版本，可以创建一个 Pull Request：

### 步骤

1. 在 GitHub 上创建一个 Pull Request：`feature/project-import` → `main`
2. Render 会自动检测到 PR 并创建一个 **Preview Environment**（预览环境）
3. 预览环境有独立的 URL（类似 `https://report-check-pr-xx.onrender.com`）
4. 在预览环境测试功能
5. 测试通过后，在 GitHub 上合并 PR
6. 合并后，手动部署 main 分支

### 注意事项
- Preview Environment 会消耗 Render 的资源配额
- Preview Environment 的数据（Excel文件）是独立的
- 如果您的 Render 是免费套餐，可能不支持 Preview Environment

---

## 常见问题

### Q1: 切换分支后 Render 上的数据会丢失吗？

**不会丢失。** Render 的磁盘存储是持久化的，切换分支只会重新部署代码，不会影响您的 Excel 数据文件。

### Q2: 部署需要多长时间？

通常 **2-5 分钟**。主要耗时在：
- `apt-get update && apt-get install -y git`（安装依赖）
- `pip install -r requirements.txt`（安装 Python 包）
- 启动服务

### Q3: 如何查看部署日志？

1. 进入 Render 控制台 → 服务详情页
2. 点击 **Logs**（日志）选项卡
3. 可以看到实时的构建和运行日志

### Q4: 部署失败了怎么办？

1. 查看 **Logs** 中的错误信息
2. 常见问题：
   - 依赖安装失败 → 检查 `requirements.txt`
   - 端口错误 → 确认启动命令是 `python 协作服务器_安全版.py 10000`
   - 内存不足 → Render 免费套餐只有 512MB，大文件处理可能 OOM
3. 如果无法解决，可以 **Rollback**（回滚）到上一个成功的部署

### Q5: 如何回滚到上一个版本？

1. 进入服务详情页 → **Events**（事件）选项卡
2. 找到上一个成功的部署事件
3. 点击 **Deploy**（部署）按钮
4. 或者在 Settings 中将分支切回 `main` 并重新部署

---

## 当前分支信息速查

| 项目 | 值 |
|------|-----|
| 主干分支 | `main` |
| 主干备份 Hash | `e8f0ca7a14b8e563f85adb50c91453cbc67c8889` |
| 新功能分支 | `feature/project-import` |
| 新功能最新提交 | `3fcba0d` |
| Render 自动部署 | ❌ 已关闭（需手动） |
| Render 服务名 | `report-check` |
