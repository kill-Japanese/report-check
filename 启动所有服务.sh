#!/bin/bash
# ============================================================
# 超声波户表点检系统 - Mac/Linux 一键启动脚本
# 用法：chmod +x 启动所有服务.sh && ./启动所有服务.sh
# ============================================================

cd "$(dirname "$0")"

echo ""
echo "============================================"
echo "  超声波户表点检系统 - 启动中"
echo "============================================"
echo ""

# 1. 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "[错误] 未检测到 Python3，请先安装 Python 3.8+"
    exit 1
fi

# 2. 安装依赖（首次运行）
if [ ! -f ".deps_ok" ]; then
    echo "[1/4] 正在安装依赖..."
    pip3 install pandas openpyxl --break-system-packages -q
    echo "ok" > .deps_ok
    echo "  依赖安装完成"
else
    echo "[1/4] 依赖已就绪"
fi

# 3. 启动协作服务器
echo ""
echo "[2/4] 正在启动协作服务器 (端口 8080)..."
nohup python3 协作服务器_安全版.py 8080 > collab.log 2>&1 &
COLLAB_PID=$!
echo "  协作服务器已启动 (PID: $COLLAB_PID)"

# 4. 尝试启动 Cloudflare Tunnel（如果已安装）
echo ""
echo "[3/4] 检查 Cloudflare Tunnel..."
if command -v cloudflared &> /dev/null; then
    if [ -f "config.yml" ]; then
        nohup cloudflared tunnel --config config.yml run > cloudflare.log 2>&1 &
        CF_PID=$!
        echo "  Cloudflare Tunnel 已启动 (PID: $CF_PID)"
    else
        echo "  未检测到 config.yml，跳过 Cloudflare Tunnel"
        echo "  如需公网访问，请运行: cloudflared tunnel --url http://localhost:8080"
    fi
else
    echo "  未安装 cloudflared，跳过"
    echo "  安装: brew install cloudflared (Mac) 或参考文档"
fi

# 5. 设置定时任务
echo ""
echo "[4/4] 设置每日定时任务 (8:20)..."
SCRIPT_DIR=$(pwd)
CRON_CMD="20 8 * * * cd $SCRIPT_DIR && $(which python3) 更新点检表.py >> $SCRIPT_DIR/cron.log 2>&1"
(crontab -l 2>/dev/null | grep -v "更新点检表"; echo "$CRON_CMD") | crontab -
echo "  定时任务已设置"

# 完成
echo ""
echo "============================================"
echo "  启动完成！"
echo "============================================"
echo ""
echo "  本机访问: http://localhost:8080"
echo "  用户管理: http://localhost:8080/admin/users"
echo "  默认账号: admin / admin123"
echo ""
echo "  管理命令:"
echo "    查看日志: tail -f collab.log"
echo "    停止服务: kill $COLLAB_PID"
echo "    手动生成: python3 更新点检表.py"
echo ""

# 尝试打开浏览器
if command -v open &> /dev/null; then
    open http://localhost:8080
elif command -v xdg-open &> /dev/null; then
    xdg-open http://localhost:8080
fi
