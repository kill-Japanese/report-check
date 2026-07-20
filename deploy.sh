#!/bin/bash
# ============================================================
# 超声波户表点检系统 - 一键部署脚本
# 适用：Ubuntu 22.04+ / Debian 12+
# 用法：chmod +x deploy.sh && ./deploy.sh
# ============================================================

set -e

echo ""
echo "============================================"
echo "🚀 超声波户表点检系统 - 一键部署"
echo "============================================"

# 检测是否为 root
if [ "$EUID" -ne 0 ]; then
    SUDO="sudo"
else
    SUDO=""
fi

# 1. 安装基础依赖
echo ""
echo "[1/6] 安装系统依赖..."
$SUDO apt update -qq
$SUDO apt install -y -qq python3-pip python3-venv git cron nginx curl
echo "✅ 系统依赖安装完成"

# 2. 安装 Python 依赖
echo ""
echo "[2/6] 安装 Python 依赖..."
pip3 install pandas openpyxl --break-system-packages -q
echo "✅ Python 依赖安装完成"

# 3. 创建数据目录
echo ""
echo "[3/6] 初始化目录..."
mkdir -p .uploads .trae
touch cron.log collab.log mail.log
echo "✅ 目录初始化完成"

# 4. 设置定时任务（每天北京时间 8:20 = UTC 0:20）
echo ""
echo "[4/6] 设置定时任务..."
SCRIPT_DIR=$(pwd)
PYTHON_BIN=$(which python3)
CRON_CMD="20 0 * * * cd $SCRIPT_DIR && $PYTHON_BIN 更新点检表.py >> $SCRIPT_DIR/cron.log 2>&1"
(crontab -l 2>/dev/null | grep -v "更新点检表"; echo "$CRON_CMD") | crontab -
echo "✅ 定时任务已设置（每天 8:20 自动生成报表）"

# 5. 配置协作服务为 systemd 服务
echo ""
echo "[5/6] 配置协作服务..."
CURRENT_USER=$(whoami)
$SUDO tee /etc/systemd/system/report-collab.service > /dev/null <<EOF
[Unit]
Description=项目点检表协作服务器（安全版）
After=network.target

[Service]
User=$CURRENT_USER
WorkingDirectory=$SCRIPT_DIR
ExecStart=$PYTHON_BIN $SCRIPT_DIR/协作服务器_安全版.py 8080
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
EOF

$SUDO systemctl daemon-reload
$SUDO systemctl enable --now report-collab
echo "✅ 协作服务已启动（端口 8080）"

# 6. 配置 Nginx 反向代理
echo ""
echo "[6/6] 配置 Nginx..."
$SUDO tee /etc/nginx/sites-available/report > /dev/null <<'EOF'
server {
    listen 80 default_server;
    
    client_max_body_size 50M;
    
    location / {
        proxy_pass http://127.0.0.1:8080;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
        proxy_read_timeout 300s;
    }
}
EOF

$SUDO rm -f /etc/nginx/sites-enabled/default
$SUDO ln -sf /etc/nginx/sites-available/report /etc/nginx/sites-enabled/
$SUDO nginx -t -q && $SUDO systemctl restart nginx
echo "✅ Nginx 配置完成"

# 输出结果
SERVER_IP=$(curl -s ifconfig.me 2>/dev/null || curl -s ipinfo.io/ip 2>/dev/null || echo "你的服务器IP")
echo ""
echo "============================================"
echo "🎉 部署完成！"
echo "============================================"
echo ""
echo "📡 访问地址："
echo "   本机:   http://localhost:8080"
echo "   公网:   http://$SERVER_IP"
echo ""
echo "⏰ 定时任务："
echo "   每天 8:20 (北京时间) 自动生成报表"
echo ""
echo "📋 管理命令："
echo "   查看协作日志:  tail -f $SCRIPT_DIR/collab.log"
echo "   查看定时日志:  tail -f $SCRIPT_DIR/cron.log"
echo "   重启协作服务:  $SUDO systemctl restart report-collab"
echo "   手动生成报表:  python3 更新点检表.py"
echo "   用户管理页:    http://$SERVER_IP/admin/users"
echo "   默认账号:      admin / admin123"
echo ""
echo "💡 提示：如需 HTTPS，请配置域名后运行："
echo "   $SUDO apt install -y certbot python3-certbot-nginx"
echo "   $SUDO certbot --nginx -d 你的域名"
echo "============================================"
