#!/bin/bash
# 每日8:20定时生成超声波户表项目延期点检报表

cd /workspace

# 日志文件
LOG_FILE="/workspace/.trae/report_cron.log"
mkdir -p /workspace/.trae

echo "========================================" >> "$LOG_FILE"
echo "[$(date '+%Y-%m-%d %H:%M:%S')] 开始生成点检报表..." >> "$LOG_FILE"

# 优先使用 data_source.json 中记录的数据源，否则自动查找
DATA_SOURCE=""
if [ -f "/workspace/.trae/data_source.json" ]; then
    DATA_SOURCE=$(python3 -c "
import json, os
try:
    with open('/workspace/.trae/data_source.json') as f:
        info = json.load(f)
    fp = info.get('file_path', '')
    # 如果文件不存在，尝试在 .uploads 目录找
    if fp and os.path.exists(fp):
        print(fp)
    else:
        import glob
        uploads = glob.glob('/workspace/.uploads/*.xlsx')
        script_files = [f for f in uploads if '超声波户表脚本' in f]
        if script_files:
            print(max(script_files, key=os.path.getmtime))
        elif uploads:
            print(max(uploads, key=os.path.getmtime))
except:
    pass
" 2>/dev/null)
fi

# 如果还没找到，自动查找
if [ -z "$DATA_SOURCE" ] || [ ! -f "$DATA_SOURCE" ]; then
    DATA_SOURCE=$(ls -t /workspace/.uploads/*超声波户表脚本*.xlsx 2>/dev/null | head -1)
fi

if [ -z "$DATA_SOURCE" ] || [ ! -f "$DATA_SOURCE" ]; then
    DATA_SOURCE=$(ls -t /workspace/.uploads/*.xlsx 2>/dev/null | head -1)
fi

echo "[$(date '+%Y-%m-%d %H:%M:%S')] 数据源: $DATA_SOURCE" >> "$LOG_FILE"

# 运行报表生成
if [ -n "$DATA_SOURCE" ] && [ -f "$DATA_SOURCE" ]; then
    python3 /data/user/skills/chaoshengbo-report/更新点检表.py "$DATA_SOURCE" >> "$LOG_FILE" 2>&1
    
    # 确保文件在 workspace（脚本可能输出到当前目录）
    if [ -f "/workspace/项目延期点检表.html" ]; then
        echo "[$(date '+%Y-%m-%d %H:%M:%S')] ✅ 报表生成成功" >> "$LOG_FILE"
    else
        # 尝试从 /data 复制
        if [ -f "/data/项目延期点检表.html" ]; then
            cp /data/项目延期点检表.html /workspace/ 2>/dev/null
            cp /data/项目延期点检表.xlsx /workspace/ 2>/dev/null
            echo "[$(date '+%Y-%m-%d %H:%M:%S')] ✅ 报表已复制到工作区" >> "$LOG_FILE"
        else
            echo "[$(date '+%Y-%m-%d %H:%M:%S')] ❌ 报表生成失败" >> "$LOG_FILE"
        fi
    fi
else
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] ❌ 未找到数据源文件" >> "$LOG_FILE"
fi

echo "========================================" >> "$LOG_FILE"
