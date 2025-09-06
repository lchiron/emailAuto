#!/bin/bash

# ServiceNow邮件自动批准程序启动脚本

# 获取脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" &> /dev/null && pwd)"

# 进入程序目录
cd "$SCRIPT_DIR"

# 检查Python环境
if command -v python3 &> /dev/null; then
    PYTHON_CMD="python3"
elif command -v python &> /dev/null; then
    PYTHON_CMD="python"
else
    echo "错误: 未找到Python环境"
    exit 1
fi

echo "使用Python: $PYTHON_CMD"

# 检查虚拟环境
if [ -d "venv" ]; then
    echo "激活虚拟环境..."
    source venv/bin/activate
fi

# 检查依赖
echo "检查Python依赖..."
$PYTHON_CMD -c "import watchdog" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "安装依赖包..."
    pip install -r requirements.txt
fi

# 检查Thunderbird文件夹设置
echo "检查Thunderbird文件夹设置..."
$PYTHON_CMD detect_thunderbird_folders.py

echo ""
read -p "是否继续启动程序？(y/n): " -n 1 -r
echo ""
if [[ ! $REPLY =~ ^[Yy]$ ]]; then
    echo "程序已取消。请先设置好Thunderbird文件夹结构。"
    exit 1
fi

# 启动程序
echo "启动ServiceNow邮件自动批准程序..."
echo "按 Ctrl+C 停止程序"
echo "日志文件: email_auto_approve.log"
echo "----------------------------------------"

$PYTHON_CMD email_auto_approve.py