#!/bin/bash
# Mac 一键启动脚本
# 双击或在终端运行：bash run-mac.sh

set -e
cd "$(dirname "$0")"

# 颜色
BLUE='\033[0;34m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${BLUE}==========================================${NC}"
echo -e "${BLUE}  财务分析工具 - macOS 启动${NC}"
echo -e "${BLUE}==========================================${NC}"
echo ""

# 1. 检查 Python
echo -e "${YELLOW}📋 步骤 1/3: 检查 Python...${NC}"
if ! command -v python3 >/dev/null 2>&1; then
    echo -e "${RED}❌ 未安装 Python 3${NC}"
    echo "安装方法: brew install python3"
    echo "或: https://www.python.org/downloads/"
    exit 1
fi
echo -e "${GREEN}✅ Python: $(python3 --version)${NC}"
echo ""

# 2. 安装依赖
echo -e "${YELLOW}📋 步骤 2/3: 检查依赖...${NC}"
if ! python3 -c "import PyQt5" 2>/dev/null; then
    echo -e "${YELLOW}⚠️  PyQt5 未安装，正在安装所有依赖...${NC}"
    echo -e "${YELLOW}   (首次约 2-3 分钟)${NC}"
    pip3 install -r requirements.txt
fi
if ! python3 -c "import PyQt5, pdfplumber" 2>/dev/null; then
    echo -e "${RED}❌ 依赖安装失败${NC}"
    exit 1
fi
echo -e "${GREEN}✅ 依赖已就绪${NC}"
echo ""

# 3. 启动 GUI
echo -e "${YELLOW}📋 步骤 3/3: 启动 GUI...${NC}"
echo -e "${GREEN}==========================================${NC}"
echo -e "${GREEN}  🚀 启动中...（关闭窗口退出）${NC}"
echo -e "${GREEN}==========================================${NC}"
python3 main.py
