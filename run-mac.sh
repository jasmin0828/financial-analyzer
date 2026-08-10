#!/bin/bash
# ============================================================
# 财务分析工具 - Mac 启动脚本
# ============================================================
# 功能：
# 1. 自动检测 Python 环境
# 2. 没装 Python 时弹出引导对话框
# 3. 自动装/更新依赖
# 4. 启动 GUI
# ============================================================

set -e
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# 颜色
BLUE='\033[0;34m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${BLUE}=========================================="
echo -e "  财务分析工具 v$(cat _version.py | grep -oE '"[0-9.]+"' | head -1 | tr -d '"')"
echo -e "  Mac 启动脚本"
echo -e "==========================================${NC}"
echo ""

# ===== 步骤 1: 检查 Python =====
echo -e "${YELLOW}📋 步骤 1/4: 检查 Python 环境...${NC}"

# 检查 python3
PYTHON_CMD=""
if command -v python3 &> /dev/null; then
    PY_VER=$(python3 --version 2>&1 | awk '{print $2}')
    PY_MAJOR=$(echo "$PY_VER" | cut -d. -f1)
    PY_MINOR=$(echo "$PY_VER" | cut -d. -f2)
    if [ "$PY_MAJOR" -ge 3 ] && [ "$PY_MINOR" -ge 10 ]; then
        echo -e "${GREEN}✅ Python $PY_VER${NC}"
        PYTHON_CMD="python3"
    else
        echo -e "${YELLOW}⚠️  Python $PY_VER 版本过低（需要 3.10+）${NC}"
    fi
fi

# 检查 python（部分 Mac 旧版可能只有 python）
if [ -z "$PYTHON_CMD" ] && command -v python &> /dev/null; then
    PY_VER=$(python --version 2>&1 | awk '{print $2}')
    PY_MAJOR=$(echo "$PY_VER" | cut -d. -f1)
    PY_MINOR=$(echo "$PY_VER" | cut -d. -f2)
    if [ "$PY_MAJOR" -ge 3 ] && [ "$PY_MINOR" -ge 10 ]; then
        echo -e "${GREEN}✅ Python $PY_VER${NC}"
        PYTHON_CMD="python"
    fi
fi

# 没找到合适的 Python
if [ -z "$PYTHON_CMD" ]; then
    echo -e "${RED}❌ 未检测到 Python 3.10+${NC}"
    echo ""
    echo -e "${YELLOW}本工具需要 Python 3.10 或更高版本${NC}"
    echo ""
    echo -e "请按以下步骤安装："
    echo -e "  1. 访问 ${BLUE}https://www.python.org/downloads/${NC}"
    echo -e "  2. 下载 Python 3.11+ (macOS)"
    echo -e "  3. 运行安装包，${RED}勾选 'Add Python to PATH'${NC}"
    echo -e "  4. 安装完成后重新运行本脚本"
    echo ""

    # 弹原生对话框询问是否打开下载页
    if command -v osascript &> /dev/null; then
        RESULT=$(osascript -e 'display dialog "财务分析工具需要 Python 3.10+ 才能运行。\n\n是否现在打开 Python 官网下载？" buttons {"取消", "打开下载页"} default button "打开下载页" with title "需要安装 Python"' 2>&1)
        if [[ "$RESULT" == *"打开下载页"* ]]; then
            echo -e "${BLUE}→ 打开 python.org/downloads/${NC}"
            open "https://www.python.org/downloads/"
        fi
    fi
    echo ""
    echo "安装 Python 后，请重新运行此脚本。"
    exit 1
fi

# ===== 步骤 2: 检查依赖 =====
echo -e "${YELLOW}📋 步骤 2/4: 检查依赖...${NC}"
if ! $PYTHON_CMD -c "import PyQt5, pandas, pdfplumber, openpyxl" 2>/dev/null; then
    echo -e "${YELLOW}⚠️  依赖未安装，正在安装...${NC}"
    echo -e "${BLUE}   (首次约 2-3 分钟，PyQt5 较大)${NC}"
    echo ""

    # 升级 pip
    $PYTHON_CMD -m pip install --upgrade pip 2>&1 | tail -3

    # 安装依赖
    $PYTHON_CMD -m pip install -r requirements.txt 2>&1 | tail -5

    # 验证
    if ! $PYTHON_CMD -c "import PyQt5, pandas, pdfplumber, openpyxl" 2>/dev/null; then
        echo -e "${RED}❌ 依赖安装失败${NC}"
        echo "请手动运行: $PYTHON_CMD -m pip install -r requirements.txt"
        exit 1
    fi
    echo -e "${GREEN}✅ 依赖安装完成${NC}"
else
    echo -e "${GREEN}✅ 依赖已就绪${NC}"
fi
echo ""

# ===== 步骤 3: 启动 GUI =====
echo -e "${YELLOW}📋 步骤 3/4: 启动 GUI...${NC}"
echo -e "${BLUE}==========================================${NC}"
echo -e "${GREEN}  🚀 启动中...（关闭窗口退出）${NC}"
echo -e "${BLUE}==========================================${NC}"
echo ""

# ===== 步骤 4: 清理 =====
echo -e "${YELLOW}📋 步骤 4/4: 完成${NC}"
echo ""

# 真正启动
$PYTHON_CMD main.py
