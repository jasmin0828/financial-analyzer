#!/bin/bash
# 一键部署到 GitHub 脚本
# 用法：在项目根目录运行 bash setup-github.sh

set -e

# 颜色定义
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m'

echo -e "${BLUE}==========================================${NC}"
echo -e "${BLUE}  财务分析工具 - GitHub 一键部署${NC}"
echo -e "${BLUE}==========================================${NC}"
echo ""

# 1. 检查环境
echo -e "${YELLOW}📋 步骤 1/5: 检查环境...${NC}"
if ! command -v git >/dev/null 2>&1; then
    echo -e "${RED}❌ 未安装 Git${NC}"
    echo "安装方法: xcode-select --install"
    exit 1
fi
echo -e "${GREEN}✅ Git: $(git --version)${NC}"

HAS_GH=false
if command -v gh >/dev/null 2>&1; then
    HAS_GH=true
    if gh auth status >/dev/null 2>&1; then
        echo -e "${GREEN}✅ GitHub CLI: 已登录${NC}"
    else
        echo -e "${YELLOW}⚠️ GitHub CLI 未登录${NC}"
        HAS_GH=false
    fi
else
    echo -e "${YELLOW}⚠️ 未安装 GitHub CLI（推荐装，可全自动）${NC}"
    echo "安装方法: brew install gh"
fi
echo ""

# 2. 收集仓库信息
echo -e "${YELLOW}📋 步骤 2/5: 收集仓库信息${NC}"
if [ "$HAS_GH" = true ]; then
    GH_USER=$(gh api user --jq .login)
    echo -e "GitHub 用户名: ${GREEN}$GH_USER${NC} (自动识别)"
else
    read -p "请输入你的 GitHub 用户名: " GH_USER
fi
read -p "仓库名 (默认: financial-analyzer): " REPO_NAME
REPO_NAME=${REPO_NAME:-financial-analyzer}
read -p "公开还是私有？(public/private, 默认: public): " VISIBILITY
VISIBILITY=${VISIBILITY:-public}
echo -e "配置: ${GREEN}https://github.com/$GH_USER/$REPO_NAME${NC} ($VISIBILITY)"
echo ""

# 3. 创建仓库
echo -e "${YELLOW}📋 步骤 3/5: 创建 GitHub 仓库${NC}"
if [ "$HAS_GH" = true ]; then
    gh repo create "$REPO_NAME" \
        --$VISIBILITY \
        --source=. \
        --remote=origin \
        --description "本地财务分析工具 v2.1 - 支持多年趋势分析，自动生成 Excel 图表" \
        || echo -e "${YELLOW}仓库可能已存在，继续...${NC}"
    echo -e "${GREEN}✅ 仓库已创建${NC}"
else
    echo -e "${YELLOW}请在浏览器打开: ${BLUE}https://github.com/new${NC}"
    echo -e "配置:"
    echo -e "  Repository name: ${GREEN}$REPO_NAME${NC}"
    echo -e "  Visibility: ${GREEN}$VISIBILITY${NC}"
    echo -e "  其他都不要勾选"
    echo ""
    echo -e "创建完成后按回车继续..."
    read -p ""
fi
echo ""

# 4. 提交并推送
echo -e "${YELLOW}📋 步骤 4/5: 提交并推送代码${NC}"
cd "$(dirname "$0")"

# 检查是否已经是 git 仓库
if [ ! -d .git ]; then
    git init
    git branch -M main
    echo -e "  ${GREEN}✓${NC} 初始化 Git 仓库"
fi

# 配置远程
if ! git remote get-url origin >/dev/null 2>&1; then
    git remote add origin "https://github.com/$GH_USER/$REPO_NAME.git"
    echo -e "  ${GREEN}✓${NC} 添加远程仓库"
fi

# 提交
git add .
if git diff --cached --quiet; then
    echo -e "  ${YELLOW}⚠ 没有新文件需要提交${NC}"
else
    git commit -m "init: 财务分析工具 v2.1

- 多年财报支持（自动按文件名识别年份）
- 42 项财务指标（盈利/偿债/营运/现金流/发展/杜邦）
- 趋势分析（YoY 同比 + CAGR 复合增长率）
- openpyxl 原生图表（折线图 + 饼图）
- 全本地运行，数据不外传"
    echo -e "  ${GREEN}✓${NC} 提交代码"
fi

# 推送
echo -e "  正在推送到 GitHub..."
if git push -u origin main 2>&1 | tail -5; then
    echo -e "${GREEN}✅ 代码推送成功${NC}"
else
    echo -e "${RED}❌ 推送失败${NC}"
    echo "可能原因："
    echo "  1. 仓库未创建 → 回到步骤 3 创建"
    echo "  2. 网络问题"
    echo "  3. GitHub 认证问题"
    exit 1
fi
echo ""

# 5. 完成提示
echo -e "${YELLOW}📋 步骤 5/5: 等待 GitHub Actions 打包${NC}"
echo -e "${GREEN}==========================================${NC}"
echo -e "${GREEN}  🎉 上传完成！${NC}"
echo -e "${GREEN}==========================================${NC}"
echo ""
echo -e "接下来:"
echo -e "  1. 打开 ${BLUE}https://github.com/$GH_USER/$REPO_NAME/actions${NC}"
echo -e "  2. 看到 'Build Windows EXE' 在跑（黄点=进行中，✅=完成）"
echo -e "  3. 等待 5-10 分钟"
echo -e "  4. 点进 run → 底部 Artifacts → 下载 ${GREEN}财务分析工具-windows-exe.zip${NC}"
echo -e "  5. 解压得到 .exe，复制到 Windows 双击运行"
echo ""
echo -e "${YELLOW}💡 小提示：${NC}"
echo -e "  - 首次构建较慢（装依赖），后续约 1-2 分钟"
echo -e "  - 以后改代码: ${BLUE}git add . && git commit -m 'msg' && git push${NC}"
echo -e "  - Actions 会自动重新打包"
echo ""

# 询问是否打开浏览器
read -p "是否打开 GitHub Actions 页面？(y/n): " OPEN_BROWSER
if [ "$OPEN_BROWSER" = "y" ] || [ "$OPEN_BROWSER" = "Y" ]; then
    open "https://github.com/$GH_USER/$REPO_NAME/actions" 2>/dev/null \
        || xdg-open "https://github.com/$GH_USER/$REPO_NAME/actions" 2>/dev/null \
        || echo "请手动打开: https://github.com/$GH_USER/$REPO_NAME/actions"
fi
