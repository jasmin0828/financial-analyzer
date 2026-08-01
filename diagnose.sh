#!/bin/bash
# 诊断脚本：检查为什么 Release 没创建
set -e
cd "$(dirname "$0")"

BLUE='\033[0;34m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

echo -e "${BLUE}==========================================${NC}"
echo -e "${BLUE}  财务分析工具 - 部署诊断${NC}"
echo -e "${BLUE}==========================================${NC}"
echo ""

# 1. 本地状态
echo -e "${YELLOW}📋 1. 本地 Git 状态${NC}"
echo "当前分支: $(git branch --show-current)"
echo ""
echo "最近 5 次提交:"
git log --oneline -5
echo ""
echo "未推送的提交数: $(git log --oneline @{u}.. 2>/dev/null | wc -l | tr -d ' ')"
echo ""

# 2. 远程状态
echo -e "${YELLOW}📋 2. 远程状态${NC}"
git status -sb
echo ""
git remote -v
echo ""

# 3. build.yml 当前内容
echo -e "${YELLOW}📋 3. build.yml 中的 Release 条件（关键）${NC}"
grep -A 6 "创建 Release" .github/workflows/build.yml | head -10
echo ""

# 4. 检查 GITHUB_TOKEN 权限设置（GitHub UI 上的）
echo -e "${YELLOW}📋 4. GITHUB_TOKEN 权限设置${NC}"
echo "⚠️  这个需要手动检查 GitHub 网页："
echo "   1. 打开 https://github.com/jasmin0828/financial-analyzer/settings/actions"
echo "   2. 滚到 'Workflow permissions'"
echo "   3. 检查是 'Read repository contents and packages' 还是 'Read and write'"
echo "   ❌ 如果是前者 → 改成 'Read and write permissions'"
echo ""

# 5. 检查默认分支
echo -e "${YELLOW}📋 5. 默认分支是什么？${NC}"
echo "⚠️  也要手动检查："
echo "   打开 https://github.com/jasmin0828/financial-analyzer"
echo "   看左下角分支选择器显示的默认分支名（main 或 master）"
echo ""

echo -e "${GREEN}==========================================${NC}"
echo -e "${GREEN}  诊断完成！把上面的输出 + 上面 4 和 5 的答案发给我${NC}"
echo -e "${GREEN}==========================================${NC}"
