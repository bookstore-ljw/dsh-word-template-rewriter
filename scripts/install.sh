#!/usr/bin/env bash
# 一键安装（macOS / Linux）
#
#   bash scripts/install.sh            # pip 安装 + 注册技能
#   bash scripts/install.sh --render   # 额外装 pymupdf（表格列宽按字宽计算更准）
#
# 注意：Word 版面复核（retypeset --verify / render）依赖 Word COM，只有 Windows 可用；
# 其余能力（retypeset 本体、check、analyze）跨平台可用。
set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT"
echo "仓库：$ROOT"

WITH_RENDER=0
for arg in "$@"; do
  case "$arg" in
    --render) WITH_RENDER=1 ;;
  esac
done

echo
echo "[1/3] 安装 Python 包"
if [ "$WITH_RENDER" = "1" ]; then
  python3 -m pip install ".[render]"
else
  python3 -m pip install .
fi

echo
echo "[2/3] 注册 DSH 技能"
SKILLS_SRC="$ROOT/plugin/skills"
PROFILES="$HOME/.dsh/profiles"
REGISTERED=0
if [ -d "$PROFILES" ]; then
  for profile in "$PROFILES"/*/; do
    if [ -e "${profile}node_modules/@local/dsh-word-pipeline" ]; then
      echo "  已通过插件链接可见：$(basename "$profile")"
      REGISTERED=$((REGISTERED + 1))
      continue
    fi
    mkdir -p "${profile}skills"
    for skill in "$SKILLS_SRC"/*/; do
      name="$(basename "$skill")"
      if [ ! -e "${profile}skills/$name" ]; then
        cp -R "$skill" "${profile}skills/$name"
        echo "  复制技能 $name → $(basename "$profile")"
        REGISTERED=$((REGISTERED + 1))
      fi
    done
  done
fi
if [ "$REGISTERED" = "0" ]; then
  echo "  未发现 DSH 配置目录，跳过（仅 Python CLI 可用）"
  echo "  技能文件在：$SKILLS_SRC"
fi

echo
echo "[3/3] 自检"
python3 -m dsh_word retypeset --help | head -3
python3 -m unittest discover -s tests -p "test_*.py"

echo
echo "完成。用法示例："
echo "  dsh-word retypeset --template 模板.docx --source 原文.docx --out 成稿.docx --profile ieee-to-ttp"
