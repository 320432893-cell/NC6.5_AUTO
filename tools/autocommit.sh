#!/usr/bin/env bash
# 自动小步提交(机器钩子,非 AI 决定)——见 .ai-config rules/review/review.index.md §1
# 设计:编译过即提到工作分支;非语法只 flag 不 block;不直怼 main。
# 用法:存盘后 / 每 N 分钟由 watch 或编辑器钩子调用 `tools/autocommit.sh`。
set -euo pipefail
cd "$(dirname "$0")/.."

branch=$(git rev-parse --abbrev-ref HEAD)
if [ "$branch" = "main" ] || [ "$branch" = "master" ]; then
  echo "[autocommit] 在 $branch 上,拒绝自动提交;请先开工作分支。"; exit 0
fi

# 无改动则退
if git diff --quiet && git diff --cached --quiet && [ -z "$(git status --porcelain --untracked-files=normal | grep -v '^!!')" ]; then
  exit 0
fi

# 硬闸:改动的 .py 必须能编译,否则不提(绝不提坏语法)
mapfile -t pyfiles < <(git status --porcelain | awk '{print $2}' | grep -E '\.py$' | grep -v 'archive/' || true)
for f in "${pyfiles[@]:-}"; do
  [ -f "$f" ] || continue
  if ! python3 -m py_compile "$f" 2>/dev/null; then
    echo "[autocommit] $f 编译不过,不提交(等你修好)。"; exit 0
  fi
done

# 软 flag:ruff(不 block,只提示)
if [ -x .venv/bin/ruff ] && [ "${#pyfiles[@]:-0}" -gt 0 ]; then
  .venv/bin/ruff check "${pyfiles[@]}" 2>/dev/null | tail -3 || true
fi

git add -A
ts=$(date "+%Y-%m-%d %H:%M")
git commit -q -m "wip: 自动小步提交 $ts" && echo "[autocommit] 已提交 ($branch, $ts)"
