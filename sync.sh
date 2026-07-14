#!/bin/bash
# 將儀表板資料同步到 GitHub Pages。由猿田彥（Hermes）在更新 data.json 後呼叫。
cd "$(dirname "$0")" || exit 1

if [ ! -d .git ]; then
  echo "尚未設定 git repo，僅更新本機資料（跳過同步）"
  exit 0
fi

git add -A
if git diff --cached --quiet; then
  echo "沒有變更，不需同步"
  exit 0
fi

git commit -m "更新儀表板資料 $(date '+%Y-%m-%d %H:%M')" --quiet
if git push --quiet 2>&1; then
  echo "已同步到線上儀表板"
else
  echo "推送失敗，資料已存在本機，稍後可重試 bash sync.sh"
  exit 1
fi
