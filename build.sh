#!/usr/bin/env bash
# 把介面骨架、排班邏輯與 ExcelJS 打包成單一自足的 index.html（雙擊即用、離線）。
# refactor 移除 src/head.html 後，介面骨架（HTML/CSS）直接維護在 index.html；
# build 從現有 index.html 取出骨架（第一個 <script> 之前的部分），再重新注入
# ExcelJS + src/scheduler.js + src/app.js，組回單一自足的成品。
# 需求：node_modules/exceljs（npm install），或以 EXCELJS_PATH 指定 exceljs.min.js。
set -e
cd "$(dirname "$0")"

EXCELJS="${EXCELJS_PATH:-node_modules/exceljs/dist/exceljs.min.js}"
if [ ! -f "$EXCELJS" ]; then
  echo "找不到 exceljs.min.js，請先執行：npm install（或設定 EXCELJS_PATH）" >&2
  exit 1
fi
if [ ! -f index.html ]; then
  echo "找不到 index.html（介面骨架來源）" >&2
  exit 1
fi

# 介面骨架 = index.html 開頭到第一個 <script> 之前（即原 src/head.html 的內容，現維護於 index.html）。
# awk 讀 index.html、輸出寫到 .tmp（讀寫不同檔，避免自舉衝突），組完再一次覆蓋。
{
  awk '/<script>/{exit} {print}' index.html
  echo '<script>'; cat "$EXCELJS"; echo '</script>'
  echo '<script>'; cat src/scheduler.js; echo '</script>'
  echo '<script>'; cat src/app.js; echo '</script>'
  echo '</body></html>'
} > index.html.tmp
mv index.html.tmp index.html

echo "已產生 index.html ($(wc -c < index.html) bytes)"
