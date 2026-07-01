#!/usr/bin/env bash
# 把 src/ 的介面、排班邏輯與 ExcelJS 打包成單一自足的 index.html（雙擊即用、離線）。
# 需求：node_modules/exceljs（npm install），或以 EXCELJS_PATH 指定 exceljs.min.js。
set -e
cd "$(dirname "$0")"

EXCELJS="${EXCELJS_PATH:-node_modules/exceljs/dist/exceljs.min.js}"
if [ ! -f "$EXCELJS" ]; then
  echo "找不到 exceljs.min.js，請先執行：npm install（或設定 EXCELJS_PATH）" >&2
  exit 1
fi

{
  cat src/head.html
  echo '<script>'; cat "$EXCELJS"; echo '</script>'
  echo '<script>'; cat src/scheduler.js; echo '</script>'
  echo '<script>'; cat src/app.js; echo '</script>'
  echo '</body></html>'
} > index.html

echo "已產生 index.html ($(wc -c < index.html) bytes)"
