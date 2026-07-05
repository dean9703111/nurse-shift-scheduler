# 護理師排班產生器（純前端 · 離線單檔）

上傳「上月 Excel 班表」，一鍵產生「下月班表」。整個工具是**單一 HTML 檔、雙擊即用**，不需安裝、不需 Python、不需後端；班表資料**全程留在本機瀏覽器、不上傳**。

> 輸出檔以上傳檔為範本，**格式、樣式、COUNTIF／驗證公式完整保留**；日期、星期、國定假日依真實日曆自動產生。

## 特色

- 📄 **單一檔案離線可用**：`index.html` 內嵌 ExcelJS 與排班邏輯，雙擊即可在 Chrome / Edge / Safari 開啟。
- 🗓️ **自動推算下一週期**：讀取上傳檔的年月與視窗長度，整體往後推一個月自動預填（例：上傳 5/10–6/30 → 產生 6/10–7/31）。
- 🧮 **保留公式與格式**：只覆寫排班資料格，範本的統計／驗證公式原封不動。
- 🇹🇼 **國定假日自動標註**：內建 2026 年假日表，能連網時抓取最新資料（[ruyut/TaiwanCalendar](https://github.com/ruyut/TaiwanCalendar)）覆蓋；只標視窗實際涵蓋到的假日。
- 👥 **人員 × 日期完整格線**：依階層分組、上月對照區塊、違規紅色標示、每日人力摘要。
- 🛌 **預先請假**：自動載入所有人員，可個別設定預假日並重新產生。
- ✅ **硬規則驗證**：最低人力、階層上限、小組長、連續上班/大夜、換班休息（11 小時）、N-off-D（大夜後僅休 1 天不得接白班）、休假視窗，違規即時標示。
- 🧑‍🎓 **新人不排班**：預假面板可勾選「不排班」，整月排除該員（輸出留白、不計入人力）。

## 使用方式

1. 下載或 clone 本 repo，直接**雙擊 `index.html`**（或見下方 GitHub Pages）。
2. 上傳上月班表 `.xlsx`。
3. 確認自動預填的年月、起始日與規則。
4. （可選）在「預先請假設定」填入個別人員預假日。
5. 按「產生下月班表」→ 檢視結果 → 下載。

### GitHub Pages（線上版）

本 repo 已是靜態網頁，可於 **Settings → Pages → Deploy from branch → main / root** 啟用，之後即可用網址開啟 `index.html`。

## 開發

原始碼在 `src/`，`index.html` 是打包後的成品。

```
├── index.html          # 打包後的單檔應用（可直接用）
├── build.sh            # 重新打包 index.html
├── package.json
└── src/
    ├── head.html       # 介面 HTML / CSS
    ├── scheduler.js    # 排班核心：解析 / 產生 / 驗證 / 寫回 / 日曆 / 假日
    ├── app.js          # 介面控制器
    ├── test_node.js    # 核心邏輯測試（Node）
    ├── test_html.js    # 單檔 HTML 瀏覽器模擬測試（jsdom）
    └── test_rules.js   # 需求規則驗證：規則常數對照範本公式＋驗證器正反例＋產生結果獨立稽核
```

重新打包與測試：

```bash
npm install          # 安裝 exceljs、jsdom
npm run build        # 重新產生 index.html
npm test             # 執行測試（需在根目錄放一份 sample.xlsx）
```

> 測試需要一份符合格式的班表，命名為 `sample.xlsx` 放在專案根目錄。基於個資考量，實際班表**不會**提交進 repo（已列入 `.gitignore`）。

## 技術

| 功能 | 技術 |
|------|------|
| 讀寫 xlsx（保留公式+樣式） | [ExcelJS](https://github.com/exceljs/exceljs)（內嵌） |
| 排班求解 | 純 JavaScript 啟發式（貪婪＋公平性/純班偏好） |
| 國定假日 | 內建表 + [ruyut/TaiwanCalendar](https://github.com/ruyut/TaiwanCalendar) |

排班為啟發式（非最佳解）；遇到較難的衝突會以紅色標示，供人工微調。

## 授權

MIT License，詳見 [LICENSE](./LICENSE)。
