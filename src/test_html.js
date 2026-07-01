/* 用 jsdom 在瀏覽器模擬環境載入單檔 HTML，驗證內嵌的 ExcelJS 與 Scheduler
 * 能在「瀏覽器」中正常載入並跑完整流程（解析->產生->驗證->寫回->公式保留）。
 */
const fs = require("fs");
const path = require("path");
const { JSDOM } = require("jsdom");

const HTML = path.resolve(__dirname, "..", "index.html");
const SRC = path.resolve(__dirname, "..", "sample.xlsx");

(async () => {
  const html = fs.readFileSync(HTML, "utf8");
  const dom = new JSDOM(html, { runScripts: "dangerously", resources: "usable" });
  const win = dom.window;
  // 等待內嵌 script 執行完成
  await new Promise((r) => setTimeout(r, 800));

  const checks = [];
  checks.push(["內嵌 ExcelJS 載入", typeof win.ExcelJS !== "undefined", typeof win.ExcelJS]);
  checks.push(["內嵌 Scheduler 載入", typeof win.Scheduler !== "undefined", typeof win.Scheduler]);
  checks.push(["UI 元素存在(上傳/產生鈕)",
    !!win.document.getElementById("gen") && !!win.document.getElementById("file"), "ok"]);

  // 用瀏覽器環境內的 ExcelJS/Scheduler 跑一次完整流程（等同按下產生鈕的邏輯）
  const nb = fs.readFileSync(SRC);
  // 在 jsdom realm 內建立 ArrayBuffer（真實瀏覽器單一 realm 無此問題）
  const u8 = new win.Uint8Array(nb.byteLength);
  u8.set(new Uint8Array(nb.buffer, nb.byteOffset, nb.byteLength));
  const wb = new win.ExcelJS.Workbook();
  await wb.xlsx.load(u8.buffer);
  const ws = wb.worksheets[0];
  const model = win.Scheduler.parseWorkbook(ws);
  checks.push(["瀏覽器端解析天數", model.nDays === 52, `n_days=${model.nDays}`]);

  // 上月最後 7 天資料
  checks.push(["上月尾巴7天(日期)", model.tailDates.length === 7, `tailDates=${model.tailDates.join(",")}`]);
  const p0 = model.people.find((p) => p.active);
  checks.push(["每人 rawTail 7 格", p0.rawTail.length === 7, `${p0.label}: ${p0.rawTail.join("/")}`]);

  // 灰色對照區塊 = 匯入第一個月段（5/10~5/31）
  checks.push(["灰色區塊=匯入月份10~月底", model.greyLen === 22 && String(model.greyDates[0]) === "10" && String(model.greyDates[model.greyLen - 1]) === "31" && model.greyMonth === 5,
    `greyLen=${model.greyLen} 首=${model.greyDates[0]} 末=${model.greyDates[model.greyLen - 1]} 月=${model.greyMonth}`]);
  checks.push(["每人整月原始班別 rowRaw", p0.rowRaw.length === model.nDays, `rowRaw=${p0.rowRaw.length}格`]);

  // 預先請假(fixed) 生效驗證
  const tgt = p0.label;
  const fixed = {}; fixed[tgt] = { 0: "OFF", 1: "OFF", 2: "OFF" };
  const asgF = win.Scheduler.generate(model, win.Scheduler.DEFAULT_RULES, fixed);
  const leaveOK = asgF[tgt][0] === null && asgF[tgt][1] === null && asgF[tgt][2] === null;
  checks.push(["預假日確實排休", leaveOK, `${tgt} 前3天=${asgF[tgt].slice(0,3).map((x)=>x||"休").join(",")}`]);

  // 結構化違規：故意用過高人力需求觸發，確認 dayFlags/cellFlags 有座標
  const hardRules = JSON.parse(JSON.stringify(win.Scheduler.DEFAULT_RULES));
  hardRules.minManpower = { D: 40, E: 40, N: 40 }; // 不可能滿足 -> 應產生 dayFlags
  const asgH = win.Scheduler.generate(model, hardRules);
  const vf = win.Scheduler.validate(model, asgH, hardRules);
  const hasDayFlags = Object.keys(vf.dayFlags).length > 0 && Array.isArray(vf.cellFlags);
  checks.push(["違規回傳結構化座標", hasDayFlags, `dayFlags=${Object.keys(vf.dayFlags).length}天, cellFlags=${vf.cellFlags.length}`]);

  let f0 = 0; ws.eachRow({ includeEmpty: true }, (row) => row.eachCell({ includeEmpty: true }, (c) => { if (c.formula) f0++; }));
  const assignment = win.Scheduler.generate(model, win.Scheduler.DEFAULT_RULES);
  const { passed, errors } = win.Scheduler.validate(model, assignment, win.Scheduler.DEFAULT_RULES);
  checks.push(["瀏覽器端排班通過規則", passed, passed ? "無違規" : `${errors.length} 項`]);

  // 年月偵測（含起始日）
  checks.push(["偵測範本年月日", model.detectedYear && model.detectedStartMonth && model.detectedStartDay,
    `年=${model.detectedYear} 月=${model.detectedStartMonth} 日=${model.detectedStartDay} 醫院=${model.detectedHospital}`]);

  // 真實日曆：以 民國115/6/1 起算（涵蓋端午節 6/19），驗證日期從 1 起、星期真實、假日標註
  const opts = { hospital: "台大醫院", rocYear: 115, month: 6, day: 1, holidays: win.Scheduler.BUILTIN_HOLIDAYS };
  const { title, cal } = win.Scheduler.writeCalendar(ws, model, opts);
  checks.push(["日期從 1 號起算", cal[0].d === 1, `第1欄日期=${cal[0].d}`]);
  checks.push(["星期取自真實日曆", cal[0].w === "一", `6/1=星期${cal[0].w}（2026/6/1應為一）`]);
  checks.push(["標題年月正確(主要月6月)", title === "台大醫院115年6月護理師班表", `"${title}"`]);
  // 端午節在第19天(6/19)，應標註於月份標記列
  const dcol = model.dateCols[18];
  const holCell = String(ws.getCell(model.monthLabelRow, dcol).value || "");
  checks.push(["國定假日自動標註(端午節)", holCell.includes("端午節"), `6/19格="${holCell}"`]);
  // 跨月標記 7月 應出現在 7/1 (第31天)
  const jul = String(ws.getCell(model.monthLabelRow, model.dateCols[30]).value || "");
  checks.push(["跨月標記(7月)", jul.includes("7月"), `7/1格="${jul}"`]);
  // 上月7天：6/1 前 7 天應為 5/25..5/31
  const tail = win.Scheduler.buildTail(115, 6, 1, 7).map((c) => c.d).join(",");
  checks.push(["上月7天為真實前7日", tail === "25,26,27,28,29,30,31", `tail=${tail}`]);

  win.Scheduler.writeInto(ws, model, assignment, win.Scheduler.DEFAULT_RULES);
  const outBuf = await wb.xlsx.writeBuffer();
  const wb2 = new win.ExcelJS.Workbook();
  await wb2.xlsx.load(outBuf);
  const ws2 = wb2.worksheets[0];
  let f1 = 0; ws2.eachRow({ includeEmpty: true }, (row) => row.eachCell({ includeEmpty: true }, (c) => { if (c.formula) f1++; }));
  checks.push(["輸出(writeBuffer)公式保留", f1 >= f0, `${f0} -> ${f1}`]);
  const t2 = String(ws2.getCell(model.titleCell.r, model.titleCell.c).value || "");
  checks.push(["輸出檔標題含年月", t2.includes("115年6月"), `"${t2}"`]);

  // === 實際驅動 UI（上傳 -> 產生），驗證 _app.js 的 DOM 渲染不出錯 ===
  win.URL.createObjectURL = () => "blob:fake";
  win.URL.revokeObjectURL = () => {};
  if (!win.File.prototype.arrayBuffer) {
    win.File.prototype.arrayBuffer = function () { return Promise.resolve(u8.buffer); };
  }
  try {
    const fileInput = win.document.getElementById("file");
    const f = new win.File([u8], "上月班表.xlsx");
    f.arrayBuffer = () => Promise.resolve(u8.buffer);
    Object.defineProperty(fileInput, "files", { value: [f], configurable: true });
    fileInput.dispatchEvent(new win.Event("change"));
    await new Promise((r) => setTimeout(r, 400));
    const leaveInputs = win.document.querySelectorAll("#leaveTable input").length;
    checks.push(["UI:預假面板載入所有人員", leaveInputs === model.people.filter((p) => p.active).length, `${leaveInputs} 個輸入`]);

    // 自動推算下一週期：偵測起始月5 -> 預填6
    const uiMonth = win.document.getElementById("h-month").value;
    const uiDay = win.document.getElementById("h-day").value;
    checks.push(["UI:自動推算下一週期(月+1)", uiMonth === "6" && uiDay === "10", `預填 ${uiMonth}月${uiDay}日`]);

    // 設一個人的預假，按重新產生
    const li = win.document.getElementById("lv_" + p0.label);
    if (li) li.value = "2,3";
    win.document.getElementById("regen").click();
    await new Promise((r) => setTimeout(r, 500));
    const gridCells = win.document.querySelectorAll("#grid tbody td").length;
    checks.push(["UI:結果格線已渲染", gridCells > 100, `${gridCells} 格`]);
    // 月份帶：應含 6月 與 7月
    const bandTxt = Array.from(win.document.querySelectorAll("#grid thead tr.bandRow th.mBand")).map((e) => e.textContent).join(",");
    checks.push(["UI:上方月份帶顯示", bandTxt.includes("6月") && bandTxt.includes("7月"), `月份帶=${bandTxt}`]);
    const leaveCells = win.document.querySelectorAll("#grid td.sLeave").length;
    checks.push(["UI:預假格上黃色標記", leaveCells >= 2, `${leaveCells} 格 sLeave`]);
    const headPrev = win.document.querySelectorAll("#grid thead tr.headRow th.prev").length;
    checks.push(["UI:灰色區塊=匯入首月(22欄)", headPrev === 22, `${headPrev} 欄`]);
    const bandPrev = win.document.querySelector("#grid thead tr.bandRow th.prev");
    checks.push(["UI:灰色帶標示匯入月份", bandPrev && bandPrev.textContent.includes("5月"), `"${bandPrev ? bandPrev.textContent : ""}"`]);
    const dlNote = win.document.getElementById("dlnote").textContent;
    checks.push(["UI:下載就緒(公式數)", /\d/.test(dlNote), dlNote]);
  } catch (e) {
    checks.push(["UI 驅動流程", false, e.message]);
  }

  console.log("=".repeat(56));
  console.log("單檔 HTML 瀏覽器模擬(jsdom) 驗證報告");
  console.log("=".repeat(56));
  let ok = true;
  checks.forEach(([n, p, d]) => { console.log(`${p ? "✅" : "❌"}  ${n.padEnd(22)} ${d}`); ok = ok && p; });
  console.log("=".repeat(56));
  console.log("結果：", ok ? "✅ 單檔 HTML 可在瀏覽器正常運作" : "❌ 有問題");
  process.exit(ok ? 0 : 1);
})().catch((e) => { console.error("執行失敗:", e); process.exit(1); });
