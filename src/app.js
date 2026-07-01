/* UI 控制器：上傳 -> 解析 -> 預假設定 -> 產生/重新產生 -> 彩色格線+違規 -> 下載。
 * 依賴全域 ExcelJS 與 Scheduler。 */
(function () {
  "use strict";
  const $ = (id) => document.getElementById(id);
  let arrayBuffer = null;
  let model = null;
  let downloadUrl = null;

  function readRules() {
    const r = JSON.parse(JSON.stringify(Scheduler.DEFAULT_RULES));
    r.minManpower = {
      D: parseInt($("r-minD").value) || 16,
      E: parseInt($("r-minE").value) || 16,
      N: parseInt($("r-minN").value) || 14,
    };
    r.maxConsecutiveWork = parseInt($("r-cw").value) || 5;
    r.maxConsecutiveN = parseInt($("r-cn").value) || 4;
    r.minLeaderPerShift = isNaN(parseInt($("r-ld").value)) ? 2 : parseInt($("r-ld").value);
    return r;
  }

  async function loadWorkbook(buf) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    return wb;
  }

  // 國定假日：以內建表為底，能連網時抓取當年度最新資料覆蓋
  async function loadHolidays(years) {
    const map = Object.assign({}, Scheduler.BUILTIN_HOLIDAYS);
    let online = false;
    for (const y of years) {
      try {
        const r = await fetch(`https://cdn.jsdelivr.net/gh/ruyut/TaiwanCalendar/data/${y}.json`, { cache: "no-store" });
        if (r.ok) {
          const arr = await r.json();
          if (Array.isArray(arr) && arr.length) { arr.forEach((o) => { if (o.description) map[o.date] = o.description; }); online = true; }
        }
      } catch (e) { /* 離線：略過，改用內建表 */ }
    }
    return { map, online };
  }
  function countFormulas(ws) {
    let f = 0;
    ws.eachRow({ includeEmpty: true }, (row) => row.eachCell({ includeEmpty: true }, (c) => { if (c.formula) f++; }));
    return f;
  }

  // 班別 -> 樣式 class
  function shiftClass(text) {
    if (!text) return "sOff";
    const up = String(text).toUpperCase();
    if (up.includes("D")) return "sD";
    if (up.includes("E")) return "sE";
    if (up.includes("N")) return "sN";
    return "sOff";
  }

  function renderLeaveTable() {
    const active = model.people.filter((p) => p.active);
    let html = "<tr><th style='width:60px'>階層</th><th style='width:80px'>姓名</th><th>預假日（逗號分隔，1 起算）</th></tr>";
    active.forEach((p) => {
      const cur = document.getElementById("lv_" + p.label);
      const val = cur ? cur.value : "";
      html += `<tr><td>${p.group}</td><td><b>${p.label}</b></td>` +
              `<td><input id="lv_${p.label}" type="text" placeholder="例：3,4,12" value="${val}"></td></tr>`;
    });
    $("leaveTable").innerHTML = html;
  }

  function readFixed() {
    const fixed = {};
    model.people.filter((p) => p.active).forEach((p) => {
      const el = document.getElementById("lv_" + p.label);
      if (!el || !el.value.trim()) return;
      const days = {};
      el.value.split(/[,，\s]+/).forEach((tok) => {
        const n = parseInt(tok, 10);
        if (!isNaN(n) && n >= 1 && n <= model.nDays) days[n - 1] = "OFF";
      });
      if (Object.keys(days).length) fixed[p.label] = days;
    });
    return fixed;
  }

  async function onFile(file) {
    try {
      arrayBuffer = await file.arrayBuffer();
      const wb = await loadWorkbook(arrayBuffer);
      const ws = wb.worksheets[0];
      model = Scheduler.parseWorkbook(ws);
      const active = model.people.filter((p) => p.active).length;
      $("m-days").textContent = model.nDays;
      $("m-people").textContent = model.people.length;
      $("m-active").textContent = active;
      $("m-formula").textContent = countFormulas(ws);
      $("metrics").classList.remove("hidden");
      $("gen").disabled = false;
      $("drop").textContent = "已載入：" + file.name + "（可重新選擇）";
      // 以範本偵測到的年月/醫院預填，方便使用者微調
      if (model.detectedHospital && !$("h-hospital").value) $("h-hospital").value = model.detectedHospital;
      // 自動推算「下一個週期」：整體往後移一個月（同起始日、同視窗長度），並預填
      let ny = model.detectedYear, nm = (model.detectedStartMonth || 1) + 1, nd = model.detectedStartDay || 1;
      if (nm > 12) { nm -= 12; ny = (ny || 0) + 1; }
      $("h-year").value = ny || "";
      $("h-month").value = nm;
      $("h-day").value = nd;
      $("winlen").textContent = model.nDays;
      $("rollNote").textContent =
        `偵測上傳週期：${model.detectedYear}年${model.detectedStartMonth}月${model.detectedStartDay}日起 ${model.nDays} 天 → 自動推算下一週期：${ny}年${nm}月${nd}日起。`;
      renderLeaveTable();
      $("leaveCard").classList.remove("hidden");
      $("result").classList.add("hidden");
    } catch (e) {
      alert("無法解析班表：" + e.message);
    }
  }

  function renderGrid(assignment, disp, fixed, flags) {
    const active = model.people.filter((p) => p.active);
    const greyLen = model.greyLen || model.tailLen;      // 匯入第一個月段長度（例：5/10~5/31 = 22）
    const greyDates = model.greyDates || model.tailDates;
    const greyWeekdays = model.greyWeekdays || model.tailWeekdays;
    const greyMonth = model.greyMonth || "";
    const vioSet = new Set(flags.cellFlags.map((f) => f.label + "|" + f.day));
    const dayVio = flags.dayFlags || {};

    // 月份帶（最上方）：把連續同月的欄位合併顯示 X月
    function segs(arr) {
      const out = [];
      (arr || []).forEach((m) => {
        if (out.length && out[out.length - 1].m === m) out[out.length - 1].n++;
        else out.push({ m, n: 1 });
      });
      return out;
    }
    let band = "<tr class='bandRow'><th class='cName'></th><th class='cGrp'></th>";
    // 灰色對照區塊：匯入月份（單一月）
    band += `<th class='prev sep' colspan='${greyLen}'>${greyMonth}月（匯入）</th>`;
    segs(model.monthsByDay).forEach((s) => { band += `<th class='mBand' colspan='${s.n}'>${s.m}月</th>`; });
    band += "</tr>";

    // 日期/星期列
    let head = "<thead>" + band + "<tr class='headRow'><th class='cName'>姓名</th><th class='cGrp'>階</th>";
    for (let i = 0; i < greyLen; i++) {
      const sep = i === greyLen - 1 ? " sep" : "";
      head += `<th class='prev${sep}'>${greyDates[i] != null ? greyDates[i] : ""}<br><span style="font-weight:400;color:#94a3b8">${greyWeekdays[i] || ""}</span></th>`;
    }
    const hol = model.holInfo || [];
    for (let d = 0; d < model.nDays; d++) {
      let cls = dayVio[d] ? "dayVio" : (hol[d] && hol[d].name ? "holiday" : (hol[d] && hol[d].weekend ? "weekend" : ""));
      const holDiv = hol[d] && hol[d].name ? `<div class='hol'>${hol[d].name}</div>` : "";
      head += `<th class='${cls}'>${model.dates[d] != null ? model.dates[d] : d + 1}<br><span style="font-weight:400;color:#94a3b8">${model.weekdays[d] || ""}</span>${holDiv}</th>`;
    }
    head += "</tr></thead>";

    // 表身：依階層分組
    let body = "<tbody>";
    let lastGrp = null;
    const totalCols = 2 + greyLen + model.nDays;
    active.forEach((p) => {
      if (p.group !== lastGrp) {
        body += `<tr class='grpRow'><td class='cName'>${p.group}階</td><td class='cGrp'></td><td colspan='${totalCols - 2}'>資深(A) → 資淺(E)</td></tr>`;
        lastGrp = p.group;
      }
      const rowVio = flags.rowFlags[p.label] ? " rowVio" : "";
      const rowTitle = flags.rowFlags[p.label] ? ` title="${flags.rowFlags[p.label].join('；')}"` : "";
      body += `<tr><td class='cName${rowVio}'${rowTitle}>${p.label}</td><td class='cGrp'>${p.group}</td>`;
      // 灰色對照：匯入月份 10~月底 的實際班別
      const gr = p.rowRaw || [];
      for (let i = 0; i < greyLen; i++) {
        const t = gr[i] || "";
        const sep = i === greyLen - 1 ? " sep" : "";
        body += `<td class='${shiftClass(t)}${sep}' style="opacity:.55">${t}</td>`;
      }
      // 本月
      for (let d = 0; d < model.nDays; d++) {
        const text = disp[p.label][d];
        const isLeave = fixed[p.label] && fixed[p.label][d] === "OFF";
        let cls = isLeave ? "sLeave" : shiftClass(text);
        if (vioSet.has(p.label + "|" + d)) cls = "vio";
        const flagMsg = flags.cellFlags.filter((f) => f.label === p.label && f.day === d).map((f) => f.msg).join("；");
        const title = flagMsg ? ` title="${flagMsg}"` : (isLeave ? ' title="預先請假"' : "");
        body += `<td class='${cls}'${title}>${text}</td>`;
      }
      body += "</tr>";
    });
    body += "</tbody>";
    $("grid").innerHTML = head + body;
  }

  function renderSummary(rows, dayFlags) {
    let html = "<tr><th>日</th><th>D 白</th><th>E 小夜</th><th>N 大夜</th><th>休假</th></tr>";
    rows.forEach((r, i) => {
      const vio = dayFlags[i] ? " class='dayVio'" : "";
      const tip = dayFlags[i] ? ` title="${dayFlags[i].join('；')}"` : "";
      html += `<tr${vio}${tip}><td>${r.day}</td><td>${r.D}</td><td>${r.E}</td><td>${r.N}</td><td>${r.off}</td></tr>`;
    });
    $("summary").innerHTML = html;
  }

  async function generateAndRender() {
    if (!arrayBuffer) return;
    $("gen").disabled = true; $("regen").disabled = true;
    const oldTxt = $("gen").textContent; $("gen").textContent = "排班中…";
    await new Promise((r) => setTimeout(r, 30));
    try {
      const rules = readRules();
      const fixed = readFixed();
      const wb = await loadWorkbook(arrayBuffer);
      const ws = wb.worksheets[0];
      model = Scheduler.parseWorkbook(ws);

      // 依起始真實日期建立日曆 + 抓取假日
      const hy = { rocYear: $("h-year").value, month: $("h-month").value, day: $("h-day").value };
      const cal = Scheduler.buildCalendar(parseInt(hy.rocYear) || model.detectedYear, parseInt(hy.month) || model.detectedStartMonth, parseInt(hy.day) || model.detectedStartDay, model.nDays);
      const tailCal = Scheduler.buildTail(parseInt(hy.rocYear) || model.detectedYear, parseInt(hy.month) || model.detectedStartMonth, parseInt(hy.day) || model.detectedStartDay, model.tailLen);
      const years = Array.from(new Set(cal.map((c) => c.y)));
      const { map: holMap, online } = await loadHolidays(years);
      $("holStatus").textContent = online
        ? `🗓️ 國定假日：已連網抓取 ${years.join("、")} 年最新資料 ✅`
        : `🗓️ 國定假日：離線，使用內建表（涵蓋 2026 年）`;
      model.dates = cal.map((c) => c.d);
      model.weekdays = cal.map((c) => c.w);
      model.tailDates = tailCal.map((c) => c.d);
      model.tailWeekdays = tailCal.map((c) => c.w);
      model.holInfo = cal.map((c) => ({ name: Scheduler.holidayName(c.key, holMap), weekend: c.w === "六" || c.w === "日" }));
      model.tailMonths = tailCal.map((c) => c.m);
      model.monthsByDay = cal.map((c) => c.m);
      model._holMap = holMap;

      const t0 = performance.now();
      const assignment = Scheduler.generate(model, rules, fixed);
      const ms = Math.round(performance.now() - t0);
      const flags = Scheduler.validate(model, assignment, rules);

      const banner = $("banner");
      if (flags.passed) {
        banner.className = "banner ok";
        banner.textContent = `✅ 排班完成，通過所有硬規則驗證（${ms} ms）`;
        $("errbox").innerHTML = "";
      } else {
        banner.className = "banner warn";
        banner.textContent = `⚠️ 已產生班表，但有 ${flags.errors.length} 項規則未滿足（紅色標示；可放寬規則或調整預假後重排）`;
        $("errbox").innerHTML =
          "<details><summary>違規明細（點開）</summary>" +
          flags.errors.slice(0, 120).map((e) => `<div class="err">• ${e}</div>`).join("") + "</details>";
      }

      const disp = {};
      Object.keys(assignment).forEach((l) => (disp[l] = Scheduler.labelOffs(assignment[l], rules)));
      renderGrid(assignment, disp, fixed, flags);
      renderSummary(Scheduler.summarizeDaily(model, assignment), flags.dayFlags);

      // 寫回範本供下載（真實日曆：日期/星期/假日/標題）
      const { title } = Scheduler.writeCalendar(ws, model, {
        hospital: $("h-hospital").value,
        rocYear: hy.rocYear, month: hy.month, day: hy.day,
        holidays: model._holMap,
      });
      Scheduler.writeInto(ws, model, assignment, rules);
      const outBuf = await wb.xlsx.writeBuffer();
      if (downloadUrl) URL.revokeObjectURL(downloadUrl);
      downloadUrl = URL.createObjectURL(new Blob([outBuf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      }));
      const fname = (title || "下月班表").replace(/\s+/g, "") + ".xlsx";
      $("dl").onclick = () => { const a = document.createElement("a"); a.href = downloadUrl; a.download = fname; a.click(); };
      $("dlnote").textContent = `標題：${title}　·　保留公式 ${countFormulas(ws)} 個`;
      $("result").classList.remove("hidden");
      if ($("result").scrollIntoView) $("result").scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (e) {
      alert("排班失敗：" + e.message); console.error(e);
    } finally {
      $("gen").disabled = false; $("regen").disabled = false; $("gen").textContent = oldTxt;
    }
  }

  // 事件綁定
  $("drop").addEventListener("click", () => $("file").click());
  $("file").addEventListener("change", (e) => { if (e.target.files[0]) onFile(e.target.files[0]); });
  $("drop").addEventListener("dragover", (e) => e.preventDefault());
  $("drop").addEventListener("drop", (e) => { e.preventDefault(); if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); });
  $("gen").addEventListener("click", generateAndRender);
  $("regen").addEventListener("click", generateAndRender);
  $("clearLeave").addEventListener("click", () => {
    model.people.filter((p) => p.active).forEach((p) => { const el = document.getElementById("lv_" + p.label); if (el) el.value = ""; });
  });
})();
