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
    const w14 = isNaN(parseInt($("r-w14").value)) ? 2 : parseInt($("r-w14").value);
    const w28 = isNaN(parseInt($("r-w28").value)) ? 8 : parseInt($("r-w28").value);
    r.restWindows = [[14, w14], [28, w28]];
    return r;
  }

  async function loadWorkbook(buf) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buf);
    return wb;
  }

  // 國定假日：以內建表為底，能連網時抓取當年度最新資料覆蓋（同年度快取，點擊重排不重抓）
  const holCache = {};
  async function loadHolidays(years) {
    const key = years.join(",");
    if (holCache[key]) return holCache[key];
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
    return (holCache[key] = { map, online });
  }
  function countFormulas(ws) {
    let f = 0;
    ws.eachRow({ includeEmpty: true }, (row) => row.eachCell({ includeEmpty: true }, (c) => { if (c.formula) f++; }));
    return f;
  }

  // 違規原因懸浮提示：滑鼠移到帶 data-tip 的元素即顯示（原生 title 延遲太久）
  const tipEl = document.createElement("div");
  tipEl.id = "tip";
  document.body.appendChild(tipEl);
  document.addEventListener("mouseover", (e) => {
    const t = e.target.closest ? e.target.closest("[data-tip]") : null;
    if (!t) { tipEl.style.display = "none"; return; }
    tipEl.textContent = t.getAttribute("data-tip").split("；").join("\n");
    tipEl.style.display = "block";
  });
  document.addEventListener("mousemove", (e) => {
    if (tipEl.style.display !== "block") return;
    const pad = 14;
    let x = e.clientX + pad, y = e.clientY + pad;
    if (x + tipEl.offsetWidth > window.innerWidth - 8) x = e.clientX - tipEl.offsetWidth - pad;
    if (y + tipEl.offsetHeight > window.innerHeight - 8) y = e.clientY - tipEl.offsetHeight - pad;
    tipEl.style.left = x + "px"; tipEl.style.top = y + "px";
  });
  function tipAttr(msgs) {
    if (!msgs || !msgs.length) return "";
    const text = (Array.isArray(msgs) ? msgs.join("；") : String(msgs)).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
    return ` data-tip="${text}"`;
  }

  // 班別 -> 樣式 class（前綴判斷，與範本 COUNTIF 語意一致：病E/生理d 視為休）
  function shiftClass(text) {
    if (!text) return "sOff";
    const up = String(text).toUpperCase();
    if (up.startsWith("D")) return "sD";
    if (up.startsWith("E")) return "sE";
    if (up.startsWith("N")) return "sN";
    return "sOff";
  }

  function renderLeaveTable() {
    const active = model.people.filter((p) => p.active);
    let html = "<tr><th style='width:60px'>階層</th><th style='width:80px'>姓名</th><th style='width:130px'>包班（純班）</th><th style='width:70px'>不排班</th><th>預假日（排班月的日期，逗號分隔）</th></tr>";
    active.forEach((p) => {
      const cur = document.getElementById("lv_" + p.label);
      const val = cur ? cur.value : "";
      const curLk = document.getElementById("lk_" + p.label);
      const lockVal = curLk ? curLk.value : (p.suggestedLock || "");
      const curNs = document.getElementById("ns_" + p.label);
      const nsChecked = curNs ? curNs.checked : false;
      const opts = [["", "不包班"], ["D", "D 白班"], ["E", "E 小夜"], ["N", "N 大夜"]]
        .map(([v, t]) => {
          const mark = p.suggestedLock === v ? "（上月偵測）" : "";
          return `<option value="${v}"${v === lockVal ? " selected" : ""}>${t}${mark}</option>`;
        }).join("");
      html += `<tr><td>${p.group}</td><td><b>${p.label}</b></td>` +
              `<td><select id="lk_${p.label}">${opts}</select></td>` +
              `<td style="text-align:center"><input id="ns_${p.label}" type="checkbox"${nsChecked ? " checked" : ""} style="width:auto" title="新人／留停等整月不排班，該列輸出留白"></td>` +
              `<td><input id="lv_${p.label}" type="text" placeholder="例：3,4,12" value="${val}"></td></tr>`;
    });
    $("leaveTable").innerHTML = html;
  }

  // 讀取預假輸入（排班月的「日期」，如 3,4,12）；於日曆確定後再轉成天序
  function readLeaveDays() {
    const leave = {};
    model.people.filter((p) => p.active).forEach((p) => {
      const el = document.getElementById("lv_" + p.label);
      if (!el || !el.value.trim()) return;
      const days = [];
      el.value.split(/[,，\s]+/).forEach((tok) => {
        const n = parseInt(tok, 10);
        if (!isNaN(n) && n >= 1 && n <= 31) days.push(n);
      });
      if (days.length) leave[p.label] = days;
    });
    return leave;
  }

  // 預假日期 -> 新產生天數（銜接段之後）中相同「日」的天序
  function mapLeaveToFixed(leave) {
    const fixed = {};
    const carryLen = model.carryLen || 0;
    Object.keys(leave).forEach((l) => {
      const days = {};
      leave[l].forEach((n) => {
        for (let i = carryLen; i < model.nDays; i++) {
          if (parseInt(model.dates[i], 10) === n) { days[i] = "OFF"; break; }
        }
      });
      if (Object.keys(days).length) fixed[l] = days;
    });
    return fixed;
  }

  // 解析一行 CSV（支援雙引號包欄位與 "" 跳脫）
  function parseCsvLine(line) {
    const out = [];
    let cur = "", inQ = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (inQ) {
        if (c === '"') { if (line[i + 1] === '"') { cur += '"'; i++; } else inQ = false; }
        else cur += c;
      } else if (c === '"') inQ = true;
      else if (c === ",") { out.push(cur); cur = ""; }
      else cur += c;
    }
    out.push(cur);
    return out;
  }

  // 解析預排假 CSV（格式：姓名,預排假日期 兩欄；日期欄為逗號分隔的日）→ {姓名:[日期,...]}
  function parseLeaveCsv(text) {
    const data = {};
    text.replace(/^\uFEFF/, "").split(/\r?\n/).forEach((line, i) => {
      if (!line.trim()) return;
      const cells = parseCsvLine(line);
      const name = (cells[0] || "").trim();
      if (!name || (i === 0 && /姓名/.test(name))) return; // 略過標題列
      // 日期若未加引號會被拆成多欄，合併第 2 欄之後的所有欄位
      data[name] = cells.slice(1).join(",").split(/[,、\s]+/).filter(Boolean);
    });
    return data;
  }

  // 載入預排假 JSON/CSV（JSON 格式：{"A1":[3,4,12], ...}；CSV 格式：姓名,日期清單），填入預假輸入框
  async function onLeaveJson(file) {
    try {
      const text = await file.text();
      const data = /\.csv$/i.test(file.name) ? parseLeaveCsv(text) : JSON.parse(text);
      if (!data || typeof data !== "object" || Array.isArray(data)) throw new Error('格式應為 {"姓名":[日期,...]}');
      let applied = 0;
      const unknown = [];
      Object.keys(data).forEach((name) => {
        const el = document.getElementById("lv_" + name);
        if (!el) { unknown.push(name); return; }
        const days = (Array.isArray(data[name]) ? data[name] : [data[name]])
          .map((n) => parseInt(n, 10)).filter((n) => !isNaN(n) && n >= 1 && n <= 31);
        el.value = days.join(",");
        applied++;
      });
      let msg = `已載入 ${file.name}：套用 ${applied} 人`;
      if (unknown.length) msg += `；${unknown.length} 個姓名不在名單（略過）：${unknown.slice(0, 8).join("、")}${unknown.length > 8 ? "…" : ""}`;
      $("leaveLoadNote").textContent = msg;
    } catch (e) {
      alert("無法載入預排假檔案：" + e.message);
    }
  }

  function readLocks() {
    const locks = {};
    model.people.filter((p) => p.active).forEach((p) => {
      const el = document.getElementById("lk_" + p.label);
      if (el && el.value) locks[p.label] = el.value;
    });
    return locks;
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
      // 自動推算「下一個週期」並預填
      const CYCLE_START = 10; // 此院排班週期每月 10 號為界（銜接段＝上月 10 號到月底）
      const dim = (y, m) => new Date(y + 1911, m, 0).getDate();
      let ny, nm, nd, winLen, note;
      if (model.detectedStartDay === 1) {
        // 單月表（1 號到月底）：自動補上月銜接，產生「上月 10 號 → 新月月底」的跨月格式
        ny = model.detectedYear; nm = model.detectedStartMonth; nd = CYCLE_START;
        const nextM = nm === 12 ? 1 : nm + 1, nextY = nm === 12 ? ny + 1 : ny;
        winLen = (dim(ny, nm) - CYCLE_START + 1) + dim(nextY, nextM);
        note = `偵測為單月表（${model.detectedStartMonth}月）→ 自動補上月銜接：產生 ${nm}月${nd}日～${nextM}月底，共 ${winLen} 天（下載 Excel 會自動加寬容納銜接段）。`;
      } else {
        // 跨月表：整體往後移一個月（同起始日、同視窗長度）
        ny = model.detectedYear; nm = (model.detectedStartMonth || 1) + 1; nd = model.detectedStartDay || 1;
        if (nm > 12) { nm -= 12; ny = (ny || 0) + 1; }
        winLen = model.nDays;
        note = `偵測上傳週期：${model.detectedYear}年${model.detectedStartMonth}月${model.detectedStartDay}日起 ${model.nDays} 天 → 自動推算下一週期：${ny}年${nm}月${nd}日起。`;
      }
      $("h-year").value = ny || "";
      $("h-month").value = nm;
      $("h-day").value = nd;
      $("winlen").textContent = winLen;
      $("rollNote").textContent = note;
      renderLeaveTable();
      $("leaveCard").classList.remove("hidden");
      $("result").classList.add("hidden");
    } catch (e) {
      alert("無法解析班表：" + e.message);
    }
  }

  function renderGrid(assignment, disp, fixed, flags) {
    const active = model.people.filter((p) => p.active);
    const vioSet = new Set(flags.cellFlags.map((f) => f.label + "|" + f.day));
    const dayVio = flags.dayFlags || {};

    // 銜接段：新視窗與上傳檔重疊的天數（照抄上月實際班別），以灰階呈現。
    // 單月表經 expandTemplateColumns 補上銜接段後，carryLen>0，與跨月表走同一套顯示。
    const carryLen = model.carryLen || 0;

    // 月份帶（最上方）：把連續同月的欄位合併顯示 X月；銜接段以灰色標示
    function segs(arr) {
      const out = [];
      (arr || []).forEach((m) => {
        if (out.length && out[out.length - 1].m === m) out[out.length - 1].n++;
        else out.push({ m, n: 1 });
      });
      return out;
    }
    let band = "<tr class='bandRow'><th class='cName'></th><th class='cGrp'></th>";
    let acc = 0;
    segs(model.monthsByDay).forEach((s) => {
      const isCarry = acc < carryLen;                 // 整段落在銜接段內
      const sep = acc + s.n === carryLen ? " sep" : "";
      const note = isCarry ? " <span style=\"font-weight:400;font-size:10px\">·銜接上月</span>" : "";
      band += `<th class='${isCarry ? "prev" : "mBand"}${sep}' colspan='${s.n}'>${s.m}月${note}</th>`;
      acc += s.n;
    });
    band += "</tr>";

    // 日期/星期列
    let head = "<thead>" + band + "<tr class='headRow'><th class='cName'>姓名</th><th class='cGrp'>階</th>";
    const hol = model.holInfo || [];
    for (let d = 0; d < model.nDays; d++) {
      const carry = d < carryLen ? " carry" : "";
      const sep = d === carryLen - 1 ? " sep" : "";
      let cls = dayVio[d] ? "dayVio" : (hol[d] && hol[d].name ? "holiday" : (hol[d] && hol[d].weekend ? "weekend" : ""));
      const holDiv = hol[d] && hol[d].name ? `<div class='hol'>${hol[d].name}</div>` : "";
      head += `<th class='${cls}${carry}${sep}'${tipAttr(dayVio[d])}>${model.dates[d] != null ? model.dates[d] : d + 1}<br><span style="font-weight:400;color:#94a3b8">${model.weekdays[d] || ""}</span>${holDiv}</th>`;
    }
    head += "</tr></thead>";

    // 表身：依階層分組
    let body = "<tbody>";
    let lastGrp = null;
    const totalCols = 2 + model.nDays;
    active.forEach((p) => {
      if (p.group !== lastGrp) {
        const grpName = p.group === "A" ? "A階（小組長）" : `${p.group}階`;
        body += `<tr class='grpRow'><td class='cName'>${grpName}</td><td class='cGrp'></td><td colspan='${totalCols - 2}'>資深(A) → 資淺(E)</td></tr>`;
        lastGrp = p.group;
      }
      const rowVio = flags.rowFlags[p.label] ? " rowVio" : "";
      body += `<tr><td class='cName${rowVio}'${tipAttr(flags.rowFlags[p.label])}>${p.label}</td><td class='cGrp'>${p.group}</td>`;
      for (let d = 0; d < model.nDays; d++) {
        const text = disp[p.label][d];
        const isLeave = d >= carryLen && fixed[p.label] && fixed[p.label][d] === "OFF";
        let cls = isLeave ? "sLeave" : shiftClass(text);
        if (vioSet.has(p.label + "|" + d)) cls = "vio";
        const carry = d < carryLen ? " carry" : "";       // 銜接段灰階
        const sep = d === carryLen - 1 ? " sep" : "";
        const clk = d >= carryLen ? ` data-l="${p.label}" data-d="${d}"` : ""; // 可點擊切換預假
        const flagMsg = flags.cellFlags.filter((f) => f.label === p.label && f.day === d).map((f) => f.msg).join("；");
        const tip = flagMsg ? tipAttr(flagMsg) : (isLeave ? tipAttr("預先請假（點擊取消並重排）") : "");
        body += `<td class='${cls}${carry}${sep}'${clk}${tip}>${text}</td>`;
      }
      body += "</tr>";
    });
    body += "</tbody>";
    $("grid").innerHTML = head + body;
  }

  /* 人力不足時的預假調整建議：列出缺口日「當天有預假、且可支援該班別」的人選。
   * 點人名 = 取消該員該日預假並自動重排（與點班表格子同一套 toggleLeave）。 */
  function renderSuggestions(assignment, userLeave, rules, locks) {
    const carryLen = model.carryLen || 0;
    const schedDays = model.nDays - carryLen;
    const active = model.people.filter((p) => p.active);
    const kindsMap = {};
    active.forEach((p) => {
      kindsMap[p.label] = new Set((assignment[p.label] || []).slice(carryLen).filter((s) => ["D", "E", "N"].includes(s)));
    });
    // 回來上 s 班是否合法（包班別、至多2種、E+N 組合限制）
    function canTake(p, s) {
      const lk = locks[p.label];
      if (lk) return lk === s;
      const k = kindsMap[p.label];
      if (k.has(s)) return true;
      if (k.size >= rules.maxShiftKinds) return false;
      if (rules.enPairMaxWork != null && s !== "D" && !k.has("D") && k.has(s === "E" ? "N" : "E")) {
        const lv = userLeave[p.label] ? Object.keys(userLeave[p.label]).length : 0;
        if (schedDays - lv > rules.enPairMaxWork) return false; // 會形成不允許的 E+N 組合
      }
      return true;
    }
    const out = [];
    for (let d = carryLen; d < model.nDays; d++) {
      const c = { D: 0, E: 0, N: 0 }, ca = { D: 0, E: 0, N: 0 };
      active.forEach((p) => {
        const v = assignment[p.label][d];
        if (c[v] !== undefined) { c[v]++; if (p.group === "A") ca[v]++; }
      });
      ["D", "E", "N"].forEach((s) => {
        const lackMan = Math.max(0, rules.minManpower[s] - c[s]);
        const lackLead = Math.max(0, rules.minLeaderPerShift - ca[s]);
        if (!lackMan && !lackLead) return;
        const onlyLead = !lackMan && lackLead > 0;
        const cands = active
          .filter((p) => userLeave[p.label] && userLeave[p.label][d] === "OFF")
          .filter((p) => (!onlyLead || p.group === "A") && canTake(p, s))
          .sort((a, b) => {
            if (lackLead) { const ga = a.group === "A" ? 0 : 1, gb = b.group === "A" ? 0 : 1; if (ga !== gb) return ga - gb; }
            const la = locks[a.label] === s ? 0 : 1, lb = locks[b.label] === s ? 0 : 1;
            return la - lb;
          });
        const lackTxt = [lackMan ? `缺 ${lackMan} 人` : "", lackLead ? `A階小組長缺 ${lackLead} 人` : ""].filter(Boolean).join("、");
        const chips = cands.map((p) => {
          const lk = locks[p.label];
          const tag = (p.group === "A" ? "A階·" : "") + (lk ? `包${lk}` : (kindsMap[p.label].size ? Array.from(kindsMap[p.label]).join("+") : "未定"));
          return `<span class="sug" data-l="${p.label}" data-d="${d}" title="點擊取消 ${p.label} 這天的預假並自動重排">${p.label}（${tag}）</span>`;
        }).join("");
        out.push(`<div class="sugRow"><b>${model.monthsByDay[d]}/${model.dates[d]}（${model.weekdays[d]}）${s} 班 ${lackTxt}</b> → 當日預假可支援：${chips || "無（該日預假者皆無法支援此班別，需從其他日調度或改包班設定）"}</div>`);
      });
    }
    $("suggest").innerHTML = out.length
      ? `<div class="sugBox"><b>💡 人力不足：建議調整預排假</b>（點人名即取消該員當日預假並自動重排；點下方班表格子也可加/取消預假）${out.join("")}</div>`
      : "";
  }

  function renderSummary(rows, dayFlags, rules) {
    let html = "<tr><th>日期</th><th>D 白</th><th>E 小夜</th><th>N 大夜</th><th>休假</th></tr>";
    rows.forEach((r) => {
      const flags = dayFlags[r.dayIdx];
      // 人力不足的「班別」格單獨標紅並註明缺額
      const cells = ["D", "E", "N"].map((s) => {
        const need = rules.minManpower[s];
        if (r[s] < need)
          return `<td class='dayVio'${tipAttr(`${s}班僅 ${r[s]} 人，低於最低人力 ${need} 人（缺 ${need - r[s]} 人）`)}>${r[s]} ⚠</td>`;
        return `<td>${r[s]}</td>`;
      }).join("");
      const vio = flags ? " class='dayVio'" : "";
      html += `<tr${vio}${tipAttr(flags)}><td>${r.label}</td>${cells}<td>${r.off}</td></tr>`;
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
      const leave = readLeaveDays();
      const locks = readLocks();
      const wb = await loadWorkbook(arrayBuffer);
      const ws = wb.worksheets[0];
      model = Scheduler.parseWorkbook(ws);

      // 「不排班」勾選（新人／留停等）：整列排除排班與人力計算，輸出留白
      model.people.forEach((p) => {
        const el = document.getElementById("ns_" + p.label);
        if (el && el.checked) p.active = false;
      });

      // 依起始真實日期建立日曆 + 抓取假日
      const hy = { rocYear: $("h-year").value, month: $("h-month").value, day: $("h-day").value };

      // 單月表：擴充模板欄位以容納「上月銜接段 + 新月」，之後與跨月表走同一套銜接段流程
      const origNDays = model.nDays;                    // 上傳檔原始天數（buildCarry 的 upCal 用）
      if (model.detectedStartDay === 1) {
        const sy = parseInt(hy.rocYear) || model.detectedYear;
        const sm = parseInt(hy.month) || model.detectedStartMonth;
        const sd = parseInt(hy.day) || model.detectedStartDay;
        const nextM = sm === 12 ? 1 : sm + 1, nextY = sm === 12 ? sy + 1 : sy;
        const targetDays = Math.round((new Date(nextY + 1911, nextM, 0) - new Date(sy + 1911, sm - 1, sd)) / 864e5) + 1;
        Scheduler.expandTemplateColumns(ws, model, targetDays - origNDays);
      }
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

      // 銜接段：新視窗與上傳檔重疊的日期，照抄上傳檔實際班別（不重排、不報違規）。
      // upCal 用上傳檔原始天數（單月表擴欄後 model.nDays 已變大，但上傳資料仍是原始天數）。
      const upCal = Scheduler.buildCalendar(model.detectedYear, model.detectedStartMonth, model.detectedStartDay, origNDays);
      const carry = Scheduler.buildCarry(model, upCal, cal, tailCal);
      model.carryLen = carry.overlapLen;
      model.people.forEach((p) => { if (carry.history[p.label]) p.tail = carry.history[p.label]; });

      const userLeave = mapLeaveToFixed(leave); // 預假(排班月日期) -> 天序
      const fixed = {};
      Object.keys(userLeave).forEach((l) => (fixed[l] = Object.assign({}, userLeave[l])));
      Object.keys(carry.preset).forEach((l) => {
        fixed[l] = Object.assign({}, fixed[l] || {});
        Object.assign(fixed[l], carry.preset[l]); // 銜接段優先於使用者預假
      });

      const t0 = performance.now();
      const assignment = Scheduler.generate(model, rules, fixed, locks);
      const ms = Math.round(performance.now() - t0);
      const flags = Scheduler.validate(model, assignment, rules, locks);

      const banner = $("banner");
      if (flags.passed) {
        banner.className = "banner ok";
        banner.textContent = `✅ 排班完成，通過所有硬規則驗證（${ms} ms）`;
        $("errbox").innerHTML = "";
      } else {
        banner.className = "banner warn";
        const byKind = {};
        flags.errors.forEach((e) => { const m = e.match(/^\[([^\]]+)\]/); const k = m ? m[1] : "其他"; byKind[k] = (byKind[k] || 0) + 1; });
        const kindTxt = Object.keys(byKind).map((k) => `${k}×${byKind[k]}`).join("、");
        banner.textContent = `⚠️ 已產生班表，但有 ${flags.errors.length} 項規則未滿足：${kindTxt}（滑鼠移到紅色格子/日期可查看細節）`;
        $("errbox").innerHTML =
          "<details><summary>違規明細（點開）</summary>" +
          flags.errors.slice(0, 120).map((e) => `<div class="err">• ${e}</div>`).join("") + "</details>";
      }

      const disp = {};
      Object.keys(assignment).forEach((l) => (disp[l] = Scheduler.labelOffs(assignment[l], rules)));
      // 銜接段顯示上傳檔原始代碼（特/國/公…）
      Object.keys(disp).forEach((l) => {
        const raw = carry.rawByDay[l];
        if (!raw) return;
        for (let i = 0; i < carry.overlapLen; i++) disp[l][i] = raw[i] || "";
      });
      renderGrid(assignment, disp, fixed, flags);
      renderSummary(Scheduler.summarizeDaily(model, assignment), flags.dayFlags, rules);
      renderSuggestions(assignment, userLeave, rules, locks);

      // 寫回範本供下載（真實日曆：日期/星期/假日/標題）
      const { title } = Scheduler.writeCalendar(ws, model, {
        hospital: $("h-hospital").value,
        rocYear: hy.rocYear, month: hy.month, day: hy.day,
        holidays: model._holMap,
      });
      Scheduler.writeInto(ws, model, assignment, rules, carry.rawByDay, userLeave);
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
    $("leaveLoadNote").textContent = "";
  });
  $("loadLeave").addEventListener("click", () => $("leaveFile").click());
  $("leaveFile").addEventListener("change", async (e) => {
    if (e.target.files[0]) await onLeaveJson(e.target.files[0]);
    e.target.value = ""; // 允許重選同一檔案
  });

  // 點擊班表格子 / 建議面板人名：切換該員該日預排假並自動重排（銜接段不可點）
  let toggling = false;
  async function toggleLeave(label, dayIdx) {
    if (!model || toggling) return;
    const carryLen = model.carryLen || 0;
    if (isNaN(dayIdx) || dayIdx < carryLen || dayIdx >= model.nDays) return;
    const el = document.getElementById("lv_" + label);
    if (!el) return;
    const dateNum = parseInt(model.dates[dayIdx], 10);
    const days = [];
    el.value.split(/[,，\s]+/).forEach((t) => {
      const n = parseInt(t, 10);
      if (!isNaN(n) && n >= 1 && n <= 31 && days.indexOf(n) < 0) days.push(n);
    });
    const at = days.indexOf(dateNum);
    const added = at < 0;
    if (added) days.push(dateNum); else days.splice(at, 1);
    el.value = days.sort((a, b) => a - b).join(",");
    toggling = true;
    try {
      $("leaveLoadNote").textContent = `${added ? "➕ 新增" : "➖ 取消"} ${label} ${model.monthsByDay[dayIdx]}/${dateNum} 預假，重排中…`;
      await generateAndRender();
      $("leaveLoadNote").textContent = `${added ? "➕ 已新增" : "➖ 已取消"} ${label} ${model.monthsByDay[dayIdx]}/${dateNum} 預假並完成重排`;
    } finally { toggling = false; }
  }
  document.addEventListener("click", (e) => {
    const t = e.target.closest ? e.target.closest("[data-l][data-d]") : null;
    if (!t) return;
    toggleLeave(t.getAttribute("data-l"), parseInt(t.getAttribute("data-d"), 10));
  });
})();
