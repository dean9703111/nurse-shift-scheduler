/* 護理師排班 — 純前端核心邏輯（無框架，Node 與瀏覽器共用）。
 * 依賴 ExcelJS 的 worksheet 物件（Node 與瀏覽器 API 一致）。
 * 匯出：parseWorkbook, generate, validate, writeInto, summarizeDaily, DEFAULT_RULES
 */
(function (root) {
  "use strict";

  const SHIFTS = ["D", "E", "N"];
  const GROUPS = ["A", "B", "C", "D", "E"];
  const TIER_START = ["B", "C", "D", "E"];
  const PERSON_RE = /^[A-Ea-e]\d+$/;
  const TAIL_LEN = 7; // 顯示上月最後 7 天

  const DEFAULT_RULES = {
    minManpower: { D: 16, E: 16, N: 14 },
    maxConsecutiveWork: 5,
    maxConsecutiveN: 4,
    minLeaderPerShift: 2,
    tierCaps: { D: [14, 11, 8, 4], E: [14, 11, 8, 4], N: [12, 9, 6, 3] },
    // 更換班別需連續 11 小時休息（D 08-16 / E 16-24 / N 00-08 值到隔日早上）
    // → 禁止的「前一天班別 -> 今日班別」組合：N 下班後接 D(0h)/E(8h)、E 下班後接 D(8h)
    forbidTransition: { N: ["D", "E"], E: ["D"] },
    maxShiftKinds: 2, // 未包班人員每月至多 2 種班別
    lockThreshold: 15, // 同班別 >15 天視為包班（純班）
    restWindows: [[14, 2], [28, 8]],
    statutoryPeriod: 7,
    offStatutory: "例",
    offRest: "息",
  };

  function cellText(v) {
    if (v === null || v === undefined) return "";
    if (typeof v === "object") {
      if (v.result !== undefined && v.result !== null) return String(v.result);
      if (v.text !== undefined) return String(v.text);
      if (v.richText) return v.richText.map((t) => t.text).join("");
      if (v.formula !== undefined) return "";
      return "";
    }
    return String(v).trim();
  }

  function classifyCell(v) {
    const s = cellText(v);
    if (!s || s.startsWith("=")) return null;
    if (["例", "息", "國", "特", "補", "公", "離職"].includes(s)) return null;
    const up = s.toUpperCase();
    for (const sh of SHIFTS) if (up.includes(sh)) return sh;
    return null;
  }

  function groupsFrom(start) {
    return GROUPS.slice(GROUPS.indexOf(start));
  }

  function parseWorkbook(ws, tailLen) {
    tailLen = tailLen || TAIL_LEN;
    let dateRow = null;
    for (let r = 1; r <= 12; r++) {
      if (cellText(ws.getCell(r, 3).value) === "日期") { dateRow = r; break; }
    }
    if (dateRow === null) throw new Error("找不到『日期』列，無法辨識表格結構。");
    const weekdayRow = dateRow + 1;

    const maxCol = ws.columnCount || 200;
    const dateColsAll = [];
    for (let c = 1; c <= maxCol; c++) {
      const v = ws.getCell(dateRow, c).value;
      if (typeof v === "number" && Number.isFinite(v)) dateColsAll.push(c);
    }
    if (!dateColsAll.length) throw new Error("日期列沒有偵測到任何日期數字。");
    const firstCol = dateColsAll[0];
    let nDays = 1;
    while (dateColsAll.includes(firstCol + nDays)) nDays++;

    const dateCols = Array.from({ length: nDays }, (_, i) => firstCol + i);
    const dates = dateCols.map((c) => cellText(ws.getCell(dateRow, c).value));
    const weekdays = dateCols.map((c) => cellText(ws.getCell(weekdayRow, c).value));

    // 標題與年月偵測
    const monthLabelRow = dateRow - 1;
    // 標題：第 1 列中含「班表」的儲存格（通常為合併儲存格的主格）
    let titleCell = null;
    for (let c = 1; c <= maxCol; c++) {
      if (cellText(ws.getCell(1, c).value).includes("班表")) { titleCell = { r: 1, c }; break; }
    }
    // 月份區塊起點欄：第一欄 + 日期為 1 的欄（跨月處）
    const monthStartCols = [firstCol];
    for (let i = 1; i < nDays; i++) {
      if (cellText(ws.getCell(dateRow, dateCols[i]).value) === "1") monthStartCols.push(dateCols[i]);
    }
    // 從範本推測預設年月
    const titleText = titleCell ? cellText(ws.getCell(titleCell.r, titleCell.c).value) : "";
    const yM = titleText.match(/(\d+)\s*年/);
    const detectedYear = yM ? parseInt(yM[1], 10) : null;
    const firstLabel = monthLabelRow >= 1 ? cellText(ws.getCell(monthLabelRow, firstCol).value) : "";
    const mM = firstLabel.match(/(\d+)\s*月/);
    const detectedStartMonth = mM ? parseInt(mM[1], 10) : null;
    const hM = titleText.match(/^\s*(.*?)\s*\d+\s*年/);
    const detectedHospital = hM ? hM[1] : "";
    const detectedStartDay = parseInt(dates[0], 10) || 1;

    // 上月最後 tailLen 天的日期/星期（供對照顯示）
    const tailCols = [];
    for (let k = tailLen; k >= 1; k--) tailCols.push(firstCol + nDays - k);
    const tailDates = tailCols.map((c) => cellText(ws.getCell(dateRow, c).value));
    const tailWeekdays = tailCols.map((c) => cellText(ws.getCell(weekdayRow, c).value));

    const people = [];
    const maxRow = ws.rowCount || 300;
    // 包班偵測範圍：只看上傳檔「最後一個月份段」（最後一個日期=1 之後），避免把銜接段算進去
    let lockFrom = 0;
    for (let i = nDays - 1; i >= 1; i--) {
      if (cellText(ws.getCell(dateRow, dateCols[i]).value) === "1") { lockFrom = i; break; }
    }
    for (let r = dateRow + 1; r <= maxRow; r++) {
      const b = ws.getCell(r, 2).value;
      const c = cellText(ws.getCell(r, 3).value);
      if (typeof b === "number" && PERSON_RE.test(c)) {
        const label = c.toUpperCase();
        const group = label[0];
        const tail = tailCols.map((col) => classifyCell(ws.getCell(r, col).value)); // D/E/N/null
        const rawTail = tailCols.map((col) => cellText(ws.getCell(r, col).value)); // 原始顯示文字
        const rowRaw = dateCols.map((col) => cellText(ws.getCell(r, col).value)); // 整月原始班別
        // 依上月班表偵測包班（同班別 > lockThreshold 天）
        const shiftCnt = { D: 0, E: 0, N: 0 };
        dateCols.slice(lockFrom).forEach((col) => { const s = classifyCell(ws.getCell(r, col).value); if (s) shiftCnt[s]++; });
        let suggestedLock = null;
        for (const sh of SHIFTS) if (shiftCnt[sh] > DEFAULT_RULES.lockThreshold) suggestedLock = sh;
        let active = true;
        for (let i = 0; i < nDays; i++) {
          if (cellText(ws.getCell(r, firstCol + i).value).includes("離職")) { active = false; break; }
        }
        people.push({ row: r, label, group, active, tail, rawTail, rowRaw, shiftCnt, suggestedLock });
      }
    }
    if (!people.length) throw new Error("找不到任何人員列。");

    // 匯入檔「第一個月段」(起始日~該月底)：作為網頁灰色對照區塊
    let greyLen = nDays;
    for (let i = 1; i < nDays; i++) { if (String(dates[i]) === "1") { greyLen = i; break; } }
    const greyDates = dates.slice(0, greyLen);
    const greyWeekdays = weekdays.slice(0, greyLen);
    const greyMonth = detectedStartMonth;

    return { dateRow, weekdayRow, firstCol, nDays, people, dateCols,
             dates, weekdays, tailLen, tailDates, tailWeekdays,
             greyLen, greyDates, greyWeekdays, greyMonth,
             monthLabelRow, titleCell, monthStartCols,
             detectedYear, detectedStartMonth, detectedStartDay, detectedHospital };
  }

  /* ---------- 真實日曆 ---------- */
  const WEEK = ["日", "一", "二", "三", "四", "五", "六"];
  function rocToGreg(y) { return y + 1911; }
  function keyOf(dt) {
    return "" + dt.getFullYear() + String(dt.getMonth() + 1).padStart(2, "0") + String(dt.getDate()).padStart(2, "0");
  }
  // 由民國年/月/日起算 nDays 天的真實日曆
  function buildCalendar(rocYear, month, day, nDays) {
    const start = new Date(rocToGreg(rocYear), month - 1, day);
    const out = [];
    for (let i = 0; i < nDays; i++) {
      const dt = new Date(start); dt.setDate(start.getDate() + i);
      out.push({ y: dt.getFullYear(), roc: dt.getFullYear() - 1911, m: dt.getMonth() + 1, d: dt.getDate(), w: WEEK[dt.getDay()], key: keyOf(dt) });
    }
    return out;
  }
  // 起始日之前 tailLen 天（最舊 -> 最新）
  function buildTail(rocYear, month, day, tailLen) {
    const start = new Date(rocToGreg(rocYear), month - 1, day);
    const out = [];
    for (let k = tailLen; k >= 1; k--) {
      const dt = new Date(start); dt.setDate(start.getDate() - k);
      out.push({ y: dt.getFullYear(), roc: dt.getFullYear() - 1911, m: dt.getMonth() + 1, d: dt.getDate(), w: WEEK[dt.getDay()], key: keyOf(dt) });
    }
    return out;
  }

  /* 以上傳檔的「實際班別」建立新視窗的銜接資料。
   * uploadedCal：上傳檔視窗的真實日曆（與 people[].rowRaw 逐欄對齊）
   * newCal / tailCal：新視窗日曆與其起始日前 tailLen 天
   * 回傳 {
   *   overlapLen,  // 新視窗前段與上傳檔重疊的天數（銜接段長度）
   *   preset,      // { label: { dayIdx: 'D'|'E'|'N'|'OFF' } } 銜接段預填
   *   rawByDay,    // { label: { dayIdx: 原始代碼(特/國/公…) } } 供顯示與寫回
   *   history,     // { label: [tailLen 天 D/E/N/null] } 新視窗起始日前的實際班別
   * }
   */
  function buildCarry(model, uploadedCal, newCal, tailCal) {
    const idxByKey = {};
    uploadedCal.forEach((c, i) => { idxByKey[c.key] = i; });
    let overlapLen = 0;
    for (let i = 0; i < newCal.length; i++) {
      if (idxByKey[newCal[i].key] !== undefined) overlapLen = i + 1;
      else break;
    }
    const preset = {}, rawByDay = {}, history = {};
    model.people.forEach((p) => {
      const ps = {}, raw = {};
      for (let i = 0; i < overlapLen; i++) {
        const ui = idxByKey[newCal[i].key];
        const t = p.rowRaw[ui] || "";
        ps[i] = classifyCell(t) || "OFF";
        raw[i] = t;
      }
      preset[p.label] = ps;
      rawByDay[p.label] = raw;
      history[p.label] = tailCal.map((c) => {
        const ui = idxByKey[c.key];
        return ui === undefined ? null : classifyCell(p.rowRaw[ui] || "");
      });
    });
    return { overlapLen, preset, rawByDay, history };
  }

  // 內建國定假日（名稱，YYYYMMDD -> 說明）；能連網時由 app 以最新資料覆蓋。
  const BUILTIN_HOLIDAYS = {
    "20260101": "開國紀念日", "20260215": "小年夜", "20260216": "農曆除夕",
    "20260217": "春節", "20260218": "春節", "20260219": "春節", "20260220": "補假",
    "20260227": "補假", "20260228": "和平紀念日", "20260403": "補假", "20260404": "兒童節",
    "20260405": "清明節", "20260406": "補假", "20260501": "勞動節", "20260619": "端午節",
    "20260925": "中秋節", "20260928": "孔子誕辰紀念日/教師節", "20261009": "補假",
    "20261010": "國慶日", "20261025": "臺灣光復暨金門古寧頭大捷紀念日", "20261026": "補假",
    "20261225": "行憲紀念日",
  };
  function holidayName(key, holidays) {
    if (holidays && holidays[key]) return holidays[key];
    return BUILTIN_HOLIDAYS[key] || "";
  }

  /* 依起始真實日期，重寫日期列、星期列、月份標記與國定假日標註，並更新標題。
   * opts: { hospital, rocYear, month, day, holidays }
   * 回傳 { title, cal }
   */
  function writeCalendar(ws, model, opts) {
    opts = opts || {};
    const hospital = opts.hospital != null ? opts.hospital : (model.detectedHospital || "");
    const rocYear = parseInt(opts.rocYear, 10) || model.detectedYear || 0;
    const month = parseInt(opts.month, 10) || model.detectedStartMonth || 1;
    const day = parseInt(opts.day, 10) || model.detectedStartDay || 1;
    const cal = buildCalendar(rocYear, month, day, model.nDays);

    // 清空月份標記列（僅日期欄範圍），再重寫
    if (model.monthLabelRow >= 1) model.dateCols.forEach((c) => (ws.getCell(model.monthLabelRow, c).value = null));

    // 日期列 + 星期列；日期列底色依月份自動標示：上月銜接段灰底、本月藍底（同範本慣例）
    const FILL_DATE_PREV = { type: "pattern", pattern: "solid", fgColor: { theme: 2 } };
    const FILL_DATE_CUR = { type: "pattern", pattern: "solid", fgColor: { argb: "FFCCCCFF" } };
    const carryLen = model.carryLen || 0;
    cal.forEach((c, i) => {
      const col = model.dateCols[i];
      const dCell = ws.getCell(model.dateRow, col);
      dCell.value = c.d;
      dCell.style = Object.assign({}, dCell.style, { fill: i < carryLen ? FILL_DATE_PREV : FILL_DATE_CUR });
      ws.getCell(model.weekdayRow, col).value = c.w;
    });

    // 月份標記 + 假日標註
    cal.forEach((c, i) => {
      const col = model.dateCols[i];
      const parts = [];
      if (i === 0 || c.d === 1) parts.push(c.m + "月");
      const hn = holidayName(c.key, opts.holidays);
      if (hn) parts.push(hn);
      if (parts.length && model.monthLabelRow >= 1) ws.getCell(model.monthLabelRow, col).value = parts.join(" ");
    });

    // 底部驗證區也有自己的「日期」列（供 COUNTIF 對照），一併同步成新日期
    const maxRow = ws.rowCount || 300;
    for (let r = model.dateRow + 1; r <= maxRow; r++) {
      if (cellText(ws.getCell(r, 3).value) === "日期") {
        cal.forEach((c, i) => { ws.getCell(r, model.dateCols[i]).value = c.d; });
      }
    }

    // 標題：取涵蓋天數最多的月份
    const spans = {};
    cal.forEach((c) => { const k = c.roc * 100 + c.m; spans[k] = (spans[k] || 0) + 1; });
    let best = -1, titleY = rocYear, titleM = month;
    Object.keys(spans).forEach((k) => { if (spans[k] > best) { best = spans[k]; titleY = Math.floor(k / 100); titleM = k % 100; } });
    const title = `${hospital}${titleY}年${titleM}月護理師班表`;
    if (model.titleCell) ws.getCell(model.titleCell.r, model.titleCell.c).value = title;
    return { title, cal };
  }

  /* 啟發式排班。
   * fixed: { label: { dayIndex: 'OFF'|'D'|'E'|'N' } } —— 預先請假(OFF) 或銜接段實際班別。
   * locks: { label: 'D'|'E'|'N' } —— 包班人員（只排該班別）。
   * model.carryLen（選填）：前 carryLen 天為銜接段，直接沿用 fixed 的班別，不重新排班。
   * 回傳 { label: [D/E/N/null,...] }
   */
  function generate(model, rules, fixed, locks) {
    rules = rules || DEFAULT_RULES;
    fixed = fixed || {};
    locks = locks || {};
    const carryLen = model.carryLen || 0;
    const people = model.people.filter((p) => p.active);
    const N = model.nDays;
    const labels = people.map((p) => p.label);
    const byLabel = {};
    people.forEach((p) => (byLabel[p.label] = p));

    const fixedOff = {}; // label -> Set(day)
    const fixedShift = {}; // label -> { day: 'D'|'E'|'N' }（銜接段實際班別）
    labels.forEach((l) => { fixedOff[l] = new Set(); fixedShift[l] = {}; });
    Object.keys(fixed).forEach((l) => {
      if (!fixedOff[l]) return;
      Object.keys(fixed[l]).forEach((d) => {
        const di = parseInt(d, 10);
        if (di < 0 || di >= N) return;
        const v = fixed[l][d];
        if (v === "OFF") fixedOff[l].add(di);
        else if (SHIFTS.includes(v)) fixedShift[l][di] = v;
      });
    });

    const assigned = {};
    const consecWork = {}, consecN = {}, lastShift = {}, workCount = {}, kindsUsed = {};
    labels.forEach((l) => {
      assigned[l] = new Array(N).fill(null);
      kindsUsed[l] = new Set();
      const t = byLabel[l].tail;
      let cw = 0; for (let i = t.length - 1; i >= 0; i--) { if (t[i]) cw++; else break; }
      let cn = 0; for (let i = t.length - 1; i >= 0; i--) { if (t[i] === "N") cn++; else break; }
      consecWork[l] = cw; consecN[l] = cn;
      lastShift[l] = t.length ? t[t.length - 1] : null;
      workCount[l] = 0;
    });

    for (let d = 0; d < N; d++) {
      if (d < carryLen) {
        // 銜接段：照抄上傳檔的實際班別，只更新連續狀態；
        // 不參與人力補齊、不計入本月班別種類與工作量
        labels.forEach((l) => {
          const s = fixedShift[l][d] || null;
          assigned[l][d] = s;
          if (s) {
            consecWork[l] += 1;
            consecN[l] = s === "N" ? consecN[l] + 1 : 0;
            lastShift[l] = s;
          } else {
            consecWork[l] = 0; consecN[l] = 0; lastShift[l] = null;
          }
        });
        continue;
      }
      const todayShift = {};
      const cnt = { D: 0, E: 0, N: 0 };
      const leaderCnt = { D: 0, E: 0, N: 0 };
      const tierCnt = { D: [0, 0, 0, 0], E: [0, 0, 0, 0], N: [0, 0, 0, 0] };
      // 休假保底：若某視窗內「已休天數 + 今天之後還能休的天數」不足下限，今天必須休，
      // 否則該視窗註定違規。只需檢查以今天為尾端的視窗（其餘視窗餘裕更大）。
      function mustRest(l) {
        for (const [w, mo] of rules.restWindows) {
          if (w > N || !mo) continue;
          const s = Math.max(0, d - w + 1);
          let k = 0;
          for (let i = s; i < d; i++) if (assigned[l][i] === null) k++;
          if (k + (s + w - 1 - d) < mo) return true;
        }
        return false;
      }
      const forcedOff = {};
      labels.forEach((l) => {
        forcedOff[l] = consecWork[l] >= rules.maxConsecutiveWork || fixedOff[l].has(d) || mustRest(l);
      });

      function tierIndexesFor(group) {
        const idx = [];
        TIER_START.forEach((st, i) => { if (groupsFrom(st).includes(group)) idx.push(i); });
        return idx;
      }
      function capOK(group, s) {
        const caps = rules.tierCaps[s];
        for (const i of tierIndexesFor(group)) if (tierCnt[s][i] + 1 > caps[i]) return false;
        return true;
      }
      function eligible(l, s) {
        if (todayShift[l] || forcedOff[l]) return false;
        const prev = d === 0 ? byLabel[l].tail[byLabel[l].tail.length - 1] : assigned[l][d - 1];
        if (prev && (rules.forbidTransition[prev] || []).includes(s)) return false;
        if (s === "N" && consecN[l] >= rules.maxConsecutiveN) return false;
        const lock = locks[l];
        if (lock && s !== lock) return false; // 包班：只排包定班別
        if (!lock && !kindsUsed[l].has(s) && kindsUsed[l].size >= rules.maxShiftKinds) return false; // 至多 2 種班別
        if (!capOK(byLabel[l].group, s)) return false;
        return true;
      }
      function cmpFor(s) {
        // 排序原則：工作量少者最優先（讓所有人休假天數盡量拉平，目標差距 1~2 天內），
        // 同工作量才考慮「不必開第二種班別者」與資淺者。
        // 包班在 eligible 中限制班別即可，不給優先權——否則包班者會被排滿、未包班者吃剩。
        return (a, b) => {
          if (workCount[a] !== workCount[b]) return workCount[a] - workCount[b]; // 工作量平均
          // 已上過該班、或還沒有任何班種者優先；避免懲罰月初尚未排班的人
          const ka = kindsUsed[a].has(s) || kindsUsed[a].size === 0 ? 0 : 1;
          const kb = kindsUsed[b].has(s) || kindsUsed[b].size === 0 ? 0 : 1;
          if (ka !== kb) return ka - kb;
          const ga = GROUPS.indexOf(byLabel[a].group), gb = GROUPS.indexOf(byLabel[b].group);
          if (ga !== gb) return gb - ga; // 同分時資淺優先，資深保留給小組長缺口
          const pa = lastShift[a] === s ? 0 : 1, pb = lastShift[b] === s ? 0 : 1;
          return pa - pb;
        };
      }
      function doAssign(l, s) {
        todayShift[l] = s; cnt[s]++;
        kindsUsed[l].add(s);
        if (byLabel[l].group === "A") leaderCnt[s]++;
        for (const i of tierIndexesFor(byLabel[l].group)) tierCnt[s][i]++;
      }

      for (const s of ["N", "D", "E"]) {
        while (leaderCnt[s] < rules.minLeaderPerShift) {
          const cands = labels.filter((l) => byLabel[l].group === "A" && eligible(l, s));
          if (!cands.length) break;
          cands.sort(cmpFor(s));
          doAssign(cands[0], s);
        }
      }
      const order = ["N", "D", "E"].sort((a, b) => rules.minManpower[b] - rules.minManpower[a]);
      for (const s of order) {
        while (cnt[s] < rules.minManpower[s]) {
          let cands = labels.filter((l) => eligible(l, s));
          if (!cands.length) break;
          cands.sort(cmpFor(s));
          doAssign(cands[0], s);
        }
      }

      labels.forEach((l) => {
        const s = todayShift[l] || null;
        assigned[l][d] = s;
        if (s) {
          consecWork[l] += 1;
          consecN[l] = s === "N" ? consecN[l] + 1 : 0;
          lastShift[l] = s;
          workCount[l] += 1;
        } else {
          consecWork[l] = 0; consecN[l] = 0; lastShift[l] = null;
        }
      });
    }
    return assigned;
  }

  /* 驗證器。回傳結構化違規供上色：
   * { passed, errors[], cellFlags:[{label,day,msg}], dayFlags:{day:[msg]}, rowFlags:{label:[msg]} }
   */
  function validate(model, assignment, rules, locks) {
    rules = rules || DEFAULT_RULES;
    locks = locks || {};
    const carryLen = model.carryLen || 0; // 銜接段為既成事實，不對其報違規
    const people = model.people.filter((p) => p.active);
    const N = model.nDays;
    // 違規訊息以真實日期呈現（app 有設定日曆時），否則退回「第N天」
    const dayLabel = (d) => (model.monthsByDay && model.dates)
      ? `${model.monthsByDay[d]}/${model.dates[d]}`
      : `第${d + 1}天`;
    const errors = [];
    const cellFlags = [];
    const dayFlags = {};
    const rowFlags = {};
    const grp = {}; people.forEach((p) => (grp[p.label] = p.group));
    const tail = {}; people.forEach((p) => (tail[p.label] = p.tail));
    const addDay = (d, msg) => { (dayFlags[d] = dayFlags[d] || []).push(msg); errors.push(msg); };
    const addCell = (l, d, msg) => { cellFlags.push({ label: l, day: d, msg }); errors.push(msg); };
    const addRow = (l, msg) => { (rowFlags[l] = rowFlags[l] || []).push(msg); errors.push(msg); };

    function counts(day, labs) {
      const c = { D: 0, E: 0, N: 0 };
      labs.forEach((l) => { const v = assignment[l][day]; if (SHIFTS.includes(v)) c[v]++; });
      return c;
    }
    const allLabels = people.map((p) => p.label);
    const aLabels = people.filter((p) => p.group === "A").map((p) => p.label);

    for (let d = carryLen; d < N; d++) {
      const c = counts(d, allLabels);
      for (const s of SHIFTS)
        if (c[s] < rules.minManpower[s]) addDay(d, `[最低人力] ${dayLabel(d)} ${s}班 僅 ${c[s]} 人 (<${rules.minManpower[s]})`);
      const ca = counts(d, aLabels);
      for (const s of SHIFTS)
        if (ca[s] < rules.minLeaderPerShift) addDay(d, `[小組長] ${dayLabel(d)} ${s}班 A階僅 ${ca[s]} 人`);
      for (const s of SHIFTS) {
        rules.tierCaps[s].forEach((cap, i) => {
          const labs = allLabels.filter((l) => groupsFrom(TIER_START[i]).includes(grp[l]));
          const cc = counts(d, labs)[s];
          if (cc > cap) addDay(d, `[階層上限] ${dayLabel(d)} ${s}班 ${TIER_START[i]}階以下 ${cc} 人 (>${cap})`);
        });
      }
    }

    people.forEach((p) => {
      const lab = p.label, seq = assignment[lab], h = tail[lab], base = h.length;
      const prevLast = h.length ? h[h.length - 1] : null;
      if (carryLen === 0 && prevLast && seq[0] && (rules.forbidTransition[prevLast] || []).includes(seq[0]))
        addCell(lab, 0, `[換班休息] ${lab} 上月末(${prevLast})後首日排${seq[0]}（休息不足11小時）`);
      for (let d = Math.max(1, carryLen); d < N; d++)
        if (seq[d] && seq[d - 1] && (rules.forbidTransition[seq[d - 1]] || []).includes(seq[d]))
          addCell(lab, d, `[換班休息] ${lab} ${dayLabel(d)}排${seq[d]} 前一天為${seq[d - 1]}（休息不足11小時）`);

      // 包班純班 / 未包班至多 2 種班別（僅計新產生的天數）
      const lock = locks[lab];
      if (lock) {
        for (let d = carryLen; d < N; d++)
          if (SHIFTS.includes(seq[d]) && seq[d] !== lock)
            addCell(lab, d, `[包班] ${lab} 包${lock}班，${dayLabel(d)}卻排${seq[d]}`);
      } else {
        const kinds = new Set(seq.slice(carryLen).filter((s) => SHIFTS.includes(s)));
        if (kinds.size > rules.maxShiftKinds)
          addRow(lab, `[班別種類] ${lab} 本月使用 ${kinds.size} 種班別（${Array.from(kinds).join("/")}）(>${rules.maxShiftKinds})`);
      }

      const wseq = h.map((c) => (SHIFTS.includes(c) ? 1 : 0)).concat(seq.map((c) => (SHIFTS.includes(c) ? 1 : 0)));
      let run = 0;
      for (let i = 0; i < wseq.length; i++) {
        run = wseq[i] ? run + 1 : 0;
        if (run > rules.maxConsecutiveWork && i >= base + carryLen) { addCell(lab, i - base, `[連續上班] ${lab} 連續上班超過 ${rules.maxConsecutiveWork} 天`); break; }
      }
      const nseq = h.map((c) => (c === "N" ? 1 : 0)).concat(seq.map((c) => (c === "N" ? 1 : 0)));
      run = 0;
      for (let i = 0; i < nseq.length; i++) {
        run = nseq[i] ? run + 1 : 0;
        if (run > rules.maxConsecutiveN && i >= base + carryLen) { addCell(lab, i - base, `[連續大夜] ${lab} 連續大夜超過 ${rules.maxConsecutiveN} 天`); break; }
      }
      const off = seq.map((c) => (SHIFTS.includes(c) ? 0 : 1));
      for (const [w, mo] of rules.restWindows) {
        if (w <= N) {
          // 只檢查「尾端落在新產生天數」的視窗；完全落在銜接段內的屬上月既成事實
          for (let st = Math.max(0, carryLen - w + 1); st <= N - w; st++) {
            let sum = 0; for (let i = st; i < st + w; i++) sum += off[i];
            if (sum < mo) { addRow(lab, `[休假] ${lab} ${dayLabel(st)}~${dayLabel(st + w - 1)} 休假 <${mo} 日`); break; }
          }
        }
      }
    });

    return { passed: errors.length === 0, errors, cellFlags, dayFlags, rowFlags };
  }

  function labelOffs(seq, rules) {
    rules = rules || DEFAULT_RULES;
    const out = [];
    const taken = {};
    for (let d = 0; d < seq.length; d++) {
      const code = seq[d];
      if (SHIFTS.includes(code)) { out.push(code); continue; }
      const block = Math.floor(d / rules.statutoryPeriod);
      if (!taken[block]) { out.push(rules.offStatutory); taken[block] = true; }
      else out.push(rules.offRest);
    }
    return out;
  }

  /* carryRaw（選填）: { label: { dayIdx: 原始代碼 } } —— 銜接段寫回上傳檔原始內容。
   * userLeave（選填）: { label: { dayIdx: 'OFF' } } —— 使用者預排假。
   * 底色規則（同醫院慣例）：僅「本月預排假」標黃；其餘（含上月銜接段）一律清除底色，
   * 避免範本殘留的舊黃格跟新班表對不上。 */
  const FILL_LEAVE = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
  const FILL_NONE = { type: "pattern", pattern: "none" };
  function writeInto(ws, model, assignment, rules, carryRaw, userLeave) {
    rules = rules || DEFAULT_RULES;
    model.people.forEach((p) => {
      if (assignment[p.label]) {
        const disp = labelOffs(assignment[p.label], rules);
        const raw = carryRaw && carryRaw[p.label];
        const leave = userLeave && userLeave[p.label];
        model.dateCols.forEach((c, i) => {
          const isCarry = raw && raw[i] !== undefined;
          const v = isCarry ? (raw[i] || null) : disp[i];
          const cell = ws.getCell(p.row, c);
          cell.value = v;
          const isLeave = !isCarry && leave && leave[i] === "OFF";
          // 範本儲存格會共用樣式物件，直接改 cell.fill 會互相污染；須整份複製再替換
          cell.style = Object.assign({}, cell.style, { fill: isLeave ? FILL_LEAVE : FILL_NONE });
        });
      } else if (!p.active) {
        model.dateCols.forEach((c) => {
          const cell = ws.getCell(p.row, c);
          cell.value = null;
          cell.style = Object.assign({}, cell.style, { fill: FILL_NONE });
        });
      }
    });
  }

  function summarizeDaily(model, assignment) {
    const carryLen = model.carryLen || 0; // 摘要只列實際排班月（略過銜接段）
    const people = model.people.filter((p) => p.active);
    const rows = [];
    for (let d = carryLen; d < model.nDays; d++) {
      const c = { D: 0, E: 0, N: 0 }; let off = 0;
      people.forEach((p) => { const v = assignment[p.label][d]; if (SHIFTS.includes(v)) c[v]++; else off++; });
      const label = (model.monthsByDay && model.dates) ? `${model.monthsByDay[d]}/${model.dates[d]}` : String(d + 1);
      rows.push({ day: d + 1, dayIdx: d, label, D: c.D, E: c.E, N: c.N, off });
    }
    return rows;
  }

  const api = { SHIFTS, GROUPS, TAIL_LEN, DEFAULT_RULES, BUILTIN_HOLIDAYS, classifyCell, parseWorkbook, generate, validate, labelOffs, writeInto, writeCalendar, buildCalendar, buildTail, buildCarry, holidayName, summarizeDaily };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Scheduler = api;
})(typeof window !== "undefined" ? window : this);
