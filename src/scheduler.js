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
    forbidNextDayDAfter: ["N", "E"],
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
    for (let r = dateRow + 1; r <= maxRow; r++) {
      const b = ws.getCell(r, 2).value;
      const c = cellText(ws.getCell(r, 3).value);
      if (typeof b === "number" && PERSON_RE.test(c)) {
        const label = c.toUpperCase();
        const group = label[0];
        const tail = tailCols.map((col) => classifyCell(ws.getCell(r, col).value)); // D/E/N/null
        const rawTail = tailCols.map((col) => cellText(ws.getCell(r, col).value)); // 原始顯示文字
        const rowRaw = dateCols.map((col) => cellText(ws.getCell(r, col).value)); // 整月原始班別
        let active = true;
        for (let i = 0; i < nDays; i++) {
          if (cellText(ws.getCell(r, firstCol + i).value).includes("離職")) { active = false; break; }
        }
        people.push({ row: r, label, group, active, tail, rawTail, rowRaw });
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

    // 日期列 + 星期列
    cal.forEach((c, i) => {
      const col = model.dateCols[i];
      ws.getCell(model.dateRow, col).value = c.d;
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
   * fixed: { label: { dayIndex: 'OFF' } } —— 預先請假(強制休假)。
   * 回傳 { label: [D/E/N/null,...] }
   */
  function generate(model, rules, fixed) {
    rules = rules || DEFAULT_RULES;
    fixed = fixed || {};
    const people = model.people.filter((p) => p.active);
    const N = model.nDays;
    const labels = people.map((p) => p.label);
    const byLabel = {};
    people.forEach((p) => (byLabel[p.label] = p));

    const fixedOff = {}; // label -> Set(day)
    labels.forEach((l) => (fixedOff[l] = new Set()));
    Object.keys(fixed).forEach((l) => {
      if (!fixedOff[l]) return;
      Object.keys(fixed[l]).forEach((d) => {
        const di = parseInt(d, 10);
        if (di >= 0 && di < N && fixed[l][d] === "OFF") fixedOff[l].add(di);
      });
    });

    const assigned = {};
    const consecWork = {}, consecN = {}, lastShift = {}, workCount = {};
    labels.forEach((l) => {
      assigned[l] = new Array(N).fill(null);
      const t = byLabel[l].tail;
      let cw = 0; for (let i = t.length - 1; i >= 0; i--) { if (t[i]) cw++; else break; }
      let cn = 0; for (let i = t.length - 1; i >= 0; i--) { if (t[i] === "N") cn++; else break; }
      consecWork[l] = cw; consecN[l] = cn;
      lastShift[l] = t.length ? t[t.length - 1] : null;
      workCount[l] = 0;
    });

    for (let d = 0; d < N; d++) {
      const todayShift = {};
      const cnt = { D: 0, E: 0, N: 0 };
      const leaderCnt = { D: 0, E: 0, N: 0 };
      const tierCnt = { D: [0, 0, 0, 0], E: [0, 0, 0, 0], N: [0, 0, 0, 0] };
      const forcedOff = {};
      labels.forEach((l) => {
        forcedOff[l] = consecWork[l] >= rules.maxConsecutiveWork || fixedOff[l].has(d);
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
        if (s === "D" && rules.forbidNextDayDAfter.includes(prev)) return false;
        if (s === "N" && consecN[l] >= rules.maxConsecutiveN) return false;
        if (!capOK(byLabel[l].group, s)) return false;
        return true;
      }
      function doAssign(l, s) {
        todayShift[l] = s; cnt[s]++;
        if (byLabel[l].group === "A") leaderCnt[s]++;
        for (const i of tierIndexesFor(byLabel[l].group)) tierCnt[s][i]++;
      }

      for (const s of ["N", "D", "E"]) {
        while (leaderCnt[s] < rules.minLeaderPerShift) {
          const cands = labels.filter((l) => byLabel[l].group === "A" && eligible(l, s));
          if (!cands.length) break;
          cands.sort((a, b) => workCount[a] - workCount[b]);
          doAssign(cands[0], s);
        }
      }
      const order = ["N", "D", "E"].sort((a, b) => rules.minManpower[b] - rules.minManpower[a]);
      for (const s of order) {
        while (cnt[s] < rules.minManpower[s]) {
          let cands = labels.filter((l) => eligible(l, s));
          if (!cands.length) break;
          cands.sort((a, b) => {
            if (workCount[a] !== workCount[b]) return workCount[a] - workCount[b];
            const pa = lastShift[a] === s ? 0 : 1;
            const pb = lastShift[b] === s ? 0 : 1;
            return pa - pb;
          });
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
  function validate(model, assignment, rules) {
    rules = rules || DEFAULT_RULES;
    const people = model.people.filter((p) => p.active);
    const N = model.nDays;
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

    for (let d = 0; d < N; d++) {
      const c = counts(d, allLabels);
      for (const s of SHIFTS)
        if (c[s] < rules.minManpower[s]) addDay(d, `[最低人力] 第${d + 1}天 ${s}班 僅 ${c[s]} 人 (<${rules.minManpower[s]})`);
      const ca = counts(d, aLabels);
      for (const s of SHIFTS)
        if (ca[s] < rules.minLeaderPerShift) addDay(d, `[小組長] 第${d + 1}天 ${s}班 A階僅 ${ca[s]} 人`);
      for (const s of SHIFTS) {
        rules.tierCaps[s].forEach((cap, i) => {
          const labs = allLabels.filter((l) => groupsFrom(TIER_START[i]).includes(grp[l]));
          const cc = counts(d, labs)[s];
          if (cc > cap) addDay(d, `[階層上限] 第${d + 1}天 ${s}班 ${TIER_START[i]}階以下 ${cc} 人 (>${cap})`);
        });
      }
    }

    people.forEach((p) => {
      const lab = p.label, seq = assignment[lab], h = tail[lab], base = h.length;
      const prevLast = h.length ? h[h.length - 1] : null;
      if (rules.forbidNextDayDAfter.includes(prevLast) && seq[0] === "D")
        addCell(lab, 0, `[換班休息] ${lab} 上月末(${prevLast})後首日排D`);
      for (let d = 1; d < N; d++)
        if (seq[d] === "D" && rules.forbidNextDayDAfter.includes(seq[d - 1]))
          addCell(lab, d, `[換班休息] ${lab} 第${d + 1}天D 前一天為${seq[d - 1]}`);

      const wseq = h.map((c) => (SHIFTS.includes(c) ? 1 : 0)).concat(seq.map((c) => (SHIFTS.includes(c) ? 1 : 0)));
      let run = 0;
      for (let i = 0; i < wseq.length; i++) {
        run = wseq[i] ? run + 1 : 0;
        if (run > rules.maxConsecutiveWork && i >= base) { addCell(lab, i - base, `[連續上班] ${lab} 連續上班超過 ${rules.maxConsecutiveWork} 天`); break; }
      }
      const nseq = h.map((c) => (c === "N" ? 1 : 0)).concat(seq.map((c) => (c === "N" ? 1 : 0)));
      run = 0;
      for (let i = 0; i < nseq.length; i++) {
        run = nseq[i] ? run + 1 : 0;
        if (run > rules.maxConsecutiveN && i >= base) { addCell(lab, i - base, `[連續大夜] ${lab} 連續大夜超過 ${rules.maxConsecutiveN} 天`); break; }
      }
      const off = seq.map((c) => (SHIFTS.includes(c) ? 0 : 1));
      for (const [w, mo] of rules.restWindows) {
        if (w <= N) {
          for (let st = 0; st <= N - w; st++) {
            let sum = 0; for (let i = st; i < st + w; i++) sum += off[i];
            if (sum < mo) { addRow(lab, `[休假] ${lab} 第${st + 1}~${st + w}天 休假 <${mo} 日`); break; }
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

  function writeInto(ws, model, assignment, rules) {
    rules = rules || DEFAULT_RULES;
    model.people.forEach((p) => {
      if (assignment[p.label]) {
        const disp = labelOffs(assignment[p.label], rules);
        model.dateCols.forEach((c, i) => { ws.getCell(p.row, c).value = disp[i]; });
      } else if (!p.active) {
        model.dateCols.forEach((c) => { ws.getCell(p.row, c).value = null; });
      }
    });
  }

  function summarizeDaily(model, assignment) {
    const people = model.people.filter((p) => p.active);
    const rows = [];
    for (let d = 0; d < model.nDays; d++) {
      const c = { D: 0, E: 0, N: 0 }; let off = 0;
      people.forEach((p) => { const v = assignment[p.label][d]; if (SHIFTS.includes(v)) c[v]++; else off++; });
      rows.push({ day: d + 1, D: c.D, E: c.E, N: c.N, off });
    }
    return rows;
  }

  const api = { SHIFTS, GROUPS, TAIL_LEN, DEFAULT_RULES, BUILTIN_HOLIDAYS, classifyCell, parseWorkbook, generate, validate, labelOffs, writeInto, writeCalendar, buildCalendar, buildTail, holidayName, summarizeDaily };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Scheduler = api;
})(typeof window !== "undefined" ? window : this);
