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
  const MAX_COLS = 16384; // Excel 工作表欄數上限（XFD）

  const DEFAULT_RULES = {
    minManpower: { D: 16, E: 16, N: 14 },
    maxConsecutiveWork: 5,
    maxConsecutiveN: 4,
    minLeaderPerShift: 2,
    tierCaps: { D: [14, 11, 8, 4], E: [14, 11, 8, 4], N: [12, 9, 6, 3] },
    // 更換班別需連續 11 小時休息（D 08-16 / E 16-24 / N 00-08 值到隔日早上）
    // → 禁止的「前一天班別 -> 今日班別」組合：N 下班後接 D(0h)/E(8h)、E 下班後接 D(8h)
    forbidTransition: { N: ["D", "E"], E: ["D"] },
    forbidNOffD: true, // 不能 N off D：大夜後僅休 1 天不得接白班（需求明列；真實班表亦無此模式）
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
    // 與範本 COUNTIF("D*"/"E*"/"N*") 的前綴語意一致：
    // 「D/特」「N/息」算上班；「病E」「生理d」「新N」等前綴代碼不算上班
    for (const sh of SHIFTS) if (up.startsWith(sh)) return sh;
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

    // 日期列 + 星期列；日期列底色只在「換月份」處變色（同範本慣例：月份段交替 灰底 / 藍紫底）
    const FILL_DATE_A = { type: "pattern", pattern: "solid", fgColor: { theme: 2 } };
    const FILL_DATE_B = { type: "pattern", pattern: "solid", fgColor: { argb: "FFCCCCFF" } };
    let monthSeg = -1;
    const dateFillByDay = cal.map((c, i) => {
      if (i === 0 || c.d === 1) monthSeg++;
      return monthSeg % 2 === 0 ? FILL_DATE_A : FILL_DATE_B;
    });
    cal.forEach((c, i) => {
      const col = model.dateCols[i];
      const dCell = ws.getCell(model.dateRow, col);
      dCell.value = c.d;
      dCell.style = Object.assign({}, dCell.style, { fill: dateFillByDay[i] });
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

    // 底部驗證區也有自己的「日期」列（供 COUNTIF 對照），日期與「換月變色」底色一併同步
    const maxRow = ws.rowCount || 300;
    for (let r = model.dateRow + 1; r <= maxRow; r++) {
      if (cellText(ws.getCell(r, 3).value) === "日期") {
        cal.forEach((c, i) => {
          const cell = ws.getCell(r, model.dateCols[i]);
          cell.value = c.d;
          cell.style = Object.assign({}, cell.style, { fill: dateFillByDay[i] });
        });
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
    const consecWork = {}, consecN = {}, lastShift = {}, workCount = {}, kindsUsed = {}, weekendWork = {}, weekendOffNew = {};
    // 六日休假平均：需 model.weekdays（app 於產生前以真實日曆設定；無資料時此偏好自動停用）
    const weekendByDay = (model.weekdays || []).map((w) => w === "六" || w === "日");
    // 六日休假保底：每人在新排班段內至少 1 個六日休假 —— wkAfter[d] = d 之後還剩幾個週末日
    const wkAfter = new Array(N).fill(0);
    for (let i = N - 2; i >= 0; i--) wkAfter[i] = wkAfter[i + 1] + (weekendByDay[i + 1] ? 1 : 0);
    labels.forEach((l) => {
      assigned[l] = new Array(N).fill(null);
      kindsUsed[l] = new Set();
      const t = byLabel[l].tail;
      let cw = 0; for (let i = t.length - 1; i >= 0; i--) { if (t[i]) cw++; else break; }
      let cn = 0; for (let i = t.length - 1; i >= 0; i--) { if (t[i] === "N") cn++; else break; }
      consecWork[l] = cw; consecN[l] = cn;
      lastShift[l] = t.length ? t[t.length - 1] : null;
      workCount[l] = 0;
      weekendWork[l] = 0;
      weekendOffNew[l] = 0;
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
            if (weekendByDay[d]) weekendWork[l] += 1; // 銜接段的六日上班也計入，公平性以整個視窗計
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
        forcedOff[l] = consecWork[l] >= rules.maxConsecutiveWork || fixedOff[l].has(d) || mustRest(l) ||
          // 六日保底：已到本段最後一個週末日、此人週末休假仍為 0 → 強制休，避免整月六日全上班
          (weekendByDay[d] && wkAfter[d] === 0 && weekendOffNew[l] === 0);
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
        if (s === "D" && rules.forbidNOffD) {
          // 不能 N off D：昨天休、前天大夜 → 今天不得排白班（跨上月尾巴也要查）
          const t = byLabel[l].tail;
          const p1 = d >= 1 ? assigned[l][d - 1] : t[t.length - 1];
          const p2 = d >= 2 ? assigned[l][d - 2] : (d === 1 ? t[t.length - 1] : t[t.length - 2]);
          if (!p1 && p2 === "N") return false;
        }
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
          // 六日休假平均：週末日以「週末上班次數少者先上」為主排序，避免同一批人整月佔滿週末
          // （週間仍以工作量為主，總工作量會在週間自然拉回平衡）
          if (weekendByDay[d] && weekendWork[a] !== weekendWork[b]) return weekendWork[a] - weekendWork[b];
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
          if (weekendByDay[d]) weekendWork[l] += 1;
        } else {
          consecWork[l] = 0; consecN[l] = 0; lastShift[l] = null;
          if (weekendByDay[d]) weekendOffNew[l] += 1;
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

    // 跨月邊界（月底→月初）的換班休息，即使落在銜接段也要抓（銜接段內同月才視為既成不報）。
    // 依真實日曆月份判斷；無月份資料時（純邏輯測試）維持原「銜接段整段不報」的行為。
    const months = model.monthsByDay, tailMonths = model.tailMonths;
    const crossInSeq = (d) => !!months && months[d] !== months[d - 1];
    const crossAtFirst = carryLen > 0 && !!months && !!tailMonths && tailMonths.length > 0 &&
      tailMonths[tailMonths.length - 1] !== months[0];

    people.forEach((p) => {
      const lab = p.label, seq = assignment[lab], h = tail[lab], base = h.length;
      const prevLast = h.length ? h[h.length - 1] : null;
      // 首日 vs 上月末：carryLen===0 一律檢查；carryLen>0 僅在跨月邊界（首日為月初1號）也檢查
      if (prevLast && seq[0] && (rules.forbidTransition[prevLast] || []).includes(seq[0]) && (carryLen === 0 || crossAtFirst))
        addCell(lab, 0, `[換班休息] ${lab} 上月末(${prevLast})後首日排${seq[0]}（休息不足11小時）`);
      // 相鄰兩日換班：銜接段內同月屬既成事實不報，跨月邊界（月底→月初）仍要抓
      for (let d = 1; d < N; d++) {
        if (d < carryLen && !crossInSeq(d)) continue;
        if (seq[d] && seq[d - 1] && (rules.forbidTransition[seq[d - 1]] || []).includes(seq[d]))
          addCell(lab, d, `[換班休息] ${lab} ${dayLabel(d)}排${seq[d]} 前一天為${seq[d - 1]}（休息不足11小時）`);
      }

      // 不能 N off D：大夜後僅休 1 天接白班（含跨上月尾巴/銜接段邊界）
      if (rules.forbidNOffD) {
        const ext = h.concat(seq); // tail + 本期
        for (let i = base + carryLen; i < ext.length; i++)
          if (ext[i] === "D" && i >= 2 && !ext[i - 1] && ext[i - 2] === "N")
            addCell(lab, i - base, `[N-off-D] ${lab} ${dayLabel(i - base)}排D 前為大夜+僅休1天（禁止 N-off-D）`);
      }

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

  // 逐格顯示：上班班別原樣保留；休假日一律留空白（不自動填「例/息」）。
  // 例假與休息日之區分由護理長依實際狀況自行填寫（銜接段照抄上傳檔原始內容，不經此函式）。
  function labelOffs(seq) {
    return seq.map((code) => (SHIFTS.includes(code) ? code : ""));
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

  /* 欄位字母 <-> 數字；平移公式中「欄 >= fromCol」的儲存格引用 +by 欄。 */
  function colToNum(s) { let n = 0; for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function numToCol(n) { let s = ""; while (n > 0) { s = String.fromCharCode(65 + (n - 1) % 26) + s; n = Math.floor((n - 1) / 26); } return s; }
  function shiftFormulaCols(f, by, fromCol) {
    return f.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (m, a, cl, b, row) => {
      const n = colToNum(cl); return n >= fromCol ? a + numToCol(n + by) + b + row : m;
    });
  }
  function refOf(ref) { const mt = ref.match(/([A-Z]+)(\d+)/); return { col: colToNum(mt[1]), row: +mt[2] }; }
  function cloneValidation(v) {
    return Object.assign({}, v, { formulae: v && v.formulae ? v.formulae.slice() : undefined });
  }
  function shiftAddressRef(ref, by, fromCol) {
    const mt = String(ref).match(/^(\$?)([A-Z]{1,3})(\$?)(\d+)$/);
    if (!mt) return ref;
    const n = colToNum(mt[2]);
    return n >= fromCol ? mt[1] + numToCol(n + by) + mt[3] + mt[4] : ref;
  }
  function shiftSqref(ref, by, fromCol) {
    return String(ref || "").split(/\s+/).filter(Boolean).map((part) =>
      part.split(":").map((addr) => shiftAddressRef(addr, by, fromCol)).join(":")
    ).join(" ");
  }
  function cloneConditionalFormattings(list) {
    return (list || []).map((cf) => Object.assign({}, cf, {
      rules: (cf.rules || []).map((rule) => Object.assign({}, rule, {
        formulae: rule.formulae ? rule.formulae.slice() : rule.formulae,
      })),
    }));
  }
  /* 依「插欄前的快照」重建整份資料驗證。
   * ExcelJS 的 spliceColumns 不平移 dataValidations，驗證會留在原欄，而檢核欄早已右移
   * （AS/AV → BX/CA），等於套用到錯誤的欄；新插入的日期欄則完全沒有驗證。
   * 日期區內的驗證（如底部 N 班人力上下限）本就該套用到每個日期欄，故整列擴展到含新插入欄的
   * 完整日期區；右側檢核欄整體 +insertCount 平移，左側固定欄不動。 */
  function restoreExpandedDataValidations(ws, snapshot, firstCol, originalNDays, insertCount) {
    const originalLastCol = firstCol + originalNDays - 1;
    const out = {};
    const dateRowVals = {}; // 日期區：列 -> 驗證（整列擴展）
    const groups = {};      // 日期區外：同欄同驗證 -> 合併連續列
    Object.keys(snapshot || {}).forEach((addr) => {
      const p = refOf(addr);
      const v = cloneValidation(snapshot[addr]);
      if (p.col >= firstCol && p.col <= originalLastCol) {
        dateRowVals[p.row] = dateRowVals[p.row] || v;
        return;
      }
      const col = p.col > originalLastCol ? p.col + insertCount : p.col;
      const sig = col + "|" + JSON.stringify(v);
      if (!groups[sig]) groups[sig] = { col, validation: v, rows: [] };
      groups[sig].rows.push(p.row);
    });
    const dl = numToCol(firstCol), dr = numToCol(originalLastCol + insertCount);
    Object.keys(dateRowVals).forEach((row) => { out[dl + row + ":" + dr + row] = dateRowVals[row]; });
    Object.keys(groups).forEach((sig) => {
      const g = groups[sig], c = numToCol(g.col);
      const rows = Array.from(new Set(g.rows)).sort((a, b) => a - b);
      let start = rows[0], prev = rows[0];
      for (let i = 1; i <= rows.length; i++) {
        if (rows[i] === prev + 1) { prev = rows[i]; continue; }
        out[c + start + (prev === start ? "" : ":" + c + prev)] = cloneValidation(g.validation);
        start = rows[i]; prev = rows[i];
      }
    });
    ws.dataValidations.model = out;
  }

  /* 在日期區前插入 insertCount 欄，讓「上傳單月表」也能容納「上月銜接段 + 新月」的跨月格式。
   * 關鍵：spliceColumns 不會調整公式引用，故插欄後把整份表的公式欄引用整體 +insertCount
   * （欄 >= firstCol），如此 E:AH→AA:BD、分段 E:J→AA:AF、周末欄… 都精準平移到新月對應位置、
   * 語義不變。另重建合併(含主格標題值)、資料驗證、條件格式、複製日期區樣式，
   * 並更新 model.dateCols / nDays。*/
  function expandTemplateColumns(ws, model, insertCount) {
    if (!insertCount || insertCount <= 0) return;
    const firstCol = model.firstCol;
    const originalNDays = model.nDays;
    const originalDataValidations = Object.assign({}, ws.dataValidations && ws.dataValidations.model);
    const originalConditionalFormattings = cloneConditionalFormattings(ws.conditionalFormattings);
    const merges = (ws.model.merges || []).slice();
    const mergeVals = merges.map((mg) => { const s = refOf(mg.split(":")[0]); const c = ws.getCell(s.row, s.col); return { row: s.row, col: s.col, val: c.formula ? null : c.value }; });
    // 解除共享公式（否則 spliceColumns 會破壞 master/clone 關係而寫檔失敗）
    ws.eachRow({ includeEmpty: true }, (r) => r.eachCell({ includeEmpty: true }, (c) => { if (c.formula) c.value = { formula: c.formula }; }));
    ws.spliceColumns(firstCol, 0, ...Array.from({ length: insertCount }, () => []));
    /* spliceColumns 把整段欄定義往右搬，連範本那條「涵蓋到工作表最後一欄」的 <col min=137 max=16384>
     * 也 +insertCount，寫出 <col max="16415"> —— 超過 Excel 欄上限 16384(XFD)，Excel 即判定檔案
     * 損毀而拒絕開啟（單月表輸出打不開的主因；跨月表不走此函式故無恙）。ExcelJS 自身無上限保護，
     * 故手動截去溢位的欄定義。另 splice 會把新插入欄的定義清成 null，須補回日期欄寬，
     * 否則銜接段各欄會縮成預設寬度。 */
    const dateColDefn = ws.getColumn(firstCol + insertCount).defn; // 原日期區首欄，已右移到此
    for (let c = firstCol; c < firstCol + insertCount; c++) ws.getColumn(c).defn = dateColDefn;
    if (ws._columns.length > MAX_COLS) ws._columns.length = MAX_COLS;
    // spliceColumns 不調整引用，逐格把公式欄引用 +insertCount
    ws.eachRow({ includeEmpty: true }, (r) => r.eachCell({ includeEmpty: true }, (c) => {
      if (c.formula) { const nf = shiftFormulaCols(c.formula, insertCount, firstCol); if (nf !== c.formula) c.value = { formula: nf }; }
    }));
    // 重建合併（欄 >= firstCol 的邊界 +insertCount）並恢復主格值（如標題）
    merges.forEach((mg, i) => {
      const [s, e] = mg.split(":"); const sc = refOf(s), ec = refOf(e);
      const nsc = sc.col >= firstCol ? sc.col + insertCount : sc.col, nec = ec.col >= firstCol ? ec.col + insertCount : ec.col;
      const range = numToCol(nsc) + sc.row + ":" + numToCol(nec) + ec.row;
      try { ws.unMergeCells(range); } catch (e) { /* 無殘留合併 */ }
      try { ws.mergeCells(range); } catch (e) { /* 已合併 */ }
      const mv = mergeVals[i], nmc = mv.col >= firstCol ? mv.col + insertCount : mv.col;
      if (mv.val != null) ws.getCell(mv.row, nmc).value = mv.val;
    });
    // spliceColumns 完全不平移資料驗證/條件格式，兩者會留在原欄位置、套用到錯誤的欄。
    restoreExpandedDataValidations(ws, originalDataValidations, firstCol, originalNDays, insertCount);
    ws.conditionalFormattings = originalConditionalFormattings.map((cf) =>
      Object.assign({}, cf, { ref: shiftSqref(cf.ref, insertCount, firstCol) })
    );
    // 複製日期區樣式＋「逐日統計公式」到新欄，補齊銜接段各日的底部逐日階層統計
    // （COUNTIF(該欄6:76,"D*")…），否則那幾天統計會空白。
    // 參考欄取「日期區第二個資料欄」而非首欄：首欄常是月首、統計行範圍可能特殊（本院月首那欄
    // 多含 2 行），而銜接段各日都不是月首，應套用一般日的行範圍；月首欄的特殊處理由主平移保留。
    const ref = firstCol + insertCount + 1;
    for (let c = firstCol; c < firstCol + insertCount; c++) {
      const backShift = c - ref;                        // 把參考欄的欄引用平移到本欄
      for (let r = 1; r <= ws.rowCount; r++) {
        const src = ws.getCell(r, ref), dst = ws.getCell(r, c);
        dst.style = Object.assign({}, src.style);
        if (src.formula) dst.value = { formula: shiftFormulaCols(src.formula, backShift, firstCol) };
      }
    }
    model.nDays += insertCount;
    model.dateCols = Array.from({ length: model.nDays }, (_, i) => firstCol + i);
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

  const api = { SHIFTS, GROUPS, TAIL_LEN, DEFAULT_RULES, BUILTIN_HOLIDAYS, classifyCell, parseWorkbook, generate, validate, labelOffs, writeInto, writeCalendar, buildCalendar, buildTail, buildCarry, holidayName, summarizeDaily, expandTemplateColumns };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Scheduler = api;
})(typeof window !== "undefined" ? window : this);
