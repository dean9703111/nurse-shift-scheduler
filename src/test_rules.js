/* 需求規則驗證程式（對照：案例_醫事人員排班_V2.docx）
 *
 * A. 規格常數對照 —— DEFAULT_RULES 逐項對照需求文件數值與範本 Excel 的 COUNTIF 公式
 * B. 驗證器正反例 —— 每條規則建立「合規應通過 / 違規必須被抓到」的單元測試
 * C. 端對端獨立稽核 —— 以獨立實作（不共用 validate 程式碼）稽核 generate() 產出，
 *    涵蓋三情境（無預假 / 同日多人預假壓力 / 隨機預假），並與 validate() 交叉比對
 * D. 歷史資料佐證 —— 真實班表驗證規則解讀（如 N-off-D 為三天模式）
 *
 * 執行：node src/test_rules.js （需根目錄 sample.xlsx）
 */
const path = require("path");
const ExcelJS = require("exceljs");
const JSZip = require("jszip");
const S = require("./scheduler.js");

const SRC = path.resolve(__dirname, "..", "sample.xlsx");
const SRC_MONTHLY = path.resolve(__dirname, "..", "sample-monthly.xlsx");
const SHIFTS = ["D", "E", "N"];
const HOURS_PER_SHIFT = 8; // D 08-16 / E 16-24 / N 00-08，單班 8 小時

const checks = [];
let section = "";
function check(name, pass, detail) {
  checks.push([section, name, !!pass, detail || ""]);
}

/* ============ 共用 ============ */
function cellStr(ws, r, c) {
  const v = ws.getCell(r, c).value;
  if (v == null) return "";
  if (typeof v === "object") {
    if (v.result !== undefined && v.result !== null) return String(v.result);
    if (v.richText) return v.richText.map((t) => t.text).join("");
    return "";
  }
  return String(v).trim();
}
function formulaOf(ws, r, c) {
  const cell = ws.getCell(r, c);
  return cell.formula || (cell.value && cell.value.formula) || "";
}
function mulberry32(seed) {
  return function () {
    seed |= 0; seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/* ============ A. 規格常數對照 ============ */
function partA(ws, model) {
  section = "A.規格常數";
  const R = S.DEFAULT_RULES;
  // 需求：每班最低人力 D:16／E:16／N:14
  check("最低人力 D16/E16/N14（需求文件）",
    R.minManpower.D === 16 && R.minManpower.E === 16 && R.minManpower.N === 14,
    JSON.stringify(R.minManpower));
  // 需求：連續上班 D/E 最多 5 天、N 最多 4 天
  check("連續上班上限 5 天、連續大夜上限 4 天（需求文件）",
    R.maxConsecutiveWork === 5 && R.maxConsecutiveN === 4,
    `work<=${R.maxConsecutiveWork}, N<=${R.maxConsecutiveN}`);
  // 需求：每日 D/E/N 小組長 >= 2（範本「>=2」列）
  check("小組長下限 2（需求文件）", R.minLeaderPerShift === 2, `>=${R.minLeaderPerShift}`);
  // 需求：包班（同班別>15）、未包班每月至多 2 種班別
  check("包班門檻 >15、未包班至多 2 種班別（需求文件）",
    R.lockThreshold === 15 && R.maxShiftKinds === 2,
    `lock>${R.lockThreshold}, kinds<=${R.maxShiftKinds}`);
  // 需求：換班需連續 11 小時休息 → 禁止 N→D(0h)、N→E(8h)、E→D(8h)；其餘組合皆 >=16h
  const ft = R.forbidTransition;
  check("11小時休息 → 禁止 N→D/N→E/E→D（需求文件）",
    JSON.stringify(ft.N) === JSON.stringify(["D", "E"]) && JSON.stringify(ft.E) === JSON.stringify(["D"]) && !ft.D,
    JSON.stringify(ft));
  check("禁止 N-off-D（需求文件：不能 N off D）", R.forbidNOffD === true, String(R.forbidNOffD));
  // 需求：2週內至少2日例假、4週內至少8日例假及休息日
  check("休假視窗 [14天,2日] [28天,8日]（需求文件）",
    JSON.stringify(R.restWindows) === JSON.stringify([[14, 2], [28, 8]]),
    JSON.stringify(R.restWindows));

  // 範本 Excel 驗證區：階層上限標籤（如 B-E 列的 D14/E14/N12）與 DEFAULT_RULES.tierCaps 對照
  const TIER_ORDER = ["B-E", "C-E", "D-E", "E"];
  const found = {}; // tier -> {D:cap, E:cap, N:cap, rows:{shift:row}}
  const maxRow = ws.rowCount || 300;
  for (let r = 1; r <= maxRow; r++) {
    const t1 = cellStr(ws, r, 1);
    const lab = cellStr(ws, r, 3);
    const m = lab.match(/^([DEN])(\d+)$/);
    if (TIER_ORDER.includes(t1) && m) {
      found[t1] = found[t1] || { rows: {} };
      found[t1][m[1]] = parseInt(m[2], 10);
      found[t1].rows[m[1]] = r;
    }
  }
  TIER_ORDER.forEach((tier, i) => {
    const f = found[tier];
    const ok = f && SHIFTS.every((s) => R.tierCaps[s][i] === f[s]);
    check(`階層上限 ${tier} 對照範本（${f ? SHIFTS.map((s) => s + f[s]).join("/") : "未找到"}）`,
      ok, `rules=${SHIFTS.map((s) => s + R.tierCaps[s][i]).join("/")}`);
  });

  // 範本 COUNTIF 範圍語意：B-E 列公式應從第一位 B 階人員列起算到最後一位人員列（累進計數）
  const firstRowOf = {}; let lastPersonRow = 0;
  model.people.forEach((p) => {
    if (firstRowOf[p.group] === undefined) firstRowOf[p.group] = p.row;
    lastPersonRow = Math.max(lastPersonRow, p.row);
  });
  const tierStart = { "B-E": "B", "C-E": "C", "D-E": "D", "E": "E" };
  TIER_ORDER.forEach((tier) => {
    const f = found[tier];
    if (!f) return check(`範本公式範圍 ${tier}`, false, "未找到標籤列");
    const fx = formulaOf(ws, f.rows.D, model.dateCols[0]);
    const m = fx.match(/COUNTIF\([A-Z]+(\d+):[A-Z]+(\d+),"D\*"\)/i);
    const ok = m && parseInt(m[1], 10) === firstRowOf[tierStart[tier]] && parseInt(m[2], 10) === lastPersonRow;
    check(`範本公式範圍 ${tier}＝第一位${tierStart[tier]}階(列${firstRowOf[tierStart[tier]]})~最後人員(列${lastPersonRow})`,
      ok, fx || "無公式");
  });

  // 範本小組長列：「>=2」且 公式 = 全體 - B階以下（即 A 階人數）
  let leaderRow = 0;
  for (let r = 1; r <= maxRow; r++) if (cellStr(ws, r, 3) === "D小組長") { leaderRow = r; break; }
  const leaderTag = leaderRow ? cellStr(ws, leaderRow, 1) : "";
  check("範本小組長檢核列存在且標示 >=2", leaderRow > 0 && leaderTag.includes(">=2"),
    leaderRow ? `列${leaderRow} 標示「${leaderTag}」` : "未找到");
}

/* ============ B. 驗證器正反例 ============ */
// 寬鬆基準規則：把不相關的限制放到最鬆，逐項覆寫測試目標規則
function baseRules(over) {
  return Object.assign({
    minManpower: { D: 0, E: 0, N: 0 },
    maxConsecutiveWork: 99, maxConsecutiveN: 99, minLeaderPerShift: 0,
    tierCaps: { D: [99, 99, 99, 99], E: [99, 99, 99, 99], N: [99, 99, 99, 99] },
    forbidTransition: { N: ["D", "E"], E: ["D"] }, forbidNOffD: true,
    maxShiftKinds: 2, lockThreshold: 15,
    restWindows: [], statutoryPeriod: 7, offStatutory: "例", offRest: "息",
  }, over || {});
}
function makeModel(peopleSpec, nDays, carryLen) {
  return {
    nDays, carryLen: carryLen || 0,
    people: peopleSpec.map((p) => ({
      label: p.label, group: p.group || p.label[0], active: true,
      tail: p.tail || [null, null, null, null, null, null, null],
    })),
  };
}
function kindsOf(flags) {
  const set = new Set();
  flags.errors.forEach((e) => { const m = e.match(/^\[([^\]]+)\]/); if (m) set.add(m[1]); });
  return set;
}
function expectFlag(name, model, asg, rules, locks, kind, where) {
  const flags = S.validate(model, asg, rules, locks || {});
  const ks = kindsOf(flags);
  let posOK = true, posDesc = "";
  if (where && where.cell) {
    posOK = flags.cellFlags.some((f) => f.label === where.cell[0] && f.day === where.cell[1] && f.msg.startsWith(`[${kind}]`));
    posDesc = `@${where.cell[0]}/day${where.cell[1]}`;
  }
  check(`【${kind}】違規必須被抓到：${name}`, ks.has(kind) && posOK,
    `flags=${Array.from(ks).join(",") || "無"}${posDesc}`);
}
function expectClean(name, model, asg, rules, locks, kind) {
  const flags = S.validate(model, asg, rules, locks || {});
  const ks = kindsOf(flags);
  check(`【${kind}】合規不得誤報：${name}`, !ks.has(kind), `flags=${Array.from(ks).join(",") || "無"}`);
}
function partB() {
  section = "B.驗證器";
  // --- 最低人力 ---
  {
    const m = makeModel([{ label: "A1" }], 1);
    const r = baseRules({ minManpower: { D: 1, E: 0, N: 0 } });
    expectFlag("D 需 1 人實排 0 人", m, { A1: [null] }, r, null, "最低人力");
    expectClean("D 需 1 人實排 1 人", m, { A1: ["D"] }, r, null, "最低人力");
  }
  // --- 小組長 ---
  {
    const m = makeModel([{ label: "A1" }, { label: "A2" }, { label: "A3" }, { label: "B1" }], 1);
    const r = baseRules({ minLeaderPerShift: 1 });
    expectClean("每班各 1 位 A 階", m, { A1: ["D"], A2: ["E"], A3: ["N"], B1: ["D"] }, r, null, "小組長");
    expectFlag("E 班 A 階 0 人（B 階不算小組長）", m, { A1: ["D"], A2: [null], A3: ["N"], B1: ["E"] }, r, null, "小組長");
  }
  // --- 階層上限（累進語意：B-E 上限計 B~E 全部）---
  {
    const m = makeModel([{ label: "C1" }, { label: "E1" }], 1);
    const r = baseRules({ tierCaps: { D: [1, 99, 99, 99], E: [99, 99, 99, 99], N: [99, 99, 99, 99] } });
    expectFlag("B階以下上限1，C+E 兩人同日排D", m, { C1: ["D"], E1: ["D"] }, r, null, "階層上限");
    expectClean("B階以下上限1，僅 E1 排D", m, { C1: [null], E1: ["D"] }, r, null, "階層上限");
  }
  {
    const m = makeModel([{ label: "D1" }, { label: "E1" }], 1);
    const r = baseRules({ tierCaps: { D: [99, 99, 99, 99], E: [99, 99, 99, 99], N: [99, 99, 99, 1] } });
    expectFlag("E階上限1，E1+E2 同日排N", makeModel([{ label: "E1" }, { label: "E2" }], 1),
      { E1: ["N"], E2: ["N"] }, r, null, "階層上限");
    expectClean("E階上限1，D1(D階)+E1 排N", m, { D1: ["N"], E1: ["N"] }, r, null, "階層上限");
  }
  // --- 換班休息（11 小時）---
  {
    const m = makeModel([{ label: "B1" }], 2);
    const r = baseRules();
    expectFlag("N→D 相鄰", m, { B1: ["N", "D"] }, r, null, "換班休息", { cell: ["B1", 1] });
    expectFlag("N→E 相鄰", m, { B1: ["N", "E"] }, r, null, "換班休息", { cell: ["B1", 1] });
    expectFlag("E→D 相鄰", m, { B1: ["E", "D"] }, r, null, "換班休息", { cell: ["B1", 1] });
    expectClean("D→E 相鄰（休 24h 合法）", m, { B1: ["D", "E"] }, r, null, "換班休息");
    expectClean("D→N 相鄰（休 32h 合法）", m, { B1: ["D", "N"] }, r, null, "換班休息");
    expectClean("E→N 相鄰（休 24h 合法）", m, { B1: ["E", "N"] }, r, null, "換班休息");
    // 上月尾巴 → 本月第一天
    const mt = makeModel([{ label: "B1", tail: [null, null, null, null, null, null, "N"] }], 1);
    expectFlag("上月末 N → 本月首日 D", mt, { B1: ["D"] }, baseRules(), null, "換班休息", { cell: ["B1", 0] });
  }
  // --- N-off-D ---
  {
    const r = baseRules();
    const m3 = makeModel([{ label: "B1" }], 3);
    expectFlag("N→休→D", m3, { B1: ["N", null, "D"] }, r, null, "N-off-D", { cell: ["B1", 2] });
    expectClean("N→休→E（合法）", m3, { B1: ["N", null, "E"] }, r, null, "N-off-D");
    expectClean("D→休→D（合法）", m3, { B1: ["D", null, "D"] }, r, null, "N-off-D");
    const m4 = makeModel([{ label: "B1" }], 4);
    expectClean("N→休→休→D（休 2 天合法）", m4, { B1: ["N", null, null, "D"] }, r, null, "N-off-D");
    // 跨上月邊界
    const mt1 = makeModel([{ label: "B1", tail: [null, null, null, null, null, "N", null] }], 1);
    expectFlag("上月末 N,休 → 本月首日 D", mt1, { B1: ["D"] }, r, null, "N-off-D", { cell: ["B1", 0] });
    const mt2 = makeModel([{ label: "B1", tail: [null, null, null, null, null, null, "N"] }], 2);
    expectFlag("上月末 N → 本月休,D", mt2, { B1: [null, "D"] }, r, null, "N-off-D", { cell: ["B1", 1] });
    // 銜接段內的既成違規不報、跨銜接段邊界要報
    const mc = makeModel([{ label: "B1" }], 4, 2);
    expectFlag("銜接段 N,休 → 新排 D（跨界要抓）", mc, { B1: ["N", null, "D", null] }, r, null, "N-off-D", { cell: ["B1", 2] });
    const mc2 = makeModel([{ label: "B1" }], 4, 3);
    expectClean("N,休,D 全在銜接段（既成事實不報）", mc2, { B1: ["N", null, "D", null] }, r, null, "N-off-D");
  }
  // --- 包班純班 / 未包班至多 2 種 ---
  {
    const m = makeModel([{ label: "B1" }], 3);
    const r = baseRules();
    expectFlag("包 D 卻排 E", m, { B1: ["D", null, "E"] }, r, { B1: "D" }, "包班", { cell: ["B1", 2] });
    expectClean("包 D 全月純 D", m, { B1: ["D", "D", null] }, r, { B1: "D" }, "包班");
    expectFlag("未包班用 3 種班別", m, { B1: ["D", "E", "N"] }, r, null, "班別種類");
    expectClean("未包班用 2 種班別", m, { B1: ["D", "E", null] }, r, null, "班別種類");
  }
  // --- 未包班 E+N 組合：全月上班需 <14 天（差一天可包班，薪差大；以 D+E / D+N 為主）---
  {
    const r = baseRules({ enPairMaxWork: 13 });
    const m = makeModel([{ label: "B1" }], 28);
    const en14 = Array.from({ length: 28 }, (_, i) => (i < 7 ? "E" : i < 14 ? "N" : null)); // E7+N7=14 天
    expectFlag("未包班 E+N 共上 14 天（≥14）", m, { B1: en14 }, r, null, "E+N組合");
    const en13 = en14.slice(); en13[13] = null; // E7+N6=13 天
    expectClean("未包班 E+N 共上 13 天（<14 允許）", m, { B1: en13 }, r, null, "E+N組合");
    const de20 = Array.from({ length: 28 }, (_, i) => (i % 3 === 2 ? null : i % 3 === 0 ? "D" : "E")); // D+E 19 天
    expectClean("未包班 D+E 不受 14 天限制", m, { B1: de20 }, r, null, "E+N組合");
    const dn20 = de20.map((s) => (s === "E" ? "N" : s)); // D+N，含 D→N 合法轉換
    expectClean("未包班 D+N 不受 14 天限制", m, { B1: dn20 }, r, null, "E+N組合");
    expectClean("包班 E 上 14 天不受限", m, { B1: en14.map((s) => (s ? "E" : null)) }, r, { B1: "E" }, "E+N組合");
    // 銜接段的班別不計入本月 E+N 總數
    const mc = makeModel([{ label: "B1" }], 28, 14);
    expectClean("銜接段 E+N 14 天 + 新月 0 天（不計入）", mc, { B1: en14 }, r, null, "E+N組合");
  }
  // generate 端：E/N 高壓情境下，未包班者不得形成 ≥14 天的 E+N 組合（寧缺勿破）
  {
    const ppl = Array.from({ length: 6 }, (_, i) => ({ label: "B" + (i + 1) }));
    const m = makeModel(ppl, 28);
    const r = baseRules({ enPairMaxWork: 13, minManpower: { D: 0, E: 2, N: 2 }, maxConsecutiveWork: 5, restWindows: [[14, 2], [28, 8]] });
    const asg = S.generate(m, r, {}, {});
    const flags = S.validate(m, asg, r, {});
    const combos = ppl.map((p) => {
      const seg = asg[p.label].filter((s) => SHIFTS.includes(s));
      return `${p.label}:${Array.from(new Set(seg)).sort().join("+") || "-"}(${seg.length})`;
    }).join(" ");
    check("generate 端：E/N 高壓下不形成 E+N≥14", !kindsOf(flags).has("E+N組合"), combos);
  }
  // --- 連續上班 ---
  {
    const r = baseRules({ maxConsecutiveWork: 5 });
    const m6 = makeModel([{ label: "B1" }], 6);
    expectFlag("連 6 天上班", m6, { B1: ["D", "D", "D", "D", "D", "D"] }, r, null, "連續上班");
    expectClean("連 5 天上班+1 休", m6, { B1: ["D", "D", "D", "D", "D", null] }, r, null, "連續上班");
    // 混班種連上（D..E）也要計入
    expectFlag("連 6 天上班（D/E 混班）", m6, { B1: ["D", "D", "D", "E", "E", "E"] }, r, null, "連續上班");
    // 跨上月：尾巴連 3 + 本月連 3 = 6
    const mt = makeModel([{ label: "B1", tail: [null, null, null, null, "D", "D", "D"] }], 3);
    expectFlag("上月連 3 + 本月連 3 = 連 6", mt, { B1: ["D", "D", "D"] }, r, null, "連續上班");
    expectClean("上月連 3 + 本月連 2", mt, { B1: ["D", "D", null] }, r, null, "連續上班");
  }
  // --- 連續大夜 ---
  {
    const r = baseRules({ maxConsecutiveN: 4 });
    const m5 = makeModel([{ label: "B1" }], 5);
    expectFlag("連 5 天大夜", m5, { B1: ["N", "N", "N", "N", "N"] }, r, null, "連續大夜");
    expectClean("連 4 天大夜+1 休", m5, { B1: ["N", "N", "N", "N", null] }, r, null, "連續大夜");
    const mt = makeModel([{ label: "B1", tail: [null, null, null, null, null, "N", "N"] }], 3);
    expectFlag("上月連 2N + 本月連 3N = 連 5N", mt, { B1: ["N", "N", "N"] }, r, null, "連續大夜");
  }
  // --- 休假視窗 ---
  {
    const r = baseRules({ restWindows: [[14, 2], [28, 8]], maxConsecutiveWork: 99 });
    const m14 = makeModel([{ label: "B1" }], 14);
    const work13 = Array(14).fill("D"); work13[6] = null; // 14 天只休 1 天
    expectFlag("14 天僅休 1 日（<2）", m14, { B1: work13 }, r, null, "休假");
    const work12 = Array(14).fill("D"); work12[5] = null; work12[11] = null;
    expectClean("14 天休 2 日", m14, { B1: work12 }, r, null, "休假");
    const m28 = makeModel([{ label: "B1" }], 28);
    const w28 = Array(28).fill("D");
    [3, 7, 11, 15, 19, 23, 27].forEach((i) => (w28[i] = null)); // 28 天休 7 日（滿足[14,2]但<8）
    expectFlag("28 天僅休 7 日（<8）", m28, { B1: w28 }, r, null, "休假");
    const w28ok = w28.slice(); w28ok[24] = null; // 第 8 日休
    expectClean("28 天休 8 日", m28, { B1: w28ok }, r, null, "休假");
  }
  // --- generate() 遵守 N-off-D（產生端，不只驗證端）---
  {
    const model = makeModel([{ label: "B1", tail: [null, null, null, null, null, "N", null] }], 1);
    model.people[0].active = true;
    const r = baseRules({ minManpower: { D: 1, E: 0, N: 0 } }); // 逼排 D 的壓力
    const asg = S.generate(model, r, {}, {});
    check("【N-off-D】generate 端：上月 N,休 之後寧缺 D 也不排 D",
      asg.B1[0] !== "D", `day0=${asg.B1[0] || "休"}`);
  }
  // --- 六日休假保底（generate 端）---
  {
    const model = makeModel([{ label: "B1" }], 7);
    model.weekdays = ["一", "二", "三", "四", "五", "六", "日"]; // 末兩天為週末
    const r = baseRules({ minManpower: { D: 1, E: 0, N: 0 } }); // 人力壓力：天天都想排他
    const asg = S.generate(model, r, {}, {});
    check("【六日保底】generate 端：唯一人力天天被需要，仍保證至少 1 個週末日休",
      !asg.B1[5] || !asg.B1[6], `六=${asg.B1[5] || "休"} 日=${asg.B1[6] || "休"}`);
  }
  // --- labelOffs：上班班別原樣保留、休假日留空白（不自動填例/息，由護理長自填）---
  {
    const seq = ["D", null, "D", null, "D", "D", "D", null, "D", "D", "D", "D", "D", null];
    const out = S.labelOffs(seq);
    const ok = out.every((v, i) => (S.SHIFTS.includes(seq[i]) ? v === seq[i] : v === ""));
    check("labelOffs：上班保留、休假留空白（不填例/息）", ok, out.join(","));
  }
}

/* ============ C. 端對端獨立稽核 ============ */
/* 與 scheduler.validate 完全獨立實作的稽核器（雙重驗證）。回傳 [{kind, label, day}] */
function auditAll(model, asg, rules, locks) {
  const out = [];
  const add = (kind, label, day) => out.push({ kind, label, day });
  const carryLen = model.carryLen || 0;
  const people = model.people.filter((p) => p.active);
  const N = model.nDays;
  const GROUPS = ["A", "B", "C", "D", "E"];
  const TIER_START = ["B", "C", "D", "E"];

  // 每日：最低人力 / 小組長 / 階層上限
  for (let d = carryLen; d < N; d++) {
    const cnt = { D: 0, E: 0, N: 0 }, aCnt = { D: 0, E: 0, N: 0 };
    const tierCnt = { D: [0, 0, 0, 0], E: [0, 0, 0, 0], N: [0, 0, 0, 0] };
    people.forEach((p) => {
      const s = asg[p.label][d];
      if (!SHIFTS.includes(s)) return;
      cnt[s]++;
      if (p.group === "A") aCnt[s]++;
      TIER_START.forEach((st, i) => { if (GROUPS.indexOf(p.group) >= GROUPS.indexOf(st)) tierCnt[s][i]++; });
    });
    SHIFTS.forEach((s) => {
      if (cnt[s] < rules.minManpower[s]) add("最低人力", null, d);
      if (aCnt[s] < rules.minLeaderPerShift) add("小組長", null, d);
      rules.tierCaps[s].forEach((cap, i) => { if (tierCnt[s][i] > cap) add("階層上限", null, d); });
    });
  }

  people.forEach((p) => {
    const lab = p.label;
    const ext = p.tail.concat(asg[lab]); // tail + 本期
    const base = p.tail.length;
    const isNew = (i) => i - base >= carryLen;
    // 換班休息（相鄰）+ N-off-D
    for (let i = 1; i < ext.length; i++) {
      if (!isNew(i)) continue;
      const prev = ext[i - 1], cur = ext[i];
      if (prev && cur && (rules.forbidTransition[prev] || []).includes(cur)) add("換班休息", lab, i - base);
      if (rules.forbidNOffD && cur === "D" && i >= 2 && !ext[i - 1] && ext[i - 2] === "N") add("N-off-D", lab, i - base);
    }
    // 連續上班 / 連續大夜（跨尾巴與銜接段）
    let run = 0, nrun = 0;
    for (let i = 0; i < ext.length; i++) {
      run = ext[i] ? run + 1 : 0;
      nrun = ext[i] === "N" ? nrun + 1 : 0;
      if (isNew(i)) {
        if (run > rules.maxConsecutiveWork) add("連續上班", lab, i - base);
        if (nrun > rules.maxConsecutiveN) add("連續大夜", lab, i - base);
      }
    }
    // 包班純班 / 班別種類（僅新產生天數）
    const lock = locks[lab];
    const newSeq = asg[lab].slice(carryLen);
    if (lock) {
      newSeq.forEach((s, i) => { if (SHIFTS.includes(s) && s !== lock) add("包班", lab, carryLen + i); });
    } else {
      const kinds = new Set(newSeq.filter((s) => SHIFTS.includes(s)));
      if (kinds.size > rules.maxShiftKinds) add("班別種類", lab, null);
    }
    // 休假視窗（滑動，凡覆蓋到新天數的視窗都查）
    const off = asg[lab].map((s) => (SHIFTS.includes(s) ? 0 : 1));
    (rules.restWindows || []).forEach(([w, mo]) => {
      if (w > N || !mo) return;
      for (let st = Math.max(0, carryLen - w + 1); st <= N - w; st++) {
        let sum = 0;
        for (let i = st; i < st + w; i++) sum += off[i];
        if (sum < mo) { add("休假", lab, st); break; }
      }
    });
    // 工時：每日 <=10h（單班 8h、一天最多一班）、任 28 天 <=160h
    for (let d = carryLen; d < N; d++) {
      const dayHours = SHIFTS.includes(asg[lab][d]) ? HOURS_PER_SHIFT : 0;
      if (dayHours > 10) add("工時-日", lab, d);
    }
    if (N >= 28) {
      for (let st = Math.max(0, carryLen - 27); st <= N - 28; st++) {
        let hrs = 0;
        for (let i = st; i < st + 28; i++) if (SHIFTS.includes(asg[lab][i])) hrs += HOURS_PER_SHIFT;
        if (hrs > 160) { add("工時-4週", lab, st); break; }
      }
    }
  });
  return out;
}

/* 稽核 vs validate 交叉比對：
 * 稽核發現的每一種 (kind, label|day) 違規，validate 也必須報到；反之亦然。 */
const KIND_MAP = { // audit kind -> validate tag（工時為結構保證，validate 無對應標籤）
  "最低人力": "最低人力", "小組長": "小組長", "階層上限": "階層上限",
  "換班休息": "換班休息", "N-off-D": "N-off-D", "連續上班": "連續上班",
  "連續大夜": "連續大夜", "包班": "包班", "班別種類": "班別種類", "休假": "休假",
};
function crossCheck(name, model, asg, rules, locks) {
  const audit = auditAll(model, asg, rules, locks);
  const flags = S.validate(model, asg, rules, locks);
  const vKeys = new Set();
  flags.errors.forEach((e) => {
    const m = e.match(/^\[([^\]]+)\]/);
    if (m) vKeys.add(m[1]);
  });
  flags.cellFlags.forEach((f) => { const m = f.msg.match(/^\[([^\]]+)\]/); if (m) vKeys.add(m[1] + "|" + f.label); });
  flags.rowFlags && Object.keys(flags.rowFlags).forEach((l) =>
    (flags.rowFlags[l] || []).forEach((e) => { const m = e.match(/^\[([^\]]+)\]/); if (m) vKeys.add(m[1] + "|" + l); }));

  const missed = [];
  audit.forEach((a) => {
    const tag = KIND_MAP[a.kind];
    if (!tag) return; // 工時類：validate 無此標籤（由休假視窗結構保證），下方單獨斷言為 0
    const hit = a.label ? (vKeys.has(tag + "|" + a.label) || vKeys.has(tag)) : vKeys.has(tag);
    if (!hit) missed.push(`${a.kind}@${a.label || "day" + a.day}`);
  });
  const auditKinds = new Set(audit.map((a) => a.kind));
  const extra = [];
  Array.from(vKeys).forEach((k) => {
    const kind = k.split("|")[0];
    if (KIND_MAP[kind] && !auditKinds.has(kind)) extra.push(kind);
  });
  check(`交叉比對(${name})：獨立稽核發現的違規 validate 全數報到`, missed.length === 0,
    missed.length ? "遺漏：" + missed.slice(0, 5).join("、") : `稽核 ${audit.length} 項均對應`);
  check(`交叉比對(${name})：validate 未憑空多報`, extra.length === 0,
    extra.length ? "多報：" + Array.from(new Set(extra)).join("、") : "一致");
  return { audit, flags };
}

async function partC() {
  section = "C.端對端";
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(SRC);
  const ws = wb.worksheets[0];

  function freshModel() {
    const model = S.parseWorkbook(ws);
    // 依 app.js 流程：推算下一週期（月+1、同起始日、同視窗長度）
    let ny = model.detectedYear, nm = model.detectedStartMonth + 1, nd = model.detectedStartDay;
    if (nm > 12) { nm -= 12; ny += 1; }
    const cal = S.buildCalendar(ny, nm, nd, model.nDays);
    const tailCal = S.buildTail(ny, nm, nd, model.tailLen);
    model.dates = cal.map((c) => c.d);
    model.weekdays = cal.map((c) => c.w);
    model.monthsByDay = cal.map((c) => c.m);
    const upCal = S.buildCalendar(model.detectedYear, model.detectedStartMonth, model.detectedStartDay, model.nDays);
    const carry = S.buildCarry(model, upCal, cal, tailCal);
    model.carryLen = carry.overlapLen;
    model.people.forEach((p) => { if (carry.history[p.label]) p.tail = carry.history[p.label]; });
    return { model, carry, cal };
  }
  function locksOf(model) {
    const locks = {};
    model.people.forEach((p) => { if (p.active && p.suggestedLock) locks[p.label] = p.suggestedLock; });
    return locks;
  }
  function fixedOf(model, carry, extraOff) {
    const fixed = {};
    Object.keys(extraOff || {}).forEach((l) => (fixed[l] = Object.assign({}, extraOff[l])));
    Object.keys(carry.preset).forEach((l) => {
      fixed[l] = Object.assign({}, fixed[l] || {});
      Object.assign(fixed[l], carry.preset[l]);
    });
    return fixed;
  }

  // --- 情境 1：無預假（正常月）→ 必須 0 違規 ---
  {
    const { model, carry } = freshModel();
    const locks = locksOf(model);
    const asg = S.generate(model, S.DEFAULT_RULES, fixedOf(model, carry), locks);
    const { audit } = crossCheck("無預假", model, asg, S.DEFAULT_RULES, locks);
    check("情境[無預假]：獨立稽核 0 違規（含工時/N-off-D/階層/休假全項）", audit.length === 0,
      audit.length ? audit.slice(0, 6).map((a) => a.kind + "@" + (a.label || "day" + a.day)).join("、") : "全部通過");

    // 休假分布結構：新月份內每個 7 日區塊至少 1 天休假；14 日區塊的兩個子區塊各至少 1 天休假
    // （對應需求「2週內至少2日例假」。休假日已改為留空，直接以「非上班」計算）
    const carryLen = model.carryLen;
    let blk7ok = true, blk14ok = true, detail = "";
    const isOff = (x) => !S.SHIFTS.includes(x);
    model.people.filter((p) => p.active).forEach((p) => {
      const seq = asg[p.label];
      for (let st = Math.ceil(carryLen / 7) * 7; st + 7 <= model.nDays; st += 7) {
        const c = seq.slice(st, st + 7).filter(isOff).length;
        if (c < 1) { blk7ok = false; detail = detail || `${p.label}@${st}`; }
      }
      for (let st = Math.ceil(carryLen / 14) * 14; st + 14 <= model.nDays; st += 14) {
        const a = seq.slice(st, st + 7).filter(isOff).length;
        const b = seq.slice(st + 7, st + 14).filter(isOff).length;
        if (a < 1 || b < 1) { blk14ok = false; detail = detail || `${p.label}@${st}`; }
      }
    });
    check("情境[無預假]：每 7 日區塊 ≥1 天休假", blk7ok, detail || "全員符合");
    check("情境[無預假]：每 14 日區塊兩子區塊各 ≥1 天休假（2週≥2休）", blk14ok, detail || "全員符合");

    // 六日休假平均（軟性需求：盡量平均）→ 量測新月份的週末休假次數散佈
    const wkIdx = [];
    for (let d = carryLen; d < model.nDays; d++) if (["六", "日"].includes(model.weekdays[d])) wkIdx.push(d);
    const offs = model.people.filter((p) => p.active).map((p) => wkIdx.filter((d) => !asg[p.label][d]).length);
    const mn = Math.min(...offs), mx = Math.max(...offs);
    check(`情境[無預假]：六日休假保底——無人為 0（週末日共 ${wkIdx.length} 天）`, mn >= 1, `min=${mn}`);
    // 軟性需求「盡量平均」驗收：散佈不劣於人工排班（上傳檔 6 月為 min1/max5）
    check("情境[無預假]：六日休假次數盡量平均（軟性）", mx - mn <= 4,
      `min=${mn} max=${mx}（差 ${mx - mn}；對照人工排班 6 月：min1 max5 差 4）`);
  }

  // --- 情境 2：痛點壓力（同一天多人預假）→ 允許人力缺口，但硬規則不得破 ---
  {
    const { model, carry } = freshModel();
    const locks = locksOf(model);
    const carryLen = model.carryLen;
    // 取 18 位人員在新月份同一天預假（模擬「同一天太多人預假」痛點）
    const day = carryLen + 9;
    const extra = {};
    model.people.filter((p) => p.active).slice(0, 18).forEach((p) => (extra[p.label] = { [day]: "OFF" }));
    const asg = S.generate(model, S.DEFAULT_RULES, fixedOf(model, carry, extra), locks);
    const { audit } = crossCheck("同日18人預假", model, asg, S.DEFAULT_RULES, locks);
    const kinds = new Set(audit.map((a) => a.kind));
    const allowed = new Set(["最低人力", "小組長"]);
    const bad = Array.from(kinds).filter((k) => !allowed.has(k));
    check("情境[同日18人預假]：預假日全數未被排班", model.people.filter((p) => extra[p.label]).every((p) => !asg[p.label][day]), "18 人該日皆休");
    check("情境[同日18人預假]：僅允許人力/小組長缺口，其餘硬規則不得破（工時/連上/換班/休假/階層）",
      bad.length === 0, bad.length ? "破規：" + bad.join("、") : `違規種類=${Array.from(kinds).join("、") || "無"}`);
    const wkIdx2 = [];
    for (let d = carryLen; d < model.nDays; d++) if (["六", "日"].includes(model.weekdays[d])) wkIdx2.push(d);
    const wkMin2 = Math.min(...model.people.filter((p) => p.active).map((p) => wkIdx2.filter((d) => !asg[p.label][d]).length));
    check("情境[同日18人預假]：六日休假保底——無人為 0", wkMin2 >= 1, `min=${wkMin2}`);
  }

  // --- 情境 3：隨機預假（固定種子，可重現）---
  {
    const { model, carry } = freshModel();
    const locks = locksOf(model);
    const carryLen = model.carryLen;
    const rnd = mulberry32(20260705);
    const extra = {};
    model.people.filter((p) => p.active).forEach((p) => {
      const n = Math.floor(rnd() * 3); // 0~2 天
      for (let k = 0; k < n; k++) {
        const d = carryLen + Math.floor(rnd() * (model.nDays - carryLen));
        extra[p.label] = extra[p.label] || {};
        extra[p.label][d] = "OFF";
      }
    });
    const asg = S.generate(model, S.DEFAULT_RULES, fixedOf(model, carry, extra), locks);
    const { audit } = crossCheck("隨機預假", model, asg, S.DEFAULT_RULES, locks);
    const kinds = new Set(audit.map((a) => a.kind));
    const notAllowed = Array.from(kinds).filter((k) => k !== "最低人力" && k !== "小組長");
    check("情境[隨機預假]：預假全數生效", Object.keys(extra).every((l) => Object.keys(extra[l]).every((d) => !asg[l][+d])), `${Object.keys(extra).length} 人設定預假`);
    check("情境[隨機預假]：硬規則不得破（允許人力缺口）", notAllowed.length === 0,
      notAllowed.length ? "破規：" + notAllowed.join("、") : `違規種類=${Array.from(kinds).join("、") || "無"}`);
    const wkIdx3 = [];
    for (let d = carryLen; d < model.nDays; d++) if (["六", "日"].includes(model.weekdays[d])) wkIdx3.push(d);
    const wkMin3 = Math.min(...model.people.filter((p) => p.active).map((p) => wkIdx3.filter((d) => !asg[p.label][d]).length));
    check("情境[隨機預假]：六日休假保底——無人為 0", wkMin3 >= 1, `min=${wkMin3}`);
  }
}

/* ============ E. 產出檔日期列底色：換月份時才換色 ============ */
async function partE() {
  section = "E.輸出底色";
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(SRC);
  const ws = wb.worksheets[0];
  const model = S.parseWorkbook(ws);
  let ny = model.detectedYear, nm = model.detectedStartMonth + 1, nd = model.detectedStartDay;
  if (nm > 12) { nm -= 12; ny += 1; }
  const cal = S.buildCalendar(ny, nm, nd, model.nDays);
  const tailCal = S.buildTail(ny, nm, nd, model.tailLen);
  model.dates = cal.map((c) => c.d);
  model.weekdays = cal.map((c) => c.w);
  model.monthsByDay = cal.map((c) => c.m);
  const upCal = S.buildCalendar(model.detectedYear, model.detectedStartMonth, model.detectedStartDay, model.nDays);
  const carry = S.buildCarry(model, upCal, cal, tailCal);
  model.carryLen = carry.overlapLen;
  model.people.forEach((p) => { if (carry.history[p.label]) p.tail = carry.history[p.label]; });
  const locks = {};
  model.people.forEach((p) => { if (p.active && p.suggestedLock) locks[p.label] = p.suggestedLock; });
  const fixed = {};
  Object.keys(carry.preset).forEach((l) => (fixed[l] = Object.assign({}, carry.preset[l])));
  const asg = S.generate(model, S.DEFAULT_RULES, fixed, locks);
  S.writeCalendar(ws, model, { rocYear: ny, month: nm, day: nd });
  S.writeInto(ws, model, asg, S.DEFAULT_RULES, carry.rawByDay, {});

  const fillSig = (r, c) => {
    const f = ws.getCell(r, c).style && ws.getCell(r, c).style.fill;
    return f && f.fgColor ? JSON.stringify(f.fgColor) : "none";
  };
  // 找出所有「日期」列（上方主日期列 + 底部驗證區日期列）
  const dateRows = [model.dateRow];
  for (let r = model.dateRow + 1; r <= (ws.rowCount || 300); r++) {
    let t = ws.getCell(r, 3).value;
    t = t == null ? "" : String(typeof t === "object" ? (t.result != null ? t.result : "") : t).trim();
    if (t === "日期") dateRows.push(r);
  }
  check("輸出檔含底部驗證區日期列（同步對象）", dateRows.length >= 2, `日期列 rows=${dateRows.join(",")}`);
  dateRows.forEach((r) => {
    const sig = model.dateCols.map((c) => fillSig(r, c));
    let ok = true, detail = "";
    for (let i = 1; i < model.nDays; i++) {
      const changed = sig[i] !== sig[i - 1];
      const monthChanged = cal[i].d === 1;
      if (changed !== monthChanged) { ok = false; detail = `第${i + 1}格(${cal[i].m}/${cal[i].d}) 變色=${changed} 換月=${monthChanged}`; break; }
    }
    check(`日期列(列${r})底色僅在換月份處變色`, ok, detail || `月界=${cal.filter((c, i) => i > 0 && c.d === 1).map((c) => c.m + "/1").join(",")}`);
  });
  const sigTop = model.dateCols.map((c) => fillSig(dateRows[0], c)).join("|");
  const sigBot = model.dateCols.map((c) => fillSig(dateRows[dateRows.length - 1], c)).join("|");
  check("上方與底部日期列底色一致", sigTop === sigBot, sigTop === sigBot ? "一致" : "不一致");

  const wbM = new ExcelJS.Workbook();
  await wbM.xlsx.readFile(SRC_MONTHLY);
  const wsM = wbM.worksheets[0];
  const modelM = S.parseWorkbook(wsM);
  const originalDays = modelM.nDays;
  const targetDays = Math.round((new Date(115 + 1911, 8, 0) - new Date(115 + 1911, 7 - 1, 10)) / 864e5) + 1;
  const insertCount = targetDays - originalDays;
  S.expandTemplateColumns(wsM, modelM, insertCount);
  const outBuf = await wbM.xlsx.writeBuffer();
  const zip = await JSZip.loadAsync(outBuf);
  const sheetName = Object.keys(zip.files).find((n) => /^xl\/worksheets\/sheet\d+[.]xml$/.test(n));
  const xml = await zip.file(sheetName).async("string");
  const dvSqrefs = Array.from(xml.matchAll(/<dataValidation\b[^>]*\bsqref="([^"]+)"/g)).map((m) => m[1]).sort();
  const cfSqrefs = Array.from(xml.matchAll(/<conditionalFormatting\b[^>]*\bsqref="([^"]+)"/g)).map((m) => m[1]).sort();
  check("單月表加寬後資料驗證不重疊、不殘留舊欄", dvSqrefs.length === 6 &&
    dvSqrefs.includes("D80:BD80") && dvSqrefs.includes("BO3:BO4") &&
    dvSqrefs.includes("BO79:BO80") && dvSqrefs.includes("BO82:BO103") &&
    dvSqrefs.includes("BR2:BR80") && dvSqrefs.includes("BR82:BR103"),
    `sqref=${dvSqrefs.join(" | ")}`);
  check("單月表加寬後條件格式平移到右側檢核欄", cfSqrefs.includes("BI5:BI76") &&
    cfSqrefs.includes("BL5:BL76 BU5:BU76") && cfSqrefs.includes("BO5:BO76 BV5:BV76") &&
    cfSqrefs.includes("BP5:BQ76") && cfSqrefs.includes("BR5:BR76"),
    `sqref=${cfSqrefs.join(" | ")}`);

  // 範本最後一條 <col> 涵蓋到工作表末欄(16384)；插欄若把它一起右移，max 會超過 Excel 欄上限，
  // Excel 便判檔案損毀、拒絕開啟——單月表輸出打不開的主因。
  const colDefs = Array.from(xml.matchAll(/<col [^>]*min="(\d+)"[^>]*max="(\d+)"[^>]*width="([\d.]+)"/g))
    .map((m) => ({ min: +m[1], max: +m[2], width: +m[3] }));
  const overflow = colDefs.filter((c) => c.max > 16384);
  check("單月表加寬後欄定義不超過 Excel 欄上限 16384（否則 Excel 拒絕開檔）", overflow.length === 0,
    overflow.length ? `越界 ${JSON.stringify(overflow)}` : `末條 max=${colDefs[colDefs.length - 1].max}`);
  // spliceColumns 把新插入欄的定義清成 null；未補回則銜接段欄寬縮成預設值
  const covering = (col) => colDefs.find((c) => c.min <= col && c.max >= col);
  const firstIns = covering(modelM.firstCol), refCol = covering(modelM.firstCol + insertCount);
  check("單月表加寬後銜接段欄沿用日期欄寬", !!firstIns && !!refCol && firstIns.width === refCol.width,
    firstIns && refCol ? `插入欄 width=${firstIns.width}，日期欄 width=${refCol.width}` : "缺欄寬定義");
}

/* ============ D. 歷史資料規則解讀佐證 ============ */
async function partD() {
  section = "D.歷史佐證";
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(SRC);
  const ws = wb.worksheets[0];
  const model = S.parseWorkbook(ws);
  // 用上傳檔全部 52 天的實際班別（前綴語意）檢視醫院實務
  const trans = {}, gap1 = {};
  model.people.forEach((p) => {
    const seq = p.rowRaw.map((t) => S.classifyCell(t));
    for (let i = 1; i < seq.length; i++)
      if (seq[i - 1] && seq[i] && seq[i - 1] !== seq[i]) trans[seq[i - 1] + ">" + seq[i]] = (trans[seq[i - 1] + ">" + seq[i]] || 0) + 1;
    for (let i = 2; i < seq.length; i++)
      if (seq[i - 2] && !seq[i - 1] && seq[i]) gap1[seq[i - 2] + "-off-" + seq[i]] = (gap1[seq[i - 2] + "-off-" + seq[i]] || 0) + 1;
  });
  check("歷史班表 0 次 N→D / N→E / E→D 相鄰（佐證 11 小時禁接規則）",
    !trans["N>D"] && !trans["N>E"] && !trans["E>D"],
    "相鄰轉換：" + (Object.entries(trans).map(([k, v]) => `${k}×${v}`).join(" ") || "無"));
  check("歷史班表 0 次 N-off-D（佐證「不能 N off D」為三天模式）",
    !gap1["N-off-D"],
    "隔1休模式：" + Object.entries(gap1).map(([k, v]) => `${k}×${v}`).join(" "));
}

/* ============ 主流程 ============ */
(async () => {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(SRC);
  const ws = wb.worksheets[0];
  const model = S.parseWorkbook(ws);

  partA(ws, model);
  partB();
  await partC();
  await partD();
  await partE();

  console.log("=".repeat(72));
  console.log("需求規則驗證報告（案例_醫事人員排班_V2.docx ↔ scheduler.js）");
  console.log("=".repeat(72));
  let ok = true, lastSec = "";
  checks.forEach(([sec, name, pass, det]) => {
    if (sec !== lastSec) { console.log("--- " + sec + " ---"); lastSec = sec; }
    console.log(`${pass ? "✅" : "❌"}  ${name}${det ? "　·　" + det : ""}`);
    ok = ok && pass;
  });
  console.log("=".repeat(72));
  const nPass = checks.filter((c) => c[2]).length;
  console.log(`結果：${ok ? "✅ 全部通過" : "❌ 有項目未通過"}（${nPass}/${checks.length}）`);
  process.exit(ok ? 0 : 1);
})().catch((e) => { console.error("執行失敗:", e); process.exit(1); });
