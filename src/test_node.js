/* Node 端自我驗證：用 ExcelJS 讀真實班表，跑純 JS 排班+驗證+寫回，檢查公式/格式保留。
 * 執行：node web/test_node.js   （需先 npm install exceljs）
 */
const path = require("path");
const ExcelJS = require("exceljs");
const S = require("./scheduler.js");

const SRC = path.resolve(__dirname, "..", "sample.xlsx");
const OUT = path.resolve(__dirname, "..", "output_test.xlsx");

function countFormulasAndStyles(ws) {
  let f = 0, styled = 0;
  ws.eachRow({ includeEmpty: true }, (row) =>
    row.eachCell({ includeEmpty: true }, (c) => {
      if (c.formula) f++;
      if (c.fill && c.fill.type) styled++;
    })
  );
  return { f, styled };
}

(async () => {
  const checks = [];
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(SRC);
  const ws = wb.worksheets[0];

  const model = S.parseWorkbook(ws);
  const active = model.people.filter((p) => p.active).length;
  checks.push(["自動偵測排班天數", model.nDays > 0, `n_days=${model.nDays}`]);
  checks.push(["偵測人員數", active > 0, `active=${active}`]);

  const before = countFormulasAndStyles(ws);

  // 依上月班表自動偵測包班（同班別 >15 天 → 純班）
  const locks = {};
  model.people.forEach((p) => { if (p.active && p.suggestedLock) locks[p.label] = p.suggestedLock; });
  checks.push(["偵測包班人員", true, `locked=${Object.keys(locks).length}`]);

  const t0 = Date.now();
  const assignment = S.generate(model, S.DEFAULT_RULES, {}, locks);
  const genMs = Date.now() - t0;
  checks.push(["排班產生完成", true, `${genMs} ms`]);

  const { passed, errors } = S.validate(model, assignment, S.DEFAULT_RULES, locks);
  checks.push(["排班符合所有硬規則", passed, passed ? "無違規" : `${errors.length} 項違規`]);

  // 寫回 + 存檔
  S.writeInto(ws, model, assignment, S.DEFAULT_RULES);
  await wb.xlsx.writeFile(OUT);

  // 重讀驗證公式/格式保留
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.readFile(OUT);
  const ws2 = wb2.worksheets[0];
  const after = countFormulasAndStyles(ws2);
  checks.push(["輸出檔公式完整保留", after.f >= before.f, `${before.f} -> ${after.f}`]);
  checks.push(["輸出檔樣式完整保留", after.styled >= before.styled, `${before.styled} -> ${after.styled}`]);

  const model2 = S.parseWorkbook(ws2);
  checks.push(["輸出檔天數一致", model2.nDays === model.nDays, `${model.nDays} == ${model2.nDays}`]);

  // 寫入/讀出一致
  let mism = 0;
  model2.people.forEach((p) => {
    if (!assignment[p.label]) return;
    const rb = model2.dateCols.map((c) => S.classifyCell(ws2.getCell(p.row, c).value));
    if (JSON.stringify(rb) !== JSON.stringify(assignment[p.label])) mism++;
  });
  checks.push(["寫入/讀出班別一致", mism === 0, mism === 0 ? "完全一致" : `${mism} 位不一致`]);

  console.log("=".repeat(58));
  console.log("純前端(JS) 排班自我驗證報告");
  console.log("=".repeat(58));
  let ok = true;
  checks.forEach(([n, pass, det]) => {
    console.log(`${pass ? "✅" : "❌"}  ${n.padEnd(18)} ${det}`);
    ok = ok && pass;
  });
  console.log("-".repeat(58));
  console.log("每日人力摘要（前 7 天）");
  S.summarizeDaily(model, assignment).slice(0, 7).forEach((r) =>
    console.log(`  第${String(r.day).padStart(2)}天  D=${String(r.D).padStart(2)} E=${String(r.E).padStart(2)} N=${String(r.N).padStart(2)} 休=${String(r.off).padStart(2)}`)
  );
  if (!passed) {
    console.log("-".repeat(58));
    console.log("違規明細（前 15 條）：");
    errors.slice(0, 15).forEach((e) => console.log("  -", e));
  }
  console.log("=".repeat(58));
  console.log("結果：", ok ? "✅ 全部通過" : "❌ 有項目未通過");
  process.exit(ok ? 0 : 1);
})();
