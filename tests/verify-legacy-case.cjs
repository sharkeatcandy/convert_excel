const fs = require("fs");
const assert = require("assert");
const XLSX = require("../vendor/xlsx.full.min.js");
const ExcelJS = require("../vendor/exceljs.min.js");

(async () => {
  const sourceName = fs.readdirSync(".").find((name) => name.endsWith("TC300熱感紙.xls"));
  assert(sourceName, "legacy case source workbook not found");
  const sourceBook = XLSX.read(fs.readFileSync(sourceName), { type: "buffer", cellDates: true });
  const source = sourceBook.Sheets[sourceBook.SheetNames[0]];
  const templateBook = new ExcelJS.Workbook();
  await templateBook.xlsx.load(fs.readFileSync("案件空白2.xlsx"));
  const target = templateBook.worksheets[0];
  const headerStyle = target.getCell("B1").style;
  const get = (address) => source[address]?.v ?? null;
  const projectName = String(get("C5") ?? "");
  const candidate = projectName.split("_")[2] || "";
  const orderNumber = /^\d{10}$/.test(candidate) ? candidate : null;
  const common = [["C2", "B"], ["H10", "C"], ["H3", "D"], ["C5", "E"], ["C4", "F"], ["H5", "T"]];
  const range = XLSX.utils.decode_range(source["!ref"]);
  let targetRow = 2;
  const warnings = [];

  for (let sourceRow = 24; sourceRow <= range.e.r + 1; sourceRow += 1) {
    const vendor = get(`B${sourceRow}`);
    if (vendor === null || vendor === undefined || String(vendor).trim() === "") continue;
    for (const [sourceAddress, targetColumn] of common) target.getCell(`${targetColumn}${targetRow}`).value = get(sourceAddress);
    target.getCell(`G${targetRow}`).value = orderNumber;
    target.getCell(`H${targetRow}`).value = vendor;
    const invoiceSource = String(get(`F${sourceRow}`) ?? "");
    const invoiceMatch = invoiceSource.match(/\b[A-Za-z]{2}\d{8}\b/);
    if (invoiceMatch) target.getCell(`P${targetRow}`).value = invoiceMatch[0];
    else warnings.push(`F${sourceRow}:${invoiceSource}`);
    target.getCell(`U${targetRow}`).value = get(`D${sourceRow}`);
    targetRow += 1;
  }

  const output = await templateBook.xlsx.writeBuffer();
  const roundTrip = new ExcelJS.Workbook();
  await roundTrip.xlsx.load(output);
  const result = roundTrip.worksheets[0];

  assert.strictEqual(targetRow - 2, 1);
  assert.deepStrictEqual(warnings, []);
  assert.strictEqual(result.getCell("B2").value, "林文洪");
  assert.strictEqual(result.getCell("C2").value, "環保耗材銷售");
  assert.strictEqual(result.getCell("D2").value, "NO2026081803");
  assert.strictEqual(result.getCell("E2").value, "冠新_新光醫院_6300029697_門診_TC300熱感紙");
  assert.strictEqual(result.getCell("F2").value, "新光醫院");
  assert.strictEqual(result.getCell("G2").value, "6300029697");
  assert.strictEqual(result.getCell("H2").value, "金驊");
  assert.strictEqual(result.getCell("P2").value, "ZY19557293");
  assert.strictEqual(result.getCell("T2").value, 5400);
  assert.strictEqual(result.getCell("U2").value, 3150);
  assert.deepStrictEqual(result.getCell("B1").style, headerStyle);

  console.log("PASS: 1 legacy vendor; identifiers, mappings, and header styles survived XLSX round-trip");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
