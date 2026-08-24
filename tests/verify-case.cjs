const fs = require("fs");
const assert = require("assert");
const ExcelJS = require("../vendor/exceljs.min.js");

function value(sheet, column, row) {
  const cellValue = sheet.getCell(`${column}${row}`).value;
  if (cellValue && typeof cellValue === "object" && "result" in cellValue) return cellValue.result;
  if (cellValue && typeof cellValue === "object" && ("formula" in cellValue || "ref" in cellValue)) return null;
  return cellValue;
}

(async () => {
  const sourceName = fs.readdirSync(".").find((name) => name.includes("1150823-02") && name.endsWith(".xlsx"));
  assert(sourceName, "case source workbook not found");
  const sourceBook = new ExcelJS.Workbook();
  const templateBook = new ExcelJS.Workbook();
  await sourceBook.xlsx.load(fs.readFileSync(sourceName));
  await templateBook.xlsx.load(fs.readFileSync("案件空白2.xlsx"));
  const source = sourceBook.worksheets[0];
  const target = templateBook.worksheets[0];
  const headerStyle = target.getCell("B1").style;
  const common = [["C", 2, "B"], ["H", 10, "C"], ["H", 3, "D"], ["D", 3, "E"], ["C", 4, "F"], ["C", 17, "G"], ["H", 5, "T"]];
  let targetRow = 2;

  for (let sourceRow = 28; sourceRow <= source.rowCount; sourceRow += 2) {
    const vendor = value(source, "B", sourceRow) ?? value(source, "M", sourceRow);
    if (vendor === null || vendor === undefined || String(vendor).trim() === "") continue;
    for (const [column, row, targetColumn] of common) target.getCell(`${targetColumn}${targetRow}`).value = value(source, column, row);
    target.getCell(`H${targetRow}`).value = vendor;
    target.getCell(`P${targetRow}`).value = value(source, "C", sourceRow);
    target.getCell(`U${targetRow}`).value = value(source, "D", sourceRow);
    targetRow += 1;
  }

  const output = await templateBook.xlsx.writeBuffer();
  const roundTrip = new ExcelJS.Workbook();
  await roundTrip.xlsx.load(output);
  const result = roundTrip.worksheets[0];

  assert.strictEqual(targetRow - 2, 2);
  assert.strictEqual(result.getCell("B2").value, "林文洪");
  assert.strictEqual(result.getCell("C2").value, "含軟硬體銷售");
  assert.strictEqual(result.getCell("D2").value, "NO2026082302");
  assert.strictEqual(result.getCell("H2").value, "精技");
  assert.strictEqual(result.getCell("U2").value, 2363);
  assert.strictEqual(result.getCell("H3").value, "原價屋");
  assert.strictEqual(result.getCell("P3").value, "DD68551026");
  assert.strictEqual(result.getCell("U3").value, 70370);
  assert.strictEqual(result.getCell("T2").value, 76500);
  assert.deepStrictEqual(result.getCell("B1").style, headerStyle);

  console.log("PASS: 2 vendors; case mappings and header styles survived XLSX round-trip");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
