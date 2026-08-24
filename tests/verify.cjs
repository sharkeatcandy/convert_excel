const fs = require("fs");
const assert = require("assert");
const ExcelJS = require("../vendor/exceljs.min.js");

(async () => {
  const sourceBook = new ExcelJS.Workbook();
  const templateBook = new ExcelJS.Workbook();
  await sourceBook.xlsx.load(fs.readFileSync("資料.xlsx"));
  await templateBook.xlsx.load(fs.readFileSync("空白檔.xlsx"));
  const source = sourceBook.worksheets[0];
  const target = templateBook.worksheets[0];
  const headerStyle = target.getCell("A1").style;
  const map = [["C", "C"], ["V", "M"], ["Z", "V"], ["AD", "S"]];
  let targetRow = 2;

  for (let sourceRow = 4; sourceRow <= source.rowCount; sourceRow += 1) {
    const get = (column) => source.getCell(`${column}${sourceRow}`).value;
    if (!["A", "B", "C", "S", "T", "V", "W", "Z", "AD", "AE", "AN"].some((column) => get(column) !== null && get(column) !== undefined && get(column) !== "")) continue;
    const match = String(get("B")).match(/^(\d{4})年(\d{1,2})月份?$/);
    assert(match, `invalid period at row ${sourceRow}`);
    for (const [from, to] of map) target.getCell(`${to}${targetRow}`).value = get(from);
    target.getCell(`L${targetRow}`).value = (get("S") || 0) + (get("T") || 0) - (get("AE") || 0);
    target.getCell(`N${targetRow}`).value = (get("W") || 0) + (get("AN") || 0);
    target.getCell(`J${targetRow}`).value = Number(match[1]) - 1911;
    target.getCell(`K${targetRow}`).value = Number(match[2]);
    targetRow += 1;
  }

  const output = await templateBook.xlsx.writeBuffer();
  const roundTrip = new ExcelJS.Workbook();
  await roundTrip.xlsx.load(output);
  const result = roundTrip.worksheets[0];

  assert.strictEqual(targetRow - 2, 93);
  assert.strictEqual(result.getCell("C2").value, source.getCell("C4").value);
  assert.strictEqual(result.getCell("J2").value, 115);
  assert.strictEqual(result.getCell("K2").value, 1);
  for (const [from, to] of map.slice(1)) assert.strictEqual(result.getCell(`${to}2`).value, source.getCell(`${from}4`).value);
  assert.strictEqual(result.getCell("L2").value, (source.getCell("S4").value || 0) + (source.getCell("T4").value || 0) - (source.getCell("AE4").value || 0));
  assert.strictEqual(result.getCell("N2").value, (source.getCell("W4").value || 0) + (source.getCell("AN4").value || 0));
  assert.strictEqual(result.getCell("J94").value, 115);
  assert.strictEqual(result.getCell("K94").value, 7);
  assert.deepStrictEqual(result.getCell("A1").style, headerStyle);

  console.log("PASS: 93 rows; mappings, ROC periods, and header styles survived XLSX round-trip");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
