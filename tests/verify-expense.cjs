const fs = require("fs");
const assert = require("assert");
const ExcelJS = require("../vendor/exceljs.min.js");

(async () => {
  const sourceBook = new ExcelJS.Workbook();
  const templateBook = new ExcelJS.Workbook();
  await sourceBook.xlsx.load(fs.readFileSync("支出單資料.xlsx"));
  await templateBook.xlsx.load(fs.readFileSync("1精技支出申請單2.xlsx"));
  const source = sourceBook.worksheets[0];
  const target = templateBook.worksheets[0];
  const expectedDocumentDate = source.getCell("A2").value;
  const headerStyle = target.getCell("B5").style;
  const originalVendorStyles = [];
  for (let row = 7; row <= 12; row += 1) originalVendorStyles.push(JSON.parse(JSON.stringify(target.getCell(`C${row}`).style || {})));
  const samples = [
    ["精技", "115.08.01", "AB12345678", 1000, "NO-001"],
    ["精技", "115.08.02", "CD12345678", 2000, "NO-002"],
    ["精技", "115.08.03", "EF12345678", 3000, "NO-003"]
  ];
  samples.forEach((values, index) => {
    ["B", "C", "D", "E", "F"].forEach((column, columnIndex) => {
      source.getCell(`${column}${index + 2}`).value = values[columnIndex];
    });
  });

  const details = [2, 3, 4];
  target.getCell("D3").value = source.getCell("A2").value;
  target.unMergeCells("C7:C12");
  for (let row = 7; row <= 12; row += 1) target.getCell(`C${row}`).style = originalVendorStyles[row - 7];
  for (let row = 7; row <= 12; row += 1) {
    for (const column of ["B", "C", "D", "E", "G", "H", "M"]) target.getCell(`${column}${row}`).value = null;
  }
  details.forEach((sourceRow, index) => {
    const row = 7 + index;
    target.getCell(`B${row}`).value = index + 1;
    if (index === 0) target.getCell(`C${row}`).value = source.getCell(`B${sourceRow}`).value;
    target.getCell(`D${row}`).value = `${source.getCell(`C${sourceRow}`).text} ${source.getCell(`D${sourceRow}`).text}`;
    target.getCell(`E${row}`).value = source.getCell(`E${sourceRow}`).value;
    target.getCell(`G${row}`).value = { formula: `E${row}` };
    target.getCell(`H${row}`).value = { formula: `G${row}` };
    target.getCell(`M${row}`).value = source.getCell(`F${sourceRow}`).value;
  });
  target.mergeCells("C7:C9");
  target.getCell("H13").value = { formula: "SUM(H2:H12)" };
  const currentView = templateBook.views && templateBook.views[0] ? templateBook.views[0] : {};
  templateBook.views = [{ ...currentView, activeTab: 0, firstSheet: 0, visibility: "visible" }];

  const output = await templateBook.xlsx.writeBuffer();
  const roundTrip = new ExcelJS.Workbook();
  await roundTrip.xlsx.load(output);
  const result = roundTrip.worksheets[0];

  assert.strictEqual(roundTrip.views[0].activeTab, 0);
  assert.strictEqual(roundTrip.views[0].firstSheet, 0);
  assert.strictEqual(result.getCell("D3").value, expectedDocumentDate);
  assert.deepStrictEqual([result.getCell("B7").value, result.getCell("B8").value, result.getCell("B9").value], [1, 2, 3]);
  assert.strictEqual(result.getCell("C7").value, "精技");
  assert(result.model.merges.includes("C7:C9"));
  assert(!result.model.merges.includes("C7:C12"));
  for (let row = 10; row <= 12; row += 1) {
    assert.deepStrictEqual(result.getCell(`C${row}`).fill, originalVendorStyles[row - 7].fill);
    assert.deepStrictEqual(result.getCell(`C${row}`).border, originalVendorStyles[row - 7].border);
  }
  assert.strictEqual(result.getCell("D7").value, "115.08.01 AB12345678");
  assert.strictEqual(result.getCell("E8").value, 2000);
  assert.strictEqual(result.getCell("M9").value, "NO-003");
  assert.strictEqual(result.getCell("G9").value.formula, "E9");
  assert.strictEqual(result.getCell("H9").value.formula, "G9");
  assert.strictEqual(result.getCell("H13").value.formula, "SUM(H2:H12)");
  assert.deepStrictEqual(result.getCell("B5").style, headerStyle);

  console.log("PASS: 3 expense rows; merge, sequence, mappings, formulas, and styles survived XLSX round-trip");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
