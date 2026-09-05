const fs = require("fs");
const vm = require("vm");
const assert = require("assert");
const ExcelJS = require("../vendor/exceljs.min.js");

(async () => {
  const source = new ExcelJS.Workbook();
  const sheet = source.addWorksheet("來源");
  const template = new ExcelJS.Workbook();
  const target = template.addWorksheet("案件");
  const samples = [
    ["115.7.29  CN29044516", "CN29044516", "115.7.29"],
    ["115.7.9 CN29044516", "CN29044516", "115.7.9"],
    ["CN29044516 115.07.9", "CN29044516", "115.07.9"],
    ["AB12345678\n115.08.23", "AB12345678", "115.08.23"],
    ["日期：115.09.05 發票：cd68551026", "cd68551026", "115.09.05"],
    ["115.08.24", null, "115.08.24"],
    ["DD68551026", "DD68551026", null],
    ["尚未開票", null, null],
    ["ABC12345678 / 12345678 / A12345678", null, null],
    [null, null, null],
    ["2026.09.05", null, null],
  ];
  samples.forEach(([input], index) => {
    sheet.getCell(`B${28 + index * 2}`).value = `廠商${index}`;
    sheet.getCell(`C${28 + index * 2}`).value = input;
    target.getCell(`O${3 + index}`).value = "舊日期";
    target.getCell(`P${3 + index}`).value = "舊發票";
  });
  const elements = new Map();
  const document = {
    querySelector(selector) {
      if (!elements.has(selector)) elements.set(selector, {
        files: [], addEventListener(event, handler) { this[event] = handler; },
      });
      return elements.get(selector);
    },
    querySelectorAll() { return []; },
    body: { appendChild() {} },
    createElement() { return { click() {}, remove() {} }; },
  };
  document.querySelector("#case-source-file").files = [{ arrayBuffer: () => source.xlsx.writeBuffer() }];
  document.querySelector("#case-template-file").files = [{ arrayBuffer: () => template.xlsx.writeBuffer() }];
  let downloaded;
  vm.runInNewContext(fs.readFileSync(require.resolve("../app.js"), "utf8"), {
    document, ExcelJS, Blob, console, setTimeout: () => {},
    URL: { createObjectURL(blob) { downloaded = blob; return "blob:test"; }, revokeObjectURL() {} },
  });
  await elements.get("#case-convert").click();
  assert(downloaded, elements.get("#case-status").textContent);
  const result = new ExcelJS.Workbook();
  await result.xlsx.load(await downloaded.arrayBuffer());
  samples.forEach(([, invoice, date], index) => {
    assert.strictEqual(result.worksheets[0].getCell(`P${index + 3}`).value, invoice);
    assert.strictEqual(result.worksheets[0].getCell(`O${index + 3}`).value, date);
  });
  console.log(`PASS: actual case conversion separates invoices and dates for ${samples.length} cases through XLSX round-trip`);
})().catch((error) => { console.error(error); process.exitCode = 1; });
