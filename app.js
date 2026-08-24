(() => {
  "use strict";

  const sourceInput = document.querySelector("#source-file");
  const templateInput = document.querySelector("#template-file");
  const convertButton = document.querySelector("#convert");
  const status = document.querySelector("#status");
  const caseSourceInput = document.querySelector("#case-source-file");
  const caseTemplateInput = document.querySelector("#case-template-file");
  const caseConvertButton = document.querySelector("#case-convert");
  const caseStatus = document.querySelector("#case-status");
  const legacySourceInput = document.querySelector("#legacy-source-file");
  const legacyTemplateInput = document.querySelector("#legacy-template-file");
  const legacyConvertButton = document.querySelector("#legacy-convert");
  const legacyStatus = document.querySelector("#legacy-status");
  const expenseSourceInput = document.querySelector("#expense-source-file");
  const expenseTemplateInput = document.querySelector("#expense-template-file");
  const expenseConvertButton = document.querySelector("#expense-convert");
  const expenseStatus = document.querySelector("#expense-status");

  const mappings = [
    ["C", "C"], ["V", "M"], ["Z", "V"], ["AD", "S"]
  ];

  function setStatus(element, message, type = "") {
    element.textContent = message;
    element.className = type;
  }

  function updateFile(input, nameId, zoneId, inputs, button, statusElement) {
    const file = input.files[0];
    document.querySelector(nameId).textContent = file ? file.name : "尚未選擇";
    document.querySelector(zoneId).classList.toggle("selected", Boolean(file));
    button.disabled = !inputs.every((item) => item.files[0]);
    setStatus(statusElement, button.disabled ? "請先選擇兩個 Excel 檔案" : "檔案已就緒，可以開始轉錄");
  }

  sourceInput.addEventListener("change", () => updateFile(sourceInput, "#source-name", "#source-zone", [sourceInput, templateInput], convertButton, status));
  templateInput.addEventListener("change", () => updateFile(templateInput, "#template-name", "#template-zone", [sourceInput, templateInput], convertButton, status));
  caseSourceInput.addEventListener("change", () => updateFile(caseSourceInput, "#case-source-name", "#case-source-zone", [caseSourceInput, caseTemplateInput], caseConvertButton, caseStatus));
  caseTemplateInput.addEventListener("change", () => updateFile(caseTemplateInput, "#case-template-name", "#case-template-zone", [caseSourceInput, caseTemplateInput], caseConvertButton, caseStatus));
  legacySourceInput.addEventListener("change", () => updateFile(legacySourceInput, "#legacy-source-name", "#legacy-source-zone", [legacySourceInput, legacyTemplateInput], legacyConvertButton, legacyStatus));
  legacyTemplateInput.addEventListener("change", () => updateFile(legacyTemplateInput, "#legacy-template-name", "#legacy-template-zone", [legacySourceInput, legacyTemplateInput], legacyConvertButton, legacyStatus));
  expenseSourceInput.addEventListener("change", () => updateFile(expenseSourceInput, "#expense-source-name", "#expense-source-zone", [expenseSourceInput, expenseTemplateInput], expenseConvertButton, expenseStatus));
  expenseTemplateInput.addEventListener("change", () => updateFile(expenseTemplateInput, "#expense-template-name", "#expense-template-zone", [expenseSourceInput, expenseTemplateInput], expenseConvertButton, expenseStatus));

  const tabs = [...document.querySelectorAll("[role='tab']")];
  function activateTab(tab) {
    tabs.forEach((item) => {
      const selected = item === tab;
      item.classList.toggle("active", selected);
      item.setAttribute("aria-selected", String(selected));
      item.tabIndex = selected ? 0 : -1;
      document.querySelector(`#${item.dataset.panel}`).hidden = !selected;
    });
  }
  tabs.forEach((tab, index) => {
    tab.addEventListener("click", () => activateTab(tab));
    tab.addEventListener("keydown", (event) => {
      if (!["ArrowLeft", "ArrowRight"].includes(event.key)) return;
      event.preventDefault();
      const offset = event.key === "ArrowRight" ? 1 : -1;
      const next = tabs[(index + offset + tabs.length) % tabs.length];
      activateTab(next);
      next.focus();
    });
  });

  document.querySelectorAll(".dropzone").forEach((zone) => {
    ["dragenter", "dragover"].forEach((event) => zone.addEventListener(event, () => zone.classList.add("dragging")));
    ["dragleave", "drop"].forEach((event) => zone.addEventListener(event, () => zone.classList.remove("dragging")));
  });

  function readWorkbook(file) {
    return file.arrayBuffer().then(async (buffer) => {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);
      return workbook;
    });
  }

  function readLegacyWorkbook(file) {
    return file.arrayBuffer().then((buffer) => XLSX.read(buffer, { type: "array", cellDates: true }));
  }

  function getCellValue(sheet, column, row) {
    const value = sheet.getCell(`${column}${row}`).value;
    if (value && typeof value === "object" && "result" in value) return value.result;
    if (value && typeof value === "object" && ("formula" in value || "ref" in value)) return null;
    if (value && typeof value === "object" && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text).join("");
    }
    return value;
  }

  function getCellText(sheet, column, row) {
    const cell = sheet.getCell(`${column}${row}`);
    return String(cell.text ?? getCellValue(sheet, column, row) ?? "").trim();
  }

  function writeValue(sheet, address, value) {
    sheet.getCell(address).value = value === undefined ? null : value;
  }

  function getLegacyValue(sheet, address) {
    const cell = sheet[address];
    return cell ? cell.v : null;
  }

  function parseSalaryPeriod(value, sourceRow) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return { year: value.getFullYear() - 1911, month: value.getMonth() + 1 };
    }
    const normalized = String(value ?? "").trim().replace(/\s+/g, "");
    const match = normalized.match(/^(\d{4})[年\/-](\d{1,2})(?:月(?:份)?)?$/);
    if (!match) throw new Error(`資料第 ${sourceRow} 列的薪資月份「${normalized || "空白"}」無法辨識`);
    const westernYear = Number(match[1]);
    const month = Number(match[2]);
    if (westernYear <= 1911 || month < 1 || month > 12) {
      throw new Error(`資料第 ${sourceRow} 列的薪資月份「${normalized}」不在有效範圍`);
    }
    return { year: westernYear - 1911, month };
  }

  function numericValue(value, column, sourceRow) {
    if (value === null || value === undefined || value === "") return 0;
    if (typeof value === "number" && Number.isFinite(value)) return value;
    const parsed = Number(String(value).replaceAll(",", "").trim());
    if (!Number.isFinite(parsed)) throw new Error(`資料第 ${sourceRow} 列的 ${column} 欄不是有效數字`);
    return parsed;
  }

  function isDataRow(sheet, row) {
    return ["A", "B", "C", "S", "T", "V", "W", "Z", "AD", "AE", "AN"].some((column) => {
      const value = getCellValue(sheet, column, row);
      return value !== null && value !== undefined && value !== "";
    });
  }

  function downloadWorkbook(buffer, filename) {
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function fileTimestamp(date = new Date()) {
    const pad = (value) => String(value).padStart(2, "0");
    return `${date.getFullYear()}${pad(date.getMonth() + 1)}${pad(date.getDate())}_${pad(date.getHours())}${pad(date.getMinutes())}`;
  }

  function normalizeWorkbookView(workbook) {
    const currentView = workbook.views && workbook.views[0] ? workbook.views[0] : {};
    workbook.views = [{ ...currentView, activeTab: 0, firstSheet: 0, visibility: "visible" }];
    if (workbook.worksheets[0]) workbook.worksheets[0].state = "visible";
  }

  async function convertSalary() {
    if (typeof ExcelJS === "undefined") {
      setStatus(status, "Excel 元件未載入，請確認 vendor 資料夾完整", "error");
      return;
    }
    convertButton.disabled = true;
    setStatus(status, "正在讀取並核對資料…");
    try {
      const [sourceBook, templateBook] = await Promise.all([
        readWorkbook(sourceInput.files[0]), readWorkbook(templateInput.files[0])
      ]);
      const sourceSheet = sourceBook.worksheets[0];
      const templateSheet = templateBook.worksheets[0];
      if (!sourceSheet || !templateSheet) throw new Error("Excel 檔案中找不到工作表");

      let targetRow = 2;
      let count = 0;
      for (let sourceRow = 4; sourceRow <= sourceSheet.rowCount; sourceRow += 1) {
        if (!isDataRow(sourceSheet, sourceRow)) continue;
        const period = parseSalaryPeriod(getCellValue(sourceSheet, "B", sourceRow), sourceRow);
        mappings.forEach(([from, to]) => writeValue(templateSheet, `${to}${targetRow}`, getCellValue(sourceSheet, from, sourceRow)));
        writeValue(templateSheet, `L${targetRow}`,
          numericValue(getCellValue(sourceSheet, "S", sourceRow), "S", sourceRow)
          + numericValue(getCellValue(sourceSheet, "T", sourceRow), "T", sourceRow)
          - numericValue(getCellValue(sourceSheet, "AE", sourceRow), "AE", sourceRow));
        writeValue(templateSheet, `N${targetRow}`,
          numericValue(getCellValue(sourceSheet, "W", sourceRow), "W", sourceRow)
          + numericValue(getCellValue(sourceSheet, "AN", sourceRow), "AN", sourceRow));
        writeValue(templateSheet, `J${targetRow}`, period.year);
        writeValue(templateSheet, `K${targetRow}`, period.month);
        targetRow += 1;
        count += 1;
      }
      if (!count) throw new Error("資料檔第 4 列之後沒有可轉錄的資料");

      const stamp = fileTimestamp();
      normalizeWorkbookView(templateBook);
      const output = await templateBook.xlsx.writeBuffer();
      downloadWorkbook(output, `轉錄完成_${stamp}.xlsx`);
      setStatus(status, `完成：已轉錄 ${count} 筆資料並開始下載`, "success");
    } catch (error) {
      console.error(error);
      setStatus(status, error instanceof Error ? error.message : "轉錄失敗，請檢查檔案格式", "error");
    } finally {
      convertButton.disabled = !(sourceInput.files[0] && templateInput.files[0]);
    }
  }

  function hasValue(value) {
    return value !== null && value !== undefined && String(value).trim() !== "";
  }

  async function convertCase() {
    if (typeof ExcelJS === "undefined") {
      setStatus(caseStatus, "Excel 元件未載入，請確認 vendor 資料夾完整", "error");
      return;
    }
    caseConvertButton.disabled = true;
    setStatus(caseStatus, "正在讀取案件與進貨廠商資料…");
    try {
      const [sourceBook, templateBook] = await Promise.all([
        readWorkbook(caseSourceInput.files[0]), readWorkbook(caseTemplateInput.files[0])
      ]);
      const sourceSheet = sourceBook.worksheets[0];
      const templateSheet = templateBook.worksheets[0];
      if (!sourceSheet || !templateSheet) throw new Error("Excel 檔案中找不到工作表");

      const commonMappings = [
        ["C2", "B"], ["H10", "C"], ["H3", "D"], ["D3", "E"],
        ["L4", "F"]
      ];
      let targetRow = 2;
      for (const [sourceAddress, targetColumn] of commonMappings) {
        writeValue(templateSheet, `${targetColumn}${targetRow}`, getCellValue(sourceSheet, sourceAddress.match(/[A-Z]+/)[0], Number(sourceAddress.match(/\d+/)[0])));
      }
      writeValue(templateSheet, `G${targetRow}`, getCellValue(sourceSheet, "C", 17));
      writeValue(templateSheet, `H${targetRow}`, "冠新");
      writeValue(templateSheet, `I${targetRow}`, "004銷貨收入");
      writeValue(templateSheet, `K${targetRow}`, "銀行存款");
      writeValue(templateSheet, `L${targetRow}`, "應收帳款");
      writeValue(templateSheet, `T${targetRow}`, getCellValue(sourceSheet, "H", 5));
      writeValue(templateSheet, `U${targetRow}`, null);
      targetRow += 1;
      let vendorCount = 0;
      for (let sourceRow = 28; sourceRow <= sourceSheet.rowCount; sourceRow += 2) {
        const vendor = getCellValue(sourceSheet, "B", sourceRow) ?? getCellValue(sourceSheet, "M", sourceRow);
        if (!hasValue(vendor)) continue;
        for (const [sourceAddress, targetColumn] of commonMappings) {
          writeValue(templateSheet, `${targetColumn}${targetRow}`, getCellValue(sourceSheet, sourceAddress.match(/[A-Z]+/)[0], Number(sourceAddress.match(/\d+/)[0])));
        }
        writeValue(templateSheet, `G${targetRow}`, null);
        writeValue(templateSheet, `H${targetRow}`, vendor);
        writeValue(templateSheet, `I${targetRow}`, "004銷貨成本");
        writeValue(templateSheet, `K${targetRow}`, "應付帳款");
        writeValue(templateSheet, `L${targetRow}`, "銀行存款");
        writeValue(templateSheet, `P${targetRow}`, getCellValue(sourceSheet, "C", sourceRow));
        writeValue(templateSheet, `T${targetRow}`, null);
        writeValue(templateSheet, `U${targetRow}`, getCellValue(sourceSheet, "D", sourceRow));
        targetRow += 1;
        vendorCount += 1;
      }
      if (!vendorCount) throw new Error("在來源第 28 列之後找不到有效的進貨廠商");

      const stamp = fileTimestamp();
      normalizeWorkbookView(templateBook);
      const output = await templateBook.xlsx.writeBuffer();
      downloadWorkbook(output, `案件轉錄完成_${stamp}.xlsx`);
      setStatus(caseStatus, `完成：已產生 1 筆冠新資料與 ${vendorCount} 筆廠商資料`, "success");
    } catch (error) {
      console.error(error);
      setStatus(caseStatus, error instanceof Error ? error.message : "案件轉錄失敗，請檢查檔案格式", "error");
    } finally {
      caseConvertButton.disabled = !(caseSourceInput.files[0] && caseTemplateInput.files[0]);
    }
  }

  async function convertLegacyCase() {
    if (typeof XLSX === "undefined" || typeof ExcelJS === "undefined") {
      setStatus(legacyStatus, "Excel 元件未載入，請確認 vendor 資料夾完整", "error");
      return;
    }
    legacyConvertButton.disabled = true;
    setStatus(legacyStatus, "正在讀取舊版案件與進貨廠商資料…");
    try {
      const [sourceBook, templateBook] = await Promise.all([
        readLegacyWorkbook(legacySourceInput.files[0]), readWorkbook(legacyTemplateInput.files[0])
      ]);
      const sourceSheet = sourceBook.Sheets[sourceBook.SheetNames[0]];
      const templateSheet = templateBook.worksheets[0];
      if (!sourceSheet || !templateSheet) throw new Error("Excel 檔案中找不到工作表");

      const warnings = [];
      const projectName = String(getLegacyValue(sourceSheet, "C5") ?? "");
      const orderCandidate = projectName.split("_")[2] || "";
      const orderNumber = /^\d{10}$/.test(orderCandidate) ? orderCandidate : null;
      if (!orderNumber) warnings.push(`C5：「${projectName || "空白"}」`);

      const commonMappings = [
        ["C2", "B"], ["H10", "C"], ["H3", "D"], ["C5", "E"],
        ["C4", "F"]
      ];
      const range = XLSX.utils.decode_range(sourceSheet["!ref"] || "A1");
      let targetRow = 2;
      for (const [sourceAddress, targetColumn] of commonMappings) {
        writeValue(templateSheet, `${targetColumn}${targetRow}`, getLegacyValue(sourceSheet, sourceAddress));
      }
      writeValue(templateSheet, `G${targetRow}`, orderNumber);
      writeValue(templateSheet, `H${targetRow}`, "冠新");
      writeValue(templateSheet, `I${targetRow}`, "004銷貨收入");
      writeValue(templateSheet, `K${targetRow}`, "銀行存款");
      writeValue(templateSheet, `L${targetRow}`, "應收帳款");
      writeValue(templateSheet, `T${targetRow}`, getLegacyValue(sourceSheet, "H5"));
      writeValue(templateSheet, `U${targetRow}`, null);
      targetRow += 1;
      let vendorCount = 0;
      for (let sourceRow = 24; sourceRow <= range.e.r + 1; sourceRow += 1) {
        const vendor = getLegacyValue(sourceSheet, `B${sourceRow}`);
        if (!hasValue(vendor)) continue;
        for (const [sourceAddress, targetColumn] of commonMappings) {
          writeValue(templateSheet, `${targetColumn}${targetRow}`, getLegacyValue(sourceSheet, sourceAddress));
        }
        writeValue(templateSheet, `G${targetRow}`, null);
        writeValue(templateSheet, `H${targetRow}`, vendor);
        writeValue(templateSheet, `I${targetRow}`, "004銷貨成本");
        writeValue(templateSheet, `K${targetRow}`, "應付帳款");
        writeValue(templateSheet, `L${targetRow}`, "銀行存款");
        const invoiceSource = String(getLegacyValue(sourceSheet, `F${sourceRow}`) ?? "");
        const invoiceMatch = invoiceSource.match(/\b[A-Za-z]{2}\d{8}\b/);
        if (invoiceMatch) writeValue(templateSheet, `P${targetRow}`, invoiceMatch[0]);
        else warnings.push(`F${sourceRow}：「${invoiceSource || "空白"}」`);
        writeValue(templateSheet, `T${targetRow}`, null);
        writeValue(templateSheet, `U${targetRow}`, getLegacyValue(sourceSheet, `D${sourceRow}`));
        targetRow += 1;
        vendorCount += 1;
      }
      if (!vendorCount) throw new Error("在來源第 24 列之後找不到有效的進貨廠商");

      normalizeWorkbookView(templateBook);
      const output = await templateBook.xlsx.writeBuffer();
      downloadWorkbook(output, `舊版案件轉錄完成_${fileTimestamp()}.xlsx`);
      const warningText = warnings.length ? `\n已跳過不符規則的欄位和值：${warnings.join("、")}` : "";
      setStatus(legacyStatus, `完成：已產生 1 筆冠新資料與 ${vendorCount} 筆廠商資料${warningText}`, "success");
    } catch (error) {
      console.error(error);
      setStatus(legacyStatus, error instanceof Error ? error.message : "舊版案件轉錄失敗，請檢查檔案格式", "error");
    } finally {
      legacyConvertButton.disabled = !(legacySourceInput.files[0] && legacyTemplateInput.files[0]);
    }
  }

  async function convertExpense() {
    if (typeof ExcelJS === "undefined") {
      setStatus(expenseStatus, "Excel 元件未載入，請確認 vendor 資料夾完整", "error");
      return;
    }
    expenseConvertButton.disabled = true;
    setStatus(expenseStatus, "正在讀取支出明細並建立申請單…");
    try {
      const [sourceBook, templateBook] = await Promise.all([
        readWorkbook(expenseSourceInput.files[0]), readWorkbook(expenseTemplateInput.files[0])
      ]);
      const sourceSheet = sourceBook.worksheets[0];
      const templateSheet = templateBook.worksheets[0];
      if (!sourceSheet || !templateSheet) throw new Error("Excel 檔案中找不到工作表");

      const details = [];
      for (let sourceRow = 2; sourceRow <= sourceSheet.rowCount; sourceRow += 1) {
        const hasDetail = ["B", "C", "D", "E", "F"].some((column) => hasValue(getCellValue(sourceSheet, column, sourceRow)));
        if (hasDetail) details.push(sourceRow);
      }
      if (details.length > 6) throw new Error(`範本最多可容納 6 筆明細，目前共有 ${details.length} 筆`);

      writeValue(templateSheet, "D3", getCellValue(sourceSheet, "A", 2));
      const vendorCellStyles = [];
      for (let row = 7; row <= 12; row += 1) {
        vendorCellStyles.push(JSON.parse(JSON.stringify(templateSheet.getCell(`C${row}`).style || {})));
      }
      try { templateSheet.unMergeCells("C7:C12"); } catch (_) { /* 範本若未合併則略過 */ }
      for (let row = 7; row <= 12; row += 1) {
        templateSheet.getCell(`C${row}`).style = vendorCellStyles[row - 7];
      }
      for (let targetRow = 7; targetRow <= 12; targetRow += 1) {
        for (const column of ["B", "C", "D", "E", "G", "H", "M"]) writeValue(templateSheet, `${column}${targetRow}`, null);
      }

      details.forEach((sourceRow, index) => {
        const targetRow = 7 + index;
        const invoiceDate = getCellText(sourceSheet, "C", sourceRow);
        const invoiceNumber = getCellText(sourceSheet, "D", sourceRow);
        writeValue(templateSheet, `B${targetRow}`, index + 1);
        if (index === 0) writeValue(templateSheet, `C${targetRow}`, getCellValue(sourceSheet, "B", sourceRow));
        writeValue(templateSheet, `D${targetRow}`, [invoiceDate, invoiceNumber].filter(Boolean).join(" "));
        writeValue(templateSheet, `E${targetRow}`, getCellValue(sourceSheet, "E", sourceRow));
        writeValue(templateSheet, `G${targetRow}`, { formula: `E${targetRow}` });
        writeValue(templateSheet, `H${targetRow}`, { formula: `G${targetRow}` });
        writeValue(templateSheet, `M${targetRow}`, getCellValue(sourceSheet, "F", sourceRow));
      });
      if (details.length > 1) templateSheet.mergeCells(`C7:C${6 + details.length}`);
      writeValue(templateSheet, "H13", { formula: "SUM(H2:H12)" });
      templateBook.calcProperties.fullCalcOnLoad = true;
      templateBook.calcProperties.forceFullCalc = true;

      normalizeWorkbookView(templateBook);
      const output = await templateBook.xlsx.writeBuffer();
      downloadWorkbook(output, `支出申請單_${fileTimestamp()}.xlsx`);
      setStatus(expenseStatus, `完成：已填入製單日期與 ${details.length} 筆支出明細`, "success");
    } catch (error) {
      console.error(error);
      setStatus(expenseStatus, error instanceof Error ? error.message : "支出單轉錄失敗，請檢查檔案格式", "error");
    } finally {
      expenseConvertButton.disabled = !(expenseSourceInput.files[0] && expenseTemplateInput.files[0]);
    }
  }

  convertButton.addEventListener("click", convertSalary);
  caseConvertButton.addEventListener("click", convertCase);
  legacyConvertButton.addEventListener("click", convertLegacyCase);
  expenseConvertButton.addEventListener("click", convertExpense);
})();
