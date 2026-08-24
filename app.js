(() => {
  "use strict";

  const sourceInput = document.querySelector("#source-file");
  const templateInput = document.querySelector("#template-file");
  const convertButton = document.querySelector("#convert");
  const status = document.querySelector("#status");

  const mappings = [
    ["C", "C"], ["V", "M"], ["Z", "V"], ["AD", "S"]
  ];

  function setStatus(message, type = "") {
    status.textContent = message;
    status.className = type;
  }

  function updateFile(input, nameId, zoneId) {
    const file = input.files[0];
    document.querySelector(nameId).textContent = file ? file.name : "尚未選擇";
    document.querySelector(zoneId).classList.toggle("selected", Boolean(file));
    convertButton.disabled = !(sourceInput.files[0] && templateInput.files[0]);
    setStatus(convertButton.disabled ? "請先選擇兩個 Excel 檔案" : "檔案已就緒，可以開始轉錄");
  }

  sourceInput.addEventListener("change", () => updateFile(sourceInput, "#source-name", "#source-zone"));
  templateInput.addEventListener("change", () => updateFile(templateInput, "#template-name", "#template-zone"));

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

  function getCellValue(sheet, column, row) {
    const value = sheet.getCell(`${column}${row}`).value;
    if (value && typeof value === "object" && "result" in value) return value.result;
    if (value && typeof value === "object" && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text).join("");
    }
    return value;
  }

  function writeValue(sheet, address, value) {
    sheet.getCell(address).value = value === undefined ? null : value;
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

  async function convert() {
    if (typeof ExcelJS === "undefined") {
      setStatus("Excel 元件未載入，請確認 vendor 資料夾完整", "error");
      return;
    }
    convertButton.disabled = true;
    setStatus("正在讀取並核對資料…");
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

      const stamp = new Date().toISOString().slice(0, 10).replaceAll("-", "");
      const output = await templateBook.xlsx.writeBuffer();
      downloadWorkbook(output, `轉錄完成_${stamp}.xlsx`);
      setStatus(`完成：已轉錄 ${count} 筆資料並開始下載`, "success");
    } catch (error) {
      console.error(error);
      setStatus(error instanceof Error ? error.message : "轉錄失敗，請檢查檔案格式", "error");
    } finally {
      convertButton.disabled = !(sourceInput.files[0] && templateInput.files[0]);
    }
  }

  convertButton.addEventListener("click", convert);
})();
