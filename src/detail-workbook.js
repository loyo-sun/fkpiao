import ExcelJS from "exceljs";

const FIELD_ALIASES = {
  fileName: ["文件名", "源文件", "发票文件"],
  invoiceType: ["发票类型", "票据类型"],
  expenseCategory: ["费用类型", "费用类别", "报销类型", "报销类别", "费用科目", "科目", "类别"],
  invoiceCode: ["发票代码", "票据代码"],
  invoiceNumber: ["发票号码", "发票号", "票据号码", "票号"],
  issueDate: ["开票日期", "发票日期", "日期", "乘车日期"],
  buyerName: ["购买方名称", "购方名称", "报销单位", "购买方"],
  sellerName: ["销售方名称", "销方名称", "商户名称", "销售方"],
  summary: ["项目名称", "商品名称", "费用内容", "摘要", "明细"],
  amountWithoutTax: ["不含税金额", "金额", "税前金额"],
  taxAmount: ["税额", "税金"],
  totalAmount: ["价税合计", "含税金额", "总金额", "报销金额", "票价"],
  trainNumber: ["车次", "列车车次"],
  departure: ["出发站", "始发站", "出发地"],
  arrival: ["到达站", "终到站", "目的地"],
  route: ["行程", "路线", "区间"],
};

export const DETAIL_FIELDS = Object.freeze(Object.keys(FIELD_ALIASES));

function cloneStyle(sourceCell, targetCell) {
  if (!sourceCell?.style) return;
  targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
  targetCell.numFmt = sourceCell.numFmt;
  targetCell.alignment = sourceCell.alignment
    ? JSON.parse(JSON.stringify(sourceCell.alignment))
    : undefined;
  if (sourceCell.dataValidation) {
    targetCell.dataValidation = JSON.parse(JSON.stringify(sourceCell.dataValidation));
  }
}

function compactText(value, maxLength = 80) {
  return String(value ?? "").replace(/\s+/g, " ").trim().slice(0, maxLength);
}

function detectSelectableColumns(worksheets) {
  let best = null;
  for (const worksheet of worksheets) {
    const maxRow = Math.min(Math.max(worksheet.rowCount, 1), 30);
    const maxColumn = Math.min(Math.max(worksheet.columnCount, 1), 30);
    for (let row = 1; row <= maxRow; row += 1) {
      const columns = [];
      let styled = 0;
      for (let column = 1; column <= maxColumn; column += 1) {
        const cell = worksheet.getCell(row, column);
        if (cell.isMerged && cell.master?.address !== cell.address) continue;
        const label = compactText(cell.text, 40);
        if (!label || label.length > 24) continue;
        columns.push({ column, label });
        if (cell.font?.bold || cell.fill?.type === "pattern") styled += 1;
      }
      if (columns.length < 2) continue;
      const score = columns.length * 3 + styled;
      if (!best || score > best.score) {
        best = { sheetName: worksheet.name, headerRow: row, columns, score };
      }
    }
  }
  return best
    ? { sheetName: best.sheetName, headerRow: best.headerRow, columns: best.columns }
    : { sheetName: worksheets[0]?.name || "", headerRow: 1, columns: [] };
}

export async function inspectDetailTemplate(templateBytes) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBytes);
  const worksheets = workbook.worksheets.slice(0, 8);
  return {
    selectableColumns: detectSelectableColumns(worksheets),
    snapshot: {
      sheets: worksheets.map((worksheet) => {
      const maxRow = Math.min(Math.max(worksheet.rowCount, 1), 50);
      const maxColumn = Math.min(Math.max(worksheet.columnCount, 1), 30);
      const cells = [];
      for (let row = 1; row <= maxRow && cells.length < 280; row += 1) {
        for (let column = 1; column <= maxColumn && cells.length < 280; column += 1) {
          const cell = worksheet.getCell(row, column);
          if (cell.isMerged && cell.master?.address !== cell.address) continue;
          const text = compactText(cell.text);
          if (text || cell.formula) {
            cells.push({
              r: row,
              c: column,
              v: text,
              f: compactText(cell.formula, 120),
              s: {
                bold: Boolean(cell.font?.bold),
                size: cell.font?.size || 10,
                fill: colorToCss(
                  cell.fill?.type === "pattern" ? cell.fill.fgColor : null,
                  "",
                ),
                align: cell.alignment?.horizontal || "",
                wrap: Boolean(cell.alignment?.wrapText),
                numFmt: compactText(cell.numFmt, 40),
              },
            });
          }
        }
      }
      return {
        name: compactText(worksheet.name, 60),
        rowCount: maxRow,
        columnCount: maxColumn,
        cells,
        merges: (worksheet.model.merges || []).slice(0, 80),
        columnWidths: Array.from({ length: maxColumn }, (_, index) =>
          Math.round((worksheet.getColumn(index + 1).width || 10) * 10) / 10,
        ),
        rowHeights: Array.from({ length: maxRow }, (_, index) =>
          Math.round((worksheet.getRow(index + 1).height || 18) * 10) / 10,
        ),
        pageSetup: {
          orientation: worksheet.pageSetup?.orientation || "",
          paperSize: worksheet.pageSetup?.paperSize || null,
          printArea: worksheet.pageSetup?.printArea || "",
        },
        dataValidations: Object.entries(worksheet.dataValidations?.model || {})
          .slice(0, 80)
          .map(([range, rule]) => ({
            range,
            type: rule.type || "",
            operator: rule.operator || "",
            formulae: (rule.formulae || []).slice(0, 4),
            allowBlank: Boolean(rule.allowBlank),
          })),
      };
      }),
    },
    previewModels: worksheets.slice(0, 4).map((worksheet) =>
      createWorksheetPreviewModel(
        worksheet,
        Array.from(
          { length: Math.min(Math.max(worksheet.rowCount, 1), 40) },
          (_, index) => index + 1,
        ),
      ),
    ),
  };
}

function validatePlan(workbook, plan) {
  const worksheet = workbook.getWorksheet(plan?.targetSheet);
  if (!worksheet) throw new Error("AI 未能在模板中定位目标工作表");
  const headerRow = Number(plan.headerRow);
  const firstDataRow = Number(plan.firstDataRow);
  const styleSourceRow = Number(plan.styleSourceRow || firstDataRow);
  if (
    !Number.isInteger(headerRow) ||
    !Number.isInteger(firstDataRow) ||
    !Number.isInteger(styleSourceRow) ||
    headerRow < 1 ||
    firstDataRow <= headerRow ||
    firstDataRow > 200 ||
    styleSourceRow > 200
  ) {
    throw new Error("AI 返回的模板行位置无效");
  }
  const requiredColumns = new Set(
    (plan.requiredColumns || [])
      .map((item) => Number(item.column))
      .filter((column) => Number.isInteger(column) && column >= 1 && column <= 80),
  );
  const mappings = new Map();
  for (const item of plan.mappings || []) {
    const column = Number(item.column);
    if (
      DETAIL_FIELDS.includes(item.field) &&
      Number.isInteger(column) &&
      column >= 1 &&
      column <= 80
    ) {
      if (!requiredColumns.size || requiredColumns.has(column)) mappings.set(column, item.field);
    }
  }
  const lastColumn = Math.max(
    ...mappings.keys(),
    Math.min(Number(plan.lastColumn) || worksheet.columnCount, 80),
  );
  return {
    worksheet,
    headerRow,
    firstDataRow,
    styleSourceRow,
    lastColumn,
    mappings,
    requiredColumns,
  };
}

function columnLetter(column) {
  let result = "";
  for (let value = column; value > 0; value = Math.floor((value - 1) / 26)) {
    result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
  }
  return result;
}

function invoiceTypeKind(value) {
  const text = String(value || "");
  if (/高铁|动车|城际|G\d|D\d|C\d/i.test(text)) return "highSpeed";
  if (/火车|铁路|车票/.test(text)) return "train";
  if (/专用|专票/.test(text)) return "special";
  if (/普通|普票/.test(text)) return "ordinary";
  return "";
}

function normalizedValue(value) {
  return String(value || "")
    .replace(/[\s（）()【】[\]、，,./_-]/g, "")
    .toLowerCase();
}

function expenseCategoryHints(invoice) {
  const text = Object.values(invoice || {}).join(" ");
  const hints = [];
  if (/高铁|动车|火车|铁路|车票|机票|出租|网约车|滴滴|公交|地铁|交通/.test(text)) {
    hints.push("交通", "差旅", "车船");
  }
  if (/酒店|宾馆|住宿|房费/.test(text)) hints.push("住宿", "差旅");
  if (/餐饮|餐费|饭店|食品/.test(text)) hints.push("餐饮", "招待");
  if (/办公|文具|耗材|打印|纸张/.test(text)) hints.push("办公");
  if (/服务|咨询|技术/.test(text)) hints.push("服务");
  if (/加油|汽油|柴油|燃油/.test(text)) hints.push("燃油", "车辆", "交通");
  if (/电话|通信|宽带|网络/.test(text)) hints.push("通信");
  return hints;
}

function matchAllowedValue(allowedValues, value, invoice) {
  const normalized = normalizedValue(value);
  const exact = allowedValues.find((allowed) => normalizedValue(allowed) === normalized);
  if (exact) return exact;
  const contained = allowedValues.find((allowed) => {
    const candidate = normalizedValue(allowed);
    return candidate && normalized && (candidate.includes(normalized) || normalized.includes(candidate));
  });
  if (contained) return contained;
  const kind = invoiceTypeKind(value);
  const sameInvoiceKind = kind
    ? allowedValues.find((allowed) => invoiceTypeKind(allowed) === kind)
    : null;
  if (sameInvoiceKind) return sameInvoiceKind;
  const hints = expenseCategoryHints(invoice);
  return allowedValues.find((allowed) =>
    hints.some((hint) => normalizedValue(allowed).includes(normalizedValue(hint))),
  );
}

function allowedCellValues(workbook, cell) {
  const rule = cell.dataValidation;
  const source = rule?.type === "list" ? rule.formulae?.[0] : null;
  if (typeof source !== "string") return [];
  if (source.startsWith('"') && source.endsWith('"')) {
    return source
      .slice(1, -1)
      .split(/[,，]/)
      .map((value) => value.trim())
      .filter(Boolean);
  }
  const match = /^(?:'([^']+)'|([^!]+))!\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i.exec(
    source.replace(/^=/, ""),
  );
  if (!match) return [];
  const worksheet = workbook.getWorksheet(match[1] || match[2]);
  if (!worksheet) return [];
  const startColumn = columnNumber(match[3].toUpperCase());
  const endColumn = columnNumber(match[5].toUpperCase());
  const values = [];
  for (let row = Number(match[4]); row <= Number(match[6]); row += 1) {
    for (let column = startColumn; column <= endColumn; column += 1) {
      const value = worksheet.getCell(row, column).text.trim();
      if (value) values.push(value);
    }
  }
  return values;
}

export async function fillDetailWorkbook(templateBytes, invoices, plan) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBytes);
  const {
    worksheet,
    headerRow,
    firstDataRow,
    styleSourceRow,
    lastColumn,
    mappings,
    requiredColumns,
  } = validatePlan(workbook, plan);
  const styleRow = worksheet.getRow(styleSourceRow);
  const aiRecords = new Map(
    (plan.records || []).map((record) => [
      Number(record.sourceIndex),
      {
        values: Object.fromEntries(
          (record.values || [])
            .filter((item) => DETAIL_FIELDS.includes(item.field))
            .map((item) => [item.field, item.value]),
        ),
        cells: new Map(
          (record.cells || [])
            .filter(
              (item) =>
                Number.isInteger(Number(item.column)) &&
                Number(item.column) >= 1 &&
                Number(item.column) <= lastColumn &&
                (!requiredColumns.size || requiredColumns.has(Number(item.column))),
            )
            .map((item) => [Number(item.column), item.value]),
        ),
      },
    ]),
  );
  for (const item of plan.templateCells || []) {
    const row = Number(item.row);
    const column = Number(item.column);
    if (
      !Number.isInteger(row) ||
      !Number.isInteger(column) ||
      row < 1 ||
      row > 200 ||
      column < 1 ||
      column > lastColumn ||
      row === headerRow
    ) {
      continue;
    }
    const cell = worksheet.getCell(row, column);
    if (!cell.formula && (!cell.text || /待填写|请输入|未填写|____|—{2,}/.test(cell.text))) {
      cell.value = item.value ?? "未识别";
    }
  }

  invoices.forEach((invoice, index) => {
    const aiRecord = aiRecords.get(index) || { values: {}, cells: new Map() };
    const resolvedInvoice = { ...invoice, ...aiRecord.values };
    const targetRow = worksheet.getRow(firstDataRow + index);
    const baseRowHeight = styleRow.height || 18;
    let requiredRowHeight = baseRowHeight;
    for (let column = 1; column <= lastColumn; column += 1) {
      cloneStyle(styleRow.getCell(column), targetRow.getCell(column));
    }
    mappings.forEach((field, column) => {
      const targetCell = targetRow.getCell(column);
      let value = resolvedInvoice[field] ?? "";
      if (field === "invoiceType" || field === "expenseCategory") {
        const allowedValues = allowedCellValues(workbook, targetCell);
        if (allowedValues.length && !allowedValues.includes(String(value))) {
          const matched = matchAllowedValue(allowedValues, value, resolvedInvoice);
          value = matched || (targetCell.dataValidation?.allowBlank ? "" : allowedValues[0]);
        }
      }
      if (value === "" || value == null) value = "未识别";
      targetCell.value = ["invoiceCode", "invoiceNumber", "trainNumber"].includes(field)
        ? { richText: [{ text: String(value) }] }
        : value;
      if (
        ["totalAmount", "amountWithoutTax", "taxAmount"].includes(field) &&
        resolvedInvoice[field] != null
      ) {
        targetCell.numFmt = "#,##0.00";
      } else if (
        ["invoiceCode", "invoiceNumber", "trainNumber", "issueDate"].includes(field)
      ) {
        targetCell.numFmt = "@";
      }
      const displayText = String(value ?? "");
      const columnWidth = worksheet.getColumn(column).width || 10;
      const estimatedLines = Math.max(1, Math.ceil(displayText.length / Math.max(4, columnWidth)));
      if (estimatedLines > 1) {
        targetCell.alignment = {
          ...(targetCell.alignment || {}),
          wrapText: true,
          vertical: "middle",
        };
        requiredRowHeight = Math.max(requiredRowHeight, baseRowHeight * estimatedLines);
      }
    });
    aiRecord.cells.forEach((rawValue, column) => {
      const targetCell = targetRow.getCell(column);
      let value = rawValue ?? "未识别";
      const allowedValues = allowedCellValues(workbook, targetCell);
      if (allowedValues.length && !allowedValues.includes(String(value))) {
        value =
          matchAllowedValue(allowedValues, value, resolvedInvoice) ||
          (targetCell.dataValidation?.allowBlank ? "" : allowedValues[0]);
      }
      targetCell.value = value;
    });
    for (let column = 1; column <= lastColumn; column += 1) {
      if (requiredColumns.size && !requiredColumns.has(column)) continue;
      const header = worksheet.getCell(headerRow, column).text.trim();
      const targetCell = targetRow.getCell(column);
      if (header && !targetCell.formula && !targetCell.text.trim()) {
        targetCell.value = /序号|行号|编号/.test(header)
          ? index + 1
          : /备注/.test(header)
            ? "无"
            : "未识别";
      }
      const displayText = targetCell.text;
      const columnWidth = worksheet.getColumn(column).width || 10;
      const estimatedLines = Math.max(1, Math.ceil(displayText.length / Math.max(4, columnWidth)));
      if (estimatedLines > 1) {
        targetCell.alignment = {
          ...(targetCell.alignment || {}),
          wrapText: true,
          vertical: "middle",
        };
        requiredRowHeight = Math.max(requiredRowHeight, baseRowHeight * estimatedLines);
      }
    }
    targetRow.height = Math.min(150, requiredRowHeight);
    targetRow.commit();
  });

  const populatedLastRow = Math.max(worksheet.rowCount, firstDataRow + invoices.length - 1);
  const mergeColumns = new Set(
    (plan.mergeColumns || [])
      .map(Number)
      .filter(
        (column) =>
          Number.isInteger(column) &&
          column >= 1 &&
          column <= lastColumn &&
          (!requiredColumns.size || requiredColumns.has(column)),
      ),
  );
  for (const column of mergeColumns) {
    let startRow = firstDataRow;
    for (let row = firstDataRow + 1; row <= firstDataRow + invoices.length; row += 1) {
      const previous = worksheet.getCell(row - 1, column).text.trim();
      const current = row < firstDataRow + invoices.length
        ? worksheet.getCell(row, column).text.trim()
        : "";
      if (current === previous && current && current !== "未识别") continue;
      if (row - startRow > 1) {
        const cellsAlreadyMerged = Array.from(
          { length: row - startRow },
          (_, offset) => worksheet.getCell(startRow + offset, column).isMerged,
        ).some(Boolean);
        if (!cellsAlreadyMerged) {
          worksheet.mergeCells(startRow, column, row - 1, column);
          worksheet.getCell(startRow, column).alignment = {
            ...(worksheet.getCell(startRow, column).alignment || {}),
            horizontal: "center",
            vertical: "middle",
            wrapText: true,
          };
        }
      }
      startRow = row;
    }
  }
  worksheet.pageSetup = {
    ...worksheet.pageSetup,
    paperSize: 11,
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    horizontalCentered: true,
    verticalCentered: false,
    printArea: `A1:${columnLetter(lastColumn)}${populatedLastRow}`,
    margins: {
      left: 0.2,
      right: 0.2,
      top: 0.25,
      bottom: 0.25,
      header: 0.1,
      footer: 0.1,
    },
  };
  worksheet.views = [{ state: "normal", showGridLines: false }];
  workbook.creator = "发克票";
  workbook.modified = new Date();
  const previews = [
    createWorksheetPreviewModel(
      worksheet,
      Array.from({ length: populatedLastRow }, (_, index) => index + 1),
      lastColumn,
    ),
  ];
  const bytes = await workbook.xlsx.writeBuffer();
  return { bytes, previews };
}

function colorToCss(color, fallback) {
  if (!color) return fallback;
  if (color.argb) return `#${color.argb.slice(-6)}`;
  return fallback;
}

function columnNumber(letters) {
  return [...letters].reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0);
}

function mergedCellLayout(worksheet) {
  const masters = new Map();
  const covered = new Set();
  for (const range of worksheet.model.merges || []) {
    const match = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/i.exec(range);
    if (!match) continue;
    const startColumn = columnNumber(match[1].toUpperCase());
    const startRow = Number(match[2]);
    const endColumn = columnNumber(match[3].toUpperCase());
    const endRow = Number(match[4]);
    masters.set(`${startRow}:${startColumn}`, {
      rowSpan: endRow - startRow + 1,
      columnSpan: endColumn - startColumn + 1,
    });
    for (let row = startRow; row <= endRow; row += 1) {
      for (let column = startColumn; column <= endColumn; column += 1) {
        if (row !== startRow || column !== startColumn) covered.add(`${row}:${column}`);
      }
    }
  }
  return { masters, covered };
}

function createWorksheetPreviewModel(
  worksheet,
  requestedRows = Array.from(
    { length: Math.min(Math.max(worksheet.rowCount, 1), 60) },
    (_, index) => index + 1,
  ),
  requestedColumnCount = worksheet.columnCount,
) {
  const rowNumbers = requestedRows.slice(0, 200);
  const rowCount = rowNumbers.length;
  const columnCount = Math.max(1, Math.min(requestedColumnCount, 24));
  const columnWidths = Array.from({ length: columnCount }, (_, index) => {
    const width = worksheet.getColumn(index + 1).width || 10;
    return Math.max(42, Math.min(180, width * 7));
  });
  const rowHeights = rowNumbers.map((rowNumber) => {
    const height = worksheet.getRow(rowNumber).height || 18;
    return Math.max(22, Math.min(220, height * 1.35));
  });
  const cells = [];
  const merged = mergedCellLayout(worksheet);
  const visibleRowIndex = new Map(rowNumbers.map((rowNumber, index) => [rowNumber, index]));

  for (const row of rowNumbers) {
    for (let column = 1; column <= columnCount; column += 1) {
      if (merged.covered.has(`${row}:${column}`)) continue;
      const cell = worksheet.getCell(row, column);
      const span = merged.masters.get(`${row}:${column}`);
      const fill =
        cell.fill?.type === "pattern" ? colorToCss(cell.fill.fgColor, "#ffffff") : "#ffffff";
      cells.push({
        row: visibleRowIndex.get(row),
        column: column - 1,
        text: cell.text || "",
        fill,
        color: colorToCss(cell.font?.color, "#27352e"),
        bold: Boolean(cell.font?.bold),
        fontSize: Math.max(8, Math.min(16, cell.font?.size || 10)),
        horizontal: cell.alignment?.horizontal || "left",
        vertical: cell.alignment?.vertical || "middle",
        wrapText: Boolean(cell.alignment?.wrapText),
        rowSpan: Math.min(span?.rowSpan || 1, rowCount - visibleRowIndex.get(row)),
        columnSpan: span?.columnSpan || 1,
      });
    }
  }

  return {
    sheetName: worksheet.name,
    rowCount,
    columnCount,
    columnWidths,
    rowHeights,
    cells,
  };
}

export function detailOutputName(templateName) {
  const base = templateName.replace(/\.xlsx$/i, "") || "发票明细表";
  return `${base}_已填写_${new Date().toISOString().slice(0, 10)}.xlsx`;
}
