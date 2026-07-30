import ExcelJS from "exceljs";

const FIELD_ALIASES = {
  fileName: ["文件名", "源文件", "发票文件"],
  invoiceType: ["发票类型", "票据类型", "类型"],
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

const normalizeHeader = (value) =>
  String(value ?? "")
    .replace(/\s+/g, "")
    .replace(/[：:（）()【】[\]_-]/g, "")
    .toLowerCase();

const aliasLookup = new Map(
  Object.entries(FIELD_ALIASES).flatMap(([field, aliases]) =>
    aliases.map((alias) => [normalizeHeader(alias), field]),
  ),
);

function findHeaderRow(worksheet) {
  let best = { rowNumber: 0, columns: new Map(), score: 0 };
  const maxRow = Math.min(worksheet.rowCount || 30, 30);
  const maxColumn = Math.min(worksheet.columnCount || 50, 50);

  for (let rowNumber = 1; rowNumber <= maxRow; rowNumber += 1) {
    const columns = new Map();
    const row = worksheet.getRow(rowNumber);
    for (let column = 1; column <= maxColumn; column += 1) {
      const value = row.getCell(column).text;
      const field = aliasLookup.get(normalizeHeader(value));
      if (field) columns.set(column, field);
    }
    if (columns.size > best.score) best = { rowNumber, columns, score: columns.size };
  }
  return best;
}

function cloneStyle(sourceCell, targetCell) {
  if (!sourceCell?.style) return;
  targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
  targetCell.numFmt = sourceCell.numFmt;
  targetCell.alignment = sourceCell.alignment
    ? JSON.parse(JSON.stringify(sourceCell.alignment))
    : undefined;
}

function compactText(value, maxLength = 80) {
  return String(value ?? "").replace(/\s+/g, " ").trim().slice(0, maxLength);
}

export async function inspectDetailTemplate(templateBytes) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBytes);
  return {
    sheets: workbook.worksheets.slice(0, 8).map((worksheet) => {
      const maxRow = Math.min(Math.max(worksheet.rowCount, 1), 50);
      const maxColumn = Math.min(Math.max(worksheet.columnCount, 1), 30);
      const cells = [];
      for (let row = 1; row <= maxRow && cells.length < 280; row += 1) {
        for (let column = 1; column <= maxColumn && cells.length < 280; column += 1) {
          const cell = worksheet.getCell(row, column);
          if (cell.isMerged && cell.master?.address !== cell.address) continue;
          const text = compactText(cell.text);
          if (text) cells.push({ r: row, c: column, v: text });
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
      };
    }),
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
  const mappings = new Map();
  for (const item of plan.mappings || []) {
    const column = Number(item.column);
    if (
      DETAIL_FIELDS.includes(item.field) &&
      Number.isInteger(column) &&
      column >= 1 &&
      column <= 80
    ) {
      mappings.set(column, item.field);
    }
  }
  if (mappings.size < 2) throw new Error("AI 未能识别足够的明细表字段");
  return { worksheet, headerRow, firstDataRow, styleSourceRow, mappings };
}

function columnLetter(column) {
  let result = "";
  for (let value = column; value > 0; value = Math.floor((value - 1) / 26)) {
    result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
  }
  return result;
}

export async function fillDetailWorkbook(templateBytes, invoices, plan) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBytes);
  const { worksheet, firstDataRow, styleSourceRow, mappings } = validatePlan(workbook, plan);
  const styleRow = worksheet.getRow(styleSourceRow);

  invoices.forEach((invoice, index) => {
    const targetRow = worksheet.getRow(firstDataRow + index);
    targetRow.height = styleRow.height;
    mappings.forEach((field, column) => {
      const targetCell = targetRow.getCell(column);
      cloneStyle(styleRow.getCell(column), targetCell);
      targetCell.value = invoice[field] ?? "";
      if (["totalAmount", "amountWithoutTax", "taxAmount"].includes(field) && invoice[field] != null) {
        targetCell.numFmt = "#,##0.00";
      }
    });
    targetRow.commit();
  });

  const populatedLastRow = Math.max(worksheet.rowCount, firstDataRow + invoices.length - 1);
  const mappedLastColumn = Math.max(...mappings.keys(), worksheet.columnCount);
  worksheet.pageSetup = {
    ...worksheet.pageSetup,
    paperSize: 11,
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 1,
    horizontalCentered: true,
    verticalCentered: false,
    printArea: `A1:${columnLetter(mappedLastColumn)}${populatedLastRow}`,
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
  const preview = createWorksheetPreviewModel(worksheet, populatedLastRow);
  const bytes = await workbook.xlsx.writeBuffer();
  return { bytes, preview };
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

function createWorksheetPreviewModel(worksheet, populatedRowCount) {
  const rowCount = Math.max(1, Math.min(populatedRowCount, 60));
  const columnCount = Math.max(1, Math.min(worksheet.columnCount, 24));
  const columnWidths = Array.from({ length: columnCount }, (_, index) => {
    const width = worksheet.getColumn(index + 1).width || 10;
    return Math.max(42, Math.min(180, width * 7));
  });
  const rowHeights = Array.from({ length: rowCount }, (_, index) => {
    const height = worksheet.getRow(index + 1).height || 18;
    return Math.max(22, Math.min(68, height * 1.35));
  });
  const cells = [];
  const merged = mergedCellLayout(worksheet);

  for (let row = 1; row <= rowCount; row += 1) {
    for (let column = 1; column <= columnCount; column += 1) {
      if (merged.covered.has(`${row}:${column}`)) continue;
      const cell = worksheet.getCell(row, column);
      const span = merged.masters.get(`${row}:${column}`);
      const fill =
        cell.fill?.type === "pattern" ? colorToCss(cell.fill.fgColor, "#ffffff") : "#ffffff";
      cells.push({
        row: row - 1,
        column: column - 1,
        text: cell.text || "",
        fill,
        color: colorToCss(cell.font?.color, "#27352e"),
        bold: Boolean(cell.font?.bold),
        fontSize: Math.max(8, Math.min(16, cell.font?.size || 10)),
        horizontal: cell.alignment?.horizontal || "left",
        vertical: cell.alignment?.vertical || "middle",
        wrapText: Boolean(cell.alignment?.wrapText),
        rowSpan: span?.rowSpan || 1,
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
