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

export async function fillDetailWorkbook(templateBytes, invoices) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBytes);

  let selected = null;
  workbook.worksheets.forEach((worksheet) => {
    const candidate = findHeaderRow(worksheet);
    if (!selected || candidate.score > selected.header.score) {
      selected = { worksheet, header: candidate };
    }
  });
  if (!selected || selected.header.score < 2) {
    throw new Error("模板中至少需要两个可识别的表头，例如“发票号码、价税合计”");
  }

  const { worksheet, header } = selected;
  const firstDataRow = header.rowNumber + 1;
  const styleRow = worksheet.getRow(firstDataRow);

  invoices.forEach((invoice, index) => {
    const targetRow = worksheet.getRow(firstDataRow + index);
    targetRow.height = styleRow.height;
    header.columns.forEach((field, column) => {
      const targetCell = targetRow.getCell(column);
      cloneStyle(styleRow.getCell(column), targetCell);
      targetCell.value = invoice[field] ?? "";
      if (["totalAmount", "amountWithoutTax", "taxAmount"].includes(field) && invoice[field] != null) {
        targetCell.numFmt = "#,##0.00";
      }
    });
    targetRow.commit();
  });

  workbook.creator = "发克票";
  workbook.modified = new Date();
  const preview = createWorksheetPreviewModel(
    worksheet,
    Math.max(worksheet.rowCount, firstDataRow + invoices.length - 1),
  );
  const bytes = await workbook.xlsx.writeBuffer();
  return { bytes, preview };
}

function colorToCss(color, fallback) {
  if (!color) return fallback;
  if (color.argb) return `#${color.argb.slice(-6)}`;
  return fallback;
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

  for (let row = 1; row <= rowCount; row += 1) {
    for (let column = 1; column <= columnCount; column += 1) {
      const cell = worksheet.getCell(row, column);
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
