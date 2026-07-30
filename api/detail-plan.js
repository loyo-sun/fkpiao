const ALLOWED_FIELDS = [
  "fileName",
  "invoiceType",
  "expenseCategory",
  "invoiceCode",
  "invoiceNumber",
  "issueDate",
  "buyerName",
  "sellerName",
  "summary",
  "amountWithoutTax",
  "taxAmount",
  "totalAmount",
  "trainNumber",
  "departure",
  "arrival",
  "route",
];
const MAX_TEMPLATE_IMAGE_LENGTH = 1_800_000;

function send(res, status, body) {
  res.status(status).json(body);
}

function apiUrl() {
  const base = (process.env.OPENAI_BASE_URL || "https://api.openai-hub.net/v1").replace(
    /\/+$/,
    "",
  );
  return `${base}/chat/completions`;
}

function parseContent(result) {
  const content = result.choices?.[0]?.message?.content || "";
  return JSON.parse(content.replace(/^```json\s*|\s*```$/g, ""));
}

function validatePlan(plan, sheetNames, invoiceCount) {
  if (!sheetNames.includes(plan?.targetSheet)) return false;
  if (!Number.isInteger(plan.headerRow) || !Number.isInteger(plan.firstDataRow)) return false;
  if (plan.firstDataRow <= plan.headerRow) return false;
  if (
    !Number.isInteger(plan.rowsPerPage) ||
    plan.rowsPerPage < 4 ||
    plan.rowsPerPage > 24 ||
    !Number.isInteger(plan.lastColumn) ||
    plan.lastColumn < 1 ||
    plan.lastColumn > 80
  ) {
    return false;
  }
  const mappings = Array.isArray(plan.mappings) ? plan.mappings : [];
  const valid = mappings.filter(
    (item) =>
      ALLOWED_FIELDS.includes(item?.field) &&
      Number.isInteger(item?.column) &&
      item.column >= 1 &&
      item.column <= 80,
  );
  const records = Array.isArray(plan.records) ? plan.records : [];
  const sourceIndexes = new Set(records.map((record) => record?.sourceIndex));
  return (
    new Set(valid.map((item) => item.field)).size >= 2 &&
    Array.from({ length: invoiceCount }, (_, index) => index).every((index) =>
      sourceIndexes.has(index),
    )
  );
}

export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return send(res, 405, { error: "Method not allowed" });
  }
  if (!process.env.OPENAI_API_KEY) {
    return send(res, 503, { error: "AI_NOT_CONFIGURED" });
  }

  let body;
  try {
    body = typeof req.body === "string" ? JSON.parse(req.body) : req.body || {};
  } catch {
    return send(res, 400, { error: "Invalid JSON" });
  }
  const template = body.template;
  const invoices = Array.isArray(body.invoices) ? body.invoices.slice(0, 80) : [];
  const templateImages = Array.isArray(body.templateImages)
    ? body.templateImages.slice(0, 4)
    : [];
  if (!Array.isArray(template?.sheets) || !template.sheets.length || !invoices.length) {
    return send(res, 400, { error: "Invalid template snapshot" });
  }
  if (
    templateImages.some(
      (image) =>
        !String(image?.dataUrl || "").startsWith("data:image/jpeg;base64,") ||
        String(image.dataUrl).length > MAX_TEMPLATE_IMAGE_LENGTH,
    )
  ) {
    return send(res, 400, { error: "Invalid template image" });
  }

  const compactTemplate = {
    sheets: template.sheets.slice(0, 8).map((sheet) => ({
      name: String(sheet.name || "").slice(0, 60),
      rowCount: Number(sheet.rowCount) || 1,
      columnCount: Number(sheet.columnCount) || 1,
      cells: Array.isArray(sheet.cells) ? sheet.cells.slice(0, 280) : [],
      merges: Array.isArray(sheet.merges) ? sheet.merges.slice(0, 80) : [],
      columnWidths: Array.isArray(sheet.columnWidths) ? sheet.columnWidths.slice(0, 30) : [],
      rowHeights: Array.isArray(sheet.rowHeights) ? sheet.rowHeights.slice(0, 50) : [],
      pageSetup: sheet.pageSetup || {},
      dataValidations: Array.isArray(sheet.dataValidations)
        ? sheet.dataValidations.slice(0, 80)
        : [],
    })),
  };
  const sheetNames = compactTemplate.sheets.map((sheet) => sheet.name);
  const content = [
    {
      type: "text",
      text: JSON.stringify({
        task: "分析用户上传的Excel明细模板及发票内容，生成清晰可读的A5填写与分页方案",
        allowedFields: ALLOWED_FIELDS,
        fieldDefinitions: {
          invoiceType: "票据本身的类型，例如普票、专票、火车票、高铁票",
          expenseCategory:
            "模板中的费用类型、报销类别或费用科目；必须根据发票内容选择模板数据验证范围中的值",
        },
        requirements: [
          "严格沿用模板的标题、表头、列顺序、合并关系和视觉风格",
          "选择A5横向打印时仍能清晰阅读的每页数据行数，禁止把全部记录强压成一页",
          "内容无法在A5页面中完整显示时必须换行，所有内容都要完整展示，不得截断或省略",
          "字段只映射到模板实际存在且语义明确的列",
          "模板字段存在数据验证或允许值范围时，输出值必须严格属于该范围，尤其是类型字段",
          "分析全部发票；records必须为每张发票返回一条记录，并完整提供模板所需的标准字段和自定义列内容",
          "模板中无法映射到标准字段的明细列，使用records.cells按列号填写，不允许因为字段不在allowedFields中而遗漏",
          "标题区、汇总区或其他非明细填写区域使用templateCells填写，但不得覆盖模板标题、表头、公式和固定说明文字",
          "确实无法从发票确认的信息填写“未识别”，不得留空，也不得编造金额、号码、日期、单位或行程",
        ],
        template: compactTemplate,
        invoices,
      }),
    },
  ];
  for (const image of templateImages) {
    content.push({
      type: "text",
      text: `工作表“${String(image.sheetName || "").slice(0, 60)}”的视觉预览：`,
    });
    content.push({
      type: "image_url",
      image_url: { url: image.dataUrl, detail: "high" },
    });
  }

  try {
    const response = await fetch(apiUrl(), {
      method: "POST",
      headers: {
        Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model:
          process.env.OPENAI_DETAIL_MODEL ||
          process.env.OPENAI_INVOICE_MODEL ||
          "gpt-5.6-luna",
        store: false,
        reasoning_effort: "none",
        max_completion_tokens: Math.min(16_000, 1800 + invoices.length * 180),
        messages: [
          {
            role: "system",
            content:
              "你是专业的Excel报销明细表设计与内容分析器。你会同时阅读模板的结构化数据和视觉预览，理解标题、表头、数据区、字体、列宽、行高、合并单元格和打印范围。输出填写及A5分页计划，不重新设计用户模板。必须完整填写用户模板：每张发票对应一条records记录，所有需要填写的明细列都必须有内容；标准字段写入values，自定义列写入cells，非明细填写区域写入templateCells。不得遗漏模板列。每页必须清晰可读；宽表优先A5横向，按合理行数分页并重复标题和表头。输出内容无法在A5页面中完整展示时必须换行，确保全部展示。必须区分invoiceType（普票、专票、火车票、高铁票等票据类型）和expenseCategory（交通费、办公费、差旅费等费用类别）。模板中的类型、类别或科目设置了允许值范围时，必须从该范围选择。只能从发票中提取或合理归纳内容，确实无法确认时填写“未识别”，绝不编造金额、号码、日期、单位或行程。",
          },
          {
            role: "user",
            content,
          },
        ],
        response_format: {
          type: "json_schema",
          json_schema: {
            name: "detail_template_plan",
            strict: true,
            schema: {
              type: "object",
              additionalProperties: false,
              properties: {
                targetSheet: { type: "string" },
                headerRow: { type: "integer", minimum: 1, maximum: 200 },
                firstDataRow: { type: "integer", minimum: 2, maximum: 200 },
                styleSourceRow: { type: "integer", minimum: 1, maximum: 200 },
                lastColumn: { type: "integer", minimum: 1, maximum: 80 },
                rowsPerPage: { type: "integer", minimum: 4, maximum: 24 },
                mappings: {
                  type: "array",
                  minItems: 2,
                  maxItems: 15,
                  items: {
                    type: "object",
                    additionalProperties: false,
                    properties: {
                      field: { type: "string", enum: ALLOWED_FIELDS },
                      column: { type: "integer", minimum: 1, maximum: 80 },
                    },
                    required: ["field", "column"],
                  },
                },
                templateCells: {
                  type: "array",
                  maxItems: 100,
                  items: {
                    type: "object",
                    additionalProperties: false,
                    properties: {
                      row: { type: "integer", minimum: 1, maximum: 200 },
                      column: { type: "integer", minimum: 1, maximum: 80 },
                      value: { type: ["string", "number", "null"] },
                    },
                    required: ["row", "column", "value"],
                  },
                },
                records: {
                  type: "array",
                  maxItems: 80,
                  items: {
                    type: "object",
                    additionalProperties: false,
                    properties: {
                      sourceIndex: { type: "integer", minimum: 0, maximum: 79 },
                      values: {
                        type: "array",
                        minItems: 1,
                        maxItems: 15,
                        items: {
                          type: "object",
                          additionalProperties: false,
                          properties: {
                            field: { type: "string", enum: ALLOWED_FIELDS },
                            value: { type: ["string", "number", "null"] },
                          },
                          required: ["field", "value"],
                        },
                      },
                      cells: {
                        type: "array",
                        maxItems: 80,
                        items: {
                          type: "object",
                          additionalProperties: false,
                          properties: {
                            column: { type: "integer", minimum: 1, maximum: 80 },
                            value: { type: ["string", "number", "null"] },
                          },
                          required: ["column", "value"],
                        },
                      },
                    },
                    required: ["sourceIndex", "values", "cells"],
                  },
                },
                confidence: { type: "number", minimum: 0, maximum: 1 },
              },
              required: [
                "targetSheet",
                "headerRow",
                "firstDataRow",
                "styleSourceRow",
                "lastColumn",
                "rowsPerPage",
                "mappings",
                "templateCells",
                "records",
                "confidence",
              ],
            },
          },
        },
      }),
      signal: AbortSignal.timeout(25_000),
    });
    if (!response.ok) return send(res, 502, { error: "AI_UPSTREAM_ERROR" });
    const plan = parseContent(await response.json());
    if (!validatePlan(plan, sheetNames, invoices.length)) {
      return send(res, 502, { error: "AI_INVALID_RESPONSE" });
    }
    return send(res, 200, { data: plan });
  } catch {
    return send(res, 502, { error: "AI_REQUEST_FAILED" });
  }
}
