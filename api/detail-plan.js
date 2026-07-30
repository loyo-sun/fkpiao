const ALLOWED_FIELDS = [
  "fileName",
  "invoiceType",
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

function validatePlan(plan, sheetNames) {
  if (!sheetNames.includes(plan?.targetSheet)) return false;
  if (!Number.isInteger(plan.headerRow) || !Number.isInteger(plan.firstDataRow)) return false;
  if (plan.firstDataRow <= plan.headerRow) return false;
  const mappings = Array.isArray(plan.mappings) ? plan.mappings : [];
  const valid = mappings.filter(
    (item) =>
      ALLOWED_FIELDS.includes(item?.field) &&
      Number.isInteger(item?.column) &&
      item.column >= 1 &&
      item.column <= 80,
  );
  return new Set(valid.map((item) => item.field)).size >= 2;
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
  const invoices = Array.isArray(body.invoices) ? body.invoices.slice(0, 3) : [];
  if (!Array.isArray(template?.sheets) || !template.sheets.length || !invoices.length) {
    return send(res, 400, { error: "Invalid template snapshot" });
  }

  const compactTemplate = {
    sheets: template.sheets.slice(0, 8).map((sheet) => ({
      name: String(sheet.name || "").slice(0, 60),
      rowCount: Number(sheet.rowCount) || 1,
      columnCount: Number(sheet.columnCount) || 1,
      cells: Array.isArray(sheet.cells) ? sheet.cells.slice(0, 280) : [],
      merges: Array.isArray(sheet.merges) ? sheet.merges.slice(0, 80) : [],
    })),
  };
  const sheetNames = compactTemplate.sheets.map((sheet) => sheet.name);

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
        max_completion_tokens: 700,
        messages: [
          {
            role: "system",
            content:
              "你是Excel报销明细表模板解析器。仅判断填表位置，不改模板。根据非空单元格识别表头行、首个数据行、可复制样式的数据行，并将允许字段映射到1起始列号。只映射模板实际存在且语义明确的列；优先保留原模板结构。",
          },
          {
            role: "user",
            content: JSON.stringify({
              task: "生成发票明细表填写计划",
              allowedFields: ALLOWED_FIELDS,
              template: compactTemplate,
              invoiceSamples: invoices,
            }),
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
                confidence: { type: "number", minimum: 0, maximum: 1 },
              },
              required: [
                "targetSheet",
                "headerRow",
                "firstDataRow",
                "styleSourceRow",
                "mappings",
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
    if (!validatePlan(plan, sheetNames)) {
      return send(res, 502, { error: "AI_INVALID_RESPONSE" });
    }
    return send(res, 200, { data: plan });
  } catch {
    return send(res, 502, { error: "AI_REQUEST_FAILED" });
  }
}
