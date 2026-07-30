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

export default async function handler(req, res) {
  res.setHeader("Cache-Control", "no-store");
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return send(res, 405, { error: "Method not allowed" });
  }
  if (!process.env.OPENAI_API_KEY) return send(res, 503, { error: "AI_NOT_CONFIGURED" });

  let body;
  try {
    body = typeof req.body === "string" ? JSON.parse(req.body) : req.body || {};
  } catch {
    return send(res, 400, { error: "Invalid JSON" });
  }
  const sheets = Array.isArray(body.template?.sheets) ? body.template.sheets.slice(0, 8) : [];
  const images = Array.isArray(body.templateImages) ? body.templateImages.slice(0, 4) : [];
  if (!sheets.length) return send(res, 400, { error: "Invalid template" });
  if (
    images.some(
      (image) =>
        !String(image?.dataUrl || "").startsWith("data:image/jpeg;base64,") ||
        String(image.dataUrl).length > MAX_TEMPLATE_IMAGE_LENGTH,
    )
  ) {
    return send(res, 400, { error: "Invalid template image" });
  }

  const template = {
    sheets: sheets.map((sheet) => ({
      name: String(sheet.name || "").slice(0, 60),
      rowCount: Number(sheet.rowCount) || 1,
      columnCount: Number(sheet.columnCount) || 1,
      cells: Array.isArray(sheet.cells) ? sheet.cells.slice(0, 280) : [],
      merges: Array.isArray(sheet.merges) ? sheet.merges.slice(0, 80) : [],
      columnWidths: Array.isArray(sheet.columnWidths) ? sheet.columnWidths.slice(0, 30) : [],
      rowHeights: Array.isArray(sheet.rowHeights) ? sheet.rowHeights.slice(0, 50) : [],
      dataValidations: Array.isArray(sheet.dataValidations)
        ? sheet.dataValidations.slice(0, 80)
        : [],
    })),
  };
  const content = [
    {
      type: "text",
      text: JSON.stringify({
        task: "识别用户Excel明细模板中真正的明细表表头",
        requirements: [
          "选择实际用于逐条填写发票明细的工作表",
          "识别真正的表头行，不得把报表标题、部门、姓名、制表日期、说明、汇总或签字区域当作表头",
          "columns只返回该表头行中实际用于填写明细的非空列，label必须使用Excel中的原始表头文字",
          "column使用1起始Excel列号，按从左到右顺序返回",
        ],
        template,
      }),
    },
  ];
  for (const image of images) {
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
        max_completion_tokens: 1200,
        messages: [
          {
            role: "system",
            content:
              "你是Excel模板结构分析器。结合单元格结构和视觉预览，准确定位逐行填写数据的明细表表头。表头是描述每条明细记录各列含义的一行；报表标题、基本信息区、说明、合计和签字区都不是表头。只返回Excel中真实存在的表头文字和列号。",
          },
          { role: "user", content },
        ],
        response_format: {
          type: "json_schema",
          json_schema: {
            name: "template_columns",
            strict: true,
            schema: {
              type: "object",
              additionalProperties: false,
              properties: {
                targetSheet: { type: "string" },
                headerRow: { type: "integer", minimum: 1, maximum: 200 },
                columns: {
                  type: "array",
                  minItems: 1,
                  maxItems: 30,
                  items: {
                    type: "object",
                    additionalProperties: false,
                    properties: {
                      column: { type: "integer", minimum: 1, maximum: 80 },
                      label: { type: "string" },
                    },
                    required: ["column", "label"],
                  },
                },
                confidence: { type: "number", minimum: 0, maximum: 1 },
              },
              required: ["targetSheet", "headerRow", "columns", "confidence"],
            },
          },
        },
      }),
      signal: AbortSignal.timeout(25_000),
    });
    if (!response.ok) return send(res, 502, { error: "AI_UPSTREAM_ERROR" });
    const result = parseContent(await response.json());
    const sheet = template.sheets.find((item) => item.name === result.targetSheet);
    const seen = new Set();
    const columns = (result.columns || [])
      .map((item) => {
        const sourceCell = sheet?.cells.find(
          (cell) => cell.r === result.headerRow && cell.c === item.column,
        );
        return {
          column: item.column,
          label: String(sourceCell?.v || "").trim(),
        };
      })
      .filter(
        (item) =>
          Number.isInteger(item.column) &&
          item.column >= 1 &&
          item.column <= (sheet?.columnCount || 80) &&
          item.label &&
          !seen.has(item.column) &&
          seen.add(item.column),
      );
    if (!sheet || !Number.isInteger(result.headerRow) || !columns.length) {
      return send(res, 502, { error: "AI_INVALID_RESPONSE" });
    }
    return send(res, 200, { data: { ...result, columns } });
  } catch {
    return send(res, 502, { error: "AI_REQUEST_FAILED" });
  }
}
