const MAX_TEXT_LENGTH = 5000;
const MAX_IMAGE_LENGTH = 2_500_000;
const ALLOWED_TYPES = ["ordinary", "special", "train", "highSpeed"];

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
  const text = String(body.text || "").slice(0, MAX_TEXT_LENGTH);
  const fileName = String(body.fileName || "").slice(0, 160);
  const imageDataUrl = String(body.imageDataUrl || "");
  if (
    imageDataUrl &&
    (!imageDataUrl.startsWith("data:image/jpeg;base64,") ||
      imageDataUrl.length > MAX_IMAGE_LENGTH)
  ) {
    return send(res, 400, { error: "Invalid image" });
  }
  if (!text && !imageDataUrl) {
    return send(res, 400, { error: "No invoice content" });
  }

  const content = [
    {
      type: "text",
      text: `识别中国报销票据。文件名：${fileName}\n只提取可确认的信息，不要猜测。\n${text}`,
    },
  ];
  if (imageDataUrl) {
    content.push({
      type: "image_url",
      image_url: { url: imageDataUrl, detail: "low" },
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
        model: process.env.OPENAI_INVOICE_MODEL || "gpt-5.6-luna",
        store: false,
        reasoning_effort: "none",
        max_completion_tokens: 300,
        messages: [
          {
            role: "system",
            content:
              "你是票据字段提取器。类型只允许 ordinary、special、train、highSpeed。金额返回数字，无法确认返回 null。",
          },
          { role: "user", content },
        ],
        response_format: {
          type: "json_schema",
          json_schema: {
            name: "invoice_analysis",
            strict: true,
            schema: {
              type: "object",
              additionalProperties: false,
              properties: {
                type: { type: "string", enum: ALLOWED_TYPES },
                confidence: { type: "number", minimum: 0, maximum: 1 },
                totalAmount: { type: ["number", "null"] },
                invoiceNumber: { type: "string" },
                issueDate: { type: "string" },
                buyerName: { type: "string" },
                sellerName: { type: "string" },
                trainNumber: { type: "string" },
                route: { type: "string" },
                summary: { type: "string" },
              },
              required: [
                "type",
                "confidence",
                "totalAmount",
                "invoiceNumber",
                "issueDate",
                "buyerName",
                "sellerName",
                "trainNumber",
                "route",
                "summary",
              ],
            },
          },
        },
      }),
      signal: AbortSignal.timeout(25_000),
    });

    if (!response.ok) {
      return send(res, 502, { error: "AI_UPSTREAM_ERROR" });
    }
    const result = await response.json();
    const parsed = JSON.parse(result.choices?.[0]?.message?.content || "{}");
    if (!ALLOWED_TYPES.includes(parsed.type)) {
      return send(res, 502, { error: "AI_INVALID_RESPONSE" });
    }
    return send(res, 200, { data: parsed });
  } catch {
    return send(res, 502, { error: "AI_REQUEST_FAILED" });
  }
}
