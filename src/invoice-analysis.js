export const INVOICE_TYPES = {
  ordinary: "普票",
  special: "专票",
  train: "火车票",
  highSpeed: "高铁票",
};

const compact = (value = "") => value.replace(/\s+/g, " ").trim();
const firstMatch = (text, patterns) => {
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match?.[1]) return compact(match[1]);
  }
  return "";
};

function parseAmount(value) {
  if (!value) return null;
  const amount = Number(String(value).replace(/[^\d.-]/g, ""));
  return Number.isFinite(amount) ? amount : null;
}

export function classifyInvoice(text, fileName = "") {
  const source = `${fileName} ${text}`.replace(/\s+/g, "").toUpperCase();
  const railSignal =
    /铁路电子客票|中国铁路|火车票|车次|检票口|始发|终到|[GCD]\d{1,4}次?/.test(source);

  if (railSignal) {
    const highSpeed = /高铁|动车|[GCD]\d{1,4}次?/.test(source);
    return {
      type: highSpeed ? "highSpeed" : "train",
      confidence: highSpeed ? 0.96 : 0.9,
      reason: highSpeed ? "识别到高铁/动车车次" : "识别到铁路客票特征",
    };
  }
  if (/增值税专用发票|专用发票/.test(source)) {
    return { type: "special", confidence: 0.98, reason: "识别到专用发票标题" };
  }
  if (/增值税普通发票|普通发票|电子发票/.test(source)) {
    return { type: "ordinary", confidence: 0.96, reason: "识别到普通/电子发票标题" };
  }
  return {
    type: "ordinary",
    confidence: 0.35,
    reason: "未找到明确类型，暂按普票处理",
  };
}

export function extractInvoiceFields(text, fileName, classification) {
  const source = compact(text);
  const invoiceNumber = firstMatch(source, [
    /发票号码[：:\s]*([0-9]{8,24})/i,
    /票号[：:\s]*([0-9]{8,24})/i,
  ]);
  const invoiceCode = firstMatch(source, [/发票代码[：:\s]*([0-9]{10,14})/i]);
  const issueDate = firstMatch(source, [
    /开票日期[：:\s]*([0-9]{4}[年./-][0-9]{1,2}[月./-][0-9]{1,2}日?)/i,
    /乘车日期[：:\s]*([0-9]{4}[年./-][0-9]{1,2}[月./-][0-9]{1,2}日?)/i,
    /([0-9]{4}[-./年][0-9]{1,2}[-./月][0-9]{1,2}日?)/,
  ]);
  const totalAmount = parseAmount(
    firstMatch(source, [
      /价税合计[^¥￥\d]{0,16}[¥￥]?\s*([0-9,]+\.\d{2})/i,
      /(?:金额|票价)[：:\s¥￥]*([0-9,]+\.\d{2})/i,
      /[¥￥]\s*([0-9,]+\.\d{2})/,
    ]),
  );
  const taxAmount = parseAmount(
    firstMatch(source, [/(?:税额合计|税额)[：:\s¥￥]*([0-9,]+\.\d{2})/i]),
  );
  const buyerName = firstMatch(source, [
    /购买方信息.*?名称[：:\s]*([^：:]{2,60}?)(?=统一社会信用代码|纳税人识别号|销售方信息)/i,
    /购买方名称[：:\s]*([^：:]{2,60}?)(?=统一社会信用代码|纳税人识别号|$)/i,
  ]);
  const sellerName = firstMatch(source, [
    /销售方信息.*?名称[：:\s]*([^：:]{2,60}?)(?=统一社会信用代码|纳税人识别号|备注|$)/i,
    /销售方名称[：:\s]*([^：:]{2,60}?)(?=统一社会信用代码|纳税人识别号|$)/i,
  ]);
  const trainNumber = firstMatch(source.toUpperCase(), [
    /(?:车次[：:\s]*)?([GCDKTYZ]\d{1,4})(?:次|\s|$)/,
  ]);
  const routeMatch = source.match(
    /([\u4e00-\u9fa5]{2,12}(?:站|南|北|东|西))\s*(?:→|—|-|至)\s*([\u4e00-\u9fa5]{2,12}(?:站|南|北|东|西))/,
  );

  return {
    fileName,
    invoiceType: INVOICE_TYPES[classification.type],
    invoiceCode,
    invoiceNumber,
    issueDate,
    buyerName,
    sellerName,
    totalAmount,
    amountWithoutTax:
      totalAmount != null && taxAmount != null ? Number((totalAmount - taxAmount).toFixed(2)) : null,
    taxAmount,
    trainNumber,
    departure: routeMatch?.[1] || "",
    arrival: routeMatch?.[2] || "",
    route: routeMatch ? `${routeMatch[1]}-${routeMatch[2]}` : "",
    summary: firstMatch(source, [
      /项目名称[^\u4e00-\u9fa5]{0,20}([\u4e00-\u9fa5A-Za-z0-9（）()·\s]{2,50}?)(?=规格型号|单位|数量|单价|金额)/i,
      /货物或应税劳务、服务名称[：:\s]*([^：:]{2,50}?)(?=规格型号|单位|数量|$)/i,
    ]),
    rawText: source,
  };
}

export async function analyzePdf(pdf, fileName, maxPages = 2) {
  const parts = [];
  for (let index = 1; index <= Math.min(pdf.numPages, maxPages); index += 1) {
    const page = await pdf.getPage(index);
    const content = await page.getTextContent();
    parts.push(content.items.map((item) => item.str).join(" "));
  }
  const text = compact(parts.join("\n"));
  const classification = classifyInvoice(text, fileName);
  return {
    text,
    ...classification,
    details: extractInvoiceFields(text, fileName, classification),
  };
}
