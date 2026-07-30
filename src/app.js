import "./style.css";
import { PDFDocument, rgb } from "pdf-lib";
import { getDocument, GlobalWorkerOptions } from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import { analyzePdf, extractInvoiceFields, INVOICE_TYPES } from "./invoice-analysis.js";

GlobalWorkerOptions.workerSrc = workerUrl;

const A4 = { width: 595.28, height: 841.89 };
const PREVIEW = { width: 794, height: 1123 };
const CELL_MARGIN = 18;
const PDFJS_ASSET_BASE = `${import.meta.env.BASE_URL}pdfjs/`;

function loadLocalPdf(data) {
  return getDocument({
    data,
    cMapUrl: `${PDFJS_ASSET_BASE}cmaps/`,
    cMapPacked: true,
    standardFontDataUrl: `${PDFJS_ASSET_BASE}standard_fonts/`,
    wasmUrl: `${PDFJS_ASSET_BASE}wasm/`,
    useSystemFonts: true,
  });
}

const els = {
  fileInput: document.querySelector("#fileInput"),
  dropzone: document.querySelector("#dropzone"),
  fileList: document.querySelector("#fileList"),
  fileCount: document.querySelector("#fileCount"),
  clearButton: document.querySelector("#clearButton"),
  canvas: document.querySelector("#previewCanvas"),
  prevButton: document.querySelector("#prevButton"),
  nextButton: document.querySelector("#nextButton"),
  currentPage: document.querySelector("#currentPage"),
  totalPages: document.querySelector("#totalPages"),
  exportPageCount: document.querySelector("#exportPageCount"),
  copyCountInputs: document.querySelectorAll("[data-copy-type]"),
  detailToggle: document.querySelector("#detailToggle"),
  detailTemplateBlock: document.querySelector("#detailTemplateBlock"),
  detailTemplateInput: document.querySelector("#detailTemplateInput"),
  detailTemplateName: document.querySelector("#detailTemplateName"),
  detailExportButton: document.querySelector("#detailExportButton"),
  detailExportButtonText: document.querySelector("#detailExportButtonText"),
  detailDownloadButton: document.querySelector("#detailDownloadButton"),
  exportButton: document.querySelector("#exportButton"),
  exportButtonText: document.querySelector("#exportButtonText"),
  manualButton: document.querySelector("#manualButton"),
  successCard: document.querySelector("#successCard"),
  toast: document.querySelector("#toast"),
};

const state = {
  files: [],
  pages: [],
  currentSheet: 0,
  mode: "color",
  copyCounts: {
    ordinary: 1,
    special: 2,
    train: 2,
    highSpeed: 2,
  },
  exportDetails: false,
  detailTemplate: null,
  detailOutput: null,
  aiAvailable: null,
  exportUrl: null,
  renderToken: 0,
};

function createDemoInvoice(index) {
  return {
    id: `demo-${index}`,
    name: index === 1 ? "示例发票_办公用品.pdf" : "示例发票_技术服务.pdf",
    sizeLabel: index === 1 ? "184 KB · 1 页" : "216 KB · 1 页",
    isDemo: true,
    pageCount: 1,
    invoiceType: index === 1 ? "ordinary" : "special",
    confidence: 1,
    reason: "演示数据",
    details: {
      fileName: index === 1 ? "示例发票_办公用品.pdf" : "示例发票_技术服务.pdf",
      invoiceType: index === 1 ? "普票" : "专票",
      invoiceNumber: index === 1 ? "253120000001864921" : "253120000001864936",
      issueDate: "2026年07月18日",
      buyerName: "杭州明川科技有限公司",
      sellerName: "瑞捷机械设备有限公司",
      totalAmount: index === 1 ? 1286 : 4800,
      summary: index === 1 ? "办公用品一批" : "设备技术服务费",
    },
    accent: index === 1 ? "#2f6f5a" : "#9f563f",
    amount: index === 1 ? "¥ 1,286.00" : "¥ 4,800.00",
    number: index === 1 ? "253120000001864921" : "253120000001864936",
  };
}

function resetToDemo() {
  state.files = [createDemoInvoice(1), createDemoInvoice(2)];
  rebuildPages();
}

function rebuildPages() {
  const invoicePages = state.files.flatMap((file) => {
    const copies = state.copyCounts[file.invoiceType] ?? 1;
    return Array.from({ length: copies }, (_, copyIndex) =>
      Array.from({ length: file.pageCount }, (_, pageIndex) => ({
        kind: "invoice",
        file,
        pageIndex,
        copyIndex,
      })),
    ).flat();
  });
  const detailPages =
    state.exportDetails && state.detailOutput
      ? [{ kind: "detail", canvas: state.detailOutput.previewCanvas, pageIndex: 0 }]
      : [];
  state.pages = invoicePages.concat(detailPages);
  const sheets = Math.max(1, Math.ceil(state.pages.length / 2));
  state.currentSheet = Math.min(state.currentSheet, sheets - 1);
  renderFileList();
  updateControls();
  renderPreview();
}

function renderFileList() {
  els.fileList.innerHTML = "";
  for (const file of state.files) {
    const li = document.createElement("li");
    li.className = "file-item";
    const confidenceNote =
      file.confidence < 0.6
        ? `<small class="recognition-warning" title="${escapeHtml(file.reason)}">请确认识别结果</small>`
        : `<small title="${escapeHtml(file.reason)}">${file.sizeLabel}</small>`;
    const totalAmount = file.details?.totalAmount;
    const amountLabel =
      totalAmount == null
        ? "金额未识别"
        : new Intl.NumberFormat("zh-CN", {
            style: "currency",
            currency: "CNY",
          }).format(totalAmount);
    li.innerHTML = `
      <span class="file-icon">PDF</span>
      <span class="file-copy">
        <strong title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</strong>
        ${confidenceNote}
        <span class="file-meta-row">
          <select class="invoice-type-select" aria-label="${escapeHtml(file.name)} 的发票类型">
            ${Object.entries(INVOICE_TYPES)
              .map(
                ([value, label]) =>
                  `<option value="${value}"${file.invoiceType === value ? " selected" : ""}>${label}</option>`,
              )
              .join("")}
          </select>
          <b class="invoice-total${totalAmount == null ? " is-missing" : ""}">${amountLabel}</b>
        </span>
      </span>
      <button class="remove-file" type="button" aria-label="移除 ${escapeHtml(file.name)}">×</button>
    `;
    li.querySelector(".invoice-type-select").addEventListener("change", (event) => {
      file.invoiceType = event.target.value;
      file.confidence = 1;
      file.reason = "用户手动选择";
      file.details = extractInvoiceFields(file.analysisText || "", file.name, {
        type: file.invoiceType,
      });
      invalidateDetailOutput();
      clearGeneratedDownload();
      rebuildPages();
    });
    li.querySelector(".remove-file").addEventListener("click", () => removeFile(file.id));
    els.fileList.appendChild(li);
  }
  els.fileCount.textContent = String(state.files.length);
}

function updateControls() {
  const sheets = Math.max(1, Math.ceil(state.pages.length / 2));
  els.currentPage.textContent = String(state.currentSheet + 1);
  els.totalPages.textContent = String(sheets);
  els.exportPageCount.textContent = `${sheets} 页`;
  els.prevButton.disabled = state.currentSheet === 0;
  els.nextButton.disabled = state.currentSheet >= sheets - 1;
  els.exportButton.disabled = state.pages.length === 0;
  els.detailExportButton.disabled =
    !state.exportDetails ||
    !state.detailTemplate ||
    !state.files.some((file) => !file.isDemo);
  els.detailDownloadButton.disabled = !state.detailOutput;
  if (state.exportDetails && !state.detailOutput) {
    els.exportButton.disabled = true;
    els.exportButtonText.textContent = "请先生成明细表";
  } else if (!state.exportUrl) {
    els.exportButtonText.textContent = "生成并下载 PDF";
  }
}

function removeFile(id) {
  state.files = state.files.filter((file) => file.id !== id);
  invalidateDetailOutput();
  clearGeneratedDownload();
  rebuildPages();
}

function escapeHtml(value) {
  const div = document.createElement("div");
  div.textContent = value;
  return div.innerHTML;
}

function readableSize(bytes) {
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

async function renderFirstPageForAi(pdf) {
  const page = await pdf.getPage(1);
  const base = page.getViewport({ scale: 1 });
  const scale = Math.min(1200 / base.width, 1200 / base.height, 2);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.round(viewport.width));
  canvas.height = Math.max(1, Math.round(viewport.height));
  await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
  return canvas.toDataURL("image/jpeg", 0.68);
}

function mergeAiAnalysis(local, ai) {
  const type = local.confidence < 0.6 ? ai.type : local.type;
  const details = {
    ...local.details,
    invoiceType: INVOICE_TYPES[type],
  };
  [
    "totalAmount",
    "invoiceNumber",
    "issueDate",
    "buyerName",
    "sellerName",
    "trainNumber",
    "route",
    "summary",
  ].forEach((field) => {
    if ((details[field] == null || details[field] === "") && ai[field] != null) {
      details[field] = ai[field];
    }
  });
  return {
    ...local,
    type,
    confidence: local.confidence < 0.6 ? ai.confidence : local.confidence,
    reason: "本地规则识别，缺失字段由云端 AI 补充",
    details,
    aiUsed: true,
  };
}

async function maybeAnalyzeWithAi(pdf, fileName, localAnalysis) {
  const needsAi = localAnalysis.confidence < 0.6 || localAnalysis.details.totalAmount == null;
  if (!needsAi || state.aiAvailable === false) return localAnalysis;
  try {
    const payload = {
      fileName,
      text: localAnalysis.text.slice(0, 5000),
    };
    if (localAnalysis.text.length < 80) {
      payload.imageDataUrl = await renderFirstPageForAi(pdf);
    }
    const response = await fetch("/api/analyze", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (response.status === 503) {
      state.aiAvailable = false;
      return localAnalysis;
    }
    if (!response.ok) return localAnalysis;
    const result = await response.json();
    state.aiAvailable = true;
    return mergeAiAnalysis(localAnalysis, result.data);
  } catch (error) {
    console.warn("AI fallback unavailable", error);
    return localAnalysis;
  }
}

async function addFiles(fileList) {
  const pdfFiles = [...fileList].filter(
    (file) => file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf"),
  );
  if (!pdfFiles.length) {
    showToast("请选择 PDF 文件");
    return;
  }

  if (state.files.some((file) => file.isDemo)) {
    state.files = [];
    state.currentSheet = 0;
    invalidateDetailOutput();
    clearGeneratedDownload();
    rebuildPages();
  }

  els.dropzone.classList.add("is-loading");
  const newFiles = [];
  for (const file of pdfFiles) {
    try {
      const bytes = await file.arrayBuffer();
      const pdf = await loadLocalPdf(new Uint8Array(bytes.slice(0))).promise;
      const localAnalysis = await analyzePdf(pdf, file.name);
      const analysis = await maybeAnalyzeWithAi(pdf, file.name, localAnalysis);
      newFiles.push({
        id: `${Date.now()}-${crypto.randomUUID()}`,
        name: file.name,
        sizeLabel: `${readableSize(file.size)} · ${pdf.numPages} 页`,
        isDemo: false,
        bytes,
        pdf,
        pageCount: pdf.numPages,
        invoiceType: analysis.type,
        confidence: analysis.confidence,
        reason: analysis.reason,
        analysisText: analysis.text,
        details: analysis.details,
        aiUsed: Boolean(analysis.aiUsed),
      });
    } catch (error) {
      console.error(error);
      showToast(`${file.name} 无法读取，已跳过`);
    }
  }
  els.dropzone.classList.remove("is-loading");

  if (newFiles.length) {
    state.files = state.files.concat(newFiles);
    state.currentSheet = 0;
    invalidateDetailOutput();
    clearGeneratedDownload();
    rebuildPages();
    const uncertainCount = newFiles.filter((file) => file.confidence < 0.6).length;
    const aiCount = newFiles.filter((file) => file.aiUsed).length;
    showToast(
      aiCount
        ? `已识别 ${newFiles.length} 个 PDF，其中 ${aiCount} 个由 AI 补充`
        : uncertainCount
        ? `已加入 ${newFiles.length} 个 PDF，${uncertainCount} 个类型需要确认`
        : `已识别 ${newFiles.length} 个 PDF`,
    );
  }
  els.fileInput.value = "";
}

function drawPaperBase(ctx) {
  ctx.save();
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, PREVIEW.width, PREVIEW.height);
  ctx.strokeStyle = "#e5e8e3";
  ctx.lineWidth = 1;
  ctx.setLineDash([5, 5]);
  ctx.beginPath();
  ctx.moveTo(0, PREVIEW.height / 2);
  ctx.lineTo(PREVIEW.width, PREVIEW.height / 2);
  ctx.stroke();
  ctx.setLineDash([]);
  ctx.fillStyle = "#a1aaa4";
  ctx.font = "500 10px system-ui, sans-serif";
  ctx.fillText("A5 / 01", 18, 23);
  ctx.fillText("A5 / 02", 18, PREVIEW.height / 2 + 23);
  ctx.restore();
}

function drawEmptySlot(ctx, slotIndex) {
  const halfHeight = PREVIEW.height / 2;
  const y = slotIndex * halfHeight;
  ctx.save();
  ctx.strokeStyle = "#e6e9e5";
  ctx.setLineDash([6, 6]);
  ctx.strokeRect(55, y + 55, PREVIEW.width - 110, halfHeight - 110);
  ctx.fillStyle = "#a6aea8";
  ctx.textAlign = "center";
  ctx.font = "500 13px system-ui, sans-serif";
  ctx.fillText("等待发票", PREVIEW.width / 2, y + halfHeight / 2);
  ctx.restore();
}

function createDemoCanvas(file, scale = 2) {
  const canvas = document.createElement("canvas");
  canvas.width = 1000 * scale;
  canvas.height = 620 * scale;
  const ctx = canvas.getContext("2d");
  ctx.scale(scale, scale);
  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, 1000, 620);
  ctx.strokeStyle = "#d9ded9";
  ctx.lineWidth = 2;
  ctx.strokeRect(1, 1, 998, 618);

  ctx.fillStyle = file.accent;
  ctx.fillRect(0, 0, 12, 620);
  ctx.fillStyle = "#17221d";
  ctx.font = "700 36px system-ui, sans-serif";
  ctx.fillText("电子发票（增值税普通发票）", 48, 72);
  ctx.fillStyle = "#7f8982";
  ctx.font = "500 15px system-ui, sans-serif";
  ctx.fillText(`发票号码  ${file.number}`, 680, 55);
  ctx.fillText("开票日期  2026年07月18日", 680, 83);

  ctx.strokeStyle = "#dce1dc";
  ctx.lineWidth = 1;
  ctx.strokeRect(48, 112, 904, 112);
  ctx.fillStyle = "#8c958f";
  ctx.font = "500 14px system-ui, sans-serif";
  ctx.fillText("购买方信息", 64, 141);
  ctx.fillText("销售方信息", 526, 141);
  ctx.fillStyle = "#28352e";
  ctx.font = "600 16px system-ui, sans-serif";
  ctx.fillText("名称：杭州明川科技有限公司", 64, 176);
  ctx.fillText("名称：瑞捷机械设备有限公司", 526, 176);
  ctx.font = "400 13px system-ui, sans-serif";
  ctx.fillText("统一社会信用代码：91330100MA2B****2X", 64, 204);
  ctx.fillText("统一社会信用代码：91320594MA1****18P", 526, 204);

  const rows = [
    ["项目名称", "规格型号", "单位", "数量", "单价", "金额", "税率", "税额"],
    [file.id.endsWith("1") ? "办公用品一批" : "设备技术服务费", "—", "项", "1", file.amount.slice(2), file.amount.slice(2), "6%", file.id.endsWith("1") ? "72.79" : "271.70"],
  ];
  const xs = [48, 310, 450, 525, 600, 700, 808, 875];
  ctx.strokeStyle = "#dce1dc";
  ctx.strokeRect(48, 250, 904, 130);
  ctx.beginPath();
  ctx.moveTo(48, 295);
  ctx.lineTo(952, 295);
  xs.slice(1).forEach((x) => {
    ctx.moveTo(x, 250);
    ctx.lineTo(x, 380);
  });
  ctx.stroke();
  rows.forEach((row, rowIndex) => {
    ctx.fillStyle = rowIndex === 0 ? "#748078" : "#26342c";
    ctx.font = `${rowIndex === 0 ? 500 : 400} 13px system-ui, sans-serif`;
    row.forEach((text, index) => ctx.fillText(text, xs[index] + 10, rowIndex === 0 ? 278 : 335));
  });

  ctx.fillStyle = "#7d8780";
  ctx.font = "500 14px system-ui, sans-serif";
  ctx.fillText("价税合计（大写）", 64, 424);
  ctx.fillStyle = "#27352e";
  ctx.font = "600 15px system-ui, sans-serif";
  ctx.fillText(file.id.endsWith("1") ? "壹仟贰佰捌拾陆元整" : "肆仟捌佰元整", 210, 424);
  ctx.fillStyle = file.accent;
  ctx.font = "700 26px system-ui, sans-serif";
  ctx.fillText(file.amount, 770, 426);

  ctx.strokeStyle = "#e1e5e1";
  ctx.beginPath();
  ctx.moveTo(48, 462);
  ctx.lineTo(952, 462);
  ctx.stroke();
  ctx.fillStyle = "#69756d";
  ctx.font = "400 13px system-ui, sans-serif";
  ctx.fillText("备注：本发票为产品功能演示内容，不具备报销效力。", 64, 504);
  ctx.fillText("开票人：系统演示", 64, 552);
  ctx.fillText("收款人：系统演示", 300, 552);
  ctx.fillStyle = file.accent;
  ctx.globalAlpha = 0.13;
  ctx.beginPath();
  ctx.arc(855, 525, 55, 0, Math.PI * 2);
  ctx.fill();
  ctx.globalAlpha = 1;
  ctx.fillStyle = file.accent;
  ctx.font = "700 15px system-ui, sans-serif";
  ctx.fillText("发票专用章", 817, 531);
  return canvas;
}

function fitRect(sourceWidth, sourceHeight, boxX, boxY, boxWidth, boxHeight) {
  const scale = Math.min(boxWidth / sourceWidth, boxHeight / sourceHeight);
  const width = sourceWidth * scale;
  const height = sourceHeight * scale;
  return {
    x: boxX + (boxWidth - width) / 2,
    y: boxY + (boxHeight - height) / 2,
    width,
    height,
  };
}

function applyGrayscale(canvas) {
  const ctx = canvas.getContext("2d");
  const image = ctx.getImageData(0, 0, canvas.width, canvas.height);
  for (let i = 0; i < image.data.length; i += 4) {
    const gray = Math.round(
      image.data[i] * 0.299 + image.data[i + 1] * 0.587 + image.data[i + 2] * 0.114,
    );
    image.data[i] = gray;
    image.data[i + 1] = gray;
    image.data[i + 2] = gray;
  }
  ctx.putImageData(image, 0, 0);
  return canvas;
}

function cloneCanvas(source) {
  const canvas = document.createElement("canvas");
  canvas.width = source.width;
  canvas.height = source.height;
  canvas.getContext("2d").drawImage(source, 0, 0);
  return canvas;
}

function drawCellText(ctx, text, x, y, width, height, style, scale) {
  if (!text) return;
  const padding = 5 * scale;
  const fontSize = style.fontSize * scale;
  ctx.fillStyle = style.color;
  ctx.font = `${style.bold ? 700 : 400} ${fontSize}px "PingFang SC", "Microsoft YaHei", sans-serif`;
  ctx.textBaseline = "middle";
  ctx.textAlign =
    style.horizontal === "center" ? "center" : style.horizontal === "right" ? "right" : "left";
  const textX =
    style.horizontal === "center"
      ? x + width / 2
      : style.horizontal === "right"
        ? x + width - padding
        : x + padding;
  const maxWidth = Math.max(1, width - padding * 2);
  let display = String(text);
  while (ctx.measureText(display).width > maxWidth && display.length > 1) {
    display = `${display.slice(0, -2)}…`;
  }
  ctx.save();
  ctx.beginPath();
  ctx.rect(x + 1, y + 1, width - 2, height - 2);
  ctx.clip();
  ctx.fillText(display, textX, y + height / 2);
  ctx.restore();
}

function createDetailPreviewCanvas(model) {
  const canvas = document.createElement("canvas");
  canvas.width = 1400;
  canvas.height = 990;
  const ctx = canvas.getContext("2d");
  const margin = 55;
  const availableWidth = canvas.width - margin * 2;
  const availableHeight = canvas.height - margin * 2;
  const modelWidth = model.columnWidths.reduce((sum, width) => sum + width, 0);
  const modelHeight = model.rowHeights.reduce((sum, height) => sum + height, 0);
  const scale = Math.min(availableWidth / modelWidth, availableHeight / modelHeight, 2);
  const tableWidth = modelWidth * scale;
  const tableHeight = modelHeight * scale;
  const startX = (canvas.width - tableWidth) / 2;
  const startY = margin + Math.max(0, (availableHeight - tableHeight) / 2);
  const xs = [startX];
  const ys = [startY];

  model.columnWidths.forEach((width) => xs.push(xs.at(-1) + width * scale));
  model.rowHeights.forEach((height) => ys.push(ys.at(-1) + height * scale));

  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  for (const cell of model.cells) {
    const x = xs[cell.column];
    const y = ys[cell.row];
    const width = xs[Math.min(model.columnCount, cell.column + (cell.columnSpan || 1))] - x;
    const height = ys[Math.min(model.rowCount, cell.row + (cell.rowSpan || 1))] - y;
    ctx.fillStyle = cell.fill;
    ctx.fillRect(x, y, width, height);
    drawCellText(ctx, cell.text, x, y, width, height, cell, scale);
    ctx.strokeStyle = "#cfd6d1";
    ctx.lineWidth = Math.max(1, scale * 0.7);
    ctx.beginPath();
    ctx.rect(x, y, width, height);
    ctx.stroke();
  }
  return canvas;
}

async function renderSourcePage(pageRef, maxWidth, maxHeight, grayscale = false) {
  if (pageRef.kind === "detail") {
    const detail = cloneCanvas(pageRef.canvas);
    return grayscale ? applyGrayscale(detail) : detail;
  }
  if (pageRef.file.isDemo) {
    const demo = createDemoCanvas(pageRef.file, 1.3);
    return grayscale ? applyGrayscale(demo) : demo;
  }
  const page = await pageRef.file.pdf.getPage(pageRef.pageIndex + 1);
  const baseViewport = page.getViewport({ scale: 1 });
  const scale = Math.min(maxWidth / baseViewport.width, maxHeight / baseViewport.height, 2.2);
  const viewport = page.getViewport({ scale });
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(1, Math.floor(viewport.width));
  canvas.height = Math.max(1, Math.floor(viewport.height));
  await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
  return grayscale ? applyGrayscale(canvas) : canvas;
}

async function renderPreview() {
  const token = ++state.renderToken;
  const ctx = els.canvas.getContext("2d");
  drawPaperBase(ctx);
  const halfHeight = PREVIEW.height / 2;
  const pageRefs = state.pages.slice(state.currentSheet * 2, state.currentSheet * 2 + 2);

  for (let slot = 0; slot < 2; slot += 1) {
    const pageRef = pageRefs[slot];
    if (!pageRef) {
      drawEmptySlot(ctx, slot);
      continue;
    }
    try {
      const source = await renderSourcePage(
        pageRef,
        PREVIEW.width - 2 * CELL_MARGIN,
        halfHeight - 2 * CELL_MARGIN,
        state.mode === "grayscale",
      );
      if (token !== state.renderToken) return;
      const rect = fitRect(
        source.width,
        source.height,
        CELL_MARGIN,
        slot * halfHeight + CELL_MARGIN,
        PREVIEW.width - CELL_MARGIN * 2,
        halfHeight - CELL_MARGIN * 2,
      );
      ctx.drawImage(source, rect.x, rect.y, rect.width, rect.height);
    } catch (error) {
      console.error(error);
      drawEmptySlot(ctx, slot);
    }
  }
}

function dataUrlToBytes(dataUrl) {
  const binary = atob(dataUrl.split(",")[1]);
  return Uint8Array.from(binary, (char) => char.charCodeAt(0));
}

async function embedRasterPage(output, pageRef, grayscale) {
  const canvas = await renderSourcePage(pageRef, 1500, 1100, grayscale);
  const png = await output.embedPng(dataUrlToBytes(canvas.toDataURL("image/png")));
  return { embedded: png, width: canvas.width, height: canvas.height };
}

async function createOutputPdf() {
  const output = await PDFDocument.create();
  const sourceDocs = new Map();

  if (state.mode === "color") {
    for (const file of state.files.filter((item) => !item.isDemo)) {
      sourceDocs.set(file.id, await PDFDocument.load(file.bytes.slice(0)));
    }
  }

  for (let sheetIndex = 0; sheetIndex < Math.ceil(state.pages.length / 2); sheetIndex += 1) {
    const sheet = output.addPage([A4.width, A4.height]);
    const pageRefs = state.pages.slice(sheetIndex * 2, sheetIndex * 2 + 2);

    for (let slot = 0; slot < pageRefs.length; slot += 1) {
      const pageRef = pageRefs[slot];
      const cellY = slot === 0 ? A4.height / 2 : 0;
      const box = {
        x: CELL_MARGIN / 1.33,
        y: cellY + CELL_MARGIN / 1.33,
        width: A4.width - (CELL_MARGIN * 2) / 1.33,
        height: A4.height / 2 - (CELL_MARGIN * 2) / 1.33,
      };

      if (pageRef.kind === "invoice" && state.mode === "color" && !pageRef.file.isDemo) {
        const [embedded] = await output.embedPdf(sourceDocs.get(pageRef.file.id), [
          pageRef.pageIndex,
        ]);
        const rect = fitRect(embedded.width, embedded.height, box.x, box.y, box.width, box.height);
        sheet.drawPage(embedded, rect);
      } else {
        const image = await embedRasterPage(output, pageRef, state.mode === "grayscale");
        const rect = fitRect(image.width, image.height, box.x, box.y, box.width, box.height);
        sheet.drawImage(image.embedded, rect);
      }
    }

    sheet.drawLine({
      start: { x: 0, y: A4.height / 2 },
      end: { x: A4.width, y: A4.height / 2 },
      thickness: 0.35,
      color: rgb(0.84, 0.86, 0.84),
      dashArray: [3, 3],
    });
  }
  return output.save();
}

function triggerDownload(url) {
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `发克票_双拼_${new Date().toISOString().slice(0, 10)}.pdf`;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
}

function triggerBlobDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function getDetailInvoiceRows() {
  return state.files
    .filter((file) => !file.isDemo)
    .map((file) => ({
      ...file.details,
      invoiceType: INVOICE_TYPES[file.invoiceType],
    }));
}

async function requestDetailPlan(template, invoiceRows) {
  const response = await fetch("/api/detail-plan", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      template,
      invoices: invoiceRows.slice(0, 3).map((invoice) => ({
        invoiceType: invoice.invoiceType || "",
        invoiceNumber: invoice.invoiceNumber || "",
        issueDate: invoice.issueDate || "",
        buyerName: invoice.buyerName || "",
        sellerName: invoice.sellerName || "",
        summary: invoice.summary || "",
        totalAmount: invoice.totalAmount ?? null,
        trainNumber: invoice.trainNumber || "",
        route: invoice.route || "",
      })),
    }),
  });
  const result = await response.json().catch(() => ({}));
  if (!response.ok) {
    if (result.error === "AI_NOT_CONFIGURED") {
      throw new Error("Vercel 尚未配置 AI API，无法分析明细表模板");
    }
    throw new Error("AI 未能理解该模板，请检查模板表头后重试");
  }
  return result.data;
}

async function generateDetailWorkbook() {
  if (!state.exportDetails) throw new Error("请先开启明细表导出");
  if (!state.detailTemplate) throw new Error("请先上传明细表模板");
  const { detailOutputName, fillDetailWorkbook, inspectDetailTemplate } =
    await import("./detail-workbook.js");
  const invoiceRows = getDetailInvoiceRows();
  if (!invoiceRows.length) throw new Error("演示发票不写入明细表，请先上传真实发票");
  const templateBytes = state.detailTemplate.bytes.slice(0);
  const templateSnapshot = await inspectDetailTemplate(templateBytes);
  const plan = await requestDetailPlan(templateSnapshot, invoiceRows);
  const { bytes, preview } = await fillDetailWorkbook(
    templateBytes,
    invoiceRows,
    plan,
  );
  state.detailOutput = {
    bytes,
    fileName: detailOutputName(state.detailTemplate.name),
    previewCanvas: createDetailPreviewCanvas(preview),
  };
}

async function handleGenerateDetail() {
  if (els.detailExportButton.disabled) return;
  els.detailExportButton.disabled = true;
  els.detailExportButtonText.textContent = "AI 正在分析模板…";
  try {
    await generateDetailWorkbook();
    clearGeneratedDownload();
    rebuildPages();
    state.currentSheet = Math.max(0, Math.ceil(state.pages.length / 2) - 1);
    updateControls();
    renderPreview();
    els.detailExportButtonText.textContent = "重新生成明细表";
    showToast("明细表已生成，并已加入 PDF 预览");
  } catch (error) {
    console.error(error);
    showToast(error.message || "明细表生成失败");
    els.detailExportButtonText.textContent = "生成明细表并加入预览";
  }
  updateControls();
}

function downloadDetailWorkbook() {
  if (!state.detailOutput) return;
  triggerBlobDownload(
    new Blob([state.detailOutput.bytes], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }),
    state.detailOutput.fileName,
  );
  showToast("Excel 明细表下载已开始");
}

function invalidateDetailOutput() {
  state.detailOutput = null;
  if (els.detailExportButtonText) {
    els.detailExportButtonText.textContent = "生成明细表并加入预览";
  }
}

function clearGeneratedDownload() {
  if (state.exportUrl) URL.revokeObjectURL(state.exportUrl);
  state.exportUrl = null;
  els.successCard.hidden = true;
  els.manualButton.hidden = true;
}

async function exportPdf() {
  if (!state.pages.length) return;
  if (state.exportDetails && !state.detailOutput) {
    showToast("请先生成明细表，再导出 PDF");
    return;
  }
  clearGeneratedDownload();
  els.exportButton.disabled = true;
  els.exportButtonText.textContent = "正在本地生成…";
  try {
    await new Promise((resolve) => setTimeout(resolve, 80));
    const bytes = await createOutputPdf();
    state.exportUrl = URL.createObjectURL(new Blob([bytes], { type: "application/pdf" }));
    triggerDownload(state.exportUrl);
    els.successCard.hidden = false;
    els.manualButton.hidden = false;
    showToast("PDF 已按当前预览生成，下载已开始");
  } catch (error) {
    console.error(error);
    showToast(error.message || "生成失败，请检查文件后重试");
  } finally {
    els.exportButton.disabled = state.pages.length === 0;
    els.exportButtonText.textContent = "重新生成并下载";
  }
}

let toastTimer;
function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("is-visible");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => els.toast.classList.remove("is-visible"), 2200);
}

els.fileInput.addEventListener("change", (event) => addFiles(event.target.files));
["dragenter", "dragover"].forEach((eventName) => {
  els.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropzone.classList.add("is-dragging");
  });
});
["dragleave", "drop"].forEach((eventName) => {
  els.dropzone.addEventListener(eventName, (event) => {
    event.preventDefault();
    els.dropzone.classList.remove("is-dragging");
  });
});
els.dropzone.addEventListener("drop", (event) => addFiles(event.dataTransfer.files));
els.clearButton.addEventListener("click", () => {
  state.files = [];
  invalidateDetailOutput();
  clearGeneratedDownload();
  rebuildPages();
});
els.prevButton.addEventListener("click", () => {
  if (state.currentSheet > 0) {
    state.currentSheet -= 1;
    updateControls();
    renderPreview();
  }
});
els.nextButton.addEventListener("click", () => {
  const maxSheet = Math.max(0, Math.ceil(state.pages.length / 2) - 1);
  if (state.currentSheet < maxSheet) {
    state.currentSheet += 1;
    updateControls();
    renderPreview();
  }
});
document.querySelectorAll('input[name="colorMode"]').forEach((radio) => {
  radio.addEventListener("change", (event) => {
    state.mode = event.target.value;
    clearGeneratedDownload();
    renderPreview();
  });
});
els.copyCountInputs.forEach((input) => {
  input.addEventListener("change", (event) => {
    const type = event.target.dataset.copyType;
    const count = Math.max(0, Math.min(9, Number(event.target.value) || 0));
    event.target.value = String(count);
    state.copyCounts[type] = count;
    clearGeneratedDownload();
    rebuildPages();
  });
});
els.detailToggle.addEventListener("change", (event) => {
  state.exportDetails = event.target.checked;
  els.detailTemplateBlock.hidden = !state.exportDetails;
  clearGeneratedDownload();
  rebuildPages();
});
els.detailTemplateInput.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    showToast("请选择 .xlsx 格式的明细表模板");
    event.target.value = "";
    return;
  }
  state.detailTemplate = {
    name: file.name,
    bytes: await file.arrayBuffer(),
  };
  invalidateDetailOutput();
  els.detailTemplateName.textContent = file.name;
  clearGeneratedDownload();
  rebuildPages();
  showToast("明细表模板已读取");
});
els.detailExportButton.addEventListener("click", handleGenerateDetail);
els.detailDownloadButton.addEventListener("click", downloadDetailWorkbook);
els.exportButton.addEventListener("click", exportPdf);
els.manualButton.addEventListener("click", () => {
  if (state.exportUrl) triggerDownload(state.exportUrl);
});

window.addEventListener("beforeunload", () => {
  if (state.exportUrl) URL.revokeObjectURL(state.exportUrl);
});

resetToDemo();
