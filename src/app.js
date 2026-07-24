import "./style.css";
import { PDFDocument, rgb } from "pdf-lib";
import { getDocument, GlobalWorkerOptions } from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";

GlobalWorkerOptions.workerSrc = workerUrl;

const A4 = { width: 595.28, height: 841.89 };
const PREVIEW = { width: 794, height: 1123 };
const CELL_MARGIN = 18;

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
  state.pages = state.files.flatMap((file) =>
    Array.from({ length: file.pageCount }, (_, pageIndex) => ({ file, pageIndex })),
  );
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
    li.innerHTML = `
      <span class="file-icon">PDF</span>
      <span class="file-copy">
        <strong title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</strong>
        <small>${file.sizeLabel}</small>
      </span>
      <button class="remove-file" type="button" aria-label="移除 ${escapeHtml(file.name)}">×</button>
    `;
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
}

function removeFile(id) {
  state.files = state.files.filter((file) => file.id !== id);
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

async function addFiles(fileList) {
  const pdfFiles = [...fileList].filter(
    (file) => file.type === "application/pdf" || file.name.toLowerCase().endsWith(".pdf"),
  );
  if (!pdfFiles.length) {
    showToast("请选择 PDF 文件");
    return;
  }

  els.dropzone.classList.add("is-loading");
  const newFiles = [];
  for (const file of pdfFiles) {
    try {
      const bytes = await file.arrayBuffer();
      const pdf = await getDocument({ data: new Uint8Array(bytes.slice(0)) }).promise;
      newFiles.push({
        id: `${Date.now()}-${crypto.randomUUID()}`,
        name: file.name,
        sizeLabel: `${readableSize(file.size)} · ${pdf.numPages} 页`,
        isDemo: false,
        bytes,
        pdf,
        pageCount: pdf.numPages,
      });
    } catch (error) {
      console.error(error);
      showToast(`${file.name} 无法读取，已跳过`);
    }
  }
  els.dropzone.classList.remove("is-loading");

  if (newFiles.length) {
    state.files = state.files.filter((file) => !file.isDemo).concat(newFiles);
    state.currentSheet = 0;
    clearGeneratedDownload();
    rebuildPages();
    showToast(`已加入 ${newFiles.length} 个 PDF`);
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
  ctx.moveTo(24, PREVIEW.height / 2);
  ctx.lineTo(PREVIEW.width - 24, PREVIEW.height / 2);
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

async function renderSourcePage(pageRef, maxWidth, maxHeight, grayscale = false) {
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

      if (state.mode === "color" && !pageRef.file.isDemo) {
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
      start: { x: 18, y: A4.height / 2 },
      end: { x: A4.width - 18, y: A4.height / 2 },
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

function clearGeneratedDownload() {
  if (state.exportUrl) URL.revokeObjectURL(state.exportUrl);
  state.exportUrl = null;
  els.successCard.hidden = true;
  els.manualButton.hidden = true;
}

async function exportPdf() {
  if (!state.pages.length) return;
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
    showToast("PDF 已生成，下载已开始");
  } catch (error) {
    console.error(error);
    showToast("生成失败，请检查 PDF 后重试");
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
els.exportButton.addEventListener("click", exportPdf);
els.manualButton.addEventListener("click", () => {
  if (state.exportUrl) triggerDownload(state.exportUrl);
});

window.addEventListener("beforeunload", () => {
  if (state.exportUrl) URL.revokeObjectURL(state.exportUrl);
});

resetToDemo();
