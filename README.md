# 发克票

一个完全在浏览器本地运行的发票双拼排版工具。支持一次选择多个 PDF，将发票页面按每张 A4 上下两张的方式等比缩放、预览并导出。

## 本地运行

```bash
npm install
npm run dev
```

建议使用 Node.js 20.19 或更高版本。

## 构建

```bash
npm run build
```

构建结果位于 `dist` 目录。项目已包含 `vercel.json`，可直接导入 Vercel，或在项目目录执行 Vercel CLI 部署。

## 隐私说明

PDF 的读取、预览、排版和导出均在浏览器内完成。文件不会上传到任何服务器。
