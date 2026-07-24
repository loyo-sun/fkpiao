# 发克票

<p align="center">
  <img src="./public/app-icon.png" width="120" height="120" alt="发克票 icon">
</p>

<p align="center">
  <strong>为报销而生的本地发票排版工具</strong>
</p>

<p align="center">
  <a href="https://fkpiao.loyo.work/">在线演示</a>
</p>

## 为什么做这个工具

纸质报销很麻烦。

电子发票、高铁票和各种 PDF 票据需要逐张打开、调整大小、复制排版，再组合到 A4 纸上打印。票据一多，不仅操作重复，还很容易出现比例不一致、内容被裁切或漏放的问题。

**发克票**把这些步骤合并成了一次操作：批量导入 PDF 票据，工具会自动按照每张 A4 上下两张的方式等比缩放、分页预览并生成排版后的 PDF。确认无误后直接下载打印即可。

## 功能

- 批量导入多个 PDF 票据
- 自动按 A4 上下两张进行双拼排版
- 保持原始比例，避免票据信息裁切
- 支持多页预览和翻页检查
- 支持彩色、黑白两种导出模式
- PDF 生成后自动下载，也可手动再次下载
- 全程在浏览器本地处理，不上传票据文件

## 在线体验

访问：[https://fkpiao.loyo.work/](https://fkpiao.loyo.work/)

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
