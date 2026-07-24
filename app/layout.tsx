import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "发克票 · 本地发票排版",
  description: "完全在浏览器本地运行的发票双拼排版与 PDF 导出工具",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
