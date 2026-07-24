import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "发克票 · 本地发票排版",
  description: "为报销而生的本地发票双拼排版与 PDF 导出工具",
  icons: {
    icon: "/app-icon.png",
    apple: "/app-icon.png",
  },
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
