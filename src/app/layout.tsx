import type { Metadata } from "next";
import "./globals.css";
import { Toaster } from "@/components/ui/sonner";
import { ThemeProvider } from "@/components/theme-provider";

// Use system fonts instead of Google Fonts to avoid build failures
// when Google Fonts API is unreachable (e.g. in China)
const geistSans = {
  variable: "--font-geist-sans",
  className: "",
  style: { fontFamily: "'Geist', 'Inter', system-ui, -apple-system, sans-serif" },
};

const geistMono = {
  variable: "--font-geist-mono",
  className: "",
  style: { fontFamily: "'Geist Mono', 'Fira Code', 'JetBrains Mono', monospace" },
};

export const metadata: Metadata = {
  title: "PPTX Template Editor - AI-Powered Presentation Editing",
  description: "Edit PPTX templates with ease. Upload presentations, detect template variables, fill them with AI or manually, and export polished PPTX files.",
  keywords: ["PPTX", "template editor", "AI", "presentation", "variable replacement", "Next.js"],
  authors: [{ name: "PPTX Editor Team" }],
  icons: {
    icon: "https://z-cdn.chatglm.cn/z-ai/static/logo.svg",
  },
  openGraph: {
    title: "PPTX Template Editor",
    description: "AI-powered PPTX template editing with smart variable detection",
    type: "website",
  },
  twitter: {
    card: "summary_large_image",
    title: "PPTX Template Editor",
    description: "AI-powered PPTX template editing with smart variable detection",
  },
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en" suppressHydrationWarning data-scroll-behavior="smooth">
      <body
        className={`${geistSans.variable} ${geistMono.variable} antialiased bg-background text-foreground`}
        style={{
          "--font-geist-sans": "'Geist', 'Inter', system-ui, -apple-system, sans-serif",
          "--font-geist-mono": "'Geist Mono', 'Fira Code', 'JetBrains Mono', monospace",
        } as React.CSSProperties}
      >
        <ThemeProvider
          attribute="class"
          defaultTheme="dark"
          enableSystem
          disableTransitionOnChange
        >
          {children}
        </ThemeProvider>
        <Toaster />
      </body>
    </html>
  );
}
