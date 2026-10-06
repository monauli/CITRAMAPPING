import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Mapping Keuangan",
  description: "Audit Excel dan PDF laporan keuangan",
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="id">
      <body>{children}</body>
    </html>
  );
}
