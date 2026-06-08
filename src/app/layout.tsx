import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Soteria Merlin WorkOrder Dashboard",
  description: "Administrative console for bulk work order upload operations and deletions.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
