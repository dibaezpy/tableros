import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "Visor de Excel",
  description: "Lectura de la hoja costo del archivo Excel",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="es">
      <body>{children}</body>
    </html>
  );
}
