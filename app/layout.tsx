import "./globals.css";
import type { Metadata } from "next";
export const metadata: Metadata={title:"QR Generator47",description:"Excel zu druckfertigen QR-Code-Etiketten"};
export default function RootLayout({children}:{children:React.ReactNode}){return <html lang="de"><body>{children}</body></html>}