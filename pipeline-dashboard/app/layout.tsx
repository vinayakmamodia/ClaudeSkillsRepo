import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'LeadSquared FY27 Dashboard',
  description: 'Pipeline, MRR performance and renewal risk dashboard',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
