import type { Metadata } from 'next';
import './globals.css';
import LowDeliveryReport from './LowDeliveryReport';

export const metadata: Metadata = {
  title: 'ECI Hearing & Notice Monitoring Dashboard',
  description: 'PS-wise ECI hearing, notice and DEO monitoring dashboard',
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="en"><body>{children}<LowDeliveryReport /></body></html>;
}