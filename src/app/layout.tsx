import type {Metadata} from 'next';
import { Geist, Geist_Mono } from 'next/font/google'; // Corrected import names
import './globals.css';
import { Toaster } from "@/components/ui/toaster"; // Import Toaster

const geistSans = Geist({ // Corrected function name
  variable: '--font-geist-sans',
  subsets: ['latin'],
});

const geistMono = Geist_Mono({ // Corrected function name
  variable: '--font-geist-mono',
  subsets: ['latin'],
});

export const metadata: Metadata = {
  title: 'CSV to XLS Converter', // Updated App Name
  description: 'Convert CSV files to XLS format easily.', // Updated Description
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (<html lang="en" className="dark">{/* Apply dark mode by default */}
      <body className={`${geistSans.variable} ${geistMono.variable} antialiased bg-background text-foreground`}>
        {children}
        <Toaster /> {/* Add Toaster component here */}
      </body>
    </html>
  );
}
