
"use client";

import React, { useState, useCallback } from 'react';
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Upload, Download, Loader2 } from 'lucide-react';
import { convertCsvToXls } from '@/services/excel-converter';
import { useToast } from '@/hooks/use-toast';

export default function Home() {
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);
  const [downloadFilename, setDownloadFilename] = useState<string>('converted_data.xls');
  const { toast } = useToast();

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file && file.type === 'text/csv') {
      setSelectedFile(file);
      setDownloadUrl(null);
      // Attempt to create a safe filename, replace non-alphanumeric with underscore
      const safeBaseName = file.name.replace(/\.csv$/i, '').replace(/[^a-z0-9]/gi, '_');
      setDownloadFilename(`${safeBaseName}.xls`);
    } else {
      setSelectedFile(null);
      event.target.value = ''; // Clear the input if the file is invalid
      toast({
        variant: "destructive",
        title: "Invalid File Type",
        description: "Please upload a valid CSV file.",
      });
    }
  };

  const handleConvert = useCallback(async () => {
    if (!selectedFile) {
      toast({
        variant: "destructive",
        title: "No File Selected",
        description: "Please upload a CSV file first.",
      });
      return;
    }

    setIsLoading(true);
    setDownloadUrl(null);

    try {
      const reader = new FileReader();
      reader.onload = async (event) => {
        const csvData = event.target?.result as string;
        if (!csvData) {
           throw new Error("Failed to read file content.");
        }
        try {
          const xlsBuffer = await convertCsvToXls(csvData);
          const blob = new Blob([xlsBuffer], { type: 'application/vnd.ms-excel' });
          const url = URL.createObjectURL(blob);
          setDownloadUrl(url);
          toast({
            title: "Conversion Successful",
            description: "Your XLS file is ready for download.",
          });
        } catch (convertError: any) {
           console.error("Conversion error:", convertError);
           toast({
             variant: "destructive",
             title: "Conversion Failed",
             description: convertError.message || "An error occurred during conversion.",
           });
        } finally {
          setIsLoading(false);
        }
      };
      reader.onerror = (error) => {
         console.error("File reading error:", error);
         toast({
           variant: "destructive",
           title: "File Reading Error",
           description: "Could not read the selected file.",
         });
         setIsLoading(false);
      }
      reader.readAsText(selectedFile); // Read as text for CSV parsing

    } catch (error: any) {
      console.error("Error during conversion process:", error);
      toast({
        variant: "destructive",
        title: "Conversion Error",
        description: error.message || "An unexpected error occurred.",
      });
      setIsLoading(false);
    }
  }, [selectedFile, toast]);

  return (
    <main className="flex min-h-screen flex-col items-center justify-center p-6 bg-background text-foreground">
      <Card className="w-full max-w-md shadow-lg shadow-[0_0_25px_3px_rgba(64,64,64,0.4)] rounded-xl">
        <CardHeader className="text-center">
          <CardTitle className="text-2xl font-bold">RadioDJ CSV to Report XLS</CardTitle>
          <CardDescription>Upload your CSV file to convert it into XLS format.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-6">
          <div className="space-y-2">
            <label htmlFor="csv-upload" className="block text-sm font-medium text-foreground/80">
              Upload CSV File
            </label>
            <div className="flex items-center space-x-2">
               <Input
                 id="csv-upload"
                 type="file"
                 accept=".csv"
                 onChange={handleFileChange}
                 className="flex-grow text-center file:mx-auto file:py-1 file:px-2 file:rounded-md file:border-0 file:text-xs file:font-semibold file:bg-primary/10 file:text-primary hover:file:bg-primary/20"
               />
               <Upload className="h-5 w-5 text-muted-foreground" />
            </div>
             {selectedFile && (
                <p className="text-sm text-muted-foreground pt-1 text-center">
                  Selected: {selectedFile.name}
                </p>
            )}
          </div>

          <Button
            onClick={handleConvert}
            disabled={!selectedFile || isLoading}
            className="w-full bg-primary hover:bg-primary/90 text-primary-foreground rounded-lg shadow-md"
          >
            {isLoading ? (
              <Loader2 className="mr-2 h-4 w-4 animate-spin" />
            ) : null}
            {isLoading ? 'Converting...' : 'Convert to XLS'}
          </Button>

          {downloadUrl && (
            <div className="text-center">
              <a
                href={downloadUrl}
                download={downloadFilename}
                className="inline-flex items-center justify-center px-4 py-2 border border-transparent text-sm font-medium rounded-lg text-primary-foreground bg-primary hover:bg-primary/90 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-primary shadow-md"
              >
                <Download className="mr-2 h-4 w-4" />
                Download XLS File
              </a>
            </div>
          )}
        </CardContent>
      </Card>
       <div className="mt-8 text-center text-xs text-muted-foreground opacity-70">
        <p>made by Bogdan Turlacu</p>
      </div>
    </main>
  );
}

