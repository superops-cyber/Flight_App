'use client';

// This page implements a simple logbook scanner interface.
// It allows the user to capture a photo of a logbook page, run OCR (mocked),
// display the extracted entries, and provide a tools view for reviewing and
// exporting the data.
// The implementation uses client‑side React and Next.js 13 app directory.
//
// Note: OCR and AI correction are mocked for demonstration purposes.
// In a production system, you would replace the mock logic with calls to
// a real OCR service (e.g., Tesseract.js or a backend API) and an AI
// model for data validation.

import { useState, useRef, useEffect } from "react";
import { useRouter, useSearchParams } from "next/navigation";
import Head from "next/head";

interface LogEntry {
  date: string;
  flightTime: string;
  from: string;
  to: string;
  notes: string;
  ocrRaw: string;
  corrected: boolean;
}

interface ToolsPageProps {
  entries: LogEntry[];
  onExport: () => void;
  onBack: () => void;
}

function ToolsPage({ entries, onExport, onBack }: ToolsPageProps) {
  return (
    <div className="bg-white rounded-lg shadow p-6">
      <div className="flex justify-between items-center mb-4">
        <button
          onClick={onBack}
          className="px-4 py-2 bg-gray-500 text-white rounded hover:bg-gray-600"
        >
          Back to Scanner
        </button>
        <h2 className="text-xl font-bold">Logbook Entries</h2>
      </div>

      <div className="overflow-x-auto">
        <table className="min-w-full border border-gray-200">
          <thead>
            <tr className="bg-gray-100">
              <th className="px-4 py-2 text-left border">Date</th>
              <th className="px-4 py-2 text-left border">Flight Time</th>
              <th className="px-4 py-2 text-left border">From</th>
              <th className="px-4 py-2 text-left border">To</th>
              <th className="px-4 py-2 text-left border">Notes</th>
              <th className="px-4 py-2 text-left border">Corrected</th>
            </tr>
          </thead>
          <tbody>
            {entries.map((entry, index) => (
              <tr key={index} className="hover:bg-gray-50">
                <td className="px-4 py-2 border">{entry.date}</td>
                <td className="px-4 py-2 border">{entry.flightTime}</td>
                <td className="px-4 py-2 border">{entry.from}</td>
                <td className="px-4 py-2 border">{entry.to}</td>
                <td className="px-4 py-2 border">{entry.notes}</td>
                <td className="px-4 py-2 border">{entry.corrected ? "✓" : "✗"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="mt-6 flex gap-4">
        <button
          onClick={onExport}
          className="px-6 py-3 bg-green-500 text-white rounded-lg hover:bg-green-600"
        >
          Export CSV
        </button>
        <button
          onClick={() => alert("Add new entry")}
          className="px-6 py-3 bg-blue-500 text-white rounded-lg hover:bg-blue-600"
        >
          Add Entry
        </button>
      </div>
    </div>
  );
}

export default function LogScannerPage() {
  const router = useRouter();
  const searchParams = useSearchParams();
  const viewParam = searchParams.get('view');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Handle view parameter on initial load
  useEffect(() => {
    if (viewParam === 'logscan') {
      setShowTools(false);
    }
  }, [viewParam]);
  const [photo, setPhoto] = useState<string | null>(null);
  const [entries, setEntries] = useState<LogEntry[]>([]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [showTools, setShowTools] = useState(false);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onload = () => {
        setPhoto(reader.result as string);
      };
      reader.readAsDataURL(file);
    }
  };

  const handleCapture = () => {
    fileInputRef.current?.click();
  };

  const processImage = async () => {
    if (!photo) return;
    setIsProcessing(true);
    
    try {
      // Extract base64 data from photo URL
      const base64Data = photo.split(',')[1];
      
      const response = await fetch('/api/logscan', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify({
          image: base64Data,
          sheetName: "Scanned Flights"
        }),
      });
      
      if (!response.ok) throw new Error('Failed to process image');
      
      const result = await response.json();
      
      if (result.error) {
        throw new Error(result.error);
      }
      
      // Map backend response to our frontend format
      const mappedEntries = result.rawData.map((entry: any) => ({
        date: entry.date,
        flightTime: entry.duration,
        from: entry.route?.split('→')[0]?.trim() || "",
        to: entry.route?.split('→')[1]?.trim() || "",
        notes: entry.remarks,
        ocrRaw: JSON.stringify(entry),
        corrected: false
      }));
      
      setEntries(mappedEntries);
    } catch (error) {
      console.error('Processing error:', error);
      alert(`Error processing image: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleExportCSV = () => {
    if (entries.length === 0) return;
    const headers = ["Date", "Flight Time", "From", "To", "Notes", "Corrected"];
    const csvContent =
      [headers.join(",")]
        .concat(
          entries.map((e) =>
            [e.date, e.flightTime, e.from, e.to, e.notes, e.corrected]
              .join(",")
          )
        )
        .join("\n");
    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.setAttribute("href", url);
    link.setAttribute("download", "logbook-entries.csv");
    link.style.visibility = "hidden";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const handleBack = () => {
    if (showTools) {
      setShowTools(false);
      router.push('?view=logscan', { scroll: false });
    } else {
      router.push("/");
    }
  };

  const handleToolsClick = () => {
    setShowTools(true);
    // Update URL without page reload
    router.push('?view=tools', { scroll: false });
  };

  // Mobile viewport detection
  const [isMobile, setIsMobile] = useState(false);
  
  useEffect(() => {
    const checkMobile = () => {
      setIsMobile(window.innerWidth < 768);
    };
    checkMobile();
    window.addEventListener('resize', checkMobile);
    return () => window.removeEventListener('resize', checkMobile);
  }, []);

  return (
    <div className="min-h-screen bg-gray-50 p-4">
      <Head>
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no" />
      </Head>
      <div className="max-w-2xl mx-auto">
        <div className="flex justify-between items-center mb-6">
          <button
            onClick={handleBack}
            className="px-4 py-2 bg-blue-500 text-white rounded hover:bg-blue-600"
          >
            Back
          </button>
          <h1 className="text-2xl font-bold">Logbook Scanner</h1>
          {entries.length > 0 && (
            <button
              onClick={handleToolsClick}
              className="px-4 py-2 bg-green-500 text-white rounded hover:bg-green-600"
            >
              Tools
            </button>
          )}
        </div>

        {!photo && !showTools && (
          <div className="bg-white rounded-lg shadow p-6 text-center">
            <div className="mb-4">
              <svg
                className="mx-auto h-16 w-16 text-gray-400"
                fill="none"
                stroke="currentColor"
                viewBox="0 0 24 24"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth={2}
                  d="M3 7h18M3 7v18h18M3 7l3 0 9 12 9-12 3 0"
                />
              </svg>
            </div>
            <p className="text-gray-600 mb-4">
              Take a photo of your logbook page or upload an image
            </p>
            <button
              onClick={handleCapture}
              className="px-6 py-3 bg-blue-500 text-white rounded-lg hover:bg-blue-600"
            >
              {isMobile ? "Use Camera" : "Capture Photo"}
            </button>
            <p className="text-sm text-gray-500 mt-2">
              {isMobile ? "Tap to access your camera" : "Or use your device camera to take a picture"}
            </p>
          </div>
        )}

        {photo && !showTools && (
          <div className="bg-white rounded-lg shadow p-6">
            <div className="mb-4">
              <img
                src={photo}
                alt="Captured logbook page"
                className="max-w-full h-auto rounded"
              />
            </div>
            {isProcessing ? (
              <div className="text-center">
                <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-500 mx-auto"></div>
                <p className="mt-4 text-gray-600">
                  Processing with OCR and AI...
                </p>
              </div>
            ) : (
              <button
                onClick={processImage}
                className="w-full px-6 py-3 bg-green-500 text-white rounded-lg hover:bg-green-600"
              >
               {isMobile ? "Process Now" : "Process Logbook Page"}
             </button>
            )}
          </div>
        )}

        {showTools && (
          <ToolsPage entries={entries} onExport={handleExportCSV} onBack={handleBack} />
        )}

        <input
          ref={fileInputRef}
          type="file"
          accept="image/*"
          capture="environment"
          onChange={handleFileSelect}
          className="hidden"
        />
      </div>
    </div>
  );
}
