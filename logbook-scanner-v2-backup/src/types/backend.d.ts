declare module '../../Backend_LogbookScanner' {
  export function processLogbookScan(
    base64Image: string,
    sheetId?: string
  ): {
    success?: boolean;
    error?: string;
    count?: number;
    sheetId?: string;
    rawData?: Array<{
      date: string;
      aircraftType: string;
      tailNumber: string;
      route: string;
      departure: string;
      arrival: string;
      duration: string;
      picName: string;
      landings: string;
      remarks: string;
    }>;
  };
}