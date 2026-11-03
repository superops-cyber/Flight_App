import { NextResponse } from 'next/server';
// @ts-ignore
const { processLogbookScan } = require('../../../../../Backend_LogbookScanner');

export async function POST(request: Request) {
  try {
    const { image, sheetName } = await request.json();
    
    if (!image) {
      return NextResponse.json(
        { error: 'No image provided' },
        { status: 400 }
      );
    }

    // Process the image using existing backend function
    const result = processLogbookScan(image, sheetName);
    
    if (result.error) {
      return NextResponse.json(
        { error: result.error },
        { status: 500 }
      );
    }

    return NextResponse.json(result);
  } catch (error) {
    console.error('API Error:', error);
    return NextResponse.json(
      { error: 'Internal server error' },
      { status: 500 }
    );
  }
}

export const dynamic = 'force-dynamic'; // Ensure dynamic route handling