import { NextResponse } from 'next/server';

export async function POST(request: Request) {
  try {
    const body = await request.json();
    if (!body) return NextResponse.json({ error: 'No data provided' }, { status: 400 });

    const jsonString = JSON.stringify(body, null, 2);
    return new NextResponse(jsonString, {
      status: 200,
      headers: {
        'Content-Type': 'application/json',
        'Content-Disposition': 'attachment; filename="slide-data.json"',
        'Content-Length': Buffer.byteLength(jsonString).toString(),
      },
    });
  } catch (error) {
    console.error('PPTX save-json error:', error);
    const message = error instanceof Error ? error.message : 'Failed to save JSON';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
