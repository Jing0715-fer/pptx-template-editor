import { NextResponse } from 'next/server';
import { access, stat } from 'fs/promises';
import path from 'path';

export const dynamic = 'force-dynamic';

export async function GET(request: Request) {
  try {
    const { searchParams } = new URL(request.url);
    const fileId = searchParams.get('fileId');

    if (!fileId) return NextResponse.json({ error: '缺少 fileId' }, { status: 400 });

    const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!uuidRegex.test(fileId)) return NextResponse.json({ error: '无效的 fileId' }, { status: 400 });

    const tempDir = path.join(process.cwd(), 'temp-uploads');
    const pptxPath = path.join(tempDir, `${fileId}.pptx`);

    try {
      await access(pptxPath);
      const fileStat = await stat(pptxPath);
      return NextResponse.json({ exists: true, fileSize: fileStat.size });
    } catch {
      return NextResponse.json({ exists: false });
    }
  } catch (error) {
    console.error('File check error:', error);
    return NextResponse.json({ error: '检查失败' }, { status: 500 });
  }
}
