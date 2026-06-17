import { NextResponse } from 'next/server';
import { readFile } from 'fs/promises';
import path from 'path';
import { applyModificationsAndExport, testExportRoundTrip } from '@/lib/pptx-replacer';
import type { PptxModification, ImageModification } from '@/lib/pptx-replacer';

export const maxDuration = 120;
export const dynamic = 'force-dynamic';

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const { fileId, modifications, imageModifications } = body as {
      fileId: string;
      modifications: PptxModification[];
      imageModifications?: ImageModification[];
    };

    if (!fileId) return NextResponse.json({ error: '缺少 fileId' }, { status: 400 });

    if ((!modifications || !Array.isArray(modifications) || modifications.length === 0) &&
        (!imageModifications || !Array.isArray(imageModifications) || imageModifications.length === 0)) {
      return NextResponse.json({ error: '缺少修改内容' }, { status: 400 });
    }

    const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!uuidRegex.test(fileId)) return NextResponse.json({ error: '无效的 fileId' }, { status: 400 });

    const tempDir = path.join(process.cwd(), 'temp-uploads');
    const pptxPath = path.join(tempDir, `${fileId}.pptx`);

    let pptxBuffer: Buffer;
    try { pptxBuffer = await readFile(pptxPath); } catch {
      return NextResponse.json({ error: '文件不存在或已过期，请重新上传' }, { status: 404 });
    }

    if (imageModifications && imageModifications.length > 0) {
      for (const imgMod of imageModifications) {
        console.log(`Image mod: slide=${imgMod.slideIndex}, rId=${imgMod.imageRid}, type=${imgMod.newImageType}, dataLen=${imgMod.newImageData?.length || 0}`);
      }
    }

    const outputBuffer = await applyModificationsAndExport(
      pptxBuffer,
      modifications || [],
      imageModifications || [],
    );

    // Run round-trip test for additional verification
    const roundTripResult = await testExportRoundTrip(pptxBuffer, outputBuffer);
    if (!roundTripResult.success) {
      console.error('Round-trip test FAILED:', roundTripResult.details);
      // Still return the file but log the warning
    } else {
      console.log('Round-trip test passed:', roundTripResult.details);
    }

    // Return the PPTX as binary - avoid Content-Encoding header as it can cause issues
    // with proxies and browsers misinterpreting the binary data
    return new Response(outputBuffer, {
      status: 200,
      headers: {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'Content-Disposition': 'attachment; filename="modified.pptx"',
        'Content-Length': String(outputBuffer.length),
        'Cache-Control': 'no-cache, no-store, must-revalidate',
      },
    });
  } catch (error) {
    console.error('PPTX export error:', error);
    const message = error instanceof Error ? error.message : '导出失败';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
