import { NextResponse } from 'next/server';
import { readFile, mkdir, readdir, unlink, rmdir } from 'fs/promises';
import { execFile } from 'child_process';
import { promisify } from 'util';
import path from 'path';
import { randomUUID } from 'crypto';

const execFileAsync = promisify(execFile);
export const maxDuration = 120;

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const { fileId } = body as { fileId: string };

    if (!fileId) return NextResponse.json({ error: '缺少 fileId' }, { status: 400 });

    const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!uuidRegex.test(fileId)) return NextResponse.json({ error: '无效的 fileId' }, { status: 400 });

    const tempDir = path.join(process.cwd(), 'temp-uploads');
    const pptxPath = path.join(tempDir, `${fileId}.pptx`);

    let pptxBuffer: Buffer;
    try { pptxBuffer = await readFile(pptxPath); } catch {
      return NextResponse.json({ error: '文件不存在或已过期' }, { status: 404 });
    }

    const conversionId = randomUUID();
    const outputDir = path.join(process.cwd(), 'temp-uploads', `preview_${conversionId}`);
    await mkdir(outputDir, { recursive: true });

    try {
      const libreOfficePath = process.env.LIBREOFFICE_PATH || 'libreoffice';
      await execFileAsync(libreOfficePath, ['--headless', '--convert-to', 'png', '--outdir', outputDir, pptxPath], { timeout: 60000 });

      const files = await readdir(outputDir);
      const pngFiles = files.filter((f) => f.toLowerCase().endsWith('.png')).sort((a, b) => {
        const numA = parseInt(a.match(/\d+/)?.[0] || '0', 10);
        const numB = parseInt(b.match(/\d+/)?.[0] || '0', 10);
        return numA - numB;
      });

      const previewImages: string[] = [];
      for (const pngFile of pngFiles) {
        const pngPath = path.join(outputDir, pngFile);
        const pngBuffer = await readFile(pngPath);
        previewImages.push(`data:image/png;base64,${pngBuffer.toString('base64')}`);
      }

      try { for (const f of files) await unlink(path.join(outputDir, f)); await rmdir(outputDir); } catch { /* ignore */ }

      if (previewImages.length === 0)
        return NextResponse.json({ previewImages: [], warning: '无法生成预览图片，LibreOffice 可能未正确安装' });

      return NextResponse.json({ previewImages });
    } catch (conversionError) {
      try { const files = await readdir(outputDir); for (const f of files) await unlink(path.join(outputDir, f)); await rmdir(outputDir); } catch { /* ignore */ }
      console.error('LibreOffice conversion error:', conversionError);
      return NextResponse.json({ previewImages: [], warning: '预览生成失败' });
    }
  } catch (error) {
    console.error('Preview generation error:', error);
    const message = error instanceof Error ? error.message : '预览生成失败';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
