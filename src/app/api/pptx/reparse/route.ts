import { NextResponse } from 'next/server';
import { readFile, writeFile, mkdir, readdir, unlink, rmdir } from 'fs/promises';
import path from 'path';
import { parsePptx, getImageAsBase64 } from '@/lib/pptx-parser';
import { randomUUID } from 'crypto';
import { execFile } from 'child_process';
import { promisify } from 'util';

const execFileAsync = promisify(execFile);
export const maxDuration = 120;
export const dynamic = 'force-dynamic';
const MAX_PREVIEW_SIZE = 500 * 1024;

async function generatePreviewImages(pptxPath: string): Promise<string[]> {
  const conversionId = randomUUID();
  const outputDir = path.join(path.dirname(pptxPath), `preview_${conversionId}`);

  try {
    await mkdir(outputDir, { recursive: true });
    const libreOfficePath = process.env.LIBREOFFICE_PATH || 'libreoffice';

    const pdfFileName = path.basename(pptxPath, '.pptx') + '.pdf';
    const pdfPath = path.join(outputDir, pdfFileName);

    await execFileAsync(libreOfficePath, ['--headless', '--norestore', '--convert-to', 'pdf', '--outdir', outputDir, pptxPath], { timeout: 90000 });
    try { await readFile(pdfPath); } catch { return generatePreviewImagesFallback(pptxPath); }

    const imgPrefix = path.join(outputDir, 'slide');
    try {
      await execFileAsync('pdftoppm', ['-jpeg', '-jpegopt', 'quality=75', '-r', '96', '-hide-annotations', pdfPath, imgPrefix], { timeout: 60000 });
    } catch {
      try { await execFileAsync('convert', ['-density', '96', '-quality', '75', pdfPath, `${imgPrefix}.jpg`], { timeout: 60000 }); }
      catch { return generatePreviewImagesFallback(pptxPath); }
    }

    const files = await readdir(outputDir);
    const imageFiles = files
      .filter((f) => /\.(jpg|jpeg|png)$/i.test(f))
      .sort((a, b) => {
        const numA = parseInt(a.match(/(\d+)\.(?:jpg|jpeg|png)$/i)?.[1] || '0', 10);
        const numB = parseInt(b.match(/(\d+)\.(?:jpg|jpeg|png)$/i)?.[1] || '0', 10);
        return numA - numB;
      });

    const previewImages: string[] = [];
    for (const imgFile of imageFiles) {
      const imgPath = path.join(outputDir, imgFile);
      try {
        const imgData = await readFile(imgPath);
        if (imgData.length > MAX_PREVIEW_SIZE) { previewImages.push(''); continue; }
        const ext = imgFile.split('.').pop()?.toLowerCase() || 'jpg';
        const mime = ext === 'png' ? 'image/png' : 'image/jpeg';
        previewImages.push(`data:${mime};base64,${imgData.toString('base64')}`);
      } catch { previewImages.push(''); }
    }

    await cleanupDir(outputDir);
    return previewImages;
  } catch (error) {
    console.warn('LibreOffice preview generation failed for reparse:', error);
    await cleanupDir(outputDir);
    return generatePreviewImagesFallback(pptxPath);
  }
}

async function generatePreviewImagesFallback(pptxPath: string): Promise<string[]> {
  const conversionId = randomUUID();
  const outputDir = path.join(path.dirname(pptxPath), `preview_fb_${conversionId}`);
  try {
    await mkdir(outputDir, { recursive: true });
    const libreOfficePath = process.env.LIBREOFFICE_PATH || 'libreoffice';
    await execFileAsync(libreOfficePath, ['--headless', '--norestore', '--convert-to', 'png', '--outdir', outputDir, pptxPath], { timeout: 60000 });
    const files = await readdir(outputDir);
    const pngFiles = files.filter((f) => f.toLowerCase().endsWith('.png')).sort();
    const previewImages: string[] = [];
    for (const pngFile of pngFiles) {
      try {
        const pngData = await readFile(path.join(outputDir, pngFile));
        if (pngData.length > MAX_PREVIEW_SIZE) continue;
        previewImages.push(`data:image/png;base64,${pngData.toString('base64')}`);
      } catch { /* skip */ }
    }
    await cleanupDir(outputDir);
    return previewImages;
  } catch { await cleanupDir(outputDir); return []; }
}

async function cleanupDir(dirPath: string) {
  try { const files = await readdir(dirPath); for (const f of files) await unlink(path.join(dirPath, f)); await rmdir(dirPath); } catch { /* ignore */ }
}

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
      return NextResponse.json({ error: '文件不存在或已过期，请重新上传' }, { status: 404 });
    }

    let parseResult;
    try { parseResult = await parsePptx(pptxBuffer, `${fileId}.pptx`); } catch (parseError) {
      console.error('PPTX reparse error:', parseError);
      const msg = parseError instanceof Error ? parseError.message : '未知解析错误';
      return NextResponse.json({ error: `文件解析失败: ${msg}` }, { status: 500 });
    }

    if (!parseResult.slides || parseResult.slides.length === 0)
      return NextResponse.json({ error: '未在文件中找到幻灯片内容' }, { status: 400 });

    const newFileId = randomUUID();
    const newPptxPath = path.join(tempDir, `${newFileId}.pptx`);
    await writeFile(newPptxPath, pptxBuffer);

    let previewImages: string[] = [];
    try {
      previewImages = await generatePreviewImages(newPptxPath);
      console.log(`Reparse: Generated ${previewImages.filter(p => p).length}/${parseResult.slides.length} preview images`);
    } catch (err) { console.warn('Preview generation failed for reparse:', err); }

    const slides = parseResult.slides.map((slide, index) => ({
      slideNumber: slide.slideNumber,
      elements: slide.elements.map((element) => {
        if (element.type === 'image') {
          let imageData: string | null = null;
          try { imageData = getImageAsBase64(parseResult._rawEntries, element.imageName); } catch (err) {
            console.warn(`Failed to extract image ${element.imageName}:`, err);
          }
          return { ...element, imageData };
        }
        return element;
      }),
      previewImage: previewImages[index] || null,
    }));

    return NextResponse.json({ fileId: newFileId, fileName: parseResult.fileName, slideCount: parseResult.slideCount, slides });
  } catch (error) {
    console.error('PPTX reparse error:', error);
    const message = error instanceof Error ? error.message : '文件处理失败';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
