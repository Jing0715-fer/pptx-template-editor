import { NextResponse } from 'next/server';
import { randomUUID } from 'crypto';
import { parsePptx, getImageAsBase64 } from '@/lib/pptx-parser';
import { readFile, writeFile, mkdir, readdir, unlink, rmdir } from 'fs/promises';
import { execFile } from 'child_process';
import { promisify } from 'util';
import path from 'path';

const execFileAsync = promisify(execFile);
export const maxDuration = 120;
const MAX_PREVIEW_SIZE = 500 * 1024;

async function ensureTempDir() {
  const tempDir = path.join(process.cwd(), 'temp-uploads');
  try { await mkdir(tempDir, { recursive: true }); } catch { /* exists */ }
  return tempDir;
}

async function generatePreviewImages(pptxPath: string): Promise<string[]> {
  const conversionId = randomUUID();
  const outputDir = path.join(path.dirname(pptxPath), `preview_${conversionId}`);

  try {
    await mkdir(outputDir, { recursive: true });
    const libreOfficePath = process.env.LIBREOFFICE_PATH || 'libreoffice';

    const pdfFileName = path.basename(pptxPath, '.pptx') + '.pdf';
    const pdfPath = path.join(outputDir, pdfFileName);

    await execFileAsync(libreOfficePath, ['--headless', '--norestore', '--convert-to', 'pdf', '--outdir', outputDir, pptxPath], { timeout: 90000 });

    try { await readFile(pdfPath); } catch {
      return generatePreviewImagesFallback(pptxPath);
    }

    const imgPrefix = path.join(outputDir, 'slide');
    try {
      await execFileAsync('pdftoppm', ['-jpeg', '-jpegopt', 'quality=75', '-r', '96', '-hide-annotations', pdfPath, imgPrefix], { timeout: 60000 });
    } catch {
      try {
        await execFileAsync('convert', ['-density', '96', '-quality', '75', pdfPath, `${imgPrefix}.jpg`], { timeout: 60000 });
      } catch { return generatePreviewImagesFallback(pptxPath); }
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
    console.warn('LibreOffice preview generation failed:', error);
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
      const pngPath = path.join(outputDir, pngFile);
      try {
        const pngData = await readFile(pngPath);
        if (pngData.length > MAX_PREVIEW_SIZE) continue;
        previewImages.push(`data:image/png;base64,${pngData.toString('base64')}`);
      } catch { /* skip */ }
    }

    await cleanupDir(outputDir);
    return previewImages;
  } catch {
    await cleanupDir(outputDir);
    return [];
  }
}

async function cleanupDir(dirPath: string) {
  try {
    const files = await readdir(dirPath);
    for (const f of files) await unlink(path.join(dirPath, f));
    await rmdir(dirPath);
  } catch { /* ignore */ }
}

export async function POST(request: Request) {
  try {
    const contentType = request.headers.get('content-type') || '';
    if (!contentType.includes('multipart/form-data'))
      return NextResponse.json({ error: '请使用 multipart/form-data 格式上传文件' }, { status: 400 });

    let formData: FormData;
    try { formData = await request.formData(); } catch {
      return NextResponse.json({ error: '无法读取上传数据，请重试' }, { status: 400 });
    }

    const file = formData.get('file') as File | null;
    if (!file) return NextResponse.json({ error: '请选择文件' }, { status: 400 });

    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.pptx'))
      return NextResponse.json({ error: '仅支持 .pptx 格式文件' }, { status: 400 });
    if (file.size > 100 * 1024 * 1024)
      return NextResponse.json({ error: '文件大小不能超过 100MB' }, { status: 400 });
    if (file.size === 0)
      return NextResponse.json({ error: '文件内容为空' }, { status: 400 });

    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    if (buffer.length === 0)
      return NextResponse.json({ error: '文件内容为空' }, { status: 400 });

    let parseResult;
    try { parseResult = await parsePptx(buffer, file.name); } catch (parseError) {
      console.error('PPTX parse error:', parseError);
      const msg = parseError instanceof Error ? parseError.message : '未知解析错误';
      return NextResponse.json({ error: `文件解析失败: ${msg}` }, { status: 500 });
    }

    if (!parseResult.slides || parseResult.slides.length === 0)
      return NextResponse.json({ error: '未在文件中找到幻灯片内容' }, { status: 400 });

    const fileId = randomUUID();
    let pptxPath: string | null = null;
    try {
      const tempDir = await ensureTempDir();
      pptxPath = path.join(tempDir, `${fileId}.pptx`);
      await writeFile(pptxPath, buffer);
    } catch (err) { console.error('Failed to save temp file:', err); }

    let previewImages: string[] = [];
    if (pptxPath) {
      try {
        previewImages = await generatePreviewImages(pptxPath);
        console.log(`Generated ${previewImages.filter(p => p).length}/${parseResult.slides.length} preview images`);
      } catch (err) { console.warn('Preview generation failed:', err); }
    }

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

    return NextResponse.json({ fileId, fileName: parseResult.fileName, slideCount: parseResult.slideCount, slides, slideSize: parseResult.slideSize });
  } catch (error) {
    console.error('PPTX upload error:', error);
    const message = error instanceof Error ? error.message : '文件处理失败';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
