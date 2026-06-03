import { NextResponse } from 'next/server';
import { readFile, mkdir, readdir, unlink, rmdir } from 'fs/promises';
import { execFile } from 'child_process';
import { promisify } from 'util';
import path from 'path';
import { randomUUID } from 'crypto';
import { existsSync } from 'fs';

const execFileAsync = promisify(execFile);
export const maxDuration = 120;
export const dynamic = 'force-dynamic';

// Cache thumbnails in memory for 10 minutes
const thumbnailCache = new Map<string, { data: Buffer; timestamp: number }>();
const CACHE_TTL = 10 * 60 * 1000;

async function cleanupDir(dirPath: string) {
  try {
    const files = await readdir(dirPath);
    for (const f of files) await unlink(path.join(dirPath, f));
    await rmdir(dirPath);
  } catch { /* ignore */ }
}

/**
 * Generate a small JPEG thumbnail for the first slide of a PPTX file.
 * Returns the image data directly (not base64 JSON) for use as <img src>.
 */
export async function GET(request: Request) {
  try {
    const { searchParams } = new URL(request.url);
    const fileId = searchParams.get('fileId');

    if (!fileId) {
      return new NextResponse('Missing fileId', { status: 400 });
    }

    const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!uuidRegex.test(fileId)) {
      return new NextResponse('Invalid fileId', { status: 400 });
    }

    // Check memory cache
    const cached = thumbnailCache.get(fileId);
    if (cached && Date.now() - cached.timestamp < CACHE_TTL) {
      return new NextResponse(cached.data, {
        headers: {
          'Content-Type': 'image/jpeg',
          'Cache-Control': 'public, max-age=600, s-maxage=600',
          'X-Content-Type-Options': 'nosniff',
        },
      });
    }

    const tempDir = path.join(process.cwd(), 'temp-uploads');
    const pptxPath = path.join(tempDir, `${fileId}.pptx`);

    if (!existsSync(pptxPath)) {
      return new NextResponse('File not found', { status: 404 });
    }

    const conversionId = randomUUID();
    const outputDir = path.join(tempDir, `thumb_${conversionId}`);
    await mkdir(outputDir, { recursive: true });

    try {
      const libreOfficePath = process.env.LIBREOFFICE_PATH || 'libreoffice';

      // Convert first page to PNG
      await execFileAsync(libreOfficePath, [
        '--headless', '--norestore',
        '--convert-to', 'png',
        '--outdir', outputDir,
        pptxPath,
      ], { timeout: 60000 });

      const files = await readdir(outputDir);
      // Sort to get the first slide
      const pngFiles = files
        .filter((f) => f.toLowerCase().endsWith('.png'))
        .sort((a, b) => {
          const numA = parseInt(a.match(/\d+/)?.[0] || '0', 10);
          const numB = parseInt(b.match(/\d+/)?.[0] || '0', 10);
          return numA - numB;
        });

      if (pngFiles.length === 0) {
        await cleanupDir(outputDir);
        return new NextResponse('No preview generated', { status: 500 });
      }

      // Read the first slide PNG
      const firstPng = path.join(outputDir, pngFiles[0]);
      const pngBuffer = await readFile(firstPng);

      // Try to convert to small JPEG using ImageMagick for better quality/size
      let jpegBuffer: Buffer;
      const thumbPath = path.join(outputDir, 'thumb.jpg');
      try {
        await execFileAsync('convert', [
          pngBuffer.toString().length > 0 ? firstPng : firstPng,
          '-resize', '320x180>',
          '-quality', '60',
          '-background', 'white',
          '-flatten',
          thumbPath,
        ], { timeout: 30000 });
        jpegBuffer = await readFile(thumbPath);
      } catch {
        // If ImageMagick fails, try with ffmpeg
        try {
          await execFileAsync('ffmpeg', [
            '-i', firstPng,
            '-vf', 'scale=320:180:force_original_aspect_ratio=decrease,pad=320:180:(ow-iw)/2:(oh-ih)/2:white',
            '-q:v', '4',
            '-y', thumbPath,
          ], { timeout: 30000 });
          jpegBuffer = await readFile(thumbPath);
        } catch {
          // Fallback: return the PNG as-is (it's still a valid image)
          await cleanupDir(outputDir);
          // Cache the PNG too
          thumbnailCache.set(fileId, { data: pngBuffer, timestamp: Date.now() });
          return new NextResponse(pngBuffer, {
            headers: {
              'Content-Type': 'image/png',
              'Cache-Control': 'public, max-age=600, s-maxage=600',
              'X-Content-Type-Options': 'nosniff',
            },
          });
        }
      }

      await cleanupDir(outputDir);

      // Cache the thumbnail
      thumbnailCache.set(fileId, { data: jpegBuffer, timestamp: Date.now() });

      return new NextResponse(jpegBuffer, {
        headers: {
          'Content-Type': 'image/jpeg',
          'Cache-Control': 'public, max-age=600, s-maxage=600',
          'X-Content-Type-Options': 'nosniff',
        },
      });
    } catch (conversionError) {
      await cleanupDir(outputDir);
      console.error('Thumbnail generation error:', conversionError);
      return new NextResponse('Thumbnail generation failed', { status: 500 });
    }
  } catch (error) {
    console.error('Thumbnail API error:', error);
    return new NextResponse('Internal server error', { status: 500 });
  }
}
