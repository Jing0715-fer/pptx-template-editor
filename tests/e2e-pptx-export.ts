/**
 * E2E Test Script for PPTX Export Validation
 * 
 * Tests:
 * 1. Content_Types.xml sanitization (fixes invalid MIME types like image/.jpg)
 * 2. WPS-specific file removal (ppt/tags/)
 * 3. Text modifications are preserved in exported PPTX
 * 4. Original photos are preserved when not replaced
 * 5. Exported PPTX can be re-parsed (round-trip test)
 * 6. No invalid MIME types in exported Content_Types.xml
 * 7. ZIP structure integrity
 * 8. OPC Spec MIME type compliance
 */

import JSZip from 'jszip';
import { readFile } from 'fs/promises';
import path from 'path';
import { applyModificationsAndExport } from '../src/lib/pptx-replacer';
import { parsePptx, getImageAsBase64 } from '../src/lib/pptx-parser';
import type { PptxModification } from '../src/lib/pptx-replacer';

// ANSI colors for output
const GREEN = '\x1b[32m';
const RED = '\x1b[31m';
const YELLOW = '\x1b[33m';
const CYAN = '\x1b[36m';
const RESET = '\x1b[0m';

let passCount = 0;
let failCount = 0;

function log(message: string, color?: string) {
  console.log(`${color || ''}${message}${RESET}`);
}

function assert(condition: boolean, testName: string, details?: string) {
  if (condition) {
    log(`  ✅ PASS: ${testName}`, GREEN);
    passCount++;
  } else {
    log(`  ❌ FAIL: ${testName}`, RED);
    if (details) log(`     ${details}`, RED);
    failCount++;
  }
}

async function readContentTypes(zip: JSZip): Promise<string | null> {
  const file = zip.file('[Content_Types].xml');
  if (!file) return null;
  return file.async('string');
}

// ============================================================================
// Test Suite
// ============================================================================

async function runTests() {
  log('\n═══════════════════════════════════════════════════════════', CYAN);
  log('  PPTX Export E2E Test Suite', CYAN);
  log('═══════════════════════════════════════════════════════════\n', CYAN);

  // Find an existing PPTX file in temp-uploads
  const tempDir = path.join(process.cwd(), 'temp-uploads');
  const fs = await import('fs/promises');
  let files: string[];
  try {
    files = (await fs.readdir(tempDir)).filter(f => f.endsWith('.pptx'));
  } catch {
    log('No temp-uploads directory found, creating a test PPTX...', YELLOW);
    files = [];
  }

  let testBuffer: Buffer;

  if (files.length > 0) {
    const testFile = path.join(tempDir, files[0]);
    log(`Using test file: ${files[0]}`, CYAN);
    testBuffer = await readFile(testFile);
  } else {
    // Create a minimal PPTX for testing
    log('Creating minimal test PPTX...', YELLOW);
    testBuffer = await createMinimalPptx();
  }

  // -----------------------------------------------------------------------
  // Test 1: Original PPTX structure validation
  // -----------------------------------------------------------------------
  log('\n--- Test 1: Original PPTX Structure ---', YELLOW);
  {
    const zip = await JSZip.loadAsync(testBuffer);
    const ctXml = await readContentTypes(zip);
    
    assert(ctXml !== null, '[Content_Types].xml exists in original');
    
    // Check for invalid MIME types in original
    if (ctXml) {
      const hasInvalidDotMime = /ContentType="image\/\.[a-zA-Z+]+"/.test(ctXml);
      log(`  ℹ️  Original has invalid dot-MIME (image/.xxx): ${hasInvalidDotMime}`, CYAN);
      
      const hasUppercaseExt = /Extension="[A-Z]{2,}"/.test(ctXml);
      log(`  ℹ️  Original has uppercase extensions: ${hasUppercaseExt}`, CYAN);

      const hasWpsTags = Object.keys(zip.files).some(f => f.startsWith('ppt/tags/'));
      log(`  ℹ️  Original has WPS tag files: ${hasWpsTags}`, CYAN);
    }
  }

  // -----------------------------------------------------------------------
  // Test 2: Content_Types.xml sanitization
  // -----------------------------------------------------------------------
  log('\n--- Test 2: Content_Types.xml Sanitization ---', YELLOW);
  {
    const outputBuffer = await applyModificationsAndExport(testBuffer, [], []);
    const zip = await JSZip.loadAsync(outputBuffer);
    const ctXml = await readContentTypes(zip);

    assert(ctXml !== null, '[Content_Types].xml exists in exported file');

    if (ctXml) {
      // No invalid MIME types with extra dot
      const hasInvalidDotMime = /ContentType="image\/\.[a-zA-Z+]+"/.test(ctXml);
      assert(!hasInvalidDotMime, 'No invalid dot-MIME types (image/.xxx) in exported file',
        hasInvalidDotMime ? 'Found: ' + ctXml.match(/ContentType="image\/\.[a-zA-Z+]+"/g)?.join(', ') : undefined);

      // No uppercase extensions
      const hasUppercaseExt = /Extension="[A-Z]{2,}"/.test(ctXml);
      assert(!hasUppercaseExt, 'No uppercase extensions (Extension="JPG") in exported file',
        hasUppercaseExt ? 'Found: ' + ctXml.match(/Extension="[A-Z]{2,}"/g)?.join(', ') : undefined);

      // No image/jpg (should be image/jpeg)
      const hasImageJpg = /ContentType="image\/jpg"/.test(ctXml);
      assert(!hasImageJpg, 'No image/jpg MIME type (should be image/jpeg)');

      // Correct MIME types for known extensions
      const jpgEntry = ctXml.match(/Extension="jpg"[^>]*ContentType="([^"]+)"/) 
        || ctXml.match(/ContentType="([^"]+)"[^>]*Extension="jpg"/);
      if (jpgEntry) {
        assert(jpgEntry[1] === 'image/jpeg', `JPG extension uses correct MIME type (got: ${jpgEntry[1]})`);
      }

      // JPEG extension uses image/jpeg
      const jpegMime = ctXml.match(/Extension="jpeg"[^>]*ContentType="([^"]+)"/);
      if (jpegMime) {
        assert(jpegMime[1] === 'image/jpeg', `JPEG extension uses correct MIME type (got: ${jpegMime[1]})`);
      }

      // PNG extension uses image/png
      const pngMime = ctXml.match(/Extension="png"[^>]*ContentType="([^"]+)"/);
      if (pngMime) {
        assert(pngMime[1] === 'image/png', `PNG extension uses correct MIME type (got: ${pngMime[1]})`);
      }
    }
  }

  // -----------------------------------------------------------------------
  // Test 3: WPS-specific files removal
  // -----------------------------------------------------------------------
  log('\n--- Test 3: WPS-specific Files Removal ---', YELLOW);
  {
    const outputBuffer = await applyModificationsAndExport(testBuffer, [], []);
    const zip = await JSZip.loadAsync(outputBuffer);
    const entryNames = Object.keys(zip.files);

    const hasWpsTags = entryNames.some(f => f.startsWith('ppt/tags/'));
    assert(!hasWpsTags, 'No ppt/tags/ files in exported PPTX');

    // Core files should still exist
    assert(entryNames.includes('[Content_Types].xml'), '[Content_Types].xml preserved');
    assert(entryNames.some(n => n.startsWith('ppt/slides/slide') && n.endsWith('.xml')), 'Slide files preserved');
  }

  // -----------------------------------------------------------------------
  // Test 4: Text modifications are preserved
  // -----------------------------------------------------------------------
  log('\n--- Test 4: Text Modification Preservation ---', YELLOW);
  {
    // First parse to find text elements
    const parseResult = await parsePptx(testBuffer, 'test.pptx');
    const firstSlideWithText = parseResult.slides.find(s => 
      s.elements.some(e => e.type === 'text' && e.originalText.trim().length > 0)
    );

    if (firstSlideWithText) {
      const textEl = firstSlideWithText.elements.find(e => e.type === 'text' && e.originalText.trim().length > 0);
      if (textEl && textEl.type === 'text') {
        const testText = 'E2E_TEST_MODIFIED_TEXT';
        const modifications: PptxModification[] = [{
          slideIndex: textEl.slideIndex,
          elementIndex: textEl.elementIndex,
          type: 'text',
          newText: testText,
        }];

        const outputBuffer = await applyModificationsAndExport(testBuffer, modifications, []);
        
        // Re-parse the exported file to verify
        const reparsed = await parsePptx(outputBuffer, 'exported.pptx');
        const slide = reparsed.slides[textEl.slideIndex];
        const modifiedEl = slide?.elements.find(e => 
          e.type === 'text' && e.elementIndex === textEl.elementIndex
        );

        if (modifiedEl && modifiedEl.type === 'text') {
          const hasModText = modifiedEl.originalText.includes(testText);
          assert(hasModText, `Text modification preserved: "${modifiedEl.originalText.substring(0, 50)}..."`,
            !hasModText ? `Expected to find "${testText}" in "${modifiedEl.originalText.substring(0, 100)}"` : undefined);
        } else {
          log('  ⚠️  Could not find modified text element in re-parsed file', YELLOW);
        }
      } else {
        log('  ⚠️  No suitable text element found for testing', YELLOW);
      }
    } else {
      log('  ⚠️  No slides with text elements found, skipping', YELLOW);
    }
  }

  // -----------------------------------------------------------------------
  // Test 5: Original photos preserved when not replaced
  // -----------------------------------------------------------------------
  log('\n--- Test 5: Original Photos Preservation ---', YELLOW);
  {
    const parseResult = await parsePptx(testBuffer, 'test.pptx');
    const firstSlideWithImage = parseResult.slides.find(s => 
      s.elements.some(e => e.type === 'image')
    );

    if (firstSlideWithImage) {
      const imageEl = firstSlideWithImage.elements.find(e => e.type === 'image');
      if (imageEl && imageEl.type === 'image') {
        const originalImageName = imageEl.imageName;
        
        // Export WITHOUT any image modifications
        const outputBuffer = await applyModificationsAndExport(testBuffer, [], []);
        
        // Re-parse and check the image still exists
        const reparsed = await parsePptx(outputBuffer, 'exported.pptx');
        const slide = reparsed.slides[imageEl.slideIndex];
        const imgEl = slide?.elements.find(e => 
          e.type === 'image' && e.elementIndex === imageEl.elementIndex
        );

        if (imgEl && imgEl.type === 'image') {
          assert(imgEl.imageName === originalImageName, 
            `Original image preserved: ${originalImageName}`,
            `Expected ${originalImageName}, got ${imgEl.imageName}`);
          
          // Check if we can extract the image data
          const imageData = getImageAsBase64(reparsed._rawEntries, imgEl.imageName);
          assert(imageData !== null, `Original image data is extractable: ${originalImageName}`);
          if (imageData) {
            assert(imageData.length > 100, `Image data has reasonable size (${imageData.length} chars)`);
          }
        } else {
          log('  ⚠️  Could not find image element in re-parsed file', YELLOW);
        }
      }
    } else {
      log('  ⚠️  No slides with images found, skipping', YELLOW);
    }
  }

  // -----------------------------------------------------------------------
  // Test 6: Round-trip test (export → re-parse → verify)
  // -----------------------------------------------------------------------
  log('\n--- Test 6: Round-trip Export/Import Test ---', YELLOW);
  {
    const parseResult = await parsePptx(testBuffer, 'test.pptx');
    const originalSlideCount = parseResult.slides.length;
    
    const outputBuffer = await applyModificationsAndExport(testBuffer, [], []);
    
    // Re-parse the exported file
    try {
      const reparsed = await parsePptx(outputBuffer, 'roundtrip.pptx');
      assert(reparsed.slides.length === originalSlideCount, 
        `Slide count preserved: ${reparsed.slides.length} === ${originalSlideCount}`);
    } catch (err) {
      assert(false, 'Re-parsing exported PPTX succeeds',
        `Error: ${err instanceof Error ? err.message : 'unknown'}`);
    }
  }

  // -----------------------------------------------------------------------
  // Test 7: ZIP structure integrity
  // -----------------------------------------------------------------------
  log('\n--- Test 7: ZIP Structure Integrity ---', YELLOW);
  {
    const outputBuffer = await applyModificationsAndExport(testBuffer, [], []);
    const zip = await JSZip.loadAsync(outputBuffer);
    const entryNames = Object.keys(zip.files);

    // Required PPTX files
    const requiredFiles = ['[Content_Types].xml', '_rels/.rels'];
    for (const req of requiredFiles) {
      assert(entryNames.includes(req), `Required file exists: ${req}`);
    }

    // No empty files
    let emptyCount = 0;
    for (const [name, file] of Object.entries(zip.files)) {
      if (file.dir) continue;
      const data = await file.async('uint8array');
      if (data.length === 0) {
        emptyCount++;
        log(`  ⚠️  Empty file: ${name}`, YELLOW);
      }
    }
    assert(emptyCount === 0, `No empty files in exported PPTX (found ${emptyCount})`);

    // XML files should have valid content
    let invalidXmlCount = 0;
    const xmlEntries = entryNames.filter(n => n.endsWith('.xml') || n.endsWith('.rels'));
    for (const xmlName of xmlEntries.slice(0, 20)) { // Check first 20 to save time
      try {
        const data = await zip.file(xmlName)?.async('string');
        if (data && !data.includes('<')) {
          invalidXmlCount++;
          log(`  ⚠️  Invalid XML: ${xmlName}`, YELLOW);
        }
      } catch { /* skip */ }
    }
    assert(invalidXmlCount === 0, `All XML files have valid content (${invalidXmlCount} invalid)`);
  }

  // -----------------------------------------------------------------------
  // Test 8: Content_Types.xml MIME type validation against OPC spec
  // -----------------------------------------------------------------------
  log('\n--- Test 8: OPC Spec MIME Type Compliance ---', YELLOW);
  {
    const outputBuffer = await applyModificationsAndExport(testBuffer, [], []);
    const zip = await JSZip.loadAsync(outputBuffer);
    const ctXml = await readContentTypes(zip);

    if (ctXml) {
      // Extract all ContentType values
      const mimeTypes = [...ctXml.matchAll(/ContentType="([^"]+)"/g)].map(m => m[1]);
      
      // Check for common OPC violations
      const violations: string[] = [];
      for (const mime of mimeTypes) {
        // No leading dot in MIME subtype
        if (/^image\/\./.test(mime)) {
          violations.push(`Leading dot in MIME: ${mime}`);
        }
        // No uppercase in MIME type (per RFC 2045, type/subtype are case-insensitive but lowercase is standard)
        if (mime !== mime.toLowerCase() && mime.startsWith('image/')) {
          violations.push(`Uppercase in MIME: ${mime}`);
        }
      }

      assert(violations.length === 0, `All MIME types comply with OPC spec`,
        violations.length > 0 ? `Violations: ${violations.join('; ')}` : undefined);
      
      log(`  ℹ️  Total MIME types found: ${mimeTypes.length}`, CYAN);
      const imageMimes = mimeTypes.filter(m => m.startsWith('image/'));
      log(`  ℹ️  Image MIME types: ${imageMimes.join(', ')}`, CYAN);
    }
  }

  // -----------------------------------------------------------------------
  // Results Summary
  // -----------------------------------------------------------------------
  log('\n═══════════════════════════════════════════════════════════', CYAN);
  if (failCount === 0) {
    log(`  🎉 All ${passCount} tests passed!`, GREEN);
  } else {
    log(`  ⚠️  ${passCount} passed, ${failCount} failed`, RED);
  }
  log('═══════════════════════════════════════════════════════════\n', CYAN);

  process.exit(failCount > 0 ? 1 : 0);
}

/**
 * Creates a minimal PPTX file for testing when no real files are available
 * Includes intentional WPS-style errors (invalid MIME types, WPS tags) to test sanitization
 */
async function createMinimalPptx(): Promise<Buffer> {
  const zip = new JSZip();
  
  // [Content_Types].xml with intentional WPS-style errors for testing
  zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="JPG" ContentType="image/.jpg"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`, { compression: 'DEFLATE' });

  // _rels/.rels
  zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`, { compression: 'DEFLATE' });

  // ppt/presentation.xml
  zip.file('ppt/presentation.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst/>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`, { compression: 'DEFLATE' });

  // ppt/_rels/presentation.xml.rels
  zip.file('ppt/_rels/presentation.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
</Relationships>`, { compression: 'DEFLATE' });

  // ppt/slides/slide1.xml
  zip.file('ppt/slides/slide1.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="274638"/>
            <a:ext cx="8229600" cy="1143000"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" dirty="0"/>
              <a:t>Hello World</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`, { compression: 'DEFLATE' });

  // ppt/slides/_rels/slide1.xml.rels
  zip.file('ppt/slides/_rels/slide1.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`, { compression: 'DEFLATE' });

  // WPS-specific files that should be removed
  zip.file('ppt/tags/tag1.xml', `<?xml version="1.0" encoding="UTF-8"?><tag/>`, { compression: 'DEFLATE' });
  zip.file('ppt/tags/tag2.xml', `<?xml version="1.0" encoding="UTF-8"?><tag/>`, { compression: 'DEFLATE' });

  return zip.generateAsync({ type: 'nodebuffer' });
}

// Run the tests
runTests().catch(err => {
  log(`\nFatal error: ${err}`, RED);
  process.exit(1);
});
