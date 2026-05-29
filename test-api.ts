import { readFileSync, writeFileSync } from 'fs';

async function testApiExport() {
  // Test the actual API endpoint
  const fileId = '7139432a-818b-4d9c-839a-d749945faa16';
  
  const modifications = [{
    slideIndex: 0,
    elementIndex: 6,
    type: 'text',
    newText: 'API_TEST_MODIFICATION'
  }];
  
  try {
    const response = await fetch('http://localhost:3000/api/pptx/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileId, modifications, imageModifications: [] }),
    });
    
    console.log('Response status:', response.status);
    console.log('Response headers:');
    response.headers.forEach((value, key) => {
      console.log(`  ${key}: ${value}`);
    });
    
    if (!response.ok) {
      const text = await response.text();
      console.log('Error response:', text.substring(0, 500));
      return;
    }
    
    const buffer = Buffer.from(await response.arrayBuffer());
    console.log('Response body size:', buffer.length);
    console.log('ZIP signature:', buffer[0] === 0x50 && buffer[1] === 0x4B ? 'VALID' : 'INVALID');
    
    // Save the downloaded file
    writeFileSync('/tmp/test-api-export.pptx', buffer);
    
    // Try to parse it
    const { parsePptx } = await import('./src/lib/pptx-parser');
    const parseResult = await parsePptx(buffer, 'test.pptx');
    console.log('Parse result: slides =', parseResult.slideCount);
    
    // Check if modification was applied
    const el = parseResult.slides[0]?.elements.find(e => e.elementIndex === 6);
    if (el && el.type === 'text') {
      console.log('Element 6 text:', el.originalText.substring(0, 80));
      console.log('Modification applied:', el.originalText.includes('API_TEST_MODIFICATION'));
    }
    
    // Try LibreOffice validation
    const { execFile } = await import('child_process');
    const { promisify } = await import('util');
    const execFileAsync = promisify(execFile);
    
    try {
      await execFileAsync('libreoffice', ['--headless', '--norestore', '--convert-to', 'pdf', '--outdir', '/tmp', '/tmp/test-api-export.pptx'], { timeout: 60000 });
      console.log('LibreOffice conversion: SUCCESS');
    } catch (e: any) {
      console.log('LibreOffice conversion: FAILED:', e.message?.substring(0, 200));
    }
    
  } catch (e: any) {
    console.error('API test FAILED:', e.message);
  }
}

testApiExport();
