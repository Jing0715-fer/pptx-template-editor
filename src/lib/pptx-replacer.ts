import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';

// ============================================================================
// Types
// ============================================================================

export interface PptxModification {
  slideIndex: number;
  elementIndex: number;
  type: 'text' | 'table' | 'image';
  newText?: string;
  tableCells?: { row: number; col: number; text: string }[];
  newImageData?: string;
  newImageType?: string;
}

export interface ImageModification {
  slideIndex: number;
  imageRid: string;
  newImageData: string; // base64 or data URL
  newImageType: string; // 'png', 'jpeg', etc.
}

// ============================================================================
// XML Tree Navigation Helpers (for preserveOrder mode)
// ============================================================================

type XmlNode = Record<string, unknown>;
type Children = XmlNode[];

const CONTENT_ELEMENT_TAGS = new Set([
  'p:sp',
  'p:grpSp',
  'p:graphicFrame',
  'p:pic',
  'p:cxnSp',
]);

function findChild(parent: Children, tagName: string): Children | null {
  for (const item of parent) {
    if (item[tagName] !== undefined && Array.isArray(item[tagName])) {
      return item[tagName] as Children;
    }
  }
  return null;
}

function getContentElements(children: Children): { tagName: string; content: Children }[] {
  const result: { tagName: string; content: Children }[] = [];
  for (const item of children) {
    for (const key of Object.keys(item)) {
      if (key === ':@' || key === '#text') continue;
      if (Array.isArray(item[key])) {
        result.push({ tagName: key, content: item[key] as Children });
      }
    }
  }
  return result;
}

// ============================================================================
// XML String Manipulation Helpers
// ============================================================================

function escapeXmlText(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function findElementEnd(xml: string, start: number, tagName: string): number {
  let pos = start + 1;
  let inQuote = false;
  let quoteChar = '';

  while (pos < xml.length) {
    const ch = xml[pos];
    if (inQuote) {
      if (ch === quoteChar) inQuote = false;
    } else if (ch === '"' || ch === "'") {
      inQuote = true;
      quoteChar = ch;
    } else if (ch === '>') {
      if (xml[pos - 1] === '/') return pos + 1;
      pos++;
      break;
    }
    pos++;
  }

  let depth = 1;
  let iterations = 0;
  const maxIterations = 100000;
  while (depth > 0 && pos < xml.length && iterations < maxIterations) {
    iterations++;
    const nextTagPos = xml.indexOf('<', pos);
    if (nextTagPos === -1) break;

    if (xml[nextTagPos + 1] === '/') {
      const closeTag = `</${tagName}>`;
      if (xml.substring(nextTagPos, nextTagPos + closeTag.length) === closeTag) {
        depth--;
        if (depth === 0) return nextTagPos + closeTag.length;
      }
      pos = nextTagPos + 1;
    } else if (xml[nextTagPos + 1] === '!' || xml[nextTagPos + 1] === '?') {
      pos = nextTagPos + 2;
      const endDelim = xml[nextTagPos + 1] === '!' ? '-->' : '?>';
      const endPos = xml.indexOf(endDelim, pos);
      if (endPos === -1) break;
      pos = endPos + endDelim.length;
    } else {
      const tagMatch = xml.substring(nextTagPos).match(/^<([a-zA-Z][a-zA-Z0-9:]*)/);
      if (tagMatch && tagMatch[1] === tagName) depth++;
      const tagEndPos = xml.indexOf('>', nextTagPos);
      if (tagEndPos !== -1) {
        if (xml[tagEndPos - 1] === '/' && tagMatch && tagMatch[1] === tagName) {
          depth--;
        }
        pos = tagEndPos + 1;
      } else {
        pos = nextTagPos + 1;
      }
    }
  }

  return -1;
}

function findNthContentElementInString(
  xml: string,
  elementIndex: number,
): { tag: string; start: number; end: number } | null {
  const spTreeOpenMatch = xml.match(/<p:spTree[^>]*>/);
  if (!spTreeOpenMatch || spTreeOpenMatch.index === undefined) return null;

  const spTreeContentStart = spTreeOpenMatch.index + spTreeOpenMatch[0].length;
  let count = 0;
  let pos = spTreeContentStart;

  while (pos < xml.length) {
    const ltPos = xml.indexOf('<', pos);
    if (ltPos === -1) break;
    if (xml.substring(ltPos, ltPos + 11) === '</p:spTree>') break;
    if (xml[ltPos + 1] === '/') { pos = ltPos + 1; continue; }
    if (xml[ltPos + 1] === '!' || xml[ltPos + 1] === '?') { pos = ltPos + 2; continue; }

    const tagMatch = xml.substring(ltPos).match(/^<([a-zA-Z][a-zA-Z0-9:]*)/);
    if (!tagMatch) { pos = ltPos + 1; continue; }

    const tagName = tagMatch[1];

    if (CONTENT_ELEMENT_TAGS.has(tagName)) {
      const elementEnd = findElementEnd(xml, ltPos, tagName);
      if (elementEnd === -1) { pos = ltPos + 1; continue; }
      if (count === elementIndex) return { tag: tagName, start: ltPos, end: elementEnd };
      count++;
      pos = elementEnd;
    } else {
      const elementEnd = findElementEnd(xml, ltPos, tagName);
      pos = elementEnd !== -1 ? elementEnd : ltPos + 1;
    }
  }

  return null;
}

function findATagsInRegion(xml: string, regionStart: number, regionEnd: number): {
  start: number; end: number; textStart: number; textEnd: number;
}[] {
  const results: { start: number; end: number; textStart: number; textEnd: number }[] = [];
  let pos = regionStart;

  while (pos < regionEnd) {
    const aTagPos = xml.indexOf('<a:t', pos);
    if (aTagPos === -1 || aTagPos >= regionEnd) break;

    const afterName = xml[aTagPos + 4];
    if (afterName !== '>' && afterName !== ' ' && afterName !== '/') {
      pos = aTagPos + 1;
      continue;
    }

    const openTagEnd = xml.indexOf('>', aTagPos);
    if (openTagEnd === -1 || openTagEnd >= regionEnd) break;

    if (xml[openTagEnd - 1] === '/') {
      results.push({ start: aTagPos, end: openTagEnd + 1, textStart: openTagEnd + 1, textEnd: openTagEnd + 1 });
      pos = openTagEnd + 1;
      continue;
    }

    const textStart = openTagEnd + 1;
    const closeTagPos = xml.indexOf('</a:t>', textStart);
    if (closeTagPos === -1 || closeTagPos >= regionEnd) break;

    results.push({ start: aTagPos, end: closeTagPos + 6, textStart, textEnd: closeTagPos });
    pos = closeTagPos + 6;
  }

  return results;
}

function replaceTextInRegion(
  xml: string, regionStart: number, regionEnd: number, newText: string,
): string {
  const newParagraphs = newText.split('\n');
  const paragraphRegions: { start: number; end: number }[] = [];
  let pos = regionStart;

  while (pos < regionEnd) {
    const pTagPos = xml.indexOf('<a:p', pos);
    if (pTagPos === -1 || pTagPos >= regionEnd) break;
    const afterName = xml[pTagPos + 4];
    if (afterName !== '>' && afterName !== ' ') { pos = pTagPos + 1; continue; }
    const pEnd = findElementEnd(xml, pTagPos, 'a:p');
    if (pEnd === -1 || pEnd > regionEnd) break;
    paragraphRegions.push({ start: pTagPos, end: pEnd });
    pos = pEnd;
  }

  if (paragraphRegions.length === 0) return xml;

  const replacements: { from: number; to: number; replacement: string }[] = [];

  for (let pi = 0; pi < paragraphRegions.length; pi++) {
    const pRegion = paragraphRegions[pi];
    const aTags = findATagsInRegion(xml, pRegion.start, pRegion.end);
    if (aTags.length === 0) continue;
    const paraText = pi < newParagraphs.length ? newParagraphs[pi] : '';

    for (let ti = 0; ti < aTags.length; ti++) {
      const tag = aTags[ti];
      const text = ti === aTags.length - 1 ? paraText : '';
      replacements.push({ from: tag.textStart, to: tag.textEnd, replacement: escapeXmlText(text) });
    }
  }

  if (newParagraphs.length > paragraphRegions.length && paragraphRegions.length > 0) {
    const lastPRegion = paragraphRegions[paragraphRegions.length - 1];
    const lastATags = findATagsInRegion(xml, lastPRegion.start, lastPRegion.end);
    if (lastATags.length > 0) {
      const lastTag = lastATags[lastATags.length - 1];
      const extraText = newParagraphs.slice(paragraphRegions.length - 1).join('\n');
      const existingReplacement = replacements.find(
        (r) => r.from === lastTag.textStart && r.to === lastTag.textEnd
      );
      if (existingReplacement) existingReplacement.replacement = escapeXmlText(extraText);
    }
  }

  replacements.sort((a, b) => b.from - a.from);
  let result = xml;
  for (const { from, to, replacement } of replacements) {
    result = result.substring(0, from) + replacement + result.substring(to);
  }
  return result;
}

function replaceTableCellsInRegion(
  xml: string, regionStart: number, regionEnd: number,
  tableCells: { row: number; col: number; text: string }[],
): string {
  let tblStart = -1;
  let tblEnd = -1;
  let pos = regionStart;

  while (pos < regionEnd) {
    const tblPos = xml.indexOf('<a:tbl', pos);
    if (tblPos === -1 || tblPos >= regionEnd) break;
    const afterName = xml[tblPos + 6];
    if (afterName !== '>' && afterName !== ' ') { pos = tblPos + 1; continue; }
    const endPos = findElementEnd(xml, tblPos, 'a:tbl');
    if (endPos === -1 || endPos > regionEnd) break;
    tblStart = tblPos;
    tblEnd = endPos;
    break;
  }

  if (tblStart === -1) return xml;

  const rowRegions: { start: number; end: number }[] = [];
  pos = tblStart;
  while (pos < tblEnd) {
    const trPos = xml.indexOf('<a:tr', pos);
    if (trPos === -1 || trPos >= tblEnd) break;
    const afterName = xml[trPos + 5];
    if (afterName !== '>' && afterName !== ' ') { pos = trPos + 1; continue; }
    const trEnd = findElementEnd(xml, trPos, 'a:tr');
    if (trEnd === -1 || trEnd > tblEnd) break;
    rowRegions.push({ start: trPos, end: trEnd });
    pos = trEnd;
  }

  const replacements: { from: number; to: number; replacement: string }[] = [];

  for (const cellMod of tableCells) {
    const { row, col, text } = cellMod;
    if (row >= rowRegions.length) continue;
    const trRegion = rowRegions[row];
    const cellRegions: { start: number; end: number }[] = [];
    let cellPos = trRegion.start;

    while (cellPos < trRegion.end) {
      const tcPos = xml.indexOf('<a:tc', cellPos);
      if (tcPos === -1 || tcPos >= trRegion.end) break;
      const afterName = xml[tcPos + 5];
      if (afterName !== '>' && afterName !== ' ') { cellPos = tcPos + 1; continue; }
      const tcEnd = findElementEnd(xml, tcPos, 'a:tc');
      if (tcEnd === -1 || tcEnd > trRegion.end) break;
      cellRegions.push({ start: tcPos, end: tcEnd });
      cellPos = tcEnd;
    }

    if (col >= cellRegions.length) continue;
    const tcRegion = cellRegions[col];
    const aTags = findATagsInRegion(xml, tcRegion.start, tcRegion.end);

    for (let ti = 0; ti < aTags.length; ti++) {
      const tag = aTags[ti];
      const tagText = ti === aTags.length - 1 ? text : '';
      replacements.push({ from: tag.textStart, to: tag.textEnd, replacement: escapeXmlText(tagText) });
    }
  }

  replacements.sort((a, b) => b.from - a.from);
  let result = xml;
  for (const { from, to, replacement } of replacements) {
    result = result.substring(0, from) + replacement + result.substring(to);
  }
  return result;
}

// ============================================================================
// Image Replacement Helpers
// ============================================================================

function parseSlideRels(relsXml: string): Record<string, string> {
  const parser = new XMLParser({ ignoreAttributes: false, attributeNamePrefix: '@_' });
  const result: Record<string, string> = {};
  try {
    const parsed = parser.parse(relsXml);
    const relationships = parsed?.Relationships?.Relationship;
    if (relationships) {
      const relsArray = Array.isArray(relationships) ? relationships : [relationships];
      for (const rel of relsArray) {
        if (rel['@_Id'] && rel['@_Target']) result[rel['@_Id']] = rel['@_Target'];
      }
    }
  } catch { /* ignore */ }
  return result;
}

// ============================================================================
// Element Index Mapping (handle group shapes correctly)
// ============================================================================

function buildElementIndexMapping(parsed: Children): Map<number, number> {
  const mapping = new Map<number, number>();
  const sldChildren = findChild(parsed, 'p:sld');
  if (!sldChildren) return mapping;
  const cSldChildren = findChild(sldChildren, 'p:cSld');
  if (!cSldChildren) return mapping;
  const spTreeChildren = findChild(cSldChildren, 'p:spTree');
  if (!spTreeChildren) return mapping;

  const contentElements = getContentElements(spTreeChildren);
  let elementIndex = 0;
  let spTreeIndex = 0;

  for (const { tagName, content } of contentElements) {
    if (!CONTENT_ELEMENT_TAGS.has(tagName)) continue;
    if (tagName === 'p:grpSp') {
      const groupContentElements = getContentElements(content);
      for (const { tagName: childTag } of groupContentElements) {
        if (childTag === 'p:sp' || CONTENT_ELEMENT_TAGS.has(childTag)) {
          mapping.set(elementIndex, spTreeIndex);
          elementIndex++;
        }
      }
      spTreeIndex++;
    } else {
      mapping.set(elementIndex, spTreeIndex);
      elementIndex++;
      spTreeIndex++;
    }
  }

  return mapping;
}

function getSpTreeIndexForElement(parsed: Children, elementIndex: number): number {
  const mapping = buildElementIndexMapping(parsed);
  const spTreeIndex = mapping.get(elementIndex);
  return spTreeIndex !== undefined ? spTreeIndex : elementIndex;
}

function getGroupChildInfo(
  parsed: Children, elementIndex: number,
): { groupSpTreeIndex: number; childIndexInGroup: number } | null {
  const sldChildren = findChild(parsed, 'p:sld');
  if (!sldChildren) return null;
  const cSldChildren = findChild(sldChildren, 'p:cSld');
  if (!cSldChildren) return null;
  const spTreeChildren = findChild(cSldChildren, 'p:spTree');
  if (!spTreeChildren) return null;

  const contentElements = getContentElements(spTreeChildren);
  let currentElementIndex = 0;
  let spTreeIndex = 0;

  for (const { tagName, content } of contentElements) {
    if (!CONTENT_ELEMENT_TAGS.has(tagName)) continue;
    if (tagName === 'p:grpSp') {
      const groupContentElements = getContentElements(content);
      let childIdx = 0;
      for (const { tagName: childTag } of groupContentElements) {
        if (childTag === 'p:sp' || CONTENT_ELEMENT_TAGS.has(childTag)) {
          if (currentElementIndex === elementIndex) {
            return { groupSpTreeIndex: spTreeIndex, childIndexInGroup: childIdx };
          }
          currentElementIndex++;
          childIdx++;
        }
      }
      spTreeIndex++;
    } else {
      if (currentElementIndex === elementIndex) return null;
      currentElementIndex++;
      spTreeIndex++;
    }
  }

  return null;
}

function findNthContentElementInGroup(
  xml: string, groupStart: number, groupEnd: number, childIndex: number,
): { tag: string; start: number; end: number } | null {
  let count = 0;
  const groupOpenEnd = xml.indexOf('>', groupStart);
  if (groupOpenEnd === -1) return null;
  const contentStart = groupOpenEnd + 1;
  let pos = contentStart;

  while (pos < groupEnd) {
    const ltPos = xml.indexOf('<', pos);
    if (ltPos === -1 || ltPos >= groupEnd) break;
    if (xml[ltPos + 1] === '/') { pos = ltPos + 1; continue; }
    if (xml[ltPos + 1] === '!' || xml[ltPos + 1] === '?') { pos = ltPos + 2; continue; }

    const tagMatch = xml.substring(ltPos).match(/^<([a-zA-Z][a-zA-Z0-9:]*)/);
    if (!tagMatch) { pos = ltPos + 1; continue; }
    const tagName = tagMatch[1];

    if (CONTENT_ELEMENT_TAGS.has(tagName)) {
      const elementEnd = findElementEnd(xml, ltPos, tagName);
      if (elementEnd === -1 || elementEnd > groupEnd) { pos = ltPos + 1; continue; }
      if (count === childIndex) return { tag: tagName, start: ltPos, end: elementEnd };
      count++;
      pos = elementEnd;
    } else {
      const elementEnd = findElementEnd(xml, ltPos, tagName);
      pos = elementEnd !== -1 && elementEnd <= groupEnd ? elementEnd : ltPos + 1;
    }
  }

  return null;
}

// ============================================================================
// BOM Handling
// ============================================================================

const UTF8_BOM = Buffer.from([0xEF, 0xBB, 0xBF]);

function hasUtf8Bom(data: Buffer): boolean {
  return data.length >= 3 && data[0] === 0xEF && data[1] === 0xBB && data[2] === 0xBF;
}

function bufferToString(data: Buffer): string {
  let start = 0;
  if (hasUtf8Bom(data)) start = 3;
  return data.subarray(start).toString('utf-8');
}

function stringToBuffer(str: string, withBom: boolean): Buffer {
  const strBuf = Buffer.from(str, 'utf-8');
  if (!withBom) return strBuf;
  return Buffer.concat([UTF8_BOM, strBuf]);
}

// ============================================================================
// ZIP Verification (using JSZip for consistency)
// ============================================================================

async function verifyPptxBuffer(buffer: Buffer): Promise<{ valid: boolean; details: string }> {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const entryNames = Object.keys(zip.files);
    if (entryNames.length === 0) return { valid: false, details: 'ZIP file has no entries' };

    if (!entryNames.includes('[Content_Types].xml')) return { valid: false, details: 'Missing [Content_Types].xml' };
    if (!entryNames.includes('_rels/.rels')) return { valid: false, details: 'Missing _rels/.rels' };
    if (!entryNames.some((n) => n.startsWith('ppt/slides/slide') && n.endsWith('.xml')))
      return { valid: false, details: 'No slide XML files found' };

    try {
      const ctData = await zip.file('[Content_Types].xml')?.async('string');
      if (!ctData || !ctData.includes('<Types') || !ctData.includes('</Types>'))
        return { valid: false, details: '[Content_Types].xml has invalid structure' };
    } catch {
      return { valid: false, details: 'Failed to read [Content_Types].xml' };
    }

    // Verify all slide XML files are valid XML
    const slideEntries = entryNames.filter(n => n.startsWith('ppt/slides/slide') && n.endsWith('.xml'));
    let invalidSlides = 0;
    for (const slideName of slideEntries) {
      try {
        const slideData = await zip.file(slideName)?.async('string');
        if (!slideData || !slideData.includes('<p:sld')) {
          invalidSlides++;
        }
      } catch {
        invalidSlides++;
      }
    }
    if (invalidSlides > 0) {
      return { valid: false, details: `${invalidSlides} slides have invalid XML` };
    }

    return {
      valid: true,
      details: `Valid PPTX with ${entryNames.length} entries, ${slideEntries.length} slides`,
    };
  } catch (err) {
    return { valid: false, details: `ZIP parse error: ${err instanceof Error ? err.message : 'unknown'}` };
  }
}

// ============================================================================
// Main Export Function
//
// Key design decisions for Office compatibility:
// 1. Per-file compression: XML files use DEFLATE, binary files use STORE
// 2. All modifications applied before writing back to ZIP
// 3. BOM preservation for XML files
// 4. Round-trip verification after generation
// ============================================================================

export async function applyModificationsAndExport(
  pptxBuffer: Buffer,
  modifications: PptxModification[],
  imageModifications?: ImageModification[],
): Promise<Buffer> {
  const zip = await JSZip.loadAsync(pptxBuffer);

  // Track which files we've modified (for per-file compression)
  const modifiedFiles = new Set<string>();

  // Step 1: Apply text/table modifications
  const textTableMods = modifications.filter((m) => m.type === 'text' || m.type === 'table');
  const modsBySlide = new Map<number, PptxModification[]>();
  for (const mod of textTableMods) {
    if (!modsBySlide.has(mod.slideIndex)) modsBySlide.set(mod.slideIndex, []);
    modsBySlide.get(mod.slideIndex)!.push(mod);
  }

  for (const [slideIndex, slideMods] of Array.from(modsBySlide.entries())) {
    const slidePath = `ppt/slides/slide${slideIndex + 1}.xml`;
    const slideFile = zip.file(slidePath);
    if (!slideFile) continue;

    const rawXmlData = Buffer.from(await slideFile.async('uint8array'));
    const bomPresent = hasUtf8Bom(rawXmlData);
    let slideXml = bufferToString(rawXmlData);

    const parser = new XMLParser({
      preserveOrder: true,
      ignoreAttributes: false,
      attributeNamePrefix: '@_',
      processEntities: true,
      htmlEntities: true,
    });

    let parsed: Children;
    try {
      parsed = parser.parse(slideXml) as Children;
    } catch {
      console.error(`Failed to parse slide ${slideIndex + 1} for structure analysis`);
      continue;
    }

    // CRITICAL FIX: Collect ALL modifications first, then apply them in reverse position order
    // This prevents position offsets from being invalidated by earlier replacements
    const pendingReplacements: {
      regionStart: number;
      regionEnd: number;
      mod: PptxModification;
      groupChildInfo: { groupSpTreeIndex: number; childIndexInGroup: number } | null;
    }[] = [];

    for (const mod of slideMods) {
      try {
        const groupChildInfo = getGroupChildInfo(parsed, mod.elementIndex);
        pendingReplacements.push({
          regionStart: -1,
          regionEnd: -1,
          mod,
          groupChildInfo,
        });
      } catch (err) {
        console.error(`Error preparing modification for element ${mod.elementIndex} on slide ${slideIndex + 1}:`, err);
      }
    }

    // Find all element regions in the CURRENT xml (before any modifications)
    for (const pending of pendingReplacements) {
      try {
        let elementRegion: { tag: string; start: number; end: number } | null;

        if (pending.groupChildInfo) {
          const groupRegion = findNthContentElementInString(slideXml, pending.groupChildInfo.groupSpTreeIndex);
          if (!groupRegion) continue;
          elementRegion = findNthContentElementInGroup(slideXml, groupRegion.start, groupRegion.end, pending.groupChildInfo.childIndexInGroup);
        } else {
          const spTreeIndex = getSpTreeIndexForElement(parsed, pending.mod.elementIndex);
          elementRegion = findNthContentElementInString(slideXml, spTreeIndex);
        }

        if (!elementRegion) continue;
        pending.regionStart = elementRegion.start;
        pending.regionEnd = elementRegion.end;
      } catch (err) {
        console.error(`Error finding region for element ${pending.mod.elementIndex} on slide ${slideIndex + 1}:`, err);
      }
    }

    // Sort by region start position in DESCENDING order (apply from end to start)
    // This ensures earlier replacements don't shift positions of later ones
    const validReplacements = pendingReplacements.filter(r => r.regionStart >= 0);
    validReplacements.sort((a, b) => b.regionStart - a.regionStart);

    // Apply modifications from end to start
    for (const pending of validReplacements) {
      try {
        if (pending.mod.type === 'text' && pending.mod.newText !== undefined) {
          slideXml = replaceTextInRegion(slideXml, pending.regionStart, pending.regionEnd, pending.mod.newText);
        } else if (pending.mod.type === 'table' && pending.mod.tableCells) {
          slideXml = replaceTableCellsInRegion(slideXml, pending.regionStart, pending.regionEnd, pending.mod.tableCells);
        }
      } catch (err) {
        console.error(`Error applying modification for element ${pending.mod.elementIndex} on slide ${slideIndex + 1}:`, err);
      }
    }

    const outputData = stringToBuffer(slideXml, bomPresent);
    // Use DEFLATE for XML files (they compress well)
    zip.file(slidePath, outputData, { compression: 'DEFLATE', compressionOptions: { level: 6 } });
    modifiedFiles.add(slidePath);
  }

  // Step 2: Apply image modifications
  if (imageModifications && imageModifications.length > 0) {
    for (const imgMod of imageModifications) {
      try {
        const slideFileName = `slide${imgMod.slideIndex + 1}`;
        const relsPathWithXml = `ppt/slides/_rels/${slideFileName}.xml.rels`;
        const relsPathWithoutXml = `ppt/slides/_rels/${slideFileName}.rels`;

        let relsPath = relsPathWithXml;
        let relsFile = zip.file(relsPathWithXml);
        if (!relsFile) {
          relsPath = relsPathWithoutXml;
          relsFile = zip.file(relsPathWithoutXml);
        }

        if (!relsFile) {
          console.error(`Rels file not found for slide ${imgMod.slideIndex + 1}`);
          continue;
        }

        const relsXml = await relsFile.async('string');
        const relsMap = parseSlideRels(relsXml);
        const originalTarget = relsMap[imgMod.imageRid];

        if (!originalTarget) {
          console.error(`Image rId ${imgMod.imageRid} not found in rels`);
          continue;
        }

        let base64Data = imgMod.newImageData;
        if (base64Data.includes(',')) base64Data = base64Data.split(',')[1];
        const imageBuffer = Buffer.from(base64Data, 'base64');

        if (imageBuffer.length === 0) {
          console.error(`Decoded image buffer is empty for rId ${imgMod.imageRid}`);
          continue;
        }

        const mediaPath = originalTarget.startsWith('../')
          ? `ppt/${originalTarget.substring(3)}`
          : originalTarget.startsWith('/')
          ? originalTarget.substring(1)
          : `ppt/slides/${originalTarget}`;

        const originalExt = originalTarget.split('.').pop()?.toLowerCase() || '';
        const newExt = imgMod.newImageType.toLowerCase() === 'jpg' ? 'jpeg' : imgMod.newImageType.toLowerCase();

        if (originalExt !== newExt) {
          const newTarget = originalTarget.replace(/\.[^.]+$/, `.${newExt}`);
          const escapedRid = imgMod.imageRid.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const elementRegex = new RegExp(`<Relationship\\s[^>]*Id="${escapedRid}"[^>]*/?\\s*>`);
          const elementMatch = relsXml.match(elementRegex);

          let updatedRelsXml = relsXml;
          if (elementMatch) {
            const updatedElement = elementMatch[0].replace(/Target="[^"]*"/, `Target="${newTarget}"`);
            updatedRelsXml = relsXml.replace(elementMatch[0], updatedElement);
          }
          zip.file(relsPath, Buffer.from(updatedRelsXml, 'utf-8'), { compression: 'DEFLATE', compressionOptions: { level: 6 } });
          modifiedFiles.add(relsPath);

          // Update Content_Types.xml if the new extension is not registered
          const contentTypesFile = zip.file('[Content_Types].xml');
          if (contentTypesFile) {
            const contentTypesXml = await contentTypesFile.async('string');
            if (!contentTypesXml.includes(`Extension="${newExt}"`)) {
              const mimeType = `image/${newExt}`;
              const newContentType = `<Default Extension="${newExt}" ContentType="${mimeType}"/>`;
              const updatedContentTypes = contentTypesXml.replace('</Types>', `${newContentType}</Types>`);
              zip.file('[Content_Types].xml', Buffer.from(updatedContentTypes, 'utf-8'), { compression: 'DEFLATE', compressionOptions: { level: 6 } });
              modifiedFiles.add('[Content_Types].xml');
            }
          }

          const newMediaPath = newTarget.startsWith('../')
            ? `ppt/${newTarget.substring(3)}`
            : newTarget.startsWith('/')
            ? newTarget.substring(1)
            : `ppt/slides/${newTarget}`;

          // Remove old media file and add new one
          zip.remove(mediaPath);
          // Use STORE for binary image files (they're already compressed)
          zip.file(newMediaPath, imageBuffer, { compression: 'STORE' });
          modifiedFiles.add(newMediaPath);
        } else {
          // Use STORE for binary image files (they're already compressed)
          zip.file(mediaPath, imageBuffer, { compression: 'STORE' });
          modifiedFiles.add(mediaPath);
        }
      } catch (err) {
        console.error(`Error replacing image rId ${imgMod.imageRid} on slide ${imgMod.slideIndex + 1}:`, err);
      }
    }
  }

  // Step 3: Ensure unmodified XML files keep DEFLATE compression
  // JSZip's generateAsync defaults to STORE, but PPTX files should use
  // DEFLATE for XML content. We need to re-compress unmodified XML files.
  const xmlExtensions = ['.xml', '.rels'];
  for (const [filePath, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    if (modifiedFiles.has(filePath)) continue; // Already handled

    const isXml = xmlExtensions.some(ext => filePath.toLowerCase().endsWith(ext));
    if (isXml) {
      // Re-add the file with DEFLATE compression
      const data = await file.async('uint8array');
      zip.file(filePath, data, { compression: 'DEFLATE', compressionOptions: { level: 6 } });
    }
    // Non-XML, non-modified files keep their original compression (STORE by default)
  }

  // Generate output buffer without global compression (per-file settings will apply)
  const outputBuffer = await zip.generateAsync({
    type: 'nodebuffer',
    // Don't set global compression - we've set per-file compression above
    // This defaults to STORE for files without explicit compression,
    // but our XML files and modified files have explicit compression settings
  });

  const verification = await verifyPptxBuffer(outputBuffer);
  if (!verification.valid) {
    console.error('Export verification FAILED:', verification.details);
    throw new Error(`导出文件验证失败: ${verification.details}`);
  }

  // Additional round-trip check: verify the output can be parsed back
  try {
    const { parsePptx } = await import('./pptx-parser');
    const testParse = await parsePptx(outputBuffer, 'verify.pptx');
    if (testParse.slides.length === 0) {
      throw new Error('导出文件无法解析出任何幻灯片');
    }
  } catch (err) {
    console.error('Round-trip parse verification FAILED:', err);
    throw new Error(`导出文件验证失败(重新解析): ${err instanceof Error ? err.message : 'unknown'}`);
  }

  console.log('Export verification passed:', verification.details);
  return outputBuffer;
}

// ============================================================================
// Round-trip test
// ============================================================================

export async function testExportRoundTrip(
  originalBuffer: Buffer,
  modifiedBuffer: Buffer,
): Promise<{ success: boolean; details: string }> {
  try {
    const verification = await verifyPptxBuffer(modifiedBuffer);
    if (!verification.valid) return { success: false, details: `验证失败: ${verification.details}` };

    const zip = await JSZip.loadAsync(modifiedBuffer);
    const entryNames = Object.keys(zip.files);

    const originalZip = await JSZip.loadAsync(originalBuffer);
    const originalSlideEntries = Object.keys(originalZip.files).filter(
      (n) => n.startsWith('ppt/slides/slide') && n.endsWith('.xml')
    );

    let allSlidesPresent = true;
    for (const slideName of originalSlideEntries) {
      if (!entryNames.includes(slideName)) { allSlidesPresent = false; break; }
    }
    if (!allSlidesPresent) return { success: false, details: '部分幻灯片在导出文件中缺失' };

    const mediaEntries = entryNames.filter((n) => n.startsWith('ppt/media/') && !zip.files[n].dir);
    let emptyImages = 0;
    for (const mediaName of mediaEntries) {
      const data = await zip.files[mediaName].async('uint8array');
      if (data.length === 0) emptyImages++;
    }
    if (emptyImages > 0) return { success: false, details: `${emptyImages} 个图片文件为空` };

    // Try to parse with our parser
    try {
      const { parsePptx } = await import('./pptx-parser');
      const parseResult = await parsePptx(modifiedBuffer, 'roundtrip.pptx');
      if (parseResult.slides.length === 0) {
        return { success: false, details: '重新解析后没有找到幻灯片' };
      }
    } catch (err) {
      return { success: false, details: `重新解析失败: ${err instanceof Error ? err.message : 'unknown'}` };
    }

    return {
      success: true,
      details: `验证通过: ${entryNames.length} 个文件, ${originalSlideEntries.length} 页幻灯片完整, ${mediaEntries.length} 个媒体文件`,
    };
  } catch (err) {
    return { success: false, details: `重新解析失败: ${err instanceof Error ? err.message : 'unknown'}` };
  }
}

// Legacy wrapper
export async function replaceImagesInPptx(
  pptxBuffer: Buffer,
  imageModifications: ImageModification[],
): Promise<Buffer> {
  return applyModificationsAndExport(pptxBuffer, [], imageModifications);
}
