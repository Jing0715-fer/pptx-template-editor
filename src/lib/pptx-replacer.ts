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
  // Step 1: Strip XML-forbidden control characters (U+0000-U+0008, U+000B-U+000C,
  // U+000E-U+001F). These cannot appear in well-formed XML 1.0 character data.
  // PowerPoint's strict parser will fail or silently truncate text containing them,
  // which is a common cause of "blank slide after text replacement" bugs.
  // U+0009 (tab), U+000A (LF), U+000D (CR) are allowed and preserved.
  // U+FFFE / U+FFFF (non-characters) are also stripped for safety.
  // eslint-disable-next-line no-control-regex
  let result = text.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, '');

  // Step 2: Escape the three XML-reserved characters.
  // Note: we do NOT escape quotes (") or apostrophes (') because <a:t> text
  // content is character data, not an attribute value. Escaping them would
  // produce literal &quot; / &apos; in the rendered slide text, which looks
  // broken to end users.
  // We also normalize the three "safe" line terminators that PowerPoint
  // sometimes writes: vertical tab (U+000B) and form feed (U+000C) are
  // already stripped above; we additionally collapse NEL (U+0085) to LF.
  result = result.replace(/\u0085/g, '\n');

  return result
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
// Content_Types.xml Sanitization
// ============================================================================

// Proper MIME type mapping for image extensions (OPC spec compliant)
const IMAGE_MIME_MAP: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  svg: 'image/svg+xml',
  tiff: 'image/tiff',
  tif: 'image/tiff',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  webp: 'image/webp',
  ico: 'image/x-icon',
};

/**
 * Sanitizes [Content_Types].xml to fix invalid MIME types and ensure Office compatibility.
 * Fixes issues found in WPS-created PPTX files:
 * - "image/.jpg" (extra dot) → "image/jpeg"
 * - "image/image/png" (doubled prefix) → "image/png"
 * - Case-insensitive MIME types → lowercase
 * - Incorrect MIME type for jpg → "image/jpeg" (not "image/jpg")
 * - Uppercase extensions like Extension="JPG" → Extension="jpg"
 * - Extension containing slashes like Extension="image/png" → remove (invalid)
 * - **Duplicate <Default> ContentType** (OPC spec violation: two <Default> entries
 *   pointing to the same ContentType is illegal and causes Windows Office to
 *   trigger a "needs repair" dialog, which in turn may lose parts of the
 *   presentation — the root cause of the "修复后变空白" bug.)
 */
function sanitizeContentTypesXml(xml: string): string {
  let result = xml;

  // 0. Fix doubled MIME prefix: ContentType="image/image/png" → ContentType="image/png"
  // WPS bug: writes ContentType="image/image/png" with Extension="image/png"
  result = result.replace(
    /ContentType="image\/image\/([a-zA-Z+]+)"/g,
    (_match: string, sub: string) => {
      const corrected = IMAGE_MIME_MAP[sub.toLowerCase()] || `image/${sub.toLowerCase()}`;
      return `ContentType="${corrected}"`;
    }
  );

  // 1. Fix patterns like "image/.jpg", "image/.png", etc. (extra dot before extension)
  // This is the PRIMARY bug: WPS writes ContentType="image/.jpg" instead of "image/jpeg"
  result = result.replace(
    /ContentType="image\/\.([a-zA-Z+]+)"/g,
    (_match: string, ext: string) => {
      const corrected = IMAGE_MIME_MAP[ext.toLowerCase()] || `image/${ext.toLowerCase()}`;
      return `ContentType="${corrected}"`;
    }
  );

  // 2. Fix ContentType="image/jpg" → ContentType="image/jpeg"
  result = result.replace(
    /ContentType="image\/jpg"/gi,
    'ContentType="image/jpeg"'
  );

  // 3. Fix case-insensitive MIME types (e.g., "Image/JPEG" → "image/jpeg")
  result = result.replace(
    /ContentType="Image\/JPEG"/gi,
    'ContentType="image/jpeg"'
  );
  result = result.replace(
    /ContentType="Image\/PNG"/gi,
    'ContentType="image/png"'
  );
  result = result.replace(
    /ContentType="Image\/GIF"/gi,
    'ContentType="image/gif"'
  );
  result = result.replace(
    /ContentType="Image\/BMP"/gi,
    'ContentType="image/bmp"'
  );
  result = result.replace(
    /ContentType="Image\/TIFF"/gi,
    'ContentType="image/tiff"'
  );
  result = result.replace(
    /ContentType="Image\/SVG\+XML"/gi,
    'ContentType="image/svg+xml"'
  );
  result = result.replace(
    /ContentType="Image\/X-EMF"/gi,
    'ContentType="image/x-emf"'
  );
  result = result.replace(
    /ContentType="Image\/X-WMF"/gi,
    'ContentType="image/x-wmf"'
  );
  result = result.replace(
    /ContentType="Image\/WEBP"/gi,
    'ContentType="image/webp"'
  );

  // 4. Normalize uppercase extensions to lowercase (e.g., Extension="JPG" → Extension="jpg")
  // and ensure correct MIME type mapping
  result = result.replace(
    /<Default\s+Extension="([A-Z]{2,})"\s+ContentType="([^"]+)"/g,
    (_match: string, ext: string, contentType: string) => {
      const lowerExt = ext.toLowerCase();
      const correctMime = IMAGE_MIME_MAP[lowerExt] || contentType;
      return `<Default Extension="${lowerExt}" ContentType="${correctMime}"`;
    }
  );

  // Also handle Extension and ContentType in reverse attribute order
  result = result.replace(
    /<Default\s+ContentType="([^"]+)"\s+Extension="([A-Z]{2,})"/g,
    (_match: string, contentType: string, ext: string) => {
      const lowerExt = ext.toLowerCase();
      const correctMime = IMAGE_MIME_MAP[lowerExt] || contentType;
      return `<Default ContentType="${correctMime}" Extension="${lowerExt}"`;
    }
  );

  // 5. Remove <Default> entries with invalid Extension containing a slash
  // WPS bug: writes Extension="image/png" ContentType="image/png"
  // After step 0 above, ContentType is fixed but Extension="image/png" is still
  // invalid (OPC spec says Extension must be a simple file extension, no slashes).
  // This will also create a duplicate ContentType with the valid
  // <Default Extension="png" ContentType="image/png"/> entry, which the
  // deduplication in step 8 would handle — but it's safer to remove the
  // malformed entry entirely because dedup may not match the slash-containing
  // extension when rewriting <Override> PartName references.
  result = result.replace(
    /<Default\s+Extension="[^"]*\/[^"]*"\s+ContentType="[^"]*"[^/]*\/>/g,
    ''
  );
  // Also handle reversed attribute order: ContentType before Extension
  result = result.replace(
    /<Default\s+ContentType="[^"]*"\s+Extension="[^"]*\/[^"]*"[^/]*\/>/g,
    ''
  );

  // 5b. Remove any <Default> entries for ppt/tags/ if they exist
  // (WPS adds ContentType entries for its non-standard tag files)
  result = result.replace(
    /<Default\s+Extension="tag"\s+ContentType="[^"]*"[^/]*\/>/g,
    ''
  );

  // 6. Remove <Override> entries for WPS tag files (ppt/tags/tag*.xml)
  // These are non-standard and reference files that don't exist in the ZIP.
  // Windows Office strict OPC parser rejects the package when it encounters
  // Override entries for Parts that don't exist in the archive.
  // WPS typically creates 81 of these: tag1.xml through tag81.xml
  result = result.replace(
    /<Override\s+PartName="\/ppt\/tags\/tag\d+\.xml"\s+ContentType="[^"]*"[^/]*\/?>/g,
    ''
  );
  // Also handle attribute order variation: ContentType before PartName
  result = result.replace(
    /<Override\s+ContentType="[^"]*"\s+PartName="\/ppt\/tags\/tag\d+\.xml"[^/]*\/?>/g,
    ''
  );

  // 7. Remove any <Override> entries for ppt/tags/ directory itself
  result = result.replace(
    /<Override\s+PartName="\/ppt\/tags[^"]*"\s+ContentType="[^"]*"[^/]*\/?>/g,
    ''
  );
  result = result.replace(
    /<Override\s+ContentType="[^"]*"\s+PartName="\/ppt\/tags[^"]*"[^/]*\/?>/g,
    ''
  );

  // 8. CRITICAL FIX: Deduplicate <Default> entries that share the same ContentType.
  //
  // OPC spec (ECMA-376 Part 1, §11.3.2.2) says implementations SHOULD NOT have
  // multiple <Default> elements with the same ContentType. Windows Office's
  // strict OPC parser treats this as a malformed package: it shows the
  // "PowerPoint found a problem with content. PowerPoint can attempt to repair
  // the presentation." dialog. The auto-repair process then rebuilds the slide
  // list and may drop slides that reference images whose Default ContentType
  // mapping is "ambiguous" — resulting in blank pages.
  //
  // Real-world trigger: original file contains
  //   <Default Extension="jpeg" ContentType="image/jpeg"/>
  //   <Default Extension="JPG"  ContentType="image/.jpg"/>   ← WPS
  // After step 4 above, the second becomes:
  //   <Default Extension="jpg"  ContentType="image/jpeg"/>   ← dup ContentType
  // We keep the FIRST occurrence (lowest Extension sort order = canonical) and
  // remove the rest, rewriting affected <Override> entries that referenced the
  // dropped Extension so they point to the kept one.
  const beforeDedupe = result;
  result = dedupeDefaultContentTypes(result);
  if (result !== beforeDedupe) {
    console.log('[Content_Types].xml dedup: removed duplicate <Default> ContentType entries');
  }

  // 9. Clean up any blank lines left by removed entries
  result = result.replace(/\n\s*\n/g, '\n');

  return result;
}

/**
 * Removes duplicate <Default> entries that share the same ContentType.
 * Returns the rewritten XML.
 *
 * Strategy:
 *  1. Collect all <Default> entries (with their order).
 *  2. Group by ContentType; keep the FIRST entry per ContentType, drop the rest.
 *  3. Rewrite <Override> entries that referenced a dropped Extension so they
 *     use the kept Extension instead. This is safe because both Extensions map
 *     to the same ContentType, so the Override is functionally identical.
 *  4. Remove the dropped <Default> elements from the XML.
 */
function dedupeDefaultContentTypes(xml: string): string {
  // Match <Default Extension="X" ContentType="Y"/> or with reversed attribute order.
  // Allow optional whitespace around attributes.
  const defaultRegex = /<Default\s+(?:Extension="([^"]+)"\s+ContentType="([^"]+)"|ContentType="([^"]+)"\s+Extension="([^"]+)")\s*\/>/g;

  const allDefaults: { ext: string; contentType: string; raw: string; start: number; end: number }[] = [];
  let m: RegExpExecArray | null;
  while ((m = defaultRegex.exec(xml)) !== null) {
    const ext = m[1] ?? m[4];
    const contentType = m[2] ?? m[3];
    allDefaults.push({ ext, contentType, raw: m[0], start: m.index, end: m.index + m[0].length });
  }

  if (allDefaults.length === 0) return xml;

  // Group by ContentType, preserve first-seen order
  const seenContentTypes = new Map<string, typeof allDefaults[number]>();
  const droppedExts = new Map<string, string>(); // droppedExt -> keptExt (for Override rewrite)

  for (const d of allDefaults) {
    const existing = seenContentTypes.get(d.contentType);
    if (!existing) {
      seenContentTypes.set(d.contentType, d);
    } else {
      // Drop this one; record the mapping
      droppedExts.set(d.ext, existing.ext);
    }
  }

  if (droppedExts.size === 0) return xml;

  // Rewrite <Override> PartName references: ppt/media/foo.OLD_EXT -> ppt/media/foo.NEW_EXT
  // We only rewrite if the new extension is in our kept set; this preserves file naming.
  let result = xml;
  for (const [oldExt, newExt] of droppedExts) {
    if (oldExt === newExt) continue;
    // Match PartName="..." ending with .oldExt, and update the extension.
    // Use a safe pattern: only change the trailing extension, don't touch directories.
    const extRegex = new RegExp(
      `(PartName="[^"]+\\.)${oldExt.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}(")`,
      'g'
    );
    result = result.replace(extRegex, `$1${newExt}$2`);
  }

  // Remove the dropped <Default> entries (process from end to start to preserve offsets)
  const droppedRaw = allDefaults
    .filter(d => droppedExts.has(d.ext))
    .sort((a, b) => b.start - a.start);
  for (const d of droppedRaw) {
    result = result.substring(0, d.start) + result.substring(d.end);
  }

  return result;
}

/**
 * Removes WPS-specific files from the ZIP that may cause issues in Windows Office.
 * WPS adds non-standard files like ppt/tags/tag*.xml which are not part of
 * the OOXML specification and may cause Windows Office to reject the file.
 */
function removeWpsSpecificFiles(zip: JSZip): void {
  const entriesToRemove: string[] = [];

  for (const [filePath, file] of Object.entries(zip.files)) {
    // Remove WPS tag files: ppt/tags/tag*.xml
    if (filePath.startsWith('ppt/tags/')) {
      entriesToRemove.push(filePath);
    }
  }

  for (const entry of entriesToRemove) {
    zip.remove(entry);
  }

  if (entriesToRemove.length > 0) {
    console.log(`Removed ${entriesToRemove.length} WPS-specific files`);
  }
}

/**
 * Removes references to WPS tag files (../tags/tag*.xml) from ALL .rels files.
 *
 * CRITICAL for Windows Office compatibility:
 * WPS-created PPTX files declare relationships to ppt/tags/tag*.xml in their
 * .rels files (slideLayout, slideMaster, slide, etc.), but the tag files often
 * don't exist in the ZIP archive. Windows Office's strict OPC parser validates
 * every relationship target exists — if a referenced Part is missing, the entire
 * package is rejected. macOS Office is more lenient and silently ignores missing Parts.
 *
 * This function iterates ALL .rels files in the ZIP and removes <Relationship>
 * elements whose Target references ../tags/tag*.xml.
 */
async function removeTagReferencesFromRels(zip: JSZip, modifiedFiles?: Set<string>): Promise<void> {
  let totalRemoved = 0;

  for (const [filePath, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    // Process ALL .rels files in the entire ZIP
    if (!filePath.endsWith('.rels')) continue;

    try {
      const rawData = Buffer.from(await file.async('uint8array'));
      const bomPresent = hasUtf8Bom(rawData);
      let relsXml = bufferToString(rawData);

      // Check if this .rels file contains any tag references
      if (!relsXml.includes('tags/tag')) continue;

      // Remove <Relationship> elements referencing ../tags/tag*.xml
      // Pattern matches various attribute orders and whitespace
      const before = relsXml;

      // Handle: <Relationship Id="rId3" Type="..." Target="../tags/tag1.xml"/>
      relsXml = relsXml.replace(
        /<Relationship\s[^>]*Target="[^"]*tags\/tag\d+\.xml"[^>]*\/?>/g,
        ''
      );

      // Handle: Target before other attributes (less common but valid XML)
      // Actually the above regex already covers this since [^>]* matches any attributes before Target
      // and [^>]* matches any attributes after Target, as long as Target contains tags/tagN.xml

      // Clean up blank lines left by removed entries
      relsXml = relsXml.replace(/\n\s*\n/g, '\n');

      if (relsXml !== before) {
        const outputData = stringToBuffer(relsXml, bomPresent);
        zip.file(filePath, outputData, { compression: 'DEFLATE', compressionOptions: { level: 6 } });
        if (modifiedFiles) modifiedFiles.add(filePath);
        const removedCount = (before.match(/tags\/tag\d+\.xml/g) || []).length -
                            (relsXml.match(/tags\/tag\d+\.xml/g) || []).length;
        totalRemoved += removedCount;
        console.log(`Cleaned ${removedCount} tag references from ${filePath}`);
      }
    } catch {
      // Skip files that can't be processed
    }
  }

  if (totalRemoved > 0) {
    console.log(`Total: removed ${totalRemoved} WPS tag references from .rels files`);
  }
}

/**
 * Sanitizes WPS-specific XML attributes and elements from presentation files.
 * Removes non-standard namespaces, custom WPS properties, and WPS-only
 * elements that Windows Office's strict OPC parser may reject.
 *
 * Scans BOTH the always-processed core files AND every .xml / .rels file in
 * the package — because WPS often embeds its custom namespaces (wps, wpg, etc.)
 * inside slide-level XML (e.g. <a:extLst><asvg:svgBlip .../></a:extLst>), and
 * failing to clean those will still cause the "needs repair" dialog on Windows.
 */
async function sanitizeWpsXmlAttributes(zip: JSZip): Promise<void> {
  const alwaysSanitize = [
    'ppt/presentation.xml',
    'ppt/presProps.xml',
    'ppt/viewProps.xml',
    'docProps/core.xml',
    'docProps/app.xml',
  ];

  // Track files we touched (so we can skip redundant scans of slide XMLs
  // and to record changes for the round-trip verification)
  const seen = new Set<string>(alwaysSanitize);

  // Collect all XML / .rels files in the package
  const toScan: string[] = [...alwaysSanitize];
  for (const [filePath, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    if (seen.has(filePath)) continue;
    const lower = filePath.toLowerCase();
    if (lower.endsWith('.xml') || lower.endsWith('.rels')) {
      toScan.push(filePath);
    }
  }

  for (const filePath of toScan) {
    const file = zip.file(filePath);
    if (!file) continue;

    try {
      const rawData = Buffer.from(await file.async('uint8array'));
      const bomPresent = hasUtf8Bom(rawData);
      let content = bufferToString(rawData);

      const before = content;

      // ---- KSOProductBuildVer removal (WPS custom property) ----
      // The original regex `<property[^>]*KSOProductBuildVer[^>]*>[\s\S]*?<\/property>`
      // was unsafe because `[\s\S]*?` is non-greedy but would happily stop at the
      // FIRST `</property>` even if KSOProductBuildVer lives inside a parent
      // container. We now anchor on the property name with strict bounds and
      // allow the value element to vary (vt:lpstr / vt:lpwstr / vt:filetime / etc.)
      if (content.includes('KSOProductBuildVer')) {
        content = content.replace(
          /<property\b[^>]*\bname="KSOProductBuildVer"[^>]*>[\s\S]*?<\/property>/g,
          ''
        );
        // Fallback: name appears in any attribute order
        content = content.replace(
          /<property\b(?=[^>]*\bKSOProductBuildVer)[\s\S]*?<\/property>/g,
          (match) => (match.includes('KSOProductBuildVer') ? '' : match)
        );
      }

      // ---- KSO-related custom properties (broader cleanup) ----
      // WPS writes a whole set of "KSO*" properties (KSOProductBuildVer,
      // KSOWPSVersion, etc.). They are meaningless to Office and may confuse
      // strict parsers. Only remove well-known safe ones to avoid losing user
      // data.
      const SAFE_KSO_PROPS = ['KSOProductBuildVer', 'KSOWPSVersion', 'KSOGoBackRev'];
      for (const propName of SAFE_KSO_PROPS) {
        const re = new RegExp(
          `<property\\b[^>]*\\bname="${propName}"[^>]*>[\\s\\S]*?<\\/property>`,
          'g'
        );
        content = content.replace(re, '');
      }

      // ---- WPS/Kingsoft namespace declarations ----
      // These typically appear as xmlns:wps="http://www.wps.cn/...". They are
      // declared on root elements but ALSO inline on sub-elements when WPS
      // patches existing slides. Both forms are scrubbed.
      content = content.replace(
        /\s+xmlns:[a-zA-Z0-9]+="http[^"]*wps[^"]*"/gi,
        ''
      );
      content = content.replace(
        /\s+xmlns:[a-zA-Z0-9]+="http[^"]*kingsoft[^"]*"/gi,
        ''
      );

      // ---- WPS-prefixed XML elements ----
      // <wps:xxx ...>, <wpg:xxx ...>, etc. Strip them entirely (open + close).
      // Use a more careful regex that handles namespaces declared inline.
      content = content.replace(
        /<\/?(?:wps|wpg|wpc|wpi):[a-zA-Z][a-zA-Z0-9]*\b[^>]*>/g,
        ''
      );

      // ---- Inline xmlns declarations attached to WPS elements (cleanup) ----
      // After removing <wps:xxx> elements, their inline xmlns: declarations
      // may become orphaned. E.g. <foo xmlns:wps="..."> with no wps: children.
      // We don't aggressively remove these (they're harmless) but DO remove
      // them when on a single line / when they match the standard WPS URI.
      content = content.replace(
        /\s+xmlns:(?:wps|wpg|wpc|wpi)="[^"]*"/g,
        ''
      );

      // ---- Normalize excessive whitespace left by removals ----
      // Multiple consecutive blank lines → single newline
      content = content.replace(/\n\s*\n/g, '\n');

      if (content !== before) {
        const outputData = stringToBuffer(content, bomPresent);
        zip.file(filePath, outputData, { compression: 'DEFLATE', compressionOptions: { level: 6 } });
      }
    } catch {
      // Skip files that can't be processed
    }
  }
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

      // Check for invalid MIME types that Windows Office strictly rejects
      const invalidMimeMatch = ctData.match(/ContentType="image\/\.[a-zA-Z+]+"/);
      if (invalidMimeMatch) {
        return { valid: false, details: `[Content_Types].xml contains invalid MIME type: ${invalidMimeMatch[0]} (extra dot before extension)` };
      }

      // Check for doubled MIME prefix: ContentType="image/image/png"
      const doubledMimeMatch = ctData.match(/ContentType="image\/image\//);
      if (doubledMimeMatch) {
        return { valid: false, details: '[Content_Types].xml contains doubled MIME prefix (image/image/)' };
      }

      // Check for Extension containing slashes: Extension="image/png"
      const slashExtMatch = ctData.match(/Extension="[^"]*\/[^"]*"/);
      if (slashExtMatch) {
        return { valid: false, details: `[Content_Types].xml contains invalid Extension with slash: ${slashExtMatch[0]}` };
      }

      // Check for WPS tag Override entries that reference non-existent files
      const tagOverrideMatch = ctData.match(/PartName="\/ppt\/tags\/tag\d+\.xml"/);
      if (tagOverrideMatch) {
        return { valid: false, details: `[Content_Types].xml contains WPS tag Override for non-existent file: ${tagOverrideMatch[0]}` };
      }
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

    // Check for dangling .rels references to WPS tag files
    // ⚠️ NOT a fatal error: macOS PowerPoint's internal consistency check tolerates
    // these dangling refs as long as the .rels files themselves are present and
    // well-formed. We only WARN about them, not fail. See the comment in
    // applyModificationsAndExport near the disabled removeWpsSpecificFiles calls
    // for the full explanation.
    let danglingRelsCount = 0;
    for (const [filePath, file] of Object.entries(zip.files)) {
      if (file.dir || !filePath.endsWith('.rels')) continue;
      try {
        const relsData = await file.async('string');
        if (relsData.includes('tags/tag')) {
          const tagRefs = relsData.match(/Target="[^"]*tags\/tag\d+\.xml"/g);
          if (tagRefs) danglingRelsCount += tagRefs.length;
        }
      } catch { /* skip */ }
    }
    // Note: danglingRelsCount is reported in the details string but is no longer fatal.
    if (danglingRelsCount > 0) {
      console.warn(`Found ${danglingRelsCount} dangling .rels references to WPS tag files (kept for PowerPoint compatibility)`);
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

        // Extract the actual file extension from the target path.
        // Handle WPS non-standard nested paths like "../media/image5.image/png"
        // where the "extension" is the format name after the last slash inside
        // a directory named "*.image". For standard paths like "../media/image3.png",
        // the extension is simply the part after the last dot.
        const originalExt = (() => {
          // Check for WPS-style .image/ directory structure (e.g., "image5.image/png")
          const wpsImageMatch = originalTarget.match(/\.image\/([a-zA-Z0-9]+)$/);
          if (wpsImageMatch) return wpsImageMatch[1].toLowerCase();
          // Standard path: take the part after the last dot
          const lastDotIdx = originalTarget.lastIndexOf('.');
          if (lastDotIdx >= 0) return originalTarget.substring(lastDotIdx + 1).toLowerCase();
          return '';
        })();

        // Normalize newImageType: accept both MIME types ("image/png") and plain extensions ("png")
        const newExt = (() => {
          let ext = imgMod.newImageType.toLowerCase().trim();
          // Strip MIME type prefix: "image/png" → "png", "image/jpeg" → "jpeg"
          if (ext.includes('/')) {
            ext = ext.split('/').pop() || ext;
          }
          // Normalize jpg → jpeg
          if (ext === 'jpg') ext = 'jpeg';
          return ext;
        })();

        if (originalExt !== newExt) {
          // Compute the new target path. For WPS-style nested paths like
          // "../media/image5.image/png", produce a flat path like "../media/image5.jpeg"
          // instead of another nested path. For standard paths like "../media/image3.png",
          // just replace the extension.
          const newTarget = (() => {
            const wpsImageMatch = originalTarget.match(/^(.+)\.image\/[a-zA-Z0-9]+$/);
            if (wpsImageMatch) {
              // WPS nested path: replace ".image/png" with ".{newExt}"
              return `${wpsImageMatch[1]}.${newExt}`;
            }
            // Standard path: just replace the last extension
            return originalTarget.replace(/\.[^.]+$/, `.${newExt}`);
          })();
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
              const mimeType = IMAGE_MIME_MAP[newExt] || `image/${newExt}`;
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
          // Same extension — but check if the original is a WPS nested path like
          // "../media/image5.image/png". If so, normalize to a flat path for
          // Office compatibility. The .image/ directory structure is non-standard
          // and causes Content_Types mapping failures because the part name
          // "image5.image/png" doesn't match any <Default Extension="..."> entry.
          const wpsImageMatch = originalTarget.match(/^(.+)\.image\/[a-zA-Z0-9]+$/);
          if (wpsImageMatch) {
            const flatTarget = `${wpsImageMatch[1]}.${newExt}`;
            const flatMediaPath = flatTarget.startsWith('../')
              ? `ppt/${flatTarget.substring(3)}`
              : flatTarget.startsWith('/')
              ? flatTarget.substring(1)
              : `ppt/slides/${flatTarget}`;

            // Update the .rels target
            const escapedRid = imgMod.imageRid.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const elementRegex = new RegExp(`<Relationship\\s[^>]*Id="${escapedRid}"[^>]*/?\\s*>`);
            const elementMatch = relsXml.match(elementRegex);
            let updatedRelsXml = relsXml;
            if (elementMatch) {
              const updatedElement = elementMatch[0].replace(/Target="[^"]*"/, `Target="${flatTarget}"`);
              updatedRelsXml = relsXml.replace(elementMatch[0], updatedElement);
            }
            zip.file(relsPath, Buffer.from(updatedRelsXml, 'utf-8'), { compression: 'DEFLATE', compressionOptions: { level: 6 } });
            modifiedFiles.add(relsPath);

            // Remove old nested file, add new flat file
            zip.remove(mediaPath);
            zip.file(flatMediaPath, imageBuffer, { compression: 'STORE' });
            modifiedFiles.add(flatMediaPath);
          } else {
            // Standard flat path: just replace the file content
            zip.file(mediaPath, imageBuffer, { compression: 'STORE' });
            modifiedFiles.add(mediaPath);
          }
        }
      } catch (err) {
        console.error(`Error replacing image rId ${imgMod.imageRid} on slide ${imgMod.slideIndex + 1}:`, err);
      }
    }
  }

  // Step 2b: Clean up WPS .image/ directory entries from the ZIP.
  // WPS creates non-standard directory structures like ppt/media/image5.image/
  // with the actual image file inside as ppt/media/image5.image/png. When we
  // replace images, we normalize these to flat paths (ppt/media/image5.jpeg),
  // but the directory entry may remain as an empty leftover. Office rejects
  // packages with these empty non-standard directories, so we remove them.
  const wpsImageDirs: string[] = [];
  for (const [filePath, file] of Object.entries(zip.files)) {
    if (file.dir && filePath.match(/\.image\/$/)) {
      wpsImageDirs.push(filePath);
    }
  }
  for (const dir of wpsImageDirs) {
    // Only remove if no files exist inside the directory
    const hasChildren = Object.keys(zip.files).some(
      f => f.startsWith(dir) && !zip.files[f].dir
    );
    if (!hasChildren) {
      zip.remove(dir);
      console.log(`Removed empty WPS .image/ directory: ${dir}`);
    }
  }

  // Step 3: Sanitize [Content_Types].xml to fix WPS-specific invalid MIME types
  // This is critical for Windows Office compatibility:
  // - Fixes "image/.jpg" → "image/jpeg" (OPC spec violation)
  // - Normalizes uppercase extensions (Extension="JPG" → Extension="jpg")
  // - Ensures correct MIME mapping for all image types
  const contentTypesFile = zip.file('[Content_Types].xml');
  if (contentTypesFile) {
    const ctRawData = Buffer.from(await contentTypesFile.async('uint8array'));
    const ctBomPresent = hasUtf8Bom(ctRawData);
    const ctXml = bufferToString(ctRawData);
    const sanitizedXml = sanitizeContentTypesXml(ctXml);
    if (sanitizedXml !== ctXml) {
      console.log('[Content_Types].xml sanitized: fixed invalid MIME types or extensions');
      const ctOutputData = stringToBuffer(sanitizedXml, ctBomPresent);
      zip.file('[Content_Types].xml', ctOutputData, { compression: 'DEFLATE', compressionOptions: { level: 6 } });
      modifiedFiles.add('[Content_Types].xml');
    }
  }

  // Step 4: WPS file removal DISABLED (breaks macOS PowerPoint)
  // removeWpsSpecificFiles(zip);

  // Step 4b: WPS tag refs removal DISABLED
  // await removeTagReferencesFromRels(zip, modifiedFiles);

  // Step 5: Sanitize WPS-specific XML attributes from presentation files
  await sanitizeWpsXmlAttributes(zip);

  // Step 6: Ensure unmodified XML files keep DEFLATE compression
  const xmlExtensions = ['.xml', '.rels'];
  for (const [filePath, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    if (modifiedFiles.has(filePath)) continue;
    const isXml = xmlExtensions.some(ext => filePath.toLowerCase().endsWith(ext));
    if (isXml) {
      const data = await file.async('uint8array');
      zip.file(filePath, data, { compression: 'DEFLATE', compressionOptions: { level: 6 } });
    }
  }

  const outputBuffer = await zip.generateAsync({ type: 'nodebuffer' });

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
