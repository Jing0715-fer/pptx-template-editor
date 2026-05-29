import JSZip from 'jszip';
import { XMLParser } from 'fast-xml-parser';

// ============================================================================
// Types
// ============================================================================

export interface ElementPosition {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PptxRun {
  originalText: string;
  bold: boolean;
  italic: boolean;
  fontSize: number | null;
  fontColor: string | null;
  fontName: string | null;
}

export interface PptxParagraph {
  originalText: string;
  runs: PptxRun[];
}

export interface PptxTextElement {
  type: 'text';
  id: string;
  shapeName: string;
  originalText: string;
  paragraphs: PptxParagraph[];
  position: ElementPosition;
  slideIndex: number;
  elementIndex: number;
}

export interface PptxTableRow {
  cells: PptxTableCell[];
}

export interface PptxTableCell {
  text: string;
  originalText: string;
  rowSpan: number;
  colSpan: number;
  gridCol: number;
}

export interface PptxTableElement {
  type: 'table';
  id: string;
  shapeName: string;
  rows: PptxTableRow[];
  position: ElementPosition;
  slideIndex: number;
  elementIndex: number;
}

export interface PptxImageElement {
  type: 'image';
  id: string;
  shapeName: string;
  imageName: string;
  imageType: string;
  imageRid: string;
  position: ElementPosition;
  slideIndex: number;
  elementIndex: number;
}

export type PptxElement = PptxTextElement | PptxTableElement | PptxImageElement;

export interface PptxSlideData {
  slideNumber: number;
  elements: PptxElement[];
}

export interface SlideSize {
  width: number;  // EMU
  height: number; // EMU
}

export interface PptxParseResult {
  fileName: string;
  slideCount: number;
  slides: PptxSlideData[];
  slideSize: SlideSize;
  _rawEntries: Record<string, Uint8Array>;
}

// ============================================================================
// XML Tree Navigation Helpers (for preserveOrder mode)
// ============================================================================

type XmlNode = Record<string, unknown>;
type Children = XmlNode[];

function findChild(parent: Children, tagName: string): Children | null {
  for (const item of parent) {
    if (item[tagName] !== undefined && Array.isArray(item[tagName])) {
      return item[tagName] as Children;
    }
  }
  return null;
}

function findAllChildren(parent: Children, tagName: string): Children[] {
  const result: Children[] = [];
  for (const item of parent) {
    if (item[tagName] !== undefined && Array.isArray(item[tagName])) {
      result.push(item[tagName] as Children);
    }
  }
  return result;
}

function getAttrs(children: Children): Record<string, string> {
  for (const child of children) {
    if (child[':@'] !== undefined && typeof child[':@'] === 'object') {
      return child[':@'] as Record<string, string>;
    }
  }
  return {};
}

function findChildAttrs(parent: Children, tagName: string): Record<string, string> {
  for (const item of parent) {
    if (item[tagName] !== undefined && Array.isArray(item[tagName])) {
      if (item[':@'] !== undefined && typeof item[':@'] === 'object') {
        return item[':@'] as Record<string, string>;
      }
      return getAttrs(item[tagName] as Children);
    }
  }
  return {};
}

function findChildWithAttrs(parent: Children, tagName: string): { children: Children; attrs: Record<string, string> } | null {
  for (const item of parent) {
    if (item[tagName] !== undefined && Array.isArray(item[tagName])) {
      const children = item[tagName] as Children;
      const attrs = (item[':@'] !== undefined && typeof item[':@'] === 'object')
        ? item[':@'] as Record<string, string>
        : getAttrs(children);
      return { children, attrs };
    }
  }
  return null;
}

function getText(children: Children): string {
  for (const child of children) {
    if (child['#text'] !== undefined) {
      return String(child['#text']);
    }
  }
  return '';
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
// Position Extraction
// ============================================================================

function extractPositionFromSpPr(spPrChildren: Children | null): ElementPosition {
  if (!spPrChildren) return { x: 0, y: 0, width: 0, height: 0 };

  const xfrm = findChild(spPrChildren, 'a:xfrm');
  if (!xfrm) return { x: 0, y: 0, width: 0, height: 0 };

  const offAttrs = findChildAttrs(xfrm, 'a:off');
  const extAttrs = findChildAttrs(xfrm, 'a:ext');

  return {
    x: parseInt(offAttrs['@_x'] || offAttrs['x'] || '0', 10),
    y: parseInt(offAttrs['@_y'] || offAttrs['y'] || '0', 10),
    width: parseInt(extAttrs['@_cx'] || extAttrs['cx'] || '0', 10),
    height: parseInt(extAttrs['@_cy'] || extAttrs['cy'] || '0', 10),
  };
}

function extractPositionFromXfrm(xfrmChildren: Children | null): ElementPosition {
  if (!xfrmChildren) return { x: 0, y: 0, width: 0, height: 0 };

  const offAttrs = findChildAttrs(xfrmChildren, 'a:off');
  const extAttrs = findChildAttrs(xfrmChildren, 'a:ext');

  return {
    x: parseInt(offAttrs['@_x'] || offAttrs['x'] || '0', 10),
    y: parseInt(offAttrs['@_y'] || offAttrs['y'] || '0', 10),
    width: parseInt(extAttrs['@_cx'] || extAttrs['cx'] || '0', 10),
    height: parseInt(extAttrs['@_cy'] || extAttrs['cy'] || '0', 10),
  };
}

// ============================================================================
// Text Extraction
// ============================================================================

function extractRunProps(rPrInfo: { children: Children; attrs: Record<string, string> } | null): {
  bold: boolean;
  italic: boolean;
  fontSize: number | null;
  fontColor: string | null;
  fontName: string | null;
} {
  const result = {
    bold: false,
    italic: false,
    fontSize: null as number | null,
    fontColor: null as string | null,
    fontName: null as string | null,
  };

  if (!rPrInfo) return result;

  const { children: rPrChildren, attrs } = rPrInfo;
  result.bold = attrs['@_b'] === '1' || attrs['@_b'] === 'true' || attrs['b'] === '1' || attrs['b'] === 'true';
  result.italic = attrs['@_i'] === '1' || attrs['@_i'] === 'true' || attrs['i'] === '1' || attrs['i'] === 'true';

  const szVal = attrs['@_sz'] || attrs['sz'];
  if (szVal) result.fontSize = parseInt(szVal, 10) / 100;

  const latinInfo = findChildWithAttrs(rPrChildren, 'a:latin');
  if (latinInfo) {
    const typeface = latinInfo.attrs['@_typeface'] || latinInfo.attrs['typeface'];
    if (typeface) result.fontName = typeface;
  }

  const solidFillInfo = findChildWithAttrs(rPrChildren, 'a:solidFill');
  if (solidFillInfo) {
    const srgbClrInfo = findChildWithAttrs(solidFillInfo.children, 'a:srgbClr');
    if (srgbClrInfo) {
      const val = srgbClrInfo.attrs['@_val'] || srgbClrInfo.attrs['val'];
      if (val) result.fontColor = val;
    }
  }

  return result;
}

function extractParagraph(pChildren: Children): PptxParagraph {
  const runs: PptxRun[] = [];
  let fullText = '';
  const contentElements = getContentElements(pChildren);

  for (const { tagName, content } of contentElements) {
    if (tagName === 'a:r') {
      const rPrInfo = findChildWithAttrs(content, 'a:rPr');
      const props = extractRunProps(rPrInfo);
      const tChildren = findChild(content, 'a:t');
      const text = tChildren ? getText(tChildren) : '';
      runs.push({ originalText: text, ...props });
      fullText += text;
    } else if (tagName === 'a:fld') {
      const tChildren = findChild(content, 'a:t');
      const text = tChildren ? getText(tChildren) : '';
      runs.push({ originalText: text, bold: false, italic: false, fontSize: null, fontColor: null, fontName: null });
      fullText += text;
    } else if (tagName === 'a:br') {
      runs.push({ originalText: '\n', bold: false, italic: false, fontSize: null, fontColor: null, fontName: null });
      fullText += '\n';
    }
  }

  return { originalText: fullText, runs };
}

function extractTextBody(txBodyChildren: Children): { paragraphs: PptxParagraph[]; fullText: string } {
  const paragraphs: PptxParagraph[] = [];
  let fullText = '';

  const pAll = findAllChildren(txBodyChildren, 'a:p');
  for (const pChildren of pAll) {
    const para = extractParagraph(pChildren);
    paragraphs.push(para);
    if (fullText.length > 0) fullText += '\n';
    fullText += para.originalText;
  }

  return { paragraphs, fullText };
}

// ============================================================================
// Table Extraction
// ============================================================================

function extractTableCell(tcChildren: Children): PptxTableCell {
  const txBody = findChild(tcChildren, 'a:txBody');
  let text = '';
  let originalText = '';

  if (txBody) {
    const result = extractTextBody(txBody);
    text = result.fullText;
    originalText = result.fullText;
  }

  const attrs = getAttrs(tcChildren);
  return {
    text,
    originalText,
    rowSpan: parseInt(attrs['@_rowSpan'] || attrs['rowSpan'] || '1', 10),
    colSpan: parseInt(attrs['@_gridSpan'] || attrs['gridSpan'] || '1', 10),
    gridCol: parseInt(attrs['@_gridCol'] || attrs['gridCol'] || '0', 10),
  };
}

function extractTable(tblChildren: Children): PptxTableRow[] {
  const rows: PptxTableRow[] = [];
  const trAll = findAllChildren(tblChildren, 'a:tr');

  for (const trChildren of trAll) {
    const cells: PptxTableCell[] = [];
    const tcAll = findAllChildren(trChildren, 'a:tc');
    for (const tcChildren of tcAll) {
      cells.push(extractTableCell(tcChildren));
    }
    rows.push({ cells });
  }

  return rows;
}

// ============================================================================
// Rels Parsing
// ============================================================================

interface RelsMap {
  [rId: string]: string;
}

function parseRels(relsXml: string): RelsMap {
  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
  });

  const result: RelsMap = {};
  try {
    const parsed = parser.parse(relsXml);
    const relationships = parsed?.Relationships?.Relationship;
    if (relationships) {
      const relsArray = Array.isArray(relationships) ? relationships : [relationships];
      for (const rel of relsArray) {
        if (rel['@_Id'] && rel['@_Target']) {
          result[rel['@_Id']] = rel['@_Target'];
        }
      }
    }
  } catch {
    // Ignore parsing errors for rels
  }
  return result;
}

// ============================================================================
// Slide Parsing
// ============================================================================

const CONTENT_ELEMENT_TAGS = new Set([
  'p:sp',
  'p:grpSp',
  'p:graphicFrame',
  'p:pic',
  'p:cxnSp',
]);

function parseSlide(
  slideXml: string,
  slideIndex: number,
  slideNumber: number,
  relsMap: RelsMap,
): PptxSlideData {
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
    return { slideNumber, elements: [] };
  }

  const elements: PptxElement[] = [];

  const sldChildren = findChild(parsed, 'p:sld');
  if (!sldChildren) return { slideNumber, elements: [] };

  const cSldChildren = findChild(sldChildren, 'p:cSld');
  if (!cSldChildren) return { slideNumber, elements: [] };

  const spTreeChildren = findChild(cSldChildren, 'p:spTree');
  if (!spTreeChildren) return { slideNumber, elements: [] };

  const contentElements = getContentElements(spTreeChildren);
  let elementIndex = 0;

  for (const { tagName, content } of contentElements) {
    if (!CONTENT_ELEMENT_TAGS.has(tagName)) continue;

    try {
      if (tagName === 'p:sp') {
        const nvSpPr = findChild(content, 'p:nvSpPr');
        const shapeName = nvSpPr ? findChildAttrs(nvSpPr, 'p:cNvPr')['@_name'] || '' : '';
        const txBody = findChild(content, 'p:txBody');
        const spPr = findChild(content, 'p:spPr');
        const position = extractPositionFromSpPr(spPr);

        if (txBody) {
          const { paragraphs, fullText } = extractTextBody(txBody);
          elements.push({
            type: 'text',
            id: `slide${slideIndex}_el${elementIndex}`,
            shapeName,
            originalText: fullText,
            paragraphs,
            position,
            slideIndex,
            elementIndex,
          });
        }
        elementIndex++;
      } else if (tagName === 'p:grpSp') {
        const groupContentElements = getContentElements(content);
        for (const { tagName: childTag, content: childContent } of groupContentElements) {
          if (childTag === 'p:sp') {
            const nvSpPr = findChild(childContent, 'p:nvSpPr');
            const shapeName = nvSpPr ? findChildAttrs(nvSpPr, 'p:cNvPr')['@_name'] || '' : '';
            const txBody = findChild(childContent, 'p:txBody');
            const spPr = findChild(childContent, 'p:spPr');
            const position = extractPositionFromSpPr(spPr);

            if (txBody) {
              const { paragraphs, fullText } = extractTextBody(txBody);
              elements.push({
                type: 'text',
                id: `slide${slideIndex}_grpEl${elementIndex}`,
                shapeName,
                originalText: fullText,
                paragraphs,
                position,
                slideIndex,
                elementIndex,
              });
            }
            elementIndex++;
          } else if (CONTENT_ELEMENT_TAGS.has(childTag)) {
            elementIndex++;
          }
        }
      } else if (tagName === 'p:graphicFrame') {
        const nvGfxFramePr = findChild(content, 'p:nvGraphicFramePr');
        const shapeName = nvGfxFramePr ? findChildAttrs(nvGfxFramePr, 'p:cNvPr')['@_name'] || '' : '';
        const xfrm = findChild(content, 'p:xfrm');
        const position = extractPositionFromXfrm(xfrm);
        const graphic = findChild(content, 'a:graphic');
        if (graphic) {
          const graphicData = findChild(graphic, 'a:graphicData');
          if (graphicData) {
            const tbl = findChild(graphicData, 'a:tbl');
            if (tbl) {
              const rows = extractTable(tbl);
              elements.push({
                type: 'table',
                id: `slide${slideIndex}_el${elementIndex}`,
                shapeName,
                rows,
                position,
                slideIndex,
                elementIndex,
              });
            }
          }
        }
        elementIndex++;
      } else if (tagName === 'p:pic') {
        const nvPicPr = findChild(content, 'p:nvPicPr');
        const shapeName = nvPicPr ? findChildAttrs(nvPicPr, 'p:cNvPr')['@_name'] || '' : '';
        const spPr = findChild(content, 'p:spPr');
        const position = extractPositionFromSpPr(spPr);
        const blipFill = findChild(content, 'p:blipFill');
        let imageRid = '';
        if (blipFill) {
          const blipAttrs = findChildAttrs(blipFill, 'a:blip');
          imageRid = blipAttrs['@_r:embed'] || blipAttrs['r:embed'] || blipAttrs['@_embed'] || '';
        }

        let imageName = '';
        let imageType = '';
        if (imageRid && relsMap[imageRid]) {
          const target = relsMap[imageRid];
          const parts = target.split('/');
          imageName = parts[parts.length - 1];
          const ext = imageName.split('.').pop() || '';
          imageType = ext.toLowerCase();
        }

        elements.push({
          type: 'image',
          id: `slide${slideIndex}_el${elementIndex}`,
          shapeName,
          imageName,
          imageType,
          imageRid,
          position,
          slideIndex,
          elementIndex,
        });
        elementIndex++;
      } else if (tagName === 'p:cxnSp') {
        elementIndex++;
      }
    } catch (err) {
      console.warn(`Error parsing element ${tagName} at index ${elementIndex} in slide ${slideNumber}:`, err);
      elementIndex++;
    }
  }

  return { slideNumber, elements };
}

// ============================================================================
// Main Parse Function
// ============================================================================

export async function parsePptx(buffer: Buffer, fileName: string): Promise<PptxParseResult> {
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(buffer);
  } catch {
    throw new Error('无法读取 PPTX 文件，文件可能已损坏或不是有效的 ZIP 格式');
  }

  const rawEntries: Record<string, Uint8Array> = {};
  const slideFileNames: string[] = [];

  for (const [entryPath, entry] of Object.entries(zip.files)) {
    if (!entry.dir) {
      try {
        const data = await entry.async('uint8array');
        rawEntries[entryPath] = data;

        const slideMatch = entryPath.match(/^ppt\/slides\/slide(\d+)\.xml$/);
        if (slideMatch) {
          slideFileNames.push(entryPath);
        }
      } catch {
        console.warn(`Skipping unreadable entry: ${entryPath}`);
      }
    }
  }

  if (slideFileNames.length === 0) {
    throw new Error('未在文件中找到幻灯片，请确认是有效的 PPTX 文件');
  }

  slideFileNames.sort((a, b) => {
    const numA = parseInt(a.match(/slide(\d+)\.xml$/)![1], 10);
    const numB = parseInt(b.match(/slide(\d+)\.xml$/)![1], 10);
    return numA - numB;
  });

  const slides: PptxSlideData[] = [];

  for (let i = 0; i < slideFileNames.length; i++) {
    const slidePath = slideFileNames[i];
    try {
      const slideXml = new TextDecoder('utf-8').decode(rawEntries[slidePath]);
      const slideFileName = slidePath.split('/').pop()!;
      const relsPathWithXml = `ppt/slides/_rels/${slideFileName}.rels`;
      const relsPathWithoutXml = `ppt/slides/_rels/${slideFileName.replace('.xml', '')}.rels`;
      let relsMap: RelsMap = {};

      if (rawEntries[relsPathWithXml]) {
        const relsXml = new TextDecoder('utf-8').decode(rawEntries[relsPathWithXml]);
        relsMap = parseRels(relsXml);
      } else if (rawEntries[relsPathWithoutXml]) {
        const relsXml = new TextDecoder('utf-8').decode(rawEntries[relsPathWithoutXml]);
        relsMap = parseRels(relsXml);
      }

      const slideData = parseSlide(slideXml, i, i + 1, relsMap);
      slides.push(slideData);
    } catch (err) {
      console.error(`Error parsing slide ${i + 1}:`, err);
      slides.push({ slideNumber: i + 1, elements: [] });
    }
  }

  // Extract slide size from presentation.xml
  let slideSize: SlideSize = { width: 12192000, height: 6858000 }; // default 16:9
  try {
    const presXmlData = rawEntries['ppt/presentation.xml'];
    if (presXmlData) {
      const presXml = new TextDecoder('utf-8').decode(presXmlData);
      // Extract <p:sldSz cx="..." cy="..."/>
      const sldSzMatch = presXml.match(/<p:sldSz[^>]*\bcx\s*=\s*"(\d+)"[^>]*\bcy\s*=\s*"(\d+)"/);
      if (sldSzMatch) {
        slideSize = { width: parseInt(sldSzMatch[1], 10), height: parseInt(sldSzMatch[2], 10) };
      } else {
        // Try reversed attribute order
        const sldSzMatch2 = presXml.match(/<p:sldSz[^>]*\bcy\s*=\s*"(\d+)"[^>]*\bcx\s*=\s*"(\d+)"/);
        if (sldSzMatch2) {
          slideSize = { width: parseInt(sldSzMatch2[2], 10), height: parseInt(sldSzMatch2[1], 10) };
        }
      }
    }
  } catch { /* use default */ }

  return { fileName, slideCount: slides.length, slides, slideSize, _rawEntries: rawEntries };
}

// ============================================================================
// Image Extraction Helper
// ============================================================================

export function getImageAsBase64(rawEntries: Record<string, Uint8Array>, imageName: string): string | null {
  for (const [path, data] of Object.entries(rawEntries)) {
    if (path.endsWith('/' + imageName) || path === imageName) {
      const base64 = Buffer.from(data).toString('base64');
      const ext = imageName.split('.').pop()?.toLowerCase() || '';
      const mimeMap: Record<string, string> = {
        png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
        gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml',
        tiff: 'image/tiff', tif: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf',
      };
      const mime = mimeMap[ext] || 'application/octet-stream';
      return `data:${mime};base64,${base64}`;
    }
  }
  return null;
}
