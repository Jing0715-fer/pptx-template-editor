import ZAI from 'z-ai-web-dev-sdk';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import { parsePptx } from '@/lib/pptx-parser';
import { readFile } from 'fs/promises';
import path from 'path';

export interface AiModification {
  slideIndex: number;
  elementIndex: number;
  type: 'text' | 'table';
  newText?: string;
  tableCells?: { row: number; col: number; text: string }[];
}

export interface AiGenerateResult {
  modifications: AiModification[];
  summary?: string;
}

export interface PptxTemplateSummary {
  slides: {
    slideIndex: number;
    slideNumber: number;
    elements: {
      elementIndex: number;
      type: 'text' | 'table';
      originalText?: string;
      rows?: { cells: { text: string }[] }[];
    }[];
  }[];
}

async function parseDocx(buffer: Buffer): Promise<string> {
  try {
    const result = await mammoth.extractRawText({ buffer });
    return result.value;
  } catch (err) {
    throw new Error(`无法解析 Word 文件: ${err instanceof Error ? err.message : '未知错误'}`);
  }
}

function parseXlsx(buffer: Buffer): string {
  try {
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetsData: Record<string, unknown[]> = {};
    for (const sheetName of workbook.SheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      if (worksheet) sheetsData[sheetName] = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    }
    return JSON.stringify(sheetsData, null, 2);
  } catch (err) {
    throw new Error(`无法解析 Excel 文件: ${err instanceof Error ? err.message : '未知错误'}`);
  }
}

export async function parseDataSource(buffer: Buffer, fileName: string): Promise<string> {
  const ext = fileName.toLowerCase().split('.').pop();
  if (ext === 'docx') return parseDocx(buffer);
  if (ext === 'xlsx' || ext === 'xls') return parseXlsx(buffer);
  throw new Error(`不支持的数据源文件格式: .${ext}，仅支持 .docx 和 .xlsx`);
}

export function buildTemplateSummary(
  slides: { slideNumber: number; elements: { type: string; originalText?: string; rows?: { cells: { text: string }[] }[]; slideIndex: number; elementIndex: number }[] }[]
): PptxTemplateSummary {
  return {
    slides: slides.map((slide) => ({
      slideIndex: slide.elements[0]?.slideIndex ?? 0,
      slideNumber: slide.slideNumber,
      elements: slide.elements
        .filter((el) => el.type === 'text' || el.type === 'table')
        .map((el) => {
          if (el.type === 'text') return { elementIndex: el.elementIndex, type: 'text' as const, originalText: el.originalText || '' };
          if (el.type === 'table' && el.rows) return { elementIndex: el.elementIndex, type: 'table' as const, rows: el.rows.map((r) => ({ cells: r.cells.map((c) => ({ text: c.text })) })) };
          return null;
        })
        .filter((el): el is NonNullable<typeof el> => el !== null),
    })),
  };
}

export async function readPptxTemplate(fileId: string): Promise<{
  slides: { slideNumber: number; elements: { type: string; originalText?: string; rows?: { cells: { text: string }[] }[]; slideIndex: number; elementIndex: number }[] }[];
  fileName: string;
}> {
  const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  if (!uuidRegex.test(fileId)) throw new Error('无效的 fileId 格式');

  const tempDir = path.join(process.cwd(), 'temp-uploads');
  const pptxPath = path.join(tempDir, `${fileId}.pptx`);

  let pptxBuffer: Buffer;
  try { pptxBuffer = await readFile(pptxPath); } catch { throw new Error('PPTX 文件不存在或已过期，请重新上传'); }

  const parseResult = await parsePptx(pptxBuffer, 'template.pptx');
  return { slides: parseResult.slides, fileName: parseResult.fileName };
}

function buildSystemPrompt(): string {
  return `You are a professional report generation assistant. Your task is to intelligently map data from a source document to the appropriate fields in a PPTX template.

## Your Responsibilities:
1. Read and understand the PPTX template structure (text placeholders and table cells)
2. Read and understand the source document content (from Word or Excel files)
3. Match relevant data from the source document to the most appropriate PPTX template fields
4. Return a JSON object with the modifications

## Important Rules:
- For text fields: provide the complete replacement text
- For table cells: provide cell-by-cell replacements with row and column indices
- Only map data that clearly matches a template field - do NOT make up or fabricate data
- If the source document contains information that doesn't fit any template field, skip it
- Preserve the meaning and accuracy of the source data
- For tables, map data row by row, column by column from the source
- If a template field has no matching data in the source, do NOT include it in modifications
- For numeric data, keep the numbers accurate
- For dates, preserve the date format from the source
- Be concise - do not add extra commentary or explanation in the replacement text

## Output Format:
Return ONLY a valid JSON object (no markdown, no code blocks, no extra text) in this exact format:
{
  "modifications": [
    { "slideIndex": 0, "elementIndex": 0, "type": "text", "newText": "extracted value" },
    { "slideIndex": 0, "elementIndex": 1, "type": "table", "tableCells": [{"row": 0, "col": 0, "text": "value"}, {"row": 0, "col": 1, "text": "value"}] }
  ]
}`;
}

function buildUserPrompt(
  templateSummary: PptxTemplateSummary,
  documentContent: string,
  customPrompt?: string
): string {
  let prompt = `## PPTX Template Structure:\n${JSON.stringify(templateSummary, null, 2)}\n\n## Source Document Content:\n${documentContent}`;
  if (customPrompt && customPrompt.trim()) prompt += `\n\n## Additional Instructions from User:\n${customPrompt.trim()}`;
  prompt += `\n\n## Task:\nAnalyze the PPTX template structure and the source document content above. Map the data from the source document to the appropriate PPTX template fields. Return ONLY a valid JSON object with the modifications array.`;
  return prompt;
}

function parseLlmResponse(responseContent: string): AiModification[] {
  if (!responseContent || typeof responseContent !== 'string') throw new Error('LLM 返回了空响应');
  let jsonStr = responseContent.trim();

  const codeBlockMatch = jsonStr.match(/```(?:json)?\s*\n?([\s\S]*?)\n?\s*```/);
  if (codeBlockMatch) jsonStr = codeBlockMatch[1].trim();

  const jsonMatch = jsonStr.match(/\{[\s\S]*\}/);
  if (!jsonMatch) throw new Error('LLM 响应中未找到有效的 JSON 对象');
  jsonStr = jsonMatch[0];

  let parsed: { modifications?: AiModification[] };
  try { parsed = JSON.parse(jsonStr); } catch { throw new Error('LLM 返回的 JSON 格式无效，无法解析'); }

  if (!parsed.modifications || !Array.isArray(parsed.modifications)) throw new Error('LLM 返回的 JSON 中缺少 modifications 数组');

  const validModifications: AiModification[] = [];
  for (const mod of parsed.modifications) {
    if (typeof mod.slideIndex === 'number' && typeof mod.elementIndex === 'number' && (mod.type === 'text' || mod.type === 'table')) {
      if (mod.type === 'text' && typeof mod.newText === 'string') {
        validModifications.push({ slideIndex: mod.slideIndex, elementIndex: mod.elementIndex, type: 'text', newText: mod.newText });
      } else if (mod.type === 'table' && Array.isArray(mod.tableCells)) {
        const validCells = mod.tableCells.filter(
          (cell: { row?: number; col?: number; text?: string }) =>
            typeof cell.row === 'number' && typeof cell.col === 'number' && typeof cell.text === 'string'
        );
        if (validCells.length > 0) {
          validModifications.push({ slideIndex: mod.slideIndex, elementIndex: mod.elementIndex, type: 'table', tableCells: validCells as { row: number; col: number; text: string }[] });
        }
      }
    }
  }
  return validModifications;
}

export async function generateAiReport(
  fileId: string,
  dataSourceBuffer: Buffer,
  dataSourceFileName: string,
  customPrompt?: string
): Promise<AiGenerateResult> {
  const documentContent = await parseDataSource(dataSourceBuffer, dataSourceFileName);
  if (!documentContent.trim()) throw new Error('数据源文件内容为空，无法提取有效数据');

  const template = await readPptxTemplate(fileId);
  if (!template.slides || template.slides.length === 0) throw new Error('PPTX 模板中没有幻灯片');

  const templateSummary = buildTemplateSummary(template.slides);
  const totalFields = templateSummary.slides.reduce((acc, slide) => acc + slide.elements.length, 0);
  if (totalFields === 0) throw new Error('PPTX 模板中没有可编辑的文本或表格字段');

  const systemPrompt = buildSystemPrompt();
  const userPrompt = buildUserPrompt(templateSummary, documentContent, customPrompt);

  let zai: InstanceType<typeof ZAI>;
  try { zai = await ZAI.create(); } catch (err) { throw new Error(`AI 服务初始化失败: ${err instanceof Error ? err.message : '未知错误'}`); }

  let completion: { choices?: { message?: { content?: string } }[] };
  try {
    completion = await zai.chat.completions.create({
      messages: [
        { role: 'system', content: systemPrompt },
        { role: 'user', content: userPrompt },
      ],
      thinking: { type: 'disabled' },
    });
  } catch (err) { throw new Error(`AI 模型调用失败: ${err instanceof Error ? err.message : '未知错误'}`); }

  const responseContent = completion.choices?.[0]?.message?.content;
  if (!responseContent) throw new Error('AI 模型返回了空响应，请重试');

  const modifications = parseLlmResponse(responseContent);
  const validatedModifications = modifications.filter((mod) => {
    const slide = templateSummary.slides.find((s) => s.slideIndex === mod.slideIndex);
    if (!slide) return false;
    const element = slide.elements.find((e) => e.elementIndex === mod.elementIndex && e.type === mod.type);
    return !!element;
  });

  const textModCount = validatedModifications.filter((m) => m.type === 'text').length;
  const tableModCount = validatedModifications.filter((m) => m.type === 'table').length;
  const summary = `已生成 ${validatedModifications.length} 项修改（${textModCount} 个文本字段，${tableModCount} 个表格）`;

  return { modifications: validatedModifications, summary };
}
