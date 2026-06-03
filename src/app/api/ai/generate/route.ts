import { NextResponse } from 'next/server';
import { generateAiReport } from '@/lib/ai-report-generator';
import type { LlmProvider } from '@/app/api/ai/config/route';

export const maxDuration = 120;
export const dynamic = 'force-dynamic';

export async function POST(request: Request) {
  try {
    const contentType = request.headers.get('content-type') || '';
    if (!contentType.includes('multipart/form-data'))
      return NextResponse.json({ error: '请使用 multipart/form-data 格式提交数据' }, { status: 400 });

    let formData: FormData;
    try { formData = await request.formData(); } catch {
      return NextResponse.json({ error: '无法读取请求数据，请重试' }, { status: 400 });
    }

    const fileId = formData.get('fileId') as string | null;
    const dataSource = formData.get('dataSource') as File | null;
    const prompt = formData.get('prompt') as string | null;
    const providerStr = formData.get('provider') as string | null;

    // Validate provider
    let provider: LlmProvider | undefined;
    if (providerStr) {
      if (providerStr !== 'openai' && providerStr !== 'anthropic') {
        return NextResponse.json({ error: '无效的 AI 提供商参数' }, { status: 400 });
      }
      provider = providerStr;
    }

    if (!fileId || typeof fileId !== 'string')
      return NextResponse.json({ error: '缺少 fileId 参数' }, { status: 400 });

    const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!uuidRegex.test(fileId))
      return NextResponse.json({ error: '无效的 fileId 格式' }, { status: 400 });

    if (!dataSource)
      return NextResponse.json({ error: '请上传数据源文件（.docx 或 .xlsx）' }, { status: 400 });

    const fileName = dataSource.name.toLowerCase();
    if (!fileName.endsWith('.docx') && !fileName.endsWith('.xlsx') && !fileName.endsWith('.xls'))
      return NextResponse.json({ error: '数据源文件仅支持 .docx 和 .xlsx 格式' }, { status: 400 });
    if (dataSource.size > 50 * 1024 * 1024)
      return NextResponse.json({ error: '数据源文件大小不能超过 50MB' }, { status: 400 });
    if (dataSource.size === 0)
      return NextResponse.json({ error: '数据源文件内容为空' }, { status: 400 });

    const arrayBuffer = await dataSource.arrayBuffer();
    const dataSourceBuffer = Buffer.from(arrayBuffer);

    const result = await generateAiReport(fileId, dataSourceBuffer, dataSource.name, prompt || undefined, provider);

    if (result.modifications.length === 0)
      return NextResponse.json({ modifications: [], summary: 'AI 未能从数据源中找到匹配模板字段的内容' });

    return NextResponse.json({ modifications: result.modifications, summary: result.summary });
  } catch (error) {
    console.error('AI report generation error:', error);
    const message = error instanceof Error ? error.message : 'AI 报告生成失败';
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
