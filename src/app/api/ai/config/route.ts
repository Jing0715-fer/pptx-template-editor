import { NextResponse } from 'next/server';
import { readFile, writeFile } from 'fs/promises';
import path from 'path';

export const dynamic = 'force-dynamic';

const CONFIG_PATH = path.join(process.cwd(), '.z-ai-config');

// ============================================================================
// Types
// ============================================================================

export type LlmProvider = 'openai' | 'anthropic';

export interface ProviderConfig {
  baseUrl: string;
  apiKey: string;
  model: string;
}

export interface AiConfig {
  defaultProvider: LlmProvider;
  openai: ProviderConfig;
  anthropic: ProviderConfig;
}

export const DEFAULT_CONFIG: AiConfig = {
  defaultProvider: 'openai',
  openai: {
    baseUrl: 'https://api.openai.com/v1',
    apiKey: '',
    model: 'gpt-4o',
  },
  anthropic: {
    baseUrl: 'https://api.anthropic.com',
    apiKey: '',
    model: 'claude-sonnet-4-20250514',
  },
};

// ============================================================================
// Helpers
// ============================================================================

function maskApiKey(key: string): string {
  if (!key) return '';
  if (key.length <= 8) return '****';
  return key.slice(0, 4) + '****' + key.slice(-4);
}

function normalizeOpenaiUrl(url: string): string {
  let normalized = url.trim().replace(/\/+$/, '');
  // For OpenAI-compatible, ensure /v1 suffix
  if (!normalized.endsWith('/v1')) {
    normalized += '/v1';
  }
  return normalized;
}

function normalizeAnthropicUrl(url: string): string {
  let normalized = url.trim().replace(/\/+$/, '');
  // For Anthropic, strip /v1 if present (we'll add the correct path later)
  normalized = normalized.replace(/\/v1$/, '');
  return normalized;
}

async function readConfig(): Promise<AiConfig> {
  try {
    const raw = await readFile(CONFIG_PATH, 'utf-8');
    const parsed = JSON.parse(raw);
    // Merge with defaults to ensure all fields exist
    return {
      defaultProvider: parsed.defaultProvider || DEFAULT_CONFIG.defaultProvider,
      openai: { ...DEFAULT_CONFIG.openai, ...parsed.openai },
      anthropic: { ...DEFAULT_CONFIG.anthropic, ...parsed.anthropic },
    };
  } catch {
    return { ...DEFAULT_CONFIG };
  }
}

function maskConfig(config: AiConfig) {
  return {
    defaultProvider: config.defaultProvider,
    openai: {
      baseUrl: config.openai.baseUrl,
      apiKeyMasked: maskApiKey(config.openai.apiKey),
      model: config.openai.model,
      configured: !!(config.openai.baseUrl && config.openai.apiKey),
    },
    anthropic: {
      baseUrl: config.anthropic.baseUrl,
      apiKeyMasked: maskApiKey(config.anthropic.apiKey),
      model: config.anthropic.model,
      configured: !!(config.anthropic.baseUrl && config.anthropic.apiKey),
    },
    configured: !!(config.openai.apiKey || config.anthropic.apiKey),
  };
}

// ============================================================================
// API Routes
// ============================================================================

/**
 * GET /api/ai/config — Read current AI config (apiKeys are masked)
 */
export async function GET() {
  try {
    const config = await readConfig();
    return NextResponse.json(maskConfig(config));
  } catch (error) {
    console.error('Failed to read AI config:', error);
    return NextResponse.json({ error: '读取配置失败' }, { status: 500 });
  }
}

/**
 * POST /api/ai/config — Save AI config
 * Body: { provider: 'openai' | 'anthropic', baseUrl, apiKey, model, defaultProvider? }
 */
export async function POST(request: Request) {
  try {
    const body = await request.json();
    const { provider, baseUrl, apiKey, model, defaultProvider } = body as {
      provider?: LlmProvider;
      baseUrl?: string;
      apiKey?: string;
      model?: string;
      defaultProvider?: LlmProvider;
    };

    if (!provider || (provider !== 'openai' && provider !== 'anthropic')) {
      return NextResponse.json({ error: '请选择 API 提供商' }, { status: 400 });
    }
    if (!baseUrl || typeof baseUrl !== 'string') {
      return NextResponse.json({ error: '请输入 API Base URL' }, { status: 400 });
    }
    if (!apiKey || typeof apiKey !== 'string') {
      return NextResponse.json({ error: '请输入 API Key' }, { status: 400 });
    }
    if (!model || typeof model !== 'string') {
      return NextResponse.json({ error: '请输入模型名称' }, { status: 400 });
    }

    const config = await readConfig();

    // Update the specified provider
    if (provider === 'openai') {
      config.openai = {
        baseUrl: normalizeOpenaiUrl(baseUrl),
        apiKey: apiKey.trim(),
        model: model.trim(),
      };
    } else {
      config.anthropic = {
        baseUrl: normalizeAnthropicUrl(baseUrl),
        apiKey: apiKey.trim(),
        model: model.trim(),
      };
    }

    // Update default provider if specified
    if (defaultProvider && (defaultProvider === 'openai' || defaultProvider === 'anthropic')) {
      config.defaultProvider = defaultProvider;
    }

    await writeFile(CONFIG_PATH, JSON.stringify(config, null, 2), 'utf-8');

    return NextResponse.json({
      success: true,
      ...maskConfig(config),
    });
  } catch (error) {
    console.error('Failed to save AI config:', error);
    return NextResponse.json({ error: '保存配置失败' }, { status: 500 });
  }
}

/**
 * DELETE /api/ai/config — Remove a specific provider's config or all
 * Query: ?provider=openai|anthropic|all
 */
export async function DELETE(request: Request) {
  try {
    const url = new URL(request.url);
    const provider = url.searchParams.get('provider') || 'all';

    if (provider === 'all') {
      await writeFile(CONFIG_PATH, JSON.stringify(DEFAULT_CONFIG, null, 2), 'utf-8');
      return NextResponse.json({ success: true, ...maskConfig(DEFAULT_CONFIG) });
    }

    if (provider !== 'openai' && provider !== 'anthropic') {
      return NextResponse.json({ error: '无效的提供商参数' }, { status: 400 });
    }

    const config = await readConfig();
    config[provider] = { ...DEFAULT_CONFIG[provider] };

    // If deleting the default provider, switch to the other one
    if (config.defaultProvider === provider) {
      config.defaultProvider = provider === 'openai' ? 'anthropic' : 'openai';
    }

    await writeFile(CONFIG_PATH, JSON.stringify(config, null, 2), 'utf-8');

    return NextResponse.json({ success: true, ...maskConfig(config) });
  } catch (error) {
    console.error('Failed to delete AI config:', error);
    return NextResponse.json({ error: '删除配置失败' }, { status: 500 });
  }
}
