'use client';

import { useState, useCallback, useRef, useEffect, type DragEvent, type ChangeEvent } from 'react';
import { usePptxStore, type PptxModification } from '@/lib/pptx-store';
import { cn } from '@/lib/utils';
import { motion, AnimatePresence } from 'framer-motion';
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
  DialogTrigger,
} from '@/components/ui/dialog';
import { Button } from '@/components/ui/button';
import { Textarea } from '@/components/ui/textarea';
import { Badge } from '@/components/ui/badge';
import { Separator } from '@/components/ui/separator';
import { ScrollArea } from '@/components/ui/scroll-area';
import {
  Sparkles,
  Upload,
  Loader2,
  FileText,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle,
  X,
  Wand2,
  Settings2,
  Zap,
} from 'lucide-react';

// ============================================================================
// Types
// ============================================================================

type LlmProvider = 'openai' | 'anthropic';

interface AiGenerateDialogProps {
  children?: React.ReactNode;
  open?: boolean;
  onOpenChange?: (open: boolean) => void;
  onOpenSettings?: () => void;
}

interface AiResult {
  modifications: PptxModification[];
  summary: string;
}

interface ConfigInfo {
  configured: boolean;
  defaultProvider: LlmProvider;
  openai: { configured: boolean; model: string };
  anthropic: { configured: boolean; model: string };
}

// ============================================================================
// Animation variants
// ============================================================================

const fadeSlideUp = {
  initial: { opacity: 0, y: 12 },
  animate: { opacity: 1, y: 0 },
  exit: { opacity: 0, y: -8 },
  transition: { duration: 0.25, ease: [0.25, 0.46, 0.45, 0.94] },
};

const scaleIn = {
  initial: { opacity: 0, scale: 0.95 },
  animate: { opacity: 1, scale: 1 },
  exit: { opacity: 0, scale: 0.95 },
  transition: { duration: 0.2, ease: [0.25, 0.46, 0.45, 0.94] },
};

const staggerContainer = {
  animate: { transition: { staggerChildren: 0.06 } },
};

const staggerItem = {
  initial: { opacity: 0, y: 8 },
  animate: { opacity: 1, y: 0 },
};

// ============================================================================
// Provider config
// ============================================================================

const PROVIDERS: Record<LlmProvider, { label: string; icon: React.ReactNode; color: string; bgColor: string }> = {
  openai: {
    label: 'OpenAI',
    icon: <Sparkles className="size-3.5" />,
    color: 'text-emerald-600 dark:text-emerald-400',
    bgColor: 'bg-emerald-50 dark:bg-emerald-900/20 border-emerald-200 dark:border-emerald-800',
  },
  anthropic: {
    label: 'Anthropic',
    icon: <Zap className="size-3.5" />,
    color: 'text-amber-600 dark:text-amber-400',
    bgColor: 'bg-amber-50 dark:bg-amber-900/20 border-amber-200 dark:border-amber-800',
  },
};

// ============================================================================
// Component
// ============================================================================

export function AiGenerateDialog({ children, open, onOpenChange, onOpenSettings }: AiGenerateDialogProps) {
  const { fileId, slides, applyAiModifications, fileName } = usePptxStore();

  // ── State ────────────────────────────────────────────────────────
  const [dataSource, setDataSource] = useState<File | null>(null);
  const [prompt, setPrompt] = useState('');
  const [isGenerating, setIsGenerating] = useState(false);
  const [result, setResult] = useState<AiResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [isDragOver, setIsDragOver] = useState(false);
  const [isApplied, setIsApplied] = useState(false);
  const [selectedProvider, setSelectedProvider] = useState<LlmProvider>('openai');
  const [configInfo, setConfigInfo] = useState<ConfigInfo | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ── Load config when dialog opens ──────────────────────────────
  useEffect(() => {
    if (open) {
      fetch('/api/ai/config')
        .then((res) => res.json())
        .then((data) => {
          setConfigInfo(data);
          setSelectedProvider(data.defaultProvider || 'openai');
        })
        .catch(() => {});
    }
  }, [open]);

  // ── Template info ────────────────────────────────────────────────
  const slideCount = slides.length;
  const totalElements = slides.reduce((acc, s) => acc + s.elements.length, 0);

  // ── Config status ────────────────────────────────────────────────
  const currentProviderConfig = configInfo?.[selectedProvider];
  const isCurrentProviderConfigured = currentProviderConfig?.configured ?? false;
  const anyProviderConfigured = configInfo?.openai?.configured || configInfo?.anthropic?.configured;

  // ── Reset state when dialog closes ───────────────────────────────
  const handleOpenChange = useCallback(
    (nextOpen: boolean) => {
      if (!nextOpen) {
        setTimeout(() => {
          setDataSource(null);
          setPrompt('');
          setResult(null);
          setError(null);
          setIsDragOver(false);
          setIsApplied(false);
          if (fileInputRef.current) fileInputRef.current.value = '';
        }, 200);
      }
      onOpenChange?.(nextOpen);
    },
    [onOpenChange],
  );

  // ── File handling ────────────────────────────────────────────────
  const validateFile = (file: File): string | null => {
    const name = file.name.toLowerCase();
    if (!name.endsWith('.docx') && !name.endsWith('.xlsx') && !name.endsWith('.xls')) {
      return '仅支持 .docx, .xlsx, .xls 格式文件';
    }
    if (file.size > 50 * 1024 * 1024) {
      return '文件大小不能超过 50MB';
    }
    if (file.size === 0) {
      return '文件内容为空';
    }
    return null;
  };

  const handleFileSelect = useCallback(
    (file: File) => {
      const validationError = validateFile(file);
      if (validationError) {
        setError(validationError);
        return;
      }
      setDataSource(file);
      setError(null);
      setResult(null);
      setIsApplied(false);
    },
    [],
  );

  const handleInputChange = useCallback(
    (e: ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) handleFileSelect(file);
    },
    [handleFileSelect],
  );

  const handleDragOver = useCallback((e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e: DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragOver(false);
  }, []);

  const handleDrop = useCallback(
    (e: DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      setIsDragOver(false);
      const file = e.dataTransfer.files?.[0];
      if (file) handleFileSelect(file);
    },
    [handleFileSelect],
  );

  const removeFile = useCallback(() => {
    setDataSource(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  }, []);

  // ── AI Generation ────────────────────────────────────────────────
  const handleGenerate = useCallback(async () => {
    if (!fileId || !dataSource) return;

    setIsGenerating(true);
    setError(null);
    setResult(null);
    setIsApplied(false);

    try {
      const formData = new FormData();
      formData.append('fileId', fileId);
      formData.append('dataSource', dataSource);
      formData.append('provider', selectedProvider);
      if (prompt.trim()) {
        formData.append('prompt', prompt.trim());
      }

      const response = await fetch('/api/ai/generate', {
        method: 'POST',
        body: formData,
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || 'AI 生成失败');
      }

      setResult({
        modifications: data.modifications || [],
        summary: data.summary || '',
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : 'An unexpected error occurred');
    } finally {
      setIsGenerating(false);
    }
  }, [fileId, dataSource, prompt, selectedProvider]);

  // ── Apply modifications ──────────────────────────────────────────
  const handleApply = useCallback(() => {
    if (!result || result.modifications.length === 0) return;
    applyAiModifications(result.modifications);
    setIsApplied(true);
  }, [result, applyAiModifications]);

  // ── Reset / Regenerate ───────────────────────────────────────────
  const handleReset = useCallback(() => {
    setResult(null);
    setError(null);
    setIsApplied(false);
  }, []);

  // ── Helpers ──────────────────────────────────────────────────────
  const getFileIcon = (name: string) => {
    if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
      return <FileSpreadsheet className="size-4" />;
    }
    return <FileText className="size-4" />;
  };

  const formatFileSize = (bytes: number) => {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  };

  const canGenerate = fileId && dataSource && !isGenerating && isCurrentProviderConfigured;
  const hasResult = result !== null;
  const hasError = error !== null;

  // ── Modification label helper ────────────────────────────────────
  const getModLabel = (mod: PptxModification): string => {
    const slide = slides[mod.slideIndex];
    const element = slide?.elements.find(
      (el) => el.slideIndex === mod.slideIndex && el.elementIndex === mod.elementIndex,
    );
    if (element) {
      if (element.type === 'text') {
        const text = element.originalText || element.shapeName;
        return text.length > 40 ? text.slice(0, 40) + '...' : text;
      }
      if (element.type === 'table') {
        return `Table: ${element.shapeName}`;
      }
    }
    return `Slide ${mod.slideIndex + 1} - Element ${mod.elementIndex + 1}`;
  };

  const getModNewValue = (mod: PptxModification): string => {
    if (mod.type === 'text' && mod.newText) {
      return mod.newText.length > 50 ? mod.newText.slice(0, 50) + '...' : mod.newText;
    }
    if (mod.type === 'table' && mod.tableCells) {
      return `${mod.tableCells.length} cell${mod.tableCells.length !== 1 ? 's' : ''} updated`;
    }
    return '-';
  };

  // ====================================================================
  // Render
  // ====================================================================

  return (
    <Dialog open={open} onOpenChange={handleOpenChange}>
      {children && <DialogTrigger asChild>{children}</DialogTrigger>}

      <DialogContent
        showCloseButton={false}
        className="sm:max-w-2xl w-[calc(100%-2rem)] p-0 gap-0 overflow-hidden rounded-2xl border-0 shadow-2xl bg-white dark:bg-zinc-950"
      >
        {/* ── Header ─────────────────────────────────────────────── */}
        <div className="relative px-6 pt-6 pb-4">
          <div className="absolute inset-0 overflow-hidden rounded-t-2xl">
            <div className="absolute -top-20 -right-20 size-60 rounded-full bg-gradient-to-br from-violet-200/40 to-fuchsia-200/30 blur-3xl dark:from-violet-900/30 dark:to-fuchsia-900/20" />
            <div className="absolute -bottom-10 -left-10 size-40 rounded-full bg-gradient-to-tr from-amber-200/30 to-rose-200/20 blur-3xl dark:from-amber-900/20 dark:to-rose-900/10" />
          </div>

          <div className="relative flex items-start justify-between">
            <div className="flex items-center gap-3.5">
              <motion.div
                initial={{ rotate: -10, scale: 0.8 }}
                animate={{ rotate: 0, scale: 1 }}
                transition={{ type: 'spring', stiffness: 260, damping: 20, delay: 0.1 }}
                className="flex items-center justify-center size-11 rounded-xl bg-gradient-to-br from-violet-500 to-fuchsia-600 shadow-lg shadow-violet-500/25"
              >
                <Wand2 className="size-5.5 text-white" />
              </motion.div>
              <div>
                <DialogTitle className="text-xl font-bold tracking-tight text-zinc-900 dark:text-zinc-50">
                  AI 内容生成
                </DialogTitle>
                <DialogDescription className="text-sm text-zinc-500 dark:text-zinc-400 mt-0.5">
                  上传数据源，让 AI 智能填充模板
                </DialogDescription>
              </div>
            </div>

            <button
              onClick={() => handleOpenChange(false)}
              className="flex items-center justify-center size-8 rounded-lg text-zinc-400 hover:text-zinc-600 hover:bg-zinc-100 dark:hover:text-zinc-300 dark:hover:bg-zinc-800 transition-all duration-150 mt-0.5"
            >
              <X className="size-4" />
            </button>
          </div>
        </div>

        <Separator className="bg-zinc-100 dark:bg-zinc-800" />

        {/* ── Body ───────────────────────────────────────────────── */}
        <ScrollArea className="max-h-[65vh]">
          <div className="px-6 py-5 space-y-5">
            {/* Template Info Card */}
            <motion.div {...fadeSlideUp} className="relative">
              <div className="rounded-xl border border-zinc-200/80 dark:border-zinc-800/80 bg-zinc-50/80 dark:bg-zinc-900/50 p-3.5">
                <div className="flex items-center gap-3">
                  <div className="flex items-center justify-center size-9 rounded-lg bg-gradient-to-br from-zinc-200 to-zinc-300 dark:from-zinc-700 dark:to-zinc-600 shrink-0">
                    <FileText className="size-4 text-zinc-600 dark:text-zinc-300" />
                  </div>
                  <div className="min-w-0 flex-1">
                    <p className="text-sm font-semibold text-zinc-800 dark:text-zinc-200 truncate">
                      {fileName || 'Untitled Template'}
                    </p>
                    <p className="text-xs text-zinc-500 dark:text-zinc-400 mt-0.5">
                      {slideCount} 页幻灯片 &middot; {totalElements} 个元素
                    </p>
                  </div>
                  <Badge
                    variant="secondary"
                    className="shrink-0 text-[10px] font-medium bg-violet-100/80 text-violet-700 border-violet-200/50 dark:bg-violet-900/30 dark:text-violet-300 dark:border-violet-800/40"
                  >
                    Active
                  </Badge>
                </div>
              </div>
            </motion.div>

            {/* Provider Selection */}
            <motion.div {...fadeSlideUp} transition={{ ...fadeSlideUp.transition, delay: 0.03 }}>
              <div className="flex items-center justify-between mb-2">
                <label className="text-sm font-medium text-zinc-700 dark:text-zinc-300">
                  AI 模型 <span className="text-rose-500">*</span>
                </label>
                <button
                  onClick={onOpenSettings}
                  className="flex items-center gap-1 text-xs text-violet-600 dark:text-violet-400 hover:text-violet-700 dark:hover:text-violet-300 transition-colors"
                >
                  <Settings2 className="size-3" />
                  设置
                </button>
              </div>
              <div className="flex gap-2">
                {(Object.keys(PROVIDERS) as LlmProvider[]).map((provider) => {
                  const info = PROVIDERS[provider];
                  const isActive = selectedProvider === provider;
                  const isConfigured = configInfo?.[provider]?.configured ?? false;

                  return (
                    <button
                      key={provider}
                      onClick={() => setSelectedProvider(provider)}
                      className={cn(
                        'flex-1 flex items-center justify-center gap-2 px-3 py-2.5 rounded-xl text-sm font-medium transition-all duration-200 border',
                        isActive
                          ? cn(info.bgColor, info.color, 'shadow-sm')
                          : 'border-transparent text-zinc-500 dark:text-zinc-400 hover:bg-zinc-50 dark:hover:bg-zinc-900',
                      )}
                    >
                      {info.icon}
                      {info.label}
                      {isConfigured ? (
                        <CheckCircle2 className="size-3 opacity-60" />
                      ) : (
                        <AlertCircle className="size-3 opacity-40" />
                      )}
                    </button>
                  );
                })}
              </div>
              {/* Provider not configured warning */}
              {configInfo && !isCurrentProviderConfigured && (
                <div className="mt-2 rounded-lg border border-amber-200 dark:border-amber-800/60 bg-amber-50 dark:bg-amber-900/20 px-3 py-2 flex items-center gap-2">
                  <AlertCircle className="size-3.5 text-amber-500 shrink-0" />
                  <p className="text-xs text-amber-700 dark:text-amber-400">
                    {PROVIDERS[selectedProvider].label} 未配置，请先
                    <button
                      onClick={onOpenSettings}
                      className="underline font-medium hover:text-amber-800 dark:hover:text-amber-300 mx-0.5"
                    >
                      设置 API Key
                    </button>
                  </p>
                </div>
              )}
              {/* Show current model when configured */}
              {isCurrentProviderConfigured && currentProviderConfig && (
                <p className="mt-1.5 text-xs text-zinc-400 dark:text-zinc-500">
                  当前模型: <span className="font-mono">{currentProviderConfig.model}</span>
                </p>
              )}
            </motion.div>

            {/* Data Source Upload */}
            <motion.div {...fadeSlideUp} transition={{ ...fadeSlideUp.transition, delay: 0.05 }}>
              <label className="block text-sm font-medium text-zinc-700 dark:text-zinc-300 mb-2">
                数据源 <span className="text-rose-500">*</span>
              </label>

              <AnimatePresence mode="wait">
                {dataSource ? (
                  <motion.div
                    key="file-selected"
                    initial={{ opacity: 0, scale: 0.98 }}
                    animate={{ opacity: 1, scale: 1 }}
                    exit={{ opacity: 0, scale: 0.98 }}
                    className="rounded-xl border border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-900 p-3.5"
                  >
                    <div className="flex items-center gap-3">
                      <div className="flex items-center justify-center size-10 rounded-lg bg-emerald-50 dark:bg-emerald-900/30 shrink-0">
                        {getFileIcon(dataSource.name)}
                      </div>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-medium text-zinc-800 dark:text-zinc-200 truncate">
                          {dataSource.name}
                        </p>
                        <p className="text-xs text-zinc-500 dark:text-zinc-400 mt-0.5">
                          {formatFileSize(dataSource.size)}
                        </p>
                      </div>
                      <Button
                        variant="ghost"
                        size="icon"
                        className="size-8 shrink-0 rounded-lg text-zinc-400 hover:text-rose-500 hover:bg-rose-50 dark:hover:bg-rose-900/20 transition-colors"
                        onClick={removeFile}
                      >
                        <X className="size-3.5" />
                      </Button>
                    </div>
                  </motion.div>
                ) : (
                  <motion.div
                    key="upload-area"
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    exit={{ opacity: 0 }}
                  >
                    <div
                      onDragOver={handleDragOver}
                      onDragLeave={handleDragLeave}
                      onDrop={handleDrop}
                      onClick={() => fileInputRef.current?.click()}
                      className={cn(
                        'relative rounded-xl border-2 border-dashed p-6 text-center cursor-pointer transition-all duration-200',
                        isDragOver
                          ? 'border-violet-400 bg-violet-50/60 dark:border-violet-500 dark:bg-violet-900/20 scale-[1.01]'
                          : 'border-zinc-200 dark:border-zinc-700 hover:border-violet-300 hover:bg-violet-50/30 dark:hover:border-violet-600 dark:hover:bg-violet-900/10',
                      )}
                    >
                      <input
                        ref={fileInputRef}
                        type="file"
                        accept=".docx,.xlsx,.xls"
                        onChange={handleInputChange}
                        className="hidden"
                      />

                      <motion.div
                        animate={isDragOver ? { y: -4, scale: 1.05 } : { y: 0, scale: 1 }}
                        transition={{ type: 'spring', stiffness: 300, damping: 20 }}
                        className="flex flex-col items-center gap-2.5"
                      >
                        <div
                          className={cn(
                            'flex items-center justify-center size-12 rounded-xl transition-colors duration-200',
                            isDragOver
                              ? 'bg-violet-100 dark:bg-violet-800/40'
                              : 'bg-zinc-100 dark:bg-zinc-800',
                          )}
                        >
                          <Upload
                            className={cn(
                              'size-5 transition-colors duration-200',
                              isDragOver
                                ? 'text-violet-600 dark:text-violet-300'
                                : 'text-zinc-400 dark:text-zinc-500',
                            )}
                          />
                        </div>
                        <div>
                          <p className="text-sm font-medium text-zinc-700 dark:text-zinc-300">
                            {isDragOver ? '将文件拖放到此处' : '拖放或点击上传'}
                          </p>
                          <p className="text-xs text-zinc-400 dark:text-zinc-500 mt-1">
                            支持 .docx, .xlsx, .xls &mdash; 最大 50MB
                          </p>
                        </div>
                      </motion.div>
                    </div>
                  </motion.div>
                )}
              </AnimatePresence>
            </motion.div>

            {/* Custom Prompt */}
            <motion.div {...fadeSlideUp} transition={{ ...fadeSlideUp.transition, delay: 0.1 }}>
              <label className="block text-sm font-medium text-zinc-700 dark:text-zinc-300 mb-2">
                自定义提示词 <span className="text-xs font-normal text-zinc-400">(可选)</span>
              </label>
              <div className="relative">
                <Textarea
                  value={prompt}
                  onChange={(e) => setPrompt(e.target.value)}
                  placeholder="描述 AI 应如何理解和填充数据到模板..."
                  rows={3}
                  className="resize-none rounded-xl border-zinc-200 dark:border-zinc-700 bg-white dark:bg-zinc-900 text-sm placeholder:text-zinc-400 dark:placeholder:text-zinc-500 focus-visible:ring-violet-500/30 focus-visible:border-violet-400 dark:focus-visible:ring-violet-500/20 dark:focus-visible:border-violet-500 transition-all duration-200"
                />
                <div className="absolute bottom-2.5 right-2.5">
                  <Badge
                    variant="secondary"
                    className="text-[10px] font-normal bg-zinc-100 text-zinc-400 border-0 dark:bg-zinc-800 dark:text-zinc-500"
                  >
                    {prompt.length}/500
                  </Badge>
                </div>
              </div>
            </motion.div>

            {/* Generate Button */}
            <motion.div {...fadeSlideUp} transition={{ ...fadeSlideUp.transition, delay: 0.15 }}>
              <Button
                onClick={handleGenerate}
                disabled={!canGenerate}
                className={cn(
                  'w-full h-11 rounded-xl text-sm font-semibold transition-all duration-300 shadow-md',
                  canGenerate
                    ? 'bg-gradient-to-r from-violet-600 to-fuchsia-600 hover:from-violet-700 hover:to-fuchsia-700 text-white shadow-violet-500/25 hover:shadow-lg hover:shadow-violet-500/30 hover:scale-[1.01] active:scale-[0.99]'
                    : 'bg-zinc-100 dark:bg-zinc-800 text-zinc-400 dark:text-zinc-500 shadow-none cursor-not-allowed',
                )}
              >
                {isGenerating ? (
                  <motion.span
                    initial={{ opacity: 0 }}
                    animate={{ opacity: 1 }}
                    className="flex items-center gap-2"
                  >
                    <Loader2 className="size-4 animate-spin" />
                    使用 {PROVIDERS[selectedProvider].label} 生成中...
                  </motion.span>
                ) : (
                  <span className="flex items-center gap-2">
                    <Sparkles className="size-4" />
                    使用 {PROVIDERS[selectedProvider].label} 生成
                  </span>
                )}
              </Button>
            </motion.div>

            {/* Error Display */}
            <AnimatePresence>
              {hasError && (
                <motion.div
                  initial={{ opacity: 0, y: -8, height: 0 }}
                  animate={{ opacity: 1, y: 0, height: 'auto' }}
                  exit={{ opacity: 0, y: -8, height: 0 }}
                  transition={{ duration: 0.2 }}
                >
                  <div className="rounded-xl border border-rose-200 dark:border-rose-800/60 bg-rose-50 dark:bg-rose-900/20 p-3.5">
                    <div className="flex items-start gap-2.5">
                      <AlertCircle className="size-4 text-rose-500 dark:text-rose-400 shrink-0 mt-0.5" />
                      <div className="min-w-0">
                        <p className="text-sm font-medium text-rose-700 dark:text-rose-300">
                          生成失败
                        </p>
                        <p className="text-xs text-rose-600/80 dark:text-rose-400/80 mt-0.5 break-words">
                          {error}
                        </p>
                      </div>
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {/* Result Section */}
            <AnimatePresence>
              {hasResult && (
                <motion.div
                  initial={{ opacity: 0, y: 12 }}
                  animate={{ opacity: 1, y: 0 }}
                  exit={{ opacity: 0, y: -8 }}
                  transition={{ duration: 0.3, ease: [0.25, 0.46, 0.45, 0.94] }}
                  className="space-y-4"
                >
                  {/* Success Header */}
                  <div className="rounded-xl border border-emerald-200/80 dark:border-emerald-800/50 bg-gradient-to-r from-emerald-50 to-teal-50 dark:from-emerald-900/20 dark:to-teal-900/15 p-4">
                    <div className="flex items-start gap-3">
                      <motion.div
                        initial={{ scale: 0 }}
                        animate={{ scale: 1 }}
                        transition={{ type: 'spring', stiffness: 300, damping: 15, delay: 0.1 }}
                      >
                        <CheckCircle2 className="size-5 text-emerald-600 dark:text-emerald-400 shrink-0 mt-0.5" />
                      </motion.div>
                      <div className="min-w-0 flex-1">
                        <p className="text-sm font-semibold text-emerald-800 dark:text-emerald-300">
                          {result.modifications.length > 0
                            ? `找到 ${result.modifications.length} 项修改`
                            : '未找到匹配的修改'}
                        </p>
                        {result.summary && (
                          <p className="text-xs text-emerald-700/70 dark:text-emerald-400/70 mt-1 leading-relaxed">
                            {result.summary}
                          </p>
                        )}
                      </div>
                      <Badge className="shrink-0 text-[10px] font-medium bg-emerald-100/80 text-emerald-700 border-emerald-200/50 dark:bg-emerald-800/40 dark:text-emerald-300 dark:border-emerald-700/30">
                        AI
                      </Badge>
                    </div>
                  </div>

                  {/* Modification Table */}
                  {result.modifications.length > 0 && (
                    <motion.div {...scaleIn} className="rounded-xl border border-zinc-200 dark:border-zinc-800 overflow-hidden">
                      <div className="bg-zinc-50 dark:bg-zinc-900/50 px-4 py-2.5 border-b border-zinc-200 dark:border-zinc-800">
                        <p className="text-xs font-semibold uppercase tracking-wider text-zinc-500 dark:text-zinc-400">
                          修改详情
                        </p>
                      </div>
                      <div className="divide-y divide-zinc-100 dark:divide-zinc-800/80">
                        <motion.div
                          variants={staggerContainer}
                          initial="initial"
                          animate="animate"
                        >
                          {result.modifications.map((mod) => (
                            <motion.div
                              key={`${mod.slideIndex}-${mod.elementIndex}-${mod.type}`}
                              variants={staggerItem}
                              transition={{ duration: 0.15 }}
                              className="flex items-start gap-3 px-4 py-3 hover:bg-zinc-50/50 dark:hover:bg-zinc-800/30 transition-colors"
                            >
                              <div className="shrink-0 mt-0.5">
                                <Badge
                                  variant="outline"
                                  className={cn(
                                    'text-[10px] font-medium px-1.5 py-0 h-5 rounded-md',
                                    mod.type === 'text'
                                      ? 'border-sky-200 bg-sky-50 text-sky-700 dark:border-sky-800 dark:bg-sky-900/30 dark:text-sky-300'
                                      : 'border-amber-200 bg-amber-50 text-amber-700 dark:border-amber-800 dark:bg-amber-900/30 dark:text-amber-300',
                                  )}
                                >
                                  {mod.type}
                                </Badge>
                              </div>
                              <div className="min-w-0 flex-1">
                                <p className="text-xs text-zinc-500 dark:text-zinc-400">
                                  幻灯片 {mod.slideIndex + 1}
                                </p>
                                <p className="text-sm text-zinc-800 dark:text-zinc-200 mt-0.5 break-words">
                                  {getModLabel(mod)}
                                </p>
                              </div>
                              <div className="shrink-0 text-right max-w-[200px]">
                                <p className="text-xs text-zinc-400 dark:text-zinc-500">→</p>
                                <p className="text-xs font-medium text-violet-700 dark:text-violet-300 mt-0.5 break-words">
                                  {getModNewValue(mod)}
                                </p>
                              </div>
                            </motion.div>
                          ))}
                        </motion.div>
                      </div>
                    </motion.div>
                  )}

                  {/* Action Buttons */}
                  <div className="flex items-center gap-3">
                    {result.modifications.length > 0 && (
                      <Button
                        onClick={handleApply}
                        disabled={isApplied}
                        className={cn(
                          'flex-1 h-10 rounded-xl text-sm font-semibold transition-all duration-300',
                          isApplied
                            ? 'bg-emerald-100 dark:bg-emerald-900/30 text-emerald-700 dark:text-emerald-300 hover:bg-emerald-100 dark:hover:bg-emerald-900/30 cursor-default'
                            : 'bg-gradient-to-r from-violet-600 to-fuchsia-600 hover:from-violet-700 hover:to-fuchsia-700 text-white shadow-md shadow-violet-500/20 hover:shadow-lg hover:shadow-violet-500/25 hover:scale-[1.01] active:scale-[0.99]',
                        )}
                      >
                        {isApplied ? (
                          <span className="flex items-center gap-2">
                            <CheckCircle2 className="size-4" />
                            已应用
                          </span>
                        ) : (
                          <span className="flex items-center gap-2">
                            <Wand2 className="size-4" />
                            应用修改
                          </span>
                        )}
                      </Button>
                    )}
                    <Button
                      variant="outline"
                      onClick={handleReset}
                      className={cn(
                        'h-10 rounded-xl text-sm font-medium transition-all duration-200',
                        'border-zinc-200 dark:border-zinc-700 text-zinc-600 dark:text-zinc-400',
                        'hover:bg-zinc-50 dark:hover:bg-zinc-800 hover:text-zinc-800 dark:hover:text-zinc-200',
                        result.modifications.length === 0 && 'flex-1',
                      )}
                    >
                      <span className="flex items-center gap-2">
                        <Sparkles className="size-3.5" />
                        {hasResult ? '重新生成' : '重试'}
                      </span>
                    </Button>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        </ScrollArea>

        {/* ── Footer ─────────────────────────────────────────────── */}
        <Separator className="bg-zinc-100 dark:bg-zinc-800" />
        <div className="px-6 py-3">
          <p className="text-[11px] text-center text-zinc-400 dark:text-zinc-500">
            AI 生成内容可能不准确，请在应用前仔细检查所有修改
          </p>
        </div>
      </DialogContent>
    </Dialog>
  );
}
