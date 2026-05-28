'use client'

import React from 'react';
import { usePptxStore, type PptxSlideData } from '@/lib/pptx-store';
import { cn } from '@/lib/utils';
import { toast } from 'sonner';
import { motion, AnimatePresence } from 'framer-motion';
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
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
} from 'lucide-react';

interface AiGenerateDialogProps {
  children?: React.ReactNode;
  open?: boolean;
  onOpenChange?: (open: boolean) => void;
}

interface AiModificationResult {
  slideIndex: number;
  elementIndex: number;
  type: 'text' | 'table';
  newText?: string;
  tableCells?: { row: number; col: number; text: string }[];
}

export function AiGenerateDialog({ children, open: controlledOpen, onOpenChange }: AiGenerateDialogProps) {
  const { fileId, slides, applyAiModifications } = usePptxStore();
  const [internalOpen, setInternalOpen] = React.useState(false);
  const isOpen = controlledOpen !== undefined ? controlledOpen : internalOpen;
  const setIsOpen = onOpenChange || setInternalOpen;

  const [dataSource, setDataSource] = React.useState<File | null>(null);
  const [customPrompt, setCustomPrompt] = React.useState('');
  const [isGenerating, setIsGenerating] = React.useState(false);
  const [result, setResult] = React.useState<{
    modifications: AiModificationResult[];
    summary?: string;
  } | null>(null);
  const [error, setError] = React.useState<string | null>(null);
  const [applied, setApplied] = React.useState(false);

  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const ext = file.name.toLowerCase().split('.').pop();
    if (ext !== 'docx' && ext !== 'xlsx' && ext !== 'xls') {
      toast.error('数据源文件仅支持 .docx 和 .xlsx 格式');
      return;
    }
    if (file.size > 50 * 1024 * 1024) {
      toast.error('数据源文件大小不能超过 50MB');
      return;
    }

    setDataSource(file);
    setResult(null);
    setError(null);
    setApplied(false);

    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleGenerate = async () => {
    if (!fileId) {
      toast.error('请先上传 PPTX 文件');
      return;
    }
    if (!dataSource) {
      toast.error('请选择数据源文件');
      return;
    }

    setIsGenerating(true);
    setError(null);
    setResult(null);
    setApplied(false);

    try {
      const formData = new FormData();
      formData.append('fileId', fileId);
      formData.append('dataSource', dataSource);
      if (customPrompt.trim()) {
        formData.append('prompt', customPrompt.trim());
      }

      const response = await fetch('/api/ai/generate', {
        method: 'POST',
        body: formData,
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({ error: 'AI 生成失败' }));
        throw new Error(errorData.error || 'AI 生成失败');
      }

      const data = await response.json();

      if (data.modifications && data.modifications.length > 0) {
        setResult({ modifications: data.modifications, summary: data.summary });
        toast.success(data.summary || `已生成 ${data.modifications.length} 项修改`);
      } else {
        setError('AI 未能从数据源中找到匹配模板字段的内容');
        toast.warning('未找到匹配内容');
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : 'AI 生成失败，请重试';
      setError(message);
      toast.error(message);
    } finally {
      setIsGenerating(false);
    }
  };

  const handleApply = () => {
    if (!result || !result.modifications.length) return;

    applyAiModifications(result.modifications);
    setApplied(true);
    toast.success(`已应用 ${result.modifications.length} 项 AI 修改`);
  };

  const handleClose = () => {
    setIsOpen(false);
    // Don't reset state immediately to allow smooth close animation
    setTimeout(() => {
      if (!applied) {
        setDataSource(null);
        setCustomPrompt('');
        setResult(null);
        setError(null);
        setApplied(false);
      }
    }, 300);
  };

  const handleReset = () => {
    setDataSource(null);
    setCustomPrompt('');
    setResult(null);
    setError(null);
    setApplied(false);
  };

  // Count template fields for info display
  const templateFieldCount = React.useMemo(() => {
    let count = 0;
    for (const slide of slides) {
      for (const el of slide.elements) {
        if (el.type === 'text' || el.type === 'table') count++;
      }
    }
    return count;
  }, [slides]);

  return (
    <Dialog open={isOpen} onOpenChange={setIsOpen}>
      {children && (
        <DialogTrigger asChild>
          {children}
        </DialogTrigger>
      )}
      <DialogContent className="sm:max-w-[560px] max-h-[85vh] flex flex-col p-0">
        <DialogHeader className="px-6 pt-6 pb-0">
          <DialogTitle className="flex items-center gap-2">
            <div className="w-8 h-8 rounded-lg bg-primary/10 flex items-center justify-center">
              <Wand2 className="w-4 h-4 text-primary" />
            </div>
            AI 智能填充
          </DialogTitle>
          <DialogDescription>
            上传数据源文件，AI 将自动匹配模板字段并填充内容
          </DialogDescription>
        </DialogHeader>

        <ScrollArea className="flex-1 px-6 py-4">
          <div className="space-y-5">
            {/* Template info */}
            <div className="flex items-center gap-3 p-3 rounded-lg bg-muted/50">
              <FileText className="w-5 h-5 text-muted-foreground shrink-0" />
              <div className="flex-1 min-w-0">
                <p className="text-xs font-medium">当前模板</p>
                <p className="text-[10px] text-muted-foreground">
                  {slides.length} 页幻灯片 · {templateFieldCount} 个可编辑字段
                </p>
              </div>
            </div>

            {/* Data source upload */}
            <div className="space-y-2">
              <label className="text-xs font-medium">数据源文件</label>
              <input
                ref={fileInputRef}
                type="file"
                accept=".docx,.xlsx,.xls"
                className="hidden"
                onChange={handleFileSelect}
              />

              {dataSource ? (
                <div className="flex items-center gap-2 p-3 rounded-lg border bg-muted/30">
                  {dataSource.name.toLowerCase().endsWith('.docx') ? (
                    <FileText className="w-5 h-5 text-blue-500 shrink-0" />
                  ) : (
                    <FileSpreadsheet className="w-5 h-5 text-emerald-500 shrink-0" />
                  )}
                  <div className="flex-1 min-w-0">
                    <p className="text-xs font-medium truncate">{dataSource.name}</p>
                    <p className="text-[10px] text-muted-foreground">
                      {(dataSource.size / 1024).toFixed(1)} KB
                    </p>
                  </div>
                  <Button
                    variant="ghost"
                    size="sm"
                    className="h-6 w-6 p-0 shrink-0"
                    onClick={() => {
                      setDataSource(null);
                      setResult(null);
                      setError(null);
                    }}
                  >
                    <X className="w-3 h-3" />
                  </Button>
                </div>
              ) : (
                <button
                  className="w-full flex flex-col items-center gap-2 p-6 rounded-lg border-2 border-dashed hover:bg-muted/30 transition-colors cursor-pointer"
                  onClick={() => fileInputRef.current?.click()}
                >
                  <Upload className="w-6 h-6 text-muted-foreground" />
                  <div className="text-center">
                    <p className="text-xs font-medium">点击上传数据源文件</p>
                    <p className="text-[10px] text-muted-foreground mt-0.5">
                      支持 .docx 和 .xlsx 格式，最大 50MB
                    </p>
                  </div>
                </button>
              )}
            </div>

            {/* Custom prompt */}
            <div className="space-y-2">
              <label className="text-xs font-medium">
                自定义指令
                <span className="text-muted-foreground font-normal ml-1">(可选)</span>
              </label>
              <Textarea
                value={customPrompt}
                onChange={(e) => setCustomPrompt(e.target.value)}
                placeholder="例如：请重点关注财务数据部分，忽略日期字段..."
                className="min-h-[60px] text-xs resize-y"
                rows={2}
              />
            </div>

            {/* Error display */}
            <AnimatePresence>
              {error && (
                <motion.div
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  exit={{ opacity: 0, height: 0 }}
                  className="overflow-hidden"
                >
                  <div className="flex items-start gap-2 p-3 rounded-lg bg-destructive/10 text-destructive text-xs">
                    <AlertCircle className="w-4 h-4 shrink-0 mt-0.5" />
                    <p>{error}</p>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {/* Result display */}
            <AnimatePresence>
              {result && (
                <motion.div
                  initial={{ opacity: 0, height: 0 }}
                  animate={{ opacity: 1, height: 'auto' }}
                  exit={{ opacity: 0, height: 0 }}
                  className="overflow-hidden"
                >
                  <div className="space-y-3">
                    {/* Summary */}
                    <div className="flex items-center gap-2 p-3 rounded-lg bg-emerald-50 dark:bg-emerald-900/20 border border-emerald-200 dark:border-emerald-800">
                      <CheckCircle2 className="w-4 h-4 text-emerald-600 dark:text-emerald-400 shrink-0" />
                      <div className="flex-1">
                        <p className="text-xs font-medium text-emerald-700 dark:text-emerald-300">
                          {result.summary || `已生成 ${result.modifications.length} 项修改`}
                        </p>
                      </div>
                    </div>

                    {/* Modification list */}
                    <div className="max-h-48 overflow-y-auto rounded-lg border">
                      <table className="w-full text-[10px]">
                        <thead className="bg-muted/50 sticky top-0">
                          <tr>
                            <th className="px-2 py-1.5 text-left font-medium">页</th>
                            <th className="px-2 py-1.5 text-left font-medium">元素</th>
                            <th className="px-2 py-1.5 text-left font-medium">类型</th>
                            <th className="px-2 py-1.5 text-left font-medium">内容预览</th>
                          </tr>
                        </thead>
                        <tbody>
                          {result.modifications.map((mod, i) => {
                            const slide = slides[mod.slideIndex];
                            const element = slide?.elements.find(
                              (el) => el.slideIndex === mod.slideIndex && el.elementIndex === mod.elementIndex
                            );
                            const preview =
                              mod.type === 'text'
                                ? mod.newText?.substring(0, 40) + (mod.newText && mod.newText.length > 40 ? '...' : '')
                                : `${mod.tableCells?.length || 0} 个单元格`;

                            return (
                              <tr key={i} className="border-t">
                                <td className="px-2 py-1.5">{mod.slideIndex + 1}</td>
                                <td className="px-2 py-1.5">{element?.shapeName || `#${mod.elementIndex}`}</td>
                                <td className="px-2 py-1.5">
                                  <Badge variant="outline" className="text-[8px] px-1 py-0">
                                    {mod.type === 'text' ? '文本' : '表格'}
                                  </Badge>
                                </td>
                                <td className="px-2 py-1.5 truncate max-w-[160px]">{preview}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>

                    {/* Apply button */}
                    <div className="flex items-center gap-2">
                      <Button
                        className="flex-1 gap-2"
                        size="sm"
                        onClick={handleApply}
                        disabled={applied}
                      >
                        {applied ? (
                          <>
                            <CheckCircle2 className="w-3.5 h-3.5" />
                            已应用
                          </>
                        ) : (
                          <>
                            <Sparkles className="w-3.5 h-3.5" />
                            应用修改
                          </>
                        )}
                      </Button>
                      <Button
                        variant="outline"
                        size="sm"
                        onClick={handleReset}
                        disabled={isGenerating}
                      >
                        重新生成
                      </Button>
                    </div>
                  </div>
                </motion.div>
              )}
            </AnimatePresence>

            {/* Generate button */}
            {!result && (
              <Button
                className="w-full gap-2"
                size="sm"
                onClick={handleGenerate}
                disabled={!dataSource || isGenerating}
              >
                {isGenerating ? (
                  <>
                    <Loader2 className="w-3.5 h-3.5 animate-spin" />
                    AI 生成中...
                  </>
                ) : (
                  <>
                    <Sparkles className="w-3.5 h-3.5" />
                    开始生成
                  </>
                )}
              </Button>
            )}
          </div>
        </ScrollArea>

        {/* Footer */}
        <div className="px-6 py-3 border-t bg-muted/30">
          <div className="flex items-center justify-between">
            <p className="text-[10px] text-muted-foreground">
              AI 生成的结果仅供参考，请仔细核对后再应用
            </p>
            <Button variant="ghost" size="sm" className="h-7 text-xs" onClick={handleClose}>
              关闭
            </Button>
          </div>
        </div>
      </DialogContent>
    </Dialog>
  );
}
