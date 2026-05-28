'use client'

import React from 'react';
import { usePptxStore, type PptxModification, type PptxImageModification } from '@/lib/pptx-store';
import { cn } from '@/lib/utils';
import { toast } from 'sonner';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Separator } from '@/components/ui/separator';
import {
  Tooltip,
  TooltipContent,
  TooltipProvider,
  TooltipTrigger,
} from '@/components/ui/tooltip';
import {
  AlertDialog,
  AlertDialogAction,
  AlertDialogCancel,
  AlertDialogContent,
  AlertDialogDescription,
  AlertDialogFooter,
  AlertDialogHeader,
  AlertDialogTitle,
  AlertDialogTrigger,
} from '@/components/ui/alert-dialog';
import {
  ArrowLeft,
  Download,
  FileJson,
  RotateCcw,
  Sparkles,
  Save,
  FileText,
  Loader2,
} from 'lucide-react';

interface ToolbarProps {
  onAiGenerate: () => void;
}

export function Toolbar({ onAiGenerate }: ToolbarProps) {
  const {
    fileName,
    fileId,
    slides,
    step,
    setStep,
    reset,
    getModifications,
    getImageModifications,
    getTotalModificationCount,
    resetAllModifications,
  } = usePptxStore();

  const [isExporting, setIsExporting] = React.useState(false);
  const [isSavingJson, setIsSavingJson] = React.useState(false);
  const modCount = getTotalModificationCount();
  const textMods = getModifications();
  const imageMods = getImageModifications();

  const handleBack = () => {
    if (modCount > 0) {
      // Will show confirmation via AlertDialog
      return;
    }
    reset();
  };

  const handleReset = () => {
    resetAllModifications();
    toast.success('已重置所有修改');
  };

  const handleSaveJson = async () => {
    setIsSavingJson(true);
    try {
      const jsonData = {
        fileName,
        fileId,
        slideCount: slides.length,
        modifications: textMods,
        slides: slides.map((slide) => ({
          slideNumber: slide.slideNumber,
          elements: slide.elements.map((el) => {
            const base: Record<string, unknown> = {
              type: el.type,
              id: el.id,
              shapeName: el.shapeName,
              slideIndex: el.slideIndex,
              elementIndex: el.elementIndex,
            };
            if (el.type === 'text') {
              base.originalText = el.originalText;
              if (el.currentText !== undefined) base.currentText = el.currentText;
              base.paragraphs = el.paragraphs;
            } else if (el.type === 'table') {
              base.rows = el.rows;
              if (el.currentRows) base.currentRows = el.currentRows;
            } else if (el.type === 'image') {
              base.imageName = el.imageName;
              base.imageType = el.imageType;
              base.imageRid = el.imageRid;
            }
            return base;
          }),
        })),
      };

      const response = await fetch('/api/pptx/save-json', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(jsonData),
      });

      if (!response.ok) throw new Error('保存失败');

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${fileName?.replace('.pptx', '') || 'slide'}-data.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      toast.success('JSON 数据已保存');
    } catch (err) {
      toast.error('保存 JSON 失败');
    } finally {
      setIsSavingJson(false);
    }
  };

  const handleExport = async () => {
    if (!fileId) {
      toast.error('文件 ID 缺失，请重新上传');
      return;
    }

    const modifications = getModifications();
    const imageModifications = getImageModifications();

    if (modifications.length === 0 && imageModifications.length === 0) {
      toast.error('没有修改内容可导出');
      return;
    }

    setIsExporting(true);

    try {
      // CRITICAL FIX: Pass imageModifications directly.
      // They are already in the correct format with base64 strings.
      // No conversion needed - the replacer handles base64 strings directly.
      const response = await fetch('/api/pptx/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          fileId,
          modifications,
          imageModifications,
        }),
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({ error: '导出失败' }));
        throw new Error(errorData.error || '导出失败');
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${fileName?.replace('.pptx', '') || 'modified'}.pptx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      toast.success(`成功导出 PPTX（${modifications.length} 项文本修改，${imageModifications.length} 项图片修改）`);
    } catch (err) {
      toast.error(err instanceof Error ? err.message : '导出失败，请重试');
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="flex items-center gap-2 px-4 py-2 border-b bg-background/95 backdrop-blur supports-[backdrop-filter]:bg-background/60">
      {/* Left section: Back + file info */}
      <div className="flex items-center gap-2 min-w-0 flex-1">
        {modCount > 0 ? (
          <AlertDialog>
            <AlertDialogTrigger asChild>
              <Button variant="ghost" size="sm" className="h-8 gap-1.5 shrink-0">
                <ArrowLeft className="w-4 h-4" />
                <span className="hidden sm:inline">返回</span>
              </Button>
            </AlertDialogTrigger>
            <AlertDialogContent>
              <AlertDialogHeader>
                <AlertDialogTitle>确认返回？</AlertDialogTitle>
                <AlertDialogDescription>
                  您有 {modCount} 处未保存的修改，返回将丢失所有修改。确定要返回吗？
                </AlertDialogDescription>
              </AlertDialogHeader>
              <AlertDialogFooter>
                <AlertDialogCancel>取消</AlertDialogCancel>
                <AlertDialogAction onClick={reset} className="bg-destructive text-destructive-foreground hover:bg-destructive/90">
                  确认返回
                </AlertDialogAction>
              </AlertDialogFooter>
            </AlertDialogContent>
          </AlertDialog>
        ) : (
          <Button variant="ghost" size="sm" className="h-8 gap-1.5 shrink-0" onClick={reset}>
            <ArrowLeft className="w-4 h-4" />
            <span className="hidden sm:inline">返回</span>
          </Button>
        )}

        <Separator orientation="vertical" className="h-5" />

        <div className="flex items-center gap-2 min-w-0">
          <FileText className="w-4 h-4 text-muted-foreground shrink-0" />
          <span className="text-sm font-medium truncate max-w-[200px]">
            {fileName || '未命名文件'}
          </span>
          {modCount > 0 && (
            <Badge variant="secondary" className="text-[10px] px-1.5 shrink-0 bg-orange-100 text-orange-700 dark:bg-orange-900/30 dark:text-orange-400">
              {modCount} 处修改
            </Badge>
          )}
        </div>
      </div>

      {/* Right section: Actions */}
      <div className="flex items-center gap-1.5">
        {/* Reset */}
        <TooltipProvider>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="ghost"
                size="sm"
                className="h-8 gap-1.5"
                disabled={modCount === 0}
                onClick={handleReset}
              >
                <RotateCcw className="w-3.5 h-3.5" />
                <span className="hidden md:inline">重置</span>
              </Button>
            </TooltipTrigger>
            <TooltipContent>重置所有修改</TooltipContent>
          </Tooltip>
        </TooltipProvider>

        {/* Save JSON */}
        <TooltipProvider>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="ghost"
                size="sm"
                className="h-8 gap-1.5"
                disabled={isSavingJson}
                onClick={handleSaveJson}
              >
                {isSavingJson ? (
                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                ) : (
                  <FileJson className="w-3.5 h-3.5" />
                )}
                <span className="hidden md:inline">JSON</span>
              </Button>
            </TooltipTrigger>
            <TooltipContent>保存为 JSON</TooltipContent>
          </Tooltip>
        </TooltipProvider>

        <Separator orientation="vertical" className="h-5" />

        {/* AI Generate */}
        <TooltipProvider>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="outline"
                size="sm"
                className="h-8 gap-1.5"
                onClick={onAiGenerate}
              >
                <Sparkles className="w-3.5 h-3.5" />
                <span className="hidden sm:inline">AI 填充</span>
              </Button>
            </TooltipTrigger>
            <TooltipContent>AI 智能填充</TooltipContent>
          </Tooltip>
        </TooltipProvider>

        {/* Export PPTX */}
        <TooltipProvider>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="default"
                size="sm"
                className="h-8 gap-1.5"
                disabled={isExporting || (textMods.length === 0 && imageMods.length === 0)}
                onClick={handleExport}
              >
                {isExporting ? (
                  <Loader2 className="w-3.5 h-3.5 animate-spin" />
                ) : (
                  <Download className="w-3.5 h-3.5" />
                )}
                <span className="hidden sm:inline">导出 PPTX</span>
              </Button>
            </TooltipTrigger>
            <TooltipContent>导出修改后的 PPTX 文件</TooltipContent>
          </Tooltip>
        </TooltipProvider>
      </div>
    </div>
  );
}
