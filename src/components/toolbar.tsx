'use client';

import { useState, useCallback } from 'react';
import { usePptxStore, type PptxModification, type PptxImageModification } from '@/lib/pptx-store';
import { toast } from 'sonner';
import { cn } from '@/lib/utils';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Separator } from '@/components/ui/separator';
import {
  Tooltip,
  TooltipTrigger,
  TooltipContent,
  TooltipProvider,
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
  FileText,
  Loader2,
  Settings2,
} from 'lucide-react';

interface ToolbarProps {
  onAiGenerate: () => void;
  onAiSettings?: () => void;
}

export default function Toolbar({ onAiGenerate, onAiSettings }: ToolbarProps) {
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

  const [isSavingJson, setIsSavingJson] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  const modCount = getTotalModificationCount();
  const hasModifications = modCount > 0;

  // ── Back navigation ──────────────────────────────────────────────
  const handleBack = useCallback(() => {
    reset();
    setStep('upload');
  }, [reset, setStep]);

  // ── Reset all modifications ──────────────────────────────────────
  const handleResetModifications = useCallback(() => {
    resetAllModifications();
    toast.success('All modifications have been reset');
  }, [resetAllModifications]);

  // ── Save as JSON ────────────────────────────────────────────────
  const handleSaveJson = useCallback(async () => {
    if (!slides.length) return;

    setIsSavingJson(true);
    try {
      const modifications = getModifications();
      const imageModifications = getImageModifications();

      const jsonData = {
        fileName,
        fileId,
        slideCount: slides.length,
        modifications,
        imageModifications,
        slides: slides.map((slide) => ({
          slideNumber: slide.slideNumber,
          elements: slide.elements.map((el) => ({
            type: el.type,
            id: el.id,
            shapeName: el.shapeName,
            ...(el.type === 'text'
              ? {
                  originalText: el.originalText,
                  currentText: el.currentText,
                  paragraphs: el.paragraphs,
                }
              : {}),
            ...(el.type === 'table'
              ? {
                  rows: el.rows,
                  currentRows: el.currentRows,
                }
              : {}),
            ...(el.type === 'image'
              ? {
                  imageName: el.imageName,
                  imageType: el.imageType,
                }
              : {}),
            position: el.position,
            slideIndex: el.slideIndex,
            elementIndex: el.elementIndex,
          })),
        })),
      };

      const response = await fetch('/api/pptx/save-json', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(jsonData),
      });

      if (!response.ok) throw new Error('Failed to save JSON');

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${fileName?.replace(/\.pptx$/i, '') || 'slide-data'}.json`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      toast.success('JSON file saved successfully');
    } catch (error) {
      console.error('Save JSON error:', error);
      toast.error('Failed to save JSON file');
    } finally {
      setIsSavingJson(false);
    }
  }, [slides, fileName, fileId, getModifications, getImageModifications]);

  // ── Export PPTX ─────────────────────────────────────────────────
  const handleExportPptx = useCallback(async () => {
    if (!fileId) return;

    const modifications = getModifications();
    const imageModifications = getImageModifications();

    if (modifications.length === 0 && imageModifications.length === 0) {
      toast.error('No modifications to export');
      return;
    }

    setIsExporting(true);
    try {
      const response = await fetch('/api/pptx/export', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileId, modifications, imageModifications }),
      });

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({}));
        throw new Error(errorData.error || 'Export failed');
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `${fileName?.replace(/\.pptx$/i, '') || 'modified'}.pptx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      toast.success('PPTX exported successfully');
    } catch (error) {
      console.error('Export error:', error);
      toast.error(error instanceof Error ? error.message : 'Failed to export PPTX');
    } finally {
      setIsExporting(false);
    }
  }, [fileId, fileName, getModifications, getImageModifications]);

  const isBusy = isSavingJson || isExporting;

  return (
    <TooltipProvider delayDuration={300}>
      <div className="border-b border-border/40 px-4 py-2.5 bg-background/95 backdrop-blur-md flex-shrink-0 shadow-sm shadow-black/[0.02] dark:shadow-black/10">
        <div className="flex items-center justify-between gap-4">
          {/* ── Left Section ─────────────────────────────────────── */}
          <div className="flex items-center gap-3 min-w-0">
            {/* Back Button */}
            {hasModifications ? (
              <AlertDialog>
                <Tooltip>
                  <TooltipTrigger asChild>
                    <AlertDialogTrigger asChild>
                      <Button
                        variant="ghost"
                        size="icon"
                        className="size-8 shrink-0 rounded-lg text-muted-foreground hover:text-foreground transition-colors"
                      >
                        <ArrowLeft className="size-4" />
                      </Button>
                    </AlertDialogTrigger>
                  </TooltipTrigger>
                  <TooltipContent side="bottom">
                    <p>Go back</p>
                  </TooltipContent>
                </Tooltip>

                <AlertDialogContent>
                  <AlertDialogHeader>
                    <AlertDialogTitle>Discard changes?</AlertDialogTitle>
                    <AlertDialogDescription>
                      You have {modCount} unsaved modification{modCount !== 1 ? 's' : ''}.
                      Going back will discard all changes. This action cannot be undone.
                    </AlertDialogDescription>
                  </AlertDialogHeader>
                  <AlertDialogFooter>
                    <AlertDialogCancel>Cancel</AlertDialogCancel>
                    <AlertDialogAction
                      onClick={handleBack}
                      className="bg-destructive text-white hover:bg-destructive/90"
                    >
                      Discard & Go Back
                    </AlertDialogAction>
                  </AlertDialogFooter>
                </AlertDialogContent>
              </AlertDialog>
            ) : (
              <Tooltip>
                <TooltipTrigger asChild>
                  <Button
                    variant="ghost"
                    size="icon"
                    className="size-8 shrink-0 rounded-lg text-muted-foreground hover:text-foreground transition-colors"
                    onClick={handleBack}
                  >
                    <ArrowLeft className="size-4" />
                  </Button>
                </TooltipTrigger>
                <TooltipContent side="bottom">
                  <p>Go back</p>
                </TooltipContent>
              </Tooltip>
            )}

            {/* File Name Display */}
            <div className="flex items-center gap-2 min-w-0">
              <div className="flex items-center justify-center size-7 rounded-md bg-muted/50 shrink-0 ring-1 ring-border/30">
                <FileText className="size-3.5 text-muted-foreground" />
              </div>
              <span className="text-sm font-medium truncate max-w-[200px] sm:max-w-[300px] md:max-w-[400px] tracking-tight">
                {fileName || 'Untitled'}
              </span>
            </div>

            {/* Modification Count Badge */}
            {modCount > 0 && (
              <Badge
                className={cn(
                  'shrink-0 rounded-full px-2 py-0 h-5 text-[11px] font-semibold',
                  'bg-amber-100 text-amber-700 border-amber-200/60',
                  'dark:bg-amber-900/40 dark:text-amber-300 dark:border-amber-700/40',
                  'transition-all duration-200'
                )}
              >
                {modCount} edit{modCount !== 1 ? 's' : ''}
              </Badge>
            )}
          </div>

          {/* ── Right Section ────────────────────────────────────── */}
          <div className="flex items-center gap-1.5 shrink-0">
            {/* Reset Modifications */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="sm"
                  className={cn(
                    'h-8 gap-1.5 rounded-lg text-muted-foreground transition-all duration-200',
                    'hover:text-foreground hover:bg-accent/60',
                    !hasModifications && 'opacity-40 pointer-events-none'
                  )}
                  disabled={!hasModifications || isBusy}
                  onClick={handleResetModifications}
                >
                  <RotateCcw className="size-3.5" />
                  <span className="hidden sm:inline">Reset</span>
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Reset all modifications</p>
              </TooltipContent>
            </Tooltip>

            <Separator orientation="vertical" className="h-5 mx-1" />

            {/* Save as JSON */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="sm"
                  className="h-8 gap-1.5 rounded-lg text-muted-foreground hover:text-foreground hover:bg-accent/60 transition-all duration-200"
                  disabled={isBusy || !slides.length}
                  onClick={handleSaveJson}
                >
                  {isSavingJson ? (
                    <Loader2 className="size-3.5 animate-spin" />
                  ) : (
                    <FileJson className="size-3.5" />
                  )}
                  <span className="hidden sm:inline">JSON</span>
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Save as JSON</p>
              </TooltipContent>
            </Tooltip>

            <Separator orientation="vertical" className="h-5 mx-1" />

            {/* AI Settings */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="icon"
                  className="size-8 shrink-0 rounded-lg text-muted-foreground hover:text-violet-600 hover:bg-violet-50 dark:hover:text-violet-400 dark:hover:bg-violet-900/20 transition-all duration-200"
                  onClick={onAiSettings}
                >
                  <Settings2 className="size-4" />
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>AI 模型设置</p>
              </TooltipContent>
            </Tooltip>

            {/* AI Generate */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="outline"
                  size="sm"
                  className={cn(
                    'h-8 gap-1.5 rounded-lg transition-all duration-200',
                    'border-emerald-200/60 bg-emerald-50/50 text-emerald-700',
                    'hover:bg-emerald-100/70 hover:border-emerald-300/70 hover:text-emerald-800',
                    'dark:border-emerald-700/40 dark:bg-emerald-900/20 dark:text-emerald-300',
                    'dark:hover:bg-emerald-900/30 dark:hover:border-emerald-600/50 dark:hover:text-emerald-200'
                  )}
                  disabled={isBusy || !slides.length}
                  onClick={onAiGenerate}
                >
                  <Sparkles className="size-3.5" />
                  <span className="hidden sm:inline">AI Generate</span>
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>AI-powered content generation</p>
              </TooltipContent>
            </Tooltip>

            {/* Export PPTX */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  size="sm"
                  className={cn(
                    'h-8 gap-1.5 rounded-lg transition-all duration-300',
                    'bg-gradient-to-r from-emerald-600 to-teal-600 text-white',
                    'hover:from-emerald-700 hover:to-teal-700',
                    'dark:from-emerald-500 dark:to-teal-500',
                    'dark:hover:from-emerald-600 dark:hover:to-teal-600',
                    'shadow-sm shadow-emerald-500/20 hover:shadow-md hover:shadow-emerald-500/30',
                    'disabled:opacity-50 disabled:cursor-not-allowed',
                    hasModifications && 'animate-export-pulse'
                  )}
                  disabled={isBusy || !hasModifications}
                  onClick={handleExportPptx}
                >
                  {isExporting ? (
                    <Loader2 className="size-3.5 animate-spin" />
                  ) : (
                    <Download className="size-3.5" />
                  )}
                  <span className="hidden sm:inline">Export PPTX</span>
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Export modified PPTX file</p>
              </TooltipContent>
            </Tooltip>
          </div>
        </div>
      </div>

    </TooltipProvider>
  );
}
