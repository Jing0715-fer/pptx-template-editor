'use client';

import React, { useState, useCallback } from 'react';
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
  Undo2,
  Redo2,
  Keyboard,
  Play,
} from 'lucide-react';

interface ToolbarProps {
  onAiGenerate: () => void;
  onAiSettings?: () => void;
  onKeyboardShortcuts?: () => void;
  onSaveJson?: () => void;
  onExportPptx?: () => void;
  onPresent?: () => void;
  onZoomIn?: () => void;
  onZoomOut?: () => void;
  onZoomReset?: () => void;
}

export default function Toolbar({ onAiGenerate, onAiSettings, onKeyboardShortcuts, onSaveJson, onExportPptx, onPresent, onZoomIn, onZoomOut, onZoomReset }: ToolbarProps) {
  const {
    fileName,
    fileId,
    slides,
    step,
    setStep,
    reset,
    currentSlideIndex,
    setCurrentSlide,
    getModifications,
    getImageModifications,
    getTotalModificationCount,
    resetAllModifications,
    undo,
    redo,
    canUndo,
    canRedo,
  } = usePptxStore();

  const [isSavingJson, setIsSavingJson] = useState(false);
  const [isExporting, setIsExporting] = useState(false);

  const modCount = getTotalModificationCount();
  const hasModifications = modCount > 0;
  const hasUndo = canUndo();
  const hasRedo = canRedo();

  // Modification breakdown by type
  const modBreakdown = React.useMemo(() => {
    const mods = getModifications();
    const imgMods = getImageModifications();
    const textMods = mods.filter((m) => m.type === 'text').length;
    const tableCellMods = mods.filter((m) => m.type === 'table').reduce((sum, m) => sum + (m.tableCells?.length ?? 0), 0);
    const imageMods = imgMods.length;
    return { textMods, tableCellMods, imageMods };
  }, [getModifications, getImageModifications]);

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

  // ── Keyboard shortcuts ─────────────────────────────────────────────
  React.useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      // Don't handle slide nav if user is typing in an input/textarea
      const target = e.target as HTMLElement;
      const isInput = target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable;

      if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey) {
        e.preventDefault();
        if (canUndo()) undo();
      }
      if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey))) {
        e.preventDefault();
        if (canRedo()) redo();
      }
      if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        if (slides.length) handleSaveJson();
      }
      if ((e.ctrlKey || e.metaKey) && e.key === 'e') {
        e.preventDefault();
        if (hasModifications && fileId) handleExportPptx();
      }
      // F5 to enter presentation mode
      if (e.key === 'F5') {
        e.preventDefault();
        if (slides.length && onPresent) onPresent();
      }
      // Zoom shortcuts (Ctrl+Plus, Ctrl+Minus, Ctrl+0)
      if ((e.ctrlKey || e.metaKey) && (e.key === '=' || e.key === '+')) {
        e.preventDefault();
        if (onZoomIn) onZoomIn();
      }
      if ((e.ctrlKey || e.metaKey) && e.key === '-') {
        e.preventDefault();
        if (onZoomOut) onZoomOut();
      }
      if ((e.ctrlKey || e.metaKey) && e.key === '0') {
        e.preventDefault();
        if (onZoomReset) onZoomReset();
      }
      // Slide navigation shortcuts (only when not in input fields)
      if (!isInput && !e.ctrlKey && !e.metaKey && !e.altKey) {
        if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
          e.preventDefault();
          if (currentSlideIndex > 0) setCurrentSlide(currentSlideIndex - 1);
        }
        if (e.key === 'ArrowRight' || e.key === 'PageDown') {
          e.preventDefault();
          if (currentSlideIndex < slides.length - 1) setCurrentSlide(currentSlideIndex + 1);
        }
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [undo, redo, canUndo, canRedo, handleSaveJson, handleExportPptx, slides.length, hasModifications, fileId, currentSlideIndex, setCurrentSlide, onPresent, onZoomIn, onZoomOut, onZoomReset]);

  return (
    <TooltipProvider delayDuration={300}>
      <div className="border-b border-border/30 dark:border-b dark:border-white/5 px-2.5 h-9 bg-gradient-to-r from-background via-background to-muted/10 dark:bg-muted/30 backdrop-blur-md flex-shrink-0 flex items-center overflow-x-auto">
        <div className="flex items-center justify-between gap-2 min-w-0">
          {/* ── Left Section ─────────────────────────────────────── */}
          <div className="flex items-center gap-1.5 min-w-0">
            {/* Back Button */}
            {hasModifications ? (
              <AlertDialog>
                <Tooltip>
                  <TooltipTrigger asChild>
                    <AlertDialogTrigger asChild>
                      <Button
                        variant="ghost"
                        size="icon"
                        aria-label="Go back"
                        className="size-7 shrink-0 rounded-md text-muted-foreground hover:text-foreground hover:bg-accent/50 transition-all active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30"
                      >
                        <ArrowLeft className="size-3.5" />
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
                    aria-label="Go back"
                    className="size-7 shrink-0 rounded-md text-muted-foreground hover:text-foreground hover:bg-accent/50 transition-all active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30"
                    onClick={handleBack}
                  >
                    <ArrowLeft className="size-3.5" />
                  </Button>
                </TooltipTrigger>
                <TooltipContent side="bottom">
                  <p>Go back</p>
                </TooltipContent>
              </Tooltip>
            )}

            {/* File Name Display */}
            <div className="flex items-center gap-1.5 min-w-0">
              <div className="flex items-center justify-center size-5 rounded bg-gradient-to-br from-emerald-500/15 to-teal-500/10 dark:from-emerald-500/20 dark:to-teal-500/15 shrink-0">
                <FileText className="size-2.5 text-emerald-600 dark:text-emerald-400" />
              </div>
              <span className="text-xs font-medium truncate max-w-[160px] sm:max-w-[260px] md:max-w-[360px] text-muted-foreground">
                {fileName || 'Untitled'}
              </span>
            </div>

            {/* Modification Count Badge */}
            {modCount > 0 && (
              <Tooltip>
                <TooltipTrigger asChild>
                  <Badge
                    className={cn(
                      'shrink-0 rounded-full px-1.5 py-0 h-4 text-[10px] font-semibold cursor-default',
                      'bg-amber-100/80 text-amber-700 border-amber-200/50',
                      'dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-700/30'
                    )}
                  >
                    {modCount} edit{modCount !== 1 ? 's' : ''}
                  </Badge>
                </TooltipTrigger>
                <TooltipContent side="bottom" className="text-[9px]">
                  <div className="space-y-0.5">
                    {modBreakdown.textMods > 0 && (
                      <span className="flex items-center gap-1"><span className="size-1.5 rounded-full bg-emerald-500" />{modBreakdown.textMods} text modified</span>
                    )}
                    {modBreakdown.tableCellMods > 0 && (
                      <span className="flex items-center gap-1"><span className="size-1.5 rounded-full bg-amber-500" />{modBreakdown.tableCellMods} table cell{modBreakdown.tableCellMods !== 1 ? 's' : ''}</span>
                    )}
                    {modBreakdown.imageMods > 0 && (
                      <span className="flex items-center gap-1"><span className="size-1.5 rounded-full bg-cyan-500" />{modBreakdown.imageMods} image{modBreakdown.imageMods !== 1 ? 's' : ''} replaced</span>
                    )}
                  </div>
                </TooltipContent>
              </Tooltip>
            )}
          </div>

          {/* ── Right Section ────────────────────────────────────── */}
          <div className="flex items-center gap-0.5 shrink-0">
            {/* Undo */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="icon"
                  aria-label="Undo"
                  className={cn(
                    'size-7 shrink-0 rounded-md transition-all active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                    hasUndo ? 'text-muted-foreground hover:text-foreground hover:bg-accent/50 dark:hover:bg-white/10' : 'opacity-35 dark:opacity-60 pointer-events-none'
                  )}
                  disabled={!hasUndo || isBusy}
                  onClick={undo}
                >
                  <Undo2 className="size-3.5" />
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Undo <kbd className="ml-1 text-[9px] px-0.5 py-px rounded bg-muted border border-border/50">Ctrl+Z</kbd></p>
              </TooltipContent>
            </Tooltip>

            {/* Redo */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="icon"
                  aria-label="Redo"
                  className={cn(
                    'size-7 shrink-0 rounded-md transition-all active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                    hasRedo ? 'text-muted-foreground hover:text-foreground hover:bg-accent/50 dark:hover:bg-white/10' : 'opacity-35 dark:opacity-60 pointer-events-none'
                  )}
                  disabled={!hasRedo || isBusy}
                  onClick={redo}
                >
                  <Redo2 className="size-3.5" />
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Redo <kbd className="ml-1 text-[9px] px-0.5 py-px rounded bg-muted border border-border/50">Ctrl+Y</kbd></p>
              </TooltipContent>
            </Tooltip>

            <Separator orientation="vertical" className="h-4 mx-0.5" />

            {/* Reset Modifications */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="sm"
                  className={cn(
                    'h-7 gap-1 rounded-md text-muted-foreground transition-all active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                    'hover:text-foreground hover:bg-accent/50 dark:hover:bg-white/10',
                    !hasModifications && 'opacity-40 dark:opacity-60 pointer-events-none'
                  )}
                  disabled={!hasModifications || isBusy}
                  onClick={handleResetModifications}
                >
                  <RotateCcw className="size-3" />
                  <span className="hidden sm:inline text-[11px]">Reset</span>
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Reset all modifications</p>
              </TooltipContent>
            </Tooltip>

            <Separator orientation="vertical" className="h-4 mx-0.5" />

            {/* Save as JSON */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="sm"
                  className="h-7 gap-1 rounded-md text-muted-foreground hover:text-foreground hover:bg-accent/50 dark:hover:bg-white/10 transition-all active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30"
                  disabled={isBusy || !slides.length}
                  onClick={handleSaveJson}
                >
                  {isSavingJson ? (
                    <Loader2 className="size-3 animate-spin" />
                  ) : (
                    <FileJson className="size-3" />
                  )}
                  <span className="hidden sm:inline text-[11px]">JSON</span>
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Save as JSON</p>
              </TooltipContent>
            </Tooltip>

            <Separator orientation="vertical" className="h-4 mx-0.5" />

            {/* Keyboard Shortcuts */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="icon"
                  className="size-7 shrink-0 rounded-md text-muted-foreground hover:text-foreground hover:bg-accent/50 dark:hover:bg-white/10 transition-all active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30"
                  onClick={onKeyboardShortcuts}
                  aria-label="Keyboard shortcuts"
                >
                  <Keyboard className="size-3.5" />
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Keyboard shortcuts</p>
              </TooltipContent>
            </Tooltip>

            {/* AI Settings */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="icon"
                  className="size-7 shrink-0 rounded-md text-muted-foreground hover:text-violet-600 hover:bg-violet-50/50 dark:hover:text-violet-400 dark:hover:bg-violet-900/20 transition-all active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-violet-500/30"
                  onClick={onAiSettings}
                  aria-label="AI settings"
                >
                  <Settings2 className="size-3.5" />
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>AI Settings</p>
              </TooltipContent>
            </Tooltip>

            {/* AI Generate */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="outline"
                  size="sm"
                  className={cn(
                    'h-7 gap-1 rounded-md transition-all active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                    'border-emerald-200/50 bg-emerald-50/40 text-emerald-700 text-[11px]',
                    'hover:bg-emerald-100/60 hover:border-emerald-300/60',
                    'dark:border-emerald-700/30 dark:bg-emerald-900/15 dark:text-emerald-300',
                    'dark:hover:bg-emerald-900/25 dark:shadow-[0_0_12px_rgba(16,185,129,0.15)]'
                  )}
                  disabled={isBusy || !slides.length}
                  onClick={onAiGenerate}
                >
                  <Sparkles className="size-3" />
                  <span className="hidden sm:inline">AI Generate</span>
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>AI-powered content generation</p>
              </TooltipContent>
            </Tooltip>

            {/* Present */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  size="sm"
                  className={cn(
                    'h-7 gap-1 rounded-md transition-all text-[11px] active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                    'bg-gradient-to-r from-emerald-600 to-teal-600 text-white',
                    'hover:from-emerald-700 hover:to-teal-700',
                    'dark:from-emerald-500 dark:to-teal-500',
                    'dark:hover:from-emerald-600 dark:hover:to-teal-600',
                    'shadow-sm shadow-emerald-500/15 dark:shadow-[0_0_12px_rgba(16,185,129,0.15)]',
                    'disabled:opacity-50 disabled:cursor-not-allowed'
                  )}
                  disabled={!slides.length}
                  onClick={onPresent}
                  aria-label="Enter presentation mode"
                >
                  <Play className="size-3" />
                  <span className="hidden sm:inline">Present</span>
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Enter presentation mode <kbd className="ml-1 text-[9px] px-0.5 py-px rounded bg-muted border border-border/50">F5</kbd></p>
              </TooltipContent>
            </Tooltip>

            {/* Export PPTX */}
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  size="sm"
                  className={cn(
                    'h-7 gap-1 rounded-md transition-all text-[11px] active:scale-95 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                    'bg-gradient-to-r from-emerald-600 to-teal-600 text-white',
                    'hover:from-emerald-700 hover:to-teal-700',
                    'dark:from-emerald-500 dark:to-teal-500',
                    'dark:hover:from-emerald-600 dark:hover:to-teal-600',
                    'shadow-sm shadow-emerald-500/15 dark:shadow-[0_0_12px_rgba(16,185,129,0.15)]',
                    'disabled:opacity-50 disabled:cursor-not-allowed'
                  )}
                  disabled={isBusy || !hasModifications}
                  onClick={handleExportPptx}
                >
                  {isExporting ? (
                    <Loader2 className="size-3 animate-spin" />
                  ) : (
                    <Download className="size-3" />
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
