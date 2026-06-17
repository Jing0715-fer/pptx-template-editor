'use client';

import { useState, useCallback, useRef, useMemo, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  type DragStartEvent,
  type DragEndEvent,
} from '@dnd-kit/core';
import {
  SortableContext,
  sortableKeyboardCoordinates,
  verticalListSortingStrategy,
  useSortable,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import {
  Copy,
  RotateCcw,
  Type,
  Table2,
  Image as ImageIcon,
  MousePointerClick,
  FileText,
  ChevronDown,
  ChevronRight,
  Upload,
  X,
  ZoomIn,
  Search,
  Layers,
  Replace,
  CaseSensitive,
  Braces,
  GripVertical,
  StickyNote,
  ArrowUp,
  ArrowDown,
  CopyPlus,
  BarChart3,
  Hash,
  AlignLeft,
  Percent,
  FilterX,
  Eye,
  EyeOff,
  MessageSquare,
  ListChecks,
  Check,
  ClipboardList,
} from 'lucide-react';
import { toast } from 'sonner';
import { cn } from '@/lib/utils';
import {
  type PptxTextElement,
  type PptxTableElement,
  type PptxImageElement,
  type PptxElement,
  type PptxSlideData,
  usePptxStore,
} from '@/lib/pptx-store';
import { isEmptyElement } from '@/components/slide-preview';
import {
  TemplateVariableHighlighter,
  VariableChip,
  extractVariables,
  countVariables,
} from '@/components/template-variable-highlighter';
import { Card } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Textarea } from '@/components/ui/textarea';
import { Separator } from '@/components/ui/separator';
import {
  Tooltip,
  TooltipProvider,
  TooltipTrigger,
  TooltipContent,
} from '@/components/ui/tooltip';
import {
  Popover,
  PopoverTrigger,
  PopoverContent,
} from '@/components/ui/popover';

// ============================================================================
// Sortable Element Wrapper
// ============================================================================

function SortableElementWrapper({
  id,
  children,
}: {
  id: string;
  children: (props: {
    isDragging: boolean;
    handleProps: Record<string, unknown>;
  }) => React.ReactNode;
}) {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
    setActivatorNodeRef,
  } = useSortable({ id });

  const style: React.CSSProperties = {
    transform: CSS.Transform.toString(transform),
    transition,
  };

  return (
    <div
      ref={setNodeRef}
      style={style}
      className={cn(
        isDragging && 'opacity-90 z-50'
      )}
    >
      {children({
        isDragging,
        handleProps: {
          ref: setActivatorNodeRef,
          ...listeners,
          ...attributes,
        },
      })}
    </div>
  );
}

// ============================================================================
// Element Ref Registry
// ============================================================================

const elementRefs = new Map<string, HTMLDivElement>();

function registerElementRef(id: string, el: HTMLDivElement | null) {
  if (el) {
    elementRefs.set(id, el);
  } else {
    elementRefs.delete(id);
  }
}

export function scrollToElement(id: string) {
  const el = elementRefs.get(id);
  if (el) {
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    el.classList.add('highlight-flash');
    el.addEventListener('animationend', () => {
      el.classList.remove('highlight-flash');
    }, { once: true });
  }
}

// ============================================================================
// Text Element Editor
// ============================================================================

interface TextElementEditorProps {
  element: PptxTextElement;
  isExpanded: boolean;
  onToggle: () => void;
  isDragging?: boolean;
  dragHandleProps?: Record<string, unknown>;
  batchMode?: boolean;
  isBatchSelected?: boolean;
}

function TextElementEditor({ element, isExpanded, onToggle: _onToggle, isDragging = false, dragHandleProps, batchMode = false, isBatchSelected = false }: TextElementEditorProps) {
  const { updateText, selectElement, selectedElementId, duplicateTextElement, currentSlideIndex, reorderElements, slides, toggleElementVisibility, hiddenElementIds, updateElementComment, elementComments } = usePptxStore();
  const isSelected = selectedElementId === element.id;
  const isHidden = hiddenElementIds.has(element.id);
  const [showOriginal, setShowOriginal] = useState(false);
  const [showDiff, setShowDiff] = useState(false);
  const [showCommentInput, setShowCommentInput] = useState(false);
  const [commentDraft, setCommentDraft] = useState(elementComments[element.id] || '');

  const currentText = element.currentText ?? element.originalText;
  const isModified = element.currentText !== undefined && element.currentText !== element.originalText;
  const variableCount = useMemo(() => countVariables(element.originalText), [element.originalText]);

  // Text stats
  const charCount = currentText.length;
  const wordCount = useMemo(() => currentText.trim().split(/\s+/).filter(Boolean).length, [currentText]);
  const exceedsCharLimit = charCount > 200;

  // Case transform handlers
  const handleUppercase = useCallback(() => {
    updateText(element.id, currentText.toUpperCase());
  }, [updateText, element.id, currentText]);

  const handleLowercase = useCallback(() => {
    updateText(element.id, currentText.toLowerCase());
  }, [updateText, element.id, currentText]);

  const handleTitleCase = useCallback(() => {
    updateText(element.id, currentText.replace(/\w\S*/g, (txt) => txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase()));
  }, [updateText, element.id, currentText]);

  const handleSentenceCase = useCallback(() => {
    const result = currentText.charAt(0).toUpperCase() + currentText.slice(1).replace(/([.!?]\s*)([a-z])/g, (_, p, c) => p + c.toUpperCase());
    updateText(element.id, result);
  }, [updateText, element.id, currentText]);

  const handleTextChange = useCallback(
    (e: React.ChangeEvent<HTMLTextAreaElement>) => {
      updateText(element.id, e.target.value);
    },
    [updateText, element.id]
  );

  const handleReset = useCallback(() => {
    updateText(element.id, element.originalText);
  }, [updateText, element.id, element.originalText]);

  const [showCopyDropdown, setShowCopyDropdown] = useState(false);
  const copyDropdownRef = useRef<HTMLDivElement>(null);

  // Close dropdown on outside click
  useEffect(() => {
    if (!showCopyDropdown) return;
    const handler = (e: MouseEvent) => {
      if (copyDropdownRef.current && !copyDropdownRef.current.contains(e.target as Node)) {
        setShowCopyDropdown(false);
      }
    };
    document.addEventListener('mousedown', handler);
    return () => document.removeEventListener('mousedown', handler);
  }, [showCopyDropdown]);

  const handleCopyFormat = useCallback(async (format: 'plain' | 'markdown' | 'html' | 'variables') => {
    let textToCopy = '';
    const isTitleElement = /title|heading/i.test(element.shapeName);
    switch (format) {
      case 'plain':
        textToCopy = currentText;
        break;
      case 'markdown':
        if (isTitleElement) {
          textToCopy = `**${currentText}**`;
        } else {
          textToCopy = currentText;
        }
        break;
      case 'html':
        if (isTitleElement) {
          textToCopy = `<p style="font-weight:bold">${currentText}</p>`;
        } else {
          textToCopy = `<p>${currentText}</p>`;
        }
        break;
      case 'variables': {
        const vars = extractVariables(element.originalText);
        if (vars.length > 0) {
          textToCopy = currentText.replace(/\{\{\s*([^}]+)\s*\}\}/g, '<<$1>>');
        } else {
          textToCopy = currentText;
        }
        break;
      }
    }
    try {
      await navigator.clipboard.writeText(textToCopy);
      const formatLabels: Record<string, string> = { plain: 'plain text', markdown: 'Markdown', html: 'HTML', variables: 'with variables' };
      toast.success(`Copied as ${formatLabels[format]}!`);
    } catch {
      // fallback
    }
    setShowCopyDropdown(false);
  }, [currentText, element.shapeName, element.originalText, setShowCopyDropdown]);

  // Quick action: move element up/down within same type
  const handleMoveUp = useCallback(() => {
    const slide = slides[currentSlideIndex];
    if (!slide) return;
    const sameTypeEls = slide.elements.filter((el) => el.type === 'text');
    const idxInType = sameTypeEls.findIndex((el) => el.id === element.id);
    if (idxInType <= 0) return;
    const prevEl = sameTypeEls[idxInType - 1];
    const oldIndex = slide.elements.findIndex((el) => el.id === element.id);
    const newIndex = slide.elements.findIndex((el) => el.id === prevEl.id);
    reorderElements(currentSlideIndex, oldIndex, newIndex);
  }, [slides, currentSlideIndex, element.id, reorderElements]);

  const handleMoveDown = useCallback(() => {
    const slide = slides[currentSlideIndex];
    if (!slide) return;
    const sameTypeEls = slide.elements.filter((el) => el.type === 'text');
    const idxInType = sameTypeEls.findIndex((el) => el.id === element.id);
    if (idxInType === -1 || idxInType >= sameTypeEls.length - 1) return;
    const nextEl = sameTypeEls[idxInType + 1];
    const oldIndex = slide.elements.findIndex((el) => el.id === element.id);
    const newIndex = slide.elements.findIndex((el) => el.id === nextEl.id);
    reorderElements(currentSlideIndex, oldIndex, newIndex);
  }, [slides, currentSlideIndex, element.id, reorderElements]);

  const handleDuplicate = useCallback(() => {
    duplicateTextElement(element.id);
  }, [duplicateTextElement, element.id]);

  // Compute move up/down availability
  const canMoveUp = useMemo(() => {
    const slide = slides[currentSlideIndex];
    if (!slide) return false;
    const sameTypeEls = slide.elements.filter((el) => el.type === 'text');
    const idx = sameTypeEls.findIndex((el) => el.id === element.id);
    return idx > 0;
  }, [slides, currentSlideIndex, element.id]);

  const canMoveDown = useMemo(() => {
    const slide = slides[currentSlideIndex];
    if (!slide) return false;
    const sameTypeEls = slide.elements.filter((el) => el.type === 'text');
    const idx = sameTypeEls.findIndex((el) => el.id === element.id);
    return idx !== -1 && idx < sameTypeEls.length - 1;
  }, [slides, currentSlideIndex, element.id]);

  return (
    <Card
      ref={(el) => registerElementRef(element.id, el)}
      className={cn(
        'group relative overflow-hidden transition-all duration-200 p-px gap-px cursor-pointer',
        'border border-border/40 hover:border-emerald-300/40 dark:hover:border-emerald-700/30',
        'hover:shadow-sm hover:shadow-emerald-500/8',
        'hover:scale-[1.005] transition-transform duration-150',
        'hover:border-l-emerald-400/50 dark:hover:border-l-emerald-500/30',
        isModified && 'border-amber-300/60 dark:border-amber-600/30',
        isModified && 'animate-pulse-border-amber',
        isSelected && 'ring-1 ring-emerald-500/30 border-l-2 border-l-emerald-500 bg-emerald-50/30 dark:bg-emerald-950/20',
        isDragging && 'shadow-lg ring-2 ring-emerald-500/30',
        isHidden && 'opacity-60 border-l-2 border-l-amber-400/60 dark:border-l-amber-500/40',
        isBatchSelected && 'bg-emerald-50/40 dark:bg-emerald-950/25 border-emerald-300/50 dark:border-emerald-700/40'
      )}
      onClick={() => { if (!batchMode) selectElement(element.id); }}
    >
      {/* Batch checkbox (visible in batch mode) */}
      {batchMode && (
        <div
          className="absolute left-1 top-1 z-10"
          onClick={(e) => { e.stopPropagation(); toggleBatchSelect(element.id); }}
        >
          <div className={cn(
            'flex items-center justify-center size-3 rounded-sm border transition-all cursor-pointer',
            isBatchSelected
              ? 'bg-emerald-500 border-emerald-500 text-white'
              : 'border-muted-foreground/30 hover:border-muted-foreground/60'
          )}>
            {isBatchSelected && <Check className="size-2" strokeWidth={3} />}
          </div>
        </div>
      )}
      {/* Quick-action bar (visible when selected) */}
      <AnimatePresence>
        {isSelected && (
          <motion.div
            initial={{ opacity: 0, x: 8 }}
            animate={{ opacity: 1, x: 0 }}
            exit={{ opacity: 0, x: 8 }}
            transition={{ duration: 0.15 }}
            className="absolute -top-1 -right-1 z-10 flex items-center gap-px bg-background/90 backdrop-blur-sm border border-border/40 rounded-md shadow-sm px-0.5 py-px"
            onClick={(e) => e.stopPropagation()}
          >
            {canMoveUp && (
              <TooltipProvider delayDuration={300}>
                <Tooltip>
                  <TooltipTrigger asChild>
                    <Button
                      variant="ghost"
                      size="icon"
                      className="size-4 text-muted-foreground hover:text-foreground"
                      onClick={handleMoveUp}
                    >
                      <ArrowUp className="size-2" />
                    </Button>
                  </TooltipTrigger>
                  <TooltipContent side="top"><p>Move up</p></TooltipContent>
                </Tooltip>
              </TooltipProvider>
            )}
            {canMoveDown && (
              <TooltipProvider delayDuration={300}>
                <Tooltip>
                  <TooltipTrigger asChild>
                    <Button
                      variant="ghost"
                      size="icon"
                      className="size-4 text-muted-foreground hover:text-foreground"
                      onClick={handleMoveDown}
                    >
                      <ArrowDown className="size-2" />
                    </Button>
                  </TooltipTrigger>
                  <TooltipContent side="top"><p>Move down</p></TooltipContent>
                </Tooltip>
              </TooltipProvider>
            )}
            <TooltipProvider delayDuration={300}>
              <Tooltip>
                <TooltipTrigger asChild>
                  <Button
                    variant="ghost"
                    size="icon"
                    className="size-4 text-muted-foreground hover:text-emerald-600 dark:hover:text-emerald-400"
                    onClick={handleDuplicate}
                  >
                    <CopyPlus className="size-2" />
                  </Button>
                </TooltipTrigger>
                <TooltipContent side="top"><p>Duplicate</p></TooltipContent>
              </Tooltip>
            </TooltipProvider>
          </motion.div>
        )}
      </AnimatePresence>
      {/* Row 1: Drag handle + Icon + Name + Modified badge + action buttons */}
      <div className="flex items-center gap-1.5 px-2 py-1">
        {/* Drag handle */}
        <div
          {...dragHandleProps}
          className={cn(
            'shrink-0 cursor-grab active:cursor-grabbing rounded',
            'opacity-0 group-hover:opacity-100 transition-opacity duration-150',
            'hover:bg-muted/40 p-0.5 -m-0.5',
            dragHandleProps && 'opacity-60'
          )}
          onClick={(e) => e.stopPropagation()}
        >
          <GripVertical className="size-2 text-muted-foreground/50" />
        </div>

        <div className={cn(
          'flex size-4 shrink-0 items-center justify-center rounded',
          'bg-emerald-500/10 text-emerald-600 dark:bg-emerald-500/15 dark:text-emerald-400'
        )}>
          <Type className="size-2" />
        </div>

        <div className="flex-1 min-w-0 flex items-center gap-1">
          <span className="text-[10px] font-medium truncate">{element.shapeName}</span>
          {variableCount > 0 && (
            <Badge
              className={cn(
                'shrink-0 rounded-full px-0.5 py-0 h-3 text-[7px] font-semibold',
                'bg-emerald-100/80 text-emerald-700 border-emerald-200/50',
                'dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-700/30'
              )}
            >
              {variableCount} var{variableCount !== 1 ? 's' : ''}
            </Badge>
          )}
          {isModified && (
            <Badge
              className={cn(
                'shrink-0 rounded-full px-0.5 py-0 h-3 text-[7px] font-semibold',
                'bg-amber-100/80 text-amber-700 border-amber-200/50',
                'dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-700/30'
              )}
            >
              Mod
            </Badge>
          )}
        </div>

        <div className="flex items-center gap-px shrink-0">
          <TooltipProvider delayDuration={300}>
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="icon"
                  className={cn(
                    'size-3.5',
                    isHidden
                      ? 'text-amber-500/70 hover:text-amber-500'
                      : 'text-muted-foreground/50 hover:text-muted-foreground'
                  )}
                  onClick={(e) => { e.stopPropagation(); toggleElementVisibility(element.id); }}
                >
                  {isHidden ? <EyeOff className="size-2.5" /> : <Eye className="size-2.5" />}
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>{isHidden ? 'Show in preview' : 'Hide in preview'}</p>
              </TooltipContent>
            </Tooltip>
          </TooltipProvider>

          {/* Comment toggle */}
          <TooltipProvider delayDuration={300}>
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="icon"
                  className={cn(
                    'size-3.5',
                    elementComments[element.id]
                      ? 'text-amber-500 hover:text-amber-600 dark:text-amber-400 dark:hover:text-amber-300'
                      : 'text-muted-foreground/40 hover:text-muted-foreground'
                  )}
                  onClick={(e) => { e.stopPropagation(); setShowCommentInput(v => !v); }}
                >
                  {elementComments[element.id]
                    ? <MessageSquare className="size-2" fill="currentColor" />
                    : <MessageSquare className="size-2" />
                  }
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>{elementComments[element.id] ? 'Edit comment' : 'Add comment'}</p>
              </TooltipContent>
            </Tooltip>
          </TooltipProvider>

          <div className="relative" ref={copyDropdownRef}>
            <TooltipProvider delayDuration={300}>
              <Tooltip>
                <TooltipTrigger asChild>
                  <Button
                    variant="ghost"
                    size="icon"
                    className={cn(
                      'size-3.5 text-muted-foreground hover:text-foreground',
                      showCopyDropdown && 'text-emerald-600 dark:text-emerald-400 bg-emerald-500/10'
                    )}
                    onClick={(e) => { e.stopPropagation(); setShowCopyDropdown((v) => !v); }}
                  >
                    <Copy className="size-1.5" />
                  </Button>
                </TooltipTrigger>
                <TooltipContent side="bottom">
                  <p>Copy text</p>
                </TooltipContent>
              </Tooltip>
            </TooltipProvider>
            <AnimatePresence>
              {showCopyDropdown && (
                <motion.div
                  initial={{ opacity: 0, scale: 0.95, y: -2 }}
                  animate={{ opacity: 1, scale: 1, y: 0 }}
                  exit={{ opacity: 0, scale: 0.95, y: -2 }}
                  transition={{ duration: 0.1 }}
                  className="absolute right-0 top-full z-20 mt-1 min-w-[120px] rounded-md border border-border/40 bg-background/95 backdrop-blur-sm shadow-md py-0.5"
                  onClick={(e) => e.stopPropagation()}
                >
                  {[
                    { key: 'plain' as const, label: 'Copy plain text' },
                    { key: 'markdown' as const, label: 'Copy as Markdown' },
                    { key: 'html' as const, label: 'Copy as HTML' },
                    { key: 'variables' as const, label: 'Copy with variables' },
                  ].map((item) => (
                    <button
                      key={item.key}
                      onClick={() => handleCopyFormat(item.key)}
                      className="flex w-full items-center h-5 px-2 text-[8px] text-muted-foreground hover:text-foreground hover:bg-emerald-500/10 transition-colors focus-visible:outline-none"
                    >
                      {item.label}
                    </button>
                  ))}
                </motion.div>
              )}
            </AnimatePresence>
          </div>

          {isModified && (
            <TooltipProvider delayDuration={300}>
              <Tooltip>
                <TooltipTrigger asChild>
                  <Button
                    variant="ghost"
                    size="icon"
                    className={cn(
                      'size-3.5',
                      'text-amber-600 hover:text-amber-700 dark:text-amber-400 dark:hover:text-amber-300'
                    )}
                    onClick={handleReset}
                  >
                    <RotateCcw className="size-1.5" />
                  </Button>
                </TooltipTrigger>
                <TooltipContent side="bottom">
                  <p>Reset to original</p>
                </TooltipContent>
              </Tooltip>
            </TooltipProvider>
          )}
        </div>
      </div>

      {/* Row 2: Inline textarea (always visible) */}
      <div className="px-2 pb-1">
        <Textarea
          value={currentText}
          onChange={handleTextChange}
          rows={1}
          className={cn(
            'resize-y text-[10px] leading-tight min-h-[22px] py-0.5',
            'focus-visible:ring-emerald-500/20 focus-visible:border-emerald-300/40 dark:focus-visible:ring-emerald-400/25 dark:focus-visible:border-emerald-500/40',
            'bg-muted/15 dark:bg-muted/20 hover:bg-muted/25 dark:hover:bg-muted/25 transition-colors',
            isModified && 'border-amber-300/60 dark:border-amber-500/25 bg-amber-50/20 dark:bg-amber-900/10'
          )}
        />
        {/* Diff indicator - collapsible */}
        {isModified && (
          <div className="flex items-center gap-0.5 mt-0.5">
            <button
              onClick={() => setShowDiff((v) => !v)}
              className="text-[7px] text-amber-500/60 hover:text-amber-500 transition-colors focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-amber-500/30 rounded px-0.5"
            >
              diff
            </button>
            <AnimatePresence>
              {showDiff && (
                <motion.div
                  initial={{ opacity: 0, width: 0 }}
                  animate={{ opacity: 1, width: 'auto' }}
                  exit={{ opacity: 0, width: 0 }}
                  transition={{ duration: 0.12 }}
                  className="overflow-hidden whitespace-nowrap"
                >
                  <span className="text-[8px]">
                    <span className="line-through text-muted-foreground/40">{element.originalText.length > 30 ? element.originalText.slice(0, 30) + '…' : element.originalText}</span>
                    <span className="mx-0.5 text-amber-500/50">→</span>
                    <span className="text-foreground/70">{currentText.length > 30 ? currentText.slice(0, 30) + '…' : currentText}</span>
                  </span>
                </motion.div>
              )}
            </AnimatePresence>
          </div>
        )}
      </div>

      {/* Mini formatting toolbar (always visible with textarea) */}
      <div className="flex items-center gap-0.5 px-2 py-0.5 bg-muted/15 rounded-b-md border-t border-border/10">
          {/* Case transform buttons */}
          <button
            onClick={handleUppercase}
            className="h-4 px-1 text-[7px] font-mono bg-muted/20 hover:bg-muted/40 rounded transition-all duration-100 hover:text-emerald-600 dark:hover:text-emerald-400 active:scale-90 active:bg-emerald-500/20"
            title="UPPERCASE"
          >
            AA
          </button>
          <button
            onClick={handleLowercase}
            className="h-4 px-1 text-[7px] font-mono bg-muted/20 hover:bg-muted/40 rounded transition-all duration-100 hover:text-emerald-600 dark:hover:text-emerald-400 active:scale-90 active:bg-emerald-500/20"
            title="lowercase"
          >
            aa
          </button>
          <button
            onClick={handleTitleCase}
            className="h-4 px-1 text-[7px] font-mono bg-muted/20 hover:bg-muted/40 rounded transition-all duration-100 hover:text-emerald-600 dark:hover:text-emerald-400 active:scale-90 active:bg-emerald-500/20"
            title="Title Case"
          >
            Aa
          </button>
          <button
            onClick={handleSentenceCase}
            className="h-4 px-1 text-[7px] font-mono bg-muted/20 hover:bg-muted/40 rounded transition-all duration-100 hover:text-emerald-600 dark:hover:text-emerald-400 active:scale-90 active:bg-emerald-500/20"
            title="Sentence case"
          >
            Aa.
          </button>

          <Separator orientation="vertical" className="h-2.5 mx-0.5" />

          {/* Text stats */}
          <div className="flex items-center gap-1 text-[7px] text-muted-foreground ml-auto">
            {exceedsCharLimit && (
              <span className="flex items-center gap-0.5 text-amber-500">
                <span className="inline-block size-1 rounded-full bg-amber-500" />
                {charCount} chars
              </span>
            )}
            <span>{charCount} chars</span>
            <span>{wordCount} words</span>
          </div>
        </div>

      {/* Show Original toggle + display */}
      <div className="px-2 pb-1">
        <Button
          variant="ghost"
          size="sm"
          className="h-3 gap-0.5 text-[7px] text-muted-foreground hover:text-emerald-600 dark:hover:text-emerald-400 px-0.5"
          onClick={() => setShowOriginal((v) => !v)}
        >
          <FileText className="size-1.5" />
          {showOriginal ? 'Hide' : 'Orig'}
        </Button>

        <AnimatePresence>
          {showOriginal && (
            <motion.div
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: 'auto', opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              transition={{ duration: 0.15 }}
              className="overflow-hidden"
            >
              <div className={cn(
                'rounded border p-1 text-[10px] leading-relaxed mt-0.5',
                'bg-muted/25 border-border/40 text-muted-foreground'
              )}>
                <span className="text-[8px] font-semibold uppercase tracking-wider text-muted-foreground/60 block mb-0.5">
                  Original
                </span>
                <TemplateVariableHighlighter text={element.originalText} />
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>

      {/* Comment input */}
      <AnimatePresence>
        {showCommentInput && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.15 }}
            className="overflow-hidden"
          >
            <div className="px-2 pb-1">
              <div className="flex items-center gap-1 rounded border bg-amber-50/30 dark:bg-amber-900/10 border-amber-200/30 dark:border-amber-700/20 px-1.5 py-0.5">
                <MessageSquare className="size-2 text-amber-500 shrink-0" />
                <input
                  type="text"
                  value={commentDraft}
                  onChange={(e) => setCommentDraft(e.target.value)}
                  onBlur={() => {
                    updateElementComment(element.id, commentDraft);
                    if (!commentDraft.trim()) setShowCommentInput(false);
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      updateElementComment(element.id, commentDraft);
                      if (!commentDraft.trim()) setShowCommentInput(false);
                    }
                    if (e.key === 'Escape') {
                      setCommentDraft(elementComments[element.id] || '');
                      setShowCommentInput(false);
                    }
                  }}
                  placeholder="Add a note..."
                  className="flex-1 h-5 bg-transparent text-[8px] outline-none placeholder:text-amber-400/40 dark:placeholder:text-amber-500/30"
                  autoFocus
                />
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </Card>
  );
}

// ============================================================================
// Table Element Editor
// ============================================================================

interface TableElementEditorProps {
  element: PptxTableElement;
  isExpanded: boolean;
  onToggle: () => void;
  isDragging?: boolean;
  dragHandleProps?: Record<string, unknown>;
  batchMode?: boolean;
  isBatchSelected?: boolean;
}

function TableElementEditor({ element, isExpanded, onToggle, isDragging = false, dragHandleProps, batchMode = false, isBatchSelected = false }: TableElementEditorProps) {
  const { updateTableCell, selectElement, selectedElementId, toggleElementVisibility, hiddenElementIds, updateElementComment, elementComments } = usePptxStore();
  const isSelected = selectedElementId === element.id;
  const isHidden = hiddenElementIds.has(element.id);
  const [showOriginal, setShowOriginal] = useState(false);
  const [showCommentInput, setShowCommentInput] = useState(false);
  const [commentDraft, setCommentDraft] = useState(elementComments[element.id] || '');

  const currentRows = element.currentRows ?? element.rows;
  const isModified = element.currentRows !== undefined;

  const modifiedCellCount = useMemo(() => {
    if (!element.currentRows) return 0;
    let count = 0;
    for (let ri = 0; ri < element.currentRows.length; ri++) {
      const origRow = element.rows[ri];
      const curRow = element.currentRows[ri];
      if (!origRow || !curRow) continue;
      for (let ci = 0; ci < curRow.cells.length; ci++) {
        const origCell = origRow.cells[ci];
        const curCell = curRow.cells[ci];
        if (!origCell || !curCell) continue;
        if (curCell.text !== origCell.text) count++;
      }
    }
    return count;
  }, [element.currentRows, element.rows]);

  const totalCells = useMemo(
    () => element.rows.reduce((sum, row) => sum + row.cells.length, 0),
    [element.rows]
  );

  const handleCellChange = useCallback(
    (rowIndex: number, colIndex: number, text: string) => {
      updateTableCell(element.id, rowIndex, colIndex, text);
    },
    [updateTableCell, element.id]
  );

  const handleReset = useCallback(() => {
    for (let ri = 0; ri < element.rows.length; ri++) {
      for (let ci = 0; ci < element.rows[ri].cells.length; ci++) {
        const origText = element.rows[ri].cells[ci].text;
        const curText = currentRows[ri]?.cells[ci]?.text;
        if (curText !== origText) {
          updateTableCell(element.id, ri, ci, origText);
        }
      }
    }
  }, [element.id, element.rows, currentRows, updateTableCell]);

  const handleCopyTable = useCallback(async () => {
    const text = currentRows
      .map((row) => row.cells.map((cell) => cell.text).join('\t'))
      .join('\n');
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      // fallback
    }
  }, [currentRows]);

  const isCellModified = useCallback(
    (rowIndex: number, colIndex: number) => {
      if (!element.currentRows) return false;
      const origCell = element.rows[rowIndex]?.cells[colIndex];
      const curCell = element.currentRows[rowIndex]?.cells[colIndex];
      if (!origCell || !curCell) return false;
      return curCell.text !== origCell.text;
    },
    [element.currentRows, element.rows]
  );

  return (
    <Card
      ref={(el) => registerElementRef(element.id, el)}
      className={cn(
        'group overflow-hidden transition-all duration-200 p-px gap-px cursor-pointer',
        'border border-border/40 hover:border-emerald-300/40 dark:hover:border-emerald-700/30',
        'hover:shadow-sm hover:shadow-emerald-500/8',
        'hover:scale-[1.005] transition-transform duration-150',
        'hover:border-l-emerald-400/50 dark:hover:border-l-emerald-500/30',
        isModified && 'border-amber-300/60 dark:border-amber-500/25',
        isModified && 'animate-pulse-border-amber',
        isExpanded && 'shadow-sm shadow-emerald-500/5',
        isSelected && 'ring-1 ring-emerald-500/30 border-l-2 border-l-emerald-500 bg-emerald-50/30 dark:bg-emerald-950/20',
        isDragging && 'shadow-lg ring-2 ring-emerald-500/30',
        isHidden && 'opacity-60 border-l-2 border-l-amber-400/60 dark:border-l-amber-500/40',
        isBatchSelected && 'bg-emerald-50/40 dark:bg-emerald-950/25 border-emerald-300/50 dark:border-emerald-700/40'
      )}
      onClick={() => { if (!batchMode) selectElement(element.id); }}
    >
      {/* Batch checkbox (visible in batch mode) */}
      {batchMode && (
        <div
          className="absolute left-1 top-1.5 z-10"
          onClick={(e) => { e.stopPropagation(); toggleBatchSelect(element.id); }}
        >
          <div className={cn(
            'flex items-center justify-center size-3 rounded-sm border transition-all cursor-pointer',
            isBatchSelected
              ? 'bg-emerald-500 border-emerald-500 text-white'
              : 'border-muted-foreground/30 hover:border-muted-foreground/60'
          )}>
            {isBatchSelected && <Check className="size-2" strokeWidth={3} />}
          </div>
        </div>
      )}
      {/* Header */}
      <button
        onClick={(e) => { e.stopPropagation(); onToggle(); }}
        className="flex w-full items-center gap-1.5 px-2 py-1.5 text-left transition-colors hover:bg-emerald-50/20 dark:hover:bg-emerald-950/10"
    >
        {/* Drag handle */}
        <div
          {...dragHandleProps}
          className={cn(
            'shrink-0 cursor-grab active:cursor-grabbing rounded',
            'opacity-0 group-hover:opacity-100 transition-opacity duration-150',
            'hover:bg-muted/40 p-0.5 -m-0.5',
            dragHandleProps && 'opacity-60'
          )}
          onClick={(e) => e.stopPropagation()}
        >
          <GripVertical className="size-2 text-muted-foreground/50" />
        </div>

        <div className={cn(
          'flex size-5 shrink-0 items-center justify-center rounded',
          'bg-amber-500/10 text-amber-600 dark:bg-amber-500/15 dark:text-amber-400'
        )}>
          <Table2 className="size-2.5" />
        </div>

        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-1">
            <span className="text-[10px] font-medium truncate">{element.shapeName}</span>
            {isModified && (
              <Badge
                className={cn(
                  'shrink-0 rounded-full px-0.5 py-0 h-3 text-[7px] font-semibold',
                  'bg-amber-100/80 text-amber-700 border-amber-200/50',
                  'dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-700/30'
                )}
              >
                {modifiedCellCount}cell{modifiedCellCount !== 1 ? 's' : ''}
              </Badge>
            )}
          </div>
          <p className="text-[8px] text-muted-foreground truncate mt-px">
            {element.rows.length}×{element.rows[0]?.cells.length ?? 0} · {totalCells} cells
          </p>
        </div>

        <motion.div
          animate={{ rotate: isExpanded ? 0 : -90 }}
          transition={{ duration: 0.15 }}
          className="shrink-0 text-muted-foreground"
        >
          <ChevronDown className="size-3" />
        </motion.div>

        {/* Visibility toggle */}
        <TooltipProvider delayDuration={300}>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="ghost"
                size="icon"
                className={cn(
                  'size-4',
                  isHidden
                    ? 'text-amber-500/70 hover:text-amber-500'
                    : 'text-muted-foreground/50 hover:text-muted-foreground'
                )}
                onClick={(e) => { e.stopPropagation(); toggleElementVisibility(element.id); }}
              >
                {isHidden ? <EyeOff className="size-2.5" /> : <Eye className="size-2.5" />}
              </Button>
            </TooltipTrigger>
            <TooltipContent side="bottom">
              <p>{isHidden ? 'Show in preview' : 'Hide in preview'}</p>
            </TooltipContent>
          </Tooltip>
        </TooltipProvider>

        {/* Comment toggle */}
        <TooltipProvider delayDuration={300}>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="ghost"
                size="icon"
                className={cn(
                  'size-4',
                  elementComments[element.id]
                    ? 'text-amber-500 hover:text-amber-600 dark:text-amber-400 dark:hover:text-amber-300'
                    : 'text-muted-foreground/40 hover:text-muted-foreground'
                )}
                onClick={(e) => { e.stopPropagation(); setShowCommentInput(v => !v); }}
              >
                {elementComments[element.id]
                  ? <MessageSquare className="size-2" fill="currentColor" />
                  : <MessageSquare className="size-2" />
                }
              </Button>
            </TooltipTrigger>
            <TooltipContent side="bottom">
              <p>{elementComments[element.id] ? 'Edit comment' : 'Add comment'}</p>
            </TooltipContent>
          </Tooltip>
        </TooltipProvider>
      </button>

      {/* Comment input (outside header button) */}
      <AnimatePresence>
        {showCommentInput && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.15 }}
            className="overflow-hidden"
          >
            <div className="px-2 pb-1">
              <div className="flex items-center gap-1 rounded border bg-amber-50/30 dark:bg-amber-900/10 border-amber-200/30 dark:border-amber-700/20 px-1.5 py-0.5">
                <MessageSquare className="size-2 text-amber-500 shrink-0" />
                <input
                  type="text"
                  value={commentDraft}
                  onChange={(e) => setCommentDraft(e.target.value)}
                  onBlur={() => {
                    updateElementComment(element.id, commentDraft);
                    if (!commentDraft.trim()) setShowCommentInput(false);
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      updateElementComment(element.id, commentDraft);
                      if (!commentDraft.trim()) setShowCommentInput(false);
                    }
                    if (e.key === 'Escape') {
                      setCommentDraft(elementComments[element.id] || '');
                      setShowCommentInput(false);
                    }
                  }}
                  placeholder="Add a note..."
                  className="flex-1 h-5 bg-transparent text-[8px] outline-none placeholder:text-amber-400/40 dark:placeholder:text-amber-500/30"
                  autoFocus
                />
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Expandable content */}
      <AnimatePresence initial={false}>
        {isExpanded && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.2, ease: [0.4, 0, 0.2, 1] }}
            className="overflow-hidden"
          >
            <div className="px-2 pb-1.5 space-y-1">
              <Separator className="opacity-30" />

              {/* Action buttons */}
              <div className="flex items-center gap-0.5">
                <TooltipProvider delayDuration={300}>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Button
                        variant="ghost"
                        size="sm"
                        className="h-5 gap-1 text-[9px] text-muted-foreground hover:text-foreground px-1"
                        onClick={handleCopyTable}
                      >
                        <Copy className="size-2" />
                        Copy
                      </Button>
                    </TooltipTrigger>
                    <TooltipContent side="bottom">
                      <p>Copy table as tab-separated</p>
                    </TooltipContent>
                  </Tooltip>
                </TooltipProvider>

                <TooltipProvider delayDuration={300}>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Button
                        variant="ghost"
                        size="sm"
                        className="h-5 gap-1 text-[9px] text-muted-foreground hover:text-foreground px-1"
                        onClick={() => setShowOriginal((v) => !v)}
                      >
                        <FileText className="size-2" />
                        {showOriginal ? 'Hide' : 'Original'}
                      </Button>
                    </TooltipTrigger>
                    <TooltipContent side="bottom">
                      <p>Toggle original table</p>
                    </TooltipContent>
                  </Tooltip>
                </TooltipProvider>

                {isModified && (
                  <TooltipProvider delayDuration={300}>
                    <Tooltip>
                      <TooltipTrigger asChild>
                        <Button
                          variant="ghost"
                          size="sm"
                          className={cn(
                            'h-5 gap-1 text-[9px] px-1',
                            'text-amber-600 hover:text-amber-700 dark:text-amber-400 dark:hover:text-amber-300'
                          )}
                          onClick={handleReset}
                        >
                          <RotateCcw className="size-2" />
                          Reset
                        </Button>
                      </TooltipTrigger>
                      <TooltipContent side="bottom">
                        <p>Reset all cells to original</p>
                      </TooltipContent>
                    </Tooltip>
                  </TooltipProvider>
                )}
              </div>

              {/* Editable table grid */}
              <div className="overflow-x-auto custom-scrollbar rounded border border-border/30">
                <table className="w-full text-[10px] border-collapse">
                  <tbody>
                    {currentRows.map((row, ri) => (
                      <tr key={ri} className="border-b border-border/20 last:border-b-0">
                        {row.cells.map((cell, ci) => {
                          const cellModified = isCellModified(ri, ci);
                          return (
                            <td
                              key={ci}
                              className={cn(
                                'border-r border-border/20 last:border-r-0 p-0',
                                cellModified && 'bg-amber-50/60 dark:bg-amber-900/15'
                              )}
                              style={{
                                rowSpan: cell.rowSpan > 1 ? cell.rowSpan : undefined,
                                colSpan: cell.colSpan > 1 ? cell.colSpan : undefined,
                              }}
                            >
                              <input
                                type="text"
                                value={cell.text}
                                onChange={(e) => handleCellChange(ri, ci, e.target.value)}
                                className={cn(
                                  'w-full px-1 py-0.5 text-[10px] bg-transparent outline-none',
                                  'placeholder:text-muted-foreground/50',
                                  'focus:bg-emerald-50/30 dark:focus:bg-emerald-950/20 focus:ring-1 focus:ring-emerald-500/15 transition-all',
                                  cellModified && 'text-amber-700 dark:text-amber-300'
                                )}
                                placeholder="—"
                              />
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {/* Original table display */}
              <AnimatePresence>
                {showOriginal && (
                  <motion.div
                    initial={{ height: 0, opacity: 0 }}
                    animate={{ height: 'auto', opacity: 1 }}
                    exit={{ height: 0, opacity: 0 }}
                    transition={{ duration: 0.15 }}
                    className="overflow-hidden"
                  >
                    <div className="overflow-x-auto custom-scrollbar rounded border border-border/30 bg-muted/5">
                      <table className="w-full text-[10px] border-collapse">
                        <tbody>
                          {element.rows.map((row, ri) => (
                            <tr key={ri} className="border-b border-border/20 last:border-b-0">
                              {row.cells.map((cell, ci) => (
                                <td
                                  key={ci}
                                  className="border-r border-border/20 last:border-r-0 px-1 py-0.5 text-muted-foreground"
                                  style={{
                                    rowSpan: cell.rowSpan > 1 ? cell.rowSpan : undefined,
                                    colSpan: cell.colSpan > 1 ? cell.colSpan : undefined,
                                  }}
                                >
                                  {cell.text || <span className="italic opacity-30">—</span>}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <span className="text-[8px] font-semibold uppercase tracking-wider text-muted-foreground/50 block mt-1">
                      Original Table
                    </span>
                  </motion.div>
                )}
              </AnimatePresence>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </Card>
  );
}

// ============================================================================
// Image Element Display
// ============================================================================

interface ImageElementDisplayProps {
  element: PptxImageElement;
  isExpanded: boolean;
  onToggle: () => void;
  isDragging?: boolean;
  dragHandleProps?: Record<string, unknown>;
  batchMode?: boolean;
  isBatchSelected?: boolean;
}

function ImageElementDisplay({ element, isExpanded, onToggle, isDragging = false, dragHandleProps, batchMode = false, isBatchSelected = false }: ImageElementDisplayProps) {
  const { updateImage, removeImage, selectElement, selectedElementId, toggleElementVisibility, hiddenElementIds, updateElementComment, elementComments } = usePptxStore();
  const isSelected = selectedElementId === element.id;
  const isHidden = hiddenElementIds.has(element.id);
  const [isZoomed, setIsZoomed] = useState(false);
  const [showCommentInput, setShowCommentInput] = useState(false);
  const [commentDraft, setCommentDraft] = useState(elementComments[element.id] || '');
  const fileInputRef = useRef<HTMLInputElement>(null);

  const isModified = !!element.replacementImageData;
  const isEmfOrWmf = /emf|wmf/i.test(element.imageType);

  const imageDataUrl = useMemo(() => {
    const toMimeType = (t: string | undefined): string => {
      if (!t) return 'image/png';
      if (t.includes('/')) return t;
      const extMap: Record<string, string> = {
        png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
        gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml',
        tiff: 'image/tiff', tif: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf',
        webp: 'image/webp',
      };
      return extMap[t.toLowerCase()] || `image/${t.toLowerCase()}`;
    };

    const buildDataUrl = (data: string | null | undefined, type: string | undefined): string | null => {
      if (!data) return null;
      if (data.startsWith('data:')) return data;
      const mime = toMimeType(type);
      return `data:${mime};base64,${data}`;
    };

    if (element.replacementImageData) {
      return buildDataUrl(element.replacementImageData, element.replacementImageType);
    }
    if (element.imageData) {
      return buildDataUrl(element.imageData, element.imageType);
    }
    return null;
  }, [element.replacementImageData, element.replacementImageType, element.imageData, element.imageType]);

  const handleReplaceImage = useCallback(() => {
    fileInputRef.current?.click();
  }, []);

  const handleFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;

      const reader = new FileReader();
      reader.onload = () => {
        const result = reader.result as string;
        const base64 = result.split(',')[1];
        const mimeType = file.type || 'image/png';
        updateImage(element.id, base64, mimeType);
      };
      reader.readAsDataURL(file);

      e.target.value = '';
    },
    [updateImage, element.id]
  );

  const handleRemoveReplacement = useCallback(() => {
    removeImage(element.id);
  }, [removeImage, element.id]);

  return (
    <Card
      ref={(el) => registerElementRef(element.id, el)}
      className={cn(
        'group overflow-hidden transition-all duration-200 p-px gap-px cursor-pointer',
        'border border-border/40 hover:border-emerald-300/40 dark:hover:border-emerald-700/30',
        'hover:shadow-sm hover:shadow-emerald-500/8',
        'hover:scale-[1.005] transition-transform duration-150',
        'hover:border-l-emerald-400/50 dark:hover:border-l-emerald-500/30',
        isModified && 'border-amber-300/60 dark:border-amber-500/25',
        isModified && 'animate-pulse-border-amber',
        isExpanded && 'shadow-sm shadow-emerald-500/5',
        isSelected && 'ring-1 ring-emerald-500/30 border-l-2 border-l-emerald-500 bg-emerald-50/30 dark:bg-emerald-950/20',
        isDragging && 'shadow-lg ring-2 ring-emerald-500/30',
        isHidden && 'opacity-60 border-l-2 border-l-amber-400/60 dark:border-l-amber-500/40',
        isBatchSelected && 'bg-emerald-50/40 dark:bg-emerald-950/25 border-emerald-300/50 dark:border-emerald-700/40'
      )}
      onClick={() => { if (!batchMode) selectElement(element.id); }}
    >
      {/* Batch checkbox (visible in batch mode) */}
      {batchMode && (
        <div
          className="absolute left-1 top-1.5 z-10"
          onClick={(e) => { e.stopPropagation(); toggleBatchSelect(element.id); }}
        >
          <div className={cn(
            'flex items-center justify-center size-3 rounded-sm border transition-all cursor-pointer',
            isBatchSelected
              ? 'bg-emerald-500 border-emerald-500 text-white'
              : 'border-muted-foreground/30 hover:border-muted-foreground/60'
          )}>
            {isBatchSelected && <Check className="size-2" strokeWidth={3} />}
          </div>
        </div>
      )}
      {/* Hidden file input */}
      <input
        ref={fileInputRef}
        type="file"
        accept="image/*"
        className="hidden"
        onChange={handleFileChange}
      />

      {/* Header */}
      <button
        onClick={(e) => { e.stopPropagation(); onToggle(); }}
        className="flex w-full items-center gap-1.5 px-2 py-1.5 text-left transition-colors hover:bg-emerald-50/20 dark:hover:bg-emerald-950/10"
      >
        {/* Drag handle */}
        <div
          {...dragHandleProps}
          className={cn(
            'shrink-0 cursor-grab active:cursor-grabbing rounded',
            'opacity-0 group-hover:opacity-100 transition-opacity duration-150',
            'hover:bg-muted/40 p-0.5 -m-0.5',
            dragHandleProps && 'opacity-60'
          )}
          onClick={(e) => e.stopPropagation()}
        >
          <GripVertical className="size-2 text-muted-foreground/50" />
        </div>

        <div className={cn(
          'size-5 shrink-0 rounded overflow-hidden',
          !imageDataUrl && 'flex items-center justify-center bg-cyan-500/10 text-cyan-600 dark:bg-cyan-500/15 dark:text-cyan-400'
        )}>
          {imageDataUrl ? (
            <img
              src={imageDataUrl}
              alt={element.imageName}
              className="object-cover w-full h-full"
            />
          ) : (
            <ImageIcon className="size-2.5" />
          )}
        </div>

        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-1">
            <span className="text-[10px] font-medium truncate">{element.shapeName}</span>
            {isModified && (
              <Badge
                className={cn(
                  'shrink-0 rounded-full px-0.5 py-0 h-3 text-[7px] font-semibold',
                  'bg-amber-100/80 text-amber-700 border-amber-200/50',
                  'dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-700/30'
                )}
              >
                Replaced
              </Badge>
            )}
          </div>
          <p className="text-[8px] text-muted-foreground truncate mt-px">
            {element.imageName}
          </p>
        </div>

        <motion.div
          animate={{ rotate: isExpanded ? 0 : -90 }}
          transition={{ duration: 0.15 }}
          className="shrink-0 text-muted-foreground"
        >
          <ChevronDown className="size-3" />
        </motion.div>

        {/* Visibility toggle */}
        <TooltipProvider delayDuration={300}>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="ghost"
                size="icon"
                className={cn(
                  'size-4',
                  isHidden
                    ? 'text-amber-500/70 hover:text-amber-500'
                    : 'text-muted-foreground/50 hover:text-muted-foreground'
                )}
                onClick={(e) => { e.stopPropagation(); toggleElementVisibility(element.id); }}
              >
                {isHidden ? <EyeOff className="size-2.5" /> : <Eye className="size-2.5" />}
              </Button>
            </TooltipTrigger>
            <TooltipContent side="bottom">
              <p>{isHidden ? 'Show in preview' : 'Hide in preview'}</p>
            </TooltipContent>
          </Tooltip>
        </TooltipProvider>

        {/* Comment toggle */}
        <TooltipProvider delayDuration={300}>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="ghost"
                size="icon"
                className={cn(
                  'size-4',
                  elementComments[element.id]
                    ? 'text-amber-500 hover:text-amber-600 dark:text-amber-400 dark:hover:text-amber-300'
                    : 'text-muted-foreground/40 hover:text-muted-foreground'
                )}
                onClick={(e) => { e.stopPropagation(); setShowCommentInput(v => !v); }}
              >
                {elementComments[element.id]
                  ? <MessageSquare className="size-2" fill="currentColor" />
                  : <MessageSquare className="size-2" />
                }
              </Button>
            </TooltipTrigger>
            <TooltipContent side="bottom">
              <p>{elementComments[element.id] ? 'Edit comment' : 'Add comment'}</p>
            </TooltipContent>
          </Tooltip>
        </TooltipProvider>
      </button>

      {/* Comment input (outside header button) */}
      <AnimatePresence>
        {showCommentInput && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.15 }}
            className="overflow-hidden"
          >
            <div className="px-2 pb-1">
              <div className="flex items-center gap-1 rounded border bg-amber-50/30 dark:bg-amber-900/10 border-amber-200/30 dark:border-amber-700/20 px-1.5 py-0.5">
                <MessageSquare className="size-2 text-amber-500 shrink-0" />
                <input
                  type="text"
                  value={commentDraft}
                  onChange={(e) => setCommentDraft(e.target.value)}
                  onBlur={() => {
                    updateElementComment(element.id, commentDraft);
                    if (!commentDraft.trim()) setShowCommentInput(false);
                  }}
                  onKeyDown={(e) => {
                    if (e.key === 'Enter') {
                      updateElementComment(element.id, commentDraft);
                      if (!commentDraft.trim()) setShowCommentInput(false);
                    }
                    if (e.key === 'Escape') {
                      setCommentDraft(elementComments[element.id] || '');
                      setShowCommentInput(false);
                    }
                  }}
                  placeholder="Add a note..."
                  className="flex-1 h-5 bg-transparent text-[8px] outline-none placeholder:text-amber-400/40 dark:placeholder:text-amber-500/30"
                  autoFocus
                />
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Expandable content */}
      <AnimatePresence initial={false}>
        {isExpanded && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.2, ease: [0.4, 0, 0.2, 1] }}
            className="overflow-hidden"
          >
            <div className="px-2 pb-1.5 space-y-1">
              <Separator className="opacity-30" />

              {/* Image preview */}
              <div className="space-y-0.5">
                <span className="text-[8px] font-semibold uppercase tracking-wider text-muted-foreground/60">
                  Preview
                </span>
                <div
                  className={cn(
                    'relative group/img overflow-hidden rounded border border-border/30',
                    'bg-muted/10'
                  )}
                >
                  {imageDataUrl ? (
                    <>
                      <img
                        src={imageDataUrl}
                        alt={element.imageName}
                        className={cn(
                          'w-full object-contain transition-transform duration-200',
                          isZoomed ? 'max-h-none scale-150 cursor-zoom-out' : 'max-h-36 cursor-zoom-in'
                        )}
                        onClick={() => setIsZoomed((z) => !z)}
                      />
                      <div className="absolute top-1 right-1 opacity-0 group-hover/img:opacity-100 transition-opacity">
                        <TooltipProvider delayDuration={200}>
                          <Tooltip>
                            <TooltipTrigger asChild>
                              <Button
                                variant="secondary"
                                size="icon"
                                className="size-5 shadow-sm"
                                onClick={() => setIsZoomed((z) => !z)}
                              >
                                <ZoomIn className="size-2.5" />
                              </Button>
                            </TooltipTrigger>
                            <TooltipContent side="left">
                              <p>{isZoomed ? 'Zoom out' : 'Zoom in'}</p>
                            </TooltipContent>
                          </Tooltip>
                        </TooltipProvider>
                      </div>
                    </>
                  ) : isEmfOrWmf ? (
                    <div className="flex flex-col items-center justify-center py-5 text-muted-foreground">
                      <ImageIcon className="size-5 mb-1 opacity-25" />
                      <span className="text-[9px]">EMF/WMF — no preview</span>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center justify-center py-5 text-muted-foreground">
                      <ImageIcon className="size-5 mb-1 opacity-25" />
                      <span className="text-[9px]">No preview</span>
                    </div>
                  )}
                </div>
              </div>

              {/* Action buttons */}
              <div className="flex items-center gap-0.5">
                <TooltipProvider delayDuration={300}>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Button
                        variant="outline"
                        size="sm"
                        className="h-5 gap-1 text-[9px] border-emerald-200/40 hover:border-emerald-300/50 hover:bg-emerald-50/20 dark:border-emerald-700/20 dark:hover:border-emerald-600/40 dark:hover:bg-emerald-950/15 px-1"
                        onClick={handleReplaceImage}
                      >
                        <Upload className="size-2" />
                        Replace
                      </Button>
                    </TooltipTrigger>
                    <TooltipContent side="bottom">
                      <p>Upload a replacement image</p>
                    </TooltipContent>
                  </Tooltip>
                </TooltipProvider>

                {isModified && (
                  <TooltipProvider delayDuration={300}>
                    <Tooltip>
                      <TooltipTrigger asChild>
                        <Button
                          variant="ghost"
                          size="sm"
                          className={cn(
                            'h-5 gap-1 text-[9px] px-1',
                            'text-amber-600 hover:text-amber-700 dark:text-amber-400 dark:hover:text-amber-300'
                          )}
                          onClick={handleRemoveReplacement}
                        >
                          <X className="size-2" />
                          Revert
                        </Button>
                      </TooltipTrigger>
                      <TooltipContent side="bottom">
                        <p>Revert to original image</p>
                      </TooltipContent>
                    </Tooltip>
                  </TooltipProvider>
                )}
              </div>

              {/* Image info */}
              <div className={cn(
                'rounded border p-1.5',
                'bg-muted/10 border-border/30'
              )}>
                <span className="text-[8px] font-semibold uppercase tracking-wider text-muted-foreground/50 block mb-0.5">
                  Info
                </span>
                <div className="grid grid-cols-2 gap-x-2 gap-y-px text-[9px]">
                  <span className="text-muted-foreground">Name</span>
                  <span className="truncate font-medium">{element.imageName}</span>
                  <span className="text-muted-foreground">Format</span>
                  <span className="font-medium">{element.imageType || 'Unknown'}</span>
                  <span className="text-muted-foreground">Size</span>
                  <span className="font-medium">
                    {element.position.width > 0
                      ? `${Math.round(element.position.width / 914400)}" × ${Math.round(element.position.height / 914400)}"`
                      : 'N/A'}
                  </span>
                  {isEmfOrWmf && (
                    <>
                      <span className="text-muted-foreground">Note</span>
                      <span className="font-medium text-amber-600 dark:text-amber-400">Vector</span>
                    </>
                  )}
                </div>
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </Card>
  );
}

// ============================================================================
// Section Header
// ============================================================================

interface SectionHeaderProps {
  icon: React.ReactNode;
  title: string;
  count: number;
  accentClass?: string;
  collapsed?: boolean;
  onToggleCollapse?: () => void;
  modifiedCount?: number;
  tooltipDetail?: string;
}

function SectionHeader({ icon, title, count, accentClass, collapsed = false, onToggleCollapse, modifiedCount = 0, tooltipDetail }: SectionHeaderProps) {
  const isCollapsible = !!onToggleCollapse;

  const countBadge = (
    <Badge
      variant="secondary"
      className="h-3 px-1 text-[8px] font-medium"
    >
      {count}
    </Badge>
  );

  return (
    <motion.div
      whileHover={isCollapsible ? { scale: 1.005 } : undefined}
      whileTap={isCollapsible ? { scale: 0.995 } : undefined}
      className={cn(
        'flex items-center gap-1.5 py-1 px-1 bg-muted/15 dark:bg-muted/20 rounded',
        isCollapsible && 'cursor-pointer select-none hover:bg-muted/30 transition-colors'
      )}
      onClick={isCollapsible ? onToggleCollapse : undefined}
      role={isCollapsible ? 'button' : undefined}
      tabIndex={isCollapsible ? 0 : undefined}
      onKeyDown={isCollapsible ? (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); onToggleCollapse(); } } : undefined}
      aria-expanded={isCollapsible ? !collapsed : undefined}
    >
      {isCollapsible && (
        <motion.div
          animate={{ rotate: collapsed ? -90 : 0 }}
          transition={{ duration: 0.15 }}
          className="shrink-0 text-muted-foreground/60"
        >
          <ChevronDown className="size-2.5" />
        </motion.div>
      )}
      <div className={cn(
        'flex size-4 items-center justify-center rounded shrink-0',
        accentClass ?? 'bg-muted text-muted-foreground'
      )}>
        {icon}
      </div>
      <span className="text-[9px] font-bold uppercase tracking-widest text-muted-foreground/70">
        {title}
      </span>
      {tooltipDetail ? (
        <TooltipProvider delayDuration={400}>
          <Tooltip>
            <TooltipTrigger asChild>
              <span className="inline-flex">{countBadge}</span>
            </TooltipTrigger>
            <TooltipContent side="top" className="text-[9px]">
              <p>{tooltipDetail}</p>
            </TooltipContent>
          </Tooltip>
        </TooltipProvider>
      ) : (
        countBadge
      )}
      {modifiedCount > 0 && (
        <span className="flex items-center gap-0.5 text-[8px] text-amber-600 dark:text-amber-400">
          <span className="size-1 rounded-full bg-amber-400 dark:bg-amber-500 shrink-0 animate-pulse-dot-amber" />
          {modifiedCount} edited
        </span>
      )}
      <div className="flex-1 h-px bg-gradient-to-r from-border/50 via-border/35 to-transparent" />
    </motion.div>
  );
}

// ============================================================================
// Keyboard Shortcuts Hint Bar
// ============================================================================

function KeyboardShortcutsHint() {
  const [dismissed, setDismissed] = useState(() => {
    if (typeof window === 'undefined') return false;
    try { return localStorage.getItem('pptx-kb-hint-dismissed') === '1'; } catch { return false; }
  });

  useEffect(() => {
    try { localStorage.setItem('pptx-kb-hint-dismissed', dismissed ? '1' : '0'); } catch { /* ignore */ }
  }, [dismissed]);

  if (dismissed) return null;

  const shortcuts = [
    { keys: 'Ctrl+Z', label: 'Undo' },
    { keys: 'Ctrl+Y', label: 'Redo' },
    { keys: 'Ctrl+S', label: 'Save' },
    { keys: 'Ctrl+E', label: 'Export' },
  ];

  return (
    <div className="shrink-0 flex items-center justify-center gap-2 px-2 h-5 bg-muted/10 border-t border-border/20 select-none">
      {shortcuts.map((s) => (
        <span key={s.keys} className="flex items-center gap-0.5 text-[7px] text-muted-foreground/40">
          <kbd className="px-0.5 py-px rounded bg-muted/30 border border-border/20 text-[7px] font-mono text-muted-foreground/50">
            {s.keys}
          </kbd>
          <span>{s.label}</span>
        </span>
      ))}
      <button
        onClick={() => setDismissed(true)}
        className="ml-1 text-muted-foreground/30 hover:text-muted-foreground/60 transition-colors focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-emerald-500/30 rounded"
        aria-label="Dismiss shortcuts hint"
      >
        <X className="size-2" />
      </button>
    </div>
  );
}

// ============================================================================
// Slide Editor (Main Component)
// ============================================================================

interface SlideEditorProps {
  slide: PptxSlideData;
}

export default function SlideEditor({ slide }: SlideEditorProps) {
  const {
    updateText,
    updateTableCell,
    updateImage,
    removeImage,
    selectElement,
    selectedElementId,
    hideEmpty,
    currentSlideIndex,
    reorderElements,
    updateSlideNote,
    slideNotes,
    toggleBatchSelect,
    clearBatchSelection,
    batchSelectAll,
    batchSelectedIds,
    toggleElementVisibility,
    hiddenElementIds,
  } = usePptxStore();

  const [manuallyExpandedIds, setManuallyExpandedIds] = useState<Set<string>>(new Set());
  const [searchQuery, setSearchQuery] = useState('');
  const [filterType, setFilterType] = useState<'all' | 'text' | 'table' | 'image'>('all');
  const [showFindReplace, setShowFindReplace] = useState(false);
  const [findQuery, setFindQuery] = useState('');
  const [replaceQuery, setReplaceQuery] = useState('');
  const [caseSensitive, setCaseSensitive] = useState(false);
  const [showVariablesPanel, setShowVariablesPanel] = useState(false);
  const [showNotesPanel, setShowNotesPanel] = useState(false);
  const [collapsedSections, setCollapsedSections] = useState<Set<string>>(new Set());
  const [batchMode, setBatchMode] = useState(false);

  // Track the initial slide number to skip entrance animation on first load
  const [initialSlideNumber] = useState(slide.slideNumber);
  const shouldAnimateEntrance = slide.slideNumber !== initialSlideNumber;

  // Compute auto-expanded IDs for modified table/image elements (always expand modified elements)
  const autoExpandedModifiedIds = useMemo(() => {
    const ids = new Set<string>();
    for (const el of slide.elements) {
      if (el.type === 'table' && el.currentRows) {
        ids.add(el.id);
      } else if (el.type === 'image' && el.replacementImageData) {
        ids.add(el.id);
      }
    }
    return ids;
  }, [slide.elements]);

  // DnD sensors
  const sensors = useSensors(
    useSensor(PointerSensor, {
      activationConstraint: {
        distance: 5,
      },
    }),
    useSensor(KeyboardSensor, {
      coordinateGetter: sortableKeyboardCoordinates,
    })
  );

  // Compute effective expanded set: manually expanded + auto-expanded selected element + auto-expanded modified elements
  const expandedIds = useMemo(() => {
    const ids = new Set(manuallyExpandedIds);
    if (selectedElementId) {
      ids.add(selectedElementId);
    }
    for (const id of autoExpandedModifiedIds) {
      ids.add(id);
    }
    return ids;
  }, [manuallyExpandedIds, selectedElementId, autoExpandedModifiedIds]);

  const toggleExpand = useCallback((id: string) => {
    setManuallyExpandedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) {
        next.delete(id);
      } else {
        next.add(id);
      }
      return next;
    });
  }, []);

  const toggleSectionCollapse = useCallback((key: string) => {
    setCollapsedSections((prev) => {
      const next = new Set(prev);
      if (next.has(key)) {
        next.delete(key);
      } else {
        next.add(key);
      }
      return next;
    });
  }, []);

  // Categorize elements
  const { selectedElements, textElements, tableElements, imageElements, visibleElements } = useMemo(() => {
    const selected: PptxElement[] = [];
    const text: PptxTextElement[] = [];
    const table: PptxTableElement[] = [];
    const image: PptxImageElement[] = [];

    for (const el of slide.elements) {
      if (hideEmpty && isEmptyElement(el)) continue;

      if (selectedElementId === el.id) {
        selected.push(el);
      }

      switch (el.type) {
        case 'text':
          text.push(el);
          break;
        case 'table':
          table.push(el);
          break;
        case 'image':
          image.push(el);
          break;
      }
    }

    const visible = [...text, ...table, ...image];

    return {
      selectedElements: selected,
      textElements: text,
      tableElements: table,
      imageElements: image,
      visibleElements: visible,
    };
  }, [slide.elements, hideEmpty, selectedElementId]);

  // Search filtering
  const filteredTextElements = useMemo(() => {
    if (!searchQuery.trim()) return textElements;
    const q = searchQuery.toLowerCase();
    return textElements.filter((el) =>
      el.shapeName.toLowerCase().includes(q) ||
      el.originalText.toLowerCase().includes(q) ||
      (el.currentText && el.currentText.toLowerCase().includes(q))
    );
  }, [textElements, searchQuery]);

  const filteredTableElements = useMemo(() => {
    if (!searchQuery.trim()) return tableElements;
    const q = searchQuery.toLowerCase();
    return tableElements.filter((el) =>
      el.shapeName.toLowerCase().includes(q) ||
      el.rows.some((row) => row.cells.some((cell) => cell.text.toLowerCase().includes(q)))
    );
  }, [tableElements, searchQuery]);

  const filteredImageElements = useMemo(() => {
    if (!searchQuery.trim()) return imageElements;
    const q = searchQuery.toLowerCase();
    return imageElements.filter((el) =>
      el.shapeName.toLowerCase().includes(q) ||
      el.imageName.toLowerCase().includes(q)
    );
  }, [imageElements, searchQuery]);

  const filteredTotal = filteredTextElements.length + filteredTableElements.length + filteredImageElements.length;
  const isSearching = searchQuery.trim().length > 0;

  // Filtered totals by type (respecting both search and filterType)
  const displayTextElements = filterType === 'all' || filterType === 'text' ? filteredTextElements : [];
  const displayTableElements = filterType === 'all' || filterType === 'table' ? filteredTableElements : [];
  const displayImageElements = filterType === 'all' || filterType === 'image' ? filteredImageElements : [];
  const displayTotal = displayTextElements.length + displayTableElements.length + displayImageElements.length;

  // Total counts
  const totalText = textElements.length;
  const totalTable = tableElements.length;
  const totalImage = imageElements.length;
  const totalAll = totalText + totalTable + totalImage;

  // Template variables summary across the current slide's text & table elements
  const slideVariables = useMemo(() => {
    const varMap = new Map<string, { name: string; elementIds: string[]; currentValues: string[] }>();

    for (const el of slide.elements) {
      if (el.type === 'text') {
        const vars = extractVariables(el.originalText);
        const currentText = el.currentText ?? el.originalText;
        for (const varName of vars) {
          if (!varMap.has(varName)) {
            varMap.set(varName, { name: varName, elementIds: [], currentValues: [] });
          }
          const entry = varMap.get(varName)!;
          if (!entry.elementIds.includes(el.id)) {
            entry.elementIds.push(el.id);
          }
          // Try to extract current value by matching the same variable pattern
          const varRegex = new RegExp(`\\{\\{\\s*${varName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\s*\\}\\}`, 'g');
          const currentMatch = varRegex.exec(currentText);
          entry.currentValues.push(currentMatch ? currentText : 'Not set');
        }
      } else if (el.type === 'table') {
        const rows = el.currentRows ?? el.rows;
        for (const row of rows) {
          for (const cell of row.cells) {
            const vars = extractVariables(cell.text);
            for (const varName of vars) {
              if (!varMap.has(varName)) {
                varMap.set(varName, { name: varName, elementIds: [], currentValues: [] });
              }
              const entry = varMap.get(varName)!;
              if (!entry.elementIds.includes(el.id)) {
                entry.elementIds.push(el.id);
              }
              entry.currentValues.push(cell.text !== cell.originalText ? cell.text : 'Not set');
            }
          }
        }
      }
    }

    return Array.from(varMap.values()).sort((a, b) => a.name.localeCompare(b.name));
  }, [slide.elements]);

  // Modification count for this slide
  const modCount = useMemo(() => {
    let count = 0;
    for (const el of slide.elements) {
      if (el.type === 'text' && el.currentText !== undefined && el.currentText !== el.originalText) count++;
      else if (el.type === 'table' && el.currentRows) {
        for (let ri = 0; ri < el.currentRows.length; ri++) {
          const origRow = el.rows[ri];
          const curRow = el.currentRows[ri];
          if (!origRow || !curRow) continue;
          for (let ci = 0; ci < curRow.cells.length; ci++) {
            if (!origRow.cells[ci] || !curRow.cells[ci]) continue;
            if (curRow.cells[ci].text !== origRow.cells[ci].text) count++;
          }
        }
      } else if (el.type === 'image' && el.replacementImageData) count++;
    }
    return count;
  }, [slide.elements]);

  // Per-section modification counts
  const textModCount = useMemo(() => {
    let count = 0;
    for (const el of textElements) {
      if (el.currentText !== undefined && el.currentText !== el.originalText) count++;
    }
    return count;
  }, [textElements]);

  const tableModCount = useMemo(() => {
    let count = 0;
    for (const el of tableElements) {
      if (el.currentRows) count++;
    }
    return count;
  }, [tableElements]);

  const imageModCount = useMemo(() => {
    let count = 0;
    for (const el of imageElements) {
      if (el.replacementImageData) count++;
    }
    return count;
  }, [imageElements]);

  // Find & Replace: count matches
  const findMatches = useMemo(() => {
    if (!findQuery) return 0;
    let count = 0;
    const q = caseSensitive ? findQuery : findQuery.toLowerCase();
    for (const el of slide.elements) {
      if (el.type === 'text') {
        const text = el.currentText ?? el.originalText;
        const hay = caseSensitive ? text : text.toLowerCase();
        const origHay = caseSensitive ? el.originalText : el.originalText.toLowerCase();
        let idx = hay.indexOf(q);
        while (idx !== -1) { count++; idx = hay.indexOf(q, idx + 1); }
        if (el.currentText !== undefined && el.currentText !== el.originalText) {
          idx = origHay.indexOf(q);
          while (idx !== -1) { count++; idx = origHay.indexOf(q, idx + 1); }
        }
      } else if (el.type === 'table') {
        const rows = el.currentRows ?? el.rows;
        for (const row of rows) {
          for (const cell of row.cells) {
            const hay = caseSensitive ? cell.text : cell.text.toLowerCase();
            let idx = hay.indexOf(q);
            while (idx !== -1) { count++; idx = hay.indexOf(q, idx + 1); }
          }
        }
      }
    }
    return count;
  }, [slide.elements, findQuery, caseSensitive]);

  // Find & Replace: replace all matches
  const handleReplaceAll = useCallback(() => {
    if (!findQuery) return;
    let totalReplacements = 0;
    const q = caseSensitive ? findQuery : findQuery.toLowerCase();

    for (const el of slide.elements) {
      if (el.type === 'text') {
        const text = el.currentText ?? el.originalText;
        const hay = caseSensitive ? text : text.toLowerCase();
        if (hay.includes(q)) {
          let newText = text;
          if (caseSensitive) {
            newText = text.split(findQuery).join(replaceQuery);
          } else {
            const regex = new RegExp(findQuery.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
            newText = text.replace(regex, replaceQuery);
          }
          const count = (hay.match(new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), caseSensitive ? 'g' : 'gi')) || []).length;
          totalReplacements += count;
          updateText(el.id, newText);
        }
      } else if (el.type === 'table') {
        const rows = el.currentRows ?? el.rows;
        for (let ri = 0; ri < rows.length; ri++) {
          for (let ci = 0; ci < rows[ri].cells.length; ci++) {
            const cellText = rows[ri].cells[ci].text;
            const hay = caseSensitive ? cellText : cellText.toLowerCase();
            if (hay.includes(q)) {
              let newText = cellText;
              if (caseSensitive) {
                newText = cellText.split(findQuery).join(replaceQuery);
              } else {
                const regex = new RegExp(findQuery.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi');
                newText = cellText.replace(regex, replaceQuery);
              }
              const count = (hay.match(new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), caseSensitive ? 'g' : 'gi')) || []).length;
              totalReplacements += count;
              updateTableCell(el.id, ri, ci, newText);
            }
          }
        }
      }
    }

    if (totalReplacements > 0) {
      toast.success(`Replaced ${totalReplacements} occurrence${totalReplacements !== 1 ? 's' : ''}`);
    } else {
      toast.info('No matches found');
    }
  }, [findQuery, replaceQuery, caseSensitive, slide.elements, updateText, updateTableCell]);

  // DnD handlers - reorder elements of the same type
  const handleDragStart = useCallback((_event: DragStartEvent) => {
    // Drag started - visual feedback handled by SortableElementWrapper
  }, []);

  const handleDragEnd = useCallback((event: DragEndEvent, elementType: 'text' | 'table' | 'image') => {
    const { active, over } = event;

    if (!over || active.id === over.id) return;

    // Find indices in slide.elements for the active and over items
    const activeEl = slide.elements.find((el) => el.id === active.id);
    const overEl = slide.elements.find((el) => el.id === over.id);

    if (!activeEl || !overEl) return;

    // Only allow same-type reordering
    if (activeEl.type !== overEl.type || activeEl.type !== elementType) return;

    const oldIndex = slide.elements.findIndex((el) => el.id === active.id);
    const newIndex = slide.elements.findIndex((el) => el.id === over.id);

    if (oldIndex === -1 || newIndex === -1) return;

    reorderElements(currentSlideIndex, oldIndex, newIndex);
  }, [slide.elements, currentSlideIndex, reorderElements]);

  // Create type-specific drag end handlers
  const handleTextDragEnd = useCallback((event: DragEndEvent) => handleDragEnd(event, 'text'), [handleDragEnd]);
  const handleTableDragEnd = useCallback((event: DragEndEvent) => handleDragEnd(event, 'table'), [handleDragEnd]);
  const handleImageDragEnd = useCallback((event: DragEndEvent) => handleDragEnd(event, 'image'), [handleDragEnd]);

  return (
    <div className="flex flex-col h-full overflow-hidden max-w-full dark:bg-muted/10">
      {/* Header - fixed height to match preview header */}
      <div className="shrink-0 flex items-center gap-1.5 border-b border-border/30 px-2 h-[30px]">
        <div className="flex size-5 items-center justify-center rounded bg-gradient-to-br from-emerald-500 to-teal-600 shrink-0">
          <FileText className="size-2.5 text-white" />
        </div>
        <span className="text-[11px] font-semibold tracking-tight">Slide {slide.slideNumber}</span>
        <Badge variant="secondary" className="h-3.5 px-1 text-[9px] font-medium bg-emerald-50/40 dark:bg-emerald-900/15 text-emerald-600 dark:text-emerald-400 border-emerald-200/30 dark:border-emerald-700/20">
          {totalAll} el
        </Badge>
        {/* Element type distribution bar */}
        {totalAll > 0 && (
          <TooltipProvider delayDuration={400}>
            <Tooltip>
              <TooltipTrigger asChild>
                <div className="flex w-[50px] h-1 rounded-full overflow-hidden bg-muted/30 shrink-0">
                  {totalText > 0 && (
                    <div
                      className="h-full bg-emerald-500 dark:bg-emerald-400 transition-all duration-300"
                      style={{ width: `${(totalText / totalAll) * 100}%` }}
                    />
                  )}
                  {totalTable > 0 && (
                    <div
                      className="h-full bg-amber-500 dark:bg-amber-400 transition-all duration-300"
                      style={{ width: `${(totalTable / totalAll) * 100}%` }}
                    />
                  )}
                  {totalImage > 0 && (
                    <div
                      className="h-full bg-cyan-500 dark:bg-cyan-400 transition-all duration-300"
                      style={{ width: `${(totalImage / totalAll) * 100}%` }}
                    />
                  )}
                </div>
              </TooltipTrigger>
              <TooltipContent side="bottom" className="text-[9px]">
                <div className="flex items-center gap-2">
                  <span className="flex items-center gap-1"><span className="size-1.5 rounded-full bg-emerald-500" />Text: {totalText}</span>
                  <span className="flex items-center gap-1"><span className="size-1.5 rounded-full bg-amber-500" />Table: {totalTable}</span>
                  <span className="flex items-center gap-1"><span className="size-1.5 rounded-full bg-cyan-500" />Image: {totalImage}</span>
                </div>
              </TooltipContent>
            </Tooltip>
          </TooltipProvider>
        )}
        <div className="flex-1" />
        {/* Element Statistics Popover */}
        <Popover>
          <PopoverTrigger asChild>
            <button
              className={cn(
                'flex items-center justify-center size-4 rounded shrink-0 transition-all duration-150',
                'text-muted-foreground/50 hover:text-muted-foreground hover:bg-muted/30',
                'focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30'
              )}
              aria-label="Element statistics"
            >
              <BarChart3 className="size-2.5" />
            </button>
          </PopoverTrigger>
          <PopoverContent side="bottom" align="end" className="w-52 p-2 text-[9px]">
            <div className="space-y-1.5">
              <div className="flex items-center gap-1 text-[10px] font-semibold text-foreground/80">
                <BarChart3 className="size-3 text-emerald-500" />
                Slide Statistics
              </div>
              <Separator className="opacity-30" />
              <div className="space-y-1">
                <div className="flex items-center gap-1.5">
                  <Type className="size-2.5 text-emerald-500 shrink-0" />
                  <span className="flex-1 text-muted-foreground">Text</span>
                  <span className="font-semibold text-foreground">{totalText}</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <Table2 className="size-2.5 text-amber-500 shrink-0" />
                  <span className="flex-1 text-muted-foreground">Table</span>
                  <span className="font-semibold text-foreground">{totalTable}</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <ImageIcon className="size-2.5 text-cyan-500 shrink-0" />
                  <span className="flex-1 text-muted-foreground">Image</span>
                  <span className="font-semibold text-foreground">{totalImage}</span>
                </div>
              </div>
              <Separator className="opacity-30" />
              <div className="space-y-1">
                <div className="flex items-center gap-1.5">
                  <AlignLeft className="size-2.5 text-muted-foreground shrink-0" />
                  <span className="flex-1 text-muted-foreground">Words</span>
                  <span className="font-semibold text-foreground">{useMemo(() => {
                    let words = 0;
                    for (const el of slide.elements) {
                      if (el.type === 'text') words += (el.currentText ?? el.originalText).split(/\s+/).filter(Boolean).length;
                      else if (el.type === 'table') {
                        const rows = el.currentRows ?? el.rows;
                        for (const row of rows) for (const cell of row.cells) words += cell.text.split(/\s+/).filter(Boolean).length;
                      }
                    }
                    return words;
                  }, [slide.elements])}</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <Hash className="size-2.5 text-muted-foreground shrink-0" />
                  <span className="flex-1 text-muted-foreground">Characters</span>
                  <span className="font-semibold text-foreground">{useMemo(() => {
                    let chars = 0;
                    for (const el of slide.elements) {
                      if (el.type === 'text') chars += (el.currentText ?? el.originalText).length;
                      else if (el.type === 'table') {
                        const rows = el.currentRows ?? el.rows;
                        for (const row of rows) for (const cell of row.cells) chars += cell.text.length;
                      }
                    }
                    return chars;
                  }, [slide.elements])}</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <Percent className="size-2.5 text-amber-500 shrink-0" />
                  <span className="flex-1 text-muted-foreground">Modified</span>
                  <span className="font-semibold text-foreground">{totalAll > 0 ? Math.round((modCount / totalAll) * 100) : 0}%</span>
                </div>
                <div className="flex items-center gap-1.5">
                  <Braces className="size-2.5 text-emerald-500 shrink-0" />
                  <span className="flex-1 text-muted-foreground">Variables</span>
                  <span className="font-semibold text-foreground">{slideVariables.length}</span>
                </div>
              </div>
              {/* Modification progress bar */}
              {totalAll > 0 && (
                <>
                  <Separator className="opacity-30" />
                  <div className="space-y-0.5">
                    <span className="text-muted-foreground/60">Modification progress</span>
                    <div className="w-full h-1.5 rounded-full bg-muted/40 overflow-hidden">
                      <div
                        className="h-full rounded-full bg-gradient-to-r from-amber-400 to-amber-500 dark:from-amber-500 dark:to-amber-400 transition-all duration-300"
                        style={{ width: `${Math.min(100, (modCount / totalAll) * 100)}%` }}
                      />
                    </div>
                  </div>
                </>
              )}
            </div>
          </PopoverContent>
        </Popover>
        {modCount > 0 && (
          <div className="flex items-center gap-1">
            <div className="w-8 h-1 rounded-full bg-muted/40 overflow-hidden">
              <div
                className="h-full rounded-full bg-gradient-to-r from-amber-400 to-amber-500 dark:from-amber-500 dark:to-amber-400 transition-all duration-300"
                style={{ width: `${Math.min(100, (modCount / totalAll) * 100)}%` }}
              />
            </div>
            <Badge
              className={cn(
                'rounded-full px-1.5 py-0 h-4 text-[9px] font-semibold',
                'bg-amber-100/80 text-amber-700 border-amber-200/50',
                'dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-700/30'
              )}
            >
              {modCount} edit{modCount !== 1 ? 's' : ''}
            </Badge>
          </div>
        )}
      </div>

      {/* Search bar */}
      <div className="shrink-0 px-1.5 pt-1.5">
        <div className="relative">
          <Search className="absolute left-2 top-1/2 -translate-y-1/2 size-3 text-muted-foreground/60" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search elements..."
            className={cn(
              'w-full h-6 pl-6 pr-6 text-[10px] rounded-md border border-border/30',
              'bg-muted/15 dark:bg-muted/25 dark:border-white/8 placeholder:text-muted-foreground/50 dark:placeholder:text-muted-foreground/50',
              'focus:outline-none focus:ring-1 focus:ring-emerald-500/20 focus:border-emerald-300/40',
              'dark:focus-visible:ring-emerald-400/25 dark:focus-visible:border-emerald-500/40 transition-all'
            )}
          />
          {isSearching && (
            <button
              onClick={() => setSearchQuery('')}
              aria-label="Clear search"
              className="absolute right-1.5 top-1/2 -translate-y-1/2 size-3.5 flex items-center justify-center rounded text-muted-foreground/60 hover:text-foreground transition-colors focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30"
            >
              <X className="size-2.5" />
            </button>
          )}
        </div>
        {isSearching && (
          <div className="flex items-center gap-1 mt-1 text-[9px] text-muted-foreground/60">
            <span>{filteredTotal} result{filteredTotal !== 1 ? 's' : ''}</span>
            <span>·</span>
            <button
              onClick={() => setSearchQuery('')}
              className="text-emerald-600 dark:text-emerald-400 hover:underline"
            >
              Clear
            </button>
          </div>
        )}
      </div>

      {/* Element type filter tabs + Find & Replace toggle */}
      <div className="shrink-0 px-1.5 pt-1">
        <div className="flex items-center gap-0.5" role="tablist" aria-label="Element type filter">
          {([
            { type: 'all' as const, icon: <Layers className="size-2" />, label: 'All', count: totalAll },
            { type: 'text' as const, icon: <Type className="size-2" />, label: 'Text', count: totalText },
            { type: 'table' as const, icon: <Table2 className="size-2" />, label: 'Table', count: totalTable },
            { type: 'image' as const, icon: <ImageIcon className="size-2" />, label: 'Image', count: totalImage },
          ] as const).map(({ type, icon, label, count }) => (
            <button
              key={type}
              role="tab"
              aria-selected={filterType === type}
              tabIndex={filterType === type ? 0 : -1}
              onClick={() => setFilterType(type)}
              onKeyDown={(e) => {
                const types = ['all', 'text', 'table', 'image'] as const;
                const idx = types.indexOf(type);
                if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
                  e.preventDefault();
                  const next = types[(idx + 1) % types.length];
                  setFilterType(next);
                  (e.target as HTMLElement).parentElement?.querySelector(`[data-type="${next}"]`)?.focus();
                } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
                  e.preventDefault();
                  const prev = types[(idx - 1 + types.length) % types.length];
                  setFilterType(prev);
                  (e.target as HTMLElement).parentElement?.querySelector(`[data-type="${prev}"]`)?.focus();
                }
              }}
              data-type={type}
              className={cn(
                'flex items-center gap-0.5 h-5 px-1.5 rounded text-[9px] font-medium transition-all duration-150',
                'focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                filterType === type
                  ? 'bg-emerald-500/15 text-emerald-700 dark:bg-emerald-500/20 dark:text-emerald-300 shadow-sm shadow-emerald-500/5'
                  : 'text-muted-foreground/60 hover:text-muted-foreground hover:bg-muted/20'
              )}
            >
              {icon}
              <span>{label}</span>
              {count > 0 && (
                <span className={cn(
                  'inline-flex items-center justify-center min-w-[10px] h-2.5 px-[3px] rounded-full text-[7px] font-semibold leading-none',
                  filterType === type
                    ? 'bg-emerald-500/25 text-emerald-700 dark:text-emerald-300'
                    : 'bg-muted/40 text-muted-foreground/50'
                )}>
                  {count}
                </span>
              )}
            </button>
          ))}
          <div className="flex-1" />
          <TooltipProvider delayDuration={300}>
            <Tooltip>
              <TooltipTrigger asChild>
                <button
                  onClick={() => {
                    setBatchMode((v) => {
                      if (v) clearBatchSelection();
                      return !v;
                    });
                  }}
                  aria-label="Batch select mode"
                  className={cn(
                    'flex items-center justify-center h-5 w-5 rounded transition-all duration-150',
                    'focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                    batchMode
                      ? 'bg-emerald-500/15 text-emerald-700 dark:bg-emerald-500/20 dark:text-emerald-300'
                      : 'text-muted-foreground/50 hover:text-muted-foreground hover:bg-muted/20'
                  )}
                >
                  <ListChecks className="size-2.5" />
                </button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Batch select mode</p>
              </TooltipContent>
            </Tooltip>
          </TooltipProvider>
          <TooltipProvider delayDuration={300}>
            <Tooltip>
              <TooltipTrigger asChild>
                <button
                  onClick={() => setShowFindReplace((v) => !v)}
                  aria-label="Find and replace"
                  className={cn(
                    'flex items-center justify-center h-5 w-5 rounded transition-all duration-150',
                    'focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                    showFindReplace
                      ? 'bg-emerald-500/15 text-emerald-700 dark:bg-emerald-500/20 dark:text-emerald-300'
                      : 'text-muted-foreground/50 hover:text-muted-foreground hover:bg-muted/20'
                  )}
                >
                  <Replace className="size-2.5" />
                </button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Find & Replace</p>
              </TooltipContent>
            </Tooltip>
          </TooltipProvider>
        </div>
      </div>

      {/* Find & Replace bar */}
      <AnimatePresence>
        {showFindReplace && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.15 }}
            className="shrink-0 overflow-hidden"
          >
            <div className="px-1.5 pt-1 space-y-0.5">
              <div className="flex items-center gap-0.5">
                <input
                  type="text"
                  value={findQuery}
                  onChange={(e) => setFindQuery(e.target.value)}
                  placeholder="Find..."
                  className={cn(
                    'flex-1 h-5 px-1.5 text-[9px] rounded border border-border/30',
                    'bg-muted/15 dark:bg-muted/20 placeholder:text-muted-foreground/50',
                    'focus:outline-none focus:ring-1 focus:ring-emerald-500/20 focus:border-emerald-300/40',
                    'dark:focus:border-emerald-600/30 transition-all'
                  )}
                />
                <input
                  type="text"
                  value={replaceQuery}
                  onChange={(e) => setReplaceQuery(e.target.value)}
                  placeholder="Replace..."
                  className={cn(
                    'flex-1 h-5 px-1.5 text-[9px] rounded border border-border/30',
                    'bg-muted/15 dark:bg-muted/20 placeholder:text-muted-foreground/50',
                    'focus:outline-none focus:ring-1 focus:ring-emerald-500/20 focus:border-emerald-300/40',
                    'dark:focus:border-emerald-600/30 transition-all'
                  )}
                />
              </div>
              <div className="flex items-center gap-0.5">
                <button
                  onClick={() => setCaseSensitive((v) => !v)}
                  className={cn(
                    'flex items-center justify-center h-4 w-4 rounded transition-all',
                    caseSensitive
                      ? 'bg-emerald-500/15 text-emerald-700 dark:bg-emerald-500/20 dark:text-emerald-300'
                      : 'text-muted-foreground/40 hover:text-muted-foreground hover:bg-muted/20'
                  )}
                  title={caseSensitive ? 'Case sensitive' : 'Case insensitive'}
                >
                  <CaseSensitive className="size-2" />
                </button>
                <button
                  onClick={handleReplaceAll}
                  disabled={!findQuery || findMatches === 0}
                  className={cn(
                    'h-4 px-1.5 rounded text-[8px] font-semibold transition-all',
                    'bg-emerald-500/15 text-emerald-700 dark:bg-emerald-500/20 dark:text-emerald-300',
                    'hover:bg-emerald-500/25 dark:hover:bg-emerald-500/30',
                    'disabled:opacity-30 disabled:cursor-not-allowed'
                  )}
                >
                  Replace All
                </button>
                <div className="flex-1" />
                {findQuery && (
                  <span className="text-[8px] text-muted-foreground/60">
                    {findMatches} match{findMatches !== 1 ? 'es' : ''}
                  </span>
                )}
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Variables summary panel */}
      {slideVariables.length > 0 && (
        <div className="shrink-0 px-1.5 pt-1">
          <button
            onClick={() => setShowVariablesPanel((v) => !v)}
            className={cn(
              'flex items-center gap-1 w-full h-5 px-1.5 rounded text-[9px] font-semibold transition-all duration-150',
              'focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
              showVariablesPanel
                ? 'bg-emerald-500/10 text-emerald-700 dark:bg-emerald-500/15 dark:text-emerald-300'
                : 'bg-muted/15 text-muted-foreground/70 hover:bg-muted/25 hover:text-muted-foreground'
            )}
          >
            <motion.div
              animate={{ rotate: showVariablesPanel ? 90 : 0 }}
              transition={{ duration: 0.15 }}
              className="shrink-0"
            >
              <ChevronRight className="size-2" />
            </motion.div>
            <Braces className="size-2" />
            <span className="uppercase tracking-wider">Variables</span>
            <Badge
              className={cn(
                'shrink-0 rounded-full px-0.5 py-0 h-3 text-[7px] font-semibold',
                'bg-emerald-100/80 text-emerald-700 border-emerald-200/50',
                'dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-700/30'
              )}
            >
              {slideVariables.length}
            </Badge>
            <div className="flex-1 h-px bg-gradient-to-r from-border/40 via-border/25 to-transparent" />
          </button>
          <AnimatePresence>
            {showVariablesPanel && (
              <motion.div
                initial={{ height: 0, opacity: 0 }}
                animate={{ height: 'auto', opacity: 1 }}
                exit={{ height: 0, opacity: 0 }}
                transition={{ duration: 0.15 }}
                className="overflow-hidden"
              >
                <div className="py-1 space-y-px max-h-32 overflow-y-auto custom-scrollbar">
                  {slideVariables.map((v) => (
                    <VariableChip
                      key={v.name}
                      name={v.name}
                      usageCount={v.elementIds.length}
                      currentValue={
                        v.currentValues.every((val) => val === v.currentValues[0])
                          ? v.currentValues[0]
                          : undefined
                      }
                      onClick={() => {
                        if (v.elementIds[0]) {
                          scrollToElement(v.elementIds[0]);
                          selectElement(v.elementIds[0]);
                        }
                      }}
                    />
                  ))}
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      )}

      {/* Scrollable element list */}
      {/* Notes panel - collapsible, between Variables and element list */}
      <div className="shrink-0 px-1.5 pt-1">
        <button
          onClick={() => setShowNotesPanel((v) => !v)}
          className={cn(
            'flex items-center gap-1 w-full h-5 px-1.5 rounded text-[9px] font-semibold transition-all duration-150',
            'focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
            showNotesPanel
              ? 'bg-amber-500/10 text-amber-700 dark:bg-amber-500/15 dark:text-amber-300'
              : 'bg-muted/15 text-muted-foreground/70 hover:bg-muted/25 hover:text-muted-foreground'
          )}
        >
          <motion.div
            animate={{ rotate: showNotesPanel ? 90 : 0 }}
            transition={{ duration: 0.15 }}
            className="shrink-0"
          >
            <ChevronRight className="size-2" />
          </motion.div>
          <StickyNote className="size-2" />
          <span className="uppercase tracking-wider">Notes</span>
          {(() => {
            const noteLen = (slideNotes[slide.slideNumber] || '').length;
            return noteLen > 0 ? (
              <Badge
                className={cn(
                  'shrink-0 rounded-full px-0.5 py-0 h-3 text-[7px] font-semibold',
                  'bg-amber-100/80 text-amber-700 border-amber-200/50',
                  'dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-700/30'
                )}
              >
                {noteLen}
              </Badge>
            ) : null;
          })()}
          <div className="flex-1 h-px bg-gradient-to-r from-border/40 via-border/25 to-transparent" />
        </button>
        <AnimatePresence>
          {showNotesPanel && (
            <motion.div
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: 'auto', opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              transition={{ duration: 0.15 }}
              className="overflow-hidden"
            >
              <div className="py-1">
                <Textarea
                  value={slideNotes[slide.slideNumber] || ''}
                  onChange={(e) => updateSlideNote(slide.slideNumber, e.target.value)}
                  rows={3}
                  placeholder="Add notes for this slide..."
                  className={cn(
                    'resize-y text-[9px] leading-relaxed min-h-[44px] py-1 px-1.5',
                    'focus-visible:ring-amber-500/20 focus-visible:border-amber-300/40 dark:focus-visible:border-amber-600/30',
                    'bg-muted/15 dark:bg-muted/20 hover:bg-muted/25 dark:hover:bg-muted/25 transition-colors',
                    'placeholder:text-muted-foreground/40'
                  )}
                />
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>

      {/* Scrollable element list (continued) */}
      <div
        role="tabpanel"
        className="flex-1 overflow-y-auto custom-scrollbar editor-dot-pattern px-1.5 py-1.5 space-y-1.5"
        onClick={(e) => { if (e.target === e.currentTarget) selectElement(null); }}
      >
        <AnimatePresence mode="wait">
          {visibleElements.length === 0 ? (
            <motion.div
              key="empty"
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              transition={{ duration: 0.15 }}
              className="flex flex-col items-center justify-center py-8 text-muted-foreground"
            >
              <div className={cn(
                'flex size-8 items-center justify-center rounded-lg mb-2',
                'bg-emerald-500/10'
              )}>
                <FileText className="size-4 text-emerald-500/40" />
              </div>
              <span className="text-[10px] font-medium">No editable elements</span>
              <span className="text-[9px] mt-0.5 opacity-60 max-w-[160px] text-center leading-relaxed">
                {hideEmpty ? 'Hidden empty elements — toggle to show all' : 'This slide has no content elements'}
              </span>
            </motion.div>
          ) : displayTotal === 0 ? (
            <motion.div
              key={`filter-empty-${filterType}`}
              initial={{ opacity: 0, y: 8, scale: 0.98 }}
              animate={{ opacity: 1, y: 0, scale: 1 }}
              exit={{ opacity: 0, y: -8, scale: 0.98 }}
              transition={{ duration: 0.2, ease: [0.4, 0, 0.2, 1] }}
              className="flex flex-col items-center justify-center py-10 text-muted-foreground"
            >
              <motion.div
                initial={{ scale: 0.8, opacity: 0 }}
                animate={{ scale: 1, opacity: 1 }}
                transition={{ duration: 0.25, delay: 0.05 }}
                className={cn(
                  'flex size-10 items-center justify-center rounded-xl mb-3',
                  'bg-emerald-500/10'
                )}
              >
                {filterType === 'text' && <Type className="size-5 text-emerald-500/40" />}
                {filterType === 'table' && <Table2 className="size-5 text-emerald-500/40" />}
                {filterType === 'image' && <ImageIcon className="size-5 text-emerald-500/40" />}
              </motion.div>
              <motion.span
                initial={{ opacity: 0, y: 4 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.2, delay: 0.1 }}
                className="text-[11px] font-medium text-foreground/70"
              >
                No elements visible
              </motion.span>
              <motion.span
                initial={{ opacity: 0, y: 4 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.2, delay: 0.15 }}
                className="text-[9px] mt-0.5 opacity-60 max-w-[180px] text-center leading-relaxed"
              >
                Try adjusting your search or filter
              </motion.span>
              <motion.div
                initial={{ opacity: 0, y: 4 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.2, delay: 0.2 }}
              >
                <button
                  onClick={() => { setSearchQuery(''); setFilterType('all'); }}
                  className={cn(
                    'flex items-center gap-1 mt-2 px-2 py-1 rounded-md text-[9px] font-medium transition-all',
                    'bg-emerald-500/10 text-emerald-700 dark:text-emerald-300',
                    'hover:bg-emerald-500/20 hover:shadow-sm hover:shadow-emerald-500/5',
                    'focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30'
                  )}
                >
                  <FilterX className="size-2.5" />
                  Reset filters
                </button>
              </motion.div>
            </motion.div>
          ) : (
            <motion.div
              key={`elements-${filterType}`}
              initial={{ opacity: 0, y: 4 }}
              animate={{ opacity: 1, y: 0 }}
              exit={{ opacity: 0, y: -4 }}
              transition={{ duration: 0.15 }}
              className="space-y-1.5"
            >
              {/* Selected element section */}
              {selectedElements.length > 0 && (
                <div className="space-y-1.5">
                  <SectionHeader
                    icon={<MousePointerClick className="size-2" />}
                    title="Selected"
                    count={selectedElements.length}
                    accentClass="bg-primary/10 text-primary"
                    collapsed={collapsedSections.has('selected')}
                    onToggleCollapse={() => toggleSectionCollapse('selected')}
                  />
                  <AnimatePresence initial={false}>
                    {!collapsedSections.has('selected') && (
                      <motion.div
                        initial={{ height: 0, opacity: 0 }}
                        animate={{ height: 'auto', opacity: 1 }}
                        exit={{ height: 0, opacity: 0 }}
                        transition={{ duration: 0.2, ease: [0.4, 0, 0.2, 1] }}
                        className="overflow-hidden space-y-1.5"
                      >
                        {selectedElements.map((el) => {
                          if (el.type === 'text') {
                            return (
                              <TextElementEditor
                                key={el.id}
                                element={el}
                                isExpanded={expandedIds.has(el.id)}
                                onToggle={() => toggleExpand(el.id)}
                                batchMode={batchMode}
                                isBatchSelected={batchSelectedIds.has(el.id)}
                              />
                            );
                          }
                          if (el.type === 'table') {
                            return (
                              <TableElementEditor
                                key={el.id}
                                element={el}
                                isExpanded={expandedIds.has(el.id)}
                                onToggle={() => toggleExpand(el.id)}
                                batchMode={batchMode}
                                isBatchSelected={batchSelectedIds.has(el.id)}
                              />
                            );
                          }
                          if (el.type === 'image') {
                            return (
                              <ImageElementDisplay
                                key={el.id}
                                element={el}
                                isExpanded={expandedIds.has(el.id)}
                                onToggle={() => toggleExpand(el.id)}
                                batchMode={batchMode}
                                isBatchSelected={batchSelectedIds.has(el.id)}
                              />
                            );
                          }
                          return null;
                        })}
                      </motion.div>
                    )}
                  </AnimatePresence>
                </div>
              )}

              {/* Text elements section */}
              {displayTextElements.length > 0 && (
                <DndContext
                  sensors={sensors}
                  collisionDetection={closestCenter}
                  onDragStart={handleDragStart}
                  onDragEnd={handleTextDragEnd}
                >
                  <SortableContext
                    items={displayTextElements.map((el) => el.id)}
                    strategy={verticalListSortingStrategy}
                  >
                    <div className="space-y-1.5">
                    <SectionHeader
                      icon={<Type className="size-2" />}
                      title="Text"
                      count={displayTextElements.length}
                      accentClass="bg-emerald-500/10 text-emerald-600 dark:bg-emerald-500/15 dark:text-emerald-400"
                      collapsed={collapsedSections.has('text')}
                      onToggleCollapse={() => toggleSectionCollapse('text')}
                      modifiedCount={textModCount}
                      tooltipDetail={`${displayTextElements.length} element${displayTextElements.length !== 1 ? 's' : ''}${textModCount > 0 ? `, ${textModCount} modified` : ''} · ${displayTextElements.reduce((sum, el) => sum + (el.currentText ?? el.originalText).split(/\s+/).filter(Boolean).length, 0)} words`}
                    />
                    <AnimatePresence initial={false}>
                      {!collapsedSections.has('text') && (
                        <motion.div
                          initial={{ height: 0, opacity: 0 }}
                          animate={{ height: 'auto', opacity: 1 }}
                          exit={{ height: 0, opacity: 0 }}
                          transition={{ duration: 0.2, ease: [0.4, 0, 0.2, 1] }}
                          className="overflow-hidden space-y-1.5"
                        >
                          {displayTextElements.map((el, index) => (
                            <motion.div
                              key={el.id}
                              initial={shouldAnimateEntrance ? { opacity: 0, y: 8 } : false}
                              animate={{ opacity: 1, y: 0 }}
                              transition={{ delay: index * 0.05, duration: 0.2 }}
                            >
                            <SortableElementWrapper id={el.id}>
                              {({ isDragging, handleProps }) => (
                                <TextElementEditor
                                  element={el}
                                  isExpanded={expandedIds.has(el.id)}
                                  onToggle={() => toggleExpand(el.id)}
                                  isDragging={isDragging}
                                  dragHandleProps={handleProps}
                                  batchMode={batchMode}
                                  isBatchSelected={batchSelectedIds.has(el.id)}
                                />
                              )}
                            </SortableElementWrapper>
                            </motion.div>
                          ))}
                        </motion.div>
                      )}
                    </AnimatePresence>
                    </div>
                  </SortableContext>
                </DndContext>
              )}

              {/* Table elements section */}
              {displayTableElements.length > 0 && (
                <DndContext
                  sensors={sensors}
                  collisionDetection={closestCenter}
                  onDragStart={handleDragStart}
                  onDragEnd={handleTableDragEnd}
                >
                  <SortableContext
                    items={displayTableElements.map((el) => el.id)}
                    strategy={verticalListSortingStrategy}
                  >
                    <div className="space-y-1.5">
                    <SectionHeader
                      icon={<Table2 className="size-2" />}
                      title="Tables"
                      count={displayTableElements.length}
                      accentClass="bg-amber-500/10 text-amber-600 dark:bg-amber-500/15 dark:text-amber-400"
                      collapsed={collapsedSections.has('table')}
                      onToggleCollapse={() => toggleSectionCollapse('table')}
                      modifiedCount={tableModCount}
                      tooltipDetail={`${displayTableElements.length} element${displayTableElements.length !== 1 ? 's' : ''}${tableModCount > 0 ? `, ${tableModCount} modified` : ''} · ${displayTableElements.reduce((sum, el) => sum + el.rows.reduce((rSum, row) => rSum + row.cells.length, 0), 0)} cells`}
                    />
                    <AnimatePresence initial={false}>
                      {!collapsedSections.has('table') && (
                        <motion.div
                          initial={{ height: 0, opacity: 0 }}
                          animate={{ height: 'auto', opacity: 1 }}
                          exit={{ height: 0, opacity: 0 }}
                          transition={{ duration: 0.2, ease: [0.4, 0, 0.2, 1] }}
                          className="overflow-hidden space-y-1.5"
                        >
                          {displayTableElements.map((el, index) => (
                            <motion.div
                              key={el.id}
                              initial={shouldAnimateEntrance ? { opacity: 0, y: 8 } : false}
                              animate={{ opacity: 1, y: 0 }}
                              transition={{ delay: (displayTextElements.length + index) * 0.05, duration: 0.2 }}
                            >
                            <SortableElementWrapper id={el.id}>
                              {({ isDragging, handleProps }) => (
                                <TableElementEditor
                                  element={el}
                                  isExpanded={expandedIds.has(el.id)}
                                  onToggle={() => toggleExpand(el.id)}
                                  isDragging={isDragging}
                                  dragHandleProps={handleProps}
                                  batchMode={batchMode}
                                  isBatchSelected={batchSelectedIds.has(el.id)}
                                />
                              )}
                            </SortableElementWrapper>
                            </motion.div>
                          ))}
                        </motion.div>
                      )}
                    </AnimatePresence>
                    </div>
                  </SortableContext>
                </DndContext>
              )}

              {/* Image elements section */}
              {displayImageElements.length > 0 && (
                <DndContext
                  sensors={sensors}
                  collisionDetection={closestCenter}
                  onDragStart={handleDragStart}
                  onDragEnd={handleImageDragEnd}
                >
                  <SortableContext
                    items={displayImageElements.map((el) => el.id)}
                    strategy={verticalListSortingStrategy}
                  >
                    <div className="space-y-1.5">
                    <SectionHeader
                      icon={<ImageIcon className="size-2" />}
                      title="Images"
                      count={displayImageElements.length}
                      accentClass="bg-cyan-500/10 text-cyan-600 dark:bg-cyan-500/15 dark:text-cyan-400"
                      collapsed={collapsedSections.has('image')}
                      onToggleCollapse={() => toggleSectionCollapse('image')}
                      modifiedCount={imageModCount}
                      tooltipDetail={`${displayImageElements.length} element${displayImageElements.length !== 1 ? 's' : ''}${imageModCount > 0 ? `, ${imageModCount} replaced` : ''}`}
                    />
                    <AnimatePresence initial={false}>
                      {!collapsedSections.has('image') && (
                        <motion.div
                          initial={{ height: 0, opacity: 0 }}
                          animate={{ height: 'auto', opacity: 1 }}
                          exit={{ height: 0, opacity: 0 }}
                          transition={{ duration: 0.2, ease: [0.4, 0, 0.2, 1] }}
                          className="overflow-hidden space-y-1.5"
                        >
                          {displayImageElements.map((el, index) => (
                            <motion.div
                              key={el.id}
                              initial={shouldAnimateEntrance ? { opacity: 0, y: 8 } : false}
                              animate={{ opacity: 1, y: 0 }}
                              transition={{ delay: (displayTextElements.length + displayTableElements.length + index) * 0.05, duration: 0.2 }}
                            >
                            <SortableElementWrapper id={el.id}>
                              {({ isDragging, handleProps }) => (
                                <ImageElementDisplay
                                  element={el}
                                  isExpanded={expandedIds.has(el.id)}
                                  onToggle={() => toggleExpand(el.id)}
                                  isDragging={isDragging}
                                  dragHandleProps={handleProps}
                                  batchMode={batchMode}
                                  isBatchSelected={batchSelectedIds.has(el.id)}
                                />
                              )}
                            </SortableElementWrapper>
                            </motion.div>
                          ))}
                        </motion.div>
                      )}
                    </AnimatePresence>
                    </div>
                  </SortableContext>
                </DndContext>
              )}

              {/* No search results */}
              {isSearching && filteredTotal === 0 && (
                <div className="flex flex-col items-center justify-center py-6 text-muted-foreground">
                  <Search className="size-4 mb-1.5 text-muted-foreground/50" />
                  <span className="text-[10px] font-medium">No elements match &ldquo;{searchQuery}&rdquo;</span>
                  <button
                    onClick={() => setSearchQuery('')}
                    className="text-[9px] text-emerald-600 dark:text-emerald-400 hover:underline mt-0.5 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30 rounded"
                  >
                    Clear search
                  </button>
                </div>
              )}
            </motion.div>
          )}
        </AnimatePresence>
      </div>

      {/* Keyboard shortcuts hint bar */}
      <KeyboardShortcutsHint />

      {/* Floating batch action bar */}
      <AnimatePresence>
        {batchMode && batchSelectedIds.size > 0 && (
          <motion.div
            initial={{ y: 40, opacity: 0 }}
            animate={{ y: 0, opacity: 1 }}
            exit={{ y: 40, opacity: 0 }}
            transition={{ duration: 0.2, ease: [0.4, 0, 0.2, 1] }}
            className="shrink-0 border-t border-emerald-200/30 dark:border-emerald-700/20 bg-gradient-to-r from-emerald-50/80 via-teal-50/60 to-emerald-50/80 dark:from-emerald-950/40 dark:via-teal-950/30 dark:to-emerald-950/40 backdrop-blur-sm px-2 py-1"
          >
            <div className="flex items-center gap-1 h-6">
              <Badge
                className={cn(
                  'shrink-0 rounded-full px-1.5 py-0 h-4 text-[8px] font-semibold',
                  'bg-emerald-500/20 text-emerald-700 border-emerald-300/40',
                  'dark:bg-emerald-500/25 dark:text-emerald-300 dark:border-emerald-600/30'
                )}
              >
                {batchSelectedIds.size} selected
              </Badge>
              <button
                onClick={() => {
                  const allIds = visibleElements.map((el) => el.id);
                  batchSelectAll(allIds);
                }}
                className="h-5 px-1.5 rounded text-[8px] font-semibold bg-emerald-500/15 text-emerald-700 dark:bg-emerald-500/20 dark:text-emerald-300 hover:bg-emerald-500/25 dark:hover:bg-emerald-500/30 transition-colors focus-visible:outline-none"
              >
                Select All
              </button>
              <button
                onClick={() => clearBatchSelection()}
                className="h-5 px-1.5 rounded text-[8px] font-semibold bg-muted/30 text-muted-foreground hover:bg-muted/50 hover:text-foreground transition-colors focus-visible:outline-none"
              >
                Clear
              </button>
              <button
                onClick={async () => {
                  const texts: string[] = [];
                  for (const el of slide.elements) {
                    if (!batchSelectedIds.has(el.id)) continue;
                    if (el.type === 'text') {
                      texts.push(el.currentText ?? el.originalText);
                    } else if (el.type === 'table') {
                      const rows = el.currentRows ?? el.rows;
                      texts.push(rows.map((r) => r.cells.map((c) => c.text).join('\t')).join('\n'));
                    }
                  }
                  try {
                    await navigator.clipboard.writeText(texts.join('\n\n'));
                    toast.success(`Copied ${batchSelectedIds.size} element${batchSelectedIds.size !== 1 ? 's' : ''}!`);
                  } catch {
                    // fallback
                  }
                }}
                className="h-5 px-1.5 rounded text-[8px] font-semibold bg-emerald-500/15 text-emerald-700 dark:bg-emerald-500/20 dark:text-emerald-300 hover:bg-emerald-500/25 dark:hover:bg-emerald-500/30 transition-colors flex items-center gap-0.5 focus-visible:outline-none"
              >
                <ClipboardList className="size-2.5" />
                Copy All
              </button>
              <button
                onClick={() => {
                  batchSelectedIds.forEach((id) => {
                    toggleElementVisibility(id);
                  });
                  clearBatchSelection();
                }}
                className="h-5 px-1.5 rounded text-[8px] font-semibold bg-amber-500/15 text-amber-700 dark:bg-amber-500/20 dark:text-amber-300 hover:bg-amber-500/25 dark:hover:bg-amber-500/30 transition-colors flex items-center gap-0.5 focus-visible:outline-none"
              >
                <EyeOff className="size-2.5" />
                Hide All
              </button>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </div>
  );
}
