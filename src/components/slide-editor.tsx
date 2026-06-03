'use client';

import { useState, useCallback, useRef, useMemo } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import {
  Copy,
  RotateCcw,
  Type,
  Table2,
  Image as ImageIcon,
  MousePointerClick,
  FileText,
  Bold,
  Italic,
  ChevronDown,
  ChevronRight,
  Upload,
  X,
  ZoomIn,
} from 'lucide-react';
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
import { Card, CardContent } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Textarea } from '@/components/ui/textarea';
import { Input } from '@/components/ui/input';
import { ScrollArea } from '@/components/ui/scroll-area';
import { Separator } from '@/components/ui/separator';
import {
  Tooltip,
  TooltipProvider,
  TooltipTrigger,
  TooltipContent,
} from '@/components/ui/tooltip';

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

function scrollToElement(id: string) {
  const el = elementRefs.get(id);
  if (el) {
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    // Brief highlight effect
    el.classList.add('ring-2', 'ring-primary/50');
    setTimeout(() => {
      el.classList.remove('ring-2', 'ring-primary/50');
    }, 1500);
  }
}

// ============================================================================
// Text Element Editor
// ============================================================================

interface TextElementEditorProps {
  element: PptxTextElement;
  isExpanded: boolean;
  onToggle: () => void;
}

function TextElementEditor({ element, isExpanded: _isExpanded, onToggle: _onToggle }: TextElementEditorProps) {
  const { updateText } = usePptxStore();
  const [showOriginal, setShowOriginal] = useState(false);

  const currentText = element.currentText ?? element.originalText;
  const isModified = element.currentText !== undefined && element.currentText !== element.originalText;

  const handleTextChange = useCallback(
    (e: React.ChangeEvent<HTMLTextAreaElement>) => {
      updateText(element.id, e.target.value);
    },
    [updateText, element.id]
  );

  const handleReset = useCallback(() => {
    updateText(element.id, element.originalText);
  }, [updateText, element.id, element.originalText]);

  const handleCopy = useCallback(async () => {
    try {
      await navigator.clipboard.writeText(currentText);
    } catch {
      // fallback
    }
  }, [currentText]);

  return (
    <Card
      ref={(el) => registerElementRef(element.id, el)}
      className={cn(
        'group overflow-hidden transition-all duration-200',
        'border border-border/60 hover:border-emerald-300/50 dark:hover:border-emerald-700/40 hover:shadow-sm hover:shadow-emerald-500/5',
        isModified && 'border-amber-300/70 dark:border-amber-600/40 hover:border-amber-400/80 dark:hover:border-amber-500/50'
      )}
    >
      {/* Row 1: Icon + Name + Modified badge + action buttons */}
      <div className="flex items-center gap-2 px-3 py-2">
        <div className={cn(
          'flex size-6 shrink-0 items-center justify-center rounded-md',
          'bg-emerald-100 text-emerald-600',
          'dark:bg-emerald-900/40 dark:text-emerald-400'
        )}>
          <Type className="size-3" />
        </div>

        <div className="flex-1 min-w-0 flex items-center gap-1.5">
          <span className="text-xs font-medium truncate">{element.shapeName}</span>
          {isModified && (
            <Badge
              className={cn(
                'shrink-0 rounded-full px-1 py-0 h-3.5 text-[9px] font-semibold',
                'bg-amber-100 text-amber-700 border-amber-200/60',
                'dark:bg-amber-900/40 dark:text-amber-300 dark:border-amber-700/40'
              )}
            >
              Modified
            </Badge>
          )}
        </div>

        <div className="flex items-center gap-0.5 shrink-0">
          <TooltipProvider delayDuration={300}>
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="icon"
                  className="size-6 text-muted-foreground hover:text-foreground"
                  onClick={handleCopy}
                >
                  <Copy className="size-3" />
                </Button>
              </TooltipTrigger>
              <TooltipContent side="bottom">
                <p>Copy text</p>
              </TooltipContent>
            </Tooltip>
          </TooltipProvider>

          {isModified && (
            <TooltipProvider delayDuration={300}>
              <Tooltip>
                <TooltipTrigger asChild>
                  <Button
                    variant="ghost"
                    size="icon"
                    className={cn(
                      'size-6',
                      'text-amber-600 hover:text-amber-700 dark:text-amber-400 dark:hover:text-amber-300'
                    )}
                    onClick={handleReset}
                  >
                    <RotateCcw className="size-3" />
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
      <div className="px-3 pb-2">
        <Textarea
          value={currentText}
          onChange={handleTextChange}
          rows={2}
          className={cn(
            'resize-y text-xs leading-relaxed',
            'focus-visible:ring-emerald-500/30',
            isModified && 'border-amber-300/70 dark:border-amber-600/40'
          )}
        />
      </div>

      {/* Show Original toggle + display */}
      <div className="px-3 pb-2">
        <Button
          variant="ghost"
          size="sm"
          className="h-5 gap-1 text-[10px] text-muted-foreground hover:text-foreground px-1"
          onClick={() => setShowOriginal((v) => !v)}
        >
          <FileText className="size-2.5" />
          {showOriginal ? 'Hide Original' : 'Show Original'}
        </Button>

        <AnimatePresence>
          {showOriginal && (
            <motion.div
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: 'auto', opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              transition={{ duration: 0.2 }}
              className="overflow-hidden"
            >
              <div className={cn(
                'rounded-md border p-2 text-[11px] leading-relaxed mt-1',
                'bg-muted/40 border-border/50 text-muted-foreground'
              )}>
                <span className="text-[9px] font-semibold uppercase tracking-wider text-muted-foreground/70 block mb-1">
                  Original
                </span>
                {element.originalText || <span className="italic opacity-50">Empty</span>}
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </div>
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
}

function TableElementEditor({ element, isExpanded, onToggle }: TableElementEditorProps) {
  const { updateTableCell } = usePptxStore();
  const [showOriginal, setShowOriginal] = useState(false);

  const currentRows = element.currentRows ?? element.rows;
  const isModified = element.currentRows !== undefined;

  // Count modified cells
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
    // Reset by re-applying original values
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
        'group overflow-hidden transition-all duration-200',
        'border border-border/60 hover:border-emerald-300/50 dark:hover:border-emerald-700/40 hover:shadow-sm hover:shadow-emerald-500/5',
        isModified && 'border-amber-300/70 dark:border-amber-600/40 hover:border-amber-400/80 dark:hover:border-amber-500/50',
        isExpanded && 'shadow-sm shadow-emerald-500/5'
      )}
    >
      {/* Header */}
      <button
        onClick={onToggle}
        className="flex w-full items-center gap-3 px-4 py-3 text-left transition-colors hover:bg-emerald-50/30 dark:hover:bg-emerald-950/20"
      >
        <div className={cn(
          'flex size-8 shrink-0 items-center justify-center rounded-lg',
          'bg-amber-100 text-amber-600',
          'dark:bg-amber-900/40 dark:text-amber-400'
        )}>
          <Table2 className="size-4" />
        </div>

        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2">
            <span className="text-sm font-medium truncate">{element.shapeName}</span>
            {isModified && (
              <Badge
                className={cn(
                  'shrink-0 rounded-full px-1.5 py-0 h-4 text-[10px] font-semibold',
                  'bg-amber-100 text-amber-700 border-amber-200/60',
                  'dark:bg-amber-900/40 dark:text-amber-300 dark:border-amber-700/40'
                )}
              >
                {modifiedCellCount} cell{modifiedCellCount !== 1 ? 's' : ''}
              </Badge>
            )}
          </div>
          <p className="text-xs text-muted-foreground truncate mt-0.5">
            {element.rows.length} × {element.rows[0]?.cells.length ?? 0} · {totalCells} cells
          </p>
        </div>

        <motion.div
          animate={{ rotate: isExpanded ? 0 : -90 }}
          transition={{ duration: 0.2 }}
          className="shrink-0 text-muted-foreground"
        >
          <ChevronDown className="size-4" />
        </motion.div>
      </button>

      {/* Expandable content */}
      <AnimatePresence initial={false}>
        {isExpanded && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.25, ease: [0.4, 0, 0.2, 1] }}
            className="overflow-hidden"
          >
            <div className="px-4 pb-4 space-y-3">
              <Separator className="opacity-50" />

              {/* Action buttons */}
              <div className="flex items-center gap-1.5">
                <TooltipProvider delayDuration={300}>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Button
                        variant="ghost"
                        size="sm"
                        className="h-7 gap-1.5 text-xs text-muted-foreground hover:text-foreground"
                        onClick={handleCopyTable}
                      >
                        <Copy className="size-3" />
                        Copy Table
                      </Button>
                    </TooltipTrigger>
                    <TooltipContent side="bottom">
                      <p>Copy table content as tab-separated text</p>
                    </TooltipContent>
                  </Tooltip>
                </TooltipProvider>

                <TooltipProvider delayDuration={300}>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Button
                        variant="ghost"
                        size="sm"
                        className="h-7 gap-1.5 text-xs text-muted-foreground hover:text-foreground"
                        onClick={() => setShowOriginal((v) => !v)}
                      >
                        <FileText className="size-3" />
                        {showOriginal ? 'Hide Original' : 'Show Original'}
                      </Button>
                    </TooltipTrigger>
                    <TooltipContent side="bottom">
                      <p>Toggle original table display</p>
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
                            'h-7 gap-1.5 text-xs',
                            'text-amber-600 hover:text-amber-700 dark:text-amber-400 dark:hover:text-amber-300'
                          )}
                          onClick={handleReset}
                        >
                          <RotateCcw className="size-3" />
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
              <div className="overflow-x-auto custom-scrollbar rounded-lg border border-border/50">
                <table className="w-full text-xs border-collapse">
                  <tbody>
                    {currentRows.map((row, ri) => (
                      <tr key={ri} className="border-b border-border/30 last:border-b-0">
                        {row.cells.map((cell, ci) => {
                          const cellModified = isCellModified(ri, ci);
                          return (
                            <td
                              key={ci}
                              className={cn(
                                'border-r border-border/30 last:border-r-0 p-0',
                                cellModified && 'bg-amber-50/80 dark:bg-amber-900/20'
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
                                  'w-full px-2 py-1.5 text-xs bg-transparent outline-none',
                                  'placeholder:text-muted-foreground/40',
                                  'focus:bg-accent/30 transition-colors',
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
                    transition={{ duration: 0.2 }}
                    className="overflow-hidden"
                  >
                    <div className="overflow-x-auto custom-scrollbar rounded-lg border border-border/50">
                      <table className="w-full text-xs border-collapse">
                        <tbody>
                          {element.rows.map((row, ri) => (
                            <tr key={ri} className="border-b border-border/30 last:border-b-0">
                              {row.cells.map((cell, ci) => (
                                <td
                                  key={ci}
                                  className="border-r border-border/30 last:border-r-0 px-2 py-1.5 text-muted-foreground"
                                  style={{
                                    rowSpan: cell.rowSpan > 1 ? cell.rowSpan : undefined,
                                    colSpan: cell.colSpan > 1 ? cell.colSpan : undefined,
                                  }}
                                >
                                  {cell.text || <span className="italic opacity-40">—</span>}
                                </td>
                              ))}
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    <span className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground/70 block mt-2">
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
}

function ImageElementDisplay({ element, isExpanded, onToggle }: ImageElementDisplayProps) {
  const { updateImage, removeImage } = usePptxStore();
  const [isZoomed, setIsZoomed] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const isModified = !!element.replacementImageData;
  const isEmfOrWmf = /emf|wmf/i.test(element.imageType);

  const imageDataUrl = useMemo(() => {
    // Helper to ensure imageType is a proper MIME type
    const toMimeType = (t: string | undefined): string => {
      if (!t) return 'image/png';
      if (t.includes('/')) return t; // Already a MIME type like "image/png"
      const extMap: Record<string, string> = {
        png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg',
        gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml',
        tiff: 'image/tiff', tif: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf',
        webp: 'image/webp',
      };
      return extMap[t.toLowerCase()] || `image/${t.toLowerCase()}`;
    };

    // Helper: if data already contains a full data URL, return as-is; otherwise build one
    const buildDataUrl = (data: string | null | undefined, type: string | undefined): string | null => {
      if (!data) return null;
      // Already a full data URL (backward compat / edge case)
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

      // Reset input so the same file can be selected again
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
        'group overflow-hidden transition-all duration-200',
        'border border-border/60 hover:border-emerald-300/50 dark:hover:border-emerald-700/40 hover:shadow-sm hover:shadow-emerald-500/5',
        isModified && 'border-amber-300/70 dark:border-amber-600/40 hover:border-amber-400/80 dark:hover:border-amber-500/50',
        isExpanded && 'shadow-sm shadow-emerald-500/5'
      )}
    >
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
        onClick={onToggle}
        className="flex w-full items-center gap-3 px-4 py-3 text-left transition-colors hover:bg-emerald-50/30 dark:hover:bg-emerald-950/20"
      >
        <div className={cn(
          'size-8 shrink-0 rounded-lg overflow-hidden',
          !imageDataUrl && 'flex items-center justify-center bg-cyan-100 text-cyan-600 dark:bg-cyan-900/40 dark:text-cyan-400'
        )}>
          {imageDataUrl ? (
            <img
              src={imageDataUrl}
              alt={element.imageName}
              className="object-cover w-full h-full"
            />
          ) : (
            <ImageIcon className="size-4" />
          )}
        </div>

        <div className="flex-1 min-w-0">
          <div className="flex items-center gap-2">
            <span className="text-sm font-medium truncate">{element.shapeName}</span>
            {isModified && (
              <Badge
                className={cn(
                  'shrink-0 rounded-full px-1.5 py-0 h-4 text-[10px] font-semibold',
                  'bg-amber-100 text-amber-700 border-amber-200/60',
                  'dark:bg-amber-900/40 dark:text-amber-300 dark:border-amber-700/40'
                )}
              >
                Replaced
              </Badge>
            )}
          </div>
          <p className="text-xs text-muted-foreground truncate mt-0.5">
            {element.imageName}
          </p>
        </div>

        <motion.div
          animate={{ rotate: isExpanded ? 0 : -90 }}
          transition={{ duration: 0.2 }}
          className="shrink-0 text-muted-foreground"
        >
          <ChevronDown className="size-4" />
        </motion.div>
      </button>

      {/* Expandable content */}
      <AnimatePresence initial={false}>
        {isExpanded && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.25, ease: [0.4, 0, 0.2, 1] }}
            className="overflow-hidden"
          >
            <div className="px-4 pb-4 space-y-3">
              <Separator className="opacity-50" />

              {/* Image preview */}
              <div className="space-y-1.5">
                <span className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground/70">
                  Preview
                </span>
                <div
                  className={cn(
                    'relative group/img overflow-hidden rounded-lg border border-border/50',
                    'bg-muted/20'
                  )}
                >
                  {imageDataUrl ? (
                    <>
                      <img
                        src={imageDataUrl}
                        alt={element.imageName}
                        className={cn(
                          'w-full object-contain transition-transform duration-300',
                          isZoomed ? 'max-h-none scale-150 cursor-zoom-out' : 'max-h-52 cursor-zoom-in'
                        )}
                        onClick={() => setIsZoomed((z) => !z)}
                      />
                      <div className="absolute top-2 right-2 opacity-0 group-hover/img:opacity-100 transition-opacity">
                        <TooltipProvider delayDuration={200}>
                          <Tooltip>
                            <TooltipTrigger asChild>
                              <Button
                                variant="secondary"
                                size="icon"
                                className="size-7 shadow-sm"
                                onClick={() => setIsZoomed((z) => !z)}
                              >
                                <ZoomIn className="size-3.5" />
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
                    <div className="flex flex-col items-center justify-center py-8 text-muted-foreground">
                      <ImageIcon className="size-8 mb-2 opacity-30" />
                      <span className="text-xs">
                        EMF/WMF format — preview unavailable
                      </span>
                    </div>
                  ) : (
                    <div className="flex flex-col items-center justify-center py-8 text-muted-foreground">
                      <ImageIcon className="size-8 mb-2 opacity-30" />
                      <span className="text-xs">No preview available</span>
                    </div>
                  )}
                </div>
              </div>

              {/* Action buttons */}
              <div className="flex items-center gap-1.5">
                <TooltipProvider delayDuration={300}>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Button
                        variant="outline"
                        size="sm"
                        className="h-7 gap-1.5 text-xs"
                        onClick={handleReplaceImage}
                      >
                        <Upload className="size-3" />
                        Replace Image
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
                            'h-7 gap-1.5 text-xs',
                            'text-amber-600 hover:text-amber-700 dark:text-amber-400 dark:hover:text-amber-300'
                          )}
                          onClick={handleRemoveReplacement}
                        >
                          <X className="size-3" />
                          Remove Replacement
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
                'rounded-lg border p-3 space-y-1',
                'bg-muted/30 border-border/40'
              )}>
                <span className="text-[10px] font-semibold uppercase tracking-wider text-muted-foreground/70 block mb-1.5">
                  Image Info
                </span>
                <div className="grid grid-cols-2 gap-x-4 gap-y-1 text-xs">
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
                      <span className="font-medium text-amber-600 dark:text-amber-400">
                        Vector format
                      </span>
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
}

function SectionHeader({ icon, title, count, accentClass }: SectionHeaderProps) {
  return (
    <div className="flex items-center gap-2.5 py-2.5 px-1">
      <div className={cn(
        'flex size-5 items-center justify-center rounded-md shrink-0 shadow-sm',
        accentClass ?? 'bg-muted text-muted-foreground'
      )}>
        {icon}
      </div>
      <span className="text-[11px] font-bold uppercase tracking-widest text-muted-foreground/70">
        {title}
      </span>
      <Badge
        variant="secondary"
        className="h-4 px-1.5 text-[10px] font-medium shadow-sm"
      >
        {count}
      </Badge>
      <div className="flex-1 h-px bg-gradient-to-r from-border/40 via-border/20 to-transparent" />
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
  } = usePptxStore();

  const [manuallyExpandedIds, setManuallyExpandedIds] = useState<Set<string>>(new Set());

  // Compute effective expanded set: manually expanded + auto-expanded selected element
  const expandedIds = useMemo(() => {
    const ids = new Set(manuallyExpandedIds);
    if (selectedElementId) {
      ids.add(selectedElementId);
    }
    return ids;
  }, [manuallyExpandedIds, selectedElementId]);

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

  // Total counts
  const totalText = textElements.length;
  const totalTable = tableElements.length;
  const totalImage = imageElements.length;
  const totalAll = totalText + totalTable + totalImage;

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

  return (
    <div className="flex flex-col h-full overflow-hidden max-w-full">
      {/* Header */}
      <div className="shrink-0 border-b border-border/40 px-4 py-3 editor-header-gradient">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="flex size-7 items-center justify-center rounded-lg bg-gradient-to-br from-emerald-500 to-teal-600 shadow-sm shadow-emerald-500/20">
              <FileText className="size-3.5 text-white" />
            </div>
            <div>
              <span className="text-sm font-semibold tracking-tight">Slide {slide.slideNumber}</span>
              <Badge variant="secondary" className="h-5 px-2 text-[10px] font-medium ml-1.5 shadow-sm">
                {totalAll} element{totalAll !== 1 ? 's' : ''}
              </Badge>
            </div>
          </div>
          {modCount > 0 && (
            <Badge
              className={cn(
                'rounded-full px-2 py-0 h-5 text-[11px] font-semibold',
                'bg-amber-100 text-amber-700 border-amber-200/60',
                'dark:bg-amber-900/40 dark:text-amber-300 dark:border-amber-700/40'
              )}
            >
              {modCount} edit{modCount !== 1 ? 's' : ''}
            </Badge>
          )}
        </div>
        <div className="flex items-center gap-3 mt-2 ml-[38px]">
          {totalText > 0 && (
            <div className="flex items-center gap-1.5">
              <div className="size-2 rounded-full bg-emerald-500" />
              <span className="text-[11px] text-muted-foreground">{totalText} text</span>
            </div>
          )}
          {totalTable > 0 && (
            <div className="flex items-center gap-1.5">
              <div className="size-2 rounded-full bg-amber-500" />
              <span className="text-[11px] text-muted-foreground">{totalTable} table{totalTable !== 1 ? 's' : ''}</span>
            </div>
          )}
          {totalImage > 0 && (
            <div className="flex items-center gap-1.5">
              <div className="size-2 rounded-full bg-cyan-500" />
              <span className="text-[11px] text-muted-foreground">{totalImage} image{totalImage !== 1 ? 's' : ''}</span>
            </div>
          )}
        </div>
      </div>

      {/* Scrollable element list */}
      <div className="flex-1 overflow-y-auto custom-scrollbar px-3 py-3 space-y-2">
        {visibleElements.length === 0 ? (
          <div className="flex flex-col items-center justify-center py-16 text-muted-foreground">
            <div className={cn(
              'flex size-14 items-center justify-center rounded-xl mb-4',
              'bg-gradient-to-br from-muted/60 to-muted/30 ring-1 ring-border/20'
            )}>
              <FileText className="size-6 opacity-30" />
            </div>
            <span className="text-sm font-medium">No editable elements</span>
            <span className="text-xs mt-1.5 opacity-50 max-w-[200px] text-center leading-relaxed">
              {hideEmpty ? 'Hidden empty elements — toggle to show all' : 'This slide has no content elements'}
            </span>
          </div>
        ) : (
          <>
            {/* Selected element section */}
            {selectedElements.length > 0 && (
              <div className="space-y-2">
                <SectionHeader
                  icon={<MousePointerClick className="size-3" />}
                  title="Selected"
                  count={selectedElements.length}
                  accentClass="bg-primary/10 text-primary"
                />
                {selectedElements.map((el) => {
                  if (el.type === 'text') {
                    return (
                      <TextElementEditor
                        key={el.id}
                        element={el}
                        isExpanded={expandedIds.has(el.id)}
                        onToggle={() => toggleExpand(el.id)}
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
                      />
                    );
                  }
                  return null;
                })}
              </div>
            )}

            {/* Text elements section */}
            {textElements.length > 0 && (
              <div className="space-y-2">
                <SectionHeader
                  icon={<Type className="size-3" />}
                  title="Text"
                  count={totalText}
                  accentClass="bg-emerald-100 text-emerald-600 dark:bg-emerald-900/40 dark:text-emerald-400"
                />
                {textElements.map((el) => (
                  <TextElementEditor
                    key={el.id}
                    element={el}
                    isExpanded={expandedIds.has(el.id)}
                    onToggle={() => toggleExpand(el.id)}
                  />
                ))}
              </div>
            )}

            {/* Table elements section */}
            {tableElements.length > 0 && (
              <div className="space-y-2">
                <SectionHeader
                  icon={<Table2 className="size-3" />}
                  title="Tables"
                  count={totalTable}
                  accentClass="bg-amber-100 text-amber-600 dark:bg-amber-900/40 dark:text-amber-400"
                />
                {tableElements.map((el) => (
                  <TableElementEditor
                    key={el.id}
                    element={el}
                    isExpanded={expandedIds.has(el.id)}
                    onToggle={() => toggleExpand(el.id)}
                  />
                ))}
              </div>
            )}

            {/* Image elements section */}
            {imageElements.length > 0 && (
              <div className="space-y-2">
                <SectionHeader
                  icon={<ImageIcon className="size-3" />}
                  title="Images"
                  count={totalImage}
                  accentClass="bg-cyan-100 text-cyan-600 dark:bg-cyan-900/40 dark:text-cyan-400"
                />
                {imageElements.map((el) => (
                  <ImageElementDisplay
                    key={el.id}
                    element={el}
                    isExpanded={expandedIds.has(el.id)}
                    onToggle={() => toggleExpand(el.id)}
                  />
                ))}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

// Export the scroll-to utility for external use
export { scrollToElement, registerElementRef };
