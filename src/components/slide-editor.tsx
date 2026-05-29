'use client'

import React from 'react';
import {
  type PptxTextElement,
  type PptxTableElement,
  type PptxImageElement,
  type PptxElement,
  type PptxSlideData,
  usePptxStore,
} from '@/lib/pptx-store';
import { isEmptyElement } from '@/components/slide-preview';
import { cn } from '@/lib/utils';
import { toast } from 'sonner';
import { motion, AnimatePresence } from 'framer-motion';
import { Card, CardContent } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Textarea } from '@/components/ui/textarea';
import { Input } from '@/components/ui/input';
import { ScrollArea } from '@/components/ui/scroll-area';
import { Separator } from '@/components/ui/separator';
import {
  Tooltip,
  TooltipContent,
  TooltipProvider,
  TooltipTrigger,
} from '@/components/ui/tooltip';
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

// ============================================================================
// Element Ref Registry (for scroll-to-element functionality)
// ============================================================================

const elementRefs = new Map<string, HTMLDivElement>();

export function registerElementRef(id: string, element: HTMLDivElement | null): void {
  if (element) {
    elementRefs.set(id, element);
  } else {
    elementRefs.delete(id);
  }
}

export function scrollToElement(id: string): void {
  const el = elementRefs.get(id);
  if (el) {
    el.scrollIntoView({ behavior: 'smooth', block: 'center' });
    el.classList.add('ring-2', 'ring-primary', 'ring-offset-2');
    setTimeout(() => {
      el.classList.remove('ring-2', 'ring-primary', 'ring-offset-2');
    }, 1500);
  }
}

// ============================================================================
// AnimatedCard wrapper
// ============================================================================

function AnimatedCard({
  children,
  className,
  delay = 0,
  id,
  ref: refProp,
}: {
  children: React.ReactNode;
  className?: string;
  delay?: number;
  id?: string;
  ref?: React.Ref<HTMLDivElement>;
}) {
  return (
    <motion.div
      id={id}
      ref={refProp}
      initial={{ opacity: 0, y: 12 }}
      animate={{ opacity: 1, y: 0 }}
      exit={{ opacity: 0, y: -8 }}
      transition={{ duration: 0.25, delay, ease: 'easeOut' }}
      className={cn('rounded-lg border bg-card text-card-foreground shadow-sm transition-all', className)}
    >
      {children}
    </motion.div>
  );
}

// ============================================================================
// Font info helper
// ============================================================================

function getFontInfo(el: PptxTextElement): string[] {
  const info: string[] = [];
  const fontNames = new Set<string>();
  const fontSizes = new Set<number>();
  let hasBold = false;
  let hasItalic = false;

  for (const para of el.paragraphs) {
    for (const run of para.runs) {
      if (run.fontName) fontNames.add(run.fontName);
      if (run.fontSize) fontSizes.add(run.fontSize);
      if (run.bold) hasBold = true;
      if (run.italic) hasItalic = true;
    }
  }

  if (fontNames.size > 0) info.push([...fontNames].join(', '));
  if (fontSizes.size > 0) info.push([...fontSizes].map((s) => `${s}pt`).join(', '));
  const styles: string[] = [];
  if (hasBold) styles.push('B');
  if (hasItalic) styles.push('I');
  if (styles.length > 0) info.push(styles.join(' '));

  return info;
}

// ============================================================================
// TextElementEditor
// ============================================================================

function TextElementEditor({ element }: { element: PptxTextElement }) {
  const { updateText, selectElement, selectedElementId } = usePptxStore();
  const [text, setText] = React.useState(element.currentText ?? element.originalText);
  const [isExpanded, setIsExpanded] = React.useState(true);
  const textAreaRef = React.useRef<HTMLTextAreaElement>(null);
  const isSelected = selectedElementId === element.id;

  const isModified = element.currentText !== undefined && element.currentText !== element.originalText;
  const fontInfo = getFontInfo(element);

  React.useEffect(() => {
    setText(element.currentText ?? element.originalText);
  }, [element.currentText, element.originalText]);

  const handleTextChange = (newText: string) => {
    setText(newText);
    updateText(element.id, newText);
  };

  const handleReset = () => {
    setText(element.originalText);
    updateText(element.id, element.originalText);
    toast.success('已重置为原始文本');
  };

  const handleCopy = async () => {
    try {
      await navigator.clipboard.writeText(element.originalText);
      toast.success('已复制原始文本');
    } catch {
      toast.error('复制失败');
    }
  };

  const lineCount = text.split('\n').length;
  const charCount = text.length;

  return (
    <AnimatedCard
      id={`element-${element.id}`}
      ref={(el) => registerElementRef(element.id, el)}
      className={cn(
        'overflow-hidden',
        isSelected && 'ring-2 ring-primary/50',
        isModified && 'border-orange-300 dark:border-orange-700',
      )}
    >
      {/* Header */}
      <button
        className="w-full flex items-center gap-2 px-4 py-3 hover:bg-muted/50 transition-colors"
        onClick={() => setIsExpanded(!isExpanded)}
      >
        <div className="flex items-center justify-center w-6 h-6 rounded bg-indigo-100 dark:bg-indigo-900/30 shrink-0">
          <Type className="w-3.5 h-3.5 text-indigo-600 dark:text-indigo-400" />
        </div>
        <div className="flex-1 min-w-0 text-left">
          <div className="flex items-center gap-2">
            <span className="text-sm font-medium truncate">
              {element.shapeName || '文本框'}
            </span>
            {isModified && (
              <Badge variant="secondary" className="text-[10px] px-1.5 py-0 bg-orange-100 text-orange-700 dark:bg-orange-900/30 dark:text-orange-400 shrink-0">
                已修改
              </Badge>
            )}
          </div>
          {fontInfo.length > 0 && (
            <div className="flex items-center gap-1 mt-0.5">
              {fontInfo.map((info, i) => (
                <span key={i} className="text-[10px] text-muted-foreground">
                  {i > 0 && <span className="mx-0.5">·</span>}
                  {info}
                </span>
              ))}
            </div>
          )}
        </div>
        <div className="flex items-center gap-1 shrink-0">
          {isExpanded ? (
            <ChevronDown className="w-4 h-4 text-muted-foreground" />
          ) : (
            <ChevronRight className="w-4 h-4 text-muted-foreground" />
          )}
        </div>
      </button>

      {/* Content */}
      <AnimatePresence>
        {isExpanded && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.2 }}
          >
            <div className="px-4 pb-4 space-y-3">
              {/* Original text preview */}
              <div className="space-y-1.5">
                <div className="flex items-center justify-between">
                  <label className="text-xs font-medium text-muted-foreground">原始文本</label>
                  <TooltipProvider>
                    <Tooltip>
                      <TooltipTrigger asChild>
                        <Button variant="ghost" size="sm" className="h-6 px-1.5" onClick={handleCopy}>
                          <Copy className="w-3 h-3" />
                        </Button>
                      </TooltipTrigger>
                      <TooltipContent>复制原始文本</TooltipContent>
                    </Tooltip>
                  </TooltipProvider>
                </div>
                <div className="text-xs text-muted-foreground bg-muted/50 rounded p-2.5 max-h-24 overflow-y-auto whitespace-pre-wrap break-all leading-relaxed">
                  {element.originalText || <span className="italic text-muted-foreground/50">空文本</span>}
                </div>
              </div>

              {/* Edit area */}
              <div className="space-y-1.5">
                <div className="flex items-center justify-between">
                  <label className="text-xs font-medium text-muted-foreground">编辑文本</label>
                  {isModified && (
                    <TooltipProvider>
                      <Tooltip>
                        <TooltipTrigger asChild>
                          <Button variant="ghost" size="sm" className="h-6 px-1.5 text-orange-600 hover:text-orange-700" onClick={handleReset}>
                            <RotateCcw className="w-3 h-3" />
                          </Button>
                        </TooltipTrigger>
                        <TooltipContent>重置为原始文本</TooltipContent>
                      </Tooltip>
                    </TooltipProvider>
                  )}
                </div>
                <Textarea
                  ref={textAreaRef}
                  value={text}
                  onChange={(e) => handleTextChange(e.target.value)}
                  placeholder="输入替换文本..."
                  className={cn(
                    'min-h-[80px] text-sm leading-relaxed resize-y',
                    isModified && 'border-orange-300 focus-visible:ring-orange-300 dark:border-orange-700',
                  )}
                  rows={Math.min(Math.max(lineCount + 1, 3), 12)}
                />
                <div className="flex items-center justify-between text-[10px] text-muted-foreground">
                  <span>{charCount} 字符 · {lineCount} 行</span>
                  {isModified && (
                    <span className="text-orange-600 dark:text-orange-400">
                      已修改 {element.currentText!.length - element.originalText.length > 0 ? '+' : ''}
                      {element.currentText!.length - element.originalText.length} 字符
                    </span>
                  )}
                </div>
              </div>

              {/* Paragraph runs detail */}
              {element.paragraphs.length > 1 && (
                <details className="group">
                  <summary className="text-[10px] text-muted-foreground cursor-pointer hover:text-foreground transition-colors flex items-center gap-1">
                    <ChevronRight className="w-3 h-3 group-open:rotate-90 transition-transform" />
                    段落详情 ({element.paragraphs.length} 段)
                  </summary>
                  <div className="mt-2 space-y-1.5 max-h-40 overflow-y-auto">
                    {element.paragraphs.map((para, pi) => (
                      <div key={pi} className="text-[10px] bg-muted/30 rounded p-2">
                        <div className="flex items-center gap-1 mb-1 text-muted-foreground">
                          <span>段落 {pi + 1}</span>
                          {para.runs.length > 1 && <span>· {para.runs.length} 个文本段</span>}
                        </div>
                        <div className="flex flex-wrap gap-1">
                          {para.runs.map((run, ri) => (
                            <span
                              key={ri}
                              className={cn(
                                'inline-block px-1 py-0.5 rounded text-[9px] border',
                                run.bold && 'font-bold',
                                run.italic && 'italic',
                                'bg-background border-border',
                              )}
                              style={run.fontColor ? { color: `#${run.fontColor}` } : undefined}
                            >
                              {run.originalText || '(空)'}
                              {run.fontSize && <span className="ml-0.5 opacity-50">{run.fontSize}pt</span>}
                            </span>
                          ))}
                        </div>
                      </div>
                    ))}
                  </div>
                </details>
              )}
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </AnimatedCard>
  );
}

// ============================================================================
// TableElementEditor
// ============================================================================

function TableElementEditor({ element }: { element: PptxTableElement }) {
  const { updateTableCell, selectElement, selectedElementId } = usePptxStore();
  const [isExpanded, setIsExpanded] = React.useState(true);
  const isSelected = selectedElementId === element.id;

  const currentRows = element.currentRows ?? element.rows;
  const isModified =
    element.currentRows !== undefined &&
    element.currentRows.some((r, ri) =>
      r.cells.some((c, ci) => c.text !== element.rows[ri]?.cells[ci]?.text)
    );

  const modifiedCellCount = React.useMemo(() => {
    if (!element.currentRows) return 0;
    let count = 0;
    for (let ri = 0; ri < element.currentRows.length; ri++) {
      const origRow = element.rows[ri];
      const curRow = element.currentRows[ri];
      if (!origRow || !curRow) continue;
      for (let ci = 0; ci < curRow.cells.length; ci++) {
        if (curRow.cells[ci]?.text !== origRow.cells[ci]?.text) count++;
      }
    }
    return count;
  }, [element.currentRows, element.rows]);

  const handleCellChange = (row: number, col: number, text: string) => {
    updateTableCell(element.id, row, col, text);
  };

  const handleReset = () => {
    for (let ri = 0; ri < element.rows.length; ri++) {
      for (let ci = 0; ci < element.rows[ri].cells.length; ci++) {
        updateTableCell(element.id, ri, ci, element.rows[ri].cells[ci].text);
      }
    }
    toast.success('已重置表格为原始数据');
  };

  const handleCopyTable = async () => {
    try {
      const tsv = element.rows
        .map((r) => r.cells.map((c) => c.text).join('\t'))
        .join('\n');
      await navigator.clipboard.writeText(tsv);
      toast.success('已复制表格数据');
    } catch {
      toast.error('复制失败');
    }
  };

  const maxCols = Math.max(...currentRows.map((r) => r.cells.length));

  return (
    <AnimatedCard
      id={`element-${element.id}`}
      ref={(el) => registerElementRef(element.id, el)}
      className={cn(
        'overflow-hidden',
        isSelected && 'ring-2 ring-emerald-500/50',
        isModified && 'border-orange-300 dark:border-orange-700',
      )}
    >
      {/* Header */}
      <button
        className="w-full flex items-center gap-2 px-4 py-3 hover:bg-muted/50 transition-colors"
        onClick={() => setIsExpanded(!isExpanded)}
      >
        <div className="flex items-center justify-center w-6 h-6 rounded bg-emerald-100 dark:bg-emerald-900/30 shrink-0">
          <Table2 className="w-3.5 h-3.5 text-emerald-600 dark:text-emerald-400" />
        </div>
        <div className="flex-1 min-w-0 text-left">
          <div className="flex items-center gap-2">
            <span className="text-sm font-medium truncate">
              {element.shapeName || '表格'}
            </span>
            {isModified && (
              <Badge variant="secondary" className="text-[10px] px-1.5 py-0 bg-orange-100 text-orange-700 dark:bg-orange-900/30 dark:text-orange-400 shrink-0">
                {modifiedCellCount} 处修改
              </Badge>
            )}
          </div>
          <span className="text-[10px] text-muted-foreground">
            {currentRows.length} 行 × {maxCols} 列
          </span>
        </div>
        <div className="flex items-center gap-1 shrink-0">
          {isExpanded ? (
            <ChevronDown className="w-4 h-4 text-muted-foreground" />
          ) : (
            <ChevronRight className="w-4 h-4 text-muted-foreground" />
          )}
        </div>
      </button>

      {/* Content */}
      <AnimatePresence>
        {isExpanded && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.2 }}
          >
            <div className="px-4 pb-4 space-y-3">
              {/* Action buttons */}
              <div className="flex items-center gap-2">
                <TooltipProvider>
                  <Tooltip>
                    <TooltipTrigger asChild>
                      <Button variant="outline" size="sm" className="h-7 text-xs" onClick={handleCopyTable}>
                        <Copy className="w-3 h-3 mr-1" /> 复制
                      </Button>
                    </TooltipTrigger>
                    <TooltipContent>复制为 TSV 格式</TooltipContent>
                  </Tooltip>
                </TooltipProvider>
                {isModified && (
                  <TooltipProvider>
                    <Tooltip>
                      <TooltipTrigger asChild>
                        <Button variant="outline" size="sm" className="h-7 text-xs text-orange-600 hover:text-orange-700" onClick={handleReset}>
                          <RotateCcw className="w-3 h-3 mr-1" /> 重置
                        </Button>
                      </TooltipTrigger>
                      <TooltipContent>重置表格为原始数据</TooltipContent>
                    </Tooltip>
                  </TooltipProvider>
                )}
              </div>

              {/* Table grid */}
              <div className="overflow-x-auto rounded border">
                <table className="w-full text-xs border-collapse">
                  <tbody>
                    {currentRows.map((row, ri) => {
                      const origRow = element.rows[ri];
                      return (
                        <tr key={ri} className={ri % 2 === 0 ? 'bg-muted/30' : ''}>
                          <td className="px-2 py-1 text-muted-foreground text-[10px] text-center border-r bg-muted/50 font-medium w-8">
                            {ri + 1}
                          </td>
                          {row.cells.map((cell, ci) => {
                            const origCell = origRow?.cells[ci];
                            const cellModified = origCell && cell.text !== origCell.text;
                            return (
                              <td key={ci} className="border-r last:border-r-0 p-0">
                                <input
                                  type="text"
                                  value={cell.text}
                                  onChange={(e) => handleCellChange(ri, ci, e.target.value)}
                                  className={cn(
                                    'w-full px-2 py-1.5 text-xs bg-transparent outline-none',
                                    'focus:bg-background focus:ring-1 focus:ring-emerald-400',
                                    cellModified && 'bg-orange-50 dark:bg-orange-950/20 text-orange-700 dark:text-orange-300',
                                  )}
                                  placeholder="..."
                                />
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>

              {/* Original table toggle */}
              <details className="group">
                <summary className="text-[10px] text-muted-foreground cursor-pointer hover:text-foreground transition-colors flex items-center gap-1">
                  <ChevronRight className="w-3 h-3 group-open:rotate-90 transition-transform" />
                  查看原始表格数据
                </summary>
                <div className="mt-2 overflow-x-auto rounded border">
                  <table className="w-full text-[10px] border-collapse">
                    <tbody>
                      {element.rows.map((row, ri) => (
                        <tr key={ri} className={ri % 2 === 0 ? 'bg-muted/20' : ''}>
                          {row.cells.map((cell, ci) => (
                            <td key={ci} className="border px-2 py-1 text-muted-foreground">
                              {cell.text || <span className="italic opacity-50">空</span>}
                            </td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </details>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </AnimatedCard>
  );
}

// ============================================================================
// ImageElementDisplay
// ============================================================================

function ImageElementDisplay({ element }: { element: PptxImageElement }) {
  const { updateImage, removeImage, selectedElementId } = usePptxStore();
  const [isExpanded, setIsExpanded] = React.useState(true);
  const [previewUrl, setPreviewUrl] = React.useState<string | null>(null);
  const [isZoomed, setIsZoomed] = React.useState(false);
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const isSelected = selectedElementId === element.id;
  const isModified = !!element.replacementImageData;

  React.useEffect(() => {
    if (element.replacementImageData) {
      // replacementImageData may be a full data URL or raw base64
      if (element.replacementImageData.startsWith('data:')) {
        // Already a complete data URL (from FileReader.readAsDataURL)
        setPreviewUrl(element.replacementImageData);
      } else {
        // Raw base64, needs data URL prefix
        const type = element.replacementImageType || 'png';
        const mime = type === 'jpg' ? 'jpeg' : type;
        setPreviewUrl(`data:image/${mime};base64,${element.replacementImageData}`);
      }
    } else if (element.imageData) {
      // imageData from server already includes data URL prefix
      setPreviewUrl(element.imageData);
    } else {
      setPreviewUrl(null);
    }
  }, [element.replacementImageData, element.replacementImageType, element.imageData]);

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (!file.type.startsWith('image/')) {
      toast.error('请选择图片文件');
      return;
    }

    if (file.size > 10 * 1024 * 1024) {
      toast.error('图片文件大小不能超过 10MB');
      return;
    }

    const reader = new FileReader();
    reader.onload = (ev) => {
      const base64 = ev.target?.result as string;
      const ext = file.name.split('.').pop()?.toLowerCase() || 'png';
      const typeMap: Record<string, string> = {
        png: 'png', jpg: 'jpeg', jpeg: 'jpeg', gif: 'gif', bmp: 'bmp', webp: 'webp', svg: 'svg',
      };
      const imageType = typeMap[ext] || 'png';
      updateImage(element.id, base64, imageType);
      toast.success('图片已替换');
    };
    reader.readAsDataURL(file);

    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleRemoveReplacement = () => {
    removeImage(element.id);
    toast.success('已恢复原始图片');
  };

  return (
    <AnimatedCard
      id={`element-${element.id}`}
      ref={(el) => registerElementRef(element.id, el)}
      className={cn(
        'overflow-hidden',
        isSelected && 'ring-2 ring-purple-500/50',
        isModified && 'border-orange-300 dark:border-orange-700',
      )}
    >
      {/* Header */}
      <button
        className="w-full flex items-center gap-2 px-4 py-3 hover:bg-muted/50 transition-colors"
        onClick={() => setIsExpanded(!isExpanded)}
      >
        <div className="flex items-center justify-center w-6 h-6 rounded bg-purple-100 dark:bg-purple-900/30 shrink-0">
          <ImageIcon className="w-3.5 h-3.5 text-purple-600 dark:text-purple-400" />
        </div>
        <div className="flex-1 min-w-0 text-left">
          <div className="flex items-center gap-2">
            <span className="text-sm font-medium truncate">
              {element.shapeName || '图片'}
            </span>
            {isModified && (
              <Badge variant="secondary" className="text-[10px] px-1.5 py-0 bg-orange-100 text-orange-700 dark:bg-orange-900/30 dark:text-orange-400 shrink-0">
                已替换
              </Badge>
            )}
          </div>
          <span className="text-[10px] text-muted-foreground">
            {element.imageName || '未命名图片'} · {element.imageType.toUpperCase()}
          </span>
        </div>
        <div className="flex items-center gap-1 shrink-0">
          {isExpanded ? (
            <ChevronDown className="w-4 h-4 text-muted-foreground" />
          ) : (
            <ChevronRight className="w-4 h-4 text-muted-foreground" />
          )}
        </div>
      </button>

      {/* Content */}
      <AnimatePresence>
        {isExpanded && (
          <motion.div
            initial={{ height: 0, opacity: 0 }}
            animate={{ height: 'auto', opacity: 1 }}
            exit={{ height: 0, opacity: 0 }}
            transition={{ duration: 0.2 }}
          >
            <div className="px-4 pb-4 space-y-3">
              {/* Image preview */}
              <div className="relative group rounded-md overflow-hidden border bg-muted/30 flex items-center justify-center min-h-[100px]">
                {previewUrl ? (
                  <>
                    <img
                      src={previewUrl}
                      alt={element.imageName || 'Image'}
                      className="max-w-full max-h-48 object-contain cursor-pointer"
                      onClick={() => setIsZoomed(true)}
                    />
                    <Button
                      variant="secondary"
                      size="sm"
                      className="absolute top-2 right-2 h-7 w-7 p-0 opacity-0 group-hover:opacity-100 transition-opacity"
                      onClick={() => setIsZoomed(true)}
                    >
                      <ZoomIn className="w-3.5 h-3.5" />
                    </Button>
                  </>
                ) : (
                  <div className="flex flex-col items-center gap-2 py-6 text-muted-foreground">
                    <ImageIcon className="w-8 h-8 opacity-30" />
                    <span className="text-xs">无法预览此图片</span>
                  </div>
                )}
              </div>

              {/* Actions */}
              <div className="flex items-center gap-2">
                <input
                  ref={fileInputRef}
                  type="file"
                  accept="image/*"
                  className="hidden"
                  onChange={handleFileSelect}
                />
                <Button
                  variant="outline"
                  size="sm"
                  className="h-7 text-xs"
                  onClick={() => fileInputRef.current?.click()}
                >
                  <Upload className="w-3 h-3 mr-1" /> 替换图片
                </Button>
                {isModified && (
                  <Button
                    variant="outline"
                    size="sm"
                    className="h-7 text-xs text-orange-600 hover:text-orange-700"
                    onClick={handleRemoveReplacement}
                  >
                    <RotateCcw className="w-3 h-3 mr-1" /> 恢复原图
                  </Button>
                )}
              </div>

              {/* Image info */}
              <div className="text-[10px] text-muted-foreground space-y-0.5">
                <div className="flex justify-between">
                  <span>文件名</span>
                  <span className="font-mono">{element.imageName}</span>
                </div>
                <div className="flex justify-between">
                  <span>格式</span>
                  <span className="font-mono">{element.imageType.toUpperCase()}</span>
                </div>
                <div className="flex justify-between">
                  <span>引用 ID</span>
                  <span className="font-mono">{element.imageRid}</span>
                </div>
              </div>
            </div>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Zoom dialog */}
      {isZoomed && previewUrl && (
        <div
          className="fixed inset-0 z-50 bg-black/80 flex items-center justify-center p-4"
          onClick={() => setIsZoomed(false)}
        >
          <div className="relative max-w-[90vw] max-h-[90vh]">
            <img
              src={previewUrl}
              alt={element.imageName || 'Image'}
              className="max-w-full max-h-[85vh] object-contain rounded-lg"
            />
            <Button
              variant="secondary"
              size="sm"
              className="absolute -top-3 -right-3 h-8 w-8 rounded-full p-0"
              onClick={() => setIsZoomed(false)}
            >
              <X className="w-4 h-4" />
            </Button>
          </div>
        </div>
      )}
    </AnimatedCard>
  );
}

// ============================================================================
// EmptyState
// ============================================================================

function EmptyState() {
  return (
    <div className="flex flex-col items-center justify-center py-16 px-4 text-center">
      <div className="w-16 h-16 rounded-full bg-muted/50 flex items-center justify-center mb-4">
        <MousePointerClick className="w-8 h-8 text-muted-foreground/40" />
      </div>
      <h3 className="text-sm font-medium text-muted-foreground mb-1">选择元素进行编辑</h3>
      <p className="text-xs text-muted-foreground/60 max-w-[240px]">
        点击左侧预览图上的元素，或滚动浏览下方所有可编辑元素
      </p>
    </div>
  );
}

// ============================================================================
// SlideEditor (main component)
// ============================================================================

interface SlideEditorProps {
  slide: PptxSlideData;
}

export function SlideEditor({ slide }: SlideEditorProps) {
  const { hideEmpty, selectElement, selectedElementId } = usePptxStore();
  const scrollRef = React.useRef<HTMLDivElement>(null);

  const visibleElements = slide.elements.filter((el) => !(hideEmpty && isEmptyElement(el)));

  const textElements = visibleElements.filter((el): el is PptxTextElement => el.type === 'text');
  const tableElements = visibleElements.filter((el): el is PptxTableElement => el.type === 'table');
  const imageElements = visibleElements.filter((el): el is PptxImageElement => el.type === 'image');

  const selectedElement = visibleElements.find((el) => el.id === selectedElementId);

  // Scroll to selected element when it changes
  React.useEffect(() => {
    if (selectedElementId) {
      const timer = setTimeout(() => scrollToElement(selectedElementId), 50);
      return () => clearTimeout(timer);
    }
  }, [selectedElementId]);

  return (
    <div className="flex flex-col h-full">
      {/* Header */}
      <div className="px-4 py-3 border-b bg-muted/30">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-2">
            <FileText className="w-4 h-4 text-muted-foreground" />
            <span className="text-sm font-medium">第 {slide.slideNumber} 页编辑器</span>
          </div>
          <div className="flex items-center gap-2 text-[10px] text-muted-foreground">
            <span className="flex items-center gap-0.5">
              <Type className="w-3 h-3" /> {textElements.length}
            </span>
            <span className="flex items-center gap-0.5">
              <Table2 className="w-3 h-3" /> {tableElements.length}
            </span>
            <span className="flex items-center gap-0.5">
              <ImageIcon className="w-3 h-3" /> {imageElements.length}
            </span>
          </div>
        </div>
      </div>

      {/* Scrollable element list */}
      <ScrollArea className="flex-1" ref={scrollRef}>
        <div className="p-4 space-y-3">
          {visibleElements.length === 0 && (
            <div className="text-center py-12 text-muted-foreground text-sm">
              本页没有可编辑的元素
            </div>
          )}

          {/* Selected element section (pinned to top) */}
          {selectedElement && (
            <div className="space-y-2">
              <div className="flex items-center gap-1.5 text-[10px] text-muted-foreground font-medium uppercase tracking-wider">
                <div className="w-3 h-0.5 bg-primary rounded-full" />
                当前选中
              </div>
              {selectedElement.type === 'text' && <TextElementEditor element={selectedElement} />}
              {selectedElement.type === 'table' && <TableElementEditor element={selectedElement} />}
              {selectedElement.type === 'image' && <ImageElementDisplay element={selectedElement} />}
              <Separator className="my-3" />
            </div>
          )}

          {/* Text elements */}
          {textElements.length > 0 && (
            <div className="space-y-2">
              <div className="flex items-center gap-1.5 text-[10px] text-muted-foreground font-medium uppercase tracking-wider">
                <Type className="w-3 h-3" /> 文本元素 ({textElements.length})
              </div>
              {textElements
                .filter((el) => el.id !== selectedElementId)
                .map((el) => (
                  <TextElementEditor key={el.id} element={el} />
                ))}
            </div>
          )}

          {/* Table elements */}
          {tableElements.length > 0 && (
            <div className="space-y-2">
              <div className="flex items-center gap-1.5 text-[10px] text-muted-foreground font-medium uppercase tracking-wider">
                <Table2 className="w-3 h-3" /> 表格元素 ({tableElements.length})
              </div>
              {tableElements
                .filter((el) => el.id !== selectedElementId)
                .map((el) => (
                  <TableElementEditor key={el.id} element={el} />
                ))}
            </div>
          )}

          {/* Image elements */}
          {imageElements.length > 0 && (
            <div className="space-y-2">
              <div className="flex items-center gap-1.5 text-[10px] text-muted-foreground font-medium uppercase tracking-wider">
                <ImageIcon className="w-3 h-3" /> 图片元素 ({imageElements.length})
              </div>
              {imageElements
                .filter((el) => el.id !== selectedElementId)
                .map((el) => (
                  <ImageElementDisplay key={el.id} element={el} />
                ))}
            </div>
          )}
        </div>
      </ScrollArea>
    </div>
  );
}
