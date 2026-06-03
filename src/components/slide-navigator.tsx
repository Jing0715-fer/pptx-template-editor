'use client';

import { useState, useMemo, useCallback } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { Eye, EyeOff, FileText, ChevronLeft, ChevronRight, Presentation } from 'lucide-react';
import { cn } from '@/lib/utils';
import {
  type PptxElement,
  type PptxTextElement,
  type PptxImageElement,
  type PptxSlideData,
  usePptxStore,
} from '@/lib/pptx-store';
import { isEmptyElement } from '@/components/slide-preview';
import { scrollToElement } from '@/components/slide-editor';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import {
  Tooltip,
  TooltipTrigger,
  TooltipContent,
} from '@/components/ui/tooltip';

interface SlideNavigatorProps {
  slides: PptxSlideData[];
  currentSlide: PptxSlideData | undefined;
}

// Helper to count modifications for a single slide
function getSlideModificationCount(slide: PptxSlideData): number {
  let count = 0;
  for (const el of slide.elements) {
    if (el.type === 'text' && el.currentText !== undefined && el.currentText !== el.originalText) {
      count++;
    } else if (el.type === 'table' && el.currentRows) {
      for (let ri = 0; ri < el.currentRows.length; ri++) {
        const origRow = el.rows[ri];
        const curRow = el.currentRows[ri];
        if (!origRow || !curRow) continue;
        for (let ci = 0; ci < curRow.cells.length; ci++) {
          if (!origRow.cells[ci] || !curRow.cells[ci]) continue;
          if (curRow.cells[ci].text !== origRow.cells[ci].text) count++;
        }
      }
    } else if (el.type === 'image' && el.replacementImageData) {
      count++;
    }
  }
  return count;
}

function isElementModified(el: PptxElement): boolean {
  if (el.type === 'text') {
    return (el as PptxTextElement).currentText !== undefined &&
           (el as PptxTextElement).currentText !== (el as PptxTextElement).originalText;
  }
  if (el.type === 'image') {
    return !!(el as PptxImageElement).replacementImageData;
  }
  return false;
}

// Build data URL for image element
function buildImageDataUrl(element: PptxImageElement): string | null {
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

  const buildUrl = (data: string | null | undefined, type: string | undefined): string | null => {
    if (!data) return null;
    if (data.startsWith('data:')) return data;
    const mime = toMimeType(type);
    return `data:${mime};base64,${data}`;
  };

  if (element.replacementImageData) {
    return buildUrl(element.replacementImageData, element.replacementImageType);
  }
  if (element.imageData) {
    return buildUrl(element.imageData, element.imageType);
  }
  return null;
}

// Color scheme for element overlays
const ELEMENT_COLORS = {
  text: {
    border: 'rgba(20, 184, 166, 0.55)',
    bg: 'rgba(20, 184, 166, 0.07)',
    hoverBg: 'rgba(20, 184, 166, 0.13)',
    selectedBorder: 'rgb(20, 184, 166)',
    glow: 'rgba(20, 184, 166, 0.40)',
    dot: 'rgb(20, 184, 166)',
  },
  table: {
    border: 'rgba(16, 185, 129, 0.55)',
    bg: 'rgba(16, 185, 129, 0.07)',
    hoverBg: 'rgba(16, 185, 129, 0.13)',
    selectedBorder: 'rgb(16, 185, 129)',
    glow: 'rgba(16, 185, 129, 0.40)',
    dot: 'rgb(16, 185, 129)',
  },
  image: {
    border: 'rgba(139, 92, 246, 0.55)',
    bg: 'rgba(139, 92, 246, 0.07)',
    hoverBg: 'rgba(139, 92, 246, 0.13)',
    selectedBorder: 'rgb(139, 92, 246)',
    glow: 'rgba(139, 92, 246, 0.40)',
    dot: 'rgb(139, 92, 246)',
  },
} as const;

type ElementTypeKey = 'text' | 'table' | 'image';
function getElementColors(type: ElementTypeKey) {
  return ELEMENT_COLORS[type];
}

export default function SlideNavigator({ slides, currentSlide }: SlideNavigatorProps) {
  const [collapsed, setCollapsed] = useState(false);
  const [previewErrors, setPreviewErrors] = useState<Set<number>>(new Set());
  const {
    currentSlideIndex,
    setCurrentSlide,
    hideEmpty,
    toggleHideEmpty,
    getTotalModificationCount,
    selectedElementId,
    selectElement,
    slideSize,
  } = usePptxStore();

  const totalModifications = getTotalModificationCount();
  const { width: slideW, height: slideH } = slideSize;

  const handleDoubleClick = useCallback(
    (elementId: string) => {
      selectElement(elementId);
      requestAnimationFrame(() => {
        scrollToElement(elementId);
      });
    },
    [selectElement],
  );

  // Visible elements for current slide preview (filter empty + background-style)
  const visibleElements = useMemo(() => {
    if (!currentSlide) return [];
    return (hideEmpty
      ? currentSlide.elements.filter((el) => !isEmptyElement(el))
      : currentSlide.elements
    ).filter((el) => el.position.width > 0 && el.position.height > 0).filter((el) => {
      const wPct = el.position.width / slideW;
      const hPct = el.position.height / slideH;
      const coversMost = wPct >= 0.9 && hPct >= 0.9;
      const isLargeDecor = wPct >= 0.1 && hPct >= 0.1 && (wPct * hPct) >= 0.05;
      if (el.type === 'image' && coversMost && !el.replacementImageData) return false;
      if (el.type === 'text' && isEmptyElement(el)) {
        if (coversMost) return false;
        if (isLargeDecor) return false;
      }
      return true;
    });
  }, [currentSlide, hideEmpty, slideW, slideH]);

  // Decorative overlays: small empty text shapes (WPS template placeholders)
  const decorOverlays = useMemo(() => {
    if (!currentSlide || !hideEmpty) return [];
    return currentSlide.elements.filter((el) => {
      if (!isEmptyElement(el)) return false;
      if (el.position.width <= 0 || el.position.height <= 0) return false;
      const wPct = el.position.width / slideW;
      const hPct = el.position.height / slideH;
      const coversMost = wPct >= 0.9 && hPct >= 0.9;
      if (coversMost) return false;
      const isLargeDecor = wPct >= 0.1 && hPct >= 0.1 && (wPct * hPct) >= 0.05;
      if (isLargeDecor) return false;
      return wPct < 0.5 && hPct < 0.5;
    });
  }, [currentSlide, hideEmpty, slideW, slideH]);

  return (
    <motion.aside
      className="flex h-full flex-col overflow-hidden bg-background border-r border-border/40"
      animate={{ width: collapsed ? 52 : 280 }}
      transition={{ type: 'spring', stiffness: 300, damping: 30 }}
    >
      {/* ── Header ── */}
      <div className="flex items-center gap-2 border-b border-border/50 px-3 py-3 bg-gradient-to-r from-emerald-500/5 via-teal-500/3 to-transparent flex-shrink-0">
        <AnimatePresence mode="wait">
          {!collapsed && (
            <motion.div
              key="header-content"
              initial={{ opacity: 0, x: -8 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -8 }}
              transition={{ duration: 0.15 }}
              className="flex flex-1 items-center gap-2 overflow-hidden"
            >
              <FileText className="size-4 shrink-0 text-muted-foreground" />
              <span className="truncate text-sm font-semibold">幻灯片</span>
              <Badge variant="secondary" className="ml-auto shrink-0 text-[10px] px-1.5 py-0">
                {slides.length}
              </Badge>
            </motion.div>
          )}
        </AnimatePresence>

        {/* Action buttons */}
        <div className={cn('flex items-center gap-0.5', !collapsed && 'ml-auto')}>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="ghost"
                size="icon"
                className="size-7"
                onClick={toggleHideEmpty}
              >
                {hideEmpty ? (
                  <EyeOff className="size-3.5 text-muted-foreground" />
                ) : (
                  <Eye className="size-3.5 text-muted-foreground" />
                )}
              </Button>
            </TooltipTrigger>
            <TooltipContent side="right" sideOffset={8}>
              {hideEmpty ? '显示空元素' : '隐藏空元素'}
            </TooltipContent>
          </Tooltip>

          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="ghost"
                size="icon"
                className="size-7"
                onClick={() => setCollapsed((c) => !c)}
              >
                {collapsed ? (
                  <ChevronRight className="size-3.5 text-muted-foreground" />
                ) : (
                  <ChevronLeft className="size-3.5 text-muted-foreground" />
                )}
              </Button>
            </TooltipTrigger>
            <TooltipContent side="right" sideOffset={8}>
              {collapsed ? '展开侧栏' : '收起侧栏'}
            </TooltipContent>
          </Tooltip>
        </div>
      </div>

      {/* ── Slide list ── */}
      <div className="custom-scrollbar flex-1 min-h-0 overflow-y-auto px-2 py-2">
        <AnimatePresence mode="popLayout">
          {slides.map((slide, index) => {
            const isActive = index === currentSlideIndex;
            const modCount = getSlideModificationCount(slide);

            // Collapsed: compact number strip
            if (collapsed) {
              return (
                <Tooltip key={slide.slideNumber}>
                  <TooltipTrigger asChild>
                    <motion.button
                      layout
                      whileHover={{ scale: 1.08 }}
                      whileTap={{ scale: 0.95 }}
                      onClick={() => setCurrentSlide(index)}
                      className={cn(
                        'relative mb-1 flex size-10 items-center justify-center rounded-lg text-xs font-medium transition-all duration-200',
                        isActive
                          ? 'bg-primary text-primary-foreground shadow-sm shadow-primary/25'
                          : 'text-muted-foreground hover:bg-accent hover:text-accent-foreground'
                      )}
                    >
                      {slide.slideNumber}
                      {modCount > 0 && (
                        <span className="absolute -right-0.5 -top-0.5 flex size-3.5 items-center justify-center rounded-full bg-amber-500 text-[8px] font-bold text-white">
                          {modCount}
                        </span>
                      )}
                    </motion.button>
                  </TooltipTrigger>
                  <TooltipContent side="right" sideOffset={8}>
                    第 {slide.slideNumber} 页
                    {modCount > 0 && ` · ${modCount} 处修改`}
                  </TooltipContent>
                </Tooltip>
              );
            }

            // Expanded: compact card row
            return (
              <motion.div
                key={slide.slideNumber}
                layout
                whileHover={{ scale: 1.01 }}
                whileTap={{ scale: 0.98 }}
                initial={{ opacity: 0, y: 8 }}
                animate={{ opacity: 1, y: 0 }}
                exit={{ opacity: 0, y: -8 }}
                transition={{ duration: 0.2 }}
                onClick={() => setCurrentSlide(index)}
                className={cn(
                  'group relative mb-1.5 cursor-pointer overflow-hidden rounded-lg border transition-all duration-200',
                  isActive
                    ? 'border-primary/60 bg-primary/5 shadow-sm shadow-primary/10'
                    : 'border-transparent hover:border-border hover:bg-accent/40'
                )}
              >
                {isActive && (
                  <motion.div
                    layoutId="slide-accent-bar"
                    className="absolute bottom-0 left-0 right-0 h-[2px] bg-gradient-to-r from-emerald-500 to-teal-500"
                    transition={{ type: 'spring', stiffness: 400, damping: 30 }}
                  />
                )}

                <div className="flex items-center gap-2.5 px-2.5 py-2">
                  {/* Thumbnail */}
                  <div className="relative w-16 h-9 overflow-hidden rounded bg-muted/50 flex-shrink-0">
                    {slide.previewImage && !previewErrors.has(slide.slideNumber) ? (
                      <img
                        src={slide.previewImage}
                        alt={`Slide ${slide.slideNumber}`}
                        className="h-full w-full object-cover"
                        draggable={false}
                        onError={() => {
                          setPreviewErrors(prev => new Set(prev).add(slide.slideNumber));
                        }}
                      />
                    ) : (
                      <div className="flex h-full w-full items-center justify-center bg-gradient-to-br from-muted/40 to-muted/10">
                        <FileText className="size-4 text-muted-foreground/20" />
                      </div>
                    )}
                    <div
                      className={cn(
                        'absolute bottom-0.5 left-0.5 flex items-center justify-center rounded px-1 py-0 text-[8px] font-semibold backdrop-blur-sm',
                        isActive
                          ? 'bg-primary/90 text-primary-foreground'
                          : 'bg-black/50 text-white'
                      )}
                    >
                      {slide.slideNumber}
                    </div>
                  </div>

                  {/* Info */}
                  <div className="flex-1 min-w-0">
                    <span className={cn(
                      'text-xs font-medium block truncate',
                      isActive ? 'text-emerald-600 dark:text-emerald-400' : 'text-foreground'
                    )}>
                      第 {slide.slideNumber} 页
                    </span>
                    <span className="text-[10px] text-muted-foreground block truncate">
                      {slide.elements.length} 个元素
                    </span>
                  </div>

                  {/* Modification badge */}
                  {modCount > 0 && (
                    <Badge className="bg-amber-500 text-white border-0 text-[8px] px-1 py-0 h-4 shrink-0 shadow-sm hover:bg-amber-600">
                      {modCount}
                    </Badge>
                  )}
                </div>
              </motion.div>
            );
          })}
        </AnimatePresence>
      </div>

      {/* ── Current Slide Preview (interactive) ── */}
      {!collapsed && currentSlide && (
        <div className="flex-shrink-0 border-t border-border/50">
          <div className="px-2.5 py-2">
            <div className="flex items-center gap-1.5 mb-1.5">
              <Presentation className="size-3 text-muted-foreground" />
              <span className="text-[11px] font-semibold text-muted-foreground">实时预览</span>
              <Badge variant="secondary" className="text-[9px] px-1 py-0 h-3.5 ml-auto">
                第 {currentSlide.slideNumber} 页
              </Badge>
            </div>

            {/* Preview container with aspect ratio */}
            <div
              className="relative w-full overflow-hidden rounded-lg border border-border/40 bg-muted/30 shadow-md shadow-black/5 dark:shadow-black/20"
              style={{ aspectRatio: `${slideW} / ${slideH}` }}
            >
              {/* Background preview image */}
              {currentSlide.previewImage ? (
                <img
                  src={currentSlide.previewImage}
                  alt={`Slide ${currentSlide.slideNumber} preview`}
                  className="h-full w-full object-cover select-none"
                  draggable={false}
                />
              ) : (
                <div
                  className="absolute inset-0"
                  style={{
                    backgroundImage: `
                      linear-gradient(to right, hsl(var(--border) / 0.25) 1px, transparent 1px),
                      linear-gradient(to bottom, hsl(var(--border) / 0.25) 1px, transparent 1px)
                    `,
                    backgroundSize: '10% 10%',
                  }}
                />
              )}

              {/* Decorative (empty) element outlines — very faint, non-interactive */}
              {decorOverlays.map((el) => {
                const left = (el.position.x / slideW) * 100;
                const top = (el.position.y / slideH) * 100;
                const width = (el.position.width / slideW) * 100;
                const height = (el.position.height / slideH) * 100;
                if (width < 0.1 || height < 0.1) return null;
                return (
                  <div
                    key={`decor-${el.id}`}
                    className="pointer-events-none absolute rounded-[2px]"
                    style={{
                      left: `${left}%`,
                      top: `${top}%`,
                      width: `${width}%`,
                      height: `${height}%`,
                      border: '1px dotted rgba(120, 120, 120, 0.35)',
                      backgroundColor: 'transparent',
                      zIndex: 1,
                    }}
                    title={`${el.shapeName || '装饰区域'} (不可编辑)`}
                  />
                );
              })}

              {/* Element overlays — interactive, with real-time preview */}
              {visibleElements.map((el) => {
                const colors = getElementColors(el.type);
                const isSelected = selectedElementId === el.id;
                const isModified = isElementModified(el);

                const left = (el.position.x / slideW) * 100;
                const top = (el.position.y / slideH) * 100;
                const width = (el.position.width / slideW) * 100;
                const height = (el.position.height / slideH) * 100;

                // For image elements with replacement data, show the new image
                const isImageWithReplacement = el.type === 'image' && (el as PptxImageElement).replacementImageData;
                const imageUrl = isImageWithReplacement ? buildImageDataUrl(el as PptxImageElement) : null;

                return (
                  <div
                    key={el.id}
                    className={cn(
                      'group/overlay absolute cursor-pointer transition-all duration-150 ease-out',
                      'rounded-[2px]',
                    )}
                    style={{
                      left: `${left}%`,
                      top: `${top}%`,
                      width: `${width}%`,
                      height: `${height}%`,
                      border: isSelected
                        ? `1.5px solid ${colors.selectedBorder}`
                        : `1px solid ${colors.border}`,
                      boxShadow: isSelected
                        ? `0 0 0 1.5px ${colors.glow}, 0 0 8px ${colors.glow}`
                        : 'none',
                      zIndex: isSelected ? 10 : 1,
                      backgroundColor: isImageWithReplacement && imageUrl
                        ? 'transparent'
                        : isSelected
                        ? colors.bg
                        : 'transparent',
                    }}
                    onClick={(e) => {
                      e.stopPropagation();
                      selectElement(el.id);
                    }}
                    onDoubleClick={(e) => {
                      e.stopPropagation();
                      handleDoubleClick(el.id);
                    }}
                    onMouseEnter={(e) => {
                      if (!isSelected) {
                        (e.currentTarget as HTMLDivElement).style.backgroundColor = colors.hoverBg;
                        (e.currentTarget as HTMLDivElement).style.borderColor = colors.selectedBorder;
                      }
                    }}
                    onMouseLeave={(e) => {
                      if (!isSelected) {
                        (e.currentTarget as HTMLDivElement).style.backgroundColor = isImageWithReplacement && imageUrl ? 'transparent' : 'transparent';
                        (e.currentTarget as HTMLDivElement).style.borderColor = colors.border;
                      }
                    }}
                  >
                    {/* Show replacement image inside overlay */}
                    {isImageWithReplacement && imageUrl && (
                      <img
                        src={imageUrl}
                        alt="Replacement"
                        className="absolute inset-0 w-full h-full object-cover rounded-[1px]"
                        style={{ opacity: 0.85 }}
                      />
                    )}

                    {/* Show text content overlay for text elements when selected or modified */}
                    {el.type === 'text' && (isSelected || isModified) && (() => {
                      const textEl = el as PptxTextElement;
                      const displayText = textEl.currentText ?? textEl.originalText;
                      if (!displayText) return null;
                      return (
                        <div
                          className="absolute inset-0 overflow-hidden rounded-[1px]"
                          style={{
                            backgroundColor: 'rgba(0,0,0,0.6)',
                            padding: '1px',
                          }}
                        >
                          <span
                            className="block text-white leading-tight"
                            style={{
                              fontSize: '6px',
                              overflow: 'hidden',
                              textOverflow: 'ellipsis',
                              display: '-webkit-box',
                              WebkitLineClamp: 3,
                              WebkitBoxOrient: 'vertical',
                              wordBreak: 'break-all',
                            }}
                          >
                            {displayText}
                          </span>
                        </div>
                      );
                    })()}

                    {/* Type indicator dot */}
                    <span
                      className={cn(
                        'pointer-events-none absolute -left-0.5 -top-0.5 flex items-center justify-center',
                        'rounded-[2px] px-0.5 py-[0.5px] text-[6px] font-semibold leading-none text-white opacity-0',
                        'transition-opacity duration-100',
                        isSelected && 'opacity-100',
                        'group-hover/overlay:opacity-100',
                      )}
                      style={{ backgroundColor: colors.dot }}
                    >
                      {el.type === 'text' ? 'T' : el.type === 'table' ? '#' : '🖼'}
                    </span>

                    {/* Modified indicator */}
                    {isModified && (
                      <span className="pointer-events-none absolute -right-0.5 -top-0.5 flex h-2.5 w-2.5 items-center justify-center">
                        <span
                          className="block h-2 w-2 rounded-full bg-orange-500"
                          style={{
                            boxShadow: '0 0 0 1px rgba(255,255,255,0.9), 0 0 4px rgba(249,115,22,0.5)',
                          }}
                        />
                      </span>
                    )}
                  </div>
                );
              })}

              {/* Click on empty area deselects */}
              <div
                className="absolute inset-0"
                style={{ zIndex: 0 }}
                onClick={() => selectElement(null)}
              />
            </div>

            {/* Legend */}
            <div className="flex items-center justify-center gap-3 text-[9px] text-muted-foreground/80 mt-1.5">
              <span className="inline-flex items-center gap-1">
                <span className="block h-2 w-2 rounded-[1px]" style={{ backgroundColor: ELEMENT_COLORS.text.dot }} />
                文本
              </span>
              <span className="inline-flex items-center gap-1">
                <span className="block h-2 w-2 rounded-[1px]" style={{ backgroundColor: ELEMENT_COLORS.table.dot }} />
                表格
              </span>
              <span className="inline-flex items-center gap-1">
                <span className="block h-2 w-2 rounded-[1px]" style={{ backgroundColor: ELEMENT_COLORS.image.dot }} />
                图片
              </span>
              <span className="inline-flex items-center gap-1">
                <span className="block h-2 w-2 rounded-full bg-orange-500" />
                已修改
              </span>
            </div>
          </div>
        </div>
      )}

      {/* ── Footer ── */}
      <AnimatePresence mode="wait">
        {!collapsed && (
          <motion.div
            key="footer"
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.15 }}
            className="border-t border-border/50 px-3 py-1.5 flex-shrink-0"
          >
            <div className="flex items-center justify-between text-[10px] text-muted-foreground">
              <span>{slides.reduce((sum, slide) => sum + slide.elements.length, 0)} 个元素</span>
              <span>
                {totalModifications > 0 ? (
                  <span className="font-medium text-amber-600 dark:text-amber-400">
                    {totalModifications} 处修改
                  </span>
                ) : (
                  '无修改'
                )}
              </span>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </motion.aside>
  );
}
