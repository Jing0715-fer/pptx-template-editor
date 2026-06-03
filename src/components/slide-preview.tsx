'use client';

import React, { useCallback } from 'react';
import { cn } from '@/lib/utils';
import {
  type PptxElement,
  type PptxTextElement,
  type PptxTableElement,
  type PptxImageElement,
  type PptxSlideData,
  usePptxStore,
} from '@/lib/pptx-store';
import { scrollToElement } from '@/components/slide-editor';

// ============================================================================
// Color scheme — warmer tones
// ============================================================================

const ELEMENT_COLORS = {
  text: {
    border: 'rgba(20, 184, 166, 0.55)',    // teal-500
    bg: 'rgba(20, 184, 166, 0.07)',
    hoverBg: 'rgba(20, 184, 166, 0.13)',
    selectedBorder: 'rgb(20, 184, 166)',
    glow: 'rgba(20, 184, 166, 0.40)',
    dot: 'rgb(20, 184, 166)',
    label: 'teal',
  },
  table: {
    border: 'rgba(16, 185, 129, 0.55)',    // emerald-500
    bg: 'rgba(16, 185, 129, 0.07)',
    hoverBg: 'rgba(16, 185, 129, 0.13)',
    selectedBorder: 'rgb(16, 185, 129)',
    glow: 'rgba(16, 185, 129, 0.40)',
    dot: 'rgb(16, 185, 129)',
    label: 'emerald',
  },
  image: {
    border: 'rgba(139, 92, 246, 0.55)',    // violet-500
    bg: 'rgba(139, 92, 246, 0.07)',
    hoverBg: 'rgba(139, 92, 246, 0.13)',
    selectedBorder: 'rgb(139, 92, 246)',
    glow: 'rgba(139, 92, 246, 0.40)',
    dot: 'rgb(139, 92, 246)',
    label: 'violet',
  },
} as const;

type ElementTypeKey = 'text' | 'table' | 'image';

function getElementColors(type: ElementTypeKey) {
  return ELEMENT_COLORS[type];
}

// ============================================================================
// isEmptyElement — background / filler element detection
// ============================================================================

export function isEmptyElement(el: PptxElement): boolean {
  // Empty text rectangles
  if (el.type === 'text') {
    const text = (el as PptxTextElement).originalText?.trim() ?? '';
    if (text === '') return true;
  }

  // Empty tables
  if (el.type === 'table') {
    return el.rows.length === 0;
  }

  return false;
}

function hasPosition(el: PptxElement): boolean {
  return el.position.width > 0 && el.position.height > 0;
}

/**
 * Determine if an element is a "background-style" element that should NOT
 * be rendered as an interactive highlight overlay. Two patterns:
 *
 *  1. Full-bleed (>=90% on both axes) empty/no-replacement elements — these
 *     are full-slide backgrounds (WPS stores them as <p:pic> or empty rects).
 *  2. Large empty text rects (>=10% on each axis AND area >= 5%) — these
 *     are the WPS scientific-template card / panel decorations.
 *
 * The slide preview image already includes these shapes; rendering an
 * overlay on top adds no value and only confuses the user.
 */
function isBackgroundElement(el: PptxElement, slideW: number, slideH: number): boolean {
  if (el.position.width === 0 || el.position.height === 0) return false;
  const wPct = el.position.width / slideW;
  const hPct = el.position.height / slideH;
  const coversMostOfSlide = wPct >= 0.9 && hPct >= 0.9;
  const isLargeDecor = wPct >= 0.1 && hPct >= 0.1 && (wPct * hPct) >= 0.05;

  if (el.type === 'image') {
    if (coversMostOfSlide) {
      const imgEl = el as PptxImageElement;
      if (!imgEl.replacementImageData) return true;
    }
    return false;
  }

  if (el.type === 'text') {
    if (coversMostOfSlide && isEmptyElement(el)) return true;
    if (isLargeDecor && isEmptyElement(el)) return true;
  }

  return false;
}

// ============================================================================
// Modified-element helper
// ============================================================================

function isElementModified(el: PptxElement): boolean {
  if (el.type === 'text') {
    return (el as PptxTextElement).currentText !== undefined &&
           (el as PptxTextElement).currentText !== (el as PptxTextElement).originalText;
  }
  if (el.type === 'table') {
    const tableEl = el as PptxTableElement;
    if (!tableEl.currentRows) return false;
    for (let ri = 0; ri < tableEl.currentRows.length; ri++) {
      const origRow = tableEl.rows[ri];
      const curRow = tableEl.currentRows[ri];
      if (!origRow || !curRow) continue;
      for (let ci = 0; ci < curRow.cells.length; ci++) {
        if (!origRow.cells[ci] || !curRow.cells[ci]) continue;
        if (curRow.cells[ci].text !== origRow.cells[ci].text) return true;
      }
    }
    return false;
  }
  if (el.type === 'image') {
    return !!(el as PptxImageElement).replacementImageData;
  }
  return false;
}

// ============================================================================
// SlidePreview component
// ============================================================================

interface SlidePreviewProps {
  slide: PptxSlideData;
  className?: string;
}

export function SlidePreview({ slide, className }: SlidePreviewProps) {
  const { selectedElementId, selectElement, hideEmpty, slideSize } = usePptxStore();
  const [imageError, setImageError] = React.useState(false);

  const { width: slideW, height: slideH } = slideSize;

  // Filter visible elements (hideEmpty + drop background-style decorations)
  const visibleElements = (hideEmpty
    ? slide.elements.filter((el) => !isEmptyElement(el))
    : slide.elements
  ).filter((el) => hasPosition(el) && !isBackgroundElement(el, slideW, slideH));

  // Decorative overlays: small empty text shapes (WPS template placeholders).
  // Render as a separate faint dotted layer so the user can see WPS template
  // placeholder rectangles without confusing them with editable content.
  const decorOverlays = slide.elements.filter((el) => {
    if (!hideEmpty) return false;
    if (!isEmptyElement(el)) return false;
    if (!hasPosition(el)) return false;
    if (isBackgroundElement(el, slideW, slideH)) return false;
    const wPct = el.position.width / slideW;
    const hPct = el.position.height / slideH;
    return wPct < 0.5 && hPct < 0.5;
  });

  // Reset error state when slide changes
  React.useEffect(() => {
    setImageError(false);
    // Debug: log preview image status
    if (slide.previewImage) {
      console.log(`[SlidePreview] Slide ${slide.slideNumber}: previewImage exists, length=${slide.previewImage.length}, startsWith=${slide.previewImage.substring(0, 30)}`);
    } else {
      console.log(`[SlidePreview] Slide ${slide.slideNumber}: NO previewImage`);
    }
  }, [slide.slideNumber, slide.previewImage]);

  const handleClick = useCallback(
    (elementId: string) => {
      selectElement(elementId);
    },
    [selectElement],
  );

  const handleDoubleClick = useCallback(
    (elementId: string) => {
      selectElement(elementId);
      // Defer scroll so the store has time to update selectedElementId
      requestAnimationFrame(() => {
        scrollToElement(elementId);
      });
    },
    [selectElement],
  );

  return (
    <div className={cn('flex flex-col gap-2', className)}>
      {/* ── Preview container ── */}
      <div
        className="relative w-full overflow-hidden rounded-xl border border-border/40 bg-muted/30 shadow-lg shadow-black/5 dark:shadow-black/20"
        style={{ aspectRatio: `${slideW} / ${slideH}` }}
      >
        {/* Preview image or grid fallback */}
        {slide.previewImage && !imageError ? (
          <img
            src={slide.previewImage}
            alt={`Slide ${slide.slideNumber} preview`}
            className="h-full w-full object-cover select-none"
            draggable={false}
            onError={() => {
              console.warn(`Preview image failed to load for slide ${slide.slideNumber}, data length: ${slide.previewImage?.length ?? 0}`);
              setImageError(true);
            }}
            onLoad={() => {
              setImageError(false);
            }}
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

        {/* Decorative (empty) element outlines — very faint, non-interactive.
            Helps the user see WPS template card / panel decorations without
            confusing them with editable content. */}
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

        {/* Element overlays */}
        {visibleElements.map((el) => {
          const colors = getElementColors(el.type);
          const isSelected = selectedElementId === el.id;
          const isModified = isElementModified(el);

          // Percentage-based positioning
          const left = (el.position.x / slideW) * 100;
          const top = (el.position.y / slideH) * 100;
          const width = (el.position.width / slideW) * 100;
          const height = (el.position.height / slideH) * 100;

          return (
            <div
              key={el.id}
              className={cn(
                'group/overlay absolute cursor-pointer transition-all duration-200 ease-out',
                'rounded-[3px]',
              )}
              style={{
                left: `${left}%`,
                top: `${top}%`,
                width: `${width}%`,
                height: `${height}%`,
                backgroundColor: isSelected ? colors.bg : 'transparent',
                border: isSelected
                  ? `1.5px solid ${colors.selectedBorder}`
                  : `1px solid ${colors.border}`,
                boxShadow: isSelected
                  ? `0 0 0 2px ${colors.glow}, 0 0 16px ${colors.glow}, inset 0 0 8px ${colors.bg}`
                  : 'none',
                zIndex: isSelected ? 10 : 1,
              }}
              onClick={(e) => {
                e.stopPropagation();
                handleClick(el.id);
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
                  (e.currentTarget as HTMLDivElement).style.backgroundColor = 'transparent';
                  (e.currentTarget as HTMLDivElement).style.borderColor = colors.border;
                }
              }}
            >
              {/* Type indicator — top-left corner */}
              <span
                className={cn(
                  'pointer-events-none absolute -left-0.5 -top-0.5 flex items-center justify-center',
                  'rounded-[3px] px-1 py-[1px] text-[8px] font-semibold leading-none text-white opacity-0',
                  'transition-opacity duration-150',
                  isSelected && 'opacity-100',
                  'group-hover/overlay:opacity-100',
                )}
                style={{ backgroundColor: colors.dot }}
              >
                {el.type === 'text' ? 'T' : el.type === 'table' ? '#' : '🖼'}
              </span>

              {/* Modified indicator — orange dot */}
              {isModified && (
                <span
                  className="pointer-events-none absolute -right-1 -top-1 flex h-3 w-3 items-center justify-center"
                >
                  <span
                    className="block h-2.5 w-2.5 rounded-full bg-orange-500 shadow-sm"
                    style={{
                      boxShadow: '0 0 0 1.5px rgba(255,255,255,0.9), 0 0 6px rgba(249,115,22,0.5)',
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

      {/* ── Legend ── */}
      <div className="flex items-center justify-center gap-5 text-[11px] text-muted-foreground/80 bg-muted/20 rounded-lg py-1.5 px-3">
        <LegendItem color={ELEMENT_COLORS.text.dot} label="文本" />
        <LegendItem color={ELEMENT_COLORS.table.dot} label="表格" />
        <LegendItem color={ELEMENT_COLORS.image.dot} label="图片" />
        <LegendItem color="rgb(249, 115, 22)" label="已修改" dotStyle="ring" />
      </div>
    </div>
  );
}

// ============================================================================
// Legend item
// ============================================================================

function LegendItem({
  color,
  label,
  dotStyle = 'solid',
}: {
  color: string;
  label: string;
  dotStyle?: 'solid' | 'ring';
}) {
  return (
    <span className="inline-flex items-center gap-1.5">
      {dotStyle === 'solid' ? (
        <span
          className="block h-2.5 w-2.5 rounded-[2px]"
          style={{ backgroundColor: color }}
        />
      ) : (
        <span
          className="block h-2.5 w-2.5 rounded-full"
          style={{
            backgroundColor: color,
            boxShadow: '0 0 0 1.5px rgba(255,255,255,0.9)',
          }}
        />
      )}
      <span>{label}</span>
    </span>
  );
}
