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
import { EyeOff } from 'lucide-react';
import { scrollToElement } from '@/components/slide-editor';

// ============================================================================
// Color scheme
// ============================================================================

const ELEMENT_COLORS = {
  text: {
    border: 'rgba(20, 184, 166, 0.5)',
    bg: 'rgba(20, 184, 166, 0.06)',
    hoverBg: 'rgba(20, 184, 166, 0.12)',
    selectedBorder: 'rgb(20, 184, 166)',
    glow: 'rgba(20, 184, 166, 0.35)',
    dot: 'rgb(20, 184, 166)',
    label: 'teal',
  },
  table: {
    border: 'rgba(16, 185, 129, 0.5)',
    bg: 'rgba(16, 185, 129, 0.06)',
    hoverBg: 'rgba(16, 185, 129, 0.12)',
    selectedBorder: 'rgb(16, 185, 129)',
    glow: 'rgba(16, 185, 129, 0.35)',
    dot: 'rgb(16, 185, 129)',
    label: 'emerald',
  },
  image: {
    border: 'rgba(139, 92, 246, 0.5)',
    bg: 'rgba(139, 92, 246, 0.06)',
    hoverBg: 'rgba(139, 92, 246, 0.12)',
    selectedBorder: 'rgb(139, 92, 246)',
    glow: 'rgba(139, 92, 246, 0.35)',
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
  if (el.type === 'text') {
    const text = (el as PptxTextElement).originalText?.trim() ?? '';
    if (text === '') return true;
    if (text.length <= 1 && !text.trim()) return true;
  }

  if (el.type === 'table') {
    if (el.rows.length === 0) return true;
    const allEmpty = el.rows.every(row =>
      row.cells.every(cell => !cell.text?.trim())
    );
    if (allEmpty) return true;
  }

  return false;
}

function hasPosition(el: PptxElement): boolean {
  return el.position.width > 0 && el.position.height > 0;
}

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
  showGridOverlay?: boolean;
}

export function SlidePreview({ slide, className, showGridOverlay }: SlidePreviewProps) {
  const { selectedElementId, selectElement, hideEmpty, slideSize, hiddenElementIds } = usePptxStore();
  const [imageError, setImageError] = React.useState(false);

  const { width: slideW, height: slideH } = slideSize;

  const visibleElements = (hideEmpty
    ? slide.elements.filter((el) => !isEmptyElement(el))
    : slide.elements
  ).filter((el) => hasPosition(el) && !isBackgroundElement(el, slideW, slideH) && !hiddenElementIds.has(el.id));

  const hiddenElements = (hideEmpty
    ? slide.elements.filter((el) => !isEmptyElement(el))
    : slide.elements
  ).filter((el) => hasPosition(el) && !isBackgroundElement(el, slideW, slideH) && hiddenElementIds.has(el.id));

  const decorOverlays = slide.elements.filter((el) => {
    if (!hideEmpty) return false;
    if (!isEmptyElement(el)) return false;
    if (!hasPosition(el)) return false;
    if (isBackgroundElement(el, slideW, slideH)) return false;
    const wPct = el.position.width / slideW;
    const hPct = el.position.height / slideH;
    return wPct < 0.5 && hPct < 0.5;
  });

  React.useEffect(() => {
    setImageError(false);
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
      requestAnimationFrame(() => {
        scrollToElement(elementId);
      });
    },
    [selectElement],
  );

  return (
    <div className={cn('flex flex-col gap-1.5', className)}>
      {/* ── Preview container ── */}
      <div
        className="relative w-full overflow-hidden rounded-lg border border-emerald-200/20 dark:border-emerald-700/15 bg-muted/15"
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
                linear-gradient(to right, hsl(var(--border) / 0.2) 1px, transparent 1px),
                linear-gradient(to bottom, hsl(var(--border) / 0.2) 1px, transparent 1px)
              `,
              backgroundSize: '10% 10%',
            }}
          />
        )}

        {/* Decorative (empty) element outlines */}
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
                border: '1px dashed rgba(16, 185, 129, 0.25)',
                backgroundColor: 'rgba(16, 185, 129, 0.03)',
                zIndex: 1,
              }}
              title={`${el.shapeName || 'Decorative'} (not editable)`}
            />
          );
        })}

        {/* Element overlays */}
        {visibleElements.map((el) => {
          const colors = getElementColors(el.type);
          const isSelected = selectedElementId === el.id;
          const isModified = isElementModified(el);

          const left = (el.position.x / slideW) * 100;
          const top = (el.position.y / slideH) * 100;
          const width = (el.position.width / slideW) * 100;
          const height = (el.position.height / slideH) * 100;

          return (
            <div
              key={el.id}
              className={cn(
                'group/overlay absolute cursor-pointer transition-all duration-150 ease-out',
                'rounded-[2px]',
                isSelected && 'animate-element-select-pulse',
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
                  ? `0 0 0 1.5px ${colors.glow}, 0 0 10px ${colors.glow}`
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
                  'rounded-[2px] px-0.5 py-[1px] text-[7px] font-semibold leading-none text-white opacity-0',
                  'transition-opacity duration-100',
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
                  className="pointer-events-none absolute -right-0.5 -top-0.5 flex h-2 w-2 items-center justify-center"
                >
                  <span
                    className="block h-2 w-2 rounded-full bg-orange-500 shadow-sm"
                    style={{
                      boxShadow: '0 0 0 1px rgba(255,255,255,0.9), 0 0 4px rgba(249,115,22,0.4)',
                    }}
                  />
                </span>
              )}
            </div>
          );
        })}

        {/* Hidden element overlays — dashed border, low opacity, no click */}
        {hiddenElements.map((el) => {
          const colors = getElementColors(el.type);

          const left = (el.position.x / slideW) * 100;
          const top = (el.position.y / slideH) * 100;
          const width = (el.position.width / slideW) * 100;
          const height = (el.position.height / slideH) * 100;

          if (width < 0.1 || height < 0.1) return null;

          return (
            <div
              key={`hidden-${el.id}`}
              className="pointer-events-none absolute rounded-[2px] opacity-20"
              style={{
                left: `${left}%`,
                top: `${top}%`,
                width: `${width}%`,
                height: `${height}%`,
                border: `1px dashed ${colors.border}`,
                backgroundColor: 'transparent',
                zIndex: 0,
              }}
            >
              {/* Hidden indicator — eye-off icon in corner */}
              <span
                className="pointer-events-none absolute -right-0.5 -top-0.5 flex items-center justify-center rounded-[2px] bg-amber-500/80 px-0.5 py-[1px]"
              >
                <EyeOff className="size-[7px] text-white" />
              </span>
            </div>
          );
        })}

        {/* Grid overlay */}
        {showGridOverlay && (
          <div
            className="absolute inset-0 pointer-events-none"
            style={{ zIndex: 5 }}
          >
            {/* Vertical lines every 10% */}
            {Array.from({ length: 9 }, (_, i) => i + 1).map((pct) => (
              <div
                key={`v-${pct}`}
                className="absolute top-0 bottom-0"
                style={{
                  left: `${pct * 10}%`,
                  width: 0,
                  borderLeft: pct === 5
                    ? '1px solid rgba(16, 185, 129, 0.3)'
                    : '1px solid rgba(16, 185, 129, 0.15)',
                }}
              />
            ))}
            {/* Horizontal lines every 10% */}
            {Array.from({ length: 9 }, (_, i) => i + 1).map((pct) => (
              <div
                key={`h-${pct}`}
                className="absolute left-0 right-0"
                style={{
                  top: `${pct * 10}%`,
                  height: 0,
                  borderTop: pct === 5
                    ? '1px solid rgba(16, 185, 129, 0.3)'
                    : '1px solid rgba(16, 185, 129, 0.15)',
                }}
              />
            ))}
            {/* Dot indicators at grid intersections */}
            {Array.from({ length: 9 }, (_, vi) => vi + 1).flatMap((vpct) =>
              Array.from({ length: 9 }, (_, hi) => hi + 1).map((hpct) => (
                <div
                  key={`d-${vpct}-${hpct}`}
                  className="absolute"
                  style={{
                    left: `${vpct * 10}%`,
                    top: `${hpct * 10}%`,
                    width: 1,
                    height: 1,
                    backgroundColor: 'rgba(16, 185, 129, 0.2)',
                    transform: 'translate(-0.5px, -0.5px)',
                  }}
                />
              ))
            )}
          </div>
        )}

        {/* Click on empty area deselects */}
        <div
          className="absolute inset-0"
          style={{ zIndex: 0 }}
          onClick={() => selectElement(null)}
        />
      </div>

      {/* ── Legend ── */}
      <div className="flex items-center justify-center gap-4 text-[9px] text-muted-foreground/70">
        <LegendItem color={ELEMENT_COLORS.text.dot} label="Text" />
        <LegendItem color={ELEMENT_COLORS.table.dot} label="Table" />
        <LegendItem color={ELEMENT_COLORS.image.dot} label="Image" />
        <LegendItem color="rgb(249, 115, 22)" label="Modified" dotStyle="ring" />
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
    <span className="inline-flex items-center gap-1">
      {dotStyle === 'solid' ? (
        <span
          className="block h-2 w-2 rounded-[1px]"
          style={{ backgroundColor: color }}
        />
      ) : (
        <span
          className="block h-2 w-2 rounded-full"
          style={{
            backgroundColor: color,
            boxShadow: '0 0 0 1px rgba(255,255,255,0.9)',
          }}
        />
      )}
      <span>{label}</span>
    </span>
  );
}
