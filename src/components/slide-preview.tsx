'use client'

import React from 'react';
import {
  type PptxElement,
  type PptxTextElement,
  type PptxTableElement,
  type PptxImageElement,
  type PptxSlideData,
  usePptxStore,
} from '@/lib/pptx-store';
import { scrollToElement } from '@/components/slide-editor';
import { cn } from '@/lib/utils';

const SLIDE_16_9 = { width: 12192000, height: 6858000 };
const SLIDE_4_3 = { width: 9144000, height: 6858000 };

const ELEMENT_COLORS = {
  text: { border: 'rgba(99, 102, 241, 0.5)', active: 'rgb(99, 102, 241)', activeBg: 'rgba(99, 102, 241, 0.15)', glow: 'rgba(99, 102, 241, 0.25)' },
  table: { border: 'rgba(16, 185, 129, 0.5)', active: 'rgb(16, 185, 129)', activeBg: 'rgba(16, 185, 129, 0.15)', glow: 'rgba(16, 185, 129, 0.25)' },
  image: { border: 'rgba(168, 85, 247, 0.5)', active: 'rgb(168, 85, 247)', activeBg: 'rgba(168, 85, 247, 0.15)', glow: 'rgba(168, 85, 247, 0.25)' },
};

export function isEmptyElement(el: PptxElement): boolean {
  if (el.type === 'text') return !el.originalText.trim();
  if (el.type === 'table') return el.rows.length === 0;
  return false;
}

function hasPosition(el: PptxElement): boolean {
  return el.position.width > 0 && el.position.height > 0;
}

function detectSlideSize(elements: PptxElement[]): { width: number; height: number } {
  let maxX = 0, maxY = 0;
  for (const el of elements) {
    if (hasPosition(el)) {
      maxX = Math.max(maxX, el.position.x + el.position.width);
      maxY = Math.max(maxY, el.position.y + el.position.height);
    }
  }
  if (maxX > 0 && maxY > 0) {
    const ratio = maxX / maxY;
    if (ratio > 1.5) return { width: Math.max(maxX + 457200, SLIDE_16_9.width), height: Math.max(maxY + 457200, SLIDE_16_9.height) };
    return { width: Math.max(maxX + 457200, SLIDE_4_3.width), height: Math.max(maxY + 457200, SLIDE_4_3.height) };
  }
  return SLIDE_16_9;
}

function getElementLabel(el: PptxElement): string {
  if (el.type === 'text') return el.shapeName || '文本框';
  if (el.type === 'table') return el.shapeName || '表格';
  return el.shapeName || '图片';
}

interface SlidePreviewProps { slide: PptxSlideData; }

export function SlidePreview({ slide }: SlidePreviewProps) {
  const { selectedElementId, selectElement, hideEmpty } = usePptxStore();
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [containerWidth, setContainerWidth] = React.useState(0);

  React.useEffect(() => {
    const container = containerRef.current;
    if (!container) return;
    const observer = new ResizeObserver((entries) => {
      for (const entry of entries) setContainerWidth(entry.contentRect.width);
    });
    observer.observe(container);
    return () => observer.disconnect();
  }, []);

  const slideSize = detectSlideSize(slide.elements);
  const aspectRatio = slideSize.width / slideSize.height;
  const previewWidth = containerWidth;
  const previewHeight = previewWidth / aspectRatio;
  const scale = previewWidth / slideSize.width;

  const visibleElements = slide.elements.filter((el) => !(hideEmpty && isEmptyElement(el)));
  const elementsWithPosition = visibleElements.filter(hasPosition);
  const hasPreviewImage = !!slide.previewImage;

  const handleDoubleClick = (el: PptxElement) => {
    selectElement(el.id);
    setTimeout(() => scrollToElement(el.id), 100);
  };

  return (
    <div ref={containerRef} className="w-full">
      <div className="relative w-full rounded-lg border border-border/50 overflow-hidden shadow-lg" style={{ paddingBottom: `${(1 / aspectRatio) * 100}%` }}>
        <div className="absolute inset-0">
          {hasPreviewImage && (
            <img src={slide.previewImage!} alt={`Slide ${slide.slideNumber} preview`} className="absolute inset-0 w-full h-full object-fill" style={{ zIndex: 0 }} />
          )}
          {!hasPreviewImage && (
            <>
              <div className="absolute inset-0 bg-white" />
              <div className="absolute inset-0 opacity-[0.04]" style={{
                backgroundImage: 'linear-gradient(rgba(99, 102, 241, 0.3) 1px, transparent 1px), linear-gradient(90deg, rgba(99, 102, 241, 0.3) 1px, transparent 1px)',
                backgroundSize: `${100 * scale}px ${100 * scale}px`,
              }} />
              <div className="absolute inset-0 flex items-center justify-center">
                <span className="text-[10px] text-muted-foreground/30 font-medium">第 {slide.slideNumber} 页 · 无预览截图</span>
              </div>
            </>
          )}

          {elementsWithPosition.map((el) => {
            const colors = ELEMENT_COLORS[el.type];
            const isSelected = selectedElementId === el.id;
            const isModified = (el.type === 'text' && el.currentText !== undefined && el.currentText !== el.originalText) ||
              (el.type === 'table' && el.currentRows !== undefined && el.currentRows.some((r, ri) => r.cells.some((c, ci) => c.text !== el.rows[ri]?.cells[ci]?.text))) ||
              (el.type === 'image' && !!(el as PptxImageElement).replacementImageData);

            const left = el.position.x * scale;
            const top = el.position.y * scale;
            const width = el.position.width * scale;
            const height = el.position.height * scale;
            if (width < 2 || height < 2) return null;

            return (
              <button key={el.id} className="absolute rounded-[2px] transition-all duration-150 cursor-pointer"
                style={{
                  left: `${left}px`, top: `${top}px`, width: `${width}px`, height: `${height}px`,
                  backgroundColor: isSelected ? colors.activeBg : 'transparent',
                  border: isSelected ? `2px solid ${colors.active}` : isModified ? '2px solid rgba(249, 115, 22, 0.7)' : '1.5px solid transparent',
                  boxShadow: isSelected ? `0 0 0 3px ${colors.glow}, 0 0 12px ${colors.glow}` : 'none',
                  zIndex: isSelected ? 20 : 10,
                }}
                onClick={(e) => { e.stopPropagation(); selectElement(el.id); }}
                onDoubleClick={(e) => { e.stopPropagation(); handleDoubleClick(el); }}
                onMouseEnter={(e) => { if (!isSelected) { e.currentTarget.style.backgroundColor = colors.activeBg; e.currentTarget.style.border = isModified ? '2px solid rgba(249, 115, 22, 0.8)' : `1.5px solid ${colors.border}`; } }}
                onMouseLeave={(e) => { if (!isSelected) { e.currentTarget.style.backgroundColor = 'transparent'; e.currentTarget.style.border = isModified ? '2px solid rgba(249, 115, 22, 0.7)' : '1.5px solid transparent'; } }}
                title={`${getElementLabel(el)} (双击跳转编辑)`}
              >
                {isSelected && height > 12 && width > 30 && (
                  <span className="absolute -top-4 left-0 text-[8px] font-bold px-1 py-px rounded-t whitespace-nowrap shadow-md"
                    style={{ backgroundColor: colors.active, color: 'white', fontSize: '8px', lineHeight: '12px' }}>
                    {getElementLabel(el)}
                  </span>
                )}
                {isModified && !isSelected && (
                  <span className="absolute -top-0.5 -right-0.5 w-2 h-2 rounded-full bg-orange-500 shadow-sm" style={{ border: '1px solid white' }} />
                )}
              </button>
            );
          })}
        </div>
      </div>

      <div className="flex items-center gap-2 mt-1.5 text-[8px] text-muted-foreground">
        <span className="flex items-center gap-0.5">
          <span className="inline-block w-1.5 h-1.5 rounded-sm" style={{ backgroundColor: 'rgba(99, 102, 241, 0.3)', border: '1px solid rgba(99, 102, 241, 0.5)' }} /> 文本
        </span>
        <span className="flex items-center gap-0.5">
          <span className="inline-block w-1.5 h-1.5 rounded-sm" style={{ backgroundColor: 'rgba(16, 185, 129, 0.3)', border: '1px solid rgba(16, 185, 129, 0.5)' }} /> 表格
        </span>
        <span className="flex items-center gap-0.5">
          <span className="inline-block w-1.5 h-1.5 rounded-sm" style={{ backgroundColor: 'rgba(168, 85, 247, 0.3)', border: '1px solid rgba(168, 85, 247, 0.5)' }} /> 图片
        </span>
        <span className="flex items-center gap-0.5">
          <span className="inline-block w-1.5 h-1.5 rounded-full" style={{ backgroundColor: 'rgb(249, 115, 22)', border: '1px solid white' }} /> 已修改
        </span>
      </div>
    </div>
  );
}
