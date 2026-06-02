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

const ELEMENT_COLORS = {
  text: { border: 'rgba(99, 102, 241, 0.85)', active: 'rgb(99, 102, 241)', activeBg: 'rgba(99, 102, 241, 0.18)', glow: 'rgba(99, 102, 241, 0.35)' },
  table: { border: 'rgba(16, 185, 129, 0.85)', active: 'rgb(16, 185, 129)', activeBg: 'rgba(16, 185, 129, 0.18)', glow: 'rgba(16, 185, 129, 0.35)' },
  image: { border: 'rgba(168, 85, 247, 0.85)', active: 'rgb(168, 85, 247)', activeBg: 'rgba(168, 85, 247, 0.18)', glow: 'rgba(168, 85, 247, 0.35)' },
};

export function isEmptyElement(el: PptxElement): boolean {
  if (el.type === 'text') return !el.originalText.trim();
  if (el.type === 'table') return el.rows.length === 0;
  return false;
}

function hasPosition(el: PptxElement): boolean {
  return el.position.width > 0 && el.position.height > 0;
}

/**
 * Determine if an element is a "background-style" element that should NOT
 * be rendered as an interactive highlight overlay.
 *
 * Two cases are treated as background:
 *  1. Image elements that cover ≥90% of the slide AND have no replacement
 *     — these are usually the slide's main background image (WPS stores
 *     full-bleed backgrounds as <p:pic> elements rather than slide masters).
 *     Rendering them as a highlight overlay would obscure the rest of the
 *     slide and click events on them are not meaningful.
 *  2. Empty text rectangles that cover ≥90% of the slide — these are
 *     "page background" rectangles with no content; they're decorative
 *     and should not show up as an editable element.
 */
function isBackgroundElement(el: PptxElement, slideW: number, slideH: number): boolean {
  if (el.position.width === 0 || el.position.height === 0) return false;
  const wPct = el.position.width / slideW;
  const hPct = el.position.height / slideH;
  const coversMostOfSlide = wPct >= 0.9 && hPct >= 0.9;

  if (!coversMostOfSlide) return false;

  if (el.type === 'image') {
    // Background images: full-bleed, no replacement
    const imgEl = el as PptxImageElement;
    if (!imgEl.replacementImageData) return true;
  }
  if (el.type === 'text') {
    // Empty full-bleed text rectangles (no actual content)
    if (isEmptyElement(el)) return true;
  }
  return false;
}

function getElementLabel(el: PptxElement): string {
  if (el.type === 'text') return el.shapeName || '文本框';
  if (el.type === 'table') return el.shapeName || '表格';
  return el.shapeName || '图片';
}

interface SlidePreviewProps { slide: PptxSlideData; }

export function SlidePreview({ slide }: SlidePreviewProps) {
  const { selectedElementId, selectElement, hideEmpty, slideSize } = usePptxStore();

  // Use the actual slide dimensions from the PPTX file (parsed from presentation.xml)
  const aspectRatio = slideSize.width / slideSize.height;

  // Use percentage-based positioning to avoid pixel scaling errors
  // This ensures overlay frames align perfectly with the preview image at any size
  const slideW = slideSize.width;
  const slideH = slideSize.height;

  const visibleElements = slide.elements.filter((el) => !(hideEmpty && isEmptyElement(el)));
  const elementsWithPosition = visibleElements.filter((el) => hasPosition(el) && !isBackgroundElement(el, slideW, slideH));
  const hasPreviewImage = !!slide.previewImage;

  const handleDoubleClick = (el: PptxElement) => {
    selectElement(el.id);
    setTimeout(() => scrollToElement(el.id), 100);
  };

  return (
    <div className="w-full">
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
                backgroundSize: `${100 / slideW * 914400}px ${100 / slideH * 914400}px`,
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

            // For image elements with replacement, show the replacement image as thumbnail
            const imgEl = el.type === 'image' ? (el as PptxImageElement) : null;
            const replacementThumbnail = imgEl?.replacementImageData
              ? (imgEl.replacementImageData.startsWith('data:')
                ? imgEl.replacementImageData
                : `data:image/${imgEl.replacementImageType || 'png'};base64,${imgEl.replacementImageData}`)
              : null;

            // Use percentage-based positioning for accurate alignment
            const leftPct = (el.position.x / slideW) * 100;
            const topPct = (el.position.y / slideH) * 100;
            const widthPct = (el.position.width / slideW) * 100;
            const heightPct = (el.position.height / slideH) * 100;

            // Skip very small elements
            if (widthPct < 0.1 || heightPct < 0.1) return null;

            return (
              <button key={el.id} className="absolute rounded-[2px] transition-all duration-150 cursor-pointer overflow-hidden"
                style={{
                  left: `${leftPct}%`, top: `${topPct}%`, width: `${widthPct}%`, height: `${heightPct}%`,
                  backgroundColor: isSelected ? colors.activeBg : 'transparent',
                  // Default: dashed colored border so the user can see ALL editable
                  // regions at a glance. Selected / modified get a solid stronger border.
                  border: isSelected
                    ? `2px solid ${colors.active}`
                    : isModified
                      ? '2px solid rgba(249, 115, 22, 0.9)'
                      : `1.5px dashed ${colors.border}`,
                  boxShadow: isSelected ? `0 0 0 3px ${colors.glow}, 0 0 12px ${colors.glow}` : 'none',
                  zIndex: isSelected ? 20 : 10,
                }}
                onClick={(e) => { e.stopPropagation(); selectElement(el.id); }}
                onDoubleClick={(e) => { e.stopPropagation(); handleDoubleClick(el); }}
                onMouseEnter={(e) => {
                  if (!isSelected) {
                    e.currentTarget.style.backgroundColor = colors.activeBg;
                    e.currentTarget.style.border = isModified
                      ? '2px solid rgba(249, 115, 22, 1)'
                      : `1.5px solid ${colors.active}`;
                  }
                }}
                onMouseLeave={(e) => {
                  if (!isSelected) {
                    e.currentTarget.style.backgroundColor = 'transparent';
                    e.currentTarget.style.border = isModified
                      ? '2px solid rgba(249, 115, 22, 0.9)'
                      : `1.5px dashed ${colors.border}`;
                  }
                }}
                title={`${getElementLabel(el)} (双击跳转编辑)`}
              >
                {/* Show replacement image as thumbnail overlay */}
                {replacementThumbnail && (
                  <img
                    src={replacementThumbnail}
                    alt="Replaced"
                    className="absolute inset-0 w-full h-full object-cover opacity-60 pointer-events-none"
                  />
                )}
                {isSelected && heightPct > 3 && widthPct > 5 && (
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
          <span className="inline-block w-1.5 h-1.5 rounded-sm" style={{ backgroundColor: 'rgba(99, 102, 241, 0.5)', border: '1px solid rgba(99, 102, 241, 0.85)' }} /> 文本
        </span>
        <span className="flex items-center gap-0.5">
          <span className="inline-block w-1.5 h-1.5 rounded-sm" style={{ backgroundColor: 'rgba(16, 185, 129, 0.5)', border: '1px solid rgba(16, 185, 129, 0.85)' }} /> 表格
        </span>
        <span className="flex items-center gap-0.5">
          <span className="inline-block w-1.5 h-1.5 rounded-sm" style={{ backgroundColor: 'rgba(168, 85, 247, 0.5)', border: '1px solid rgba(168, 85, 247, 0.85)' }} /> 图片
        </span>
        <span className="flex items-center gap-0.5">
          <span className="inline-block w-1.5 h-1.5 rounded-full" style={{ backgroundColor: 'rgb(249, 115, 22)', border: '1px solid white' }} /> 已修改
        </span>
      </div>
    </div>
  );
}
