'use client'

import React from 'react';
import { type PptxSlideData, usePptxStore } from '@/lib/pptx-store';
import { cn } from '@/lib/utils';
import { ScrollArea } from '@/components/ui/scroll-area';
import { Badge } from '@/components/ui/badge';
import { Eye, EyeOff, FileText, ChevronLeft, ChevronRight } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Tooltip, TooltipContent, TooltipProvider, TooltipTrigger } from '@/components/ui/tooltip';
import { motion } from 'framer-motion';

interface SlideNavigatorProps {
  slides: PptxSlideData[];
}

export function SlideNavigator({ slides }: SlideNavigatorProps) {
  const { currentSlideIndex, setCurrentSlide, hideEmpty, toggleHideEmpty, selectedElementId, selectElement } = usePptxStore();
  const [isCollapsed, setIsCollapsed] = React.useState(false);

  const getSlideModificationCount = (slide: PptxSlideData): number => {
    let count = 0;
    for (const el of slide.elements) {
      if (el.type === 'text' && el.currentText !== undefined && el.currentText !== el.originalText) count++;
      else if (el.type === 'table' && el.currentRows) {
        for (let ri = 0; ri < el.currentRows.length; ri++) {
          const origRow = el.rows[ri];
          const curRow = el.currentRows[ri];
          if (!origRow || !curRow) continue;
          for (let ci = 0; ci < curRow.cells.length; ci++) {
            if (curRow.cells[ci]?.text !== origRow.cells[ci]?.text) count++;
          }
        }
      } else if (el.type === 'image' && (el as { replacementImageData?: unknown }).replacementImageData) count++;
    }
    return count;
  };

  const handleSlideClick = (index: number) => {
    setCurrentSlide(index);
    selectElement(null);
  };

  if (isCollapsed) {
    return (
      <div className="flex flex-col items-center w-12 h-full min-h-0 border-r bg-muted/30 py-2 gap-2">
        <TooltipProvider>
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="ghost"
                size="sm"
                className="h-8 w-8 p-0"
                onClick={() => setIsCollapsed(false)}
              >
                <ChevronRight className="w-4 h-4" />
              </Button>
            </TooltipTrigger>
            <TooltipContent side="right">展开导航</TooltipContent>
          </Tooltip>
        </TooltipProvider>
        <div className="w-6 h-px bg-border" />
        {slides.map((slide, index) => (
          <TooltipProvider key={slide.slideNumber}>
            <Tooltip>
              <TooltipTrigger asChild>
                <button
                  className={cn(
                    'w-8 h-6 rounded text-[10px] font-medium transition-all',
                    index === currentSlideIndex
                      ? 'bg-primary text-primary-foreground'
                      : 'bg-muted hover:bg-muted/80 text-muted-foreground',
                  )}
                  onClick={() => handleSlideClick(index)}
                >
                  {index + 1}
                </button>
              </TooltipTrigger>
              <TooltipContent side="right">第 {slide.slideNumber} 页</TooltipContent>
            </Tooltip>
          </TooltipProvider>
        ))}
      </div>
    );
  }

  return (
    <div className="flex flex-col w-64 h-full min-h-0 border-r bg-muted/30">
      {/* Header */}
      <div className="flex items-center justify-between px-3 py-2 border-b">
        <div className="flex items-center gap-1.5">
          <FileText className="w-3.5 h-3.5 text-muted-foreground" />
          <span className="text-xs font-medium">幻灯片</span>
          <Badge variant="secondary" className="text-[9px] px-1 py-0 h-4">
            {slides.length}
          </Badge>
        </div>
        <div className="flex items-center gap-0.5">
          <TooltipProvider>
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="sm"
                  className={cn('h-6 w-6 p-0', hideEmpty && 'text-primary')}
                  onClick={toggleHideEmpty}
                >
                  {hideEmpty ? <EyeOff className="w-3 h-3" /> : <Eye className="w-3 h-3" />}
                </Button>
              </TooltipTrigger>
              <TooltipContent side="right">
                {hideEmpty ? '显示空元素' : '隐藏空元素'}
              </TooltipContent>
            </Tooltip>
          </TooltipProvider>
          <TooltipProvider>
            <Tooltip>
              <TooltipTrigger asChild>
                <Button
                  variant="ghost"
                  size="sm"
                  className="h-6 w-6 p-0"
                  onClick={() => setIsCollapsed(true)}
                >
                  <ChevronLeft className="w-3 h-3" />
                </Button>
              </TooltipTrigger>
              <TooltipContent side="right">折叠导航</TooltipContent>
            </Tooltip>
          </TooltipProvider>
        </div>
      </div>

      {/* Slide list */}
      <div className="flex-1 min-h-0 overflow-y-auto">
        <div className="p-2 space-y-2">
          {slides.map((slide, index) => {
            const isActive = index === currentSlideIndex;
            const modCount = getSlideModificationCount(slide);
            const hasPreview = !!slide.previewImage;
            const elementCount = slide.elements.length;

            return (
              <motion.button
                key={slide.slideNumber}
                className={cn(
                  'w-full rounded-lg border-2 overflow-hidden transition-all text-left',
                  'hover:shadow-md focus:outline-none focus:ring-2 focus:ring-primary/30',
                  isActive
                    ? 'border-primary shadow-md ring-1 ring-primary/20'
                    : 'border-transparent hover:border-border',
                )}
                onClick={() => handleSlideClick(index)}
                whileHover={{ scale: 1.02 }}
                whileTap={{ scale: 0.98 }}
                layout
              >
                {/* Thumbnail area */}
                <div className="relative aspect-video bg-white">
                  {hasPreview ? (
                    <img
                      src={slide.previewImage!}
                      alt={`Slide ${slide.slideNumber}`}
                      className="w-full h-full object-fill"
                    />
                  ) : (
                    <div className="w-full h-full flex items-center justify-center bg-gradient-to-br from-muted/50 to-muted">
                      <div className="text-center">
                        <FileText className="w-5 h-5 text-muted-foreground/30 mx-auto mb-1" />
                        <span className="text-[8px] text-muted-foreground/40">无预览</span>
                      </div>
                    </div>
                  )}

                  {/* Slide number overlay */}
                  <div className={cn(
                    'absolute top-1 left-1 text-[9px] font-bold px-1.5 py-0.5 rounded',
                    isActive
                      ? 'bg-primary text-primary-foreground'
                      : 'bg-black/60 text-white',
                  )}>
                    {index + 1}
                  </div>

                  {/* Modification badge */}
                  {modCount > 0 && (
                    <div className="absolute top-1 right-1">
                      <Badge className="text-[8px] px-1 py-0 h-3.5 bg-orange-500 text-white border-0">
                        {modCount}
                      </Badge>
                    </div>
                  )}

                  {/* Active indicator */}
                  {isActive && (
                    <div className="absolute bottom-0 left-0 right-0 h-0.5 bg-primary" />
                  )}
                </div>

                {/* Slide info */}
                <div className="px-2 py-1.5 bg-card">
                  <div className="flex items-center justify-between">
                    <span className="text-[10px] font-medium truncate">
                      第 {slide.slideNumber} 页
                    </span>
                    <span className="text-[9px] text-muted-foreground">
                      {elementCount} 个元素
                    </span>
                  </div>
                  {modCount > 0 && (
                    <span className="text-[9px] text-orange-600 dark:text-orange-400">
                      {modCount} 处修改
                    </span>
                  )}
                </div>
              </motion.button>
            );
          })}
        </div>
      </div>

      {/* Footer stats */}
      <div className="px-3 py-2 border-t text-[9px] text-muted-foreground">
        <div className="flex items-center justify-between">
          <span>{slides.reduce((acc, s) => acc + s.elements.length, 0)} 个元素</span>
          <span>{slides.reduce((acc, s) => acc + getSlideModificationCount(s), 0)} 处修改</span>
        </div>
      </div>
    </div>
  );
}
