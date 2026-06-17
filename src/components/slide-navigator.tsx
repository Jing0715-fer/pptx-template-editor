'use client';

import { useState, useCallback } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { Eye, EyeOff, FileText, ChevronLeft, ChevronRight, GripVertical } from 'lucide-react';
import {
  DndContext,
  closestCenter,
  KeyboardSensor,
  PointerSensor,
  useSensor,
  useSensors,
  type DragEndEvent,
} from '@dnd-kit/core';
import {
  SortableContext,
  sortableKeyboardCoordinates,
  useSortable,
  verticalListSortingStrategy,
  arrayMove,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { cn } from '@/lib/utils';
import {
  type PptxSlideData,
  usePptxStore,
} from '@/lib/pptx-store';
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

// Helper to get element type counts for a slide
function getElementTypeCounts(slide: PptxSlideData) {
  let text = 0;
  let table = 0;
  let image = 0;
  for (const el of slide.elements) {
    if (el.type === 'text') text++;
    else if (el.type === 'table') table++;
    else if (el.type === 'image') image++;
  }
  return { text, table, image, total: text + table + image };
}

// ── Sortable Slide Card (Expanded) ──
function SortableSlideCard({
  slide,
  index,
  isActive,
  modCount,
  previewErrors,
  onPreviewError,
  onClick,
}: {
  slide: PptxSlideData;
  index: number;
  isActive: boolean;
  modCount: number;
  previewErrors: Set<number>;
  onPreviewError: (slideNumber: number) => void;
  onClick: () => void;
}) {
  const {
    attributes,
    listeners,
    setNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id: `slide-${slide.slideNumber}` });

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
  };

  const typeCounts = getElementTypeCounts(slide);
  // Only show type bar if there are more than 1 element type
  const showTypeBar = typeCounts.total > 0 && [typeCounts.text > 0, typeCounts.table > 0, typeCounts.image > 0].filter(Boolean).length > 1;

  return (
    <motion.div
      ref={setNodeRef}
      style={style}
      layout
      initial={{ opacity: 0, y: 6 }}
      animate={{ opacity: isDragging ? 0.9 : 1, y: 0, scale: isDragging ? 1.03 : 1 }}
      exit={{ opacity: 0, y: -6 }}
      transition={{ duration: 0.15 }}
      onClick={onClick}
      className={cn(
        'group relative mb-1 cursor-pointer overflow-hidden rounded-md border transition-all duration-200',
        isDragging && 'shadow-lg z-10 opacity-90',
        isActive
          ? 'border-emerald-400/40 bg-emerald-50/40 dark:border-emerald-500/30 dark:bg-emerald-950/20 shadow-[0_0_8px_rgba(16,185,129,0.3)]'
          : 'border-transparent hover:border-border/40 hover:bg-accent/30'
      )}
    >
      {/* Left emerald bar indicator for active slide */}
      {isActive && (
        <div className="absolute left-0 top-0 bottom-0 w-[2px] bg-emerald-500 z-10" />
      )}

      {/* Faint left border indicator on hover for non-active */}
      {!isActive && (
        <div className="absolute left-0 top-0 bottom-0 w-[2px] bg-emerald-500/0 group-hover:bg-emerald-500/30 transition-all duration-200 z-10" />
      )}

      {/* Bottom gradient bar for active */}
      {isActive && (
        <motion.div
          layoutId="slide-accent-bar"
          className="absolute bottom-0 left-0 right-0 h-[2px] bg-gradient-to-r from-emerald-500 to-teal-500"
          transition={{ type: 'spring', stiffness: 400, damping: 30 }}
        />
      )}

      <div className="flex items-center gap-2 px-2 py-1.5">
        {/* Drag handle */}
        <div
          className="opacity-0 group-hover:opacity-100 transition-opacity duration-150 cursor-grab active:cursor-grabbing flex-shrink-0"
          {...attributes}
          {...listeners}
        >
          <GripVertical className="size-2.5 text-muted-foreground/50 hover:text-muted-foreground/80" />
        </div>

        {/* Thumbnail */}
        <div className={cn(
          'relative w-14 h-8 overflow-hidden rounded bg-muted/40 flex-shrink-0 transition-all duration-200',
          'hover:scale-[1.02]',
          'hover:shadow-[0_0_8px_rgba(16,185,129,0.15)]'
        )}>
          {slide.previewImage && !previewErrors.has(slide.slideNumber) ? (
            <img
              src={slide.previewImage}
              alt={`Slide ${slide.slideNumber}`}
              className="h-full w-full object-cover"
              draggable={false}
              onError={() => onPreviewError(slide.slideNumber)}
            />
          ) : (
            <div className="flex h-full w-full items-center justify-center bg-muted/20 dark:bg-muted/30">
              <FileText className="size-3 text-muted-foreground/30" />
            </div>
          )}
          <div
            className={cn(
              'absolute bottom-0 left-0.5 flex items-center justify-center rounded-sm px-0.5 text-[7px] font-semibold',
              isActive
                ? 'bg-emerald-500/90 text-white'
                : 'bg-black/40 text-white'
            )}
          >
            {slide.slideNumber}
          </div>
        </div>

        {/* Info */}
        <div className="flex-1 min-w-0">
          <span className={cn(
            'text-[10px] font-medium block truncate',
            isActive ? 'text-emerald-600 dark:text-emerald-400' : 'text-foreground'
          )}>
            Slide {slide.slideNumber}
          </span>
          <span className="text-[9px] text-muted-foreground block truncate">
            {typeCounts.total} el
          </span>
          {/* Element type breakdown bar */}
          {showTypeBar && (
            <div className="mt-0.5 w-full h-1 rounded-full overflow-hidden flex bg-muted/30">
              <div
                className="bg-emerald-500 dark:bg-emerald-400 transition-all duration-300"
                style={{ width: `${(typeCounts.text / typeCounts.total) * 100}%` }}
              />
              <div
                className="bg-amber-500 dark:bg-amber-400 transition-all duration-300"
                style={{ width: `${(typeCounts.table / typeCounts.total) * 100}%` }}
              />
              <div
                className="bg-cyan-500 dark:bg-cyan-400 transition-all duration-300"
                style={{ width: `${(typeCounts.image / typeCounts.total) * 100}%` }}
              />
            </div>
          )}
        </div>

        {/* Modification badge */}
        {modCount > 0 && (
          <Badge className="bg-amber-500 text-white border-0 text-[7px] px-0.5 py-0 h-3.5 shrink-0">
            {modCount}
          </Badge>
        )}
      </div>
    </motion.div>
  );
}

export default function SlideNavigator({ slides, currentSlide: _currentSlide }: SlideNavigatorProps) {
  const [collapsed, setCollapsed] = useState(false);
  const [previewErrors, setPreviewErrors] = useState<Set<number>>(new Set());
  const {
    currentSlideIndex,
    setCurrentSlide,
    hideEmpty,
    toggleHideEmpty,
    getTotalModificationCount,
    reorderSlides,
  } = usePptxStore();

  const totalModifications = getTotalModificationCount();

  // Compute total element count for progress bar
  const totalElements = slides.reduce((sum, slide) => sum + slide.elements.length, 0);

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

  const handleDragEnd = useCallback((event: DragEndEvent) => {
    const { active, over } = event;
    if (!over || active.id === over.id) return;

    // Extract indices from sortable IDs
    const activeId = String(active.id);
    const overId = String(over.id);
    const activeIndex = slides.findIndex((s) => `slide-${s.slideNumber}` === activeId);
    const overIndex = slides.findIndex((s) => `slide-${s.slideNumber}` === overId);

    if (activeIndex !== -1 && overIndex !== -1) {
      reorderSlides(activeIndex, overIndex);
    }
  }, [slides, reorderSlides]);

  const handlePreviewError = useCallback((slideNumber: number) => {
    setPreviewErrors(prev => new Set(prev).add(slideNumber));
  }, []);

  return (
    <motion.aside
      className="flex h-full flex-col overflow-hidden bg-background dark:bg-muted/10 border-r border-border/30"
      animate={{ width: collapsed ? 44 : 200 }}
      transition={{ type: 'spring', stiffness: 300, damping: 30 }}
    >
      {/* ── Header ── */}
      <div className="flex items-center gap-1.5 border-b border-border/30 px-2 h-[30px] flex-shrink-0">
        <AnimatePresence mode="wait">
          {!collapsed && (
            <motion.div
              key="header-content"
              initial={{ opacity: 0, x: -6 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: -6 }}
              transition={{ duration: 0.12 }}
              className="flex flex-1 items-center gap-1.5 overflow-hidden"
            >
              <FileText className="size-3.5 shrink-0 text-muted-foreground" />
              <span className="truncate text-[11px] font-semibold">Slides</span>
              <Badge variant="secondary" className="ml-auto shrink-0 text-[9px] px-1 py-0 h-3.5">
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
                className="size-6 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30"
                onClick={toggleHideEmpty}
                aria-label={hideEmpty ? 'Show empty elements' : 'Hide empty elements'}
              >
                {hideEmpty ? (
                  <EyeOff className="size-3 text-muted-foreground" />
                ) : (
                  <Eye className="size-3 text-muted-foreground" />
                )}
              </Button>
            </TooltipTrigger>
            <TooltipContent side="right" sideOffset={6}>
              {hideEmpty ? 'Show Empty' : 'Hide Empty'}
            </TooltipContent>
          </Tooltip>

          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="ghost"
                size="icon"
                className="size-6 focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30"
                onClick={() => setCollapsed((c) => !c)}
                aria-label={collapsed ? 'Expand slide navigator' : 'Collapse slide navigator'}
              >
                {collapsed ? (
                  <ChevronRight className="size-3 text-muted-foreground" />
                ) : (
                  <ChevronLeft className="size-3 text-muted-foreground" />
                )}
              </Button>
            </TooltipTrigger>
            <TooltipContent side="right" sideOffset={6}>
              {collapsed ? 'Expand' : 'Collapse'}
            </TooltipContent>
          </Tooltip>
        </div>
      </div>

      {/* ── Slide list ── */}
      <div className="custom-scrollbar flex-1 min-h-0 overflow-y-auto px-1.5 py-1.5">
        {slides.length === 0 && !collapsed && (
          <div className="flex flex-col items-center justify-center py-8 text-muted-foreground">
            <div className="flex size-8 items-center justify-center rounded-lg bg-emerald-500/10 mb-2">
              <FileText className="size-4 text-emerald-500/50" />
            </div>
            <span className="text-[10px] font-medium opacity-70 dark:opacity-80">No slides loaded</span>
          </div>
        )}

        {collapsed ? (
          // Collapsed: compact number strip (no drag)
          <AnimatePresence mode="popLayout">
            {slides.map((slide, index) => {
              const isActive = index === currentSlideIndex;
              const modCount = getSlideModificationCount(slide);

              return (
                <Tooltip key={slide.slideNumber}>
                  <TooltipTrigger asChild>
                    <motion.button
                      layout
                      whileHover={{ scale: 1.06 }}
                      whileTap={{ scale: 0.95 }}
                      onClick={() => setCurrentSlide(index)}
                      aria-label={`Slide ${slide.slideNumber}, ${slide.elements.length} elements`}
                      className={cn(
                        'relative mb-1 flex size-8 items-center justify-center rounded-md text-[10px] font-medium transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                        isActive
                          ? 'bg-emerald-500/15 text-emerald-600 dark:text-emerald-400 border border-emerald-500/30'
                          : 'text-muted-foreground hover:bg-accent/50 hover:text-accent-foreground border border-transparent'
                      )}
                    >
                      {slide.slideNumber}
                      {/* Dot indicator for modified slides */}
                      {modCount > 0 && (
                        <span className="absolute -bottom-0.5 left-1/2 -translate-x-1/2 size-1.5 rounded-full bg-amber-500" />
                      )}
                    </motion.button>
                  </TooltipTrigger>
                  <TooltipContent side="right" sideOffset={6}>
                    Slide {slide.slideNumber} · {slide.elements.length} element{slide.elements.length !== 1 ? 's' : ''}
                    {modCount > 0 && ` · ${modCount} edits`}
                  </TooltipContent>
                </Tooltip>
              );
            })}
          </AnimatePresence>
        ) : (
          // Expanded: draggable slide cards
          <DndContext
            sensors={sensors}
            collisionDetection={closestCenter}
            onDragEnd={handleDragEnd}
          >
            <SortableContext
              items={slides.map((s) => `slide-${s.slideNumber}`)}
              strategy={verticalListSortingStrategy}
            >
              <AnimatePresence mode="popLayout">
                {slides.map((slide, index) => {
                  const isActive = index === currentSlideIndex;
                  const modCount = getSlideModificationCount(slide);

                  return (
                    <SortableSlideCard
                      key={slide.slideNumber}
                      slide={slide}
                      index={index}
                      isActive={isActive}
                      modCount={modCount}
                      previewErrors={previewErrors}
                      onPreviewError={handlePreviewError}
                      onClick={() => setCurrentSlide(index)}
                    />
                  );
                })}
              </AnimatePresence>
            </SortableContext>
          </DndContext>
        )}
      </div>

      {/* ── Footer ── */}
      <AnimatePresence mode="wait">
        {!collapsed && (
          <motion.div
            key="footer"
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.1 }}
            className="border-t border-border/30 flex-shrink-0"
          >
            {/* Modification progress bar */}
            {totalModifications > 0 && totalElements > 0 && (
              <div className="mx-2 mt-1">
                <div className="w-full h-[1px] rounded-full bg-muted overflow-hidden">
                  <div
                    className="h-full bg-gradient-to-r from-amber-400 to-amber-500 transition-all duration-500"
                    style={{ width: `${Math.min((totalModifications / totalElements) * 100, 100)}%` }}
                  />
                </div>
              </div>
            )}
            <div className="px-2 py-1 flex items-center justify-between text-[9px] text-muted-foreground dark:text-muted-foreground/90">
              <span>{totalElements} elements</span>
              <span>
                {totalModifications > 0 ? (
                  <span className="font-medium text-amber-600 dark:text-amber-400">
                    {totalModifications} edits
                  </span>
                ) : (
                  'No edits'
                )}
              </span>
            </div>
          </motion.div>
        )}
      </AnimatePresence>
    </motion.aside>
  );
}
