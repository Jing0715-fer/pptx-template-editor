'use client';

import React, { useState, useEffect, useCallback, useRef } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import { X } from 'lucide-react';
import { Button } from '@/components/ui/button';
import { SlidePreview } from '@/components/slide-preview';
import { type PptxSlideData } from '@/lib/pptx-store';
import { cn } from '@/lib/utils';

// ============================================================================
// PresentationMode — Full-screen slideshow overlay
// ============================================================================

interface PresentationModeProps {
  slides: PptxSlideData[];
  initialSlideIndex: number;
  onExit: () => void;
  onSlideChange: (index: number) => void;
}

export default function PresentationMode({
  slides,
  initialSlideIndex,
  onExit,
  onSlideChange,
}: PresentationModeProps) {
  const [currentIndex, setCurrentIndex] = useState(initialSlideIndex);
  const [showControls, setShowControls] = useState(true);
  const [cursorHidden, setCursorHidden] = useState(false);
  const cursorTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const controlsTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);

  const totalSlides = slides.length;
  const currentSlide = slides[currentIndex];

  // ── Auto-hide controls after 3 seconds of inactivity ──
  const scheduleControlsHide = useCallback(() => {
    if (controlsTimerRef.current) clearTimeout(controlsTimerRef.current);
    controlsTimerRef.current = setTimeout(() => {
      setShowControls(false);
    }, 3000);
  }, []);

  // ── Auto-hide cursor after 3 seconds of inactivity ──
  const scheduleCursorHide = useCallback(() => {
    if (cursorTimerRef.current) clearTimeout(cursorTimerRef.current);
    cursorTimerRef.current = setTimeout(() => {
      setCursorHidden(true);
    }, 3000);
  }, []);

  // ── Start initial timers on mount ──
  useEffect(() => {
    scheduleControlsHide();
    scheduleCursorHide();
    return () => {
      if (cursorTimerRef.current) clearTimeout(cursorTimerRef.current);
      if (controlsTimerRef.current) clearTimeout(controlsTimerRef.current);
    };
  }, [scheduleControlsHide, scheduleCursorHide]);

  // ── Navigate slides ──
  const goToSlide = useCallback(
    (index: number) => {
      const clamped = Math.max(0, Math.min(totalSlides - 1, index));
      if (clamped !== currentIndex) {
        setCurrentIndex(clamped);
        onSlideChange(clamped);
      }
    },
    [currentIndex, totalSlides, onSlideChange],
  );

  const goNext = useCallback(() => goToSlide(currentIndex + 1), [currentIndex, goToSlide]);
  const goPrev = useCallback(() => goToSlide(currentIndex - 1), [currentIndex, goToSlide]);

  // ── Keyboard navigation ──
  useEffect(() => {
    const handler = (e: KeyboardEvent) => {
      switch (e.key) {
        case 'Escape':
          e.preventDefault();
          onExit();
          break;
        case 'ArrowRight':
        case 'PageDown':
        case ' ':
        case 'Enter':
          e.preventDefault();
          goNext();
          break;
        case 'ArrowLeft':
        case 'PageUp':
        case 'Backspace':
          e.preventDefault();
          goPrev();
          break;
        case 'Home':
          e.preventDefault();
          goToSlide(0);
          break;
        case 'End':
          e.preventDefault();
          goToSlide(totalSlides - 1);
          break;
      }
      // Show controls on any key press
      setShowControls(true);
      scheduleControlsHide();
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [onExit, goNext, goPrev, goToSlide, totalSlides, scheduleControlsHide]);

  // ── Mouse move handler ──
  const handleMouseMove = useCallback(() => {
    setShowControls(true);
    setCursorHidden(false);
    scheduleControlsHide();
    scheduleCursorHide();
  }, [scheduleControlsHide, scheduleCursorHide]);

  // ── Click-to-navigate on left/right halves ──
  const handleClick = useCallback(
    (e: React.MouseEvent) => {
      const target = e.target as HTMLElement;
      // Don't navigate if clicking the exit button or progress bar
      if (target.closest('[data-presentation-control]')) return;

      const rect = containerRef.current?.getBoundingClientRect();
      if (!rect) return;
      const x = e.clientX - rect.left;
      const midX = rect.width / 2;
      if (x < midX) {
        goPrev();
      } else {
        goNext();
      }
    },
    [goNext, goPrev],
  );

  // ── Progress percentage ──
  const progressPct = totalSlides > 1 ? ((currentIndex) / (totalSlides - 1)) * 100 : 100;

  return (
    <div
      ref={containerRef}
      className={cn(
        'fixed inset-0 z-50 bg-black flex items-center justify-center',
        cursorHidden && 'cursor-none',
      )}
      onClick={handleClick}
      onMouseMove={handleMouseMove}
      role="dialog"
      aria-label="Presentation mode"
      aria-modal="true"
    >
      {/* ── Slide display ── */}
      <div className="w-full h-full flex items-center justify-center p-4 sm:p-8">
        <AnimatePresence mode="wait">
          <motion.div
            key={currentIndex}
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.3, ease: 'easeInOut' }}
            className="w-full h-full flex items-center justify-center"
          >
            {currentSlide && (
              <SlidePreview
                slide={currentSlide}
                className="w-full max-w-[1200px]"
              />
            )}
          </motion.div>
        </AnimatePresence>
      </div>

      {/* ── Exit button (top-right, appears on hover/mouse move) ── */}
      <AnimatePresence>
        {showControls && (
          <motion.div
            data-presentation-control
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.2 }}
            className="absolute top-4 right-4 z-50"
          >
            <Button
              variant="ghost"
              size="icon"
              onClick={(e) => {
                e.stopPropagation();
                onExit();
              }}
              className={cn(
                'size-9 rounded-full',
                'bg-white/10 hover:bg-white/20 text-white/70 hover:text-white',
                'backdrop-blur-sm border border-white/10',
                'transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-400/50',
              )}
              aria-label="Exit presentation mode"
            >
              <X className="size-4" />
            </Button>
          </motion.div>
        )}
      </AnimatePresence>

      {/* ── Slide counter (bottom-right, semi-transparent) ── */}
      <AnimatePresence>
        {showControls && (
          <motion.div
            data-presentation-control
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            transition={{ duration: 0.2 }}
            className="absolute bottom-6 right-6 z-50"
          >
            <span className="text-white/50 text-xs font-medium tabular-nums bg-white/5 backdrop-blur-sm rounded px-2 py-1 border border-white/5">
              Slide {currentIndex + 1} / {totalSlides}
            </span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* ── Navigation hint arrows (left/right sides, appear on hover) ── */}
      <AnimatePresence>
        {showControls && (
          <>
            {/* Left side hint */}
            {currentIndex > 0 && (
              <motion.div
                data-presentation-control
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                transition={{ duration: 0.2 }}
                className="absolute left-3 top-1/2 -translate-y-1/2 z-40 pointer-events-none"
              >
                <div className="text-white/15">
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <polyline points="15 18 9 12 15 6" />
                  </svg>
                </div>
              </motion.div>
            )}
            {/* Right side hint */}
            {currentIndex < totalSlides - 1 && (
              <motion.div
                data-presentation-control
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                exit={{ opacity: 0 }}
                transition={{ duration: 0.2 }}
                className="absolute right-3 top-1/2 -translate-y-1/2 z-40 pointer-events-none"
              >
                <div className="text-white/15">
                  <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                    <polyline points="9 18 15 12 9 6" />
                  </svg>
                </div>
              </motion.div>
            )}
          </>
        )}
      </AnimatePresence>

      {/* ── Progress bar (bottom, thin emerald/teal gradient) ── */}
      <div className="absolute bottom-0 left-0 right-0 z-40">
        <div className="h-[2px] w-full bg-white/5">
          <motion.div
            className="h-full bg-gradient-to-r from-emerald-500 to-teal-500"
            initial={{ width: 0 }}
            animate={{ width: `${progressPct}%` }}
            transition={{ duration: 0.3, ease: 'easeOut' }}
          />
        </div>
      </div>
    </div>
  );
}
