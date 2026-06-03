'use client'

import React, { useEffect } from 'react';
import { usePptxStore } from '@/lib/pptx-store';
import { UploadZone } from '@/components/upload-zone';
import SlideNavigator from '@/components/slide-navigator';
import SlideEditor from '@/components/slide-editor';
import Toolbar from '@/components/toolbar';
import { Presentation, Loader2 } from 'lucide-react';
import { AiGenerateDialog } from '@/components/ai-generate-dialog';
import { AiSettingsDialog } from '@/components/ai-settings-dialog';
import { motion } from 'framer-motion';

export default function Home() {
  const {
    step,
    setStep,
    setParsedData,
    loadFromJson,
    slides,
    currentSlideIndex,
    setCurrentSlide,
    selectElement,
    getTotalModificationCount,
  } = usePptxStore();

  const [aiDialogOpen, setAiDialogOpen] = React.useState(false);
  const [aiSettingsOpen, setAiSettingsOpen] = React.useState(false);

  // Keyboard navigation
  useEffect(() => {
    if (step !== 'editing') return;

    const handleKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement;
      if (target.tagName === 'INPUT' || target.tagName === 'TEXTAREA' || target.isContentEditable) {
        if (e.key === 'Escape') {
          selectElement(null);
          target.blur();
        }
        return;
      }

      switch (e.key) {
        case 'ArrowLeft':
        case 'ArrowUp':
          e.preventDefault();
          setCurrentSlide(Math.max(0, currentSlideIndex - 1));
          break;
        case 'ArrowRight':
        case 'ArrowDown':
          e.preventDefault();
          setCurrentSlide(Math.min(slides.length - 1, currentSlideIndex + 1));
          break;
        case 'Escape':
          selectElement(null);
          break;
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [step, currentSlideIndex, slides.length, setCurrentSlide, selectElement]);

  const currentSlide = slides[currentSlideIndex];

  // Loading state
  if (step === 'loading') {
    return (
      <div className="h-screen flex items-center justify-center bg-hero-gradient bg-grid-pattern relative overflow-hidden">
        <div className="absolute top-1/4 left-1/4 w-64 h-64 bg-emerald-400/10 rounded-full blur-3xl animate-float" />
        <div className="absolute bottom-1/4 right-1/4 w-48 h-48 bg-amber-400/10 rounded-full blur-3xl animate-float-delayed" />

        <motion.div
          initial={{ opacity: 0, scale: 0.9 }}
          animate={{ opacity: 1, scale: 1 }}
          transition={{ duration: 0.5, ease: 'easeOut' }}
          className="flex flex-col items-center gap-6 z-10"
        >
          <div className="relative">
            <div className="w-20 h-20 rounded-2xl bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center shadow-xl shadow-emerald-500/20">
              <Loader2 className="w-10 h-10 text-white animate-spin" />
            </div>
            <div className="absolute -inset-2 rounded-3xl bg-gradient-to-br from-emerald-500/20 to-teal-600/20 blur-lg animate-pulse" />
          </div>
          <div className="text-center">
            <p className="text-lg font-semibold text-foreground">正在解析 PPTX 文件</p>
            <p className="text-sm text-muted-foreground mt-2">提取文字、表格和图片内容...</p>
          </div>
          <div className="flex items-center gap-1.5 mt-2">
            {[0, 1, 2].map((i) => (
              <motion.div
                key={i}
                className="w-2 h-2 rounded-full bg-emerald-500"
                animate={{ scale: [1, 1.5, 1], opacity: [0.5, 1, 0.5] }}
                transition={{ duration: 1, repeat: Infinity, delay: i * 0.2 }}
              />
            ))}
          </div>
        </motion.div>
      </div>
    );
  }

  // Upload state - landing page hero
  if (step === 'upload') {
    return (
      <div className="h-screen flex flex-col overflow-hidden max-w-full">
        <main className="flex-1 flex items-center justify-center min-h-0 overflow-hidden w-full">
          <UploadZone onAiGenerate={() => setAiDialogOpen(true)} onAiSettings={() => setAiSettingsOpen(true)} />
        </main>
        <footer className="border-t bg-gradient-to-r from-muted/20 via-background to-muted/20 py-3 px-4 text-center text-xs text-muted-foreground flex-shrink-0">
          <div className="flex items-center justify-center gap-1.5">
            <div className="w-4 h-4 rounded-md bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center">
              <Presentation className="w-2.5 h-2.5 text-white" />
            </div>
            PPTX 模板编辑器 · 上传 → 编辑 → 导出 · 保持原有排版格式
          </div>
        </footer>
        <AiGenerateDialog open={aiDialogOpen} onOpenChange={setAiDialogOpen} onOpenSettings={() => { setAiDialogOpen(false); setAiSettingsOpen(true); }} />
        <AiSettingsDialog open={aiSettingsOpen} onOpenChange={setAiSettingsOpen} />
      </div>
    );
  }

  // Editing state - TWO-COLUMN layout
  return (
    <div className="h-screen flex flex-col bg-background overflow-hidden max-w-full">
      <Toolbar onAiGenerate={() => setAiDialogOpen(true)} onAiSettings={() => setAiSettingsOpen(true)} />

      {/* Two-column content area */}
      <div className="flex-1 min-h-0 flex overflow-hidden">
        {/* Left column: Slide list + Preview */}
        <SlideNavigator slides={slides} currentSlide={currentSlide} />

        {/* Right column: Element editor — fills remaining space */}
        <div className="flex-1 min-w-0 h-full border-l border-border/40 bg-background">
          {currentSlide ? (
            <SlideEditor slide={currentSlide} />
          ) : (
            <div className="flex items-center justify-center h-full text-muted-foreground">
              <div className="text-center">
                <div className="mx-auto mb-4 flex size-14 items-center justify-center rounded-xl bg-gradient-to-br from-muted/50 to-muted/20 ring-1 ring-border/15">
                  <Presentation className="w-7 h-7 opacity-20" />
                </div>
                <p className="text-sm font-medium">选择一个幻灯片开始编辑</p>
                <p className="text-xs mt-1 text-muted-foreground/50">使用 ← → 键盘快捷键切换</p>
              </div>
            </div>
          )}
        </div>
      </div>

      {/* Status bar */}
      <div className="border-t border-border/40 bg-muted/15 px-4 py-1.5 text-[11px] text-muted-foreground/50 flex-shrink-0 flex items-center justify-between max-w-full overflow-hidden">
        <div className="flex items-center gap-3">
          <span>← → 切换幻灯片</span>
          <span className="text-border/40">·</span>
          <span>Esc 取消选中</span>
        </div>
        {slides.length > 0 && (
          <div className="flex items-center gap-2">
            <span>{slides.length} 页幻灯片</span>
            {getTotalModificationCount() > 0 && (
              <>
                <span className="text-border/40">·</span>
                <span className="text-amber-600/70 dark:text-amber-400/70">{getTotalModificationCount()} 处修改</span>
              </>
            )}
          </div>
        )}
      </div>

      <AiGenerateDialog open={aiDialogOpen} onOpenChange={setAiDialogOpen} onOpenSettings={() => { setAiDialogOpen(false); setAiSettingsOpen(true); }} />
      <AiSettingsDialog open={aiSettingsOpen} onOpenChange={setAiSettingsOpen} onConfigChange={() => {}} />
    </div>
  );
}
