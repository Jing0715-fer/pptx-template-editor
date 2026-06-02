'use client'

import React, { useCallback, useEffect } from 'react';
import { usePptxStore, type PptxJsonData } from '@/lib/pptx-store';
import { UploadZone } from '@/components/upload-zone';
import { SlideNavigator } from '@/components/slide-navigator';
import { SlideEditor } from '@/components/slide-editor';
import { Toolbar } from '@/components/toolbar';
import { SlidePreview } from '@/components/slide-preview';
import { Sheet, SheetContent, SheetHeader, SheetTitle } from '@/components/ui/sheet';
import { Presentation, Type, Table as TableIcon, Image as ImageIcon, Loader2 } from 'lucide-react';
import { toast } from 'sonner';
import { addFileHistory } from '@/lib/pptx-history';
import { AiGenerateDialog } from '@/components/ai-generate-dialog';

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
  } = usePptxStore();

  const [mobileNavOpen, setMobileNavOpen] = React.useState(false);
  const [aiDialogOpen, setAiDialogOpen] = React.useState(false);

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
          setMobileNavOpen(false);
          break;
        case 'ArrowRight':
        case 'ArrowDown':
          e.preventDefault();
          setCurrentSlide(Math.min(slides.length - 1, currentSlideIndex + 1));
          setMobileNavOpen(false);
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
      <div className="h-screen flex items-center justify-center bg-gradient-to-br from-background via-background to-muted/30">
        <div className="flex flex-col items-center gap-4">
          <div className="w-16 h-16 rounded-2xl bg-primary/10 flex items-center justify-center">
            <Loader2 className="w-8 h-8 text-primary animate-spin" />
          </div>
          <div className="text-center">
            <p className="text-sm font-medium">正在解析 PPTX 文件</p>
            <p className="text-xs text-muted-foreground mt-1">提取文字、表格和图片内容...</p>
          </div>
        </div>
      </div>
    );
  }

  // Upload state
  if (step === 'upload') {
    return (
      <div className="min-h-screen flex flex-col">
        <main className="flex-1 flex items-center justify-center">
          <UploadZone onAiGenerate={() => setAiDialogOpen(true)} />
        </main>
        <footer className="border-t bg-gradient-to-r from-muted/20 via-background to-muted/20 py-3 px-4 text-center text-xs text-muted-foreground flex-shrink-0">
          <div className="flex items-center justify-center gap-1.5">
            <div className="w-4 h-4 rounded-md bg-primary/20 flex items-center justify-center">
              <Presentation className="w-2.5 h-2.5 text-primary/70" />
            </div>
            PPTX 模板编辑器 · 上传 → 编辑 → 导出 · 保持原有排版格式
          </div>
        </footer>
        <AiGenerateDialog open={aiDialogOpen} onOpenChange={setAiDialogOpen} />
      </div>
    );
  }

  // Editing state
  return (
    <div className="h-screen flex flex-col bg-background overflow-hidden">
      <Toolbar onAiGenerate={() => setAiDialogOpen(true)} />

      <div className="flex-1 flex overflow-hidden">
        {/* Desktop slide navigator */}
        <div className="hidden md:flex md:flex-col md:h-full min-h-0">
          <SlideNavigator slides={slides} />
        </div>

        {/* Mobile slide navigator (Sheet) */}
        <Sheet open={mobileNavOpen} onOpenChange={setMobileNavOpen}>
          <SheetContent side="left" className="w-72 p-0">
            <SheetHeader className="p-3 border-b bg-gradient-to-r from-primary/5 to-transparent">
              <SheetTitle className="flex items-center gap-2 text-sm">
                <div className="w-6 h-6 rounded-lg bg-gradient-to-br from-primary to-primary/80 flex items-center justify-center">
                  <Presentation className="w-3 h-3 text-primary-foreground" />
                </div>
                幻灯片 ({slides.length})
              </SheetTitle>
            </SheetHeader>
            <div className="overflow-y-auto p-2 space-y-1.5">
              {slides.map((slide, index) => {
                const textCount = slide.elements.filter((e) => e.type === 'text').length;
                const tableCount = slide.elements.filter((e) => e.type === 'table').length;
                const imageCount = slide.elements.filter((e) => e.type === 'image').length;
                const isActive = index === currentSlideIndex;
                const modificationCount = slide.elements.filter((el) => {
                  if (el.type === 'text') return el.currentText !== undefined && el.currentText !== el.originalText;
                  if (el.type === 'table' && el.currentRows) {
                    return el.currentRows.some((r, ri) => r.cells.some((c, ci) => c.text !== el.rows[ri]?.cells[ci]?.text));
                  }
                  if (el.type === 'image' && el.replacementImageData) return true;
                  return false;
                }).length;
                const hasChanges = modificationCount > 0;

                return (
                  <button
                    key={index}
                    onClick={() => { setCurrentSlide(index); setMobileNavOpen(false); }}
                    className={`w-full text-left rounded-xl p-2.5 transition-all duration-200 ease-out group relative ${
                      isActive
                        ? 'bg-gradient-to-r from-primary/10 to-primary/5 border border-primary/25 shadow-md shadow-primary/10'
                        : hasChanges
                        ? 'border border-orange-200/60 bg-gradient-to-r from-orange-50/40 to-amber-50/20 hover:from-orange-50/60 hover:to-amber-50/40 shadow-sm'
                        : 'border border-transparent hover:bg-primary/3 hover:border-primary/10'
                    }`}
                  >
                    <div className="flex items-center gap-2">
                      <span className={`inline-flex items-center justify-center w-6 h-6 rounded-lg text-[11px] font-bold shadow-sm transition-all duration-200 ${
                        isActive ? 'bg-gradient-to-br from-primary to-primary/80 text-primary-foreground scale-105' : hasChanges ? 'bg-gradient-to-br from-orange-400 to-amber-500 text-white' : 'bg-muted/60 text-muted-foreground group-hover:bg-muted/80'
                      }`}>
                        {index + 1}
                      </span>
                      <span className={`text-sm font-semibold transition-colors duration-200 ${isActive ? 'text-primary' : ''}`}>第 {index + 1} 页</span>
                      {hasChanges && (
                        <span className="ml-auto flex items-center gap-0.5 text-[10px] text-orange-600 font-bold">
                          <span className="w-1.5 h-1.5 rounded-full bg-orange-500 animate-pulse" />
                          {modificationCount}
                        </span>
                      )}
                    </div>
                    <div className="flex items-center gap-2 text-[10px] text-muted-foreground ml-8 mt-0.5">
                      {textCount > 0 && <span className="flex items-center gap-0.5"><Type className="w-2.5 h-2.5 text-blue-500/60" />{textCount}</span>}
                      {tableCount > 0 && <span className="flex items-center gap-0.5"><TableIcon className="w-2.5 h-2.5 text-emerald-500/60" />{tableCount}</span>}
                      {imageCount > 0 && <span className="flex items-center gap-0.5"><ImageIcon className="w-2.5 h-2.5 text-purple-500/60" />{imageCount}</span>}
                    </div>
                  </button>
                );
              })}
            </div>
          </SheetContent>
        </Sheet>

        {/* Main content area */}
        <div className="flex-1 flex min-h-0">
          {/* Preview panel (center) */}
          <div className="flex-1 overflow-y-auto p-4 sm:p-6 bg-gradient-to-b from-muted/20 to-muted/5">
            <div className="max-w-3xl mx-auto">
              {currentSlide ? (
                <SlidePreview slide={currentSlide} />
              ) : (
                <div className="flex items-center justify-center h-full text-muted-foreground">
                  <div className="text-center">
                    <Presentation className="w-12 h-12 mx-auto mb-3 opacity-20" />
                    <p className="text-sm">选择一个幻灯片开始编辑</p>
                  </div>
                </div>
              )}
            </div>
          </div>

          {/* Elements list (right side) */}
          <div className="hidden lg:block w-[420px] border-l overflow-y-auto">
            {currentSlide ? (
              <SlideEditor slide={currentSlide} />
            ) : (
              <div className="flex items-center justify-center h-full text-muted-foreground">
                <p className="text-sm">选择幻灯片查看编辑器</p>
              </div>
            )}
          </div>
        </div>

        {/* Mobile editor (shown below preview on small screens) */}
        <div className="lg:hidden border-t overflow-hidden max-h-[50vh]">
          {currentSlide && <SlideEditor slide={currentSlide} />}
        </div>
      </div>

      {/* Footer */}
      <footer className="border-t bg-gradient-to-r from-muted/20 via-background to-muted/20 py-2 px-4 text-center text-xs text-muted-foreground flex-shrink-0">
        <div className="flex items-center justify-center gap-1.5">
          <div className="w-4 h-4 rounded-md bg-primary/20 flex items-center justify-center">
            <Presentation className="w-2.5 h-2.5 text-primary/70" />
          </div>
          PPTX 模板编辑器 · ← → 切换幻灯片 · Esc 取消选中
        </div>
      </footer>

      <AiGenerateDialog open={aiDialogOpen} onOpenChange={setAiDialogOpen} />
    </div>
  );
}
