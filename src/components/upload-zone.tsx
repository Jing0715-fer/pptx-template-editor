'use client';

import { useState, useCallback, useRef, useEffect } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
import {
  Upload,
  FileUp,
  Sparkles,
  Code2,
  History,
  Trash2,
  RotateCcw,
  X,
  FileText,
  Loader2,
  ChevronRight,
  Clock,
  Layers,
  FileSpreadsheet,
  Play,
  Settings2,
} from 'lucide-react';
import { usePptxStore, type PptxJsonData } from '@/lib/pptx-store';
import {
  getFileHistory,
  addFileHistory,
  removeFileHistory,
  clearFileHistory,
  formatTimeAgo,
  type FileHistoryEntry,
} from '@/lib/pptx-history';
// Thumbnail API endpoint is used directly as <img src> — no need for client-side resize
import { toast } from 'sonner';
import { cn } from '@/lib/utils';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Input } from '@/components/ui/input';
import { Textarea } from '@/components/ui/textarea';
import { Separator } from '@/components/ui/separator';

// ============================================================================
// Types
// ============================================================================

interface UploadZoneProps {
  onAiGenerate: () => void;
  onAiSettings?: () => void;
}

// ============================================================================
// Animation variants
// ============================================================================

const containerVariants = {
  hidden: { opacity: 0 },
  visible: {
    opacity: 1,
    transition: { staggerChildren: 0.08, delayChildren: 0.1 },
  },
};

const itemVariants = {
  hidden: { opacity: 0, y: 20 },
  visible: {
    opacity: 1,
    y: 0,
    transition: { duration: 0.5, ease: [0.25, 0.46, 0.45, 0.94] },
  },
};

const floatingOrbVariants = {
  animate: {
    y: [0, -15, 5, 0],
    scale: [1, 1.05, 0.98, 1],
    transition: { duration: 8, repeat: Infinity, ease: 'easeInOut' },
  },
};

const historyCardVariants = {
  hidden: { opacity: 0, y: 20, scale: 0.95 },
  visible: (i: number) => ({
    opacity: 1,
    y: 0,
    scale: 1,
    transition: { duration: 0.4, delay: i * 0.06, ease: [0.25, 0.46, 0.45, 0.94] },
  }),
};

// ============================================================================
// HistoryThumbnail — loads thumbnail from API endpoint with error handling
// ============================================================================

function HistoryThumbnail({ fileId, fileName }: { fileId: string; fileName: string }) {
  const [status, setStatus] = useState<'loading' | 'loaded' | 'error'>('loading');
  const thumbnailUrl = `/api/pptx/thumbnail?fileId=${fileId}`;

  return (
    <>
      {status !== 'loaded' && (
        <div className="absolute inset-0 flex flex-col items-center justify-center bg-amber-500/5 z-10">
          {status === 'loading' ? (
            <Loader2 className="size-5 text-amber-400/50 animate-spin" />
          ) : (
            <>
              <FileSpreadsheet className="size-7 text-amber-400/40" />
              <span className="text-[9px] text-muted-foreground/40 mt-0.5">暂无预览</span>
            </>
          )}
        </div>
      )}
      <img
        src={thumbnailUrl}
        alt={`${fileName} 预览`}
        className="w-full h-full object-cover transition-transform duration-500 group-hover:scale-105"
        onLoad={() => setStatus('loaded')}
        onError={() => setStatus('error')}
        loading="lazy"
      />
    </>
  );
}

// ============================================================================
// Component
// ============================================================================

export function UploadZone({ onAiGenerate, onAiSettings }: UploadZoneProps) {
  // ---- State ----
  const [isDragging, setIsDragging] = useState(false);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [jsonFileId, setJsonFileId] = useState('');
  const [jsonText, setJsonText] = useState('');
  const [fileHistory, setFileHistory] = useState<FileHistoryEntry[]>([]);
  const [isJsonExpanded, setIsJsonExpanded] = useState(false);
  const [reopeningFileId, setReopeningFileId] = useState<string | null>(null);

  // ---- Store ----
  const { setStep, setParsedData, loadFromJson } = usePptxStore();

  // ---- Refs ----
  const fileInputRef = useRef<HTMLInputElement>(null);

  // ---- Load file history on mount ----
  useEffect(() => {
    setFileHistory(getFileHistory());
  }, []);

  // ========================================================================
  // File Upload
  // ========================================================================

  const handleFileUpload = useCallback(
    async (file: File) => {
      if (!file.name.toLowerCase().endsWith('.pptx')) {
        toast.error('仅支持 .pptx 格式文件');
        return;
      }
      if (file.size > 100 * 1024 * 1024) {
        toast.error('文件大小不能超过 100MB');
        return;
      }
      if (file.size === 0) {
        toast.error('文件内容为空');
        return;
      }

      setIsUploading(true);
      setUploadProgress(0);
      setStep('loading');

      try {
        const formData = new FormData();
        formData.append('file', file);

        const xhr = new XMLHttpRequest();

        const uploadPromise = new Promise<{ fileId: string; fileName: string; slideCount: number; slides: any[]; slideSize?: any }>(
          (resolve, reject) => {
            xhr.upload.addEventListener('progress', (e) => {
              if (e.lengthComputable) {
                setUploadProgress(Math.round((e.loaded / e.total) * 100));
              }
            });

            xhr.addEventListener('load', () => {
              if (xhr.status >= 200 && xhr.status < 300) {
                try {
                  const data = JSON.parse(xhr.responseText);
                  if (data.error) {
                    reject(new Error(data.error));
                  } else {
                    resolve(data);
                  }
                } catch {
                  reject(new Error('服务器响应解析失败'));
                }
              } else {
                try {
                  const errData = JSON.parse(xhr.responseText);
                  reject(new Error(errData.error || '上传失败'));
                } catch {
                  reject(new Error('上传失败'));
                }
              }
            });

            xhr.addEventListener('error', () => reject(new Error('网络错误，请检查连接')));
            xhr.addEventListener('abort', () => reject(new Error('上传已取消')));

            xhr.open('POST', '/api/pptx/parse');
            xhr.send(formData);
          }
        );

        const result = await uploadPromise;

        setParsedData(result.fileId, result.fileName, result.slides, result.slideSize);

        addFileHistory({
          fileId: result.fileId,
          fileName: result.fileName,
          slideCount: result.slideCount,
          openedAt: Date.now(),
        });
        setFileHistory(getFileHistory());

        toast.success(`已加载 "${result.fileName}"（${result.slideCount} 页幻灯片）`);
      } catch (error) {
        setStep('upload');
        toast.error(error instanceof Error ? error.message : '文件上传失败');
      } finally {
        setIsUploading(false);
        setUploadProgress(0);
      }
    },
    [setStep, setParsedData]
  );

  // ========================================================================
  // Drag & Drop
  // ========================================================================

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
      setIsDragging(false);

      const files = e.dataTransfer.files;
      if (files.length > 0) {
        handleFileUpload(files[0]);
      }
    },
    [handleFileUpload]
  );

  const handleFileSelect = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const files = e.target.files;
      if (files && files.length > 0) {
        handleFileUpload(files[0]);
      }
      // Reset input so same file can be selected again
      if (fileInputRef.current) {
        fileInputRef.current.value = '';
      }
    },
    [handleFileUpload]
  );

  // ========================================================================
  // JSON Import
  // ========================================================================

  const handleJsonImport = useCallback(() => {
    if (!jsonFileId.trim()) {
      toast.error('请输入文件 ID');
      return;
    }
    if (!jsonText.trim()) {
      toast.error('请输入 JSON 数据');
      return;
    }

    try {
      const jsonData: PptxJsonData = JSON.parse(jsonText);
      if (!jsonData.slides || !Array.isArray(jsonData.slides)) {
        toast.error('JSON 格式无效：缺少 slides 数组');
        return;
      }

      loadFromJson(jsonData);
      toast.success('JSON 数据已导入');
      setJsonFileId('');
      setJsonText('');
    } catch {
      toast.error('JSON 解析失败，请检查格式');
    }
  }, [jsonFileId, jsonText, loadFromJson]);

  // ========================================================================
  // File History
  // ========================================================================

  const handleReopenFile = useCallback(
    async (entry: FileHistoryEntry) => {
      try {
        setReopeningFileId(entry.fileId);

        // Check if the file still exists on the server
        const checkRes = await fetch(`/api/pptx/check?fileId=${entry.fileId}`);
        const checkData = await checkRes.json();

        if (!checkData.exists) {
          toast.error('文件已过期或不存在，请重新上传');
          removeFileHistory(entry.fileId);
          setFileHistory(getFileHistory());
          return;
        }

        setStep('loading');

        // Reparse the file
        const reparseRes = await fetch('/api/pptx/reparse', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ fileId: entry.fileId }),
        });
        const reparseData = await reparseRes.json();

        if (reparseData.error) {
          setStep('upload');
          toast.error(reparseData.error);
          return;
        }

        setParsedData(reparseData.fileId, reparseData.fileName, reparseData.slides, reparseData.slideSize);

        addFileHistory({
          fileId: reparseData.fileId,
          fileName: reparseData.fileName,
          slideCount: reparseData.slideCount,
          openedAt: Date.now(),
        });
        setFileHistory(getFileHistory());
        toast.success(`已重新打开 "${reparseData.fileName}"`);
      } catch {
        setStep('upload');
        toast.error('重新打开文件失败');
      } finally {
        setReopeningFileId(null);
      }
    },
    [setStep, setParsedData]
  );

  const handleDeleteHistory = useCallback((fileId: string) => {
    removeFileHistory(fileId);
    setFileHistory(getFileHistory());
    toast.success('已从历史记录中移除');
  }, []);

  const handleClearHistory = useCallback(() => {
    clearFileHistory();
    setFileHistory([]);
    toast.success('已清空历史记录');
  }, []);

  // ========================================================================
  // Render
  // ========================================================================

  return (
    <div className="h-full w-full max-w-full flex flex-col bg-hero-gradient relative overflow-hidden">
      {/* Main content */}
      <motion.div
        variants={containerVariants}
        initial="hidden"
        animate="visible"
        className="relative z-10 flex-1 flex flex-col items-center px-4 sm:px-6 py-4 sm:py-5 overflow-y-auto overflow-x-hidden custom-scrollbar"
      >
        <div className="w-full max-w-2xl flex flex-col items-center">
          {/* ---- Hero Title ---- */}
          <motion.div variants={itemVariants} className="text-center mb-4">
            <h1 className="text-2xl sm:text-3xl font-bold tracking-tight gradient-text leading-tight">
              PPTX 模板编辑器
            </h1>
            <p className="mt-1.5 text-sm text-muted-foreground">
              上传 PowerPoint 文件，可视化编辑文本与表格，一键导出
            </p>
          </motion.div>

          {/* ---- Main Content Grid ---- */}
          <div className="w-full grid grid-cols-1 lg:grid-cols-5 gap-3">
            {/* ---- Upload Zone (left, wider) ---- */}
            <motion.div variants={itemVariants} className="lg:col-span-3">
              <div className="glass-strong rounded-xl p-3 h-full">
                {/* Upload area */}
                <div
                  className={cn(
                    'relative rounded-lg border-2 border-dashed transition-all duration-300 cursor-pointer upload-zone-glow',
                    isDragging
                      ? 'border-emerald-500 bg-emerald-50/60 dark:bg-emerald-950/30 scale-[1.01] is-dragging'
                      : 'border-muted-foreground/20 hover:border-emerald-400/60 hover:bg-emerald-50/30 dark:hover:bg-emerald-950/20',
                    isUploading && 'pointer-events-none opacity-70'
                  )}
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onDrop={handleDrop}
                  onClick={() => !isUploading && fileInputRef.current?.click()}
                  role="button"
                  tabIndex={0}
                  aria-label="上传 PPTX 文件"
                  onKeyDown={(e) => {
                    if (e.key === 'Enter' || e.key === ' ') {
                      e.preventDefault();
                      if (!isUploading) fileInputRef.current?.click();
                    }
                  }}
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".pptx"
                    onChange={handleFileSelect}
                    className="hidden"
                    aria-hidden="true"
                  />

                  <div className="flex flex-col items-center justify-center py-5 sm:py-6 px-4">
                    <AnimatePresence mode="wait">
                      {isUploading ? (
                        <motion.div
                          key="uploading"
                          initial={{ opacity: 0, scale: 0.8 }}
                          animate={{ opacity: 1, scale: 1 }}
                          exit={{ opacity: 0, scale: 0.8 }}
                          className="flex flex-col items-center gap-3 w-full max-w-xs"
                        >
                          <div className="relative">
                            <Loader2 className="size-8 text-emerald-600 dark:text-emerald-400 animate-spin" />
                            <div className="absolute inset-0 flex items-center justify-center">
                              <span className="text-[10px] font-bold text-emerald-700 dark:text-emerald-300">
                                {uploadProgress}%
                              </span>
                            </div>
                          </div>
                          <div className="w-full space-y-1.5">
                            <p className="text-xs font-medium text-center text-foreground">
                              正在上传并解析...
                            </p>
                            <div className="h-1.5 w-full rounded-full bg-emerald-100 dark:bg-emerald-900/40 overflow-hidden">
                              <motion.div
                                className="h-full rounded-full bg-gradient-to-r from-emerald-500 via-teal-500 to-cyan-500"
                                initial={{ width: 0 }}
                                animate={{ width: `${uploadProgress}%` }}
                                transition={{ duration: 0.3, ease: 'easeOut' }}
                              />
                            </div>
                          </div>
                        </motion.div>
                      ) : (
                        <motion.div
                          key="idle"
                          initial={{ opacity: 0, scale: 0.8 }}
                          animate={{ opacity: 1, scale: 1 }}
                          exit={{ opacity: 0, scale: 0.8 }}
                          className="flex flex-row items-center gap-3"
                        >
                          <motion.div
                            animate={isDragging ? { scale: 1.15, rotate: 5 } : { scale: 1, rotate: 0 }}
                            transition={{ type: 'spring', stiffness: 300, damping: 20 }}
                            className={cn(
                              'size-10 rounded-lg flex items-center justify-center transition-all duration-400 upload-icon-glow shrink-0',
                              isDragging
                                ? 'bg-emerald-500 text-white shadow-lg shadow-emerald-500/30'
                                : 'bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 hover:bg-emerald-500/15',
                              !isDragging && 'animate-breathe'
                            )}
                          >
                            <FileUp className="size-5" />
                          </motion.div>
                          <div className="text-left space-y-0.5 min-w-0">
                            <p className="text-sm font-semibold text-foreground">
                              {isDragging ? '松开以上传文件' : '点击或拖拽 PPTX 文件到此处'}
                            </p>
                            <p className="text-xs text-muted-foreground flex items-center gap-1">
                              <FileText className="size-3" />
                              支持 .pptx 格式，最大 100MB
                            </p>
                          </div>
                        </motion.div>
                      )}
                    </AnimatePresence>
                  </div>
                </div>
              </div>
            </motion.div>

            {/* ---- Right Sidebar ---- */}
            <motion.div variants={itemVariants} className="lg:col-span-2 flex flex-col gap-3">
              {/* ---- AI Generate Card ---- */}
              <div className="glass-strong rounded-xl p-3">
                <div className="flex items-center gap-2.5 mb-2">
                  <div className="size-8 rounded-lg bg-gradient-to-br from-emerald-500 to-teal-500 flex items-center justify-center shrink-0">
                    <Sparkles className="size-4 text-white" />
                  </div>
                  <div className="min-w-0 flex-1">
                    <h3 className="font-semibold text-sm leading-tight">AI 智能生成</h3>
                    <p className="text-[11px] text-muted-foreground">自动填充模板内容</p>
                  </div>
                  <Button
                    variant="ghost"
                    size="icon"
                    className="size-6 shrink-0 rounded-md text-muted-foreground/50 hover:text-violet-600 hover:bg-violet-50 dark:hover:text-violet-400 dark:hover:bg-violet-900/20"
                    onClick={onAiSettings}
                    title="AI 模型设置"
                  >
                    <Settings2 className="size-3.5" />
                  </Button>
                </div>
                <Button
                  onClick={onAiGenerate}
                  className="w-full bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-700 hover:to-teal-700 text-white"
                  size="sm"
                >
                  <Sparkles className="size-3.5" />
                  开始 AI 生成
                </Button>
              </div>

              {/* ---- JSON Import Card ---- */}
              <div className="glass-strong rounded-xl overflow-hidden">
                <button
                  onClick={() => setIsJsonExpanded(!isJsonExpanded)}
                  className="w-full p-3 flex items-center gap-2.5 text-left hover:bg-white/20 dark:hover:bg-white/5 transition-colors"
                >
                  <div className="size-7 rounded-md bg-teal-500/10 flex items-center justify-center shrink-0">
                    <Code2 className="size-3.5 text-teal-600 dark:text-teal-400" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <h3 className="font-semibold text-sm leading-tight">JSON 导入</h3>
                    <p className="text-[11px] text-muted-foreground">通过文件 ID 和 JSON 数据导入</p>
                  </div>
                  <motion.div
                    animate={{ rotate: isJsonExpanded ? 90 : 0 }}
                    transition={{ duration: 0.2 }}
                  >
                    <ChevronRight className="size-4 text-muted-foreground" />
                  </motion.div>
                </button>

                <AnimatePresence>
                  {isJsonExpanded && (
                    <motion.div
                      initial={{ height: 0, opacity: 0 }}
                      animate={{ height: 'auto', opacity: 1 }}
                      exit={{ height: 0, opacity: 0 }}
                      transition={{ duration: 0.25, ease: [0.25, 0.46, 0.45, 0.94] }}
                      className="overflow-hidden"
                    >
                      <div className="px-3 pb-3 space-y-2">
                        <Separator className="!mt-0" />
                        <div className="space-y-1.5">
                          <label className="text-[11px] font-medium text-muted-foreground">文件 ID</label>
                          <Input
                            placeholder="输入文件 ID"
                            value={jsonFileId}
                            onChange={(e) => setJsonFileId(e.target.value)}
                            className="h-7 text-xs"
                          />
                        </div>
                        <div className="space-y-1.5">
                          <label className="text-[11px] font-medium text-muted-foreground">JSON 数据</label>
                          <Textarea
                            placeholder='{"slides": [...]}'
                            value={jsonText}
                            onChange={(e) => setJsonText(e.target.value)}
                            className="min-h-20 text-xs font-mono resize-none"
                          />
                        </div>
                        <Button
                          onClick={handleJsonImport}
                          variant="outline"
                          size="sm"
                          className="w-full border-teal-300/50 dark:border-teal-700/50 text-teal-700 dark:text-teal-300 hover:bg-teal-50 dark:hover:bg-teal-950/30 h-7 text-xs"
                        >
                          <Code2 className="size-3.5" />
                          导入 JSON
                        </Button>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </div>
            </motion.div>
          </div>

          {/* ---- File History Section (Full Width) ---- */}
          {fileHistory.length > 0 && (
            <motion.div variants={itemVariants} className="w-full mt-5">
              {/* Section header */}
              <div className="flex items-center justify-between mb-2.5">
                <div className="flex items-center gap-2">
                  <div className="size-6 rounded-md bg-amber-500/10 flex items-center justify-center">
                    <History className="size-3.5 text-amber-600 dark:text-amber-400" />
                  </div>
                  <h3 className="font-semibold text-sm">文件历史</h3>
                  <Badge variant="secondary" className="text-[10px] px-1.5 py-0 h-4">
                    {fileHistory.length}
                  </Badge>
                </div>
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={handleClearHistory}
                  className="text-[11px] text-muted-foreground hover:text-destructive hover:bg-destructive/5 h-6 px-2"
                >
                  <Trash2 className="size-3 mr-1" />
                  清空
                </Button>
              </div>

              {/* History card grid */}
              <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-2.5">
                {fileHistory.map((entry, index) => (
                  <motion.div
                    key={entry.fileId}
                    custom={index}
                    variants={historyCardVariants}
                    initial="hidden"
                    animate="visible"
                    whileHover={{ y: -2, scale: 1.02 }}
                    className="group relative"
                  >
                    <div className="glass-strong rounded-lg overflow-hidden cursor-pointer transition-all duration-300 hover:shadow-lg hover:shadow-emerald-500/10 hover:border-emerald-400/30 border border-transparent"
                      onClick={() => handleReopenFile(entry)}
                    >
                      {/* Thumbnail area — uses /api/pptx/thumbnail endpoint */}
                      <div className="relative aspect-[16/9] bg-muted/30 overflow-hidden">
                        <HistoryThumbnail fileId={entry.fileId} fileName={entry.fileName} />

                        {/* Hover overlay with open button */}
                        <div className="absolute inset-0 bg-black/0 group-hover:bg-black/30 transition-all duration-300 flex items-center justify-center">
                          <motion.div
                            initial={{ opacity: 0, scale: 0.8 }}
                            whileHover={{ scale: 1.1 }}
                            className="opacity-0 group-hover:opacity-100 transition-opacity duration-300"
                          >
                            <div className="size-8 rounded-full bg-white/90 dark:bg-white/80 shadow-lg flex items-center justify-center">
                              {reopeningFileId === entry.fileId ? (
                                <Loader2 className="size-3.5 text-emerald-600 animate-spin" />
                              ) : (
                                <Play className="size-3.5 text-emerald-600 ml-0.5" />
                              )}
                            </div>
                          </motion.div>
                        </div>

                        {/* Slide count badge */}
                        <div className="absolute top-1.5 right-1.5">
                          <Badge className="text-[9px] px-1.5 py-0 h-4 bg-black/50 text-white border-0 backdrop-blur-sm">
                            <Layers className="size-2.5 mr-0.5" />
                            {entry.slideCount} 页
                          </Badge>
                        </div>
                      </div>

                      {/* Info area */}
                      <div className="p-2">
                        <p className="text-[11px] font-medium truncate leading-tight">{entry.fileName}</p>
                        <div className="flex items-center justify-between mt-0.5">
                          <span className="text-[10px] text-muted-foreground flex items-center gap-1">
                            <Clock className="size-2.5" />
                            {formatTimeAgo(entry.openedAt)}
                          </span>
                          <Button
                            variant="ghost"
                            size="icon"
                            className="size-5 opacity-0 group-hover:opacity-100 transition-opacity hover:bg-destructive/10"
                            onClick={(e) => {
                              e.stopPropagation();
                              handleDeleteHistory(entry.fileId);
                            }}
                            title="删除记录"
                          >
                            <X className="size-3 text-muted-foreground hover:text-destructive" />
                          </Button>
                        </div>
                      </div>
                    </div>
                  </motion.div>
                ))}
              </div>
            </motion.div>
          )}
        </div>
      </motion.div>

    </div>
  );
}
