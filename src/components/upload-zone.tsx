'use client'

import React from 'react';
import { usePptxStore, type PptxSlideData, type PptxJsonData } from '@/lib/pptx-store';
import {
  getFileHistory,
  addFileHistory,
  removeFileHistory,
  clearFileHistory,
  formatTimeAgo,
  type FileHistoryEntry,
} from '@/lib/pptx-history';
import { cn } from '@/lib/utils';
import { toast } from 'sonner';
import { motion, AnimatePresence } from 'framer-motion';
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from '@/components/ui/card';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Input } from '@/components/ui/input';
import { Separator } from '@/components/ui/separator';
import { ScrollArea } from '@/components/ui/scroll-area';
import {
  Upload,
  FileUp,
  Trash2,
  Clock,
  Sparkles,
  FileText,
  X,
  Loader2,
  History,
  Braces,
  Presentation,
  AlertCircle,
} from 'lucide-react';

interface UploadZoneProps {
  onAiGenerate: () => void;
}

export function UploadZone({ onAiGenerate }: UploadZoneProps) {
  const { setStep, setParsedData } = usePptxStore();
  const [isDragging, setIsDragging] = React.useState(false);
  const [isUploading, setIsUploading] = React.useState(false);
  const [uploadProgress, setUploadProgress] = React.useState(0);
  const [fileHistory, setFileHistory] = React.useState<FileHistoryEntry[]>([]);
  const [showJsonInput, setShowJsonInput] = React.useState(false);
  const [jsonText, setJsonText] = React.useState('');
  const [jsonFileId, setJsonFileId] = React.useState('');
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const jsonFileInputRef = React.useRef<HTMLInputElement>(null);

  React.useEffect(() => {
    setFileHistory(getFileHistory());
  }, []);

  const handleUpload = async (file: File) => {
    if (!file.name.toLowerCase().endsWith('.pptx')) {
      toast.error('仅支持 .pptx 格式文件');
      return;
    }
    if (file.size > 100 * 1024 * 1024) {
      toast.error('文件大小不能超过 100MB');
      return;
    }

    setIsUploading(true);
    setUploadProgress(0);
    setStep('loading');

    try {
      const formData = new FormData();
      formData.append('file', file);

      // Simulate progress
      const progressInterval = setInterval(() => {
        setUploadProgress((prev) => Math.min(prev + Math.random() * 15, 90));
      }, 300);

      const response = await fetch('/api/pptx/parse', {
        method: 'POST',
        body: formData,
      });

      clearInterval(progressInterval);
      setUploadProgress(95);

      if (!response.ok) {
        const errorData = await response.json().catch(() => ({ error: '上传失败' }));
        throw new Error(errorData.error || '上传失败');
      }

      const data = await response.json();
      setUploadProgress(100);

      // Store in history
      addFileHistory({
        fileId: data.fileId,
        fileName: data.fileName,
        slideCount: data.slideCount,
        openedAt: Date.now(),
      });
      setFileHistory(getFileHistory());

      setParsedData(data.fileId, data.fileName, data.slides);
      toast.success(`成功解析 ${data.slideCount} 页幻灯片`);
    } catch (err) {
      setStep('upload');
      toast.error(err instanceof Error ? err.message : '文件上传失败，请重试');
    } finally {
      setIsUploading(false);
      setUploadProgress(0);
    }
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) handleUpload(file);
  };

  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
  };

  const handleFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleUpload(file);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleReopenFile = async (entry: FileHistoryEntry) => {
    try {
      // Check if file still exists
      const checkRes = await fetch(`/api/pptx/check?fileId=${entry.fileId}`);
      const checkData = await checkRes.json();

      if (!checkData.exists) {
        toast.error('文件已过期，请重新上传');
        removeFileHistory(entry.fileId);
        setFileHistory(getFileHistory());
        return;
      }

      // Reparse the file
      setStep('loading');
      const reparseRes = await fetch('/api/pptx/reparse', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileId: entry.fileId }),
      });

      if (!reparseRes.ok) {
        const errorData = await reparseRes.json().catch(() => ({ error: '重新解析失败' }));
        throw new Error(errorData.error || '重新解析失败');
      }

      const data = await reparseRes.json();

      addFileHistory({
        fileId: data.fileId,
        fileName: entry.fileName,
        slideCount: data.slideCount,
        openedAt: Date.now(),
      });
      setFileHistory(getFileHistory());

      setParsedData(data.fileId, entry.fileName, data.slides);
      toast.success(`成功重新打开 ${data.slideCount} 页幻灯片`);
    } catch (err) {
      setStep('upload');
      toast.error(err instanceof Error ? err.message : '重新打开失败');
    }
  };

  const handleDeleteHistory = (fileId: string) => {
    removeFileHistory(fileId);
    setFileHistory(getFileHistory());
    toast.success('已删除历史记录');
  };

  const handleClearHistory = () => {
    clearFileHistory();
    setFileHistory([]);
    toast.success('已清除所有历史记录');
  };

  const handleJsonImport = () => {
    if (!jsonText.trim()) {
      toast.error('请输入 JSON 数据');
      return;
    }
    if (!jsonFileId.trim()) {
      toast.error('请输入文件 ID');
      return;
    }

    try {
      const jsonData: PptxJsonData = JSON.parse(jsonText);

      if (!jsonData.slides || !Array.isArray(jsonData.slides)) {
        toast.error('JSON 格式错误：缺少 slides 数组');
        return;
      }

      // Check if the file still exists
      fetch(`/api/pptx/check?fileId=${jsonFileId}`)
        .then((res) => res.json())
        .then((checkData) => {
          if (!checkData.exists) {
            toast.error('文件 ID 对应的 PPTX 已过期，请重新上传');
            return;
          }

          // Load the PPTX first via reparse
          setStep('loading');
          fetch('/api/pptx/reparse', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ fileId: jsonFileId }),
          })
            .then((res) => res.json())
            .then((data) => {
              const store = usePptxStore.getState();
              store.setParsedData(data.fileId, data.fileName, data.slides);
              // Now apply JSON modifications
              store.loadFromJson(jsonData);
              toast.success('JSON 数据导入成功');
            })
            .catch((err) => {
              setStep('upload');
              toast.error(err instanceof Error ? err.message : '加载文件失败');
            });
        });
    } catch {
      toast.error('JSON 格式无效，请检查输入');
    }
  };

  const handleJsonFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (ev) => {
      const text = ev.target?.result as string;
      setJsonText(text);

      // Try to extract fileId from JSON
      try {
        const data = JSON.parse(text);
        if (data.fileId) setJsonFileId(data.fileId);
      } catch { /* ignore */ }
    };
    reader.readAsText(file);

    if (jsonFileInputRef.current) jsonFileInputRef.current.value = '';
  };

  return (
    <div className="min-h-screen flex items-center justify-center p-4 sm:p-8 bg-gradient-to-br from-background via-background to-muted/30">
      <div className="w-full max-w-4xl space-y-6">
        {/* Title area */}
        <motion.div
          initial={{ opacity: 0, y: -20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.5 }}
          className="text-center space-y-2"
        >
          <div className="flex items-center justify-center gap-2 mb-4">
            <div className="w-10 h-10 rounded-xl bg-primary/10 flex items-center justify-center">
              <Presentation className="w-6 h-6 text-primary" />
            </div>
          </div>
          <h1 className="text-2xl sm:text-3xl font-bold tracking-tight">PPTX 模板编辑器</h1>
          <p className="text-sm text-muted-foreground max-w-md mx-auto">
            上传 PowerPoint 文件，编辑文本、表格和图片，AI 辅助生成报告内容
          </p>
        </motion.div>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
          {/* Main upload zone */}
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.1 }}
            className="lg:col-span-2"
          >
            <Card className="overflow-hidden">
              <CardContent className="p-0">
                {/* Drop zone */}
                <div
                  className={cn(
                    'relative flex flex-col items-center justify-center p-8 sm:p-12 transition-all cursor-pointer min-h-[280px]',
                    isDragging
                      ? 'bg-primary/5 border-2 border-dashed border-primary'
                      : 'hover:bg-muted/30 border-2 border-dashed border-transparent',
                    isUploading && 'pointer-events-none',
                  )}
                  onDrop={handleDrop}
                  onDragOver={handleDragOver}
                  onDragLeave={handleDragLeave}
                  onClick={() => !isUploading && fileInputRef.current?.click()}
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".pptx"
                    className="hidden"
                    onChange={handleFileSelect}
                  />

                  {isUploading ? (
                    <div className="flex flex-col items-center gap-4">
                      <Loader2 className="w-12 h-12 text-primary animate-spin" />
                      <div className="text-center">
                        <p className="text-sm font-medium">正在解析文件...</p>
                        <p className="text-xs text-muted-foreground mt-1">
                          {uploadProgress < 50 ? '上传中' : uploadProgress < 90 ? '解析中' : '即将完成'}
                        </p>
                      </div>
                      <div className="w-48 h-1.5 bg-muted rounded-full overflow-hidden">
                        <motion.div
                          className="h-full bg-primary rounded-full"
                          initial={{ width: 0 }}
                          animate={{ width: `${uploadProgress}%` }}
                          transition={{ duration: 0.3 }}
                        />
                      </div>
                    </div>
                  ) : (
                    <>
                      <motion.div
                        animate={isDragging ? { scale: 1.1, y: -5 } : { scale: 1, y: 0 }}
                        transition={{ duration: 0.2 }}
                      >
                        <div className={cn(
                          'w-16 h-16 rounded-2xl flex items-center justify-center mb-4 transition-colors',
                          isDragging ? 'bg-primary/10' : 'bg-muted',
                        )}>
                          <FileUp className={cn(
                            'w-8 h-8 transition-colors',
                            isDragging ? 'text-primary' : 'text-muted-foreground',
                          )} />
                        </div>
                      </motion.div>
                      <h3 className="text-base font-medium mb-1">
                        {isDragging ? '释放以上传文件' : '拖放 PPTX 文件到此处'}
                      </h3>
                      <p className="text-sm text-muted-foreground mb-4">
                        或点击选择文件
                      </p>
                      <Button variant="outline" size="sm" className="gap-2">
                        <Upload className="w-4 h-4" />
                        选择文件
                      </Button>
                      <p className="text-[10px] text-muted-foreground/60 mt-4">
                        支持 .pptx 格式，最大 100MB
                      </p>
                    </>
                  )}
                </div>
              </CardContent>
            </Card>
          </motion.div>

          {/* Right sidebar */}
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            animate={{ opacity: 1, y: 0 }}
            transition={{ duration: 0.5, delay: 0.2 }}
            className="space-y-4"
          >
            {/* AI Card */}
            <Card className="overflow-hidden border-primary/20 bg-gradient-to-br from-card to-primary/5">
              <CardContent className="p-4">
                <div className="flex items-center gap-2 mb-2">
                  <Sparkles className="w-4 h-4 text-primary" />
                  <span className="text-sm font-medium">AI 智能填充</span>
                </div>
                <p className="text-xs text-muted-foreground mb-3">
                  上传数据源文件，AI 自动匹配模板字段并填充内容
                </p>
                <Button
                  variant="default"
                  size="sm"
                  className="w-full gap-2"
                  onClick={onAiGenerate}
                >
                  <Sparkles className="w-3.5 h-3.5" />
                  AI 生成
                </Button>
              </CardContent>
            </Card>

            {/* JSON Import Card */}
            <Card>
              <CardContent className="p-4">
                <button
                  className="w-full flex items-center gap-2 mb-2"
                  onClick={() => setShowJsonInput(!showJsonInput)}
                >
                  <Braces className="w-4 h-4 text-muted-foreground" />
                  <span className="text-sm font-medium">导入 JSON</span>
                </button>
                <AnimatePresence>
                  {showJsonInput && (
                    <motion.div
                      initial={{ height: 0, opacity: 0 }}
                      animate={{ height: 'auto', opacity: 1 }}
                      exit={{ height: 0, opacity: 0 }}
                      transition={{ duration: 0.2 }}
                      className="overflow-hidden"
                    >
                      <div className="space-y-2 pt-2">
                        <div>
                          <label className="text-[10px] text-muted-foreground mb-1 block">文件 ID</label>
                          <Input
                            value={jsonFileId}
                            onChange={(e) => setJsonFileId(e.target.value)}
                            placeholder="输入已上传文件的 ID"
                            className="h-8 text-xs"
                          />
                        </div>
                        <div>
                          <div className="flex items-center justify-between mb-1">
                            <label className="text-[10px] text-muted-foreground">JSON 数据</label>
                            <input
                              ref={jsonFileInputRef}
                              type="file"
                              accept=".json"
                              className="hidden"
                              onChange={handleJsonFileSelect}
                            />
                            <Button
                              variant="ghost"
                              size="sm"
                              className="h-5 text-[10px] px-1"
                              onClick={() => jsonFileInputRef.current?.click()}
                            >
                              从文件加载
                            </Button>
                          </div>
                          <textarea
                            value={jsonText}
                            onChange={(e) => setJsonText(e.target.value)}
                            placeholder='{"slides": [...]}'
                            className="w-full h-20 text-[10px] font-mono bg-muted/50 rounded p-2 resize-y border border-input focus:outline-none focus:ring-1 focus:ring-primary"
                          />
                        </div>
                        <Button
                          variant="outline"
                          size="sm"
                          className="w-full h-7 text-xs"
                          onClick={handleJsonImport}
                        >
                          应用 JSON 修改
                        </Button>
                      </div>
                    </motion.div>
                  )}
                </AnimatePresence>
              </CardContent>
            </Card>

            {/* File history */}
            <Card>
              <CardHeader className="p-4 pb-2">
                <div className="flex items-center justify-between">
                  <div className="flex items-center gap-2">
                    <History className="w-4 h-4 text-muted-foreground" />
                    <CardTitle className="text-sm">历史记录</CardTitle>
                  </div>
                  {fileHistory.length > 0 && (
                    <Button
                      variant="ghost"
                      size="sm"
                      className="h-5 text-[10px] px-1 text-muted-foreground"
                      onClick={handleClearHistory}
                    >
                      清除
                    </Button>
                  )}
                </div>
              </CardHeader>
              <CardContent className="p-4 pt-0">
                {fileHistory.length === 0 ? (
                  <p className="text-xs text-muted-foreground text-center py-4">
                    暂无历史记录
                  </p>
                ) : (
                  <ScrollArea className="max-h-48">
                    <div className="space-y-1.5">
                      {fileHistory.map((entry) => (
                        <div
                          key={entry.fileId}
                          className="flex items-center gap-2 p-2 rounded-md hover:bg-muted/50 group transition-colors"
                        >
                          <FileText className="w-3.5 h-3.5 text-muted-foreground shrink-0" />
                          <div className="flex-1 min-w-0">
                            <p className="text-xs font-medium truncate">{entry.fileName}</p>
                            <p className="text-[9px] text-muted-foreground">
                              {entry.slideCount} 页 · {formatTimeAgo(entry.openedAt)}
                            </p>
                          </div>
                          <div className="flex items-center gap-0.5 opacity-0 group-hover:opacity-100 transition-opacity">
                            <Button
                              variant="ghost"
                              size="sm"
                              className="h-5 w-5 p-0"
                              onClick={() => handleReopenFile(entry)}
                              title="重新打开"
                            >
                              <Upload className="w-3 h-3" />
                            </Button>
                            <Button
                              variant="ghost"
                              size="sm"
                              className="h-5 w-5 p-0 text-destructive hover:text-destructive"
                              onClick={() => handleDeleteHistory(entry.fileId)}
                              title="删除"
                            >
                              <Trash2 className="w-3 h-3" />
                            </Button>
                          </div>
                        </div>
                      ))}
                    </div>
                  </ScrollArea>
                )}
              </CardContent>
            </Card>
          </motion.div>
        </div>
      </div>
    </div>
  );
}
