export interface FileHistoryEntry {
  fileId: string;
  fileName: string;
  slideCount: number;
  openedAt: number;
}

const STORAGE_KEY = 'pptx-file-history';
const MAX_HISTORY_ENTRIES = 20;

export function getFileHistory(): FileHistoryEntry[] {
  if (typeof window === 'undefined') return [];
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    const entries: FileHistoryEntry[] = JSON.parse(raw);
    return entries.sort((a, b) => b.openedAt - a.openedAt);
  } catch { return []; }
}

export function addFileHistory(entry: FileHistoryEntry): void {
  if (typeof window === 'undefined') return;
  try {
    const existing = getFileHistory();
    const filtered = existing.filter((e) => e.fileId !== entry.fileId);
    filtered.unshift(entry);
    const trimmed = filtered.slice(0, MAX_HISTORY_ENTRIES);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(trimmed));
  } catch { /* ignore */ }
}

export function removeFileHistory(fileId: string): void {
  if (typeof window === 'undefined') return;
  try {
    const existing = getFileHistory();
    const filtered = existing.filter((e) => e.fileId !== fileId);
    localStorage.setItem(STORAGE_KEY, JSON.stringify(filtered));
  } catch { /* ignore */ }
}

export function clearFileHistory(): void {
  if (typeof window === 'undefined') return;
  try { localStorage.removeItem(STORAGE_KEY); } catch { /* ignore */ }
}

export function formatTimeAgo(timestamp: number): string {
  const now = Date.now();
  const diff = now - timestamp;
  const seconds = Math.floor(diff / 1000);
  const minutes = Math.floor(seconds / 60);
  const hours = Math.floor(minutes / 60);
  const days = Math.floor(hours / 24);

  if (days > 7) return new Date(timestamp).toLocaleDateString('zh-CN', { month: 'short', day: 'numeric' });
  if (days > 0) return `${days} 天前`;
  if (hours > 0) return `${hours} 小时前`;
  if (minutes > 0) return `${minutes} 分钟前`;
  return '刚刚';
}
