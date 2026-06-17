'use client';

import React from 'react';
import {
  Dialog,
  DialogContent,
  DialogHeader,
  DialogTitle,
  DialogDescription,
} from '@/components/ui/dialog';
import { Keyboard } from 'lucide-react';

interface KeyboardShortcutsDialogProps {
  open: boolean;
  onOpenChange: (open: boolean) => void;
}

const SHORTCUT_GROUPS = [
  {
    title: 'Editing',
    shortcuts: [
      { keys: ['Ctrl', 'Z'], description: 'Undo' },
      { keys: ['Ctrl', 'Y'], description: 'Redo' },
      { keys: ['Ctrl', '⇧', 'Z'], description: 'Redo (alternative)' },
    ],
  },
  {
    title: 'Navigation',
    shortcuts: [
      { keys: ['←', 'PageUp'], description: 'Previous slide' },
      { keys: ['→', 'PageDown'], description: 'Next slide' },
      { keys: ['Ctrl', '/'], description: 'Show keyboard shortcuts' },
    ],
  },
  {
    title: 'Presentation',
    shortcuts: [
      { keys: ['F5'], description: 'Enter presentation mode' },
      { keys: ['Esc'], description: 'Exit presentation mode' },
      { keys: ['←', 'Space'], description: 'Next slide (presentation)' },
      { keys: ['→', 'Backspace'], description: 'Previous slide (presentation)' },
    ],
  },
  {
    title: 'Tools',
    shortcuts: [
      { keys: ['Ctrl', 'S'], description: 'Save as JSON' },
      { keys: ['Ctrl', 'E'], description: 'Export PPTX' },
      { keys: ['Ctrl', '+'], description: 'Zoom in preview' },
      { keys: ['Ctrl', '-'], description: 'Zoom out preview' },
      { keys: ['Ctrl', '0'], description: 'Reset zoom (fit)' },
    ],
  },
];

function Kbd({ children }: { children: React.ReactNode }) {
  return (
    <kbd className="inline-flex items-center justify-center min-w-[20px] h-5 px-1.5 rounded text-[10px] font-medium bg-muted border border-border/60 text-muted-foreground shadow-sm">
      {children}
    </kbd>
  );
}

export function KeyboardShortcutsDialog({ open, onOpenChange }: KeyboardShortcutsDialogProps) {
  return (
    <Dialog open={open} onOpenChange={onOpenChange}>
      <DialogContent className="sm:max-w-[400px] p-0 gap-0 overflow-hidden">
        <DialogHeader className="px-4 py-3 border-b border-border/30">
          <DialogTitle className="flex items-center gap-2 text-sm">
            <Keyboard className="size-4 text-emerald-500" />
            Keyboard Shortcuts
          </DialogTitle>
          <DialogDescription className="sr-only">
            Available keyboard shortcuts for the editor
          </DialogDescription>
        </DialogHeader>
        <div className="px-4 py-3 space-y-4">
          {SHORTCUT_GROUPS.map((group) => (
            <div key={group.title}>
              <h4 className="text-[10px] font-bold uppercase tracking-widest text-muted-foreground/60 mb-2">
                {group.title}
              </h4>
              <div className="space-y-1.5">
                {group.shortcuts.map((shortcut) => (
                  <div
                    key={shortcut.description}
                    className="flex items-center justify-between gap-4"
                  >
                    <span className="text-xs text-foreground/80">{shortcut.description}</span>
                    <div className="flex items-center gap-0.5">
                      {shortcut.keys.map((key, i) => (
                        <React.Fragment key={i}>
                          {i > 0 && (
                            <span className="text-[9px] text-muted-foreground/40 mx-px">+</span>
                          )}
                          <Kbd>{key}</Kbd>
                        </React.Fragment>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      </DialogContent>
    </Dialog>
  );
}
