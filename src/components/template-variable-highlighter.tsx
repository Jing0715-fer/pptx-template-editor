'use client';

import { useMemo } from 'react';
import { cn } from '@/lib/utils';

// ============================================================================
// Variable Detection Utilities
// ============================================================================

/** Regex to match {{variable_name}} patterns */
const VARIABLE_REGEX = /\{\{([^}]+)\}\}/g;

/**
 * Extract all unique template variable names from a text string.
 * e.g. "Hello {{name}}, welcome to {{company}}" → ["name", "company"]
 */
export function extractVariables(text: string): string[] {
  const matches = text.matchAll(/\{\{([^}]+)\}\}/g);
  const seen = new Set<string>();
  const result: string[] = [];
  for (const match of matches) {
    const varName = match[1].trim();
    if (varName && !seen.has(varName)) {
      seen.add(varName);
      result.push(varName);
    }
  }
  return result;
}

/**
 * Count the number of template variables in a text string.
 */
export function countVariables(text: string): number {
  const matches = text.match(/\{\{[^}]+\}\}/g);
  return matches ? matches.length : 0;
}

/**
 * Check if a text string contains any template variables.
 */
export function hasVariables(text: string): boolean {
  return /\{\{[^}]+\}\}/.test(text);
}

// ============================================================================
// Parsed Segment Type
// ============================================================================

interface TextSegment {
  type: 'text' | 'variable';
  content: string;
}

/**
 * Parse text into segments of plain text and template variables.
 */
function parseTextSegments(text: string): TextSegment[] {
  const segments: TextSegment[] = [];
  let lastIndex = 0;

  // Reset regex state
  const regex = new RegExp(VARIABLE_REGEX.source, 'g');
  let match: RegExpExecArray | null;

  while ((match = regex.exec(text)) !== null) {
    // Add plain text before this variable
    if (match.index > lastIndex) {
      segments.push({ type: 'text', content: text.slice(lastIndex, match.index) });
    }
    // Add the variable (with the full {{...}} syntax)
    segments.push({ type: 'variable', content: match[1].trim() });
    lastIndex = regex.lastIndex;
  }

  // Add remaining plain text
  if (lastIndex < text.length) {
    segments.push({ type: 'text', content: text.slice(lastIndex) });
  }

  return segments;
}

// ============================================================================
// TemplateVariableHighlighter Component
// ============================================================================

interface TemplateVariableHighlighterProps {
  /** The text content to parse and render */
  text: string;
  /** Additional class names for the wrapper */
  className?: string;
  /** Whether to render in inline mode (no wrapper div, just spans) */
  inline?: boolean;
}

/**
 * Parses text content to find {{variable_name}} patterns and renders them
 * as inline colored chips/badges with a distinctive visual style.
 * Non-variable text renders normally.
 */
export function TemplateVariableHighlighter({
  text,
  className,
  inline = false,
}: TemplateVariableHighlighterProps) {
  const segments = useMemo(() => parseTextSegments(text), [text]);

  if (!text) {
    return <span className={cn('italic opacity-60', className)}>Empty</span>;
  }

  const content = segments.map((segment, index) => {
    if (segment.type === 'variable') {
      return (
        <span
          key={`var-${index}`}
          className={cn(
            'inline-flex items-center align-baseline',
            'bg-emerald-100/80 text-emerald-700 border border-emerald-200/50',
            'dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-700/30',
            'rounded px-1 text-[10px] font-mono whitespace-nowrap'
          )}
        >
          {segment.content}
        </span>
      );
    }
    return <span key={`txt-${index}`}>{segment.content}</span>;
  });

  if (inline) {
    return <span className={className}>{content}</span>;
  }

  return <div className={cn('leading-relaxed', className)}>{content}</div>;
}

// ============================================================================
// VariableChip Component (for the Variables panel)
// ============================================================================

interface VariableChipProps {
  /** The variable name (without {{ }}) */
  name: string;
  /** How many elements use this variable */
  usageCount: number;
  /** Current value of the variable, or undefined if not set */
  currentValue?: string;
  /** Click handler */
  onClick?: () => void;
}

/**
 * A compact chip for displaying a template variable in the Variables panel.
 */
export function VariableChip({ name, usageCount, currentValue, onClick }: VariableChipProps) {
  return (
    <button
      onClick={onClick}
      className={cn(
        'w-full flex items-center gap-1 px-1.5 py-1 rounded text-left transition-all duration-150',
        'hover:bg-emerald-50/30 dark:hover:bg-emerald-950/15',
        'focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
        'group/var'
      )}
    >
      <span
        className={cn(
          'inline-flex items-center shrink-0',
          'bg-emerald-100/80 text-emerald-700 border border-emerald-200/50',
          'dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-700/30',
          'rounded px-1 text-[9px] font-mono'
        )}
      >
        {name}
      </span>
      <div className="flex-1 min-w-0">
        <span className="text-[8px] text-muted-foreground truncate block">
          {currentValue || 'Not set'}
        </span>
      </div>
      <span
        className={cn(
          'shrink-0 inline-flex items-center justify-center',
          'min-w-[14px] h-3 px-[3px] rounded-full text-[7px] font-semibold leading-none',
          'bg-emerald-500/10 text-emerald-600 dark:bg-emerald-500/15 dark:text-emerald-400',
          'group-hover/var:bg-emerald-500/20'
        )}
      >
        {usageCount}
      </span>
    </button>
  );
}
