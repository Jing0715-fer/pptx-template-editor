'use client'

import React, { useCallback, useState, useEffect, useRef, useMemo, useSyncExternalStore } from 'react';
import { motion, AnimatePresence, useMotionValue, useSpring, useTransform } from 'framer-motion';
import {
  Upload, Sparkles, FileJson, FileSpreadsheet, ArrowRight,
  AlertCircle, Presentation,
  Zap, Shield, Globe,
  BookOpen,
  Menu, FileText, Play,
  Code2, History, Trash2,
  ChevronLeft, ChevronRight, Edit3,
  Minus, Plus, Maximize2, Grid3x3,
} from 'lucide-react';
import { Button } from '@/components/ui/button';
import { Badge } from '@/components/ui/badge';
import { Progress } from '@/components/ui/progress';
import {
  Tooltip,
  TooltipTrigger,
  TooltipContent,
  TooltipProvider,
} from '@/components/ui/tooltip';

import {
  Accordion, AccordionContent, AccordionItem, AccordionTrigger,
} from '@/components/ui/accordion';
import { usePptxStore, type PptxSlideData, type PptxTextElement } from '@/lib/pptx-store';
import {
  getFileHistory, addFileHistory, removeFileHistory, clearFileHistory, formatTimeAgo,
  type FileHistoryEntry,
} from '@/lib/pptx-history';
import { cn } from '@/lib/utils';
import { toast } from 'sonner';
import { ThemeToggle } from '@/components/theme-toggle';
import SlideNavigator from '@/components/slide-navigator';
import SlideEditor from '@/components/slide-editor';
import { SlidePreview } from '@/components/slide-preview';
import Toolbar from '@/components/toolbar';
import { AiGenerateDialog } from '@/components/ai-generate-dialog';
import { AiSettingsDialog } from '@/components/ai-settings-dialog';
import { KeyboardShortcutsDialog } from '@/components/keyboard-shortcuts-dialog';
import PresentationMode from '@/components/presentation-mode';

// ============================================================================
// Template Data (converted to PptxSlideData format)
// ============================================================================

interface TemplateSlideData {
  slideNumber: number;
  elements: Array<{
    id: string;
    name: string;
    text: string;
  }>;
}

interface TemplateInfo {
  id: string;
  name: string;
  description: string;
  icon: string;
  slides: number;
  variables: number;
  gradient: string;
  accentColor: string;
  accentBorder: string;
  accentBg: string;
  difficulty: 'Simple' | 'Intermediate' | 'Advanced';
  isPopular: boolean;
  slideData: TemplateSlideData[];
  fileName: string;
}

const DEMO_SLIDES: TemplateSlideData[] = [
  { slideNumber: 1, elements: [
    { id: 't1-1', name: 'Title', text: '{{company_name}} Annual Report' },
    { id: 't1-2', name: 'Subtitle', text: 'Fiscal Year {{fiscal_year}}' },
    { id: 't1-3', name: 'Date', text: 'Presented by {{presenter_name}} - {{presentation_date}}' },
  ]},
  { slideNumber: 2, elements: [
    { id: 't2-1', name: 'Section Title', text: 'Executive Summary' },
    { id: 't2-2', name: 'Content', text: '{{executive_summary}}' },
  ]},
  { slideNumber: 3, elements: [
    { id: 't3-1', name: 'Metrics Title', text: 'Key Metrics' },
    { id: 't3-2', name: 'Revenue', text: 'Revenue: {{revenue_amount}}' },
    { id: 't3-3', name: 'Growth', text: 'Growth: {{growth_percentage}}%' },
    { id: 't3-4', name: 'Employees', text: 'Team: {{employee_count}} employees' },
  ]},
  { slideNumber: 4, elements: [
    { id: 't4-1', name: 'Thank You', text: 'Thank You' },
    { id: 't4-2', name: 'Contact', text: 'Contact: {{contact_email}}' },
  ]},
];

const MARKETING_SLIDES: TemplateSlideData[] = [
  { slideNumber: 1, elements: [
    { id: 'm1-1', name: 'Title', text: '{{brand_name}} Marketing Plan' },
    { id: 'm1-2', name: 'Subtitle', text: 'Campaign: {{campaign_name}}' },
    { id: 'm1-3', name: 'Date', text: '{{quarter}} - Prepared by {{marketing_lead}}' },
  ]},
  { slideNumber: 2, elements: [
    { id: 'm2-1', name: 'Section Title', text: 'Strategy Overview' },
    { id: 'm2-2', name: 'Content', text: 'Target Audience: {{target_audience}}' },
  ]},
  { slideNumber: 3, elements: [
    { id: 'm3-1', name: 'Budget Title', text: 'Budget & Timeline' },
    { id: 'm3-2', name: 'Budget', text: 'Total Budget: {{budget_amount}}' },
  ]},
];

const PROPOSAL_SLIDES: TemplateSlideData[] = [
  { slideNumber: 1, elements: [
    { id: 'p1-1', name: 'Title', text: '{{project_name}} Proposal' },
    { id: 'p1-2', name: 'Client', text: 'Prepared for {{client_name}}' },
    { id: 'p1-3', name: 'Date', text: '{{proposal_date}} | By {{project_manager}}' },
  ]},
  { slideNumber: 2, elements: [
    { id: 'p2-1', name: 'Section Title', text: 'Project Scope' },
    { id: 'p2-2', name: 'Scope', text: '{{project_scope}}' },
  ]},
  { slideNumber: 3, elements: [
    { id: 'p3-1', name: 'Timeline Title', text: 'Timeline & Deliverables' },
    { id: 'p3-2', name: 'Timeline', text: 'Estimated Duration: {{timeline_weeks}} weeks' },
    { id: 'p3-3', name: 'Cost', text: 'Estimated Cost: {{estimated_cost}}' },
  ]},
  { slideNumber: 4, elements: [
    { id: 'p4-1', name: 'Thank You', text: 'Next Steps' },
    { id: 'p4-2', name: 'Contact', text: 'Contact: {{contact_person}} at {{contact_email}}' },
  ]},
];

const TEAM_SLIDES: TemplateSlideData[] = [
  { slideNumber: 1, elements: [
    { id: 'tm1-1', name: 'Title', text: '{{team_name}} Weekly Update' },
    { id: 'tm1-2', name: 'Date', text: 'Week of {{report_week}} | Led by {{team_lead}}' },
  ]},
  { slideNumber: 2, elements: [
    { id: 'tm2-1', name: 'Section Title', text: 'Progress' },
    { id: 'tm2-2', name: 'Content', text: '{{key_accomplishment}}' },
  ]},
  { slideNumber: 3, elements: [
    { id: 'tm3-1', name: 'Next Title', text: 'Upcoming' },
    { id: 'tm3-2', name: 'Next Steps', text: '{{next_priority}}' },
  ]},
];

const PRODUCT_SLIDES: TemplateSlideData[] = [
  { slideNumber: 1, elements: [
    { id: 'pl1-1', name: 'Title', text: '{{product_name}} Launch' },
    { id: 'pl1-2', name: 'Subtitle', text: 'Prepared by {{launch_lead}}' },
    { id: 'pl1-3', name: 'Date', text: '{{launch_date}}' },
  ]},
  { slideNumber: 2, elements: [
    { id: 'pl2-1', name: 'Section Title', text: 'Key Features' },
    { id: 'pl2-2', name: 'Content', text: '{{key_features}}' },
  ]},
  { slideNumber: 3, elements: [
    { id: 'pl3-1', name: 'Section Title', text: 'Pricing & Availability' },
    { id: 'pl3-2', name: 'Pricing', text: 'Starting at {{pricing}}' },
    { id: 'pl3-3', name: 'Availability', text: 'Available {{availability_date}}' },
  ]},
];

const QUARTERLY_SLIDES: TemplateSlideData[] = [
  { slideNumber: 1, elements: [
    { id: 'q1-1', name: 'Title', text: '{{department_name}} Quarterly Review' },
    { id: 'q1-2', name: 'Quarter', text: 'Q{{quarter_number}} {{review_year}}' },
    { id: 'q1-3', name: 'Presenter', text: 'Presented by {{reviewer_name}}' },
  ]},
  { slideNumber: 2, elements: [
    { id: 'q2-1', name: 'Section Title', text: 'Achievements' },
    { id: 'q2-2', name: 'Content', text: '{{key_achievements}}' },
  ]},
  { slideNumber: 3, elements: [
    { id: 'q3-1', name: 'Section Title', text: 'Metrics' },
    { id: 'q3-2', name: 'Revenue', text: 'Revenue: {{quarter_revenue}}' },
    { id: 'q3-3', name: 'Growth', text: 'Growth: {{growth_rate}}%' },
  ]},
  { slideNumber: 4, elements: [
    { id: 'q4-1', name: 'Section Title', text: 'Next Steps' },
    { id: 'q4-2', name: 'Content', text: '{{next_steps}}' },
  ]},
];

const TRAINING_SLIDES: TemplateSlideData[] = [
  { slideNumber: 1, elements: [
    { id: 'tr1-1', name: 'Title', text: '{{course_title}} Training' },
    { id: 'tr1-2', name: 'Instructor', text: 'Instructor: {{instructor_name}}' },
    { id: 'tr1-3', name: 'Duration', text: 'Duration: {{duration_hours}} hours' },
  ]},
  { slideNumber: 2, elements: [
    { id: 'tr2-1', name: 'Section Title', text: 'Learning Objectives' },
    { id: 'tr2-2', name: 'Content', text: '{{learning_objectives}}' },
  ]},
  { slideNumber: 3, elements: [
    { id: 'tr3-1', name: 'Section Title', text: 'Assessment' },
    { id: 'tr3-2', name: 'Score', text: 'Passing score: {{passing_score}}%' },
  ]},
];

const TEMPLATES: TemplateInfo[] = [
  {
    id: 'annual-report', name: 'Annual Report', description: 'Professional annual business report',
    icon: 'BarChart3', slides: 4, variables: 9,
    gradient: 'from-emerald-500/20 to-teal-500/10', accentColor: 'text-emerald-400',
    accentBorder: 'border-emerald-500/30', accentBg: 'bg-emerald-500/10',
    difficulty: 'Intermediate', isPopular: true, slideData: DEMO_SLIDES, fileName: 'demo-template.pptx',
  },
  {
    id: 'marketing-plan', name: 'Marketing Plan', description: 'Strategic marketing campaign template',
    icon: 'Target', slides: 3, variables: 6,
    gradient: 'from-violet-500/20 to-purple-500/10', accentColor: 'text-violet-400',
    accentBorder: 'border-violet-500/30', accentBg: 'bg-violet-500/10',
    difficulty: 'Simple', isPopular: false, slideData: MARKETING_SLIDES, fileName: 'marketing-plan.pptx',
  },
  {
    id: 'project-proposal', name: 'Project Proposal', description: 'Business project pitch template',
    icon: 'Rocket', slides: 4, variables: 7,
    gradient: 'from-amber-500/20 to-orange-500/10', accentColor: 'text-amber-400',
    accentBorder: 'border-amber-500/30', accentBg: 'bg-amber-500/10',
    difficulty: 'Advanced', isPopular: false, slideData: PROPOSAL_SLIDES, fileName: 'project-proposal.pptx',
  },
  {
    id: 'team-update', name: 'Team Update', description: 'Weekly team status report',
    icon: 'Users', slides: 3, variables: 5,
    gradient: 'from-cyan-500/20 to-teal-500/10', accentColor: 'text-cyan-400',
    accentBorder: 'border-cyan-500/30', accentBg: 'bg-cyan-500/10',
    difficulty: 'Simple', isPopular: false, slideData: TEAM_SLIDES, fileName: 'team-update.pptx',
  },
  {
    id: 'product-launch', name: 'Product Launch', description: 'New product launch announcement',
    icon: 'Rocket', slides: 3, variables: 6,
    gradient: 'from-rose-500/20 to-pink-500/10', accentColor: 'text-rose-400',
    accentBorder: 'border-rose-500/30', accentBg: 'bg-rose-500/10',
    difficulty: 'Intermediate', isPopular: true, slideData: PRODUCT_SLIDES, fileName: 'product-launch.pptx',
  },
  {
    id: 'quarterly-review', name: 'Quarterly Review', description: 'Department quarterly performance review',
    icon: 'BarChart3', slides: 4, variables: 8,
    gradient: 'from-emerald-500/20 to-cyan-500/10', accentColor: 'text-teal-400',
    accentBorder: 'border-teal-500/30', accentBg: 'bg-teal-500/10',
    difficulty: 'Advanced', isPopular: false, slideData: QUARTERLY_SLIDES, fileName: 'quarterly-review.pptx',
  },
  {
    id: 'training-module', name: 'Training Module', description: 'Course training and assessment template',
    icon: 'GraduationCap', slides: 3, variables: 5,
    gradient: 'from-amber-500/20 to-yellow-500/10', accentColor: 'text-amber-400',
    accentBorder: 'border-amber-500/30', accentBg: 'bg-amber-500/10',
    difficulty: 'Simple', isPopular: false, slideData: TRAINING_SLIDES, fileName: 'training-module.pptx',
  },
];



// ── FAQ Data ──
const FAQ_ITEMS = [
  {
    question: 'What file formats are supported?',
    answer: 'PPTX Editor works with .pptx files (PowerPoint 2007+). You can upload existing presentations, generate new ones with AI, or import from JSON template definitions. All exports are in .pptx format with full formatting preservation.',
  },
  {
    question: 'How does the AI template generation work?',
    answer: 'Simply describe the presentation you need in plain English. Our AI analyzes your description, creates an appropriate slide structure, and identifies key variables that you can customize. It then generates professional values for all variables automatically.',
  },
  {
    question: 'Is my data secure?',
    answer: 'Absolutely. File processing happens on the server and your data is not shared. The AI generation uses a secure API call, and your file contents and variable data are kept safe.',
  },
  {
    question: 'Can I use my own existing PowerPoint templates?',
    answer: 'Yes! Upload any .pptx file and the editor will automatically detect all template variables (marked with {{variable_name}} syntax). You can then fill them in manually or use AI to generate appropriate values instantly.',
  },
  {
    question: 'What happens to formatting when I export?',
    answer: 'We preserve 100% of the original formatting, images, charts, and styles. For uploaded files, we modify only the variable text while keeping everything else intact. For AI-generated templates, we create clean, professional presentations.',
  },
];

// ============================================================================
// Helper: Convert template data to PptxSlideData[]
// ============================================================================

function templateToSlides(template: TemplateInfo): PptxSlideData[] {
  const EMU_PER_INCH = 914400;
  const SLIDE_W = 12192000; // 10 inches in EMU
  const SLIDE_H = 6858000;  // 7.5 inches in EMU

  return template.slideData.map((slide, slideIdx) => {
    const elCount = slide.elements.length;
    const elements: PptxTextElement[] = slide.elements.map((el, elIdx) => {
      // Distribute elements vertically across the slide
      const topMargin = EMU_PER_INCH * 0.8;
      const usableHeight = SLIDE_H - topMargin * 2;
      const elHeight = Math.min(EMU_PER_INCH * 0.8, usableHeight / elCount);
      const yPos = topMargin + elIdx * (elHeight + EMU_PER_INCH * 0.15);

      return {
        type: 'text' as const,
        id: el.id,
        shapeName: el.name,
        originalText: el.text,
        paragraphs: [{
          originalText: el.text,
          runs: [{
            originalText: el.text,
            bold: slideIdx === 0 && elIdx === 0,
            italic: false,
            fontSize: slideIdx === 0 && elIdx === 0 ? 28 : 14,
            fontColor: null,
            fontName: null,
          }],
        }],
        position: {
          x: EMU_PER_INCH * 0.8,
          y: yPos,
          width: SLIDE_W - EMU_PER_INCH * 1.6,
          height: elHeight,
        },
        slideIndex: slideIdx,
        elementIndex: elIdx,
      };
    });

    return {
      slideNumber: slide.slideNumber,
      elements,
      previewImage: null,
    };
  });
}

// ============================================================================
// Helper Components
// ============================================================================

// ── Typing Animation Component ──
function TypingText({ texts, className }: { texts: string[]; className?: string }) {
  const [textIndex, setTextIndex] = useState(0);
  const [charIndex, setCharIndex] = useState(0);
  const [isDeleting, setIsDeleting] = useState(false);

  useEffect(() => {
    const currentText = texts[textIndex];
    const timeout = isDeleting ? 40 : 80;

    if (!isDeleting && charIndex === currentText.length) {
      const pause = setTimeout(() => setIsDeleting(true), 2000);
      return () => clearTimeout(pause);
    }

    if (isDeleting && charIndex === 0) {
      const switchTimer = setTimeout(() => {
        setIsDeleting(false);
        setTextIndex((prev) => (prev + 1) % texts.length);
      }, 0);
      return () => clearTimeout(switchTimer);
    }

    const timer = setTimeout(() => {
      setCharIndex((prev) => prev + (isDeleting ? -1 : 1));
    }, timeout);

    return () => clearTimeout(timer);
  }, [charIndex, isDeleting, textIndex, texts]);

  const displayText = texts[textIndex].slice(0, charIndex);

  return (
    <span className={className}>
      {displayText}
      <motion.span
        animate={{ opacity: [1, 0] }}
        transition={{ duration: 0.5, repeat: Infinity, repeatType: 'reverse' }}
        className="inline-block w-[2px] h-[1em] bg-emerald-400 ml-0.5 align-middle"
      />
    </span>
  );
}

// ── Editor Mockup Component (CSS-only product visual) ──
function EditorMockup() {
  const ref = useRef<HTMLDivElement>(null);
  const mouseX = useMotionValue(0.5);
  const mouseY = useMotionValue(0.5);
  const tiltX = useSpring(useTransform(mouseY, [0, 1], [4, -4]), { stiffness: 150, damping: 20 });
  const tiltY = useSpring(useTransform(mouseX, [0, 1], [-4, 4]), { stiffness: 150, damping: 20 });

  const handleMouseMove = useCallback((e: React.MouseEvent) => {
    if (!ref.current) return;
    const rect = ref.current.getBoundingClientRect();
    mouseX.set((e.clientX - rect.left) / rect.width);
    mouseY.set((e.clientY - rect.top) / rect.height);
  }, [mouseX, mouseY]);

  const handleMouseLeave = useCallback(() => {
    mouseX.set(0.5);
    mouseY.set(0.5);
  }, [mouseX, mouseY]);

  const variableChips: Array<{ text: string; x: string; y: string; delay: number }> = [];

  return (
    <motion.div
      ref={ref}
      initial={{ opacity: 0, y: 40, scale: 0.95 }}
      animate={{ opacity: 1, y: 0, scale: 1 }}
      transition={{ duration: 0.9, delay: 0.3, ease: [0.22, 1, 0.36, 1] }}
      onMouseMove={handleMouseMove}
      onMouseLeave={handleMouseLeave}
      style={{ rotateX: tiltX, rotateY: tiltY, transformPerspective: 1200 }}
      className="relative"
    >
      {/* Floating variable chips */}
      {variableChips.map((chip) => (
        <motion.div
          key={chip.text}
          className="absolute z-10 hidden lg:flex items-center gap-1 px-2.5 py-1 rounded-full bg-emerald-500/15 dark:bg-emerald-400/20 border border-emerald-500/30 dark:border-emerald-400/25 text-[10px] font-mono font-medium text-emerald-600 dark:text-emerald-400 whitespace-nowrap backdrop-blur-sm"
          style={{ left: chip.x, top: chip.y }}
          animate={{ opacity: [0, 1, 1, 0], y: [4, 0, 0, -4] }}
          transition={{ duration: 4, delay: chip.delay, repeat: Infinity, repeatDelay: 2 }}
        >
          <span className="inline-block w-1.5 h-1.5 rounded-full bg-emerald-400 animate-pulse" />
          {chip.text}
        </motion.div>
      ))}

      {/* Animated gradient border */}
      <div className="absolute -inset-[2px] rounded-2xl overflow-hidden">
        <motion.div
          className="absolute inset-0"
          style={{
            background: 'conic-gradient(from 0deg, rgba(16,185,129,0.6), rgba(20,184,166,0.3), rgba(6,182,212,0.6), rgba(16,185,129,0.3), rgba(20,184,166,0.6))',
          }}
          animate={{ rotate: 360 }}
          transition={{ duration: 8, repeat: Infinity, ease: 'linear' }}
        />
      </div>

      {/* Main mockup window */}
      <div className="relative rounded-2xl overflow-hidden bg-white dark:bg-slate-900 shadow-2xl shadow-emerald-500/10 border border-slate-200/80 dark:border-slate-700/60">
        {/* Window chrome bar */}
        <div className="flex items-center gap-2 px-4 py-2.5 bg-slate-100 dark:bg-slate-800 border-b border-slate-200/60 dark:border-slate-700/40">
          <div className="flex gap-1.5">
            <div className="w-2.5 h-2.5 rounded-full bg-red-400/80" />
            <div className="w-2.5 h-2.5 rounded-full bg-amber-400/80" />
            <div className="w-2.5 h-2.5 rounded-full bg-emerald-400/80" />
          </div>
          <div className="flex-1 flex justify-center">
            <div className="px-8 py-0.5 rounded-md bg-slate-200/60 dark:bg-slate-700/60 text-[9px] text-slate-400 font-medium">
              template.pptx — PPTX Editor
            </div>
          </div>
        </div>

        {/* Editor layout */}
        <div className="flex h-48 sm:h-56 md:h-64">
          {/* Left sidebar - slide thumbnails */}
          <div className="w-14 sm:w-16 flex-shrink-0 bg-slate-50 dark:bg-slate-800/50 border-r border-slate-200/40 dark:border-slate-700/30 py-2 px-1.5 space-y-1.5">
            {[
              { bg: 'from-emerald-400 to-teal-400', active: true },
              { bg: 'from-slate-200 to-slate-300 dark:from-slate-600 dark:to-slate-700', active: false },
              { bg: 'from-slate-200 to-slate-300 dark:from-slate-600 dark:to-slate-700', active: false },
              { bg: 'from-slate-200 to-slate-300 dark:from-slate-600 dark:to-slate-700', active: false },
            ].map((slide, i) => (
              <div
                key={i}
                className={cn(
                  'aspect-[16/10] rounded-sm overflow-hidden',
                  slide.active ? 'ring-1 ring-emerald-400 shadow-sm' : 'opacity-60'
                )}
              >
                <div className={cn('w-full h-full bg-gradient-to-br', slide.bg, 'flex items-center justify-center')}>
                  <span className="text-[7px] font-bold text-white/80">{i + 1}</span>
                </div>
              </div>
            ))}
          </div>

          {/* Main slide area */}
          <div className="flex-1 bg-white dark:bg-slate-850 p-3 sm:p-4 flex flex-col items-center justify-center relative">
            {/* Slide placeholder */}
            <div className="w-full max-w-[200px] sm:max-w-[260px] aspect-[16/10] rounded-lg bg-gradient-to-br from-emerald-50 to-teal-50 dark:from-emerald-900/20 dark:to-teal-900/10 border border-emerald-200/40 dark:border-emerald-700/20 p-2 sm:p-3 flex flex-col gap-1.5 sm:gap-2">
              {/* Title line */}
              <div className="h-2 sm:h-2.5 w-3/4 rounded-full bg-emerald-400/30 dark:bg-emerald-400/20" />
              {/* Subtitle line */}
              <div className="h-1.5 sm:h-2 w-1/2 rounded-full bg-slate-300/50 dark:bg-slate-600/40" />
              {/* Body lines */}
              <div className="h-1 sm:h-1.5 w-full rounded-full bg-slate-200/60 dark:bg-slate-700/40" />
              <div className="h-1 sm:h-1.5 w-5/6 rounded-full bg-slate-200/60 dark:bg-slate-700/40" />
              {/* Chart shape */}
              <div className="flex-1 flex items-end gap-0.5 pt-1">
                <motion.div
                  className="flex-1 bg-emerald-400/25 dark:bg-emerald-400/15 rounded-t-sm"
                  animate={{ height: ['30%', '60%', '30%'] }}
                  transition={{ duration: 3, repeat: Infinity, ease: 'easeInOut' }}
                />
                <motion.div
                  className="flex-1 bg-teal-400/25 dark:bg-teal-400/15 rounded-t-sm"
                  animate={{ height: ['50%', '40%', '50%'] }}
                  transition={{ duration: 3, repeat: Infinity, ease: 'easeInOut', delay: 0.3 }}
                />
                <motion.div
                  className="flex-1 bg-cyan-400/25 dark:bg-cyan-400/15 rounded-t-sm"
                  animate={{ height: ['40%', '70%', '40%'] }}
                  transition={{ duration: 3, repeat: Infinity, ease: 'easeInOut', delay: 0.6 }}
                />
              </div>
            </div>

            {/* Element overlay indicator */}
            <motion.div
              className="absolute top-4 right-4 sm:top-6 sm:right-6 px-2 py-0.5 rounded-md bg-emerald-400/20 border border-emerald-400/30"
              animate={{ opacity: [0.5, 1, 0.5] }}
              transition={{ duration: 2, repeat: Infinity }}
            >
              <span className="text-[8px] font-semibold text-emerald-600 dark:text-emerald-400">3 elements</span>
            </motion.div>
          </div>

          {/* Right panel - form elements */}
          <div className="w-20 sm:w-28 flex-shrink-0 bg-slate-50 dark:bg-slate-800/50 border-l border-slate-200/40 dark:border-slate-700/30 p-2 space-y-2">
            {/* Section label */}
            <div className="text-[7px] font-semibold text-slate-400 uppercase tracking-wider">Variables</div>
            {/* Form fields */}
            {[
              { label: 'Title', value: 'Annual Report', filled: true },
              { label: 'Year', value: '{{fiscal_year}}', filled: false },
              { label: 'Name', value: '{{presenter}}', filled: false },
            ].map((field, i) => (
              <div key={i} className="space-y-0.5">
                <div className="text-[7px] text-slate-400 font-medium">{field.label}</div>
                <motion.div
                  className={cn(
                    'h-4 sm:h-5 rounded px-1.5 flex items-center text-[7px] sm:text-[8px] border',
                    field.filled
                      ? 'bg-emerald-50 dark:bg-emerald-900/20 border-emerald-300/40 dark:border-emerald-700/30 text-emerald-600 dark:text-emerald-400'
                      : 'bg-white dark:bg-slate-700/40 border-slate-200/60 dark:border-slate-600/30 text-slate-300 dark:text-slate-500 font-mono'
                  )}
                  animate={field.filled ? { borderColor: ['rgba(16,185,129,0.4)', 'rgba(16,185,129,0.8)', 'rgba(16,185,129,0.4)'] } : {}}
                  transition={{ duration: 2, repeat: Infinity, delay: i * 0.5 }}
                >
                  {field.value}
                </motion.div>
              </div>
            ))}
            {/* AI button */}
            <div className="pt-1">
              <motion.div
                className="h-5 sm:h-6 rounded-md bg-gradient-to-r from-emerald-400 to-teal-400 flex items-center justify-center gap-0.5"
                animate={{ opacity: [0.7, 1, 0.7] }}
                transition={{ duration: 2, repeat: Infinity }}
              >
                <Sparkles className="w-2 h-2 sm:w-2.5 sm:h-2.5 text-white" />
                <span className="text-[7px] sm:text-[8px] font-semibold text-white">AI Fill</span>
              </motion.div>
            </div>
          </div>
        </div>
      </div>
    </motion.div>
  );
}

// ── Floating Particles Component ──
function seededRandom(seed: number) {
  const x = Math.sin(seed * 9301 + 49297) * 49297;
  return x - Math.floor(x);
}

const emptySubscribe = () => () => {};

function FloatingParticles() {
  const mounted = useSyncExternalStore(emptySubscribe, () => true, () => false);

  const particles = useMemo(() =>
    Array.from({ length: 30 }, (_, i) => ({
      id: i,
      x: seededRandom(i * 3 + 1) * 100,
      delay: seededRandom(i * 3 + 2) * 10,
      duration: 8 + seededRandom(i * 3 + 3) * 12,
      size: 1 + seededRandom(i * 5 + 7) * 2.5,
      opacity: 0.1 + seededRandom(i * 5 + 11) * 0.3,
    })),
  []);

  if (!mounted) return null;

  return (
    <div className="pointer-events-none absolute inset-0 overflow-hidden">
      {particles.map((p) => (
        <motion.div
          key={p.id}
          className="absolute rounded-full bg-emerald-400"
          style={{
            width: p.size,
            height: p.size,
            left: `${p.x}%`,
            bottom: '-5%',
            opacity: p.opacity,
          }}
          animate={{
            y: [0, -1200],
            opacity: [p.opacity, 0],
          }}
          transition={{
            duration: p.duration,
            delay: p.delay,
            repeat: Infinity,
            ease: 'linear',
          }}
        />
      ))}
    </div>
  );
}

// ── Count Up Animation ──
function CountUp({ end, duration = 2, suffix = '' }: { end: number; duration?: number; suffix?: string }) {
  const [count, setCount] = useState(end);
  const ref = useRef<HTMLSpanElement>(null);
  const hasAnimated = useRef(false);

  useEffect(() => {
    if (hasAnimated.current) return;
    const observer = new IntersectionObserver(
      ([entry]) => {
        if (entry.isIntersecting && !hasAnimated.current) {
          hasAnimated.current = true;
          setCount(0);
          let start = 0;
          const step = end / (duration * 60);
          const interval = setInterval(() => {
            start += step;
            if (start >= end) {
              setCount(end);
              clearInterval(interval);
            } else {
              setCount(Math.floor(start));
            }
          }, 1000 / 60);
        }
      },
      { threshold: 0.3 }
    );
    if (ref.current) observer.observe(ref.current);
    return () => observer.disconnect();
  }, [end, duration]);

  return (
    <span ref={ref}>
      {count}{suffix}
    </span>
  );
}

// ── Upload Zone Component ──
function UploadZone({
  dragOver,
  setDragOver,
  onDrop,
  onFileSelect,
}: {
  dragOver: boolean;
  setDragOver: (v: boolean) => void;
  onDrop: (e: React.DragEvent) => void;
  onFileSelect: (file: File) => void;
}) {
  const inputRef = useRef<HTMLInputElement>(null);

  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setDragOver(true); }}
      onDragLeave={() => setDragOver(false)}
      onDrop={onDrop}
      onClick={() => inputRef.current?.click()}
      className={cn(
        'flex flex-col items-center justify-center rounded-xl border-2 border-dashed p-8 transition-all duration-300 cursor-pointer',
        dragOver
          ? 'border-emerald-400 bg-emerald-500/10 scale-[1.02]'
          : 'border-slate-300 dark:border-white/20 hover:border-emerald-400 dark:hover:border-white/40 hover:bg-slate-50 dark:hover:bg-white/5',
      )}
    >
      <Upload className={cn('h-10 w-10 mb-3', dragOver ? 'text-emerald-400' : 'text-slate-500')} />
      <p className="text-sm font-medium text-slate-700 dark:text-slate-300">
        {dragOver ? 'Drop your file here' : 'Drag & drop a .pptx file'}
      </p>
      <p className="text-xs text-slate-500 mt-1">or click to browse</p>
      <input
        ref={inputRef}
        type="file"
        accept=".pptx"
        className="hidden"
        onChange={(e) => {
          const file = e.target.files?.[0];
          if (file) onFileSelect(file);
        }}
      />
    </div>
  );
}

// ── Magnetic GlassCard Component ──
interface GlassCardProps {
  icon: React.ReactNode;
  title: string;
  description: string;
  onClick: () => void;
  active?: boolean;
  gradient?: string;
  delay?: number;
  badge?: string;
  stepNumber?: number;
  iconAnimation?: 'rotate' | 'pulse' | 'bounce';
}

function GlassCard({ icon, title, description, onClick, active, gradient, delay = 0, badge, stepNumber, iconAnimation }: GlassCardProps) {
  const ref = useRef<HTMLDivElement>(null);
  const x = useMotionValue(0);
  const y = useMotionValue(0);

  const rotateX = useSpring(useTransform(y, [-0.5, 0.5], [6, -6]), { stiffness: 200, damping: 20 });
  const rotateY = useSpring(useTransform(x, [-0.5, 0.5], [-6, 6]), { stiffness: 200, damping: 20 });

  const handleMouseMove = useCallback(
    (e: React.MouseEvent) => {
      if (!ref.current) return;
      const rect = ref.current.getBoundingClientRect();
      const px = (e.clientX - rect.left) / rect.width - 0.5;
      const py = (e.clientY - rect.top) / rect.height - 0.5;
      x.set(px);
      y.set(py);
    },
    [x, y]
  );

  const handleMouseLeave = useCallback(() => {
    x.set(0);
    y.set(0);
  }, [x, y]);

  const iconAnimationClass = iconAnimation === 'rotate'
    ? 'group-hover:rotate-12'
    : iconAnimation === 'pulse'
    ? 'group-hover:animate-pulse'
    : iconAnimation === 'bounce'
    ? 'group-hover:animate-bounce'
    : '';

  return (
    <motion.div
      ref={ref}
      initial={{ opacity: 0, y: 40 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true }}
      transition={{ duration: 0.6, delay, ease: [0.22, 1, 0.36, 1] }}
      whileHover={{ scale: 1.03 }}
      whileTap={{ scale: 0.97 }}
      onClick={onClick}
      onMouseMove={handleMouseMove}
      onMouseLeave={handleMouseLeave}
      style={{ rotateX, rotateY, transformPerspective: 1000 }}
      className={cn(
        'group relative cursor-pointer rounded-2xl p-6 sm:p-8 transition-all duration-300',
        'border border-slate-200/50 dark:border-white/20',
        'shadow-lg shadow-black/5 hover:shadow-2xl hover:shadow-black/10',
        gradient || 'bg-white/80 dark:bg-white/10 backdrop-blur-xl',
        active && 'ring-2 ring-emerald-400/60 bg-white/90 dark:bg-white/20',
        !active && 'hover:bg-white/90 dark:hover:bg-white/15'
      )}
    >
      {/* Animated gradient border on hover */}
      <div className="pointer-events-none absolute inset-0 rounded-2xl opacity-0 group-hover:opacity-100 transition-opacity duration-500"
        style={{
          background: 'linear-gradient(135deg, rgba(16,185,129,0.4), rgba(20,184,166,0.2), rgba(6,182,212,0.4))',
          padding: '1px',
          WebkitMask: 'linear-gradient(#fff 0 0) content-box, linear-gradient(#fff 0 0)',
          WebkitMaskComposite: 'xor',
          maskComposite: 'exclude',
        }}
      />
      <div className="pointer-events-none absolute inset-0 rounded-2xl bg-gradient-to-br from-slate-200/30 dark:from-white/10 via-transparent to-transparent opacity-0 transition-opacity duration-300 group-hover:opacity-100" />
      <div className="pointer-events-none absolute inset-0 rounded-2xl opacity-0 group-hover:opacity-100 transition-opacity duration-700 bg-gradient-to-br from-emerald-500/5 via-transparent to-teal-500/5" />

      {/* Progress indicator line at bottom */}
      <div className="absolute bottom-0 left-0 right-0 h-[2px] rounded-b-2xl overflow-hidden">
        <motion.div
          className="h-full bg-gradient-to-r from-emerald-400 via-teal-400 to-cyan-400"
          initial={{ width: '0%' }}
          whileHover={{ width: '100%' }}
          transition={{ duration: 0.4, ease: 'easeOut' }}
        />
      </div>

      {/* Step number indicator */}
      {stepNumber && (
        <div className="absolute -top-3 -left-3 z-20 flex h-7 w-7 items-center justify-center rounded-full bg-gradient-to-br from-emerald-400 to-teal-500 text-xs font-bold text-white shadow-lg shadow-emerald-500/30">
          {stepNumber}
        </div>
      )}

      <div className="relative z-10">
        <div className="mb-5 flex items-center gap-3">
          <div className={cn(
            'inline-flex h-12 w-12 sm:h-14 sm:w-14 items-center justify-center rounded-xl bg-gradient-to-br from-slate-100 dark:from-white/20 to-slate-50 dark:to-white/5 backdrop-blur-sm transition-transform duration-300',
            iconAnimationClass
          )}>
            {icon}
          </div>
          {badge && (
            <span className="rounded-full bg-emerald-500/20 border border-emerald-500/30 px-2.5 py-0.5 text-[10px] font-semibold text-emerald-400 uppercase tracking-wider">
              {badge}
            </span>
          )}
        </div>
        <h3 className="mb-2 text-lg sm:text-xl font-bold text-slate-900 dark:text-white tracking-tight">{title}</h3>
        <p className="text-sm leading-relaxed text-slate-500 dark:text-slate-400">{description}</p>
        <div className="mt-5 flex items-center text-sm font-medium text-emerald-400 transition-all duration-300 group-hover:gap-2.5 gap-1.5">
          Get started <ArrowRight className="h-4 w-4 transition-transform duration-300 group-hover:translate-x-1" />
        </div>
      </div>
    </motion.div>
  );
}



// ── How It Works Steps ──
function HowItWorks() {
  const [activeStep, setActiveStep] = useState(0);
  const steps = [
    { number: '1', label: 'Upload', description: 'Import your PPTX or pick a template', icon: <Upload className="h-6 w-6" /> },
    { number: '2', label: 'Edit', description: 'Fill variables with AI or manually', icon: <Sparkles className="h-6 w-6" /> },
    { number: '3', label: 'Export', description: 'Download in PPTX format', icon: <FileSpreadsheet className="h-6 w-6" /> },
  ];

  useEffect(() => {
    const interval = setInterval(() => {
      setActiveStep((prev) => (prev + 1) % steps.length);
    }, 3000);
    return () => clearInterval(interval);
  }, [steps.length]);

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      whileInView={{ opacity: 1, y: 0 }}
      viewport={{ once: true }}
      transition={{ delay: 0.6, duration: 0.6 }}
      className="mt-12 sm:mt-16"
    >
      <div className="flex flex-col sm:flex-row items-center justify-center gap-6 sm:gap-0">
        {steps.map((step, i) => (
          <React.Fragment key={step.number}>
            <motion.div
              className="flex flex-col items-center gap-3"
              animate={{ scale: activeStep === i ? 1.08 : 1 }}
              transition={{ duration: 0.4, ease: 'easeOut' }}
            >
              <motion.div
                className={cn(
                  'relative flex h-16 w-16 sm:h-18 sm:w-18 items-center justify-center rounded-full border-2 transition-all duration-500',
                  activeStep === i
                    ? 'bg-gradient-to-br from-emerald-400/30 to-teal-500/30 border-emerald-400/70 text-emerald-300'
                    : 'bg-gradient-to-br from-emerald-400/10 to-teal-500/10 border-emerald-500/20 text-emerald-400/50'
                )}
                animate={activeStep === i ? { boxShadow: ['0 0 20px rgba(16,185,129,0.2)', '0 0 35px rgba(16,185,129,0.4)', '0 0 20px rgba(16,185,129,0.2)'] } : {}}
                transition={activeStep === i ? { duration: 2, repeat: Infinity, ease: 'easeInOut' } : {}}
              >
                {step.icon}
                <span className={cn(
                  'absolute -top-2 -right-2 flex h-6 w-6 items-center justify-center rounded-full text-[11px] font-bold text-white transition-all duration-300',
                  activeStep === i
                    ? 'bg-emerald-400 shadow-lg shadow-emerald-400/50'
                    : 'bg-emerald-600/60'
                )}>
                  {step.number}
                </span>
              </motion.div>
              <div className="text-center max-w-[150px]">
                <p className={cn('text-sm font-semibold transition-colors duration-300', activeStep === i ? 'text-slate-900 dark:text-white' : 'text-slate-500 dark:text-slate-400')}>{step.label}</p>
                <p className={cn('text-[11px] mt-0.5 transition-colors duration-300', activeStep === i ? 'text-slate-600 dark:text-slate-300' : 'text-slate-400 dark:text-slate-600')}>{step.description}</p>
              </div>
            </motion.div>
            {i < steps.length - 1 && (
              <div className="flex-shrink-0 flex items-center sm:mx-4 my-2 sm:my-0">
                <div className="relative w-12 sm:w-24 h-8 flex items-center justify-center">
                  <svg className="absolute w-full h-4" viewBox="0 0 80 16">
                    <line x1="0" y1="8" x2="80" y2="8" stroke="rgba(16,185,129,0.3)" strokeWidth="2" strokeDasharray="6 4">
                      <animate attributeName="stroke-dashoffset" from="0" to="-20" dur="1s" repeatCount="indefinite" />
                    </line>
                  </svg>
                  <motion.div
                    animate={{ x: [0, 4, 0] }}
                    transition={{ duration: 1.5, repeat: Infinity, ease: 'easeInOut' }}
                  >
                    <ArrowRight className="h-4 w-4 text-emerald-500/50 absolute right-0 top-1/2 -translate-y-1/2" />
                  </motion.div>
                </div>
              </div>
            )}
          </React.Fragment>
        ))}
      </div>
    </motion.div>
  );
}



// ============================================================================
// Main HomePage Component
// ============================================================================

export default function HomePage() {
  const { step, setStep, setParsedData, slides, currentSlideIndex, setCurrentSlide, getTotalModificationCount, fileName, reset } = usePptxStore();

  const [dragOver, setDragOver] = useState(false);
  const [uploadProgress, setUploadProgress] = useState(0);
  const [uploadError, setUploadError] = useState<string | null>(null);
  const [scrollProgress, setScrollProgress] = useState(0);
  const [headerVisible, setHeaderVisible] = useState(true);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
  const [aiDialogOpen, setAiDialogOpen] = useState(false);
  const [aiSettingsOpen, setAiSettingsOpen] = useState(false);
  const [kbShortcutsOpen, setKbShortcutsOpen] = useState(false);
  const [presentationMode, setPresentationMode] = useState(false);
  const [previewZoom, setPreviewZoom] = useState(1);
  const [showGridOverlay, setShowGridOverlay] = useState(false);
  const [jsonInput, setJsonInput] = useState('');
  const [fileHistory, setFileHistory] = useState<FileHistoryEntry[]>([]);
  const lastScrollY = useRef(0);
  const [slideTransition, setSlideTransition] = useState<number | null>(null);
  const prevSlideIndexRef = useRef(currentSlideIndex);
  const isInitialMountRef = useRef(true);

  // Load file history
  useEffect(() => {
    setFileHistory(getFileHistory());
  }, [step]);

  // Slide transition indicator
  useEffect(() => {
    if (isInitialMountRef.current) {
      isInitialMountRef.current = false;
      prevSlideIndexRef.current = currentSlideIndex;
      return;
    }
    if (prevSlideIndexRef.current !== currentSlideIndex) {
      prevSlideIndexRef.current = currentSlideIndex;
      setSlideTransition(currentSlideIndex + 1);
      const timer = setTimeout(() => setSlideTransition(null), 600);
      return () => clearTimeout(timer);
    }
  }, [currentSlideIndex]);

  // ── Global keyboard shortcuts ──────────────────────────────────────────
  useEffect(() => {
    if (step !== 'editing') return;
    const handler = (e: KeyboardEvent) => {
      // Ctrl+/ to show keyboard shortcuts
      if ((e.ctrlKey || e.metaKey) && e.key === '/') {
        e.preventDefault();
        setKbShortcutsOpen(true);
      }
    };
    window.addEventListener('keydown', handler);
    return () => window.removeEventListener('keydown', handler);
  }, [step]);

  // Scroll progress tracking & header visibility
  useEffect(() => {
    if (step !== 'upload') return;
    const handleScroll = () => {
      const scrollTop = window.scrollY;
      const docHeight = document.documentElement.scrollHeight - window.innerHeight;
      const prog = docHeight > 0 ? (scrollTop / docHeight) * 100 : 0;
      setScrollProgress(Math.min(100, Math.max(0, prog)));

      if (scrollTop < 100) {
        setHeaderVisible(true);
      } else if (scrollTop > lastScrollY.current + 5) {
        setHeaderVisible(false);
      } else if (scrollTop < lastScrollY.current - 5) {
        setHeaderVisible(true);
      }
      lastScrollY.current = scrollTop;
    };
    window.addEventListener('scroll', handleScroll, { passive: true });
    return () => window.removeEventListener('scroll', handleScroll);
  }, [step]);

  // ── File Upload Handler (uses /api/pptx/parse) ──
  const handleFileUpload = useCallback(
    async (file: File) => {
      if (!file.name.endsWith('.pptx')) {
        setUploadError('Please upload a .pptx file');
        toast.error('Invalid file type', { description: 'Please upload a .pptx file' });
        return;
      }

      setStep('loading');
      setUploadError(null);
      setUploadProgress(0);

      try {
        const formData = new FormData();
        formData.append('file', file);

        // Use XHR for progress tracking
        const result = await new Promise<{ fileId: string; fileName: string; slides: PptxSlideData[]; slideSize?: { width: number; height: number } }>((resolve, reject) => {
          const xhr = new XMLHttpRequest();
          xhr.open('POST', '/api/pptx/parse');

          xhr.upload.onprogress = (e) => {
            if (e.lengthComputable) {
              const pct = Math.round((e.loaded / e.total) * 70);
              setUploadProgress(pct);
            }
          };

          xhr.onload = () => {
            if (xhr.status === 200) {
              try {
                const data = JSON.parse(xhr.responseText);
                setUploadProgress(90);
                resolve(data);
              } catch {
                reject(new Error('Invalid response from server'));
              }
            } else {
              try {
                const errData = JSON.parse(xhr.responseText);
                reject(new Error(errData.error || 'Upload failed'));
              } catch {
                reject(new Error('Upload failed'));
              }
            }
          };

          xhr.onerror = () => reject(new Error('Network error'));
          xhr.send(formData);
        });

        setUploadProgress(100);

        // Save to file history
        addFileHistory({
          fileId: result.fileId,
          fileName: result.fileName,
          slideCount: result.slides.length,
          openedAt: Date.now(),
        });

        // Load parsed data into the store
        setParsedData(result.fileId, result.fileName, result.slides, result.slideSize);

        toast.success('File uploaded successfully', {
          description: `Found ${result.slides.length} slides with ${result.slides.reduce((sum, s) => sum + s.elements.length, 0)} elements`,
        });
      } catch (err) {
        setUploadError(err instanceof Error ? err.message : 'Failed to parse PPTX');
        setStep('upload');
        toast.error('Upload failed', { description: err instanceof Error ? err.message : 'Failed to parse PPTX' });
      }
    },
    [setStep, setParsedData]
  );

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      const file = e.dataTransfer.files[0];
      if (file) handleFileUpload(file);
    },
    [handleFileUpload]
  );

  // ── Template Load Handler ──
  const handleTemplateLoad = useCallback((template: TemplateInfo) => {
    const slides = templateToSlides(template);
    const fileId = `template-${template.id}-${Date.now()}`;
    setParsedData(fileId, template.fileName, slides);

    addFileHistory({
      fileId,
      fileName: template.fileName,
      slideCount: slides.length,
      openedAt: Date.now(),
    });

    toast.success(`${template.name} loaded`, {
      description: `Template with ${template.variables} variables across ${template.slides} slides`,
    });
  }, [setParsedData]);

  // ── JSON Import Handler ──
  const handleJsonImport = useCallback(() => {
    if (!jsonInput.trim()) return;
    try {
      const data = JSON.parse(jsonInput);
      if (data.slides && Array.isArray(data.slides)) {
        // Convert JSON slides to PptxSlideData format if they match
        const slides: PptxSlideData[] = data.slides.map((slide: { slideNumber?: number; elements?: unknown[] }, idx: number) => ({
          slideNumber: slide.slideNumber ?? idx + 1,
          elements: (slide.elements || []) as PptxSlideData['elements'],
          previewImage: null,
        }));
        const fileId = `json-import-${Date.now()}`;
        const fileName = data.fileName || 'imported.json';
        setParsedData(fileId, fileName, slides);
        addFileHistory({ fileId, fileName, slideCount: slides.length, openedAt: Date.now() });
        toast.success('JSON template imported');
      } else {
        toast.error('Invalid JSON structure');
      }
    } catch {
      toast.error('Invalid JSON format');
    }
  }, [jsonInput, setParsedData]);

  // ── Reopen from history ──
  const handleReopenFile = useCallback(async (entry: FileHistoryEntry) => {
    try {
      const res = await fetch(`/api/pptx/check?fileId=${entry.fileId}`);
      const data = await res.json();
      if (data.exists) {
        // Re-parse the file
        setStep('loading');
        setUploadProgress(20);
        const repRes = await fetch('/api/pptx/reparse', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ fileId: entry.fileId }),
        });
        setUploadProgress(60);
        const repData = await repRes.json();
        if (repData.slides) {
          setUploadProgress(100);
          setParsedData(entry.fileId, entry.fileName, repData.slides, repData.slideSize);
          addFileHistory({ fileId: entry.fileId, fileName: entry.fileName, slideCount: repData.slides.length, openedAt: Date.now() });
          toast.success('File reopened');
        } else {
          throw new Error(repData.error || 'Failed to reopen');
        }
      } else {
        toast.error('File no longer available on server');
        removeFileHistory(entry.fileId);
        setFileHistory(getFileHistory());
      }
    } catch (err) {
      toast.error('Failed to reopen file', { description: err instanceof Error ? err.message : 'Unknown error' });
      setStep('upload');
    }
  }, [setStep, setParsedData]);

  // ====================================================================
  // EDITING STATE
  // ====================================================================
  if (step === 'editing') {
    const currentSlide = slides[currentSlideIndex];
    return (
      <div className="flex flex-col h-screen bg-background overflow-hidden">
        <Toolbar
          onAiGenerate={() => setAiDialogOpen(true)}
          onAiSettings={() => setAiSettingsOpen(true)}
          onKeyboardShortcuts={() => setKbShortcutsOpen(true)}
          onPresent={() => setPresentationMode(true)}
          onZoomIn={() => setPreviewZoom(z => Math.min(3, +(z + 0.1).toFixed(1)))}
          onZoomOut={() => setPreviewZoom(z => Math.max(0.5, +(z - 0.1).toFixed(1)))}
          onZoomReset={() => setPreviewZoom(1)}
        />
        <div className="flex flex-1 min-h-0 overflow-hidden">
          {/* Left Column: Navigator + Preview */}
          <div className="flex flex-1 min-w-0 min-h-0 overflow-hidden">
            <SlideNavigator slides={slides} currentSlide={currentSlide} />
            {/* Slide Preview (main area) */}
            <div className="flex-1 min-w-0 flex flex-col min-h-0 overflow-hidden bg-muted/10">
              {/* Preview header - same height as editor header */}
              <div className="shrink-0 flex items-center gap-1.5 border-b border-border/20 px-3 h-[30px] bg-muted/10">
                <span className="text-[10px] font-semibold text-muted-foreground/70 uppercase tracking-widest">Preview</span>
                <Badge variant="secondary" className="h-3.5 px-1 text-[9px] font-medium bg-emerald-50/40 dark:bg-emerald-900/15 text-emerald-600 dark:text-emerald-400 border-emerald-200/30 dark:border-emerald-700/20">
                  {currentSlideIndex + 1} / {slides.length}
                </Badge>
                {/* Grid overlay toggle */}
                <button
                  onClick={() => setShowGridOverlay(v => !v)}
                  className={"h-5 w-5 flex items-center justify-center rounded transition-colors " + (showGridOverlay ? "bg-emerald-500/20 text-emerald-600 dark:text-emerald-400" : "bg-muted/30 text-muted-foreground hover:bg-muted/50 hover:text-muted-foreground")}
                  title="Toggle grid overlay"
                  aria-label="Toggle grid overlay"
                >
                  <Grid3x3 className="size-2.5" />
                </button>
                {/* Zoom controls */}
                <div className="ml-auto flex items-center gap-0.5">
                  {/* Zoom out */}
                  <button
                    onClick={() => setPreviewZoom(z => Math.max(0.5, +(z - 0.1).toFixed(1)))}
                    className="h-5 w-5 flex items-center justify-center rounded bg-muted/30 hover:bg-muted/50 text-muted-foreground/70 hover:text-muted-foreground transition-colors"
                    aria-label="Zoom out"
                  >
                    <Minus className="size-2.5" />
                  </button>
                  {/* Zoom slider */}
                  <input
                    type="range"
                    min={0.5}
                    max={3}
                    step={0.1}
                    value={previewZoom}
                    onChange={e => setPreviewZoom(+e.target.value)}
                    className="w-12 h-1 accent-emerald-500 cursor-pointer"
                    aria-label="Zoom level"
                  />
                  {/* Zoom in */}
                  <button
                    onClick={() => setPreviewZoom(z => Math.min(3, +(z + 0.1).toFixed(1)))}
                    className="h-5 w-5 flex items-center justify-center rounded bg-muted/30 hover:bg-muted/50 text-muted-foreground/70 hover:text-muted-foreground transition-colors"
                    aria-label="Zoom in"
                  >
                    <Plus className="size-2.5" />
                  </button>
                  {/* Zoom percentage */}
                  <span className="text-[8px] text-muted-foreground/60 min-w-[28px] text-center tabular-nums">
                    {Math.round(previewZoom * 100)}%
                  </span>
                  {/* Fit button (reset to 100%) */}
                  <button
                    onClick={() => setPreviewZoom(1)}
                    className={"h-5 w-5 flex items-center justify-center rounded bg-muted/30 hover:bg-muted/50 text-muted-foreground/70 hover:text-muted-foreground transition-colors" + (previewZoom === 1 ? " text-emerald-600 dark:text-emerald-400" : "")}
                    aria-label="Fit to view"
                    title="Fit to view (Ctrl+0)"
                  >
                    <Maximize2 className="size-2.5" />
                  </button>
                  {/* Actual size button */}
                  <button
                    onClick={() => setPreviewZoom(2)}
                    className={"h-5 px-1 flex items-center justify-center rounded bg-muted/30 hover:bg-muted/50 text-muted-foreground/70 hover:text-muted-foreground transition-colors" + (Math.round(previewZoom * 100) === 200 ? " text-emerald-600 dark:text-emerald-400" : "")}
                    aria-label="Actual size (200%)"
                    title="Actual size (200%)"
                  >
                    <span className="text-[8px] font-semibold">1:1</span>
                  </button>
                </div>
              </div>
              <div className={"flex-1 flex items-center justify-center p-3 relative" + (previewZoom > 1 ? " overflow-auto" : " overflow-hidden")}>
                {/* Subtle checkerboard pattern for transparency indication */}
                <div className="absolute inset-0 opacity-[0.02] dark:opacity-[0.03]" style={{
                  backgroundImage: 'linear-gradient(45deg, rgba(128,128,128,0.3) 25%, transparent 25%, transparent 75%, rgba(128,128,128,0.3) 75%), linear-gradient(45deg, rgba(128,128,128,0.3) 25%, transparent 25%, transparent 75%, rgba(128,128,128,0.3) 75%)',
                  backgroundSize: '20px 20px',
                  backgroundPosition: '0 0, 10px 10px',
                }} />
                {/* Slide transition indicator */}
                <AnimatePresence>
                  {slideTransition !== null && (
                    <motion.div
                      key={`slide-transition-${slideTransition}`}
                      initial={{ opacity: 0, scale: 0.9 }}
                      animate={{ opacity: 1, scale: 1 }}
                      exit={{ opacity: 0, scale: 1.05 }}
                      transition={{ duration: 0.2, ease: 'easeOut' }}
                      className="absolute inset-0 flex items-center justify-center z-20 pointer-events-none"
                    >
                      <span className="text-2xl font-bold text-foreground/30 select-none">
                        Slide {slideTransition}
                      </span>
                    </motion.div>
                  )}
                </AnimatePresence>
                {currentSlide ? (
                  <div
                    className="w-full max-w-3xl relative z-10 transition-transform duration-200"
                    style={{ transform: `scale(${previewZoom})`, transformOrigin: 'center center' }}
                  >
                    <SlidePreview slide={currentSlide} className="w-full" showGridOverlay={showGridOverlay} />
                  </div>
                ) : (
                  <div className="flex flex-col items-center justify-center text-muted-foreground/50 relative z-10">
                    <FileText className="size-12 mb-3 opacity-20" />
                    <span className="text-sm font-medium">No slide selected</span>
                    <span className="text-xs mt-1 opacity-60">Select a slide from the navigator</span>
                  </div>
                )}
              </div>
            </div>
          </div>
          {/* Right Column: Editor - responsive width */}
          <div className="w-full md:w-[320px] lg:w-[380px] shrink-0 border-l border-border/30 overflow-hidden">
            {currentSlide && <SlideEditor slide={currentSlide} />}
          </div>
        </div>
        {/* Status bar */}
        <div className="flex items-center justify-between px-2 sm:px-3 h-[26px] border-t border-border/20 bg-gradient-to-r from-muted/20 via-background to-muted/20 dark:from-muted/15 dark:via-muted/5 dark:to-muted/15 text-[10px] text-muted-foreground shrink-0">
          {/* Left: File info (hidden on very small screens, abbreviated on medium) */}
          <div className="hidden sm:flex items-center gap-2 min-w-0">
            <div className="flex items-center gap-1 min-w-0">
              <FileText className="size-2.5 shrink-0 text-emerald-600/70 dark:text-emerald-400/70" />
              <span className="truncate max-w-[120px] font-medium">{fileName || 'Untitled'}</span>
            </div>
            <Badge variant="secondary" className="h-3.5 px-1 text-[8px] font-bold bg-emerald-50/50 dark:bg-emerald-900/20 text-emerald-700 dark:text-emerald-400 border-emerald-200/30 dark:border-emerald-700/20 shrink-0 uppercase tracking-wider">
              .pptx
            </Badge>
            <span className="text-muted-foreground/50">|</span>
            <span className="hidden md:inline">{slides.length} slide{slides.length !== 1 ? 's' : ''} &middot; {slides.reduce((sum, s) => sum + s.elements.length, 0)} elements</span>
            <span className="md:hidden">{slides.length} slide{slides.length !== 1 ? 's' : ''}</span>
          </div>

          {/* Center: Slide navigation */}
          <div className="flex items-center gap-1.5">
            <button
              onClick={() => setCurrentSlide(Math.max(0, currentSlideIndex - 1))}
              disabled={currentSlideIndex <= 0}
              aria-label="Previous slide"
              className={cn(
                'inline-flex items-center justify-center size-5 rounded transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                currentSlideIndex > 0
                  ? 'text-muted-foreground hover:text-foreground hover:bg-accent/50 active:bg-accent/70 cursor-pointer'
                  : 'text-muted-foreground/25 cursor-not-allowed'
              )}
            >
              <ChevronLeft className="size-3" />
            </button>
            <span className="text-[10px] font-medium tabular-nums min-w-[60px] text-center">
              Slide {currentSlideIndex + 1} / {slides.length}
            </span>
            <button
              onClick={() => setCurrentSlide(Math.min(slides.length - 1, currentSlideIndex + 1))}
              disabled={currentSlideIndex >= slides.length - 1}
              aria-label="Next slide"
              className={cn(
                'inline-flex items-center justify-center size-5 rounded transition-all focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-emerald-500/30',
                currentSlideIndex < slides.length - 1
                  ? 'text-muted-foreground hover:text-foreground hover:bg-accent/50 active:bg-accent/70 cursor-pointer'
                  : 'text-muted-foreground/25 cursor-not-allowed'
              )}
            >
              <ChevronRight className="size-3" />
            </button>
          </div>

          {/* Right: Modification progress */}
          <div className="flex items-center gap-2">
            {(() => {
              const modCount = getTotalModificationCount();
              const totalElements = slides.reduce((sum, s) => sum + s.elements.length, 0);
              const progressPct = totalElements > 0 ? Math.round((modCount / totalElements) * 100) : 0;
              return (
                <>
                  <div className="hidden sm:flex items-center gap-1">
                    <Edit3 className="size-2.5 text-amber-500/70 dark:text-amber-400/70" />
                    <span>{modCount} of {totalElements} modified</span>
                  </div>
                  <TooltipProvider delayDuration={300}>
                    <Tooltip>
                      <TooltipTrigger asChild>
                        <div className="w-12 h-1 rounded-full bg-muted-foreground/10 overflow-hidden cursor-default">
                          <div
                            className={cn(
                              'h-full rounded-full transition-all duration-300',
                              progressPct > 0
                                ? 'bg-gradient-to-r from-amber-400 to-amber-500 dark:from-amber-400 dark:to-amber-500'
                                : 'bg-transparent'
                            )}
                            style={{ width: `${progressPct}%` }}
                          />
                        </div>
                      </TooltipTrigger>
                      <TooltipContent side="top">
                        <p>{modCount} of {totalElements} elements modified ({progressPct}%)</p>
                      </TooltipContent>
                    </Tooltip>
                  </TooltipProvider>
                </>
              );
            })()}
            <span className="text-muted-foreground/40 hidden md:inline">← → navigate</span>
          </div>
        </div>

        {/* AI Dialogs */}
        <AiGenerateDialog
          open={aiDialogOpen}
          onOpenChange={setAiDialogOpen}
          onOpenSettings={() => {
            setAiDialogOpen(false);
            setAiSettingsOpen(true);
          }}
        />
        <AiSettingsDialog
          open={aiSettingsOpen}
          onOpenChange={setAiSettingsOpen}
          onConfigChange={() => {}}
        />

        {/* Keyboard Shortcuts Dialog */}
        <KeyboardShortcutsDialog
          open={kbShortcutsOpen}
          onOpenChange={setKbShortcutsOpen}
        />

        {/* Presentation Mode Overlay */}
        {presentationMode && (
          <PresentationMode
            slides={slides}
            initialSlideIndex={currentSlideIndex}
            onExit={() => setPresentationMode(false)}
            onSlideChange={(index) => setCurrentSlide(index)}
          />
        )}
      </div>
    );
  }

  // ====================================================================
  // LOADING STATE
  // ====================================================================
  if (step === 'loading') {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-slate-50 via-white to-slate-50 dark:from-slate-950 dark:via-slate-900 dark:to-slate-950">
        <div className="flex flex-col items-center gap-6">
          <motion.div
            animate={{ rotate: 360 }}
            transition={{ duration: 2, repeat: Infinity, ease: 'linear' }}
            className="relative"
          >
            <div className="h-16 w-16 rounded-2xl bg-gradient-to-br from-emerald-400 to-teal-500 flex items-center justify-center shadow-xl shadow-emerald-500/20">
              <Presentation className="h-8 w-8 text-white" />
            </div>
          </motion.div>
          <div className="text-center">
            <h2 className="text-xl font-bold text-slate-900 dark:text-white mb-2">Processing your file</h2>
            <p className="text-sm text-slate-500 dark:text-slate-400">Parsing slides and extracting elements...</p>
          </div>
          <div className="w-64">
            <Progress value={uploadProgress} className="h-2" />
            <p className="text-xs text-slate-400 mt-2 text-center">{uploadProgress}%</p>
          </div>
        </div>
      </div>
    );
  }

  // ====================================================================
  // UPLOAD / LANDING PAGE STATE
  // ====================================================================
  return (
    <div className="min-h-screen flex flex-col bg-gradient-to-br from-slate-50 via-white to-slate-50 dark:from-slate-950 dark:via-slate-900 dark:to-slate-950 relative overflow-hidden">
      {/* Scroll Progress Indicator */}
      <div className="fixed top-0 left-0 right-0 z-50 h-[2px]">
        <motion.div
          className="h-full bg-gradient-to-r from-emerald-400 via-teal-400 to-cyan-400"
          style={{
            width: `${scrollProgress}%`,
          }}
        />
      </div>

      {/* Animated Background Orbs */}
      <div className="pointer-events-none fixed inset-0 overflow-hidden">
        <div className="absolute -top-1/4 -right-1/4 w-[600px] h-[600px] rounded-full bg-emerald-500/5 dark:bg-emerald-500/10 blur-3xl" />
        <div className="absolute -bottom-1/4 -left-1/4 w-[500px] h-[500px] rounded-full bg-teal-500/5 dark:bg-teal-500/10 blur-3xl" />
        <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[800px] h-[800px] rounded-full bg-cyan-500/3 dark:bg-cyan-500/5 blur-3xl" />
      </div>

      {/* ── Sticky Header ── */}
      <motion.header
        initial={{ y: 0 }}
        animate={{ y: headerVisible ? 0 : -100 }}
        transition={{ duration: 0.3 }}
        className="sticky top-0 z-40 backdrop-blur-xl bg-white/80 dark:bg-slate-900/80 border-b border-slate-200/50 dark:border-white/10"
      >
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <div className="flex items-center gap-3">
              <div className="flex items-center justify-center h-9 w-9 rounded-xl bg-gradient-to-br from-emerald-400 to-teal-500 shadow-lg shadow-emerald-500/20">
                <Presentation className="h-5 w-5 text-white" />
              </div>
              <span className="text-lg font-bold tracking-tight text-slate-900 dark:text-white">PPTX Editor</span>
              <Badge variant="secondary" className="text-[10px] hidden sm:inline-flex animate-float-badge">Beta</Badge>
            </div>

            {/* Desktop Nav */}
            <nav className="hidden md:flex items-center gap-6 text-sm text-slate-600 dark:text-slate-400">
              <a href="#features" className="hover:text-emerald-400 transition-colors">Features</a>
              <a href="#how-it-works" className="hover:text-emerald-400 transition-colors">How It Works</a>
              <a href="#faq" className="hover:text-emerald-400 transition-colors">FAQ</a>
            </nav>

            <div className="flex items-center gap-2">
              <ThemeToggle />
              <Button
                variant="ghost"
                size="icon"
                className="md:hidden"
                onClick={() => setMobileMenuOpen(!mobileMenuOpen)}
              >
                <Menu className="h-5 w-5" />
              </Button>
            </div>
          </div>
        </div>

        {/* Mobile Nav */}
        <AnimatePresence>
          {mobileMenuOpen && (
            <motion.div
              initial={{ height: 0, opacity: 0 }}
              animate={{ height: 'auto', opacity: 1 }}
              exit={{ height: 0, opacity: 0 }}
              className="md:hidden overflow-hidden border-t border-slate-200/50 dark:border-white/10"
            >
              <nav className="flex flex-col px-4 py-3 gap-3 text-sm text-slate-600 dark:text-slate-400">
                <a href="#features" className="hover:text-emerald-400 transition-colors" onClick={() => setMobileMenuOpen(false)}>Features</a>
                <a href="#how-it-works" className="hover:text-emerald-400 transition-colors" onClick={() => setMobileMenuOpen(false)}>How It Works</a>
                <a href="#faq" className="hover:text-emerald-400 transition-colors" onClick={() => setMobileMenuOpen(false)}>FAQ</a>
              </nav>
            </motion.div>
          )}
        </AnimatePresence>
      </motion.header>

      {/* ── Hero Section ── */}
      <section className="relative pt-12 sm:pt-20 lg:pt-28 pb-8 sm:pb-16 px-4 sm:px-6 lg:px-8 overflow-hidden">
        <FloatingParticles />

        {/* Gradient mesh background */}
        <div className="pointer-events-none absolute inset-0 overflow-hidden">
          <div className="absolute top-0 left-1/4 w-[500px] h-[500px] rounded-full bg-emerald-500/8 dark:bg-emerald-500/5 blur-[100px]" />
          <div className="absolute bottom-0 right-1/4 w-[400px] h-[400px] rounded-full bg-teal-500/8 dark:bg-teal-500/5 blur-[100px]" />
          <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[600px] h-[600px] rounded-full bg-cyan-500/5 dark:bg-cyan-500/3 blur-[120px]" />
        </div>

        <div className="relative max-w-7xl mx-auto">
          <div className="flex flex-col lg:flex-row items-center gap-10 lg:gap-16">
            {/* Left: Text + CTA */}
            <div className="flex-1 text-center lg:text-left max-w-2xl lg:max-w-none">
              <motion.div
                initial={{ opacity: 0, y: 30 }}
                animate={{ opacity: 1, y: 0 }}
                transition={{ duration: 0.8, ease: [0.22, 1, 0.36, 1] }}
              >
                <Badge variant="secondary" className="mb-6 px-4 py-1.5 text-sm bg-emerald-500/10 border-emerald-500/30 text-emerald-400">
                  <Sparkles className="h-3.5 w-3.5 mr-1.5" />
                  AI-Powered Template Editor
                </Badge>

                <h1 className="text-4xl sm:text-5xl lg:text-6xl font-extrabold tracking-tight text-slate-900 dark:text-white mb-6">
                  Edit PPTX templates
                  <br />
                  <span className="bg-gradient-to-r from-emerald-400 via-teal-400 to-cyan-400 bg-clip-text text-transparent">
                    <TypingText
                      texts={['with variables', 'with AI assist', 'losslessly', 'in seconds']}
                    />
                  </span>
                </h1>

                <p className="text-lg sm:text-xl text-slate-600 dark:text-slate-400 mb-10 max-w-xl mx-auto lg:mx-0 leading-relaxed">
                  Upload any .pptx file, fill in template variables manually or with AI, then export a perfectly formatted presentation. No formatting loss, no hassle.
                </p>

                {/* Hero Actions with shimmer */}
                <div className="flex flex-col sm:flex-row items-center justify-center lg:justify-start gap-4 mb-12">
                  <div className="relative group">
                    <Button
                      size="lg"
                      className="bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-600 hover:to-teal-600 text-white shadow-lg shadow-emerald-500/25 hover:shadow-xl hover:shadow-emerald-500/30 transition-all duration-300 px-8 h-12 text-base relative overflow-hidden"
                      onClick={() => {
                        const el = document.getElementById('upload-section');
                        el?.scrollIntoView({ behavior: 'smooth' });
                      }}
                    >
                      <span className="absolute inset-0 overflow-hidden rounded-md">
                        <span className="absolute inset-0 bg-gradient-to-r from-transparent via-white/20 to-transparent translate-x-[-100%] group-hover:translate-x-[100%] transition-transform duration-700" />
                      </span>
                      <Upload className="h-5 w-5 mr-2 relative z-10" />
                      <span className="relative z-10">Upload PPTX</span>
                    </Button>
                  </div>
                  <div className="relative group">
                    <Button
                      variant="outline"
                      size="lg"
                      className="border-slate-300 dark:border-white/20 hover:border-emerald-400 dark:hover:border-emerald-400 hover:bg-emerald-50 dark:hover:bg-emerald-900/20 transition-all duration-300 px-8 h-12 text-base relative overflow-hidden"
                      onClick={() => handleTemplateLoad(TEMPLATES[0])}
                    >
                      <span className="absolute inset-0 overflow-hidden rounded-md">
                        <span className="absolute inset-0 bg-gradient-to-r from-transparent via-emerald-400/10 to-transparent translate-x-[-100%] group-hover:translate-x-[100%] transition-transform duration-700" />
                      </span>
                      <Play className="h-5 w-5 mr-2 relative z-10" />
                      <span className="relative z-10">Try Demo</span>
                    </Button>
                  </div>
                </div>
              </motion.div>
            </div>

            {/* Right: Product Visual */}
            <div className="flex-1 w-full max-w-lg lg:max-w-none">
              <EditorMockup />
            </div>
          </div>
        </div>
      </section>

      {/* ── Stats Bar ── */}
      <section className="relative py-6 sm:py-8 px-4 sm:px-6 lg:px-8">
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          transition={{ duration: 0.6, delay: 0.4 }}
          className="max-w-4xl mx-auto"
        >
          <div className="flex items-center justify-center gap-6 sm:gap-0 sm:divide-x sm:divide-slate-200/60 dark:sm:divide-white/10 rounded-2xl bg-white/60 dark:bg-white/5 backdrop-blur-sm border border-slate-200/40 dark:border-white/10 px-6 sm:px-0 py-4">
            <div className="flex items-center gap-3 sm:px-6 sm:py-2">
              <div className="flex items-center justify-center h-10 w-10 rounded-xl bg-emerald-500/10 border border-emerald-500/20">
                <Shield className="h-5 w-5 text-emerald-500" />
              </div>
              <div>
                <div className="text-xl sm:text-2xl font-bold text-slate-900 dark:text-white">
                  <CountUp end={100} suffix="%" />
                </div>
                <div className="text-xs text-slate-500 dark:text-slate-400">Format Preserved</div>
              </div>
            </div>
            <div className="flex items-center gap-3 sm:px-6 sm:py-2">
              <div className="flex items-center justify-center h-10 w-10 rounded-xl bg-amber-500/10 border border-amber-500/20">
                <Zap className="h-5 w-5 text-amber-500" />
              </div>
              <div>
                <div className="text-xl sm:text-2xl font-bold text-slate-900 dark:text-white">
                  <CountUp end={10} suffix="x" />
                </div>
                <div className="text-xs text-slate-500 dark:text-slate-400">Faster Edits</div>
              </div>
            </div>
            <div className="flex items-center gap-3 sm:px-6 sm:py-2">
              <div className="flex items-center justify-center h-10 w-10 rounded-xl bg-violet-500/10 border border-violet-500/20">
                <Sparkles className="h-5 w-5 text-violet-500" />
              </div>
              <div>
                <div className="text-xl sm:text-2xl font-bold text-slate-900 dark:text-white">AI-Powered</div>
                <div className="text-xs text-slate-500 dark:text-slate-400">Smart Generation</div>
              </div>
            </div>
          </div>
        </motion.div>
      </section>

      {/* ── Upload & Features Section ── */}
      <section id="upload-section" className="relative py-16 sm:py-24 px-4 sm:px-6 lg:px-8">
        <div className="max-w-6xl mx-auto">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            transition={{ duration: 0.6 }}
            className="text-center mb-12"
          >
            <h2 className="text-3xl sm:text-4xl font-bold text-slate-900 dark:text-white mb-4">Get Started</h2>
            <p className="text-slate-500 dark:text-slate-400 max-w-xl mx-auto">Choose how you want to create or edit your presentation</p>
          </motion.div>

          {/* Upload Zone */}
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            transition={{ duration: 0.5 }}
            className="max-w-xl mx-auto mb-12"
          >
            <UploadZone
              dragOver={dragOver}
              setDragOver={setDragOver}
              onDrop={handleDrop}
              onFileSelect={handleFileUpload}
            />
            {uploadError && (
              <div className="mt-3 flex items-center gap-2 text-sm text-red-500">
                <AlertCircle className="h-4 w-4 shrink-0" />
                {uploadError}
              </div>
            )}
          </motion.div>

          {/* File History */}
          {fileHistory.length > 0 && (
            <motion.div
              initial={{ opacity: 0 }}
              whileInView={{ opacity: 1 }}
              viewport={{ once: true }}
              className="max-w-xl mx-auto mb-16"
            >
              <div className="flex items-center justify-between mb-3">
                <div className="flex items-center gap-2 text-sm font-medium text-slate-700 dark:text-slate-300">
                  <History className="h-4 w-4" />
                  Recent Files
                </div>
                <Button
                  variant="ghost"
                  size="sm"
                  className="h-7 text-xs text-slate-400 hover:text-red-400"
                  onClick={() => { clearFileHistory(); setFileHistory([]); }}
                >
                  <Trash2 className="h-3 w-3 mr-1" />
                  Clear
                </Button>
              </div>
              <div className="space-y-1.5 max-h-48 overflow-y-auto custom-scrollbar">
                {fileHistory.slice(0, 5).map((entry) => (
                  <button
                    key={entry.fileId}
                    onClick={() => handleReopenFile(entry)}
                    className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg border border-slate-200/50 dark:border-white/10 bg-white/60 dark:bg-white/5 hover:bg-white/80 dark:hover:bg-white/10 transition-all duration-200 text-left"
                  >
                    <FileText className="h-4 w-4 text-emerald-400 shrink-0" />
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium text-slate-700 dark:text-slate-300 truncate">{entry.fileName}</p>
                      <p className="text-xs text-slate-400">{entry.slideCount} slides &middot; {formatTimeAgo(entry.openedAt)}</p>
                    </div>
                    <ArrowRight className="h-3.5 w-3.5 text-slate-400 shrink-0" />
                  </button>
                ))}
              </div>
            </motion.div>
          )}

          {/* Feature Cards */}
          <div id="features" className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
            <GlassCard
              icon={<Upload className="h-6 w-6 text-emerald-400" />}
              title="Upload PPTX"
              description="Upload any .pptx file to parse and edit its template variables"
              onClick={() => {
                const el = document.getElementById('upload-section');
                el?.scrollIntoView({ behavior: 'smooth' });
              }}
              delay={0}
              stepNumber={1}
              iconAnimation="bounce"
            />
            <GlassCard
              icon={<Sparkles className="h-6 w-6 text-violet-400" />}
              title="AI Generate"
              description="Upload a data source and let AI fill template variables automatically"
              onClick={() => { setAiDialogOpen(true); }}
              delay={0.1}
              badge="AI"
              stepNumber={2}
              iconAnimation="pulse"
            />
            <GlassCard
              icon={<FileJson className="h-6 w-6 text-amber-400" />}
              title="JSON Import"
              description="Import slide data from a JSON definition file"
              onClick={() => {
                const json = prompt('Paste your JSON template data:');
                if (json) { setJsonInput(json); handleJsonImport(); }
              }}
              delay={0.2}
              stepNumber={3}
              iconAnimation="rotate"
            />
            <GlassCard
              icon={<Play className="h-6 w-6 text-cyan-400" />}
              title="Try Demo"
              description="Load the Annual Report demo template to explore the editor"
              onClick={() => handleTemplateLoad(TEMPLATES[0])}
              delay={0.3}
              stepNumber={4}
              iconAnimation="bounce"
            />
          </div>
        </div>
      </section>

      {/* ── How It Works Section ── */}
      <section id="how-it-works" className="py-16 sm:py-24 px-4 sm:px-6 lg:px-8">
        <div className="max-w-4xl mx-auto text-center">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            transition={{ duration: 0.6 }}
          >
            <h2 className="text-3xl sm:text-4xl font-bold text-slate-900 dark:text-white mb-4">How It Works</h2>
            <p className="text-slate-500 dark:text-slate-400 max-w-xl mx-auto mb-4">Three simple steps to create professional presentations</p>
          </motion.div>
          <HowItWorks />
        </div>
      </section>

      {/* ── Features Grid ── */}
      <section className="py-16 sm:py-24 px-4 sm:px-6 lg:px-8">
        <div className="max-w-6xl mx-auto">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            transition={{ duration: 0.6 }}
            className="text-center mb-12"
          >
            <h2 className="text-3xl sm:text-4xl font-bold text-slate-900 dark:text-white mb-4">Why PPTX Editor?</h2>
            <p className="text-slate-500 dark:text-slate-400 max-w-xl mx-auto">Built for professionals who need fast, accurate template editing</p>
          </motion.div>

          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6">
            {[
              { icon: <Zap className="h-6 w-6 text-amber-400" />, title: 'Lightning Fast', desc: 'Parse and edit PPTX files in seconds, not minutes' },
              { icon: <Shield className="h-6 w-6 text-emerald-400" />, title: 'Lossless Export', desc: '100% formatting preservation on export — fonts, colors, layouts, images' },
              { icon: <Sparkles className="h-6 w-6 text-violet-400" />, title: 'AI-Powered', desc: 'Auto-fill template variables using OpenAI or Anthropic models' },
              { icon: <Globe className="h-6 w-6 text-cyan-400" />, title: 'Multi-Provider AI', desc: 'Support for OpenAI, Anthropic, and compatible API endpoints' },
              { icon: <BookOpen className="h-6 w-6 text-rose-400" />, title: 'Template Library', desc: '7+ pre-built templates for common business presentations' },
              { icon: <Code2 className="h-6 w-6 text-teal-400" />, title: 'JSON Import/Export', desc: 'Full JSON support for programmatic template editing' },
            ].map((feature, i) => (
              <motion.div
                key={feature.title}
                initial={{ opacity: 0, y: 30 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true }}
                transition={{ duration: 0.5, delay: 0.1 * i }}
                className="group rounded-xl border border-slate-200/50 dark:border-white/10 bg-white/60 dark:bg-white/5 backdrop-blur-sm p-6 transition-all duration-300 hover:border-slate-300/50 dark:hover:border-white/20 hover:bg-white/80 dark:hover:bg-white/8 hover:shadow-lg"
              >
                <div className="mb-4 inline-flex h-12 w-12 items-center justify-center rounded-xl bg-slate-100 dark:bg-white/10 group-hover:scale-110 transition-transform duration-300">
                  {feature.icon}
                </div>
                <h3 className="text-lg font-bold text-slate-900 dark:text-white mb-2">{feature.title}</h3>
                <p className="text-sm text-slate-500 dark:text-slate-400 leading-relaxed">{feature.desc}</p>
              </motion.div>
            ))}
          </div>
        </div>
      </section>

      {/* ── FAQ Section ── */}
      <section id="faq" className="py-16 sm:py-24 px-4 sm:px-6 lg:px-8">
        <div className="max-w-3xl mx-auto">
          <motion.div
            initial={{ opacity: 0, y: 20 }}
            whileInView={{ opacity: 1, y: 0 }}
            viewport={{ once: true }}
            transition={{ duration: 0.6 }}
            className="text-center mb-12"
          >
            <h2 className="text-3xl sm:text-4xl font-bold text-slate-900 dark:text-white mb-4">Frequently Asked Questions</h2>
            <p className="text-slate-500 dark:text-slate-400">Everything you need to know</p>
          </motion.div>

          <Accordion type="single" collapsible className="space-y-3">
            {FAQ_ITEMS.map((item, index) => (
              <motion.div
                key={index}
                initial={{ opacity: 0, y: 10 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true }}
                transition={{ duration: 0.3, delay: index * 0.05 }}
              >
                <AccordionItem
                  value={`faq-${index}`}
                  className="border border-slate-200/50 dark:border-white/10 rounded-xl px-6 bg-white/60 dark:bg-white/5 backdrop-blur-sm data-[state=open]:bg-white/80 dark:data-[state=open]:bg-white/8 data-[state=open]:border-emerald-200/50 dark:data-[state=open]:border-emerald-500/20 data-[state=open]:shadow-md data-[state=open]:shadow-emerald-500/5 transition-all duration-300"
                >
                  <AccordionTrigger className="text-left text-sm font-semibold text-slate-900 dark:text-white hover:no-underline py-4 transition-colors duration-200 data-[state=open]:text-emerald-600 dark:data-[state=open]:text-emerald-400">
                    {item.question}
                  </AccordionTrigger>
                  <AccordionContent className="text-sm text-slate-500 dark:text-slate-400 leading-relaxed pb-4 data-[state=open]:animate-in data-[state=open]:fade-in data-[state=open]:slide-in-from-top-2 data-[state=closed]:animate-out data-[state=closed]:fade-out data-[state=closed]:slide-out-to-top-2">
                    {item.answer}
                  </AccordionContent>
                </AccordionItem>
              </motion.div>
            ))}
          </Accordion>
        </div>
      </section>

      {/* ── CTA Section ── */}
      <section className="py-16 sm:py-24 px-4 sm:px-6 lg:px-8">
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          transition={{ duration: 0.6 }}
          className="max-w-4xl mx-auto text-center"
        >
          <div className="rounded-3xl border border-emerald-200/50 dark:border-emerald-500/20 bg-gradient-to-br from-emerald-50 to-teal-50 dark:from-emerald-900/20 dark:to-teal-900/10 p-8 sm:p-12 lg:p-16 relative overflow-hidden">
            <div className="pointer-events-none absolute inset-0">
              <div className="absolute -top-20 -right-20 w-60 h-60 rounded-full bg-emerald-300/20 blur-3xl" />
              <div className="absolute -bottom-20 -left-20 w-60 h-60 rounded-full bg-teal-300/20 blur-3xl" />
            </div>

            <div className="relative">
              <h2 className="text-3xl sm:text-4xl font-bold text-slate-900 dark:text-white mb-4">Ready to edit your templates?</h2>
              <p className="text-lg text-slate-600 dark:text-slate-400 mb-8 max-w-lg mx-auto">
                Upload a PPTX file or choose a template to get started in seconds.
              </p>
              <div className="flex flex-col sm:flex-row items-center justify-center gap-4">
                <div className="relative group">
                  <Button
                    size="lg"
                    className="bg-gradient-to-r from-emerald-500 to-teal-500 hover:from-emerald-600 hover:to-teal-600 text-white shadow-lg shadow-emerald-500/25 hover:shadow-xl transition-all duration-300 px-8 h-12 text-base relative overflow-hidden"
                    onClick={() => {
                      const el = document.getElementById('upload-section');
                      el?.scrollIntoView({ behavior: 'smooth' });
                    }}
                  >
                    <span className="absolute inset-0 overflow-hidden rounded-md">
                      <span className="absolute inset-0 bg-gradient-to-r from-transparent via-white/20 to-transparent translate-x-[-100%] group-hover:translate-x-[100%] transition-transform duration-700" />
                    </span>
                    <Upload className="h-5 w-5 mr-2 relative z-10" />
                    <span className="relative z-10">Upload PPTX</span>
                  </Button>
                </div>
                <div className="relative group">
                  <Button
                    variant="outline"
                    size="lg"
                    className="border-emerald-300 dark:border-emerald-700 text-emerald-600 dark:text-emerald-400 hover:bg-emerald-100 dark:hover:bg-emerald-900/30 transition-all duration-300 px-8 h-12 text-base relative overflow-hidden"
                    onClick={() => handleTemplateLoad(TEMPLATES[0])}
                  >
                    <span className="absolute inset-0 overflow-hidden rounded-md">
                      <span className="absolute inset-0 bg-gradient-to-r from-transparent via-emerald-400/10 to-transparent translate-x-[-100%] group-hover:translate-x-[100%] transition-transform duration-700" />
                    </span>
                    <Play className="h-5 w-5 mr-2 relative z-10" />
                    <span className="relative z-10">Try Demo</span>
                  </Button>
                </div>
              </div>
            </div>
          </div>
        </motion.div>
      </section>

      {/* ── Footer ── */}
      <footer className="mt-auto border-t border-slate-200/50 dark:border-white/10 bg-white/60 dark:bg-white/5 backdrop-blur-sm">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
          <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="flex items-center justify-center h-8 w-8 rounded-lg bg-gradient-to-br from-emerald-400 to-teal-500">
                <Presentation className="h-4 w-4 text-white" />
              </div>
              <span className="text-sm font-semibold text-slate-900 dark:text-white">PPTX Editor</span>
            </div>
            <div className="flex items-center gap-6 text-sm text-slate-500 dark:text-slate-400">
              <span>Built with Next.js &amp; TypeScript</span>
              <span>&middot;</span>
              <span>Powered by AI</span>
            </div>
            <div className="flex items-center gap-1">
              <ThemeToggle />
            </div>
          </div>
        </div>
      </footer>

      {/* AI Dialogs (available on landing page too) */}
      <AiGenerateDialog
        open={aiDialogOpen}
        onOpenChange={setAiDialogOpen}
        onOpenSettings={() => {
          setAiDialogOpen(false);
          setAiSettingsOpen(true);
        }}
      />
      <AiSettingsDialog
        open={aiSettingsOpen}
        onOpenChange={setAiSettingsOpen}
        onConfigChange={() => {}}
      />
    </div>
  );
}
