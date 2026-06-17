import { create } from 'zustand';

// ============================================================================
// Types
// ============================================================================

export interface ElementPosition {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface PptxRun {
  originalText: string;
  bold: boolean;
  italic: boolean;
  fontSize: number | null;
  fontColor: string | null;
  fontName: string | null;
}

export interface PptxParagraph {
  originalText: string;
  runs: PptxRun[];
}

export interface PptxTextElement {
  type: 'text';
  id: string;
  shapeName: string;
  originalText: string;
  paragraphs: PptxParagraph[];
  position: ElementPosition;
  slideIndex: number;
  elementIndex: number;
  currentText?: string;
}

export interface PptxTableCell {
  text: string;
  originalText: string;
  rowSpan: number;
  colSpan: number;
  gridCol: number;
}

export interface PptxTableRow {
  cells: PptxTableCell[];
}

export interface PptxTableElement {
  type: 'table';
  id: string;
  shapeName: string;
  rows: PptxTableRow[];
  position: ElementPosition;
  slideIndex: number;
  elementIndex: number;
  currentRows?: PptxTableRow[];
}

export interface PptxImageElement {
  type: 'image';
  id: string;
  shapeName: string;
  imageName: string;
  imageType: string;
  imageRid: string;
  imageData?: string | null;
  position: ElementPosition;
  slideIndex: number;
  elementIndex: number;
  replacementImageData?: string | null;
  replacementImageType?: string | null;
}

export type PptxElement = PptxTextElement | PptxTableElement | PptxImageElement;

export interface PptxSlideData {
  slideNumber: number;
  elements: PptxElement[];
  previewImage?: string | null;
}

export interface SlideSize {
  width: number;  // EMU
  height: number; // EMU
}

export interface PptxModification {
  slideIndex: number;
  elementIndex: number;
  type: 'text' | 'table';
  newText?: string;
  tableCells?: { row: number; col: number; text: string }[];
}

export interface PptxImageModification {
  slideIndex: number;
  imageRid: string;
  newImageData: string;
  newImageType: string;
}

interface JsonElement {
  type: 'text' | 'table' | 'image';
  id: string;
  shapeName?: string;
  originalText?: string;
  currentText?: string;
  paragraphs?: PptxParagraph[];
  rows?: PptxTableRow[];
  currentRows?: PptxTableRow[];
  position?: ElementPosition;
  slideIndex?: number;
  elementIndex?: number;
  imageName?: string;
  imageType?: string;
}

interface JsonSlide {
  slideNumber: number;
  elements: JsonElement[];
}

export interface PptxJsonData {
  fileName?: string;
  fileId?: string;
  slideCount?: number;
  modifications?: PptxModification[];
  slides?: JsonSlide[];
}

// ============================================================================
// Store
// ============================================================================

type AppStep = 'upload' | 'loading' | 'editing';

const MAX_HISTORY = 50;

type HistoryEntry = { slides: PptxSlideData[] };

interface PptxStore {
  step: AppStep;
  fileId: string | null;
  fileName: string | null;
  slides: PptxSlideData[];
  slideSize: SlideSize;
  currentSlideIndex: number;
  selectedElementId: string | null;
  hideEmpty: boolean;
  undoStack: HistoryEntry[];
  redoStack: HistoryEntry[];
  slideNotes: Record<number, string>;
  hiddenElementIds: Set<string>;
  elementComments: Record<string, string>;
  batchSelectedIds: Set<string>;

  setStep: (step: AppStep) => void;
  setParsedData: (fileId: string, fileName: string, slides: PptxSlideData[], slideSize?: SlideSize) => void;
  setCurrentSlide: (index: number) => void;
  selectElement: (id: string | null) => void;
  updateText: (elementId: string, newText: string) => void;
  updateTableCell: (elementId: string, row: number, col: number, text: string) => void;
  updateImage: (elementId: string, imageData: string, imageType: string) => void;
  removeImage: (elementId: string) => void;
  getModifications: () => PptxModification[];
  getImageModifications: () => PptxImageModification[];
  getTotalModificationCount: () => number;
  loadFromJson: (jsonData: PptxJsonData) => void;
  resetAllModifications: () => void;
  toggleHideEmpty: () => void;
  updateSlidePreviews: (previewImages: (string | null)[]) => void;
  reorderElements: (slideIndex: number, oldIndex: number, newIndex: number) => void;
  reorderSlides: (oldIndex: number, newIndex: number) => void;
  applyAiModifications: (modifications: { slideIndex: number; elementIndex: number; type: 'text' | 'table'; newText?: string; tableCells?: { row: number; col: number; text: string }[] }[]) => void;
  updateSlideNote: (slideNumber: number, note: string) => void;
  updateElementComment: (elementId: string, comment: string) => void;
  duplicateTextElement: (elementId: string) => void;
  toggleElementVisibility: (id: string) => void;
  isElementHidden: (id: string) => boolean;
  toggleBatchSelect: (elementId: string) => void;
  clearBatchSelection: () => void;
  batchSelectAll: (elementIds: string[]) => void;
  undo: () => void;
  redo: () => void;
  canUndo: () => boolean;
  canRedo: () => boolean;
  reset: () => void;
}

function cloneSlides(slides: PptxSlideData[]): PptxSlideData[] {
  return JSON.parse(JSON.stringify(slides));
}

export const usePptxStore = create<PptxStore>((set, get) => ({
  step: 'upload',
  fileId: null,
  fileName: null,
  slides: [],
  slideSize: { width: 12192000, height: 6858000 },
  currentSlideIndex: 0,
  selectedElementId: null,
  hideEmpty: true,
  undoStack: [],
  redoStack: [],
  slideNotes: {},
  hiddenElementIds: new Set<string>(),
  elementComments: {},
  batchSelectedIds: new Set<string>(),

  setStep: (step) => set({ step }),

  setParsedData: (fileId, fileName, slides, slideSize) =>
    set({ fileId, fileName, slides, slideSize: slideSize || { width: 12192000, height: 6858000 }, currentSlideIndex: 0, selectedElementId: null, step: 'editing', undoStack: [], redoStack: [] }),

  setCurrentSlide: (index) => set({ currentSlideIndex: index, selectedElementId: null, hiddenElementIds: new Set<string>(), batchSelectedIds: new Set<string>() }),
  selectElement: (id) => set({ selectedElementId: id }),

  updateText: (elementId, newText) =>
    set((state) => {
      const undoEntry = { slides: cloneSlides(state.slides) };
      return {
        undoStack: [...state.undoStack.slice(-(MAX_HISTORY - 1)), undoEntry],
        redoStack: [],
        slides: state.slides.map((slide) => ({
          ...slide,
          elements: slide.elements.map((el) =>
            el.id === elementId && el.type === 'text' ? { ...el, currentText: newText } : el
          ),
        })),
      };
    }),

  updateTableCell: (elementId, row, col, text) =>
    set((state) => {
      const undoEntry = { slides: cloneSlides(state.slides) };
      return {
        undoStack: [...state.undoStack.slice(-(MAX_HISTORY - 1)), undoEntry],
        redoStack: [],
        slides: state.slides.map((slide) => ({
          ...slide,
          elements: slide.elements.map((el) => {
            if (el.id === elementId && el.type === 'table') {
              const currentRows = el.currentRows || el.rows.map((r) => ({
                ...r, cells: r.cells.map((c) => ({ ...c })),
              }));
              const newRows = currentRows.map((r, ri) =>
                ri === row ? { ...r, cells: r.cells.map((c, ci) => ci === col ? { ...c, text } : c) } : r
              );
              return { ...el, currentRows: newRows };
            }
            return el;
          }),
        })),
      };
    }),

  updateImage: (elementId, imageData, imageType) =>
    set((state) => {
      const undoEntry = { slides: cloneSlides(state.slides) };
      return {
        undoStack: [...state.undoStack.slice(-(MAX_HISTORY - 1)), undoEntry],
        redoStack: [],
        slides: state.slides.map((slide) => ({
          ...slide,
          elements: slide.elements.map((el) =>
            el.id === elementId && el.type === 'image'
              ? { ...el, replacementImageData: imageData, replacementImageType: imageType }
              : el
          ),
        })),
      };
    }),

  removeImage: (elementId) =>
    set((state) => {
      const undoEntry = { slides: cloneSlides(state.slides) };
      return {
        undoStack: [...state.undoStack.slice(-(MAX_HISTORY - 1)), undoEntry],
        redoStack: [],
        slides: state.slides.map((slide) => ({
          ...slide,
          elements: slide.elements.map((el) =>
            el.id === elementId && el.type === 'image'
              ? { ...el, replacementImageData: null, replacementImageType: null }
              : el
          ),
        })),
      };
    }),

  getModifications: () => {
    const { slides } = get();
    const modifications: PptxModification[] = [];
    for (const slide of slides) {
      for (const el of slide.elements) {
        if (el.type === 'text' && el.currentText !== undefined && el.currentText !== el.originalText) {
          modifications.push({ slideIndex: el.slideIndex, elementIndex: el.elementIndex, type: 'text', newText: el.currentText });
        } else if (el.type === 'table' && el.currentRows) {
          const cellMods: { row: number; col: number; text: string }[] = [];
          for (let ri = 0; ri < el.currentRows.length; ri++) {
            const origRow = el.rows[ri];
            const curRow = el.currentRows[ri];
            if (!origRow || !curRow) continue;
            for (let ci = 0; ci < curRow.cells.length; ci++) {
              const origCell = origRow.cells[ci];
              const curCell = curRow.cells[ci];
              if (!origCell || !curCell) continue;
              if (curCell.text !== origCell.text) cellMods.push({ row: ri, col: ci, text: curCell.text });
            }
          }
          if (cellMods.length > 0) {
            modifications.push({ slideIndex: el.slideIndex, elementIndex: el.elementIndex, type: 'table', tableCells: cellMods });
          }
        }
      }
    }
    return modifications;
  },

  getImageModifications: () => {
    const { slides } = get();
    const imageModifications: PptxImageModification[] = [];
    for (const slide of slides) {
      for (const el of slide.elements) {
        if (el.type === 'image' && el.replacementImageData && el.replacementImageType) {
          imageModifications.push({
            slideIndex: el.slideIndex,
            imageRid: el.imageRid,
            newImageData: el.replacementImageData,
            newImageType: el.replacementImageType,
          });
        }
      }
    }
    return imageModifications;
  },

  getTotalModificationCount: () => {
    const { slides } = get();
    let count = 0;
    for (const slide of slides) {
      for (const el of slide.elements) {
        if (el.type === 'text' && el.currentText !== undefined && el.currentText !== el.originalText) count++;
        else if (el.type === 'table' && el.currentRows) {
          for (let ri = 0; ri < el.currentRows.length; ri++) {
            const origRow = el.rows[ri];
            const curRow = el.currentRows[ri];
            if (!origRow || !curRow) continue;
            for (let ci = 0; ci < curRow.cells.length; ci++) {
              if (!origRow.cells[ci] || !curRow.cells[ci]) continue;
              if (curRow.cells[ci].text !== origRow.cells[ci].text) count++;
            }
          }
        } else if (el.type === 'image' && el.replacementImageData) count++;
      }
    }
    return count;
  },

  loadFromJson: (jsonData) => {
    const { slides } = get();
    if (!slides.length || !jsonData.slides) return;
    const jsonElementMap = new Map<string, JsonElement>();
    for (const slide of jsonData.slides) {
      for (const el of slide.elements) jsonElementMap.set(el.id, el);
    }
    const updatedSlides = slides.map((slide) => ({
      ...slide,
      elements: slide.elements.map((el) => {
        const jsonEl = jsonElementMap.get(el.id);
        if (!jsonEl) return el;
        if (el.type === 'text' && jsonEl.type === 'text' && jsonEl.currentText !== undefined) return { ...el, currentText: jsonEl.currentText };
        if (el.type === 'table' && jsonEl.type === 'table' && jsonEl.currentRows) return { ...el, currentRows: jsonEl.currentRows };
        return el;
      }),
    }));
    set({ slides: updatedSlides });
  },

  resetAllModifications: () =>
    set((state) => {
      const undoEntry = { slides: cloneSlides(state.slides) };
      return {
        undoStack: [...state.undoStack.slice(-(MAX_HISTORY - 1)), undoEntry],
        redoStack: [],
        slides: state.slides.map((slide) => ({
          ...slide,
          elements: slide.elements.map((el) => {
            if (el.type === 'text' && el.currentText !== undefined) return { ...el, currentText: undefined };
            if (el.type === 'table' && el.currentRows !== undefined) return { ...el, currentRows: undefined };
            if (el.type === 'image' && el.replacementImageData !== undefined) return { ...el, replacementImageData: null, replacementImageType: null };
            return el;
          }),
        })),
        selectedElementId: null,
      };
    }),

  toggleHideEmpty: () => set((state) => ({ hideEmpty: !state.hideEmpty })),

  updateSlidePreviews: (previewImages) =>
    set((state) => ({
      slides: state.slides.map((slide, index) => ({ ...slide, previewImage: previewImages[index] || null })),
    })),

  reorderElements: (slideIndex, oldIndex, newIndex) =>
    set((state) => {
      if (oldIndex === newIndex) return state;
      const undoEntry = { slides: cloneSlides(state.slides) };
      const updatedSlides = state.slides.map((slide, si) => {
        if (si !== slideIndex) return slide;
        const elements = [...slide.elements];
        const [moved] = elements.splice(oldIndex, 1);
        elements.splice(newIndex, 0, moved);
        return { ...slide, elements };
      });
      return {
        undoStack: [...state.undoStack.slice(-(MAX_HISTORY - 1)), undoEntry],
        redoStack: [],
        slides: updatedSlides,
      };
    }),

  reorderSlides: (oldIndex, newIndex) =>
    set((state) => {
      if (oldIndex === newIndex || oldIndex < 0 || newIndex < 0 || oldIndex >= state.slides.length || newIndex >= state.slides.length) return state;
      const undoEntry = { slides: cloneSlides(state.slides) };
      const slides = [...state.slides];
      const [moved] = slides.splice(oldIndex, 1);
      slides.splice(newIndex, 0, moved);
      // Update currentSlideIndex if the current slide was moved
      let currentSlideIndex = state.currentSlideIndex;
      if (currentSlideIndex === oldIndex) {
        currentSlideIndex = newIndex;
      } else if (oldIndex < currentSlideIndex && newIndex >= currentSlideIndex) {
        currentSlideIndex--;
      } else if (oldIndex > currentSlideIndex && newIndex <= currentSlideIndex) {
        currentSlideIndex++;
      }
      return {
        undoStack: [...state.undoStack.slice(-(MAX_HISTORY - 1)), undoEntry],
        redoStack: [],
        slides,
        currentSlideIndex,
      };
    }),

  applyAiModifications: (modifications) =>
    set((state) => {
      const undoEntry = { slides: cloneSlides(state.slides) };
      const updatedSlides = state.slides.map((slide) => ({
        ...slide,
        elements: slide.elements.map((el) => {
          const match = modifications.find(
            (mod) => mod.slideIndex === el.slideIndex && mod.elementIndex === el.elementIndex && mod.type === el.type
          );
          if (!match) return el;
          if (el.type === 'text' && match.type === 'text' && match.newText !== undefined) return { ...el, currentText: match.newText };
          if (el.type === 'table' && match.type === 'table' && match.tableCells) {
            const currentRows = el.currentRows || el.rows.map((r) => ({ ...r, cells: r.cells.map((c) => ({ ...c })) }));
            const newRows = currentRows.map((r, ri) => ({
              ...r,
              cells: r.cells.map((c, ci) => {
                const cellMod = match.tableCells!.find((cell) => cell.row === ri && cell.col === ci);
                return cellMod ? { ...c, text: cellMod.text } : c;
              }),
            }));
            return { ...el, currentRows: newRows };
          }
          return el;
        }),
      }));
      return {
        undoStack: [...state.undoStack.slice(-(MAX_HISTORY - 1)), undoEntry],
        redoStack: [],
        slides: updatedSlides,
        selectedElementId: null,
      };
    }),

  undo: () => {
    const { undoStack } = get();
    if (undoStack.length === 0) return;
    const previous = undoStack[undoStack.length - 1];
    const redoEntry = { slides: cloneSlides(get().slides) };
    set({
      slides: previous.slides,
      undoStack: undoStack.slice(0, -1),
      redoStack: [...get().redoStack, redoEntry],
      selectedElementId: null,
    });
  },

  redo: () => {
    const { redoStack } = get();
    if (redoStack.length === 0) return;
    const next = redoStack[redoStack.length - 1];
    const undoEntry = { slides: cloneSlides(get().slides) };
    set({
      slides: next.slides,
      undoStack: [...get().undoStack, undoEntry],
      redoStack: redoStack.slice(0, -1),
      selectedElementId: null,
    });
  },

  updateSlideNote: (slideNumber, note) =>
    set((state) => ({
      slideNotes: { ...state.slideNotes, [slideNumber]: note },
    })),

  updateElementComment: (elementId, comment) =>
    set((state) => {
      const updated = { ...state.elementComments };
      if (comment.trim() === '') {
        delete updated[elementId];
      } else {
        updated[elementId] = comment;
      }
      return { elementComments: updated };
    }),

  duplicateTextElement: (elementId) =>
    set((state) => {
      const undoEntry = { slides: cloneSlides(state.slides) };
      const updatedSlides = state.slides.map((slide) => {
        const elIndex = slide.elements.findIndex((el) => el.id === elementId && el.type === 'text');
        if (elIndex === -1) return slide;
        const original = slide.elements[elIndex] as PptxTextElement;
        const copy: PptxTextElement = {
          ...original,
          id: `${original.id}-copy`,
          currentText: original.currentText ?? original.originalText,
        };
        const elements = [...slide.elements];
        elements.splice(elIndex + 1, 0, copy);
        return { ...slide, elements };
      });
      return {
        undoStack: [...state.undoStack.slice(-(MAX_HISTORY - 1)), undoEntry],
        redoStack: [],
        slides: updatedSlides,
      };
    }),

  toggleElementVisibility: (id) =>
    set((state) => {
      const newSet = new Set(state.hiddenElementIds);
      if (newSet.has(id)) {
        newSet.delete(id);
      } else {
        newSet.add(id);
      }
      return { hiddenElementIds: newSet };
    }),

  isElementHidden: (id) => get().hiddenElementIds.has(id),

  toggleBatchSelect: (elementId) =>
    set((state) => {
      const newSet = new Set(state.batchSelectedIds);
      if (newSet.has(elementId)) {
        newSet.delete(elementId);
      } else {
        newSet.add(elementId);
      }
      return { batchSelectedIds: newSet };
    }),

  clearBatchSelection: () => set({ batchSelectedIds: new Set<string>() }),

  batchSelectAll: (elementIds) => set({ batchSelectedIds: new Set(elementIds) }),

  canUndo: () => get().undoStack.length > 0,
  canRedo: () => get().redoStack.length > 0,

  reset: () => set({
    step: 'upload', fileId: null, fileName: null, slides: [],
    slideSize: { width: 12192000, height: 6858000 },
    currentSlideIndex: 0, selectedElementId: null, hideEmpty: true,
    undoStack: [], redoStack: [], slideNotes: {}, hiddenElementIds: new Set<string>(), elementComments: {},
    batchSelectedIds: new Set<string>(),
  }),
}));
