---
Task ID: 1
Agent: Main Agent
Task: Fix PPTX export bug (files can't be opened by Office), fix image replacement display, fix preview alignment, and create E2E tests

Work Log:
- Investigated full codebase: pptx-replacer.ts, pptx-parser.ts, export route, parse route, slide-editor, slide-preview, toolbar
- Ran diagnostic scripts to test export with actual PPTX files
- Identified key issues and applied fixes

- Fixed pptx-replacer.ts:
  - Changed to per-file compression (DEFLATE for XML, STORE for binary/images)
  - Fixed text modification order: collect all modifications first, find regions before any changes, apply in reverse position order
  - Added re-compression of unmodified XML files with DEFLATE for Office compatibility
  - Added round-trip parse verification using parsePptx after generation

- Fixed export API route:
  - Removed Content-Encoding: identity header
  - Added round-trip test verification call

- Fixed slide-editor.tsx (ImageElementDisplay):
  - Used React.useMemo for preview URL computation
  - Changed MIME type detection from file extension to data URL MIME type
  - Added FileReader onerror handler and image load error handling
  - Added visual "已替换" badge on replaced images

- Fixed slide-preview.tsx:
  - Changed from pixel-based to percentage-based positioning
  - Eliminates cumulative scaling errors that caused frames to be too high on lower part

- Created comprehensive E2E test suites (15 + 5 tests), all passed

Stage Summary:
- All export bugs fixed and verified with LibreOffice and round-trip tests
- Image replacement display improved with error handling
- Preview alignment fixed with percentage-based positioning
