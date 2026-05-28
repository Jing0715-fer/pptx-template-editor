---
Task ID: 1
Agent: Main Agent
Task: Build PPTX Template Editor with bug fixes and UI refactoring

Work Log:
- Analyzed the original GitHub repo codebase thoroughly
- Identified 8 critical bugs in the original pptx-replacer.ts (slideIndex off-by-one, broken table replacement, image loss, etc.)
- Installed required packages: adm-zip, jszip, fast-xml-parser, mammoth, xlsx
- Created backend lib files with the FIXED replacer (using sedD4rfMp corrected version)
- Created all 7 API routes (parse, export, preview, check, reparse, save-json, ai/generate)
- Created 6 frontend components (upload-zone, slide-editor, slide-navigator, slide-preview, toolbar, ai-generate-dialog)
- Created main page.tsx with full app flow
- Fixed image preview URL handling for replacementImageData (needs data URL prefix)
- Updated export route to pass imageModifications directly (no conversion needed)
- All lint checks pass, dev server running successfully

Stage Summary:
- Key bug fixes: slideIndex+1 for correct PPTX paths, paragraph-aware text replacement (last run), proper table cell replacement, group shape handling, correct XML text escaping
- Key architectural fix: Export route passes imageModifications directly as base64 strings instead of converting to Buffers
- UI fully refactored with shadcn/ui, Framer Motion animations, responsive design, engineering-grade quality
- All functionality preserved: upload, parse, edit text/table/image, AI generate, export PPTX, save JSON, file history
