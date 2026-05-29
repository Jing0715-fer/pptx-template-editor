---
Task ID: 1
Agent: Main Agent
Task: Fix PPTX export, image preview, and alignment issues

Work Log:
- Diagnosed exported PPTX files not opening in Office - caused by AdmZip library producing unreliable OOXML output
- Rewrote pptx-replacer.ts to use JSZip instead of AdmZip for both reading and writing PPTX files
- Fixed image replacement not displaying in UI - double data URL prefix bug in slide-editor.tsx useEffect
- Fixed preview alignment issues by parsing actual slide dimensions from ppt/presentation.xml instead of guessing
- Added SlideSize type to parser, store, and API responses
- Updated SlidePreview component to use actual slide dimensions from the store
- All exports validated with python-pptx - text modification, image replacement, cross-type image replacement, and combined modifications all work correctly
- Lint check passes

Stage Summary:
- pptx-replacer.ts: Switched from AdmZip to JSZip for reliable OOXML output
- slide-editor.tsx: Fixed double data URL prefix bug in image preview useEffect
- pptx-parser.ts: Added SlideSize extraction from presentation.xml
- pptx-store.ts: Added slideSize to store state and setParsedData
- parse/route.ts & reparse/route.ts: Pass slideSize in API responses
- upload-zone.tsx: Pass slideSize to setParsedData
- slide-preview.tsx: Use actual slideSize from store instead of heuristic detection
