---
Task ID: 1
Agent: Main Agent
Task: Fix PPTX export corruption (Windows Office can't open, re-upload parse failure)

Work Log:
- Investigated the PPTX export code in pptx-replacer.ts
- Identified root cause: WPS-created PPTX files contain invalid MIME types in [Content_Types].xml
  - `Extension="JPG" ContentType="image/.jpg"` (extra dot before extension - OPC spec violation)
  - Uppercase extensions like `Extension="JPG"` instead of `Extension="jpg"`
  - 82 WPS-specific `ppt/tags/tag*.xml` files
- Windows Office strictly validates OPC spec and rejects files with invalid ContentType values
- macOS Office is more lenient and can bypass these errors

Stage Summary:
- Added `sanitizeContentTypesXml()` function to fix invalid MIME types
- Added `removeWpsSpecificFiles()` to remove `ppt/tags/` directory entries
- Added `sanitizeWpsXmlAttributes()` to remove WPS-specific XML namespace declarations
- Fixed naive MIME construction: `image/${newExt}` → `IMAGE_MIME_MAP[newExt]`
- Added MIME type validation in `verifyPptxBuffer()`
- All 21 e2e tests pass

---
Task ID: 2
Agent: Main Agent
Task: Fix image replacement not displaying in UI

Work Log:
- Found that EMF/WMF images from PPTX files can't be displayed in browsers
- The `computedPreviewUrl` now filters out non-displayable MIME types
- Added `isBrowserDisplayable` flag to show proper error messages
- Added replacement image thumbnail overlay in SlidePreview component

Stage Summary:
- Image replacement preview works correctly for browser-displayable formats
- Non-displayable formats show helpful error messages
- Replacement images are visible in both editor panel and slide preview

---
Task ID: 3
Agent: Main Agent
Task: Complete e2e testing

Work Log:
- Created comprehensive e2e test script at tests/e2e-pptx-export.ts
- Test suite covers 8 test categories with 21 assertions
- Added `test:e2e` script to package.json
- All 21 tests pass with real WPS-created PPTX files

Stage Summary:
- Full e2e test suite created and verified
- Confirmed that `image/.jpg` → `image/jpeg` fix works correctly
- Confirmed that WPS tag files are properly removed
- Added npm script `bun run test:e2e` for easy re-running
