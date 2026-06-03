#!/usr/bin/env python3
"""
PDF → Image converter using PyMuPDF (fitz).

Faster than pdftoppm and has no native dependencies on macOS.
Usage: pdf2img.py <pdf_path> <output_prefix> <dpi> <quality>
Output: writes <output_prefix>-1.jpg, <output_prefix>-2.jpg, ...
"""
import sys
import os
import json
import base64
import io

def main():
    if len(sys.argv) < 5:
        print(json.dumps({"success": False, "error": "Usage: pdf2img.py <pdf> <out_prefix> <dpi> <quality>"}))
        sys.exit(1)
    pdf_path = sys.argv[1]
    out_prefix = sys.argv[2]
    dpi = int(sys.argv[3])
    quality = int(sys.argv[4])
    max_bytes = int(sys.argv[5]) if len(sys.argv) > 5 else 500 * 1024

    try:
        import fitz  # PyMuPDF
    except ImportError:
        print(json.dumps({"success": False, "error": "PyMuPDF not installed. Run: pip install pymupdf"}))
        sys.exit(1)

    try:
        doc = fitz.open(pdf_path)
        written = []
        for i, page in enumerate(doc, start=1):
            # Render at given DPI
            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            # Encode as JPEG with given quality
            img_bytes = pix.tobytes("jpeg", jpg_quality=quality)
            if len(img_bytes) > max_bytes:
                # Try lower quality
                img_bytes = pix.tobytes("jpeg", jpg_quality=max(40, quality - 20))
            out_path = f"{out_prefix}-{i}.jpg"
            with open(out_path, "wb") as f:
                f.write(img_bytes)
            written.append(out_path)
        doc.close()
        print(json.dumps({"success": True, "files": written, "count": len(written)}))
    except Exception as e:
        print(json.dumps({"success": False, "error": str(e)}))
        sys.exit(1)

if __name__ == "__main__":
    main()
