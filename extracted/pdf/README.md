# PDF Extraction — `467_26536540990502.pdf`

Corporate brochure for **بهین تجربه آرتیمان (BTA / BA)** — HVAC, chillers, cooling towers, cleanrooms, warehouse fit-out, and related projects.

## Facts
- **32 pages**, landscape scan (`Microsoft: Print To PDF`), **no native text layer**
- Each page embeds **4 JPEG strips** (~2680×500) → **128 unique images**
- Printed page numbers (Persian) generally run cover → ۱…۳۰ → back cover

## Layout of this folder
| Path | Contents |
|------|----------|
| `images/` | Raw embedded JPEG strips per page (`page_XX_img_YY_xrefZ.jpg`) |
| `pages_hires/` | Stitched full-page JPEGs from the 4 strips (~2680×1850) |
| `page_renders/` | 150 DPI PNG renders for preview/vision (~1754×1240) |
| `pages/page_XX.md` | Per-page text/image extraction notes (Persian key terms + English gloss) |
| `full_text.md` | All page notes concatenated |
| `manifest.json` | Machine-readable page/image inventory |
| `images_catalog.csv` | Flat catalog of embedded images |

## Notes
- Tesseract OCR (`fas`/`eng`) was unreliable on this designed layout (angled text, gold-on-black, photo backgrounds); text was extracted via vision from page renders.
- Many interior pages are designed for **portrait reading** while stored landscape in the PDF (rotate 90° CCW to read).
